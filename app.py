from flask import Flask, redirect, request, session, send_file
import os
import re
from io import BytesIO

import pandas as pd
import requests

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "supersecret")

CLIENT_ID = os.environ.get("EXACT_CLIENT_ID")
CLIENT_SECRET = os.environ.get("EXACT_CLIENT_SECRET")
REDIRECT_URI = os.environ.get("EXACT_REDIRECT_URI")

AUTH_URL = "https://start.exactonline.nl/api/oauth2/auth"
TOKEN_URL = "https://start.exactonline.nl/api/oauth2/token"
BASE_URL = "https://start.exactonline.nl/api/v1"

TARGET = "mission freight"


def normalize(x):
    return (x or "").strip().lower()


def is_target(name):
    return TARGET in normalize(name)


def date_fix(v):
    if not v:
        return ""
    m = re.search(r"/Date\((\d+)", str(v))
    if m:
        try:
            return pd.to_datetime(int(m.group(1)), unit="ms").strftime("%Y-%m-%d")
        except Exception:
            return str(v)
    return str(v)


def safe_json(response):
    text = response.text or ""
    if not text.strip():
        return None
    try:
        return response.json()
    except Exception:
        return None


def extract_results(data):
    if isinstance(data, dict):
        d = data.get("d")
        if isinstance(d, dict):
            return d.get("results", [])
        if isinstance(d, list):
            return d
    return []


def get_division(headers):
    res = requests.get(f"{BASE_URL}/current/Me", headers=headers, timeout=30)

    if res.status_code != 200:
        raise Exception(f"Division ophalen mislukt: {res.text}")

    data = safe_json(res)
    if data:
        try:
            return str(data["d"]["results"][0]["CurrentDivision"])
        except Exception:
            pass

    text = res.text or ""
    match = re.search(r"<d:CurrentDivision>(\d+)</d:CurrentDivision>", text)
    if match:
        return match.group(1)

    raise Exception(f"Division niet gevonden: {text[:300]}")


def fetch_all(url, headers):
    all_rows = []
    skip = 0
    page_size = 50

    while True:
        full = f"{url}&$top={page_size}&$skip={skip}"
        res = requests.get(full, headers=headers, timeout=60)

        if res.status_code != 200:
            print("ERROR:", res.status_code, res.text[:300])
            break

        data = safe_json(res)
        if not data:
            print("GEEN JSON:", res.text[:300])
            break

        rows = extract_results(data)

        if not rows:
            break

        all_rows.extend(rows)

        if len(rows) < page_size:
            break

        skip += page_size

    return all_rows


@app.route("/")
def home():
    return """
    <h2>Exact Invoice Tool</h2>
    <a href="/login">Login met Exact</a>
    """


@app.route("/login")
def login():
    if not CLIENT_ID or not CLIENT_SECRET or not REDIRECT_URI:
        return (
            "Environment variables ontbreken. Zet EXACT_CLIENT_ID, "
            "EXACT_CLIENT_SECRET en EXACT_REDIRECT_URI in Render."
        )

    url = (
        f"{AUTH_URL}"
        f"?client_id={CLIENT_ID}"
        f"&redirect_uri={REDIRECT_URI}"
        f"&response_type=code"
        f"&scope=exactonlineapi%20offline_access"
    )
    return redirect(url)


@app.route("/callback")
def callback():
    error = request.args.get("error")
    code = request.args.get("code")

    if error:
        return f"Exact fout: {error}"

    if not code:
        return "Geen code ontvangen van Exact."

    data = {
        "grant_type": "authorization_code",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI,
    }

    res = requests.post(TOKEN_URL, data=data, timeout=30)
    token = safe_json(res)

    if not token:
        return f"Token response niet leesbaar: {res.text}"

    access_token = token.get("access_token")
    if not access_token:
        return f"Geen access token ontvangen: {token}"

    session["access_token"] = access_token
    session["refresh_token"] = token.get("refresh_token")

    return redirect("/sync")


@app.route("/sync")
def sync():
    try:
        token = session.get("access_token")
        if not token:
            return redirect("/login")

        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }

        division = get_division(headers)

        url = (
            f"{BASE_URL}/{division}/purchaseentry/PurchaseEntries"
            f"?$select=InvoiceNumber,EntryDate,AmountDC,SupplierName,Description,Currency,EntryID,Status"
        )

        rows = fetch_all(url, headers)

        results = []

        for r in rows:
            leverancier = r.get("SupplierName", "")

            if not is_target(leverancier):
                continue

            results.append({
                "Factuurnummer": r.get("InvoiceNumber", ""),
                "Datum": date_fix(r.get("EntryDate")),
                "Leverancier": leverancier,
                "Omschrijving": r.get("Description", ""),
                "Totaal": r.get("AmountDC", 0),
                "Valuta": r.get("Currency", ""),
                "Status": r.get("Status", ""),
                "EntryID": r.get("EntryID", ""),
            })

        if not results:
            return "⚠️ Geen Mission Freight boekingen gevonden"

        df = pd.DataFrame(results)

        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        return send_file(
            output,
            download_name="mission_freight_exact.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        return f"Fout: {str(e)}"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
