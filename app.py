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

TARGET_SUPPLIER = "mission freight"


def normalize(text):
    return (text or "").strip().lower()


def is_mission_freight(name):
    return TARGET_SUPPLIER in normalize(name)


def exact_date_to_text(value):
    if not value:
        return ""
    m = re.search(r"/Date\((\d+)", str(value))
    if m:
        return pd.to_datetime(int(m.group(1)), unit="ms").strftime("%Y-%m-%d")
    return str(value)


def parse_exact(data):
    if isinstance(data, dict):
        d = data.get("d")
        if isinstance(d, dict):
            return d.get("results", [])
        if isinstance(d, list):
            return d
    return []


def fetch_all(url, headers):
    results = []
    skip = 0

    while True:
        full_url = f"{url}&$top=50&$skip={skip}"
        res = requests.get(full_url, headers=headers, timeout=60)

        if res.status_code != 200:
            print("ERROR:", res.text[:300])
            break

        try:
            data = res.json()
        except Exception:
            print("JSON FAIL:", res.text[:300])
            break

        rows = parse_exact(data)

        if not rows:
            break

        results.extend(rows)

        if len(rows) < 50:
            break

        skip += 50

    return results


def get_division(headers):
    res = requests.get(f"{BASE_URL}/current/Me", headers=headers)

    try:
        data = res.json()
        return str(data["d"]["results"][0]["CurrentDivision"])
    except Exception:
        raise Exception("Division ophalen mislukt")


@app.route("/")
def home():
    return """
    <h2>Exact Invoice Tool</h2>
    <a href="/login">Login met Exact</a>
    """


@app.route("/login")
def login():
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
    code = request.args.get("code")

    data = {
        "grant_type": "authorization_code",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI,
    }

    res = requests.post(TOKEN_URL, data=data)

    token = res.json()

    session["access_token"] = token["access_token"]

    return redirect("/sync")


@app.route("/sync")
def sync():
    try:
        token = session.get("access_token")
        headers = {"Authorization": f"Bearer {token}"}

        division = get_division(headers)

        url = (
            f"{BASE_URL}/{division}/purchaseentry/PurchaseEntries?"
            f"$select=InvoiceNumber,EntryDate,AmountDC,SupplierName,Description"
        )

        rows = fetch_all(url, headers)

        results = []

        for r in rows:
            leverancier = r.get("SupplierName", "")

            if not is_mission_freight(leverancier):
                continue

            results.append({
                "Factuurnummer": r.get("InvoiceNumber"),
                "Datum": exact_date_to_text(r.get("EntryDate")),
                "Leverancier": leverancier,
                "Omschrijving": r.get("Description"),
                "Totaal": r.get("AmountDC")
            })

        # 🔥 fallback zodat je nooit leeg blijft
        if not results:
            return "⚠️ Geen Mission Freight facturen gevonden — check naam in Exact"

        df = pd.DataFrame(results)

        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="mission_freight.xlsx",
            as_attachment=True
        )

    except Exception as e:
        return f"Fout: {str(e)}"


if __name__ == "__main__":
    app.run()
