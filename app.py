from flask import Flask, redirect, request, session, send_file
import os
import re
from io import BytesIO
import pandas as pd
import requests

app = Flask(__name__)
app.secret_key = "supersecret"

CLIENT_ID = os.environ.get("EXACT_CLIENT_ID")
CLIENT_SECRET = os.environ.get("EXACT_CLIENT_SECRET")
REDIRECT_URI = os.environ.get("EXACT_REDIRECT_URI")

AUTH_URL = "https://start.exactonline.nl/api/oauth2/auth"
TOKEN_URL = "https://start.exactonline.nl/api/oauth2/token"
BASE_URL = "https://start.exactonline.nl/api/v1"

TARGET = "mission freight"


def normalize(x):
    return (x or "").lower().strip()


def is_target(name):
    return TARGET in normalize(name)


def date_fix(v):
    if not v:
        return ""
    m = re.search(r"/Date\((\d+)", str(v))
    if m:
        return pd.to_datetime(int(m.group(1)), unit="ms").strftime("%Y-%m-%d")
    return str(v)


def get_division(headers):
    res = requests.get(f"{BASE_URL}/current/Me", headers=headers)
    data = res.json()
    return str(data["d"]["results"][0]["CurrentDivision"])


def fetch_all(url, headers):
    all_rows = []
    skip = 0

    while True:
        full = f"{url}&$top=50&$skip={skip}"
        res = requests.get(full, headers=headers)

        if res.status_code != 200:
            print(res.text)
            break

        data = res.json()
        rows = data.get("d", {}).get("results", [])

        if not rows:
            break

        all_rows.extend(rows)

        if len(rows) < 50:
            break

        skip += 50

    return all_rows


@app.route("/")
def home():
    return '<a href="/login">Start</a>'


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

        url = f"{BASE_URL}/{division}/purchaseentry/PurchaseEntries?"

        rows = fetch_all(url, headers)

        results = []

        for r in rows:
            leverancier = r.get("SupplierName", "")

            if not is_target(leverancier):
                continue

            results.append({
                "Factuurnummer": r.get("InvoiceNumber"),
                "Datum": date_fix(r.get("EntryDate")),
                "Leverancier": leverancier,
                "Omschrijving": r.get("Description"),
                "Totaal": r.get("AmountDC"),
                "Status": r.get("Status"),
                "Valuta": r.get("Currency"),
                "EntryID": r.get("EntryID"),
            })

        if not results:
            return "⚠️ Geen Mission Freight boekingen gevonden (controleer Exact omgeving)"

        df = pd.DataFrame(results)

        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="mission_freight_exact.xlsx",
            as_attachment=True
        )

    except Exception as e:
        return f"Fout: {str(e)}"


if __name__ == "__main__":
    app.run()
