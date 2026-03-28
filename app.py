from flask import Flask, redirect, request, session, send_file
import os
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


@app.route("/")
def home():
    return """
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Exact Invoice Tool</title>
      </head>
      <body style="font-family: Arial, sans-serif; padding: 24px;">
        <h2>Exact Invoice Tool</h2>
        <p><a href="/login">Login met Exact & download facturen</a></p>
      </body>
    </html>
    """


@app.route("/login")
def login():
    if not CLIENT_ID or not CLIENT_SECRET or not REDIRECT_URI:
        return (
            "Environment variables ontbreken. Zet EXACT_CLIENT_ID, "
            "EXACT_CLIENT_SECRET en EXACT_REDIRECT_URI in Render.",
            500,
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
    code = request.args.get("code")
    error = request.args.get("error")

    if error:
        return f"Exact gaf een fout terug: {error}", 400

    if not code:
        return "Geen authorization code ontvangen van Exact.", 400

    data = {
        "grant_type": "authorization_code",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI,
    }

    response = requests.post(TOKEN_URL, data=data, timeout=30)

    try:
        token = response.json()
    except Exception:
        return f"Token response niet leesbaar: {response.text}", 400

    access_token = token.get("access_token")
    refresh_token = token.get("refresh_token")

    if not access_token:
        return f"Geen access token ontvangen: {token}", 400

    session["access_token"] = access_token
    session["refresh_token"] = refresh_token

    return redirect("/sync")


@app.route("/sync")
def sync():
    token = session.get("access_token")

    if not token:
        return redirect("/login")

    headers = {"Authorization": f"Bearer {token}"}
    division = "110"

    url = f"{BASE_URL}/{division}/purchaseentry/PurchaseEntries?$top=100"

    response = requests.get(url, headers=headers, timeout=30)

    try:
        res = response.json()
    except Exception:
        return f"Exact response niet leesbaar: {response.text}", 400

    results = []

    for item in res.get("d", {}).get("results", []):
        leverancier = item.get("SupplierName", "") or item.get("Supplier", "")
        totaal = item.get("AmountDC", 0)
        factuur = item.get("InvoiceNumber", "")
        datum = item.get("EntryDate", "")

        results.append(
            {
                "Factuurnummer": factuur,
                "Datum": datum,
                "Leverancier": leverancier,
                "Totaal": totaal,
                "KG": 1,
                "Prijs/kg": float(totaal) if totaal else 0,
            }
        )

    if not results:
        results.append(
            {
                "Factuurnummer": "",
                "Datum": "",
                "Leverancier": "Geen resultaten gevonden",
                "Totaal": 0,
                "KG": 0,
                "Prijs/kg": 0,
            }
        )

    df = pd.DataFrame(results)

    output = BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)

    return send_file(
        output,
        download_name="exact_invoices.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
