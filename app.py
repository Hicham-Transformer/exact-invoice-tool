from flask import Flask, request, redirect, session, send_file
import requests
import pandas as pd
import os
from io import BytesIO

app = Flask(__name__)
app.secret_key = "supersecret"

# ENV VARS
CLIENT_ID = os.environ.get("EXACT_CLIENT_ID")
CLIENT_SECRET = os.environ.get("EXACT_CLIENT_SECRET")
REDIRECT_URI = os.environ.get("EXACT_REDIRECT_URI")

AUTH_URL = "https://start.exactonline.nl/api/oauth2/auth"
TOKEN_URL = "https://start.exactonline.nl/api/oauth2/token"
BASE_URL = "https://start.exactonline.nl/api/v1"

# 👉 LOGIN MET EXACT
@app.route("/login")
def login():
    url = f"{AUTH_URL}?client_id={CLIENT_ID}&redirect_uri={REDIRECT_URI}&response_type=code"
    return redirect(url)

# 👉 CALLBACK
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

    response = requests.post(TOKEN_URL, data=data)
    token = response.json()

    session["access_token"] = token.get("access_token")

    return redirect("/sync")

# 👉 FACTUREN OPHALEN
@app.route("/sync")
def sync():
    token = session.get("access_token")

    if not token:
        return redirect("/login")

    headers = {"Authorization": f"Bearer {token}"}

    division = "110"  # jouw administratie

    url = f"{BASE_URL}/{division}/purchaseentry/PurchaseEntries?$filter=Supplier eq 'Mission Freight B.V.'"

    res = requests.get(url, headers=headers).json()

    results = []

    for item in res.get("d", {}).get("results", []):
        totaal = item.get("AmountDC", 0)
        factuur = item.get("InvoiceNumber", "")
        datum = item.get("EntryDate", "")

        # Simpele berekening placeholder
        kg = 1
        prijs_per_kg = totaal / kg if kg else 0

        results.append({
            "Factuurnummer": factuur,
            "Datum": datum,
            "Totaal": totaal,
            "KG": kg,
            "Prijs/kg": prijs_per_kg
        })

    df = pd.DataFrame(results)

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="exact_invoices.xlsx",
        as_attachment=True
    )

@app.route("/")
def home():
    return '<a href="/login">Login met Exact & download facturen</a>'
