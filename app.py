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


def fetch_all_pages(url, headers):
    all_rows = []
    skip = 0
    page_size = 20  # 🔥 kleiner = stabieler

    while True:
        full_url = f"{url}&$top={page_size}&$skip={skip}"

        res = requests.get(full_url, headers=headers)

        if res.status_code != 200:
            print("❌ pagina error:", res.text)
            break  # ❗ NIET crashen, gewoon stoppen

        try:
            data = res.json()
        except:
            print("❌ json error:", res.text)
            break

        rows = data.get("d", {}).get("results", [])

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
    <a href="/login">Login met Exact & download facturen</a>
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

    token = requests.post(TOKEN_URL, data=data).json()

    session["access_token"] = token.get("access_token")

    return redirect("/sync")


@app.route("/sync")
def sync():
    try:
        token = session.get("access_token")
        headers = {"Authorization": f"Bearer {token}"}

        # division ophalen
        me = requests.get(f"{BASE_URL}/current/Me", headers=headers).json()
        division = str(me["d"]["results"][0]["CurrentDivision"])

        # 🔥 SIMPELE QUERY (geen complexe select)
        url_entries = f"{BASE_URL}/{division}/purchaseentry/PurchaseEntries"
        url_invoices = f"{BASE_URL}/{division}/purchaseinvoice/PurchaseInvoices"

        entries = fetch_all_pages(url_entries, headers)
        invoices = fetch_all_pages(url_invoices, headers)

        results = []

        # Entries
        for item in entries:
            leverancier = (item.get("SupplierName") or "").lower()

            if "mission freight" not in leverancier:
                continue

            results.append({
                "Bron": "Entries",
                "Factuurnummer": item.get("InvoiceNumber"),
                "Datum": item.get("EntryDate"),
                "Leverancier": item.get("SupplierName"),
                "Totaal": item.get("AmountDC"),
            })

        # Invoices
        for item in invoices:
            leverancier = (item.get("SupplierName") or "").lower()

            if "mission freight" not in leverancier:
                continue

            results.append({
                "Bron": "Invoices",
                "Factuurnummer": item.get("InvoiceNumber"),
                "Datum": item.get("InvoiceDate"),
                "Leverancier": item.get("SupplierName"),
                "Totaal": item.get("AmountDC"),
            })

        if not results:
            return "Geen facturen gevonden"

        df = pd.DataFrame(results)

        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="exact_invoices.xlsx",
            as_attachment=True
        )

    except Exception as e:
        return f"Fout in sync: {str(e)}"


if __name__ == "__main__":
    app.run()
