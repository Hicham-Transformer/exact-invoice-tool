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

MISSION_FREIGHT_NAME = "mission freight"


def exact_date_to_text(value):
    if not value:
        return ""
    if isinstance(value, str):
        m = re.search(r"/Date\((\d+)", value)
        if m:
            try:
                return pd.to_datetime(int(m.group(1)), unit="ms").strftime("%Y-%m-%d")
            except Exception:
                return value
    return str(value)


def extract_results(data):
    d = data.get("d")
    if isinstance(d, dict):
        return d.get("results", [])
    if isinstance(d, list):
        return d
    return []


def fetch_all_pages(base_url: str, headers: dict, page_size: int = 60) -> list:
    all_rows = []
    skip = 0

    while True:
        joiner = "&" if "?" in base_url else "?"
        url = f"{base_url}{joiner}$top={page_size}&$skip={skip}"

        response = requests.get(url, headers=headers, timeout=60)

        if response.status_code != 200:
            raise RuntimeError(f"Fout bij ophalen pagina: {response.text}")

        try:
            data = response.json()
        except Exception:
            raise RuntimeError(f"Response niet leesbaar: {response.text}")

        page_rows = extract_results(data)

        if not page_rows:
            break

        all_rows.extend(page_rows)

        if len(page_rows) < page_size:
            break

        skip += page_size

    return all_rows


def get_current_division(headers: dict) -> str:
    me_res = requests.get(f"{BASE_URL}/current/Me", headers=headers, timeout=30)

    if me_res.status_code != 200:
        raise RuntimeError(f"Fout bij ophalen division: {me_res.text}")

    division = None

    try:
        me_data = me_res.json()
        d = me_data.get("d")

        if isinstance(d, dict) and d.get("results"):
            division = str(d["results"][0]["CurrentDivision"])
        elif isinstance(d, list) and len(d) > 0:
            division = str(d[0]["CurrentDivision"])
    except Exception:
        pass

    if not division:
        text = me_res.text
        match = re.search(r"<d:CurrentDivision>(\d+)</d:CurrentDivision>", text)
        if match:
            division = match.group(1)

    if not division:
        raise RuntimeError(f"Division niet gevonden: {me_res.text}")

    return division


def is_mission_freight(name: str) -> bool:
    return MISSION_FREIGHT_NAME in (name or "").strip().lower()


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
    try:
        token = session.get("access_token")

        if not token:
            return redirect("/login")

        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }

        division = get_current_division(headers)

        purchase_entries_url = (
            f"{BASE_URL}/{division}/purchaseentry/PurchaseEntries"
            f"?$select="
            f"EntryID,"
            f"InvoiceNumber,"
            f"EntryNumber,"
            f"EntryDate,"
            f"AmountDC,"
            f"AmountFC,"
            f"Currency,"
            f"Supplier,"
            f"SupplierName,"
            f"Description,"
            f"YourRef,"
            f"OrderNumber,"
            f"DueDate,"
            f"Journal,"
            f"PaymentCondition,"
            f"Created,"
            f"Modified"
        )

        purchase_invoices_url = (
            f"{BASE_URL}/{division}/purchaseinvoice/PurchaseInvoices"
            f"?$select="
            f"InvoiceID,"
            f"InvoiceNumber,"
            f"InvoiceDate,"
            f"AmountDC,"
            f"AmountFC,"
            f"Currency,"
            f"Supplier,"
            f"SupplierName,"
            f"Description,"
            f"YourRef,"
            f"DueDate,"
            f"Created,"
            f"Modified"
        )

        entry_rows = fetch_all_pages(purchase_entries_url, headers)
        invoice_rows = fetch_all_pages(purchase_invoices_url, headers)

        results = []

        for item in entry_rows:
            leverancier = (item.get("SupplierName") or "").strip()
            if not is_mission_freight(leverancier):
                continue

            totaal_dc = item.get("AmountDC", 0)
            totaal_fc = item.get("AmountFC", 0)

            try:
                prijs_per_kg = float(totaal_dc) if totaal_dc is not None else 0
            except Exception:
                prijs_per_kg = 0

            results.append(
                {
                    "Bron": "PurchaseEntries",
                    "Factuurnummer": item.get("InvoiceNumber", ""),
                    "Boekingsnummer": item.get("EntryNumber", ""),
                    "Document ID": item.get("EntryID", ""),
                    "Factuurdatum": exact_date_to_text(item.get("EntryDate", "")),
                    "Leverancier": leverancier,
                    "Leverancier ID": item.get("Supplier", ""),
                    "Omschrijving": item.get("Description", ""),
                    "Referentie": item.get("YourRef", ""),
                    "Ordernummer": item.get("OrderNumber", ""),
                    "Vervaldatum": exact_date_to_text(item.get("DueDate", "")),
                    "Dagboek": item.get("Journal", ""),
                    "Betalingsconditie": item.get("PaymentCondition", ""),
                    "Valuta": item.get("Currency", ""),
                    "Totaal DC": totaal_dc,
                    "Totaal FC": totaal_fc,
                    "KG": 1,
                    "Prijs/kg": prijs_per_kg,
                    "Aangemaakt": exact_date_to_text(item.get("Created", "")),
                    "Gewijzigd": exact_date_to_text(item.get("Modified", "")),
                }
            )

        for item in invoice_rows:
            leverancier = (item.get("SupplierName") or "").strip()
            if not is_mission_freight(leverancier):
                continue

            totaal_dc = item.get("AmountDC", 0)
            totaal_fc = item.get("AmountFC", 0)

            try:
                prijs_per_kg = float(totaal_dc) if totaal_dc is not None else 0
            except Exception:
                prijs_per_kg = 0

            results.append(
                {
                    "Bron": "PurchaseInvoices",
                    "Factuurnummer": item.get("InvoiceNumber", ""),
                    "Boekingsnummer": "",
                    "Document ID": item.get("InvoiceID", ""),
                    "Factuurdatum": exact_date_to_text(item.get("InvoiceDate", "")),
                    "Leverancier": leverancier,
                    "Leverancier ID": item.get("Supplier", ""),
                    "Omschrijving": item.get("Description", ""),
                    "Referentie": item.get("YourRef", ""),
                    "Ordernummer": "",
                    "Vervaldatum": exact_date_to_text(item.get("DueDate", "")),
                    "Dagboek": "",
                    "Betalingsconditie": "",
                    "Valuta": item.get("Currency", ""),
                    "Totaal DC": totaal_dc,
                    "Totaal FC": totaal_fc,
                    "KG": 1,
                    "Prijs/kg": prijs_per_kg,
                    "Aangemaakt": exact_date_to_text(item.get("Created", "")),
                    "Gewijzigd": exact_date_to_text(item.get("Modified", "")),
                }
            )

        if not results:
            results.append(
                {
                    "Bron": "",
                    "Factuurnummer": "",
                    "Boekingsnummer": "",
                    "Document ID": "",
                    "Factuurdatum": "",
                    "Leverancier": "Geen Mission Freight facturen gevonden",
                    "Leverancier ID": "",
                    "Omschrijving": "",
                    "Referentie": "",
                    "Ordernummer": "",
                    "Vervaldatum": "",
                    "Dagboek": "",
                    "Betalingsconditie": "",
                    "Valuta": "",
                    "Totaal DC": 0,
                    "Totaal FC": 0,
                    "KG": 0,
                    "Prijs/kg": 0,
                    "Aangemaakt": "",
                    "Gewijzigd": "",
                }
            )

        df = pd.DataFrame(results)

        # Duplicaten eruit op bron + factuurnummer + bedrag + datum
        if not df.empty and "Geen Mission Freight facturen gevonden" not in str(df.iloc[0].get("Leverancier", "")):
            df["_dedupe_key"] = (
                df["Bron"].astype(str).fillna("")
                + "|"
                + df["Factuurnummer"].astype(str).fillna("")
                + "|"
                + df["Factuurdatum"].astype(str).fillna("")
                + "|"
                + df["Totaal DC"].astype(str).fillna("")
            )
            df = df.drop_duplicates(subset=["_dedupe_key"]).drop(columns=["_dedupe_key"])

        # Sortering
        sort_cols = [c for c in ["Factuurdatum", "Factuurnummer", "Bron"] if c in df.columns]
        if sort_cols:
            df = df.sort_values(by=sort_cols, ascending=[False] * len(sort_cols))

        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        return send_file(
            output,
            download_name="exact_invoices.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        return f"Fout in sync: {str(e)}", 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
