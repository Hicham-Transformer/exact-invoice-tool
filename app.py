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


def parse_exact_results(data):
    if isinstance(data, dict):
        d = data.get("d")
        if isinstance(d, dict):
            return d.get("results", [])
        if isinstance(d, list):
            return d
    return []


def fetch_all_pages(base_url, headers, page_size=20, max_pages=500):
    all_rows = []
    skip = 0
    page_count = 0

    while page_count < max_pages:
        joiner = "&" if "?" in base_url else "?"
        url = f"{base_url}{joiner}$top={page_size}&$skip={skip}"

        res = requests.get(url, headers=headers, timeout=60)

        if res.status_code != 200:
            print("HTTP error:", res.status_code, res.text[:300])
            break

        if not res.text or not res.text.strip():
            print("Lege response op:", url)
            break

        try:
            data = res.json()
        except Exception:
            print("JSON parse error op:", url)
            print("Response preview:", res.text[:300])
            break

        rows = parse_exact_results(data)

        if not rows:
            break

        all_rows.extend(rows)

        if len(rows) < page_size:
            break

        skip += page_size
        page_count += 1

    return all_rows


def get_division(headers):
    res = requests.get(f"{BASE_URL}/current/Me", headers=headers, timeout=30)

    if res.status_code != 200:
        raise RuntimeError(f"Fout bij ophalen division: {res.text}")

    if not res.text or not res.text.strip():
        raise RuntimeError("Lege response bij ophalen division")

    division = None

    try:
        data = res.json()
        d = data.get("d")
        if isinstance(d, dict) and d.get("results"):
            division = str(d["results"][0]["CurrentDivision"])
        elif isinstance(d, list) and len(d) > 0:
            division = str(d[0]["CurrentDivision"])
    except Exception:
        pass

    if not division:
        match = re.search(r"<d:CurrentDivision>(\d+)</d:CurrentDivision>", res.text)
        if match:
            division = match.group(1)

    if not division:
        raise RuntimeError(f"Division niet gevonden: {res.text[:500]}")

    return division


def is_mission_freight(name):
    return MISSION_FREIGHT_NAME in (name or "").strip().lower()


def build_entries_map(entry_rows):
    entries_map = {}
    for item in entry_rows:
        entry_id = str(item.get("EntryID") or "")
        if entry_id:
            entries_map[entry_id] = item
    return entries_map


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

    token_res = requests.post(TOKEN_URL, data=data, timeout=30)

    if not token_res.text or not token_res.text.strip():
        return "Lege token response van Exact", 400

    try:
        token = token_res.json()
    except Exception:
        return f"Token response niet leesbaar: {token_res.text}", 400

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

        division = get_division(headers)

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

        purchase_entry_lines_url = (
            f"{BASE_URL}/{division}/purchaseentry/PurchaseEntryLines"
            f"?$select="
            f"ID,"
            f"EntryID,"
            f"LineNumber,"
            f"Description,"
            f"AmountDC,"
            f"AmountFC,"
            f"ItemCode,"
            f"ItemDescription,"
            f"Quantity,"
            f"UnitCode"
        )

        entry_rows = fetch_all_pages(purchase_entries_url, headers, page_size=20)
        invoice_rows = fetch_all_pages(purchase_invoices_url, headers, page_size=20)
        line_rows = fetch_all_pages(purchase_entry_lines_url, headers, page_size=20)

        entries_map = build_entries_map(entry_rows)
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
                    "Regel ID": "",
                    "Regelnummer": "",
                    "Factuurdatum": exact_date_to_text(item.get("EntryDate", "")),
                    "Leverancier": leverancier,
                    "Leverancier ID": item.get("Supplier", ""),
                    "Omschrijving": item.get("Description", ""),
                    "Regelomschrijving": "",
                    "ItemCode": "",
                    "ItemOmschrijving": "",
                    "Aantal": "",
                    "Eenheid": "",
                    "Referentie": item.get("YourRef", ""),
                    "Ordernummer": item.get("OrderNumber", ""),
                    "Vervaldatum": exact_date_to_text(item.get("DueDate", "")),
                    "Dagboek": item.get("Journal", ""),
                    "Betalingsconditie": item.get("PaymentCondition", ""),
                    "Valuta": item.get("Currency", ""),
                    "Totaal DC": totaal_dc,
                    "Totaal FC": totaal_fc,
                    "Regel Totaal DC": "",
                    "Regel Totaal FC": "",
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
                    "Regel ID": "",
                    "Regelnummer": "",
                    "Factuurdatum": exact_date_to_text(item.get("InvoiceDate", "")),
                    "Leverancier": leverancier,
                    "Leverancier ID": item.get("Supplier", ""),
                    "Omschrijving": item.get("Description", ""),
                    "Regelomschrijving": "",
                    "ItemCode": "",
                    "ItemOmschrijving": "",
                    "Aantal": "",
                    "Eenheid": "",
                    "Referentie": item.get("YourRef", ""),
                    "Ordernummer": "",
                    "Vervaldatum": exact_date_to_text(item.get("DueDate", "")),
                    "Dagboek": "",
                    "Betalingsconditie": "",
                    "Valuta": item.get("Currency", ""),
                    "Totaal DC": totaal_dc,
                    "Totaal FC": totaal_fc,
                    "Regel Totaal DC": "",
                    "Regel Totaal FC": "",
                    "KG": 1,
                    "Prijs/kg": prijs_per_kg,
                    "Aangemaakt": exact_date_to_text(item.get("Created", "")),
                    "Gewijzigd": exact_date_to_text(item.get("Modified", "")),
                }
            )

        for item in line_rows:
            entry_id = str(item.get("EntryID") or "")
            header = entries_map.get(entry_id, {})

            leverancier = (header.get("SupplierName") or "").strip()
            if not is_mission_freight(leverancier):
                continue

            totaal_dc = header.get("AmountDC", 0)
            totaal_fc = header.get("AmountFC", 0)
            regel_dc = item.get("AmountDC", 0)
            regel_fc = item.get("AmountFC", 0)

            try:
                prijs_per_kg = float(regel_dc) if regel_dc is not None else 0
            except Exception:
                prijs_per_kg = 0

            results.append(
                {
                    "Bron": "PurchaseEntryLines",
                    "Factuurnummer": header.get("InvoiceNumber", ""),
                    "Boekingsnummer": header.get("EntryNumber", ""),
                    "Document ID": entry_id,
                    "Regel ID": item.get("ID", ""),
                    "Regelnummer": item.get("LineNumber", ""),
                    "Factuurdatum": exact_date_to_text(header.get("EntryDate", "")),
                    "Leverancier": leverancier,
                    "Leverancier ID": header.get("Supplier", ""),
                    "Omschrijving": header.get("Description", ""),
                    "Regelomschrijving": item.get("Description", ""),
                    "ItemCode": item.get("ItemCode", ""),
                    "ItemOmschrijving": item.get("ItemDescription", ""),
                    "Aantal": item.get("Quantity", ""),
                    "Eenheid": item.get("UnitCode", ""),
                    "Referentie": header.get("YourRef", ""),
                    "Ordernummer": header.get("OrderNumber", ""),
                    "Vervaldatum": exact_date_to_text(header.get("DueDate", "")),
                    "Dagboek": header.get("Journal", ""),
                    "Betalingsconditie": header.get("PaymentCondition", ""),
                    "Valuta": header.get("Currency", ""),
                    "Totaal DC": totaal_dc,
                    "Totaal FC": totaal_fc,
                    "Regel Totaal DC": regel_dc,
                    "Regel Totaal FC": regel_fc,
                    "KG": item.get("Quantity", 1) or 1,
                    "Prijs/kg": prijs_per_kg,
                    "Aangemaakt": exact_date_to_text(header.get("Created", "")),
                    "Gewijzigd": exact_date_to_text(header.get("Modified", "")),
                }
            )

        if not results:
            results.append(
                {
                    "Bron": "",
                    "Factuurnummer": "",
                    "Boekingsnummer": "",
                    "Document ID": "",
                    "Regel ID": "",
                    "Regelnummer": "",
                    "Factuurdatum": "",
                    "Leverancier": "Geen Mission Freight facturen gevonden",
                    "Leverancier ID": "",
                    "Omschrijving": "",
                    "Regelomschrijving": "",
                    "ItemCode": "",
                    "ItemOmschrijving": "",
                    "Aantal": "",
                    "Eenheid": "",
                    "Referentie": "",
                    "Ordernummer": "",
                    "Vervaldatum": "",
                    "Dagboek": "",
                    "Betalingsconditie": "",
                    "Valuta": "",
                    "Totaal DC": 0,
                    "Totaal FC": 0,
                    "Regel Totaal DC": "",
                    "Regel Totaal FC": "",
                    "KG": 0,
                    "Prijs/kg": 0,
                    "Aangemaakt": "",
                    "Gewijzigd": "",
                }
            )

        df = pd.DataFrame(results)

        if not df.empty and "Geen Mission Freight facturen gevonden" not in str(df.iloc[0].get("Leverancier", "")):
            df["_dedupe_key"] = (
                df["Bron"].astype(str).fillna("")
                + "|"
                + df["Factuurnummer"].astype(str).fillna("")
                + "|"
                + df["Factuurdatum"].astype(str).fillna("")
                + "|"
                + df["Document ID"].astype(str).fillna("")
                + "|"
                + df["Regelnummer"].astype(str).fillna("")
            )
            df = df.drop_duplicates(subset=["_dedupe_key"]).drop(columns=["_dedupe_key"])

        sort_cols = [c for c in ["Factuurdatum", "Factuurnummer", "Bron", "Regelnummer"] if c in df.columns]
        if sort_cols:
            df = df.sort_values(by=sort_cols, ascending=[False] * len(sort_cols))

        output = BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        return send_file(
            output,
            download_name="exact_invoices_enterprise.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        return f"Fout in sync: {str(e)}", 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
