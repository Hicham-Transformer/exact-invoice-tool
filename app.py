import os
import io
import re
import base64
from urllib.parse import urlencode

import requests
import pandas as pd
from flask import Flask, redirect, request, session, send_file

try:
    import pdfplumber
except Exception:
    pdfplumber = None

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "supersecret-dev-key")

CLIENT_ID = os.environ.get("MS_CLIENT_ID")
CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET")
REDIRECT_URI = os.environ.get("MS_REDIRECT_URI")
TENANT_ID = os.environ.get("MS_TENANT_ID", "common")

AUTH_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
GRAPH_API = "https://graph.microsoft.com/v1.0"

SCOPES = [
    "User.Read",
    "Mail.Read",
    "offline_access",
]

TARGET_FOLDER_NAME = "facturen verwerkt"
TARGET_SENDER = "s.gasior@missionfreight.nl"


@app.route("/")
def home():
    return """
    <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>Outlook PDF Tool</title>
      </head>
      <body style="font-family: Arial, sans-serif; padding: 24px;">
        <h2>Outlook PDF Tool</h2>
        <p><a href="/login">Login met Outlook</a></p>
        <p><a href="/fetch-mails">Haal facturen op</a></p>
      </body>
    </html>
    """


@app.route("/login")
def login():
    if not CLIENT_ID or not CLIENT_SECRET or not REDIRECT_URI:
        return "MS_CLIENT_ID, MS_CLIENT_SECRET of MS_REDIRECT_URI ontbreekt.", 500

    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "response_mode": "query",
        "scope": " ".join(SCOPES),
    }
    return redirect(f"{AUTH_URL}?{urlencode(params)}")


@app.route("/callback")
def callback():
    code = request.args.get("code")
    error = request.args.get("error")
    error_desc = request.args.get("error_description")

    if error:
        return f"Microsoft login fout: {error} - {error_desc}", 400

    if not code:
        return "Geen authorization code ontvangen.", 400

    data = {
        "client_id": CLIENT_ID,
        "scope": " ".join(SCOPES),
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "client_secret": CLIENT_SECRET,
    }

    response = requests.post(TOKEN_URL, data=data, timeout=30)
    token_json = response.json()

    if response.status_code != 200 or "access_token" not in token_json:
        return f"Token error: {token_json}", 400

    session["access_token"] = token_json["access_token"]
    if "refresh_token" in token_json:
        session["refresh_token"] = token_json["refresh_token"]

    return "Login succesvol ✅ Ga nu naar /fetch-mails"


def get_headers():
    token = session.get("access_token")
    if not token:
        raise RuntimeError("Geen access token in session. Log opnieuw in via /login.")
    return {"Authorization": f"Bearer {token}"}


def graph_get(url: str):
    response = requests.get(url, headers=get_headers(), timeout=30)
    data = response.json()
    if response.status_code != 200:
        raise RuntimeError(f"Graph fout {response.status_code}: {data}")
    return data


def graph_get_all(url: str):
    items = []
    while url:
        data = graph_get(url)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def find_folder_under_inbox(folder_name: str):
    folder_name = folder_name.strip().lower()
    url = f"{GRAPH_API}/me/mailFolders/inbox/childFolders?$top=200"
    folders = graph_get_all(url)

    for folder in folders:
        if (folder.get("displayName") or "").strip().lower() == folder_name:
            return folder["id"]

    return None


def parse_decimal(text):
    if not text:
        return ""
    value = str(text).strip()
    value = value.replace(".", "").replace(",", ".")
    try:
        return float(value)
    except Exception:
        return ""


def extract_pdf_text(pdf_bytes: bytes) -> str:
    if pdfplumber is None:
        raise RuntimeError("pdfplumber niet geïnstalleerd")

    text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text


def find_charge(text: str):
    patterns = [
        r"(Import warehouse charges)\s+.*?EUR\s+(\d+(?:[.,]\d+)?)",
        r"(Handling fee)\s+.*?EUR\s+(\d+(?:[.,]\d+)?)",
        r"(Handling charges)\s+.*?EUR\s+(\d+(?:[.,]\d+)?)",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            return match.group(1), parse_decimal(match.group(2))

    return "", ""


def find_kg(text: str):
    # patroon zoals in jouw voorbeeld
    goods_patterns = [
        r"\bCOLLI\b.*?E-COMMERCE\s+(\d+(?:[.,]\d+)?)\s+\d+(?:[.,]\d+)?",
        r"\bCOLLI\b.*?\bBRUTO\b.*?(\d+(?:[.,]\d+)?)",
        r"\bBRUTO\b.*?(\d+(?:[.,]\d+)?)",
        r"(\d+(?:[.,]\d+)?)\s*KG\b",
    ]

    for pattern in goods_patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            return parse_decimal(match.group(1))

    return ""


def extract_pdf_data(pdf_bytes: bytes):
    result = {
        "Factuurnummer": "",
        "Factuurdatum": "",
        "AWB": "",
        "KG": "",
        "Charge omschrijving": "",
        "Charges": "",
        "Prijs_per_KG": "",
        "Status": "",
    }

    try:
        text = extract_pdf_text(pdf_bytes)

        invoice_match = re.search(r"FACTUURNUMMER\s+([0-9]{5,})", text, re.IGNORECASE)
        if invoice_match:
            result["Factuurnummer"] = invoice_match.group(1)

        date_match = re.search(
            r"FACTUURDATUM\s+([0-9]{2}-[A-Za-z]{3}-[0-9]{4})",
            text,
            re.IGNORECASE,
        )
        if date_match:
            result["Factuurdatum"] = date_match.group(1)

        awb_match = re.search(r"\b\d{3}-\d{8}\b", text)
        if awb_match:
            result["AWB"] = awb_match.group(0)

        result["KG"] = find_kg(text)

        charge_name, charge_value = find_charge(text)
        result["Charge omschrijving"] = charge_name
        result["Charges"] = charge_value

        if result["KG"] != "" and result["Charges"] != "":
            result["Prijs_per_KG"] = round(float(result["Charges"]) / float(result["KG"]), 5)

        missing = []
        if not result["Factuurnummer"]:
            missing.append("Factuurnummer")
        if not result["AWB"]:
            missing.append("AWB")
        if result["KG"] == "":
            missing.append("KG")
        if result["Charges"] == "":
            missing.append("Charges")

        result["Status"] = "OK" if not missing else f"Ontbreekt: {', '.join(missing)}"
        return result

    except Exception as e:
        result["Status"] = f"PDF fout: {e}"
        return result


@app.route("/fetch-mails")
def fetch_mails():
    try:
        if "access_token" not in session:
            return redirect("/login")

        folder_id = find_folder_under_inbox(TARGET_FOLDER_NAME)
        if not folder_id:
            return "Map 'Inbox > facturen verwerkt' niet gevonden"

        # haalt ALLE mails op via nextLink pagination
        messages_url = (
            f"{GRAPH_API}/me/mailFolders/{folder_id}/messages"
            "?$top=100"
            "&$select=id,subject,receivedDateTime,from,hasAttachments"
            "&$orderby=receivedDateTime desc"
        )
        messages = graph_get_all(messages_url)

        rows = []

        for mail in messages:
            sender = (
                mail.get("from", {})
                .get("emailAddress", {})
                .get("address", "")
                .strip()
                .lower()
            )

            if TARGET_SENDER not in sender:
                continue

            if not mail.get("hasAttachments"):
                continue

            msg_id = mail["id"]
            attachments_url = f"{GRAPH_API}/me/messages/{msg_id}/attachments?$top=100"
            attachments = graph_get_all(attachments_url)

            for att in attachments:
                if att.get("@odata.type") != "#microsoft.graph.fileAttachment":
                    continue

                filename = att.get("name", "")
                if not filename.lower().endswith(".pdf"):
                    continue

                content_b64 = att.get("contentBytes")
                if not content_b64:
                    continue

                pdf_bytes = base64.b64decode(content_b64)
                parsed = extract_pdf_data(pdf_bytes)

                rows.append(
                    {
                        "Datum email": mail.get("receivedDateTime", ""),
                        "Afzender": sender,
                        "Onderwerp": mail.get("subject", ""),
                        "Bestandsnaam": filename,
                        "Factuurnummer": parsed.get("Factuurnummer", ""),
                        "Factuurdatum": parsed.get("Factuurdatum", ""),
                        "AWB": parsed.get("AWB", ""),
                        "KG": parsed.get("KG", ""),
                        "Charge omschrijving": parsed.get("Charge omschrijving", ""),
                        "Charges": parsed.get("Charges", ""),
                        "Prijs_per_KG": parsed.get("Prijs_per_KG", ""),
                        "Status": parsed.get("Status", ""),
                    }
                )

        if not rows:
            return "Geen PDF facturen gevonden van s.gasior@missionfreight.nl in map 'Inbox > facturen verwerkt'"

        df = pd.DataFrame(rows)

        # dubbele pdf’s / dubbele factuurnummers verminderen
        if "Factuurnummer" in df.columns:
            df = df.sort_values(by=["Datum email"], ascending=False)
            df = df.drop_duplicates(subset=["Factuurnummer", "Bestandsnaam"], keep="first")

        output = io.BytesIO()
        df.to_excel(output, index=False, engine="openpyxl")
        output.seek(0)

        return send_file(
            output,
            download_name="outlook_facturen.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        return f"Fout in fetch-mails: {e}", 500


@app.route("/health")
def health():
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
