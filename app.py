import os
import io
import re
import base64
import traceback
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

# =========================
# CONFIG
# =========================
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

TARGET_SENDER = "s.gasior@missionfreight.nl"
TARGET_FOLDER_NAME = "facturen verwerkt"


# =========================
# HOME
# =========================
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
        <p><a href="/health">Health check</a></p>
      </body>
    </html>
    """


# =========================
# LOGIN
# =========================
@app.route("/login")
def login():
    if not CLIENT_ID or not CLIENT_SECRET or not REDIRECT_URI:
        return (
            "MS_CLIENT_ID, MS_CLIENT_SECRET of MS_REDIRECT_URI ontbreekt in Render Environment.",
            500,
        )

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
        return "Geen authorization code ontvangen van Microsoft.", 400

    data = {
        "client_id": CLIENT_ID,
        "scope": " ".join(SCOPES),
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "client_secret": CLIENT_SECRET,
    }

    response = requests.post(TOKEN_URL, data=data, timeout=30)
    try:
        token_json = response.json()
    except Exception:
        return f"Token response niet leesbaar: {response.text}", 400

    if response.status_code != 200 or "access_token" not in token_json:
        return f"Token error: {token_json}", 400

    session["access_token"] = token_json["access_token"]
    if "refresh_token" in token_json:
        session["refresh_token"] = token_json["refresh_token"]

    return "Login succesvol ✅ Ga nu naar /fetch-mails"


# =========================
# HELPERS
# =========================
def get_headers():
    token = session.get("access_token")
    if not token:
        raise RuntimeError("Geen access token in session. Log opnieuw in via /login.")
    return {"Authorization": f"Bearer {token}"}


def graph_get(url):
    response = requests.get(url, headers=get_headers(), timeout=30)

    try:
        data = response.json()
    except Exception:
        raise RuntimeError(
            f"Graph response niet leesbaar. HTTP {response.status_code}: {response.text}"
        )

    if response.status_code != 200:
        raise RuntimeError(f"Graph fout {response.status_code}: {data}")

    return data


def graph_get_all(url):
    """
    Haalt alle pagina's op via @odata.nextLink
    """
    all_items = []

    while url:
        data = graph_get(url)
        all_items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")

    return all_items


def find_folder_recursive(target_name):
    """
    Zoekt recursief in ALLE mappen en submappen op displayName.
    """
    target_name = target_name.strip().lower()

    def search(folder_id=None):
        if folder_id:
            url = f"{GRAPH_API}/me/mailFolders/{folder_id}/childFolders?$top=200"
        else:
            url = f"{GRAPH_API}/me/mailFolders?$top=200"

        folders = graph_get_all(url)

        for folder in folders:
            name = (folder.get("displayName") or "").strip().lower()

            if name == target_name:
                return folder["id"]

            found = search(folder["id"])
            if found:
                return found

        return None

    return search()


def parse_decimal(value):
    """
    Zet '102,54' om naar 102.54
    """
    if value in (None, ""):
        return ""

    text = str(value).strip()
    text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except Exception:
        return ""


def extract_pdf_text(pdf_bytes):
    if pdfplumber is None:
        raise RuntimeError("pdfplumber niet geïnstalleerd. Zet pdfplumber in requirements.txt.")

    text = ""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    return text


def extract_pdf_data(pdf_bytes):
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

        # Factuurnummer
        invoice_match = re.search(r"FACTUURNUMMER\s+([0-9]{5,})", text, re.IGNORECASE)
        if invoice_match:
            result["Factuurnummer"] = invoice_match.group(1)

        # Factuurdatum
        date_match = re.search(
            r"FACTUURDATUM\s+([0-9]{2}-[A-Za-z]{3}-[0-9]{4})",
            text,
            re.IGNORECASE,
        )
        if date_match:
            result["Factuurdatum"] = date_match.group(1)

        # AWB
        awb_match = re.search(r"\b\d{3}-\d{8}\b", text)
        if awb_match:
            result["AWB"] = awb_match.group(0)

        # KG - afgestemd op Mission Freight layout
        goods_match = re.search(
            r"\bCOLLI\b.*?E-COMMERCE\s+(\d+(?:[.,]\d+)?)\s+\d+(?:[.,]\d+)?",
            text,
            re.IGNORECASE | re.DOTALL,
        )
        if goods_match:
            result["KG"] = parse_decimal(goods_match.group(1))
        else:
            kg_match = re.search(r"(\d+(?:[.,]\d+)?)\s*KG\b", text, re.IGNORECASE)
            if kg_match:
                result["KG"] = parse_decimal(kg_match.group(1))

        # Charges: warehouse / handling fee / handling charges
        charge_patterns = [
            r"(Import warehouse charges)\s+.*?EUR\s+(\d+(?:[.,]\d+)?)",
            r"(Handling fee)\s+.*?EUR\s+(\d+(?:[.,]\d+)?)",
            r"(Handling charges)\s+.*?EUR\s+(\d+(?:[.,]\d+)?)",
        ]

        for pattern in charge_patterns:
            charge_match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if charge_match:
                result["Charge omschrijving"] = charge_match.group(1)
                result["Charges"] = parse_decimal(charge_match.group(2))
                break

        # Prijs per KG
        if result["KG"] != "" and result["Charges"] != "":
            try:
                result["Prijs_per_KG"] = round(
                    float(result["Charges"]) / float(result["KG"]),
                    5,
                )
            except Exception:
                result["Prijs_per_KG"] = ""

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


# =========================
# FETCH MAILS
# =========================
@app.route("/fetch-mails")
def fetch_mails():
    try:
        if "access_token" not in session:
            return redirect("/login")

        folder_id = find_folder_recursive(TARGET_FOLDER_NAME)

        if not folder_id:
            return "Map 'facturen verwerkt' niet gevonden"

        messages_url = (
            f"{GRAPH_API}/me/mailFolders/{folder_id}/messages"
            "?$top=100"
            "&$select=id,subject,receivedDateTime,from,hasAttachments"
        )
        messages = graph_get_all(messages_url)

        rows = []

        for mail in messages:
            try:
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
                attachments_url = f"{GRAPH_API}/me/messages/{msg_id}/attachments"
                attachments = graph_get_all(attachments_url)

                for att in attachments:
                    try:
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

                    except Exception as att_err:
                        rows.append(
                            {
                                "Datum email": mail.get("receivedDateTime", ""),
                                "Afzender": sender,
                                "Onderwerp": mail.get("subject", ""),
                                "Bestandsnaam": att.get("name", ""),
                                "Factuurnummer": "",
                                "Factuurdatum": "",
                                "AWB": "",
                                "KG": "",
                                "Charge omschrijving": "",
                                "Charges": "",
                                "Prijs_per_KG": "",
                                "Status": f"Attachment fout: {att_err}",
                            }
                        )

            except Exception as mail_err:
                rows.append(
                    {
                        "Datum email": mail.get("receivedDateTime", ""),
                        "Afzender": "",
                        "Onderwerp": mail.get("subject", ""),
                        "Bestandsnaam": "",
                        "Factuurnummer": "",
                        "Factuurdatum": "",
                        "AWB": "",
                        "KG": "",
                        "Charge omschrijving": "",
                        "Charges": "",
                        "Prijs_per_KG": "",
                        "Status": f"Mail fout: {mail_err}",
                    }
                )

        if not rows:
            return (
                "Geen PDF facturen gevonden van s.gasior@missionfreight.nl "
                "in map 'facturen verwerkt'"
            )

        df = pd.DataFrame(rows)

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
        return (
            "<h3>Fout in fetch-mails</h3>"
            f"<pre>{str(e)}</pre>"
            f"<pre>{traceback.format_exc()}</pre>"
        ), 500


@app.route("/health")
def health():
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
