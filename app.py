import os
import io
import re
import base64
import requests
import pandas as pd
from flask import Flask, redirect, request, session, send_file
from urllib.parse import urlencode

try:
    import pdfplumber
except Exception:
    pdfplumber = None

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "supersecret")

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


@app.route("/")
def home():
    return """
    <h2>Outlook PDF Tool</h2>
    <p><a href="/login">Login met Outlook</a></p>
    <p><a href="/fetch-mails">Haal facturen op</a></p>
    """


@app.route("/login")
def login():
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

    if not code:
        return "Geen code ontvangen van Microsoft."

    data = {
        "client_id": CLIENT_ID,
        "scope": " ".join(SCOPES),
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "client_secret": CLIENT_SECRET,
    }

    token_res = requests.post(TOKEN_URL, data=data, timeout=30)
    token_json = token_res.json()

    if "access_token" not in token_json:
        return f"Token error: {token_json}"

    session["access_token"] = token_json["access_token"]
    return "Login succesvol ✅ Ga nu naar /fetch-mails"


def get_headers():
    token = session.get("access_token")
    return {"Authorization": f"Bearer {token}"}


def graph_get(url):
    response = requests.get(url, headers=get_headers(), timeout=30)

    try:
        data = response.json()
    except Exception:
        data = None

    if response.status_code != 200:
        raise RuntimeError(f"Graph fout {response.status_code}: {response.text}")

    return data


def find_folder_recursive(target_name):
    target_name = target_name.strip().lower()

    def search(folder_id=None):
        if folder_id:
            url = f"{GRAPH_API}/me/mailFolders/{folder_id}/childFolders?$top=200"
        else:
            url = f"{GRAPH_API}/me/mailFolders?$top=200"

        data = graph_get(url)

        for folder in data.get("value", []):
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
    Zet Europese bedragen zoals 102,54 om naar float 102.54
    """
    if not value:
        return ""

    value = str(value).strip()
    value = value.replace(".", "").replace(",", ".")
    try:
        return float(value)
    except Exception:
        return ""


def extract_pdf_data(pdf_bytes):
    result = {
        "Factuurnummer": "",
        "Factuurdatum": "",
        "AWB": "",
        "KG": "",
        "Charges": "",
        "Charge_omschrijving": "",
        "Prijs_per_KG": "",
        "Status": "",
    }

    if pdfplumber is None:
        result["Status"] = "pdfplumber niet geïnstalleerd"
        return result

    text = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
    except Exception as e:
        result["Status"] = f"PDF leesfout: {e}"
        return result

    # Factuurnummer
    invoice_match = re.search(r"FACTUURNUMMER\s+([0-9]{5,})", text, re.IGNORECASE)
    if invoice_match:
        result["Factuurnummer"] = invoice_match.group(1)

    # Factuurdatum
    date_match = re.search(r"FACTUURDATUM\s+([0-9]{2}-[A-Za-z]{3}-[0-9]{4})", text, re.IGNORECASE)
    if date_match:
        result["Factuurdatum"] = date_match.group(1)

    # AWB
    awb_match = re.search(r"\b\d{3}-\d{8}\b", text)
    if awb_match:
        result["AWB"] = awb_match.group(0)

    # KG: pak eerst de regel met COLLI / E-COMMERCE / goederen en haal het eerste hele gewicht eruit
    # Voor jouw voorbeeld: "47 COLLI E-COMMERCE 505 505,00"
    goods_match = re.search(
        r"\bCOLLI\b.*?E-COMMERCE\s+(\d+(?:[.,]\d+)?)\s+\d+(?:[.,]\d+)?",
        text,
        re.IGNORECASE | re.DOTALL,
    )
    if goods_match:
        result["KG"] = parse_decimal(goods_match.group(1))
    else:
        # fallback: eerste "xxx KG"
        kg_match = re.search(r"(\d+(?:[.,]\d+)?)\s*KG\b", text, re.IGNORECASE)
        if kg_match:
            result["KG"] = parse_decimal(kg_match.group(1))

    # Charges: Import warehouse charges / Handling fee / Handling charges
    charge_patterns = [
        r"(Import warehouse charges)\s+.*?EUR\s+(\d+(?:[.,]\d+)?)",
        r"(Handling fee)\s+.*?EUR\s+(\d+(?:[.,]\d+)?)",
        r"(Handling charges)\s+.*?EUR\s+(\d+(?:[.,]\d+)?)",
    ]

    for pattern in charge_patterns:
        charge_match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if charge_match:
            result["Charge_omschrijving"] = charge_match.group(1)
            result["Charges"] = parse_decimal(charge_match.group(2))
            break

    if result["KG"] != "" and result["Charges"] != "":
        try:
            result["Prijs_per_KG"] = round(float(result["Charges"]) / float(result["KG"]), 5)
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


@app.route("/fetch-mails")
def fetch_mails():
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
    messages_data = graph_get(messages_url)

    rows = []

    for mail in messages_data.get("value", []):
        sender = (
            mail.get("from", {})
            .get("emailAddress", {})
            .get("address", "")
            .strip()
            .lower()
        )

        # minder streng: contains in plaats van exact
        if TARGET_SENDER not in sender:
            continue

        if not mail.get("hasAttachments"):
            continue

        msg_id = mail["id"]

        attachments_url = f"{GRAPH_API}/me/messages/{msg_id}/attachments"
        attachments_data = graph_get(attachments_url)

        for att in attachments_data.get("value", []):
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
                    "Charge omschrijving": parsed.get("Charge_omschrijving", ""),
                    "Charges": parsed.get("Charges", ""),
                    "Prijs_per_KG": parsed.get("Prijs_per_KG", ""),
                    "Status": parsed.get("Status", ""),
                }
            )

    if not rows:
        return "Geen PDF facturen gevonden van s.gasior@missionfreight.nl in map 'facturen verwerkt'"

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


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
