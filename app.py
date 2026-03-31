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
TARGET_FOLDER_PATH = ["Postvak IN", "Submap", "facturen verwerkt"]


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


def get_child_folders(folder_id=None):
    if folder_id:
        url = f"{GRAPH_API}/me/mailFolders/{folder_id}/childFolders?$top=200"
    else:
        url = f"{GRAPH_API}/me/mailFolders?$top=200"

    data = graph_get(url)
    return data.get("value", [])


def find_folder_by_path(path_parts):
    current_parent_id = None

    for part in path_parts:
        folders = get_child_folders(current_parent_id)

        found = None
        for folder in folders:
            if folder.get("displayName", "").strip().lower() == part.strip().lower():
                found = folder
                break

        if not found:
            return None

        current_parent_id = found["id"]

    return current_parent_id


def extract_pdf_data(pdf_bytes):
    result = {
        "AWB": "",
        "KG": "",
        "Charges": "",
        "Prijs_per_KG": "",
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

    awb_match = re.search(r"\b\d{3}-\d{8}\b", text)

    kg_match = re.search(r"(\d+(?:[.,]\d+)?)\s*KG\b", text, re.IGNORECASE)

    charge_match = re.search(
        r"(handling charges|handling fee|warehouse import charges|import warehouse charges).*?(\d+(?:[.,]\d+)?)",
        text,
        re.IGNORECASE | re.DOTALL,
    )

    if awb_match:
        result["AWB"] = awb_match.group(0)

    if kg_match:
        kg = float(kg_match.group(1).replace(",", "."))
        result["KG"] = kg

    if charge_match:
        charges = float(charge_match.group(2).replace(",", "."))
        result["Charges"] = charges

    if result["KG"] and result["Charges"]:
        result["Prijs_per_KG"] = round(result["Charges"] / result["KG"], 4)

    result["Status"] = "OK"
    return result


@app.route("/fetch-mails")
def fetch_mails():
    if "access_token" not in session:
        return redirect("/login")

    folder_id = find_folder_by_path(TARGET_FOLDER_PATH)

    if not folder_id:
        return "Map 'Postvak IN > Submap > facturen verwerkt' niet gevonden"

    messages_url = (
        f"{GRAPH_API}/me/mailFolders/{folder_id}/messages"
        "?$top=50"
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

        if sender != TARGET_SENDER:
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
                    "Datum": mail.get("receivedDateTime", ""),
                    "Afzender": sender,
                    "Onderwerp": mail.get("subject", ""),
                    "Bestandsnaam": filename,
                    "AWB": parsed.get("AWB", ""),
                    "KG": parsed.get("KG", ""),
                    "Charges": parsed.get("Charges", ""),
                    "Prijs_per_KG": parsed.get("Prijs_per_KG", ""),
                    "Status": parsed.get("Status", ""),
                }
            )

    if not rows:
        return "Geen PDF facturen gevonden van s.gasior@missionfreight.nl in 'Postvak IN > Submap > facturen verwerkt'"

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
