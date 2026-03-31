import os
import requests
import tempfile
import re
from flask import Flask, redirect, request, session, send_file
from urllib.parse import urlencode
import pdfplumber
import pandas as pd

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "supersecret")

# ENV VARS (Render)
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
    "offline_access"
]


# ---------------- LOGIN ----------------

@app.route("/login")
def login():
    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "response_mode": "query",
        "scope": " ".join(SCOPES)
    }
    return redirect(f"{AUTH_URL}?{urlencode(params)}")


@app.route("/callback")
def callback():
    code = request.args.get("code")

    data = {
        "client_id": CLIENT_ID,
        "scope": " ".join(SCOPES),
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "client_secret": CLIENT_SECRET
    }

    token_res = requests.post(TOKEN_URL, data=data).json()

    if "access_token" not in token_res:
        return f"Token error: {token_res}"

    session["access_token"] = token_res["access_token"]

    return "Login succesvol ✅ Ga naar /fetch-mails"


# ---------------- FETCH MAILS ----------------

def get_headers():
    return {
        "Authorization": f"Bearer {session.get('access_token')}"
    }


def extract_pdf_data(pdf_path):
    data = {
        "awb": None,
        "kg": None,
        "charges": None
    }

    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"

    # AWB (voorbeeld patroon)
    awb_match = re.search(r"\d{3}-\d{8}", text)
    if awb_match:
        data["awb"] = awb_match.group()

    # KG
    kg_match = re.search(r"(\d+(\.\d+)?)\s?KG", text, re.IGNORECASE)
    if kg_match:
        data["kg"] = float(kg_match.group(1))

    # Charges (handling etc.)
    charge_match = re.search(r"(\d+(\.\d+)?)\s?EUR", text)
    if charge_match:
        data["charges"] = float(charge_match.group(1))

    return data


@app.route("/fetch-mails")
def fetch_mails():
    headers = get_headers()

    # 📂 Folder ophalen (facturen verwerkt)
    folders = requests.get(f"{GRAPH_API}/me/mailFolders", headers=headers).json()

    folder_id = None
    for f in folders.get("value", []):
        if f["displayName"].lower() == "facturen verwerkt":
            folder_id = f["id"]

    if not folder_id:
        return "Map 'facturen verwerkt' niet gevonden"

    # 📧 Mails ophalen
    mails = requests.get(
        f"{GRAPH_API}/me/mailFolders/{folder_id}/messages?$top=20",
        headers=headers
    ).json()

    results = []

    for mail in mails.get("value", []):
        sender = mail.get("from", {}).get("emailAddress", {}).get("address", "")

        # 🔒 filter op afzender
        if "missionfreight.nl" not in sender:
            continue

        msg_id = mail["id"]

        attachments = requests.get(
            f"{GRAPH_API}/me/messages/{msg_id}/attachments",
            headers=headers
        ).json()

        for att in attachments.get("value", []):
            if att.get("@odata.type") == "#microsoft.graph.fileAttachment":
                if att["name"].lower().endswith(".pdf"):

                    file_data = att["contentBytes"]

                    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                    temp.write(requests.utils.unquote_to_bytes(file_data))
                    temp.close()

                    parsed = extract_pdf_data(temp.name)

                    if parsed["kg"] and parsed["charges"]:
                        price_per_kg = parsed["charges"] / parsed["kg"]
                    else:
                        price_per_kg = None

                    parsed["price_per_kg"] = price_per_kg

                    results.append(parsed)

    if not results:
        return "Geen data gevonden"

    # 📊 Excel maken
    df = pd.DataFrame(results)

    file_path = "output.xlsx"
    df.to_excel(file_path, index=False)

    return send_file(file_path, as_attachment=True)


# ---------------- ROOT ----------------

@app.route("/")
def home():
    return """
    <h2>Outlook PDF Tool</h2>
    <a href="/login">Login met Outlook</a><br><br>
    <a href="/fetch-mails">Haal facturen op</a>
    """


# ---------------- RUN ----------------

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
