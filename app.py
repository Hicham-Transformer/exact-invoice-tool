import os
import requests
import pdfplumber
import pandas as pd
from flask import Flask, redirect, request, session, url_for, send_file

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev")

# ENV VARS
CLIENT_ID = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
TENANT_ID = os.getenv("MS_TENANT_ID", "common")
REDIRECT_URI = os.getenv("MS_REDIRECT_URI")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
TOKEN_URL = f"{AUTHORITY}/oauth2/v2.0/token"
AUTH_URL = f"{AUTHORITY}/oauth2/v2.0/authorize"

SCOPES = [
    "User.Read",
    "Mail.Read",
    "offline_access"
]

GRAPH_API = "https://graph.microsoft.com/v1.0"

# ---------------- LOGIN ----------------

@app.route("/")
def index():
    return '<a href="/login">Login met Outlook</a>'


@app.route("/login")
def login():
    params = {
        "client_id": CLIENT_ID,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "response_mode": "query",
        "scope": " ".join(SCOPES)
    }
    url = AUTH_URL + "?" + "&".join([f"{k}={v}" for k, v in params.items()])
    return redirect(url)


@app.route("/callback")
def callback():
    code = request.args.get("code")

    data = {
        "client_id": CLIENT_ID,
        "scope": " ".join(SCOPES),
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "client_secret": CLIENT_SECRET,
    }

    token = requests.post(TOKEN_URL, data=data).json()

    session["access_token"] = token.get("access_token")

    return "Login succesvol ✅ Ga naar /fetch-mails"


# ---------------- GRAPH HELPER ----------------

def graph_get(url):
    headers = {
        "Authorization": f"Bearer {session.get('access_token')}"
    }
    return requests.get(url, headers=headers).json()


# ---------------- FOLDER SEARCH ----------------

def find_folder_recursive(target_name):
    target_name = target_name.lower()

    def search(folder_id=None):
        if folder_id:
            url = f"{GRAPH_API}/me/mailFolders/{folder_id}/childFolders?$top=200"
        else:
            url = f"{GRAPH_API}/me/mailFolders?$top=200"

        data = graph_get(url)

        for folder in data.get("value", []):
            name = folder.get("displayName", "").lower()

            if name == target_name:
                return folder["id"]

            found = search(folder["id"])
            if found:
                return found

        return None

    return search()


# ---------------- PDF PARSER ----------------

def extract_data_from_pdf(file_bytes):
    results = []

    with pdfplumber.open(file_bytes) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() or ""

    # simpele parsing (kan later slimmer)
    awb = ""
    kg = ""
    handling = ""

    for line in text.split("\n"):
        if "AWB" in line:
            awb = line
        if "kg" in line.lower():
            kg = line
        if "handling" in line.lower():
            handling = line

    results.append({
        "AWB": awb,
        "KG": kg,
        "Handling": handling
    })

    return results


# ---------------- FETCH MAILS ----------------

@app.route("/fetch-mails")
def fetch_mails():
    folder_id = find_folder_recursive("facturen verwerkt")

    if not folder_id:
        return "Map 'facturen verwerkt' niet gevonden ❌"

    url = f"{GRAPH_API}/me/mailFolders/{folder_id}/messages?$top=25"
    messages = graph_get(url)

    all_data = []

    for msg in messages.get("value", []):
        msg_id = msg["id"]

        attachments_url = f"{GRAPH_API}/me/messages/{msg_id}/attachments"
        attachments = graph_get(attachments_url)

        for att in attachments.get("value", []):
            if att["@odata.type"] == "#microsoft.graph.fileAttachment":
                if att["name"].lower().endswith(".pdf"):

                    content_bytes = att["contentBytes"]

                    import base64
                    file_bytes = base64.b64decode(content_bytes)

                    import io
                    pdf_file = io.BytesIO(file_bytes)

                    data = extract_data_from_pdf(pdf_file)
                    all_data.extend(data)

    if not all_data:
        return "Geen PDF data gevonden"

    df = pd.DataFrame(all_data)

    output_file = "result.xlsx"
    df.to_excel(output_file, index=False)

    return send_file(output_file, as_attachment=True)


# ---------------- RUN ----------------

if __name__ == "__main__":
    app.run(debug=True)
