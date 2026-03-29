from flask import Flask, redirect, request, session, send_file, url_for
import requests
import os
import io
import pandas as pd
import re
from datetime import datetime

try:
    import fitz  # pymupdf
except:
    fitz = None

app = Flask(__name__)
app.secret_key = "supersecret"

# ================= MICROSOFT CONFIG =================
CLIENT_ID = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")
REDIRECT_URI = os.getenv("MS_REDIRECT_URI")

AUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
GRAPH_URL = "https://graph.microsoft.com/v1.0"


# ================= LOGIN =================
@app.route("/login")
def login():
    return redirect(
        f"{AUTH_URL}?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={REDIRECT_URI}"
        f"&scope=offline_access Mail.Read"
    )


@app.route("/callback")
def callback():
    code = request.args.get("code")

    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
    }

    r = requests.post(TOKEN_URL, data=data)
    token = r.json()

    session["access_token"] = token["access_token"]

    return redirect("/fetch")


# ================= PDF PARSER =================
def extract_pdf_data(pdf_bytes):
    if not fitz:
        return {}

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = ""

    for page in doc:
        text += page.get_text()

    factuur = re.search(r"FACTUURNUMMER\s*(\d+)", text)
    awb = re.search(r"(\d{3}-\d{8})", text)
    kg = re.search(r"bruto.*?(\d+[.,]\d+)", text.lower())
    charge = re.search(r"warehouse.*?(\d+[.,]\d+)", text.lower())

    return {
        "Factuurnummer": factuur.group(1) if factuur else "",
        "AWB": awb.group(1) if awb else "",
        "KG": float(kg.group(1).replace(",", ".")) if kg else "",
        "Charges": float(charge.group(1).replace(",", ".")) if charge else "",
    }


# ================= FETCH OUTLOOK =================
@app.route("/fetch")
def fetch():
    token = session.get("access_token")

    if not token:
        return redirect("/login")

    headers = {"Authorization": f"Bearer {token}"}

    # 🔍 Zoek map
    folders = requests.get(f"{GRAPH_URL}/me/mailFolders", headers=headers).json()

    folder_id = None
    for f in folders.get("value", []):
        if f["displayName"].lower() == "facturen verwerkt":
            folder_id = f["id"]

    if not folder_id:
        return "Map 'facturen verwerkt' niet gevonden"

    # 📥 Haal mails
    messages = requests.get(
        f"{GRAPH_URL}/me/mailFolders/{folder_id}/messages?$top=50",
        headers=headers,
    ).json()

    rows = []

    for msg in messages.get("value", []):
        sender = msg["from"]["emailAddress"]["address"]

        if sender.lower() != "s.gasior@missionfreight.nl":
            continue

        msg_id = msg["id"]

        attachments = requests.get(
            f"{GRAPH_URL}/me/messages/{msg_id}/attachments",
            headers=headers,
        ).json()

        for att in attachments.get("value", []):
            if att.get("contentType") == "application/pdf":
                content = att.get("contentBytes")

                import base64
                pdf_bytes = base64.b64decode(content)

                data = extract_pdf_data(pdf_bytes)

                rows.append({
                    "Datum": msg["receivedDateTime"],
                    "Afzender": sender,
                    "Bestandsnaam": att["name"],
                    **data
                })

    if not rows:
        return "Geen PDF facturen gevonden"

    df = pd.DataFrame(rows)

    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="outlook_facturen.xlsx",
        as_attachment=True
    )


# ================= HOME =================
@app.route("/")
def home():
    return """
    <h2>Outlook Factuur Tool</h2>
    <a href="/login">Login met Outlook</a>
    """
    

if __name__ == "__main__":
    app.run()
