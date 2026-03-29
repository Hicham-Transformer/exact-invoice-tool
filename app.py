from flask import Flask, redirect, request, session, url_for
import requests
import os

app = Flask(__name__)
app.secret_key = "supersecret"

CLIENT_ID = os.getenv("EXACT_CLIENT_ID")
CLIENT_SECRET = os.getenv("EXACT_CLIENT_SECRET")
REDIRECT_URI = os.getenv("EXACT_REDIRECT_URI")

AUTH_URL = "https://start.exactonline.nl/api/oauth2/auth"
TOKEN_URL = "https://start.exactonline.nl/api/oauth2/token"
BASE_URL = "https://start.exactonline.nl/api/v1"


# ================= LOGIN =================

@app.route("/login")
def login():
    return redirect(
        f"{AUTH_URL}?client_id={CLIENT_ID}&redirect_uri={REDIRECT_URI}&response_type=code&scope=exactonlineapi%20offline_access"
    )


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

    r = requests.post(TOKEN_URL, data=data)
    token = r.json()

    session["access_token"] = token["access_token"]

    return redirect("/")


# ================= DIVISION =================

def get_division(token):
    r = requests.get(
        f"{BASE_URL}/current/Me",
        headers={"Authorization": f"Bearer {token}"}
    )

    data = r.json()
    return data["d"]["results"][0]["CurrentDivision"]


# ================= EXACT DATA =================

@app.route("/fetch")
def fetch():
    token = session.get("access_token")

    if not token:
        return redirect("/login")

    division = get_division(token)

    url = f"{BASE_URL}/{division}/purchaseentry/PurchaseEntries?$top=50"

    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})

    if r.status_code != 200:
        return f"Fout: {r.text}"

    data = r.json()
    results = data["d"]["results"]

    output = "<h2>Exact data (eerste 50)</h2><br>"

    for item in results:
        output += f"""
        Factuur: {item.get('InvoiceNumber')}<br>
        Leverancier: {item.get('SupplierName')}<br>
        Bedrag: {item.get('AmountDC')}<br>
        Omschrijving: {item.get('Description')}<br>
        <hr>
        """

    return output


# ================= HOME =================

@app.route("/")
def index():
    if "access_token" in session:
        return """
        <h2>Exact gekoppeld</h2>
        <a href='/fetch'>👉 Haal facturen op</a>
        """
    else:
        return "<a href='/login'>Login met Exact</a>"


if __name__ == "__main__":
    app.run()
