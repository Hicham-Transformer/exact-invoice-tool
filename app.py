import io
import re
import pandas as pd
from flask import Flask, request, send_file

import pdfplumber

app = Flask(__name__)

BATCH_SIZE = 20


# =====================
# HELPERS
# =====================

def safe_float(val):
    try:
        return float(val)
    except:
        return ""


def parse_decimal(value):
    if not value:
        return ""

    s = str(value).replace("EUR", "").replace("kg", "").replace("kgs", "").strip()

    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        if len(parts) != 2:
            s = s.replace(".", "")

    return safe_float(s)


def extract_text(pdf_bytes):
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text += (page.extract_text() or "") + "\n"
    except Exception as e:
        return "", str(e)

    return text, None


# =====================
# PARSING
# =====================

def find_kg(text):
    m = re.search(r"([0-9\.,]+)\s*kgs?", text, re.IGNORECASE)
    if m:
        return parse_decimal(m.group(1))
    return ""


def find_charge(text):
    patterns = [
        r"handling fee.*?EUR\s*([0-9\.,]+)",
        r"handling charges.*?EUR\s*([0-9\.,]+)",
        r"import warehouse charges.*?EUR\s*([0-9\.,]+)",
    ]

    for p in patterns:
        m = re.search(p, text, re.IGNORECASE | re.DOTALL)
        if m:
            return parse_decimal(m.group(1))

    return ""


def find_awb(text):
    m = re.search(r"\d{3}-\d{8}", text)
    return m.group(0) if m else ""


def find_invoice(text):
    m = re.search(r"FACTUURNUMMER\s+([0-9]+)", text, re.IGNORECASE)
    return m.group(1) if m else ""


def parse_pdf(pdf_bytes):
    text, error = extract_text(pdf_bytes)

    if error:
        return {"Status": f"PDF fout: {error}"}

    kg = find_kg(text)
    charge = find_charge(text)

    prijs = ""
    if kg and charge:
        prijs = round(charge / kg, 5)

    return {
        "Factuurnummer": find_invoice(text),
        "AWB": find_awb(text),
        "KG": kg,
        "Charges": charge,
        "Prijs_per_KG": prijs,
        "Status": "OK"
    }


# =====================
# ROUTES
# =====================

@app.route("/")
def home():
    return """
    <h2>Upload PDF facturen</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
        <input type="file" name="files" multiple>
        <br><br>
        <button type="submit">Upload & Verwerken</button>
    </form>
    """


@app.route("/upload", methods=["POST"])
def upload():
    try:
        files = request.files.getlist("files")

        if not files:
            return "Geen bestanden"

        all_rows = []

        for i in range(0, len(files), BATCH_SIZE):
            batch = files[i:i+BATCH_SIZE]

            for file in batch:
                try:
                    pdf_bytes = file.read()
                    data = parse_pdf(pdf_bytes)

                    data["Bestandsnaam"] = file.filename
                    data["Batch"] = (i // BATCH_SIZE) + 1

                    all_rows.append(data)

                except Exception as e:
                    all_rows.append({
                        "Bestandsnaam": file.filename,
                        "Status": f"Crash: {str(e)}"
                    })

        df = pd.DataFrame(all_rows)

        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="facturen.xlsx",
            as_attachment=True
        )

    except Exception as e:
        return f"SERVER ERROR: {str(e)}", 500


@app.route("/health")
def health():
    return "ok"
