from flask import Flask, render_template, request, send_file, session
import pdfplumber
import pandas as pd
import os
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom.minidom import parseString
from datetime import datetime
import os
app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
app.run(debug=True)

app = Flask(__name__)
app.secret_key = 'your_secret_key'

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

bank_ledger_names = {
    "canara": "Canara Bank",
    "sbi": "SBI Bank",
    "hdfc": "HDFC Bank",
    "uco": "UCO Bank",
    "axis": "Axis Bank",
    "union": "Union Bank",
    "central": "Central Bank of India",
    "federal": "Federal Bank",
    "icici": "ICICI Bank"
}

def convert_to_tally_xml(df, bank_ledger):
    root = Element("ENVELOPE")
    header = SubElement(root, "HEADER")
    SubElement(header, "TALLYREQUEST").text = "Import Data"

    body = SubElement(root, "BODY")
    import_data = SubElement(body, "IMPORTDATA")

    request_desc = SubElement(import_data, "REQUESTDESC")
    SubElement(request_desc, "REPORTNAME").text = "Vouchers"
    static_vars = SubElement(request_desc, "STATICVARIABLES")
    SubElement(static_vars, "SVCURRENTCOMPANY").text = "My Company"

    request_data = SubElement(import_data, "REQUESTDATA")

    for _, row in df.iterrows():
        try:
            date_str = datetime.strptime(row["Date"], "%d/%m/%Y").strftime("%Y%m%d")
            amt = row["Credit"] or row["Debit"]
            amt = amt.replace(",", "").strip()

            if amt == '' or float(amt) == 0.0:
                continue

            tally_msg = SubElement(request_data, "TALLYMESSAGE")
            voucher = SubElement(tally_msg, "VOUCHER", {"VCHTYPE": "Receipt", "ACTION": "Create"})

            SubElement(voucher, "DATE").text = date_str
            SubElement(voucher, "NARRATION").text = row["Narration"]
            SubElement(voucher, "VOUCHERTYPENAME").text = "Receipt"
            SubElement(voucher, "PARTYLEDGERNAME").text = "Suspense"

            entry1 = SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            SubElement(entry1, "LEDGERNAME").text = "Suspense"
            SubElement(entry1, "ISDEEMEDPOSITIVE").text = "Yes"
            SubElement(entry1, "AMOUNT").text = f"-{amt}"

            entry2 = SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            SubElement(entry2, "LEDGERNAME").text = bank_ledger
            SubElement(entry2, "ISDEEMEDPOSITIVE").text = "No"
            SubElement(entry2, "AMOUNT").text = amt

        except Exception as e:
            continue

    xml_data = parseString(tostring(root)).toprettyxml()
    with open("output/statement.xml", "w", encoding="utf-8") as f:
        f.write(xml_data)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/preview', methods=['POST'])
def preview():
    if 'pdf' not in request.files:
        return "No file part"
    file = request.files['pdf']
    if file.filename == '':
        return "No selected file"

    bank = request.form.get('bank')
    bank_ledger = bank_ledger_names.get(bank, "Bank Account")

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)

    data = []
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if not table:
                    continue
                for row in table[1:]:
                    if not row:
                        continue
                    row = [cell.strip() if cell else '' for cell in row]
                    try:
                        if bank == 'canara' and len(row) >= 8:
                            data.append([row[1], row[4], row[5], row[6], row[7]])
                        elif bank == 'sbi' and len(row) >= 7:
                            data.append([row[0], row[2], row[3], row[4], row[5]])
                        elif bank == 'hdfc' and len(row) >= 6:
                            data.append([row[0], row[2], row[3], row[4], row[5]])
                        elif bank == 'uco' and len(row) >= 6:
                            data.append([row[0], row[2], row[3], row[4], row[5]])
                        elif bank == 'federal' and len(row) >= 6:
                            data.append([row[0], row[1], row[2], row[3], row[4]])
                        elif bank == 'axis' and len(row) >= 6:
                            data.append([row[0], row[1], row[2], row[3], row[4]])
                        elif bank == 'union' and len(row) >= 7:
                            data.append([row[0], row[2], row[3], row[4], row[5]])
                        elif bank == 'central' and len(row) >= 7:
                            data.append([row[0], row[2], row[3], row[4], row[5]])
                        elif bank == 'icici' and len(row) >= 6:
                            data.append([row[0], row[1], row[2], row[3], row[4]])
                    except:
                        continue

        if not data:
            return "❌ No valid data extracted."

        df = pd.DataFrame(data, columns=["Date", "Narration", "Debit", "Credit", "Balance"])
        df.to_excel("output/statement.xlsx", index=False)

        convert_to_tally_xml(df, bank_ledger)

        session['data'] = data
        return render_template('preview.html', data=data)

    except Exception as e:
        return f"❌ Error: {e}"

@app.route('/download/excel')
def download_excel():
    return send_file("output/statement.xlsx", as_attachment=True)

@app.route('/download/xml')
def download_xml():
    from flask import Flask, render_template, request, send_file
import os

app = Flask(__name__)

# --- Your routes and logic go here ---
@app.route('/')
def home():
    return render_template("index.html")

# Add your other routes here...

# --- This part must be at the END ---
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))

    return send_file("output/statement.xml", as_attachment=True)

if __name__ == '__main__':
    print("✅ Flask server started on http://127.0.0.1:5000")
    app.run(debug=True)
