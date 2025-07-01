from flask import Flask, request, render_template, send_file, redirect, url_for
import os
import pandas as pd
import pdfplumber
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

last_output_path = ""

@app.route('/', methods=['GET'])
def index():
    return render_template("index.html")

@app.route('/upload', methods=['POST'])
def upload_files():
    global last_output_path
    uploaded_files = request.files.getlist("pdf_file")
    ledger_file = request.files.get("ledger_file")
    all_data = []

    # Load ledger keywords from uploaded Excel with error handling
    ledger_keywords = {}
    if ledger_file and ledger_file.filename.endswith(('.xlsx', '.xls')):
        try:
            ledger_df = pd.read_excel(ledger_file)
            for name in ledger_df['Ledger Name']:
                ledger_keywords[name.upper()] = name
        except Exception as e:
            print("❌ Ledger file error:", e)
            ledger_keywords = {}

    for file in uploaded_files:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                lines = page.extract_text().split('\n')
                i = 0
                while i < len(lines):
                    line = lines[i].strip()
                    if not line or 'B/F' in line.upper():
                        i += 1
                        continue

                    combined_line = line
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if not next_line[:1].isdigit():
                            combined_line += ' ' + next_line
                            i += 1

                    parts = combined_line.split()
                    if len(parts) < 6:
                        i += 1
                        continue

                    # Extract last 3 numbers from line as withdraw, deposit, balance
                    try:
                        numbers = [float(p.replace(',', '')) for p in parts if p.replace(',', '').replace('.', '').isdigit()]
                        if len(numbers) >= 3:
                            withdraw = numbers[-3]
                            deposit = numbers[-2]
                            balance = numbers[-1]
                        else:
                            i += 1
                            continue
                    except:
                        i += 1
                        continue

                    if withdraw == 0 and deposit == 0:
                        i += 1
                        continue

                    tokens = combined_line.split()
                    narration_tokens = [t for t in tokens if not t.replace(',', '').replace('.', '').isdigit()]
                    narration = ' '.join(narration_tokens[3:]).upper()
                    if len(narration.strip()) < 3:
                        narration = "NO NARRATION"

                    date = parts[0]

                    matched_ledger = "SUSPENSE"
                    for keyword in ledger_keywords:
                        if any(k in narration for k in keyword.upper().split()):
                            matched_ledger = ledger_keywords[keyword]
                            break

                    if deposit > 0:
                        amount = deposit
                        voucher_type = "Receipt"
                        by_dr_text = ""
                        to_cr_text = matched_ledger
                    elif withdraw > 0:
                        amount = withdraw
                        voucher_type = "Payment"
                        by_dr_text = "BANK CHARGES" if amount < 50 else matched_ledger
                        to_cr_text = ""
                    else:
                        i += 1
                        continue

                    all_data.append({
                        "DATE": date,
                        "VOUCHER NO.": "",
                        "BY / DR": by_dr_text,
                        "TO / CR": to_cr_text,
                        "AMOUNT": amount,
                        "NARRATION": narration.title(),
                        "VOUCHER TYPE": voucher_type,
                        "DAY": ""
                    })

                    i += 1

    if all_data:
        df = pd.DataFrame(all_data)
        output_excel = os.path.join(OUTPUT_FOLDER, "converted_output.xlsx")
        df.to_excel(output_excel, index=False)
        last_output_path = output_excel
        table_html = df.to_html(index=False)
    else:
        table_html = "❌ No valid data found. Try another PDF or format."

    return render_template("index.html", table=table_html)

@app.route('/download')
def download_file():
    if last_output_path and os.path.exists(last_output_path):
        return send_file(last_output_path, as_attachment=True)
    return redirect(url_for('index'))

if __name__ == "__main__":
    import os
    print("✅ Flask server started on Render")
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))









