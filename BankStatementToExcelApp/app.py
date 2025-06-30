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
    all_data = []

    for file in uploaded_files:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)

        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                lines = text.split('\n')

                for line in lines:
                    if 'B/F' in line.upper() or line.strip() == "":
                        continue

                    parts = line.split()
                    if len(parts) < 6:
                        continue

                    try:
                        withdraw_str = parts[-2].replace(",", "").strip()
                        withdraw = float(withdraw_str)
                        deposit = 0.0
                    except ValueError:
                        continue

                    if withdraw == 0:
                        continue

                    narration_text = ' '.join(parts[1:-3]).upper()
                    date = parts[0]

                    is_card_sale = 'UPI' in narration_text or 'CARD' in narration_text

                    if is_card_sale:
                        voucher_type = "Receipt"
                        amount = withdraw
                        by_dr_text = ""
                        to_cr_text = "CARD SALES"
                    else:
                        voucher_type = "Payment"
                        amount = withdraw
                        by_dr_text = "BANK CHARGES" if amount < 50 else "SUSPENSE"
                        to_cr_text = ""

                    all_data.append({
                        "DATE": date,
                        "VOUCHER NO.": "",
                        "BY / DR": by_dr_text,
                        "TO / CR": to_cr_text,
                        "AMOUNT": amount,
                        "NARRATION": narration_text,
                        "VOUCHER TYPE": voucher_type,
                        "DAY": ""
                    })

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

