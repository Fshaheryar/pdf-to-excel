from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os
import pdfplumber
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "supersecretkey"
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# PDF to Excel conversion logic (same as your original function)
def pdf_to_excel(pdf_path, excel_path):
    all_tables = []
    text_fallback = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    if table:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        df['Source_Page'] = i + 1
                        all_tables.append(df)
            else:
                text = page.extract_text()
                if text:
                    text_fallback.append({'Page': i + 1, 'Text': text.strip()})

    if all_tables:
        final_df = pd.concat(all_tables, ignore_index=True)
        final_df.to_excel(excel_path, index=False)
        return "tables"
    elif text_fallback:
        fallback_df = pd.DataFrame(text_fallback)
        fallback_df.to_excel(excel_path, index=False)
        return "text"
    else:
        return "empty"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if 'pdf_file' not in request.files:
            flash("No file part")
            return redirect(request.url)
        
        file = request.files['pdf_file']
        if file.filename == '':
            flash("No file selected")
            return redirect(request.url)

        if file:
            filename = secure_filename(file.filename)
            pdf_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(pdf_path)

            base_name = os.path.splitext(filename)[0]
            excel_path = os.path.join(UPLOAD_FOLDER, f"{base_name}.xlsx")

            result = pdf_to_excel(pdf_path, excel_path)

            if result == "empty":
                flash("No content could be extracted from the PDF.")
                return redirect(request.url)
            else:
                return send_file(excel_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
