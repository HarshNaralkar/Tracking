from flask import Flask, render_template, request, send_from_directory, jsonify
import os
import pandas as pd
from docx import Document
import uuid
import requests

app = Flask(__name__, template_folder="templates", static_folder="static")

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "C:/Users/%USERPROFILE%/Downloads/"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process_files():
    word_file = request.files.get("wordFile")
    excel_file = request.files.get("excelFile")
    google_sheet_url = request.form.get("sheetUrl")
    sheet_name = request.form.get("sheetName")

    if not word_file:
        return jsonify({"success": False, "message": "Word file is required!"})

    word_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}_template.docx")
    word_file.save(word_path)

    try:
        if excel_file:
            excel_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}_data.xlsx")
            excel_file.save(excel_path)
            df = pd.read_excel(excel_path)
        elif google_sheet_url and sheet_name:
            sheet_id = google_sheet_url.split("/d/")[1].split("/")[0]
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
            response = requests.get(csv_url)
            response.raise_for_status()
            df = pd.read_csv(pd.compat.StringIO(response.text))
        else:
            return jsonify({"success": False, "message": "Please upload Excel or provide Google Sheet URL and Sheet Name."})

        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d/%m/%Y')

        generated_files = []

        for _, row in df.iterrows():
            doc = Document(word_path)
            srno = str(row.get("Srno", uuid.uuid4()))
            for para in doc.paragraphs:
                for key in row.index:
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in para.text:
                        para.text = para.text.replace(placeholder, str(row[key]))

            output_path = os.path.join(OUTPUT_FOLDER, f"{srno}.docx")
            doc.save(output_path)

            generated_files.append({
                "name": f"{srno}.docx",
                "url": f"/download/{srno}.docx"
            })

        return jsonify({"success": True, "files": generated_files})

    except Exception as e:
        return jsonify({"success": False, "message": f"Error: {str(e)}"})

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
