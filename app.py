from flask import Flask, request, render_template_string, send_file
import pandas as pd
from docxtpl import DocxTemplate
import zipfile
import io
import os
from datetime import datetime

app = Flask(__name__)

# Path to your Jinja Word template
TEMPLATE_PATH = "BTEC-Assignment-Brief-Template_jinja.docx"

# --- Frontend HTML ---
UPLOAD_PAGE = """
<!DOCTYPE html>
<html>
<head>
  <title>Generate Feedback</title>
  <style>
    body { font-family: Arial; margin: 50px; background: #f9f9f9; color: #333; }
    form { background: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 6px rgba(0,0,0,0.1); }
    input[type=file] { margin-bottom: 20px; }
    button { background: #0078d4; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; }
    button:hover { background: #005fa3; }
  </style>
</head>
<body>
  <h2>Upload Excel File to Generate Feedback Docs</h2>
  <form action="/generate" method="post" enctype="multipart/form-data">
    <input type="file" name="excel_file" accept=".xlsx,.xls" required><br>
    <button type="submit">Generate ZIP</button>
  </form>
</body>
</html>
"""

@app.route("/")
def index():
    return render_template_string(UPLOAD_PAGE)


@app.route("/generate", methods=["POST"])
def generate_docs():
    file = request.files.get("excel_file")
    if not file:
        return "No file uploaded", 400

    df = pd.read_excel(file)

    # Expect columns: studentName, programmeTitle
    required_cols = ["studentName", "programmeTitle"]
    for col in required_cols:
        if col not in df.columns:
            return f"Missing column: {col}", 400

    memory_zip = io.BytesIO()
    with zipfile.ZipFile(memory_zip, "w") as zipf:
        for _, row in df.iterrows():
            name = str(row["studentName"]).strip()
            if not name:
                continue  # skip empty rows

            context = {
                "studentName": name,
                "programmeTitle": str(row["programmeTitle"]).strip()
            }

            doc = DocxTemplate(TEMPLATE_PATH)
            doc.render(context)

            # Save each file to memory zip
            filename = f"{name.replace(' ', '_')}_Feedback.docx"
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            zipf.writestr(filename, buffer.read())

    memory_zip.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return send_file(
        memory_zip,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"Generated_Feedback_{timestamp}.zip"
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))  # Render provides PORT automatically
    app.run(host='0.0.0.0', port=port)
