from flask import Flask, request, render_template, send_file
from docxtpl import DocxTemplate
import pandas as pd
import io, zipfile, os

app = Flask(__name__)

@app.route('/download-template')
def download_template():
    # Define column headers and one sample row for guidance
    columns = [
        "Name", "assignmentTitle", "assessorName", "unitNumber",
        "programmeTitle", "dueDate", "handInDate", "overallComment"
    ]

    data = [
        {
            "Name": "Example Student",
            "assignmentTitle": "Unit 1 – Principles of Computing",
            "assessorName": "Your Name",
            "unitNumber": "1",
            "programmeTitle": "BTEC National in Computing",
            "dueDate": "2025-10-23",
            "handInDate": "2025-10-22",
            "overallComment": "Write feedback here..."
        }
    ]

    df = pd.DataFrame(data, columns=columns)
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, sheet_name="FeedbackData")
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Feedback_Template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
@app.route('/')
def home():
    return """
    <h2>BTEC Document Generator</h2>
    <p>Select a tool:</p>
    <ul>
      <li><a href='/assignmentbrief'>Generate Assignment Brief</a></li>
      <li><a href='/feedback'>Generate Feedback/Assessment Records</a></li>
    </ul>
    """

# 1️⃣ Assignment Brief Generator
@app.route('/assignmentbrief', methods=['GET', 'POST'])
def assignment_brief():
    if request.method == 'POST':
        # Collect data from form fields
        data = request.form.to_dict(flat=True)

        # Load and render the Word template
        doc = DocxTemplate("BTEC-Assignment-Brief-Template_jinja.docx")
        doc.render(data)

        # Return file as download
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name="Assignment_Brief.docx")

    return render_template("form_assignment.html")


# 2️⃣ Feedback Generator (reads Excel/CSV)
@app.route('/feedback', methods=['GET', 'POST'])
def feedback():
    if request.method == 'POST':
        file = request.files['file']
        if not file or file.filename == '':
            return "⚠️ No file uploaded.", 400

        # Determine file type
        filename = file.filename.lower()
        if filename.endswith(".xlsx"):
            df = pd.read_excel(file)
        elif filename.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            return "❌ Please upload a .csv or .xlsx file.", 400

        # Clean data
        df = df.dropna(how="all")  # drop completely empty rows
        if df.empty:
            return "⚠️ Your file appears empty or has no valid rows.", 400

        # Debug: print out first row to server log
        print("First row data:", df.iloc[0].to_dict())

        # Prepare ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for index, row in df.iterrows():
                context = {k.strip(): ("" if pd.isna(v) else str(v)) for k, v in row.items()}

                # Skip if Name is missing
                if not context.get("Name"):
                    print(f"Skipping blank Name at row {index}")
                    continue

                try:
                    doc = DocxTemplate("BTEC-Assessment-Record-TemplateJinja.docx")
                    doc.render(context)
                    temp = io.BytesIO()
                    doc.save(temp)
                    temp.seek(0)
                    zipf.writestr(f"{context['Name']}_Feedback.docx", temp.read())
                    print(f"✅ Created doc for {context['Name']}")
                except Exception as e:
                    print(f"❌ Error rendering {context.get('Name')}: {e}")

        zip_buffer.seek(0)

        if not zip_buffer.getbuffer().nbytes:
            return "⚠️ No documents were generated. Check template placeholders and Excel headers."

        return send_file(zip_buffer, as_attachment=True, download_name="Feedback_Records.zip")

    return render_template("form_feedback.html")



if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
