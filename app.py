from pathlib import Path
from dotenv import load_dotenv
import os
import io
import zipfile
import pandas as pd

from flask import Flask, request, send_file, render_template
from flask import Response, stream_with_context
import json
import time
from docxtpl import DocxTemplate
from openai import OpenAI

# ===============================
# ENV SETUP
# ===============================
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)

print("OPENAI_API_KEY loaded:", bool(os.getenv("OPENAI_API_KEY")))

# ===============================
# APP
# ===============================
app = Flask(__name__)
client = OpenAI()  # GPT-5 mini

ASSESSMENT_TEMPLATE = "BTEC-Assessment-Record-TemplateJinja.docx"
ASSIGNMENT_TEMPLATE = "BTEC-Assignment-Brief-Template_jinja.docx"


# ===============================
# HOME
# ===============================
@app.route("/")
def home():
    return render_template("home.html")

@app.route("/extract-students", methods=["POST"])
def extract_students_route():
    data = request.get_json()
    names = extract_student_names(data.get("text", ""))
    return json.dumps(names)


@app.route("/assignmentbrief", methods=["GET", "POST"])
def assignmentbrief():
    if request.method == "POST":
        data = request.form.to_dict(flat=True)

        raw_learning_aims = data.get("learningAims", "").strip()

        # ---------- AI CLEAN LEARNING AIMS ----------
        if raw_learning_aims:
            prompt = f"""
You are a UK BTEC assessor.

Rewrite the learning aims below so that they are:
- Clear
- Concise
- Written in formal academic style
- Suitable for a BTEC Assignment Brief

Do NOT invent new aims.
Do NOT add grades.
Preserve meaning.

Learning aims:
{raw_learning_aims}
"""
            response = client.chat.completions.create(
                model="gpt-5-mini",
                messages=[{"role": "user", "content": prompt}]
            )

            data["learningAims"] = response.choices[0].message.content.strip()

        # ---------- RENDER WORD DOC ----------
        doc = DocxTemplate(ASSIGNMENT_TEMPLATE)
        doc.render(data)

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="Assignment_Brief.docx"
        )

    return render_template("form_assignment_with_storage.html")



# ===============================
# 2️⃣ MANUAL FEEDBACK (EXCEL)
# ===============================
@app.route("/feedback", methods=["GET", "POST"])
def feedback():
    if request.method == "POST":
        file = request.files.get("file")

        if not file or not file.filename.endswith(".xlsx"):
            return "Please upload an Excel file (.xlsx)", 400

        xls = pd.ExcelFile(file)
        learners_df = pd.read_excel(xls, "Learners")
        criteria_df = pd.read_excel(xls, "Criteria")

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for _, learner in learners_df.iterrows():
                name = learner.get("Name")
                if pd.isna(name):
                    continue

                context = learner.to_dict()
                context["criterias"] = criteria_df[
                    criteria_df["Name"] == name
                ].to_dict(orient="records")

                # auto student email
                parts = name.lower().replace(",", "").split()
                context["studentEmailSignature"] = (
                    f"{parts[0]}.{parts[-1]}@student.cityofbristol.ac.uk"
                )

                doc = DocxTemplate(ASSESSMENT_TEMPLATE)
                doc.render(context)

                temp = io.BytesIO()
                doc.save(temp)
                temp.seek(0)

                zipf.writestr(
                    f"{name.replace(' ', '_')}_Assessment_Record.docx",
                    temp.read()
                )

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name="Assessment_Records_Manual.zip"
        )

    return render_template("form_feedback.html")


def extract_student_names(raw_text: str) -> list[str]:
    # ---------- HARD GUARD ----------
    if not raw_text or not raw_text.strip():
        return []

    prompt = f"""
You are extracting student names for an assessment system.

IMPORTANT RULES:
- Return ONLY a JSON array
- Each item must be a full student name
- If NO names are found, return an EMPTY array: []
- Do NOT explain anything
- Do NOT include messages, apologies, or guidance

Examples:
Input: "John Smith, Merit"
Output: ["John Smith"]

Input: ""
Output: []

Text to extract from:
{raw_text}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-5-mini",
            messages=[{"role": "user", "content": prompt}],
            timeout=20
        )

        content = response.choices[0].message.content.strip()
        names = json.loads(content)

        # Final safety filter
        return [
            n.strip()
            for n in names
            if isinstance(n, str) and len(n.strip()) > 2
        ]

    except Exception as e:
        print("❌ Name extraction failed:", e)
        return []
    

@app.route("/feedback-ai", methods=["GET", "POST"])
def feedback_ai():
    if request.method == "POST":

        # ---------- READ STRUCTURED STUDENT DATA ----------
        try:
            students = json.loads(request.form.get("students_json", "[]"))
        except json.JSONDecodeError:
            return "❌ Invalid student data submitted.", 400

        if not students:
            return "❌ No students submitted.", 400

        if len(students) > 20:
            return "❌ Too many students at once (maximum 20).", 400

        criteria_text = request.form.get("criteria", "")
        teacher_notes = request.form.get("teacher_notes", "")

        print(f"🔍 Starting AI feedback generation for {len(students)} students")

        # ---------- SHARED CONTEXT ----------
        context_base = {
            "programmeTitle": request.form.get("programmeTitle"),
            "assignmentTitle": request.form.get("assignmentTitle"),
            "unitNumber": request.form.get("unitNumber"),
            "assessorName": request.form.get("assessorName"),
            "assessorEmailSignature": request.form.get("assessorEmail"),
            "due_date": request.form.get("due_date"),
            "HandInDate": request.form.get("handInDate"),
            "marked_date": request.form.get("marked_date"),
            "feedback_date": request.form.get("feedback_date"),
            "signature_iv": request.form.get("signature_iv"),
            "lead_iv_signed_date": request.form.get("lead_iv_signed_date"),
        }

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for s in students:
                student = s["name"]
                grade_hint = s.get("grade", "")
                teacher_note = s.get("note", "")

                print(f"🧠 Generating AI assessment for {student}")

                prompt = f"""
You are a UK BTEC assessor completing an official Assessment Record.

Student: {student}
Teacher grading indication: {grade_hint}
Teacher note: {teacher_note}

Assignment criteria:
{criteria_text}

General teacher notes:
{teacher_notes}

TASK:
Return ONLY valid JSON.

{{
  "criterias": [
    {{
      "title": "23 / P1",
      "targetedCriteria": "Full criterion wording",
      "criteriaAchieved": "Yes or No",
      "assessmentComment": "Explain clearly to the learner, address them as you or something direct to them, why this criterion was or was not achieved.
If not achieved, state what is missing and what must be done to meet it but only based on given criteria or extra input from the teacher."
    }}
  ],
  "overallComment": "A full professional paragraph summarising, addressed to the learner as you or something direct to them, overall performance,
strengths, unmet criteria, and clear next steps, unless they achieved all the criteria and got a distinction. Don't go overboard or state anything that you don't know from evidence in comment from teacher or from the criteria. 
State the grade PASS, MERIT or DISTINCTION as first thing to be clear to them."
}}

RULES:
- Use professional BTEC assessor tone
- Use ONLY Yes or No
- Do not mention AI
"""

                try:
                    response = client.chat.completions.create(
                        model="gpt-5-mini",
                        messages=[{"role": "user", "content": prompt}],
                        timeout=30
                    )
                    data = json.loads(response.choices[0].message.content.strip())

                except Exception as e:
                    print(f"❌ AI failed for {student}: {e}")
                    continue

                # ---------- NORMALISE CRITERIA ----------
                criterias = []
                for c in data.get("criterias", []):
                    criterias.append({
                        "title": c.get("title", ""),
                        "targetedCriteria": c.get("targetedCriteria", ""),
                        "criteriaAchieved": "Yes" if c.get("criteriaAchieved") == "Yes" else "No",
                        "assessmentComment": c.get("assessmentComment", ""),
                    })

                # ---------- AUTO STUDENT EMAIL ----------
                parts = student.lower().replace(",", "").split()
                student_email = f"{parts[0]}.{parts[-1]}@student.cityofbristol.ac.uk"

                context = {
                    **context_base,
                    "Name": student,
                    "studentEmailSignature": student_email,
                    "criterias": criterias,
                    "overallComment": data.get("overallComment", "")
                }

                doc = DocxTemplate(ASSESSMENT_TEMPLATE)
                doc.render(context)

                temp = io.BytesIO()
                doc.save(temp)
                temp.seek(0)

                zipf.writestr(
                    f"{student.replace(' ', '_')}_Assessment_Record_AI.docx",
                    temp.read()
                )

                print(f"📄 Document generated for {student}")

        zip_buffer.seek(0)
        print("✅ All AI assessment records generated")

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name="Assessment_Records_AI.zip"
        )

    return render_template("form_feedback_ai.html")


# ===============================
# RUN
# ===============================
if __name__ == "__main__":
    app.run(debug=True)
