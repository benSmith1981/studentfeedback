from flask import Flask, request, jsonify
from docxtpl import DocxTemplate
from io import BytesIO
import base64
import os

app = Flask(__name__)

@app.route("/generateword", methods=["POST"])
def generate_word():
    try:
        body = request.get_json()
        data = body.get("data", {})

        template_path = "BTEC-Assignment-Brief-Template_jinja.docx"
        doc = DocxTemplate(template_path)
        doc.render(data)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        encoded = base64.b64encode(buffer.read()).decode("utf-8")
        return jsonify({"file": encoded})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
