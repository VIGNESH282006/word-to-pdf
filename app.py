import os
from flask import Flask, request, send_file, render_template
from docx2pdf import convert
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
CONVERTED_FOLDER = "converted"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        if file and file.filename.endswith(".docx"):
            filename = secure_filename(file.filename)
            docx_path = os.path.join(UPLOAD_FOLDER, filename)
            pdf_filename = filename.replace(".docx", ".pdf")
            pdf_path = os.path.join(CONVERTED_FOLDER, pdf_filename)

            file.save(docx_path)
            convert(docx_path, pdf_path)

            return send_file(pdf_path, as_attachment=True)

    return render_template("upload.html")

if __name__ == "__main__":
    app.run(debug=True)
