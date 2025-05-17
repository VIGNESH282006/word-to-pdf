import os
import subprocess
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {'docx'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_to_pdf():
    if 'file' not in request.files:
        return render_template('index.html', message="No file part")

    file = request.files['file']
    if file.filename == '':
        return render_template('index.html', message="No file selected")

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        output_pdf = os.path.join(OUTPUT_FOLDER, filename.replace('.docx', '.pdf'))

        file.save(input_path)

        try:
            # Use pandoc to convert the DOCX file to PDF
            subprocess.run(['pandoc', input_path, '-o', output_pdf], check=True)
            return send_file(output_pdf, as_attachment=True)
        except subprocess.CalledProcessError as e:
            return render_template('index.html', message="Conversion failed. Ensure pandoc is installed.")
    else:
        return render_template('index.html', message="Only .docx files are supported")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
