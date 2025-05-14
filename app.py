import os
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
from docx import Document
from weasyprint import HTML

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
        file.save(input_path)

        # Read the DOCX content
        try:
            doc = Document(input_path)
            content = ''.join(f"<p>{para.text}</p>" for para in doc.paragraphs)
        except Exception as e:
            return render_template('index.html', message=f"Failed to read DOCX: {e}")

        # Generate HTML and convert to PDF
        html = f"""
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    padding: 30px;
                }}
                p {{
                    margin-bottom: 15px;
                }}
            </style>
        </head>
        <body>
            {content}
        </body>
        </html>
        """

        output_pdf = os.path.join(OUTPUT_FOLDER, filename.replace('.docx', '.pdf'))
        try:
            HTML(string=html).write_pdf(output_pdf)
            return send_file(output_pdf, as_attachment=True)
        except Exception as e:
            return render_template('index.html', message=f"PDF creation failed: {e}")
    else:
        return render_template('index.html', message="Only .docx files are supported")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
