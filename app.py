import os
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import pypandoc

app = Flask(__name__)

# Setup file upload and output directories
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

# Allowed file extensions
ALLOWED_EXTENSIONS = {'docx'}

# Make sure the upload and output directories exist
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
        return "No file part in request"
    
    file = request.files['file']
    
    if file.filename == '':
        return "No selected file"
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        input_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(input_path)

        output_filename = os.path.splitext(filename)[0] + '.pdf'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        try:
            # Use pypandoc to convert the file
            pypandoc.convert_file(input_path, 'pdf', outputfile=output_path)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            return f"Conversion failed: {str(e)}"
    else:
        return "Invalid file type. Please upload a .docx file."

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Use the correct port on Render
    app.run(host="0.0.0.0", port=port)
