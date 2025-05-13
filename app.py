from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
import pypandoc
import os

# Ensure pandoc is available (required for Render)
pypandoc.download_pandoc()

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"

# Create folders if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def upload_file():
    return render_template('index.html')  # This should be a simple HTML form

@app.route('/convert', methods=['POST'])
def convert_to_pdf():
    if 'file' not in request.files:
        return "No file part in request"

    file = request.files['file']
    if file.filename == '':
        return "No selected file"

    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(input_path)

    output_filename = os.path.splitext(filename)[0] + '.pdf'
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    try:
        pypandoc.convert_file(input_path, 'pdf', outputfile=output_path)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return f"Conversion failed: {str(e)}"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
