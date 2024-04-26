from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import os
import zipfile
from your_extract_info_code import process_cvs

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def extract_zip(file_path, extract_dir):
    """Extracts the contents of a ZIP file to a directory."""
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file:
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)
            if file.filename.endswith('.zip'):
                extract_dir = os.path.join(app.config['UPLOAD_FOLDER'], 'extracted')
                os.makedirs(extract_dir, exist_ok=True)
                extract_zip(file_path, extract_dir)
                process_cvs(extract_dir, 'extracted_info.xlsx')
                return redirect(url_for('download_file', filename='extracted_info.xlsx'))
            else:
                process_cvs(file_path, 'extracted_info.xlsx')
                return redirect(url_for('download_file', filename='extracted_info.xlsx'))

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.secret_key = 'super_secret_key'
    app.run(debug=True, host='0.0.0.0', port=8080)

