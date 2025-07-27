from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
import zipfile
import io
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'files[]' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400
    
    files = request.files.getlist('files[]')
    if not files or files[0].filename == '':
        return jsonify({'error': 'No files selected'}), 400

    # Create a ZIP file in memory
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file in files:
            if not file.filename.endswith(('.xlsx', '.xls')):
                continue

            try:
                # Save the uploaded file
                filename = secure_filename(file.filename)
                excel_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(excel_path)

                # Convert to CSV
                df = pd.read_excel(excel_path)
                csv_filename = f"{os.path.splitext(filename)[0]}.csv"
                csv_path = os.path.join(app.config['UPLOAD_FOLDER'], csv_filename)
                df.to_csv(csv_path, index=False)

                # Add CSV to ZIP
                with open(csv_path, 'rb') as f:
                    zf.writestr(csv_filename, f.read())

                # Clean up files
                os.remove(excel_path)
                os.remove(csv_path)

            except Exception as e:
                return jsonify({'error': f'Error processing {filename}: {str(e)}'}), 500

    memory_file.seek(0)
    return send_file(
        memory_file,
        mimetype='application/zip',
        as_attachment=True,
        download_name='converted_files.zip'
    )

if __name__ == '__main__':
    app.run(debug=True, port=5001)