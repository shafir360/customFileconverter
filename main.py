import os
from flask import Flask, jsonify, request
from werkzeug.utils import secure_filename
from pptx_utils import allowed_file, extract_text_from_pptx_markdown

app = Flask(__name__)

# Set up upload configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pptx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/extract-text', methods=['POST'])
def extract_text():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in the request'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        file.save(filepath)

        try:
            extracted_text = extract_text_from_pptx_markdown(filepath)
        except Exception as e:
            os.remove(filepath)
            return jsonify({'error': f'Error processing PPTX file: {str(e)}'}), 500

        os.remove(filepath)
        return jsonify({'extracted_text': extracted_text})
    else:
        return jsonify({'error': 'Invalid file type. Only PPTX files are allowed.'}), 400

@app.route('/')
def index():
    return jsonify({"Choo Choo": "Welcome to your Flask app 🚅"})

if __name__ == '__main__':
    app.run(debug=True, port=os.getenv("PORT", default=5000))
