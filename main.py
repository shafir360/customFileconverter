import os
from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename
from pptx_utils import allowed_file, extract_text_from_pptx_markdown
from docx_generator import build_word_document  # Import our business logic for document generation

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

@app.route('/generate-docx', methods=['POST'])
def generate_docx():
    """
    Endpoint to generate a Word document from a structured JSON object.

    Expected JSON input should have at least:
      - "title": A string for the document title.
      - "sections": An array of objects (each with a "type" and "content").

    Optional parameter:
      - "include_images": Boolean (default False). If set to False, image sections will be skipped.

    Returns:
      The generated Word document as an attachment.
    """
    if not request.is_json:
        return jsonify({'error': 'Request must be in JSON format.'}), 400

    try:
        data = request.get_json()
        # Check for required keys in JSON
        if 'title' not in data or 'sections' not in data:
            return jsonify({'error': 'JSON must contain "title" and "sections" keys.'}), 400

        include_images = data.get('include_images', False)
        output_filename = "generated_document.docx"

        # Call the document generator business logic
        generated_file = build_word_document(data, output_filename=output_filename, include_images=include_images)

        # Return the generated file as an attachment
        return send_file(generated_file, as_attachment=True)

    except Exception as e:
        return jsonify({'error': f'Error generating document: {str(e)}'}), 500

@app.route('/')
def index():
    return jsonify({"Choo Choo": "Welcome to your Flask app 🚅"})

if __name__ == '__main__':
    app.run(debug=True, port=os.getenv("PORT", default=5000))
