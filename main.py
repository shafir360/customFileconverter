import os
from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename
from pptx_utils import allowed_file, extract_text_from_pptx_markdown
from docx_generator import build_word_document  # Import our business logic for document generation
import requests

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
        return jsonify({'error': 'No selected file.'}), 400

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
    
    Optional parameters:
      - "include_images": Boolean (default False). If false, image sections will be skipped.
      - "webhook_url": String (optional). If provided, it will be used to generate images.
      
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
        webhook_url = data.get('webhook_url')  # Optional webhook URL; if not provided, images are skipped.
        output_filename = "generated_document.docx"

        # Call the document generator business logic with the provided webhook_url.
        generated_file = build_word_document(
            data,
            output_filename=output_filename,
            include_images=include_images,
            webhook_url=webhook_url
        )

        # Return the generated file as an attachment.
        return send_file(generated_file, as_attachment=True)

    except Exception as e:
        return jsonify({'error': f'Error generating document: {str(e)}'}), 500
    


def is_url_reachable(url, timeout=5):
    """
    Returns True if a HEAD request to `url` succeeds with a 2xx or 3xx status.
    """
    try:
        resp = requests.head(url, timeout=timeout)
        return resp.status_code < 400
    except requests.RequestException:
        return False



@app.route('/check-connection', methods=['POST'])
def check_connection():
    """
    Expects JSON:
      {
        "url": "<full target URL>",
        "data": { ... }            # JSON body to POST if reachable
      }
    First does a HEAD to the URL; if that fails, returns 502.
    Otherwise does a POST(url, json=data) and streams back the response.
    """
    if not request.is_json:
        return jsonify({'error': 'Request must be JSON with "url" and "data"'}), 400

    payload = request.get_json()
    url = payload.get('url')
    data = payload.get('data')

    if not url:
        return jsonify({'error': 'Missing "url" in request body'}), 400

    # 1) check reachability
    if not is_url_reachable(url):
        return jsonify({'error': f'Cannot reach {url}'}), 502

    # 2) send the data
    try:
        resp = requests.post(url, json=data, timeout=10)
        resp.raise_for_status()
    except requests.RequestException as e:
        return jsonify({'error': f'Error sending data to {url}: {str(e)}'}), 502

    # 3) return whatever the remote service returned
    try:
        # if JSON, forward it
        return jsonify({
            'status': 'success',
            'remote_status_code': resp.status_code,
            'remote_response': resp.json()
        }), 200
    except ValueError:
        # non-JSON response
        return (resp.text, resp.status_code, {'Content-Type': resp.headers.get('Content-Type', 'text/plain')})

@app.route('/')
def index():
    return jsonify({"Choo Choo": "Welcome to your Flask app 🚅"})

if __name__ == '__main__':
    app.run(debug=True, port=os.getenv("PORT", default=5000))
