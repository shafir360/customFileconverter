import os
from flask import Flask, jsonify, request
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

app = Flask(__name__)

# Set up upload configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pptx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension (pptx)."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pptx_improved(file_path):
    """
    Extracts text from a PPTX file with improved formatting.
    
    Enhancements:
      - Processes each slide with a header ("Slide X:") only if there is non-empty content.
      - Iterates over text frames and table cells.
      - Skips slides that end up with no text content.
    
    :param file_path: Path to the PPTX file.
    :return: A string with the structured text extracted from the presentation.
    """
    prs = Presentation(file_path)
    slides_text = []
    
    # Process each slide, numbering them for clarity.
    for idx, slide in enumerate(prs.slides, start=1):
        slide_lines = []
        for shape in slide.shapes:
            # For shapes with a text_frame, iterate through paragraphs and runs.
            if hasattr(shape, "text_frame") and shape.text_frame is not None:
                for paragraph in shape.text_frame.paragraphs:
                    paragraph_text = "".join(run.text for run in paragraph.runs).strip()
                    if paragraph_text:
                        slide_lines.append(paragraph_text)
            # For table shapes, extract text cell by cell.
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table = shape.table
                for row in table.rows:
                    row_cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if row_cells:
                        slide_lines.append(" | ".join(row_cells))
        # Only add slide header if any text was collected.
        if slide_lines:
            header = f"Slide {idx}:"
            slides_text.append("\n".join([header] + slide_lines))
    
    return "\n\n".join(slides_text)


@app.route('/extract-text', methods=['POST'])
def extract_text():
    """
    Endpoint to extract text from an uploaded PPTX file.

    Expects a form-data POST request with the file attached under the key 'file'.
    """
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in the request'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        # Ensure the upload directory exists.
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        file.save(filepath)

        try:
            # Use the improved extraction function.
            extracted_text = extract_text_from_pptx_improved(filepath)
        except Exception as e:
            os.remove(filepath)
            return jsonify({'error': f'Error processing PPTX file: {str(e)}'}), 500

        # Clean up the temporary file.
        os.remove(filepath)
        return jsonify({'extracted_text': extracted_text})
    else:
        return jsonify({'error': 'Invalid file type. Only PPTX files are allowed.'}), 400

@app.route('/')
def index():
    return jsonify({"Choo Choo": "Welcome to your Flask app 🚅"})

if __name__ == '__main__':
    app.run(debug=True, port=os.getenv("PORT", default=5000))
