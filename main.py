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

def extract_text_from_pptx_markdown(file_path):
    """
    Extracts text from a PPTX file and returns it in Markdown format.
    
    Enhancements:
      - Processes each slide and outputs a Markdown header ("## Slide X") only if there is non-empty content.
      - Uses continuous slide numbering for slides that contain text.
      - Extracts text from text frames and tables.
      - Does not remove duplicate lines.
    
    :param file_path: Path to the PPTX file.
    :return: A Markdown formatted string containing the structured text from the presentation.
    """
    prs = Presentation(file_path)
    slides_md = []
    slide_num = 1  # Continuous numbering for non-empty slides
    
    for slide in prs.slides:
        slide_lines = []
        for shape in slide.shapes:
            # Process shapes with text frames.
            if hasattr(shape, "text_frame") and shape.text_frame is not None:
                for paragraph in shape.text_frame.paragraphs:
                    paragraph_text = "".join(run.text for run in paragraph.runs).strip()
                    if paragraph_text:
                        slide_lines.append(paragraph_text)
            # Process table shapes.
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table = shape.table
                for row in table.rows:
                    row_text = " | ".join(cell.text.strip() for cell in row.cells if cell.text.strip())
                    if row_text:
                        slide_lines.append(row_text)
        # Add slide only if there is content.
        if slide_lines:
            header = f"## Slide {slide_num}"
            # Insert a blank line after the header for better Markdown readability.
            slide_md = "\n".join([header, ""] + slide_lines)
            slides_md.append(slide_md)
            slide_num += 1
            
    return "\n\n".join(slides_md)

@app.route('/extract-text', methods=['POST'])
def extract_text():
    """
    Endpoint to extract text from an uploaded PPTX file in Markdown format.
    
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
            # Use the Markdown extraction function.
            extracted_text = extract_text_from_pptx_markdown(filepath)
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
