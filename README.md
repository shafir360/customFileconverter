# PPTX Text Extraction Service

This is a simple Flask application that extracts text from a PowerPoint (.pptx) file.

## Endpoints

- **GET /**: Returns a welcome message.
- **POST /extract**: Accepts a PPTX file via form-data (key: `file`) and returns the extracted text as JSON.

## Running Locally

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
