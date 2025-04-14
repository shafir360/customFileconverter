FROM python:3.9-slim

WORKDIR /app

# Copy requirements file and install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy the application code
COPY extract_text.py .

# Expose a port (e.g., 8000)
EXPOSE 8000

# Start the Flask application
CMD ["python", "extract_text.py"]
