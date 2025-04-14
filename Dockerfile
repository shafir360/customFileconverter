FROM python:3.10

# Install LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean

# Copy project files
WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir -r requirements.txt

# Expose port
EXPOSE 3000

# Run FastAPI with uvicorn
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "3000"]
