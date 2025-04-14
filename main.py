FROM python:3.10

# Install LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean

# Set workdir and copy code
WORKDIR /app
COPY . /app

# Install Python deps
RUN pip install --no-cache-dir -r requirements.txt

# Run using shell so $PORT gets expanded
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port $PORT"]
