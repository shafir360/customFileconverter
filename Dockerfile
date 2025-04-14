FROM python:3.10

# Install LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean

WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir -r requirements.txt

# Railway sets $PORT, but we'll default to 8000 for local
ENV PORT 8000

CMD uvicorn main:app --host 0.0.0.0 --port ${PORT}
