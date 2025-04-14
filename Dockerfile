FROM python:3.10

# Install LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean

# Set work directory and copy files
WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir -r requirements.txt

# Railway will inject PORT as an environment variable
ENV PORT=3000

# Use shell-style CMD for env substitution
CMD uvicorn main:app --host 0.0.0.0 --port $PORT
