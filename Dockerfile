FROM python:3.10

# Install system dependencies
RUN apt-get update && apt-get install -y libreoffice && apt-get clean

# Set up app
WORKDIR /app
COPY . /app
RUN pip install -r requirements.txt

# Run FastAPI
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "3000"]
