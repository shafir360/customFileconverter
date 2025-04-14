from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse
import os
import uuid
import subprocess

app = FastAPI()

@app.post("/convert")
async def convert_pptx_to_pdf(file: UploadFile = File(...)):
    # Save uploaded PPTX
    input_path = f"/tmp/{uuid.uuid4()}.pptx"
    with open(input_path, "wb") as f:
        f.write(await file.read())

    output_path = input_path.replace(".pptx", ".pdf")

    # Convert using LibreOffice CLI
    subprocess.run([
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", "/tmp",
        input_path
    ], check=True)

    return FileResponse(output_path, media_type="application/pdf", filename="converted.pdf")
