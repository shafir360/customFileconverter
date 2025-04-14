from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse
from pptx import Presentation
import io

app = FastAPI()

@app.post("/extract-pptx-text")
async def extract_pptx_text(file: UploadFile = File(...)):
    if file.content_type != "application/vnd.openxmlformats-officedocument.presentationml.presentation":
        return JSONResponse(status_code=400, content={"error": "File must be a .pptx"})

    contents = await file.read()
    prs = Presentation(io.BytesIO(contents))

    text = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text.append(shape.text)

    return {"text": "\n".join(text)}
