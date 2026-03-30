
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import Response
from fastapi.middleware.cors import CORSMiddleware
import tempfile
import os
from pdf2docx import Converter

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/convert")
async def convert_pdf(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        content = await file.read()
        tmp_pdf.write(content)
        tmp_pdf_path = tmp_pdf.name
    tmp_docx_path = tmp_pdf_path.replace(".pdf", ".docx")
    cv = Converter(tmp_pdf_path)
    cv.convert(tmp_docx_path)
    cv.close()
    with open(tmp_docx_path, "rb") as f:
        docx_content = f.read()
    os.unlink(tmp_pdf_path)
    os.unlink(tmp_docx_path)
    return Response(
        content=docx_content,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=converted.docx"}
    )

@app.get("/")
def health():
    return {"status": "ok"}
