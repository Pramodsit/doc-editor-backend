from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from tempfile import NamedTemporaryFile
from weasyprint import HTML
from docx import Document
import traceback

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

class ContentModel(BaseModel):
    content: str

@app.get("/")
async def root():
    return {"message": "Welcome to the Enhanced Doc Editor API!"}

@app.post("/export-pdf/")
async def export_pdf(payload: ContentModel):
    html_content = payload.content
    if not html_content.strip():
        raise HTTPException(status_code=400, detail="Missing HTML content")

    html_template = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; margin: 1cm; }}
            table, th, td {{ border: 1px solid black; border-collapse: collapse; padding: 5px; }}
            h1,h2,h3,h4,h5,h6 {{ font-weight: bold; }}
            span, p {{ white-space: pre-wrap; }}
        </style>
    </head>
    <body>{html_content}</body>
    </html>
    """

    try:
        tmp_pdf = NamedTemporaryFile(delete=False, suffix=".pdf")
        HTML(string=html_template).write_pdf(tmp_pdf.name)
        tmp_pdf.close()

        return FileResponse(
            tmp_pdf.name,
            media_type="application/pdf",
            filename="document.pdf",
            background=True,
        )
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"PDF generation failed: {str(e)}")

@app.post("/upload-docx/")
async def upload_docx(file: UploadFile = File(...)):
    try:
        doc = Document(file.file)
        html_output = ""

        for para in doc.paragraphs:
            html_output += "<p>"
            for run in para.runs:
                style = ""
                if run.font.color and run.font.color.rgb:
                    style += f"color: #{run.font.color.rgb};"
                if run.bold:
                    style += "font-weight: bold;"
                if run.italic:
                    style += "font-style: italic;"
                text = run.text.replace("<", "&lt;").replace(">", "&gt;")
                html_output += f'<span style="{style}">{text}</span>'
            html_output += "</p>"

        return JSONResponse(content={"html": html_output})
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Failed to convert DOCX to HTML: {str(e)}")
