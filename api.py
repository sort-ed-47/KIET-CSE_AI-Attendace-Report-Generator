import os
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, Response
from fastapi.staticfiles import StaticFiles
from openai import OpenAI

from agent import (
    DEFAULT_THRESHOLD,
    generate_defaulter_emails,
    generate_lecture_insights,
    generate_summary_report,
    parse_excel,
)
from agent import save_pdf_report as build_pdf_report


app = FastAPI(
    title="Attendance Report Generator API",
    version="1.0.0",
    description="Generate professional attendance PDF reports from KIET-format Excel files.",
)
app.mount("/static", StaticFiles(directory="static"), name="static")


def get_client() -> OpenAI:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY is not configured.")
    if not api_key.isascii():
        raise HTTPException(status_code=500, detail="OPENAI_API_KEY contains invalid characters.")
    return OpenAI(api_key=api_key)


@app.get("/")
def home():
    return FileResponse("static/index.html")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/reports")
async def create_report(
    file: UploadFile = File(...),
    threshold: float = Form(DEFAULT_THRESHOLD),
    include_emails: bool = Form(False),
):
    filename = file.filename or "attendance.xlsx"
    if Path(filename).suffix.lower() != ".xlsx":
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported.")
    if not 0 <= threshold <= 100:
        raise HTTPException(status_code=400, detail="threshold must be between 0 and 100.")

    client = get_client()

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / Path(filename).name
        input_path.write_bytes(await file.read())

        try:
            data = parse_excel(str(input_path), threshold)
            summary = generate_summary_report(client, data)
            insights = generate_lecture_insights(client, data)
            emails = generate_defaulter_emails(client, data) if include_emails else []
            pdf_name = build_pdf_report(data, summary, insights, emails, tmpdir, filename)
            pdf_bytes = (Path(tmpdir) / pdf_name).read_bytes()
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"Report generation failed: {exc}") from exc

    headers = {"Content-Disposition": f'attachment; filename="{Path(filename).stem}.pdf"'}
    return Response(content=pdf_bytes, media_type="application/pdf", headers=headers)
