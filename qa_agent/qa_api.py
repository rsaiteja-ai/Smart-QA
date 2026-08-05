from pathlib import Path
import shutil
import uuid

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

import test_case_reviewer as reviewer
import qa_agent_2 as generator

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CHECKLIST = BASE_DIR / "TestCaseReview_CheckList.xlsx"
UPLOAD_DIR = BASE_DIR / "temp_uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="QA Agent API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def safe_upload_path(filename: str) -> Path:
    suffix = Path(filename).suffix.lower()
    stem = Path(filename).stem.replace(" ", "_")
    safe_stem = "".join(char for char in stem if char.isalnum() or char in ("_", "-")) or "upload"
    return UPLOAD_DIR / f"{safe_stem}_{uuid.uuid4().hex[:8]}{suffix}"


def save_upload(file: UploadFile) -> Path:
    stored_path = safe_upload_path(file.filename or "upload")
    with stored_path.open("wb") as target:
        shutil.copyfileobj(file.file, target)
    return stored_path


def run_review_workflow(stored_path: Path):
    checklist_path = UPLOAD_DIR / f"{stored_path.stem}_review_checklist.xlsx"
    summary_path = UPLOAD_DIR / f"{stored_path.stem}_review_summary.txt"
    json_path = UPLOAD_DIR / f"{stored_path.stem}_review_summary.json"
    shutil.copyfile(DEFAULT_CHECKLIST, checklist_path)

    checklist = reviewer.update_checklist(str(stored_path), str(checklist_path))
    summary = reviewer.generate_summary_report(checklist, str(summary_path))
    reviewer.generate_json_report(checklist, str(json_path))
    return summary, checklist_path, summary_path, json_path


def generated_output_path(stored_path: Path) -> Path:
    return UPLOAD_DIR / f"{stored_path.stem}_generated_test_cases.xlsx"


def build_generation_response(file: UploadFile, stored_path: Path, generated_path: Path, summary: str,
                              checklist_path: Path, summary_path: Path):
    return {
        "message": "Test cases generated and reviewed successfully.",
        "source_file": file.filename,
        "stored_file": stored_path.name,
        "generated_file": generated_path.name,
        "reviewed_checklist": checklist_path.name,
        "summary_report": summary_path.name,
        "download_file": generated_path.name,
        "download_name": f"{Path(file.filename).stem}_generated_test_cases.xlsx",
        "download_label": "Download generated test cases",
        "summary": summary,
        "status": "success",
    }


def llm_config_from_form(llm_endpoint: str | None, api_key: str | None, model: str | None):
    return {
        "url": (llm_endpoint or "").strip(),
        "api_key": (api_key or "").strip(),
        "model": (model or "").strip(),
    }


def get_upload_artifact(filename: str) -> Path:
    artifact_path = (UPLOAD_DIR / Path(filename).name).resolve()
    if UPLOAD_DIR.resolve() not in artifact_path.parents or not artifact_path.exists():
        raise HTTPException(status_code=404, detail="Requested file was not found.")
    if not artifact_path.is_file():
        raise HTTPException(status_code=404, detail="Requested file was not found.")
    return artifact_path


@app.get("/api/health")
def health():
    return {"status": "ok", "service": "QA Agent Studio API"}


@app.post("/api/test-llm-connection")
async def test_llm_connection(
    llm_endpoint: str | None = Form(default=None),
    api_key: str | None = Form(default=None),
    model: str | None = Form(default=None),
):
    try:
        result = generator.test_llm_connection(llm_config_from_form(llm_endpoint, api_key, model))
        return {
            "status": "success",
            "message": f"LLM connection successful for model {result['model']}.",
            "model": result["model"],
        }
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/generate-from-excel")
async def generate_from_excel(
    file: UploadFile = File(...),
    llm_endpoint: str | None = Form(default=None),
    api_key: str | None = Form(default=None),
    model: str | None = Form(default=None),
):
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Please upload an Excel file (.xlsx or .xls).")

    stored_path = save_upload(file)

    try:
        llm_config = llm_config_from_form(llm_endpoint, api_key, model)
        generated_path = generator.generate_from_excel(stored_path, generated_output_path(stored_path), llm_config)
        summary, checklist_path, summary_path, _ = run_review_workflow(generated_path)
        return build_generation_response(file, stored_path, generated_path, summary, checklist_path, summary_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Excel workflow failed: {exc}") from exc


@app.post("/api/generate-from-docx")
async def generate_from_docx(
    file: UploadFile = File(...),
    llm_endpoint: str | None = Form(default=None),
    api_key: str | None = Form(default=None),
    model: str | None = Form(default=None),
):
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Please upload a DOCX file.")

    stored_path = save_upload(file)

    try:
        llm_config = llm_config_from_form(llm_endpoint, api_key, model)
        generated_path = generator.generate_from_docx(stored_path, generated_output_path(stored_path), llm_config)
        summary, checklist_path, summary_path, _ = run_review_workflow(generated_path)
        return build_generation_response(file, stored_path, generated_path, summary, checklist_path, summary_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"DOCX workflow failed: {exc}") from exc


@app.post("/api/generate-from-image")
async def generate_from_image(
    file: UploadFile = File(...),
    llm_endpoint: str | None = Form(default=None),
    api_key: str | None = Form(default=None),
    model: str | None = Form(default=None),
):
    allowed = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
    file_ext = Path(file.filename).suffix.lower() if file.filename else ""
    if file_ext not in allowed:
        raise HTTPException(status_code=400, detail="Please upload an image file (PNG, JPG, JPEG, WEBP, or BMP).")

    stored_path = save_upload(file)

    try:
        llm_config = llm_config_from_form(llm_endpoint, api_key, model)
        generated_path = generator.generate_from_image(stored_path, generated_output_path(stored_path), llm_config)
        summary, checklist_path, summary_path, _ = run_review_workflow(generated_path)
        return build_generation_response(file, stored_path, generated_path, summary, checklist_path, summary_path)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Image workflow failed: {exc}") from exc


@app.post("/api/review-results")
async def review_results(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(status_code=400, detail="Please upload a generated test cases workbook.")

    stored_path = save_upload(file)

    try:
        summary, checklist_path, summary_path, json_path = run_review_workflow(stored_path)
        return {
            "message": "Review workflow completed successfully.",
            "source_file": file.filename,
            "stored_file": stored_path.name,
            "reviewed_checklist": checklist_path.name,
            "summary_report": summary_path.name,
            "json_report": json_path.name,
            "download_file": stored_path.name,
            "download_name": Path(file.filename).name,
            "download_label": "Download test cases workbook",
            "summary": summary,
            "status": "success",
        }
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Review workflow failed: {exc}") from exc


@app.get("/api/download/{filename}")
def download_artifact(filename: str):
    artifact_path = get_upload_artifact(filename)
    return FileResponse(
        artifact_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=artifact_path.name,
    )
