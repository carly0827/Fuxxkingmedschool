from __future__ import annotations

import os
import shutil
import sys
import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

BASE_DIR = Path(__file__).resolve().parent
CODE_DIR = BASE_DIR / "code"
if str(CODE_DIR) not in sys.path:
    sys.path.insert(0, str(CODE_DIR))

from build_final_notes_from_jul_yabout import run_pipeline  # noqa: E402

APP_TITLE = "최종본 생성 사이트"
RUNS_DIR = Path(os.getenv("RUNS_DIR", str(BASE_DIR / "data" / "runs")))
RUNS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title=APP_TITLE)
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


def _safe_filename(name: str, default: str) -> str:
    cleaned = Path(name).name.strip().replace("/", "_").replace("\\", "_")
    return cleaned or default


@app.get("/healthz")
def healthz() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "title": APP_TITLE,
            "message": None,
            "error": None,
            "result": None,
        },
    )


@app.post("/generate", response_class=HTMLResponse)
async def generate(
    request: Request,
    jul_pdf: UploadFile = File(...),
    yabout_pdf: UploadFile = File(...),
    transcript_text: str = Form(...),
):
    transcript_text = (transcript_text or "").strip()
    if not transcript_text:
        return templates.TemplateResponse(
            request,
            "index.html",
            {
                "title": APP_TITLE,
                "message": None,
                "error": "전사문 텍스트를 넣어주세요.",
                "result": None,
            },
            status_code=400,
        )

    run_id = uuid.uuid4().hex[:12]
    run_dir = RUNS_DIR / run_id
    input_dir = run_dir / "input"
    output_dir = run_dir / "output"
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    jul_path = input_dir / _safe_filename(jul_pdf.filename or "jul.pdf", "jul.pdf")
    yabout_path = input_dir / _safe_filename(yabout_pdf.filename or "yabout.pdf", "yabout.pdf")

    try:
        jul_path.write_bytes(await jul_pdf.read())
        yabout_path.write_bytes(await yabout_pdf.read())
        (input_dir / "transcript.txt").write_text(transcript_text, encoding="utf-8")

        manifest = run_pipeline(
            jul_path=jul_path,
            yabout_path=yabout_path,
            transcript_text=transcript_text,
            outdir=output_dir,
            keep_merged_pdf=True,
        )
    except Exception as exc:
        shutil.rmtree(run_dir, ignore_errors=True)
        return templates.TemplateResponse(
            request,
            "index.html",
            {
                "title": APP_TITLE,
                "message": None,
                "error": f"생성 중 오류가 발생했습니다: {exc}",
                "result": None,
            },
            status_code=500,
        )

    label_map = [
        ("final_pdf", "최종본 PDF"),
        ("safe_download_zip", "권장 다운로드 ZIP"),
        ("safe_full_archive_zip", "전체 보관 ZIP"),
        ("transcript_digest_pdf", "전체 전사문 요약정리표 PDF"),
        ("transcript_digest_docx", "전체 전사문 요약정리표 DOCX"),
        ("page_plan_json", "페이지 계획 JSON"),
        ("merge_report_json", "줄·야붙 병합 검증 JSON"),
        ("transcript_completeness_txt", "전사문 누락 검토 TXT"),
    ]
    preferred = manifest.get("preferred_delivery_file")
    links = []
    for key, label in label_map:
        value = manifest.get(key)
        if not value:
            continue
        p = Path(value)
        if p.exists() and p.is_file():
            links.append(
                {
                    "label": label,
                    "href": f"/downloads/{run_id}/{p.name}",
                    "filename": p.name,
                    "preferred": bool(preferred and Path(preferred).name == p.name),
                }
            )

    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "title": APP_TITLE,
            "message": "생성이 완료되었습니다.",
            "error": None,
            "result": {
                "run_id": run_id,
                "char_count": manifest.get("transcript_char_count", 0),
                "links": links,
            },
        },
    )


@app.get("/downloads/{run_id}/{filename}")
def download(run_id: str, filename: str):
    if "/" in filename or "\\" in filename:
        raise HTTPException(status_code=400, detail="잘못된 파일명입니다.")
    target = RUNS_DIR / run_id / "output" / filename
    if not target.exists() or not target.is_file():
        raise HTTPException(status_code=404, detail="파일을 찾지 못했습니다.")
    return FileResponse(path=str(target), filename=target.name)


@app.get("/favicon.ico")
def favicon():
    return RedirectResponse(url="/static/favicon.svg")
