from pathlib import Path
import shutil
import uuid
import os
import asyncio
import logging
from typing import Optional, Dict

import httpx
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks, Header, Request
from fastapi.responses import FileResponse, JSONResponse, Response

from libreoffice import convert_pptx_to_pdf, resolve_libreoffice_path

SHOW_DOCS = os.getenv("SHOW_DOCS", "false").lower() == "true"

if SHOW_DOCS:
    app = FastAPI(title="pptx2pdf", version="1.0.0")
else:
    # 彻底禁用 Swagger, ReDoc 和 OpenAPI JSON 结构
    app = FastAPI(
        title="pptx2pdf", 
        version="1.0.0",
        docs_url=None, 
        redoc_url=None, 
        openapi_url=None
    )

BASE_DIR = Path(__file__).parent.resolve()
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "outputs"

for d in (UPLOAD_DIR, OUTPUT_DIR):
    d.mkdir(parents=True, exist_ok=True)


# Basic logging setup; control via LOG_LEVEL env var (e.g., DEBUG, INFO)
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(level=LOG_LEVEL, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logger = logging.getLogger("pptx2pdf")


def _safe_ext(filename: str) -> str:
    return Path(filename).suffix.lower()


def _cleanup_paths(*paths: Path) -> None:
    for p in paths:
        try:
            if p.exists():
                p.unlink()
        except Exception:
            # Best-effort cleanup; ignore errors
            pass


def _build_multipart_file(field_name: str, filename: str, content: bytes, content_type: str = "application/pdf"):
    """Create a multipart/form-data body containing a single file field.

    Returns (body_bytes, content_type_header_value)
    """
    boundary = f"----pptx2pdf-{uuid.uuid4().hex}"
    crlf = "\r\n"
    # Build parts
    headers = (
        f"--{boundary}{crlf}"
        f"Content-Disposition: form-data; name=\"{field_name}\"; filename=\"{filename}\"{crlf}"
        f"Content-Type: {content_type}{crlf}{crlf}"
    ).encode("utf-8")
    footer = (f"{crlf}--{boundary}--{crlf}").encode("utf-8")
    body = headers + content + footer
    return body, f"multipart/form-data; boundary={boundary}"


def _effective_parser_url(env_first: bool = True) -> str:
    env_url = os.getenv("PARSER_URL") or os.getenv("PARSE_URL")
    return env_url or "http://10.81.194.110:31007/file_parse"


def _parser_query_params(request: Request) -> Dict[str, str]:
    prefix = "parser_query_"
    params: Dict[str, str] = {}
    for key, value in request.query_params.multi_items():
        if not key.startswith(prefix):
            continue
        trimmed = key[len(prefix):]
        if not trimmed:
            raise HTTPException(status_code=400, detail="parser_query_ prefix requires a field name")
        params[trimmed] = value
    return params


@app.get("/healthz")
def healthz():
    """Basic health check with quick environment info."""
    try:
        lo_path = str(resolve_libreoffice_path())
    except Exception:
        lo_path = None
    logger.debug("Health check - libreoffice=%s, parser_url=%s", bool(lo_path), _effective_parser_url())
    return JSONResponse({
        "status": "ok",
        "libreoffice": bool(lo_path),
        "libreoffice_path": lo_path,
        "parser_url": _effective_parser_url(),
    })


@app.post("/convert")
async def convert(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """
    Upload a PPT/PPTX and receive the converted PDF.

    Returns application/pdf with a filename derived from the uploaded file name.
    """
    if not file or not file.filename:
        raise HTTPException(status_code=400, detail="No file uploaded")

    ext = _safe_ext(file.filename)
    if ext not in {".ppt", ".pptx"}:
        raise HTTPException(status_code=400, detail="Only .ppt or .pptx files are supported")

    # Store upload with a unique name to avoid collisions
    unique_stem = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{unique_stem}{ext}"

    try:
        with input_path.open("wb") as out_f:
            shutil.copyfileobj(file.file, out_f)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to save upload: {exc}")

    # Run conversion; the output PDF will use the unique stem as its base name
    try:
        pdf_path = convert_pptx_to_pdf(input_path, OUTPUT_DIR)
    except FileNotFoundError as exc:
        # Likely missing expected output file
        # Cleanup upload before raising
        _cleanup_paths(input_path)
        raise HTTPException(status_code=500, detail=str(exc))
    except Exception as exc:
        _cleanup_paths(input_path)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}")

    # Return as a file response with a friendly filename derived from original
    download_name = f"{Path(file.filename).stem}.pdf"

    # Schedule cleanup after response is sent
    background_tasks.add_task(_cleanup_paths, input_path, pdf_path)

    return FileResponse(
        path=str(pdf_path),
        media_type="application/pdf",
        filename=download_name,
        background=background_tasks,
    )


@app.post("/convert_multipart")
async def convert_multipart(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """
    Upload a PPT/PPTX and receive the converted PDF wrapped as multipart/form-data
    with a single field named "file" (compatible with downstream `file_parse`).
    """
    if not file or not file.filename:
        raise HTTPException(status_code=400, detail="No file uploaded")

    ext = _safe_ext(file.filename)
    if ext not in {".ppt", ".pptx"}:
        raise HTTPException(status_code=400, detail="Only .ppt or .pptx files are supported")

    unique_stem = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{unique_stem}{ext}"

    try:
        with input_path.open("wb") as out_f:
            shutil.copyfileobj(file.file, out_f)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to save upload: {exc}")

    try:
        pdf_path = convert_pptx_to_pdf(input_path, OUTPUT_DIR)
        pdf_bytes = pdf_path.read_bytes()
    except Exception as exc:
        _cleanup_paths(input_path)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}")

    # Friendly filename matches original stem
    download_name = f"{Path(file.filename).stem}.pdf"

    # Build multipart body with field name 'file'
    body, content_type_header = _build_multipart_file(
        field_name="file",
        filename=download_name,
        content=pdf_bytes,
        content_type="application/pdf",
    )

    # Cleanup temp files after response is sent
    background_tasks.add_task(_cleanup_paths, input_path, pdf_path)

    return Response(content=body, media_type=content_type_header, background=background_tasks)


def _resolve_parser_url(query_override: Optional[str], header_override: Optional[str]) -> str:
    # Precedence: query param > header > env vars > default
    url = (query_override or header_override or _effective_parser_url())
    if not (url.startswith("http://") or url.startswith("https://")):
        raise HTTPException(status_code=400, detail="parser_url must start with http:// or https://")
    return url


@app.post("/convert_and_parse")
async def convert_and_parse(
    request: Request,
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    parser_url: Optional[str] = None,
    x_parser_url: Optional[str] = Header(default=None, alias="X-Parser-Url"),
    # Optional tuning parameters for the downstream parser (as query params)
    return_middle_json: bool = False,
    return_model_output: bool = False,
    return_md: bool = True,
    return_images: bool = False,
    end_page_id: int = 99999,
    parse_method: str = "auto",
    start_page_id: int = 0,
    lang_list: str = "ch",
    output_dir: str = "./output",
    server_url: str = "string",
    return_content_list: bool = False,
    backend: str = "pipeline",
    table_enable: bool = True,
    formula_enable: bool = True,
):
    """
    Upload a PPT/PPTX, convert it to PDF, call an external PDF parsing service,
    and return the parsing service's JSON response.
    """
    if not file or not file.filename:
        raise HTTPException(status_code=400, detail="No file uploaded")

    ext = _safe_ext(file.filename)
    if ext not in {".ppt", ".pptx"}:
        raise HTTPException(status_code=400, detail="Only .ppt or .pptx files are supported")

    unique_stem = uuid.uuid4().hex
    input_path = UPLOAD_DIR / f"{unique_stem}{ext}"

    try:
        with input_path.open("wb") as out_f:
            shutil.copyfileobj(file.file, out_f)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to save upload: {exc}")

    # Convert in a thread to avoid blocking the event loop
    try:
        loop = asyncio.get_running_loop()
        pdf_path: Path = await loop.run_in_executor(None, convert_pptx_to_pdf, input_path, OUTPUT_DIR)
    except Exception as exc:
        _cleanup_paths(input_path)
        raise HTTPException(status_code=500, detail=f"Conversion failed: {exc}")

    download_name = f"{Path(file.filename).stem}.pdf"
    try:
        size_bytes = pdf_path.stat().st_size
    except Exception:
        size_bytes = -1
    logger.info("Converted %s -> %s (%s bytes)", input_path.name, pdf_path.name, size_bytes)

    # Determine target parsing service URL (query > header > env > default)
    target_url = _resolve_parser_url(parser_url, x_parser_url)
    logger.info("Posting to parser: %s", target_url)

    # Prepare form fields expected by the downstream parser
    form_data = {
        "return_middle_json": str(return_middle_json).lower(),
        "return_model_output": str(return_model_output).lower(),
        "return_md": str(return_md).lower(),
        "return_images": str(return_images).lower(),
        "end_page_id": str(end_page_id),
        "parse_method": parse_method,
        "start_page_id": str(start_page_id),
        "lang_list": lang_list,
        "output_dir": output_dir,
        "server_url": server_url,
        "return_content_list": str(return_content_list).lower(),
        "backend": backend,
        "table_enable": str(table_enable).lower(),
        "formula_enable": str(formula_enable).lower(),
    }
    parser_query_params = _parser_query_params(request)
    logger.debug("Parser form fields: %s", form_data)
    if parser_query_params:
        logger.debug("Parser query params: %s", parser_query_params)

    # Call downstream parsing service with the generated PDF
    try:
        timeout = httpx.Timeout(300.0, connect=20.0)
        async with httpx.AsyncClient(timeout=timeout) as client:
            headers = {"accept": "application/json"}

            # Try with field name 'files'
            with pdf_path.open("rb") as f1:
                files1 = [("files", (download_name, f1, "application/pdf"))]
                resp = await client.post(
                    target_url,
                    params=parser_query_params or None,
                    data=form_data,
                    files=files1,
                    headers=headers,
                )
                logger.info("Parser response (files): status=%s, content-type=%s", resp.status_code, resp.headers.get("content-type"))

            # If the server clearly rejects, try alternative field name 'file'
            if resp.status_code >= 400:
                with pdf_path.open("rb") as f2:
                    files2 = [("file", (download_name, f2, "application/pdf"))]
                    resp2 = await client.post(
                        target_url,
                        params=parser_query_params or None,
                        data=form_data,
                        files=files2,
                        headers=headers,
                    )
                    logger.info("Parser response (file): status=%s, content-type=%s", resp2.status_code, resp2.headers.get("content-type"))
                # Prefer successful response if fallback worked
                if 200 <= resp2.status_code < 300:
                    resp = resp2

        # Try to decode JSON; if not JSON, raise 502 with snippet
        try:
            payload = resp.json()
        except ValueError:
            text = resp.text[:1000]
            logger.error("Downstream returned non-JSON. status=%s snippet=%s", resp.status_code, text)
            raise HTTPException(status_code=502, detail=f"Downstream returned non-JSON (status {resp.status_code}). Snippet: {text}")

        # Cleanup temp files after sending response
        background_tasks.add_task(_cleanup_paths, input_path, pdf_path)

        return JSONResponse(status_code=resp.status_code, content=payload)

    except HTTPException:
        # Re-raise HTTPExceptions as-is
        logger.exception("HTTP error while calling parser")
        _cleanup_paths(input_path, pdf_path)
        raise
    except Exception as exc:
        logger.exception("Exception while calling parser")
        _cleanup_paths(input_path, pdf_path)
        raise HTTPException(status_code=502, detail=f"Failed to call parser: {exc}")


# Optional root
@app.get("/")
def root():
    return {"service": "pptx2pdf", "endpoints": ["GET /healthz", "POST /convert"]}
