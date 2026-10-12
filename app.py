import os
import shutil
import uuid
from typing import Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

import converter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IS_VERCEL = bool(os.environ.get("VERCEL") or os.environ.get("AWS_LAMBDA_FUNCTION_NAME"))
STORAGE_DIR = "/tmp/storage" if IS_VERCEL else os.path.join(BASE_DIR, "storage")
UPLOAD_DIR = os.path.join(STORAGE_DIR, "uploads")
OUTPUT_DIR = os.path.join(STORAGE_DIR, "outputs")
PREVIEWS_DIR = os.path.join(STORAGE_DIR, "previews")
STATIC_DIR = os.path.join(BASE_DIR, "public") if os.path.exists(os.path.join(BASE_DIR, "public")) else os.path.join(BASE_DIR, "static")

for d in [UPLOAD_DIR, OUTPUT_DIR, PREVIEWS_DIR]:
    os.makedirs(d, exist_ok=True)
if not IS_VERCEL:
    os.makedirs(STATIC_DIR, exist_ok=True)

app = FastAPI(title="Excel to PDF Converter & Margin Styler")

# Mount static folder
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Memory registry for tasks: task_id -> { excel_path, pdf_path, preview_dir, original_name, ... }
tasks_registry = {}


@app.get("/")
async def serve_index():
    index_file = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_file):
        with open(index_file, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse("<h1>Excel to PDF Converter</h1><p>Index file not found.</p>")


@app.get("/api/status")
async def get_system_status():
    excel_ok = converter.check_excel_availability()
    return {
        "status": "online",
        "excel_engine_ready": excel_ok,
        "engine_name": "Microsoft Excel Native COM" if excel_ok else ("ReportLab Engine (Vercel Cloud)" if IS_VERCEL else "ReportLab Fallback"),
        "platform": "Vercel Cloud" if IS_VERCEL else "Local PC Server",
        "storage": {
            "tasks_count": len(tasks_registry)
        }
    }


@app.post("/api/inspect")
async def inspect_uploaded_file(file: UploadFile = File(...)):
    """
    Saves uploaded file temporarily and extracts sheet list and metadata.
    """
    upload_id = str(uuid.uuid4())
    ext = os.path.splitext(file.filename)[1].lower()
    if not ext:
        ext = ".xlsx"
    saved_filename = f"{upload_id}{ext}"
    saved_path = os.path.join(UPLOAD_DIR, saved_filename)

    with open(saved_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    info = converter.inspect_excel_file(saved_path)
    info["upload_id"] = upload_id
    info["original_filename"] = file.filename

    # Register temp file
    tasks_registry[upload_id] = {
        "excel_path": saved_path,
        "original_filename": file.filename
    }

    return info


@app.post("/api/convert")
async def convert_file(
    file: Optional[UploadFile] = File(None),
    upload_id: Optional[str] = Form(None),
    unit: str = Form("cm"),
    top_margin: float = Form(1.91),
    bottom_margin: float = Form(1.91),
    left_margin: float = Form(1.78),
    right_margin: float = Form(1.78),
    header_margin: Optional[float] = Form(None),
    footer_margin: Optional[float] = Form(None),
    orientation: str = Form("portrait"),
    paper_size: str = Form("A4"),
    fit_mode: str = Form("fit_width"),
    center_horizontally: bool = Form(True),
    center_vertically: bool = Form(False),
    sheet_name: str = Form("all"),
    show_gridlines: bool = Form(False),
    show_headings: bool = Form(False)
):
    """
    Converts Excel file to PDF with custom margins and page settings.
    """
    excel_path = None
    original_filename = "document.xlsx"

    # 1. Resolve source Excel file
    if upload_id and upload_id in tasks_registry:
        excel_path = tasks_registry[upload_id]["excel_path"]
        original_filename = tasks_registry[upload_id]["original_filename"]
    elif file is not None:
        new_upload_id = str(uuid.uuid4())
        ext = os.path.splitext(file.filename)[1].lower() or ".xlsx"
        excel_path = os.path.join(UPLOAD_DIR, f"{new_upload_id}{ext}")
        with open(excel_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        original_filename = file.filename
    else:
        raise HTTPException(status_code=400, detail="Tidak ada file yang diunggah atau upload_id tidak valid.")

    if not os.path.exists(excel_path):
        raise HTTPException(status_code=404, detail="File sumber Excel tidak ditemukan.")

    # 2. Convert margins to points based on unit
    if unit.lower() == "inch":
        top_pt = converter.inch_to_points(top_margin)
        bottom_pt = converter.inch_to_points(bottom_margin)
        left_pt = converter.inch_to_points(left_margin)
        right_pt = converter.inch_to_points(right_margin)
        header_pt = converter.inch_to_points(header_margin) if header_margin is not None else None
        footer_pt = converter.inch_to_points(footer_margin) if footer_margin is not None else None
    else:  # Centimeters (cm)
        top_pt = converter.cm_to_points(top_margin)
        bottom_pt = converter.cm_to_points(bottom_margin)
        left_pt = converter.cm_to_points(left_margin)
        right_pt = converter.cm_to_points(right_margin)
        header_pt = converter.cm_to_points(header_margin) if header_margin is not None else None
        footer_pt = converter.cm_to_points(footer_margin) if footer_margin is not None else None

    # 3. Destination paths
    task_id = str(uuid.uuid4())
    pdf_filename = f"{os.path.splitext(original_filename)[0]}.pdf"
    output_pdf_path = os.path.join(OUTPUT_DIR, f"{task_id}.pdf")
    task_preview_dir = os.path.join(PREVIEWS_DIR, task_id)

    # 4. Perform conversion
    try:
        conv_result = converter.convert_excel_to_pdf(
            excel_path=excel_path,
            output_pdf_path=output_pdf_path,
            top_margin_pt=top_pt,
            bottom_margin_pt=bottom_pt,
            left_margin_pt=left_pt,
            right_margin_pt=right_pt,
            header_margin_pt=header_pt,
            footer_margin_pt=footer_pt,
            orientation=orientation,
            paper_size=paper_size,
            fit_mode=fit_mode,
            center_horizontally=center_horizontally,
            center_vertically=center_vertically,
            sheet_name=sheet_name,
            show_gridlines=show_gridlines,
            show_headings=show_headings
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal melakukan konversi: {str(e)}")

    # 5. Generate high-res visual previews for browser
    try:
        preview_data = converter.generate_pdf_previews(output_pdf_path, task_preview_dir)
    except Exception as e:
        preview_data = {"total_pages": 1, "previews": []}

    # Format preview URLs
    for p in preview_data.get("previews", []):
        p["url"] = f"/api/preview/{task_id}/{p['page']}"

    # Store task details
    tasks_registry[task_id] = {
        "pdf_path": output_pdf_path,
        "preview_dir": task_preview_dir,
        "download_name": pdf_filename,
        "original_filename": original_filename,
        "total_pages": preview_data.get("total_pages", 1)
    }

    return {
        "success": True,
        "task_id": task_id,
        "filename": pdf_filename,
        "pdf_size": conv_result["pdf_size"],
        "pdf_size_formatted": conv_result["pdf_size_formatted"],
        "duration_seconds": conv_result["duration_seconds"],
        "engine": conv_result["engine"],
        "total_pages": preview_data.get("total_pages", 1),
        "previews": preview_data.get("previews", []),
        "download_url": f"/api/download/{task_id}",
        "view_url": f"/api/pdf/{task_id}",
        "settings_applied": {
            "unit": unit,
            "margins": {
                "top": top_margin,
                "bottom": bottom_margin,
                "left": left_margin,
                "right": right_margin,
                "header": header_margin,
                "footer": footer_margin
            },
            "orientation": orientation,
            "paper_size": paper_size,
            "fit_mode": fit_mode,
            "sheet_name": sheet_name
        }
    }


@app.get("/api/preview/{task_id}/{page_num}")
async def get_page_preview(task_id: str, page_num: int):
    img_path = os.path.join(PREVIEWS_DIR, task_id, f"page_{page_num}.png")
    if os.path.exists(img_path):
        return FileResponse(img_path, media_type="image/png")
    raise HTTPException(status_code=404, detail="Preview halaman tidak ditemukan.")


@app.get("/api/pdf/{task_id}")
async def get_inline_pdf(task_id: str):
    task = tasks_registry.get(task_id)
    if task and os.path.exists(task["pdf_path"]):
        return FileResponse(
            task["pdf_path"],
            media_type="application/pdf",
            headers={"Content-Disposition": f"inline; filename=\"{task['download_name']}\""}
        )
    raise HTTPException(status_code=404, detail="File PDF tidak ditemukan.")


@app.get("/api/download/{task_id}")
async def download_pdf(task_id: str):
    task = tasks_registry.get(task_id)
    if task and os.path.exists(task["pdf_path"]):
        return FileResponse(
            task["pdf_path"],
            media_type="application/pdf",
            filename=task["download_name"],
            headers={"Content-Disposition": f"attachment; filename=\"{task['download_name']}\""}
        )
    raise HTTPException(status_code=404, detail="File PDF tidak ditemukan.")


import socket

def get_local_ip():
    """Returns local LAN/Wi-Fi IP for HP/Tablet connection."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"


@app.get("/api/network-info")
async def get_network_info():
    ip = get_local_ip()
    port = 8000
    return {
        "ip": ip,
        "port": port,
        "local_url": f"http://localhost:{port}",
        "network_url": f"http://{ip}:{port}"
    }


# Mount root static files so relative URLs work in both browser server and mobile browsers
app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="site_root")


if __name__ == "__main__":
    ip = get_local_ip()
    print("=" * 65)
    print("        EXCEL2PDF STUDIO - WEB APP (PC, HP & TABLET)")
    print("=" * 65)
    print(f"[*] Komputer ini      : http://localhost:8000")
    print(f"[*] Akses HP / Tablet : http://{ip}:8000")
    print("=" * 65)
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=False)
