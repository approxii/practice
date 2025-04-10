import subprocess
from io import BytesIO
from pathlib import Path
from typing import Any, Dict
from urllib.parse import quote

from fastapi import APIRouter, Depends, File, UploadFile
from fastapi.responses import Response, HTMLResponse, FileResponse

from core.api.router.powerpoint.depends import get_analyze_service, get_generate_service
from core.services.conveter import DataConverter
from core.services.powerpoint import PowerpointAnalyzeService as AnalyzeService
from core.services.powerpoint import PowerpointGenerateService as GenerateService

router = APIRouter(prefix="/powerpoint")


@router.post("/analyze/")
async def powerpoint_analyze(
    file: UploadFile = File(...),
    service: AnalyzeService = Depends(get_analyze_service),
):
    contents = await file.read()
    service.load(BytesIO(contents))
    res = service.analyze()
    return res


@router.post("/update/")
async def powerpoint_generate(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: GenerateService = Depends(get_generate_service),
):
    """
    Принимает документ и словарь в теле запроса, возвращает новый документ.
    """
    contents = await file.read()
    service.load(BytesIO(contents))
    service.update(dictionary)
    new_file = service.save_to_bytes()
    filename = quote(file.filename)
    headers = {
        "Content-Disposition": f"attachment; filename*=utf-8''{filename}",
    }
    media_type = (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)


@router.post("/update/preview/")
async def powerpoint_generate_preview(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: GenerateService = Depends(get_generate_service),
):
    """
    Обновляет презентацию, сохраняет PDF и возвращает HTML-страницу с редиректом на предпросмотр.
    """
    contents = await file.read()
    service.load(BytesIO(contents))
    service.update(dictionary)
    new_file = service.save_to_bytes()
    filename = quote(file.filename)

    tmp_dir = Path("/tmp")
    updated_pptx_path = tmp_dir / f"updated_{filename}"
    with open(updated_pptx_path, "wb") as f:
        f.write(new_file.getvalue())

    subprocess.run([
        "soffice", "--headless", "--convert-to", "pdf",
        "--outdir", str(tmp_dir), str(updated_pptx_path)
    ], check=True)

    pdf_filename = f"updated_{Path(filename).stem}.pdf"
    preview_url = f"/powerpoint/preview/pdf/{quote(pdf_filename)}"

    # HTML-редирект 
    return HTMLResponse(content=f"""
    <html>
      <head>
        <meta http-equiv="refresh" content="0; url={preview_url}">
      </head>
    </html>
    """)

@router.get("/preview/pdf/{filename}", response_class=FileResponse)
async def serve_pdf_preview(filename: str):
    file_path = Path("/tmp") / filename
    if not file_path.exists():
        return Response(content="PDF not found", status_code=404)
    return FileResponse(path=file_path, media_type="application/pdf")
