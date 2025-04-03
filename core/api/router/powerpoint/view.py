from io import BytesIO
from pathlib import Path
from typing import Any, Dict
from urllib.parse import quote
import subprocess

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
    Принимает документ и словарь, возвращает новый документ для скачивания.
    """
    contents = await file.read()
    service.load(BytesIO(contents))
    service.update(dictionary)
    new_file = service.save_to_bytes()
    filename = quote(file.filename)
    headers = {"Content-Disposition": f"attachment; filename*=utf-8''{filename}"}
    media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)


@router.post("/update/preview/")
async def powerpoint_generate_preview(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: GenerateService = Depends(get_generate_service),
):
    """
    Принимает документ и словарь, обновляет его, конвертирует в PDF и возвращает HTML-страницу для просмотра.
    """
    contents = await file.read()
    
    # Обновляем документ
    service.load(BytesIO(contents))
    service.update(dictionary)
    new_file = service.save_to_bytes()
    filename = quote(file.filename)
    
    # Сохраняем обновленный документ во временную директорию 
    tmp_dir = "/tmp"
    updated_pptx_path = Path(tmp_dir) / f"updated_{filename}"
    with open(updated_pptx_path, "wb") as f:
        f.write(new_file.getvalue())
    
    # Конвертируем PPTX в PDF с помощью LibreOffice 
    cmd = [
        "soffice", "--headless", "--convert-to", "pdf",
        "--outdir", tmp_dir, str(updated_pptx_path)
    ]
    subprocess.run(cmd, check=True)
    
    # Формируем имя PDF-файла
    pdf_filename = f"updated_{Path(filename).stem}.pdf"
    # Формируем URL для предпросмотра
    view_url = f"/powerpoint/preview/pdf/{pdf_filename}"
    
    html_content = f"""
    <html>
      <head><title>Предпросмотр презентации</title></head>
      <body>
         <h1>Предпросмотр презентации</h1>
         <p><a href="{view_url}" target="_blank"></a></p>
      </body>
    </html>
    """
    return HTMLResponse(content=html_content)
