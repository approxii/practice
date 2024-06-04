from io import BytesIO
from typing import Any, Dict
from urllib.parse import quote

from fastapi import APIRouter, Depends, File, UploadFile
from fastapi.responses import Response

from core.api.router.powerpoint.depends import (get_analyze_service,
                                                get_generate_service)
from core.services.conveter import DataConverter
from core.services.powerpoint import PowerpointAnalyzeService as AnalyzeService
from core.services.powerpoint import \
    PowerpointGenerateService as GenerateService

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
