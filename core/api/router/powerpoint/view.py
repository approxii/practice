from io import BytesIO

from fastapi import APIRouter, Depends, File, UploadFile

from core.api.router.powerpoint.depends import get_service
from core.services.powerpoint import PowerPointService as Service

router = APIRouter(prefix="/powerpoint")


@router.post("/analyze/")
async def word_analyze(
    file: UploadFile = File(...),
    service: Service = Depends(get_service),
):
    contents = await file.read()
    service.load(BytesIO(contents))
    res = service.analyze()

    return res
