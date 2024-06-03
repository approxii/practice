from io import BytesIO
from typing import Any, Dict
from urllib.parse import quote

from fastapi import APIRouter, Depends, File, UploadFile
from fastapi.responses import Response

from core.api.router.powerpoint.depends import get_service
from core.services.conveter import DataConverter
from core.services.powerpoint import PowerpointService as Service

router = APIRouter(prefix="/powerpoint")


@router.post("/update/")
async def upload_file_and_dict(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: Service = Depends(get_service),
):
    """
    Принимает презентацию и словарь в теле запроса, возвращает новую презентацию.
    """
    contents = await file.read()
    service.load(BytesIO(contents))
    service.update(dictionary)
    new_file = service.save_to_bytes()
    filename = quote(file.filename)

    headers = {
        "Content-Disposition": f"attachment; filename*=utf-8''{filename}",
    }
    media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)
