from io import BytesIO
from typing import Any, Dict

from fastapi import APIRouter, Body, Depends, File, UploadFile
from fastapi.responses import FileResponse, Response, StreamingResponse

from core.api.router.word.depends import get_service
from core.services.conveter import DataConverter
from core.services.word import WordService as Service

router = APIRouter(prefix="/word")


@router.post("/update/")
async def upload_file_and_dict(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: Service = Depends(get_service),
):
    """
    Принимает таюлицу и словарь в теле запроса, возвращает новую таблицу.
    """
    contents = await file.read()
    service.load(BytesIO(contents))
    service.update(dictionary)
    new_file = service.save_to_bytes()
    service.save_to_file('updated_document.docx')
    headers = {
        "Content-Disposition": f"attachment; filename={file.filename}",
    }
    media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)
