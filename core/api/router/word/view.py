from io import BytesIO
from typing import Any, Dict
from urllib.parse import quote

from fastapi import APIRouter, Depends, File, UploadFile, HTTPException
from fastapi.responses import Response

from core.api.router.word.depends import get_service
from core.services.conveter import DataConverter
from core.services.word import WordService as Service

router = APIRouter(prefix="/word")


@router.post("/word_generate/")
async def word_generate(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: Service = Depends(get_service),
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
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)

@router.post("/add bookmarks instead of paragraph")
async def word_clean_para_with_bookmarks(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: Service = Depends(get_service),
):

    contents = await file.read()
    service.load(BytesIO(contents))
    service.clean_para_with_bookmark(dictionary)
    new_file = service.save_to_bytes()
    filename = quote(file.filename)

    headers = {
        "Content-Disposition": f"attachment; filename*=utf-8''{filename}",
    }
    media_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)

@router.post("/get_bookmarks")
async def get_bookmarks(
        file: UploadFile = File(...),
        service: Service = Depends(get_service)
):
    try:
        contents = await file.read()
        service.load(BytesIO(contents))
        return service.extract_bookmarks()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))