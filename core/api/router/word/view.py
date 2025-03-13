from io import BytesIO
from typing import Any, Dict
from urllib.parse import quote

from fastapi import APIRouter, Depends, File, UploadFile, HTTPException
from fastapi.responses import Response
from fastapi.responses import StreamingResponse

from core.api.router.word.depends import get_service
from core.services.conveter import DataConverter
from core.services.word import WordService as Service

from pydantic import BaseModel

router = APIRouter(prefix="/word")
router.tags = ["word"]

class ExampleDictionary(BaseModel):
    key1: str = "value1"
    key2: str = "value2"

    class Config:
        schema_extra = {
            "example": {
                "key1": "example_value_1",
                "key2": "example_value_2",
            }
        }

@router.post("/generate/")
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
    """
        Принимает документ и словарь в теле запроса, возвращает новый документ(очищает нужный текст, указанный в словаре, и ставит туда закладку)
        """
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
    """
        Принимает документ в теле запроса, возвращает словарь в Response body
        """
    try:
        contents = await file.read()
        service.load(BytesIO(contents))
        return service.extract_bookmarks()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    
@router.post("/get_bookmarks_with_formatting")
async def get_bookmarks_with_formatting(
        file: UploadFile = File(...),
        service: Service = Depends(get_service)
):
    """
        Принимает документ в теле запроса, возвращает словарь в Response body
        """
    try:
        contents = await file.read()
        service.load(BytesIO(contents))
        return service.parse_with_formatting()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    
@router.post("/generate_with_formatting")
async def generate_with_formatting(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: Service = Depends(get_service),
):
    """
    Принимает документ и словарь в теле запроса, возвращает новый документ.
    """
    contents = await file.read()
    service.load(BytesIO(contents))
    service.update_with_formatting(dictionary)
    new_file = service.save_to_bytes()
    filename = quote(file.filename)

    headers = {
        "Content-Disposition": f"attachment; filename*=utf-8''{filename}",
    }
    media_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)

@router.post("/md_to_docx")
async def convert_md_to_docx(
        file: UploadFile = File(...),
        service: Service = Depends(get_service)
):
    """
    Принимает md в теле запроса, возвращает docx
    """
    try:
        # Чтение содержимого файла
        contents = await file.read()

        # Конвертация Markdown в DOCX
        service.md_to_word(contents)  # передаем содержимое файла как байты или строку

        # Сохранение документа в память
        docx_stream = BytesIO()
        service.docx_file.save(docx_stream)  # сохраним документ в поток
        docx_stream.seek(0)

        # Возвращаем DOCX файл в ответе
        return StreamingResponse(docx_stream,
                                 media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                 headers={"Content-Disposition": "attachment; filename=converted.docx"})
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))