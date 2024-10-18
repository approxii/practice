from io import BytesIO
from typing import Any, Dict
from urllib.parse import quote

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from fastapi.responses import Response

from core.api.router.excel.depends import get_service
from core.services.conveter import DataConverter
from core.services.excel import ExcelService as Service

router = APIRouter(prefix="/excel")
router.tags = ["excel"]


@router.post("/update/")
async def excel_generate(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: Service = Depends(get_service),
):
    """
    Принимает таблицу и словарь в теле запроса, возвращает новую таблицу. (Обновляет по закладкам)
    """
    contents = await file.read()
    service.load(BytesIO(contents))
    service.update(dictionary)
    new_file = service.save_to_bytes()
    filename = quote(file.filename)

    headers = {
        "Content-Disposition": f"attachment; filename*=utf-8''{filename}",
    }
    media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)


@router.post("/get_as_json/")
async def excel_as_json(
    file: UploadFile = File(...),
    sheet_name: str = Form(None),
    range: str = Form(None),
    service: Service = Depends(get_service),
):
    """
    Получение значений таблицы в JSON.
    """
    try:
        contents = await file.read()
        service.load(BytesIO(contents))
        return service.to_json(sheet_name=sheet_name, range=range)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.post("/update_from_json/")
async def excel_generate_from_json(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: Service = Depends(get_service),
):
    """
    Принимает таблицу и словарь в теле запроса, возвращает новую таблицу. (Обновляет по ячейкам)
    """

    contents = await file.read()
    try:
        service.load(BytesIO(contents))
        service.from_json(dictionary)
        new_file = service.save_to_bytes()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    filename = quote(file.filename)

    headers = {
        "Content-Disposition": f"attachment; filename*=utf-8''{filename}",
    }
    media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)


@router.post("/update_with_blocks/")
async def excel_update_with_blocks(
    file: UploadFile = File(...),
    dictionary: Dict[str, Any] = Depends(DataConverter()),
    service: Service = Depends(get_service),
):
    """
    Принимает таблицу и словарь в теле запроса, возвращает новую таблицу. (Обновляет по ячейкам)
    """

    contents = await file.read()
    try:
        service.load(BytesIO(contents))
        service.update_with_blocks(dictionary)
        new_file = service.save_to_bytes()
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))
    filename = quote(file.filename)

    headers = {
        "Content-Disposition": f"attachment; filename*=utf-8''{filename}",
    }
    media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    return Response(content=new_file.getvalue(), headers=headers, media_type=media_type)
