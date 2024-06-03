from io import BytesIO
from fastapi import FastAPI
from fastapi import Depends, File, UploadFile
from core.services.updater import parser as Service
from core.api.router.powerpoint.depends import get_service

app: FastAPI = FastAPI(title="Microsoft documents generate/analyze")

@app.post("/update/")
async def upload_file_and_dict(
    file: UploadFile = File(...),
    service: Service = Depends(get_service),
):
    contents = await file.read()
    service.load(BytesIO(contents))
    res = service.analyze()

    return res