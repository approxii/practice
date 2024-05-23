import uvicorn
from fastapi import FastAPI

from core.api import router
from core.settings.app_config import settings

app: FastAPI = FastAPI(title="Word Generate")


app.include_router(router)

# 1 практика (07.03) написал класс для чтения, обновления и сохранения xlxs файлов

# {"#Лист1!A1":"Тут должно быть ааваа","#Лист1!F8":"а тут не должно"}

if __name__ == "__main__":
    uvicorn.run(app, host=settings.APP_HOST, port=settings.APP_PORT)
