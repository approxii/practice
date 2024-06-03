import uvicorn
from fastapi import FastAPI

from core.api import router
from core.settings.app_config import settings

app: FastAPI = FastAPI(title="Micorosoft documets generate/analyze")


app.include_router(router)


if __name__ == "__main__":
    uvicorn.run(app, host=settings.APP_HOST, port=settings.APP_PORT)
