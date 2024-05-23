from fastapi import APIRouter, Depends

from core.api.router.word.view import router as excel_router

router = APIRouter()

router.include_router(excel_router)
