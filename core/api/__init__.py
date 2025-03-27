from fastapi import APIRouter, Depends

from core.api.router.excel.view import router as excel_router
from core.api.router.powerpoint.view import router as powerpoint_router
from core.api.router.word.view import router as word_router
from core.api.sso import get_auth

router = APIRouter()#dependencies=[Depends(get_auth)])

router.include_router(excel_router)
router.include_router(word_router)
router.include_router(powerpoint_router)
