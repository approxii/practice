from fastapi import APIRouter, Depends

from core.api.sso import get_auth
from core.api.router.excel.view import router as excel_router


router = APIRouter(dependencies=[Depends(get_auth)])

router.include_router(excel_router)
