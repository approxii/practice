from fastapi import APIRouter, Depends

from core.api.sso import get_auth

router = APIRouter(dependencies=[Depends(get_auth)])
