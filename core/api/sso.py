from datetime import datetime, timezone

import aiohttp
from dateutil import parser
from fastapi import Depends, HTTPException, Request
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from starlette.status import HTTP_401_UNAUTHORIZED, HTTP_403_FORBIDDEN

from core.settings.app_config import settings


class JWTBearer(HTTPBearer):
    async def __call__(self, request: Request):
        credentials: HTTPAuthorizationCredentials = await super().__call__(request)
        if not credentials or credentials.scheme != "Bearer":
            raise HTTPException(
                status_code=HTTP_401_UNAUTHORIZED,
                detail="Необходима авторизация.",
                headers={"WWW-Authenticate": "Bearer"},
            )
        return credentials.credentials


async def validate_token(token: str) -> bool:
    
    url = settings.AUTH_URL + token

    async with aiohttp.ClientSession() as session:
        async with session.get(url) as response:
            if response.status != 200:
                return False
            token_data = await response.json()
            exp = parser.parse(token_data["exp"])
            if exp < datetime.now(timezone.utc):
                raise HTTPException(
                    status_code=HTTP_401_UNAUTHORIZED,
                    detail="Токен истек.",
                )
            return True


async def get_auth(token: str = Depends(JWTBearer())):
    if not await validate_token(token):
        raise HTTPException(
            status_code=HTTP_401_UNAUTHORIZED,
            detail="Необходима авторизация.",
        )
    return True
