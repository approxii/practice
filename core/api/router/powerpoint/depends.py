from core.services.powerpoint import parser as Service

async def get_service() -> Service:
    return Service()
