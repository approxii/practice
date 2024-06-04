from core.services.powerpoint import PowerPointService as Service


async def get_service() -> Service:
    return Service()
