from core.services.powerpoint import PowerpointService as Service


async def get_service() -> Service:
    return Service()
