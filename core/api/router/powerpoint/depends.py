from core.services.updater import parser as Service

async def get_service() -> Service:
    return Service()
