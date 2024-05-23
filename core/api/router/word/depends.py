from core.services.word import WordService as Service


async def get_service() -> Service:
    return Service()
