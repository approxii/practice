from core.services.excel import ExcelService as Service


async def get_service() -> Service:
    return Service()
