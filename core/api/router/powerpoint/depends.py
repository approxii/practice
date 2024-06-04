from core.services.powerpoint import PowerpointAnalyzeService as AnalyzeService
from core.services.powerpoint import \
    PowerpointGenerateService as GenerateService


async def get_analyze_service() -> AnalyzeService:
    return AnalyzeService()


async def get_generate_service() -> GenerateService:
    return GenerateService()
