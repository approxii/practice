from dotenv import load_dotenv
from pydantic_settings import BaseSettings

load_dotenv()


class Settings(BaseSettings):
    class Config:
        case_sensitive = True

    APP_HOST: str = "0.0.0.0"
    APP_PORT: int = 8000


settings = Settings()
