from dotenv import load_dotenv
from pydantic_settings import BaseSettings

load_dotenv()


class Settings(BaseSettings):
    class Config:
        case_sensitive = True

    APP_HOST: str = "0.0.0.0"
    APP_PORT: int = 8000
    AUTH_HOST: str = "https://vega.mirea.ru/"
    AUTH_URL: str = AUTH_HOST + "authservice.php?op=parsetoken&token="
    ROOT_PATH: str = "/officesvc"


settings = Settings()
