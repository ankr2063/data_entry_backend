from pydantic_settings import BaseSettings
from typing import Optional

class Settings(BaseSettings):
    database_url: str
    secret_key: str
    algorithm: str = "HS256"
    access_token_expire_minutes: int = 30
    
    # Microsoft Graph API
    microsoft_client_id: str
    microsoft_client_secret: str
    microsoft_tenant_id: str
    
    class Config:
        env_file = ".env"

settings = Settings()