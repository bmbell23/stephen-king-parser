from pydantic import BaseSettings, HttpUrl
from typing import List, Optional
from pathlib import Path

class Settings(BaseSettings):
    base_url: HttpUrl = "https://www.stephenking.com"
    max_workers: int = 10
    rate_limit: float = 0.5
    cache_duration: int = 86400
    output_formats: List[str] = ["csv", "html"]
    cache_dir: Path = Path(".cache")
    log_level: str = "INFO"
    timeout: int = 30
    user_agent: str = "StephenKingParser/2.0"

    class Config:
        env_prefix = "KING_PARSER_"
        env_file = ".env"

def load_config(config_path: Optional[str] = None) -> Settings:
    """Load configuration from file and/or environment variables"""
    settings = Settings()

    if config_path:
        settings = Settings.parse_file(config_path)

    return settings