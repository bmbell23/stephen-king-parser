from pathlib import Path

from pydantic import BaseSettings, HttpUrl


class Settings(BaseSettings):
    base_url: HttpUrl = "https://www.stephenking.com"
    rate_limit: float = 0.5
    cache_dir: Path = Path(".cache")
    log_level: str = "INFO"
    timeout: int = 30


def load_config(config_path: Optional[str] = None) -> Settings:
    """Load configuration from file and/or environment variables"""
    settings = Settings()

    if config_path:
        settings = Settings.parse_file(config_path)

    return settings
