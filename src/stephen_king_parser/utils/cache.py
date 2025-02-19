import hashlib
import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Optional

import aiofiles


class Cache:
    def __init__(self, cache_dir: str = ".cache", ttl: int = 86400):
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.ttl = ttl

    def _get_cache_path(self, key: str) -> Path:
        """Generate cache file path from key"""
        hashed_key = hashlib.md5(key.encode()).hexdigest()
        return self.cache_dir / f"{hashed_key}.json"

    async def get(self, key: str) -> Optional[Any]:
        """Get value from cache if not expired"""
        cache_path = self._get_cache_path(key)
        if not cache_path.exists():
            return None

        try:
            async with aiofiles.open(cache_path, "r") as f:
                data = json.loads(await f.read())
                if (
                    datetime.fromisoformat(data["timestamp"])
                    + timedelta(seconds=self.ttl)
                    > datetime.now()
                ):
                    return data["content"]
        except (json.JSONDecodeError, KeyError, ValueError):
            await self.delete(key)
        return None

    async def set(self, key: str, value: Any) -> None:
        """Set value in cache with timestamp"""
        cache_path = self._get_cache_path(key)
        data = {"timestamp": datetime.now().isoformat(), "content": value}
        async with aiofiles.open(cache_path, "w") as f:
            await f.write(json.dumps(data))

    async def delete(self, key: str) -> None:
        """Delete key from cache"""
        cache_path = self._get_cache_path(key)
        try:
            cache_path.unlink(missing_ok=True)
        except OSError:
            pass

    async def clear(self) -> None:
        """Clear all cached data"""
        for cache_file in self.cache_dir.glob("*.json"):
            try:
                cache_file.unlink()
            except OSError:
                pass
