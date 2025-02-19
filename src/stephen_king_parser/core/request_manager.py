import asyncio
import aiohttp
from typing import Optional, Dict, Any
from urllib3.util.retry import Retry
from aiohttp import ClientTimeout
from ..utils.cache import Cache
import logging

class RequestManager:
    def __init__(
        self,
        rate_limit: float = 0.5,
        max_retries: int = 3,
        timeout: int = 30,
        cache: Optional[Cache] = None
    ):
        self.rate_limit = rate_limit
        self.last_request_time = 0
        self.cache = cache or Cache()
        self.timeout = ClientTimeout(total=timeout)
        self.logger = logging.getLogger(__name__)
        self._semaphore = asyncio.Semaphore(10)  # Limit concurrent requests

        # Retry configuration
        self.retry_options = {
            "total": max_retries,
            "status_forcelist": [500, 502, 503, 504],
            "backoff_factor": 1
        }

    async def get(self, url: str) -> Optional[Dict[str, Any]]:
        """Fetch URL with caching and rate limiting"""
        if cached_data := await self.cache.get(url):
            return cached_data

        async with self._semaphore:
            try:
                async with aiohttp.ClientSession(timeout=self.timeout) as session:
                    async with session.get(url) as response:
                        await asyncio.sleep(self.rate_limit)
                        response.raise_for_status()
                        data = await response.text()
                        await self.cache.set(url, data)
                        return data
            except aiohttp.ClientError as e:
                self.logger.error(f"Request failed for {url}: {str(e)}")
                return None

    async def bulk_get(self, urls: list[str]) -> Dict[str, Any]:
        """Fetch multiple URLs concurrently"""
        tasks = [self.get(url) for url in urls]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        return {url: result for url, result in zip(urls, results)
                if not isinstance(result, Exception)}