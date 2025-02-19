import requests
from time import time, sleep
from typing import Optional
import logging
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

logger = logging.getLogger(__name__)

class RequestManager:
    def __init__(self, rate_limit: float = 0.2):
        self.rate_limit = rate_limit
        self.last_request = 0
        self._setup_session()

    def _setup_session(self):
        """Configure session with optimized settings"""
        self.session = requests.Session()

        # Configure retries
        retry_strategy = Retry(
            total=3,
            backoff_factor=0.1,
            status_forcelist=[429, 500, 502, 503, 504],
        )

        # Configure connection pooling
        adapter = HTTPAdapter(
            max_retries=retry_strategy,
            pool_connections=10,
            pool_maxsize=10
        )

        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

        # Optimize headers
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,application/xhtml+xml',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive'
        })

    def get(self, url: str) -> Optional[requests.Response]:
        """Make a GET request with optimized handling"""
        current_time = time()
        sleep_time = self.rate_limit - (current_time - self.last_request)

        if sleep_time > 0:
            sleep(sleep_time)

        try:
            response = self.session.get(url, timeout=5)  # Reduced timeout
            response.raise_for_status()
            self.last_request = time()
            return response
        except requests.exceptions.RequestException as e:
            logger.error(f"Request failed for {url}: {str(e)[:100]}")
            return None
