import logging
import time
from typing import Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


class RequestManager:
    def __init__(self, rate_limit: float = 0.5):
        self.rate_limit = rate_limit
        self.last_request_time = 0
        self.session = requests.Session()
        self.logger = logging.getLogger(__name__)

        # Configure retry strategy
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)

    def get(self, url: str) -> Optional[requests.Response]:
        """Make a rate-limited GET request"""
        # Implement rate limiting
        elapsed = time.time() - self.last_request_time
        if elapsed < self.rate_limit:
            time.sleep(self.rate_limit - elapsed)

        try:
            response = self.session.get(url)
            response.raise_for_status()
            self.last_request_time = time.time()
            return response
        except Exception as e:
            self.logger.warning(f"Request failed: {str(e)}")
            return None
