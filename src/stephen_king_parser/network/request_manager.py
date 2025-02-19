import requests
import threading
import time
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

class RequestManager:
    """Manages HTTP requests with rate limiting and connection pooling"""
    def __init__(self, rate_limit: float = 0.5):
        self.rate_limit = rate_limit
        self.last_request_time = 0
        self.lock = threading.Lock()

        # Configure retry strategy
        retry_strategy = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504]
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)

        # Create session with retry strategy
        self.session = requests.Session()
        self.session.mount("http://", adapter)
        self.session.mount("https://", adapter)
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })

    def get(self, url: str) -> requests.Response:
        """Make a GET request with rate limiting"""
        with self.lock:
            current_time = time.time()
            time_since_last_request = current_time - self.last_request_time

            if time_since_last_request < self.rate_limit:
                time.sleep(self.rate_limit - time_since_last_request)

            try:
                response = self.session.get(url)
                self.last_request_time = time.time()
                return response
            except Exception as e:
                print(f"Request failed for {url}: {str(e)}")
                return None