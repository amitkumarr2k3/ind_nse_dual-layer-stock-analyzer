# This file shows the session setup to add after line 94 in Stock_Agent.py

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Create a session with proper connection pooling and retry strategy for screener.in
# This fixes timeout/SSL errors by:
# 1. Reusing TCP connections (pooling)
# 2. Increasing pool size for concurrent requests
# 3. Adding connection keep-alive
# 4. Using urllib3's built-in retry with backoff (faster than manual sleep)

SESSION = requests.Session()

# Configure connection pool: max 20 per host, block when full
adapter = HTTPAdapter(
    pool_connections=10,
    pool_maxsize=20,
    max_retries=Retry(
        total=2,                              # Total retries
        connect=2,                            # Retry on connection errors
        read=2,                               # Retry on read timeout
        backoff_factor=0.5,                   # 0.5s, 1s backoff (much faster than manual 1s, 2s, 4s)
        status_forcelist=[500, 502, 503, 504],  # Retry on server errors (not client errors)
        raise_on_status=False                 # Don't raise exception, return the response
    )
)

SESSION.mount('http://', adapter)
SESSION.mount('https://', adapter)
SESSION.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})

# Then replace:
# requests.get(url, headers=..., timeout=X)
# with:
# SESSION.get(url, timeout=X)
