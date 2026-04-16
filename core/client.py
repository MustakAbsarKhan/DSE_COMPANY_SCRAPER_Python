import requests
from requests.adapters import HTTPAdapter
from urllib3 import Retry

session = requests.Session()

retry = Retry(
    total=5,
    backoff_factor=1,
    status_forcelist=[429, 500, 502, 503, 504]
)

adapter = HTTPAdapter(max_retries=retry)
session.mount("http://", adapter)
session.mount("https://", adapter)


def get(url):
    response = session.get(url, timeout=10)
    response.raise_for_status()
    return response.text