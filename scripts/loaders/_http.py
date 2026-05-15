"""HTTP helper with retry/backoff for flaky public APIs (SICONFI ORDS pool throttles aggressively)."""
import time
import requests

DEFAULT_TIMEOUT = 60
DEFAULT_RETRIES = 6
DEFAULT_BACKOFF = 4  # seconds, doubles each retry

UA = "panorama-fiscal-df/1.0 (build script; contact via repo)"


def get_json(url, params=None, retries=DEFAULT_RETRIES, backoff=DEFAULT_BACKOFF, timeout=DEFAULT_TIMEOUT):
    """GET a URL and decode JSON. Retries on rate-limit (HTTP 429), 5xx, and SICONFI's TooManyRequests JSON body."""
    last_err = None
    delay = backoff
    for attempt in range(1, retries + 1):
        try:
            r = requests.get(url, params=params, timeout=timeout, headers={"User-Agent": UA, "Accept": "application/json"})
            if r.status_code in (429, 500, 502, 503, 504):
                last_err = f"HTTP {r.status_code}"
            else:
                try:
                    data = r.json()
                except ValueError:
                    last_err = f"non-JSON body (HTTP {r.status_code}): {r.text[:200]}"
                else:
                    if isinstance(data, dict) and data.get("code") == "TooManyRequests":
                        last_err = "SICONFI TooManyRequests"
                    else:
                        return data
        except requests.RequestException as e:
            last_err = f"{type(e).__name__}: {e}"

        if attempt < retries:
            print(f"  retry {attempt}/{retries} in {delay}s ({last_err})")
            time.sleep(delay)
            delay *= 2
    raise RuntimeError(f"GET failed after {retries} attempts: {url} — {last_err}")
