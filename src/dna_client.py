import os
import time
import logging
from typing import Optional, Dict, Any

import requests

logger = logging.getLogger(__name__)


class DNAClient:
    """Cisco DNA Center API client with token retrieval and simple caching.

    The client will try the following for authentication in order:
    - Use a token passed into the constructor or found in an environment variable
      (e.g. `EMEA_DNAC_TOKEN`).
    - If no token is available, request one from `/dna/system/api/v1/auth/token`
      using the username/password from env (`DNA_USERNAME` / `DNA_PASSWORD`).

    Tokens are cached in-memory for `token_ttl` seconds (default 25 minutes).
    """

    def __init__(
        self,
        base_url: Optional[str] = None,
        token: Optional[str] = None,
        token_env: Optional[str] = None,
        token_ttl: int = 25 * 60,
        verify_ssl: bool = True,
    ) -> None:
        self.base_url = self._normalize_base_url(base_url or os.getenv('DNA_BASE_URL'))
        if not self.base_url:
            raise ValueError('DNA_BASE_URL must be set in environment or passed to DNAClient')

        self.username = os.getenv('DNA_USERNAME')
        self.password = os.getenv('DNA_PASSWORD')
        # token can come from explicit arg or from a named env var
        self._provided_token = token or (os.getenv(token_env) if token_env else None)
        self._token: Optional[str] = None
        self._token_acquired_at: float = 0
        self._token_ttl = token_ttl
        self.verify_ssl = verify_ssl

    @staticmethod
    def _normalize_base_url(base_url: Optional[str]) -> Optional[str]:
        if not base_url:
            return base_url
        u = base_url.strip().rstrip('/')
        # Users often paste UI URLs (..../dna/home). API root should be host root.
        if u.endswith('/dna/home'):
            u = u[:-9]
        elif u.endswith('/dna'):
            u = u[:-4]
        return u

    def _token_is_valid(self) -> bool:
        if not self._token:
            return False
        return (time.time() - self._token_acquired_at) < self._token_ttl

    @staticmethod
    def _is_placeholder_token(token: Optional[str]) -> bool:
        if not token:
            return True
        t = token.strip().lower()
        return t.startswith('your_') or t in {'changeme', 'replace_me', 'token'}

    def _fetch_token(self) -> Optional[str]:
        url = f"{self.base_url.rstrip('/')}/dna/system/api/v1/auth/token"
        logger.debug('Requesting DNAC token from %s', url)
        try:
            r = requests.post(url, auth=(self.username, self.password), verify=self.verify_ssl, timeout=10)
            r.raise_for_status()
            data = r.json()
            token = data.get('Token') or data.get('token')
            if not token:
                token = r.headers.get('X-Auth-Token')
            return token
        except Exception:
            logger.exception('Failed to fetch token')
            return None

    def get_token(self) -> Optional[str]:
        # If a provided token is available, use it and treat as newly acquired
        if self._provided_token and not self._token and not self._is_placeholder_token(self._provided_token):
            self._token = self._provided_token
            self._token_acquired_at = time.time()
            return self._token

        if self._token_is_valid():
            return self._token

        token = self._fetch_token()
        if token:
            self._token = token
            self._token_acquired_at = time.time()
            logger.info('Obtained DNAC token for %s', self.base_url)
            return token

        logger.warning('Could not obtain DNAC token for %s', self.base_url)
        return None

    def _headers(self) -> Dict[str, str]:
        token = self.get_token()
        headers = {'Accept': 'application/json'}
        if token:
            headers['X-Auth-Token'] = token
        return headers

    def _request(self, method: str, path: str, **kwargs) -> Any:
        url = path if path.startswith('http') else f"{self.base_url}/{path.lstrip('/')}"
        headers = kwargs.pop('headers', {})
        headers.update(self._headers())
        try:
            r = requests.request(method, url, headers=headers, verify=self.verify_ssl, timeout=30, **kwargs)
            if r.status_code == 401:
                # try refreshing token once
                logger.info('Received 401, refreshing token and retrying')
                self._token = None
                self._provided_token = None
                headers.update(self._headers())
                r = requests.request(method, url, headers=headers, verify=self.verify_ssl, timeout=30, **kwargs)
            r.raise_for_status()
            if r.text:
                return r.json()
            return None
        except requests.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else None
            if status == 404:
                logger.debug('HTTP 404 for %s %s', method, url)
            else:
                logger.error('HTTP error for %s %s: %s', method, url, exc)
            raise
        except Exception:
            logger.exception('Request failed for %s %s', method, url)
            raise

    def get(self, path: str, params: Optional[Dict] = None) -> Any:
        return self._request('GET', path, params=params)

    def post(self, path: str, json: Optional[Dict] = None) -> Any:
        return self._request('POST', path, json=json)

    # Convenience helpers
    def list_devices(self) -> Any:
        """Return device inventory list from DNAC"""
        return self.get('/dna/intent/api/v1/network-device')

    def list_images(self) -> Any:
        """Return images in DNAC image repository (import listing)"""
        return self.get('/dna/intent/api/v1/image/import')

    def get_device_by_id(self, device_id: str) -> Any:
        return self.get(f'/dna/intent/api/v1/network-device/{device_id}')


__all__ = ['DNAClient']
