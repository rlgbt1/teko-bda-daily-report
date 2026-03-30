"""
scrapers/base_scraper.py — Shared HTTP + caching base for all scrapers.
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Any

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from src.config import REQUEST_TIMEOUT, MAX_RETRIES
from src.utils.logger import get_logger

logger = get_logger(__name__)


@dataclass
class ScrapeResult:
    data:    dict[str, Any] = field(default_factory=dict)
    success: bool = True
    error:   str | None = None


class BaseScraper:
    """
    Base class for all scrapers.

    Provides:
    - requests.Session with retry logic
    - Simple in-memory cache (TTL in seconds)
    - fetch() wrapper with cache support
    - get() convenience method
    """

    CACHE_TTL: int = 3600  # override per scraper

    def __init__(self, source_name: str) -> None:
        self.source_name = source_name
        self._cache: dict[str, tuple[float, ScrapeResult]] = {}
        self._session = self._build_session()

    # ── Session ───────────────────────────────────────────────────────

    def _build_session(self) -> requests.Session:
        session = requests.Session()
        retry = Retry(
            total=MAX_RETRIES,
            backoff_factor=0.5,
            status_forcelist=[429, 500, 502, 503, 504],
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("https://", adapter)
        session.mount("http://",  adapter)
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (compatible; BDA-DailyReport/1.0)",
        })
        return session

    def get(self, url: str, **kwargs) -> requests.Response:
        kwargs.setdefault("timeout", REQUEST_TIMEOUT)
        return self._session.get(url, **kwargs)

    # ── Cache ─────────────────────────────────────────────────────────

    def _is_cached(self, key: str) -> bool:
        if key not in self._cache:
            return False
        ts, _ = self._cache[key]
        return (time.time() - ts) < self.CACHE_TTL

    # ── Public fetch ──────────────────────────────────────────────────

    def fetch(self, cache_key: str = "") -> ScrapeResult:
        key = cache_key or self.source_name
        if self._is_cached(key):
            logger.debug("%s: returning cached result", self.source_name)
            return self._cache[key][1]

        try:
            data   = self._fetch()
            result = ScrapeResult(data=data, success=True)
        except Exception as exc:
            logger.error("%s fetch failed: %s", self.source_name, exc)
            result = ScrapeResult(data={}, success=False, error=str(exc))

        self._cache[key] = (time.time(), result)
        return result

    # ── Override in subclass ──────────────────────────────────────────

    def _fetch(self) -> dict[str, Any]:
        raise NotImplementedError