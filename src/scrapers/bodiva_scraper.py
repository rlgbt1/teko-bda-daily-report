"""
scrapers/bodiva_scraper.py — Scraper for BODIVA (Bolsa de Derivados de Angola).

Playwright is used for JS rendering. Falls back to requests if unavailable.
"""
from __future__ import annotations

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

from typing import Any
from bs4 import BeautifulSoup

from src.config import URLs
from src.scrapers.base_scraper import BaseScraper
from src.utils.logger import get_logger
from src.utils.helpers import safe_float, compute_variation

logger = get_logger(__name__)

BODIVA_STOCKS = {
    "BAIAAAAA": "BAI",
    "BFAAAAAA": "BFA",
    "BCGAAAAA": "BCGA",
    "ENSAAAAA": "ENSA",
    "BDVAAAAA": "BDV",
}


class BODIVAScraper(BaseScraper):

    CACHE_TTL = 86400  # 24 hours

    def __init__(self) -> None:
        super().__init__(source_name="BODIVA")
        self._playwright_available = self._check_playwright()

    def _check_playwright(self) -> bool:
        try:
            import playwright  # noqa: F401
            return True
        except ImportError: 
            logger.warning(
                "Playwright not installed — BODIVA scraper will use requests fallback. "
                "Run: pip install playwright && python -m playwright install chromium"
            )
            return False

    def _fetch(self) -> dict[str, Any]:
        if self._playwright_available:
            return self._fetch_with_playwright()
        return self._fetch_with_requests()

    # ── Playwright ────────────────────────────────────────────────────

    def _fetch_with_playwright(self) -> dict[str, Any]:
        from playwright.sync_api import sync_playwright

        result: dict[str, Any] = {"stocks": {}, "segments": {}}

        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            page    = browser.new_page()
            try:
                page.goto(URLs.BODIVA_MARKET, timeout=30_000)
                page.wait_for_load_state("networkidle", timeout=20_000)
                result.update(self._parse_html(page.content()))
            except Exception as exc:
                logger.error("Playwright BODIVA scrape failed: %s", exc)
            finally:
                browser.close()

        return result

    # ── Requests fallback ─────────────────────────────────────────────

    def _fetch_with_requests(self) -> dict[str, Any]:
        try:
            response = self.get(URLs.BODIVA_HOME)
            response.raise_for_status()
            return self._parse_html(response.text)
        except Exception as exc:
            logger.error("BODIVA requests fallback failed: %s", exc)
            return {"stocks": {}, "segments": {}}

    # ── Parser ────────────────────────────────────────────────────────

    def _parse_html(self, html: str) -> dict[str, Any]:
        """
        BODIVA's /mercado page is a Next.js app that renders stock data as a
        horizontal marquee ticker (no <table> elements).

        Each stock appears in a  <div class="rfm-child">  block with:
          - <h1> stock code        e.g. "BFAAAAAA"
          - <h1> current price     e.g. "108500.00" or "108500.00 %"  (strip the %)
          - <span> variation %     e.g. "-1.36000" or "5.38000 %"
          - <svg class="text-green-600|text-red-600">  (direction indicator)

        Market segment totals (Obrigações, Repos, etc.) are NOT present in the
        rendered DOM — they exist only in the server-push JSON blobs.
        We return an empty segments dict and log a warning.
        """
        soup   = BeautifulSoup(html, "html.parser")
        stocks: dict[str, dict] = {}

        children = soup.find_all("div", class_="rfm-child")
        for child in children:
            h1_texts = [h.get_text(strip=True) for h in child.find_all("h1")]
            if not h1_texts:
                continue

            # First h1 that matches a known code is the stock identifier
            code = None
            for h in h1_texts:
                candidate = h.strip().upper()
                if candidate in BODIVA_STOCKS:
                    code = candidate
                    break
            if code is None:
                continue

            # Price: the h1 that is NOT the code (strip stray "%" characters)
            price_str = next(
                (h for h in h1_texts if h.strip().upper() != code), ""
            )
            current = safe_float(price_str.replace("%", "").strip())

            # Variation: first non-empty <span> text (strip "%" too)
            spans = [s.get_text(strip=True) for s in child.find_all("span")
                     if s.get_text(strip=True)]
            var_str = spans[0].replace("%", "").strip() if spans else "0"
            change_pct = safe_float(var_str)

            # Sign: SVG colour class confirms direction (guards against missing sign)
            svg = child.find("svg")
            if svg and change_pct is not None:
                classes = " ".join(svg.get("class", []))
                if "red" in classes and change_pct > 0:
                    change_pct = -change_pct
                elif "green" in classes and change_pct < 0:
                    change_pct = abs(change_pct)

            stocks[code] = {
                "name":       BODIVA_STOCKS[code],
                "volume":     None,      # not available in the ticker widget
                "previous":   None,      # not available in the ticker widget
                "current":    current,
                "change_pct": change_pct,
            }

        if not stocks:
            logger.warning(
                "BODIVA: 0 stocks parsed — site structure may have changed "
                "or page did not fully render. Slides will show '—' placeholders."
            )
        else:
            logger.info("BODIVA: %d stocks parsed from ticker", len(stocks))

        # Segment totals are embedded in Next.js server-push JSON, not in the DOM.
        logger.debug("BODIVA: segment totals not available from rendered HTML.")

        return {"stocks": stocks, "segments": {}}

    # ── Public accessors ──────────────────────────────────────────────

    def get_stocks(self) -> dict[str, dict]:
        return self.fetch(cache_key="bodiva_stocks").data.get("stocks", {})

    def get_segments(self) -> dict[str, float | None]:
        return self.fetch(cache_key="bodiva_stocks").data.get("segments", {})