"""Scraper for BODIVA (Bolsa de Divida e Valores de Angola).

The visible market widgets on bodiva.ao are React components, but the useful
data is loaded from a public JSON endpoint. We only need the listed equities,
so the scraper reads the order-book table and filters it down to the BODIVA
stock codes used by the report.
"""
from __future__ import annotations

import os
import sys
from typing import Any

from bs4 import BeautifulSoup
import pandas as pd
import requests
import urllib3

from src.config import URLs
from src.scrapers.base_scraper import BaseScraper
from src.utils.helpers import safe_float
from src.utils.logger import get_logger

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

logger = get_logger(__name__)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BODIVA_API = "https://www.bodiva.ao/website/api"
BODIVA_ORDER_BOOK_URL = f"{BODIVA_API}/ListarLivroOrdensCPZ_no.php"

BODIVA_STOCKS = {
    "BDVAAAAA": "BODIVA",
    "BAIAAAAA": "BAI",
    "BFAAAAAA": "BFA",
    "BCGAAAAA": "BCGA",
    "ENSAAAAA": "ENSA",
}


class BODIVAScraper(BaseScraper):
    CACHE_TTL = 86400

    def __init__(self) -> None:
        super().__init__(source_name="BODIVA")
        self._playwright_available = self._check_playwright()

    def _check_playwright(self) -> bool:
        try:
            import playwright  # noqa: F401
            return True
        except ImportError:
            logger.warning(
                "Playwright not installed; BODIVA scraper will use API/requests only. "
                "Run: pip install playwright && python -m playwright install chromium"
            )
            return False

    def _fetch(self) -> dict[str, Any]:
        result = self._fetch_with_api()
        if result.get("stocks"):
            return result

        logger.warning("BODIVA API scrape returned no rows; trying rendered DOM fallback.")
        if self._playwright_available:
            return self._fetch_with_playwright()
        return self._fetch_with_requests()

    # Public JSON APIs

    def _fetch_json(self, url: str) -> Any:
        headers = {
            "Accept": "application/json,text/plain,*/*",
            "Referer": URLs.BODIVA_HOME,
        }
        try:
            response = self.get(url, headers=headers)
        except requests.exceptions.SSLError:
            logger.warning("BODIVA SSL verification failed for %s; retrying without verification.", url)
            response = self.get(url, headers=headers, verify=False)
        response.raise_for_status()
        return response.json()

    def _fetch_with_api(self) -> dict[str, Any]:
        result = self._empty_result()

        try:
            order_book = self._fetch_json(BODIVA_ORDER_BOOK_URL)
            if isinstance(order_book, list):
                parsed_order_book = self._parse_order_book(order_book)
                result["stocks"] = self._stocks_from_order_book(parsed_order_book)
                logger.info("BODIVA: %d order-book rows parsed from API", len(order_book))
        except Exception as exc:
            logger.error("BODIVA order-book API failed: %s", exc)

        return result

    # Rendered DOM fallback

    def _fetch_with_playwright(self) -> dict[str, Any]:
        from playwright.sync_api import sync_playwright

        result = self._empty_result()
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            page = browser.new_page(viewport={"width": 1600, "height": 1400})
            try:
                page.goto(URLs.BODIVA_HOME, timeout=60_000, wait_until="domcontentloaded")
                page.wait_for_timeout(6_000)
                result.update(self._parse_html(page.content()))
            except Exception as exc:
                logger.error("Playwright BODIVA scrape failed: %s", exc)
            finally:
                browser.close()
        return result

    def _fetch_with_requests(self) -> dict[str, Any]:
        try:
            response = self.get(URLs.BODIVA_HOME)
            response.raise_for_status()
            return self._parse_html(response.text)
        except Exception as exc:
            logger.error("BODIVA requests fallback failed: %s", exc)
            return self._empty_result()

    @staticmethod
    def _empty_result() -> dict[str, Any]:
        return {
            "stocks": {},
        }

    # API parsers

    def _parse_order_book(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        parsed = []
        for row in rows:
            code = str(row.get("CodigoNegociacao") or "").strip().upper()
            if not code:
                continue

            buy = row.get("CompraF") if isinstance(row.get("CompraF"), dict) else {}
            sell = row.get("VendaF") if isinstance(row.get("VendaF"), dict) else {}
            parsed.append({
                "code": code,
                "isin": row.get("Isin"),
                "typology": row.get("Tipologia"),
                "typology_ci": row.get("TipologiaCI"),
                "face_value": safe_float(row.get("faceValue")),
                "par_value": safe_float(row.get("parValue")),
                "premium_value": safe_float(row.get("premiumValue")),
                "coupon_rate": row.get("TaxaCupao"),
                "issue_date": row.get("DataEmissao"),
                "maturity_date": row.get("DataMaturidade"),
                "last_quote": safe_float(row.get("UltimaCotacao")),
                "best_bid_qty": safe_float(buy.get("Quantidade")),
                "best_bid_price": safe_float(buy.get("Preco")),
                "best_ask_qty": safe_float(sell.get("Quantidade")),
                "best_ask_price": safe_float(sell.get("Preco")),
            })
        return parsed

    def _stocks_from_order_book(self, rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
        stocks = {}
        for row in rows:
            code = row["code"]
            typology = str(row.get("typology") or "").lower()
            if "acç" not in typology and "acc" not in typology and code not in BODIVA_STOCKS:
                continue

            stocks[code] = {
                "name": BODIVA_STOCKS.get(code, code),
                "volume": row.get("face_value"),
                "previous": None,
                "current": row.get("par_value") or row.get("last_quote"),
                "change_pct": None,
                "last_quote": row.get("last_quote"),
                "best_bid_price": row.get("best_bid_price"),
                "best_ask_price": row.get("best_ask_price"),
            }
        return stocks

    # DOM fallback parser

    def _parse_html(self, html: str) -> dict[str, Any]:
        soup = BeautifulSoup(html, "html.parser")
        stocks: dict[str, dict[str, Any]] = {}

        for child in soup.find_all("div", class_="rfm-child"):
            h1_texts = [h.get_text(strip=True) for h in child.find_all("h1")]
            code = next((h.upper() for h in h1_texts if h.upper() in BODIVA_STOCKS), None)
            if not code:
                continue

            price_str = next((h for h in h1_texts if h.upper() != code), "")
            current = safe_float(price_str.replace("%", "").strip())
            spans = [s.get_text(strip=True) for s in child.find_all("span") if s.get_text(strip=True)]
            change_pct = safe_float(spans[0].replace("%", "").strip()) if spans else None

            svg = child.find("svg")
            if svg and change_pct is not None:
                classes = " ".join(svg.get("class", []))
                if "red" in classes and change_pct > 0:
                    change_pct = -change_pct
                elif "green" in classes and change_pct < 0:
                    change_pct = abs(change_pct)

            stocks[code] = {
                "name": BODIVA_STOCKS[code],
                "volume": None,
                "previous": None,
                "current": current,
                "change_pct": change_pct,
            }

        if not stocks:
            logger.warning("BODIVA: 0 stocks parsed from rendered DOM fallback.")
        return {"stocks": stocks}

    # Public accessors

    def get_stocks(self) -> dict[str, dict]:
        return self.fetch(cache_key="bodiva").data.get("stocks", {})


def scrape_bodiva() -> dict[str, Any]:
    return BODIVAScraper().fetch(cache_key="bodiva").data


def get_bodiva_stocks(force_refresh: bool = False) -> pd.DataFrame:
    scraper = BODIVAScraper()
    key = "bodiva_stocks_refresh" if force_refresh else "bodiva_stocks"
    stocks = scraper.fetch(cache_key=key).data.get("stocks", {})
    return pd.DataFrame([{"codigo": code, **info} for code, info in stocks.items()])
