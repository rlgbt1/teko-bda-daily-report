"""
scrapers/market_aggregator.py — Orchestrates all external data scraping + QA.

Public API
----------
scrape_all_external_data()
    Legacy call — returns the raw data dict only (backward compatible).

scrape_all_external_data_with_qa()
    New call — returns:
    {
        "data": {markets, commodities, crypto, luibor, fx_rates, bodiva, bna_rates},
        "packets": {<step>: ScrapePacket, ...},
        "qa":  {<step>: ScrapeQAResult, ...},
        "safe_to_proceed": bool,
    }
"""
from __future__ import annotations

import time
from typing import Any

import pandas as pd

from src.qa.schemas import QAStatus, ScrapePacket
from src.qa.validators import validate_packet
from src.utils.logger import get_logger

logger = get_logger(__name__)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _df_to_rows(df: pd.DataFrame) -> list[dict]:
    """Convert a DataFrame to list-of-dicts for ScrapePacket.parsed_data."""
    if df is None or df.empty:
        return []
    return df.to_dict(orient="records")


def _scrape_step(
    name: str,
    step: str,
    fn,
    url: str = "",
    *,
    parsed_key: str = "rows",
    extra_data: dict | None = None,
) -> tuple[Any, ScrapePacket]:
    """
    Run *fn*, capture timing, build and validate a ScrapePacket.
    Returns (raw_result, validated_packet).
    """
    t0 = time.monotonic()
    raw_excerpt = ""
    result = None

    try:
        result = fn()
        duration = time.monotonic() - t0

        # Build parsed_data
        if isinstance(result, pd.DataFrame):
            parsed = {parsed_key: _df_to_rows(result)}
        elif isinstance(result, dict):
            parsed = result
        else:
            parsed = {parsed_key: result}

        packet = ScrapePacket(
            source=name,
            step=step,
            url=url,
            parsed_data=parsed,
            raw_excerpt=raw_excerpt,
            duration_s=round(duration, 2),
        )
        validate_packet(packet)

    except Exception as exc:
        duration = time.monotonic() - t0
        logger.error("%s/%s scrape failed: %s", name, step, exc)
        packet = ScrapePacket(
            source=name,
            step=step,
            url=url,
            status=QAStatus.FAIL,
            parsed_data={},
            errors=[str(exc)],
            duration_s=round(duration, 2),
        )

    return result, packet


# ── Scrape functions ───────────────────────────────────────────────────────────

def _scrape_markets() -> tuple[pd.DataFrame, ScrapePacket]:
    from src.scrapers.yahoo_scraper import get_global_markets
    return _scrape_step("Yahoo", "markets", get_global_markets,
                        url="https://finance.yahoo.com")


def _scrape_commodities() -> tuple[pd.DataFrame, ScrapePacket]:
    from src.scrapers.yahoo_scraper import get_commodities
    return _scrape_step("Yahoo", "commodities", get_commodities,
                        url="https://finance.yahoo.com")


def _scrape_crypto() -> tuple[pd.DataFrame, ScrapePacket]:
    from src.scrapers.yahoo_scraper import get_crypto
    return _scrape_step("Yahoo", "crypto", get_crypto,
                        url="https://finance.yahoo.com")


def _scrape_bna_all() -> tuple[dict, ScrapePacket]:
    from src.scrapers.bna_scraper import scrape_once
    from src.config import BNA_URL

    def _fn():
        return scrape_once(force_refresh=True)

    return _scrape_step("BNA", "bna", _fn, url=BNA_URL)


def _bna_df_or_default(rows: list[dict], columns: list[str], defaults: list[list[str]]) -> pd.DataFrame:
    if rows:
        return pd.DataFrame(rows)
    return pd.DataFrame(dict(zip(columns, values)) for values in defaults)


def _scrape_bodiva() -> tuple[dict, ScrapePacket]:
    from src.scrapers.bodiva_scraper import BODIVAScraper
    from src.config import URLs

    def _fn():
        s = BODIVAScraper()
        return {
            "stocks":   s.get_stocks(),
            "segments": s.get_segments(),
        }

    return _scrape_step("BODIVA", "bodiva", _fn, url=URLs.BODIVA_MARKET)


# ── QA orchestration ──────────────────────────────────────────────────────────

def _run_qa(packets: dict[str, ScrapePacket]) -> dict[str, Any]:
    """
    Run WorkflowQAAgent.review_scrape_packet() for each validated packet.
    Returns dict of {step: ScrapeQAResult}.
    """
    from src.qa.qa_agent import WorkflowQAAgent
    agent = WorkflowQAAgent()
    results = {}
    for step, packet in packets.items():
        results[step] = agent.review_scrape_packet(packet)
    return results


# ── Public API ─────────────────────────────────────────────────────────────────

def scrape_all_external_data() -> dict:
    """
    Backward-compatible: scrape and return raw data dict only.
    No QA is run.
    """
    logger.info("Scraping external data (legacy mode — no QA)")

    from src.scrapers.yahoo_scraper import get_global_markets, get_commodities, get_crypto
    from src.scrapers.bna_scraper   import get_luibor_rates, get_exchange_rates, get_bna_rates

    print("Scraping global markets...")
    markets = get_global_markets()
    print("Scraping commodities...")
    commodities = get_commodities()
    print("Scraping crypto...")
    crypto = get_crypto()
    print("Scraping BNA LUIBOR rates...")
    luibor = get_luibor_rates()
    print("Scraping BNA exchange rates...")
    fx_rates = get_exchange_rates()

    return {
        "markets":    markets,
        "commodities": commodities,
        "crypto":     crypto,
        "luibor":     luibor,
        "fx_rates":   fx_rates,
    }


def scrape_all_external_data_with_qa(run_gemini_qa: bool = True) -> dict:
    """
    Full scrape + deterministic validation + optional Gemini QA.

    Returns:
    {
        "data": {
            "markets":    pd.DataFrame,
            "commodities": pd.DataFrame,
            "crypto":     pd.DataFrame,
            "luibor":     pd.DataFrame,
            "fx_rates":   pd.DataFrame,
            "bna_rates":  dict,          # taxa_bna, inflacao
            "bodiva":     dict,          # stocks, segments
        },
        "packets": {step: ScrapePacket, ...},
        "qa":      {step: ScrapeQAResult, ...},
        "safe_to_proceed": bool,
    }
    """
    logger.info("Scraping external data with QA (gemini_qa=%s)", run_gemini_qa)

    packets: dict[str, ScrapePacket] = {}

    # ── Yahoo ─────────────────────────────────────────────────────────────────
    print("Scraping global markets...")
    markets_raw, packets["markets"] = _scrape_markets()

    print("Scraping commodities...")
    commodities_raw, packets["commodities"] = _scrape_commodities()

    print("Scraping crypto...")
    crypto_raw, packets["crypto"] = _scrape_crypto()

    # ── BNA ───────────────────────────────────────────────────────────────────
    print("Scraping BNA...")
    bna_raw, packets["bna"] = _scrape_bna_all()
    bna_data = bna_raw if isinstance(bna_raw, dict) else {}
    luibor_rows = bna_data.get("luibor", [])
    fx_rows = bna_data.get("fx", [])
    bna_rates = {
        "taxa_bna": bna_data.get("taxa_bna", "N/A"),
        "inflacao": bna_data.get("inflacao", "N/A"),
    }
    luibor = _bna_df_or_default(
        luibor_rows,
        ["Maturidade", "Taxa (%)"],
        [["Overnight", "N/A"], ["1 Mês", "N/A"], ["3 Meses", "N/A"], ["6 Meses", "N/A"], ["9 Meses", "N/A"], ["12 Meses", "N/A"]],
    )
    fx_rates = _bna_df_or_default(
        fx_rows,
        ["Moeda", "Taxa (AOA)"],
        [["USD", "N/A"], ["EUR", "N/A"], ["ZAR", "N/A"]],
    )

    # ── BODIVA ────────────────────────────────────────────────────────────────
    print("Scraping BODIVA...")
    bodiva_raw, packets["bodiva"] = _scrape_bodiva()

    # ── QA ────────────────────────────────────────────────────────────────────
    if run_gemini_qa:
        print("Running Gemini scrape QA...")
        qa_results = _run_qa(packets)
    else:
        # Deterministic-only: convert packet status to minimal QAResult
        from src.qa.schemas import ScrapeQAResult
        qa_results = {}
        for step, packet in packets.items():
            safe = packet.status != QAStatus.FAIL
            qa_results[step] = ScrapeQAResult(
                source=packet.source,
                step=step,
                status=packet.status,
                confidence=0.8 if safe else 0.2,
                safe_for_report=safe,
                issues=packet.errors + packet.warnings,
                llm_used=False,
            )

    safe_to_proceed = all(r.safe_for_report for r in qa_results.values())

    if not safe_to_proceed:
        failed = [s for s, r in qa_results.items() if not r.safe_for_report]
        logger.warning("Scrape QA: NOT safe to proceed. Failed steps: %s", failed)
    else:
        logger.info("Scrape QA: all steps safe to proceed.")

    return {
        "data": {
            "markets":     markets_raw if isinstance(markets_raw, pd.DataFrame) else pd.DataFrame(),
            "commodities": commodities_raw if isinstance(commodities_raw, pd.DataFrame) else pd.DataFrame(),
            "crypto":      crypto_raw if isinstance(crypto_raw, pd.DataFrame) else pd.DataFrame(),
            "luibor":      luibor,
            "fx_rates":    fx_rates,
            "bna_rates":   bna_rates,
            "bodiva":      bodiva_raw or {"stocks": {}, "segments": {}},
        },
        "packets": packets,
        "qa":      qa_results,
        "safe_to_proceed": safe_to_proceed,
    }
