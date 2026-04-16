"""
qa/validators.py — Deterministic (Python-only) validators for scrape packets.

Philosophy:
  - Python handles hard validation (missing data, bad types, impossible values)
  - Gemini handles reasoning, interpretation, and audit judgement
  - These checks run BEFORE any LLM call

Each validate_* function:
  - accepts a ScrapePacket
  - mutates packet.checks, packet.warnings, packet.errors
  - sets packet.status to FAIL/WARNING if thresholds are crossed
  - returns the packet (mutated in-place)
"""
from __future__ import annotations

import re
from typing import Any

import pandas as pd

from src.qa.schemas import QAStatus, ScrapePacket
from src.utils.logger import get_logger

logger = get_logger(__name__)

# ── Thresholds ────────────────────────────────────────────────────────────────
MIN_LUIBOR_ROWS   = 4   # expect at least 4 of 6 maturities
MIN_FX_ROWS       = 2   # expect at least USD + EUR
MIN_MARKET_ROWS   = 6   # expect at least 6 of 9 indices
MIN_COMMODITY_ROWS = 6
MIN_CRYPTO_ROWS   = 2
MIN_BODIVA_STOCKS  = 1  # BODIVA is fragile; 1+ stock is considered partial

NA_PLACEHOLDER    = re.compile(r"^(N/A|—|-{1,3}|n/a|na|null)$", re.I)
NUMERIC_PAT       = re.compile(r"^\d{1,3}([.,]\d+)*\s*%?$")


# ── Shared helpers ─────────────────────────────────────────────────────────────

def _rows_from_packet(packet: ScrapePacket, key: str) -> list[dict]:
    """Extract list-of-dicts from parsed_data, or [] if missing."""
    raw = packet.parsed_data.get(key, [])
    if isinstance(raw, list):
        return raw
    return []


def _count_na(rows: list[dict]) -> int:
    """Count cells whose value matches the NA placeholder pattern."""
    count = 0
    for row in rows:
        for v in row.values():
            if NA_PLACEHOLDER.match(str(v).strip()):
                count += 1
    return count


def _all_na(rows: list[dict]) -> bool:
    if not rows:
        return True
    total = sum(len(r) for r in rows)
    return _count_na(rows) == total


def _mark(packet: ScrapePacket, key: str, passed: bool, warning_msg: str = "", error_msg: str = "") -> None:
    packet.checks[key] = passed
    if not passed:
        if error_msg:
            packet.errors.append(error_msg)
        elif warning_msg:
            packet.warnings.append(warning_msg)


# ── BNA validators ────────────────────────────────────────────────────────────

def validate_bna(packet: ScrapePacket) -> ScrapePacket:
    """Validate a BNA scrape packet (LUIBOR + FX + rates)."""
    luibor_rows = _rows_from_packet(packet, "luibor")
    fx_rows     = _rows_from_packet(packet, "fx")
    taxa_bna    = packet.parsed_data.get("taxa_bna", "N/A")
    inflacao    = packet.parsed_data.get("inflacao", "N/A")

    # LUIBOR row count
    _mark(packet, "luibor_row_count_ok",
          len(luibor_rows) >= MIN_LUIBOR_ROWS,
          warning_msg=f"Only {len(luibor_rows)} LUIBOR rows (expected ≥{MIN_LUIBOR_ROWS})")

    # LUIBOR all-NA check
    luibor_na = _all_na(luibor_rows)
    _mark(packet, "luibor_not_all_na", not luibor_na,
          error_msg="All LUIBOR values are N/A — scraper likely failed")

    # Expected maturities present
    expected_maturities = {"Overnight", "1 Mês", "3 Meses", "6 Meses", "9 Meses", "12 Meses"}
    present = {r.get("Maturidade", "") for r in luibor_rows}
    missing = expected_maturities - present
    _mark(packet, "luibor_maturities_complete", len(missing) == 0,
          warning_msg=f"Missing LUIBOR maturities: {missing}")

    # FX row count
    _mark(packet, "fx_row_count_ok",
          len(fx_rows) >= MIN_FX_ROWS,
          warning_msg=f"Only {len(fx_rows)} FX rows (expected ≥{MIN_FX_ROWS})")

    # FX all-NA
    fx_na = _all_na(fx_rows)
    _mark(packet, "fx_not_all_na", not fx_na,
          error_msg="All FX values are N/A — BNA FX extraction failed")

    # Taxa BNA present
    taxa_ok = bool(taxa_bna) and not NA_PLACEHOLDER.match(str(taxa_bna))
    _mark(packet, "taxa_bna_present", taxa_ok,
          warning_msg=f"Taxa BNA not extracted (got: '{taxa_bna}')")

    # Inflação present
    inf_ok = bool(inflacao) and not NA_PLACEHOLDER.match(str(inflacao))
    _mark(packet, "inflacao_present", inf_ok,
          warning_msg=f"Taxa de Inflação not extracted (got: '{inflacao}')")

    # Derive status
    if packet.errors:
        packet.status = QAStatus.FAIL
    elif packet.warnings:
        packet.status = QAStatus.WARNING
    else:
        packet.status = QAStatus.PASS

    logger.debug("BNA validation: %s | errors=%d warnings=%d",
                 packet.status, len(packet.errors), len(packet.warnings))
    return packet


# ── Yahoo validators ──────────────────────────────────────────────────────────

def validate_yahoo_markets(packet: ScrapePacket) -> ScrapePacket:
    rows = _rows_from_packet(packet, "rows")

    _mark(packet, "row_count_ok",
          len(rows) >= MIN_MARKET_ROWS,
          warning_msg=f"Only {len(rows)} market rows (expected ≥{MIN_MARKET_ROWS})")

    na_count = _count_na(rows)
    total    = sum(len(r) for r in rows) or 1
    na_ratio = na_count / total
    _mark(packet, "low_na_ratio", na_ratio < 0.4,
          warning_msg=f"{na_count}/{total} cells are N/A ({na_ratio:.0%}) — Yahoo may be rate-limited")

    _mark(packet, "not_empty", bool(rows),
          error_msg="Markets DataFrame is empty")

    packet.status = _derive_status(packet)
    return packet


def validate_yahoo_commodities(packet: ScrapePacket) -> ScrapePacket:
    rows = _rows_from_packet(packet, "rows")

    _mark(packet, "row_count_ok",
          len(rows) >= MIN_COMMODITY_ROWS,
          warning_msg=f"Only {len(rows)} commodity rows (expected ≥{MIN_COMMODITY_ROWS})")

    _mark(packet, "not_empty", bool(rows),
          error_msg="Commodities DataFrame is empty")

    packet.status = _derive_status(packet)
    return packet


def validate_yahoo_crypto(packet: ScrapePacket) -> ScrapePacket:
    rows = _rows_from_packet(packet, "rows")

    _mark(packet, "row_count_ok",
          len(rows) >= MIN_CRYPTO_ROWS,
          warning_msg=f"Only {len(rows)} crypto rows (expected ≥{MIN_CRYPTO_ROWS})")

    _mark(packet, "not_empty", bool(rows),
          error_msg="Crypto DataFrame is empty")

    packet.status = _derive_status(packet)
    return packet


# ── BODIVA validator ──────────────────────────────────────────────────────────

def validate_bodiva(packet: ScrapePacket) -> ScrapePacket:
    stocks = packet.parsed_data.get("stocks", {})

    _mark(packet, "stocks_present",
          len(stocks) >= MIN_BODIVA_STOCKS,
          warning_msg=f"Only {len(stocks)} BODIVA stocks parsed — site may have changed structure")

    if len(stocks) == 0:
        packet.errors.append("BODIVA returned 0 stocks — slides will show placeholders")

    # Check for stocks with None current price
    null_prices = [code for code, d in stocks.items() if d.get("current") is None]
    if null_prices:
        packet.warnings.append(f"Stocks with missing price: {null_prices}")
    _mark(packet, "prices_present", len(null_prices) == 0,
          warning_msg=f"Some stock prices are None: {null_prices}")

    # Segment totals are always missing from BODIVA's rendered DOM
    if not packet.parsed_data.get("segments"):
        packet.warnings.append(
            "BODIVA segment totals not available (known limitation — server-push JSON only)"
        )

    packet.status = _derive_status(packet)
    return packet


# ── Dispatch ──────────────────────────────────────────────────────────────────

def _derive_status(packet: ScrapePacket) -> QAStatus:
    if packet.errors:
        return QAStatus.FAIL
    if packet.warnings:
        return QAStatus.WARNING
    return QAStatus.PASS


_VALIDATORS = {
    "bna":               validate_bna,
    "yahoo_markets":     validate_yahoo_markets,
    "yahoo_commodities": validate_yahoo_commodities,
    "yahoo_crypto":      validate_yahoo_crypto,
    "bodiva":            validate_bodiva,
}


def validate_packet(packet: ScrapePacket) -> ScrapePacket:
    """
    Run the appropriate deterministic validator for this packet.
    Falls back to a generic check (not-empty) if no specific validator found.
    """
    key = f"{packet.source.lower()}_{packet.step.lower()}"
    # Try exact match first, then source-only
    validator = _VALIDATORS.get(key) or _VALIDATORS.get(packet.source.lower())
    if validator:
        return validator(packet)

    # Generic fallback: just check parsed_data is not empty
    _mark(packet, "not_empty", bool(packet.parsed_data),
          warning_msg=f"No specific validator for {packet.source}/{packet.step}")
    packet.status = _derive_status(packet)
    return packet
