"""
qa/template_checker.py — Template compliance QA for the generated PPTX.

Two passes:
  1. Deterministic checks  (Python / python-pptx)
  2. Gemini review         (structured export passed to LLM)

The coded template spec in pptx_builder.py is the source of truth.
When a real reference .pptx is available, add its path to
check_template_compliance(reference_pptx=...) to enable comparison.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import json

from src.llm.llm_client import review_template
from src.qa.deck_exporter import export_deck
from src.qa.prompts import TEMPLATE_QA_SYSTEM, TEMPLATE_QA_USER
from src.qa.schemas import QAStatus, TemplateQAResult
from src.utils.logger import get_logger

logger = get_logger(__name__)

# ── Template spec constants ────────────────────────────────────────────────────
EXPECTED_SLIDE_COUNT = 11

EXPECTED_TITLES = {
    1:  "resumo diário dos mercados",       # cover
    2:  "agenda",
    3:  "sumário executivo",
    4:  "liquidez",                          # partial match
    5:  "liquidez",
    6:  "liquidez",
    7:  "mercado cambial",
    8:  "bodiva",
    9:  "operações",
    10: "informação de mercados",
    11: "informação de mercados",
}

COMMENTARY_REQUIRED_SLIDES = {10, 11}
FOOTER_REQUIRED_SLIDES     = set(range(1, 12))   # all slides
MIN_SHAPES_PER_SLIDE       = 3
NA_HEAVY_THRESHOLD         = 4   # more than this many N/A cells → warning


# ── Deterministic checks ───────────────────────────────────────────────────────

def _check_slide_count(deck: dict) -> list[str]:
    n = deck.get("slide_count", 0)
    if n != EXPECTED_SLIDE_COUNT:
        return [f"Slide count is {n}, expected {EXPECTED_SLIDE_COUNT}"]
    return []


def _check_titles(deck: dict) -> list[str]:
    issues = []
    for slide in deck.get("slides", []):
        idx   = slide["index"]
        title = slide.get("title", "").lower()
        expected = EXPECTED_TITLES.get(idx, "")
        if expected and expected not in title:
            issues.append(
                f"Slide {idx}: title '{slide.get('title', '')}' "
                f"does not contain expected '{expected}'"
            )
    return issues


def _check_footers(deck: dict) -> list[str]:
    issues = []
    for slide in deck.get("slides", []):
        idx = slide["index"]
        if idx in FOOTER_REQUIRED_SLIDES and not slide.get("footer_present", False):
            issues.append(f"Slide {idx}: footer not detected")
    return issues


def _check_blank_slides(deck: dict) -> list[int]:
    blank = []
    for slide in deck.get("slides", []):
        if slide.get("text_shape_count", 0) < MIN_SHAPES_PER_SLIDE:
            blank.append(slide["index"])
    return blank


def _check_na_heavy_slides(deck: dict) -> list[int]:
    heavy = []
    for slide in deck.get("slides", []):
        if slide.get("na_cell_count", 0) > NA_HEAVY_THRESHOLD:
            heavy.append(slide["index"])
    return heavy


def _check_commentary(deck: dict) -> list[str]:
    issues = []
    for slide in deck.get("slides", []):
        idx = slide["index"]
        if idx in COMMENTARY_REQUIRED_SLIDES:
            if slide.get("commentary_length", 0) < 50:
                issues.append(
                    f"Slide {idx}: commentary block missing or too short "
                    f"({slide.get('commentary_length', 0)} chars)"
                )
    return issues


def run_deterministic_checks(deck: dict) -> tuple[list[str], list[int], list[int]]:
    """
    Returns (issues, blank_slide_indices, na_heavy_slide_indices).
    """
    issues: list[str] = []
    issues += _check_slide_count(deck)
    issues += _check_titles(deck)
    issues += _check_footers(deck)
    issues += _check_commentary(deck)
    blank   = _check_blank_slides(deck)
    na_heavy = _check_na_heavy_slides(deck)
    for b in blank:
        issues.append(f"Slide {b}: appears blank or near-blank (< {MIN_SHAPES_PER_SLIDE} text shapes)")
    for n in na_heavy:
        issues.append(f"Slide {n}: table has > {NA_HEAVY_THRESHOLD} N/A cells — may be placeholder-heavy")
    return issues, blank, na_heavy


# ── LLM QA pass ────────────────────────────────────────────────────────────────

def _llm_template_qa(deck: dict, deterministic_issues: list[str]) -> dict[str, Any] | None:
    prompt = (
        TEMPLATE_QA_SYSTEM + "\n\n"
        + TEMPLATE_QA_USER.format(
            deck_json=json.dumps(deck, indent=2, default=str)[:6000],
            deterministic_issues="\n".join(deterministic_issues) or "(none)",
        )
    )
    try:
        return review_template(prompt)
    except Exception as exc:
        logger.error("template_checker LLM call failed: %s", exc)
        return None


# ── Main entry point ───────────────────────────────────────────────────────────

def check_template_compliance(
    pptx_path: str | Path,
    reference_pptx: str | Path | None = None,
) -> TemplateQAResult:
    """
    Run full template QA on a generated PPTX.

    Args:
        pptx_path:      Path to the generated .pptx file.
        reference_pptx: Optional path to a reference template .pptx.
                        Reserved for future comparison logic.

    Returns:
        TemplateQAResult with deterministic + Gemini findings.
    """
    deck = export_deck(pptx_path)

    if "error" in deck:
        return TemplateQAResult(
            status=QAStatus.FAIL,
            deterministic_issues=[deck["error"]],
            safe_to_release=False,
        )

    # ── Pass 1: deterministic ─────────────────────────────────────────────────
    det_issues, blank_slides, na_slides = run_deterministic_checks(deck)

    # ── Pass 2: provider-agnostic LLM review ─────────────────────────────────
    raw = _llm_template_qa(deck, det_issues)
    llm_used = raw is not None

    llm_issues: list[str] = []
    llm_slides_review: list[int] = []
    llm_safe = None

    if llm_used:
        llm_issues        = raw.get("issues", [])
        llm_slides_review = raw.get("slides_needing_review", [])
        llm_safe          = raw.get("safe_to_release")

        gemini_status_str = raw.get("status", "unknown")
        try:
            gemini_status = QAStatus(gemini_status_str)
        except ValueError:
            gemini_status = QAStatus.UNKNOWN
    else:
        gemini_status = QAStatus.UNKNOWN

    # ── Derive final status ───────────────────────────────────────────────────
    # Hard fail conditions
    hard_fail = (
        deck.get("slide_count", 0) != EXPECTED_SLIDE_COUNT
        or len(blank_slides) > 2
    )

    if hard_fail:
        final_status = QAStatus.FAIL
    elif det_issues or gemini_status == QAStatus.WARNING:
        final_status = QAStatus.WARNING
    elif gemini_status == QAStatus.FAIL:
        final_status = QAStatus.FAIL
    else:
        final_status = QAStatus.PASS

    # safe_to_release: require both deterministic pass and LLM approval
    if llm_safe is not None:
        safe = (not hard_fail) and llm_safe
    else:
        # LLM unavailable — be conservative
        safe = (not hard_fail) and len(det_issues) == 0

    # Reference template comparison (future extension point)
    if reference_pptx:
        logger.info(
            "template_checker: reference_pptx provided (%s) — "
            "comparison not yet implemented", reference_pptx
        )

    result = TemplateQAResult(
        status=final_status,
        slide_count=deck.get("slide_count", 0),
        expected_slide_count=EXPECTED_SLIDE_COUNT,
        blank_slides=blank_slides,
        placeholder_heavy_slides=na_slides,
        deterministic_issues=det_issues,
        gemini_issues=llm_issues,
        slides_needing_review=llm_slides_review,
        safe_to_release=safe,
        llm_used=llm_used,
    )

    logger.info(
        "Template QA: %s (safe_to_release=%s, llm=%s, det_issues=%d, gemini_issues=%d)",
        result.status, result.safe_to_release, llm_used,
        len(det_issues), len(gemini_issues),
    )
    return result
