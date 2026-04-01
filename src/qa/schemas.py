"""
qa/schemas.py — Pydantic models for QA packets, results, and audit records.

Every scraper step, QA check, and final run produces objects defined here.
This keeps the shape of data explicit and prevents silent KeyError failures.
"""
from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


# ── Enums ─────────────────────────────────────────────────────────────────────

class QAStatus(str, Enum):
    PASS    = "pass"
    WARNING = "warning"
    FAIL    = "fail"
    UNKNOWN = "unknown"


# ── Scrape packet ─────────────────────────────────────────────────────────────

class ScrapePacket(BaseModel):
    """
    One structured unit produced by a scraper step.

    Evidence excerpt + deterministic checks are included so that:
    - Python validators can inspect the packet before any LLM call
    - Gemini QA can review evidence vs parsed values
    """
    source:      str                          # e.g. "BNA", "Yahoo", "BODIVA"
    step:        str                          # e.g. "luibor_rates", "global_markets"
    status:      QAStatus = QAStatus.UNKNOWN
    url:         str = ""
    timestamp:   datetime = Field(default_factory=datetime.utcnow)

    # Raw evidence the scraper captured
    raw_excerpt: str = ""                     # text snippet / HTML fragment

    # Parsed output — DataFrame rows as list-of-dicts, or plain dict
    parsed_data: dict[str, Any] = Field(default_factory=dict)

    # Deterministic check results (filled by validators.py)
    checks:      dict[str, bool] = Field(default_factory=dict)

    # Human-readable issues from deterministic checks
    warnings:    list[str] = Field(default_factory=list)
    errors:      list[str] = Field(default_factory=list)

    # Optional: how long the scrape took in seconds
    duration_s:  float | None = None


# ── QA result (LLM-backed) ────────────────────────────────────────────────────

class ScrapeQAResult(BaseModel):
    """
    Gemini's review of a ScrapePacket.
    The LLM never overrides deterministic failures — it adds audit judgement.
    """
    source:             str
    step:               str
    status:             QAStatus = QAStatus.UNKNOWN
    confidence:         float = 0.0           # 0.0 – 1.0
    hallucination_risk: QAStatus = QAStatus.UNKNOWN
    issues:             list[str] = Field(default_factory=list)
    recommended_action: str = ""
    safe_for_report:    bool = False
    llm_used:           bool = False          # False when Gemini was unavailable


class ContentQAResult(BaseModel):
    """
    QA result for a generated commentary block.
    Verifies that the text is grounded in the provided data.
    """
    section:         str                      # e.g. "cm_commentary"
    status:          QAStatus = QAStatus.UNKNOWN
    grounded:        bool = False
    issues:          list[str] = Field(default_factory=list)
    safe_to_include: bool = False
    fallback_used:   bool = False
    llm_used:        bool = False


class TemplateQAResult(BaseModel):
    """
    QA result for the generated PPTX against the BDA template spec.
    """
    status:                  QAStatus = QAStatus.UNKNOWN
    slide_count:             int = 0
    expected_slide_count:    int = 11
    missing_slides:          list[str] = Field(default_factory=list)
    title_issues:            list[str] = Field(default_factory=list)
    footer_issues:           list[str] = Field(default_factory=list)
    commentary_issues:       list[str] = Field(default_factory=list)
    placeholder_heavy_slides: list[int] = Field(default_factory=list)
    blank_slides:            list[int] = Field(default_factory=list)
    deterministic_issues:    list[str] = Field(default_factory=list)
    gemini_issues:           list[str] = Field(default_factory=list)
    safe_to_release:         bool = False
    llm_used:                bool = False


# ── Final audit ───────────────────────────────────────────────────────────────

class FinalAudit(BaseModel):
    """
    Full run-level audit record.
    Produced by run_tracker.build_final_audit().
    """
    run_id:              str
    timestamp:           datetime = Field(default_factory=datetime.utcnow)
    overall_status:      QAStatus = QAStatus.UNKNOWN
    overall_confidence:  float = 0.0

    safe_to_generate_ppt: bool = False
    safe_to_send_to_client: bool = False

    scrape_qa:    list[ScrapeQAResult]  = Field(default_factory=list)
    content_qa:   list[ContentQAResult] = Field(default_factory=list)
    template_qa:  TemplateQAResult | None = None

    # Human-readable summary sections
    scrape_integrity_notes:   list[str] = Field(default_factory=list)
    content_grounding_notes:  list[str] = Field(default_factory=list)
    template_compliance_notes: list[str] = Field(default_factory=list)
    slides_needing_review:    list[str] = Field(default_factory=list)
    recommended_action:       str = ""
