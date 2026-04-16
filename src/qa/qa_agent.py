"""
qa/qa_agent.py — WorkflowQAAgent: Gemini-backed scrape and content QA.

Responsibilities:
  - review_scrape_packet()  → ScrapeQAResult
  - review_commentary()     → ContentQAResult

This agent is separate from DailyReportAgent (the writer).
  DailyReportAgent  = writer (generates Portuguese commentary)
  WorkflowQAAgent   = checker / auditor (reviews evidence and output)

Fallback behaviour:
  If Gemini is unavailable the agent degrades gracefully:
  - scrape QA falls back to deterministic-only results
  - content QA marks commentary as unverified but still usable
"""
from __future__ import annotations

import json
from typing import Any

from src.llm.llm_client import review_commentary as llm_review_commentary
from src.llm.llm_client import review_scrape_packet as llm_review_scrape_packet
from src.qa.schemas import (
    ContentQAResult,
    QAStatus,
    ScrapePacket,
    ScrapeQAResult,
)
from src.qa.prompts import (
    CONTENT_QA_SYSTEM,
    CONTENT_QA_USER,
    SCRAPE_QA_SYSTEM,
    SCRAPE_QA_USER,
)
from src.utils.logger import get_logger

logger = get_logger(__name__)

class WorkflowQAAgent:
    """
    Gemini-backed audit layer for the BDA Daily Report workflow.

    All methods are safe to call even when Gemini is unavailable —
    they degrade to deterministic-only results with llm_used=False.
    """

    # ── Scrape QA ──────────────────────────────────────────────────────────────

    def review_scrape_packet(self, packet: ScrapePacket) -> ScrapeQAResult:
        """
        Review a validated ScrapePacket and return a ScrapeQAResult.

        If packet.status is already FAIL (hard deterministic failure),
        we still run Gemini for its audit judgement, but safe_for_report
        is forced to False regardless.
        """
        # Build prompt
        prompt = (
            SCRAPE_QA_SYSTEM + "\n\n"
            + SCRAPE_QA_USER.format(
                source=packet.source,
                step=packet.step,
                url=packet.url or "(no URL)",
                checks_json=json.dumps(packet.checks, indent=2),
                warnings="\n".join(packet.warnings) or "(none)",
                errors="\n".join(packet.errors) or "(none)",
                raw_excerpt=packet.raw_excerpt[:2000] or "(no excerpt)",
                parsed_data_json=json.dumps(packet.parsed_data, indent=2, default=str)[:3000],
            )
        )

        raw = llm_review_scrape_packet(prompt)
        llm_used = raw is not None

        if not llm_used:
            # Fallback: derive result entirely from deterministic checks
            raw = self._deterministic_fallback(packet)

        # Parse Gemini response defensively
        status_str = raw.get("status", "unknown")
        try:
            status = QAStatus(status_str)
        except ValueError:
            status = QAStatus.UNKNOWN

        hal_str = raw.get("hallucination_risk", "unknown")
        try:
            hal = QAStatus(hal_str)
        except ValueError:
            hal = QAStatus.UNKNOWN

        # Hard rule: if deterministic says FAIL, never mark safe_for_report
        safe = bool(raw.get("safe_for_report", False))
        if packet.status == QAStatus.FAIL:
            safe = False
            if status not in (QAStatus.FAIL, QAStatus.WARNING):
                status = QAStatus.WARNING  # at minimum warn

        result = ScrapeQAResult(
            source=packet.source,
            step=packet.step,
            status=status,
            confidence=float(raw.get("confidence", 0.5)),
            hallucination_risk=hal,
            issues=raw.get("issues", []) + packet.errors + packet.warnings,
            recommended_action=raw.get("recommended_action", ""),
            safe_for_report=safe,
            llm_used=llm_used,
        )

        logger.info(
            "Scrape QA [%s/%s]: %s (safe=%s, llm=%s)",
            packet.source, packet.step, result.status, result.safe_for_report, llm_used,
        )
        return result

    def _deterministic_fallback(self, packet: ScrapePacket) -> dict[str, Any]:
        """Build a QA result dict from deterministic checks alone (no LLM)."""
        if packet.status == QAStatus.FAIL:
            status = "fail"
            confidence = 0.2
            safe = False
        elif packet.status == QAStatus.WARNING:
            status = "warning"
            confidence = 0.6
            safe = True   # warnings still allow report generation
        else:
            status = "pass"
            confidence = 0.85
            safe = True

        issues = list(packet.errors) + list(packet.warnings)
        return {
            "status": status,
            "confidence": confidence,
            "hallucination_risk": "pass" if not packet.errors else "warning",
            "issues": issues,
            "recommended_action": "Revisar manualmente." if packet.errors else "OK.",
            "safe_for_report": safe,
        }

    # ── Content QA ─────────────────────────────────────────────────────────────

    def review_commentary(
        self,
        section: str,
        commentary: str,
        data_str: str,
    ) -> ContentQAResult:
        """
        Verify that *commentary* is grounded in *data_str*.

        Returns ContentQAResult with safe_to_include=True only if
        Gemini confirms grounding, or if Gemini is unavailable and
        commentary is non-empty (conservative pass-through).
        """
        if not commentary or not commentary.strip():
            return ContentQAResult(
                section=section,
                status=QAStatus.FAIL,
                grounded=False,
                issues=["Commentary is empty"],
                safe_to_include=False,
                fallback_used=True,
            )

        prompt = (
            CONTENT_QA_SYSTEM + "\n\n"
            + CONTENT_QA_USER.format(
                section=section,
                data_str=data_str[:2000],
                commentary=commentary[:1500],
            )
        )

        raw = llm_review_commentary(prompt)
        llm_used = raw is not None

        if not llm_used:
            # Conservative fallback: allow non-empty commentary through
            logger.warning(
                "Content QA for '%s': Gemini unavailable — using conservative pass-through",
                section,
            )
            return ContentQAResult(
                section=section,
                status=QAStatus.WARNING,
                grounded=False,   # we cannot confirm grounding without LLM
                issues=["Gemini unavailable — grounding not verified"],
                safe_to_include=True,
                fallback_used=False,
                llm_used=False,
            )

        status_str = raw.get("status", "unknown")
        try:
            status = QAStatus(status_str)
        except ValueError:
            status = QAStatus.UNKNOWN

        result = ContentQAResult(
            section=section,
            status=status,
            grounded=bool(raw.get("grounded", False)),
            issues=raw.get("issues", []),
            safe_to_include=bool(raw.get("safe_to_include", False)),
            fallback_used=False,
            llm_used=True,
        )

        logger.info(
            "Content QA [%s]: %s (grounded=%s, safe=%s)",
            section, result.status, result.grounded, result.safe_to_include,
        )
        return result

    # ── Release gating helpers ─────────────────────────────────────────────────

    @staticmethod
    def scrape_is_safe_to_proceed(results: list[ScrapeQAResult]) -> bool:
        """
        Return True if no FAIL-level scrape QA result exists.
        Warnings are allowed; a single FAIL blocks report generation.
        """
        return all(r.status != QAStatus.FAIL for r in results)

    @staticmethod
    def safe_commentary(
        section: str,
        commentary: str,
        qa_result: ContentQAResult,
        fallback: str = "",
    ) -> tuple[str, bool]:
        """
        Return (text_to_use, used_fallback).
        If content QA says the commentary is unsafe, return the fallback.
        """
        if qa_result.safe_to_include:
            return commentary, False
        logger.warning(
            "Content QA blocked commentary for '%s': %s — using fallback",
            section, qa_result.issues,
        )
        return fallback or f"Resumo de {section} não disponível.", True
