"""
llm/llm_client.py — Thin wrapper around Google Gemini.

All LLM calls in this project go through this module so that:
  - The API key is read once from the environment.
  - The model name is controlled from config.py (GEMINI_MODEL).
  - Swapping to a different model later means changing one line in config.py.

Environment variable required:
    GEMINI_API_KEY=<your key from https://aistudio.google.com/>

Usage:
    from src.llm.llm_client import generate_commentary, run_report_qa
"""

from __future__ import annotations

import json
import os
from typing import Any

from src.utils.logger import get_logger

logger = get_logger(__name__)

# ── SDK import (graceful degradation if not installed) ────────────────────────
try:
    import google.generativeai as genai
    _GENAI_AVAILABLE = True
except ImportError:
    _GENAI_AVAILABLE = False
    logger.warning(
        "google-generativeai not installed. "
        "Run: pip install google-generativeai  — LLM features will be disabled."
    )

# ── Model config ──────────────────────────────────────────────────────────────
try:
    from src.config import GEMINI_MODEL
except ImportError:
    GEMINI_MODEL = "gemini-2.0-flash"


def _get_model() -> Any | None:
    """
    Initialise and return a Gemini GenerativeModel.

    Returns None if the SDK is missing or the API key is not set.
    To set the key: add GEMINI_API_KEY=<key> to your .env file.
    """
    if not _GENAI_AVAILABLE:
        return None
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logger.warning(
            "GEMINI_API_KEY not set — LLM features disabled. "
            "Add it to your .env file."
        )
        return None
    genai.configure(api_key=api_key)
    return genai.GenerativeModel(GEMINI_MODEL)


def generate_commentary(prompt: str, fallback: str = "") -> str:
    """
    Send *prompt* to Gemini and return the text response.

    Args:
        prompt:   Full prompt to send (system instruction + data already embedded).
        fallback: String to return when Gemini is unavailable or errors out.

    Returns:
        Generated text string, or *fallback* on failure.
    """
    model = _get_model()
    if model is None:
        return fallback or "Resumo automático não disponível (Gemini não configurado)."
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as exc:
        logger.error("Gemini generate_commentary failed: %s", exc)
        return fallback or "Resumo automático não disponível."


def run_report_qa(
    brief: str,
    report_content: str,
    validation_results: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Ask Gemini to review the generated report against the brief.

    Returns a structured JSON dict with the following keys:
        overall_status        : "pass" | "pass_with_warnings" | "fail"
        missing_requirements  : list[str]
        data_consistency_issues: list[str]
        formatting_issues     : list[str]
        hallucination_risks   : list[str]
        recommended_fixes     : list[str]
        short_verdict         : str  (1-2 sentences in Portuguese)

    Returns a safe fallback dict if Gemini is unavailable.
    """
    fallback = {
        "overall_status": "unknown",
        "missing_requirements": [],
        "data_consistency_issues": [],
        "formatting_issues": [],
        "hallucination_risks": [],
        "recommended_fixes": [],
        "short_verdict": "QA automático não disponível (Gemini não configurado).",
    }

    model = _get_model()
    if model is None:
        return fallback

    val_summary = json.dumps(validation_results or {}, ensure_ascii=False, indent=2)

    prompt = f"""You are a senior financial report auditor reviewing a bank's daily executive report.

BRIEF (what the report must cover):
{brief}

REPORT CONTENT (extracted text from the generated PPTX):
{report_content}

DETERMINISTIC VALIDATION RESULTS:
{val_summary}

Your task: review the report and return ONLY valid JSON (no markdown fences) with exactly these keys:
{{
  "overall_status": "pass" | "pass_with_warnings" | "fail",
  "missing_requirements": ["..."],
  "data_consistency_issues": ["..."],
  "formatting_issues": ["..."],
  "hallucination_risks": ["..."],
  "recommended_fixes": ["..."],
  "short_verdict": "1-2 sentences in Portuguese summarising the verdict"
}}
"""

    try:
        response = model.generate_content(prompt)
        text = response.text.strip()
        # Strip markdown code fences if present
        if text.startswith("```"):
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text)
    except Exception as exc:
        logger.error("Gemini run_report_qa failed: %s", exc)
        return fallback
