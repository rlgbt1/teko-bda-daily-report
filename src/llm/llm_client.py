"""
llm/llm_client.py — Provider-agnostic LLM router.

The rest of the app should import only this module, not provider SDKs.

Supported environment variables:
    LLM_PROVIDER=openai|gemini
    OPENAI_API_KEY=...
    OPENAI_MODEL=gpt-5.4-mini
    GEMINI_API_KEY=...
    GEMINI_MODEL=gemini-2.0-flash

Public API:
    generate_commentary(prompt, fallback="")
    generate_json(prompt, schema=None, fallback=None)
    review_scrape_packet(prompt, fallback=None)
    review_commentary(prompt, fallback=None)
    review_template(prompt, fallback=None)
    run_report_qa(...)
"""
from __future__ import annotations

import json
import os
from functools import lru_cache
from typing import Any

from src.utils.logger import get_logger

logger = get_logger(__name__)

DEFAULT_PROVIDER = "openai"


def _normalise_provider(value: str | None) -> str:
    provider = (value or DEFAULT_PROVIDER).strip().lower()
    if provider in {"openai", "gemini"}:
        return provider
    logger.warning(
        "Unknown LLM_PROVIDER='%s' — defaulting to %s",
        value, DEFAULT_PROVIDER,
    )
    return DEFAULT_PROVIDER


def get_provider_name() -> str:
    """Return the configured provider name."""
    return _normalise_provider(os.getenv("LLM_PROVIDER", DEFAULT_PROVIDER))


@lru_cache(maxsize=1)
def _build_client():
    """
    Build the configured provider client once per process.

    Falls back to the other provider if the configured one is unavailable.
    Returns None only if no provider can be initialised.
    """
    provider = get_provider_name()
    order = [provider, "gemini" if provider == "openai" else "openai"]

    for name in order:
        client = _init_provider(name)
        if client and getattr(client, "available", False):
            if name != provider:
                logger.warning(
                    "Configured LLM provider '%s' unavailable — using '%s' instead",
                    provider, name,
                )
            else:
                logger.info("LLM provider selected: %s", name)
            return client

    logger.warning("No LLM provider is available; AI features will use fallbacks")
    return None


def _init_provider(name: str):
    try:
        if name == "openai":
            from src.llm.openai_client import OpenAIClient
            return OpenAIClient()
        if name == "gemini":
            from src.llm.gemini_client import GeminiClient
            return GeminiClient()
    except Exception as exc:
        logger.error("Failed to initialise provider '%s': %s", name, exc)
    return None


def get_client():
    """Return the active provider client or None."""
    return _build_client()


def generate_commentary(prompt: str, fallback: str = "") -> str:
    """Generate free-form text using the active provider."""
    client = get_client()
    if client is None:
        return fallback or "Resumo automático não disponível."
    return client.generate_text(prompt, fallback=fallback)


def generate_json(
    prompt: str,
    schema: dict[str, Any] | None = None,
    fallback: dict[str, Any] | None = None,
) -> dict[str, Any] | None:
    """
    Generate structured JSON using the active provider.

    Args:
        prompt: LLM prompt.
        schema: Optional JSON schema. Providers may use it for stronger validation.
        fallback: Safe fallback dict when generation fails.
    """
    client = get_client()
    if client is None:
        return fallback

    try:
        data = client.generate_json(prompt, schema=schema)
        if data is None:
            return fallback
        return data
    except Exception as exc:
        logger.error("generate_json failed: %s", exc)
        return fallback


def review_scrape_packet(
    prompt: str,
    fallback: dict[str, Any] | None = None,
) -> dict[str, Any] | None:
    """Structured JSON helper for scrape QA."""
    return generate_json(prompt, fallback=fallback)


def review_commentary(
    prompt: str,
    fallback: dict[str, Any] | None = None,
) -> dict[str, Any] | None:
    """Structured JSON helper for content QA."""
    return generate_json(prompt, fallback=fallback)


def review_template(
    prompt: str,
    fallback: dict[str, Any] | None = None,
) -> dict[str, Any] | None:
    """Structured JSON helper for deck/template QA."""
    return generate_json(prompt, fallback=fallback)


def run_report_qa(
    brief: str,
    report_content: str,
    validation_results: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Review the generated report against the brief using the active provider.
    """
    fallback = {
        "overall_status": "unknown",
        "missing_requirements": [],
        "data_consistency_issues": [],
        "formatting_issues": [],
        "hallucination_risks": [],
        "recommended_fixes": [],
        "short_verdict": "QA automático não disponível (LLM não configurado).",
    }

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
    return generate_json(prompt, fallback=fallback) or fallback
