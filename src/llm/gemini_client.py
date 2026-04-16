"""
llm/gemini_client.py — Gemini provider implementation.

Extracted from the original llm_client.py so the router (llm_client.py)
can pick between OpenAI and Gemini without duplication.

Environment variables:
    GEMINI_API_KEY   required
    GEMINI_MODEL     optional, default gemini-2.0-flash

Usage (internal — call via llm_client.py, not directly):
    from src.llm.gemini_client import GeminiClient
    client = GeminiClient()
    text = client.generate_text(prompt)
    data = client.generate_json(prompt)
"""
from __future__ import annotations

import json
import os
from typing import Any

from src.utils.logger import get_logger

logger = get_logger(__name__)

DEFAULT_MODEL     = "gemini-2.0-flash"
DEFAULT_TEMP_TEXT = 0.4
DEFAULT_TEMP_JSON = 0.1


class GeminiClient:
    """
    Thin wrapper around google-genai (current SDK) with a fallback to
    google-generativeai (deprecated but still functional).

    All public methods return safe fallback values on failure.
    """

    def __init__(self) -> None:
        self._model_name = os.getenv("GEMINI_MODEL", DEFAULT_MODEL)
        self._client     = None   # google.genai client (new SDK)
        self._legacy     = None   # GenerativeModel (old SDK)
        self._init()

    def _init(self) -> None:
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            logger.warning(
                "GEMINI_API_KEY not set — Gemini features disabled. "
                "Add it to your .env file."
            )
            return

        # Try new SDK first (google-genai)
        try:
            import google.genai as genai
            self._client = genai.Client(api_key=api_key)
            logger.info("Gemini client initialised via google-genai (model=%s)", self._model_name)
            return
        except ImportError:
            pass  # fall through to legacy SDK
        except Exception as exc:
            logger.warning("google-genai init failed: %s — trying legacy SDK", exc)

        # Fallback: deprecated google-generativeai
        try:
            import google.generativeai as genai_legacy  # type: ignore[import]
            genai_legacy.configure(api_key=api_key)
            self._legacy = genai_legacy.GenerativeModel(self._model_name)
            logger.info(
                "Gemini client initialised via google-generativeai (deprecated) "
                "(model=%s)", self._model_name
            )
        except ImportError:
            logger.warning(
                "Neither google-genai nor google-generativeai installed. "
                "Run: pip install google-genai  — Gemini features disabled."
            )
        except Exception as exc:
            logger.error("Gemini legacy init failed: %s", exc)

    @property
    def available(self) -> bool:
        return self._client is not None or self._legacy is not None

    # ── Text generation ───────────────────────────────────────────────────────

    def generate_text(self, prompt: str, fallback: str = "") -> str:
        if not self.available:
            return fallback or "Resumo automático não disponível (Gemini não configurado)."

        try:
            if self._client is not None:
                from google.genai import types
                response = self._client.models.generate_content(
                    model=self._model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(temperature=DEFAULT_TEMP_TEXT),
                )
                return (response.text or "").strip()

            # Legacy SDK
            response = self._legacy.generate_content(prompt)
            return (response.text or "").strip()

        except Exception as exc:
            logger.error("Gemini generate_text failed: %s", exc)
            return fallback or "Resumo automático não disponível."

    # ── JSON generation ───────────────────────────────────────────────────────

    def generate_json(
        self,
        prompt: str,
        schema: dict[str, Any] | None = None,
    ) -> dict[str, Any] | None:
        if not self.available:
            return None

        try:
            if self._client is not None:
                from google.genai import types
                response = self._client.models.generate_content(
                    model=self._model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=DEFAULT_TEMP_JSON,
                        response_mime_type="application/json",
                    ),
                )
                text = (response.text or "{}").strip()
            else:
                # Legacy SDK — plain text + manual parse
                response = self._legacy.generate_content(prompt)
                text = (response.text or "{}").strip()

            # Strip markdown fences just in case
            if text.startswith("```"):
                parts = text.split("```")
                text  = parts[1] if len(parts) > 1 else text
                if text.startswith("json"):
                    text = text[4:]

            return json.loads(text.strip())

        except Exception as exc:
            logger.error("Gemini generate_json failed: %s", exc)
            return None
