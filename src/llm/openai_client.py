"""
llm/openai_client.py — OpenAI provider implementation.

Environment variables:
    OPENAI_API_KEY   required
    OPENAI_MODEL     optional, default gpt-5.4-mini

Usage (internal — call via llm_client.py, not directly):
    from src.llm.openai_client import OpenAIClient
    client = OpenAIClient()
    text = client.generate_text(prompt)
    data = client.generate_json(prompt)
"""
from __future__ import annotations

import json
import os
from typing import Any

from src.utils.logger import get_logger

logger = get_logger(__name__)

DEFAULT_MODEL      = "gpt-5.4-mini"
DEFAULT_TEMP_TEXT  = 0.4   # commentary generation
DEFAULT_TEMP_JSON  = 0.1   # QA / structured outputs


class OpenAIClient:
    """
    Thin wrapper around the OpenAI Python SDK (v1+).

    All public methods return safe fallback values on failure —
    they never raise exceptions to the caller.
    """

    def __init__(self) -> None:
        self._client = None
        self.model   = os.getenv("OPENAI_MODEL", DEFAULT_MODEL)
        self._init()

    def _init(self) -> None:
        try:
            from openai import OpenAI
        except ImportError:
            logger.warning(
                "openai package not installed. "
                "Run: pip install openai  — OpenAI features disabled."
            )
            return

        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            logger.warning(
                "OPENAI_API_KEY not set — OpenAI features disabled. "
                "Add it to your .env file."
            )
            return

        try:
            from openai import OpenAI
            self._client = OpenAI(api_key=api_key)
            logger.info("OpenAI client initialised (model=%s)", self.model)
        except Exception as exc:
            logger.error("OpenAI client init failed: %s", exc)

    @property
    def available(self) -> bool:
        return self._client is not None

    # ── Text generation ───────────────────────────────────────────────────────

    def generate_text(self, prompt: str, fallback: str = "") -> str:
        """
        Send *prompt* to OpenAI and return the text response.
        Returns *fallback* on any failure.
        """
        if not self.available:
            return fallback or "Resumo automático não disponível (OpenAI não configurado)."

        try:
            response = self._client.responses.create(
                model=self.model,
                input=prompt,
                temperature=DEFAULT_TEMP_TEXT,
                max_output_tokens=512,
            )
            text = (getattr(response, "output_text", None) or "").strip()
            if text:
                return text
        except Exception as exc:
            logger.warning("OpenAI responses.generate_text failed: %s — trying chat fallback", exc)

        try:
            response = self._client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                temperature=DEFAULT_TEMP_TEXT,
                max_tokens=512,
            )
            text = response.choices[0].message.content or ""
            return text.strip()
        except Exception as exc:
            logger.error("OpenAI generate_text failed: %s", exc)
            return fallback or "Resumo automático não disponível."

    # ── JSON generation ───────────────────────────────────────────────────────

    def generate_json(
        self,
        prompt: str,
        schema: dict[str, Any] | None = None,
    ) -> dict[str, Any] | None:
        """
        Send *prompt* expecting a JSON response. Returns parsed dict or None.

        Tries the Responses API first, then falls back to chat completions.
        """
        if not self.available:
            return None

        try:
            text_config: dict[str, Any] = {"format": {"type": "json_object"}}
            if schema:
                text_config = {
                    "format": {
                        "type": "json_schema",
                        "name": "structured_output",
                        "schema": schema,
                        "strict": False,
                    }
                }
            response = self._client.responses.create(
                model=self.model,
                input=prompt,
                temperature=DEFAULT_TEMP_JSON,
                max_output_tokens=1024,
                text=text_config,
            )
            text = (getattr(response, "output_text", None) or "{}").strip()
            return json.loads(text)

        except Exception as exc:
            logger.warning(
                "OpenAI Responses JSON failed (%s) — retrying with chat fallback", exc
            )
            return self._generate_json_chat(prompt)

    def _generate_json_chat(self, prompt: str) -> dict[str, Any] | None:
        """Fallback: chat completion JSON mode, then plain text parse."""
        try:
            response = self._client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                temperature=DEFAULT_TEMP_JSON,
                max_tokens=1024,
                response_format={"type": "json_object"},
            )
            text = (response.choices[0].message.content or "{}").strip()
            return json.loads(text)
        except Exception as exc:
            logger.warning("OpenAI chat JSON mode failed: %s — retrying as plain text", exc)
            return self._generate_json_plain(prompt)

    def _generate_json_plain(self, prompt: str) -> dict[str, Any] | None:
        """Fallback: generate text then strip markdown fences and parse JSON."""
        try:
            response = self._client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                temperature=DEFAULT_TEMP_JSON,
                max_tokens=1024,
            )
            text = (response.choices[0].message.content or "").strip()
            # Strip markdown fences
            if text.startswith("```"):
                parts = text.split("```")
                text = parts[1] if len(parts) > 1 else text
                if text.startswith("json"):
                    text = text[4:]

            return json.loads(text.strip())
        except Exception as exc:
            logger.error("OpenAI _generate_json_plain failed: %s", exc)
            return None
