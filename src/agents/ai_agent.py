"""
agents/ai_agent.py — High-level AI commentary agent backed by Gemini.

Wraps src/llm/llm_client.py with domain-specific prompt building so the
Streamlit app and report builder can call simple, named methods.

Roles
-----
DailyReportAgent  — writer: generates Portuguese commentary blocks
WorkflowQAAgent   — checker: see src/qa/qa_agent.py

Content QA flow
---------------
Each write_and_verify_* method:
  1. generates commentary via Gemini
  2. runs content QA (WorkflowQAAgent.review_commentary)
  3. returns the commentary only if it is safe_to_include
  4. returns a safe Portuguese fallback if QA fails or Gemini is unavailable

Usage:
    from src.agents.ai_agent import DailyReportAgent
    agent = DailyReportAgent()
    text, qa_result = agent.write_and_verify_markets(markets_df)
"""

from __future__ import annotations

import pandas as pd

from src.llm.llm_client import generate_commentary
from src.utils.logger import get_logger

logger = get_logger(__name__)

_SYSTEM = (
    "Você é um analista financeiro sénior de um banco angolano (BDA). "
    "Escreva resumos claros, factuais e concisos em português europeu. "
    "Use APENAS os dados fornecidos — nunca invente valores ou factos."
)


class DailyReportAgent:
    """
    Generates Portuguese commentary blocks for each section of the daily report.

    Simple API (no QA): summarize_*()
    Verified API (with content QA): write_and_verify_*() → (str, ContentQAResult)

    If Gemini is unavailable a safe Portuguese fallback is returned instead
    of raising an exception.
    """

    def __init__(self, run_content_qa: bool = True) -> None:
        self._run_qa = run_content_qa
        self._qa_agent = None
        if run_content_qa:
            try:
                from src.qa.qa_agent import WorkflowQAAgent
                self._qa_agent = WorkflowQAAgent()
            except ImportError:
                logger.warning("WorkflowQAAgent unavailable — content QA disabled")

    # ── Internal helper ───────────────────────────────────────────────────────

    def _generate_and_verify(
        self,
        section: str,
        prompt: str,
        data_str: str,
        fallback: str,
    ):
        """
        Generate commentary then run content QA.
        Returns (final_text, ContentQAResult | None).
        """
        from src.qa.schemas import ContentQAResult, QAStatus

        text = generate_commentary(prompt, fallback=fallback)

        if not self._qa_agent:
            return text, None

        qa_result = self._qa_agent.review_commentary(
            section=section,
            commentary=text,
            data_str=data_str,
        )
        final_text, used_fallback = WorkflowQAAgent_safe_commentary(
            section, text, qa_result, fallback
        )
        if used_fallback:
            qa_result.fallback_used = True
        return final_text, qa_result

    # ── Markets ───────────────────────────────────────────────────────────────

    def summarize_markets(self, markets_df: pd.DataFrame | None) -> str:
        text, _ = self.write_and_verify_markets(markets_df)
        return text

    def write_and_verify_markets(self, markets_df: pd.DataFrame | None):
        fallback = "Dados de mercados globais não disponíveis."
        if markets_df is None or markets_df.empty:
            return fallback, None
        data_str = markets_df.to_string(index=False)
        prompt = (
            f"{_SYSTEM}\n\n"
            "Com base nos dados abaixo, escreva um resumo de 3-4 frases sobre os "
            "principais movimentos dos mercados de capitais globais. "
            "Destaque os movimentos mais significativos (maiores altas e baixas).\n\n"
            f"{data_str}\n\nResumo:"
        )
        return self._generate_and_verify("cm_commentary", prompt, data_str, fallback)

    # ── Commodities ───────────────────────────────────────────────────────────

    def summarize_commodities(self, commodities_df: pd.DataFrame | None) -> str:
        text, _ = self.write_and_verify_commodities(commodities_df)
        return text

    def write_and_verify_commodities(self, commodities_df: pd.DataFrame | None):
        fallback = "Dados de commodities não disponíveis."
        if commodities_df is None or commodities_df.empty:
            return fallback, None
        data_str = commodities_df.to_string(index=False)
        prompt = (
            f"{_SYSTEM}\n\n"
            "Com base nos dados abaixo, escreva 2-3 frases sobre os principais "
            "movimentos das commodities e minerais. Mencione petróleo, ouro e cobre "
            "se presentes.\n\n"
            f"{data_str}\n\nResumo:"
        )
        return self._generate_and_verify("commodities_commentary", prompt, data_str, fallback)

    # ── Crypto ────────────────────────────────────────────────────────────────

    def summarize_crypto(self, crypto_df: pd.DataFrame | None) -> str:
        text, _ = self.write_and_verify_crypto(crypto_df)
        return text

    def write_and_verify_crypto(self, crypto_df: pd.DataFrame | None):
        fallback = "Dados de criptomoedas não disponíveis."
        if crypto_df is None or crypto_df.empty:
            return fallback, None
        data_str = crypto_df.to_string(index=False)
        prompt = (
            f"{_SYSTEM}\n\n"
            "Com base nos dados abaixo, escreva 2 frases sobre o mercado de "
            "criptomoedas, mencionando Bitcoin, Ethereum e tendência geral.\n\n"
            f"{data_str}\n\nResumo:"
        )
        return self._generate_and_verify("crypto_commentary", prompt, data_str, fallback)

    # ── FX / Cambial ──────────────────────────────────────────────────────────

    def summarize_fx(self, fx_data: dict | None) -> str:
        if not fx_data:
            return "Dados cambiais não disponíveis."
        data_str = "\n".join(f"{k}: {v}" for k, v in fx_data.items())
        prompt = (
            f"{_SYSTEM}\n\n"
            "Com base nos dados cambiais abaixo, escreva 1-2 frases sobre a "
            "evolução do kwanza face ao dólar e euro.\n\n"
            f"{data_str}\n\nResumo:"
        )
        return generate_commentary(prompt, fallback="")

    # ── Liquidity ─────────────────────────────────────────────────────────────

    def summarize_liquidity(
        self,
        liquidez_mn: dict | None,
        liquidez_me: dict | None,
    ) -> str:
        parts = []
        if liquidez_mn:
            parts.append("MN: " + ", ".join(f"{k}={v}" for k, v in liquidez_mn.items()))
        if liquidez_me:
            parts.append("ME: " + ", ".join(f"{k}={v}" for k, v in liquidez_me.items()))
        if not parts:
            return ""
        prompt = (
            f"{_SYSTEM}\n\n"
            "Com base na posição de liquidez abaixo, escreva 1-2 frases de "
            "enquadramento executivo.\n\n"
            + "\n".join(parts)
            + "\n\nResumo:"
        )
        return generate_commentary(prompt, fallback="")

    # ── Minerals (slide 11) ───────────────────────────────────────────────────

    def summarize_minerals(self, minerals_df: pd.DataFrame | None) -> str:
        text, _ = self.write_and_verify_minerals(minerals_df)
        return text

    def write_and_verify_minerals(self, minerals_df: pd.DataFrame | None):
        fallback = "Dados de minerais não disponíveis."
        if minerals_df is None or minerals_df.empty:
            return fallback, None
        data_str = minerals_df.to_string(index=False)
        prompt = (
            f"{_SYSTEM}\n\n"
            "Com base nos dados abaixo, escreva 2 frases sobre os principais "
            "movimentos dos minerais e metais preciosos.\n\n"
            f"{data_str}\n\nResumo:"
        )
        return self._generate_and_verify("minerais_commentary", prompt, data_str, fallback)


# ── Module-level helper (avoids circular import) ──────────────────────────────

def WorkflowQAAgent_safe_commentary(section, commentary, qa_result, fallback):
    """Thin wrapper so DailyReportAgent doesn't import WorkflowQAAgent at class level."""
    try:
        from src.qa.qa_agent import WorkflowQAAgent
        return WorkflowQAAgent.safe_commentary(section, commentary, qa_result, fallback)
    except ImportError:
        return commentary, False
