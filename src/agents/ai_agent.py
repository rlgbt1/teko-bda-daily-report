"""
agents/ai_agent.py — High-level AI commentary agent backed by Gemini.

Wraps src/llm/llm_client.py with domain-specific prompt building so the
Streamlit app and report builder can call simple, named methods.

Usage:
    from src.agents.ai_agent import DailyReportAgent
    agent = DailyReportAgent()
    text  = agent.summarize_markets(markets_df)
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

    All methods return a plain string.  If Gemini is unavailable (missing key,
    network error, etc.) a safe Portuguese fallback is returned instead of
    raising an exception.
    """

    # ── Markets ───────────────────────────────────────────────────────────────

    def summarize_markets(self, markets_df: pd.DataFrame | None) -> str:
        """3-4 sentence summary of global equity / bond indices."""
        if markets_df is None or markets_df.empty:
            return "Dados de mercado não disponíveis."
        data_str = markets_df.to_string(index=False)
        prompt = (
            f"{_SYSTEM}\n\n"
            "Com base nos dados abaixo, escreva um resumo de 3-4 frases sobre os "
            "principais movimentos dos mercados de capitais globais. "
            "Destaque os movimentos mais significativos (maiores altas e baixas).\n\n"
            f"{data_str}\n\nResumo:"
        )
        return generate_commentary(
            prompt, fallback="Dados de mercados globais não disponíveis."
        )

    # ── Commodities ───────────────────────────────────────────────────────────

    def summarize_commodities(self, commodities_df: pd.DataFrame | None) -> str:
        """2-3 sentence summary of commodities and minerals."""
        if commodities_df is None or commodities_df.empty:
            return "Dados de commodities não disponíveis."
        data_str = commodities_df.to_string(index=False)
        prompt = (
            f"{_SYSTEM}\n\n"
            "Com base nos dados abaixo, escreva 2-3 frases sobre os principais "
            "movimentos das commodities e minerais. Mencione petróleo, ouro e cobre "
            "se presentes.\n\n"
            f"{data_str}\n\nResumo:"
        )
        return generate_commentary(
            prompt, fallback="Dados de commodities não disponíveis."
        )

    # ── Crypto ────────────────────────────────────────────────────────────────

    def summarize_crypto(self, crypto_df: pd.DataFrame | None) -> str:
        """2 sentence summary of cryptocurrency movements."""
        if crypto_df is None or crypto_df.empty:
            return "Dados de criptomoedas não disponíveis."
        data_str = crypto_df.to_string(index=False)
        prompt = (
            f"{_SYSTEM}\n\n"
            "Com base nos dados abaixo, escreva 2 frases sobre o mercado de "
            "criptomoedas, mencionando Bitcoin, Ethereum e tendência geral.\n\n"
            f"{data_str}\n\nResumo:"
        )
        return generate_commentary(
            prompt, fallback="Dados de criptomoedas não disponíveis."
        )

    # ── FX / Cambial ──────────────────────────────────────────────────────────

    def summarize_fx(self, fx_data: dict | None) -> str:
        """1-2 sentence comment on FX moves (USD/AOA, EUR/AOA)."""
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
        """1-2 sentence summary of BDA's liquidity position."""
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
