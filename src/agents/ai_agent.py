"""
agents/ai_agent.py — High-level AI commentary agent.

Wraps src/llm/llm_client.py with domain-specific prompt building so the
Streamlit app and report builder can call simple, named methods.

Roles
-----
DailyReportAgent  — writer: generates Portuguese commentary blocks
WorkflowQAAgent   — checker: see src/qa/qa_agent.py

Content QA flow
---------------
Each write_and_verify_* method:
  1. generates commentary via the active LLM provider
  2. runs content QA (WorkflowQAAgent.review_commentary)
  3. returns the commentary only if it is safe_to_include
  4. returns a safe Portuguese fallback if QA fails or the provider is unavailable

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
    "Use APENAS os dados fornecidos — nunca invente valores ou factos. "
    "O texto deve soar como um resumo executivo de relatório bancário, "
    "sem floreados, sem listas e sem títulos desnecessários."
)

_STYLE_TEMPLATE = (
    "Estilo desejado:\n"
    "- escreva em 2 a 4 parágrafos curtos\n"
    "- mantenha tom executivo, neutro e objectivo\n"
    "- destaque os movimentos mais importantes com percentagens dentro da frase\n"
    "- não use bullets, numeração, tabelas ou markdown\n"
    "- não repita os mesmos dados duas vezes\n"
    "- use expressões como 'subiu', 'caiu', 'avançou', 'recuou', 'manteve-se estável'\n"
    "- se houver vários blocos geográficos, organize por EUA, Ásia, Europa, América Latina\n"
    "- se houver commodities/minerais, separe o comentário por grupos afins"
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
            "Com base nos dados abaixo, escreva o comentário do slide de "
            "INFORMAÇÃO DE MERCADOS (1/2).\n"
            "Faça um resumo executivo curto dos mercados globais, tal como no "
            "template de referência: EUA, Ásia, Europa e América Latina devem "
            "ser tratados em blocos separados quando os dados existirem.\n"
            "Destaque os maiores ganhos e perdas com percentagens dentro do texto.\n\n"
            f"{_STYLE_TEMPLATE}\n\n"
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
            "Com base nos dados abaixo, escreva o comentário do slide de "
            "INFORMAÇÃO DE MERCADOS (2/2).\n"
            "O texto deve cobrir commodities e minerais no mesmo estilo do "
            "template de referência: primeiro um resumo curto dos minerais e "
            "metais preciosos, depois um resumo curto das commodities agrícolas e "
            "energéticas.\n"
            "Use frases curtas, factuais e orientadas ao movimento percentual.\n\n"
            f"{_STYLE_TEMPLATE}\n\n"
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
            "Com base nos dados abaixo, escreva um comentário curto de cripto "
            "para o bloco de INFORMAÇÃO DE MERCADOS (1/2). "
            "O texto deve ser um parágrafo executivo com 2 a 3 frases, "
            "mencionando Bitcoin e outras moedas relevantes quando existirem. "
            "Se os dados mostrarem estabilidade em stablecoins, diga isso de forma breve.\n\n"
            f"{_STYLE_TEMPLATE}\n\n"
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
            "Com base nos dados cambiais abaixo, escreva 1 a 2 frases sobre a "
            "evolução do kwanza face ao dólar e euro, em tom executivo e objectivo.\n\n"
            f"{_STYLE_TEMPLATE}\n\n"
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
            "Com base na posição de liquidez abaixo, escreva 1 a 2 frases de "
            "enquadramento executivo para um slide bancário.\n\n"
            f"{_STYLE_TEMPLATE}\n\n"
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
            "Com base nos dados abaixo, escreva o comentário curto de minerais "
            "para o slide de INFORMAÇÃO DE MERCADOS (2/2). "
            "Use um tom semelhante ao template de referência: destaque o ouro, "
            "ferro, cobre e manganês quando presentes, e diga o que subiu, caiu "
            "ou permaneceu estável.\n\n"
            f"{_STYLE_TEMPLATE}\n\n"
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
