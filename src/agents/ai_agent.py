import re

try:
    from langchain.llms import Ollama
    OLLAMA_AVAILABLE = True
except ImportError:
    OLLAMA_AVAILABLE = False

from src.config import OLLAMA_BASE_URL, OLLAMA_MODEL


class DailyReportAgent:
    def __init__(self):
        self.llm = None
        if OLLAMA_AVAILABLE:
            try:
                self.llm = Ollama(model=OLLAMA_MODEL, base_url=OLLAMA_BASE_URL)
            except Exception:
                self.llm = None

    def _call_llm(self, prompt: str) -> str:
        if self.llm is None:
            return "Resumo automático não disponível (Ollama não activo)."
        try:
            return self.llm(prompt)
        except Exception:
            return "Resumo automático não disponível."

    def summarize_markets(self, markets_df) -> str:
        if markets_df is None or markets_df.empty:
            return "Dados de mercado não disponíveis."
        data_str = markets_df.to_string(index=False)
        prompt = f"""Você é um analista financeiro de um banco angolano. 
Com base nos dados abaixo, escreva um resumo de 3-4 frases em português sobre os principais movimentos dos mercados globais. 
Use APENAS os números fornecidos. NÃO invente dados.

{data_str}

Resumo:"""
        return self._call_llm(prompt)

    def summarize_commodities(self, commodities_df) -> str:
        if commodities_df is None or commodities_df.empty:
            return "Dados de commodities não disponíveis."
        data_str = commodities_df.to_string(index=False)
        prompt = f"""Com base nos dados abaixo, escreva 2-3 frases em português sobre os principais movimentos das commodities.
Use APENAS os números fornecidos.

{data_str}

Resumo:"""
        return self._call_llm(prompt)

    def summarize_crypto(self, crypto_df) -> str:
        if crypto_df is None or crypto_df.empty:
            return "Dados de criptomoedas não disponíveis."
        data_str = crypto_df.to_string(index=False)
        prompt = f"""Com base nos dados abaixo, escreva 2 frases em português sobre o mercado de criptomoedas.
Use APENAS os números fornecidos.

{data_str}

Resumo:"""
        return self._call_llm(prompt)

    def verify_no_hallucinations(self, text: str, source_df) -> bool:
        """Basic check: ensure no numbers in text are invented."""
        if source_df is None or source_df.empty:
            return True
        numbers_in_text = re.findall(r"\d+[\.,]?\d*", text)
        source_values = source_df.to_string()
        for num in numbers_in_text:
            clean = num.replace(",", ".")
            if clean not in source_values:
                return False
        return True
