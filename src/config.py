# Bank branding
BANK_NAME = "BDA"
BANK_COLOR_HEX = "#FF8C00"
BANK_COLOR_RGB = (255, 140, 0)
WHITE = (255, 255, 255)
DARK_TEXT = (33, 33, 33)
LIGHT_GRAY = (240, 240, 240)

# Report settings
REPORT_LANGUAGE = "pt_PT"
REPORT_TITLE = "Resumo Diário dos Mercados"
REPORT_SUBTITLE = "Direcção Financeira"
BANK_ADDRESS = "Edifício BDA, Condomínio Dolce Vita, Via S8, Talatona, Luanda - Angola"

# External data URLs
BNA_URL = "https://www.bna.ao/"
BODIVA_URL = "https://www.bodiva.ao/"
INE_URL = "https://www.ine.gov.ao/"

# Yahoo Finance tickers
INDICES_TICKERS = {
    "S&P 500": "^GSPC",
    "Dow Jones": "^DJI",
    "NASDAQ": "^IXIC",
    "Nikkei 225": "^N225",
    "IBOVESPA": "^BVSP",
    "Eurostoxx 50": "^STOXX50E",
    "PSI 20": "PSI20.LS",
    "Shanghai": "000001.SS",
    "Bolsa de Londres": "^FTSE",
}

COMMODITIES_TICKERS = {
    "Petróleo (Brent)": "BZ=F",
    "Ouro": "GC=F",
    "Cobre": "HG=F",
    "Milho": "ZC=F",
    "Soja": "ZS=F",
    "Trigo": "ZW=F",
    "Café": "KC=F",
    "Açúcar": "SB=F",
    "Algodão": "CT=F",
}

CRYPTO_TICKERS = {
    "Bitcoin (BTC)": "BTC-USD",
    "Ethereum (ETH)": "ETH-USD",
    "XRP": "XRP-USD",
}

# Gemini settings (replace GEMINI_API_KEY in .env)
GEMINI_MODEL = "gemini-2.0-flash"

# Cache duration in seconds (24 hours)
CACHE_TTL = 86400

# ── Timeouts & Retries ────────────────────────────────────────────────────────
REQUEST_TIMEOUT = 30
MAX_RETRIES     = 3
RETRY_BACKOFF   = 2

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # dotenv optional — run: pip install python-dotenv

import os

# ── Timeouts & Retries ────────────────────────────────────────────────
REQUEST_TIMEOUT  = int(os.getenv("REQUEST_TIMEOUT", 30))
MAX_RETRIES      = int(os.getenv("MAX_RETRIES", 3))

# ── URLs ──────────────────────────────────────────────────────────────
class URLs:
    BODIVA_HOME   = os.getenv("BODIVA_HOME_URL",   "https://www.bodiva.ao")
    BODIVA_MARKET = os.getenv("BODIVA_MARKET_URL", "https://www.bodiva.ao/mercado")
    BNA_HOME      = os.getenv("BNA_HOME_URL",      "https://www.bna.ao")
    YAHOO_BASE    = os.getenv("YAHOO_BASE_URL",    "https://query1.finance.yahoo.com")
