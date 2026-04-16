# Teko BDA — Resumo Diário dos Mercados

Automated daily executive report system for **BDA (Banco de Desenvolvimento de Angola)**.
Scrapes external financial data, accepts internal treasury inputs, and generates a branded
11-slide PowerPoint that mirrors the bank's manual daily report.

---

## What it does

| Step | What happens |
|------|-------------|
| 1 | Scrape external data — BNA (LUIBOR, FX, inflation), Yahoo Finance (indices, commodities, crypto) |
| 2 | User enters internal treasury data in the Streamlit UI |
| 3 | Optional: AI generates Portuguese commentary for each section |
| 4 | One click generates a fully branded `.pptx` matching the BDA template |
| 5 | Download the PPTX directly from the browser |

---

## Quick Start

### 1. Clone & install

```bash
git clone <repo-url>
cd teko-bda-daily-report
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium        # for BNA scraper
```

### 2. Configure environment

```bash
cp .env.example .env
# Edit .env and add your provider API key (optional — report works without it)
```

### 3. Run

```bash
streamlit run streamlit_app/app.py
```

Open [http://localhost:8501](http://localhost:8501).

---

## Project Structure

```
teko-bda-daily-report/
│
├── streamlit_app/
│   └── app.py                  ← Streamlit frontend (3 tabs)
│
├── src/
│   ├── config.py               ← Constants, URLs, ticker lists
│   │
│   ├── scrapers/
│   │   ├── base_scraper.py     ← Abstract base with session + caching
│   │   ├── bna_scraper.py      ← BNA: LUIBOR, FX, inflation (Playwright)
│   │   ├── bodiva_scraper.py   ← BODIVA stocks + segments (JS site — partial)
│   │   ├── yahoo_scraper.py    ← Global indices, commodities, crypto (yfinance)
│   │   └── market_aggregator.py← Combines all external sources
│   │
│   ├── report_generator/
│   │   └── pptx_builder.py     ← 11-slide BDA PPTX builder (BDAReportGenerator)
│   │
│   ├── agents/
│   │   └── ai_agent.py         ← High-level commentary agent
│   │
│   ├── llm/
│   │   └── llm_client.py       ← Provider-agnostic LLM router
│   │
│   └── utils/
│       ├── logger.py
│       └── helpers.py
│
├── output/                     ← Generated PPTX files (git-ignored)
├── .env.example                ← Template for environment variables
├── requirements.txt
└── README.md
```

---

## LLM Setup

The AI commentary is **optional** — the report generates fully without it.

Default provider is OpenAI:

```env
LLM_PROVIDER=openai
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-5.4-mini
```

Gemini remains supported:

```env
LLM_PROVIDER=gemini
GEMINI_API_KEY=AIza...
GEMINI_MODEL=gemini-2.0-flash
```

Switch providers by changing only `LLM_PROVIDER` in `.env`.

---

## Report Slides

| # | Slide | Data source |
|---|-------|-------------|
| 1 | Cover | date |
| 2 | Agenda | static |
| 3 | Sumário Executivo | internal KPIs |
| 4 | Liquidez – Moeda Nacional (1/2) | internal + BNA LUIBOR |
| 5 | Liquidez – Moeda Nacional (2/2) | internal cash-flow |
| 6 | Liquidez – Moeda Estrangeira | internal |
| 7 | Mercado Cambial | internal FX + BNA |
| 8 | Mercado de Capitais – BODIVA | BODIVA scraper |
| 9 | Mercado de Capitais – Operações BDA | internal portfolio |
| 10 | Informação de Mercados (1/2) | Yahoo indices + crypto + optional AI |
| 11 | Informação de Mercados (2/2) | Yahoo commodities + minerals + optional AI |

---

## Known Limitations

- **BODIVA scraper** — BODIVA's website is heavily JavaScript-rendered.
  The Playwright path works when the site loads correctly, but is fragile.
  When it fails the slide renders with `—` placeholders and a clear warning.
  Internal BODIVA data can be entered manually via `bodiva_stocks` in the data dict.

- **Charts/images** — The original PDF includes pie charts and line charts.
  These are not yet generated programmatically; the slides contain tables only.

- **BDA logo** — The logo image is not embedded (not available as a file in the repo).
  The BDA text label is used as a placeholder in the bottom-right corner of the cover.

---

## For Malcolm (and future maintainers)

- All LLM calls go through `src/llm/llm_client.py`. Provider selection is controlled by `LLM_PROVIDER` in `.env`.
- To add a new data source: create a new scraper inheriting `BaseScraper`, add it to `market_aggregator.py`, then map its output into the `data` dict in `app.py`.
- To add a new slide: add a `_slide_xyz()` method to `BDAReportGenerator` and call it in `build()`.
- The full `data` dict schema is documented in the `BDAReportGenerator` class docstring in `src/report_generator/pptx_builder.py`.
