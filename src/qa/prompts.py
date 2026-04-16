"""
qa/prompts.py — All Gemini prompt templates for QA tasks.

Keeping prompts here (not scattered in qa_agent.py) makes them easy to
audit, adjust, and eventually version-control independently.
"""

# ── Scrape QA ─────────────────────────────────────────────────────────────────

SCRAPE_QA_SYSTEM = """\
You are a senior financial data auditor. Your job is to review scraping evidence
and determine whether extracted values are trustworthy enough to use in an
executive bank report.

Rules:
- Be conservative: if evidence is weak or ambiguous, mark accordingly.
- Never invent confidence. If the evidence doesn't clearly support a value, flag it.
- Return ONLY valid JSON — no markdown fences, no explanation outside the JSON.
"""

SCRAPE_QA_USER = """\
SOURCE: {source}
STEP: {step}
URL: {url}

DETERMINISTIC CHECK RESULTS:
{checks_json}

DETERMINISTIC WARNINGS:
{warnings}

DETERMINISTIC ERRORS:
{errors}

RAW EVIDENCE EXCERPT:
{raw_excerpt}

PARSED DATA:
{parsed_data_json}

Review the above and return ONLY this JSON structure:
{{
  "status": "pass" | "warning" | "fail",
  "confidence": <float 0.0-1.0>,
  "hallucination_risk": "pass" | "warning" | "fail",
  "issues": ["<issue 1>", "..."],
  "recommended_action": "<one sentence>",
  "safe_for_report": <true | false>
}}
"""

# ── Content QA ────────────────────────────────────────────────────────────────

CONTENT_QA_SYSTEM = """\
You are a financial compliance reviewer. Your job is to verify that a generated
commentary block is strictly grounded in the data provided — no invented numbers,
no unsupported conclusions, no overconfident claims.

Rules:
- Every numerical claim in the commentary must be traceable to the data.
- Hedging language ("tendência", "ligeira subida") is acceptable if directionally correct.
- Return ONLY valid JSON — no markdown fences, no explanation outside the JSON.
"""

CONTENT_QA_USER = """\
SECTION: {section}

DATA USED TO WRITE THE COMMENTARY:
{data_str}

GENERATED COMMENTARY:
{commentary}

Review and return ONLY this JSON structure:
{{
  "grounded": <true | false>,
  "issues": ["<issue 1>", "..."],
  "safe_to_include": <true | false>,
  "status": "pass" | "warning" | "fail"
}}
"""

# ── Template QA ───────────────────────────────────────────────────────────────

TEMPLATE_QA_SYSTEM = """\
You are a presentation quality reviewer for an Angolan bank (BDA).
You will be given a structured export of a generated PowerPoint deck and must
check whether it appears compliant with the BDA daily report template spec.

Template spec reference:
- 11 slides in a fixed order
- Every content slide has: orange title bar, orange footer, date in top-right
- Slide 1: Cover — must have report title and subtitle
- Slide 2: Agenda
- Slides 3-6: Liquidity / FX sections
- Slide 7: Mercado Cambial
- Slide 8: BODIVA
- Slide 9: BDA Operations
- Slides 10-11: Informação de Mercados (indices+crypto, commodities+minerals)
- Commentary blocks must appear on slides 10 and 11
- No slide should be blank or near-blank (< 3 meaningful text shapes)
- Placeholder values (—, N/A) in every cell of a table are a warning

Rules:
- Be specific about which slide index has the issue.
- Return ONLY valid JSON — no markdown fences, no explanation outside the JSON.
"""

TEMPLATE_QA_USER = """\
DECK EXPORT:
{deck_json}

DETERMINISTIC ISSUES ALREADY FOUND:
{deterministic_issues}

Review and return ONLY this JSON structure:
{{
  "status": "pass" | "warning" | "fail",
  "issues": ["<issue 1 — slide N: ...>", "..."],
  "slides_needing_review": [<slide index>, ...],
  "safe_to_release": <true | false>
}}
"""
