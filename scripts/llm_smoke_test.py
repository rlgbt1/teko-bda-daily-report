"""
Small smoke test for the provider-agnostic LLM layer.

Examples:
    python scripts/llm_smoke_test.py text
    python scripts/llm_smoke_test.py json
    LLM_PROVIDER=gemini python scripts/llm_smoke_test.py text
"""
from __future__ import annotations

import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from src.llm.llm_client import generate_commentary, generate_json, get_provider_name


def run_text() -> int:
    prompt = (
        "Escreva 2 frases curtas em português europeu sobre estes dados. "
        "Use apenas a informação dada.\n\n"
        "S&P 500: +1.2%\nNASDAQ: -0.4%\nBitcoin: +3.1%"
    )
    text = generate_commentary(prompt, fallback="FALLBACK")
    print(f"provider={get_provider_name()}")
    print(text)
    return 0


def run_json() -> int:
    prompt = """Return only JSON with keys:
{
  "status": "ok" | "warning",
  "summary": "short Portuguese sentence"
}

Data:
S&P 500 +1.2%, Bitcoin +3.1%
"""
    data = generate_json(
        prompt,
        fallback={"status": "warning", "summary": "Fallback usado."},
    )
    print(f"provider={get_provider_name()}")
    print(json.dumps(data, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "text"
    raise SystemExit(run_json() if mode == "json" else run_text())
