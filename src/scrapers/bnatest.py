import re
import pandas as pd
from typing import Optional
from concurrent.futures import ThreadPoolExecutor
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://www.bna.ao/"


def _run_in_thread(fn):
    with ThreadPoolExecutor(max_workers=1) as ex:
        return ex.submit(fn).result()


def _clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def _normalize_pct(value: str) -> str:
    value = value.strip().replace(" ", "")
    return value


def _extract_section_percent(block: str, start_label: str, end_label: Optional[str] = None) -> str:
    flags = re.I | re.S
    if end_label:
        pattern = rf"{re.escape(start_label)}(.*?){re.escape(end_label)}"
    else:
        pattern = rf"{re.escape(start_label)}(.*)"
    m = re.search(pattern, block, flags)
    if not m:
        return "N/A"

    section_text = m.group(1)
    pct_match = re.search(r"\b\d{1,2}[,.]\d{1,3}\s*%", section_text)

    if pct_match:
        return _normalize_pct(pct_match.group())

    return "N/A"


def _parse_fx_from_text(text: str):
    results = []
    seen = set()

    fx_pattern = re.compile(r"\b(USD|EUR|ZAR)\b\s*[:\-]?\s*([\d\.,]+)", re.I)

    for code, value in fx_pattern.findall(text):
        code = code.upper()
        key = (code, value)
        if key not in seen:
            seen.add(key)
            results.append({
                "Moeda": code,
                "Taxa (AOA)": value.strip()
            })

    return results


def _parse_luibor_rows_from_text(text: str):
    rows = []
    seen = set()

    patterns = [
        re.compile(r"\b(O/N)\b\s*(\d{1,2}[,.]\d{1,3}\s*%)", re.I),
        re.compile(r"\b(1\s*M[eê]s|3\s*Meses|6\s*Meses|9\s*Meses|12\s*Meses)\b\s*(\d{1,2}[,.]\d{1,3}\s*%)", re.I),
    ]

    for pattern in patterns:
        for mat, rate in pattern.findall(text):
            mat = re.sub(r"\s+", " ", mat).strip()
            rate = _normalize_pct(rate)
            key = (mat.lower(), rate)
            if key not in seen:
                seen.add(key)
                rows.append({
                    "Maturidade": mat,
                    "Taxa (%)": rate
                })

    return rows


def _scrape_bna():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # set to True later
        page = browser.new_page(viewport={"width": 1400, "height": 2000})

        page.goto(URL, timeout=60000)
        page.wait_for_timeout(6000)

        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except PlaywrightTimeoutError:
            pass

        full_text = _clean_text(page.inner_text("body"))

        # --- Rates ---
        taxa_bna = _extract_section_percent(full_text, "Taxa BNA", "Taxa de Inflação")
        inflacao = _extract_section_percent(full_text, "Taxa de Inflação", "Taxa de Câmbio")

        fx_rows = _parse_fx_from_text(full_text)

        # --- LUIBOR ---
        luibor_rows = []
        seen = set()

        def collect():
            text = _clean_text(page.inner_text("body"))
            parsed = _parse_luibor_rows_from_text(text)

            for item in parsed:
                key = (item["Maturidade"], item["Taxa (%)"])
                if key not in seen:
                    seen.add(key)
                    luibor_rows.append(item)

        collect()

        # Try clicking carousel next button multiple times
        for _ in range(6):
            try:
                btn = page.locator("[class*='next'], .swiper-button-next").first
                if btn.is_visible():
                    btn.click()
                    page.wait_for_timeout(1500)
                    collect()
            except:
                break

        browser.close()

        return {
            "taxa_bna": taxa_bna,
            "inflacao": inflacao,
            "fx": fx_rows,
            "luibor": luibor_rows
        }


def get_exchange_rates():
    try:
        data = _run_in_thread(_scrape_bna)
        return pd.DataFrame(data["fx"])
    except Exception as e:
        print("FX error:", e)
        return pd.DataFrame()


def get_bna_rates():
    try:
        data = _run_in_thread(_scrape_bna)
        return {
            "taxa_bna": data["taxa_bna"],
            "inflacao": data["inflacao"]
        }
    except Exception as e:
        print("Rates error:", e)
        return {}


def get_luibor_rates():
    try:
        data = _run_in_thread(_scrape_bna)
        return pd.DataFrame(data["luibor"]).drop_duplicates()
    except Exception as e:
        print("LUIBOR error:", e)
        return pd.DataFrame()


# 🔥 TEST RUNNER
if __name__ == "__main__":
    print("\n--- BNA RATES ---")
    print(get_bna_rates())

    print("\n--- EXCHANGE RATES ---")
    print(get_exchange_rates())

    print("\n--- LUIBOR ---")
    print(get_luibor_rates())