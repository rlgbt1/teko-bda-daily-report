import re
import pandas as pd
from typing import Optional
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://www.bna.ao/"

_cached_data = None


def _clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def _normalize_pct(value: str) -> str:
    return re.sub(r"\s+", "", value.strip())


def _extract_section_percent(block: str, start_label: str, end_label: Optional[str] = None) -> str:
    flags = re.I | re.S
    if end_label:
        pattern = rf"{re.escape(start_label)}(.*?){re.escape(end_label)}"
    else:
        pattern = rf"{re.escape(start_label)}(.*)"
    match = re.search(pattern, block, flags)
    if not match:
        return "N/A"

    section_text = match.group(1)
    pct_match = re.search(r"\b\d{1,2}[,.]\d{1,3}\s*%", section_text)
    return _normalize_pct(pct_match.group()) if pct_match else "N/A"


def _parse_fx_from_text(text: str) -> list:
    seen = set()
    results = []

    fx_pattern = re.compile(r"\b(USD|EUR|ZAR)\b\s*[:\-]?\s*([\d\.,]+)", re.I)

    for code, value in fx_pattern.findall(text):
        code = code.upper()
        value = value.strip()
        key = (code, value)
        if key not in seen:
            seen.add(key)
            results.append({"Moeda": code, "Taxa (AOA)": value})

    desired = ["USD", "EUR", "ZAR"]
    ordered = []
    for currency in desired:
        found = next((r for r in results if r["Moeda"] == currency), None)
        if found:
            ordered.append(found)
    return ordered


def _standardize_maturity(mat: str) -> str:
    mat = _clean_text(mat)
    mat = mat.replace("Mês", "Mes").replace("mês", "mes")

    mapping = {
        "O/N": "O/N",
        "ON": "O/N",
        "1 Mes": "1M",
        "1 Mês": "1M",
        "1M": "1M",
        "3 Meses": "3M",
        "3M": "3M",
        "6 Meses": "6M",
        "6M": "6M",
        "9 Meses": "9M",
        "9M": "9M",
        "12 Meses": "12M",
        "12M": "12M",
    }
    return mapping.get(mat, mat)


def _parse_luibor_rows_from_text(text: str) -> list:
    text = _clean_text(text)
    seen = set()
    rows = []

    patterns = [
        re.compile(r"\b(O/N)\b\s*[:\-]?\s*(\d{1,2}[,.]\d{1,3}\s*%)", re.I),
        re.compile(r"\b(1\s*M[eê]s|3\s*Meses|6\s*Meses|9\s*Meses|12\s*Meses)\b\s*[:\-]?\s*(\d{1,2}[,.]\d{1,3}\s*%)", re.I),
        re.compile(r"\b(1M|3M|6M|9M|12M)\b\s*[:\-]?\s*(\d{1,2}[,.]\d{1,3}\s*%)", re.I),
    ]

    for pattern in patterns:
        for mat, rate in pattern.findall(text):
            maturity = _standardize_maturity(mat)
            pct = _normalize_pct(rate)
            key = (maturity, pct)
            if key not in seen:
                seen.add(key)
                rows.append({"Maturidade": maturity, "Taxa (%)": pct})

    return rows


def _target_luibor_section(page):
    candidates = [
        "section:has-text('LUIBOR')",
        "div:has-text('LUIBOR')",
        "article:has-text('LUIBOR')",
        "[class*='luibor' i]",
        "[id*='luibor' i]",
    ]

    for selector in candidates:
        locator = page.locator(selector)
        if locator.count() > 0:
            return locator.first

    return page.locator("body")


def _collect_luibor_from_section(section) -> list:
    seen = set()
    rows = []

    try:
        section_text = _clean_text(section.inner_text())
        for item in _parse_luibor_rows_from_text(section_text):
            key = (item["Maturidade"], item["Taxa (%)"])
            if key not in seen:
                seen.add(key)
                rows.append(item)
    except Exception:
        pass

    try:
        tr_locator = section.locator("tr")
        for i in range(tr_locator.count()):
            row_text = _clean_text(tr_locator.nth(i).inner_text())
            parsed = _parse_luibor_rows_from_text(row_text)
            for item in parsed:
                key = (item["Maturidade"], item["Taxa (%)"])
                if key not in seen:
                    seen.add(key)
                    rows.append(item)
    except Exception:
        pass

    try:
        item_locator = section.locator("li, .item, .slide, .swiper-slide, .slick-slide, .owl-item")
        for i in range(min(item_locator.count(), 30)):
            block_text = _clean_text(item_locator.nth(i).inner_text())
            parsed = _parse_luibor_rows_from_text(block_text)
            for item in parsed:
                key = (item["Maturidade"], item["Taxa (%)"])
                if key not in seen:
                    seen.add(key)
                    rows.append(item)
    except Exception:
        pass

    return rows


def _merge_luibor_rows(existing: list, new_rows: list) -> list:
    seen = {(row["Maturidade"], row["Taxa (%)"]) for row in existing}
    for row in new_rows:
        key = (row["Maturidade"], row["Taxa (%)"])
        if key not in seen:
            seen.add(key)
            existing.append(row)
    return existing


def _click_luibor_next(section, page) -> bool:
    next_selectors = [
        ".swiper-button-next",
        ".slick-next",
        ".owl-next",
        "[aria-label*='next' i]",
        "[aria-label*='seguinte' i]",
        "button:has-text('>')",
        "[class*='next' i]",
        "[class*='right' i]",
    ]

    for selector in next_selectors:
        btn = section.locator(selector).first
        try:
            if btn.count() > 0 and btn.is_visible():
                btn.click(timeout=2000, force=True)
                page.wait_for_timeout(1200)
                return True
        except Exception:
            continue

    return False


def _sort_luibor_rows(rows: list) -> list:
    order = {"O/N": 0, "1M": 1, "3M": 2, "6M": 3, "9M": 4, "12M": 5}
    return sorted(rows, key=lambda x: order.get(x["Maturidade"], 99))


def _scrape_bna() -> dict:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 2200})

        page.goto(URL, timeout=60000, wait_until="domcontentloaded")
        page.wait_for_timeout(5000)

        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except PlaywrightTimeoutError:
            pass

        full_text = _clean_text(page.inner_text("body"))

        taxa_bna = _extract_section_percent(full_text, "Taxa BNA", "Taxa de Inflação")
        inflacao = _extract_section_percent(full_text, "Taxa de Inflação", "Taxa de Câmbio")
        fx_rows = _parse_fx_from_text(full_text)

        luibor_section = _target_luibor_section(page)
        luibor_rows = _collect_luibor_from_section(luibor_section)

        max_clicks = 10
        for _ in range(max_clicks):
            before = len(luibor_rows)
            clicked = _click_luibor_next(luibor_section, page)
            if not clicked:
                break

            new_rows = _collect_luibor_from_section(luibor_section)
            luibor_rows = _merge_luibor_rows(luibor_rows, new_rows)

            if len(luibor_rows) == before:
                break

            found_maturities = {row["Maturidade"] for row in luibor_rows}
            if {"O/N", "1M", "3M", "6M", "9M", "12M"}.issubset(found_maturities):
                break

        browser.close()

    return {
        "taxa_bna": taxa_bna,
        "inflacao": inflacao,
        "fx": fx_rows,
        "luibor": _sort_luibor_rows(luibor_rows),
    }


def scrape_once(force_refresh: bool = False) -> dict:
    global _cached_data
    if force_refresh or _cached_data is None:
        _cached_data = _scrape_bna()
    return _cached_data


def get_bna_rates(force_refresh: bool = False) -> dict:
    data = scrape_once(force_refresh=force_refresh)
    return {
        "taxa_bna": data.get("taxa_bna", "N/A"),
        "inflacao": data.get("inflacao", "N/A"),
    }


def get_exchange_rates(force_refresh: bool = False) -> pd.DataFrame:
    data = scrape_once(force_refresh=force_refresh)
    rows = data.get("fx", [])
    if not rows:
        return pd.DataFrame({
            "Moeda": ["USD", "EUR", "ZAR"],
            "Taxa (AOA)": ["N/A", "N/A", "N/A"]
        })
    return pd.DataFrame(rows)


def get_luibor_rates(force_refresh: bool = False) -> pd.DataFrame:
    data = scrape_once(force_refresh=force_refresh)
    rows = data.get("luibor", [])
    if not rows:
        return pd.DataFrame({
            "Maturidade": ["O/N", "1M", "3M", "6M", "9M", "12M"],
            "Taxa (%)": ["N/A"] * 6
        })
    return pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)


if __name__ == "__main__":
    print("\n--- BNA RATES ---")
    print(get_bna_rates(force_refresh=True))

    print("\n--- EXCHANGE RATES ---")
    print(get_exchange_rates())

    print("\n--- LUIBOR ---")
    print(get_luibor_rates())