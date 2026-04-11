import re
import pandas as pd
from typing import Optional
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://www.bna.ao/"
GOTO_TIMEOUT_MS = 15_000
LOAD_STATE_TIMEOUT_MS = 5_000

_cached_data = None


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def normalize_pct(value: str) -> str:
    return re.sub(r"\s+", "", value.strip())


def extract_section_percent(block: str, start_label: str, end_label: Optional[str] = None) -> str:
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

    if pct_match:
        return normalize_pct(pct_match.group())

    return "N/A"


def parse_fx_from_text(text: str) -> list:
    seen = set()
    results = []

    fx_pattern = re.compile(r"\b(USD|EUR|ZAR)\b\s*[:\-]?\s*([\d\.,]+)", re.I)

    for code, value in fx_pattern.findall(text):
        code = code.upper()
        value = value.strip()
        key = (code, value)
        if key not in seen:
            seen.add(key)
            results.append({
                "Moeda": code,
                "Taxa (AOA)": value
            })

    desired = ["USD", "EUR", "ZAR"]
    ordered = []
    for currency in desired:
        found = next((r for r in results if r["Moeda"] == currency), None)
        if found:
            ordered.append(found)

    return ordered


def standardize_luibor_maturity(text: str) -> str:
    text = clean_text(text)

    mapping = {
        "Overnight": "Overnight",
        "O/N": "Overnight",
        "1 Mês": "1 Mês",
        "1 Mes": "1 Mês",
        "3 Meses": "3 Meses",
        "6 Meses": "6 Meses",
        "9 Meses": "9 Meses",
        "12 Meses": "12 Meses",
    }
    return mapping.get(text, text)


def target_luibor_section(page):
    selectors = [
        "section:has-text('LUIBOR')",
        "div:has-text('LUIBOR')",
        "article:has-text('LUIBOR')",
        "[id*='luibor' i]",
        "[class*='luibor' i]",
    ]

    for selector in selectors:
        locator = page.locator(selector)
        if locator.count() > 0:
            return locator.first

    raise Exception("LUIBOR section not found")


def extract_luibor_rows(section) -> list:
    rows = []
    seen = set()

    tr_locator = section.locator("tr")
    tr_count = tr_locator.count()

    for i in range(tr_count):
        cells = tr_locator.nth(i).locator("td")
        if cells.count() >= 2:
            maturity = clean_text(cells.nth(0).inner_text())
            rate = clean_text(cells.nth(1).inner_text())

            if maturity and rate and "%" in rate:
                maturity = standardize_luibor_maturity(maturity)
                key = (maturity, rate)
                if key not in seen:
                    seen.add(key)
                    rows.append({
                        "Maturidade": maturity,
                        "Taxa (%)": rate
                    })

    if not rows:
        section_text = clean_text(section.inner_text())

        patterns = [
            re.compile(r"(Overnight)\s+(\d{1,2},\d{2}%)", re.I),
            re.compile(r"(1\s*M[eê]s|3\s*Meses|6\s*Meses|9\s*Meses|12\s*Meses)\s+(\d{1,2},\d{2}%)", re.I),
        ]

        for pattern in patterns:
            for maturity, rate in pattern.findall(section_text):
                maturity = standardize_luibor_maturity(maturity)
                key = (maturity, rate)
                if key not in seen:
                    seen.add(key)
                    rows.append({
                        "Maturidade": maturity,
                        "Taxa (%)": rate
                    })

    return rows


def click_luibor_next(section, page) -> bool:
    selectors = [
        ".swiper-button-next",
        ".slick-next",
        ".owl-next",
        "[aria-label*='next' i]",
        "[aria-label*='right' i]",
        "[class*='next' i]",
        "[class*='right' i]",
    ]

    for selector in selectors:
        btn = section.locator(selector).first
        try:
            if btn.count() > 0 and btn.is_visible():
                btn.click(force=True)
                page.wait_for_timeout(1500)
                return True
        except Exception:
            pass

    try:
        pager = section.locator(":text-matches('1/2|2/2', 'i')").first
        if pager.count() > 0:
            box = pager.bounding_box()
            if box:
                page.mouse.click(box["x"] + box["width"] + 25, box["y"] + box["height"] / 2)
                page.wait_for_timeout(1500)
                return True
    except Exception:
        pass

    return False


def sort_luibor_rows(rows: list) -> list:
    order = {
        "Overnight": 0,
        "1 Mês": 1,
        "3 Meses": 2,
        "6 Meses": 3,
        "9 Meses": 4,
        "12 Meses": 5,
    }
    return sorted(rows, key=lambda x: order.get(x["Maturidade"], 99))


def scrape_bna() -> dict:
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(viewport={"width": 1440, "height": 2200})

            # BNA can be slow or intermittently unresponsive. Use a quicker
            # commit-based navigation first so we can recover fast and keep the
            # report workflow moving if the page takes too long to finish.
            try:
                page.goto(URL, timeout=GOTO_TIMEOUT_MS, wait_until="commit")
            except PlaywrightTimeoutError:
                try:
                    page.goto(URL, timeout=GOTO_TIMEOUT_MS, wait_until="domcontentloaded")
                except PlaywrightTimeoutError:
                    # Best-effort fallback: keep going if the browser has any
                    # partial content already. This is still better than a hard fail.
                    pass

            page.wait_for_timeout(4000)

            try:
                page.wait_for_load_state("networkidle", timeout=LOAD_STATE_TIMEOUT_MS)
            except PlaywrightTimeoutError:
                pass

            full_text = clean_text(page.inner_text("body"))

            taxa_bna = extract_section_percent(full_text, "Taxa BNA", "Taxa de Inflação")
            inflacao = extract_section_percent(full_text, "Taxa de Inflação", "Taxa de Câmbio")
            fx_rows = parse_fx_from_text(full_text)

            try:
                luibor_section = target_luibor_section(page)
                all_luibor_rows = extract_luibor_rows(luibor_section)

                clicked = click_luibor_next(luibor_section, page)
                if clicked:
                    page2_rows = extract_luibor_rows(luibor_section)

                    existing = {(r["Maturidade"], r["Taxa (%)"]) for r in all_luibor_rows}
                    for row in page2_rows:
                        key = (row["Maturidade"], row["Taxa (%)"])
                        if key not in existing:
                            existing.add(key)
                            all_luibor_rows.append(row)
            except Exception:
                all_luibor_rows = []

            browser.close()

        return {
            "taxa_bna": taxa_bna,
            "inflacao": inflacao,
            "fx": fx_rows,
            "luibor": sort_luibor_rows(all_luibor_rows),
        }
    except Exception:
        # Hard fallback: return safe placeholders so the rest of the report can
        # still generate if BNA is temporarily unavailable.
        return {
            "taxa_bna": "N/A",
            "inflacao": "N/A",
            "fx": [],
            "luibor": [],
        }


def scrape_once(force_refresh: bool = False) -> dict:
    global _cached_data
    if force_refresh or _cached_data is None:
        _cached_data = scrape_bna()
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
            "Maturidade": ["Overnight", "1 Mês", "3 Meses", "6 Meses", "9 Meses", "12 Meses"],
            "Taxa (%)": ["N/A"] * 6
        })

    return pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)


def get_all_bna_data(force_refresh: bool = False) -> dict:
    data = scrape_once(force_refresh=force_refresh)

    return {
        "taxa_bna": data.get("taxa_bna", "N/A"),
        "inflacao": data.get("inflacao", "N/A"),
        "exchange_rates": pd.DataFrame(data.get("fx", [])),
        "luibor_rates": pd.DataFrame(data.get("luibor", [])),
    }


if __name__ == "__main__":
    print("\n--- BNA RATES ---")
    print(get_bna_rates(force_refresh=True))

    print("\n--- EXCHANGE RATES ---")
    print(get_exchange_rates())

    print("\n--- LUIBOR RATES ---")
    print(get_luibor_rates())

    print("\n--- ALL DATA ---")
    all_data = get_all_bna_data()
    print("Taxa BNA:", all_data["taxa_bna"])
    print("Inflação:", all_data["inflacao"])
    print("\nExchange Rates:")
    print(all_data["exchange_rates"])
    print("\nLUIBOR Rates:")
    print(all_data["luibor_rates"])
