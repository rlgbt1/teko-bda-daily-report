import re
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://www.bna.ao/"


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def standardize_maturity(text: str) -> str:
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

    # First try proper table rows
    tr_locator = section.locator("tr")
    tr_count = tr_locator.count()

    for i in range(tr_count):
        cells = tr_locator.nth(i).locator("td")
        if cells.count() >= 2:
            maturity = clean_text(cells.nth(0).inner_text())
            rate = clean_text(cells.nth(1).inner_text())

            if maturity and rate and "%" in rate:
                maturity = standardize_maturity(maturity)
                key = (maturity, rate)
                if key not in seen:
                    seen.add(key)
                    rows.append({
                        "Maturidade": maturity,
                        "Taxa (%)": rate
                    })

    # Fallback: parse visible text blocks if table tags are not present
    if not rows:
        section_text = clean_text(section.inner_text())

        patterns = [
            re.compile(r"(Overnight)\s+(\d{1,2},\d{2}%)", re.I),
            re.compile(r"(1\s*M[eê]s|3\s*Meses|6\s*Meses|9\s*Meses|12\s*Meses)\s+(\d{1,2},\d{2}%)", re.I),
        ]

        for pattern in patterns:
            for maturity, rate in pattern.findall(section_text):
                maturity = standardize_maturity(maturity)
                key = (maturity, rate)
                if key not in seen:
                    seen.add(key)
                    rows.append({
                        "Maturidade": maturity,
                        "Taxa (%)": rate
                    })

    return rows


def click_luibor_next(section, page) -> bool:
    # We try the most likely right-arrow selectors inside the LUIBOR section first
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

    # Fallback: click the right chevron text/icon area near the 1/2 pager
    try:
        pager = section.locator(":text-matches('1/2|2/2', 'i')").first
        if pager.count() > 0:
            box = pager.bounding_box()
            if box:
                # click slightly to the right of the pager text where the arrow usually sits
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


def get_luibor_rates() -> pd.DataFrame:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 2200})

        page.goto(URL, timeout=60000, wait_until="domcontentloaded")
        page.wait_for_timeout(5000)

        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except PlaywrightTimeoutError:
            pass

        luibor_section = target_luibor_section(page)

        # Page 1
        all_rows = extract_luibor_rows(luibor_section)

        # Click to page 2
        clicked = click_luibor_next(luibor_section, page)

        if clicked:
            page2_rows = extract_luibor_rows(luibor_section)

            existing = {(r["Maturidade"], r["Taxa (%)"]) for r in all_rows}
            for row in page2_rows:
                key = (row["Maturidade"], row["Taxa (%)"])
                if key not in existing:
                    existing.add(key)
                    all_rows.append(row)

        browser.close()

    return pd.DataFrame(sort_luibor_rows(all_rows))


if __name__ == "__main__":
    df = get_luibor_rates()
    print(df)