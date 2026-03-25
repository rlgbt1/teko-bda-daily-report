import re
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from playwright.sync_api import sync_playwright


def _scrape_bna():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://www.bna.ao/", timeout=30000, wait_until="domcontentloaded")
        page.wait_for_timeout(5000)
        body = page.inner_text("body")

        # Collect LUIBOR rows across both carousel pages
        luibor_rows = []
        for _ in range(2):
            rows = page.query_selector_all("tr")
            for row in rows:
                cells = row.query_selector_all("td")
                if len(cells) >= 2:
                    mat = cells[0].inner_text().strip()
                    taxa = cells[1].inner_text().strip()
                    if "Meses" in mat or "O/N" in mat or "Dias" in mat:
                        luibor_rows.append({"Maturidade": mat, "Taxa (%)": taxa})
            # Try clicking prev/next arrow in LUIBOR carousel
            for selector in [".prev", "[class*='prev']", "[class*='esquerda']", "[class*='left']"]:
                btn = page.query_selector(selector)
                if btn:
                    btn.click()
                    page.wait_for_timeout(1500)
                    break

        browser.close()
        return body, luibor_rows


def _run_in_thread(fn):
    with ThreadPoolExecutor(max_workers=1) as ex:
        return ex.submit(fn).result()


def get_exchange_rates() -> pd.DataFrame:
    try:
        body, _ = _run_in_thread(_scrape_bna)
        lines = body.split("\n")
        results = []
        for line in lines:
            # Lines like "USD\t912,131"
            for currency in ["USD", "EUR", "ZAR", "GBP"]:
                if line.strip().startswith(currency) and "\t" in line:
                    parts = line.strip().split("\t")
                    if len(parts) == 2:
                        results.append({"Moeda": parts[0].strip(), "Taxa (AOA)": parts[1].strip()})
        if results:
            return pd.DataFrame(results)
    except Exception as e:
        print(f"FX error: {e}")
    return pd.DataFrame({"Moeda": ["USD", "EUR", "ZAR"], "Taxa (AOA)": ["N/A"] * 3})


def get_bna_rates() -> dict:
    try:
        body, _ = _run_in_thread(_scrape_bna)
        lines = body.split("\n")

        taxa_bna, inflacao = "N/A", "N/A"
        pct_pattern = re.compile(r"\d+[,\.]\d+%")

        for i, line in enumerate(lines):
            # taxa_bna: find standalone "Taxa BNA" label then look ahead for a % value
            if line.strip() == "Taxa BNA" and taxa_bna == "N/A":
                for j in range(i + 1, min(i + 6, len(lines))):
                    m = pct_pattern.search(lines[j])
                    if m:
                        taxa_bna = m.group()
                        break
            # inflacao: find "Taxa de Inflação" label
            if "Taxa de Inflação" in line and inflacao == "N/A":
                for j in range(i + 1, min(i + 6, len(lines))):
                    m = pct_pattern.search(lines[j])
                    if m:
                        inflacao = m.group()
                        break

        return {"taxa_bna": taxa_bna, "inflacao": inflacao}
    except Exception as e:
        return {"taxa_bna": "N/A", "inflacao": "N/A"}


def get_luibor_rates() -> pd.DataFrame:
    try:
        _, luibor_rows = _run_in_thread(_scrape_bna)
        if luibor_rows:
            return pd.DataFrame(luibor_rows).drop_duplicates()
    except Exception as e:
        print(f"LUIBOR error: {e}")
    return pd.DataFrame({
        "Maturidade": ["O/N", "1M", "3M", "6M", "9M", "12M"],
        "Taxa (%)": ["N/A"] * 6
    })
