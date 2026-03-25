from src.scrapers.yahoo_scraper import get_global_markets, get_commodities, get_crypto
from src.scrapers.bna_scraper import get_luibor_rates, get_exchange_rates


def scrape_all_external_data() -> dict:
    """
    Scrape all external data sources.
    Returns a dict of DataFrames, one per section.
    Falls back gracefully if any source fails.
    """
    print("Scraping global markets...")
    markets = get_global_markets()

    print("Scraping commodities...")
    commodities = get_commodities()

    print("Scraping crypto...")
    crypto = get_crypto()

    print("Scraping BNA LUIBOR rates...")
    luibor = get_luibor_rates()

    print("Scraping BNA exchange rates...")
    fx_rates = get_exchange_rates()

    return {
        "markets": markets,
        "commodities": commodities,
        "crypto": crypto,
        "luibor": luibor,
        "fx_rates": fx_rates,
    }
