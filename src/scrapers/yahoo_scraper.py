import yfinance as yf
import pandas as pd
from src.config import INDICES_TICKERS, COMMODITIES_TICKERS, CRYPTO_TICKERS


def _fetch_ticker_data(tickers: dict, label_col: str) -> pd.DataFrame:
    results = []
    for name, ticker in tickers.items():
        try:
            hist = yf.Ticker(ticker).history(period="2d")
            if len(hist) >= 2:
                prev = round(hist["Close"].iloc[-2], 2)
                curr = round(hist["Close"].iloc[-1], 2)
                change = round(((curr - prev) / prev) * 100, 2)
            elif len(hist) == 1:
                curr = round(hist["Close"].iloc[-1], 2)
                prev = curr
                change = 0.0
            else:
                prev, curr, change = "N/A", "N/A", "N/A"
            results.append({label_col: name, "Anterior": prev, "Atual": curr, "Var (%)": change})
        except Exception:
            results.append({label_col: name, "Anterior": "N/A", "Atual": "N/A", "Var (%)": "N/A"})
    return pd.DataFrame(results)


def get_global_markets() -> pd.DataFrame:
    return _fetch_ticker_data(INDICES_TICKERS, "Índice")


def get_commodities() -> pd.DataFrame:
    return _fetch_ticker_data(COMMODITIES_TICKERS, "Commodity")


def get_crypto() -> pd.DataFrame:
    return _fetch_ticker_data(CRYPTO_TICKERS, "Cripto")
