## \file api_client.py
#  \brief Клиент CoinGecko (stdlib urllib) с кешированием в Excel и офлайн-фолбэком.
#
#  Стратегия данных:
#   * API освежает СПРАВОЧНИК (id / тикер / имя / текущая цена / капитализация)
#     и историю цен по монете. Именно тут хранится сопоставление
#     «тикер → id монеты» для дальнейших запросов к API.
#   * Если сети нет или API недоступен — берутся последние закешированные
#     данные из соответствующего листа Excel.
#
#  Бесплатный CoinGecko не требует ключа. Для CoinMarketCap нужен ключ —
#  он подставляется в CMC_API_KEY (по умолчанию режим CMC выключен).
#
#  \author Фахрутдинов Амир
#  \date 2026

import os
import json
import time
import datetime as dt
import urllib.request
import urllib.error

import pandas as pd

## Базовый URL бесплатного CoinGecko API (без ключа).
COINGECKO_BASE = "https://api.coingecko.com/api/v3"
## Ключ CoinMarketCap (вставьте свой, если хотите использовать CMC).
CMC_API_KEY = ""

## Сопоставление тикеров из локальной БД с идентификаторами монет CoinGecko.
#  Это и есть «имя криптовалюты и её id для обращения к API».
TICKER_TO_ID = {
    "BTC": "bitcoin", "ETH": "ethereum", "USDT": "tether",
    "BNB": "binancecoin", "SOL": "solana", "XRP": "ripple",
    "ADA": "cardano", "DOGE": "dogecoin",
}

## Лист Excel с актуальными курсами из API.
SHEET_RATES = "API_КУРСЫ"
## Лист Excel с историей цен из API.
SHEET_HISTORY = "API_ИСТОРИЯ"


## \brief Выполнить GET-запрос и вернуть разобранный JSON.
#  \param url Полный URL запроса.
#  \param timeout Таймаут в секундах.
#  \return Разобранный объект (dict/list).
#  \throws urllib.error.URLError при сетевой ошибке.
def _http_get_json(url: str, timeout: int = 15):
    req = urllib.request.Request(url, headers={"User-Agent": "BIV253-CryptoApp"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


## \brief Записать DataFrame в отдельный лист Excel, не трогая остальные листы.
#  \param df Данные для записи.
#  \param db_path Путь к Excel-файлу.
#  \param sheet Имя листа.
#  \return None
def _write_sheet(df: pd.DataFrame, db_path: str, sheet: str) -> None:
    mode = "a" if os.path.exists(db_path) else "w"
    kwargs = {"engine": "openpyxl", "mode": mode}
    if mode == "a":
        kwargs["if_sheet_exists"] = "replace"
    with pd.ExcelWriter(db_path, **kwargs) as xw:
        df.to_excel(xw, index=False, sheet_name=sheet)


## \brief Прочитать ранее закешированный лист (офлайн-фолбэк).
#  \param db_path Путь к Excel-файлу.
#  \param sheet Имя листа.
#  \return DataFrame или None, если листа нет.
def _read_cached_sheet(db_path: str, sheet: str):
    try:
        return pd.read_excel(db_path, sheet_name=sheet)
    except Exception:
        return None


## \brief Освежить справочник курсов из API и записать его в Excel.
#  \param db_path Путь к Excel-базе (туда добавится лист SHEET_RATES).
#  \param tickers Список тикеров; по умолчанию все из TICKER_TO_ID.
#  \return Кортеж (DataFrame, источник), источник = "api" | "cache" | "none".
def refresh_reference(db_path: str, tickers: list = None) -> tuple:
    tickers = tickers or list(TICKER_TO_ID.keys())
    ids = ",".join(TICKER_TO_ID[t] for t in tickers if t in TICKER_TO_ID)
    url = (f"{COINGECKO_BASE}/coins/markets?vs_currency=usd&ids={ids}"
           f"&order=market_cap_desc&per_page=250&page=1")
    try:
        raw = _http_get_json(url)
        rows = [{
            "ID_МОНЕТЫ": c.get("id"),
            "ТИКЕР": (c.get("symbol") or "").upper(),
            "НАЗВ_КРИПТО": c.get("name"),
            "ЦЕНА_USD": c.get("current_price"),
            "РЫНОЧ_КАП_USD": c.get("market_cap"),
            "ИЗМ_24Ч_%": c.get("price_change_percentage_24h"),
            "ОБНОВЛЕНО": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        } for c in raw]
        df = pd.DataFrame(rows)
        _write_sheet(df, db_path, SHEET_RATES)
        return df, "api"
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError):
        cached = _read_cached_sheet(db_path, SHEET_RATES)
        return (cached, "cache") if cached is not None else (None, "none")


## \brief Освежить историю цен по одной монете из API и записать в Excel.
#  \param db_path Путь к Excel-базе (лист SHEET_HISTORY).
#  \param ticker Тикер монеты (например, "BTC").
#  \param days Глубина истории в днях.
#  \return Кортеж (DataFrame, источник): "api" | "cache" | "none".
def refresh_history(db_path: str, ticker: str = "BTC", days: int = 30) -> tuple:
    coin_id = TICKER_TO_ID.get(ticker.upper())
    if coin_id is None:
        return None, "none"
    url = (f"{COINGECKO_BASE}/coins/{coin_id}/market_chart"
           f"?vs_currency=usd&days={days}&interval=daily")
    try:
        raw = _http_get_json(url)
        prices = raw.get("prices", [])
        volumes = dict((p[0], v[1]) for p, v in zip(prices, raw.get("total_volumes", [])))
        rows = []
        for ts, price in prices:
            rows.append({
                "ТИКЕР": ticker.upper(),
                "ДАТА": dt.datetime.utcfromtimestamp(ts / 1000).date().isoformat(),
                "ЦЕНА_ЗАКРЫТИЯ": round(price, 4),
                "ОБЪЕМ_ТОРГОВ": round(volumes.get(ts, 0.0), 2),
            })
        df = pd.DataFrame(rows)
        _write_sheet(df, db_path, SHEET_HISTORY)
        return df, "api"
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError):
        cached = _read_cached_sheet(db_path, SHEET_HISTORY)
        return (cached, "cache") if cached is not None else (None, "none")


## \brief Получить текущую цену одной монеты (быстрый запрос с фолбэком).
#  \param db_path Путь к Excel-базе (для кеш-фолбэка).
#  \param ticker Тикер монеты.
#  \return Цена в USD (float) или None.
def get_price(db_path: str, ticker: str) -> float:
    df, src = refresh_reference(db_path, [ticker.upper()])
    if df is None or df.empty:
        return None
    row = df[df["ТИКЕР"] == ticker.upper()]
    return float(row.iloc[0]["ЦЕНА_USD"]) if not row.empty else None


if __name__ == "__main__":
    import sys
    db = sys.argv[1] if len(sys.argv) > 1 else "crypto_database.xlsx"
    ref, src = refresh_reference(db)
    print(f"Справочник из источника: {src}")
    if ref is not None:
        print(ref.to_string(index=False))
