## \file news_store.py
#  \brief Хранилище новостей: локальный Excel + клиент к домашнему серверу.
#
#  Новости хранятся локально в data/news_db.xlsx. Если задан адрес сервера
#  (SERVER_URL), приложение может: забрать общие новости (GET /news) и
#  отправить свою (POST /news). Если сервер недоступен — работает офлайн на
#  локальной копии. Только стандартная библиотека + pandas/openpyxl.
#
#  \author Хотнянский Кирилл
#  \date 2026

import os
import sys
import json
import datetime as dt
import urllib.request
import urllib.error

import pandas as pd

## Каталог данных: папка database/ в корне проекта (рядом с .exe в сборке).
if getattr(sys, "frozen", False):
    _BASE = os.path.dirname(sys.executable)
else:
    _BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(_BASE, "database")
os.makedirs(DATA_DIR, exist_ok=True)
## Локальная база новостей.
NEWS_DB_PATH = os.path.join(DATA_DIR, "news_db.xlsx")
## Имя листа новостей.
NEWS_SHEET = "НОВОСТИ"
## Адрес домашнего сервера новостей. Пусто = офлайн-режим.
#  Пример: "http://192.168.0.50:8765"
SERVER_URL = ""

## Колонки таблицы новостей.
NEWS_COLUMNS = ["ID", "АВТОР", "ЗАГОЛОВОК", "ТЕКСТ", "ДАТА"]


## \brief Создать пустую базу новостей, если её ещё нет.
#  \return None
def _ensure_db() -> None:
    if not os.path.exists(NEWS_DB_PATH):
        pd.DataFrame(columns=NEWS_COLUMNS).to_excel(
            NEWS_DB_PATH, index=False, sheet_name=NEWS_SHEET)


## \brief Загрузить локальные новости из Excel.
#  \return DataFrame, отсортированный по дате (свежие сверху).
def load_local_news() -> pd.DataFrame:
    _ensure_db()
    df = pd.read_excel(NEWS_DB_PATH, sheet_name=NEWS_SHEET, dtype=str).fillna("")
    if not df.empty:
        df = df.sort_values("ДАТА", ascending=False).reset_index(drop=True)
    return df


## \brief Сохранить локальную базу новостей.
#  \param df DataFrame новостей.
#  \return None
def _save_local(df: pd.DataFrame) -> None:
    df.to_excel(NEWS_DB_PATH, index=False, sheet_name=NEWS_SHEET)


## \brief Добавить новость локально и (если задан сервер) отправить на сервер.
#  \param author Автор (никнейм).
#  \param title Заголовок.
#  \param text Текст новости.
#  \return Кортеж (успех, сообщение).
def add_news(author: str, title: str, text: str) -> tuple:
    if not title.strip():
        return False, "Заголовок не может быть пустым."

    df = load_local_news()
    new_id = (df["ID"].astype(int).max() + 1) if not df.empty else 1
    record = {
        "ID": str(new_id), "АВТОР": author,
        "ЗАГОЛОВОК": title.strip(), "ТЕКСТ": text.strip(),
        "ДАТА": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
    _save_local(df)

    if SERVER_URL:
        ok, msg = _post_remote(record)
        return ok, ("Новость опубликована (сервер)." if ok
                    else f"Сохранено локально, сервер недоступен: {msg}")
    return True, "Новость сохранена локально."


## \brief Отправить новость на сервер (POST /news).
#  \param record Словарь новости.
#  \return Кортеж (успех, сообщение).
def _post_remote(record: dict) -> tuple:
    try:
        data = json.dumps(record).encode("utf-8")
        req = urllib.request.Request(
            f"{SERVER_URL}/news", data=data, method="POST",
            headers={"Content-Type": "application/json",
                     "ngrok-skip-browser-warning": "true"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            return resp.status == 200, "ok"
    except (urllib.error.URLError, urllib.error.HTTPError, OSError) as e:
        return False, str(e)


## \brief Забрать общие новости с сервера и слить их в локальную базу.
#  \return Кортеж (DataFrame, источник): "server" | "local".
def sync_from_server() -> tuple:
    if not SERVER_URL:
        return load_local_news(), "local"
    try:
        req = urllib.request.Request(
            f"{SERVER_URL}/news", method="GET",
            headers={"ngrok-skip-browser-warning": "true"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            remote = json.loads(resp.read().decode("utf-8"))
        rdf = pd.DataFrame(remote)
        local = load_local_news()
        merged = (pd.concat([local, rdf], ignore_index=True)
                    .drop_duplicates(subset=["ID", "АВТОР", "ДАТА"], keep="last"))
        _save_local(merged)
        return load_local_news(), "server"
    except (urllib.error.URLError, urllib.error.HTTPError, OSError, ValueError):
        return load_local_news(), "local"
