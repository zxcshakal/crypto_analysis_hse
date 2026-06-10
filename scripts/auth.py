## \file auth.py
#  \brief Локальная система аккаунтов: регистрация, вход, профиль, избранное, KYC.
#
#  Модуль хранит пользователей в Excel-файле (users_db.xlsx) и кеширует
#  активную сессию в pickle-файле (session.pkl). Пароли НЕ хранятся в открытом
#  виде — сохраняется только соль и SHA-256-хеш (соль + пароль).
#
#  Используются только стандартная библиотека + pandas/openpyxl, поэтому
#  ограничение курсовой «только разрешённые библиотеки» не нарушается.
#
#  \author Прохачев Никита
#  \date 2026

import os
import sys
import pickle
import hashlib
import secrets
import datetime as dt

import pandas as pd

## Каталог данных: папка database/ в корне проекта (рядом с .exe в сборке).
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "database")
os.makedirs(DATA_DIR, exist_ok=True)

## Путь к Excel-базе пользователей (логины/пароли/профиль).
USERS_DB_PATH = os.path.join(DATA_DIR, "users_db.xlsx")
## Путь к pickle-файлу активной сессии («запомнить вход»).
SESSION_PATH = os.path.join(DATA_DIR, "session.pkl")
## Имя листа Excel со всеми пользователями.
USERS_SHEET = "ПОЛЬЗОВАТЕЛИ"

## Полный набор колонок таблицы пользователей.
#  КYC по умолчанию «Нет» и проставляется администратором вручную в Excel.
USER_COLUMNS = [
    "ЛОГИН", "ХЕШ_ПАРОЛЯ", "СОЛЬ", "ДАТА_РЕГ",
    "ФИО", "ДАТА_РОЖДЕНИЯ", "ТЕЛЕФОН", "СТРАНА",
    "KYC", "АВАТАР", "ИЗБР_КРИПТО", "ИЗБР_БИРЖИ",
]


## \brief Вычислить SHA-256-хеш пароля с заданной солью.
#  \param password Пароль в открытом виде.
#  \param salt Шестнадцатеричная соль (строка).
#  \return Шестнадцатеричный хеш (строка).
def _hash_password(password: str, salt: str) -> str:
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


## \brief Создать пустую базу пользователей, если файла ещё нет.
#  \return None
def _ensure_db() -> None:
    if not os.path.exists(USERS_DB_PATH):
        df = pd.DataFrame(columns=USER_COLUMNS)
        df.to_excel(USERS_DB_PATH, index=False, sheet_name=USERS_SHEET)


## \brief Загрузить таблицу пользователей из Excel.
#  \return DataFrame с колонками USER_COLUMNS (возможно пустой).
def load_users() -> pd.DataFrame:
    _ensure_db()
    df = pd.read_excel(USERS_DB_PATH, sheet_name=USERS_SHEET, dtype=str)
    for col in USER_COLUMNS:                      # гарантируем наличие всех колонок
        if col not in df.columns:
            df[col] = ""
    return df[USER_COLUMNS].fillna("")


## \brief Сохранить таблицу пользователей обратно в Excel.
#  \param df DataFrame с колонками USER_COLUMNS.
#  \return None
def save_users(df: pd.DataFrame) -> None:
    df.to_excel(USERS_DB_PATH, index=False, sheet_name=USERS_SHEET)


## \brief Проверить, есть ли хотя бы один зарегистрированный пользователь.
#  \return True, если база пуста (нужна первичная регистрация).
def is_first_launch() -> bool:
    return load_users().empty


## \brief Зарегистрировать нового пользователя.
#  \param login Уникальный логин (никнейм).
#  \param password Пароль в открытом виде (будет захеширован).
#  \param fio ФИО (необязательно).
#  \return Кортеж (успех, сообщение).
def register(login: str, password: str, fio: str = "") -> tuple:
    login = (login or "").strip()
    if len(login) < 3:
        return False, "Логин должен содержать минимум 3 символа."
    if len(password) < 4:
        return False, "Пароль должен содержать минимум 4 символа."

    df = load_users()
    if (df["ЛОГИН"].str.lower() == login.lower()).any():
        return False, "Такой логин уже занят."

    salt = secrets.token_hex(16)
    new_row = {
        "ЛОГИН": login,
        "ХЕШ_ПАРОЛЯ": _hash_password(password, salt),
        "СОЛЬ": salt,
        "ДАТА_РЕГ": dt.date.today().isoformat(),
        "ФИО": fio, "ДАТА_РОЖДЕНИЯ": "", "ТЕЛЕФОН": "", "СТРАНА": "",
        "KYC": "Нет",          # проставляется администратором вручную
        "АВАТАР": "", "ИЗБР_КРИПТО": "", "ИЗБР_БИРЖИ": "",
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save_users(df)
    return True, "Регистрация успешна."


## \brief Проверить пару логин/пароль.
#  \param login Логин.
#  \param password Пароль в открытом виде.
#  \return Кортеж (успех, сообщение). При успехе сессия НЕ сохраняется здесь —
#          это делает caller через save_session().
def login(login: str, password: str) -> tuple:
    df = load_users()
    row = df[df["ЛОГИН"].str.lower() == (login or "").strip().lower()]
    if row.empty:
        return False, "Пользователь не найден."
    row = row.iloc[0]
    if _hash_password(password, row["СОЛЬ"]) != row["ХЕШ_ПАРОЛЯ"]:
        return False, "Неверный пароль."
    return True, "Вход выполнен."


## \brief Получить словарь профиля пользователя по логину.
#  \param login Логин.
#  \return dict с полями профиля или None, если не найден.
def get_profile(login: str) -> dict:
    df = load_users()
    row = df[df["ЛОГИН"] == login]
    if row.empty:
        return None
    return row.iloc[0].to_dict()


## \brief Обновить произвольные поля профиля пользователя.
#  \param login Логин.
#  \param fields Словарь {колонка: значение}; KYC и пароль игнорируются.
#  \return True при успехе.
def update_profile(login: str, fields: dict) -> bool:
    df = load_users()
    idx = df.index[df["ЛОГИН"] == login]
    if len(idx) == 0:
        return False
    protected = {"ЛОГИН", "ХЕШ_ПАРОЛЯ", "СОЛЬ", "KYC"}  # KYC меняет только админ
    for key, val in fields.items():
        if key in df.columns and key not in protected:
            df.loc[idx, key] = str(val)
    save_users(df)
    return True


## \brief Добавить/убрать элемент в избранном (крипто или биржа).
#  \param login Логин.
#  \param column Колонка "ИЗБР_КРИПТО" или "ИЗБР_БИРЖИ".
#  \param item Название элемента.
#  \return Новый список избранного.
def toggle_favorite(login: str, column: str, item: str) -> list:
    df = load_users()
    idx = df.index[df["ЛОГИН"] == login]
    if len(idx) == 0:
        return []
    raw = df.loc[idx[0], column] or ""
    items = [x for x in raw.split(";") if x]
    if item in items:
        items.remove(item)
    else:
        items.append(item)
    df.loc[idx, column] = ";".join(items)
    save_users(df)
    return items


## \brief Прочитать список избранного как Python-список.
#  \param profile Словарь профиля (из get_profile).
#  \param column Колонка избранного.
#  \return Список строк.
def get_favorites(profile: dict, column: str) -> list:
    raw = (profile or {}).get(column, "") or ""
    return [x for x in raw.split(";") if x]


## \brief Сохранить активную сессию в pickle.
#  \param login Логин вошедшего пользователя.
#  \return None
def save_session(login: str) -> None:
    with open(SESSION_PATH, "wb") as f:
        pickle.dump({"login": login, "ts": dt.datetime.now().isoformat()}, f)


## \brief Загрузить сохранённую сессию (если есть).
#  \return Логин из сессии или None.
def load_session() -> str:
    if not os.path.exists(SESSION_PATH):
        return None
    try:
        with open(SESSION_PATH, "rb") as f:
            return pickle.load(f).get("login")
    except Exception:
        return None


## \brief Удалить сохранённую сессию (выход из аккаунта).
#  \return None
def clear_session() -> None:
    if os.path.exists(SESSION_PATH):
        os.remove(SESSION_PATH)
