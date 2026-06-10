## \file data_loader.py
#  \brief Загрузка справочников из Excel и построение денормализованной таблицы.
#
#  Excel выступает локальной базой и офлайн-фолбэком: даже без доступа к API
#  приложение работает на этих данных. Листы API_КУРСЫ / API_ИСТОРИЯ,
#  добавляемые api_client.py, здесь намеренно не загружаются в основную
#  модель, чтобы не ломать существующие отчёты и графики.
#
#  \author Фахрутдинов Амир
#  \date 2026

import pandas as pd
import os
import sys

## Корень проекта: на уровень выше папки scripts/ (или рядом с .exe).
if getattr(sys, "frozen", False):
    _ROOT = os.path.dirname(sys.executable)
else:
    _ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
## Путь к Excel-базе по умолчанию (папка database/ в корне проекта).
DEFAULT_DB_PATH = os.path.join(_ROOT, "database", "crypto_database.xlsx")


## \brief Загрузить все справочники из Excel в словарь DataFrame-ов.
#  \param db_path Путь к Excel-файлу; None — использовать DEFAULT_DB_PATH.
#  \return Словарь с ключами categories/cryptos/exchanges/pairs/trades/apis.
#  \throws FileNotFoundError если файл базы не найден.
def load_data(db_path: str = None) -> dict:
    if db_path is None:
        db_path = DEFAULT_DB_PATH

    if not os.path.exists(db_path):
        raise FileNotFoundError(f"Файл базы данных не найден: {db_path}")

    xl = pd.ExcelFile(db_path)

    data = {
        "categories": xl.parse("КАТЕГОРИИ"),
        "cryptos":    xl.parse("КРИПТОВАЛЮТЫ"),
        "exchanges":  xl.parse("БИРЖИ"),
        "pairs":      xl.parse("ТОРГОВЫЕ_ПАРЫ"),
        "trades":     xl.parse("ТОРГОВЫЕ_ДАННЫЕ"),
        "apis":       xl.parse("API"),
    }

    data["trades"]["ДАТА"] = pd.to_datetime(data["trades"]["ДАТА"])

    return data


## \brief Собрать единую денормализованную таблицу из всех справочников.
#  \param data Словарь DataFrame-ов из load_data().
#  \return Денормализованный DataFrame (торговые данные + все атрибуты).
def build_full_table(data: dict) -> pd.DataFrame:
    df = data["trades"].merge(data["cryptos"],   on="ID_КРИПТО", how="left")
    df = df.merge(data["categories"], on="ID_КАТ",   how="left")
    df = df.merge(data["exchanges"],  on="ID_БИРЖИ",  how="left", suffixes=("", "_БИР"))
    df = df.merge(data["pairs"],      on="ID_ПАРЫ",   how="left")
    return df


if __name__ == "__main__":
    import sys

    path = sys.argv[1] if len(sys.argv) > 1 else None
    try:
        data = load_data(path)
    except FileNotFoundError as e:
        print(e)
        sys.exit(1)

    print("Загружены таблицы:")
    for name, df in data.items():
        print(f"  {name:15s} — {len(df)} строк, {len(df.columns)} столбцов")

    full = build_full_table(data)
    print(f"\nДенормализованная таблица: {len(full)} строк, {len(full.columns)} столбцов")
    print("Столбцы:", list(full.columns))
