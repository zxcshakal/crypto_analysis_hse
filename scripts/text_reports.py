## \file text_reports.py
#  \brief Текстовые/табличные отчёты: проекция с фильтром, частоты, описательные
#         статистики и сводная таблица. Логика отчётов не менялась.
#
#  \author Прохачев Никита
#  \date 2026

import pandas as pd
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_loader import load_data, build_full_table


## \brief Сохранить отчёт в файл выбранного формата (xlsx/csv/txt).
#  \param df Данные отчёта.
#  \param path Путь сохранения; расширение определяет формат.
#  \param sheet_name Имя листа для xlsx.
#  \return None
#  \throws ValueError если DataFrame пустой.
def save_report(df: pd.DataFrame, path: str, sheet_name: str = "Отчёт") -> None:
    if df is None or df.empty:
        raise ValueError("DataFrame пустой — нечего сохранять.")

    ext = os.path.splitext(path)[-1].lower()

    if ext == ".xlsx":
        df.to_excel(path, index=True, sheet_name=sheet_name)
    elif ext == ".csv":
        df.to_csv(path, index=True, sep=";", encoding="utf-8-sig")
    elif ext == ".txt":
        df.to_csv(path, index=True, sep="\t", encoding="utf-8")
    else:
        df.to_excel(path + ".xlsx", index=True, sheet_name=sheet_name)


## \brief Отчёт «Сделки»: проекция нужных колонок с фильтрами.
#  \param data Словарь справочников.
#  \param crypto_names Список тикеров для фильтра (None — все).
#  \param date_from Нижняя граница даты (строка) или None.
#  \param date_to Верхняя граница даты (строка) или None.
#  \param exchange_name Название биржи для фильтра или None.
#  \return Отфильтрованный и отсортированный DataFrame.
def report_simple(
        data: dict,
        crypto_names: list = None,
        date_from:    str  = None,
        date_to:      str  = None,
        exchange_name: str = None,
) -> pd.DataFrame:
    df   = build_full_table(data)
    mask = pd.Series(True, index=df.index)

    if crypto_names:
        mask &= df["ТИКЕР"].isin(crypto_names)
    if date_from:
        mask &= df["ДАТА"] >= pd.to_datetime(date_from)
    if date_to:
        mask &= df["ДАТА"] <= pd.to_datetime(date_to)
    if exchange_name:
        mask &= df["НАЗВ_БИРЖИ"] == exchange_name

    cols = [
        "ДАТА", "НАЗВ_КРИПТО", "ТИКЕР", "НАЗВ_КАТ",
        "НАЗВ_БИРЖИ", "СТРАНА",
        "ЦЕНА_ОТКРЫТИЯ", "ЦЕНА_ЗАКРЫТИЯ", "МАКС_ЦЕНА", "МИН_ЦЕНА",
        "ОБЪЕМ_ТОРГОВ", "КОЛ_СДЕЛОК",
    ]
    return (df.loc[mask, cols]
              .sort_values(["ДАТА", "ТИКЕР"])
              .reset_index(drop=True))


## \brief Частотный анализ качественного атрибута.
#  \param data Словарь справочников.
#  \param column Имя качественного столбца.
#  \return DataFrame: Значение / Частота / Процент (%).
#  \throws KeyError если столбец не найден.
def report_qualitative_stats(
        data: dict,
        column: str = "НАЗВ_КАТ",
) -> pd.DataFrame:
    df = build_full_table(data)
    if column not in df.columns:
        raise KeyError(f"Атрибут '{column}' не найден в таблице.")

    freq = df[column].value_counts().reset_index()
    freq.columns = ["Значение", "Частота"]
    freq["Процент (%)"] = (freq["Частота"] / freq["Частота"].sum() * 100).round(2)
    return freq.sort_values("Значение").reset_index(drop=True)


## \brief Описательные статистики количественных переменных.
#  \param data Словарь справочников.
#  \param columns Список количественных столбцов (None — набор по умолчанию).
#  \return DataFrame со статистиками (макс/мин/среднее/дисперсия/СКО).
def report_quantitative_stats(
        data: dict,
        columns: list = None,
) -> pd.DataFrame:
    df = build_full_table(data)
    if columns is None:
        columns = [
            "ЦЕНА_ОТКРЫТИЯ", "ЦЕНА_ЗАКРЫТИЯ",
            "МАКС_ЦЕНА", "МИН_ЦЕНА", "ОБЪЕМ_ТОРГОВ", "КОЛ_СДЕЛОК",
        ]

    return pd.DataFrame({
        "Переменная":             columns,
        "Максимум":               [df[c].max()  for c in columns],
        "Минимум":                [df[c].min()  for c in columns],
        "Среднее":                [df[c].mean() for c in columns],
        "Дисперсия (выб.)":       [df[c].var()  for c in columns],
        "Стандартное отклонение": [df[c].std()  for c in columns],
    }).round(4).set_index("Переменная")


## \brief Сводная таблица (pivot_table) с итогами по строкам и столбцам.
#  \param data Словарь справочников.
#  \param index Качественный атрибут строк.
#  \param columns Качественный атрибут столбцов.
#  \param values Количественный атрибут значений.
#  \param aggfunc Функция агрегации (mean/sum/count/...).
#  \return Сводная таблица с полем «ИТОГО».
def report_pivot(
        data: dict,
        index:   str = "НАЗВ_КРИПТО",
        columns: str = "НАЗВ_БИРЖИ",
        values:  str = "ОБЪЕМ_ТОРГОВ",
        aggfunc: str = "mean",
) -> pd.DataFrame:
    df = build_full_table(data)
    return pd.pivot_table(
        df,
        index=index,
        columns=columns,
        values=values,
        aggfunc=aggfunc,
        fill_value=0,
        margins=True,
        margins_name="ИТОГО",
    ).round(2)


if __name__ == "__main__":
    db_path = sys.argv[1] if len(sys.argv) > 1 else None
    try:
        data = load_data(db_path)
    except FileNotFoundError as e:
        print(e); sys.exit(1)

    r1 = report_simple(data, crypto_names=["BTC", "ETH"],
                       date_from="2026-01-01", date_to="2026-01-07",
                       exchange_name="Binance")
    print(r1.head().to_string())
    r2 = report_qualitative_stats(data, column="НАЗВ_КАТ")
    print(r2.to_string(index=False))
    r3 = report_quantitative_stats(data)
    print(r3)
    r4 = report_pivot(data)
    print(r4)
