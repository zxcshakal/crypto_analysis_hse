## \file graphic_reports.py
#  \brief Пять типов графиков (matplotlib) в тёмной палитре Catppuccin.
#         Логика построения не менялась; уточнены только подписи/заголовки.
#
#  \author Фахрутдиноа Амир
#  \date 2026

import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from data_loader import load_data, build_full_table

## Палитра серий графиков.
COLORS = [
    "#89b4fa", "#a6e3a1", "#f38ba8", "#fab387",
    "#cba6f7", "#94e2d5", "#f9e2af", "#eba0ac",
]
BG_FIGURE  = "#1e1e2e"   ##< Фон фигуры.
BG_AXES    = "#181825"   ##< Фон области осей.
FG_TEXT    = "#cdd6f4"   ##< Цвет текста.
FG_SPINE   = "#313244"   ##< Цвет рамок осей.

plt.rcParams.update({
    "font.family":      "DejaVu Sans",
    "axes.titlesize":   12,
    "axes.labelsize":   10,
    "xtick.labelsize":  8,
    "ytick.labelsize":  8,
    "figure.dpi":       110,
})


## \brief Линейный график динамики показателя во времени.
#  \param data Словарь справочников.
#  \param date_col Колонка с датой (по умолчанию "ДАТА").
#  \param value_col Количественный атрибут для отображения.
#  \param group_col Качественный атрибут для группировки (разные линии).
#  \param aggfunc Функция агрегации для группировки по датам.
#  \return Фигура matplotlib.
def chart_line(
        data: dict,
        date_col: str = "ДАТА",
        value_col: str = "ЦЕНА_ЗАКРЫТИЯ",
        group_col: str = "ТИКЕР",
        aggfunc: str = "mean",
) -> plt.Figure:
    df = build_full_table(data)

    # Группируем по дате и группе, агрегируем значения
    grouped = df.groupby([date_col, group_col])[value_col].agg(aggfunc).reset_index()

    # Сортируем по дате
    grouped[date_col] = pd.to_datetime(grouped[date_col])
    grouped = grouped.sort_values(date_col)

    # Получаем уникальные группы (линии)
    groups = grouped[group_col].unique()

    fig, ax = plt.subplots(figsize=(12, 6))
    _style_ax(ax, fig)

    for i, grp in enumerate(groups):
        subset = grouped[grouped[group_col] == grp]
        ax.plot(subset[date_col], subset[value_col],
                marker='o', linewidth=2, markersize=4,
                color=COLORS[i % len(COLORS)],
                label=str(grp), alpha=0.85)

    ax.set_xlabel("Дата")
    ax.set_ylabel(f"{aggfunc.capitalize()}({value_col})")
    ax.set_title(f"Динамика «{value_col}» по датам (в разрезе «{group_col}»)")
    ax.legend(title=group_col, bbox_to_anchor=(1.01, 1), loc="upper left",
              facecolor=BG_AXES, labelcolor=FG_TEXT,
              title_fontsize=8, fontsize=8)

    # Форматирование оси Y
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:,.0f}"))

    # Поворот подписей дат для лучшей читаемости
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')

    # Добавляем сетку для удобства чтения
    ax.grid(True, alpha=0.3, color=FG_SPINE, linestyle='--')

    fig.tight_layout()
    return fig

## \brief Применить единое тёмное оформление к осям.
#  \param ax Объект осей matplotlib.
#  \param fig Фигура matplotlib.
#  \return None
def _style_ax(ax, fig):
    fig.patch.set_facecolor(BG_FIGURE)
    ax.set_facecolor(BG_AXES)
    ax.tick_params(colors=FG_TEXT)
    ax.xaxis.label.set_color(FG_TEXT)
    ax.yaxis.label.set_color(FG_TEXT)
    ax.title.set_color(FG_TEXT)
    for spine in ax.spines.values():
        spine.set_color(FG_SPINE)


## \brief Сохранить фигуру в файл; формат — по расширению пути.
#  \param fig Фигура matplotlib.
#  \param path Путь сохранения (png/jpg/pdf/svg).
#  \return None
#  \throws ValueError если фигура не задана.
def save_figure(fig: plt.Figure, path: str) -> None:
    if fig is None:
        raise ValueError("Фигура не задана (None).")
    ext = os.path.splitext(path)[-1].lower()
    fmt_map = {".png": "png", ".jpg": "jpeg", ".jpeg": "jpeg",
               ".pdf": "pdf", ".svg": "svg"}
    fmt = fmt_map.get(ext, "png")
    fig.savefig(path, format=fmt, bbox_inches="tight",
                facecolor=fig.get_facecolor())


## \brief Кластеризованная столбчатая диаграмма (значение по X с группировкой).
#  \param data Словарь справочников.
#  \param qual_x Качественный атрибут оси X.
#  \param qual_group Качественный атрибут группировки (легенда).
#  \param value_col Количественный атрибут значения.
#  \param aggfunc Функция агрегации.
#  \return Фигура matplotlib.
def chart_clustered_bar(
        data: dict,
        qual_x:    str = "НАЗВ_КРИПТО",
        qual_group: str = "НАЗВ_БИРЖИ",
        value_col: str = "ОБЪЕМ_ТОРГОВ",
        aggfunc:   str = "mean",
) -> plt.Figure:
    df    = build_full_table(data)
    pivot = df.groupby([qual_x, qual_group])[value_col].agg(aggfunc).unstack(qual_group).fillna(0)

    cats   = pivot.index.tolist()
    groups = pivot.columns.tolist()
    x      = np.arange(len(cats))
    width  = 0.8 / max(len(groups), 1)

    fig, ax = plt.subplots(figsize=(11, 5))
    _style_ax(ax, fig)

    for i, grp in enumerate(groups):
        offset = (i - len(groups) / 2 + 0.5) * width
        ax.bar(x + offset, pivot[grp], width * 0.9,
               label=str(grp), color=COLORS[i % len(COLORS)], alpha=0.88)

    ax.set_xticks(x)
    ax.set_xticklabels(cats, rotation=25, ha="right", color=FG_TEXT)
    ax.set_xlabel(qual_x)
    ax.set_ylabel(f"{aggfunc.capitalize()}({value_col})")
    ax.set_title(f"Сравнение «{value_col}»: {qual_x} в разрезе {qual_group}")
    ax.legend(title=qual_group, bbox_to_anchor=(1.01, 1), loc="upper left",
              facecolor=BG_AXES, labelcolor=FG_TEXT, title_fontsize=8, fontsize=8)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    fig.tight_layout()
    return fig


## \brief Категоризированная гистограмма: распределение по уровням атрибута.
#  \param data Словарь справочников.
#  \param quant_col Количественный атрибут.
#  \param qual_col Качественный атрибут (по нему делятся подграфики).
#  \param bins Число интервалов.
#  \return Фигура matplotlib.
def chart_categorized_histogram(
        data: dict,
        quant_col: str = "ЦЕНА_ЗАКРЫТИЯ",
        qual_col:  str = "ТИКЕР",
        bins:      int = 25,
) -> plt.Figure:
    df     = build_full_table(data)
    levels = sorted(df[qual_col].dropna().unique())

    ncols = 2
    nrows = max((len(levels) + 1) // 2, 1)
    fig, axes = plt.subplots(nrows=nrows, ncols=ncols,
                             figsize=(12, 3.4 * nrows))
    fig.patch.set_facecolor(BG_FIGURE)
    axes = np.array(axes).flatten()

    for i, lvl in enumerate(levels):
        subset = df.loc[df[qual_col] == lvl, quant_col].dropna()
        ax = axes[i]
        ax.set_facecolor(BG_AXES)
        ax.hist(subset, bins=bins,
                color=COLORS[i % len(COLORS)], alpha=0.84,
                edgecolor=BG_FIGURE, linewidth=0.5)
        ax.set_title(f"{qual_col} = {lvl}", color=FG_TEXT, fontsize=9)
        ax.set_xlabel(quant_col, color=FG_TEXT, fontsize=8)
        ax.set_ylabel("Частота", color=FG_TEXT, fontsize=8)
        ax.tick_params(colors=FG_TEXT, labelsize=7)
        for sp in ax.spines.values():
            sp.set_color(FG_SPINE)
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:,.0f}"))

    for j in range(len(levels), len(axes)):
        axes[j].set_visible(False)

    fig.suptitle(f"Распределение «{quant_col}» по уровням «{qual_col}»",
                 color=FG_TEXT, fontsize=13, y=1.01)
    fig.tight_layout()
    return fig


## \brief Диаграмма Бокса-Вискера: разброс величины по группам.
#  \param data Словарь справочников.
#  \param quant_col Количественный атрибут (ось Y).
#  \param qual_col Качественный атрибут (ось X / группы).
#  \return Фигура matplotlib.
def chart_boxplot(
        data: dict,
        quant_col: str = "ОБЪЕМ_ТОРГОВ",
        qual_col:  str = "НАЗВ_КРИПТО",
) -> plt.Figure:
    df     = build_full_table(data)
    levels = sorted(df[qual_col].dropna().unique())
    groups_data = [df.loc[df[qual_col] == lvl, quant_col].dropna().values
                   for lvl in levels]

    fig, ax = plt.subplots(figsize=(11, 5))
    _style_ax(ax, fig)

    bp = ax.boxplot(
        groups_data,
        patch_artist=True,
        notch=False,
        vert=True,
        showfliers=True,
        flierprops=dict(marker="o", markersize=3, alpha=0.4,
                        markerfacecolor=COLORS[2]),
    )
    for patch, color in zip(bp["boxes"], COLORS):
        patch.set_facecolor(color)
        patch.set_alpha(0.76)
    for element in ("whiskers", "caps", "medians"):
        for line in bp[element]:
            line.set_color(FG_TEXT)

    ax.set_xticks(range(1, len(levels) + 1))
    ax.set_xticklabels(levels, rotation=25, ha="right", color=FG_TEXT)
    ax.set_xlabel(qual_col)
    ax.set_ylabel(quant_col)
    ax.set_title(f"Разброс «{quant_col}» по «{qual_col}» (Бокса-Вискера)")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    fig.tight_layout()
    return fig


## \brief Диаграмма рассеивания: связь двух количественных переменных.
#  \param data Словарь справочников.
#  \param quant_x Количественный атрибут оси X.
#  \param quant_y Количественный атрибут оси Y.
#  \param qual_col Качественный атрибут (цвет точек).
#  \return Фигура matplotlib.
def chart_scatter(
        data: dict,
        quant_x:  str = "ОБЪЕМ_ТОРГОВ",
        quant_y:  str = "ЦЕНА_ЗАКРЫТИЯ",
        qual_col: str = "ТИКЕР",
) -> plt.Figure:
    df     = build_full_table(data)
    levels = sorted(df[qual_col].dropna().unique())

    fig, ax = plt.subplots(figsize=(11, 6))
    _style_ax(ax, fig)

    handles = []
    for i, lvl in enumerate(levels):
        sub = df[df[qual_col] == lvl]
        ax.scatter(sub[quant_x], sub[quant_y],
                   color=COLORS[i % len(COLORS)],
                   alpha=0.38, s=20, label=str(lvl))
        handles.append(mpatches.Patch(color=COLORS[i % len(COLORS)], label=str(lvl)))

    ax.set_xlabel(quant_x)
    ax.set_ylabel(quant_y)
    ax.set_title(f"Связь «{quant_x}» и «{quant_y}» (цвет — {qual_col})")
    ax.legend(handles=handles, title=qual_col,
              bbox_to_anchor=(1.01, 1), loc="upper left",
              facecolor=BG_AXES, labelcolor=FG_TEXT,
              title_fontsize=8, fontsize=8)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:,.0f}"))
    fig.tight_layout()
    return fig


## \brief Круговая (кольцевая) диаграмма долей с группировкой «Прочие».
#  \param data Словарь справочников.
#  \param qual_col Качественный атрибут (срезы).
#  \param value_col Количественный атрибут (доли).
#  \param aggfunc Функция агрегации.
#  \param top_n Максимум срезов; остальное сводится в «Прочие».
#  \return Фигура matplotlib.
def chart_pie(
        data: dict,
        qual_col:  str = "НАЗВ_КАТ",
        value_col: str = "ОБЪЕМ_ТОРГОВ",
        aggfunc:   str = "sum",
        top_n:     int = 8,
) -> plt.Figure:
    df     = build_full_table(data)
    series = df.groupby(qual_col)[value_col].agg(aggfunc).sort_values(ascending=False)

    if len(series) > top_n:
        top    = series.iloc[:top_n]
        others = pd.Series({"Прочие": series.iloc[top_n:].sum()})
        series = pd.concat([top, others])

    labels = series.index.tolist()
    values = series.values

    fig, ax = plt.subplots(figsize=(8, 6))
    fig.patch.set_facecolor(BG_FIGURE)
    ax.set_facecolor(BG_FIGURE)

    wedge_colors = COLORS[:len(labels)]
    wedges, texts, autotexts = ax.pie(
        values,
        labels=labels,
        colors=wedge_colors,
        autopct="%1.1f%%",
        startangle=140,
        pctdistance=0.80,
        wedgeprops=dict(width=0.55, edgecolor=BG_FIGURE, linewidth=1.2),
    )
    for txt in texts:
        txt.set_color(FG_TEXT)
        txt.set_fontsize(8)
    for atxt in autotexts:
        atxt.set_color(BG_FIGURE)
        atxt.set_fontsize(7.5)
        atxt.set_fontweight("bold")

    ax.set_title(
        f"Доли «{value_col}» ({aggfunc}) по «{qual_col}»",
        color=FG_TEXT, fontsize=12, pad=16,
    )
    fig.tight_layout()
    return fig


if __name__ == "__main__":
    db_path = sys.argv[1] if len(sys.argv) > 1 else None
    try:
        data = load_data(db_path)
    except FileNotFoundError as e:
        print(e); sys.exit(1)

    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_charts")
    os.makedirs(out_dir, exist_ok=True)

    charts = [
        (chart_clustered_bar,           "chart1_bar.png"),
        (chart_categorized_histogram,   "chart2_hist.png"),
        (chart_boxplot,                 "chart3_box.png"),
        (chart_scatter,                 "chart4_scatter.png"),
        (chart_pie,                     "chart5_pie.png"),
    ]

    for func, fname in charts:
        fig = func(data)
        path = os.path.join(out_dir, fname)
        save_figure(fig, path)
        plt.close(fig)
        print(f"{path}")

    print(f"\nВсе графики сохранены: {out_dir}")
