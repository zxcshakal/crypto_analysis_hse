## \file main.py
#  \brief Главное GUI-приложение «Криптовалютный аналитик» (Tkinter).
#
#  Возможности:
#   * Регистрация при первом запуске / вход по логину и паролю (auth.py).
#   * Личный кабинет (клик по нику в шапке): профиль, KYC, избранное, аватар.
#   * Обновление данных из CoinGecko API с офлайн-фолбэком (api_client.py).
#   * Вкладка новостей с публикацией своих новостей (news_store.py + сервер).
#   * Тёмная / светлая тема с переключением и лёгкими анимациями.
#   * Существующие отчёты и графики сохранены без изменения логики.
#
#  Используются только стандартная библиотека + pandas/numpy/matplotlib/openpyxl.
#
#  \author Хотнянский Кирилл, Прохачев Никита, Фахрутдинов Амир
#  \date 2026

import os
import sys
import pickle
import shutil

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Модули проекта вынесены в подпапку scripts/ — добавляем её в путь поиска.
if not getattr(sys, "frozen", False):
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts"))
from data_loader import load_data, build_full_table
from text_reports import (
    report_simple, report_qualitative_stats,
    report_quantitative_stats, report_pivot, save_report,
)
from graphic_reports import (
    chart_clustered_bar, chart_categorized_histogram,
    chart_boxplot, chart_scatter, chart_pie, save_figure,
)
import auth
import api_client
import news_store

## Каталоги.
#  Корень проекта — рядом с main.py (или рядом с .exe в сборке).
#  Все данные (основная БД, аккаунты, новости, сессия, аватары) — в database/.
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR      = os.path.join(BASE_DIR, "database")
DATA_DIR    = DB_DIR
AVATAR_DIR  = os.path.join(DB_DIR, "avatars")
CONFIG_PATH = os.path.join(DB_DIR, "config.pkl")
DB_PATH     = os.path.join(DB_DIR, "crypto_database.xlsx")
os.makedirs(DB_DIR, exist_ok=True)
os.makedirs(AVATAR_DIR, exist_ok=True)

# В .exe база лежит внутри сборки (_MEIPASS). При первом запуске копируем её
# в database/, чтобы можно было дописывать листы API и работать офлайн.
if getattr(sys, "frozen", False) and not os.path.exists(DB_PATH):
    _bundled = os.path.join(getattr(sys, "_MEIPASS", BASE_DIR), "crypto_database.xlsx")
    if os.path.exists(_bundled):
        shutil.copy(_bundled, DB_PATH)

## Тёмная и светлая палитры интерфейса.
THEMES = {
    "dark": {
        "bg": "#1e1e2e", "bg2": "#11111b", "surface": "#313244",
        "overlay": "#45475a", "text": "#cdd6f4", "subtext": "#6c7086",
        "tree_bg": "#181825",
        "blue": "#89b4fa", "green": "#a6e3a1", "red": "#f38ba8",
        "peach": "#fab387", "mauve": "#cba6f7",
    },
    "light": {
        "bg": "#eff1f5", "bg2": "#dce0e8", "surface": "#ccd0da",
        "overlay": "#bcc0cc", "text": "#4c4f69", "subtext": "#6c6f85",
        "tree_bg": "#ffffff",
        "blue": "#1e66f5", "green": "#40a02b", "red": "#d20f39",
        "peach": "#fe640b", "mauve": "#8839ef",
    },
}

## Активная палитра (мутируется на месте при смене темы, чтобы все ссылки на C
#  получали актуальные цвета после пересборки интерфейса).
C = dict(THEMES["dark"])

QUAL_COLS  = ["НАЗВ_КРИПТО", "ТИКЕР", "НАЗВ_КАТ", "НАЗВ_БИРЖИ", "СТРАНА", "АЛГОРИТМ"]
QUANT_COLS = ["ОБЪЕМ_ТОРГОВ", "ЦЕНА_ЗАКРЫТИЯ", "КОЛ_СДЕЛОК", "МАКС_ЦЕНА", "МИН_ЦЕНА", "ЦЕНА_ОТКРЫТИЯ"]
AGG_OPTS   = ["mean", "sum", "count", "max", "min"]

SAVE_FORMATS_TEXT  = [("Excel", "*.xlsx"), ("CSV", "*.csv"), ("Текст (TSV)", "*.txt")]
SAVE_FORMATS_IMAGE = [("PNG", "*.png"), ("JPEG", "*.jpg"), ("PDF", "*.pdf"), ("SVG", "*.svg")]


## \brief Прочитать конфиг приложения (тема, адрес сервера).
#  \return Словарь конфигурации.
def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "rb") as f:
                return pickle.load(f)
        except Exception:
            pass
    return {"theme": "dark", "server_url": ""}


## \brief Сохранить конфиг приложения.
#  \param cfg Словарь конфигурации.
#  \return None
def save_config(cfg: dict) -> None:
    with open(CONFIG_PATH, "wb") as f:
        pickle.dump(cfg, f)


## \class App
#  \brief Корневое окно: управляет экраном входа и главным интерфейсом.
class App(tk.Tk):

    ## \brief Инициализация приложения, загрузка конфига и показ нужного экрана.
    def __init__(self):
        super().__init__()
        self.title("Криптовалютный аналитик — БИВ253")
        self.geometry("1240x780")
        self.configure(bg=C["bg"])

        self.cfg = load_config()
        self._set_theme(self.cfg.get("theme", "dark"), rebuild=False)
        news_store.SERVER_URL = self.cfg.get("server_url", "")

        self.user = auth.load_session()      # «запомнить вход»
        self._avatar_cache = None            # держим ссылку на PhotoImage

        if self.user and auth.get_profile(self.user):
            self.show_main()
        else:
            self.show_auth()

    # ------------------------------------------------------------------ темы
    ## \brief Установить активную тему (мутирует C на месте).
    #  \param name "dark" или "light".
    #  \param rebuild Пересобрать ли интерфейс немедленно.
    #  \return None
    def _set_theme(self, name: str, rebuild: bool = True):
        C.clear(); C.update(THEMES.get(name, THEMES["dark"]))
        self.cfg["theme"] = name
        save_config(self.cfg)
        self.configure(bg=C["bg"])
        if rebuild:
            if self.user:
                self.show_main()
            else:
                self.show_auth()

    ## \brief Переключить тему «тёмная ↔ светлая».
    #  \return None
    def toggle_theme(self):
        self._set_theme("light" if self.cfg.get("theme") == "dark" else "dark")

    ## \brief Применить ttk-стили под текущую тему.
    #  \return None
    def _apply_style(self):
        st = ttk.Style()
        st.theme_use("default")
        st.configure("TNotebook", background=C["bg"], borderwidth=0)
        st.configure("TNotebook.Tab", background=C["surface"], foreground=C["text"],
                     padding=[12, 5], font=("Arial", 10))
        st.map("TNotebook.Tab",
               background=[("selected", C["blue"])],
               foreground=[("selected", C["bg"])])
        st.configure("Treeview", background=C["tree_bg"], foreground=C["text"],
                     fieldbackground=C["tree_bg"], rowheight=22, font=("Arial", 9))
        st.configure("Treeview.Heading", background=C["surface"],
                     foreground=C["blue"], font=("Arial", 9, "bold"))
        st.map("Treeview", background=[("selected", C["overlay"])])
        st.configure("TCombobox", fieldbackground=C["surface"], background=C["surface"])

    # ------------------------------------------------------------- анимации
    ## \brief Плавное появление окна (alpha 0 → 1).
    #  \param win Окно tk.Tk/Toplevel.
    #  \return None
    def _fade_in(self, win):
        try:
            win.attributes("-alpha", 0.0)
            def step(a=0.0):
                a = min(a + 0.12, 1.0)
                win.attributes("-alpha", a)
                if a < 1.0:
                    win.after(16, lambda: step(a))
            step()
        except tk.TclError:
            pass     # некоторые WM не поддерживают alpha — не критично

    ## \brief Навесить hover-подсветку на кнопку.
    #  \param btn Кнопка.
    #  \param base Базовый цвет.
    #  \param hot Цвет при наведении.
    #  \return Кнопка (для цепочечного вызова).
    def _hover(self, btn, base, hot):
        btn.bind("<Enter>", lambda e: btn.config(bg=hot))
        btn.bind("<Leave>", lambda e: btn.config(bg=base))
        return btn

    ## \brief Всплывающее уведомление (toast) внизу справа.
    #  \param msg Текст.
    #  \param color Цвет фона.
    #  \return None
    def _toast(self, msg, color=None):
        color = color or C["green"]
        t = tk.Toplevel(self)
        t.overrideredirect(True)
        t.configure(bg=color)
        tk.Label(t, text=msg, bg=color, fg=C["bg"],
                 font=("Arial", 10, "bold"), padx=16, pady=8).pack()
        self.update_idletasks()
        x = self.winfo_x() + self.winfo_width() - t.winfo_reqwidth() - 30
        y = self.winfo_y() + self.winfo_height() - t.winfo_reqheight() - 40
        t.geometry(f"+{x}+{y}")
        self._fade_in(t)
        t.after(2200, t.destroy)

    # --------------------------------------------------------- экран входа
    ## \brief Очистить корневое окно от всех виджетов.
    #  \return None
    def _clear_root(self):
        for w in self.winfo_children():
            w.destroy()

    ## \brief Построить экран входа/регистрации.
    #  \return None
    def show_auth(self):
        self._clear_root()
        self._apply_style()
        self._fade_in(self)

        wrap = tk.Frame(self, bg=C["bg"])
        wrap.place(relx=0.5, rely=0.5, anchor="center")

        card = tk.Frame(wrap, bg=C["bg2"], padx=40, pady=34)
        card.pack()

        tk.Label(card, text="₿  Криптовалютный аналитик", font=("Arial", 18, "bold"),
                 fg=C["blue"], bg=C["bg2"]).pack(pady=(0, 4))
        first = auth.is_first_launch()
        sub = "Первый запуск — создайте аккаунт" if first else "Вход в систему"
        tk.Label(card, text=sub, font=("Arial", 10),
                 fg=C["subtext"], bg=C["bg2"]).pack(pady=(0, 18))

        self.auth_login = tk.StringVar()
        self.auth_pass  = tk.StringVar()
        self.auth_fio   = tk.StringVar()
        self.auth_mode  = "register" if first else "login"

        def field(label, var, show=None):
            tk.Label(card, text=label, fg=C["text"], bg=C["bg2"],
                     font=("Arial", 9), anchor="w").pack(fill="x")
            e = tk.Entry(card, textvariable=var, show=show, width=30,
                         bg=C["surface"], fg=C["text"], insertbackground=C["text"],
                         relief="flat", font=("Arial", 11))
            e.pack(pady=(2, 12), ipady=4)
            return e

        field("Логин (никнейм)", self.auth_login)
        field("Пароль", self.auth_pass, show="•")
        self.auth_fio_entry = None
        if self.auth_mode == "register":
            self.auth_fio_entry = field("ФИО (необязательно)", self.auth_fio)

        btn = tk.Button(card,
                        text="Зарегистрироваться" if first else "Войти",
                        command=self._submit_auth,
                        bg=C["green"], fg=C["bg"], font=("Arial", 11, "bold"),
                        relief="flat", width=26, pady=6, cursor="hand2")
        btn.pack(pady=(6, 8))
        self._hover(btn, C["green"], C["blue"])

        self.auth_switch = tk.Label(
            card,
            text=("Уже есть аккаунт? Войти" if first
                  else "Нет аккаунта? Зарегистрироваться"),
            fg=C["mauve"], bg=C["bg2"], font=("Arial", 9, "underline"),
            cursor="hand2")
        self.auth_switch.pack()
        self.auth_switch.bind("<Button-1>", lambda e: self._toggle_auth_mode())

        tk.Button(card, text="◐ Сменить тему", command=self.toggle_theme,
                  bg=C["surface"], fg=C["text"], relief="flat",
                  font=("Arial", 8), cursor="hand2").pack(pady=(16, 0))

        self.bind("<Return>", lambda e: self._submit_auth())

    ## \brief Переключить режим экрана между входом и регистрацией.
    #  \return None
    def _toggle_auth_mode(self):
        self.auth_mode = "login" if self.auth_mode == "register" else "register"
        # перерисуем экран в нужном режиме
        self._clear_root(); self._apply_style()
        # эмулируем «не первый запуск» в выбранном режиме
        wrap = tk.Frame(self, bg=C["bg"]); wrap.place(relx=0.5, rely=0.5, anchor="center")
        card = tk.Frame(wrap, bg=C["bg2"], padx=40, pady=34); card.pack()
        reg = self.auth_mode == "register"
        tk.Label(card, text="₿  Криптовалютный аналитик", font=("Arial", 18, "bold"),
                 fg=C["blue"], bg=C["bg2"]).pack(pady=(0, 4))
        tk.Label(card, text=("Создание аккаунта" if reg else "Вход в систему"),
                 font=("Arial", 10), fg=C["subtext"], bg=C["bg2"]).pack(pady=(0, 18))

        def field(label, var, show=None):
            tk.Label(card, text=label, fg=C["text"], bg=C["bg2"],
                     font=("Arial", 9), anchor="w").pack(fill="x")
            tk.Entry(card, textvariable=var, show=show, width=30,
                     bg=C["surface"], fg=C["text"], insertbackground=C["text"],
                     relief="flat", font=("Arial", 11)).pack(pady=(2, 12), ipady=4)

        field("Логин (никнейм)", self.auth_login)
        field("Пароль", self.auth_pass, show="•")
        if reg:
            field("ФИО (необязательно)", self.auth_fio)

        b = tk.Button(card, text=("Зарегистрироваться" if reg else "Войти"),
                      command=self._submit_auth, bg=C["green"], fg=C["bg"],
                      font=("Arial", 11, "bold"), relief="flat", width=26,
                      pady=6, cursor="hand2")
        b.pack(pady=(6, 8)); self._hover(b, C["green"], C["blue"])

        sw = tk.Label(card, text=("Уже есть аккаунт? Войти" if reg
                                  else "Нет аккаунта? Зарегистрироваться"),
                      fg=C["mauve"], bg=C["bg2"], font=("Arial", 9, "underline"),
                      cursor="hand2")
        sw.pack(); sw.bind("<Button-1>", lambda e: self._toggle_auth_mode())

    ## \brief Обработать отправку формы входа/регистрации.
    #  \return None
    def _submit_auth(self):
        login = self.auth_login.get().strip()
        pwd   = self.auth_pass.get()
        if self.auth_mode == "register":
            ok, msg = auth.register(login, pwd, self.auth_fio.get())
            if not ok:
                messagebox.showerror("Регистрация", msg); return
            auth.save_session(login)
            self.user = login
            self._toast("Аккаунт создан!")
            self.show_main()
        else:
            ok, msg = auth.login(login, pwd)
            if not ok:
                messagebox.showerror("Вход", msg); return
            auth.save_session(login)
            self.user = login
            self.show_main()

    ## \brief Выйти из аккаунта и вернуться на экран входа.
    #  \return None
    def logout(self):
        auth.clear_session()
        self.user = None
        self.show_auth()

    # --------------------------------------------------------- главный экран
    ## \brief Построить главный интерфейс (шапка + вкладки).
    #  \return None
    def show_main(self):
        self._clear_root()
        self._apply_style()
        self.unbind("<Return>")
        try:
            self.data = load_data(DB_PATH)
        except FileNotFoundError:
            messagebox.showerror("Ошибка", "crypto_database.xlsx не найден рядом с main.py")
            return
        self.full_df = build_full_table(self.data)
        self._build_header()
        self._build_notebook()
        self._fade_in(self)

    ## \brief Построить шапку: заголовок, кликабельный ник, тема, выход.
    #  \return None
    def _build_header(self):
        hdr = tk.Frame(self, bg=C["bg2"])
        hdr.pack(fill="x")
        left = tk.Frame(hdr, bg=C["bg2"]); left.pack(side="left", padx=14, pady=8)
        tk.Label(left, text="Аналитическое приложение — Криптовалютный рынок",
                 font=("Arial", 14, "bold"), fg=C["text"], bg=C["bg2"]).pack(anchor="w")
        tk.Label(left, text="Группа БИВ253", font=("Arial", 9),
                 fg=C["subtext"], bg=C["bg2"]).pack(anchor="w")

        right = tk.Frame(hdr, bg=C["bg2"]); right.pack(side="right", padx=14)
        theme_btn = tk.Button(right, text="◐ Тема", command=self.toggle_theme,
                              bg=C["surface"], fg=C["text"], relief="flat",
                              font=("Arial", 9), cursor="hand2", padx=8)
        theme_btn.pack(side="left", padx=4)
        out_btn = tk.Button(right, text="Выйти", command=self.logout,
                            bg=C["red"], fg=C["bg"], relief="flat",
                            font=("Arial", 9, "bold"), cursor="hand2", padx=8)
        out_btn.pack(side="left", padx=4)
        self._hover(out_btn, C["red"], C["peach"])

        # кликабельный никнейм -> личный кабинет
        nick = tk.Label(right, text=f"  👤 {self.user}  ", fg=C["bg"], bg=C["blue"],
                        font=("Arial", 10, "bold"), cursor="hand2", padx=6, pady=2)
        nick.pack(side="left", padx=8)
        nick.bind("<Button-1>", lambda e: self.open_cabinet())
        self._hover(nick, C["blue"], C["mauve"])

    ## \brief Построить notebook со всеми вкладками.
    #  \return None
    def _build_notebook(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        tabs = {
            "tab_data":   ("Данные / API",        self._build_tab_data),
            "tab_news":   ("Новости",             self._build_tab_news),
            "tab_r1":     ("Торговый отчёт",      self._build_tab_r1),
            "tab_r2":     ("Статистика",          self._build_tab_r2),
            "tab_r3":     ("Сводная таблица",     self._build_tab_r3),
            "tab_g_bar":  ("Столбчатая",          self._build_tab_bar),
            "tab_g_hist": ("Гистограмма",         self._build_tab_hist),
            "tab_g_box":  ("Бокса-Вискера",       self._build_tab_box),
            "tab_g_sct":  ("Рассеивание",         self._build_tab_scatter),
            "tab_g_pie":  ("Круговая",            self._build_tab_pie),
        }
        for attr, (label, builder) in tabs.items():
            frame = tk.Frame(nb, bg=C["bg"])
            setattr(self, attr, frame)
            nb.add(frame, text=f" {label} ")
            builder(frame)

    # ------------------------------------------------- личный кабинет
    ## \brief Открыть окно личного кабинета (профиль, KYC, избранное, аватар).
    #  \return None
    def open_cabinet(self):
        prof = auth.get_profile(self.user) or {}
        win = tk.Toplevel(self)
        win.title("Личный кабинет")
        win.geometry("520x640")
        win.configure(bg=C["bg"])
        self._fade_in(win)

        # --- аватар ---
        top = tk.Frame(win, bg=C["bg2"]); top.pack(fill="x", pady=(0, 8))
        self.cab_avatar_lbl = tk.Label(top, bg=C["surface"], width=14, height=7)
        self.cab_avatar_lbl.pack(pady=12)
        self._load_avatar_into(self.cab_avatar_lbl, prof.get("АВАТАР", ""))
        ch = tk.Button(top, text="Сменить аватар", command=lambda: self._change_avatar(win),
                       bg=C["blue"], fg=C["bg"], relief="flat",
                       font=("Arial", 9, "bold"), cursor="hand2")
        ch.pack(pady=(0, 10)); self._hover(ch, C["blue"], C["mauve"])
        tk.Label(top, text=self.user, font=("Arial", 14, "bold"),
                 fg=C["text"], bg=C["bg2"]).pack(pady=(0, 10))

        # --- избранное ---
        favf = tk.LabelFrame(win, text="Избранное", fg=C["blue"], bg=C["bg"],
                             font=("Arial", 10, "bold"), padx=10, pady=6)
        favf.pack(fill="x", padx=14, pady=6)
        fav_c = auth.get_favorites(prof, "ИЗБР_КРИПТО")
        fav_e = auth.get_favorites(prof, "ИЗБР_БИРЖИ")
        tk.Label(favf, text="Криптовалюты: " + (", ".join(fav_c) or "—"),
                 fg=C["text"], bg=C["bg"], font=("Arial", 9), anchor="w").pack(fill="x")
        tk.Label(favf, text="Биржи: " + (", ".join(fav_e) or "—"),
                 fg=C["text"], bg=C["bg"], font=("Arial", 9), anchor="w").pack(fill="x")

        # --- личная информация ---
        info = tk.LabelFrame(win, text="Личная информация", fg=C["blue"], bg=C["bg"],
                             font=("Arial", 10, "bold"), padx=10, pady=6)
        info.pack(fill="x", padx=14, pady=6)

        self.cab_vars = {}
        rows = [("ФИО", "ФИО"), ("Дата рождения", "ДАТА_РОЖДЕНИЯ"),
                ("Телефон", "ТЕЛЕФОН"), ("Страна", "СТРАНА")]
        for label, key in rows:
            r = tk.Frame(info, bg=C["bg"]); r.pack(fill="x", pady=3)
            tk.Label(r, text=label, width=16, anchor="w",
                     fg=C["text"], bg=C["bg"], font=("Arial", 9)).pack(side="left")
            v = tk.StringVar(value=prof.get(key, ""))
            self.cab_vars[key] = v
            tk.Entry(r, textvariable=v, bg=C["surface"], fg=C["text"],
                     insertbackground=C["text"], relief="flat",
                     font=("Arial", 10)).pack(side="left", fill="x", expand=True, ipady=2)

        # KYC — только чтение (ставит администратор вручную в Excel)
        kyc = prof.get("KYC", "Нет")
        kyc_color = C["green"] if str(kyc).lower().startswith(("да", "yes", "verif")) else C["red"]
        r = tk.Frame(info, bg=C["bg"]); r.pack(fill="x", pady=8)
        tk.Label(r, text="Верификация KYC", width=16, anchor="w",
                 fg=C["text"], bg=C["bg"], font=("Arial", 9)).pack(side="left")
        tk.Label(r, text=f" {kyc} ", fg=C["bg"], bg=kyc_color,
                 font=("Arial", 9, "bold")).pack(side="left")
        tk.Label(info, text="(KYC задаётся администратором в базе данных)",
                 fg=C["subtext"], bg=C["bg"], font=("Arial", 8)).pack(anchor="w")

        save = tk.Button(win, text="Сохранить профиль",
                         command=lambda: self._save_cabinet(win),
                         bg=C["green"], fg=C["bg"], relief="flat",
                         font=("Arial", 10, "bold"), cursor="hand2", pady=6)
        save.pack(pady=14, padx=14, fill="x")
        self._hover(save, C["green"], C["blue"])

    ## \brief Загрузить аватар (PNG/GIF) в Label, либо показать заглушку.
    #  \param label Целевой Label.
    #  \param path Путь к файлу аватара.
    #  \return None
    def _load_avatar_into(self, label, path):
        if path and os.path.exists(path):
            try:
                img = tk.PhotoImage(file=path)
                # ужимаем крупные изображения до миниатюры
                factor = max(img.width() // 120, 1)
                if factor > 1:
                    img = img.subsample(factor, factor)
                self._avatar_cache = img
                label.config(image=img, text="", width=img.width(), height=img.height())
                return
            except tk.TclError:
                pass
        label.config(image="", text="нет\nфото", fg=C["subtext"],
                     font=("Arial", 10), width=14, height=7)

    ## \brief Выбрать новый файл аватара, скопировать и сохранить путь.
    #  \param win Окно кабинета (для обновления).
    #  \return None
    def _change_avatar(self, win):
        path = filedialog.askopenfilename(
            title="Выберите изображение (PNG или GIF)",
            filetypes=[("Изображения", "*.png *.gif")])
        if not path:
            return
        ext = os.path.splitext(path)[-1].lower()
        dest = os.path.join(AVATAR_DIR, f"{self.user}{ext}")
        try:
            shutil.copyfile(path, dest)
        except OSError as e:
            messagebox.showerror("Аватар", str(e)); return
        auth.update_profile(self.user, {"АВАТАР": dest})
        self._load_avatar_into(self.cab_avatar_lbl, dest)
        self._toast("Аватар обновлён")

    ## \brief Сохранить изменённые поля профиля из кабинета.
    #  \param win Окно кабинета.
    #  \return None
    def _save_cabinet(self, win):
        fields = {k: v.get() for k, v in self.cab_vars.items()}
        auth.update_profile(self.user, fields)
        self._toast("Профиль сохранён")
        win.destroy()

    # ----------------------------------------------------- вкладка новостей
    ## \brief Построить вкладку новостей: список + форма публикации.
    #  \param f Родительский фрейм.
    #  \return None
    def _build_tab_news(self, f):
        self._section_title(f, "Новости сообщества")

        bar = tk.Frame(f, bg=C["bg"]); bar.pack(fill="x", padx=12, pady=4)
        self._btn(bar, "Обновить с сервера", self._refresh_news, C["blue"]).pack(side="left", padx=4)
        self._btn(bar, "Указать сервер", self._set_server, C["surface"]).pack(side="left", padx=4)
        srv = news_store.SERVER_URL or "офлайн (только локально)"
        tk.Label(bar, text=f"Сервер: {srv}", fg=C["subtext"], bg=C["bg"],
                 font=("Arial", 8)).pack(side="left", padx=8)

        # форма публикации
        form = self._param_panel(f, "Написать новость")
        form.pack(fill="x", padx=12, pady=4)
        self.news_title = tk.StringVar()
        r = tk.Frame(form, bg=C["bg"]); r.pack(fill="x", pady=2)
        tk.Label(r, text="Заголовок:", width=12, anchor="w",
                 fg=C["text"], bg=C["bg"]).pack(side="left")
        tk.Entry(r, textvariable=self.news_title, bg=C["surface"], fg=C["text"],
                 insertbackground=C["text"], relief="flat").pack(side="left",
                 fill="x", expand=True, ipady=2)
        self.news_text = tk.Text(form, height=3, bg=C["surface"], fg=C["text"],
                                 insertbackground=C["text"], relief="flat",
                                 font=("Arial", 10))
        self.news_text.pack(fill="x", pady=4)
        self._btn(form, "Опубликовать", self._post_news, C["green"]).pack(anchor="e", pady=2)

        # список новостей
        self.news_list = tk.Frame(f, bg=C["bg"])
        self.news_list.pack(fill="both", expand=True, padx=12, pady=6)
        self._render_news()

    ## \brief Отрисовать список новостей в виде карточек.
    #  \return None
    def _render_news(self):
        for w in self.news_list.winfo_children():
            w.destroy()
        canvas = tk.Canvas(self.news_list, bg=C["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(self.news_list, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=C["bg"])
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        df = news_store.load_local_news()
        if df.empty:
            tk.Label(inner, text="Пока нет новостей. Будьте первым!",
                     fg=C["subtext"], bg=C["bg"], font=("Arial", 11)).pack(pady=20)
            return
        for _, row in df.iterrows():
            card = tk.Frame(inner, bg=C["bg2"], padx=12, pady=8)
            card.pack(fill="x", pady=4, padx=2)
            tk.Label(card, text=row["ЗАГОЛОВОК"], font=("Arial", 11, "bold"),
                     fg=C["blue"], bg=C["bg2"], anchor="w").pack(fill="x")
            tk.Label(card, text=f"{row['АВТОР']} · {row['ДАТА']}", font=("Arial", 8),
                     fg=C["subtext"], bg=C["bg2"], anchor="w").pack(fill="x")
            if row["ТЕКСТ"]:
                tk.Label(card, text=row["ТЕКСТ"], font=("Arial", 10), justify="left",
                         fg=C["text"], bg=C["bg2"], anchor="w", wraplength=900).pack(fill="x", pady=(4, 0))

    ## \brief Опубликовать новость от текущего пользователя.
    #  \return None
    def _post_news(self):
        ok, msg = news_store.add_news(self.user, self.news_title.get(),
                                      self.news_text.get("1.0", "end").strip())
        if not ok:
            messagebox.showwarning("Новости", msg); return
        self.news_title.set(""); self.news_text.delete("1.0", "end")
        self._toast(msg)
        self._render_news()

    ## \brief Синхронизировать новости с сервера и перерисовать список.
    #  \return None
    def _refresh_news(self):
        _, src = news_store.sync_from_server()
        self._toast("Обновлено: " + ("сервер" if src == "server" else "локально"))
        self._render_news()

    ## \brief Указать адрес сервера новостей и сохранить его в конфиг.
    #  \return None
    def _set_server(self):
        cur = news_store.SERVER_URL
        url = simpledialog.askstring(
            "Сервер новостей",
            "Адрес сервера (например http://192.168.0.10:8765).\n"
            "Оставьте пустым для офлайн-режима:",
            initialvalue=cur, parent=self)
        if url is None:
            return
        url = url.strip().rstrip("/")
        news_store.SERVER_URL = url
        self.cfg["server_url"] = url
        save_config(self.cfg)
        self._toast("Сервер сохранён" if url else "Офлайн-режим")
        self.show_main()

    # ----------------------------------------------- общие GUI-помощники
    @staticmethod
    def _lbl(parent, text, **kw):
        defaults = dict(fg=C["text"], bg=C["bg"], font=("Arial", 10))
        defaults.update(kw)
        return tk.Label(parent, text=text, **defaults)

    def _btn(self, parent, text, command, color=None, **kw):
        c = color or C["green"]
        b = tk.Button(parent, text=text, command=command, bg=c, fg=C["bg"],
                      font=("Arial", 10, "bold"), relief="flat", padx=12,
                      cursor="hand2", **kw)
        self._hover(b, c, C["blue"])
        return b

    def _combo(self, parent, var, values, width=18):
        return ttk.Combobox(parent, textvariable=var, values=values,
                            width=width, state="readonly")

    def _row(self, parent, label_text, var, values, width_lbl=20, width_cb=18):
        r = tk.Frame(parent, bg=C["bg"]); r.pack(fill="x", pady=2)
        self._lbl(r, label_text, width=width_lbl, anchor="w").pack(side="left")
        self._combo(r, var, values, width=width_cb).pack(side="left", padx=4)
        return r

    def _treeview_frame(self, parent):
        f = tk.Frame(parent, bg=C["bg"]); f.pack(fill="both", expand=True, padx=12, pady=4)
        sb_y = ttk.Scrollbar(f, orient="vertical")
        sb_x = ttk.Scrollbar(f, orient="horizontal")
        tree = ttk.Treeview(f, yscrollcommand=sb_y.set,
                            xscrollcommand=sb_x.set, show="headings")
        sb_y.config(command=tree.yview); sb_x.config(command=tree.xview)
        sb_y.pack(side="right", fill="y"); sb_x.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        return tree

    def _fill_tree(self, tree, df: pd.DataFrame):
        tree.delete(*tree.get_children())
        df = df.reset_index() if df.index.name or isinstance(df.index, pd.MultiIndex) else df
        cols = list(df.columns)
        tree["columns"] = cols
        for col in cols:
            tree.heading(col, text=str(col))
            max_w = max(len(str(col)),
                        df[col].astype(str).str.len().max() if len(df) else len(str(col)))
            tree.column(col, width=min(int(max_w * 8) + 10, 220), minwidth=55)
        for _, row in df.iterrows():
            tree.insert("", "end", values=[str(v) for v in row])

    def _show_figure(self, fig, frame):
        for w in frame.winfo_children():
            w.destroy()
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        plt.close(fig)

    def _save_fig_dialog(self, fig):
        if fig is None:
            messagebox.showwarning("Пусто", "Сначала постройте график"); return
        path = filedialog.asksaveasfilename(defaultextension=".png",
                                            filetypes=SAVE_FORMATS_IMAGE, initialfile="chart")
        if path:
            save_figure(fig, path); self._toast("График сохранён")

    def _save_df_dialog(self, df, init="report"):
        if df is None or (isinstance(df, pd.DataFrame) and df.empty):
            messagebox.showwarning("Пусто", "Сначала сгенерируйте отчёт"); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=SAVE_FORMATS_TEXT, initialfile=init)
        if path:
            save_report(df, path); self._toast("Отчёт сохранён")

    def _section_title(self, parent, text):
        tk.Label(parent, text=text, font=("Arial", 12, "bold"),
                 fg=C["blue"], bg=C["bg"]).pack(pady=(12, 4))

    def _param_panel(self, parent, title="Параметры"):
        return tk.LabelFrame(parent, text=title, fg=C["text"], bg=C["bg"],
                             font=("Arial", 10), padx=10, pady=6)

    def _btn_row(self, parent, buttons: list):
        row = tk.Frame(parent, bg=C["bg"]); row.pack(pady=6)
        for text, cmd, color in buttons:
            self._btn(row, text, cmd, color).pack(side="left", padx=6)
        return row

    def _info_label(self, parent):
        lbl = tk.Label(parent, text="", fg=C["green"], bg=C["bg"], font=("Arial", 9))
        lbl.pack(); return lbl

    # ----------------------------------------------------- вкладка данных + API
    ## \brief Вкладка просмотра справочников + обновление из API.
    #  \param f Родительский фрейм.
    #  \return None
    def _build_tab_data(self, f):
        self._section_title(f, "Справочники базы данных и обновление из API")

        row = tk.Frame(f, bg=C["bg"]); row.pack(pady=4)
        self._lbl(row, "Таблица:").pack(side="left", padx=6)
        self.data_table_var = tk.StringVar(value="КРИПТОВАЛЮТЫ")
        tables = ["КРИПТОВАЛЮТЫ", "БИРЖИ", "КАТЕГОРИИ",
                  "ТОРГОВЫЕ_ПАРЫ", "ТОРГОВЫЕ_ДАННЫЕ", "API"]
        self._combo(row, self.data_table_var, tables, width=22).pack(side="left", padx=4)
        self._btn(row, "Показать", self._show_data_table).pack(side="left", padx=8)
        self._btn(row, "Обновить курсы из API", self._refresh_api, C["peach"]).pack(side="left", padx=8)

        self.data_info = self._info_label(f)
        self.data_tree = self._treeview_frame(f)
        self._show_data_table()

    ## \brief Запросить актуальные курсы из CoinGecko и сообщить источник.
    #  \return None
    def _refresh_api(self):
        df, src = api_client.refresh_reference(DB_PATH)
        if src == "api":
            self._toast("Курсы обновлены из CoinGecko")
        elif src == "cache":
            self._toast("API недоступен — показан кеш", C["peach"])
        else:
            self._toast("Нет данных: ни API, ни кеша", C["red"]); return
        self.data_table_var.set("КРИПТОВАЛЮТЫ")
        if df is not None:
            self._fill_tree(self.data_tree, df)
            self.data_info.config(text=f"API_КУРСЫ  |  источник: {src}  |  строк: {len(df)}")

    ## \brief Показать выбранный справочник базы.
    #  \return None
    def _show_data_table(self):
        sheet_map = {
            "КРИПТОВАЛЮТЫ": "cryptos", "БИРЖИ": "exchanges", "КАТЕГОРИИ": "categories",
            "ТОРГОВЫЕ_ПАРЫ": "pairs", "ТОРГОВЫЕ_ДАННЫЕ": "trades", "API": "apis",
        }
        name = self.data_table_var.get()
        df = self.data[sheet_map[name]].copy()
        if "ДАТА" in df.columns:
            df["ДАТА"] = df["ДАТА"].dt.strftime("%d.%m.%Y")
        self._fill_tree(self.data_tree, df)
        self.data_info.config(
            text=f"Таблица: {name}  |  Строк: {len(df)}  |  Столбцов: {len(df.columns)}")

    # --------------------------------------------------------- отчёт 1
    def _build_tab_r1(self, f):
        self._section_title(f, "Отчёт 1 — Торговые данные (проекция + фильтр)")
        panel = self._param_panel(f, "Параметры фильтрации"); panel.pack(fill="x", padx=12, pady=4)
        tickers   = ["Все"] + sorted(self.data["cryptos"]["ТИКЕР"].tolist())
        exchanges = ["Все"] + sorted(self.data["exchanges"]["НАЗВ_БИРЖИ"].tolist())
        self.r1_ticker = tk.StringVar(value="Все"); self.r1_exchange = tk.StringVar(value="Все")
        self.r1_from = tk.StringVar(value="2026-01-01"); self.r1_to = tk.StringVar(value="2026-01-31")

        r1 = tk.Frame(panel, bg=C["bg"]); r1.pack(fill="x", pady=3)
        self._lbl(r1, "Криптовалюта:", width=16, anchor="w").pack(side="left")
        self._combo(r1, self.r1_ticker, tickers, 14).pack(side="left", padx=4)
        r2 = tk.Frame(panel, bg=C["bg"]); r2.pack(fill="x", pady=3)
        self._lbl(r2, "Биржа:", width=16, anchor="w").pack(side="left")
        self._combo(r2, self.r1_exchange, exchanges, 14).pack(side="left", padx=4)
        r3 = tk.Frame(panel, bg=C["bg"]); r3.pack(fill="x", pady=3)
        self._lbl(r3, "Дата с:", width=16, anchor="w").pack(side="left")
        tk.Entry(r3, textvariable=self.r1_from, width=12, bg=C["surface"], fg=C["text"],
                 insertbackground=C["text"]).pack(side="left", padx=4)
        self._lbl(r3, "по:").pack(side="left", padx=4)
        tk.Entry(r3, textvariable=self.r1_to, width=12, bg=C["surface"], fg=C["text"],
                 insertbackground=C["text"]).pack(side="left", padx=4)

        self._btn_row(f, [
            ("Сгенерировать", self._gen_r1, C["green"]),
            ("Сохранить Excel", lambda: self._save_df_dialog(self.r1_df, "report1_trades"), C["blue"]),
        ])
        self.r1_info = self._info_label(f); self.r1_df = None
        self.r1_tree = self._treeview_frame(f); self._gen_r1()

    def _gen_r1(self):
        t = self.r1_ticker.get(); e = self.r1_exchange.get()
        df = report_simple(self.data, crypto_names=[t] if t != "Все" else None,
                           date_from=self.r1_from.get() or None,
                           date_to=self.r1_to.get() or None,
                           exchange_name=e if e != "Все" else None)
        df["ДАТА"] = df["ДАТА"].dt.strftime("%d.%m.%Y")
        self.r1_df = df; self._fill_tree(self.r1_tree, df)
        self.r1_info.config(text=f"Найдено строк: {len(df)}")

    # --------------------------------------------------------- отчёт 2
    def _build_tab_r2(self, f):
        self._section_title(f, "Отчёт 2 — Статистический анализ")
        panel = self._param_panel(f, "Параметры"); panel.pack(fill="x", padx=12, pady=4)
        self.stat_type = tk.StringVar(value="Качественные (частоты)")
        r1 = tk.Frame(panel, bg=C["bg"]); r1.pack(fill="x", pady=3)
        self._lbl(r1, "Тип:", width=18, anchor="w").pack(side="left")
        self._combo(r1, self.stat_type, ["Качественные (частоты)", "Количественные (описательные)"],
                    28).pack(side="left", padx=4)
        self.stat_col = tk.StringVar(value="НАЗВ_КАТ")
        r2 = tk.Frame(panel, bg=C["bg"]); r2.pack(fill="x", pady=3)
        self._lbl(r2, "Атрибут:", width=18, anchor="w").pack(side="left")
        self.stat_col_cb = self._combo(r2, self.stat_col, QUAL_COLS[:5], 28)
        self.stat_col_cb.pack(side="left", padx=4)
        self.stat_type.trace_add("write", self._update_stat_cols)
        self._btn_row(f, [
            ("Сгенерировать", self._gen_r2, C["green"]),
            ("Сохранить", lambda: self._save_df_dialog(self.r2_df, "report2_stats"), C["blue"]),
        ])
        self.r2_info = self._info_label(f); self.r2_df = None
        self.r2_tree = self._treeview_frame(f)

    def _update_stat_cols(self, *_):
        if self.stat_type.get().startswith("Кач"):
            cols = ["НАЗВ_КАТ", "АЛГОРИТМ", "НАЗВ_БИРЖИ", "СТРАНА", "ТИКЕР", "ТИП_ПАРЫ"]
        else:
            cols = QUANT_COLS
        self.stat_col.set(cols[0]); self.stat_col_cb["values"] = cols

    def _gen_r2(self):
        col = self.stat_col.get()
        if self.stat_type.get().startswith("Кач"):
            df = report_qualitative_stats(self.data, column=col)
            info = f"Таблица частот: {col}  |  Уровней: {len(df)}"
        else:
            df = report_quantitative_stats(self.data); info = "Описательные статистики"
        self.r2_df = df; self._fill_tree(self.r2_tree, df); self.r2_info.config(text=info)

    # --------------------------------------------------------- отчёт 3
    def _build_tab_r3(self, f):
        self._section_title(f, "Отчёт 3 — Сводная таблица (pivot_table)")
        panel = self._param_panel(f, "Параметры"); panel.pack(fill="x", padx=12, pady=4)
        val_cols = ["ЦЕНА_ЗАКРЫТИЯ", "ОБЪЕМ_ТОРГОВ", "КОЛ_СДЕЛОК", "МАКС_ЦЕНА", "МИН_ЦЕНА"]
        self.piv_index = tk.StringVar(value="НАЗВ_КРИПТО"); self.piv_cols = tk.StringVar(value="НАЗВ_БИРЖИ")
        self.piv_value = tk.StringVar(value="ОБЪЕМ_ТОРГОВ"); self.piv_agg = tk.StringVar(value="mean")
        self._row(panel, "Строки (index):", self.piv_index, QUAL_COLS)
        self._row(panel, "Столбцы (columns):", self.piv_cols, QUAL_COLS)
        self._row(panel, "Значение:", self.piv_value, val_cols)
        self._row(panel, "Агрегация:", self.piv_agg, AGG_OPTS, width_cb=10)
        self._btn_row(f, [
            ("Сгенерировать", self._gen_r3, C["green"]),
            ("Сохранить", lambda: self._save_df_dialog(self.r3_df, "report3_pivot"), C["blue"]),
        ])
        self.r3_info = self._info_label(f); self.r3_df = None
        self.r3_tree = self._treeview_frame(f)

    def _gen_r3(self):
        try:
            df = report_pivot(self.data, index=self.piv_index.get(), columns=self.piv_cols.get(),
                              values=self.piv_value.get(), aggfunc=self.piv_agg.get())
            self.r3_df = df
            flat = df.reset_index(); flat.columns = [str(c) for c in flat.columns]
            self._fill_tree(self.r3_tree, flat)
            self.r3_info.config(text=f"{self.piv_index.get()} × {self.piv_cols.get()} "
                                     f"| {self.piv_agg.get()}({self.piv_value.get()})")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # --------------------------------------------------------- графики
    def _build_tab_bar(self, f):
        self._section_title(f, "График 1 — Кластеризованная столбчатая диаграмма")
        panel = self._param_panel(f); panel.pack(fill="x", padx=12, pady=2)
        val_cols = ["ОБЪЕМ_ТОРГОВ", "ЦЕНА_ЗАКРЫТИЯ", "КОЛ_СДЕЛОК", "МАКС_ЦЕНА"]
        self.bar_x = tk.StringVar(value="НАЗВ_КРИПТО"); self.bar_grp = tk.StringVar(value="НАЗВ_БИРЖИ")
        self.bar_val = tk.StringVar(value="ОБЪЕМ_ТОРГОВ"); self.bar_agg = tk.StringVar(value="mean")
        self._row(panel, "Ось X (кач.):", self.bar_x, QUAL_COLS)
        self._row(panel, "Группа (кач.):", self.bar_grp, QUAL_COLS)
        self._row(panel, "Значение (кол.):", self.bar_val, val_cols)
        self._row(panel, "Агрегация:", self.bar_agg, AGG_OPTS, width_cb=10)
        self.bar_fig = None
        self._btn_row(f, [("Построить", self._draw_bar, C["green"]),
                          ("Сохранить граф", lambda: self._save_fig_dialog(self.bar_fig), C["red"])])
        self.bar_frame = tk.Frame(f, bg=C["bg"]); self.bar_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self._draw_bar()

    def _draw_bar(self):
        fig = chart_clustered_bar(self.data, qual_x=self.bar_x.get(), qual_group=self.bar_grp.get(),
                                  value_col=self.bar_val.get(), aggfunc=self.bar_agg.get())
        self.bar_fig = fig; self._show_figure(fig, self.bar_frame)

    def _build_tab_hist(self, f):
        self._section_title(f, "График 2 — Категоризированная гистограмма")
        panel = self._param_panel(f); panel.pack(fill="x", padx=12, pady=2)
        self.hist_quant = tk.StringVar(value="ЦЕНА_ЗАКРЫТИЯ"); self.hist_qual = tk.StringVar(value="ТИКЕР")
        self.hist_bins = tk.IntVar(value=25)
        self._row(panel, "Количественный:", self.hist_quant, QUANT_COLS)
        self._row(panel, "Качественный:", self.hist_qual, ["ТИКЕР", "НАЗВ_КАТ", "НАЗВ_БИРЖИ", "СТРАНА"])
        r3 = tk.Frame(panel, bg=C["bg"]); r3.pack(fill="x", pady=2)
        self._lbl(r3, "Число бинов:", width=20, anchor="w").pack(side="left")
        tk.Spinbox(r3, from_=5, to=150, textvariable=self.hist_bins, width=6,
                   bg=C["surface"], fg=C["text"], buttonbackground=C["surface"]).pack(side="left", padx=4)
        self.hist_fig = None
        self._btn_row(f, [("Построить", self._draw_hist, C["green"]),
                          ("Сохранить граф", lambda: self._save_fig_dialog(self.hist_fig), C["red"])])
        self.hist_frame = tk.Frame(f, bg=C["bg"]); self.hist_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self._draw_hist()

    def _draw_hist(self):
        fig = chart_categorized_histogram(self.data, quant_col=self.hist_quant.get(),
                                          qual_col=self.hist_qual.get(), bins=self.hist_bins.get())
        self.hist_fig = fig; self._show_figure(fig, self.hist_frame)

    def _build_tab_box(self, f):
        self._section_title(f, "График 3 — Диаграмма Бокса-Вискера")
        panel = self._param_panel(f); panel.pack(fill="x", padx=12, pady=2)
        self.box_quant = tk.StringVar(value="ОБЪЕМ_ТОРГОВ"); self.box_qual = tk.StringVar(value="НАЗВ_КАТ")
        self._row(panel, "Количественный (Y):", self.box_quant, QUANT_COLS)
        self._row(panel, "Качественный (X):", self.box_qual,
                  ["НАЗВ_КРИПТО", "ТИКЕР", "НАЗВ_КАТ", "НАЗВ_БИРЖИ", "СТРАНА"])
        self.box_fig = None
        self._btn_row(f, [("Построить", self._draw_box, C["green"]),
                          ("Сохранить граф", lambda: self._save_fig_dialog(self.box_fig), C["red"])])
        self.box_frame = tk.Frame(f, bg=C["bg"]); self.box_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self._draw_box()

    def _draw_box(self):
        fig = chart_boxplot(self.data, quant_col=self.box_quant.get(), qual_col=self.box_qual.get())
        self.box_fig = fig; self._show_figure(fig, self.box_frame)

    def _build_tab_scatter(self, f):
        self._section_title(f, "График 4 — Диаграмма рассеивания")
        panel = self._param_panel(f); panel.pack(fill="x", padx=12, pady=2)
        self.sct_x = tk.StringVar(value="ОБЪЕМ_ТОРГОВ"); self.sct_y = tk.StringVar(value="ЦЕНА_ЗАКРЫТИЯ")
        self.sct_col = tk.StringVar(value="ТИКЕР")
        self._row(panel, "Ось X (кол.):", self.sct_x, QUANT_COLS)
        self._row(panel, "Ось Y (кол.):", self.sct_y, QUANT_COLS)
        self._row(panel, "Цвет (кач.):", self.sct_col, ["ТИКЕР", "НАЗВ_КАТ", "НАЗВ_БИРЖИ", "СТРАНА"])
        self.sct_fig = None
        self._btn_row(f, [("Построить", self._draw_scatter, C["green"]),
                          ("Сохранить граф", lambda: self._save_fig_dialog(self.sct_fig), C["red"])])
        self.sct_frame = tk.Frame(f, bg=C["bg"]); self.sct_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self._draw_scatter()

    def _draw_scatter(self):
        fig = chart_scatter(self.data, quant_x=self.sct_x.get(), quant_y=self.sct_y.get(),
                            qual_col=self.sct_col.get())
        self.sct_fig = fig; self._show_figure(fig, self.sct_frame)

    def _build_tab_pie(self, f):
        self._section_title(f, "График 5 — Круговая диаграмма")
        panel = self._param_panel(f); panel.pack(fill="x", padx=12, pady=2)
        self.pie_qual = tk.StringVar(value="НАЗВ_КАТ"); self.pie_val = tk.StringVar(value="ОБЪЕМ_ТОРГОВ")
        self.pie_agg = tk.StringVar(value="sum"); self.pie_topn = tk.IntVar(value=8)
        self._row(panel, "Срезы (кач.):", self.pie_qual, ["НАЗВ_КАТ", "ТИКЕР", "НАЗВ_БИРЖИ", "СТРАНА", "АЛГОРИТМ"])
        self._row(panel, "Значение (кол.):", self.pie_val, QUANT_COLS)
        self._row(panel, "Агрегация:", self.pie_agg, AGG_OPTS, width_cb=10)
        r4 = tk.Frame(panel, bg=C["bg"]); r4.pack(fill="x", pady=2)
        self._lbl(r4, "Макс. срезов:", width=20, anchor="w").pack(side="left")
        tk.Spinbox(r4, from_=2, to=20, textvariable=self.pie_topn, width=5,
                   bg=C["surface"], fg=C["text"], buttonbackground=C["surface"]).pack(side="left", padx=4)
        self.pie_fig = None
        self._btn_row(f, [("Построить", self._draw_pie, C["green"]),
                          ("Сохранить граф", lambda: self._save_fig_dialog(self.pie_fig), C["red"])])
        self.pie_frame = tk.Frame(f, bg=C["bg"]); self.pie_frame.pack(fill="both", expand=True, padx=8, pady=4)
        self._draw_pie()

    def _draw_pie(self):
        fig = chart_pie(self.data, qual_col=self.pie_qual.get(), value_col=self.pie_val.get(),
                        aggfunc=self.pie_agg.get(), top_n=self.pie_topn.get())
        self.pie_fig = fig; self._show_figure(fig, self.pie_frame)


if __name__ == "__main__":
    app = App()
    app.mainloop()
