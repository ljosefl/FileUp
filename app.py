"""
Excel → SQL Server: загрузка данных и создание таблиц через GUI.
"""
from __future__ import annotations

import json
import os
import re
import threading
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
from tkinter import filedialog
from urllib.parse import quote_plus

import pandas as pd
import pyodbc
import ttkbootstrap as ttk
from sqlalchemy import create_engine
from ttkbootstrap.constants import BOTH, E, END, EW, LEFT, NSEW, RIGHT, VERTICAL, W
from ttkbootstrap.dialogs import Messagebox

APP_NAME = "FileUp"
APP_VERSION = "1.0"

# Публичный репозиторий на GitHub для проверки обновлений (API releases/latest).
# Переопределение: переменная окружения FILEUP_GITHUB_REPO.
GITHUB_REPO_FOR_UPDATES = "ljosefl/FileUp"

# Распространённые типы SQL Server (полное объявление для CREATE TABLE)
SQL_TYPES = (
    "BIGINT",
    "INT",
    "SMALLINT",
    "TINYINT",
    "BIT",
    "DECIMAL(18, 2)",
    "DECIMAL(19, 4)",
    "FLOAT",
    "REAL",
    "MONEY",
    "NVARCHAR(50)",
    "NVARCHAR(255)",
    "NVARCHAR(MAX)",
    "VARCHAR(100)",
    "CHAR(36)",
    "DATE",
    "DATETIME2(7)",
    "DATETIME2(0)",
    "SMALLDATETIME",
    "TIME(7)",
    "UNIQUEIDENTIFIER",
)


def bracket_ident(name: str) -> str:
    """Безопасное имя для SQL Server: [идентификатор]."""
    n = name.strip().replace("]", "]]")
    return f"[{n}]"


def _normalize_version_tuple(v: str) -> tuple[int, ...]:
    s = v.strip().lstrip("vV")
    nums: list[int] = []
    for part in re.split(r"[^\d]+", s):
        if part.isdigit():
            nums.append(int(part))
    return tuple(nums) if nums else (0,)


def _version_is_newer(remote: str, local: str) -> bool:
    tr, tl = _normalize_version_tuple(remote), _normalize_version_tuple(local)
    n = max(len(tr), len(tl))
    tr = tr + (0,) * (n - len(tr))
    tl = tl + (0,) * (n - len(tl))
    return tr > tl


def _resolve_github_repo() -> str | None:
    env = os.environ.get("FILEUP_GITHUB_REPO", "").strip()
    if env and "/" in env:
        return env
    cfg = (GITHUB_REPO_FOR_UPDATES or "").strip()
    if cfg and "/" in cfg:
        return cfg
    return None


def fetch_latest_github_release(repo: str) -> dict | None:
    """GET /repos/{owner}/{repo}/releases/latest"""
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    req = urllib.request.Request(url, headers={"Accept": "application/vnd.github+json", "User-Agent": f"{APP_NAME}/{APP_VERSION}"})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read().decode())
    except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, json.JSONDecodeError, OSError):
        return None


def pick_release_download_url(release: dict) -> str:
    """Прямая ссылка на .exe из вложений релиза, иначе страница релиза."""
    page = release.get("html_url") or ""
    for asset in release.get("assets") or []:
        name = (asset.get("name") or "").lower()
        if name.endswith(".exe"):
            return asset.get("browser_download_url") or page
    return page


class CreateTableDialog(ttk.Toplevel):
    """Диалог: схема, имя таблицы, опциональный Id IDENTITY, список столбцов."""

    def __init__(self, master: "UploadApp"):
        super().__init__(master)
        self.app = master
        self.title("Новая таблица")
        self.geometry("720x560")
        self.minsize(620, 480)
        self.transient(master)
        self.grab_set()

        self.table_schema = tk.StringVar(value=master.schema.get().strip() or "dbo")
        self.table_name = tk.StringVar()
        self.add_identity = tk.BooleanVar(value=False)
        self.column_rows: list[dict] = []

        pad = {"padx": 10, "pady": 6}
        frm = ttk.Frame(self, padding=15)
        frm.pack(fill=BOTH, expand=True)

        ttk.Label(frm, text="Создание таблицы", font=("", 13, "bold")).pack(anchor=W)

        meta = ttk.Frame(frm)
        meta.pack(fill=EW, **pad)
        meta.columnconfigure(1, weight=1)
        meta.columnconfigure(3, weight=1)

        ttk.Label(meta, text="Схема").grid(row=0, column=0, sticky=W, padx=(0, 8))
        ttk.Entry(meta, textvariable=self.table_schema, width=18).grid(row=0, column=1, sticky=EW)

        ttk.Label(meta, text="Имя таблицы").grid(row=0, column=2, sticky=W, padx=(16, 8))
        ttk.Entry(meta, textvariable=self.table_name, width=28).grid(row=0, column=3, sticky=EW)

        ttk.Checkbutton(
            frm,
            text="Добавить столбец Id — BIGINT IDENTITY(1,1) PRIMARY KEY",
            variable=self.add_identity,
        ).pack(anchor=W, **pad)

        hdr = ttk.Frame(frm)
        hdr.pack(fill=EW, padx=0, pady=(10, 4))
        ttk.Label(hdr, text="Столбцы", font=("", 11, "bold")).pack(side=LEFT)
        ttk.Button(hdr, text="+ Столбец", command=self._add_row, bootstyle="secondary-outline").pack(side=RIGHT)

        tip = ttk.Label(
            frm,
            text="Укажите имя столбца, тип SQL Server и допускает ли столбец NULL.",
            bootstyle="secondary",
        )
        tip.pack(anchor=W, pady=(0, 6))

        head = ttk.Frame(frm)
        head.pack(fill=EW)
        ttk.Label(head, text="Имя", width=26).pack(side=LEFT)
        ttk.Label(head, text="Тип данных", width=28).pack(side=LEFT, padx=(8, 0))
        ttk.Label(head, text="NULL").pack(side=LEFT, padx=(12, 0))

        scroll_host = ttk.Frame(frm)
        scroll_host.pack(fill=BOTH, expand=True, pady=(4, 8))

        canvas = tk.Canvas(scroll_host, highlightthickness=0, height=220)
        vsb = ttk.Scrollbar(scroll_host, orient=VERTICAL, command=canvas.yview)
        self._rows_inner = ttk.Frame(canvas)
        self._rows_inner.bind(
            "<Configure>",
            lambda _e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self._rows_inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)

        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        vsb.pack(side=RIGHT, fill="y")

        def _wheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_wheel(_e):
            canvas.bind_all("<MouseWheel>", _wheel)

        def _unbind_wheel(_e):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)

        btn_row = ttk.Frame(frm)
        btn_row.pack(fill=EW, pady=(8, 0))
        ttk.Button(btn_row, text="Удалить последний столбец", command=self._remove_last_row, bootstyle="secondary").pack(
            side=LEFT
        )
        ttk.Button(btn_row, text="Создать таблицу", command=self._apply, bootstyle="success").pack(side=RIGHT)
        ttk.Button(btn_row, text="Отмена", command=self._on_close, bootstyle="outline").pack(side=RIGHT, padx=(0, 8))

        self._canvas = canvas
        self._add_row()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        try:
            for w in (getattr(self, "_canvas", None),):
                if w is not None:
                    w.unbind_all("<MouseWheel>")
        except tk.TclError:
            pass
        self.destroy()

    def _add_row(self):
        row_f = ttk.Frame(self._rows_inner)
        row_f.pack(fill=EW, pady=2)

        name_v = tk.StringVar()
        type_v = tk.StringVar(value="INT")
        null_v = tk.BooleanVar(value=True)

        ent = ttk.Entry(row_f, textvariable=name_v, width=28)
        ent.pack(side=LEFT)

        cb = ttk.Combobox(row_f, textvariable=type_v, values=SQL_TYPES, width=26, state="readonly")
        cb.pack(side=LEFT, padx=(8, 0))

        chk = ttk.Checkbutton(row_f, variable=null_v, width=6)
        chk.pack(side=LEFT, padx=(12, 0))

        self.column_rows.append({"frame": row_f, "name": name_v, "sql_type": type_v, "nullable": null_v})

    def _remove_last_row(self):
        if len(self.column_rows) <= 1:
            return
        row = self.column_rows.pop()
        row["frame"].destroy()

    def _build_sql(self) -> str:
        schema = self.table_schema.get().strip() or "dbo"
        tbl = self.table_name.get().strip()
        if not tbl:
            raise ValueError("Укажите имя таблицы.")

        parts: list[str] = []
        if self.add_identity.get():
            parts.append(f"{bracket_ident('Id')} BIGINT NOT NULL IDENTITY(1,1) PRIMARY KEY")

        for row in self.column_rows:
            cname = row["name"].get().strip()
            if not cname:
                continue
            col_type = row["sql_type"].get().strip()
            if not col_type:
                raise ValueError(f"Укажите тип для столбца «{cname}».")
            null_sql = "NULL" if row["nullable"].get() else "NOT NULL"
            parts.append(f"{bracket_ident(cname)} {col_type} {null_sql}")

        if not parts:
            raise ValueError("Добавьте хотя бы один столбец или включите ключ Id.")

        inner = ",\n    ".join(parts)
        return f"CREATE TABLE {bracket_ident(schema)}.{bracket_ident(tbl)} (\n    {inner}\n);"

    def _apply(self):
        db = self.app.database.get().strip()
        if not db:
            Messagebox.show_warning("Сначала выберите базу данных и подключитесь.", "База не выбрана")
            return

        try:
            sql = self._build_sql()
        except ValueError as e:
            Messagebox.show_warning(str(e), "Проверьте данные")
            return

        try:
            conn_str = self.app._build_pyodbc_conn_str(database=db)
            self.app.log("Выполнение CREATE TABLE...")
            self.app.log(sql)
            with pyodbc.connect(conn_str, timeout=30) as conn:
                conn.cursor().execute(sql)
                conn.commit()
        except Exception as exc:
            Messagebox.show_error(str(exc), "Ошибка создания таблицы")
            self.app.log(f"Ошибка: {exc}")
            return

        schema = self.table_schema.get().strip() or "dbo"
        tbl = self.table_name.get().strip()
        self.app.log(f"Таблица создана: {schema}.{tbl}")
        self.app.schema.set(schema)
        self.app.table.set(f"{schema}.{tbl}")

        Messagebox.show_info("Таблица успешно создана.", "Готово")
        self.app.refresh_tables()
        self._on_close()


class UploadApp(ttk.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title(f"{APP_NAME} — Excel → SQL Server")
        self.geometry("920x740")
        self.minsize(780, 620)

        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.server = tk.StringVar(value="D-HDP-SBOX-111")
        self.driver = tk.StringVar(value="ODBC Driver 17 for SQL Server")
        self.trusted = tk.BooleanVar(value=True)
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.database = tk.StringVar()
        self.table = tk.StringVar()
        self.schema = tk.StringVar(value="dbo")
        self.if_exists = tk.StringVar(value="replace")

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._build_ui()
        self.after(800, self._schedule_update_check)

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        outer = ttk.Frame(self)
        outer.grid(row=0, column=0, sticky=NSEW)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        self._update_banner_host = ttk.Frame(outer)
        self._update_banner_host.columnconfigure(0, weight=1)

        root = ttk.Frame(outer, padding=18)
        root.grid(row=1, column=0, sticky=NSEW)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(5, weight=1)

        foot = ttk.Frame(outer, padding=(18, 10))
        foot.grid(row=2, column=0, sticky=EW)
        foot.columnconfigure(0, weight=1)

        ttk.Label(foot, text=APP_NAME, font=("Segoe UI", 10)).grid(row=0, column=0, sticky=W)
        ttk.Label(foot, text=f"v{APP_VERSION}", bootstyle="secondary").grid(row=0, column=1, sticky=E)

        title_f = ttk.Frame(root)
        title_f.grid(row=0, column=0, sticky=EW, pady=(0, 12))
        ttk.Label(title_f, text="Загрузка Excel в SQL Server", font=("Segoe UI", 16, "bold")).pack(anchor=W)
        ttk.Label(
            title_f,
            text="Подключитесь к серверу, выберите базу и таблицу, при необходимости создайте таблицу из интерфейса.",
            bootstyle="secondary",
        ).pack(anchor=W, pady=(4, 0))

        pad = {"padx": 0, "pady": 8}

        conn_frame = ttk.Labelframe(root, text=" Подключение ", padding=12)
        conn_frame.grid(row=1, column=0, sticky=EW, **pad)
        conn_frame.columnconfigure(1, weight=1)
        conn_frame.columnconfigure(3, weight=1)

        ttk.Label(conn_frame, text="Сервер").grid(row=0, column=0, sticky=W, padx=(0, 8))
        ttk.Entry(conn_frame, textvariable=self.server).grid(row=0, column=1, sticky=EW)

        ttk.Label(conn_frame, text="ODBC драйвер").grid(row=0, column=2, sticky=W, padx=(16, 8))
        ttk.Entry(conn_frame, textvariable=self.driver).grid(row=0, column=3, sticky=EW)

        ttk.Checkbutton(
            conn_frame,
            text="Windows Authentication (Trusted)",
            variable=self.trusted,
            command=self._toggle_auth_fields,
        ).grid(row=1, column=0, columnspan=2, sticky=W, pady=(10, 0))

        ttk.Label(conn_frame, text="Логин").grid(row=1, column=2, sticky=W, padx=(16, 8), pady=(10, 0))
        self.username_entry = ttk.Entry(conn_frame, textvariable=self.username, width=28, state="disabled")
        self.username_entry.grid(row=1, column=3, sticky=EW, pady=(10, 0))

        ttk.Label(conn_frame, text="Пароль").grid(row=2, column=2, sticky=W, padx=(16, 8), pady=(6, 0))
        self.password_entry = ttk.Entry(conn_frame, textvariable=self.password, show="•", state="disabled")
        self.password_entry.grid(row=2, column=3, sticky=EW, pady=(6, 0))

        ttk.Button(conn_frame, text="Подключиться", command=self.connect, bootstyle="primary").grid(
            row=2, column=0, sticky=W, pady=(10, 0)
        )

        db_frame = ttk.Labelframe(root, text=" База данных и таблица ", padding=12)
        db_frame.grid(row=2, column=0, sticky=EW, **pad)
        db_frame.columnconfigure(1, weight=1)
        db_frame.columnconfigure(3, weight=1)

        ttk.Label(db_frame, text="База").grid(row=0, column=0, sticky=W, padx=(0, 8))
        self.db_combo = ttk.Combobox(db_frame, textvariable=self.database, state="readonly")
        self.db_combo.grid(row=0, column=1, sticky=EW)
        self.db_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_tables())

        ttk.Label(db_frame, text="Схема по умолчанию").grid(row=0, column=2, sticky=W, padx=(16, 8))
        ttk.Entry(db_frame, textvariable=self.schema, width=14).grid(row=0, column=3, sticky=W)

        ttk.Label(db_frame, text="Таблица").grid(row=1, column=0, sticky=W, pady=(10, 0))
        self.table_combo = ttk.Combobox(db_frame, textvariable=self.table)
        self.table_combo.grid(row=1, column=1, sticky=EW, pady=(10, 0))

        bf = ttk.Frame(db_frame)
        bf.grid(row=1, column=2, columnspan=2, sticky=EW, padx=(16, 0), pady=(10, 0))
        ttk.Button(bf, text="Обновить список", command=self.refresh_tables, bootstyle="secondary-outline").pack(
            side=LEFT, padx=(0, 8)
        )
        ttk.Button(bf, text="Создать таблицу…", command=self._open_create_table, bootstyle="info-outline").pack(
            side=LEFT
        )

        excel_frame = ttk.Labelframe(root, text=" Файл Excel ", padding=12)
        excel_frame.grid(row=3, column=0, sticky=EW, **pad)
        excel_frame.columnconfigure(1, weight=1)

        ttk.Label(excel_frame, text="Файл").grid(row=0, column=0, sticky=W, padx=(0, 8))
        ttk.Entry(excel_frame, textvariable=self.file_path).grid(row=0, column=1, sticky=EW)
        ttk.Button(excel_frame, text="Обзор…", command=self.browse_file, bootstyle="secondary").grid(
            row=0, column=2, padx=(8, 0)
        )

        ttk.Label(excel_frame, text="Лист").grid(row=0, column=3, sticky=W, padx=(16, 8))
        self.sheet_combo = ttk.Combobox(excel_frame, textvariable=self.sheet_name, state="readonly", width=26)
        self.sheet_combo.grid(row=0, column=4, sticky=W)

        opt_row = ttk.Frame(root)
        opt_row.grid(row=4, column=0, sticky=EW, **pad)

        ttk.Label(opt_row, text="Режим загрузки").pack(side=LEFT)
        ttk.Combobox(
            opt_row,
            textvariable=self.if_exists,
            values=["replace", "append"],
            state="readonly",
            width=14,
        ).pack(side=LEFT, padx=(12, 0))

        ttk.Button(
            opt_row,
            text="Загрузить в SQL Server",
            command=self.upload,
            bootstyle="success",
        ).pack(side=RIGHT)

        log_frame = ttk.Labelframe(root, text=" Журнал ", padding=8)
        log_frame.grid(row=5, column=0, sticky=NSEW, **pad)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=14, wrap="word", relief="flat", padx=8, pady=8)
        self.log_text.grid(row=0, column=0, sticky=NSEW)

        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        try:
            bg = self.style.colors.get("bg")
            fg = self.style.colors.get("fg")
            self.log_text.configure(background=bg, foreground=fg, insertbackground=fg)
        except Exception:
            pass

    def _schedule_update_check(self):
        repo = _resolve_github_repo()
        if not repo:
            return

        def worker():
            data = fetch_latest_github_release(repo)
            if not data:
                return
            tag = (data.get("tag_name") or "").strip()
            if not tag or not _version_is_newer(tag, APP_VERSION):
                return
            url = pick_release_download_url(data)
            label = tag.lstrip("vV")
            self.after(0, lambda v=label, u=url: self._show_update_banner(v, u))

        threading.Thread(target=worker, daemon=True).start()

    def _show_update_banner(self, remote_ver: str, download_url: str):
        for w in self._update_banner_host.winfo_children():
            w.destroy()

        self._update_banner_host.grid(row=0, column=0, sticky=EW, padx=18, pady=(14, 0))

        wrap = ttk.Labelframe(self._update_banner_host, text=" Обновление ", padding=(12, 10))
        wrap.pack(fill=EW)
        wrap.columnconfigure(0, weight=1)

        msg = f"Доступна новая версия {remote_ver}. Установлено: v{APP_VERSION}. Скачайте сборку со страницы релиза GitHub."
        ttk.Label(wrap, text=msg, wraplength=720, anchor=W).grid(row=0, column=0, sticky=EW)

        btns = ttk.Frame(wrap)
        btns.grid(row=1, column=0, sticky=W, pady=(10, 0))

        def open_dl():
            if download_url:
                webbrowser.open(download_url)

        ttk.Button(btns, text="Скачать обновление", command=open_dl, bootstyle="info").pack(side=LEFT)
        ttk.Button(btns, text="Скрыть", command=self._hide_update_banner, bootstyle="link").pack(side=LEFT, padx=(12, 0))

    def _hide_update_banner(self):
        self._update_banner_host.grid_remove()
        for w in self._update_banner_host.winfo_children():
            w.destroy()

    def _toggle_auth_fields(self):
        state = "disabled" if self.trusted.get() else "normal"
        self.username_entry.configure(state=state)
        self.password_entry.configure(state=state)

    def _open_create_table(self):
        db = self.database.get().strip()
        if not db:
            Messagebox.show_warning("Выберите базу данных на вкладке подключения.", "База не выбрана")
            return
        CreateTableDialog(self)

    def log(self, msg: str):
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        self.file_path.set(path)
        self._load_sheets(path)

    def _load_sheets(self, path):
        try:
            xls = pd.ExcelFile(path)
        except Exception as exc:
            Messagebox.show_error(f"Не удалось прочитать файл: {exc}", "Ошибка")
            return
        sheets = xls.sheet_names
        self.sheet_combo.configure(values=sheets)
        if sheets:
            self.sheet_name.set(sheets[0])
        self.log(f"Найдено листов: {len(sheets)}")

    def _build_pyodbc_conn_str(self, database=None):
        server = self.server.get().strip()
        driver = self.driver.get().strip()
        if not server:
            raise ValueError("Укажите сервер.")
        if not driver:
            raise ValueError("Укажите ODBC драйвер.")

        parts = [f"DRIVER={{{driver}}}", f"SERVER={server}"]
        if database:
            parts.append(f"DATABASE={database}")

        if self.trusted.get():
            parts.append("Trusted_Connection=yes")
        else:
            user = self.username.get().strip()
            pwd = self.password.get()
            if not user or not pwd:
                raise ValueError("Укажите логин и пароль.")
            parts.append("Trusted_Connection=no")
            parts.append(f"UID={user}")
            parts.append(f"PWD={pwd}")

        return ";".join(parts) + ";"

    def connect(self):
        try:
            conn_str = self._build_pyodbc_conn_str(database="master")
            with pyodbc.connect(conn_str, timeout=5) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sys.databases ORDER BY name")
                dbs = [row[0] for row in cursor.fetchall()]
            self.db_combo.configure(values=dbs)
            if dbs and not self.database.get():
                self.database.set(dbs[0])
            self.log(f"Подключено. Баз данных: {len(dbs)}")
            self.refresh_tables()
        except Exception as exc:
            Messagebox.show_error(str(exc), "Ошибка подключения")

    def refresh_tables(self):
        db = self.database.get().strip()
        if not db:
            return
        try:
            conn_str = self._build_pyodbc_conn_str(database=db)
            with pyodbc.connect(conn_str, timeout=10) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    """
                    SELECT TABLE_SCHEMA, TABLE_NAME
                    FROM INFORMATION_SCHEMA.TABLES
                    WHERE TABLE_TYPE = 'BASE TABLE'
                    ORDER BY TABLE_SCHEMA, TABLE_NAME
                    """
                )
                tables = [f"{row[0]}.{row[1]}" for row in cursor.fetchall()]
            self.table_combo.configure(values=tables)
            if tables and not self.table.get():
                self.table.set(tables[0])
            self.log(f"Таблиц в «{db}»: {len(tables)}")
        except Exception as exc:
            Messagebox.show_error(f"Не удалось получить таблицы: {exc}", "Ошибка")

    def _parse_table(self):
        raw = self.table.get().strip()
        if not raw:
            raise ValueError("Укажите таблицу.")
        if "." in raw:
            schema, table = raw.split(".", 1)
            return schema or "dbo", table
        return (self.schema.get().strip() or "dbo"), raw

    def upload(self):
        path = self.file_path.get().strip()
        sheet = self.sheet_name.get().strip()
        db = self.database.get().strip()
        mode = self.if_exists.get().strip()

        if not path:
            Messagebox.show_warning("Выберите файл Excel.", "Файл")
            return
        if not sheet:
            Messagebox.show_warning("Выберите лист Excel.", "Лист")
            return
        if not db:
            Messagebox.show_warning("Выберите базу данных.", "База")
            return

        try:
            schema, table = self._parse_table()
            conn_str = self._build_pyodbc_conn_str(database=db)
            url = "mssql+pyodbc:///?odbc_connect=" + quote_plus(conn_str)
            engine = create_engine(url, fast_executemany=True)

            self.log("Чтение Excel...")
            df = pd.read_excel(path, sheet_name=sheet)
            self.log(f"Строк: {len(df)}. Колонок: {len(df.columns)}")

            self.log(f"Загрузка в {schema}.{table} ({mode})...")
            df.to_sql(table, engine, if_exists=mode, index=False, schema=schema)

            self.log("Готово.")
            Messagebox.show_info("Данные успешно загружены.", "Успех")
        except Exception as exc:
            Messagebox.show_error(str(exc), "Ошибка загрузки")


def main():
    app = UploadApp()
    app.mainloop()


if __name__ == "__main__":
    main()
