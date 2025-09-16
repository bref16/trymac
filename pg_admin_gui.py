import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import Dict, Any, Optional, List
import datetime as dt

from sqlalchemy import create_engine, text, Table, MetaData, inspect, select
from sqlalchemy.engine import Engine
from sqlalchemy import Integer, Float, Boolean, Text, LargeBinary, Date, DateTime, Time, Numeric

# Optional (ttkbootstrap) – light theme only
try:
    from ttkbootstrap import Style
except Exception:
    Style = None  # fallback if not installed

APP_TITLE = "PostgreSQL Администратор (Десктоп)"
DEFAULTS = {"host": "127.0.0.1", "port": "5432", "database": "postgres", "user": "postgres", "password": ""}

# ----------------- Utility -----------------
def coerce_value(col, raw: str):
    if raw == "" or raw is None:
        return None
    t = col.type
    try:
        if isinstance(t, Boolean): return raw.lower() in ("true","t","1","yes","y","on","да")
        if isinstance(t, (Integer,)): return int(raw)
        if isinstance(t, (Float, Numeric)): return float(raw)
        if isinstance(t, (DateTime,)): return dt.datetime.fromisoformat(raw)
        if isinstance(t, (Date,)): return dt.date.fromisoformat(raw)
        if isinstance(t, (Time,)): return dt.time.fromisoformat(raw)
        return raw
    except Exception:
        return raw

class EditDialog(tk.Toplevel):
    """Диалог редактирования одной строки (все поля, PK только для чтения)."""
    def __init__(self, master, engine: Engine, table: Table, pk_col: str, pk_value: Any, row_data: Dict[str, Any]):
        super().__init__(master)
        self.title(f"Редактирование — {table.name}")
        self.engine = engine
        self.table = table
        self.pk_col = pk_col
        self.pk_value = pk_value
        self.inputs: Dict[str, tk.Entry] = {}

        self.columnconfigure(1, weight=1)

        r = 0
        for col in table.c:
            ttk.Label(self, text=col.name).grid(row=r, column=0, sticky="w", padx=6, pady=4)
            ent = ttk.Entry(self)
            val = row_data.get(col.name)
            ent.insert(0, "" if val is None else str(val))
            if col.name == pk_col:
                ent.configure(state="disabled")
            ent.grid(row=r, column=1, sticky="we", padx=6, pady=4)
            self.inputs[col.name] = ent
            r += 1

        btnbar = ttk.Frame(self)
        btnbar.grid(row=r, column=0, columnspan=2, sticky="e", padx=6, pady=8)
        ttk.Button(btnbar, text="Отмена", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btnbar, text="Сохранить", command=self.on_save).pack(side="right", padx=4)

        self.transient(master)
        self.grab_set()
        self.focus_set()

    def on_save(self):
        payload = {}
        for name, ent in self.inputs.items():
            if name == self.pk_col:
                continue
            col = self.table.c[name]
            payload[name] = coerce_value(col, ent.get())
        try:
            with self.engine.begin() as conn:
                conn.execute(self.table.update().where(self.table.c[self.pk_col] == self.pk_value).values(**payload))
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось сохранить изменения:\n{e}")

class BulkUpdateDialog(tk.Toplevel):
    """Массовое изменение значения одного столбца для выбранных строк."""
    def __init__(self, master, engine: Engine, table: Table, pk_col: str, pk_values: List[Any]):
        super().__init__(master)
        self.title(f"Массовое изменение — {table.name}")
        self.engine = engine
        self.table = table
        self.pk_col = pk_col
        self.pk_values = pk_values

        ttk.Label(self, text=f"Выбрано строк: {len(pk_values)}").grid(row=0, column=0, columnspan=2, sticky="w", padx=6, pady=4)
        ttk.Label(self, text="Столбец").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(self, text="Новое значение").grid(row=2, column=0, sticky="w", padx=6, pady=4)

        self.col_var = tk.StringVar()
        self.val_entry = ttk.Entry(self)

        cols = [c.name for c in table.c if c.name != pk_col]
        self.col_box = ttk.Combobox(self, values=cols, textvariable=self.col_var, state="readonly")
        self.col_box.grid(row=1, column=1, sticky="we", padx=6, pady=4)
        self.val_entry.grid(row=2, column=1, sticky="we", padx=6, pady=4)

        btns = ttk.Frame(self)
        btns.grid(row=3, column=0, columnspan=2, sticky="e", padx=6, pady=8)
        ttk.Button(btns, text="Отмена", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btns, text="Применить", command=self.on_apply).pack(side="right", padx=4)

        self.columnconfigure(1, weight=1)
        self.transient(master); self.grab_set(); self.focus_set()

    def on_apply(self):
        colname = self.col_var.get()
        if not colname:
            messagebox.showwarning(APP_TITLE, "Выберите столбец для изменения.")
            return
        col = self.table.c[colname]
        new_value = coerce_value(col, self.val_entry.get())
        try:
            with self.engine.begin() as conn:
                conn.execute(
                    self.table.update().where(self.table.c[self.pk_col].in_(self.pk_values)).values({colname: new_value})
                )
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось применить массовое изменение:\n{e}")

class AddDialog(tk.Toplevel):
    """Диалог добавления новой строки."""
    def __init__(self, master, engine: Engine, table: Table, pk_col: Optional[str] = None):
        super().__init__(master)
        self.title(f"Добавление — {table.name}")
        self.engine = engine
        self.table = table
        self.pk_col = pk_col
        self.inputs: Dict[str, tk.Entry] = {}

        self.columnconfigure(1, weight=1)

        r = 0
        for col in table.c:
            ttk.Label(self, text=col.name).grid(row=r, column=0, sticky="w", padx=6, pady=4)
            ent = ttk.Entry(self)
            ent.grid(row=r, column=1, sticky="we", padx=6, pady=4)
            self.inputs[col.name] = ent
            r += 1

        hint = ttk.Label(self, text="Оставьте поле пустым, чтобы использовать NULL/DEFAULT.", foreground="gray")
        hint.grid(row=r, column=0, columnspan=2, sticky="w", padx=6, pady=(0,6))
        r += 1

        btnbar = ttk.Frame(self)
        btnbar.grid(row=r, column=0, columnspan=2, sticky="e", padx=6, pady=8)
        ttk.Button(btnbar, text="Отмена", command=self.destroy).pack(side="right", padx=4)
        ttk.Button(btnbar, text="Добавить", command=self.on_insert).pack(side="right", padx=4)

        self.transient(master); self.grab_set(); self.focus_set()

    def on_insert(self):
        payload: Dict[str, Any] = {}
        for name, ent in self.inputs.items():
            raw = ent.get()
            if raw == "":   # пропускаем пустые, даём шансу DEFAULT/NULL
                continue
            payload[name] = coerce_value(self.table.c[name], raw)

        try:
            with self.engine.begin() as conn:
                if self.pk_col:
                    # Пытаемся вернуть PK (PostgreSQL)
                    try:
                        result = conn.execute(self.table.insert().values(**payload).returning(self.table.c[self.pk_col]))
                        _ = result.scalar()
                    except Exception:
                        conn.execute(self.table.insert().values(**payload))
                else:
                    conn.execute(self.table.insert().values(**payload))
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось добавить запись:\n{e}")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1200x780")
        self.minsize(980, 620)

        # ---- Light theme / default ttk (dark disabled) ----
        self.style = None
        if Style is not None:
            try:
                # use a light ttkbootstrap theme if available
                self.style = Style(master=self, theme="flatly")
            except Exception:
                self.style = Style(theme="flatly")
        # else: default ttk theme

        # Configure some styling
        style = ttk.Style(self)
        style.configure("TFrame", padding=0)
        style.configure("Header.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Subheader.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview", rowheight=28, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

        self.engine: Optional[Engine] = None
        self.inspector = None
        self.row_cache: Dict[str, Dict[str, Any]] = {}
        self.current_pk: Optional[str] = None

        self._build_connection_bar()
        self._build_main_panes()

    # ---------- Connection ----------
    def _build_connection_bar(self):
        frame = ttk.Frame(self, padding=10)
        frame.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(frame, text=APP_TITLE, style="Header.TLabel").grid(row=0, column=0, sticky="w", padx=4, pady=(0,4))

        self.host_var = tk.StringVar(value=DEFAULTS["host"])
        self.port_var = tk.StringVar(value=DEFAULTS["port"])
        self.db_var   = tk.StringVar(value=DEFAULTS["database"])
        self.user_var = tk.StringVar(value=DEFAULTS["user"])
        self.pw_var   = tk.StringVar(value=DEFAULTS["password"])

        row = 1
        labels = [("Адрес", self.host_var, 18),
                  ("Порт", self.port_var, 8),
                  ("База данных", self.db_var, 18),
                  ("Пользователь", self.user_var, 18),
                  ("Пароль", self.pw_var, 18)]
        for i,(label,var,width) in enumerate(labels):
            ttk.Label(frame, text=label).grid(row=row, column=2*i, sticky="w", padx=4)
            e = ttk.Entry(frame, textvariable=var, width=width, show="•" if label=="Пароль" else "")
            e.grid(row=row+1, column=2*i, sticky="w", padx=4)

        ttk.Button(frame, text="Подключиться", style="Accent.TButton", command=self.on_connect).grid(row=row+1, column=10, padx=8)
        ttk.Button(frame, text="Отключиться", command=self.on_disconnect).grid(row=row+1, column=11, padx=4)

        for c in range(12):
            frame.grid_columnconfigure(c, weight=0)
        frame.grid_columnconfigure(12, weight=1)

    def on_connect(self):
        if self.engine is not None:
            self.on_disconnect()
        url = f"postgresql+psycopg://{self.user_var.get()}:{self.pw_var.get()}@{self.host_var.get()}:{self.port_var.get()}/{self.db_var.get()}"
        try:
            self.engine = create_engine(url, pool_pre_ping=True)
            with self.engine.connect() as conn:
                conn.execute(text("SELECT 1"))
            self.inspector = inspect(self.engine)
            self.load_tables()
            messagebox.showinfo(APP_TITLE, "Подключение успешно.")
        except Exception as e:
            self.engine = None
            messagebox.showerror(APP_TITLE, f"Не удалось подключиться:\n{e}")

    def on_disconnect(self):
        self.tables_list.delete(0, tk.END)
        self.clear_tree(self.tree)
        self.engine = None
        self.inspector = None
        self.row_cache.clear()
        self.current_pk = None
        # disable action buttons
        self.btn_edit["state"] = "disabled"
        self.btn_bulk["state"] = "disabled"
        self.btn_del["state"]  = "disabled"
        self.btn_add["state"]  = "disabled"

    # ---------- Main panes ----------
    def _build_main_panes(self):
        container = ttk.Frame(self, padding=(8, 0, 8, 8))
        container.pack(fill=tk.BOTH, expand=True)

        self.pw = ttk.Panedwindow(container, orient=tk.HORIZONTAL)
        self.pw.pack(fill=tk.BOTH, expand=True)

        # Left: tables
        left = ttk.Frame(self.pw, padding=(6,6))
        ttk.Label(left, text="Таблицы", style="Subheader.TLabel").pack(anchor="w", pady=(0,4))
        self.tables_list = tk.Listbox(left, exportselection=False)
        self.tables_list.bind("<<ListboxSelect>>", self.on_table_select)
        sby = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.tables_list.yview)
        self.tables_list.configure(yscrollcommand=sby.set)
        self.tables_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sby.pack(side=tk.RIGHT, fill=tk.Y)
        self.pw.add(left, weight=1)

        # Right: rows & actions
        right = ttk.Frame(self.pw, padding=(6,6))
        topbar = ttk.Frame(right)
        topbar.pack(fill=tk.X)

        ttk.Label(topbar, text="Строки", style="Subheader.TLabel").pack(side=tk.LEFT)
        ttk.Label(topbar, text="Лимит").pack(side=tk.LEFT, padx=(16,4))
        self.limit_var = tk.IntVar(value=100)
        ttk.Spinbox(topbar, from_=1, to=5000, textvariable=self.limit_var, width=6).pack(side=tk.LEFT)

        ttk.Label(topbar, text="Смещение").pack(side=tk.LEFT, padx=(12,4))
        self.offset_var = tk.IntVar(value=0)
        ttk.Spinbox(topbar, from_=0, to=1_000_000_000, textvariable=self.offset_var, width=8, increment=100).pack(side=tk.LEFT)
        ttk.Button(topbar, text="Обновить", command=self.refresh_rows).pack(side=tk.LEFT, padx=8)

        # Row height control
        ttk.Label(topbar, text="Высота строки").pack(side=tk.LEFT, padx=(12,4))
        self.rowheight_var = tk.IntVar(value=28)
        self.rowheight_scale = ttk.Scale(topbar, from_=18, to=60, orient="horizontal",
                                         command=self.on_rowheight_change, value=self.rowheight_var.get())
        self.rowheight_scale.pack(side=tk.LEFT, padx=(0,8))

        # Action buttons
        self.btn_add  = ttk.Button(topbar, text="Добавить…", command=self.on_add_row, state="disabled")
        self.btn_edit = ttk.Button(topbar, text="Редактировать…", command=self.on_edit_selected, state="disabled")
        self.btn_bulk = ttk.Button(topbar, text="Массовое изменение…", command=self.on_bulk_update, state="disabled")
        self.btn_del  = ttk.Button(topbar, text="Удалить выбранные", command=self.on_delete_selected, state="disabled")
        for b in (self.btn_del, self.btn_bulk, self.btn_edit, self.btn_add):
            b.pack(side=tk.RIGHT, padx=4)

        # Tree with multi-select
        mid = ttk.Frame(right)
        mid.pack(fill=tk.BOTH, expand=True, pady=6)
        self.tree = ttk.Treeview(mid, columns=(), show="headings", selectmode="extended")
        sby2 = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=self.tree.yview)
        sbx2 = ttk.Scrollbar(mid, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=sby2.set, xscrollcommand=sbx2.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sby2.pack(side=tk.RIGHT, fill=tk.Y)
        sbx2.pack(side=tk.BOTTOM, fill=tk.X)

        # Events
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Button-3>", self.on_right_click)

        self.pw.add(right, weight=3)

        # Context menu
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Добавить…", command=self.on_add_row)
        self.menu.add_separator()
        self.menu.add_command(label="Редактировать…", command=self.on_edit_selected)
        self.menu.add_command(label="Массовое изменение…", command=self.on_bulk_update)
        self.menu.add_separator()
        self.menu.add_command(label="Удалить выбранные", command=self.on_delete_selected)

    # ---------- Helpers ----------
    def on_rowheight_change(self, value):
        try:
            h = int(float(value))
        except Exception:
            h = int(self.rowheight_var.get())
        style = ttk.Style(self)
        style.configure("Treeview", rowheight=h)

    def clear_tree(self, tree: ttk.Treeview):
        for i in tree.get_children():
            tree.delete(i)
        tree["columns"] = ()

    def on_right_click(self, event):
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def on_tree_select(self, event=None):
        sel = self.tree.selection()
        has_sel = len(sel) > 0
        self.btn_edit["state"] = ("normal" if has_sel else "disabled")
        self.btn_bulk["state"] = ("normal" if len(sel) > 1 else "disabled")
        self.btn_del["state"]  = ("normal" if has_sel else "disabled")

    def on_tree_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        row = self.row_cache.get(iid, {})
        if not self.current_pk:
            messagebox.showinfo(APP_TITLE, "Для редактирования требуется первичный ключ из одного столбца.")
            return
        pk_val = row.get(self.current_pk)
        if pk_val is None:
            messagebox.showwarning(APP_TITLE, "Не удалось определить значение первичного ключа для строки.")
            return
        table = self._get_current_table()
        dlg = EditDialog(self, self.engine, table, self.current_pk, pk_val, row)
        self.wait_window(dlg)
        self.refresh_rows()

    def on_add_row(self):
        table = self._get_current_table()
        if table is None:
            return
        dlg = AddDialog(self, self.engine, table, pk_col=self.current_pk)
        self.wait_window(dlg)
        self.refresh_rows()

    def on_edit_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        if len(sel) > 1:
            messagebox.showinfo(APP_TITLE, "Выбрано несколько строк. Используйте «Массовое изменение…».")
            return
        self.on_tree_double_click(None)

    def on_bulk_update(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not self.current_pk:
            messagebox.showinfo(APP_TITLE, "Массовое изменение требует первичного ключа из одного столбца.")
            return
        pk_vals = []
        for iid in sel:
            row = self.row_cache.get(iid, {})
            if self.current_pk in row:
                pk_vals.append(row[self.current_pk])
        if not pk_vals:
            messagebox.showwarning(APP_TITLE, "Не удалось собрать значения первичного ключа.")
            return
        table = self._get_current_table()
        dlg = BulkUpdateDialog(self, self.engine, table, self.current_pk, pk_vals)
        self.wait_window(dlg)
        self.refresh_rows()

    def on_delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        if not self.current_pk:
            messagebox.showinfo(APP_TITLE, "Удаление требует первичного ключа из одного столбца.")
            return
        if not messagebox.askyesno(APP_TITLE, f"Удалить выбранные строки ({len(sel)})? Это действие необратимо."):
            return
        table = self._get_current_table()
        pk_vals = []
        for iid in sel:
            row = self.row_cache.get(iid, {})
            if self.current_pk in row:
                pk_vals.append(row[self.current_pk])
        try:
            with self.engine.begin() as conn:
                conn.execute(table.delete().where(table.c[self.current_pk].in_(pk_vals)))
            self.refresh_rows()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось удалить строки:\n{e}")

    # ---------- Tables & Rows ----------
    def load_tables(self):
        self.tables_list.delete(0, tk.END)
        try:
            names = sorted(self.inspector.get_table_names())
            for n in names:
                self.tables_list.insert(tk.END, n)
            if names:
                self.tables_list.selection_clear(0, tk.END)
                self.tables_list.selection_set(0)
                self.on_table_select()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось получить список таблиц:\n{e}")

    def _get_current_table(self) -> Optional[Table]:
        if self.engine is None:
            return None
        sel = self.tables_list.curselection()
        if not sel:
            return None
        name = self.tables_list.get(sel[0])
        md = MetaData()
        try:
            tbl = Table(name, md, autoload_with=self.engine)
            return tbl
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось загрузить таблицу:\n{name}\n{e}")
            return None

    def on_table_select(self, event=None):
        self.refresh_rows()

    def _detect_single_pk(self, table: Table) -> Optional[str]:
        info = self.inspector.get_pk_constraint(table.name)
        cols = info.get("constrained_columns") or []
        return cols[0] if len(cols) == 1 else None

    def refresh_rows(self):
        table = self._get_current_table()
        if table is None:
            self.btn_add["state"] = "disabled"
            return
        self.current_pk = self._detect_single_pk(table)

        limit = int(self.limit_var.get() or 100)
        offset = int(self.offset_var.get() or 0)
        stmt = select(table).limit(limit).offset(offset)
        try:
            with self.engine.connect() as conn:
                rows = [dict(r._mapping) for r in conn.execute(stmt)]
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось выполнить запрос:\n{e}")
            return
        self.populate_tree(table, rows)
        self.btn_add["state"] = "normal"

    def populate_tree(self, table: Table, rows: List[Dict[str, Any]]):
        self.clear_tree(self.tree)
        self.row_cache.clear()

        cols = [c.name for c in table.c]
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by(c, False))
            self.tree.column(col, width=180, anchor="w")

        # light zebra stripes
        self.tree.tag_configure("oddrow", background="#f7f9fb")
        self.tree.tag_configure("evenrow", background="#ffffff")

        for idx, row in enumerate(rows):
            values = ["" if row.get(c) is None else str(row.get(c)) for c in cols]
            tag = "oddrow" if idx % 2 else "evenrow"
            iid = self.tree.insert("", tk.END, values=values, tags=(tag,))
            self.row_cache[iid] = row

        self.on_tree_select()

    def sort_by(self, col: str, descending: bool):
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children("")]
        try:
            data.sort(key=lambda t: float(t[0]) if t[0] not in ("", None) else float("-inf"), reverse=descending)
        except Exception:
            data.sort(key=lambda t: t[0], reverse=descending)
        for idx, (_, child) in enumerate(data):
            self.tree.move(child, "", idx)
        self.tree.heading(col, command=lambda: self.sort_by(col, not descending))

if __name__ == "__main__":
    app = App()
    app.mainloop()
