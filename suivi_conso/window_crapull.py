import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import csv
import re
import glob
import threading
import traceback
import time
import unicodedata
import calendar
from collections import OrderedDict
from datetime import datetime

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

from .constants import *
from .theme import *
from .utils import *

class CrapullWindow(tk.Toplevel):
    """Fenêtre de visualisation et filtrage des données Crapull (PDC JH + EUR)."""

    def __init__(self, parent, path_jh, path_eur):
        super().__init__(parent)
        self.title("Vérification données Crapull")
        self.configure(bg=BG_MAIN)
        self.resizable(True, True)
        self.minsize(1000, 700)

        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w, h = min(1300, sw - 80), min(820, sh - 80)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        self.path_jh  = path_jh
        self.path_eur = path_eur
        self._df_jh   = None
        self._df_eur  = None
        self._filter_vars  = {}
        self._real_col_names = {}

        self._build_ui()
        self.after(100, self._load_data)

    def _load_data(self):
        self._status.set("Chargement des fichiers PDC…")
        self.update_idletasks()
        try:
            self._df_jh  = self._read_pdc(self.path_jh)
            self._df_eur = self._read_pdc(self.path_eur)
            self._status.set(
                f"JH : {len(self._df_jh)} lignes  |  EUR : {len(self._df_eur)} lignes")
            self._populate_filters()
            self._refresh_headings()
            self._apply_filters()
        except Exception as e:
            self._status.set(f"Erreur : {e}")
            messagebox.showerror("Erreur chargement", str(e), parent=self)

    def _read_pdc(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            enc = "utf-8-sig"
            for candidate in ("utf-8-sig", "utf-8", "windows-1252", "latin-1"):
                try:
                    with open(path, "r", encoding=candidate) as f:
                        sample = f.read(8192)
                    enc = candidate
                    break
                except UnicodeDecodeError:
                    continue
            sep = ";" if sample.count(";") >= sample.count(",") else ","
            df = pd.read_csv(path, dtype=str, sep=sep, encoding=enc,
                              encoding_errors="replace", on_bad_lines="skip")
        else:
            df = pd.read_excel(path, dtype=str, header=3)

        df.columns = [c.strip() for c in df.columns]

        def _norm(s):
            return s.lower().replace(" ", "").replace("_", "")

        col_map = {_norm(c): c for c in df.columns}

        rename   = {}
        real_map = {}

        for crapull_col in CRAPULL_COLS:
            key = _norm(crapull_col)
            if crapull_col in df.columns:
                real_map[crapull_col] = crapull_col
            elif key in col_map:
                real_name = col_map[key]
                rename[real_name] = crapull_col
                real_map[crapull_col] = real_name
            else:
                df[crapull_col] = ""
                real_map[crapull_col] = crapull_col

        if rename:
            df = df.rename(columns=rename)

        if not hasattr(self, "_real_col_names"):
            self._real_col_names = real_map
        else:
            for k, v in real_map.items():
                self._real_col_names.setdefault(k, v)

        return df[CRAPULL_COLS].fillna("").astype(str)

    def _build_ui(self):
        hdr = tk.Frame(self, bg=BG_MAIN, pady=8)
        hdr.pack(fill="x", padx=16)
        tk.Label(hdr, text="🔍  VÉRIFICATION DONNÉES CRAPULL",
                 bg=BG_MAIN, fg=ACCENT2, font=FONT_HEAD).pack(side="left")
        tk.Label(hdr, text=os.path.basename(self.path_jh),
                 bg=BG_MAIN, fg=TEXT_SEC, font=FONT_SMALL).pack(side="right")

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        filter_frame = tk.LabelFrame(
            self, text="  Filtres  ", bg=BG_PANEL, fg=ACCENT2,
            font=FONT_SMALL, bd=1, relief="flat",
            highlightbackground=BORDER, highlightthickness=1)
        filter_frame.pack(fill="x", padx=10, pady=(8, 4), ipady=6)

        self._combo_widgets = {}
        for i, col in enumerate(FILTER_COLS):
            col_frame = tk.Frame(filter_frame, bg=BG_PANEL)
            col_frame.grid(row=0, column=i*2, padx=(10, 2), pady=4, sticky="w")
            tk.Label(col_frame, text=col, bg=BG_PANEL,
                     fg=TEXT_SEC, font=FONT_SMALL).pack(anchor="w")
            var = tk.StringVar(value="(Tous)")
            self._filter_vars[col] = var

            style = ttk.Style()
            style.theme_use("default")
            style.configure("Dark.TCombobox",
                            fieldbackground=BG_ENTRY, background=BG_CARD,
                            foreground=TEXT_PRI, selectbackground=ACCENT2,
                            font=FONT_SMALL)

            cb = ttk.Combobox(col_frame, textvariable=var,
                               state="readonly", width=22,
                               style="Dark.TCombobox", font=FONT_SMALL)
            cb.pack()
            cb.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())
            self._combo_widgets[col] = cb

        reset_frame = tk.Frame(filter_frame, bg=BG_PANEL)
        reset_frame.grid(row=0, column=len(FILTER_COLS)*2, padx=10)
        make_button(reset_frame, "↺  Réinitialiser",
                    self._reset_filters, width=16).pack(pady=8)

        self._build_table_section("📋  Consommés JH", "jh")
        self._build_table_section("💶  Consommés EUR (TTNR)", "eur")

        self._status = tk.StringVar(value="Chargement…")
        tk.Label(self, textvariable=self._status, bg=BG_CARD,
                 fg=TEXT_SEC, font=FONT_SMALL, anchor="w",
                 padx=12).pack(fill="x", side="bottom")

    def _build_table_section(self, title, tag):
        section = tk.Frame(self, bg=BG_MAIN)
        section.pack(fill="both", expand=True, padx=10, pady=(4, 0))

        tk.Label(section, text=title, bg=BG_MAIN,
                 fg=ACCENT2, font=FONT_BODY).pack(anchor="w", padx=4)

        count_var = tk.StringVar(value="")
        tk.Label(section, textvariable=count_var, bg=BG_MAIN,
                 fg=TEXT_SEC, font=FONT_SMALL).pack(anchor="w", padx=4)

        tbl_frame = tk.Frame(section, bg=BG_CARD,
                              highlightbackground=BORDER, highlightthickness=1)
        tbl_frame.pack(fill="both", expand=True, pady=(2, 6))

        vsb = tk.Scrollbar(tbl_frame, orient="vertical")
        hsb = tk.Scrollbar(tbl_frame, orient="horizontal")
        vsb.pack(side="right",  fill="y")
        hsb.pack(side="bottom", fill="x")

        tv = ttk.Treeview(tbl_frame,
                           columns=CRAPULL_COLS,
                           show="headings",
                           yscrollcommand=vsb.set,
                           xscrollcommand=hsb.set)
        tv.pack(fill="both", expand=True)
        vsb.config(command=tv.yview)
        hsb.config(command=tv.xview)

        style = ttk.Style()
        style.configure("Dark.Treeview",
                         background=BG_CARD, foreground=TEXT_PRI,
                         fieldbackground=BG_CARD, rowheight=22,
                         font=FONT_SMALL)
        style.configure("Dark.Treeview.Heading",
                         background=BG_ENTRY, foreground=ACCENT2,
                         font=FONT_SMALL_BOLD)
        style.map("Dark.Treeview",
                  background=[("selected", ACCENT2)],
                  foreground=[("selected", "#000000")])
        tv.config(style="Dark.Treeview")

        col_widths = {
            "Intervenant": 180, "Type Intervenant": 210, "Code UT": 80,
            "Year of charge": 100, "Mois": 55, "NbJours": 90,
            "PRJ_CodeProjet": 110, "PRJ_LibelleProjet": 220,
        }
        center_cols = {"Mois", "NbJours", "Code UT", "Year of charge"}
        for col in CRAPULL_COLS:
            real_name = getattr(self, "_real_col_names", {}).get(col, col)
            tv.heading(col, text=real_name,
                        command=lambda c=col, t=tv: self._sort_col(t, c))
            anchor = "center" if col in center_cols else "w"
            tv.column(col, width=col_widths.get(col, 100),
                       anchor=anchor, stretch=True)

        tv.tag_configure("odd",  background=BG_CARD)
        tv.tag_configure("even", background=BG_ENTRY)

        setattr(self, f"_tv_{tag}", tv)
        setattr(self, f"_count_{tag}", count_var)

    def _refresh_headings(self):
        for tag in ("jh", "eur"):
            tv = getattr(self, f"_tv_{tag}", None)
            if tv is None:
                continue
            for col in CRAPULL_COLS:
                real_name = self._real_col_names.get(col, col)
                tv.heading(col, text=real_name)

    def _populate_filters(self):
        if self._df_jh is None:
            return
        combined = pd.concat([self._df_jh, self._df_eur], ignore_index=True)
        for col, cb in self._combo_widgets.items():
            vals = sorted(combined[col].unique().tolist())
            cb["values"] = ["(Tous)"] + vals

    def _reset_filters(self):
        for var in self._filter_vars.values():
            var.set("(Tous)")
        self._apply_filters()

    def _apply_filters(self):
        if self._df_jh is None or self._df_eur is None:
            return

        def _filter(df):
            mask = pd.Series([True] * len(df), index=df.index)
            for col, var in self._filter_vars.items():
                val = var.get()
                if val != "(Tous)" and col in df.columns:
                    mask &= (df[col] == val)
            return df[mask]

        df_jh_f  = _filter(self._df_jh)
        df_eur_f = _filter(self._df_eur)

        self._fill_table(self._tv_jh,  df_jh_f,  is_eur=False)
        self._fill_table(self._tv_eur, df_eur_f, is_eur=True)
        self._count_jh.set(f"{len(df_jh_f)} ligne(s)")
        self._count_eur.set(f"{len(df_eur_f)} ligne(s)")
        self._status.set(
            f"Filtre appliqué — JH : {len(df_jh_f)} lignes  |  EUR : {len(df_eur_f)} lignes")

    def _fill_table(self, tv, df, is_eur=False):
        tv.delete(*tv.get_children())
        for i, (_, row) in enumerate(df.iterrows()):
            tag = "even" if i % 2 == 0 else "odd"
            values = []
            for c in CRAPULL_COLS:
                val = row[c]
                if is_eur and c == "NbJours":
                    try:
                        num = float(str(val).replace(",", ".").strip())
                        formatted = f"{num:,.2f}".replace(",", " ")
                        val = formatted
                    except (ValueError, TypeError):
                        pass
                values.append(val)
            tv.insert("", "end", values=values, tags=(tag,))

    def _sort_col(self, tv, col):
        data = [(tv.set(k, col), k) for k in tv.get_children("")]
        data.sort(reverse=getattr(tv, f"_sort_rev_{col}", False))
        for i, (_, k) in enumerate(data):
            tv.move(k, "", i)
            tv.item(k, tags=("even" if i % 2 == 0 else "odd",))
        setattr(tv, f"_sort_rev_{col}",
                not getattr(tv, f"_sort_rev_{col}", False))


# ─────────────────────────────────────────────
#  FENÊTRE VÉRIFICATION DIANA
# ─────────────────────────────────────────────
DIANA_COLS = [
    "Project Code", "Project Is Alive", "WorkPackage Name", "Task Name",
    "Organization Unit Name", "User Country Name", "User Contract Type",
    "TimeSheet Mode Name", "Status", "Is Published", "Application Cluster Name",
    "User First Name", "User Last Name",
    "Project Name",
    "Timesheet Entry Details Calendar Date",
    "Timesheet Entry Details Value",
]
DIANA_FILTER_COLS = [
    "User First Name", "User Last Name",
    "Project Code", "Project Name",
    "WorkPackage Name", "Task Name",
]
DIANA_COL_WIDTHS = {
    "Project Code": 90, "Project Is Alive": 80, "WorkPackage Name": 160,
    "Task Name": 160, "Organization Unit Name": 180, "User Country Name": 100,
    "User Contract Type": 120, "TimeSheet Mode Name": 130, "Status": 80,
    "Is Published": 80, "Application Cluster Name": 160,
    "User First Name": 100, "User Last Name": 130, "Project Name": 200,
    "Timesheet Entry Details Calendar Date": 130,
    "Timesheet Entry Details Value": 100,
}
