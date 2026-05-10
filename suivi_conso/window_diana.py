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

class DianaWindow(tk.Toplevel):

    def __init__(self, parent, path):
        super().__init__(parent)
        self.title("Vérification données Diana")
        self.configure(bg=BG_MAIN)
        self.resizable(True, True)
        self.minsize(1000, 680)
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w, h = min(1400, sw - 80), min(830, sh - 80)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        self.path         = path
        self._df          = None
        self._real_cols   = {}
        self._filter_vars = {}

        self._build_ui()
        self.after(100, self._load_data)

    def _load_data(self):
        self._status.set("Chargement du fichier Diana…")
        self.update_idletasks()
        try:
            enc = "utf-8-sig"
            for candidate in ("utf-8-sig", "utf-8", "windows-1252", "latin-1"):
                try:
                    with open(self.path, "r", encoding=candidate) as f:
                        f.read(8192)
                    enc = candidate
                    break
                except UnicodeDecodeError:
                    continue

            # Détecter le séparateur (~ pour .txt, ; pour CSV filtré)
            ext = os.path.splitext(self.path)[1].lower()
            sep = SEP_INPUT if ext == ".txt" else SEP_OUTPUT

            with open(self.path, "r", encoding=enc, errors="replace") as f:
                lines = f.readlines()

            if not lines:
                raise ValueError("Fichier vide.")

            raw_headers = [h.strip().strip('"') for h in lines[0].split(sep)]

            def _norm(s):
                return s.lower().replace(" ", "").replace("_", "")

            col_map = {_norm(h): h for h in raw_headers}

            col_indices = {}
            self._real_cols = {}
            for dc in DIANA_COLS:
                key = _norm(dc)
                if dc in raw_headers:
                    col_indices[dc] = raw_headers.index(dc)
                    self._real_cols[dc] = dc
                elif key in col_map:
                    real = col_map[key]
                    col_indices[dc] = raw_headers.index(real)
                    self._real_cols[dc] = real
                else:
                    col_indices[dc] = -1
                    self._real_cols[dc] = dc

            records = []
            for line in lines[1:]:
                line = line.rstrip("\n\r")
                if not line.strip():
                    continue
                fields = [f.strip().strip('"') for f in line.split(sep)]
                while len(fields) < len(raw_headers):
                    fields.append("")
                row = {}
                for dc in DIANA_COLS:
                    idx = col_indices[dc]
                    row[dc] = fields[idx] if idx >= 0 and idx < len(fields) else ""
                records.append(row)

            self._df = pd.DataFrame(records, columns=DIANA_COLS)
            self._status.set(f"{len(self._df)} lignes chargées.")
            self._populate_filters()
            self._refresh_headings()
            self._apply_filters()

        except Exception as e:
            self._status.set(f"Erreur : {e}")
            messagebox.showerror("Erreur chargement", str(e), parent=self)

    def _build_ui(self):
        hdr = tk.Frame(self, bg=BG_MAIN, pady=8)
        hdr.pack(fill="x", padx=16)
        tk.Label(hdr, text="📋  VÉRIFICATION DONNÉES DIANA",
                 bg=BG_MAIN, fg=ACCENT2, font=FONT_HEAD).pack(side="left")
        tk.Label(hdr, text=os.path.basename(self.path),
                 bg=BG_MAIN, fg=TEXT_SEC, font=FONT_SMALL).pack(side="right")
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        filter_frame = tk.LabelFrame(
            self, text="  Filtres  ", bg=BG_PANEL, fg=ACCENT2,
            font=FONT_SMALL, bd=1, relief="flat",
            highlightbackground=BORDER, highlightthickness=1)
        filter_frame.pack(fill="x", padx=10, pady=(8, 4), ipady=6)

        self._combo_widgets = {}
        for i, col in enumerate(DIANA_FILTER_COLS):
            col_frame = tk.Frame(filter_frame, bg=BG_PANEL)
            col_frame.grid(row=0, column=i, padx=(10, 2), pady=4, sticky="w")
            tk.Label(col_frame, text=col, bg=BG_PANEL,
                     fg=TEXT_SEC, font=FONT_SMALL).pack(anchor="w")
            var = tk.StringVar(value="(Tous)")
            self._filter_vars[col] = var
            cb = ttk.Combobox(col_frame, textvariable=var,
                               state="readonly", width=20,
                               font=FONT_SMALL)
            cb.pack()
            cb.bind("<<ComboboxSelected>>", lambda e: self._apply_filters())
            self._combo_widgets[col] = cb

        reset_frame = tk.Frame(filter_frame, bg=BG_PANEL)
        reset_frame.grid(row=0, column=len(DIANA_FILTER_COLS), padx=10)
        make_button(reset_frame, "↺  Réinitialiser",
                    self._reset_filters, width=16).pack(pady=8)

        section = tk.Frame(self, bg=BG_MAIN)
        section.pack(fill="both", expand=True, padx=10, pady=(4, 0))

        self._count_var = tk.StringVar(value="")
        tk.Label(section, textvariable=self._count_var,
                 bg=BG_MAIN, fg=TEXT_SEC, font=FONT_SMALL).pack(anchor="w", padx=4)

        tbl_frame = tk.Frame(section, bg=BG_CARD,
                              highlightbackground=BORDER, highlightthickness=1)
        tbl_frame.pack(fill="both", expand=True, pady=(2, 6))
        vsb = tk.Scrollbar(tbl_frame, orient="vertical")
        hsb = tk.Scrollbar(tbl_frame, orient="horizontal")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self._tv = ttk.Treeview(tbl_frame, columns=DIANA_COLS,
                                 show="headings",
                                 yscrollcommand=vsb.set,
                                 xscrollcommand=hsb.set)
        self._tv.pack(fill="both", expand=True)
        vsb.config(command=self._tv.yview)
        hsb.config(command=self._tv.xview)

        style = ttk.Style()
        style.configure("Diana.Treeview",
                         background=BG_CARD, foreground=TEXT_PRI,
                         fieldbackground=BG_CARD, rowheight=22, font=FONT_SMALL)
        style.configure("Diana.Treeview.Heading",
                         background=BG_ENTRY, foreground=ACCENT2,
                         font=FONT_SMALL_BOLD)
        style.map("Diana.Treeview",
                  background=[("selected", ACCENT2)],
                  foreground=[("selected", "#000000")])
        self._tv.config(style="Diana.Treeview")

        for col in DIANA_COLS:
            self._tv.heading(col, text=col,
                              command=lambda c=col: self._sort_col(c))
            self._tv.column(col, width=DIANA_COL_WIDTHS.get(col, 100),
                             anchor="w", stretch=True)
        self._tv.tag_configure("odd",  background=BG_CARD)
        self._tv.tag_configure("even", background=BG_ENTRY)

        self._status = tk.StringVar(value="Chargement…")
        tk.Label(self, textvariable=self._status, bg=BG_CARD,
                 fg=TEXT_SEC, font=FONT_SMALL, anchor="w",
                 padx=12).pack(fill="x", side="bottom")

    def _refresh_headings(self):
        for col in DIANA_COLS:
            real = self._real_cols.get(col, col)
            self._tv.heading(col, text=real)

    def _populate_filters(self):
        if self._df is None:
            return
        for col, cb in self._combo_widgets.items():
            if col in self._df.columns:
                vals = sorted(self._df[col].dropna().unique().tolist())
                cb["values"] = ["(Tous)"] + vals

    def _reset_filters(self):
        for var in self._filter_vars.values():
            var.set("(Tous)")
        self._apply_filters()

    def _apply_filters(self):
        if self._df is None:
            return
        mask = pd.Series([True] * len(self._df), index=self._df.index)
        for col, var in self._filter_vars.items():
            val = var.get()
            if val != "(Tous)" and col in self._df.columns:
                mask &= (self._df[col] == val)
        df_f = self._df[mask]
        self._fill_table(df_f)
        self._count_var.set(f"{len(df_f)} ligne(s)")
        self._status.set(f"Filtre appliqué — {len(df_f)} ligne(s) affichée(s)")

    def _fill_table(self, df):
        self._tv.delete(*self._tv.get_children())
        for i, (_, row) in enumerate(df.iterrows()):
            tag = "even" if i % 2 == 0 else "odd"
            self._tv.insert("", "end",
                             values=[row[c] for c in DIANA_COLS],
                             tags=(tag,))

    def _sort_col(self, col):
        data = [(self._tv.set(k, col), k) for k in self._tv.get_children("")]
        data.sort(reverse=getattr(self._tv, f"_sort_rev_{col}", False))
        for i, (_, k) in enumerate(data):
            self._tv.move(k, "", i)
            self._tv.item(k, tags=("even" if i % 2 == 0 else "odd",))
        setattr(self._tv, f"_sort_rev_{col}",
                not getattr(self._tv, f"_sort_rev_{col}", False))

