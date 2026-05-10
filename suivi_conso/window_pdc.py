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

class PdcMajWindow(tk.Toplevel):

    def __init__(self, parent, path_jh, path_eur,
                 cache_usernames=None, cache_jh=None,
                 cache_eur=None, cache_eur_detail=None,
                 diana_jh=None, diana_eur=None, histo_tjm=None,
                 cache_ready=False, nucleus_exclusions=None):
        super().__init__(parent)
        self.title("Mise à jour Plan de charges")
        self.configure(bg=BG_MAIN)
        self.resizable(True, True)
        self.minsize(1100, 700)
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w, h = min(1500, sw - 60), min(860, sh - 60)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        self.path_jh   = path_jh
        self.path_eur  = path_eur
        self._df_jh    = None
        self._df_eur   = None
        self._interv_var      = tk.StringVar()
        self._proj_filter_var = tk.StringVar(value="Tous")
        self._annee       = datetime.now().year
        self._mois_cur    = _mois_effectif()
        self._modifs      = {}
        self._added_rows  = set()
        self._cells       = {}
        self._cache_ready     = cache_ready
        self._cache_usernames = cache_usernames or []
        self._cache_jh        = cache_jh  or {}
        self._cache_eur       = cache_eur or {}
        self._cache_eur_detail = cache_eur_detail or {}
        self._diana_jh           = diana_jh  or {}
        self._diana_eur          = diana_eur or {}
        self._histo_tjm          = histo_tjm or {}
        self._nucleus_exclusions = nucleus_exclusions or set()
        self._nucleus_only_var   = tk.BooleanVar(value=False)
        self._top_mode = "JH"    # "JH" | "TTNR"  — toggle cadran haut-gauche
        self._bot_mode = "TTNR"  # "JH" | "TTNR"  — toggle cadran bas (gauche+droite)
        # État de tri par tableau : {"jh": ("col_key", "asc"|"desc"), ...}
        # col_key : "code" | "lib" | 1..12 | "total"
        self._sort_state = {}    # clé : "jh" | "ttnr" | "eur" | "jh_global"
        self._update_top_toggle  = None   # callable(mode) — rafraîchit l'apparence du toggle haut
        self._update_bot_toggle  = None   # callable(mode) — rafraîchit l'apparence du toggle bas
        self._log_path = os.path.join(
            os.path.dirname(os.path.abspath(path_jh)),
            "Modifs_plan_de_charge.log")

        self._budgets, self._survcomm_projects, self._vacance_projects, self._analytic_projects = self._load_budgets()
        self._survcomm_intervenants = set()
        if self._survcomm_projects:
            for u_orig in (self._cache_usernames or []):
                u_norm = _norm_name(u_orig)
                projs  = set(self._cache_jh.get(u_norm, {}).keys())
                projs_hors_vacance = projs - self._vacance_projects
                if projs_hors_vacance and projs_hors_vacance.issubset(self._survcomm_projects):
                    self._survcomm_intervenants.add(u_orig)
        self._analytic_intervenants = set()
        if self._analytic_projects:
            for u_orig in (self._cache_usernames or []):
                u_norm = _norm_name(u_orig)
                projs  = set(self._cache_jh.get(u_norm, {}).keys())
                projs_hors_vacance = projs - self._vacance_projects
                if projs_hors_vacance and projs_hors_vacance.issubset(self._analytic_projects):
                    self._analytic_intervenants.add(u_orig)
        self._intervenants_orig, self._intervenants_norm = self._load_all_intervenants()
        self._build_ui()
        self.after(100, self._load_data)

    def _load_budgets(self):
        import os, csv
        log = self.master._log

        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(os.getcwd(), "Budgets_2026.csv"),
            os.path.join(script_dir,  "Budgets_2026.csv"),
        ]
        path = next((p for p in candidates if os.path.isfile(p)), None)

        result   = {}
        survcomm = set()
        vacance  = set()
        analytic = set()

        if path is None:
            log("  ⚠  Budgets_2026.csv introuvable. Chemins recherchés :", "warn")
            for p in candidates:
                log(f"      • {p}", "warn")
            return result, survcomm, vacance, analytic

        enc = "utf-8-sig"
        for c in ("utf-8-sig", "utf-8", "windows-1252", "latin-1"):
            try:
                with open(path, "r", encoding=c) as f:
                    f.read(1024)
                enc = c
                break
            except (UnicodeDecodeError, LookupError):
                continue
        try:
            with open(path, "r", encoding=enc, errors="replace", newline="") as f:
                reader = csv.DictReader(f, delimiter=";")
                raw_fields = reader.fieldnames or []
                norm_fields = {
                    fn.strip().lower().replace(" ", "_"): fn
                    for fn in raw_fields
                }
                col_pc  = norm_fields.get("project_code", "Project Code")
                col_bud = norm_fields.get("budget",       "Budget")
                col_top      = norm_fields.get("topsurvcomm",  "TopSurvComm")
                col_vac      = norm_fields.get("topabsence",
                               norm_fields.get("top_absence",
                               norm_fields.get("topvacance",
                               norm_fields.get("top_vacance",  "TopAbsence"))))
                col_analytic = norm_fields.get("topanalytic", "TopAnalytic")

                for row in reader:
                    pc  = str(row.get(col_pc,  "")).strip()
                    raw = str(row.get(col_bud, "")).strip().replace(",", ".")
                    top   = str(row.get(col_top,      "")).strip()
                    vac   = str(row.get(col_vac,      "")).strip().split(".")[0]
                    top_a = str(row.get(col_analytic, "")).strip()
                    if not pc:
                        continue
                    try:
                        result[pc] = float(raw)
                    except ValueError:
                        pass
                    if top:
                        survcomm.add(pc)
                    if vac == "1":
                        vacance.add(pc)
                    if top_a:
                        analytic.add(pc)

            log("━" * 54, "section")
            log(f"  Budgets_2026.csv chargé depuis {path} (encodage : {enc})", "ok")
            log(f"  Structure : {len(raw_fields)} colonne(s) → "
                f"{', '.join(raw_fields)}", "info")
            log(f"  Colonnes résolues : Project Code='{col_pc}' | "
                f"Budget='{col_bud}' | TopSurvComm='{col_top}' | "
                f"TopAbsence='{col_vac}' | TopAnalytic='{col_analytic}'", "info")
            with open(path, "r", encoding=enc, errors="replace", newline="") as f2:
                raw_lines = [l.rstrip("\n\r") for l in f2.readlines()]
            log(f"  Contenu brut ({len(raw_lines)} ligne(s)) :", "info")
            for line in raw_lines:
                log(f"    {line}", "info")
            log(f"  {len(result)} projet(s) avec budget, "
                f"{len(survcomm)} projet(s) SurvComm, "
                f"{len(vacance)} projet(s) TopAbsence, "
                f"{len(analytic)} projet(s) Analytic.", "ok")

            survcomm_users = []
            for u_orig in (self._cache_usernames or []):
                u_norm = _norm_name(u_orig)
                projs  = set(self._cache_jh.get(u_norm, {}).keys())
                projs_hors_vacance = projs - vacance
                if projs_hors_vacance and projs_hors_vacance.issubset(survcomm):
                    survcomm_users.append(u_orig)
            if survcomm_users:
                log(f"  Intervenant(s) exclusivement SurvComm "
                    f"({len(survcomm_users)}) — exclus des écarts JH/JO :", "info")
                for u in survcomm_users:
                    log(f"      • {u}", "info")
            else:
                log("  Aucun intervenant exclusivement SurvComm.", "info")

        except Exception as e:
            log(f"  ✘  Erreur lecture Budgets_2026.csv : {e}", "error")

        return result, survcomm, vacance, analytic

    def _load_all_intervenants(self):
        """Charge intervenants.csv et retourne (set noms originaux, set noms normalisés)
        pour TOUS les intervenants présents dans le fichier (sans filtre de colonne)."""
        import os, csv as _csv
        log = self.master._log
        path = getattr(self.master, "file4_path", None)
        path = path.get() if path else ""
        if not path or not os.path.isfile(path):
            log("  ⚠  intervenants.csv introuvable — restriction P02508 non-intervenants désactivée.", "warn")
            return set(), set()
        orig = set()
        norm = set()
        try:
            with open(path, encoding="utf-8-sig", errors="replace") as f:
                reader = _csv.DictReader(f, delimiter=";")
                for row in reader:
                    name = str(row.get("Intervenant", "")).strip()
                    if name:
                        orig.add(name)
                        norm.add(_norm_name(name))
        except Exception as e:
            log(f"  ✘  Erreur lecture intervenants.csv : {e}", "error")
        return orig, norm

    def _load_data(self):
        if self._cache_ready and self._cache_usernames:
            _sorted = sorted(self._cache_usernames, key=_norm_name) + ["Tous"]
            self._cb_interv["values"] = _sorted
            if _sorted:
                self._interv_var.set(_sorted[0])
                self._populate_proj_combo()
                self._refresh_tables()
            self._status.set(
                f"Cache utilisé — {len(self._cache_usernames)} username(s) RGU")
            return

        self._status.set("Chargement des fichiers PDC…")
        self.update_idletasks()
        try:
            self._df_jh  = self._read_pdc_df(self.path_jh)
            self._df_eur = self._read_pdc_df(self.path_eur)

            cols_jh = list(self._df_jh.columns)
            self._status.set(f"Colonnes JH : {cols_jh}")
            self.update_idletasks()

            def _norm_col(s):
                return s.lower().replace(" ", "").replace("_", "")
            col_map_jh = {_norm_col(c): c for c in cols_jh}

            interv_col = col_map_jh.get(_norm_col(COL_INTERVENANT), COL_INTERVENANT)
            mois_col   = col_map_jh.get(_norm_col(COL_MOIS), COL_MOIS)
            yr_col     = next((c for c in cols_jh if "year" in c.lower()), None)
            ou_col     = next((c for c in cols_jh
                               if "organization" in c.lower()
                               or _norm_col(c) in ("codeut", "code ut")
                               or "type" in c.lower() and "intervenant" in c.lower()), None)

            self._interv_col = interv_col
            self._mois_col   = mois_col
            self._yr_col     = yr_col
            self._ou_col     = ou_col

            def _to_mois(v):
                v2 = str(v).strip().lower()
                r  = MOIS_ORDRE.get(v2)
                if r: return r
                try: return int(float(v2))
                except: return None

            self._df_jh["_mois_num"]  = self._df_jh[mois_col].apply(_to_mois)
            self._df_eur["_mois_num"] = self._df_eur[col_map_jh.get(_norm_col(mois_col), mois_col)
                                                      if mois_col in self._df_eur.columns
                                                      else COL_MOIS].apply(_to_mois) \
                if mois_col in self._df_eur.columns else pd.Series([None]*len(self._df_eur))

            RESP_COL = "PRJ_RespNiv1"
            RESP_VAL = "GUITTARD, Raphael"
            resp_col = next(
                (c for c in cols_jh if _norm_col(c) == _norm_col(RESP_COL)), None)
            if resp_col:
                resp_mask = (self._df_jh[resp_col].astype(str).str.strip() == RESP_VAL)
                intervenants = sorted(
                    self._df_jh[resp_mask][interv_col].dropna()
                    .replace("", pd.NA).dropna().unique().tolist())
                msg = (f"Chargé — {len(intervenants)} intervenant(s) "
                       f"({resp_col} = '{RESP_VAL}')")
            else:
                intervenants = sorted(
                    self._df_jh[interv_col].dropna()
                    .replace("", pd.NA).dropna().unique().tolist())
                msg = (f"⚠ Colonne '{RESP_COL}' introuvable — "
                       f"{len(intervenants)} intervenant(s) sans filtre. "
                       f"Colonnes : {cols_jh[:8]}")
            intervenants_with_tous = intervenants + ["Tous"]
            self._cb_interv["values"] = intervenants_with_tous
            if intervenants_with_tous:
                self._interv_var.set(intervenants_with_tous[0])
                self._populate_proj_combo()
                self._refresh_tables()
            self._status.set(msg)
        except Exception as e:
            self._status.set(f"Erreur : {e}")
            messagebox.showerror("Erreur chargement", traceback.format_exc(), parent=self)

    # ── Helpers filtre projet ─────────────────────────────────────────────

    def _populate_proj_combo(self):
        """Peuple la ComboBox 'Projet' depuis le cache JH (lib_lookup)."""
        lib_lookup = {}
        for prj_dict in self._cache_jh.values():
            for prj, prj_data in prj_dict.items():
                if prj not in lib_lookup and isinstance(prj_data, dict):
                    lib_lookup[prj] = prj_data.get("_lib", "")
        items = ["Tous"] + sorted(
            f"{pc} - {lib}" if lib else pc
            for pc, lib in lib_lookup.items()
            if pc not in EXCLUDED_PROJECTS
        )
        self._cb_proj["values"] = items
        self._proj_filter_var.set("Tous")

    def _get_proj_filter(self):
        """Retourne le code projet sélectionné, ou None si 'Tous'."""
        val = self._proj_filter_var.get()
        if not val or val == "Tous":
            return None
        return val.split(" - ")[0].strip()

    def _on_proj_filter_change(self, *_):
        """Rappelé au changement de la ComboBox Projet — rafraîchit le cadran haut."""
        interv = self._interv_var.get()
        if not interv or not (self._cache_ready or self._df_jh is not None):
            return
        self._modifs.clear()
        self._added_rows.clear()
        if self._top_mode == "JH":
            self._build_table_jh(interv)
        else:
            self._build_table_ttnr_interv(interv)

    def _read_pdc_df(self, path):
        ext = path.lower().rsplit(".", 1)[-1]
        if ext == "csv":
            enc = "utf-8-sig"
            for c in ("utf-8-sig", "utf-8", "windows-1252", "latin-1"):
                try:
                    with open(path, "r", encoding=c) as f: f.read(8192)
                    enc = c; break
                except UnicodeDecodeError: continue
            with open(path, "r", encoding=enc, errors="replace") as f:
                sample = f.read(8192)
            sep = ";" if sample.count(";") >= sample.count(",") else ","
            df = pd.read_csv(path, dtype=str, sep=sep, encoding=enc,
                              encoding_errors="replace", on_bad_lines="skip")
        else:
            df = pd.read_excel(path, dtype=str, header=3)
        df.columns = [c.strip() for c in df.columns]
        return df.fillna("")

    def _refresh_tables(self, *_):
        interv = self._interv_var.get()
        if not interv:
            return
        if not self._cache_ready and self._df_jh is None:
            return
        self._modifs.clear()
        self._added_rows.clear()
        if self._top_mode == "JH":
            self._build_table_jh(interv)
        else:
            self._build_table_ttnr_interv(interv)
        if self._bot_mode == "TTNR":
            self._build_table_eur(interv)
        else:
            self._build_table_jh_global()
        self._build_ecart_list()

    def _on_interv_change(self, *_):
        """Rappelé au changement d'intervenant via la ComboBox.
        Rafraîchit uniquement le cadran haut — le cadran bas, la sélection
        projet et le tableau de détail sont conservés intacts."""
        interv = self._interv_var.get()
        if not interv:
            return
        if not self._cache_ready and self._df_jh is None:
            return
        self._modifs.clear()
        self._added_rows.clear()
        if self._top_mode == "JH":
            self._build_table_jh(interv)
        else:
            self._build_table_ttnr_interv(interv)

    def _get_jh_data(self, interv):
        if self._cache_ready:
            mois_cur = self._mois_cur

            if interv == "Tous":
                result = {}
                # Mois passés : diana_jh agrégé tous intervenants
                for u_norm, diana_data in self._diana_jh.items():
                    for prj, prj_data in diana_data.items():
                        if prj in EXCLUDED_PROJECTS:
                            continue
                        if prj not in result:
                            result[prj] = {"_lib": ""}
                        for m in range(1, mois_cur):
                            result[prj][m] = result[prj].get(m, 0.0) + prj_data.get(m, 0.0)
                # Mois futurs : cache_jh agrégé tous intervenants
                for u_norm, pdc_data in self._cache_jh.items():
                    for prj, prj_data in pdc_data.items():
                        if prj in EXCLUDED_PROJECTS:
                            continue
                        if prj not in result:
                            result[prj] = {"_lib": prj_data.get("_lib", "")}
                        elif not result[prj].get("_lib"):
                            result[prj]["_lib"] = prj_data.get("_lib", "")
                        for m in range(mois_cur, 13):
                            result[prj][m] = result[prj].get(m, 0.0) + prj_data.get(m, 0.0)
                return result

            interv_norm = _norm_name(interv)
            pdc_data    = self._cache_jh.get(interv_norm, {})
            diana_data  = self._diana_jh.get(interv_norm, {})

            result = {}
            all_projs = set(pdc_data.keys()) | set(diana_data.keys())
            for prj in all_projs:
                lib = pdc_data.get(prj, {}).get("_lib", "")
                result[prj] = {"_lib": lib}
                for m in range(1, mois_cur):
                    result[prj][m] = diana_data.get(prj, {}).get(m, 0.0)
                for m in range(mois_cur, 13):
                    result[prj][m] = pdc_data.get(prj, {}).get(m, 0.0)
            return result

        ic = getattr(self, "_interv_col", COL_INTERVENANT)
        yc = getattr(self, "_yr_col",     None)

        def _find_col(df, name):
            if name in df.columns: return name
            k = name.lower().replace(" ","").replace("_","")
            return next((c for c in df.columns
                         if c.lower().replace(" ","").replace("_","") == k), name)

        pc_col  = _find_col(self._df_jh, COL_PROJET)
        jh_col  = _find_col(self._df_jh, COL_NBJOURS)
        lib_col = next((c for c in self._df_jh.columns
                        if "libelle" in c.lower() and "projet" in c.lower()), None)

        mask = self._df_jh[ic].astype(str).str.strip() == interv
        df   = self._df_jh[mask].copy()
        if yc and yc in df.columns:
            df = df[df[yc].astype(str).str.strip() == str(self._annee)]

        result = {}
        for _, row in df.iterrows():
            pc   = str(row.get(pc_col, "")).strip()
            mois = row.get("_mois_num")
            jh   = 0.0
            try: jh = float(str(row.get(jh_col, "0")).replace(",", "."))
            except: pass
            if not pc or mois is None: continue
            if pc not in result:
                lib = str(row.get(lib_col, "")) if lib_col else ""
                result[pc] = {"_lib": lib}
            result[pc][int(mois)] = result[pc].get(int(mois), 0.0) + jh
        return result

    def _get_eur_data(self, interv):
        if self._cache_ready:
            interv_norm = _norm_name(interv)
            projets_interv = set(self._cache_jh.get(interv_norm, {}).keys())
            return {p: v for p, v in self._cache_eur.items()
                    if p in projets_interv}

        def _find_col(df, name):
            if name in df.columns: return name
            k = name.lower().replace(" ","").replace("_","")
            return next((c for c in df.columns
                         if c.lower().replace(" ","").replace("_","") == k), name)

        pc_col  = _find_col(self._df_eur, COL_PROJET)
        eur_col = _find_col(self._df_eur, COL_NBJOURS)
        yc      = getattr(self, "_yr_col", None)

        df = self._df_eur.copy()
        if yc and yc in df.columns:
            df = df[df[yc].astype(str).str.strip() == str(self._annee)]
        result = {}
        for _, row in df.iterrows():
            pc  = str(row.get(pc_col, "")).strip()
            eur = 0.0
            try: eur = float(str(row.get(eur_col, "0")).replace(",", "."))
            except: pass
            if not pc: continue
            result[pc] = result.get(pc, 0.0) + eur
        return result

    def _build_ecart_list(self):
        for w in self._frame_ecart.winfo_children():
            w.destroy()

        JOURS_OUVRES  = _calc_jours_ouvres(self._annee)
        mois_futurs   = list(range(self._mois_cur, 13))

        nucleus_only = self._nucleus_only_var.get()
        ecarts = []
        for u_orig in (self._cache_usernames or []):
            if nucleus_only and u_orig in (self._survcomm_intervenants | self._analytic_intervenants):
                continue
            u_norm   = _norm_name(u_orig)
            pdc_prjs = self._cache_jh.get(u_norm, {})
            ecart_total   = 0.0
            has_ecart_fut = False
            for m in mois_futurs:
                jh_m = sum(
                    prj_data.get(m, 0.0)
                    for prj, prj_data in pdc_prjs.items()
                    if isinstance(prj_data, dict)
                    # Règle 1a : GUITTARD/ABDELLI — exclure leur JH sur projets SurvComm
                    and not (nucleus_only
                             and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM
                             and prj in self._survcomm_projects)
                    # Règle 1b : GUITTARD/TCHIEGAING — exclure leur JH sur projets Analytic
                    and not (nucleus_only
                             and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM
                             and prj in self._analytic_projects)
                )
                e = jh_m - JOURS_OUVRES[m - 1]
                ecart_total += e
                if abs(e) > 0.01:
                    has_ecart_fut = True
            if has_ecart_fut:
                ecarts.append((u_orig, ecart_total))

        C_HDR    = CA_GREEN
        W_INTERV = 28
        W_ECART  = 10

        # ── Tri ───────────────────────────────────────────────────────────
        col_e, dir_e = self._sort_state.get("ecart", ("interv", "asc"))
        rev_e = (dir_e == "desc")
        if col_e == "ecart":
            ecarts.sort(key=lambda x: x[1], reverse=rev_e)
        else:
            ecarts.sort(key=lambda x: _norm_name(x[0]), reverse=rev_e)

        # ── En-têtes cliquables ───────────────────────────────────────────
        for col_key, txt, col, w, anc in [
            ("interv", "Intervenant", 0, W_INTERV, "w"),
            ("ecart",  "Écart JH",   1, W_ECART,  "e"),
        ]:
            lbl = tk.Label(self._frame_ecart,
                           text=self._sort_indicator("ecart", col_key, txt),
                           bg=C_HDR, fg="#ffffff", font=FONT_SMALL_BOLD,
                           width=w, anchor=anc, relief="flat", bd=1,
                           cursor="hand2")
            lbl.grid(row=0, column=col, sticky="ew", padx=1, pady=1)
            lbl.bind("<Button-1>",
                     lambda e, ck=col_key: self._on_sort(
                         "ecart", ck, self._build_ecart_list))

        self._ecart_usernames = []
        for i, (u_orig, ecart) in enumerate(ecarts, start=1):
            bg_r  = "#ffffff" if i % 2 == 0 else "#F4F4F4"
            sign  = "+" if ecart > 0 else ""
            e_str = f"{sign}{round(ecart)}"
            fg_e  = WARN if ecart > 0.01 else (ACCENT if ecart < -0.01 else TEXT_PRI)

            lbl_n = tk.Label(self._frame_ecart, text=u_orig[:W_INTERV],
                             bg=bg_r, fg="#001a33", font=FONT_SMALL,
                             width=W_INTERV, anchor="w")
            lbl_n.grid(row=i, column=0, sticky="ew", padx=1, pady=1)
            lbl_e = tk.Label(self._frame_ecart, text=e_str,
                             bg=bg_r, fg=fg_e, font=FONT_SMALL,
                             width=W_ECART, anchor="e")
            lbl_e.grid(row=i, column=1, sticky="ew", padx=1, pady=1)

            self._ecart_usernames.append(u_orig)

            def _on_click(event, u=u_orig):
                if u == self._interv_var.get():
                    return
                self._interv_var.set(u)
                self._refresh_tables()

            lbl_n.bind("<Button-1>", _on_click)
            lbl_e.bind("<Button-1>", _on_click)

    def _sync_ecart_height(self):
        pass

    # ── Helpers tri ─────────────────────────────────────────────────────────
    def _on_sort(self, table_key, col_key, rebuild_fn):
        """Inverse le sens de tri si même colonne, sinon repart en 'asc'."""
        cur_col, cur_dir = self._sort_state.get(table_key, (None, None))
        if cur_col == col_key:
            new_dir = "desc" if cur_dir == "asc" else "asc"
        else:
            new_dir = "asc"
        self._sort_state[table_key] = (col_key, new_dir)
        rebuild_fn()

    def _sort_indicator(self, table_key, col_key, label):
        """Retourne le libellé avec indicateur ▲/▼ si colonne active."""
        cur_col, cur_dir = self._sort_state.get(table_key, (None, None))
        if cur_col == col_key:
            return f"{label} {'▲' if cur_dir == 'asc' else '▼'}"
        return label

    def _apply_sort(self, table_key, projs, data_fn, default_key="code"):
        """Trie la liste projs selon l'état de tri du tableau.
        data_fn(pc, col_key) → valeur comparable (float ou str)."""
        col_key, direction = self._sort_state.get(table_key, (default_key, "asc"))
        reverse = (direction == "desc")
        return sorted(projs, key=lambda pc: data_fn(pc, col_key), reverse=reverse)

    def _build_table_jh(self, interv):
        for w in self._frame_jh.winfo_children():
            w.destroy()
        self._cells.clear()

        proj_filter = self._get_proj_filter()
        # Mode "lignes par intervenant" : projet sélectionné + "Tous" intervenants
        interv_mode = (proj_filter is not None and interv == "Tous"
                       and self._cache_ready)

        now_m = self._mois_cur

        C_HDR    = CA_GREEN
        C_PAST   = "#c8e6c9"
        C_PAST_H = "#2e5c2e"
        C_OVER   = "#f0883e"

        JOURS_OUVRES = _calc_jours_ouvres(self._annee)

        def _jh_hdr(text, col_key, col, width, anchor="w", bg=C_HDR, fg="#ffffff"):
            lbl = tk.Label(self._frame_jh,
                           text=self._sort_indicator("jh", col_key, text),
                           bg=bg, fg=fg, font=FONT_SMALL_BOLD,
                           width=width, anchor=anchor,
                           relief="flat", bd=1, cursor="hand2")
            lbl.grid(row=0, column=col, sticky="ew", padx=1, pady=1)
            lbl.bind("<Button-1>",
                     lambda e, ck=col_key: self._on_sort(
                         "jh", ck, lambda: self._build_table_jh(interv)))

        if interv_mode:
            # ── Mode intervenant : 1 ligne par intervenant pour le projet ──
            norm_to_orig = {_norm_name(u): u for u in (self._cache_usernames or [])}
            all_u_norms  = set(self._cache_jh.keys()) | set(self._diana_jh.keys())

            data_interv = {}
            for u_norm in all_u_norms:
                u_orig  = norm_to_orig.get(u_norm, u_norm)
                pdc_d   = self._cache_jh.get(u_norm, {}).get(proj_filter, {})
                dia_d   = self._diana_jh.get(u_norm, {}).get(proj_filter, {})
                entry   = {}
                for m in range(1, now_m):
                    v = dia_d.get(m, 0.0)
                    if v: entry[m] = v
                for m in range(now_m, 13):
                    v = pdc_d.get(m, 0.0)
                    if v: entry[m] = v
                if any(entry.get(m, 0.0) for m in range(1, 13)):
                    data_interv[u_orig] = entry

            def _sort_val_i(u_orig, col_key):
                if col_key in ("code", "lib"): return u_orig.lower()
                if col_key == "total":
                    return sum(data_interv[u_orig].get(m, 0.0) for m in range(1, 13))
                return data_interv[u_orig].get(col_key, 0.0)

            users = self._apply_sort("jh", list(data_interv.keys()),
                                     _sort_val_i, default_key="code")

            # En-têtes : Intervenant (col 1, col 0 vide) + mois + Total
            _jh_hdr("Intervenant", "code", 1, 42, anchor="w")
            for m in range(1, 13):
                bg_h = C_PAST if m < now_m else C_HDR
                fg_h = C_PAST_H if m < now_m else "#ffffff"
                _jh_hdr(MOIS_ABBREV[m-1], m, m+1, 7, anchor="center", bg=bg_h, fg=fg_h)
            _jh_hdr("Total", "total", 14, 11, anchor="center")

            for ri, u_orig in enumerate(users, start=1):
                entry = data_interv[u_orig]
                bg_r  = "#ffffff" if ri % 2 == 0 else "#F4F4F4"
                tk.Label(self._frame_jh, text=u_orig[:42], bg=bg_r, fg="#000000",
                         font=FONT_SMALL, width=42, anchor="w").grid(
                    row=ri, column=1, sticky="ew", padx=1, pady=1)
                row_total = 0.0
                for m in range(1, 13):
                    val  = entry.get(m, 0.0)
                    row_total += val
                    bg_c = C_PAST if m < now_m else bg_r
                    fg_c = C_PAST_H if m < now_m else "#000000"
                    disp = str(int(val)) if val == int(val) else f"{val:.1f}"
                    tk.Label(self._frame_jh, text=disp if val else "",
                             bg=bg_c, fg=fg_c, font=FONT_SMALL,
                             width=7, anchor="center").grid(
                        row=ri, column=m+1, sticky="ew", padx=1, pady=1)
                tot_disp = str(int(row_total)) if row_total == int(row_total) else f"{row_total:.1f}"
                tk.Label(self._frame_jh, text=tot_disp, bg=bg_r, fg="#000000",
                         font=FONT_SMALL_BOLD, width=11, anchor="center").grid(
                    row=ri, column=14, sticky="ew", padx=1, pady=1)

            tot_row = len(users) + 1
            C_TOT   = "#F7F7F7"
            tk.Label(self._frame_jh, text="Total", bg=C_TOT, fg="#001a33",
                     font=FONT_SMALL_BOLD, width=42, anchor="e").grid(
                row=tot_row, column=1, sticky="ew", padx=1, pady=1)
            grand_total = 0.0
            for m in range(1, 13):
                col_tot = sum(data_interv[u].get(m, 0.0) for u in users)
                grand_total += col_tot
                jo   = JOURS_OUVRES[m-1]
                over = abs(col_tot - jo) > 0.01 and m >= now_m
                bg_t = C_PAST if m < now_m else (C_OVER if over else C_TOT)
                fg_t = "#000000" if over else (C_PAST_H if m < now_m else "#000000")
                disp = str(int(col_tot)) if col_tot == int(col_tot) else f"{col_tot:.1f}"
                tk.Label(self._frame_jh, text=disp if col_tot else "",
                         bg=bg_t, fg=fg_t, font=FONT_SMALL_BOLD,
                         width=7, anchor="center").grid(
                    row=tot_row, column=m+1, sticky="ew", padx=1, pady=1)
            gt = str(int(grand_total)) if grand_total == int(grand_total) else f"{grand_total:.1f}"
            tk.Label(self._frame_jh, text=gt, bg=C_TOT, fg="#000000",
                     font=FONT_SMALL_BOLD, width=11, anchor="center").grid(
                row=tot_row, column=14, sticky="ew", padx=1, pady=1)
            return

        # ── Mode normal (Project Code | Project Name | mois | Total) ───────
        data  = self._get_jh_data(interv)

        # ── Lib map (nécessaire avant le tri par libellé) ─────────────────
        lib_map = {}
        if self._cache_ready:
            interv_norm = _norm_name(interv) if interv != "Tous" else None
            src = self._cache_jh.get(interv_norm, {}) if interv_norm else {}
            for pc2, prj_dict in src.items():
                lib_map[pc2] = prj_dict.get("_lib", prj_dict.get("_libelle", ""))
            if interv == "Tous":
                for pd2 in self._cache_jh.values():
                    for pc2, prj_dict in pd2.items():
                        if pc2 not in lib_map and isinstance(prj_dict, dict):
                            lib_map[pc2] = prj_dict.get("_lib", "")

        def _jh_sort_val(pc, col_key):
            if col_key == "code":  return pc
            if col_key == "lib":   return lib_map.get(pc, data[pc].get("_lib", "")).lower()
            if col_key == "total": return sum(data[pc].get(m, 0.0) for m in range(1, 13))
            return data[pc].get(col_key, 0.0)  # col_key = mois 1..12

        projs = self._apply_sort("jh", list(data.keys()), _jh_sort_val, default_key="code")
        if proj_filter is not None:
            projs = [p for p in projs if p == proj_filter]
        self._proj_rows = projs[:]

        _jh_hdr("Project Code", "code",  0,  12, anchor="w")
        _jh_hdr("Project Name", "lib",   1,  30, anchor="w")
        for m in range(1, 13):
            bg_h = C_PAST if m < now_m else C_HDR
            fg_h = C_PAST_H if m < now_m else "#ffffff"
            _jh_hdr(MOIS_ABBREV[m-1], m, m+1, 7, anchor="center", bg=bg_h, fg=fg_h)
        _jh_hdr("Total", "total", 14, 11, anchor="center")

        for ri, pc in enumerate(projs, start=1):
            info = data[pc]
            lib  = lib_map.get(pc, info.get("_lib", info.get("_libelle", "")))
            if not lib:
                if self._cache_ready:
                    for u_norm, pd2 in self._cache_jh.items():
                        if pc in pd2 and pd2[pc].get("_lib"):
                            lib = pd2[pc]["_lib"]; break
            bg_r = "#ffffff" if ri % 2 == 0 else "#F4F4F4"
            tk.Label(self._frame_jh, text=pc, bg=bg_r, fg="#000000",
                     font=FONT_SMALL, width=12, anchor="w").grid(
                row=ri, column=0, sticky="ew", padx=1, pady=1)
            tk.Label(self._frame_jh, text=lib[:36], bg=bg_r, fg="#000000",
                     font=FONT_SMALL, width=30, anchor="w").grid(
                row=ri, column=1, sticky="ew", padx=1, pady=1)
            row_total = 0.0
            for m in range(1, 13):
                val = info.get(m, 0.0)
                row_total += val
                editable = (m < now_m)
                bg_c = C_PAST if m < now_m else bg_r
                fg_c = C_PAST_H if m < now_m else "#000000"
                disp = str(int(val)) if val == int(val) else f"{val:.1f}"
                if editable:
                    var = tk.StringVar(value=disp)
                    e = tk.Entry(self._frame_jh, textvariable=var,
                                 bg=bg_c, fg=fg_c, font=FONT_SMALL,
                                 width=7, justify="center", bd=0,
                                 insertbackground=ACCENT,
                                 relief="flat",
                                 highlightbackground=BORDER,
                                 highlightthickness=1)
                    e.grid(row=ri, column=m+1, sticky="ew", padx=1, pady=1)
                    e.bind("<FocusOut>",
                           lambda ev, p=pc, mo=m, v=var, orig=val:
                           self._on_cell_edit(p, mo, v, orig))
                    self._cells[(pc, m)] = (var, val)
                else:
                    tk.Label(self._frame_jh, text=disp if val else "",
                             bg=bg_c, fg=fg_c, font=FONT_SMALL,
                             width=7, anchor="center").grid(
                        row=ri, column=m+1, sticky="ew", padx=1, pady=1)
            tot_disp = str(int(row_total)) if row_total == int(row_total) else f"{row_total:.1f}"
            tk.Label(self._frame_jh, text=tot_disp, bg=bg_r, fg="#000000",
                     font=FONT_SMALL_BOLD, width=11, anchor="center").grid(
                row=ri, column=14, sticky="ew", padx=1, pady=1)

        n_rows  = len(projs)
        tot_row = n_rows + 1
        C_TOT = "#F7F7F7"
        tk.Label(self._frame_jh, text="Total", bg=C_TOT, fg="#001a33",
                 font=FONT_SMALL_BOLD, width=12, anchor="e").grid(
            row=tot_row, column=0, columnspan=2, sticky="ew", padx=1, pady=1)
        grand_total = 0.0
        for m in range(1, 13):
            col_tot = sum(data[pc].get(m, 0.0) for pc in projs)
            grand_total += col_tot
            jo   = JOURS_OUVRES[m-1]
            over = abs(col_tot - jo) > 0.01 and m >= now_m
            bg_t = C_PAST if m < now_m else (C_OVER if over else C_TOT)
            fg_t = "#000000" if over else (C_PAST_H if m < now_m else "#000000")
            disp = str(int(col_tot)) if col_tot == int(col_tot) else f"{col_tot:.1f}"
            tk.Label(self._frame_jh, text=disp if col_tot else "",
                     bg=bg_t, fg=fg_t,
                     font=FONT_SMALL_BOLD, width=7, anchor="center").grid(
                row=tot_row, column=m+1, sticky="ew", padx=1, pady=1)
        gt = str(int(grand_total)) if grand_total == int(grand_total) else f"{grand_total:.1f}"
        tk.Label(self._frame_jh, text=gt, bg=C_TOT, fg="#000000",
                 font=FONT_SMALL_BOLD, width=11, anchor="center").grid(
            row=tot_row, column=14, sticky="ew", padx=1, pady=1)

        if interv != "Tous":
            jo_row = tot_row + 1
            tk.Label(self._frame_jh, text="Jours ouvrés", bg=C_TOT, fg=TEXT_SEC,
                     font=FONT_SMALL, width=12, anchor="e").grid(
                row=jo_row, column=0, columnspan=2, sticky="ew", padx=1, pady=1)
            jo_total = 0
            for m in range(1, 13):
                jo = JOURS_OUVRES[m-1]
                jo_total += jo
                bg_j = C_PAST if m < now_m else C_TOT
                tk.Label(self._frame_jh, text=str(jo), bg=bg_j, fg=TEXT_SEC,
                         font=FONT_SMALL, width=7, anchor="center").grid(
                    row=jo_row, column=m+1, sticky="ew", padx=1, pady=1)
            tk.Label(self._frame_jh, text=str(jo_total), bg=C_TOT, fg=TEXT_SEC,
                     font=FONT_SMALL, width=11, anchor="center").grid(
                row=jo_row, column=14, sticky="ew", padx=1, pady=1)

    def _build_table_eur(self, interv):
        for w in self._frame_eur.winfo_children():
            w.destroy()
        for w in self._frame_detail.winfo_children():
            w.destroy()
        self._selected_proj = None

        mois_passes = list(range(1, self._mois_cur))

        all_projs = set()
        if self._cache_ready:
            for prj_dict in self._diana_eur.values():
                all_projs.update(prj_dict.keys())
            for (u_norm, prj, m) in self._cache_eur_detail:
                all_projs.add(prj)
        raw_projs = [p for p in all_projs if p not in EXCLUDED_PROJECTS]

        lib_lookup = {}
        if self._cache_ready:
            for prj_dict in self._cache_jh.values():
                for prj, prj_data in prj_dict.items():
                    if prj not in lib_lookup and isinstance(prj_data, dict):
                        lib_lookup[prj] = prj_data.get("_lib", "")

        def _fmt(v):
            return f"{round(v):,}".replace(",", "\u202f") if v else ""

        mois_futurs  = list(range(self._mois_cur, 13))
        nucleus_only = self._nucleus_only_var.get()

        # ── Pré-calcul Consommé / RAF / Atterrissage / Budget par projet ──
        eur_vals = {}   # {pc: (conso_date, raf, atterr, budget)}
        for pc in raw_projs:
            conso_date = 0.0
            if self._cache_ready and self._diana_eur:
                for u_norm, prj_dict in self._diana_eur.items():
                    if nucleus_only and u_norm in self._nucleus_exclusions:
                        continue
                    if nucleus_only and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM and pc in self._survcomm_projects:
                        continue
                    if nucleus_only and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM and pc in self._analytic_projects:
                        continue
                    for m, eur_v in prj_dict.get(pc, {}).items():
                        if m in mois_passes:
                            conso_date += eur_v
            elif self._cache_ready and self._cache_eur_detail:
                for (u_norm, prj, m), eur_val in self._cache_eur_detail.items():
                    if nucleus_only and u_norm in self._nucleus_exclusions:
                        continue
                    if nucleus_only and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM and prj in self._survcomm_projects:
                        continue
                    if nucleus_only and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM and prj in self._analytic_projects:
                        continue
                    if prj == pc and m in mois_passes:
                        conso_date += eur_val

            raf = 0.0
            if self._cache_ready and self._cache_eur_detail:
                for (u_norm, prj, m), eur_val in self._cache_eur_detail.items():
                    if nucleus_only and u_norm in self._nucleus_exclusions:
                        continue
                    if nucleus_only and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM and prj in self._survcomm_projects:
                        continue
                    if nucleus_only and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM and prj in self._analytic_projects:
                        continue
                    if nucleus_only and prj not in _NUCLEUS_ALLOWED_PROJECTS and u_norm not in self._intervenants_norm:
                        continue
                    if prj == pc and m in mois_futurs:
                        raf += eur_val

            atterr  = round(conso_date + raf, 2)
            budget  = self._budgets.get(pc) or 0.0
            eur_vals[pc] = (conso_date, raf, atterr, budget)

        # ── Tri ───────────────────────────────────────────────────────────
        _col_idx = {"conso": 0, "raf": 1, "atterr": 2, "budget": 3}
        def _eur_sort_val(pc, col_key):
            if col_key == "code":  return pc
            if col_key == "lib":   return lib_lookup.get(pc, "").lower()
            idx = _col_idx.get(col_key)
            if idx is not None:    return eur_vals[pc][idx]
            return 0.0
        # Tri par défaut : budget décroissant (comportement d'origine)
        if "eur" not in self._sort_state:
            projs = sorted(raw_projs, key=lambda p: eur_vals[p][3], reverse=True)
        else:
            projs = self._apply_sort("eur", raw_projs, _eur_sort_val, default_key="budget")

        # ── En-têtes cliquables ───────────────────────────────────────────
        C_HDR = CA_GREEN
        self._eur_row_widgets = {}
        col_defs = [
            ("Project Code",    "code",  12, "w"),
            ("Project Name",    "lib",   34, "w"),
            ("Consommé à date", "conso", 16, "e"),
            ("RAF",             "raf",   16, "e"),
            ("Atterrissage",    "atterr",16, "e"),
            ("Budget",          "budget",14, "e"),
        ]
        for ci, (txt, col_key, w, anc) in enumerate(col_defs):
            lbl = tk.Label(self._frame_eur,
                           text=self._sort_indicator("eur", col_key, txt),
                           bg=C_HDR, fg="#ffffff", font=FONT_SMALL_BOLD,
                           width=w, anchor=anc, relief="flat", bd=1,
                           cursor="hand2")
            lbl.grid(row=0, column=ci, sticky="ew", padx=1, pady=1)
            lbl.bind("<Button-1>",
                     lambda e, ck=col_key: self._on_sort(
                         "eur", ck,
                         lambda: self._build_table_eur(self._interv_var.get())))

        # ── Lignes projets ────────────────────────────────────────────────
        for ri, pc in enumerate(projs, start=1):
            lib = lib_lookup.get(pc, "")
            conso_date, raf, atterr, budget = eur_vals[pc]

            bg_r = "#ffffff" if ri % 2 == 0 else "#F4F4F4"
            row_data = [
                (pc,              12, "w", "#000000"),
                (lib[:40],        34, "w", "#000000"),
                (_fmt(conso_date),16, "e", "#000000"),
                (_fmt(raf),       16, "e", "#000000"),
                (_fmt(atterr),    16, "e", ACCENT2),
                (_fmt(budget),    14, "e", "#000000"),
            ]
            row_lbls = []
            for ci, (txt, w, anc, fg) in enumerate(row_data):
                lbl = tk.Label(self._frame_eur, text=txt, bg=bg_r,
                               fg=fg, font=FONT_SMALL, width=w, anchor=anc,
                               cursor="hand2")
                lbl.grid(row=ri, column=ci, sticky="ew", padx=1, pady=1)
                lbl.bind("<Button-1>",
                         lambda e, p=pc, bg=bg_r: self._on_proj_click(p, bg))
                row_lbls.append(lbl)
            self._eur_row_widgets[pc] = (row_lbls, bg_r)

    def _on_proj_click(self, pc, bg_orig):
        if self._selected_proj and self._selected_proj in self._eur_row_widgets:
            lbls, bg_r = self._eur_row_widgets[self._selected_proj]
            for l in lbls: l.config(bg=bg_r)
        if pc in self._eur_row_widgets:
            lbls, _ = self._eur_row_widgets[pc]
            for l in lbls: l.config(bg="#c8e6c9")
        self._selected_proj = pc
        self._show_proj_detail(pc)

    def _compute_nucleus_exclusions(self):
        """Retourne les noms normalisés des intervenants à exclure du filtre Nucleus.
        Intervenants exclusifs SurvComm ET intervenants exclusifs Analytic.
        """
        return {_norm_name(u) for u in self._survcomm_intervenants | self._analytic_intervenants}

    def _on_nucleus_filter(self):
        """Rappelé au clic sur la checkbox — calcule les exclusions, log, rafraîchit le détail."""
        log = getattr(self.master, "_log", None)
        checked = self._nucleus_only_var.get()

        if log:
            log("─" * 50, "section")
            log(f"  Périmètre Nucleus uniquement : {'✔ activé' if checked else '✘ désactivé'}", "info")

        if checked:
            # Calcul dynamique depuis le cache JH
            self._nucleus_exclusions = self._compute_nucleus_exclusions()
            if self._nucleus_exclusions:
                # Récupérer le nom original pour l'affichage
                norm_to_orig = {_norm_name(u): u for u in (self._cache_usernames or [])}
                if log:
                    log(f"  Intervenants exclus (périmètre SurvComm/Analytic exclusif — {len(self._nucleus_exclusions)}) :", "info")
                    for norm in sorted(self._nucleus_exclusions):
                        display = norm_to_orig.get(norm, norm)
                        log(f"    • {display}", "info")
            else:
                if log:
                    log("  ⚠  Aucun intervenant à exclure — cache JH vide ou tous multi-périmètres.", "warn")

            # Exclusions en dur : GUITTARD/ABDELLI — projets SurvComm
            if log:
                log(f"  Exclusions partielles (multiprojets — projets SurvComm) :", "info")
                for name in _NUCLEUS_HARDCODED_SURVCOMM_ORIG:
                    log(f"    • {name} : consommés JH et EUR sur projets SurvComm exclus", "info")
            # Exclusions en dur : GUITTARD/TCHIEGAING — projets Analytic
            if log:
                log(f"  Exclusions partielles (multiprojets — projets Analytic) :", "info")
                for name in _NUCLEUS_HARDCODED_ANALYTIC_ORIG:
                    log(f"    • {name} : consommés JH et EUR sur projets Analytic exclus", "info")

            # Exclusions en dur : non-intervenants — RAF hors P02508
            if log:
                non_interv = [
                    u for u in (self._cache_usernames or [])
                    if _norm_name(u) not in self._intervenants_norm
                    and _norm_name(u) not in self._nucleus_exclusions
                    and _norm_name(u) not in _NUCLEUS_HARDCODED_NORM
                ]
                if non_interv:
                    log(f"  Intervenants hors intervenants.csv — RAF limité aux projets {sorted(_NUCLEUS_ALLOWED_PROJECTS)} ({len(non_interv)}) :", "info")
                    for u in non_interv:
                        log(f"    • {u}", "info")
                else:
                    log(f"  Aucun intervenant hors intervenants.csv dans le cache.", "info")
        else:
            self._nucleus_exclusions = set()

        # Rafraîchir les tableaux
        interv = self._interv_var.get()
        if interv:
            if log:
                log(f"  Rafraîchissement des tableaux pour : {interv}", "info")
            prev_proj = self._selected_proj
            self._build_ecart_list()
            if self._top_mode == "JH":
                self._build_table_jh(interv)
            else:
                self._build_table_ttnr_interv(interv)
            if self._bot_mode == "TTNR":
                self._build_table_eur(interv)
            else:
                self._build_table_jh_global()
            if prev_proj:
                self._selected_proj = prev_proj
                if prev_proj in self._eur_row_widgets:
                    lbls, _ = self._eur_row_widgets[prev_proj]
                    for l in lbls:
                        l.config(bg="#c8e6c9")
                self._show_proj_detail(prev_proj)
        else:
            if log:
                log("  ℹ  Aucun intervenant sélectionné.", "info")

    # ── _build_table_ttnr_interv ─────────────────────────────────────────
    def _build_table_ttnr_interv(self, interv):
        """Cadran haut-gauche mode TTNR : EUR par mois de l'intervenant (12 mois + Total).
        Passé = diana_eur (réel), Futur = cache_eur_detail (RAF PDC).
        Même structure visuelle que le tableau JH."""
        for w in self._frame_jh.winfo_children():
            w.destroy()

        if not self._cache_ready:
            return

        proj_filter  = self._get_proj_filter()
        tous         = (interv == "Tous")
        interv_norm  = None if tous else _norm_name(interv)
        now_m        = self._mois_cur
        mois_passes  = list(range(1, now_m))
        mois_futurs  = list(range(now_m, 13))
        nucleus_only = self._nucleus_only_var.get()
        # Mode "lignes par intervenant" : projet sélectionné + "Tous" intervenants
        interv_mode  = (proj_filter is not None and tous)

        C_HDR    = CA_GREEN
        C_PAST   = "#c8e6c9"
        C_PAST_H = "#2e5c2e"

        lib_lookup = {}
        for prj_dict in self._cache_jh.values():
            for prj, prj_data in prj_dict.items():
                if prj not in lib_lookup and isinstance(prj_data, dict):
                    lib_lookup[prj] = prj_data.get("_lib", "")

        def _fmt(v):
            """Format EUR : séparateur milliers espace insécable."""
            return f"{round(v):,}".replace(",", "\u202f") if v else ""

        def _ttnr_hdr(text, col_key, col, width, anchor="w", bg=C_HDR, fg="#ffffff"):
            lbl = tk.Label(self._frame_jh,
                           text=self._sort_indicator("ttnr", col_key, text),
                           bg=bg, fg=fg, font=FONT_SMALL_BOLD,
                           width=width, anchor=anchor,
                           relief="flat", bd=1, cursor="hand2")
            lbl.grid(row=0, column=col, sticky="ew", padx=1, pady=1)
            lbl.bind("<Button-1>",
                     lambda e, ck=col_key: self._on_sort(
                         "ttnr", ck, lambda: self._build_table_ttnr_interv(interv)))

        if interv_mode:
            # ── Mode intervenant : 1 ligne par intervenant pour le projet ──
            norm_to_orig = {_norm_name(u): u for u in (self._cache_usernames or [])}
            all_u_norms  = (set(self._diana_eur.keys())
                            | {un for (un, pr, m) in self._cache_eur_detail})

            # diana_pm_u : {u_norm: {m: eur}} pour le projet sélectionné
            # raf_pm_u   : {u_norm: {m: eur}} pour le projet sélectionné
            diana_pm_u = {}
            for u_norm, prj_dict in self._diana_eur.items():
                if nucleus_only and u_norm in self._nucleus_exclusions:
                    continue
                month_data = prj_dict.get(proj_filter, {})
                if nucleus_only and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM \
                        and proj_filter in self._survcomm_projects:
                    continue
                if nucleus_only and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM \
                        and proj_filter in self._analytic_projects:
                    continue
                for m, v in month_data.items():
                    if m in mois_passes and v:
                        diana_pm_u.setdefault(u_norm, {})[m] = v

            raf_pm_u = {}
            for (un, pr, m), v in self._cache_eur_detail.items():
                if pr != proj_filter or m not in mois_futurs or not v:
                    continue
                if nucleus_only and un in self._nucleus_exclusions:
                    continue
                if nucleus_only and un in _NUCLEUS_HARDCODED_SURVCOMM_NORM \
                        and pr in self._survcomm_projects:
                    continue
                if nucleus_only and un in _NUCLEUS_HARDCODED_ANALYTIC_NORM \
                        and pr in self._analytic_projects:
                    continue
                if nucleus_only and pr not in _NUCLEUS_ALLOWED_PROJECTS \
                        and un not in self._intervenants_norm:
                    continue
                raf_pm_u.setdefault(un, {})[m] = raf_pm_u.get(un, {}).get(m, 0.0) + v

            data_interv = {}
            for u_norm in all_u_norms:
                u_orig = norm_to_orig.get(u_norm, u_norm)
                d = dict(diana_pm_u.get(u_norm, {}))
                d.update(raf_pm_u.get(u_norm, {}))
                if any(d.get(m, 0) for m in range(1, 13)):
                    data_interv[u_orig] = {
                        "diana": diana_pm_u.get(u_norm, {}),
                        "raf":   raf_pm_u.get(u_norm, {}),
                    }

            def _sort_val_i(u_orig, col_key):
                if col_key in ("code", "lib"): return u_orig.lower()
                d = data_interv[u_orig]
                if col_key == "total":
                    return (sum(d["diana"].get(m, 0.0) for m in mois_passes)
                            + sum(d["raf"].get(m, 0.0) for m in mois_futurs))
                if isinstance(col_key, int):
                    return (d["diana"].get(col_key, 0.0) if col_key < now_m
                            else d["raf"].get(col_key, 0.0))
                return 0.0

            users = self._apply_sort("ttnr", list(data_interv.keys()),
                                     _sort_val_i, default_key="code")

            _ttnr_hdr("Intervenant", "code", 1, 42, anchor="w")
            for m in range(1, 13):
                bg_h = C_PAST if m < now_m else C_HDR
                fg_h = C_PAST_H if m < now_m else "#ffffff"
                _ttnr_hdr(MOIS_ABBREV[m-1], m, m+1, 7, anchor="center", bg=bg_h, fg=fg_h)
            _ttnr_hdr("Total", "total", 14, 11, anchor="center")

            col_totals  = [0.0] * 12
            grand_total = 0.0
            for ri, u_orig in enumerate(users, start=1):
                d    = data_interv[u_orig]
                bg_r = "#ffffff" if ri % 2 == 0 else "#F4F4F4"
                tk.Label(self._frame_jh, text=u_orig[:42], bg=bg_r, fg="#000000",
                         font=FONT_SMALL, width=42, anchor="w").grid(
                    row=ri, column=1, sticky="ew", padx=1, pady=1)
                row_total = 0.0
                for m in range(1, 13):
                    if m < now_m:
                        val  = d["diana"].get(m, 0.0)
                        bg_c = C_PAST
                        fg_c = C_PAST_H
                    else:
                        val  = d["raf"].get(m, 0.0)
                        bg_c = bg_r
                        fg_c = "#000000"
                    row_total         += val
                    col_totals[m - 1] += val
                    tk.Label(self._frame_jh, text=_fmt(val), bg=bg_c, fg=fg_c,
                             font=FONT_SMALL, width=7, anchor="e").grid(
                        row=ri, column=m+1, sticky="ew", padx=1, pady=1)
                grand_total += row_total
                tk.Label(self._frame_jh, text=_fmt(row_total), bg=bg_r, fg="#000000",
                         font=FONT_SMALL_BOLD, width=11, anchor="e").grid(
                    row=ri, column=14, sticky="ew", padx=1, pady=1)

            tot_row = len(users) + 1
            C_TOT   = "#F7F7F7"
            tk.Label(self._frame_jh, text="Total", bg=C_TOT, fg="#001a33",
                     font=FONT_SMALL_BOLD, width=42, anchor="e").grid(
                row=tot_row, column=1, sticky="ew", padx=1, pady=1)
            for m in range(1, 13):
                bg_t = C_PAST if m < now_m else C_TOT
                fg_t = C_PAST_H if m < now_m else "#000000"
                tk.Label(self._frame_jh, text=_fmt(col_totals[m - 1]), bg=bg_t, fg=fg_t,
                         font=FONT_SMALL_BOLD, width=7, anchor="e").grid(
                    row=tot_row, column=m+1, sticky="ew", padx=1, pady=1)
            tk.Label(self._frame_jh, text=_fmt(grand_total), bg=C_TOT, fg="#000000",
                     font=FONT_SMALL_BOLD, width=11, anchor="e").grid(
                row=tot_row, column=14, sticky="ew", padx=1, pady=1)
            return

        # ── Mode normal (Project Code | Project Name | mois | Total) ───────
        # Projets (tous intervenants si "Tous", sinon l'intervenant sélectionné)
        if tous:
            projs_jh    = set()
            for pd2 in self._cache_jh.values():
                projs_jh.update(pd2.keys())
            projs_diana = set()
            for pd2 in self._diana_eur.values():
                projs_diana.update(pd2.keys())
            projs_raf   = {pr for (un, pr, m) in self._cache_eur_detail}
        else:
            projs_jh    = set(self._cache_jh.get(interv_norm, {}).keys())
            projs_diana = set(self._diana_eur.get(interv_norm, {}).keys())
            projs_raf   = {pr for (un, pr, m) in self._cache_eur_detail
                           if un == interv_norm}
        all_projs_raw = [p for p in (projs_jh | projs_diana | projs_raf)
                         if p not in EXCLUDED_PROJECTS]
        if proj_filter is not None:
            all_projs_raw = [p for p in all_projs_raw if p == proj_filter]

        # ── Pré-calcul diana_pm : consommé réel (mois passés) avec filtres nucleus ──
        diana_pm = {}  # {pc: {m: eur}}
        for u_norm, prj_dict in self._diana_eur.items():
            if not tous and u_norm != interv_norm:
                continue
            if nucleus_only and u_norm in self._nucleus_exclusions:
                continue
            for pc, month_data in prj_dict.items():
                if nucleus_only and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM \
                        and pc in self._survcomm_projects:
                    continue
                if nucleus_only and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM \
                        and pc in self._analytic_projects:
                    continue
                for m, v in month_data.items():
                    if m in mois_passes:
                        diana_pm.setdefault(pc, {})[m] = diana_pm.get(pc, {}).get(m, 0.0) + v

        # ── Pré-calcul raf_pm : RAF (mois futurs) avec filtres nucleus ────────
        raf_pm = {}   # {pc: {m: eur}}
        for (un, pr, m), v in self._cache_eur_detail.items():
            if (not tous and un != interv_norm) or m not in mois_futurs:
                continue
            if nucleus_only and un in self._nucleus_exclusions:
                continue
            if nucleus_only and un in _NUCLEUS_HARDCODED_SURVCOMM_NORM \
                    and pr in self._survcomm_projects:
                continue
            if nucleus_only and un in _NUCLEUS_HARDCODED_ANALYTIC_NORM \
                    and pr in self._analytic_projects:
                continue
            if nucleus_only and pr not in _NUCLEUS_ALLOWED_PROJECTS \
                    and un not in self._intervenants_norm:
                continue
            raf_pm.setdefault(pr, {})[m] = raf_pm.get(pr, {}).get(m, 0.0) + v

        # ── Pré-calcul valeurs pour tri ───────────────────────────────────
        def _ttnr_val(pc, m_or_key):
            if m_or_key == "code":  return pc
            if m_or_key == "lib":   return lib_lookup.get(pc, "").lower()
            if isinstance(m_or_key, int):
                if m_or_key < now_m:
                    return diana_pm.get(pc, {}).get(m_or_key, 0.0)
                return raf_pm.get(pc, {}).get(m_or_key, 0.0)
            # "total"
            past  = sum(diana_pm.get(pc, {}).get(m, 0.0) for m in mois_passes)
            futur = sum(raf_pm.get(pc, {}).get(m, 0.0) for m in mois_futurs)
            return past + futur

        projs = self._apply_sort("ttnr", all_projs_raw, _ttnr_val, default_key="code")

        _ttnr_hdr("Project Code", "code",  0,  12, anchor="w")
        _ttnr_hdr("Project Name", "lib",   1,  30, anchor="w")
        for m in range(1, 13):
            bg_h = C_PAST if m < now_m else C_HDR
            fg_h = C_PAST_H if m < now_m else "#ffffff"
            _ttnr_hdr(MOIS_ABBREV[m-1], m, m+1, 7, anchor="center", bg=bg_h, fg=fg_h)
        _ttnr_hdr("Total", "total", 14, 11, anchor="center")

        col_totals  = [0.0] * 12   # index 0 = mois 1
        grand_total = 0.0

        # ── Lignes projets ─────────────────────────────────────────────────
        for ri, pc in enumerate(projs, start=1):
            lib  = lib_lookup.get(pc, "")
            bg_r = "#ffffff" if ri % 2 == 0 else "#F4F4F4"
            tk.Label(self._frame_jh, text=pc, bg=bg_r, fg="#000000",
                     font=FONT_SMALL, width=12, anchor="w").grid(
                row=ri, column=0, sticky="ew", padx=1, pady=1)
            tk.Label(self._frame_jh, text=lib[:36], bg=bg_r, fg="#000000",
                     font=FONT_SMALL, width=30, anchor="w").grid(
                row=ri, column=1, sticky="ew", padx=1, pady=1)
            row_total = 0.0
            for m in range(1, 13):
                if m < now_m:
                    val  = diana_pm.get(pc, {}).get(m, 0.0)
                    bg_c = C_PAST
                    fg_c = C_PAST_H
                else:
                    val  = raf_pm.get(pc, {}).get(m, 0.0)
                    bg_c = bg_r
                    fg_c = "#000000"
                row_total          += val
                col_totals[m - 1]  += val
                tk.Label(self._frame_jh, text=_fmt(val), bg=bg_c, fg=fg_c,
                         font=FONT_SMALL, width=7, anchor="e").grid(
                    row=ri, column=m+1, sticky="ew", padx=1, pady=1)
            grand_total += row_total
            tk.Label(self._frame_jh, text=_fmt(row_total), bg=bg_r, fg="#000000",
                     font=FONT_SMALL_BOLD, width=11, anchor="e").grid(
                row=ri, column=14, sticky="ew", padx=1, pady=1)

        # ── Ligne Total ────────────────────────────────────────────────────
        tot_row = len(projs) + 1
        C_TOT   = "#F7F7F7"
        tk.Label(self._frame_jh, text="Total", bg=C_TOT, fg="#001a33",
                 font=FONT_SMALL_BOLD, width=12, anchor="e").grid(
            row=tot_row, column=0, columnspan=2, sticky="ew", padx=1, pady=1)
        for m in range(1, 13):
            bg_t = C_PAST if m < now_m else C_TOT
            fg_t = C_PAST_H if m < now_m else "#000000"
            tk.Label(self._frame_jh, text=_fmt(col_totals[m - 1]), bg=bg_t, fg=fg_t,
                     font=FONT_SMALL_BOLD, width=7, anchor="e").grid(
                row=tot_row, column=m+1, sticky="ew", padx=1, pady=1)
        tk.Label(self._frame_jh, text=_fmt(grand_total), bg=C_TOT, fg="#000000",
                 font=FONT_SMALL_BOLD, width=11, anchor="e").grid(
            row=tot_row, column=14, sticky="ew", padx=1, pady=1)

    # ── _build_table_jh_global ───────────────────────────────────────────
    def _build_table_jh_global(self):
        """Cadran bas-gauche mode JH : Consommé / RAF / Atterrissage JH agrégés tous intervenants."""
        for w in self._frame_eur.winfo_children():
            w.destroy()
        for w in self._frame_detail.winfo_children():
            w.destroy()
        self._selected_proj = None

        if not self._cache_ready:
            return

        mois_passes  = list(range(1, self._mois_cur))
        mois_futurs  = list(range(self._mois_cur, 13))
        nucleus_only = self._nucleus_only_var.get()

        # Collecte tous les projets connus
        all_projs = set()
        for prj_dict in self._cache_jh.values():
            all_projs.update(p for p in prj_dict if p not in EXCLUDED_PROJECTS)
        for prj_dict in self._diana_jh.values():
            all_projs.update(p for p in prj_dict if p not in EXCLUDED_PROJECTS)
        raw_projs = list(all_projs)

        lib_lookup = {}
        for prj_dict in self._cache_jh.values():
            for prj, prj_data in prj_dict.items():
                if prj not in lib_lookup and isinstance(prj_data, dict):
                    lib_lookup[prj] = prj_data.get("_lib", "")

        def _fmt_jh(v):
            return str(int(v)) if v and v == int(v) else (f"{v:.1f}" if v else "")

        # ── Pré-calcul Consommé / RAF / Atterrissage ──────────────────────
        jh_vals = {}   # {pc: (conso_jh, raf_jh, atterr_jh)}
        for pc in raw_projs:
            conso_jh = 0.0
            for u_norm, prj_dict in self._diana_jh.items():
                if nucleus_only and u_norm in self._nucleus_exclusions:
                    continue
                if nucleus_only and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM and pc in self._survcomm_projects:
                    continue
                if nucleus_only and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM and pc in self._analytic_projects:
                    continue
                for m, v in prj_dict.get(pc, {}).items():
                    if m in mois_passes:
                        conso_jh += v
            raf_jh = 0.0
            for u_norm, prj_dict in self._cache_jh.items():
                if nucleus_only and u_norm in self._nucleus_exclusions:
                    continue
                if nucleus_only and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM and pc in self._survcomm_projects:
                    continue
                if nucleus_only and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM and pc in self._analytic_projects:
                    continue
                if nucleus_only and pc not in _NUCLEUS_ALLOWED_PROJECTS and u_norm not in self._intervenants_norm:
                    continue
                if pc in prj_dict and isinstance(prj_dict[pc], dict):
                    for m, v in prj_dict[pc].items():
                        if isinstance(m, int) and m in mois_futurs:
                            raf_jh += v
            jh_vals[pc] = (conso_jh, raf_jh, round(conso_jh + raf_jh, 2))

        # ── Tri ───────────────────────────────────────────────────────────
        _jh_g_idx = {"conso": 0, "raf": 1, "atterr": 2}
        def _jh_g_sort_val(pc, col_key):
            if col_key == "code":   return pc
            if col_key == "lib":    return lib_lookup.get(pc, "").lower()
            if col_key == "budget": return self._budgets.get(pc) or 0.0
            idx = _jh_g_idx.get(col_key)
            if idx is not None:     return jh_vals[pc][idx]
            return 0.0
        if "jh_global" not in self._sort_state:
            projs = sorted(raw_projs, key=lambda p: self._budgets.get(p) or 0, reverse=True)
        else:
            projs = self._apply_sort("jh_global", raw_projs, _jh_g_sort_val, default_key="budget")

        # ── En-têtes cliquables ───────────────────────────────────────────
        C_HDR = CA_GREEN
        self._eur_row_widgets = {}
        col_defs_g = [
            ("Project Code",    "code",  12, "w"),
            ("Project Name",    "lib",   34, "w"),
            ("Consommé à date", "conso", 16, "e"),
            ("RAF",             "raf",   16, "e"),
            ("Atterrissage",    "atterr",16, "e"),
            ("Budget",          "budget",14, "e"),
        ]
        for ci, (txt, col_key, w, anc) in enumerate(col_defs_g):
            lbl = tk.Label(self._frame_eur,
                           text=self._sort_indicator("jh_global", col_key, txt),
                           bg=C_HDR, fg="#ffffff", font=FONT_SMALL_BOLD,
                           width=w, anchor=anc, relief="flat", bd=1,
                           cursor="hand2")
            lbl.grid(row=0, column=ci, sticky="ew", padx=1, pady=1)
            lbl.bind("<Button-1>",
                     lambda e, ck=col_key: self._on_sort(
                         "jh_global", ck, self._build_table_jh_global))

        # ── Lignes projets ────────────────────────────────────────────────
        for ri, pc in enumerate(projs, start=1):
            lib = lib_lookup.get(pc, "")
            conso_jh, raf_jh, atterr_jh = jh_vals[pc]
            bg_r = "#ffffff" if ri % 2 == 0 else "#F4F4F4"

            row_data = [
                (pc,                  12, "w", "#000000"),
                (lib[:40],            34, "w", "#000000"),
                (_fmt_jh(conso_jh),   16, "e", "#000000"),
                (_fmt_jh(raf_jh),     16, "e", "#000000"),
                (_fmt_jh(atterr_jh),  16, "e", ACCENT2),
                ("",                  14, "e", "#000000"),   # Budget : vide en mode JH
            ]
            row_lbls = []
            for ci, (txt, w, anc, fg) in enumerate(row_data):
                lbl = tk.Label(self._frame_eur, text=txt, bg=bg_r,
                               fg=fg, font=FONT_SMALL, width=w, anchor=anc,
                               cursor="hand2")
                lbl.grid(row=ri, column=ci, sticky="ew", padx=1, pady=1)
                lbl.bind("<Button-1>",
                         lambda e, p=pc, bg=bg_r: self._on_proj_click(p, bg))
                row_lbls.append(lbl)
            self._eur_row_widgets[pc] = (row_lbls, bg_r)

    # ── _show_proj_detail ────────────────────────────────────────────────
    def _show_proj_detail(self, pc):
        for w in self._frame_detail.winfo_children():
            w.destroy()

        C_HDR = CA_GREEN
        mois_passes_d = list(range(1, self._mois_cur))
        mois_futurs_d = list(range(self._mois_cur, 13))
        nucleus_only  = self._nucleus_only_var.get()
        mode          = self._bot_mode   # "EUR" → "TTNR", "JH" → "JH"

        if mode == "TTNR":
            # ── mode EUR ──────────────────────────────────────────────────
            def _fmt_d(v):
                return f"{round(v):,}".replace(",", "\u202f") if v else ""

            detail = {}
            if self._cache_ready:
                for u_norm, prj_dict in self._cache_jh.items():
                    if pc in prj_dict:
                        if nucleus_only and u_norm in self._nucleus_exclusions:
                            continue
                        if nucleus_only and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM and pc in self._survcomm_projects:
                            continue
                        if nucleus_only and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM and pc in self._analytic_projects:
                            continue
                        u_orig = next(
                            (v for v in (self._cache_usernames or [])
                             if _norm_name(v) == u_norm), u_norm)
                        c_date = sum(
                            self._diana_eur.get(u_norm, {}).get(pc, {}).get(m, 0.0)
                            for m in mois_passes_d)
                        raf_v = sum(
                            v for (un, pr, m), v in self._cache_eur_detail.items()
                            if un == u_norm and pr == pc and m in mois_futurs_d)
                        if nucleus_only and pc not in _NUCLEUS_ALLOWED_PROJECTS and u_norm not in self._intervenants_norm:
                            raf_v = 0.0
                        if c_date or raf_v:
                            detail[u_orig] = (c_date, raf_v)

            # Tri
            _dt_col_map = {"interv": 0, "conso": 0, "raf": 1, "atterr": 2}
            col_dt, dir_dt = self._sort_state.get("detail_ttnr", ("atterr", "desc"))
            rev_dt = (dir_dt == "desc")
            def _dt_key(item):
                u, (cd, rv) = item
                if col_dt == "interv": return _norm_name(u)
                if col_dt == "conso":  return cd
                if col_dt == "raf":    return rv
                return cd + rv   # atterr
            rows_ttnr = sorted(detail.items(), key=_dt_key, reverse=rev_dt)

            # En-têtes cliquables
            _dt_cols = [
                ("interv", "Intervenant",     29, "w"),
                ("conso",  "Consommé à date", 16, "e"),
                ("raf",    "RAF",             14, "e"),
                ("atterr", "Atterrissage",    16, "e"),
            ]
            for ci, (ck, txt, w, anc) in enumerate(_dt_cols):
                lbl = tk.Label(self._frame_detail,
                               text=self._sort_indicator("detail_ttnr", ck, txt),
                               bg=C_HDR, fg="#ffffff", font=FONT_SMALL_BOLD,
                               width=w, anchor=anc, relief="flat", bd=1,
                               cursor="hand2")
                lbl.grid(row=0, column=ci, sticky="ew", padx=1, pady=1)
                lbl.bind("<Button-1>",
                         lambda e, k=ck: self._on_sort(
                             "detail_ttnr", k, lambda: self._show_proj_detail(pc)))

            tot_cd = tot_raf = 0.0
            for ri, (u_orig, (cd, raf_v)) in enumerate(rows_ttnr, start=1):
                bg_r = "#ffffff" if ri % 2 == 0 else "#F4F4F4"
                atterr_v = round(cd + raf_v, 2)
                for ci, (val, w, fg) in enumerate([
                    (u_orig[:29],     29, "#000000"),
                    (_fmt_d(cd),      16, "#000000"),
                    (_fmt_d(raf_v),   14, "#000000"),
                    (_fmt_d(atterr_v),16, ACCENT2),
                ]):
                    tk.Label(self._frame_detail, text=val, bg=bg_r,
                             fg=fg, font=FONT_SMALL, width=w,
                             anchor="w" if ci == 0 else "e").grid(
                        row=ri, column=ci, sticky="ew", padx=1, pady=1)
                tot_cd  += cd
                tot_raf += raf_v

            tot_row    = 1 + len(detail)
            tot_atterr = round(tot_cd + tot_raf, 2)
            for ci, (val, w) in enumerate([
                ("Total",             29), (_fmt_d(tot_cd),     16),
                (_fmt_d(tot_raf),     14), (_fmt_d(tot_atterr), 16)
            ]):
                tk.Label(self._frame_detail, text=val, bg="#F7F7F7", fg="#001a33",
                         font=FONT_SMALL_BOLD, width=w,
                         anchor="e").grid(
                    row=tot_row, column=ci, sticky="ew", padx=1, pady=1)

        else:
            # ── mode JH ───────────────────────────────────────────────────
            def _fmt_jh(v):
                return str(int(v)) if v and v == int(v) else (f"{v:.1f}" if v else "")

            detail_jh = {}
            if self._cache_ready:
                for u_norm, prj_dict in self._cache_jh.items():
                    if pc not in prj_dict:
                        continue
                    if nucleus_only and u_norm in self._nucleus_exclusions:
                        continue
                    if nucleus_only and u_norm in _NUCLEUS_HARDCODED_SURVCOMM_NORM and pc in self._survcomm_projects:
                        continue
                    if nucleus_only and u_norm in _NUCLEUS_HARDCODED_ANALYTIC_NORM and pc in self._analytic_projects:
                        continue
                    u_orig = next(
                        (v for v in (self._cache_usernames or [])
                         if _norm_name(v) == u_norm), u_norm)
                    conso_jh = sum(
                        self._diana_jh.get(u_norm, {}).get(pc, {}).get(m, 0.0)
                        for m in mois_passes_d)
                    raf_jh = 0.0
                    if isinstance(prj_dict[pc], dict):
                        for m, v in prj_dict[pc].items():
                            if isinstance(m, int) and m in mois_futurs_d:
                                raf_jh += v
                    if nucleus_only and pc not in _NUCLEUS_ALLOWED_PROJECTS and u_norm not in self._intervenants_norm:
                        raf_jh = 0.0
                    if conso_jh or raf_jh:
                        detail_jh[u_orig] = (conso_jh, raf_jh)

            # Tri
            col_dj, dir_dj = self._sort_state.get("detail_jh", ("atterr", "desc"))
            rev_dj = (dir_dj == "desc")
            def _dj_key(item):
                u, (cd, rv) = item
                if col_dj == "interv": return _norm_name(u)
                if col_dj == "conso":  return cd
                if col_dj == "raf":    return rv
                return cd + rv   # atterr
            rows_jh = sorted(detail_jh.items(), key=_dj_key, reverse=rev_dj)

            # En-têtes cliquables
            _dj_cols = [
                ("interv", "Intervenant",     29, "w"),
                ("conso",  "Consommé à date", 16, "e"),
                ("raf",    "RAF",             14, "e"),
                ("atterr", "Atterrissage",    16, "e"),
            ]
            for ci, (ck, txt, w, anc) in enumerate(_dj_cols):
                lbl = tk.Label(self._frame_detail,
                               text=self._sort_indicator("detail_jh", ck, txt),
                               bg=C_HDR, fg="#ffffff", font=FONT_SMALL_BOLD,
                               width=w, anchor=anc, relief="flat", bd=1,
                               cursor="hand2")
                lbl.grid(row=0, column=ci, sticky="ew", padx=1, pady=1)
                lbl.bind("<Button-1>",
                         lambda e, k=ck: self._on_sort(
                             "detail_jh", k, lambda: self._show_proj_detail(pc)))

            tot_cd = tot_raf = 0.0
            for ri, (u_orig, (cd, raf_v)) in enumerate(rows_jh, start=1):
                bg_r = "#ffffff" if ri % 2 == 0 else "#F4F4F4"
                atterr_v = round(cd + raf_v, 2)
                for ci, (val, w, fg) in enumerate([
                    (u_orig[:29],      29, "#000000"),
                    (_fmt_jh(cd),      16, "#000000"),
                    (_fmt_jh(raf_v),   14, "#000000"),
                    (_fmt_jh(atterr_v),16, ACCENT2),
                ]):
                    tk.Label(self._frame_detail, text=val, bg=bg_r,
                             fg=fg, font=FONT_SMALL, width=w,
                             anchor="w" if ci == 0 else "e").grid(
                        row=ri, column=ci, sticky="ew", padx=1, pady=1)
                tot_cd  += cd
                tot_raf += raf_v

            tot_row    = 1 + len(detail_jh)
            tot_atterr = round(tot_cd + tot_raf, 2)
            for ci, (val, w) in enumerate([
                ("Total",               29), (_fmt_jh(tot_cd),     16),
                (_fmt_jh(tot_raf),      14), (_fmt_jh(tot_atterr), 16)
            ]):
                tk.Label(self._frame_detail, text=val, bg="#F7F7F7", fg="#001a33",
                         font=FONT_SMALL_BOLD, width=w,
                         anchor="e").grid(
                    row=tot_row, column=ci, sticky="ew", padx=1, pady=1)

    def _on_cell_edit(self, proj_code, mois, var, orig_val):
        raw = var.get().strip().replace(",", ".")
        try:
            new_val = float(raw) if raw else 0.0
        except ValueError:
            var.set(str(int(orig_val)) if orig_val == int(orig_val) else f"{orig_val:.1f}")
            return
        if new_val == orig_val:
            return
        self._modifs[(proj_code, mois)] = (orig_val, new_val)
        self._write_log(proj_code, mois, orig_val, new_val)
        self._status.set(
            f"Modif. enregistrée : {proj_code} / mois {mois:02d} : "
            f"{orig_val} → {new_val} JH")

    def _write_log(self, proj_code, mois, old_val, new_val):
        interv = self._interv_var.get()
        ts     = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line   = (f"{ts} | {interv} | {proj_code} | mois {mois:02d} | "
                  f"{old_val:.2f} JH → {new_val:.2f} JH\n")
        try:
            with open(self._log_path, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception as e:
            self._status.set(f"⚠  Impossible d'écrire le log : {e}")

    def _build_ui(self):
        hdr = tk.Frame(self, bg=CA_GREEN, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📝  MISE À JOUR PLAN DE CHARGES",
                 bg=CA_GREEN, fg="#ffffff", font=FONT_HEAD,
                 padx=16).pack(side="left")
        tk.Label(hdr, text=f"Année {self._annee}",
                 bg=CA_GREEN, fg="#ffffff", font=FONT_SMALL,
                 padx=16).pack(side="right")
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")

        sel_frame = tk.Frame(self, bg=BG_MAIN, pady=6)
        sel_frame.pack(fill="x", padx=10, pady=(6, 2))
        tk.Label(sel_frame, text="Intervenant :", bg=BG_MAIN,
                 fg=TEXT_SEC, font=FONT_SMALL).pack(side="left", padx=(10, 4))
        _cb_style = ttk.Style()
        _cb_style.configure("PdcWhite.TCombobox", fieldbackground="white", background=BG_MAIN)
        _cb_style.map("PdcWhite.TCombobox",
                      fieldbackground=[("readonly", "white")],
                      background=[("readonly", BG_MAIN)])
        self._cb_interv = ttk.Combobox(
            sel_frame, textvariable=self._interv_var,
            state="readonly", width=35, font=FONT_SMALL,
            style="PdcWhite.TCombobox")
        self._cb_interv.pack(side="left")
        self._cb_interv.bind("<<ComboboxSelected>>", self._on_interv_change)
        tk.Label(sel_frame, text="Projet :", bg=BG_MAIN,
                 fg=TEXT_SEC, font=FONT_SMALL).pack(side="left", padx=(14, 4))
        self._cb_proj = ttk.Combobox(
            sel_frame, textvariable=self._proj_filter_var,
            state="readonly", width=40, font=FONT_SMALL,
            style="PdcWhite.TCombobox")
        self._cb_proj["values"] = ["Tous"]
        self._cb_proj.pack(side="left")
        self._cb_proj.bind("<<ComboboxSelected>>", self._on_proj_filter_change)
        tk.Checkbutton(
            sel_frame, text="Périmètre Nucleus uniquement",
            variable=self._nucleus_only_var,
            bg=BG_MAIN, fg=TEXT_SEC, font=FONT_SMALL,
            selectcolor=BG_ENTRY, activebackground=BG_MAIN,
            activeforeground=TEXT_PRI,
            command=self._on_nucleus_filter,
        ).pack(side="right", padx=(0, 12))

        self._status = tk.StringVar(value="Chargement…")
        tk.Label(self, textvariable=self._status, bg=BG_CARD,
                 fg=TEXT_SEC, font=FONT_SMALL, anchor="w",
                 padx=12).pack(fill="x", side="bottom")

        self._selected_proj = None

        body = tk.Frame(self, bg=BG_MAIN)
        body.pack(fill="both", expand=True, padx=4, pady=4)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=2)
        body.rowconfigure(1, weight=3)

        def _make_quad(parent, title, col, autowidth=False, vscroll=False):
            outer = tk.Frame(parent, bg=BG_PANEL,
                             highlightbackground=BORDER, highlightthickness=1)
            outer.grid(row=0, column=col, sticky="nsew", padx=4, pady=4)
            tk.Label(outer, text=title, bg=BG_PANEL, fg=CA_GREEN,
                     font=FONT_BODY, anchor="w").pack(fill="x", padx=6, pady=(4, 2))
            tk.Frame(outer, bg=BORDER, height=1).pack(fill="x")
            cf = tk.Frame(outer, bg=BG_PANEL)
            cf.pack(fill="both", expand=True)
            if vscroll or not autowidth:
                vsb = tk.Scrollbar(cf, orient="vertical")
                vsb.pack(side="right", fill="y")
            c = tk.Canvas(cf, bg="#ffffff", bd=0, highlightthickness=0,
                          yscrollcommand=vsb.set if (vscroll or not autowidth) else None)
            if vscroll or not autowidth:
                vsb.config(command=c.yview)
            c.pack(side="left", fill="both", expand=True)
            inner = tk.Frame(c, bg="#ffffff")
            cw = c.create_window((0, 0), window=inner, anchor="nw")
            if autowidth:
                inner.bind("<Configure>",
                           lambda e, cv=c, w=cw:
                               cv.config(width=e.width, scrollregion=cv.bbox("all")))
            else:
                inner.bind("<Configure>",
                           lambda e, cv=c: cv.configure(scrollregion=cv.bbox("all")))
                c.bind("<Configure>",
                       lambda e, cv=c, w=cw: cv.itemconfig(w, width=e.width))
            if vscroll or not autowidth:
                c.bind("<Enter>",
                       lambda e, cv=c: cv.bind_all(
                           "<MouseWheel>",
                           lambda ev, cv=cv: cv.yview_scroll(
                               -1 if ev.delta > 0 else 1, "units")))
                c.bind("<Leave>",
                       lambda e, cv=c: cv.unbind_all("<MouseWheel>"))
            return inner

        # ── helper : cadran gauche avec toggle JH / TTNR ──────────────────
        def _make_left_pane(parent, col, init_mode, on_toggle, autowidth=False):
            outer = tk.Frame(parent, bg=BG_PANEL,
                             highlightbackground=BORDER, highlightthickness=1)
            outer.grid(row=0, column=col, sticky="nsew", padx=4, pady=4)

            hf = tk.Frame(outer, bg=BG_PANEL)
            hf.pack(fill="x", padx=4, pady=(4, 2))
            lbl_jh = tk.Label(hf, text="Plan de charges JH", font=FONT_BODY,
                              padx=8, pady=2, cursor="hand2", relief="flat")
            lbl_ttnr = tk.Label(hf, text="Plan de charges TTNR", font=FONT_BODY,
                                padx=8, pady=2, cursor="hand2", relief="flat")
            lbl_jh.pack(side="left", padx=(0, 4))
            lbl_ttnr.pack(side="left")

            def _upd(mode):
                if mode == "JH":
                    lbl_jh.config(bg=CA_GREEN, fg="#ffffff")
                    lbl_ttnr.config(bg=BG_PANEL, fg=TEXT_SEC)
                else:
                    lbl_jh.config(bg=BG_PANEL, fg=TEXT_SEC)
                    lbl_ttnr.config(bg=CA_GREEN, fg="#ffffff")
            _upd(init_mode)
            lbl_jh.bind("<Button-1>",   lambda e: on_toggle("JH"))
            lbl_ttnr.bind("<Button-1>", lambda e: on_toggle("TTNR"))

            tk.Frame(outer, bg=BORDER, height=1).pack(fill="x")

            cf = tk.Frame(outer, bg=BG_PANEL)
            cf.pack(fill="both", expand=True)
            vsb = tk.Scrollbar(cf, orient="vertical")
            vsb.pack(side="right", fill="y")
            c = tk.Canvas(cf, bg="#ffffff", bd=0, highlightthickness=0,
                          yscrollcommand=vsb.set)
            vsb.config(command=c.yview)
            c.pack(side="left", fill="both", expand=True)
            inner = tk.Frame(c, bg="#ffffff")
            cw = c.create_window((0, 0), window=inner, anchor="nw")
            if autowidth:
                inner.bind("<Configure>",
                           lambda e, cv=c, w=cw:
                               cv.config(width=e.width, scrollregion=cv.bbox("all")))
            else:
                inner.bind("<Configure>",
                           lambda e, cv=c: cv.configure(scrollregion=cv.bbox("all")))
                c.bind("<Configure>",
                       lambda e, cv=c, w=cw: cv.itemconfig(w, width=e.width))
            c.bind("<Enter>",
                   lambda e, cv=c: cv.bind_all(
                       "<MouseWheel>",
                       lambda ev, cv=cv: cv.yview_scroll(
                           -1 if ev.delta > 0 else 1, "units")))
            c.bind("<Leave>",
                   lambda e, cv=c: cv.unbind_all("<MouseWheel>"))
            return inner, _upd

        # ── callbacks toggle ────────────────────────────────────────────────
        def _on_top_toggle(mode):
            self._top_mode = mode
            if self._update_top_toggle:
                self._update_top_toggle(mode)
            interv = self._interv_var.get()
            if not interv:
                return
            if mode == "JH":
                self._build_table_jh(interv)
            else:
                self._build_table_ttnr_interv(interv)

        def _on_bot_toggle(mode):
            self._bot_mode = mode
            if self._update_bot_toggle:
                self._update_bot_toggle(mode)
            interv = self._interv_var.get()
            if not interv:
                return
            prev = self._selected_proj
            if mode == "TTNR":
                self._build_table_eur(interv)
            else:
                self._build_table_jh_global()
            if prev:
                # Restaurer la sélection : re-highlighter la ligne et réafficher le détail
                self._selected_proj = prev
                if prev in self._eur_row_widgets:
                    lbls, _ = self._eur_row_widgets[prev]
                    for l in lbls:
                        l.config(bg="#c8e6c9")
                self._show_proj_detail(prev)

        # ── row 1 ───────────────────────────────────────────────────────────
        row1 = tk.Frame(body, bg=BG_MAIN)
        row1.grid(row=0, column=0, sticky="nsew")
        row1.rowconfigure(0, weight=1)
        row1.columnconfigure(0, weight=0)
        row1.columnconfigure(1, weight=1)

        self._frame_jh, self._update_top_toggle = _make_left_pane(
            row1, 0, self._top_mode, _on_top_toggle, autowidth=True)
        self._frame_ecart = _make_quad(row1, "Écarts JH / JO", 1)

        # ── row 2 ───────────────────────────────────────────────────────────
        row2 = tk.Frame(body, bg=BG_MAIN)
        row2.grid(row=1, column=0, sticky="nsew")
        row2.rowconfigure(0, weight=1)
        row2.columnconfigure(0, weight=0)
        row2.columnconfigure(1, weight=1)

        self._frame_eur, self._update_bot_toggle = _make_left_pane(
            row2, 0, self._bot_mode, _on_bot_toggle, autowidth=True)

        # Cadran bas-droite : Détail intervenants (simple)
        inner_det = _make_quad(row2, "Détail intervenants", 1)
        self._frame_detail_container = inner_det
        self._frame_detail = tk.Frame(inner_det, bg="#ffffff")
        self._frame_detail.pack(anchor="nw", padx=4, pady=4)


