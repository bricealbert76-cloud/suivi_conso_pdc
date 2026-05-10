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
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

from .constants import *
from .theme import *

__all__ = [
    # flags dépendances
    "HAS_OPENPYXL", "HAS_PANDAS",
    # fonctions utilitaires
    "_norm_name", "_mois_effectif", "_calc_jours_ouvres",
    "_load_allowed_intervenants", "_load_nucleus_exclusions",
    # constantes nucleus
    "_NUCLEUS_HARDCODED_SURVCOMM_ORIG", "_NUCLEUS_HARDCODED_SURVCOMM_NORM",
    "_NUCLEUS_HARDCODED_ANALYTIC_ORIG", "_NUCLEUS_HARDCODED_ANALYTIC_NORM",
    "_NUCLEUS_HARDCODED_NORM", "_NUCLEUS_ALLOWED_PROJECTS",
    # widget helper
    "make_button",
]

def _mois_effectif():
    """Retourne le mois 'effectif' pour l'application.
    Si on est dans les 3 derniers jours du mois (j >= nb_jours - 3),
    on anticipe : le mois suivant est considéré comme mois courant.
    En décembre (mois 12), on reste sur 12 pour ne pas sortir du périmètre annuel.
    """
    today = datetime.now()
    nb_jours = calendar.monthrange(today.year, today.month)[1]
    if today.day >= nb_jours - 3 and today.month < 12:
        return today.month + 1
    return today.month


def _calc_jours_ouvres(annee):
    """Calcule les jours ouvrés par mois pour l'année donnée (France métropolitaine).
    Utilise uniquement la bibliothèque standard Python (pas de dépendance externe).
    Retourne une liste de 12 entiers (janvier → décembre).
    """
    import calendar
    from datetime import date, timedelta

    # Calcul de Pâques — algorithme grégorien anonyme
    a = annee % 19
    b = annee // 100
    c = annee % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mois_p = (h + l - 7 * m + 114) // 31
    jour_p = ((h + l - 7 * m + 114) % 31) + 1
    paques = date(annee, mois_p, jour_p)

    feries = {
        date(annee, 1,  1),               # Jour de l'an
        paques + timedelta(days=1),        # Lundi de Pâques
        date(annee, 5,  1),               # Fête du travail
        date(annee, 5,  8),               # Victoire 1945
        paques + timedelta(days=39),       # Ascension
        paques + timedelta(days=50),       # Lundi de Pentecôte
        date(annee, 7,  14),              # Fête nationale
        date(annee, 8,  15),              # Assomption
        date(annee, 11, 1),               # Toussaint
        date(annee, 11, 11),              # Armistice
        date(annee, 12, 25),              # Noël
    }

    result = []
    for mois in range(1, 13):
        nb = sum(
            1 for jour in range(1, calendar.monthrange(annee, mois)[1] + 1)
            if date(annee, mois, jour).weekday() < 5        # lundi–vendredi
            and date(annee, mois, jour) not in feries
        )
        result.append(nb)
    return result


def _norm_name(s):
    """Normalise un nom pour la jointure : supprime les accents, met en majuscules,
    remplace tirets et toutes variantes d'apostrophes par des espaces."""
    if not s:
        return ""
    # Décomposition unicode → suppression des diacritiques
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = s.upper()
    # Remplacer TOUTES les variantes d'apostrophes/guillemets simples et tirets
    # par un espace (U+0027, U+2019, U+2018, U+02BC, U+2032, U+055A, U+2010, etc.)
    _APOSTROPHES = (
        "\u0027",  # apostrophe ASCII '
        "\u2019",  # apostrophe typographique '
        "\u2018",  # apostrophe réversée '
        "\u02bc",  # apostrophe modificateur ʼ
        "\u2032",  # prime ′
        "\u055a",  # apostrophe arménienne ՚
        "\u0060",  # accent grave `
        "\u00b4",  # accent aigu ´
        "\u2010",  # trait d'union sans espace ‐
        "\u2011",  # trait d'union insécable ‑
        "-",        # tiret standard
    )
    for apo in _APOSTROPHES:
        s = s.replace(apo, " ")
    # Compresser les espaces multiples
    return " ".join(s.split())


# Intervenants multiprojets avec exclusion partielle SurvComm (filtre Nucleus en dur)
_NUCLEUS_HARDCODED_SURVCOMM_ORIG = ["GUITTARD, Raphael", "ABDELLI, Ahmed"]
_NUCLEUS_HARDCODED_SURVCOMM_NORM = frozenset(_norm_name(u) for u in _NUCLEUS_HARDCODED_SURVCOMM_ORIG)
# Intervenants multiprojets avec exclusion partielle Analytic (filtre Nucleus en dur)
_NUCLEUS_HARDCODED_ANALYTIC_ORIG = ["GUITTARD, Raphael", "TCHIEGAING KAMGUIA, Arnaud"]
_NUCLEUS_HARDCODED_ANALYTIC_NORM = frozenset(_norm_name(u) for u in _NUCLEUS_HARDCODED_ANALYTIC_ORIG)
# Union des deux pour les vérifications communes
_NUCLEUS_HARDCODED_NORM = _NUCLEUS_HARDCODED_SURVCOMM_NORM | _NUCLEUS_HARDCODED_ANALYTIC_NORM
# Projets autorisés en RAF (PDC EUR) pour les non-intervenants (filtre Nucleus)
_NUCLEUS_ALLOWED_PROJECTS = frozenset({"P02508", "P04243"})


def _load_allowed_intervenants(path):
    """Charge le fichier intervenants (path) et retourne (set_norms, list_originals)
    pour tous les intervenants non marqués dans la colonne COL_SURV_COMM."""
    allowed_norms = set()
    allowed_list  = []
    if not path or not os.path.isfile(path):
        return allowed_norms, allowed_list
    try:
        with open(path, encoding="utf-8-sig", errors="replace") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                if row.get(COL_SURV_COMM, "").strip() == "1":
                    continue
                name = row.get("Intervenant", "").strip()
                if not name:
                    continue
                allowed_list.append(name)
                allowed_norms.add(_norm_name(name))
        allowed_list.sort()
    except Exception:
        pass
    return allowed_norms, allowed_list


def _load_nucleus_exclusions(path):
    """Retourne un set de noms normalisés des intervenants à exclure du périmètre Nucleus.
    Sont exclus ceux dont la SEULE colonne périmètre cochée est COL_SURV_COMM
    ou la SEULE colonne périmètre cochée est COL_ANALYTIC.
    """
    exclusions = set()
    if not path or not os.path.isfile(path):
        return exclusions
    try:
        with open(path, encoding="utf-8-sig", errors="replace") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                name = row.get("Intervenant", "").strip()
                if not name:
                    continue
                active = {col for col in PERIMETRE_COLS
                          if row.get(col, "").strip() == "1"}
                if active == {COL_SURV_COMM} or active == {COL_ANALYTIC}:
                    exclusions.add(_norm_name(name))
    except Exception:
        pass
    return exclusions


def make_button(parent, text, command, style="neutral", width=None):
    if style == "ca":
        bg, fg, hov = CA_GREEN_HOV, "#ffffff", CA_GREEN
    elif style == "action":
        bg, fg, hov = BTN_ACT, "#ffffff", BTN_ACT_H
    else:
        bg, fg, hov = BTN_NEU, TEXT_PRI, BTN_NEU_H
    kw  = dict(text=text, command=command, bg=bg, fg=fg,
               font=FONT_HEAD, bd=0, padx=14, pady=8,
               relief="flat", cursor="hand2",
               activebackground=hov, activeforeground="#ffffff")
    if width:
        kw["width"] = width
    btn = tk.Button(parent, **kw)
    btn.bind("<Enter>", lambda e, c=hov: btn.config(bg=c))
    btn.bind("<Leave>", lambda e, c=bg:  btn.config(bg=c))
    return btn
