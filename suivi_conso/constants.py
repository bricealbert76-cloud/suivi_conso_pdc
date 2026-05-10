VERSION = "5.15"

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


# ─────────────────────────────────────────────
#  CONSTANTES MÉTIER
# ─────────────────────────────────────────────
PROJECT_CODE_FIELD = "Project Code"
COL_SURV_COMM      = "Surveillances des comm."
COL_ANALYTIC       = "Analytic"
# Colonnes périmètre dans intervenants.csv (toutes sauf "Intervenant" et "Statut")
PERIMETRE_COLS = [
    "Acquisition", "Distribution", "Exploration", "Data Services",
    "Socle&Enabler", "Core Team", COL_SURV_COMM, COL_ANALYTIC, "IA",
]
TARGET_PROJECTS    = {"P02508", "P06628", "P06629", "P04243"}
EXCLUDED_PROJECTS  = {"MAL", "VAC", "NOC", "PH"}
SEP_INPUT          = "~"
SEP_OUTPUT         = ";"
CHECK_LINES        = 1000

COL_INTERVENANT = "Intervenant"
COL_MOIS        = "Mois"
COL_PROJET      = "PRJ_CodeProjet"
COL_NBJOURS     = "NbJours"

MOIS_ORDRE = {
    "janvier": 1, "février": 2, "fevrier": 2, "mars": 3, "avril": 4,
    "mai": 5, "juin": 6, "juillet": 7, "août": 8, "aout": 8,
    "septembre": 9, "octobre": 10, "novembre": 11,
    "décembre": 12, "decembre": 12,
    1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6,
    7: 7, 8: 8, 9: 9, 10: 10, 11: 11, 12: 12,
}
MOIS_LABELS = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
]

# Colonnes conservées dans le CSV Diana _avec_filtre
# = SRC_COLS (utilisées dans _process_file1) + "Project Name"
DIANA_COLS_FILTRE = [
    "Project Code", "Project Is Alive", "WorkPackage Name", "Task Name",
    "Organization Unit Name", "User First Name", "User Last Name",
    "User Country Name", "User Contract Type", "TimeSheet Mode Name",
    "Timesheet Entry Details Calendar Date", "Status", "Is Published",
    "Timesheet Entry Details Value", "Application Cluster Name",
    "Project Name",
]

# Colonnes conservées dans les CSV PDC _avec_filtre
# Obligatoires + optionnelles (conservées si présentes dans le fichier source)
PDC_COLS_FILTRE_MANDATORY = [COL_INTERVENANT, COL_MOIS, COL_PROJET, COL_NBJOURS]
PDC_COLS_FILTRE_OPTIONAL  = [
    "PRJ_LibelleProjet", "Year of charge", "PRJ_RespNiv1",
    "Code UT", "Type Intervenant",
]


CRAPULL_COLS = ["Intervenant", "Type Intervenant", "Code UT", "Year of charge",
                "Mois", "NbJours", "PRJ_CodeProjet", "PRJ_LibelleProjet"]
FILTER_COLS  = ["Intervenant", "Mois", "PRJ_CodeProjet", "PRJ_LibelleProjet"]

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

MOIS_ABBREV = ["Jan.", "Fév.", "Mar.", "Avr.", "Mai", "Juin",
               "Juil.", "Août", "Sep.", "Oct.", "Nov.", "Déc."]
COL_LIBELLE = "PRJ_LibelleProjet"
COL_OU      = "Code UT"
