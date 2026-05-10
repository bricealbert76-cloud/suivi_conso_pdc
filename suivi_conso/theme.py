from .constants import VERSION

# ─────────────────────────────────────────────
#  THÈMES  (sélection via --style dark|light|blue)
# ─────────────────────────────────────────────
THEMES = {
    "dark": {
        "BG_MAIN"  : "#0d1117", "BG_PANEL" : "#161b22",
        "BG_CARD"  : "#1c2128", "BG_ENTRY" : "#21262d",
        "ACCENT"   : "#2ea043", "ACCENT2"  : "#388bfd",
        "WARN"     : "#f0883e", "TEXT_PRI" : "#e6edf3",
        "TEXT_SEC" : "#8b949e", "BORDER"   : "#30363d",
        "BTN_ACT"  : "#238636", "BTN_ACT_H": "#2ea043",
        "BTN_NEU"  : "#21262d", "BTN_NEU_H": "#30363d",
        "LBL_HEAD" : "#388bfd",
    },
    "light": {                                           # palette CA-CIB
        "BG_MAIN"  : "#f5f7f6", "BG_PANEL" : "#ffffff",
        "BG_CARD"  : "#edf3f0", "BG_ENTRY" : "#ffffff",
        "ACCENT"   : "#007d57", "ACCENT2"  : "#116f6a",
        "WARN"     : "#d4540a", "TEXT_PRI" : "#001a33",
        "TEXT_SEC" : "#5a6878", "BORDER"   : "#c5d4ce",
        "BTN_ACT"  : "#007d57", "BTN_ACT_H": "#005c40",
        "BTN_NEU"  : "#e8eeeb", "BTN_NEU_H": "#cddbd4",
        "LBL_HEAD" : "#000000",
    },
    "blue": {
        "BG_MAIN"  : "#0f1923", "BG_PANEL" : "#162032",
        "BG_CARD"  : "#1c2d42", "BG_ENTRY" : "#1e3a5f",
        "ACCENT"   : "#00b4d8", "ACCENT2"  : "#48cae4",
        "WARN"     : "#f77f00", "TEXT_PRI" : "#e0f2fe",
        "TEXT_SEC" : "#90c4e4", "BORDER"   : "#2a4a6b",
        "BTN_ACT"  : "#0077b6", "BTN_ACT_H": "#00b4d8",
        "BTN_NEU"  : "#1e3a5f", "BTN_NEU_H": "#2a4a6b",
        "LBL_HEAD" : "#48cae4",
    },
}

CA_GREEN     = "#007461"   # vert bouton Crédit Agricole (brand primary)
CA_GREEN_HOV = "#075144"   # vert hover Crédit Agricole

def _apply_theme(name):
    theme = THEMES.get(name, THEMES["dark"])
    globals().update(theme)

import argparse as _argparse
_parser = _argparse.ArgumentParser(add_help=False)
_parser.add_argument("--style", choices=["dark", "light", "blue"], default="light")
_args, _ = _parser.parse_known_args()
_apply_theme(_args.style)

FONT_TITLE     = ("Montserrat", 18, "bold")
FONT_HEAD      = ("Montserrat", 11, "bold")
FONT_BODY      = ("Open Sans", 10)
FONT_BODY_BOLD = ("Open Sans", 10, "bold")
FONT_SMALL     = ("Open Sans", 9)
FONT_SMALL_BOLD = ("Open Sans", 9, "bold")
FONT_MONO      = ("Courier New", 9)
FONT_TAB_UNSEL = ("Open Sans", 9)

