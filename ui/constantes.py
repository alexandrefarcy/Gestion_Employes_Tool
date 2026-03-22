"""
ui/constantes.py
================
Constantes partagées entre tous les modules UI :
- Chemins des fichiers JSON
- Palette de couleurs
- Style global Qt
- Utilitaires JSON (charger, sauvegarder, backup)
"""

import sys
import os
import json
import shutil
from datetime import datetime

# =====================================================
# CONFIGURATION — Chemins des fichiers JSON
# =====================================================
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # En mode module ui/, remonter d'un niveau par rapport à ui/
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_chemins(base):
    return {
        "employes":          os.path.join(base, "employes_contrats.json"),
        "cycles_def":        os.path.join(base, "cycles_definitions.json"),
        "cycles_emp":        os.path.join(base, "cycles_employes.json"),
        "absences":          os.path.join(base, "absences_projections.json"),
        "exceptions":        os.path.join(base, "am_mensuels_exceptions.json"),
        "planning":          os.path.join(base, "planning_historique.json"),
        "import_historique": os.path.join(base, "import_historique.json"),
    }


CHEMINS = get_chemins(BASE_DIR)


def EMPLOYES_CONTRATS_JSON():   return CHEMINS["employes"]
def CYCLES_DEFINITIONS_JSON():  return CHEMINS["cycles_def"]
def CYCLES_EMPLOYES_JSON():     return CHEMINS["cycles_emp"]
def ABSENCES_JSON():            return CHEMINS["absences"]
def EXCEPTIONS_AM_JSON():       return CHEMINS["exceptions"]
def PLANNING_HISTORIQUE_JSON(): return CHEMINS["planning"]
def IMPORT_HISTORIQUE_JSON():   return CHEMINS["import_historique"]


# =====================================================
# PALETTE DE COULEURS
# =====================================================
COULEURS = {
    "bg_principal":    "#1E1E2E",
    "bg_secondaire":   "#2A2A3E",
    "bg_carte":        "#313147",
    "accent":          "#7C6AF7",
    "accent_hover":    "#9D8FFF",
    "accent_danger":   "#F7736A",
    "accent_succes":   "#6AF7A0",
    "accent_warning":  "#F7C46A",
    "texte":           "#E8E8F5",
    "texte_secondaire":"#9A9AB5",
    "bordure":         "#45455E",
}

COULEURS_CYCLE = {
    "AM": "#4CAF50",
    "M":  "#FF9800",
    "N":  "#9C27B0",
    "J":  "#78909C",
    "WE": "#2196F3",
    "R":  "#E0E0E0",
}

COULEUR_TEXTE_CYCLE = {
    "AM": "#FFFFFF",
    "M":  "#FFFFFF",
    "N":  "#FFFFFF",
    "J":  "#FFFFFF",
    "WE": "#FFFFFF",
    "R":  "#1E1E2E",
}

COULEURS_CYCLE_HYPO = {
    "AM": "#2A6B2E",
    "M":  "#8B5200",
    "N":  "#4A1060",
    "J":  "#3A4A50",
    "WE": "#0D4A7A",
    "R":  "#8A8A8A",
}

COULEUR_VIDE    = "#2A2A3E"
COULEUR_NULL    = "#3D3D55"
COULEUR_ABSENCE = "#4A3F35"

STYLE_GLOBAL = f"""
    QMainWindow, QWidget {{
        background-color: {COULEURS['bg_principal']};
        color: {COULEURS['texte']};
        font-family: 'Segoe UI';
        font-size: 13px;
    }}
    QTabWidget::pane {{
        border: 1px solid {COULEURS['bordure']};
        border-radius: 8px;
        background-color: {COULEURS['bg_secondaire']};
    }}
    QTabBar::tab {{
        background-color: {COULEURS['bg_secondaire']};
        color: {COULEURS['texte_secondaire']};
        padding: 10px 22px;
        border: none;
        font-size: 13px;
        font-weight: 500;
    }}
    QTabBar::tab:selected {{
        background-color: {COULEURS['bg_carte']};
        color: {COULEURS['accent']};
        border-bottom: 2px solid {COULEURS['accent']};
        font-weight: 600;
    }}
    QTabBar::tab:hover {{
        color: {COULEURS['texte']};
        background-color: {COULEURS['bg_carte']};
    }}
    QPushButton {{
        background-color: {COULEURS['accent']};
        color: white;
        border: none;
        border-radius: 6px;
        padding: 8px 18px;
        font-weight: 600;
        font-size: 13px;
    }}
    QPushButton:hover {{
        background-color: {COULEURS['accent_hover']};
    }}
    QPushButton:pressed {{
        background-color: {COULEURS['accent']};
    }}
    QPushButton.danger {{
        background-color: {COULEURS['accent_danger']};
    }}
    QPushButton.succes {{
        background-color: {COULEURS['accent_succes']};
        color: #1E1E2E;
    }}
    QTableWidget {{
        background-color: {COULEURS['bg_secondaire']};
        border: 1px solid {COULEURS['bordure']};
        border-radius: 8px;
        gridline-color: {COULEURS['bordure']};
        color: {COULEURS['texte']};
        selection-background-color: {COULEURS['accent']};
    }}
    QTableWidget::item {{
        padding: 6px 10px;
    }}
    QHeaderView::section {{
        background-color: {COULEURS['bg_carte']};
        color: {COULEURS['accent']};
        padding: 8px 10px;
        border: none;
        border-bottom: 2px solid {COULEURS['accent']};
        font-weight: 600;
        font-size: 12px;
        text-transform: uppercase;
    }}
    QLineEdit, QComboBox {{
        background-color: {COULEURS['bg_carte']};
        border: 1px solid {COULEURS['bordure']};
        border-radius: 6px;
        padding: 7px 12px;
        color: {COULEURS['texte']};
        font-size: 13px;
    }}
    QLineEdit:focus, QComboBox:focus {{
        border: 1px solid {COULEURS['accent']};
    }}
    QComboBox::drop-down {{
        border: none;
        padding-right: 8px;
    }}
    QComboBox QAbstractItemView {{
        background-color: {COULEURS['bg_carte']};
        border: 1px solid {COULEURS['bordure']};
        color: {COULEURS['texte']};
        selection-background-color: {COULEURS['accent']};
    }}
    QTextEdit {{
        background-color: {COULEURS['bg_carte']};
        border: 1px solid {COULEURS['bordure']};
        border-radius: 8px;
        color: {COULEURS['texte']};
        font-family: 'Consolas', monospace;
        font-size: 12px;
        padding: 8px;
    }}
    QStatusBar {{
        background-color: {COULEURS['bg_carte']};
        color: {COULEURS['texte_secondaire']};
        font-size: 12px;
        padding: 4px 10px;
    }}
    QScrollBar:vertical {{
        background: {COULEURS['bg_secondaire']};
        width: 8px;
        border-radius: 4px;
    }}
    QScrollBar::handle:vertical {{
        background: {COULEURS['bordure']};
        border-radius: 4px;
        min-height: 20px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {COULEURS['accent']};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0px;
    }}
    QFrame[frameShape="4"], QFrame[frameShape="5"] {{
        color: {COULEURS['bordure']};
    }}
    QRadioButton {{
        color: {COULEURS['texte']};
        font-size: 12px;
    }}
    QRadioButton::indicator {{
        width: 14px;
        height: 14px;
        border-radius: 7px;
        border: 2px solid {COULEURS['texte']};
        background-color: transparent;
    }}
    QRadioButton::indicator:checked {{
        background-color: white;
        border: 2px solid white;
    }}
    QRadioButton::indicator:unchecked {{
        background-color: transparent;
    }}
"""


# =====================================================
# UTILITAIRES JSON
# =====================================================
def charger_json(chemin):
    """Charge un fichier JSON, retourne un dict vide si inexistant."""
    if not os.path.exists(chemin):
        return {}
    with open(chemin, 'r', encoding='utf-8') as f:
        return json.load(f)


def sauvegarder_json(chemin, data):
    """Sauvegarde un dict dans un fichier JSON."""
    with open(chemin, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def faire_backup(chemin_json, callback_log=None):
    """
    Crée un backup du fichier JSON dans un sous-dossier 'backup/'.
    Retourne le chemin du backup créé, ou None si le fichier source n'existe pas.
    """
    if not os.path.exists(chemin_json):
        return None
    dossier_backup = os.path.join(os.path.dirname(chemin_json), "backup")
    os.makedirs(dossier_backup, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    nom = os.path.splitext(os.path.basename(chemin_json))[0]
    chemin_backup = os.path.join(dossier_backup, f"{nom}_{ts}.json")
    shutil.copy2(chemin_json, chemin_backup)
    if callback_log:
        callback_log(f"💾 Backup : {chemin_backup}")
    return chemin_backup


# =====================================================
# CONSTANTES MÉTIER
# =====================================================
DEPARTEMENTS_LISTE = [
    "", "Fabrication", "Conditionnement", "Administratif",
    "Maintenance", "Technique", "Qualité", "Informatique", "Intérim"
]