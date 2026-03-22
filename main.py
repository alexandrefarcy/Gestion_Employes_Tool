import sys
import json
import os
import subprocess
from datetime import date
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget,
    QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit, QComboBox, QMessageBox, QFileDialog,
    QDialog, QFormLayout, QDialogButtonBox, QFrame,
    QSplitter, QTextEdit, QStatusBar, QDateEdit,
    QGroupBox, QRadioButton, QButtonGroup, QProgressBar,
    QScrollArea, QStyledItemDelegate, QStyle,QCheckBox,QSpinBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QDate
from PyQt6.QtGui import QFont, QColor, QIcon, QPalette

# ── Modules UI découpés ──────────────────────────────
from ui.constantes import (
    COULEURS, COULEURS_CYCLE, COULEUR_TEXTE_CYCLE,
    COULEURS_CYCLE_HYPO, COULEUR_VIDE, COULEUR_NULL, COULEUR_ABSENCE,
    STYLE_GLOBAL, BASE_DIR, CHEMINS,
    EMPLOYES_CONTRATS_JSON, CYCLES_DEFINITIONS_JSON, CYCLES_EMPLOYES_JSON,
    ABSENCES_JSON, EXCEPTIONS_AM_JSON, PLANNING_HISTORIQUE_JSON, IMPORT_HISTORIQUE_JSON,
    charger_json, sauvegarder_json, faire_backup,
    get_chemins, DEPARTEMENTS_LISTE,
)
from ui.onglet_synthese import OngletSynthese, WorkerCalcul
from ui.onglet_absences import OngletAbsences
from ui.onglet_cycles import OngletCyclesDefinitions, OngletCyclesEmployes, DialogueCycleCustom
from ui.onglet_employes import DialogueEmploye, OngletEmployes
from ui.onglet_planning import (
    WorkerImport, DialogueNonReconnus, DialogueDoublon,
    DialogueValidationCycles, WorkerDetectionCycles, WorkerGenerationHyp,
    OngletPlanning,
)
from ui.onglet_export import OngletExport
from ui.onglet_visu import (
    DelegateHachureHypothetique, DialogueOverrideCellule,
    DialogueCorrectionCycle, OngletVisualisationPlanning,
)
from ui.widgets import ComboSansScroll, ChampNom, ChampMatricule, ChampDateMasque







class FenetrePrincipale(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Outil RH Gestion des  Cycles")
        self.setMinimumSize(1100, 700)
        self.resize(1280, 780)

        # Widget central
        central = QWidget()
        self.setCentralWidget(central)
        layout_principal = QVBoxLayout(central)
        layout_principal.setContentsMargins(0, 0, 0, 0)
        layout_principal.setSpacing(0)

        # Bandeau titre
        bandeau = QWidget()
        bandeau.setFixedHeight(56)
        bandeau.setStyleSheet(f"background-color: {COULEURS['bg_carte']}; border-bottom: 1px solid {COULEURS['bordure']};")
        bandeau_layout = QHBoxLayout(bandeau)
        bandeau_layout.setContentsMargins(20, 0, 20, 0)

        titre_app = QLabel("🏭  Outil RH — Cycles & Absences")
        titre_app.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {COULEURS['accent']}; letter-spacing: 0.5px;")
        bandeau_layout.addWidget(titre_app)
        bandeau_layout.addStretch()

        version = QLabel("v1.0.0")
        version.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        bandeau_layout.addWidget(version)

        layout_principal.addWidget(bandeau)

        # --- Bandeau dossier de travail ---
        self.bandeau_dossier = QWidget()
        self.bandeau_dossier.setFixedHeight(44)
        self.bandeau_dossier.setStyleSheet(f"background-color: {COULEURS['bg_secondaire']}; border-bottom: 1px solid {COULEURS['bordure']};")
        dossier_layout = QHBoxLayout(self.bandeau_dossier)
        dossier_layout.setContentsMargins(20, 0, 20, 0)
        dossier_layout.setSpacing(10)

        lbl_dossier = QLabel("📁  Dossier de travail :")
        lbl_dossier.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.lbl_chemin = QLabel(BASE_DIR)
        self.lbl_chemin.setStyleSheet(f"color: {COULEURS['accent']}; font-size: 12px; font-weight: 600;")
        self.lbl_chemin.setMaximumWidth(700)

        btn_changer = QPushButton("📂  Changer de dossier")
        btn_changer.setFixedHeight(28)
        btn_changer.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 5px;
                padding: 4px 12px;
                font-size: 12px;
            }}
            QPushButton:hover {{
                border-color: {COULEURS['accent']};
                color: {COULEURS['accent']};
            }}
        """)
        btn_changer.clicked.connect(self.changer_dossier)

        btn_recharger = QPushButton("🔄  Recharger")
        btn_recharger.setFixedHeight(28)
        btn_recharger.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 5px;
                padding: 4px 12px;
                font-size: 12px;
            }}
            QPushButton:hover {{
                border-color: {COULEURS['accent_succes']};
                color: {COULEURS['accent_succes']};
            }}
        """)
        btn_recharger.clicked.connect(self.recharger_tout)

        dossier_layout.addWidget(lbl_dossier)
        dossier_layout.addWidget(self.lbl_chemin)
        dossier_layout.addStretch()
        dossier_layout.addWidget(btn_changer)
        dossier_layout.addWidget(btn_recharger)
        layout_principal.addWidget(self.bandeau_dossier)

        # Onglets
        self.onglets = QTabWidget()
        self.onglets.setDocumentMode(True)
        layout_principal.addWidget(self.onglets)

        self.onglet_employes     = OngletEmployes()
        self.onglet_cycles_def   = OngletCyclesDefinitions()
        self.onglet_cycles_emp   = OngletCyclesEmployes()
        self.onglet_absences     = OngletAbsences()
        self.onglet_planning     = OngletPlanning()
        self.onglet_visu         = OngletVisualisationPlanning()
        self.onglet_synthese     = OngletSynthese()
        self.onglet_export       = OngletExport()

        self.onglets.addTab(self.onglet_employes,   "👥  Employés")
        self.onglets.addTab(self.onglet_cycles_def, "⚙️  Définitions cycles")
        self.onglets.addTab(self.onglet_cycles_emp, "🔄  Cycles employés")
        self.onglets.addTab(self.onglet_absences,   "📋  Absences")
        self.onglets.addTab(self.onglet_planning,   "📅  Planning")
        self.onglets.addTab(self.onglet_visu,       "🗓  Visualisation")
        self.onglets.addTab(self.onglet_synthese,   "📊  Synthèse")
        self.onglets.addTab(self.onglet_export,     "📤  Export")

        # Barre de statut
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("✅  Application prête  —  Aucune erreurs détectées")

        # Détecter changement d'onglet → proposer sauvegarde si modifications en attente
        self.onglets.currentChanged.connect(self._verifier_sauvegarde_employes)

    def _verifier_sauvegarde_employes(self):
        """Propose la sauvegarde si des modifications non sauvegardées existent dans OngletEmployes."""
        if not self.onglet_employes._modifie:
            return
        rep = QMessageBox.question(
            self,
            "Modifications non sauvegardées",
            "⚠️  Des modifications dans l'onglet Employés n'ont pas été sauvegardées.\n\n"
            "Voulez-vous les sauvegarder maintenant ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if rep == QMessageBox.StandardButton.Yes:
            self.onglet_employes.sauvegarder_donnees()

    def closeEvent(self, event):
        """Propose la sauvegarde avant de fermer si des modifications sont en attente."""
        if self.onglet_employes._modifie:
            rep = QMessageBox.question(
                self,
                "Modifications non sauvegardées",
                "⚠️  Des modifications dans l'onglet Employés n'ont pas été sauvegardées.\n\n"
                "Voulez-vous les sauvegarder avant de quitter ?",
                QMessageBox.StandardButton.Yes
                | QMessageBox.StandardButton.No
                | QMessageBox.StandardButton.Cancel
            )
            if rep == QMessageBox.StandardButton.Cancel:
                event.ignore()
                return
            if rep == QMessageBox.StandardButton.Yes:
                self.onglet_employes.sauvegarder_donnees()
        event.accept()

    def changer_dossier(self):
        global BASE_DIR, CHEMINS
        dossier = QFileDialog.getExistingDirectory(self, "Sélectionner le dossier de travail", BASE_DIR)
        if dossier:
            BASE_DIR = dossier
            CHEMINS  = get_chemins(BASE_DIR)
            self.lbl_chemin.setText(BASE_DIR)
            self.recharger_tout()
            self.status.showMessage(f"✅  Dossier chargé : {BASE_DIR}")

    def recharger_tout(self):
        self.onglet_employes.charger_donnees()
        self.onglet_cycles_def.charger_donnees()
        self.onglet_cycles_emp.charger_donnees()
        self.onglet_planning._maj_statut_planning()
        self.onglet_synthese.verifier_fichiers()
        # Visualisation : ne recharge pas auto (grille lourde — l'utilisateur clique Afficher)
        self.status.showMessage(f"\U0001f504  Donn\xe9es recharg\xe9es depuis : {BASE_DIR}")



# =====================================================
# POINT D'ENTRÉE
# =====================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLE_GLOBAL)
    
    # Chemin robuste (fonctionne en .py et en .exe PyInstaller)
    def resource_path(relative_path):
        base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base, relative_path)
    
    app.setWindowIcon(QIcon(resource_path("icone.ico")))  # ← ajout
    
    fenetre = FenetrePrincipale()
    fenetre.setWindowIcon(QIcon(resource_path("icone.ico")))  # ← ajout
    fenetre.show()
    sys.exit(app.exec())