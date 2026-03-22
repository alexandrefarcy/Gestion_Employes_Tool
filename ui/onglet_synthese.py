"""
ui/onglet_synthese.py
=====================
OngletSynthese + WorkerCalcul
"""

import os
import subprocess
from datetime import date

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QGroupBox, QRadioButton, QButtonGroup,
    QTextEdit, QFrame, QDateEdit, QMessageBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QDate

from ui.constantes import (
    COULEURS, BASE_DIR,
    EMPLOYES_CONTRATS_JSON, CYCLES_DEFINITIONS_JSON, CYCLES_EMPLOYES_JSON,
    ABSENCES_JSON, EXCEPTIONS_AM_JSON,
)


# =====================================================
# WORKER THREAD — calcul en arrière-plan
# =====================================================
class WorkerCalcul(QThread):
    log_signal    = pyqtSignal(str)
    fini_signal   = pyqtSignal(str)
    erreur_signal = pyqtSignal(str)

    def __init__(self, mode: str, params: dict):
        super().__init__()
        self.mode   = mode
        self.params = params

    def run(self):
        try:
            if self.mode == 'mode1':
                from mode1_legacy import lancer_mode1
                chemin = lancer_mode1(
                    chemin_excel      = self.params['excel'],
                    chemin_contrats   = self.params['contrats'],
                    chemin_absences   = self.params['absences'],
                    chemin_exceptions = self.params['exceptions'],
                    chemin_sortie     = self.params['sortie'],
                    date_retro_depuis = self.params['date_retro'],
                    fin_projection    = self.params['date_fin'],
                    callback_log      = lambda m: self.log_signal.emit(m),
                )
            else:
                from mode2_cycles import lancer_mode2
                chemin = lancer_mode2(
                    chemin_cycles_employes = self.params['cycles_emp'],
                    chemin_cycles_def      = self.params['cycles_def'],
                    chemin_contrats        = self.params['contrats'],
                    chemin_absences        = self.params['absences'],
                    chemin_exceptions      = self.params['exceptions'],
                    chemin_sortie          = self.params['sortie'],
                    date_debut_global      = self.params['date_retro'],
                    fin_projection         = self.params['date_fin'],
                    callback_log           = lambda m: self.log_signal.emit(m),
                )
            self.fini_signal.emit(chemin)
        except Exception as e:
            self.erreur_signal.emit(str(e))


# =====================================================
# ONGLET SYNTHÈSE
# =====================================================
class OngletSynthese(QWidget):
    def __init__(self):
        super().__init__()
        self._worker = None
        self._construire_ui()

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(16, 16, 16, 16)

        titre = QLabel("\U0001f4ca  Génération de la synthèse AM")
        titre.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {COULEURS['texte']};")
        layout.addWidget(titre)

        # ---- Mode de calcul ----
        grp_mode = QGroupBox("Mode de calcul")
        grp_mode.setStyleSheet(f"""
            QGroupBox {{
                color: {COULEURS['accent']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 8px;
                margin-top: 10px;
                font-weight: 600;
                font-size: 13px;
                padding: 10px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }}
        """)
        grp_mode_layout = QVBoxLayout(grp_mode)

        self.radio_mode1 = QRadioButton(
            "\u2699\ufe0f  Mode 1 — Lecture depuis les feuilles Excel  (temporaire, reproduit le VBA)"
        )
        self.radio_mode2 = QRadioButton(
            "\U0001f504  Mode 2 — Cycles depuis les JSONs  (définitif, à activer une fois les cycles remplis)"
        )
        self.radio_mode1.setChecked(True)
        for r in [self.radio_mode1, self.radio_mode2]:
            r.setStyleSheet(f"color: {COULEURS['texte']}; font-size: 13px; padding: 4px 0;")
            grp_mode_layout.addWidget(r)

        self.btn_grp = QButtonGroup(self)
        self.btn_grp.addButton(self.radio_mode1, 1)
        self.btn_grp.addButton(self.radio_mode2, 2)
        self.btn_grp.buttonClicked.connect(self._on_mode_change)
        layout.addWidget(grp_mode)

        # ---- Panneau mode 1 ----
        self.panel_mode1 = QWidget()
        p1_layout = QVBoxLayout(self.panel_mode1)
        p1_layout.setContentsMargins(0, 0, 0, 0)
        p1_layout.setSpacing(8)
        lbl_excel = QLabel("Classeur Excel contenant les feuilles S41_, S42_... :")
        lbl_excel.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        p1_layout.addWidget(lbl_excel)
        from PyQt6.QtWidgets import QLineEdit
        row_excel = QHBoxLayout()
        self.champ_excel = QLineEdit()
        self.champ_excel.setPlaceholderText("Chemin vers le fichier .xlsm / .xlsx ...")
        self.champ_excel.setReadOnly(True)
        btn_parcourir_excel = QPushButton("\U0001f4c2  Parcourir")
        btn_parcourir_excel.setFixedWidth(120)
        btn_parcourir_excel.setStyleSheet(
            f"background-color: {COULEURS['bg_carte']}; color: {COULEURS['texte']};"
            f"border: 1px solid {COULEURS['bordure']}; border-radius: 5px; padding: 6px 10px;"
        )
        btn_parcourir_excel.clicked.connect(self._parcourir_excel)
        row_excel.addWidget(self.champ_excel)
        row_excel.addWidget(btn_parcourir_excel)
        p1_layout.addLayout(row_excel)
        layout.addWidget(self.panel_mode1)

        # ---- Panneau mode 2 ----
        self.panel_mode2 = QWidget()
        p2_layout = QVBoxLayout(self.panel_mode2)
        p2_layout.setContentsMargins(0, 0, 0, 0)
        p2_layout.setSpacing(4)
        self.lbl_cycles_emp_mode2 = QLabel()
        self.lbl_cycles_def_mode2 = QLabel()
        for lbl in [self.lbl_cycles_emp_mode2, self.lbl_cycles_def_mode2]:
            lbl.setStyleSheet("font-size: 13px; padding: 2px 0;")
            p2_layout.addWidget(lbl)
        self.panel_mode2.setVisible(False)
        layout.addWidget(self.panel_mode2)

        # ---- Paramètres ----
        grp_params = QGroupBox("Paramètres de projection")
        grp_params.setStyleSheet(grp_mode.styleSheet())
        grp_params_layout = QHBoxLayout(grp_params)
        grp_params_layout.setSpacing(20)
        lbl_retro = QLabel("Depuis (rétroactif) :")
        lbl_retro.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.date_retro = QDateEdit()
        self.date_retro.setDate(QDate(2024, 6, 2))
        self.date_retro.setDisplayFormat("dd/MM/yyyy")
        self.date_retro.setCalendarPopup(True)
        self.date_retro.setFixedWidth(130)
        lbl_fin = QLabel("Jusqu'au (projection) :")
        lbl_fin.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.date_fin = QDateEdit()
        self.date_fin.setDate(QDate(2025, 12, 31))
        self.date_fin.setDisplayFormat("dd/MM/yyyy")
        self.date_fin.setCalendarPopup(True)
        self.date_fin.setFixedWidth(130)
        grp_params_layout.addWidget(lbl_retro)
        grp_params_layout.addWidget(self.date_retro)
        grp_params_layout.addSpacing(20)
        grp_params_layout.addWidget(lbl_fin)
        grp_params_layout.addWidget(self.date_fin)
        grp_params_layout.addStretch()
        layout.addWidget(grp_params)

        # ---- Statut fichiers ----
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)
        row_statut = QHBoxLayout()
        lbl_statut = QLabel("Fichiers requis :")
        lbl_statut.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-weight: 600; font-size: 12px;")
        btn_refresh = QPushButton("\U0001f504  Vérifier")
        btn_refresh.setFixedWidth(100)
        btn_refresh.setStyleSheet(
            f"background-color: {COULEURS['bg_carte']}; color: {COULEURS['texte']};"
            f"border: 1px solid {COULEURS['bordure']}; border-radius: 5px; padding: 4px 8px; font-size: 12px;"
        )
        btn_refresh.clicked.connect(self.verifier_fichiers)
        row_statut.addWidget(lbl_statut)
        row_statut.addStretch()
        row_statut.addWidget(btn_refresh)
        layout.addLayout(row_statut)

        grp_fichiers = QGroupBox()
        grp_fichiers.setStyleSheet(
            f"QGroupBox {{ border: 1px solid {COULEURS['bordure']}; border-radius: 6px; padding: 8px; }}"
        )
        grp_fichiers_layout = QVBoxLayout(grp_fichiers)
        grp_fichiers_layout.setSpacing(2)
        self.lbl_employes   = QLabel()
        self.lbl_absences   = QLabel()
        self.lbl_exceptions = QLabel()
        for lbl in [self.lbl_employes, self.lbl_absences, self.lbl_exceptions]:
            lbl.setStyleSheet("font-size: 12px; padding: 2px 0;")
            grp_fichiers_layout.addWidget(lbl)
        layout.addWidget(grp_fichiers)

        # ---- Console ----
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setMinimumHeight(140)
        self.console.setMaximumHeight(200)
        self.console.setPlaceholderText("Les logs du calcul s'afficheront ici...")
        layout.addWidget(self.console)

        # ---- Bouton lancer ----
        self.btn_lancer = QPushButton("\u25b6\ufe0f  Lancer la synthèse")
        self.btn_lancer.setMinimumHeight(48)
        self.btn_lancer.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['accent']};
                font-size: 15px;
                font-weight: 700;
                border-radius: 8px;
            }}
            QPushButton:hover {{ background-color: {COULEURS['accent_hover']}; }}
            QPushButton:disabled {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte_secondaire']};
            }}
        """)
        self.btn_lancer.clicked.connect(self.lancer_synthese)
        layout.addWidget(self.btn_lancer)

        self.verifier_fichiers()

    def _on_mode_change(self):
        is_mode1 = self.radio_mode1.isChecked()
        self.panel_mode1.setVisible(is_mode1)
        self.panel_mode2.setVisible(not is_mode1)
        self.verifier_fichiers()

    def _parcourir_excel(self):
        chemin, _ = QFileDialog.getOpenFileName(
            self, "Sélectionner le classeur Excel",
            BASE_DIR, "Fichiers Excel (*.xlsx *.xlsm *.xls)"
        )
        if chemin:
            self.champ_excel.setText(chemin)

    def verifier_fichiers(self):
        def statut(lbl, chemin, nom):
            if os.path.exists(chemin):
                lbl.setText(f"  \u2705  {nom}")
                lbl.setStyleSheet(f"color: {COULEURS['accent_succes']}; font-size: 12px;")
            else:
                lbl.setText(f"  \u274c  {nom}  — manquant")
                lbl.setStyleSheet(f"color: {COULEURS['accent_danger']}; font-size: 12px;")
        statut(self.lbl_employes,   EMPLOYES_CONTRATS_JSON(),  "employes_contrats.json")
        statut(self.lbl_absences,   ABSENCES_JSON(),           "absences_projections.json")
        statut(self.lbl_exceptions, EXCEPTIONS_AM_JSON(),      "am_mensuels_exceptions.json")
        if not self.radio_mode1.isChecked():
            statut(self.lbl_cycles_emp_mode2, CYCLES_EMPLOYES_JSON(),    "cycles_employes.json")
            statut(self.lbl_cycles_def_mode2, CYCLES_DEFINITIONS_JSON(), "cycles_definitions.json")

    def lancer_synthese(self):
        mode = 'mode1' if self.radio_mode1.isChecked() else 'mode2'
        if mode == 'mode1':
            chemin_excel = self.champ_excel.text().strip()
            if not chemin_excel or not os.path.exists(chemin_excel):
                QMessageBox.warning(self, "Fichier manquant",
                    "Veuillez sélectionner le classeur Excel contenant les feuilles S41_, S42_...")
                return
        else:
            if not os.path.exists(CYCLES_EMPLOYES_JSON()):
                QMessageBox.warning(self, "Fichier manquant",
                    "cycles_employes.json introuvable.\n"
                    "Remplissez les cycles dans l'onglet 'Cycles employés' avant de lancer.")
                return

        chemin_sortie, _ = QFileDialog.getSaveFileName(
            self, "Enregistrer la synthèse AM",
            os.path.join(BASE_DIR, "Synthese_AM.xlsx"),
            "Fichiers Excel (*.xlsx)"
        )
        if not chemin_sortie:
            return

        qd_retro = self.date_retro.date()
        qd_fin   = self.date_fin.date()
        d_retro  = date(qd_retro.year(), qd_retro.month(), qd_retro.day())
        d_fin    = date(qd_fin.year(),   qd_fin.month(),   qd_fin.day())

        params = {
            'contrats':   EMPLOYES_CONTRATS_JSON(),
            'absences':   ABSENCES_JSON(),
            'exceptions': EXCEPTIONS_AM_JSON(),
            'sortie':     chemin_sortie,
            'date_retro': d_retro,
            'date_fin':   d_fin,
        }
        if mode == 'mode1':
            params['excel'] = chemin_excel
        else:
            params['cycles_emp'] = CYCLES_EMPLOYES_JSON()
            params['cycles_def'] = CYCLES_DEFINITIONS_JSON()

        self.console.clear()
        self.btn_lancer.setEnabled(False)
        self.btn_lancer.setText("\u23f3  Calcul en cours...")
        self._worker = WorkerCalcul(mode, params)
        self._worker.log_signal.connect(self._on_log)
        self._worker.fini_signal.connect(self._on_fini)
        self._worker.erreur_signal.connect(self._on_erreur)
        self._worker.start()

    def _on_log(self, msg: str):
        self.console.append(msg)

    def _on_fini(self, chemin: str):
        self.btn_lancer.setEnabled(True)
        self.btn_lancer.setText("\u25b6\ufe0f  Lancer la synthèse")
        self.console.append(f"\n\u2705  Synthèse générée avec succès !")
        self.console.append(f"\U0001f4c4  {chemin}")
        rep = QMessageBox.question(
            self, "Synthèse générée",
            f"\u2705  Fichier généré :\n{chemin}\n\nOuvrir le fichier maintenant ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if rep == QMessageBox.StandardButton.Yes:
            if os.name == 'nt':
                os.startfile(chemin)
            else:
                subprocess.Popen(['xdg-open', chemin])

    def _on_erreur(self, msg: str):
        self.btn_lancer.setEnabled(True)
        self.btn_lancer.setText("\u25b6\ufe0f  Lancer la synthèse")
        self.console.append(f"\n\u274c  ERREUR : {msg}")
        QMessageBox.critical(self, "Erreur lors du calcul", msg)