"""
ui/onglet_planning.py
=====================
WorkerImport + DialogueNonReconnus + DialogueDoublon
+ DialogueValidationCycles + WorkerDetectionCycles
+ WorkerGenerationHyp + OngletPlanning
"""

import os
import re as _re

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit, QDialog, QDialogButtonBox,
    QTextEdit, QFrame, QFileDialog, QMessageBox,
    QCheckBox, QGroupBox, QDateEdit, QProgressBar,
    QScrollArea,QTabWidget, QComboBox,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QDate
from PyQt6.QtGui import QColor, QFont

from ui.constantes import (
    COULEURS, COULEURS_CYCLE,COULEUR_TEXTE_CYCLE,
    PLANNING_HISTORIQUE_JSON, CYCLES_EMPLOYES_JSON,
    IMPORT_HISTORIQUE_JSON, EMPLOYES_CONTRATS_JSON,
    CYCLES_DEFINITIONS_JSON, ABSENCES_JSON,
    BASE_DIR,
    charger_json, sauvegarder_json, faire_backup,
)
from ui.widgets import ComboSansScroll, ChampDateMasque
from ui.onglet_employes import DialogueEmploye

class WorkerImport(QThread):
    log_signal    = pyqtSignal(str)
    fini_signal   = pyqtSignal(dict)   # {"importes": int, "non_reconnus": list, "doublons": list}
    erreur_signal = pyqtSignal(str)
    doublon_signal = pyqtSignal(str, str, dict, dict)  # cle_emp, cle_j, existant, nouveau

    def __init__(self, source: str, chemin_fichier: str):
        super().__init__()
        self.source         = source   # 'excel_fab' | 'pdf_adp' | 'excel_adp'
        self.chemin_fichier = chemin_fichier
        self._doublons_en_attente = []   # liste de (cle_emp, cle_j, existant, nouveau)
        self._decisions_doublons  = {}   # {(cle_emp, cle_j): True/False}  — rempli par l'UI

    def run(self):
        try:
            from importer_planning import (
                importer_excel_annuel,
                importer_pdf_adp,
                importer_excel_adp_hebdo,
            )

            non_reconnus = []
            doublons_detectes = []

            def on_non_reconnu(nom_brut, *args):
                # args pour Source 1&2 : (source_str, annee_ou_mois)
                # args pour Source 3   : (source_str, num_semaine, id_norm)
                cle_ctx   = "_".join(str(a) for a in args[:2]) if args else ""
                id_source = str(args[2]) if len(args) > 2 else ""
                non_reconnus.append({
                    "nom_brut":  nom_brut,
                    "cle":       cle_ctx,
                    "id_source": id_source,
                })
                self.log_signal.emit(f"  ❓ Non reconnu : {nom_brut}")

            def on_doublon(cle_emp, cle_j, existant, nouveau):
                doublons_detectes.append({
                    "cle_emp": cle_emp,
                    "cle_j":   cle_j,
                    "existant": existant,
                    "nouveau":  nouveau,
                })
                self.log_signal.emit(
                    f"  ⚠️ Doublon : {cle_emp.split('|')[0]} — {cle_j}  "
                    f"({existant.get('source','?')}:{existant.get('cycle','?')} "
                    f"→ {nouveau.get('source','?')}:{nouveau.get('cycle','?')})"
                )
                return False   # ne pas écraser automatiquement

            self.log_signal.emit(f"⏳ Démarrage de l'import ({self.source})…")

            if self.source == 'excel_fab':
                nb = importer_excel_annuel(
                    chemin_xlsx         = self.chemin_fichier,
                    chemin_planning     = PLANNING_HISTORIQUE_JSON(),
                    chemin_employes     = EMPLOYES_CONTRATS_JSON(),
                    callback_log        = lambda m: self.log_signal.emit(m),
                    callback_non_reconnu= on_non_reconnu,
                )
            elif self.source == 'pdf_adp':
                nb = importer_pdf_adp(
                    chemin_pdf          = self.chemin_fichier,
                    chemin_planning     = PLANNING_HISTORIQUE_JSON(),
                    chemin_employes     = EMPLOYES_CONTRATS_JSON(),
                    callback_log        = lambda m: self.log_signal.emit(m),
                    callback_non_reconnu= on_non_reconnu,
                    callback_doublon    = on_doublon,
                )
            else:  # excel_adp
                nb = importer_excel_adp_hebdo(
                    chemin_xlsx        = self.chemin_fichier,
                    chemin_planning     = PLANNING_HISTORIQUE_JSON(),
                    chemin_employes     = EMPLOYES_CONTRATS_JSON(),
                    callback_log        = lambda m: self.log_signal.emit(m),
                    callback_non_reconnu= on_non_reconnu,
                    callback_doublon    = on_doublon,
                )

            # nb est le dict résumé retourné par l'importeur
            # On extrait le bon compteur selon la source
            if isinstance(nb, dict):
                nb_importes = nb.get("employes_importes", nb.get("semaines_ecrites", 0))
            else:
                nb_importes = nb or 0

            self.fini_signal.emit({
                "importes":     nb_importes,
                "non_reconnus": non_reconnus,
                "doublons":     doublons_detectes,
                "details":      nb if isinstance(nb, dict) else {},
            })

        except Exception as e:
            import traceback
            self.erreur_signal.emit(f"{e}\n\n{traceback.format_exc()}")


# =====================================================
# DIALOGUE — Non reconnus après import
# =====================================================
class DialogueNonReconnus(QDialog):
    """
    Dialogue affiché après import quand des noms n'ont pas pu être matchés.

    Pour chaque nom non reconnu, l'utilisateur peut :
      - Ignorer (défaut)
      - Associer à un employé existant
      - Créer un nouvel employé (ouvre DialogueEmploye en sous-dialogue)

    get_decisions() → {nom_brut: "ignorer" | cle_existante | {"creer": {...}}}
    get_nouveaux_employes() → liste des dicts employés à créer
    """

    def __init__(self, non_reconnus: list, employes_connus: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Noms non reconnus lors de l'import")
        self.setMinimumSize(760, 500)
        self.setStyleSheet(f"""
            QDialog {{ background-color: {COULEURS['bg_secondaire']}; color: {COULEURS['texte']}; }}
            QLabel  {{ color: {COULEURS['texte']}; }}
        """)

        self.employes_connus = employes_connus
        self.non_reconnus    = non_reconnus
        # decisions: {nom_brut: "ignorer" | "cle_emp_choisie" | {"creer": {cle, date, sortie, dept}}}
        self.decisions = {nr["nom_brut"]: "ignorer" for nr in non_reconnus}
        # stocke les données des nouveaux employés à créer, indexées par nom_brut
        self._nouveaux = {}

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        info = QLabel(
            f"<b>{len(non_reconnus)} nom(s)</b> n'ont pas été reconnus automatiquement.<br>"
            "Pour chacun, choisissez d'associer à un employé existant, "
            "<b>de créer un nouvel employé</b>, ou d'ignorer."
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        layout.addWidget(info)

        # Tableau
        self.tableau = QTableWidget()
        self.tableau.setColumnCount(3)
        self.tableau.setHorizontalHeaderLabels(["Nom brut dans le fichier", "Clé / Période", "Action"])
        self.tableau.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tableau.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.tableau.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableau.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableau.verticalHeader().setVisible(False)
        self.tableau.verticalHeader().setDefaultSectionSize(50)
        self.tableau.setAlternatingRowColors(True)
        self.tableau.setStyleSheet(
            f"QTableWidget {{ alternate-background-color: {COULEURS['bg_carte']}; }}"
        )

        # Liste employés connus pour les combos
        self._liste_employes = sorted(
            [c for c in employes_connus if "|" in c],
            key=lambda c: c.split("|")[0]
        )
        # Options fixes : ignorer (0), créer (1), puis employés existants (2+)
        self._OPT_IGNORER = "— Ignorer —"
        self._OPT_CREER   = "➕  Créer cet employé..."
        self._options_fixes = [self._OPT_IGNORER, self._OPT_CREER]
        self._options_combo = self._options_fixes + [
            c.split("|")[0].strip() for c in self._liste_employes
        ]

        self.tableau.setRowCount(len(non_reconnus))
        self.tableau.verticalHeader().setDefaultSectionSize(36)   # hauteur suffisante pour la QComboBox
        self._combos  = []
        self._labels_crees = []   # QLabel "✅ Créé" qui remplace la combo quand créé

        for row, nr in enumerate(non_reconnus):
            self.tableau.setItem(row, 0, QTableWidgetItem(nr["nom_brut"]))
            self.tableau.setItem(row, 1, QTableWidgetItem(nr.get("cle", "")))

            combo = QComboBox()
            combo.addItems(self._options_combo)
            combo.currentIndexChanged.connect(lambda idx, r=row: self._on_combo_change(r, idx))
            self.tableau.setCellWidget(row, 2, combo)
            self._combos.append(combo)
            self._labels_crees.append(None)

        layout.addWidget(self.tableau)

        boutons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        boutons.button(QDialogButtonBox.StandardButton.Ok).setText("✅  Valider et fermer")
        boutons.accepted.connect(self.accept)
        layout.addWidget(boutons)

    # ------------------------------------------------------------------
    def _on_combo_change(self, row: int, idx: int):
        nom_brut = self.non_reconnus[row]["nom_brut"]

        if idx == 0:
            # Ignorer
            self.decisions[nom_brut] = "ignorer"
            self._nouveaux.pop(nom_brut, None)

        elif idx == 1:
            # Créer — ouvrir DialogueEmploye en sous-dialogue
            self._ouvrir_creation(row, nom_brut)

        else:
            # Associer à un employé existant
            self.decisions[nom_brut] = self._liste_employes[idx - len(self._options_fixes)]
            self._nouveaux.pop(nom_brut, None)

    def _ouvrir_creation(self, row: int, nom_brut: str):
        """Ouvre DialogueEmploye pré-rempli pour créer un nouvel employé."""
        nr = self.non_reconnus[row]

        # Construire le prefill : nom normalisé + ID si présent dans le contexte
        prefill = {
            "nom": nom_brut,
            "id":  nr.get("id_source", ""),   # fourni par Source 3 via WorkerImport
            "titre_fenetre": f"Créer l'employé — {nom_brut}",
        }

        dlg = DialogueEmploye(parent=self, prefill=prefill)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            # Annulé → retour à "Ignorer"
            self._combos[row].blockSignals(True)
            self._combos[row].setCurrentIndex(0)
            self._combos[row].blockSignals(False)
            self.decisions[nom_brut] = "ignorer"
            self._nouveaux.pop(nom_brut, None)
            return

        cle, date_debut, date_sortie, dept, archive = dlg.get_donnees()

        if not cle or "|" not in cle:
            QMessageBox.warning(
                self, "Données incomplètes",
                "Le nom ou le matricule est manquant. Employé non créé."
            )
            self._combos[row].blockSignals(True)
            self._combos[row].setCurrentIndex(0)
            self._combos[row].blockSignals(False)
            self.decisions[nom_brut] = "ignorer"
            return

        # Déduire le statut : archive > sortie > actif
        if archive and not date_sortie:
            actif = False
        elif date_sortie:
            actif = False
        else:
            actif = True

        # Stocker la décision
        self._nouveaux[nom_brut] = {
            "cle":         cle,
            "date_debut":  date_debut,
            "date_sortie": date_sortie,
            "actif":       actif,
            "departement": dept,
        }
        self.decisions[nom_brut] = {"creer": cle}

        # Remplacer la combo par un label "✅ Créé" (read-only)
        nom_affiche = cle.split("|")[0].strip()
        lbl = QLabel(f"  ✅  Créé : <b>{nom_affiche}</b>")
        lbl.setStyleSheet(
            f"color: {COULEURS['accent_succes']}; font-size: 12px; padding: 4px 8px;"
        )
        self.tableau.setCellWidget(row, 2, lbl)
        self._labels_crees[row] = lbl

    # ------------------------------------------------------------------
    def get_decisions(self):
        return self.decisions

    def get_nouveaux_employes(self):
        """Retourne la liste des nouveaux employés à créer."""
        return list(self._nouveaux.values())


# =====================================================
# DIALOGUE — Conflit doublon
# =====================================================
class DialogueDoublon(QDialog):
    def __init__(self, doublons: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Conflits détectés — doublons")
        self.setMinimumSize(700, 400)
        self.setStyleSheet(f"""
            QDialog {{ background-color: {COULEURS['bg_secondaire']}; color: {COULEURS['texte']}; }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        info = QLabel(
            f"<b>{len(doublons)} conflit(s)</b> détecté(s) : des jours déjà importés ont un cycle différent.<br>"
            "Cochez les lignes à <b>écraser</b> avec la nouvelle valeur. Les non cochées sont conservées."
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        layout.addWidget(info)

        self.tableau = QTableWidget()
        self.tableau.setColumnCount(5)
        self.tableau.setHorizontalHeaderLabels([
            "Employé", "Jour / Semaine", "Valeur existante", "Nouvelle valeur", "Écraser ?"
        ])
        self.tableau.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tableau.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableau.verticalHeader().setVisible(False)
        self.tableau.setAlternatingRowColors(True)
        self.tableau.setStyleSheet(f"QTableWidget {{ alternate-background-color: {COULEURS['bg_carte']}; }}")

        self.tableau.setRowCount(len(doublons))
        self._checkboxes = []
        for row, d in enumerate(doublons):
            nom = d["cle_emp"].split("|")[0].strip()
            existant = d["existant"]
            nouveau  = d["nouveau"]
            self.tableau.setItem(row, 0, QTableWidgetItem(nom))
            self.tableau.setItem(row, 1, QTableWidgetItem(d["cle_j"]))

            item_ex = QTableWidgetItem(
                f"{existant.get('cycle','?')}  [{existant.get('source','?')}]"
            )
            item_ex.setForeground(QColor(COULEURS['accent_succes']))
            self.tableau.setItem(row, 2, item_ex)

            item_nv = QTableWidgetItem(
                f"{nouveau.get('cycle','?')}  [{nouveau.get('source','?')}]"
            )
            item_nv.setForeground(QColor(COULEURS['accent_warning']))
            self.tableau.setItem(row, 3, item_nv)

            chk_item = QTableWidgetItem()
            chk_item.setCheckState(Qt.CheckState.Unchecked)
            chk_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tableau.setItem(row, 4, chk_item)
            self._checkboxes.append(chk_item)

        layout.addWidget(self.tableau)

        # Boutons tout cocher / tout décocher
        row_sel = QHBoxLayout()
        btn_tout = QPushButton("✅  Tout écraser")
        btn_tout.setStyleSheet(f"background-color: {COULEURS['accent_warning']}; color: #1E1E2E;")
        btn_tout.clicked.connect(self._tout_cocher)
        btn_rien = QPushButton("❌  Tout conserver")
        btn_rien.setStyleSheet(f"background-color: {COULEURS['bg_carte']}; color: {COULEURS['texte']}; border: 1px solid {COULEURS['bordure']};")
        btn_rien.clicked.connect(self._tout_decocher)
        row_sel.addWidget(btn_tout)
        row_sel.addWidget(btn_rien)
        row_sel.addStretch()
        layout.addLayout(row_sel)

        boutons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        boutons.button(QDialogButtonBox.StandardButton.Ok).setText("💾  Valider")
        boutons.accepted.connect(self.accept)
        layout.addWidget(boutons)

    def _tout_cocher(self):
        for chk in self._checkboxes:
            chk.setCheckState(Qt.CheckState.Checked)

    def _tout_decocher(self):
        for chk in self._checkboxes:
            chk.setCheckState(Qt.CheckState.Unchecked)

    def get_decisions(self):
        """Retourne la liste des indices de doublons à écraser."""
        return [i for i, chk in enumerate(self._checkboxes)
                if chk.checkState() == Qt.CheckState.Checked]


# =====================================================
# WORKER THREAD — Détection des cycles en arrière-plan
# =====================================================
# =====================================================
# DIALOGUE VALIDATION CYCLES DETECTES
# =====================================================
class DialogueValidationCycles(QDialog):
    """
    Popup unique — liste tous les cycles détectés avec checkboxes.
    Les conflits (cycle actuel != cycle détecté) sont signalés en orange.
    """

    def __init__(self, resultats: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("\U0001f50d  Validation des cycles détectés")
        self.setMinimumSize(780, 500)
        self.resize(900, 560)
        self._resultats = resultats
        self._checkboxes = []
        self._construire_ui()

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        nb_total    = len(self._resultats)
        nb_conflits = sum(1 for r in self._resultats if r['conflit'])
        titre_txt = f"\U0001f50d  {nb_total} cycle(s) détecté(s) à valider"
        if nb_conflits:
            titre_txt += f"  \u2014  \u26a0\ufe0f {nb_conflits} conflit(s) avec saisie manuelle"
        titre = QLabel(titre_txt)
        titre.setStyleSheet(
            f"font-size: 14px; font-weight: 700; color: {COULEURS['texte']}; padding-bottom: 4px;"
        )
        layout.addWidget(titre)

        self.tableau = QTableWidget()
        self.tableau.setColumnCount(6)
        self.tableau.setHorizontalHeaderLabels([
            "\u2714", "Employé", "Cycle actuel", "Cycle détecté", "Date départ", "Confiance"
        ])
        self.tableau.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableau.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableau.verticalHeader().setVisible(False)
        self.tableau.setAlternatingRowColors(True)
        self.tableau.setStyleSheet(f"""
            QTableWidget {{
                background-color: {COULEURS['bg_principal']};
                gridline-color: {COULEURS['bordure']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 6px;
            }}
            QHeaderView::section {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte_secondaire']};
                border: none;
                border-right: 1px solid {COULEURS['bordure']};
                border-bottom: 1px solid {COULEURS['bordure']};
                padding: 4px 6px;
                font-size: 11px;
                font-weight: 600;
            }}
        """)

        self.tableau.setRowCount(len(self._resultats))
        for row, r in enumerate(self._resultats):
            est_conflit = r['conflit']
            bg = QColor("#3D1F00") if est_conflit else QColor(COULEURS['bg_principal'])

            cb = QCheckBox()
            cb.setChecked(not est_conflit)
            cb.setToolTip(
                "\u26a0\ufe0f Conflit — cochez pour écraser le cycle manuel existant"
                if est_conflit else "Appliquer ce cycle"
            )
            self._checkboxes.append(cb)
            self.tableau.setCellWidget(row, 0, cb)

            item_nom = QTableWidgetItem(r['nom'])
            item_nom.setForeground(QColor(COULEURS['texte']))
            if est_conflit:
                f = QFont(); f.setBold(True); item_nom.setFont(f)
                item_nom.setBackground(bg)
            self.tableau.setItem(row, 1, item_nom)

            item_act = QTableWidgetItem(r['cycle_actuel'] or "\u2014")
            item_act.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if est_conflit:
                item_act.setForeground(QColor("#FF8A65"))
                item_act.setBackground(bg)
            else:
                item_act.setForeground(QColor(COULEURS['texte_secondaire']))
            self.tableau.setItem(row, 2, item_act)

            item_det = QTableWidgetItem(r['cycle_depart'])
            item_det.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_det.setBackground(QColor(COULEURS_CYCLE.get(r['cycle_depart'], COULEURS['bg_carte'])))
            item_det.setForeground(QColor(COULEUR_TEXTE_CYCLE.get(r['cycle_depart'], '#FFFFFF')))
            if est_conflit:
                f2 = QFont(); f2.setBold(True); item_det.setFont(f2)
            self.tableau.setItem(row, 3, item_det)

            item_date = QTableWidgetItem(r['date_depart'] or "\u2014")
            item_date.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item_date.setForeground(QColor(COULEURS['texte_secondaire']))
            if est_conflit:
                item_date.setBackground(bg)
            self.tableau.setItem(row, 4, item_date)

            pct = int(r['score'] * 100)
            item_score = QTableWidgetItem(f"{pct}%")
            item_score.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            couleur_score = (
                COULEURS['accent_succes'] if pct >= 85 else
                COULEURS.get('avertissement', '#FF9800') if pct >= 70 else
                COULEURS['accent_danger']
            )
            item_score.setForeground(QColor(couleur_score))
            if est_conflit:
                item_score.setBackground(bg)
            self.tableau.setItem(row, 5, item_score)

        self.tableau.setColumnWidth(0, 36)
        hh = self.tableau.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        for c in range(2, 6):
            hh.setSectionResizeMode(c, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.tableau)

        if nb_conflits:
            lbl_info = QLabel(
                "\u26a0\ufe0f Les lignes en orange ont un cycle saisi manuellement différent "
                "du cycle détecté. Elles sont décochées par défaut — cochez pour écraser."
            )
            lbl_info.setStyleSheet("color: #FF8A65; font-size: 11px; font-style: italic;")
            lbl_info.setWordWrap(True)
            layout.addWidget(lbl_info)

        row_checks = QHBoxLayout()
        btn_tout = QPushButton("Tout cocher")
        btn_tout.clicked.connect(lambda: [cb.setChecked(True) for cb in self._checkboxes])
        btn_aucun = QPushButton("Tout décocher")
        btn_aucun.clicked.connect(lambda: [cb.setChecked(False) for cb in self._checkboxes])
        row_checks.addWidget(btn_tout)
        row_checks.addWidget(btn_aucun)
        row_checks.addStretch()
        layout.addLayout(row_checks)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.button(QDialogButtonBox.StandardButton.Ok).setText("\u2705  Appliquer les cycles cochés")
        btns.button(QDialogButtonBox.StandardButton.Cancel).setText("Annuler")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_resultats_valides(self) -> list:
        return [r for r, cb in zip(self._resultats, self._checkboxes) if cb.isChecked()]


class WorkerDetectionCycles(QThread):
    log_signal    = pyqtSignal(str)
    fini_signal   = pyqtSignal(list)   # liste de dicts par employé (analyser_sans_ecrire)
    erreur_signal = pyqtSignal(str)

    def __init__(self, ecraser_manuel: bool = False):
        super().__init__()
        self.ecraser_manuel = ecraser_manuel

    def run(self):
        try:
            from detecter_cycles import analyser_sans_ecrire
            resultats = analyser_sans_ecrire(
                chemin_planning        = PLANNING_HISTORIQUE_JSON(),
                chemin_cycles_employes = CYCLES_EMPLOYES_JSON(),
                ecraser_manuel         = self.ecraser_manuel,
                callback_log           = lambda msg: self.log_signal.emit(msg),
            )
            self.fini_signal.emit(resultats)
        except Exception as e:
            import traceback
            self.erreur_signal.emit(f"{e}\n\n{traceback.format_exc()}")


# =====================================================
# WORKER GÉNÉRATION HYPOTHÉTIQUES (Hyp-B)
# =====================================================
class WorkerGenerationHyp(QThread):
    log_signal    = pyqtSignal(str)
    fini_signal   = pyqtSignal(dict)   # dict de stats
    erreur_signal = pyqtSignal(str)

    def __init__(self, date_debut, date_fin):
        super().__init__()
        self.date_debut = date_debut   # date Python
        self.date_fin   = date_fin     # date Python

    def run(self):
        try:
            from generer_hypothetiques import generer_hypothetiques
            stats = generer_hypothetiques(
                chemin_planning = PLANNING_HISTORIQUE_JSON(),
                chemin_cycles   = CYCLES_EMPLOYES_JSON(),
                chemin_employes = EMPLOYES_CONTRATS_JSON(),
                date_debut_gen  = self.date_debut,
                date_fin_gen    = self.date_fin,
                callback_log    = lambda msg: self.log_signal.emit(msg),
            )
            self.fini_signal.emit(stats)
        except Exception as e:
            import traceback
            self.erreur_signal.emit(f"{e}\n\n{traceback.format_exc()}")


# =====================================================
# ONGLET PLANNING
# =====================================================
class OngletPlanning(QWidget):
    def __init__(self):
        super().__init__()
        self._worker        = None
        self._worker_cycles = None
        self._derniers_doublons    = []
        self._derniers_non_reconnus = []
        self._construire_ui()

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # ---- Titre ----
        titre = QLabel("📅  Import du planning historique")
        titre.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {COULEURS['texte']};")
        layout.addWidget(titre)

        sous_titre = QLabel(
            "Importez les données de planning depuis les différentes sources. "
            "Les données réelles écrasent les hypothétiques. "
            "Un backup est créé automatiquement avant chaque import."
        )
        sous_titre.setWordWrap(True)
        sous_titre.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        layout.addWidget(sous_titre)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(sep)

        # =====================================================
        # SOUS-ONGLETS
        # =====================================================
        self._sous_onglets = QTabWidget()
        self._sous_onglets.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {COULEURS['bordure']};
                border-radius: 6px;
                background-color: {COULEURS['bg_secondaire']};
            }}
            QTabBar::tab {{
                background-color: {COULEURS['bg_secondaire']};
                color: {COULEURS['texte_secondaire']};
                padding: 7px 18px;
                border: none;
                font-size: 12px;
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
        """)
        layout.addWidget(self._sous_onglets)

        self._construire_onglet_import()
        self._construire_onglet_historique()

    # --------------------------------------------------
    # Sous-onglet 1 — Import
    # --------------------------------------------------
    def _construire_onglet_import(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(12)
        layout.setContentsMargins(14, 14, 14, 14)

        # ---- Statut fichier ----
        row_fichier = QHBoxLayout()
        lbl_f = QLabel("Fichier planning :")
        lbl_f.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px; font-weight: 600;")
        self.lbl_planning_statut = QLabel()
        self._maj_statut_planning()
        row_fichier.addWidget(lbl_f)
        row_fichier.addWidget(self.lbl_planning_statut)
        row_fichier.addStretch()
        layout.addLayout(row_fichier)

        # =====================================================
        # GROUPE — Sources d'import
        # =====================================================
        _STYLE_GRP = f"""
            QGroupBox {{
                color: {COULEURS['accent']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 8px;
                margin-top: 10px;
                font-weight: 600;
                font-size: 13px;
                padding: 14px 10px 10px 10px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }}
        """
        grp_import = QGroupBox("Sources d'import")
        grp_import.setStyleSheet(_STYLE_GRP)
        grp_layout = QVBoxLayout(grp_import)
        grp_layout.setSpacing(10)

        fmt_style = f"color: {COULEURS['texte_secondaire']}; font-size: 11px; font-style: italic;"

        # Source 1 — Excel annuel Fabrication
        col1 = QVBoxLayout(); col1.setSpacing(2)
        lbl1 = QLabel("📂  <b>Excel annuel Fabrication</b>")
        lbl1.setTextFormat(Qt.TextFormat.RichText)
        lbl1.setStyleSheet(f"color: {COULEURS['texte']}; font-size: 13px;")
        lbl1_fmt = QLabel("Format attendu : 2021.xlsx  |  2022.xlsm  (annee.xlsx/xlsm)")
        lbl1_fmt.setStyleSheet(fmt_style)
        col1.addWidget(lbl1); col1.addWidget(lbl1_fmt)
        row1 = QHBoxLayout()
        self.btn_import_excel_fab = QPushButton("Sélectionner et importer")
        self.btn_import_excel_fab.setFixedWidth(220)
        self.btn_import_excel_fab.clicked.connect(lambda: self._lancer_import('excel_fab'))
        row1.addLayout(col1); row1.addStretch(); row1.addWidget(self.btn_import_excel_fab)
        grp_layout.addLayout(row1)

        sep1 = QFrame(); sep1.setFrameShape(QFrame.Shape.HLine)
        grp_layout.addWidget(sep1)

        # Source 2 — PDF ADP mensuel
        col2 = QVBoxLayout(); col2.setSpacing(2)
        lbl2 = QLabel("📄  <b>PDF ADP mensuel Fabrication</b>")
        lbl2.setTextFormat(Qt.TextFormat.RichText)
        lbl2.setStyleSheet(f"color: {COULEURS['texte']}; font-size: 13px;")
        lbl2_fmt = QLabel("Format attendu : 01_2025.pdf  |  08_2025.pdf  (MM_AAAA.pdf)")
        lbl2_fmt.setStyleSheet(fmt_style)
        col2.addWidget(lbl2); col2.addWidget(lbl2_fmt)
        row2 = QHBoxLayout()
        self.btn_import_pdf_adp = QPushButton("Sélectionner et importer")
        self.btn_import_pdf_adp.setFixedWidth(220)
        self.btn_import_pdf_adp.clicked.connect(lambda: self._lancer_import('pdf_adp'))
        row2.addLayout(col2); row2.addStretch(); row2.addWidget(self.btn_import_pdf_adp)
        grp_layout.addLayout(row2)

        sep2 = QFrame(); sep2.setFrameShape(QFrame.Shape.HLine)
        grp_layout.addWidget(sep2)

        # Source 3 — Excel ADP hebdomadaire
        col3 = QVBoxLayout(); col3.setSpacing(2)
        lbl3 = QLabel("📊  <b>Excel ADP hebdomadaire</b>  (tous départements)")
        lbl3.setTextFormat(Qt.TextFormat.RichText)
        lbl3.setStyleSheet(f"color: {COULEURS['texte']}; font-size: 13px;")
        lbl3_fmt = QLabel("Format attendu : S41_2025.xlsx  |  S41_2025.xls  (SNN_AAAA.xlsx/xls)")
        lbl3_fmt.setStyleSheet(fmt_style)
        col3.addWidget(lbl3); col3.addWidget(lbl3_fmt)
        row3 = QHBoxLayout()
        self.btn_import_excel_adp = QPushButton("Sélectionner et importer")
        self.btn_import_excel_adp.setFixedWidth(220)
        self.btn_import_excel_adp.clicked.connect(lambda: self._lancer_import('excel_adp'))
        row3.addLayout(col3); row3.addStretch(); row3.addWidget(self.btn_import_excel_adp)
        grp_layout.addLayout(row3)

        layout.addWidget(grp_import)

        # Console de log
        lbl_console_titre = QLabel("Journal d'import :")
        lbl_console_titre.setStyleSheet(
            f"color: {COULEURS['texte_secondaire']}; font-weight: 600; font-size: 12px;"
        )
        layout.addWidget(lbl_console_titre)
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        self.console.setPlaceholderText("Les logs d'import s'afficheront ici…")
        layout.addWidget(self.console)
        row_bas = QHBoxLayout()
        btn_effacer = QPushButton("🗑  Effacer les logs")
        btn_effacer.setStyleSheet(
            f"background-color: {COULEURS['bg_carte']}; color: {COULEURS['texte_secondaire']};"
            f"border: 1px solid {COULEURS['bordure']}; font-size: 12px;"
        )
        btn_effacer.clicked.connect(self.console.clear)
        row_bas.addStretch()
        row_bas.addWidget(btn_effacer)
        layout.addLayout(row_bas)

        self._sous_onglets.addTab(page, "\U0001f4e5  Import")

    # --------------------------------------------------
    # Sous-onglet 2 — Historique & Cycles
    # --------------------------------------------------
    def _construire_onglet_historique(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(14)
        layout.setContentsMargins(14, 14, 14, 14)

        _STYLE_GRP = f"""
            QGroupBox {{
                color: {COULEURS['accent']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 8px;
                margin-top: 10px;
                font-weight: 600;
                font-size: 13px;
                padding: 14px 10px 10px 10px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }}
        """

        # =====================================================
        # GROUPE — Historique des imports
        # =====================================================
        grp_hist = QGroupBox("Historique des imports")
        grp_hist.setStyleSheet(_STYLE_GRP)
        grp_hist_layout = QVBoxLayout(grp_hist)
        grp_hist_layout.setSpacing(8)

        self.tableau_historique = QTableWidget()
        self.tableau_historique.setColumnCount(4)
        self.tableau_historique.setHorizontalHeaderLabels(["Fichier", "Date / Heure", "Source", "Entrées"])
        self.tableau_historique.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tableau_historique.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau_historique.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau_historique.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau_historique.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableau_historique.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableau_historique.verticalHeader().setVisible(False)
        self.tableau_historique.setAlternatingRowColors(True)
        self.tableau_historique.setStyleSheet(f"""
            QTableWidget {{ alternate-background-color: {COULEURS['bg_carte']}; }}
        """)
        grp_hist_layout.addWidget(self.tableau_historique)

        row_hist_bas = QHBoxLayout()
        self.lbl_hist_info = QLabel("Sélectionnez un import pour le supprimer et permettre le ré-import.")
        self.lbl_hist_info.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 11px; font-style: italic;")

        btn_suppr_import = QPushButton("🗑  Supprimer cet import")
        btn_suppr_import.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['accent_danger']};
                color: #1E1E2E;
                border-radius: 5px;
                padding: 6px 14px;
                font-weight: 700;
                font-size: 12px;
            }}
            QPushButton:hover {{ background-color: #FF9A95; }}
            QPushButton:disabled {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte_secondaire']};
            }}
        """)
        btn_suppr_import.clicked.connect(self._supprimer_import_selectionne)

        btn_refresh_hist = QPushButton("🔄  Rafraîchir")
        btn_refresh_hist.setStyleSheet(
            f"background-color: {COULEURS['bg_carte']}; color: {COULEURS['texte_secondaire']};"
            f"border: 1px solid {COULEURS['bordure']}; border-radius: 5px; padding: 6px 12px; font-size: 12px;"
        )
        btn_refresh_hist.clicked.connect(self._charger_historique_imports)

        row_hist_bas.addWidget(self.lbl_hist_info)
        row_hist_bas.addStretch()
        row_hist_bas.addWidget(btn_refresh_hist)
        row_hist_bas.addWidget(btn_suppr_import)
        grp_hist_layout.addLayout(row_hist_bas)

        layout.addWidget(grp_hist)

        # =====================================================
        # GROUPE — Détection des cycles
        # =====================================================
        grp_cycles = QGroupBox("Détection automatique des cycles")
        grp_cycles.setStyleSheet(f"""
            QGroupBox {{
                color: {COULEURS['accent_succes']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 8px;
                margin-top: 10px;
                font-weight: 600;
                font-size: 13px;
                padding: 14px 10px 10px 10px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }}
        """)
        grp_cycles_layout = QVBoxLayout(grp_cycles)
        grp_cycles_layout.setSpacing(10)

        lbl_cycles_info = QLabel(
            "Analyse <b>planning_historique.json</b> pour détecter automatiquement "
            "le cycle de chaque employé (3x8, 2x8, J, WE…) et met à jour "
            "<b>cycles_employes.json</b>."
        )
        lbl_cycles_info.setTextFormat(Qt.TextFormat.RichText)
        lbl_cycles_info.setWordWrap(True)
        lbl_cycles_info.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        grp_cycles_layout.addWidget(lbl_cycles_info)

        row_cycles = QHBoxLayout()

        from PyQt6.QtWidgets import QCheckBox as _QCB2
        self.check_ecraser = _QCB2("Écraser les cycles saisis manuellement")
        self.check_ecraser.setToolTip("Par défaut, les cycles déjà remplis manuellement sont conservés.")
        self.check_ecraser.setStyleSheet(f"""
            QCheckBox {{
                color: {COULEURS['texte_secondaire']};
                font-size: 12px;
            }}
            QCheckBox::indicator {{
                width: 15px; height: 15px;
                border: 1px solid {COULEURS['bordure']};
                border-radius: 3px;
                background-color: {COULEURS['bg_carte']};
            }}
            QCheckBox::indicator:checked {{
                background-color: {COULEURS['accent_warning']};
                border-color: {COULEURS['accent_warning']};
            }}
        """)

        self.btn_detecter_cycles = QPushButton("🔍  Détecter les cycles")
        self.btn_detecter_cycles.setFixedWidth(220)
        self.btn_detecter_cycles.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['accent_succes']};
                color: #1E1E2E;
                border-radius: 6px;
                padding: 8px 18px;
                font-weight: 700;
                font-size: 13px;
            }}
            QPushButton:hover {{ background-color: #88FFB8; }}
            QPushButton:disabled {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte_secondaire']};
            }}
        """)
        self.btn_detecter_cycles.clicked.connect(self._lancer_detection_cycles)

        row_cycles.addWidget(self.check_ecraser)
        row_cycles.addStretch()
        row_cycles.addWidget(self.btn_detecter_cycles)
        grp_cycles_layout.addLayout(row_cycles)

        layout.addWidget(grp_cycles)
        layout.addStretch()

        self._sous_onglets.addTab(page, "📋  Historique & Cycles")
        self._charger_historique_imports()


    # --------------------------------------------------
    # Statut fichier planning
    # --------------------------------------------------
    # --------------------------------------------------
    # Historique des imports
    # --------------------------------------------------
    def _charger_historique_imports(self):
        """Recharge le tableau depuis import_historique.json."""
        historique = charger_json(IMPORT_HISTORIQUE_JSON())
        if not isinstance(historique, list):
            historique = []
        self.tableau_historique.setRowCount(0)
        for entree in reversed(historique):   # plus récent en haut
            row = self.tableau_historique.rowCount()
            self.tableau_historique.insertRow(row)
            self.tableau_historique.setItem(row, 0, QTableWidgetItem(entree.get("nom_fichier", "")))
            self.tableau_historique.setItem(row, 1, QTableWidgetItem(entree.get("date_import", "")))
            source_map = {"excel_fab": "Excel Fab", "pdf_adp": "PDF ADP", "excel_adp": "Excel ADP"}
            src = source_map.get(entree.get("source", ""), entree.get("source", ""))
            self.tableau_historique.setItem(row, 2, QTableWidgetItem(src))
            nb = entree.get("nb_employes", "")
            item_nb = QTableWidgetItem(str(nb) if nb != "" else "—")
            item_nb.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tableau_historique.setItem(row, 3, item_nb)

    def _enregistrer_import_historique(self, nom_fichier: str, source: str, nb_employes: int = 0):
        """Ajoute ou met a jour une entree dans import_historique.json apres un import reussi."""
        from datetime import datetime
        historique = charger_json(IMPORT_HISTORIQUE_JSON())
        if not isinstance(historique, list):
            historique = []
        # Mettre a jour si le fichier existe deja, sinon ajouter
        for entree in historique:
            if entree.get("nom_fichier") == nom_fichier:
                entree["date_import"] = datetime.now().isoformat(timespec="seconds")
                entree["source"]      = source
                entree["nb_employes"] = nb_employes
                break
        else:
            historique.append({
                "nom_fichier": nom_fichier,
                "date_import": datetime.now().isoformat(timespec="seconds"),
                "source":      source,
                "nb_employes": nb_employes,
            })
        sauvegarder_json(IMPORT_HISTORIQUE_JSON(), historique)
        self._charger_historique_imports()

    def _supprimer_import_selectionne(self):
        """Supprime l'entree selectionnee de l'historique ET les donnees planning associees."""
        row = self.tableau_historique.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Aucune selection", "Selectionnez un import dans la liste.")
            return

        # Reconstruire l'index reel (on affiche reversed)
        historique = charger_json(IMPORT_HISTORIQUE_JSON())
        if not isinstance(historique, list):
            return
        idx_reel = len(historique) - 1 - row
        if idx_reel < 0 or idx_reel >= len(historique):
            return

        entree = historique[idx_reel]
        nom    = entree.get("nom_fichier", "")
        source = entree.get("source", "")

        rep = QMessageBox.question(
            self, "Supprimer cet import",
            f"Supprimer l'import '{nom}' ?\n\n"
            f"Cela supprimera :\n"
            f"  - L'entree de l'historique\n"
            f"  - Les donnees de planning associees (source : {source})\n\n"
            f"Un backup sera cree avant la suppression.\n"
            f"Cette action est IRREVERSIBLE. Confirmer ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if rep != QMessageBox.StandardButton.Yes:
            return

        # Backup des deux fichiers
        faire_backup(IMPORT_HISTORIQUE_JSON())
        faire_backup(PLANNING_HISTORIQUE_JSON())

        # Supprimer les donnees planning dont la source correspond
        planning = charger_json(PLANNING_HISTORIQUE_JSON())
        for cle_emp, data_emp in planning.items():
            if not isinstance(data_emp, dict):
                continue
            for sous_cle in ("semaines", "jours"):
                bloc = data_emp.get(sous_cle, {})
                if not isinstance(bloc, dict):
                    continue
                a_supprimer = [k for k, v in bloc.items()
                               if isinstance(v, dict) and v.get("source") == source]
                for k in a_supprimer:
                    del bloc[k]
        sauvegarder_json(PLANNING_HISTORIQUE_JSON(), planning)

        # Supprimer l'entree de l'historique
        historique.pop(idx_reel)
        sauvegarder_json(IMPORT_HISTORIQUE_JSON(), historique)

        self._charger_historique_imports()
        self._maj_statut_planning()
        self.console.append(f"\n\U0001f5d1  Import '{nom}' supprime. Donnees planning ({source}) effacees.")
        QMessageBox.information(self, "Suppression effectuee",
            f"L'import '{nom}' et ses donnees ont ete supprimes.\n"
            f"Vous pouvez maintenant re-importer ce fichier.")

    def _maj_statut_planning(self):
        chemin = PLANNING_HISTORIQUE_JSON()
        if os.path.exists(chemin):
            try:
                data = charger_json(chemin)
                nb   = len(data)
                self.lbl_planning_statut.setText(
                    f"✅  planning_historique.json  ({nb} employé(s) enregistré(s))"
                )
                self.lbl_planning_statut.setStyleSheet(f"color: {COULEURS['accent_succes']}; font-size: 12px;")
            except Exception:
                self.lbl_planning_statut.setText("⚠️  planning_historique.json — illisible")
                self.lbl_planning_statut.setStyleSheet(f"color: {COULEURS['accent_warning']}; font-size: 12px;")
        else:
            self.lbl_planning_statut.setText("ℹ️  planning_historique.json sera créé au premier import")
            self.lbl_planning_statut.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")

    # --------------------------------------------------
    # Lancement d'un import
    # --------------------------------------------------
    def _lancer_import(self, source: str):
        if self._worker and self._worker.isRunning():
            QMessageBox.warning(self, "Import en cours", "Un import est déjà en cours. Attendez sa fin.")
            return

        # Sélection du fichier selon la source
        if source == 'excel_fab':
            chemin, _ = QFileDialog.getOpenFileName(
                self, "Sélectionner le fichier Excel Fabrication annuel",
                BASE_DIR, "Fichiers Excel (*.xlsx *.xls)"
            )
        elif source == 'pdf_adp':
            chemin, _ = QFileDialog.getOpenFileName(
                self, "Sélectionner le PDF ADP mensuel",
                BASE_DIR, "Fichiers PDF (*.pdf)"
            )
        else:
            chemin, _ = QFileDialog.getOpenFileName(
                self, "Sélectionner le fichier Excel ADP hebdomadaire",
                BASE_DIR, "Fichiers Excel (*.xlsx *.xls)"
            )

        if not chemin:
            return

        # Validation du nom de fichier
        import re as _re
        nom = os.path.basename(chemin)
        regexes = {
            'excel_fab': (r'^\d{4}\.(xlsx|xlsm)$',   "2021.xlsx  ou  2024.xlsm"),
            'pdf_adp':   (r'^\d{2}_\d{4}\.pdf$',      "01_2025.pdf  ou  08_2025.pdf"),
            'excel_adp': (r'^S\d{2}_\d{4}\.(xlsx|xls)$', "S41_2025.xlsx  ou  S41_2025.xls"),
        }
        pattern, exemple = regexes[source]
        if not _re.match(pattern, nom, _re.IGNORECASE):
            QMessageBox.critical(
                self, "Nom de fichier invalide",
                f"Le fichier '{nom}' ne correspond pas au format attendu.\n\n"
                f"Format attendu : {exemple}\n\n"
                f"Renommez le fichier et relancez l'import."
            )
            return

        if not os.path.exists(EMPLOYES_CONTRATS_JSON()):
            QMessageBox.critical(self, "Fichier manquant",
                "employes_contrats.json introuvable.\n"
                "Vérifiez le dossier de travail dans le bandeau en haut.")
            return

        # Désactiver les boutons pendant l'import
        self._set_boutons_actifs(False)
        self.console.append(f"\n{'─'*60}")
        self.console.append(f"📂  Fichier : {os.path.basename(chemin)}")

        self._worker = WorkerImport(source, chemin)
        self._worker.log_signal.connect(self._on_log)
        self._worker.fini_signal.connect(self._on_fini)
        self._worker.erreur_signal.connect(self._on_erreur)
        self._worker.start()

    # --------------------------------------------------
    # Détection des cycles
    # --------------------------------------------------
    def _lancer_detection_cycles(self):
        if self._worker and self._worker.isRunning():
            QMessageBox.warning(self, "Import en cours",
                "Un import est déjà en cours. Attendez sa fin.")
            return
        if self._worker_cycles and self._worker_cycles.isRunning():
            QMessageBox.warning(self, "Détection en cours",
                "Une détection de cycles est déjà en cours.")
            return

        if not os.path.exists(PLANNING_HISTORIQUE_JSON()):
            QMessageBox.warning(self, "Fichier manquant",
                "planning_historique.json introuvable.\n"
                "Importez d'abord des données de planning.")
            return

        ecraser = self.check_ecraser.isChecked()
        if ecraser:
            rep = QMessageBox.question(
                self, "Confirmation",
                "⚠️  Vous allez écraser les cycles saisis manuellement.\n\n"
                "Continuer ?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if rep != QMessageBox.StandardButton.Yes:
                return

        self.console.append(f"\n{'─'*60}")
        self.console.append("🔍  Démarrage de la détection automatique des cycles…")
        if ecraser:
            self.console.append("⚠️  Mode : écrasement des saisies manuelles activé")

        self.btn_detecter_cycles.setEnabled(False)
        self.btn_detecter_cycles.setText("⏳  Détection en cours…")

        self._worker_cycles = WorkerDetectionCycles(ecraser_manuel=ecraser)
        self._worker_cycles.log_signal.connect(self._on_cycles_log)
        self._worker_cycles.fini_signal.connect(self._on_cycles_fini)
        self._worker_cycles.erreur_signal.connect(self._on_cycles_erreur)
        self._worker_cycles.start()

    def _on_cycles_log(self, msg: str):
        self.console.append(msg)

    def _on_cycles_fini(self, resultats: list):
        """
        Reçoit la liste des cycles détectés (analyser_sans_ecrire).
        Ouvre DialogueValidationCycles, puis écrit uniquement les entrées validées.
        """
        self.btn_detecter_cycles.setEnabled(True)
        self.btn_detecter_cycles.setText("\U0001f50d  Détecter les cycles")

        self.console.append(f"\n{'='*50}")

        if not resultats:
            self.console.append("ℹ️  Aucun cycle à valider (données insuffisantes ou tous déjà définis).")
            self._status("\U0001f50d  Détection terminée — aucun nouveau cycle détecté.")
            return

        nb_conflits = sum(1 for r in resultats if r['conflit'])
        self.console.append(
            f"\U0001f4ca  {len(resultats)} cycle(s) détecté(s) — "
            f"{nb_conflits} conflit(s) avec saisie manuelle"
        )
        self.console.append("⏸️  En attente de validation par l'utilisateur…")

        # Ouvrir la popup de validation
        dlg = DialogueValidationCycles(resultats, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            self.console.append("❌  Validation annulée — aucun cycle appliqué.")
            self._status("\U0001f50d  Détection annulée.")
            return

        valides = dlg.get_resultats_valides()
        if not valides:
            self.console.append("ℹ️  Aucun cycle coché — aucun cycle appliqué.")
            self._status("\U0001f50d  Détection terminée — aucun cycle appliqué.")
            return

        # ── Écriture dans cycles_employes.json ─────────────────────
        chemin = CYCLES_EMPLOYES_JSON()
        faire_backup(chemin, callback_log=self.console.append)
        cycles = charger_json(chemin)

        for r in valides:
            cle = r['cle_emp']
            if cle not in cycles:
                cycles[cle] = {}
            cycles[cle]['cycle_depart'] = r['cycle_depart']
            cycles[cle]['date_depart']  = r['date_depart']
            cycles[cle]['cycle_type']   = r['cycle_type']
            # Stocker le motif réel détecté (ex: ['M','N','AM'])
            # Utilisé par generer_hypothetiques.py pour la continuité exacte
            if r.get('motif'):
                cycles[cle]['motif'] = r['motif']

        # Tri alphabétique, COMMENTAIRE en tête
        commentaire = cycles.pop("COMMENTAIRE", None)
        cycles_tries = dict(sorted(cycles.items()))
        if commentaire:
            cycles_tries = {"COMMENTAIRE": commentaire, **cycles_tries}

        sauvegarder_json(chemin, cycles_tries)

        # ── Logs résumé ─────────────────────────────────────────────
        nb_conflits_valides = sum(1 for r in valides if r['conflit'])
        self.console.append(f"\n✅  {len(valides)} cycle(s) appliqué(s)")
        if nb_conflits_valides:
            self.console.append(f"   ⚠️  dont {nb_conflits_valides} conflit(s) écrasé(s)")
        nb_ignores = len(resultats) - len(valides)
        if nb_ignores:
            self.console.append(f"   ⏭️  {nb_ignores} décochés — ignorés")
        self.console.append(f"   💾  Sauvegardé : {os.path.basename(chemin)}")

        # ── Rafraîchir onglet Cycles Employés ───────────────────────
        try:
            fenetre = self.window()
            if hasattr(fenetre, 'onglet_cycles_emp'):
                fenetre.onglet_cycles_emp.charger_donnees()
        except Exception:
            pass

        self._status(
            f"\U0001f50d  {len(valides)} cycle(s) appliqué(s)"
            + (f" — {nb_conflits_valides} conflit(s) écrasé(s)" if nb_conflits_valides else "")
        )
        QMessageBox.information(
            self, "Cycles appliqués",
            f"✅  {len(valides)} cycle(s) appliqué(s) avec succès !\n\n"
            + (f"  ⚠️  {nb_conflits_valides} conflit(s) écrasé(s)\n" if nb_conflits_valides else "")
            + (f"  ⏭️  {nb_ignores} ignoré(s) (décochés)\n" if nb_ignores else "")
            + f"  💾  Backup créé automatiquement."
        )

    def _on_cycles_erreur(self, msg: str):
        self.btn_detecter_cycles.setEnabled(True)
        self.btn_detecter_cycles.setText("\U0001f50d  Détecter les cycles")
        self.console.append(f"\n❌  ERREUR détection cycles : {msg}")
        QMessageBox.critical(self, "Erreur lors de la détection", msg)

    def _set_boutons_actifs(self, actif: bool):
        for btn in [self.btn_import_excel_fab, self.btn_import_pdf_adp, self.btn_import_excel_adp]:
            btn.setEnabled(actif)
            if not actif:
                btn.setText("⏳ Import en cours…")
            else:
                btn.setText("Sélectionner et importer")

    # --------------------------------------------------
    # Slots worker
    # --------------------------------------------------
    def _status(self, msg: str, duree: int = 4000):
        """Affiche un message temporaire dans la barre de statut principale."""
        try:
            self.window().status.showMessage(msg, duree)
        except Exception:
            pass

    def _on_log(self, msg: str):
        self.console.append(msg)

    def _on_fini(self, resultats: dict):
        self._set_boutons_actifs(True)
        self._maj_statut_planning()

        importes     = resultats.get("importes", 0)
        non_reconnus = resultats.get("non_reconnus", [])
        doublons     = resultats.get("doublons", [])

        # Enregistrer dans l'historique des imports (uniquement si import non annule)
        if self._worker:
            self._enregistrer_import_historique(
                nom_fichier  = os.path.basename(self._worker.chemin_fichier),
                source       = self._worker.source,
                nb_employes  = importes,
            )

        # Résumé détaillé (jours/semaines écrits si disponibles)
        details = resultats.get("details", {})
        jours_ecrits = details.get("jours_ecrits", details.get("semaines_ecrites", 0))
        if jours_ecrits:
            self.console.append(f"\n✅  Import terminé — {importes} employé(s), {jours_ecrits} entrée(s) écrite(s)")
        else:
            self.console.append(f"\n✅  Import terminé — {importes} employé(s) importé(s)")
        if non_reconnus:
            self.console.append(f"  ❓ {len(non_reconnus)} nom(s) non reconnu(s)")
        if doublons:
            self.console.append(f"  ⚠️ {len(doublons)} conflit(s) doublon détecté(s) — non écrasés")

        # Afficher le dialogue non-reconnus si besoin
        if non_reconnus:
            self._afficher_non_reconnus(non_reconnus)

        # Afficher le dialogue doublons si besoin
        if doublons:
            self._afficher_doublons(doublons)

        if not non_reconnus and not doublons:
            QMessageBox.information(
                self, "Import réussi",
                f"✅  Import terminé avec succès !\n{importes} entrée(s) enregistrée(s)."
            )
        self._status(f"✅  Import terminé — {importes} entrée(s) enregistrée(s).")

    def _on_erreur(self, msg: str):
        self._set_boutons_actifs(True)
        self.console.append(f"\n❌  ERREUR : {msg}")
        QMessageBox.critical(self, "Erreur lors de l'import", msg)

    # --------------------------------------------------
    # Dialogue non-reconnus
    # --------------------------------------------------
    def _afficher_non_reconnus(self, non_reconnus: list):
        employes = charger_json(EMPLOYES_CONTRATS_JSON())
        dlg = DialogueNonReconnus(non_reconnus, employes, parent=self)
        dlg.exec()
        decisions        = dlg.get_decisions()
        nouveaux_employes = dlg.get_nouveaux_employes()

        # ── 1. Écrire les nouveaux employés dans employes_contrats.json ──
        if nouveaux_employes:
            self._creer_employes_depuis_import(nouveaux_employes)

        # ── 2. Compter et logger les résultats ───────────────────────────
        nb_crees   = len(nouveaux_employes)
        nb_associes = sum(
            1 for v in decisions.values()
            if v != "ignorer" and not isinstance(v, dict)
        )
        nb_ignores  = sum(1 for v in decisions.values() if v == "ignorer")

        if nb_crees:
            self.console.append(f"  🆕 {nb_crees} employé(s) créé(s) depuis l'import")
        if nb_associes:
            self.console.append(
                f"  🔗 {nb_associes} association(s) manuelle(s) enregistrée(s) "
                f"(ré-import nécessaire pour les appliquer)"
            )
        if nb_ignores == len(non_reconnus) and not nb_crees:
            self.console.append("  ℹ️  Tous les non-reconnus ont été ignorés.")

    def _creer_employes_depuis_import(self, nouveaux: list):
        """
        Écrit les nouveaux employés dans employes_contrats.json.
        Backup automatique avant écriture. Ignore les doublons de clé.
        """
        import shutil, datetime as _dt
        chemin = EMPLOYES_CONTRATS_JSON()

        # Backup dans backup/
        faire_backup(chemin, callback_log=self.console.append)

        employes = charger_json(chemin)
        nb_ok = 0
        nb_skip = 0

        for emp in nouveaux:
            cle = emp["cle"]
            if cle in employes:
                self.console.append(f"  ⚠️  Clé déjà existante, ignorée : {cle}")
                nb_skip += 1
                continue

            # Construire l'entrée v3
            entree = {
                "date_debut":  emp.get("date_debut", ""),
                "actif":       emp.get("actif", True),
                "departements": [],
            }
            if emp.get("date_sortie"):
                entree["date_sortie"] = emp["date_sortie"]
            if emp.get("departement"):
                entree["departements"] = [{
                    "departement": emp["departement"],
                    "debut": emp.get("date_debut", ""),
                    "fin": emp.get("date_sortie") or None,
                }]

            employes[cle] = entree
            self.console.append(f"  ✅  Employé créé : {cle}")
            nb_ok += 1

        if nb_ok:
            # Trier alphabétiquement et sauvegarder
            employes_tries = dict(sorted(employes.items()))
            sauvegarder_json(chemin, employes_tries)
            self.console.append(
                f"  💾 employes_contrats.json mis à jour ({nb_ok} ajout(s))"
            )
            # Rafraîchir l'onglet Employés si accessible
            try:
                fenetre = self.window()
                if hasattr(fenetre, 'onglet_employes'):
                    fenetre.onglet_employes.charger_donnees()
            except Exception:
                pass

    # --------------------------------------------------
    # Dialogue doublons
    # --------------------------------------------------
    def _afficher_doublons(self, doublons: list):
        dlg = DialogueDoublon(doublons, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        indices_ecraser = dlg.get_decisions()
        if not indices_ecraser:
            self.console.append("  ℹ️  Aucun doublon écrasé — données existantes conservées.")
            return

        # Appliquer les écrasements manuellement dans le planning
        try:
            planning = charger_json(PLANNING_HISTORIQUE_JSON())
            nb_ecrases = 0
            for idx in indices_ecraser:
                d = doublons[idx]
                cle_emp = d["cle_emp"]
                cle_j   = d["cle_j"]
                nouveau = d["nouveau"]
                if cle_emp in planning:
                    # Déterminer si c'est une semaine (S01_AAAA) ou un jour (AAAA-MM-JJ)
                    if cle_j.startswith("S"):
                        planning[cle_emp].setdefault("semaines", {})[cle_j] = nouveau
                    else:
                        planning[cle_emp].setdefault("jours", {})[cle_j] = nouveau
                    nb_ecrases += 1

            sauvegarder_json(PLANNING_HISTORIQUE_JSON(), planning)
            self.console.append(f"  ✅  {nb_ecrases} doublon(s) écrasé(s) et sauvegardé(s).")
            self._maj_statut_planning()
        except Exception as e:
            self.console.append(f"  ❌  Erreur lors de l'application des doublons : {e}")


# =====================================================
# ONGLET VISUALISATION PLANNING — Phase 4
# =====================================================

# Couleurs par cycle