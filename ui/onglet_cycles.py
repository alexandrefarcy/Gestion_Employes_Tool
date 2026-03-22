"""
ui/onglet_cycles.py
===================
OngletCyclesDefinitions + OngletCyclesEmployes + DialogueCycleCustom
"""

import os
from datetime import date as _date

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit, QDialog, QDialogButtonBox,
    QGroupBox, QRadioButton, QButtonGroup, QSpinBox,
    QMessageBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont

from ui.constantes import (
    COULEURS,
    CYCLES_DEFINITIONS_JSON, CYCLES_EMPLOYES_JSON, EMPLOYES_CONTRATS_JSON,
    charger_json, sauvegarder_json, faire_backup,
)
from ui.widgets import ComboSansScroll

class OngletCyclesDefinitions(QWidget):
    def __init__(self):
        super().__init__()
        self.data = {}
        self._construire_ui()
        self.charger_donnees()

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── En-tête ──────────────────────────────────────────────
        entete = QHBoxLayout()
        titre = QLabel("\u2699\ufe0f  D\xe9finitions des cycles de travail")
        titre.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {COULEURS['texte']};")
        entete.addWidget(titre)
        entete.addStretch()
        layout.addLayout(entete)

        info = QLabel(
            "Ajoutez, modifiez ou supprimez les cycles disponibles pour tous les employ\xe9s. "
            "Double-cliquez sur une cellule pour la modifier."
        )
        info.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        info.setWordWrap(True)
        layout.addWidget(info)

        # ── Tableau ──────────────────────────────────────────────
        self.tableau = QTableWidget()
        self.tableau.setColumnCount(3)
        self.tableau.setHorizontalHeaderLabels(["Nom du cycle", "Type", "Description"])
        hh = self.tableau.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.tableau.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableau.verticalHeader().setVisible(False)
        self.tableau.verticalHeader().setDefaultSectionSize(50)
        self.tableau.setAlternatingRowColors(True)
        self.tableau.setStyleSheet(self.tableau.styleSheet() + f"""
            QTableWidget {{ alternate-background-color: {COULEURS['bg_carte']}; }}
        """)
        layout.addWidget(self.tableau)

        # ── Boutons ──────────────────────────────────────────────
        boutons = QHBoxLayout()
        btn_ajouter = QPushButton("\u2795  Ajouter un cycle")
        btn_ajouter.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['accent']};
                color: #1E1E2E;
                border-radius: 6px;
                padding: 7px 16px;
                font-weight: 700;
                font-size: 12px;
            }}
            QPushButton:hover {{ background-color: #A89AFF; }}
        """)
        btn_supprimer = QPushButton("\U0001f5d1  Supprimer la ligne")
        btn_supprimer.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['accent_danger']};
                color: #1E1E2E;
                border-radius: 6px;
                padding: 7px 16px;
                font-weight: 700;
                font-size: 12px;
            }}
            QPushButton:hover {{ background-color: #FF9A95; }}
        """)
        btn_recharger   = QPushButton("\U0001f504  Recharger")
        btn_sauvegarder = QPushButton("\U0001f4be  Sauvegarder")
        btn_sauvegarder.setStyleSheet(
            f"background-color: {COULEURS['accent_succes']}; color: #1E1E2E; "
            f"border-radius: 6px; padding: 7px 16px; font-weight: 700; font-size: 12px;"
        )
        btn_ajouter.clicked.connect(self._ajouter_ligne)
        btn_supprimer.clicked.connect(self._supprimer_ligne)
        btn_recharger.clicked.connect(self.charger_donnees)
        btn_sauvegarder.clicked.connect(self.sauvegarder_donnees)

        boutons.addWidget(btn_ajouter)
        boutons.addWidget(btn_supprimer)
        boutons.addStretch()
        boutons.addWidget(btn_recharger)
        boutons.addWidget(btn_sauvegarder)
        layout.addLayout(boutons)

    def charger_donnees(self):
        self.data = charger_json(CYCLES_DEFINITIONS_JSON())
        self._remplir_tableau()

    def _remplir_tableau(self):
        self.tableau.setRowCount(0)
        for nom_cycle, details in self.data.items():
            if nom_cycle == "COMMENTAIRE":
                continue
            row = self.tableau.rowCount()
            self.tableau.insertRow(row)
            if isinstance(details, dict):
                type_val = details.get("type", "")
                desc_val = details.get("description", "")
            else:
                type_val = str(details) if details else ""
                desc_val = ""
            self.tableau.setItem(row, 0, QTableWidgetItem(nom_cycle))
            self.tableau.setItem(row, 1, QTableWidgetItem(type_val))
            self.tableau.setItem(row, 2, QTableWidgetItem(desc_val))

    def _ajouter_ligne(self):
        row = self.tableau.rowCount()
        self.tableau.insertRow(row)
        self.tableau.setItem(row, 0, QTableWidgetItem("Nouveau cycle"))
        self.tableau.setItem(row, 1, QTableWidgetItem(""))
        self.tableau.setItem(row, 2, QTableWidgetItem(""))
        self.tableau.scrollToBottom()
        self.tableau.editItem(self.tableau.item(row, 0))

    def _supprimer_ligne(self):
        row = self.tableau.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Attention", "S\xe9lectionnez une ligne \xe0 supprimer.")
            return
        nom = self.tableau.item(row, 0).text() if self.tableau.item(row, 0) else "?"
        rep = QMessageBox.question(
            self, "Supprimer le cycle",
            f"Supprimer le cycle \u00ab {nom} \u00bb ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if rep == QMessageBox.StandardButton.Yes:
            self.tableau.removeRow(row)

    def sauvegarder_donnees(self):
        try:
            nouvelle_data = {}
            for row in range(self.tableau.rowCount()):
                nom  = self.tableau.item(row, 0).text().strip() if self.tableau.item(row, 0) else ""
                type_v = self.tableau.item(row, 1).text().strip() if self.tableau.item(row, 1) else ""
                desc = self.tableau.item(row, 2).text().strip() if self.tableau.item(row, 2) else ""
                if not nom:
                    continue
                nouvelle_data[nom] = {"type": type_v, "description": desc}
            faire_backup(CYCLES_DEFINITIONS_JSON())
            sauvegarder_json(CYCLES_DEFINITIONS_JSON(), nouvelle_data)
            self.data = nouvelle_data
            QMessageBox.information(self, "Succ\xe8s", "\u2705 D\xe9finitions sauvegard\xe9es !")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"\u274c {e}")


# =====================================================
# ONGLET CYCLES EMPLOYÉS
# =====================================================


class OngletCyclesEmployes(QWidget):
    def __init__(self):
        super().__init__()
        self.data_employes = {}
        self.data_cycles   = {}
        self._construire_ui()
        self.charger_donnees()

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── En-tête : titre + recherche ──────────────────────────
        entete = QHBoxLayout()
        titre = QLabel("\U0001f504  Affectation des cycles par employé")
        titre.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {COULEURS['texte']};")
        entete.addWidget(titre)
        entete.addStretch()
        self.champ_recherche = QLineEdit()
        self.champ_recherche.setPlaceholderText("\U0001f50d  Rechercher...")
        self.champ_recherche.setFixedWidth(240)
        self.champ_recherche.textChanged.connect(self._filtrer_texte)
        entete.addWidget(self.champ_recherche)
        layout.addLayout(entete)

        # ── Barre de filtres statut + légende ────────────────────
        filtres = QHBoxLayout()

        # Radios filtre statut
        self._grp_filtre = QButtonGroup(self)
        for label, val in [("Tous", "tous"), ("✅ Définis", "definis"),
                            ("\u26a0\ufe0f Incomplet", "incomplet"), ("\u274c Non défini", "non_defini")]:
            rb = QRadioButton(label)
            rb.setProperty("filtre_val", val)
            rb.setStyleSheet(f"color: {COULEURS['texte']}; font-size: 12px;")
            self._grp_filtre.addButton(rb)
            filtres.addWidget(rb)
            if val == "tous":
                rb.setChecked(True)
        self._grp_filtre.buttonClicked.connect(self._appliquer_filtre)

        filtres.addSpacing(30)
        for couleur, texte in [
            (COULEURS['accent_succes'],  "✅ Cycle défini"),
            (COULEURS['accent_warning'], "\u26a0\ufe0f Incomplet"),
            (COULEURS['accent_danger'],  "\u274c Non défini"),
        ]:
            lbl = QLabel(texte)
            lbl.setStyleSheet(f"color: {couleur}; font-size: 12px; font-weight: 600;")
            filtres.addWidget(lbl)
            filtres.addSpacing(16)
        filtres.addStretch()
        layout.addLayout(filtres)

        # ── Tableau (5 colonnes) ─────────────────────────────────
        self.tableau = QTableWidget()
        self.tableau.setColumnCount(5)
        self.tableau.setHorizontalHeaderLabels([
            "Employé", "Cycle (legacy)", "Poste de départ", "Date de départ", "Matricule"
        ])
        hh = self.tableau.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableau.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableau.verticalHeader().setDefaultSectionSize(50)
        self.tableau.verticalHeader().setVisible(False)
        self.tableau.setAlternatingRowColors(True)
        self.tableau.setStyleSheet(self.tableau.styleSheet() + f"""
            QTableWidget {{ alternate-background-color: {COULEURS['bg_carte']}; }}
        """)
        layout.addWidget(self.tableau)

        # ── Bandeau read-only ────────────────────────────────────
        lbl_readonly = QLabel(
            "Cet onglet est en lecture seule. "
            "Les cycles sont d\xe9tect\xe9s automatiquement depuis l'onglet Planning "
            "(bouton D\xe9tecter les cycles)."
        )
        lbl_readonly.setWordWrap(True)
        lbl_readonly.setStyleSheet(
            f"color: {COULEURS['accent_warning']}; font-size: 12px; "
            f"font-weight: 600; padding: 6px 10px; "
            f"background-color: {COULEURS['bg_carte']}; border-radius: 6px;"
        )
        layout.addWidget(lbl_readonly)

        # ── Barre du bas : stats + wipe ──────────────────────────
        boutons = QHBoxLayout()
        self.lbl_stats = QLabel("")
        self.lbl_stats.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")

        btn_wipe = QPushButton("\U0001f5d1  Supprimer tous les cycles")
        btn_wipe.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['accent_danger']};
                color: #1E1E2E;
                border-radius: 6px;
                padding: 7px 16px;
                font-weight: 700;
                font-size: 12px;
            }}
            QPushButton:hover {{ background-color: #FF9A95; }}
        """)
        btn_wipe.clicked.connect(self._wipe_cycles)

        boutons.addWidget(self.lbl_stats)
        boutons.addStretch()

        btn_custom = QPushButton("\u270f\ufe0f  Cycle custom")
        btn_custom.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['accent']};
                border-radius: 6px;
                padding: 7px 16px;
                font-weight: 600;
                font-size: 12px;
            }}
            QPushButton:hover {{
                background-color: {COULEURS['accent']};
                color: #FFFFFF;
            }}
        """)
        btn_custom.clicked.connect(self._ouvrir_cycle_custom)
        boutons.addWidget(btn_custom)
        boutons.addSpacing(8)
        boutons.addWidget(btn_wipe)
        layout.addLayout(boutons)

    def _ouvrir_cycle_custom(self):
        """Ouvre le dialogue de saisie d'un cycle custom pour l'employe selectionne."""
        rows = self.tableau.selectionModel().selectedRows()
        if not rows:
            QMessageBox.information(
                self, "Cycle custom",
                "S\u00e9lectionnez un employ\u00e9 dans le tableau avant de d\u00e9finir un cycle custom."
            )
            return

        row = rows[0].row()
        item = self.tableau.item(row, 0)
        if not item:
            return
        cle_emp = item.data(Qt.ItemDataRole.UserRole)
        if not cle_emp:
            return

        # Motif actuel si d\u00e9j\u00e0 d\u00e9fini
        motif_actuel = self.data_employes.get(cle_emp, {}).get("motif", [])

        dlg = DialogueCycleCustom(cle_emp, motif_actuel, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        motif = dlg.get_motif()
        if not motif:
            return

        # Backup + sauvegarde
        chemin = CYCLES_EMPLOYES_JSON()
        faire_backup(chemin)
        data = charger_json(chemin)
        if cle_emp not in data:
            data[cle_emp] = {}
        data[cle_emp]["cycle_type"]   = "CUSTOM"
        data[cle_emp]["cycle_depart"] = motif[0]
        data[cle_emp]["motif"]        = motif
        # Conserver date_depart existante ou mettre date du jour
        if not data[cle_emp].get("date_depart"):
            from datetime import date as _date
            data[cle_emp]["date_depart"] = _date.today().strftime("%d-%m-%Y")
        sauvegarder_json(chemin, data)
        self.charger_donnees()

        nom = cle_emp.split("|")[0]
        QMessageBox.information(
            self, "Cycle custom appliqu\u00e9",
            f"\u2705  Cycle custom d\u00e9fini pour {nom}\n"
            f"   Motif : {" \u2192 ".join(motif)}\n"
            "\nRelancez \"G\u00e9n\u00e9rer les hypoth\u00e9tiques\" pour appliquer."
        )

    def charger_donnees(self):
        self.data_employes = charger_json(CYCLES_EMPLOYES_JSON())
        self.data_cycles   = charger_json(CYCLES_DEFINITIONS_JSON())
        self._remplir_tableau()

    def _wipe_cycles(self):
        rep = QMessageBox.question(
            self, "Supprimer tous les cycles",
            "ATTENTION — Wipe complet des cycles\n\n"
            "Cette action va supprimer TOUS les cycles detectes pour TOUS "
            "les employes dans cycles_employes.json.\n\n"
            "Les donnees de planning_historique.json ne seront PAS supprimees.\n\n"
            "Pour re-affecter les cycles, relancez la detection depuis l'onglet Planning.\n\n"
            "Un backup sera cree automatiquement avant la suppression.\n\n"
            "Cette action est IRREVERSIBLE. Confirmer ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if rep != QMessageBox.StandardButton.Yes:
            return

        chemin = CYCLES_EMPLOYES_JSON()
        backup = faire_backup(chemin)
        if backup:
            self.data_employes = charger_json(chemin)
            for cle, val in self.data_employes.items():
                if cle == "COMMENTAIRE" or "|" not in cle:
                    continue
                val.pop("cycle_depart", None)
                val.pop("cycle_type", None)
                val.pop("date_depart", None)
            sauvegarder_json(chemin, self.data_employes)
            self._remplir_tableau()
            QMessageBox.information(
                self, "Wipe effectue",
                f"Tous les cycles ont ete supprimes.\nBackup cree : {os.path.basename(backup)}"
            )
        else:
            QMessageBox.warning(
                self, "Erreur backup",
                "Impossible de creer le backup. Operation annulee."
            )

    def _get_liste_cycles(self):
        """Retourne les cycles définis SANS ligne vide en tête."""
        return [cle for cle in self.data_cycles if cle != "COMMENTAIRE"]

    def _get_liste_postes(self):
        """Retourne les postes de départ SANS ligne vide en tête."""
        return ["M", "AM", "N", "J"]

    def _statut_employe(self, info):
        """Retourne 'definis', 'incomplet' ou 'non_defini' selon les champs renseignés."""
        poste_ok = bool(info.get("cycle_depart", "").strip())
        legacy_ok = bool(info.get("cycle", "").strip())
        if poste_ok:
            return "definis"
        elif legacy_ok:
            return "incomplet"
        return "non_defini"

    def _remplir_tableau(self):
        self.tableau.setRowCount(0)
        liste_cycles = self._get_liste_cycles()
        liste_postes = self._get_liste_postes()
        nb_definis = nb_incomplet = nb_non_defini = 0

        # Charger employes_contrats pour la date_debut (pré-remplissage date départ)
        data_contrats = charger_json(EMPLOYES_CONTRATS_JSON())

        # Filtre statut actif
        filtre_btn = self._grp_filtre.checkedButton()
        filtre_val = filtre_btn.property("filtre_val") if filtre_btn else "tous"

        for cle, info in self.data_employes.items():
            if cle == "COMMENTAIRE":
                continue
            if "|" not in cle:
                continue

            # Ignorer entrées sans ID (doublons anciens sans matricule)
            id_part = cle.split("|")[1].strip()
            if not id_part:
                continue

            nom    = cle.split("|")[0].strip()
            statut = self._statut_employe(info)

            # Comptage total (avant filtre)
            if statut == "definis":       nb_definis += 1
            elif statut == "incomplet":   nb_incomplet += 1
            else:                         nb_non_defini += 1

            # Appliquer filtre radio
            if filtre_val != "tous" and statut != filtre_val:
                continue

            row = self.tableau.rowCount()
            self.tableau.insertRow(row)

            # Col 0 — Nom (couleur selon statut)
            item_nom = QTableWidgetItem(nom)
            item_nom.setData(Qt.ItemDataRole.UserRole, cle)
            item_nom.setFlags(item_nom.flags() & ~Qt.ItemFlag.ItemIsEditable)
            if statut == "definis":
                item_nom.setForeground(QColor(COULEURS['accent_succes']))
            elif statut == "incomplet":
                item_nom.setForeground(QColor(COULEURS['accent_warning']))
            else:
                item_nom.setForeground(QColor(COULEURS['accent_danger']))
            self.tableau.setItem(row, 0, item_nom)

            # Col 1 — Cycle legacy (read-only, texte grisé)
            legacy_val = info.get("cycle", "")
            item_legacy = QTableWidgetItem(legacy_val if legacy_val else "—")
            item_legacy.setFlags(item_legacy.flags() & ~Qt.ItemFlag.ItemIsEditable)
            item_legacy.setForeground(QColor(COULEURS['texte_secondaire']))
            item_legacy.setToolTip("Valeur legacy — lecture seule")
            self.tableau.setItem(row, 1, item_legacy)

            # Col 2 — Poste de départ (ComboSansScroll, sans ligne vide)
            combo_poste = ComboSansScroll()
            combo_poste.addItems(liste_postes)
            poste_val = info.get("cycle_depart", "").strip()
            if poste_val in liste_postes:
                combo_poste.setCurrentText(poste_val)
            else:
                combo_poste.setCurrentIndex(-1)
            self.tableau.setCellWidget(row, 2, combo_poste)

            # Col 3 — Date de départ (pré-remplie depuis date_debut si vide)
            date_val = info.get("date_depart", "").strip()
            if not date_val and cle in data_contrats:
                date_val = data_contrats[cle].get("date_debut", "")
            item_date = QTableWidgetItem(date_val)
            item_date.setFlags(item_date.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.tableau.setItem(row, 3, item_date)

            # Col 4 — Matricule (read-only)
            item_mat = QTableWidgetItem(id_part)
            item_mat.setFlags(item_mat.flags() & ~Qt.ItemFlag.ItemIsEditable)
            item_mat.setForeground(QColor(COULEURS['texte_secondaire']))
            self.tableau.setItem(row, 4, item_mat)

        total = nb_definis + nb_incomplet + nb_non_defini
        self.lbl_stats.setText(
            f"{nb_definis} définis  •  {nb_incomplet} incomplets  •  {nb_non_defini} non définis  "
            f"( {total} employés au total)"
        )

    def _appliquer_filtre(self, *_):
        """Changement filtre radio — recharge tableau complet puis applique filtre texte."""
        self._remplir_tableau()
        self._filtrer_texte(self.champ_recherche.text())

    def _filtrer_texte(self, texte: str):
        """Filtre texte seul — masque/affiche lignes sans recréer le tableau."""
        texte = texte.strip().lower()
        self.tableau.setUpdatesEnabled(False)
        for row in range(self.tableau.rowCount()):
            item = self.tableau.item(row, 0)
            masquer = bool(texte and item and texte not in item.text().lower())
            self.tableau.setRowHidden(row, masquer)
        self.tableau.setUpdatesEnabled(True)

    def filtrer_tableau(self, texte):
        """Alias pour compatibilité — délègue à _appliquer_filtre."""
        self._appliquer_filtre()

    def sauvegarder_donnees(self):
        try:
            for row in range(self.tableau.rowCount()):
                item_nom = self.tableau.item(row, 0)
                if not item_nom:
                    continue
                cle         = item_nom.data(Qt.ItemDataRole.UserRole)
                combo_poste = self.tableau.cellWidget(row, 2)
                item_date   = self.tableau.item(row, 3)

                if cle in self.data_employes:
                    self.data_employes[cle]["cycle_depart"] = combo_poste.currentText() if combo_poste else ""
                    self.data_employes[cle]["date_depart"]  = item_date.text() if item_date else ""

            sauvegarder_json(CYCLES_EMPLOYES_JSON(), self.data_employes)
            self._remplir_tableau()
            QMessageBox.information(self, "Succ\xe8s", "\u2705 Cycles sauvegard\xe9s avec succ\xe8s !")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"\u274c {e}")


class DialogueCycleCustom(QDialog):
    """Dialogue de saisie manuelle d'un cycle custom (Hyp-C).
    Permet de definir un pattern de 1 a 5 semaines pour un employe.
    Stocke dans cycles_employes.json avec cycle_type='CUSTOM'.
    """

    POSTES = ["M", "AM", "N", "J", "WE", "R"]
    NB_MAX_SEMAINES = 6
    def __init__(self, cle_employe: str, motif_actuel: list = None, parent=None):
        super().__init__(parent)
        self.cle_employe = cle_employe
        self.motif_actuel = motif_actuel or []
        self.setWindowTitle("Cycle custom")
        self.setMinimumWidth(460)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {COULEURS['bg_secondaire']};
                color: {COULEURS['texte']};
            }}
            QLabel {{ color: {COULEURS['texte']}; font-size: 13px; }}
            QComboBox {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 5px;
                padding: 5px 10px;
                font-size: 13px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                selection-background-color: {COULEURS['accent']};
            }}
            QSpinBox {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 5px;
                padding: 5px 8px;
                font-size: 13px;
            }}
        """)
        self._construire_ui()

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(20, 20, 20, 20)

        # Titre
        titre = QLabel("\u270f\ufe0f  D\u00e9finir un cycle custom")
        titre.setStyleSheet(
            f"font-size: 15px; font-weight: 700; color: {COULEURS['accent']};"
        )
        layout.addWidget(titre)

        # Nom employé
        nom = self.cle_employe.split("|")[0]
        lbl_emp = QLabel(f"Employ\u00e9 : <b>{nom}</b>")
        lbl_emp.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(lbl_emp)

        # Info
        lbl_info = QLabel(
            "Saisissez le pattern de semaines qui se r\u00e9p\u00e8te (1 \u00e0 5 semaines).\n"
            "Exemple : M / N / AM  \u2192  cycle 3x8 personnalis\u00e9"
        )
        lbl_info.setWordWrap(True)
        lbl_info.setStyleSheet(
            f"color: {COULEURS['texte_secondaire']}; font-size: 12px;"
        )
        layout.addWidget(lbl_info)

        # Nombre de semaines dans le motif
        row_nb = QHBoxLayout()
        lbl_nb = QLabel("Nombre de semaines dans le cycle :")
        self.spin_nb = QSpinBox()
        self.spin_nb.setMinimum(1)
        self.spin_nb.setMaximum(self.NB_MAX_SEMAINES)
        self.spin_nb.setValue(max(1, len(self.motif_actuel)))
        self.spin_nb.setFixedWidth(70)
        self.spin_nb.valueChanged.connect(self._actualiser_combos)
        row_nb.addWidget(lbl_nb)
        row_nb.addWidget(self.spin_nb)
        row_nb.addStretch()
        layout.addLayout(row_nb)

        # Combos semaines
        self.grp_semaines = QGroupBox("Pattern du cycle")
        self.grp_semaines.setStyleSheet(
            f"QGroupBox {{ color: {COULEURS['texte_secondaire']}; "
            f"border: 1px solid {COULEURS['bordure']}; border-radius: 6px; "
            f"margin-top: 6px; padding: 10px; font-size: 12px; }}"
        )
        self.lay_combos = QHBoxLayout(self.grp_semaines)
        self.lay_combos.setSpacing(8)
        self._combos = []
        for i in range(self.NB_MAX_SEMAINES):
            col = QVBoxLayout()
            lbl = QLabel(f"S{i+1}")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet(
                f"color: {COULEURS['texte_secondaire']}; font-size: 11px;"
            )
            cb = ComboSansScroll()
            for p in self.POSTES:
                cb.addItem(p, p)
            # Pré-remplir avec motif actuel si disponible
            if i < len(self.motif_actuel):
                idx = cb.findData(self.motif_actuel[i])
                if idx >= 0:
                    cb.setCurrentIndex(idx)
            col.addWidget(lbl)
            col.addWidget(cb)
            self._combos.append(cb)
            self.lay_combos.addLayout(col)
        layout.addWidget(self.grp_semaines)

        # Aperçu du motif
        self.lbl_apercu = QLabel()
        self.lbl_apercu.setStyleSheet(
            f"color: {COULEURS['accent_succes']}; font-size: 12px; "
            f"font-weight: 600; padding: 4px 8px;"
        )
        layout.addWidget(self.lbl_apercu)

        # Connecter combos pour aperçu
        for cb in self._combos:
            cb.currentIndexChanged.connect(self._maj_apercu)

        self._actualiser_combos(self.spin_nb.value())

        # Boutons
        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 5px;
                padding: 6px 18px;
                font-size: 13px;
            }}
            QPushButton:default {{
                background-color: {COULEURS['accent']};
                color: #FFFFFF;
                border: none;
            }}
            QPushButton:hover {{ border-color: {COULEURS['accent']}; }}
        """)
        btns.button(QDialogButtonBox.StandardButton.Ok).setText("\u2705  Appliquer")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _actualiser_combos(self, nb: int):
        """Affiche/masque les combos selon le nombre de semaines choisi."""
        for i, cb in enumerate(self._combos):
            cb.setVisible(i < nb)
            # Chercher le label au-dessus
            item = self.lay_combos.itemAt(i)
            if item and item.layout():
                lbl_item = item.layout().itemAt(0)
                if lbl_item and lbl_item.widget():
                    lbl_item.widget().setVisible(i < nb)
        self._maj_apercu()

    def _maj_apercu(self):
        """Met à jour l'aperçu du motif."""
        motif = self.get_motif()
        if motif:
            self.lbl_apercu.setText(
                "Motif : " + " \u2192 " .join(motif) + f" \u2192 (rep\u00e8te)"
            )
        else:
            self.lbl_apercu.setText("")

    def get_motif(self) -> list:
        """Retourne le motif saisi (liste de postes)."""
        nb = self.spin_nb.value()
        return [self._combos[i].currentData() for i in range(nb)]
