"""
ui/onglet_employes.py
=====================
DialogueEmploye + OngletEmployes
"""

import os

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit, QDialog, QDialogButtonBox, QFormLayout,
    QCheckBox, QRadioButton, QFrame, QMessageBox, QComboBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QFont

from ui.constantes import (
    COULEURS, DEPARTEMENTS_LISTE,
    EMPLOYES_CONTRATS_JSON, CYCLES_EMPLOYES_JSON, PLANNING_HISTORIQUE_JSON,
    charger_json, sauvegarder_json, faire_backup,
)
from ui.widgets import ComboSansScroll, ChampNom, ChampMatricule, ChampDateMasque

class DialogueEmploye(QDialog):
    def __init__(self, parent=None, employe=None, prefill=None):
        """
        employe : dict pour modification (cle, date, date_sortie, departement, archive)
        prefill : dict pour création depuis import {nom, id, titre_fenetre}
                  → pré-remplit les champs mais reste entièrement modifiable
        """
        super().__init__(parent)
        titre = "Ajouter un employe" if employe is None else "Modifier l'employe"
        if prefill and prefill.get("titre_fenetre"):
            titre = prefill["titre_fenetre"]
        self.setWindowTitle(titre)
        self.setMinimumWidth(460)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {COULEURS['bg_secondaire']};
                color: {COULEURS['texte']};
            }}
            QLabel {{
                color: {COULEURS['texte_secondaire']};
                font-size: 12px;
                font-weight: 600;
            }}
            QDialogButtonBox QPushButton {{
                min-width: 90px;
            }}
            QCheckBox {{
                color: {COULEURS['texte_secondaire']};
                font-size: 12px;
                font-weight: 600;
                spacing: 8px;
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                border: 1px solid {COULEURS['bordure']};
                border-radius: 4px;
                background-color: {COULEURS['bg_carte']};
            }}
            QCheckBox::indicator:checked {{
                background-color: {COULEURS['accent_warning']};
                border-color: {COULEURS['accent_warning']};
            }}
        """)

        from PyQt6.QtWidgets import QCheckBox as _QCheckBox
        self._QCheckBox = _QCheckBox

        layout = QFormLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(24, 24, 24, 24)

        self.champ_nom    = ChampNom()
        self.champ_prenom = ChampNom()
        self.champ_id     = ChampMatricule()

        # Champs date masqués (chiffres uniquement, tirets automatiques)
        self.champ_date   = ChampDateMasque()
        self.champ_sortie = ChampDateMasque()
        self.champ_sortie.setPlaceholderText("JJ-MM-AAAA  (vide = encore en poste)")
        self.champ_sortie.setStyleSheet(
            f"border: 1px solid {COULEURS['accent_warning']};"
        )

        self.combo_dept = QComboBox()
        self.combo_dept.addItems(DEPARTEMENTS_LISTE)

        # Checkbox Archivé
        self.check_archive = _QCheckBox("Archiver cet employé (actif: false, sans planning hypothétique)")

        if employe:
            nom_prenom, emp_id = employe["cle"].split("|")
            parties = nom_prenom.strip().split(" ", 1)
            self.champ_nom.setText(parties[0] if parties else "")
            self.champ_prenom.setText(parties[1] if len(parties) > 1 else "")
            self.champ_id.setText(emp_id.strip())
            self.champ_date.setText(employe.get("date", ""))
            self.champ_sortie.setText(employe.get("date_sortie", ""))
            dept = employe.get("departement", "")
            if dept in DEPARTEMENTS_LISTE:
                self.combo_dept.setCurrentText(dept)
            # Pré-cocher si déjà archivé (actif=False sans date_sortie)
            if employe.get("archive", False):
                self.check_archive.setChecked(True)
        elif prefill:
            nom_brut = prefill.get("nom", "")
            parties = nom_brut.strip().split(" ", 1)
            self.champ_nom.setText(parties[0] if parties else "")
            self.champ_prenom.setText(parties[1] if len(parties) > 1 else "")
            emp_id_prefill = prefill.get("id", "")
            if emp_id_prefill:
                self.champ_id.setText(str(emp_id_prefill))

        layout.addRow("Nom :", self.champ_nom)
        layout.addRow("Prenom :", self.champ_prenom)
        layout.addRow("Matricule (ID) :", self.champ_id)
        layout.addRow("Date debut contrat :", self.champ_date)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"color: {COULEURS['bordure']};")
        layout.addRow(sep)

        lbl_sortie_titre = QLabel("Date de sortie (optionnelle) :")
        lbl_sortie_titre.setStyleSheet(
            f"color: {COULEURS['accent_warning']}; font-size: 12px; font-weight: 700;"
        )
        layout.addRow(lbl_sortie_titre)
        layout.addRow("Date de sortie :", self.champ_sortie)

        lbl_info = QLabel(
            "Si renseignee, aucun planning hypothetique\n"
            "ne sera genere apres cette date."
        )
        lbl_info.setStyleSheet(
            f"color: {COULEURS['texte_secondaire']}; font-size: 11px; font-style: italic;"
        )
        layout.addRow(lbl_info)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setStyleSheet(f"color: {COULEURS['bordure']};")
        layout.addRow(sep2)

        layout.addRow("Departement :", self.combo_dept)

        sep3 = QFrame()
        sep3.setFrameShape(QFrame.Shape.HLine)
        sep3.setStyleSheet(f"color: {COULEURS['bordure']};")
        layout.addRow(sep3)

        layout.addRow(self.check_archive)

        boutons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        boutons.accepted.connect(self._valider_et_accepter)
        boutons.rejected.connect(self.reject)
        layout.addRow(boutons)

    @staticmethod
    def _parse_date(texte):
        """Convertit 'JJ-MM-AAAA' en date, retourne None si invalide."""
        from datetime import date as _date
        t = texte.strip()
        if len(t) != 10:
            return None
        try:
            j, m, a = t.split("-")
            return _date(int(a), int(m), int(j))
        except (ValueError, AttributeError):
            return None

    def _valider_et_accepter(self):
        """Valide les dates avant de fermer le dialogue. Garde le formulaire ouvert si erreur."""
        date_texte   = self.champ_date.text().strip()
        sortie_texte = self.champ_sortie.text().strip()

        # Vérifier cohérence date début / date sortie uniquement si les deux sont renseignées
        if date_texte and sortie_texte:
            d_debut  = self._parse_date(date_texte)
            d_sortie = self._parse_date(sortie_texte)
            if d_debut is None:
                QMessageBox.warning(self, "Date invalide",
                                    "La date de début de contrat n'est pas valide.\n"
                                    "Format attendu : JJ-MM-AAAA")
                self.champ_date.setFocus()
                return
            if d_sortie is None:
                QMessageBox.warning(self, "Date invalide",
                                    "La date de sortie n'est pas valide.\n"
                                    "Format attendu : JJ-MM-AAAA")
                self.champ_sortie.setFocus()
                return
            if d_sortie < d_debut:
                QMessageBox.warning(self, "Date de sortie incohérente",
                                    f"La date de sortie ({sortie_texte}) ne peut pas être\n"
                                    f"antérieure à la date de début de contrat ({date_texte}).")
                self.champ_sortie.setFocus()
                return

        self.accept()

    def get_donnees(self):
        """Retourne (cle, date, sortie, dept, archive: bool) — 5 valeurs depuis v4."""
        nom     = self.champ_nom.text().strip().upper()
        prenom  = self.champ_prenom.text().strip().upper()
        emp_id  = self.champ_id.text().strip()
        date    = self.champ_date.text().strip()
        sortie  = self.champ_sortie.text().strip()
        dept    = self.combo_dept.currentText().strip()
        archive = self.check_archive.isChecked()
        cle     = f"{nom} {prenom}|{emp_id}".strip()
        return cle, date, sortie, dept, archive


# =====================================================
# ONGLET EMPLOYÉS (v2 — format actif + archivage)
# =====================================================
class OngletEmployes(QWidget):
    def __init__(self):
        super().__init__()
        self.data = {}
        self._modifie = False   # True dès qu'une modification non sauvegardée existe
        self._construire_ui()
        self.charger_donnees()

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # --- En-tête ligne 1 : titre + radios statut ---
        entete = QHBoxLayout()
        titre = QLabel("👥  Gestion des employés")
        titre.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {COULEURS['texte']};")
        entete.addWidget(titre)
        entete.addStretch()

        self.radio_actifs   = QRadioButton("Actifs")
        self.radio_sortis   = QRadioButton("Sortis")
        self.radio_archives = QRadioButton("Archivés")
        self.radio_tous     = QRadioButton("Tous")
        self.radio_tous.setChecked(True)
        for r in [self.radio_actifs, self.radio_sortis, self.radio_archives, self.radio_tous]:
            r.setStyleSheet(f"color: {COULEURS['texte']}; font-size: 12px;")
            r.toggled.connect(self._appliquer_filtre)
            entete.addWidget(r)

        layout.addLayout(entete)

        # --- Ligne 2 : recherche + filtre département ---
        barre_filtres = QHBoxLayout()

        self.champ_recherche = QLineEdit()
        self.champ_recherche.setPlaceholderText("🔍  Rechercher par nom ou matricule…")
        self.champ_recherche.setFixedWidth(280)
        self.champ_recherche.textChanged.connect(self._appliquer_filtre)
        barre_filtres.addWidget(self.champ_recherche)

        barre_filtres.addSpacing(12)

        from PyQt6.QtWidgets import QLabel as _QLabel
        lbl_dept = _QLabel("Département :")
        lbl_dept.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        barre_filtres.addWidget(lbl_dept)

        self.combo_filtre_dept = QComboBox()
        self.combo_filtre_dept.addItem("Tous")
        self.combo_filtre_dept.addItems(DEPARTEMENTS_LISTE)
        self.combo_filtre_dept.setFixedWidth(160)
        self.combo_filtre_dept.currentIndexChanged.connect(self._appliquer_filtre)
        barre_filtres.addWidget(self.combo_filtre_dept)

        barre_filtres.addStretch()
        layout.addLayout(barre_filtres)

        # --- Compteur ---
        self.lbl_compteur = QLabel("")
        self.lbl_compteur.setStyleSheet(
            f"color: {COULEURS['texte_secondaire']}; font-size: 12px; padding: 2px 0;"
        )
        layout.addWidget(self.lbl_compteur)

        # --- Tableau ---
        self.tableau = QTableWidget()
        self.tableau.setColumnCount(4)
        self.tableau.setHorizontalHeaderLabels([
            "Nom | Prénom", "Matricule", "Date début contrat", "Statut"
        ])
        self.tableau.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tableau.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.tableau.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableau.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableau.verticalHeader().setVisible(False)
        self.tableau.setAlternatingRowColors(True)
        self.tableau.setStyleSheet(self.tableau.styleSheet() + f"""
            QTableWidget {{ alternate-background-color: {COULEURS['bg_carte']}; }}
        """)
        layout.addWidget(self.tableau)

        # --- Boutons ---
        boutons = QHBoxLayout()
        btn_ajouter   = QPushButton("＋  Ajouter")
        btn_modifier  = QPushButton("✏️  Modifier")
        btn_archiver  = QPushButton("📦  Archiver")
        btn_archiver.setStyleSheet(
            f"background-color: {COULEURS['accent_warning']}; color: #1E1E2E;"
        )
        btn_supprimer = QPushButton("🗑  Supprimer")
        btn_supprimer.setStyleSheet(f"background-color: {COULEURS['accent_danger']};")
        btn_sauvegarder = QPushButton("💾  Sauvegarder")
        btn_sauvegarder.setStyleSheet(
            f"background-color: {COULEURS['accent_succes']}; color: #1E1E2E;"
        )

        btn_ajouter.clicked.connect(self.ajouter_employe)
        btn_modifier.clicked.connect(self.modifier_employe)
        btn_archiver.clicked.connect(self.archiver_employe)
        self.tableau.itemSelectionChanged.connect(self._maj_bouton_archiver)
        self.tableau.itemDoubleClicked.connect(lambda _: self.modifier_employe())
        self.btn_archiver = btn_archiver
        btn_supprimer.clicked.connect(self.supprimer_employe)
        btn_sauvegarder.clicked.connect(self.sauvegarder_donnees)

        boutons.addWidget(btn_ajouter)
        boutons.addWidget(btn_modifier)
        boutons.addWidget(btn_archiver)
        boutons.addWidget(btn_supprimer)
        boutons.addStretch()
        boutons.addWidget(btn_sauvegarder)
        layout.addLayout(boutons)

    # --------------------------------------------------
    # Chargement & affichage
    # --------------------------------------------------

    def _marquer_modifie(self):
        """Marque l'onglet comme ayant des modifications non sauvegardées."""
        self._modifie = True

    def charger_donnees(self):
        self.data = charger_json(EMPLOYES_CONTRATS_JSON())
        self._modifie = False
        self._remplir_tableau()

    def _remplir_tableau(self):
        """Remplit le tableau depuis self.data, trie alphabetiquement. 3 statuts."""
        self.tableau.setRowCount(0)
        recherche   = self.champ_recherche.text().strip().lower()
        filtre_dept = self.combo_filtre_dept.currentText()

        nb_actifs   = 0
        nb_sortis   = 0
        nb_archives = 0

        cles_triees = sorted(
            [c for c in self.data if "|" in c],
            key=lambda c: c.split("|")[0]
        )

        for cle in cles_triees:
            info = self.data[cle]
            if isinstance(info, str):
                info = {"date_debut": info, "actif": True, "departements": []}

            actif       = info.get("actif", True)
            date_sortie = info.get("date_sortie", "").strip()
            if actif:
                statut = "actif"
                nb_actifs += 1
            elif date_sortie:
                statut = "sorti"
                nb_sortis += 1
            else:
                statut = "archive"
                nb_archives += 1

            # Filtre radio statut
            if self.radio_actifs.isChecked()   and statut != "actif":
                continue
            if self.radio_sortis.isChecked()   and statut != "sorti":
                continue
            if self.radio_archives.isChecked() and statut != "archive":
                continue
            # radio_tous : tout passe

            # Filtre département
            if filtre_dept != "Tous":
                depts_emp = [
                    d.get("departement", "") for d in info.get("departements", [])
                    if isinstance(d, dict)
                ]
                if filtre_dept not in depts_emp:
                    continue

            # Filtre recherche : nom OU matricule (partiel, sans zéros en tête obligatoires)
            nom_id = cle.split("|")
            nom    = nom_id[0].strip()
            emp_id = nom_id[1].strip() if len(nom_id) > 1 else ""
            if recherche:
                # Pour le matricule : chercher la séquence dans la version sans zéros en tête
                id_sans_zeros = emp_id.lstrip("0") or "0"
                recherche_sans_zeros = recherche.lstrip("0") or recherche
                match_nom = recherche in nom.lower()
                match_id  = (recherche in emp_id) or (recherche_sans_zeros in id_sans_zeros)
                if not match_nom and not match_id:
                    continue

            date_debut = info.get("date_debut", "")

            row = self.tableau.rowCount()
            self.tableau.insertRow(row)

            item_nom = QTableWidgetItem(nom)
            item_nom.setData(Qt.ItemDataRole.UserRole, cle)
            self.tableau.setItem(row, 0, item_nom)
            self.tableau.setItem(row, 1, QTableWidgetItem(emp_id))
            self.tableau.setItem(row, 2, QTableWidgetItem(date_debut))

            # Badge statut
            if statut == "actif":
                item_statut = QTableWidgetItem("Actif")
                item_statut.setForeground(QColor(COULEURS['accent_succes']))
            elif statut == "sorti":
                item_statut = QTableWidgetItem(f"Sorti le {date_sortie}")
                item_statut.setForeground(QColor(COULEURS['accent_warning']))
                for col in range(3):
                    it = self.tableau.item(row, col)
                    if it:
                        it.setForeground(QColor(COULEURS['accent_warning']))
            else:
                item_statut = QTableWidgetItem("Archivé")
                item_statut.setForeground(QColor(COULEURS['texte_secondaire']))
                for col in range(3):
                    it = self.tableau.item(row, col)
                    if it:
                        it.setForeground(QColor(COULEURS['texte_secondaire']))

            self.tableau.setItem(row, 3, item_statut)

        parties = [f"{nb_actifs} actif(s)"]
        if nb_sortis:
            parties.append(f"{nb_sortis} sorti(s)")
        if nb_archives:
            parties.append(f"{nb_archives} archivé(s)")
        self.lbl_compteur.setText("  •  ".join(parties))

    def _appliquer_filtre(self):
        self._remplir_tableau()

    # --------------------------------------------------
    # Sélection
    # --------------------------------------------------

    def _get_cle_selectionnee(self):
        row = self.tableau.currentRow()
        if row < 0:
            return None
        item = self.tableau.item(row, 0)
        if not item:
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def _get_info_selectionnee(self):
        cle = self._get_cle_selectionnee()
        if not cle or cle not in self.data:
            return None, None
        info = self.data[cle]
        # Si ancien format string → convertir ET réécrire dans self.data
        # pour que les modifications suivantes (archiver, modifier) fonctionnent
        if isinstance(info, str):
            info = {"date_debut": info, "actif": True, "departements": []}
            self.data[cle] = info  # ← correction : réécrire dans self.data
        return cle, info

    # --------------------------------------------------
    # Actions
    # --------------------------------------------------

    def ajouter_employe(self):
        dlg = DialogueEmploye(self)
        while True:
            if dlg.exec() != QDialog.DialogCode.Accepted:
                return  # annulation volontaire

            cle, date, sortie, dept, archive = dlg.get_donnees()

            # ── Champs obligatoires ──────────────────────────────
            nom_part = cle.split("|")[0].strip() if "|" in cle else ""
            id_part  = cle.split("|")[1].strip() if "|" in cle else ""
            if not nom_part:
                QMessageBox.warning(dlg, "Champ manquant", "Le nom et prénom sont obligatoires.")
                continue
            if not id_part:
                QMessageBox.warning(dlg, "Champ manquant", "Le matricule (ID) est obligatoire.")
                continue
            if not date:
                QMessageBox.warning(dlg, "Champ manquant", "La date d'entrée est obligatoire.")
                continue

            # ── Doublon clé exacte (même nom + même matricule) ───
            if cle in self.data:
                QMessageBox.warning(dlg, "Doublon",
                                    f"Un employé avec le nom « {nom_part} » et le matricule "
                                    f"« {id_part} » existe déjà.")
                continue

            # ── Même nom+prénom, matricule différent ─────────────
            doublon_nom = False
            for cle_ex in self.data:
                if "|" not in cle_ex:
                    continue
                if cle_ex.split("|")[0].strip() == nom_part:
                    QMessageBox.warning(dlg, "Doublon nom",
                                        f"Un employé avec le nom « {nom_part} » existe déjà "
                                        f"(matricule {cle_ex.split('|')[1].strip()}).\n"
                                        f"Vérifiez qu'il ne s'agit pas du même employé.")
                    doublon_nom = True
                    break
            if doublon_nom:
                continue

            # ── Même matricule, nom différent ─────────────────────
            doublon_mat = False
            for cle_ex, info_ex in self.data.items():
                if "|" not in cle_ex:
                    continue
                id_ex = cle_ex.split("|")[1].strip()
                if id_ex != id_part:
                    continue
                nom_ex   = cle_ex.split("|")[0].strip()
                actif_ex = info_ex.get("actif", True) if isinstance(info_ex, dict) else True
                if actif_ex:
                    QMessageBox.warning(dlg, "Matricule déjà utilisé",
                                        f"Le matricule « {id_part} » est déjà attribué à "
                                        f"« {nom_ex} » (actif).\n"
                                        f"Impossible d'ajouter deux employés actifs avec le même matricule.")
                    doublon_mat = True
                else:
                    info_ex_dict = info_ex if isinstance(info_ex, dict) else {}
                    statut_ex = "sorti" if info_ex_dict.get("date_sortie") else "archivé"
                    rep = QMessageBox.question(
                        dlg, "Matricule déjà utilisé",
                        f"\u26a0\ufe0f  Le matricule « {id_part} » est déjà attribué à\n"
                        f"« {nom_ex} » ({statut_ex}).\n\n"
                        f"Voulez-vous quand même créer cet employé ?",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                    )
                    if rep != QMessageBox.StandardButton.Yes:
                        doublon_mat = True
                break
            if doublon_mat:
                continue

            # ── Création ──────────────────────────────────────────
            departements = []
            if dept:
                departements = [{"departement": dept, "debut": date, "fin": None}]
            entree = {
                "date_debut":   date,
                "actif":        not archive,
                "departements": departements,
            }
            if sortie:
                entree["date_sortie"] = sortie
                entree["actif"] = False
            self.data[cle] = entree
            self._remplir_tableau()
            self._status(f"✅  Employé ajouté : {nom_part}")
            self._marquer_modifie()
            break

    def modifier_employe(self):
        cle, info = self._get_info_selectionnee()
        if not cle:
            QMessageBox.warning(self, "Attention", "Selectionnez un employe a modifier.")
            return

        dept_actuel = ""
        for d in reversed(info.get("departements", [])):
            if d.get("fin") is None:
                dept_actuel = d.get("departement", "")
                break

        # Détecter si archivé : inactif sans date_sortie
        est_archive = (
            not info.get("actif", True) and
            not info.get("date_sortie", "").strip()
        )

        dlg = DialogueEmploye(self, employe={
            "cle":         cle,
            "date":        info.get("date_debut", ""),
            "date_sortie": info.get("date_sortie", ""),
            "departement": dept_actuel,
            "archive":     est_archive,
        })
        if dlg.exec() == QDialog.DialogCode.Accepted:
            nouvelle_cle, nouvelle_date, nouvelle_sortie, nouveau_dept, nouveau_archive = dlg.get_donnees()
            if not nouvelle_cle or not nouvelle_date:
                QMessageBox.warning(self, "Champs manquants",
                                    "Le nom, prenom, matricule et date sont obligatoires.")
                return

            info["date_debut"] = nouvelle_date

            # Gestion statut : archive > sortie > actif
            if nouveau_archive and not nouvelle_sortie:
                info.pop("date_sortie", None)
                info["actif"] = False
            elif nouvelle_sortie:
                info["date_sortie"] = nouvelle_sortie
                info["actif"] = False
            else:
                # Ni archivé ni date sortie → réactivation
                info.pop("date_sortie", None)
                info["actif"] = True

            if nouveau_dept != dept_actuel:
                for d in info.get("departements", []):
                    if d.get("fin") is None:
                        d["fin"] = nouvelle_date
                if nouveau_dept:
                    info.setdefault("departements", []).append({
                        "departement": nouveau_dept,
                        "debut": nouvelle_date,
                        "fin": None,
                    })

            if nouvelle_cle != cle:
                del self.data[cle]
                self.data[nouvelle_cle] = info
            else:
                self.data[cle] = info

            self._remplir_tableau()
            self._marquer_modifie()

    def _maj_bouton_archiver(self):
        """Met a jour le libelle du bouton selon statut actif/sorti/archive."""
        cle, info = self._get_info_selectionnee()
        if not cle or not info:
            self.btn_archiver.setText("🚪  Sortie / Archiver")
            return
        actif       = info.get("actif", True)
        date_sortie = info.get("date_sortie", "").strip()
        if actif:
            self.btn_archiver.setText("🚪  Enregistrer une sortie")
            self.btn_archiver.setStyleSheet(
                f"background-color: {COULEURS['accent_warning']}; color: #1E1E2E;"
            )
        elif date_sortie:
            self.btn_archiver.setText("♻️  Reactivier (annuler sortie)")
            self.btn_archiver.setStyleSheet(
                f"background-color: {COULEURS['accent_succes']}; color: #1E1E2E;"
            )
        else:
            self.btn_archiver.setText("♻️  Reactivier")
            self.btn_archiver.setStyleSheet(
                f"background-color: {COULEURS['accent_succes']}; color: #1E1E2E;"
            )

    def archiver_employe(self):
        """Archive, enregistre une sortie, ou reactive selon l'etat actuel."""
        cle, info = self._get_info_selectionnee()
        if not cle:
            QMessageBox.warning(self, "Attention", "Selectionnez un employe.")
            return
        nom         = cle.split("|")[0]
        actif       = info.get("actif", True)
        date_sortie = info.get("date_sortie", "").strip()

        if actif:
            # Ouvrir le dialogue Modifier pre-rempli pour saisir la date de sortie
            dept_actuel = ""
            for d in reversed(info.get("departements", [])):
                if d.get("fin") is None:
                    dept_actuel = d.get("departement", "")
                    break
            dlg = DialogueEmploye(self, employe={
                "cle":         cle,
                "date":        info.get("date_debut", ""),
                "date_sortie": "",
                "departement": dept_actuel,
            })
            dlg.setWindowTitle(f"Enregistrer la sortie de {nom}")
            # Focus sur le champ sortie
            dlg.champ_sortie.setFocus()
            if dlg.exec() == QDialog.DialogCode.Accepted:
                _, nouvelle_date, nouvelle_sortie, nouveau_dept, _ = dlg.get_donnees()
                if not nouvelle_sortie:
                    QMessageBox.warning(self, "Date manquante",
                        "Veuillez renseigner la date de sortie.")
                    return
                info["date_debut"] = nouvelle_date
                info["date_sortie"] = nouvelle_sortie
                info["actif"] = False
                self.data[cle] = info
                self._remplir_tableau()
                self._marquer_modifie()
        else:
            # Reactiver (efface la date de sortie et repasse actif)
            label = f"sorti le {date_sortie}" if date_sortie else "archive"
            rep = QMessageBox.question(
                self, "Reactiver l'employe",
                f"Reactiver {nom} ({label}) ?\n\n"
                f"L'employe sera de nouveau marque comme actif.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if rep == QMessageBox.StandardButton.Yes:
                info.pop("date_sortie", None)
                info["actif"] = True
                self.data[cle] = info
                self._remplir_tableau()
                self._marquer_modifie()

    def supprimer_employe(self):
        cle, info = self._get_info_selectionnee()
        if not cle:
            QMessageBox.warning(self, "Attention", "Sélectionnez un employé à supprimer.")
            return
        nom = cle.split("|")[0]
        rep = QMessageBox.question(
            self, "Suppression définitive",
            f"⚠️  Supprimer définitivement {nom} ?\n\n"
            f"Cette action est IRRÉVERSIBLE.\n"
            f"L'employé sera supprimé de :\n"
            f"  • employes_contrats.json\n"
            f"  • cycles_employes.json\n"
            f"  • planning_historique.json",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if rep != QMessageBox.StandardButton.Yes:
            return

        # Suppression cascade
        erreurs = []

        # 1. employes_contrats.json (déjà en mémoire)
        del self.data[cle]

        # 2. cycles_employes.json
        try:
            cycles_emp = charger_json(CYCLES_EMPLOYES_JSON())
            if cle in cycles_emp:
                del cycles_emp[cle]
                sauvegarder_json(CYCLES_EMPLOYES_JSON(), cycles_emp)
        except Exception as e:
            erreurs.append(f"cycles_employes.json : {e}")

        # 3. planning_historique.json
        try:
            chemin_planning = PLANNING_HISTORIQUE_JSON()
            planning = charger_json(chemin_planning)
            if cle in planning:
                del planning[cle]
                sauvegarder_json(chemin_planning, planning)
        except Exception as e:
            erreurs.append(f"planning_historique.json : {e}")

        self._remplir_tableau()

        if erreurs:
            QMessageBox.warning(
                self, "Suppression partielle",
                f"Employé supprimé de employes_contrats.json.\n"
                f"Erreurs sur les autres fichiers :\n" + "\n".join(erreurs)
            )
        else:
            QMessageBox.information(
                self, "Suppression effectuée",
                f"✅ {nom} supprimé de tous les fichiers JSON."
            )
            self._status(f"🗑  {nom} supprimé de tous les fichiers.")

    def _status(self, msg: str, duree: int = 4000):
        """Affiche un message temporaire dans la barre de statut principale."""
        try:
            self.window().status.showMessage(msg, duree)
        except Exception:
            pass

    def sauvegarder_donnees(self):
        try:
            # Proposer un backup avant sauvegarde
            if os.path.exists(EMPLOYES_CONTRATS_JSON()):
                rep = QMessageBox.question(
                    self, "Backup avant sauvegarde",
                    "Voulez-vous créer un backup avant de sauvegarder ?\n\n"
                    "Le fichier sera copié dans le dossier backup/.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if rep == QMessageBox.StandardButton.Yes:
                    chemin_backup = faire_backup(EMPLOYES_CONTRATS_JSON())
                    if chemin_backup:
                        QMessageBox.information(
                            self, "Backup créé",
                            f"✅ Backup créé :\nbackup/{os.path.basename(chemin_backup)}"
                        )

            sauvegarder_json(EMPLOYES_CONTRATS_JSON(), self.data)
            self._modifie = False
            self._status("💾  Employés sauvegardés avec succès.")
            QMessageBox.information(self, "Succès", "✅ Employés sauvegardés avec succès !")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", f"❌ Erreur : {e}")


# =====================================================
# ONGLET CYCLES DÉFINITIONS
# =====================================================

# =====================================================
# ONGLET ABSENCES
# =====================================================

# =====================================================
# WORKER THREAD — import planning en arrière-plan
# =====================================================