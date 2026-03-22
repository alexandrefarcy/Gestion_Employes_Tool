"""
ui/onglet_absences.py
=====================
OngletAbsences — visualisation absences + import JSON
"""

import json as _json
from datetime import date as _date, timedelta

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QLineEdit, QComboBox, QGroupBox, QFileDialog, QMessageBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from ui.fermetures import periodes_fermetures_annee
from ui.constantes import (
    COULEURS, ABSENCES_JSON,
    charger_json, sauvegarder_json, faire_backup,
)


class ComboSansScroll(QComboBox):
    """QComboBox sans défilement à la molette."""
    def wheelEvent(self, e):
        e.ignore()


class OngletAbsences(QWidget):
    """Onglet visualisation des absences — lecture seule + import JSON."""

    def __init__(self):
        super().__init__()
        self._donnees = {}
        self._construire_ui()
        self.charger_donnees()

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # ── Titre + bouton import ────────────────────────────────
        entete = QHBoxLayout()
        titre = QLabel("\U0001f4cb  Absences & Fermetures")
        titre.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {COULEURS['texte']};")
        entete.addWidget(titre)
        entete.addStretch()

        btn_import = QPushButton("\U0001f4e5  Importer absences_projections.json")
        btn_import.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['accent']};
                border-radius: 6px;
                padding: 6px 14px;
                font-size: 12px;
                font-weight: 600;
            }}
            QPushButton:hover {{ background-color: {COULEURS['accent']}; color: #FFFFFF; }}
        """)
        btn_import.clicked.connect(self._importer_json)
        entete.addWidget(btn_import)
        layout.addLayout(entete)

        # ── Fermetures obligatoires ──────────────────────────────
        grp_fermetures = QGroupBox("Fermetures obligatoires (toutes années)")
        grp_fermetures.setStyleSheet(
            f"QGroupBox {{ color: {COULEURS['accent_warning']}; "
            f"border: 1px solid {COULEURS['accent_warning']}; "
            f"border-radius: 6px; margin-top: 6px; padding: 8px; font-size: 12px; }}"
        )
        lay_f = QHBoxLayout(grp_fermetures)
        lbl_f = QLabel(
            "🏖️  <b>Été :</b> 3 semaines dès le 1er lundi semaine pleine d'août  "
            "🎄  <b>Hiver :</b> 24 décembre → 2 janvier  "
            "<i>(fermetures automatiques, incluses dans le décompte)</i>"
        )
        lbl_f.setTextFormat(Qt.TextFormat.RichText)
        lbl_f.setWordWrap(True)
        lbl_f.setStyleSheet(f"color: {COULEURS['texte']}; font-size: 12px;")
        lay_f.addWidget(lbl_f)
        lay_f.addStretch()
        layout.addWidget(grp_fermetures)

        # ── Filtres ──────────────────────────────────────────────
        filtres = QHBoxLayout()
        self.champ_recherche = QLineEdit()
        self.champ_recherche.setPlaceholderText("\U0001f50d  Rechercher un employ\u00e9...")
        self.champ_recherche.setFixedWidth(250)
        self.champ_recherche.textChanged.connect(self._filtrer)

        lbl_annee = QLabel("Ann\u00e9e :")
        lbl_annee.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.combo_annee = ComboSansScroll()
        self.combo_annee.addItem("Toutes", 0)
        for a in range(2021, 2031):
            self.combo_annee.addItem(str(a), a)
        self.combo_annee.currentIndexChanged.connect(self._filtrer)

        self.lbl_stats = QLabel("")
        self.lbl_stats.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")

        filtres.addWidget(self.champ_recherche)
        filtres.addSpacing(16)
        filtres.addWidget(lbl_annee)
        filtres.addWidget(self.combo_annee)
        filtres.addStretch()
        filtres.addWidget(self.lbl_stats)
        layout.addLayout(filtres)

        # ── Tableau ──────────────────────────────────────────────
        self.tableau = QTableWidget()
        self.tableau.setColumnCount(4)
        self.tableau.setHorizontalHeaderLabels([
            "Employ\u00e9", "Nb p\u00e9riodes", "Nb jours total", "P\u00e9riodes d'absence"
        ])
        hh = self.tableau.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.tableau.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableau.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableau.verticalHeader().setVisible(False)
        self.tableau.setAlternatingRowColors(True)
        self.tableau.setStyleSheet(f"""
            QTableWidget {{ alternate-background-color: {COULEURS['bg_carte']}; }}
        """)
        layout.addWidget(self.tableau)

    def charger_donnees(self):
        self._donnees = charger_json(ABSENCES_JSON())
        self._filtrer()

    def _parse_date(self, s: str):
        try:
            j, m, a = s.split("-")
            return _date(int(a), int(m), int(j))
        except Exception:
            return None

    def _nb_jours(self, periodes: list, annee_filtre: int = 0) -> int:
        total = 0
        for p in periodes:
            d = self._parse_date(p.get("debut", ""))
            f = self._parse_date(p.get("fin", ""))
            if not d or not f:
                continue
            if annee_filtre and d.year != annee_filtre and f.year != annee_filtre:
                continue
            cur = d
            while cur <= f:
                if cur.weekday() < 5:
                    total += 1
                cur += timedelta(days=1)
        return total

    def _filtrer_periodes(self, periodes: list, annee: int) -> list:
        if not annee:
            return periodes
        return [
            p for p in periodes
            if self._parse_date(p.get("debut", "")) and
               self._parse_date(p.get("debut", "")).year == annee or
               self._parse_date(p.get("fin", "")) and
               self._parse_date(p.get("fin", "")).year == annee
        ]

    def _periode_deja_couverte(self, fermeture: dict, periodes_perso: list) -> bool:
        """Retourne True si la fermeture est entièrement couverte par une absence personnelle."""
        d_f = self._parse_date(fermeture.get("debut", ""))
        f_f = self._parse_date(fermeture.get("fin", ""))
        if not d_f or not f_f:
            return False
        for p in periodes_perso:
            d_p = self._parse_date(p.get("debut", ""))
            f_p = self._parse_date(p.get("fin", ""))
            if d_p and f_p and d_p <= d_f and f_p >= f_f:
                return True
        return False

    def _nb_jours_sans_doublon(self, periodes_perso: list, fermetures: list, annee_filtre: int = 0) -> int:
        """Compte les jours ouvrés d'absence en fusionnant perso + fermetures, sans doublon."""
        jours = set()
        for p in periodes_perso + fermetures:
            d = self._parse_date(p.get("debut", ""))
            f = self._parse_date(p.get("fin", ""))
            if not d or not f:
                continue
            if annee_filtre and d.year != annee_filtre and f.year != annee_filtre:
                continue
            cur = d
            while cur <= f:
                if cur.weekday() < 5:
                    jours.add(cur)
                cur += _date.resolution
        return len(jours)

    def _filtrer(self):
        texte = self.champ_recherche.text().strip().lower()
        annee = self.combo_annee.currentData() or 0

        # Fermetures obligatoires pour l'année sélectionnée
        annees_fermetures = [annee] if annee else list(range(2021, 2031))
        fermetures_periodes = []
        for a in annees_fermetures:
            fermetures_periodes.extend(periodes_fermetures_annee(a))

        lignes = []
        for cle_emp, periodes in self._donnees.items():
            if not isinstance(periodes, list):
                continue
            nom = cle_emp.split("|")[0]
            if texte and texte not in nom.lower():
                continue
            periodes_filtrees = self._filtrer_periodes(periodes, annee)
            # Fusionner avec fermetures (pour affichage et comptage)
            toutes_periodes = periodes_filtrees + [
                p for p in fermetures_periodes
                if not self._periode_deja_couverte(p, periodes_filtrees)
            ]
            if annee and not toutes_periodes:
                continue
            nb_j = self._nb_jours_sans_doublon(periodes_filtrees, fermetures_periodes, annee)
            lignes.append((nom, len(toutes_periodes), nb_j, toutes_periodes))

        lignes.sort(key=lambda x: x[0])
        self.tableau.setRowCount(len(lignes))

        for row, (nom, nb_per, nb_j, periodes) in enumerate(lignes):
            item_nom = QTableWidgetItem(nom)
            item_nom.setForeground(QColor(COULEURS["texte"]))
            self.tableau.setItem(row, 0, item_nom)

            item_nb = QTableWidgetItem(str(nb_per))
            item_nb.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tableau.setItem(row, 1, item_nb)

            item_j = QTableWidgetItem(str(nb_j))
            item_j.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.tableau.setItem(row, 2, item_j)

            def fmt(p):
                d = p.get("debut", "")
                f = p.get("fin", "")
                return d if d == f else f"{d} \u2192 {f}"
            texte_per = "  |  ".join(fmt(p) for p in periodes[:20])
            if len(periodes) > 20:
                texte_per += f"  ...  (+{len(periodes)-20})"
            item_per = QTableWidgetItem(texte_per)
            item_per.setForeground(QColor(COULEURS["texte_secondaire"]))
            self.tableau.setItem(row, 3, item_per)

        total_j = sum(x[2] for x in lignes)
        self.lbl_stats.setText(
            f"{len(lignes)} employ\u00e9(s)  \u2014  {total_j} jours d'absence"
            + (f" en {annee}" if annee else " au total")
        )

    def _importer_json(self):
        chemin, _ = QFileDialog.getOpenFileName(
            self, "Importer absences_projections.json", "", "JSON (*.json)"
        )
        if not chemin:
            return
        try:
            with open(chemin, encoding="utf-8") as f:
                data = _json.load(f)
            if not isinstance(data, dict):
                raise ValueError("Format invalide — attendu un objet JSON")
        except Exception as e:
            QMessageBox.critical(self, "Erreur import", str(e))
            return

        faire_backup(ABSENCES_JSON())
        sauvegarder_json(ABSENCES_JSON(), data)
        self._donnees = data
        self._filtrer()
        QMessageBox.information(
            self, "Import r\u00e9ussi",
            f"\u2705  {len(data)} employ\u00e9s charg\u00e9s.\n"
            "\U0001f4be  Backup cr\u00e9\u00e9 automatiquement."
        )