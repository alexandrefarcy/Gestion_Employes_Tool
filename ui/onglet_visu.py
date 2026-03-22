"""
ui/onglet_visu.py
=================
DelegateHachureHypothetique + DialogueOverrideCellule
+ DialogueCorrectionCycle + OngletVisualisationPlanning
"""

from datetime import date as _date, timedelta

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QDialog, QDialogButtonBox,
    QCheckBox, QGroupBox, QDateEdit, QMessageBox,
    QRadioButton, QButtonGroup, QSpinBox,QStyledItemDelegate, QStyle, QComboBox, QLineEdit,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QDate
from PyQt6.QtGui import QColor, QFont, QPainter, QPen

from ui.fermetures import jours_fermetures_periode
from ui.constantes import (
    COULEURS, COULEURS_CYCLE, COULEURS_CYCLE_HYPO, COULEUR_TEXTE_CYCLE,
    COULEUR_VIDE, COULEUR_NULL, COULEUR_ABSENCE,
    PLANNING_HISTORIQUE_JSON, CYCLES_EMPLOYES_JSON,
    CYCLES_DEFINITIONS_JSON, EMPLOYES_CONTRATS_JSON, ABSENCES_JSON,
    charger_json, sauvegarder_json, faire_backup,
)
from ui.widgets import ComboSansScroll

from ui.onglet_planning import WorkerGenerationHyp 

class DelegateHachureHypothetique(QStyledItemDelegate):
    """
    Delegate pour OngletVisualisationPlanning.
    Cellules hypothétiques : fond hachuré diagonal + texte italique.
    Cellules réelles       : rendu standard (fond uni).
    Cellules mixtes        : fond coupé diagonalement (dominant + secondaire).
    """

    _HYPO_ROLE  = Qt.ItemDataRole.UserRole + 10
    _MIXTE_ROLE = Qt.ItemDataRole.UserRole + 11  # couleur secondaire (hex str) si cycle mixte

    def paint(self, painter, option, index):
        # ── Cellule mixte (fond coupé diagonal) ─────────────────
        couleur_secondaire = index.data(self._MIXTE_ROLE)
        hypo = index.data(self._HYPO_ROLE)

        if couleur_secondaire and not hypo:
            painter.save()
            bg_color = index.data(Qt.ItemDataRole.BackgroundRole)
            if bg_color and hasattr(bg_color, 'color'):
                couleur_dom = bg_color.color()
            else:
                couleur_dom = QColor("#3A3A50")
            couleur_sec = QColor(couleur_secondaire)
            r = option.rect
            # Triangle haut-gauche = dominant
            from PyQt6.QtGui import QPolygon
            from PyQt6.QtCore import QPoint
            painter.setBrush(couleur_dom)
            painter.setPen(Qt.PenStyle.NoPen)
            tri_dom = QPolygon([
                QPoint(r.left(), r.top()),
                QPoint(r.right(), r.top()),
                QPoint(r.left(), r.bottom()),
            ])
            painter.drawPolygon(tri_dom)
            # Triangle bas-droit = secondaire
            painter.setBrush(couleur_sec)
            tri_sec = QPolygon([
                QPoint(r.right(), r.top()),
                QPoint(r.right(), r.bottom()),
                QPoint(r.left(), r.bottom()),
            ])
            painter.drawPolygon(tri_sec)
            # Sélection
            if option.state & QStyle.StateFlag.State_Selected:
                sel = option.palette.highlight().color()
                sel.setAlpha(80)
                painter.fillRect(option.rect, sel)
            # Texte centré (cycle dominant)
            text = index.data(Qt.ItemDataRole.DisplayRole) or ""
            fg = index.data(Qt.ItemDataRole.ForegroundRole)
            if fg and hasattr(fg, 'color'):
                text_color = fg.color()
            else:
                text_color = QColor("#FFFFFF")
            painter.setPen(text_color)
            font = QFont(option.font)
            font.setBold(True)
            painter.setFont(font)
            painter.drawText(option.rect, Qt.AlignmentFlag.AlignCenter, text)
            painter.restore()
            return

        if not hypo:
            super().paint(painter, option, index)
            return

        painter.save()

        bg_color = index.data(Qt.ItemDataRole.BackgroundRole)
        if bg_color and hasattr(bg_color, 'color'):
            base = bg_color.color()
        else:
            base = QColor("#3A3A50")
        painter.fillRect(option.rect, base)

        hatch_color = QColor(255, 255, 255, 45)
        pen = painter.pen()
        pen.setColor(hatch_color)
        pen.setWidth(1)
        painter.setPen(pen)
        step = 6
        r = option.rect
        x0, y0, x1, y1 = r.left(), r.top(), r.right(), r.bottom()
        w, h = r.width(), r.height()
        for offset in range(0, w + h, step):
            ax = x0 + offset
            ay = y0
            bx = x0
            by = y0 + offset
            if ax > x1:
                ay += ax - x1
                ax = x1
            if by > y1:
                bx += by - y1
                by = y1
            painter.drawLine(ax, ay, bx, by)

        if option.state & QStyle.StateFlag.State_Selected:
            sel = option.palette.highlight().color()
            sel.setAlpha(80)
            painter.fillRect(option.rect, sel)

        text = index.data(Qt.ItemDataRole.DisplayRole) or ""
        fg = index.data(Qt.ItemDataRole.ForegroundRole)
        if fg and hasattr(fg, 'color'):
            text_color = fg.color()
        else:
            text_color = QColor("#CCCCCC")
        painter.setPen(text_color)
        font = QFont(option.font)
        font.setItalic(True)
        painter.setFont(font)
        painter.drawText(option.rect, Qt.AlignmentFlag.AlignCenter, text)

        painter.restore()



class DialogueOverrideCellule(QDialog):
    """Popup Hyp-E2 — 3 options sur cellule hypothétique.
    A : Correction ponctuelle (reste hyp:true, écrasable)
    B : Recalibrage à partir d'ici (même motif, nouvelle phase)
    C : Nouveau cycle à partir d'ici (nouveau motif)
    """

    POSTES = ["M", "AM", "N", "J", "WE", "R"]
    NB_MAX_SEMAINES = 5

    # Constantes de mode
    MODE_A = "ponctuel"
    MODE_B = "recalibrage"
    MODE_C = "nouveau_cycle"

    def __init__(self, cle_employe: str, cle_col: str, cycle_actuel: str,
                 motif_actuel: list = None, parent=None):
        super().__init__(parent)
        self.cle_employe   = cle_employe
        self.cle_col       = cle_col
        self.cycle_actuel  = cycle_actuel
        self.motif_actuel  = motif_actuel or []
        self._mode         = self.MODE_A
        self.setWindowTitle("Modifier une cellule hypothétique")
        self.setMinimumWidth(480)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {COULEURS['bg_secondaire']};
                color: {COULEURS['texte']};
            }}
            QLabel {{ color: {COULEURS['texte']}; font-size: 13px; }}
            QRadioButton {{ color: {COULEURS['texte']}; font-size: 13px; }}
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
                padding: 4px 8px;
                font-size: 13px;
            }}
            QGroupBox {{
                color: {COULEURS['texte_secondaire']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 6px;
                margin-top: 6px;
                padding: 10px;
                font-size: 12px;
            }}
        """)
        self._construire_ui()

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        # Titre
        titre = QLabel("\u270f\ufe0f  Modifier une cellule hypoth\u00e9tique")
        titre.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {COULEURS['accent']};")
        layout.addWidget(titre)

        # Infos employé / période
        nom = self.cle_employe.split("|")[0]
        lbl_ctx = QLabel(f"Employ\u00e9 : <b>{nom}</b> — P\u00e9riode : <b>{self.cle_col}</b> — Actuel : <b>{self.cycle_actuel}</b>")
        lbl_ctx.setTextFormat(Qt.TextFormat.RichText)
        lbl_ctx.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        layout.addWidget(lbl_ctx)

        # ── Choix du mode ────────────────────────────────────────
        grp_mode = QGroupBox("Type de modification")
        lay_mode = QVBoxLayout(grp_mode)
        lay_mode.setSpacing(6)

        self._grp_radios = QButtonGroup(self)

        rb_a = QRadioButton("\U0001f4cc  A — Correction ponctuelle  (reste hypothétique, sera écrasée à la prochaine génération)")
        rb_b = QRadioButton("\U0001f504  B — Recalibrer le cycle à partir d'ici  (même motif, nouvelle phase)")
        rb_c = QRadioButton("\U0001f195  C — Nouveau cycle à partir d'ici  (nouveau motif)")
        rb_a.setChecked(True)

        for rb in [rb_a, rb_b, rb_c]:
            rb.setStyleSheet(f"color: {COULEURS['texte']}; font-size: 12px;")
            self._grp_radios.addButton(rb)
            lay_mode.addWidget(rb)

        self._rb_a = rb_a
        self._rb_b = rb_b
        self._rb_c = rb_c
        self._grp_radios.buttonClicked.connect(self._on_mode_change)
        layout.addWidget(grp_mode)

        # ── Panneau A : choix du cycle pour cette cellule ────────
        self._pan_a = QGroupBox("Cycle pour cette cellule")
        lay_a = QHBoxLayout(self._pan_a)
        lay_a.setSpacing(8)
        lbl_a = QLabel("Cycle :")
        self._combo_a = ComboSansScroll()
        for p in self.POSTES:
            self._combo_a.addItem(p, p)
        if self.cycle_actuel in self.POSTES:
            self._combo_a.setCurrentIndex(self._combo_a.findData(self.cycle_actuel))
        lay_a.addWidget(lbl_a)
        lay_a.addWidget(self._combo_a)
        lay_a.addStretch()
        layout.addWidget(self._pan_a)

        # ── Panneau B : choix du cycle d'ancrage ─────────────────
        self._pan_b = QGroupBox("Recalibrage — poste de cette semaine")
        lay_b = QHBoxLayout(self._pan_b)
        lbl_b = QLabel("Cette semaine doit \u00eatre en :")
        self._combo_b = ComboSansScroll()
        for p in self.POSTES:
            self._combo_b.addItem(p, p)
        if self.cycle_actuel in self.POSTES:
            self._combo_b.setCurrentIndex(self._combo_b.findData(self.cycle_actuel))
        lbl_b2 = QLabel("Le reste du motif existant sera recalibr\u00e9 en cons\u00e9quence.")
        lbl_b2.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 11px;")
        lay_b.addWidget(lbl_b)
        lay_b.addWidget(self._combo_b)
        lay_b.addStretch()
        vb = QVBoxLayout(self._pan_b)
        vb.addLayout(lay_b)
        vb.addWidget(lbl_b2)
        layout.addWidget(self._pan_b)
        self._pan_b.setVisible(False)

        # ── Panneau C : nouveau motif (type Hyp-C) ───────────────
        self._pan_c = QGroupBox("Nouveau cycle (1 à 5 semaines)")
        lay_c_main = QVBoxLayout(self._pan_c)

        row_nb = QHBoxLayout()
        lbl_nb = QLabel("Nombre de semaines dans le cycle :")
        self._spin_c = QSpinBox()
        self._spin_c.setMinimum(1)
        self._spin_c.setMaximum(self.NB_MAX_SEMAINES)
        self._spin_c.setValue(max(1, len(self.motif_actuel)))
        self._spin_c.setFixedWidth(70)
        self._spin_c.valueChanged.connect(self._actualiser_combos_c)
        row_nb.addWidget(lbl_nb)
        row_nb.addWidget(self._spin_c)
        row_nb.addStretch()
        lay_c_main.addLayout(row_nb)

        row_combos = QHBoxLayout()
        row_combos.setSpacing(6)
        self._combos_c = []
        self._lbls_c   = []
        for i in range(self.NB_MAX_SEMAINES):
            col = QVBoxLayout()
            lbl = QLabel(f"S{i+1}")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 11px;")
            cb = ComboSansScroll()
            for p in self.POSTES:
                cb.addItem(p, p)
            if i < len(self.motif_actuel):
                idx = cb.findData(self.motif_actuel[i])
                if idx >= 0:
                    cb.setCurrentIndex(idx)
            col.addWidget(lbl)
            col.addWidget(cb)
            self._combos_c.append(cb)
            self._lbls_c.append(lbl)
            row_combos.addLayout(col)
        lay_c_main.addLayout(row_combos)

        self._lbl_apercu_c = QLabel()
        self._lbl_apercu_c.setStyleSheet(f"color: {COULEURS['accent_succes']}; font-size: 12px; font-weight: 600;")
        lay_c_main.addWidget(self._lbl_apercu_c)

        for cb in self._combos_c:
            cb.currentIndexChanged.connect(self._maj_apercu_c)

        self._actualiser_combos_c(self._spin_c.value())
        layout.addWidget(self._pan_c)
        self._pan_c.setVisible(False)

        # ── Boutons ──────────────────────────────────────────────
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

    def _on_mode_change(self, btn):
        self._mode = (self.MODE_A if btn == self._rb_a else
                      self.MODE_B if btn == self._rb_b else self.MODE_C)
        self._pan_a.setVisible(self._mode == self.MODE_A)
        self._pan_b.setVisible(self._mode == self.MODE_B)
        self._pan_c.setVisible(self._mode == self.MODE_C)
        self.adjustSize()

    def _actualiser_combos_c(self, nb: int):
        for i, (cb, lbl) in enumerate(zip(self._combos_c, self._lbls_c)):
            cb.setVisible(i < nb)
            lbl.setVisible(i < nb)
        self._maj_apercu_c()

    def _maj_apercu_c(self):
        motif = self.get_motif_c()
        if motif:
            self._lbl_apercu_c.setText("Motif : " + " \u2192 ".join(motif) + " \u2192 (r\u00e9p\u00e8te)")

    def get_mode(self) -> str:
        return self._mode

    def get_cycle_a(self) -> str:
        return self._combo_a.currentData()

    def get_cycle_b(self) -> str:
        return self._combo_b.currentData()

    def get_motif_c(self) -> list:
        nb = self._spin_c.value()
        return [self._combos_c[i].currentData() for i in range(nb)]


class DialogueCorrectionCycle(QDialog):
    """Popup correction manuelle du cycle détecté d'un employé."""

    def __init__(self, cle_employe: str, cycle_actuel: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Correction du cycle détecté")
        self.setMinimumWidth(380)
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
                padding: 6px 10px;
                font-size: 13px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                selection-background-color: {COULEURS['accent']};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(20, 20, 20, 20)

        titre = QLabel("\U0001f504  Correction manuelle du cycle")
        titre.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {COULEURS['accent']};")
        layout.addWidget(titre)

        lbl_emp = QLabel(f"Employé : <b>{cle_employe.split('|')[0]}</b>")
        lbl_emp.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(lbl_emp)

        lbl_actuel = QLabel(f"Cycle détecté actuel : <b>{cycle_actuel or '—'}</b>")
        lbl_actuel.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(lbl_actuel)

        lbl_choix = QLabel("Nouveau cycle :")
        layout.addWidget(lbl_choix)

        self.combo = QComboBox()
        self.combo.addItem("— Effacer (aucun cycle) —", "")
        for c in ["AM", "M", "N", "J", "WE", "R"]:
            self.combo.addItem(c, c)
        # Pré-sélectionner le cycle actuel
        if cycle_actuel:
            idx = self.combo.findData(cycle_actuel)
            if idx >= 0:
                self.combo.setCurrentIndex(idx)
        layout.addWidget(self.combo)

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
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_cycle(self) -> str:
        return self.combo.currentData()






class DialogueCycleCustomVisu(QDialog):
    """
    Dialogue création d'un cycle custom depuis la Visu.
    Déclenché par double-clic sur la colonne 'Cycle détecté'.
    Pré-remplit le motif depuis la meilleure détection (même score < 60%).
    L'utilisateur peut modifier le motif et nommer le cycle librement.
    """

    POSTES      = ["M", "AM", "N", "J", "WE", "R"]
    NB_MAX_SEM  = 6

    def __init__(self, cle_employe: str, motif_propose: list,
                 nom_propose: str = "", parent=None):
        super().__init__(parent)
        self.cle_employe   = cle_employe
        self._motif_init   = motif_propose or ["M"]
        self._nom_init     = nom_propose or self._generer_nom(motif_propose)
        self.setWindowTitle("Créer un cycle personnalisé")
        self.setMinimumWidth(520)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {COULEURS['bg_secondaire']};
                color: {COULEURS['texte']};
            }}
            QLabel  {{ color: {COULEURS['texte']}; font-size: 13px; }}
            QLineEdit {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 5px;
                padding: 6px 10px;
                font-size: 13px;
            }}
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
                padding: 4px 8px;
                font-size: 13px;
            }}
            QGroupBox {{
                color: {COULEURS['texte_secondaire']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 6px;
                margin-top: 6px;
                padding: 10px;
                font-size: 12px;
            }}
        """)
        self._construire_ui()

    @staticmethod
    def _generer_nom(motif: list) -> str:
        """Génère un nom automatique depuis le motif ex: ['M','J','M'] → 'Frag_MJM'."""
        if not motif:
            return "CUSTOM"
        return "Frag_" + "".join(motif)

    def _construire_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(20, 20, 20, 20)

        # Titre
        titre = QLabel("✏️  Créer un cycle personnalisé")
        titre.setStyleSheet(
            f"font-size: 15px; font-weight: 700; color: {COULEURS['accent']};"
        )
        layout.addWidget(titre)

        # Employé
        nom_emp = self.cle_employe.split("|")[0]
        lbl_emp = QLabel(f"Employé : <b>{nom_emp}</b>")
        lbl_emp.setTextFormat(Qt.TextFormat.RichText)
        lbl_emp.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        layout.addWidget(lbl_emp)

        # Nom / libellé du cycle
        grp_nom = QGroupBox("Nom du cycle")
        lay_nom = QHBoxLayout(grp_nom)
        lbl_n = QLabel("Libellé :")
        self._champ_nom = QLineEdit(self._nom_init)
        self._champ_nom.setPlaceholderText("ex: Frag_MJJM, CUSTOM_3S...")
        self._champ_nom.textChanged.connect(self._maj_apercu)
        lay_nom.addWidget(lbl_n)
        lay_nom.addWidget(self._champ_nom)
        layout.addWidget(grp_nom)

        # Motif
        grp_motif = QGroupBox("Motif du cycle (semaines)")
        lay_motif = QVBoxLayout(grp_motif)

        row_nb = QHBoxLayout()
        lbl_nb = QLabel("Nombre de semaines :")
        self._spin = QSpinBox()
        self._spin.setMinimum(1)
        self._spin.setMaximum(self.NB_MAX_SEM)
        self._spin.setValue(min(len(self._motif_init), self.NB_MAX_SEM) or 1)
        self._spin.setFixedWidth(70)
        self._spin.valueChanged.connect(self._actualiser_combos)
        row_nb.addWidget(lbl_nb)
        row_nb.addWidget(self._spin)
        row_nb.addStretch()
        lay_motif.addLayout(row_nb)

        # Combos semaines
        row_combos = QHBoxLayout()
        row_combos.setSpacing(8)
        self._combos = []
        self._lbls   = []
        for i in range(self.NB_MAX_SEM):
            col = QVBoxLayout()
            lbl = QLabel(f"S{i+1}")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet(
                f"color: {COULEURS['texte_secondaire']}; font-size: 11px; font-weight: 600;"
            )
            cb = ComboSansScroll()
            for p in self.POSTES:
                cb.addItem(p, p)
            # Pré-remplir depuis le motif proposé
            if i < len(self._motif_init):
                idx = cb.findData(self._motif_init[i])
                if idx >= 0:
                    cb.setCurrentIndex(idx)
            cb.currentIndexChanged.connect(self._maj_apercu)
            col.addWidget(lbl)
            col.addWidget(cb)
            self._combos.append(cb)
            self._lbls.append(lbl)
            row_combos.addLayout(col)
        row_combos.addStretch()
        lay_motif.addLayout(row_combos)

        # Aperçu dynamique
        self._lbl_apercu = QLabel()
        self._lbl_apercu.setStyleSheet(
            f"color: {COULEURS['accent_succes']}; font-size: 12px; font-weight: 600;"
        )
        lay_motif.addWidget(self._lbl_apercu)
        layout.addWidget(grp_motif)

        # Info proposition script
        lbl_info = QLabel(
            "💡  Le motif ci-dessus est proposé par le script depuis les données réelles. "
            "Vous pouvez le modifier librement."
        )
        lbl_info.setWordWrap(True)
        lbl_info.setStyleSheet(
            f"color: {COULEURS['texte_secondaire']}; font-size: 11px; font-style: italic;"
        )
        layout.addWidget(lbl_info)

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
        btns.button(QDialogButtonBox.StandardButton.Ok).setText("✅  Créer le cycle")
        btns.accepted.connect(self._valider)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

        # Init affichage
        self._actualiser_combos(self._spin.value())

    def _actualiser_combos(self, nb: int):
        for i, (cb, lbl) in enumerate(zip(self._combos, self._lbls)):
            visible = i < nb
            cb.setVisible(visible)
            lbl.setVisible(visible)
        # Auto-update nom si l'utilisateur n'a pas encore modifié
        motif = self._get_motif()
        nom_auto = self._generer_nom(motif)
        if self._champ_nom.text().startswith("Frag_") or self._champ_nom.text() == "CUSTOM":
            self._champ_nom.setText(nom_auto)
        self._maj_apercu()

    def _maj_apercu(self):
        motif = self._get_motif()
        if motif:
            self._lbl_apercu.setText(
                "Motif : " + " → ".join(motif) + " → (répète)"
            )

    def _get_motif(self) -> list:
        nb = self._spin.value()
        return [self._combos[i].currentData() for i in range(nb) if i < len(self._combos)]

    def _valider(self):
        nom = self._champ_nom.text().strip()
        if not nom:
            QMessageBox.warning(self, "Nom manquant", "Veuillez saisir un nom pour le cycle.")
            return
        motif = self._get_motif()
        if not motif:
            QMessageBox.warning(self, "Motif vide", "Le motif ne peut pas être vide.")
            return
        self.accept()

    def get_nom(self) -> str:
        return self._champ_nom.text().strip()

    def get_motif(self) -> list:
        return self._get_motif()


class OngletVisualisationPlanning(QWidget):
    """
    Phase 4 — Grille de visualisation planning.
    Lignes = employés (alpha), colonnes = semaines ou jours.
    Switcher semaine/jour, filtres, correction manuelle cycle.
    """

    def __init__(self):
        super().__init__()
        self._mode = "semaine"   # "semaine" | "jour"
        self._donnees_planning  = {}
        self._donnees_employes  = {}
        self._donnees_cycles    = {}
        self._colonnes          = []   # liste triée des clés colonnes affichées
        self._construire_ui()

    # --------------------------------------------------
    # Construction UI
    # --------------------------------------------------
    def _construire_ui(self):
        from PyQt6.QtWidgets import QCheckBox as _CB
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        # ---- Titre ----
        titre = QLabel("\U0001f4c5  Visualisation du planning")
        titre.setStyleSheet(f"font-size: 18px; font-weight: 700; color: {COULEURS['texte']};")
        layout.addWidget(titre)

        # =====================================================
        # BARRE DE CONTRÔLES
        # =====================================================
        grp_ctrl = QGroupBox("Filtres et affichage")
        grp_ctrl.setStyleSheet(f"""
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
        """)
        ctrl_layout = QVBoxLayout(grp_ctrl)
        ctrl_layout.setSpacing(8)

        # Ligne 1 : switcher mode + période
        row1 = QHBoxLayout()

        # Switcher semaine / jour
        self.btn_mode_semaine = QPushButton("\U0001f4c6  Par semaine")
        self.btn_mode_jour    = QPushButton("\U0001f4c4  Par jour")
        for btn in [self.btn_mode_semaine, self.btn_mode_jour]:
            btn.setFixedHeight(30)
            btn.setFixedWidth(130)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {COULEURS['bg_carte']};
                    color: {COULEURS['texte_secondaire']};
                    border: 1px solid {COULEURS['bordure']};
                    border-radius: 5px;
                    font-size: 12px;
                    font-weight: 600;
                }}
                QPushButton:hover {{
                    border-color: {COULEURS['accent']};
                    color: {COULEURS['texte']};
                }}
            """)
        self.btn_mode_semaine.clicked.connect(lambda: self._changer_mode("semaine"))
        self.btn_mode_jour.clicked.connect(lambda: self._changer_mode("jour"))
        self._mettre_a_jour_style_mode()

        row1.addWidget(QLabel("Mode :"))
        row1.addWidget(self.btn_mode_semaine)
        row1.addWidget(self.btn_mode_jour)
        row1.addSpacing(20)

        # Période
        lbl_du = QLabel("Du :")
        lbl_du.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.date_debut = QDateEdit()
        self.date_debut.setDate(QDate(2021, 3, 5))
        self.date_debut.setDisplayFormat("dd/MM/yyyy")
        self.date_debut.setCalendarPopup(True)
        self.date_debut.setFixedWidth(120)

        lbl_au = QLabel("Au :")
        lbl_au.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.date_fin = QDateEdit()
        self.date_fin.setDate(QDate.currentDate())
        self.date_fin.setDisplayFormat("dd/MM/yyyy")
        self.date_fin.setCalendarPopup(True)
        self.date_fin.setFixedWidth(120)

        for w in [lbl_du, self.date_debut, lbl_au, self.date_fin]:
            row1.addWidget(w)

        row1.addStretch()
        btn_charger = QPushButton("\U0001f50d  Afficher")
        btn_charger.setFixedHeight(30)
        btn_charger.setFixedWidth(110)
        btn_charger.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['accent']};
                color: #FFFFFF;
                border-radius: 5px;
                font-size: 12px;
                font-weight: 700;
            }}
            QPushButton:hover {{ background-color: {COULEURS['accent_hover']}; }}
        """)
        btn_charger.clicked.connect(self.charger_donnees)
        row1.addWidget(btn_charger)
        ctrl_layout.addLayout(row1)

        # Ligne 2 : filtres cycle / hypothétique / département / recherche nom
        row2 = QHBoxLayout()

        lbl_cycle = QLabel("Cycle :")
        lbl_cycle.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.combo_filtre_cycle = QComboBox()
        self.combo_filtre_cycle.addItems(["Tous", "AM", "M", "N", "J", "WE", "R"])
        self.combo_filtre_cycle.setFixedWidth(90)
        self.combo_filtre_cycle.setStyleSheet(
            f"background-color: {COULEURS['bg_carte']}; color: {COULEURS['texte']};"
            f"border: 1px solid {COULEURS['bordure']}; border-radius: 5px; padding: 4px 8px;"
        )

        lbl_dept = QLabel("Département :")
        lbl_dept.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.combo_filtre_dept = QComboBox()
        self.combo_filtre_dept.addItems([
            "Tous", "Fabrication", "Conditionnement", "Administratif",
            "Maintenance", "Technique", "Qualité", "Informatique", "Intérim"
        ])
        self.combo_filtre_dept.setFixedWidth(140)
        self.combo_filtre_dept.setStyleSheet(self.combo_filtre_cycle.styleSheet())

        self.check_reel_seul = _CB("Réel uniquement (masquer hypothétiques)")
        self.check_reel_seul.setStyleSheet(
            f"color: {COULEURS['texte_secondaire']}; font-size: 12px;"
        )

        # Checkboxes masquage statut / département
        _style_cb = f"color: {COULEURS['texte_secondaire']}; font-size: 12px;"
        self.check_masquer_sortis   = _CB("Masquer Sortis")
        self.check_masquer_archives = _CB("Masquer Archivés")
        self.check_masquer_interim  = _CB("Masquer Intérim")
        for _cb in [self.check_masquer_sortis, self.check_masquer_archives, self.check_masquer_interim]:
            _cb.setStyleSheet(_style_cb)
            _cb.toggled.connect(self._remplir_tableau)

        lbl_nom = QLabel("Nom :")
        lbl_nom.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.champ_filtre_nom = QLineEdit()
        self.champ_filtre_nom.setPlaceholderText("Filtrer par nom…")
        self.champ_filtre_nom.setFixedWidth(160)
        self.champ_filtre_nom.setStyleSheet(
            f"background-color: {COULEURS['bg_carte']}; color: {COULEURS['texte']};"
            f"border: 1px solid {COULEURS['bordure']}; border-radius: 5px; padding: 4px 8px;"
        )
        self.champ_filtre_nom.textChanged.connect(self._filtrer_lignes)

        for w in [lbl_cycle, self.combo_filtre_cycle,
                  lbl_dept, self.combo_filtre_dept,
                  self.check_reel_seul,
                  self.check_masquer_sortis,
                  self.check_masquer_archives,
                  self.check_masquer_interim,
                  lbl_nom, self.champ_filtre_nom]:
            row2.addWidget(w)

        row2.addStretch()

        btn_reset = QPushButton("🔄  Réinitialiser")
        btn_reset.setFixedHeight(28)
        btn_reset.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte_secondaire']};
                border: 1px solid {COULEURS['bordure']};
                border-radius: 5px;
                font-size: 12px;
                padding: 4px 10px;
            }}
            QPushButton:hover {{
                border-color: {COULEURS['accent']};
                color: {COULEURS['accent']};
            }}
        """)
        btn_reset.clicked.connect(self._reinitialiser_filtres)
        row2.addWidget(btn_reset)
        ctrl_layout.addLayout(row2)

        # Ligne 3 : génération hypothétiques
        row3 = QHBoxLayout()

        lbl_hyp_info = QLabel("Génération hypothétiques :")
        lbl_hyp_info.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        row3.addWidget(lbl_hyp_info)

        lbl_du_hyp = QLabel("Du :")
        lbl_du_hyp.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.date_debut_hyp = QDateEdit()
        self.date_debut_hyp.setDate(QDate(2021, 1, 1))
        self.date_debut_hyp.setDisplayFormat("dd/MM/yyyy")
        self.date_debut_hyp.setCalendarPopup(True)
        self.date_debut_hyp.setFixedWidth(120)

        lbl_au_hyp = QLabel("Au :")
        lbl_au_hyp.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 12px;")
        self.date_fin_hyp = QDateEdit()
        self.date_fin_hyp.setDate(QDate.currentDate())
        self.date_fin_hyp.setDisplayFormat("dd/MM/yyyy")
        self.date_fin_hyp.setCalendarPopup(True)
        self.date_fin_hyp.setFixedWidth(120)

        for w in [lbl_du_hyp, self.date_debut_hyp, lbl_au_hyp, self.date_fin_hyp]:
            row3.addWidget(w)

        row3.addSpacing(12)

        self.btn_generer_hyp = QPushButton("\U00002728  Générer les hypothétiques")
        self.btn_generer_hyp.setFixedHeight(30)
        self.btn_generer_hyp.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['accent']};
                border-radius: 5px;
                font-size: 12px;
                font-weight: 600;
                padding: 0 14px;
            }}
            QPushButton:hover {{
                background-color: {COULEURS['accent']};
                color: #FFFFFF;
            }}
            QPushButton:disabled {{
                opacity: 0.5;
            }}
        """)
        self.btn_generer_hyp.clicked.connect(self._lancer_generation_hyp)
        row3.addWidget(self.btn_generer_hyp)

        row3.addSpacing(8)

        self.btn_purger_hyp = QPushButton("\U0001f5d1  Purger les hypoth\u00e9tiques")
        self.btn_purger_hyp.setFixedHeight(30)
        self.btn_purger_hyp.setStyleSheet(f"""
            QPushButton {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte']};
                border: 1px solid {COULEURS['accent_danger']};
                border-radius: 5px;
                font-size: 12px;
                font-weight: 600;
                padding: 0 14px;
            }}
            QPushButton:hover {{
                background-color: {COULEURS['accent_danger']};
                color: #FFFFFF;
            }}
            QPushButton:disabled {{
                opacity: 0.5;
            }}
        """)
        self.btn_purger_hyp.clicked.connect(self._purger_hypothetiques)
        row3.addWidget(self.btn_purger_hyp)

        row3.addStretch()
        ctrl_layout.addLayout(row3)
        layout.addWidget(grp_ctrl)

        # =====================================================
        # LÉGENDE
        # =====================================================
        row_legende = QHBoxLayout()
        row_legende.setSpacing(6)
        lbl_leg = QLabel("Légende :")
        lbl_leg.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 11px;")
        row_legende.addWidget(lbl_leg)
        for cycle, couleur in COULEURS_CYCLE.items():
            badge = QLabel(f"  {cycle}  ")
            badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
            badge.setFixedSize(36, 20)
            tc = COULEUR_TEXTE_CYCLE.get(cycle, "#FFFFFF")
            badge.setStyleSheet(
                f"background-color: {couleur}; color: {tc}; "
                f"border-radius: 3px; font-size: 11px; font-weight: 700;"
            )
            row_legende.addWidget(badge)
        # Hypothétique
        badge_hyp = QLabel("  AM  ")
        badge_hyp.setFixedSize(52, 20)
        badge_hyp.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge_hyp.setStyleSheet(
            f"background-color: {COULEURS_CYCLE_HYPO['AM']}; color: #CCCCCC;"
            f"border-radius: 3px; font-size: 11px; font-style: italic;"
        )
        lbl_hyp = QLabel("= hypothétique (fond assombri + italique)")
        lbl_hyp.setStyleSheet(f"color: {COULEURS['texte_secondaire']}; font-size: 11px;")
        row_legende.addWidget(badge_hyp)
        row_legende.addWidget(lbl_hyp)
        row_legende.addWidget(badge_hyp)
        row_legende.addWidget(lbl_hyp)
        row_legende.addStretch()
        layout.addLayout(row_legende)

        # =====================================================
        # GRILLE — double tableau (colonnes fixes + scroll)
        # =====================================================
        _style_tableau = f"""
            QTableWidget {{
                background-color: {COULEURS['bg_principal']};
                gridline-color: {COULEURS['bordure']};
                border: 1px solid {COULEURS['bordure']};
            }}
            QHeaderView::section {{
                background-color: {COULEURS['bg_carte']};
                color: {COULEURS['texte_secondaire']};
                border: none;
                border-right: 1px solid {COULEURS['bordure']};
                border-bottom: 1px solid {COULEURS['bordure']};
                padding: 4px 2px;
                font-size: 11px;
                font-weight: 600;
            }}
        """

        # Tableau gauche — colonnes fixes (Employé + Cycle détecté)
        self.tableau_fixe = QTableWidget()
        self.tableau_fixe.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableau_fixe.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableau_fixe.setAlternatingRowColors(False)
        self.tableau_fixe.verticalHeader().setVisible(False)
        self.tableau_fixe.verticalHeader().setDefaultSectionSize(34)
        self.tableau_fixe.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.tableau_fixe.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.tableau_fixe.setStyleSheet(_style_tableau + "QScrollBar:horizontal { height: 0px; }")
        self.tableau_fixe.cellDoubleClicked.connect(self._on_double_clic_cellule)

        # Tableau droit — colonnes planning (scrollable)
        self.tableau = QTableWidget()
        self.tableau.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tableau.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tableau.setAlternatingRowColors(False)
        self.tableau.verticalHeader().setVisible(False)
        self.tableau.horizontalHeader().setDefaultSectionSize(52)
        self.tableau.horizontalHeader().setMinimumSectionSize(38)
        self.tableau.verticalHeader().setDefaultSectionSize(34)
        self.tableau.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.tableau.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        self.tableau.setStyleSheet(_style_tableau)
        self.tableau.cellDoubleClicked.connect(self._on_double_clic_planning)
        self._delegate_hachure = DelegateHachureHypothetique()
        self.tableau.setItemDelegate(self._delegate_hachure)

        # Synchronisation scroll vertical
        self.tableau.verticalScrollBar().valueChanged.connect(
            self.tableau_fixe.verticalScrollBar().setValue
        )
        self.tableau_fixe.verticalScrollBar().valueChanged.connect(
            self.tableau.verticalScrollBar().setValue
        )

        # Conteneur côte à côte
        # tableau_fixe est wrappé dans un VBoxLayout avec un spacer bas
        # dont la hauteur = scrollbar horizontale du tableau — évite le décalage
        conteneur_fixe = QWidget()
        lay_fixe = QVBoxLayout(conteneur_fixe)
        lay_fixe.setContentsMargins(0, 0, 0, 0)
        lay_fixe.setSpacing(0)
        lay_fixe.addWidget(self.tableau_fixe)
        self._spacer_scrollbar = QWidget()
        self._spacer_scrollbar.setFixedHeight(
            self.tableau.horizontalScrollBar().sizeHint().height()
        )
        self._spacer_scrollbar.setStyleSheet(
            f"background-color: {COULEURS['bg_carte']};"
            f"border-top: 1px solid {COULEURS['bordure']};"
        )
        lay_fixe.addWidget(self._spacer_scrollbar)

        conteneur_grille = QWidget()
        lay_grille = QHBoxLayout(conteneur_grille)
        lay_grille.setContentsMargins(0, 0, 0, 0)
        lay_grille.setSpacing(0)
        lay_grille.addWidget(conteneur_fixe)
        lay_grille.addWidget(self.tableau, stretch=1)
        layout.addWidget(conteneur_grille)

        # Barre info bas
        self.lbl_info = QLabel("Cliquez sur 'Afficher' pour charger les données.")
        self.lbl_info.setStyleSheet(
            f"color: {COULEURS['texte_secondaire']}; font-size: 11px; font-style: italic;"
        )
        layout.addWidget(self.lbl_info)

    # --------------------------------------------------
    # Mode switcher
    # --------------------------------------------------
    def _reinitialiser_filtres(self):
        """Remet tous les filtres à leur valeur par défaut."""
        self.date_debut.setDate(QDate(2024, 1, 1))
        self.date_fin.setDate(QDate.currentDate())
        self.combo_filtre_cycle.setCurrentIndex(0)
        self.combo_filtre_dept.setCurrentIndex(0)
        self.check_reel_seul.setChecked(False)
        self.check_masquer_sortis.setChecked(False)
        self.check_masquer_archives.setChecked(False)
        self.check_masquer_interim.setChecked(False)
        self.champ_filtre_nom.clear()

    # --------------------------------------------------
    # Génération hypothétiques (Hyp-B)
    # --------------------------------------------------
    def _lancer_generation_hyp(self):
        """Lance la génération des hypothétiques dans un thread."""
        from datetime import date as _date

        d_debut = self.date_debut_hyp.date()
        d_fin   = self.date_fin_hyp.date()

        if d_fin < d_debut:
            QMessageBox.warning(self, "Plage invalide",
                                "La date de fin doit être après la date de début.")
            return

        rep = QMessageBox.question(
            self, "Générer les hypothétiques",
            f"Générer les entrées hypothétiques du {d_debut.toString('dd/MM/yyyy')} "
            f"au {d_fin.toString('dd/MM/yyyy')} ?\n\n"
            "Les données réelles existantes ne seront jamais écrasées.\n"
            "Un backup sera créé automatiquement.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if rep != QMessageBox.StandardButton.Yes:
            return

        # Backup avant génération
        faire_backup(PLANNING_HISTORIQUE_JSON())

        self.btn_generer_hyp.setEnabled(False)
        self.btn_generer_hyp.setText("⏳  Génération en cours…")

        date_debut_py = _date(d_debut.year(), d_debut.month(), d_debut.day())
        date_fin_py   = _date(d_fin.year(),   d_fin.month(),   d_fin.day())

        self._worker_hyp = WorkerGenerationHyp(date_debut_py, date_fin_py)
        self._worker_hyp.log_signal.connect(self._on_hyp_log)
        self._worker_hyp.fini_signal.connect(self._on_hyp_fini)
        self._worker_hyp.erreur_signal.connect(self._on_hyp_erreur)
        self._worker_hyp.start()

    def _on_hyp_log(self, msg: str):
        """Reçoit les logs du worker — affiche dans la statusbar."""
        try:
            fenetre = self.window()
            if hasattr(fenetre, '_status'):
                fenetre._status(msg)
        except Exception:
            pass

    def _on_hyp_fini(self, stats: dict):
        """Reçoit les stats de fin de génération."""
        self.btn_generer_hyp.setEnabled(True)
        self.btn_generer_hyp.setText("\u2728  Générer les hypothétiques")

        nb_emp  = stats.get("nb_employes", 0)
        nb_sem  = stats.get("nb_semaines", 0)
        nb_j    = stats.get("nb_jours", 0)
        nb_ign  = stats.get("nb_ignores", 0)
        erreurs = stats.get("erreurs", [])

        if erreurs:
            QMessageBox.warning(
                self, "Génération terminée avec erreurs",
                f"{nb_emp} employé(s) traité(s)\n"
                f"{nb_sem} semaines + {nb_j} jours générés\n\n"
                f"Erreurs :\n" + "\n".join(erreurs[:5]),
            )
        else:
            QMessageBox.information(
                self, "Génération terminée",
                f"✅  {nb_emp} employé(s) traité(s)\n"
                f"   📅  {nb_sem} semaines hypothétiques générées\n"
                f"   📄  {nb_j} jours hypothétiques générés\n"
                + (f"   ⏭️  {nb_ign} ignoré(s) (sans cycle défini)\n" if nb_ign else "")
                + "\n💾  Backup créé automatiquement.",
            )

        # Rafraîchir la visualisation
        self.charger_donnees()

    def _on_hyp_erreur(self, msg: str):
        """Reçoit une erreur fatale du worker."""
        self.btn_generer_hyp.setEnabled(True)
        self.btn_generer_hyp.setText("\u2728  Générer les hypothétiques")
        QMessageBox.critical(self, "Erreur génération hypothétiques", msg)

    def _purger_hypothetiques(self):
        """Supprime toutes les entrees hypothetique:true du planning_historique.json."""
        rep = QMessageBox.warning(
            self, "Purger les hypoth\u00e9tiques",
            "\u26a0\ufe0f  ATTENTION — Op\u00e9ration IRR\u00c9VERSIBLE\n\n"
            "Toutes les semaines et jours hypoth\u00e9tiques (hypothetique: true) "
            "seront supprim\u00e9s du planning.\n"
            "Les donn\u00e9es r\u00e9elles (hypothetique: false) sont conserv\u00e9es.\n\n"
            "Un backup sera cr\u00e9\u00e9 automatiquement avant la suppression.",
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if rep != QMessageBox.StandardButton.Ok:
            return

        chemin = PLANNING_HISTORIQUE_JSON()
        faire_backup(chemin)

        planning = charger_json(chemin)
        nb_sem = 0
        nb_j   = 0

        for cle_emp, emp in planning.items():
            if cle_emp == "COMMENTAIRE" or not isinstance(emp, dict):
                continue
            # Purger semaines hypothetiques
            semaines = emp.get("semaines", {})
            cles_hyp_sem = [k for k, v in semaines.items()
                            if isinstance(v, dict) and v.get("hypothetique") is True]
            for k in cles_hyp_sem:
                del semaines[k]
                nb_sem += 1
            # Purger jours hypothetiques
            jours = emp.get("jours", {})
            cles_hyp_j = [k for k, v in jours.items()
                          if isinstance(v, dict) and v.get("hypothetique") is True]
            for k in cles_hyp_j:
                del jours[k]
                nb_j += 1

        sauvegarder_json(chemin, planning)
        self.charger_donnees()

        QMessageBox.information(
            self, "Purge termin\u00e9e",
            f"\u2705  Purge effectu\u00e9e avec succ\u00e8s\n"
            f"   \U0001f4c5  {nb_sem} semaine(s) supprim\u00e9e(s)\n"
            f"   \U0001f4c4  {nb_j} jour(s) supprim\u00e9(s)\n"
            "\n\U0001f4be  Backup cr\u00e9\u00e9 automatiquement."
        )

    def _changer_mode(self, mode: str):
        self._mode = mode
        self._mettre_a_jour_style_mode()
        self.charger_donnees()

    def _mettre_a_jour_style_mode(self):
        style_actif = (
            f"background-color: {COULEURS['accent']}; color: #FFFFFF; "
            f"border-radius: 5px; font-size: 12px; font-weight: 700; "
            f"border: none;"
        )
        style_inactif = (
            f"background-color: {COULEURS['bg_carte']}; color: {COULEURS['texte_secondaire']}; "
            f"border: 1px solid {COULEURS['bordure']}; border-radius: 5px; "
            f"font-size: 12px; font-weight: 600;"
        )
        self.btn_mode_semaine.setStyleSheet(
            style_actif if self._mode == "semaine" else style_inactif
        )
        self.btn_mode_jour.setStyleSheet(
            style_actif if self._mode == "jour" else style_inactif
        )

    # --------------------------------------------------
    # Chargement des données
    # --------------------------------------------------
    def charger_donnees(self):
        """Charge les JSON et reconstruit la grille."""
        self._donnees_planning     = charger_json(PLANNING_HISTORIQUE_JSON())
        self._donnees_employes     = charger_json(EMPLOYES_CONTRATS_JSON())
        self._donnees_cycles       = charger_json(CYCLES_EMPLOYES_JSON())
        self._donnees_definitions  = charger_json(CYCLES_DEFINITIONS_JSON())
        self._donnees_absences     = charger_json(ABSENCES_JSON())
        # Pré-calculer les plages d'absence par employé en objets date
        self._cache_absences = self._construire_cache_absences()

        qd_debut = self.date_debut.date()
        qd_fin   = self.date_fin.date()
        from datetime import date as _date, timedelta
        d_debut = _date(qd_debut.year(), qd_debut.month(), qd_debut.day())
        d_fin   = _date(qd_fin.year(),   qd_fin.month(),   qd_fin.day())

        if self._mode == "semaine":
            self._colonnes = self._generer_colonnes_semaines(d_debut, d_fin)
        else:
            self._colonnes = self._generer_colonnes_jours(d_debut, d_fin)

        self._remplir_tableau()

    def _construire_cache_absences(self) -> dict:
        """
        Pré-calcule les absences par employé : set de dates ouvrées.
        Inclut absences individuelles + fermetures obligatoires (sans doublon).
        Retourne {cle_emp: set(dates ouvrées)}
        "__fermetures__" → set des fermetures seules (pour employés sans absences perso)
        """
        from datetime import date as _date, timedelta

        def parse(s):
            try:
                j, m, a = s.split("-")
                return _date(int(a), int(m), int(j))
            except Exception:
                return None

        def plage_vers_jours(d, f):
            jours = set()
            cur = d
            while cur <= f:
                if cur.weekday() < 5:
                    jours.add(cur)
                cur += timedelta(days=1)
            return jours

        # Calculer les fermetures sur la plage affichée
        qd = self.date_debut.date()
        qf = self.date_fin.date()
        d_debut_visu = _date(qd.year(), qd.month(), qd.day())
        d_fin_visu   = _date(qf.year(), qf.month(), qf.day())
        fermetures = jours_fermetures_periode(d_debut_visu, d_fin_visu)

        cache = {"__fermetures__": fermetures}

        # Absences individuelles + fermetures (union sans doublon)
        for cle_emp, periodes in self._donnees_absences.items():
            if not isinstance(periodes, list):
                continue
            jours_abs = set()
            for p in periodes:
                d = parse(p.get("debut", ""))
                f = parse(p.get("fin", ""))
                if d and f:
                    jours_abs |= plage_vers_jours(d, f)
            # Union avec fermetures — pas de doublon
            cache[cle_emp] = jours_abs | fermetures

        return cache

    def _est_absent(self, cle_emp: str, cle_col: str) -> bool:
        """
        Retourne True si l'employé est absent sur cette colonne (semaine ou jour).
        Cache = set de dates ouvrées (absences perso + fermetures obligatoires).
        """
        from datetime import date as _date, timedelta
        # Utiliser le cache de l'employé, ou les fermetures seules si pas d'entrée
        jours_abs = self._cache_absences.get(
            cle_emp,
            self._cache_absences.get("__fermetures__", set())
        )
        if not jours_abs:
            return False

        if self._mode == "semaine":
            # Semaine : absent si AU MOINS UN JOUR ouvré de la semaine est en absence
            try:
                parts = cle_col.split("_")
                lundi = _date.fromisocalendar(int(parts[1]), int(parts[0][1:]), 1)
            except Exception:
                return False
            for offset in range(5):  # lun-ven
                if (lundi + timedelta(days=offset)) in jours_abs:
                    return True
        else:
            # Jour : absent si le jour est dans le set
            try:
                jour = _date.fromisoformat(cle_col)
            except Exception:
                return False
            return jour in jours_abs
        return False

    def _generer_colonnes_semaines(self, d_debut, d_fin) -> list:
        """Génère la liste triée des clés semaines SNN_AAAA dans la période."""
        from datetime import timedelta
        cols = []
        semaines_vues = set()
        d = d_debut
        while d <= d_fin:
            iso = d.isocalendar()  # (year, week, weekday) — stdlib pure
            cle = f"S{iso[1]:02d}_{iso[0]}"
            if cle not in semaines_vues:
                semaines_vues.add(cle)
                cols.append(cle)
            d += timedelta(days=7)
        return cols

    def _generer_colonnes_jours(self, d_debut, d_fin) -> list:
        """Génère la liste des clés ISO jours AAAA-MM-JJ dans la période."""
        from datetime import timedelta
        cols = []
        d = d_debut
        while d <= d_fin:
            cols.append(d.strftime("%Y-%m-%d"))
            d += timedelta(days=1)
        return cols

    def _synthese_semaine(self, cle_emp: str, cle_sem: str):
        """
        Synthétise le cycle d'une semaine ISO depuis les données jours.
        Utilisé en mode semaine quand il n'y a pas de données 'semaines' directes.

        Règles :
          - Si sam/dim travaillé (WE) et lun-ven tous repos/absents → WE
          - Si 5 jours lun-ven du même cycle → ce cycle (fixe)
          - Si cycles mixtes lun-ven → (cycle_dominant, cycle_secondaire, hypothetique)
          - Si aucun jour réel → (None, None, None)

        Retourne (cycle_dominant, cycle_secondaire_ou_None, hypothetique: bool)
        """
        from collections import Counter
        from datetime import date as _d, timedelta

        # Calculer le lundi de la semaine ISO
        try:
            parts = cle_sem.split('_')
            annee = int(parts[1])
            num   = int(parts[0][1:])
            from datetime import datetime
            lundi = datetime.strptime(f"{annee}-W{num:02d}-1", "%G-W%V-%u").date()
        except Exception:
            return None, None, None

        data_emp = self._donnees_planning.get(cle_emp, {})
        jours_dict = data_emp.get('jours', {})

        # Collecter lun-ven et sam-dim
        cycles_semaine = []   # (cycle, hypothetique) lun-ven
        cycles_we      = []   # (cycle, hypothetique) sam-dim

        for offset in range(7):
            jour = lundi + timedelta(days=offset)
            cle_j = jour.strftime("%Y-%m-%d")
            entree = jours_dict.get(cle_j)
            if entree is None:
                continue
            c   = entree.get('cycle')
            hyp = entree.get('hypothetique', False)
            if c is None:
                continue
            if offset < 5:
                cycles_semaine.append((c, hyp))
            else:
                cycles_we.append((c, hyp))

        if not cycles_semaine and not cycles_we:
            return None, None, None

        # Séparer réels et hypothétiques
        cycles_sem_reels = [(c, h) for c, h in cycles_semaine if not h]
        cycles_we_reels  = [(c, h) for c, h in cycles_we if not h]

        # Données réelles disponibles ?
        a_reels = bool(cycles_sem_reels or cycles_we_reels)
        tous_hyp = not a_reels

        # Travailler sur les données réelles si dispo, sinon hypothétiques
        cycles_a_analyser_sem = cycles_sem_reels if a_reels else cycles_semaine
        cycles_a_analyser_we  = cycles_we_reels  if a_reels else cycles_we

        # Règle WE : pas de cycle semaine, weekend travaillé
        if not cycles_a_analyser_sem and cycles_a_analyser_we:
            cycles_we_vals = [c for c, _ in cycles_a_analyser_we if c and c != 'R']
            if cycles_we_vals:
                return 'WE', None, tous_hyp

        if not cycles_a_analyser_sem:
            return None, None, None

        # Filtrer les repos (R) — ne comptent pas pour déterminer le cycle
        cycles_travailles = [(c, h) for c, h in cycles_a_analyser_sem if c and c != 'R']

        if not cycles_travailles:
            # Que des repos en semaine → WE si weekend travaillé
            cycles_we_vals = [c for c, _ in cycles_a_analyser_we if c and c != 'R']
            if cycles_we_vals:
                return 'WE', None, tous_hyp
            return 'R', None, tous_hyp

        compteur = Counter(c for c, _ in cycles_travailles)
        dominant, nb_dom = compteur.most_common(1)[0]

        # Cycle unique ou très dominant (>= 4/5)
        if len(compteur) == 1 or nb_dom >= 4:
            secondaire = None
            if len(compteur) > 1:
                secondaire = compteur.most_common(2)[1][0]
            return dominant, secondaire, tous_hyp

        # Cycles mixtes → dominant + secondaire
        secondaire = compteur.most_common(2)[1][0]
        return dominant, secondaire, tous_hyp

    def _get_cycle_employe(self, cle_emp: str, cle_col: str):
        """
        Retourne (cycle, hypothetique) pour un employé et une colonne.

        Priorité en mode semaine :
          1. Synthèse depuis jours RÉELS (hypothetique:false) — priorité absolue
          2. Semaine réelle dans 'semaines' (hypothetique:false)
          3. Synthèse depuis jours hypothétiques
          4. Semaine hypothétique dans 'semaines'
          5. None si rien

        En mode jour : cherche dans 'jours' directement.
        """
        data_emp = self._donnees_planning.get(cle_emp, {})
        if self._mode == "semaine":
            # Priorité 1 : synthèse depuis jours réels
            cycle_dom, cycle_sec, hyp = self._synthese_semaine(cle_emp, cle_col)
            if cycle_dom is not None and not hyp:
                # Des jours réels existent pour cette semaine → priorité absolue
                return cycle_dom, False

            # Priorité 2 : semaine réelle dans 'semaines'
            entree_sem = data_emp.get("semaines", {}).get(cle_col)
            if entree_sem is not None and not entree_sem.get("hypothetique", True):
                return entree_sem.get("cycle"), False

            # Priorité 3 : synthèse depuis jours (même hypothétiques)
            if cycle_dom is not None:
                return cycle_dom, (hyp if hyp is not None else True)

            # Priorité 4 : semaine hypothétique dans 'semaines'
            if entree_sem is not None:
                return entree_sem.get("cycle"), entree_sem.get("hypothetique", True)

            return None, None
        else:
            entree = data_emp.get("jours", {}).get(cle_col)
            if entree is None:
                return None, None
            return entree.get("cycle"), entree.get("hypothetique", False)

    def _get_cycle_detecte(self, cle_emp: str) -> str:
        """
        Retourne le libelle du cycle pour la colonne 'Cycle detecte'.
        Priorite : cycle_type (ex: '3x8') > cycle_depart (ex: 'AM') > vide.
        """
        data = self._donnees_cycles.get(cle_emp, {})
        return data.get("cycle_type") or data.get("cycle") or data.get("cycle_depart") or ""

    def _get_dept_employe(self, cle_emp: str) -> str:
        """Retourne le département principal d'un employé (premier trouvé)."""
        data = self._donnees_employes.get(cle_emp, {})
        depts = data.get("departements", [])
        if not depts:
            return ""
        if isinstance(depts[0], dict):
            return depts[0].get("departement", "")
        return str(depts[0])

    # --------------------------------------------------
    # Construction du tableau
    # --------------------------------------------------
    def _remplir_tableau(self):
        """Reconstruit la grille complète à partir des données chargées."""
        from PyQt6.QtWidgets import QCheckBox as _CB

        # Fusionner les clés employés (planning + employes_contrats)
        cles_emp = sorted(set(self._donnees_employes.keys()) | set(
            k for k in self._donnees_planning.keys() if k != "COMMENTAIRE"
        ))

        # Filtrer les clés sans nom valide (artefacts legacy "|" seul)
        cles_emp = [c for c in cles_emp if c.split("|")[0].strip()]

        # Filtre département
        filtre_dept = self.combo_filtre_dept.currentText()
        if filtre_dept != "Tous":
            cles_emp = [c for c in cles_emp if self._get_dept_employe(c) == filtre_dept]

        # Filtres statut / intérim
        masquer_sortis   = self.check_masquer_sortis.isChecked()
        masquer_archives = self.check_masquer_archives.isChecked()
        masquer_interim  = self.check_masquer_interim.isChecked()

        if masquer_sortis or masquer_archives or masquer_interim:
            def _garder_employe(cle):
                info = self._donnees_employes.get(cle, {})
                # Statut
                actif       = info.get("actif", True)
                date_sortie = info.get("date_sortie", "").strip()
                if actif:
                    statut = "actif"
                elif date_sortie:
                    statut = "sorti"
                else:
                    statut = "archive"
                if masquer_sortis   and statut == "sorti":   return False
                if masquer_archives and statut == "archive": return False
                # Département intérim
                if masquer_interim:
                    dept = self._get_dept_employe(cle)
                    if dept == "Intérim":
                        return False
                return True
            cles_emp = [c for c in cles_emp if _garder_employe(c)]

        # Colonnes : "Nom" + "Cycle détecté" + colonnes planning
        nb_cols_fixes = 2
        nb_cols = nb_cols_fixes + len(self._colonnes)

        # ── Tableau fixe (colonnes Employé + Cycle détecté) ──
        self.tableau_fixe.clear()
        self.tableau_fixe.setRowCount(len(cles_emp))
        self.tableau_fixe.setColumnCount(nb_cols_fixes)
        self.tableau_fixe.setHorizontalHeaderLabels(["Employé", "Cycle détecté"])
        self.tableau_fixe.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.ResizeToContents
        )
        self.tableau_fixe.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )

        # ── Tableau planning (colonnes semaines/jours) ──
        self.tableau.clear()
        self.tableau.setRowCount(len(cles_emp))
        self.tableau.setColumnCount(len(self._colonnes))
        self.tableau.setHorizontalHeaderLabels(self._colonnes)
        for c in range(len(self._colonnes)):
            self.tableau.horizontalHeader().setSectionResizeMode(
                c, QHeaderView.ResizeMode.Fixed
            )
            self.tableau.setColumnWidth(c, 52)

        # Pré-initialiser toutes les cellules avec un item vide
        # Evite l'icône Qt native (cercle barré) sur les cellules non remplies
        for r in range(len(cles_emp)):
            for c in range(len(self._colonnes)):
                item_v = QTableWidgetItem("")
                item_v.setBackground(QColor(COULEUR_VIDE))
                self.tableau.setItem(r, c, item_v)

        filtre_cycle = self.combo_filtre_cycle.currentText()
        reel_seul    = self.check_reel_seul.isChecked()

        for row, cle_emp in enumerate(cles_emp):
            nom_affiche = cle_emp.split("|")[0]

            # Col 0 — Nom (tableau fixe)
            item_nom = QTableWidgetItem(nom_affiche)
            item_nom.setForeground(QColor(COULEURS["texte"]))
            item_nom.setData(Qt.ItemDataRole.UserRole, cle_emp)
            self.tableau_fixe.setItem(row, 0, item_nom)

            # Col 1 — Cycle détecté (tableau fixe) — affiche cycle_type ou libelle definitions
            cycle_det = self._get_cycle_detecte(cle_emp)
            # Libelle : cycle_type direct, ou chercher dans definitions via cycle_depart
            data_cycle = self._donnees_cycles.get(cle_emp, {})
            cycle_depart_raw = data_cycle.get("cycle_depart", "").strip()
            if cycle_det and cycle_det == cycle_depart_raw:
                # Pas de cycle_type : chercher un cycle_definition dont la sequence contient ce poste
                libelle_det = next(
                    (k for k, v in self._donnees_definitions.items()
                     if k != "COMMENTAIRE" and cycle_depart_raw in v.get("sequence", [])
                     and v.get("rotation") == "fixe"),
                    cycle_det
                ) if cycle_det else ""
            else:
                libelle_det = cycle_det
            item_det = QTableWidgetItem(libelle_det or "—")
            item_det.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if cycle_depart_raw:
                bg = COULEURS_CYCLE.get(cycle_depart_raw, COULEURS["bg_carte"])
                tc = COULEUR_TEXTE_CYCLE.get(cycle_depart_raw, "#FFFFFF")
                item_det.setBackground(QColor(bg))
                item_det.setForeground(QColor(tc))
            else:
                item_det.setForeground(QColor(COULEURS["texte_secondaire"]))
            item_det.setToolTip(f"Poste de départ : {cycle_depart_raw or '—'}")
            self.tableau_fixe.setItem(row, 1, item_det)

            # Colonnes planning (tableau scrollable)
            for col_idx, cle_col in enumerate(self._colonnes):
                cycle, hypothetique = self._get_cycle_employe(cle_emp, cle_col)

                # Filtre hypothétique
                if reel_seul and hypothetique:
                    item = QTableWidgetItem("")
                    item.setBackground(QColor(COULEUR_VIDE))
                    self.tableau.setItem(row, col_idx, item)
                    continue

                # Absence planifiée — fond brun + cycle visible barré
                if self._est_absent(cle_emp, cle_col):
                    cycle_abs, hyp_abs = self._get_cycle_employe(cle_emp, cle_col)
                    texte_abs = cycle_abs if cycle_abs else ""
                    item = QTableWidgetItem(texte_abs)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    item.setBackground(QColor(COULEUR_ABSENCE))
                    item.setForeground(QColor("#CCCCCC"))
                    f_abs = QFont()
                    f_abs.setStrikeOut(True)
                    if hyp_abs:
                        f_abs.setItalic(True)
                    item.setFont(f_abs)
                    item.setToolTip(f"Absence planifiée{' — cycle estimé : ' + cycle_abs if cycle_abs else ''}")
                    self.tableau.setItem(row, col_idx, item)
                    continue

                if cycle is None and hypothetique is None:
                    item = QTableWidgetItem("")
                    item.setBackground(QColor(COULEUR_VIDE))
                    self.tableau.setItem(row, col_idx, item)
                    continue

                if cycle is None:
                    item = QTableWidgetItem("\u2205")  # ∅
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    item.setBackground(QColor(COULEUR_NULL))
                    item.setForeground(QColor(COULEURS["texte_secondaire"]))
                    self.tableau.setItem(row, col_idx, item)
                    continue

                # Filtre cycle
                if filtre_cycle != "Tous" and cycle != filtre_cycle:
                    item = QTableWidgetItem("")
                    item.setBackground(QColor(COULEUR_VIDE))
                    self.tableau.setItem(row, col_idx, item)
                    continue

                # Cellule normale
                item = QTableWidgetItem(cycle)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                if hypothetique:
                    bg = COULEURS_CYCLE_HYPO.get(cycle, "#3A3A50")
                    item.setBackground(QColor(bg))
                    item.setForeground(QColor("#CCCCCC"))
                    f = QFont()
                    f.setItalic(True)
                    item.setFont(f)
                    item.setData(DelegateHachureHypothetique._HYPO_ROLE, True)
                    item.setToolTip(f"Hypothétique — cycle estimé : {cycle}")
                else:
                    bg = COULEURS_CYCLE.get(cycle, COULEURS["bg_carte"])
                    tc = COULEUR_TEXTE_CYCLE.get(cycle, "#FFFFFF")
                    item.setBackground(QColor(bg))
                    item.setForeground(QColor(tc))

                    # Cycle mixte (synthèse semaine) — fond coupé diagonal
                    if self._mode == "semaine":
                        _, cycle_sec, _ = self._synthese_semaine(cle_emp, cle_col)
                        if cycle_sec:
                            couleur_sec = COULEURS_CYCLE.get(cycle_sec, COULEURS["bg_carte"])
                            item.setData(DelegateHachureHypothetique._MIXTE_ROLE, couleur_sec)
                            item.setToolTip(
                                f"Dominant : {cycle} — Secondaire : {cycle_sec}"
                            )

                self.tableau.setItem(row, col_idx, item)

        nb_emp = len(cles_emp)
        nb_col_plan = len(self._colonnes)
        self.lbl_info.setText(
            f"{nb_emp} employé(s)  ×  {nb_col_plan} {'semaine(s)' if self._mode == 'semaine' else 'jour(s)'}  "
            f"| Double-clic sur 'Cycle détecté' pour corriger"
        )

        # Ajuster largeur tableau fixe
        self.tableau_fixe.resizeColumnsToContents()
        w = self.tableau_fixe.verticalHeader().width()
        for c in range(self.tableau_fixe.columnCount()):
            w += self.tableau_fixe.columnWidth(c)
        w += 4  # marge bordure
        self.tableau_fixe.setFixedWidth(w)

        # Appliquer filtre nom si déjà renseigné
        self._filtrer_lignes(self.champ_filtre_nom.text())

    # --------------------------------------------------
    # Filtrage dynamique nom
    # --------------------------------------------------
    def _filtrer_lignes(self, texte: str):
        texte = texte.strip().upper()
        for row in range(self.tableau_fixe.rowCount()):
            item = self.tableau_fixe.item(row, 0)
            if item is None:
                self.tableau_fixe.setRowHidden(row, False)
                self.tableau.setRowHidden(row, False)
                continue
            nom = item.text().upper()
            masquer = bool(texte and texte not in nom)
            self.tableau_fixe.setRowHidden(row, masquer)
            self.tableau.setRowHidden(row, masquer)

    # --------------------------------------------------
    # Double-clic cellule — correction cycle détecté
    # --------------------------------------------------
    def _proposer_cycle_depuis_donnees(self, cle_emp: str) -> tuple:
        """
        Analyse les données réelles de l'employé et retourne
        (motif_propose: list, nom_propose: str) pour pré-remplir
        DialogueCycleCustomVisu. Utilise le meilleur motif même si score < 60%.
        """
        try:
            import sys as _sys
            import os as _os
            # Chercher detecter_cycles dans le path
            from detecter_cycles import (
                _fusionner_semaines, _sequence_reelle, _sequence_avec_we,
                _tester_periode, PERIODES_TESTEES, _cle_vers_date,
            )
            data_emp = self._donnees_planning.get(cle_emp, {})
            semaines_brutes = data_emp.get('semaines', {})
            jours           = data_emp.get('jours', {})
            semaines        = _fusionner_semaines(semaines_brutes, jours)

            # Essayer toutes les périodes, garder le meilleur motif
            seq = _sequence_reelle(semaines)
            if not seq:
                return (["M"], "Frag_M")
            valeurs = [v.get('cycle') for k, v in seq]

            meilleur_score = 0.0
            meilleur_motif = valeurs[:1]
            for periode in PERIODES_TESTEES:
                if periode == 1:
                    continue
                from detecter_cycles import _tester_periode
                motif, score = _tester_periode(valeurs, periode)
                if score > meilleur_score and motif:
                    meilleur_score = score
                    meilleur_motif = motif

            # Si aucune rotation trouvée, utiliser les cycles les plus fréquents
            if meilleur_score == 0.0:
                from collections import Counter
                cycles_freq = [c for c, _ in Counter(valeurs).most_common()]
                meilleur_motif = cycles_freq[:3] if len(cycles_freq) >= 3 else cycles_freq

            nom_propose = "Frag_" + "".join(meilleur_motif)
            return (meilleur_motif, nom_propose)
        except Exception:
            return (["M"], "Frag_M")

    def _on_double_clic_cellule(self, row: int, col: int):
        """
        Double-clic sur 'Cycle détecté' (col 1) :
        - Si l'employé a déjà un cycle → DialogueCorrectionCycle (correction simple)
        - Si l'employé n'a pas de cycle → DialogueCycleCustomVisu (création cycle custom)
        """
        if col != 1:
            return

        item_nom = self.tableau_fixe.item(row, 0)
        if item_nom is None:
            return
        cle_emp = item_nom.data(Qt.ItemDataRole.UserRole)
        if not cle_emp:
            return

        cycle_actuel = self._get_cycle_detecte(cle_emp)
        data_cycle   = self._donnees_cycles.get(cle_emp, {})
        a_un_cycle   = bool(data_cycle.get('cycle_depart', '').strip())

        if a_un_cycle:
            # Cycle existant → correction simple (comportement original)
            dlg = DialogueCorrectionCycle(cle_emp, cycle_actuel, parent=self)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                return
            nouveau_cycle = dlg.get_cycle()
            self._appliquer_correction_cycle(cle_emp, nouveau_cycle, row)
        else:
            # Pas de cycle → création cycle custom avec proposition du script
            motif_propose, nom_propose = self._proposer_cycle_depuis_donnees(cle_emp)
            dlg = DialogueCycleCustomVisu(
                cle_emp, motif_propose, nom_propose, parent=self
            )
            if dlg.exec() != QDialog.DialogCode.Accepted:
                return
            self._appliquer_cycle_custom_visu(cle_emp, dlg.get_nom(), dlg.get_motif(), row)

    def _on_double_clic_planning(self, row: int, col: int):
        """Double-clic sur la grille planning (Hyp-E2).
        Cellule hyp -> DialogueOverrideCellule options A/B/C.
        Cellule réelle ou vide -> rien.
        """
        if col < 0 or col >= len(self._colonnes):
            return

        item_nom = self.tableau_fixe.item(row, 0)
        if item_nom is None:
            return
        cle_emp = item_nom.data(Qt.ItemDataRole.UserRole)
        if not cle_emp:
            return

        cle_col = self._colonnes[col]
        cycle, hypothetique = self._get_cycle_employe(cle_emp, cle_col)

        if not hypothetique or cycle is None:
            return

        motif_actuel = self._donnees_cycles.get(cle_emp, {}).get("motif", [])

        dlg = DialogueOverrideCellule(cle_emp, cle_col, cycle, motif_actuel, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        mode = dlg.get_mode()
        if mode == DialogueOverrideCellule.MODE_A:
            self._override_ponctuel(cle_emp, cle_col, dlg.get_cycle_a(), row, col)
        elif mode == DialogueOverrideCellule.MODE_B:
            self._override_recalibrage(cle_emp, cle_col, dlg.get_cycle_b(), motif_actuel, row, col)
        elif mode == DialogueOverrideCellule.MODE_C:
            self._override_nouveau_cycle(cle_emp, cle_col, dlg.get_motif_c(), row, col)

    def _maj_cellule_hyp(self, row: int, col: int, cycle: str):
        """Met à jour visuellement une cellule — reste hypothétique (hachuré)."""
        item = self.tableau.item(row, col)
        if not item:
            return
        item.setText(cycle)
        bg = COULEURS_CYCLE_HYPO.get(cycle, "#3A3A50")
        item.setBackground(QColor(bg))
        item.setForeground(QColor("#CCCCCC"))
        f = QFont()
        f.setItalic(True)
        item.setFont(f)
        item.setData(DelegateHachureHypothetique._HYPO_ROLE, True)
        item.setToolTip(f"Hypothétique — cycle estimé : {cycle}")

    def _override_ponctuel(
        self, cle_emp: str, cle_col: str, nouveau_cycle: str, row: int, col: int
    ):
        """Option A — Correction ponctuelle. Reste hypothetique:true."""
        chemin = PLANNING_HISTORIQUE_JSON()
        faire_backup(chemin)
        planning = charger_json(chemin)
        planning.setdefault(cle_emp, {"semaines": {}, "jours": {}})

        if self._mode == "semaine":
            planning[cle_emp].setdefault("semaines", {})[cle_col] = {
                "cycle": nouveau_cycle, "hypothetique": True, "source": "hypothetique",
            }
        else:
            planning[cle_emp].setdefault("jours", {})[cle_col] = {
                "cycle": nouveau_cycle, "hypothetique": True, "source": "hypothetique",
            }

        sauvegarder_json(chemin, planning)
        self._donnees_planning = planning
        self._maj_cellule_hyp(row, col, nouveau_cycle)

        nom = cle_emp.split("|")[0]
        QMessageBox.information(
            self, "Correction ponctuelle",
            f"\u2705  {nom} — {cle_col} : {nouveau_cycle}\n"
            "Cellule modifi\u00e9e (reste hypoth\u00e9tique).\n"
            "\n\U0001f4be  Backup cr\u00e9\u00e9 automatiquement."
        )

    def _override_recalibrage(
        self, cle_emp: str, cle_col: str, nouveau_poste: str,
        motif_actuel: list, row: int, col: int
    ):
        """Option B — Recalibrage à partir d'ici. Recale la phase du motif existant."""
        if not motif_actuel or nouveau_poste not in motif_actuel:
            QMessageBox.warning(
                self, "Recalibrage impossible",
                f"Le poste '{nouveau_poste}' n'est pas dans le motif actuel {motif_actuel}.\n"
                "Utilisez l'option C pour définir un nouveau motif."
            )
            return

        from datetime import date as _date
        from generer_hypothetiques import generer_hypothetiques_employe, _lundi_semaine

        # Calculer la date de début (lundi de cle_col)
        if self._mode == "semaine":
            parts = cle_col.split("_")
            lundi_debut = _date.fromisocalendar(int(parts[1]), int(parts[0][1:]), 1)
        else:
            lundi_debut = _date.fromisoformat(cle_col)

        # Recaler le motif : nouveau_poste devient position 0
        idx = motif_actuel.index(nouveau_poste)
        nouveau_motif = motif_actuel[idx:] + motif_actuel[:idx]

        # Backup + mise à jour cycles_employes
        chemin_c = CYCLES_EMPLOYES_JSON()
        faire_backup(chemin_c)
        cycles = charger_json(chemin_c)
        cycles.setdefault(cle_emp, {})
        cycles[cle_emp]["cycle_depart"] = nouveau_poste
        cycles[cle_emp]["motif"] = nouveau_motif
        cycles[cle_emp]["date_depart"] = lundi_debut.strftime("%d-%m-%Y")
        sauvegarder_json(chemin_c, cycles)
        self._donnees_cycles = cycles

        # Regénérer les hypothétiques à partir de cette date
        chemin_p = PLANNING_HISTORIQUE_JSON()
        faire_backup(chemin_p)
        planning = charger_json(chemin_p)
        from datetime import date as _date2
        nb_sem, nb_j = generer_hypothetiques_employe(
            planning, cle_emp, nouveau_motif,
            lundi_debut, _date2(2030, 12, 31)
        )
        sauvegarder_json(chemin_p, planning)
        self._donnees_planning = planning

        self.charger_donnees()
        QMessageBox.information(
            self, "Recalibrage effectué",
            f"\u2705  Cycle recalibré pour {cle_emp.split('|')[0]}\n"
            f"   Motif : {' \u2192 '.join(nouveau_motif)}\n"
            f"   {nb_sem} semaines + {nb_j} jours regénérés.\n"
            "\n\U0001f4be  Backup créé automatiquement."
        )

    def _override_nouveau_cycle(
        self, cle_emp: str, cle_col: str, nouveau_motif: list, row: int, col: int
    ):
        """Option C — Nouveau cycle à partir d'ici."""
        if not nouveau_motif:
            return

        from datetime import date as _date
        from generer_hypothetiques import generer_hypothetiques_employe, _lundi_semaine

        if self._mode == "semaine":
            parts = cle_col.split("_")
            lundi_debut = _date.fromisocalendar(int(parts[1]), int(parts[0][1:]), 1)
        else:
            lundi_debut = _date.fromisoformat(cle_col)

        # Backup + mise à jour cycles_employes
        chemin_c = CYCLES_EMPLOYES_JSON()
        faire_backup(chemin_c)
        cycles = charger_json(chemin_c)
        cycles.setdefault(cle_emp, {})
        cycles[cle_emp]["cycle_type"]   = "CUSTOM"
        cycles[cle_emp]["cycle_depart"] = nouveau_motif[0]
        cycles[cle_emp]["motif"]        = nouveau_motif
        cycles[cle_emp]["date_depart"]  = lundi_debut.strftime("%d-%m-%Y")
        sauvegarder_json(chemin_c, cycles)
        self._donnees_cycles = cycles

        # Regénérer les hypothétiques à partir de cette date
        chemin_p = PLANNING_HISTORIQUE_JSON()
        faire_backup(chemin_p)
        planning = charger_json(chemin_p)
        from datetime import date as _date2
        nb_sem, nb_j = generer_hypothetiques_employe(
            planning, cle_emp, nouveau_motif,
            lundi_debut, _date2(2030, 12, 31)
        )
        sauvegarder_json(chemin_p, planning)
        self._donnees_planning = planning

        self.charger_donnees()
        QMessageBox.information(
            self, "Nouveau cycle appliqué",
            f"\u2705  Nouveau cycle pour {cle_emp.split('|')[0]}\n"
            f"   Motif : {' \u2192 '.join(nouveau_motif)}\n"
            f"   {nb_sem} semaines + {nb_j} jours regénérés.\n"
            "\n\U0001f4be  Backup créé automatiquement."
        )

    def _appliquer_cycle_custom_visu(
        self, cle_emp: str, nom_cycle: str, motif: list, row: int
    ):
        """
        Écrit le cycle custom dans cycles_employes.json.
        Appelé depuis DialogueCycleCustomVisu (création depuis Visu).
        """
        if not motif:
            return

        chemin = CYCLES_EMPLOYES_JSON()
        faire_backup(chemin)
        cycles = charger_json(chemin)

        cycles[cle_emp] = {
            "cycle_depart": motif[0],
            "cycle_type":   nom_cycle,
            "date_depart":  "",
            "motif":        motif,
        }
        # Conserver cycle_legacy si présent
        if cle_emp in self._donnees_cycles:
            legacy = self._donnees_cycles[cle_emp].get("cycle_legacy", "")
            if legacy:
                cycles[cle_emp]["cycle_legacy"] = legacy

        sauvegarder_json(chemin, cycles)
        self._donnees_cycles = cycles

        # Mettre à jour visuellement la cellule Cycle détecté
        item_det = self.tableau_fixe.item(row, 1)
        if item_det:
            item_det.setText(nom_cycle)
            bg = COULEURS_CYCLE.get(motif[0], COULEURS["bg_carte"])
            tc = COULEUR_TEXTE_CYCLE.get(motif[0], "#FFFFFF")
            item_det.setBackground(QColor(bg))
            item_det.setForeground(QColor(tc))
            item_det.setToolTip(f"Motif : {' → '.join(motif)}")

        # Rafraîchir OngletCyclesEmployes
        try:
            fenetre = self.window()
            if hasattr(fenetre, 'onglet_cycles_emp'):
                fenetre.onglet_cycles_emp.charger_donnees()
        except Exception:
            pass

        nom_emp = cle_emp.split("|")[0]
        QMessageBox.information(
            self, "Cycle créé",
            f"✅  Cycle '{nom_cycle}' créé pour {nom_emp}\n"
            f"   Motif : {' → '.join(motif)}\n\n"
            "💾  Backup créé automatiquement."
        )
        try:
            self.window().status.showMessage(
                f"✅  Cycle '{nom_cycle}' créé pour {nom_emp}", 4000
            )
        except Exception:
            pass

    def _appliquer_correction_cycle(self, cle_emp: str, nouveau_cycle: str, row: int):

        """Écrit la correction dans cycles_employes.json et met à jour la cellule."""
        chemin = CYCLES_EMPLOYES_JSON()
        faire_backup(chemin)

        cycles = charger_json(chemin)

        if cle_emp not in cycles:
            cycles[cle_emp] = {}

        cycles[cle_emp]["cycle_depart"] = nouveau_cycle
        # Si le champ "cycle" existe aussi (détecté automatiquement), on le met à jour
        if "cycle" in cycles[cle_emp]:
            cycles[cle_emp]["cycle"] = nouveau_cycle

        sauvegarder_json(chemin, cycles)
        self._donnees_cycles = cycles  # rafraîchir en mémoire

        # Mettre à jour la cellule visuellement
        item_det = self.tableau_fixe.item(row, 1)
        if item_det:
            if nouveau_cycle:
                item_det.setText(nouveau_cycle)
                bg = COULEURS_CYCLE.get(nouveau_cycle, COULEURS["bg_carte"])
                tc = COULEUR_TEXTE_CYCLE.get(nouveau_cycle, "#FFFFFF")
                item_det.setBackground(QColor(bg))
                item_det.setForeground(QColor(tc))
            else:
                item_det.setText("\u2014")
                item_det.setBackground(QColor(COULEURS["bg_carte"]))
                item_det.setForeground(QColor(COULEURS["texte_secondaire"]))

        # Rafraîchir l'onglet Cycles Employés si disponible
        try:
            fenetre = self.window()
            if hasattr(fenetre, 'onglet_cycles_emp'):
                fenetre.onglet_cycles_emp.charger_donnees()
        except Exception:
            pass

        QMessageBox.information(
            self,
            "Correction enregistrée",
            f"\u2705  Cycle de {cle_emp.split('|')[0]} mis à jour : "
            f"{'—' if not nouveau_cycle else nouveau_cycle}\n\nBackup créé automatiquement."
        )
        try:
            self.window().status.showMessage(
                f"🔄  Cycle de {cle_emp.split('|')[0]} corrigé : {nouveau_cycle or '—'}", 4000
            )
        except Exception:
            pass



# =====================================================
# FENÊTRE PRINCIPALE
# =====================================================