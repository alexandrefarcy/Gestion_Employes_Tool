"""
ui/widgets.py
=============
Widgets Qt réutilisables partagés entre tous les onglets.
"""

from PyQt6.QtWidgets import QComboBox, QLineEdit
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QKeyEvent


class ComboSansScroll(QComboBox):
    """QComboBox qui ignore le scroll de la souris — évite les changements accidentels."""
    def wheelEvent(self, event):
        event.ignore()


class ChampNom(QLineEdit):
    """Champ Nom/Prénom : accepte lettres, tirets, espaces et accents. Bloque chiffres et caractères spéciaux."""
    _INTERDITS = set('0123456789,;:/!?@#$%^&*()+=[]{}|<>~`"\'\\')

    def keyPressEvent(self, event):
        touches_ok = {
            Qt.Key.Key_Backspace, Qt.Key.Key_Delete,
            Qt.Key.Key_Left, Qt.Key.Key_Right,
            Qt.Key.Key_Home, Qt.Key.Key_End,
            Qt.Key.Key_Tab, Qt.Key.Key_Backtab,
        }
        if event.key() in touches_ok:
            super().keyPressEvent(event)
            return
        char = event.text()
        if not char or char in self._INTERDITS:
            return
        super().keyPressEvent(event)


class ChampMatricule(QLineEdit):
    """Champ Matricule : accepte uniquement des chiffres (0-9)."""

    def keyPressEvent(self, event):
        touches_ok = {
            Qt.Key.Key_Backspace, Qt.Key.Key_Delete,
            Qt.Key.Key_Left, Qt.Key.Key_Right,
            Qt.Key.Key_Home, Qt.Key.Key_End,
            Qt.Key.Key_Tab, Qt.Key.Key_Backtab,
        }
        if event.key() in touches_ok:
            super().keyPressEvent(event)
            return
        char = event.text()
        if not char or not char.isdigit():
            return
        super().keyPressEvent(event)


class ChampDateMasque(QLineEdit):
    """
    Champ de saisie de date au format JJ-MM-AAAA.
    - N'accepte que des chiffres et touches de navigation.
    - Insère automatiquement les tirets aux positions 2 et 5.
    """

    MASQUE = "00-00-0000"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setPlaceholderText("JJ-MM-AAAA")
        self.setMaxLength(10)
        self.setFixedWidth(140)
        self._en_cours = False

    def keyPressEvent(self, event: QKeyEvent):
        touches_autorisees = {
            Qt.Key.Key_Backspace, Qt.Key.Key_Delete,
            Qt.Key.Key_Left, Qt.Key.Key_Right,
            Qt.Key.Key_Home, Qt.Key.Key_End,
            Qt.Key.Key_Tab, Qt.Key.Key_Backtab,
        }
        if event.key() in touches_autorisees:
            super().keyPressEvent(event)
            return
        char = event.text()
        if not char or not char.isdigit():
            return
        if self._en_cours:
            return
        self._en_cours = True
        try:
            texte  = self.text()
            pos    = self.cursorPosition()
            chiffres_avant    = sum(1 for c in texte[:pos] if c.isdigit())
            chiffres_existants = [c for c in texte if c.isdigit()]
            chiffres_existants.insert(chiffres_avant, char)
            chiffres = "".join(chiffres_existants[:8])
            formate = ""
            for i, c in enumerate(chiffres):
                if i == 2 or i == 4:
                    formate += "-"
                formate += c
            self.setText(formate)
            cible_chiffres = chiffres_avant + 1
            nouvelle_pos = 0
            compte = 0
            for i, c in enumerate(formate):
                if c.isdigit():
                    compte += 1
                    if compte == cible_chiffres:
                        nouvelle_pos = i + 1
                        break
            else:
                nouvelle_pos = len(formate)
            self.setCursorPosition(nouvelle_pos)
        finally:
            self._en_cours = False
