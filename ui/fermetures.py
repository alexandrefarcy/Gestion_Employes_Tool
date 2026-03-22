"""
ui/fermetures.py
================
Calcul des fermetures obligatoires de l'entreprise.

Règles :
  - Été   : première semaine pleine d'août (lundi) + 3 semaines (15 jours ouvrés)
            "Première semaine pleine" = premier lundi où toute la semaine est en août
            (si le 1er août est lundi → ce lundi, sinon lundi suivant)
  - Hiver : 24 décembre → 2 janvier (jours ouvrés uniquement)

Usage :
    from ui.fermetures import jours_fermetures_annee, fusionner_absences_fermetures
"""

from datetime import date, timedelta


def _premier_lundi_semaine_pleine_aout(annee: int) -> date:
    """
    Retourne le premier lundi dont toute la semaine (lun-dim) est en août.
    Si le 1er août est un lundi → ce lundi.
    Sinon → le lundi suivant (premier lundi >= 2 août tel que lundi >= 1 août).
    """
    premier_aout = date(annee, 8, 1)
    # weekday() : 0=lundi … 6=dimanche
    dow = premier_aout.weekday()
    if dow == 0:
        return premier_aout  # 1er août est lundi → semaine pleine dès le 1er
    else:
        # Aller au lundi suivant
        return premier_aout + timedelta(days=(7 - dow))


def jours_fermetures_annee(annee: int) -> set:
    """
    Retourne l'ensemble des jours ouvrés (lun-ven) de fermeture obligatoire
    pour une année donnée.

    Inclut :
      - 3 semaines d'été (15 jours ouvrés à partir du premier lundi semaine pleine d'août)
      - Hiver : 24 décembre → 2 janvier de l'année suivante (jours ouvrés)
    """
    jours = set()

    # ── Été : 3 semaines = 21 jours calendaires depuis le premier lundi ──
    lundi_debut_ete = _premier_lundi_semaine_pleine_aout(annee)
    d = lundi_debut_ete
    for _ in range(21):  # 3 semaines = 21 jours
        if d.weekday() < 5:  # lun-ven uniquement
            jours.add(d)
        d += timedelta(days=1)

    # ── Hiver : 24 décembre → 2 janvier (année suivante) ──
    d = date(annee, 12, 24)
    fin_hiver = date(annee + 1, 1, 2)
    while d <= fin_hiver:
        if d.weekday() < 5:
            jours.add(d)
        d += timedelta(days=1)

    return jours


def jours_fermetures_periode(d_debut: date, d_fin: date) -> set:
    """
    Retourne tous les jours ouvrés de fermeture obligatoire
    entre d_debut et d_fin (inclus).
    """
    jours = set()
    annees = range(d_debut.year, d_fin.year + 1)
    for annee in annees:
        for j in jours_fermetures_annee(annee):
            if d_debut <= j <= d_fin:
                jours.add(j)
    return jours


def fusionner_absences_fermetures(
    absences_emp: set,
    d_debut: date,
    d_fin: date,
) -> set:
    """
    Retourne l'union des absences personnelles et des fermetures obligatoires,
    sans doublon, filtrée sur la période.

    absences_emp : set de dates (jours ouvrés d'absence personnelle)
    """
    fermetures = jours_fermetures_periode(d_debut, d_fin)
    return absences_emp | fermetures


def periodes_fermetures_annee(annee: int) -> list:
    """
    Retourne la liste des périodes de fermeture au format
    [{"debut": "JJ-MM-AAAA", "fin": "JJ-MM-AAAA", "label": str}]
    pour affichage dans l'onglet Absences.
    """
    lundi = _premier_lundi_semaine_pleine_aout(annee)
    fin_ete = lundi + timedelta(days=20)  # 3 semaines - 1 jour

    return [
        {
            "debut": lundi.strftime("%d-%m-%Y"),
            "fin":   fin_ete.strftime("%d-%m-%Y"),
            "label": f"Fermeture été {annee}",
        },
        {
            "debut": f"24-12-{annee}",
            "fin":   f"02-01-{annee + 1}",
            "label": f"Fermeture hiver {annee}/{annee + 1}",
        },
    ]
