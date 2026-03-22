"""
generer_hypothetiques.py
========================
Génère les entrées hypothétiques dans planning_historique.json
pour chaque employé ayant un cycle détecté dans cycles_employes.json.

Principes :
  - Ancre = première donnée réelle importée (date_depart + cycle_depart)
  - Le motif rotatif est calé sur le lundi de la semaine ISO de date_depart
  - Granularité : jours ET semaines générés simultanément
  - Jamais écraser hypothetique: false (données réelles)
  - Jamais générer après date_sortie
  - Sam/dim = 'R' (repos) pour tous les cycles
  - Plage configurable (date_debut_gen, date_fin_gen)

Peut être appelé :
  - Depuis l'UI (OngletVisualisationPlanning — Hyp-B)
  - En ligne de commande : python generer_hypothetiques.py
"""

import json
import os
from datetime import date, timedelta

# ═══════════════════════════════════════════════════════════════
# CONSTANTES
# ═══════════════════════════════════════════════════════════════

# Motifs de rotation par type de cycle (semaines)
MOTIFS = {
    "3x8":         ["M", "AM", "N"],
    "2x8":         ["M", "AM"],
    "WE_FIXE":     ["WE"],
    "M_FIXE":      ["M"],
    "AM_FIXE":     ["AM"],
    "N_FIXE":      ["N"],
    "J":           ["J"],
    "J_FIXE":      ["J"],
}

# Codes courts pour les cycles fixes (cycle_depart seul suffit)
CYCLES_FIXES = {"J", "WE", "M", "AM", "N"}

# Jours de la semaine travaillés (0=lundi … 4=vendredi)
JOURS_OUVRES = {0, 1, 2, 3, 4}


# ═══════════════════════════════════════════════════════════════
# UTILITAIRES
# ═══════════════════════════════════════════════════════════════

def _parse_date_depart(s: str):
    """
    Convertit 'JJ-MM-AAAA' en date Python.
    Retourne None si invalide.
    """
    if not s:
        return None
    try:
        j, m, a = s.split("-")
        return date(int(a), int(m), int(j))
    except Exception:
        return None


def _parse_date_contrat(s: str):
    """
    Convertit 'JJ-MM-AAAA' en date Python (même format que employes_contrats).
    """
    return _parse_date_depart(s)


def _lundi_semaine(d: date) -> date:
    """Retourne le lundi de la semaine ISO contenant d."""
    return d - timedelta(days=d.weekday())



def _date_from_iso(s: str):
    """Convertit 'AAAA-MM-JJ' en date Python. Retourne None si invalide."""
    try:
        return date.fromisoformat(s)
    except Exception:
        return None


def _cle_semaine(d: date) -> str:
    """date → 'S03_2024' (semaine ISO)."""
    iso = d.isocalendar()
    return f"S{iso[1]:02d}_{iso[0]}"


def _motif_pour_cycle(cycle_type: str, cycle_depart: str) -> list:
    """
    Retourne le motif (liste de postes) pour un cycle donné.
    Pour les cycles rotatifs, on cale le motif sur cycle_depart.
    Pour les cycles fixes, on retourne [cycle_depart].
    """
    # Chercher dans les motifs connus
    if cycle_type in MOTIFS:
        motif = MOTIFS[cycle_type]
        # Caler la phase sur cycle_depart si rotatif
        if len(motif) > 1 and cycle_depart in motif:
            idx = motif.index(cycle_depart)
            return motif[idx:] + motif[:idx]
        return motif

    # Cycle ROTATION_N générique
    if cycle_type and cycle_type.startswith("ROTATION_") and cycle_depart:
        return [cycle_depart]

    # Fallback : cycle_depart seul
    if cycle_depart:
        return [cycle_depart]

    return []


def _poste_semaine(lundi_ancre: date, motif: list, lundi_cible: date) -> str:
    """
    Calcule le poste d'une semaine donnée par rapport à l'ancre.
    lundi_ancre : lundi de la semaine de date_depart (position 0 du motif)
    motif       : liste de postes ex ['M','AM','N']
    lundi_cible : lundi de la semaine à calculer
    """
    if not motif:
        return ""
    delta_semaines = (lundi_cible - lundi_ancre).days // 7
    idx = delta_semaines % len(motif)
    return motif[idx]


def _poste_jour(poste_semaine: str, jour: date) -> str:
    """
    Poste d'un jour donné selon le poste de sa semaine.
    Sam/dim → 'R', WE tous les jours → 'WE'.
    """
    if poste_semaine == "WE":
        return "WE"
    if jour.weekday() >= 5:  # sam=5, dim=6
        return "R"
    return poste_semaine



# ═══════════════════════════════════════════════════════════════
# FONCTION UTILITAIRE — REGÉNÉRATION D'UN SEUL EMPLOYÉ
# ═══════════════════════════════════════════════════════════════

def generer_hypothetiques_employe(
    planning: dict,
    cle_emp: str,
    motif: list,
    date_debut_gen: date,
    date_fin_gen: date,
    date_sortie=None,
) -> tuple:
    """
    Regénère les hypothétiques pour UN SEUL employé à partir d'une date donnée.
    Utilisé par Hyp-E2 options B et C (recalibrage / nouveau cycle).

    - Ne jamais écraser hypothetique:false
    - Tout ce qui est généré reste hypothetique:true
    - Ancre = date_debut_gen (premier lundi >= date_debut_gen), motif[0] = premier poste

    Retourne (nb_semaines, nb_jours) générés.
    Modifie planning en place (ne sauvegarde pas).
    """
    if not motif:
        return 0, 0

    lundi_ancre = _lundi_semaine(date_debut_gen)
    fin_emp = date_fin_gen
    if date_sortie and date_sortie < fin_emp:
        fin_emp = date_sortie

    if cle_emp not in planning:
        planning[cle_emp] = {"semaines": {}, "jours": {}}
    planning[cle_emp].setdefault("semaines", {})
    planning[cle_emp].setdefault("jours", {})

    emp_semaines = planning[cle_emp]["semaines"]
    emp_jours    = planning[cle_emp]["jours"]
    nb_sem = 0
    nb_j   = 0

    # Semaines
    d = lundi_ancre
    fin_lundi = _lundi_semaine(fin_emp)
    while d <= fin_lundi:
        cle_sem = _cle_semaine(d)
        poste = _poste_semaine(lundi_ancre, motif, d)
        if poste:
            existant = emp_semaines.get(cle_sem)
            if not (existant and not existant.get("hypothetique", True)):
                emp_semaines[cle_sem] = {
                    "cycle": poste, "hypothetique": True, "source": "hypothetique"
                }
                nb_sem += 1
        d += timedelta(weeks=1)

    # Jours
    d = date_debut_gen
    while d <= fin_emp:
        cle_jour = d.strftime("%Y-%m-%d")
        lundi_jour = _lundi_semaine(d)
        poste_sem = _poste_semaine(lundi_ancre, motif, lundi_jour)
        poste_j = _poste_jour(poste_sem, d)
        if poste_j:
            existant = emp_jours.get(cle_jour)
            if not (existant and not existant.get("hypothetique", True)):
                emp_jours[cle_jour] = {
                    "cycle": poste_j, "hypothetique": True, "source": "hypothetique"
                }
                nb_j += 1
        d += timedelta(days=1)

    return nb_sem, nb_j

# ═══════════════════════════════════════════════════════════════
# FONCTION PRINCIPALE
# ═══════════════════════════════════════════════════════════════

def generer_hypothetiques(
    chemin_planning: str,
    chemin_cycles: str,
    chemin_employes: str,
    date_debut_gen: date,
    date_fin_gen: date,
    callback_log=None,
) -> dict:
    """
    Génère les entrées hypothétiques dans planning_historique.json.

    Paramètres :
      chemin_planning  : chemin vers planning_historique.json
      chemin_cycles    : chemin vers cycles_employes.json
      chemin_employes  : chemin vers employes_contrats.json
      date_debut_gen   : date de début de la plage de génération
      date_fin_gen     : date de fin de la plage de génération
      callback_log     : fonction(str) pour les logs (optionnel)

    Retourne un dict de stats :
      { 'nb_employes': int, 'nb_semaines': int, 'nb_jours': int,
        'nb_ignores': int, 'erreurs': list }
    """

    def log(msg):
        if callback_log:
            callback_log(msg)

    stats = {"nb_employes": 0, "nb_semaines": 0, "nb_jours": 0,
             "nb_ignores": 0, "erreurs": []}

    # ── Chargement des fichiers ──────────────────────────────────
    try:
        with open(chemin_planning, encoding="utf-8") as f:
            planning = json.load(f)
    except Exception as e:
        msg = f"Erreur lecture planning_historique.json : {e}"
        log(f"❌  {msg}")
        stats["erreurs"].append(msg)
        return stats

    try:
        with open(chemin_cycles, encoding="utf-8") as f:
            cycles = json.load(f)
    except Exception as e:
        msg = f"Erreur lecture cycles_employes.json : {e}"
        log(f"❌  {msg}")
        stats["erreurs"].append(msg)
        return stats

    try:
        with open(chemin_employes, encoding="utf-8") as f:
            employes = json.load(f)
    except Exception as e:
        msg = f"Erreur lecture employes_contrats.json : {e}"
        log(f"❌  {msg}")
        stats["erreurs"].append(msg)
        return stats

    log(f"📂  Fichiers chargés — {len(cycles)} entrées cycles")
    log(f"📅  Plage de génération : {date_debut_gen.strftime('%d/%m/%Y')} → {date_fin_gen.strftime('%d/%m/%Y')}")

    # ── Calcul des semaines et jours de la plage ─────────────────
    # Semaines ISO de la plage
    semaines_plage = []
    d = _lundi_semaine(date_debut_gen)
    fin_lundi = _lundi_semaine(date_fin_gen)
    while d <= fin_lundi:
        semaines_plage.append(d)
        d += timedelta(weeks=1)

    # Jours de la plage
    jours_plage = []
    d = date_debut_gen
    while d <= date_fin_gen:
        jours_plage.append(d)
        d += timedelta(days=1)

    # ── Traitement par employé ────────────────────────────────────
    for cle_emp, data_cycle in cycles.items():
        if cle_emp == "COMMENTAIRE" or "|" not in cle_emp:
            continue

        cycle_depart = data_cycle.get("cycle_depart", "")
        cycle_type   = data_cycle.get("cycle_type", "")
        date_depart_s = data_cycle.get("date_depart", "")

        # Ignorer les employés sans cycle défini
        if not cycle_depart:
            stats["nb_ignores"] += 1
            continue

        # Motif : priorité au motif stocké (détecté depuis les vraies données)
        # Fallback : recalcul depuis MOTIFS[cycle_type] + cycle_depart
        motif_stocke = data_cycle.get("motif", [])
        if motif_stocke and isinstance(motif_stocke, list) and len(motif_stocke) > 0:
            motif = motif_stocke
        else:
            motif = _motif_pour_cycle(cycle_type, cycle_depart)
        if not motif:
            log(f"   ⚠️  {cle_emp} — motif introuvable (cycle_type={cycle_type!r}, depart={cycle_depart!r})")
            stats["nb_ignores"] += 1
            continue

        # Ancre : derniere semaine reelle dont le cycle est dans le motif
        # -> garantit la continuite a la jonction reel/hypothetique
        # Fallback : lundi de date_depart si aucune semaine reelle exploitable
        date_depart = _parse_date_depart(date_depart_s)
        if not date_depart:
            log(f"   ⚠️  {cle_emp} — date_depart invalide ({date_depart_s!r}), ignoré")
            stats["nb_ignores"] += 1
            continue

        emp_data_tmp     = planning.get(cle_emp, {})
        emp_semaines_tmp = emp_data_tmp.get("semaines", {})
        emp_jours_tmp    = emp_data_tmp.get("jours", {})

        reelles_valides = sorted([
            (k, v) for k, v in emp_semaines_tmp.items()
            if v.get("cycle") in motif and not v.get("hypothetique", True)
        ])
        if reelles_valides:
            # Ancre = semaine SUIVANTE apres la derniere reelle
            # Motif cale sur le poste SUIVANT -> continuite parfaite
            cle_ancre, val_ancre = reelles_valides[-1]
            parts = cle_ancre.split("_")          # ["S52", "2021"]
            lundi_derniere = date.fromisocalendar(int(parts[1]), int(parts[0][1:]), 1)
            lundi_ancre = lundi_derniere + timedelta(weeks=1)
            poste_ancre = val_ancre["cycle"]
            idx_suivant = (motif.index(poste_ancre) + 1) % len(motif)
            motif = motif[idx_suivant:] + motif[:idx_suivant]
        else:
            # Fallback : chercher l'ancre dans les jours réels
            # (cas des employés avec uniquement des données jour — Excel ADP hebdo)
            jours_reels_valides = sorted([
                k for k, v in emp_jours_tmp.items()
                if v.get("cycle") in motif and not v.get("hypothetique", True)
                and _date_from_iso(k) is not None
            ])
            if jours_reels_valides:
                dernier_jour = _date_from_iso(jours_reels_valides[-1])
                lundi_derniere = _lundi_semaine(dernier_jour)
                lundi_ancre = lundi_derniere + timedelta(weeks=1)
                # Déterminer le poste de cette dernière semaine depuis les jours
                poste_derniere_sem = None
                for cle_j in reversed(jours_reels_valides):
                    d_j = _date_from_iso(cle_j)
                    if d_j and _lundi_semaine(d_j) == lundi_derniere:
                        cycle_j = emp_jours_tmp[cle_j].get("cycle")
                        if cycle_j in motif:
                            poste_derniere_sem = cycle_j
                            break
                if poste_derniere_sem and poste_derniere_sem in motif:
                    idx_suivant = (motif.index(poste_derniere_sem) + 1) % len(motif)
                    motif = motif[idx_suivant:] + motif[:idx_suivant]
                else:
                    lundi_ancre = _lundi_semaine(date_depart)
            else:
                lundi_ancre = _lundi_semaine(date_depart)

        # Date de sortie (borne max) et date de début (borne min)
        contrat = employes.get(cle_emp, {})
        date_sortie_s = contrat.get("date_sortie", "")
        date_sortie = _parse_date_contrat(date_sortie_s) if date_sortie_s else None

        date_debut_contrat_s = contrat.get("date_debut", "")
        date_debut_contrat = _parse_date_contrat(date_debut_contrat_s) if date_debut_contrat_s else None

        # Borne effective de fin pour cet employé
        fin_emp = date_fin_gen
        if date_sortie and date_sortie < fin_emp:
            fin_emp = date_sortie

        # Borne effective de début : jamais avant la date d'entrée dans le contrat
        debut_emp = date_debut_gen
        if date_debut_contrat and date_debut_contrat > debut_emp:
            debut_emp = date_debut_contrat

        # Initialiser l'entrée planning si absente
        if cle_emp not in planning:
            planning[cle_emp] = {"semaines": {}, "jours": {}}
        if "semaines" not in planning[cle_emp]:
            planning[cle_emp]["semaines"] = {}
        if "jours" not in planning[cle_emp]:
            planning[cle_emp]["jours"] = {}

        emp_semaines = planning[cle_emp]["semaines"]
        emp_jours    = planning[cle_emp]["jours"]

        nb_sem_emp = 0
        nb_jours_emp = 0

        # ── Génération des semaines ──────────────────────────────
        for lundi in semaines_plage:
            # Respecter les bornes employé (fin ET début de contrat)
            if lundi > fin_emp:
                break
            # Ne pas générer avant la date d'entrée dans le contrat
            # (le lundi de la semaine suffit comme comparaison)
            if lundi + timedelta(days=6) < debut_emp:
                continue

            cle_sem = _cle_semaine(lundi)
            poste = _poste_semaine(lundi_ancre, motif, lundi)
            if not poste:
                continue

            # Ne jamais écraser une donnée réelle
            existant = emp_semaines.get(cle_sem)
            if existant and not existant.get("hypothetique", True):
                continue

            emp_semaines[cle_sem] = {
                "cycle":       poste,
                "hypothetique": True,
                "source":      "hypothetique",
            }
            nb_sem_emp += 1

        # ── Génération des jours ─────────────────────────────────
        for jour in jours_plage:
            if jour > fin_emp:
                break
            # Ne pas générer avant la date d'entrée dans le contrat
            if jour < debut_emp:
                continue

            cle_jour = jour.strftime("%Y-%m-%d")

            # Calculer le poste du jour
            lundi_jour = _lundi_semaine(jour)
            poste_sem = _poste_semaine(lundi_ancre, motif, lundi_jour)
            poste_j = _poste_jour(poste_sem, jour)
            if not poste_j:
                continue

            # Ne jamais écraser une donnée réelle
            existant = emp_jours.get(cle_jour)
            if existant and not existant.get("hypothetique", True):
                continue

            emp_jours[cle_jour] = {
                "cycle":        poste_j,
                "hypothetique": True,
                "source":       "hypothetique",
            }
            nb_jours_emp += 1

        if nb_sem_emp > 0 or nb_jours_emp > 0:
            stats["nb_employes"] += 1
            stats["nb_semaines"] += nb_sem_emp
            stats["nb_jours"]    += nb_jours_emp
            nom = cle_emp.split("|")[0]
            log(f"   ✅  {nom} — {nb_sem_emp} semaines + {nb_jours_emp} jours générés")
        else:
            stats["nb_ignores"] += 1

    # ── Sauvegarde ───────────────────────────────────────────────
    if stats["nb_employes"] > 0:
        try:
            with open(chemin_planning, "w", encoding="utf-8") as f:
                json.dump(planning, f, ensure_ascii=False, indent=2)
            log(f"\n💾  planning_historique.json sauvegardé")
        except Exception as e:
            msg = f"Erreur sauvegarde : {e}"
            log(f"❌  {msg}")
            stats["erreurs"].append(msg)
    else:
        log("\nℹ️  Aucune entrée générée (aucun cycle défini ou tous déjà réels).")

    return stats


# ═══════════════════════════════════════════════════════════════
# ENTRÉE CLI
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys
    base = os.path.dirname(os.path.abspath(__file__))

    chemin_planning = os.path.join(base, "planning_historique.json")
    chemin_cycles   = os.path.join(base, "cycles_employes.json")
    chemin_employes = os.path.join(base, "employes_contrats.json")

    debut = date(2021, 1, 1)
    fin   = date.today()

    print(f"Génération hypothétiques du {debut} au {fin}…\n")
    stats = generer_hypothetiques(
        chemin_planning, chemin_cycles, chemin_employes,
        debut, fin,
        callback_log=print,
    )
    print(f"\n{'='*50}")
    print(f"Résultat : {stats['nb_employes']} employés traités")
    print(f"  Semaines générées : {stats['nb_semaines']}")
    print(f"  Jours générés     : {stats['nb_jours']}")
    print(f"  Ignorés           : {stats['nb_ignores']}")
    if stats["erreurs"]:
        print(f"  Erreurs           : {stats['erreurs']}")
    sys.exit(0 if not stats["erreurs"] else 1)