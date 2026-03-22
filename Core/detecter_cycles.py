"""
detecter_cycles.py
==================
Analyse planning_historique.json pour détecter automatiquement
le cycle de travail de chaque employé, puis écrit le résultat
dans cycles_employes.json.

Principes :
  - Ne lit que les données réelles (hypothetique: false)
  - Ignore les WE pour trouver le cycle M/AM/N/J sous-jacent
  - Teste les périodes 1 (fixe), 3 (3x8 hebdo), 6 (2x8 ou 3x8 mensuel)
  - Seuil de cohérence : 70% minimum
  - Si non détecté → laisse vide, ne touche pas à l'entrée existante
  - N'écrase JAMAIS une entrée déjà remplie manuellement

Peut être appelé :
  - Depuis l'UI (OngletPlanning — Phase 4)
  - En ligne de commande : python detecter_cycles.py
"""

import json
import os
import shutil
from datetime import datetime


# ═══════════════════════════════════════════════════════════════
# CONSTANTES
# ═══════════════════════════════════════════════════════════════

CYCLES_ROTATIFS   = {'M', 'AM', 'N'}
CYCLES_FIXES      = {'J', 'WE', 'R'}   # R = repos, ignoré comme WE pour la rotation
SEUIL_COHERENCE   = 0.60   # 60% des semaines doivent correspondre au motif (Hyp-D : abaissé de 70%)
MIN_SEMAINES      = 5      # Minimum de semaines réelles pour tenter la détection
PERIODES_TESTEES  = [1, 2, 3, 4, 5, 6]  # Hyp-D : elargi de [1,3,6] à toutes périodes 1-6


# ═══════════════════════════════════════════════════════════════
# UTILITAIRES
# ═══════════════════════════════════════════════════════════════

def _parse_sem(cle):
    """'S03_2023' → (2023, 3) pour tri chronologique correct."""
    try:
        parts = cle.split('_')
        return (int(parts[1]), int(parts[0][1:]))
    except Exception:
        return (0, 0)


def _sequence_reelle(semaines_dict):
    """
    Retourne la séquence chronologique des cycles réels non-null,
    en excluant les WE (pour trouver le cycle sous-jacent).
    Retourne : liste de (cle_semaine, cycle)
    """
    entries = [
        (k, v)
        for k, v in semaines_dict.items()
        if not v.get('hypothetique', True)          # données réelles uniquement
        and v.get('cycle') is not None              # pas de null
        and v.get('cycle') not in CYCLES_FIXES      # ignorer WE/J pour la rotation
    ]
    return sorted(entries, key=lambda x: _parse_sem(x[0]))


def _sequence_avec_we(semaines_dict):
    """
    Séquence réelle incluant WE — pour détecter les cycles WE fixes.
    """
    entries = [
        (k, v)
        for k, v in semaines_dict.items()
        if not v.get('hypothetique', True)
        and v.get('cycle') is not None
    ]
    return sorted(entries, key=lambda x: _parse_sem(x[0]))


# ═══════════════════════════════════════════════════════════════
# DÉTECTION DU MOTIF
# ═══════════════════════════════════════════════════════════════

def _tester_periode(valeurs, periode):
    """
    Teste si la séquence correspond à une rotation de longueur `periode`.

    Stratégie robuste : essaie toutes les permutations possibles des cycles
    présents sur TOUS les décalages de phase. Retourne le meilleur résultat.

    Retourne (motif, score) où :
      - motif : liste de `periode` cycles (ex: ['M', 'AM', 'N'])
      - score : proportion de semaines qui correspondent (0.0 à 1.0)
    """
    from itertools import permutations as _perms

    if len(valeurs) < periode:
        return None, 0.0

    meilleur_score = 0.0
    meilleur_motif = valeurs[:periode]

    # Générer les motifs candidats selon la période
    if periode == 3 and set(v for v in valeurs if v) <= {"M", "AM", "N"}:
        candidats = [list(p) for p in _perms(["M", "AM", "N"])]
    elif periode == 6 and set(v for v in valeurs if v) <= {"M", "AM", "N"}:
        candidats = []
        for base in [["M", "AM"], ["AM", "M"]]:
            candidats.append(base * 3)
        for p in _perms(["M", "AM", "N"]):
            candidats.append(list(p) * 2)
    else:
        candidats = [valeurs[:periode]]

    for motif_base in candidats:
        for decalage in range(periode):
            motif = motif_base[decalage:] + motif_base[:decalage]
            score = sum(
                1 for i, v in enumerate(valeurs)
                if v == motif[i % periode]
            ) / len(valeurs)
            if score > meilleur_score:
                meilleur_score = score
                meilleur_motif = motif

    return meilleur_motif, meilleur_score




def _jours_vers_semaines(jours_dict):
    """
    Convertit les données jour en semaines virtuelles pour la détection de cycles.
    Pour chaque semaine ISO, prend le cycle majoritaire des jours lun-ven.
    Retourne un dict au même format que semaines_dict :
      { "S03_2025": {"cycle": "AM", "hypothetique": False, "source": "jours_agrege"} }
    Utilisé quand un employé n'a que des données jour (Excel ADP hebdo).
    """
    from datetime import date as _date
    from collections import Counter

    semaines_jours = {}  # { "S03_2025": [cycle, ...] }

    for cle_j, val in jours_dict.items():
        if val.get('hypothetique', True):
            continue
        cycle = val.get('cycle')
        if not cycle:
            continue
        try:
            d = _date.fromisoformat(cle_j)
            # Ignorer samedi (5) et dimanche (6)
            if d.weekday() >= 5:
                continue
            iso = d.isocalendar()
            cle_sem = f"S{iso[1]:02d}_{iso[0]}"
            semaines_jours.setdefault(cle_sem, []).append(cycle)
        except Exception:
            continue

    result = {}
    for cle_sem, cycles in semaines_jours.items():
        if not cycles:
            continue
        # Cycle majoritaire de la semaine
        cycle_maj = Counter(cycles).most_common(1)[0][0]
        result[cle_sem] = {
            "cycle": cycle_maj,
            "hypothetique": False,
            "source": "jours_agrege",
        }
    return result


def _fusionner_semaines(semaines_dict, jours_dict):
    """
    Fusionne semaines réelles + semaines dérivées des jours.
    Les semaines réelles ont priorité sur les semaines dérivées.
    """
    if not jours_dict:
        return semaines_dict

    semaines_from_jours = _jours_vers_semaines(jours_dict)
    if not semaines_from_jours:
        return semaines_dict

    # Partir des semaines dérivées, écraser avec les réelles
    fusion = dict(semaines_from_jours)
    fusion.update(semaines_dict)
    return fusion


def detecter_cycle_employe(semaines_dict):
    """
    Analyse les semaines d'un employé et retourne le cycle détecté.

    Retourne un dict :
    {
        "cycle":        str,   # ex: "3x8", "2x8", "M_FIXE", "WE_FIXE", ""
        "cycle_depart": str,   # premier poste : "M", "AM", "N", "J", "WE"
        "date_depart":  str,   # date JJ-MM-AAAA de la première semaine connue
        "score":        float, # cohérence 0→1 (non écrit dans le JSON final)
        "detecte":      bool,
        "note":         str,   # explication lisible
    }
    """
    # ── 0. Cas J fixe — testé avant _sequence_reelle (qui exclut J) ──────
    seq_tout_j = _sequence_avec_we(semaines_dict)
    if seq_tout_j:
        valeurs_j = [v.get('cycle') for k, v in seq_tout_j]
        nb_j = valeurs_j.count('J')
        if nb_j / len(valeurs_j) >= 0.85:
            premiere_cle_j = seq_tout_j[0][0]
            date_dep_j = _cle_vers_date(premiere_cle_j)
            return {
                "cycle": "J_FIXE", "cycle_depart": "J",
                "date_depart": date_dep_j, "score": nb_j / len(valeurs_j),
                "motif": ["J"],
                "detecte": True, "note": f"Journée fixe ({nb_j/len(valeurs_j)*100:.0f}%)"
            }

    # ── 1. Cas WE fixe ──────────────────────────────────────────
    seq_tout = _sequence_avec_we(semaines_dict)
    if seq_tout:
        valeurs_tout = [v.get('cycle') for k, v in seq_tout]
        nb_we = valeurs_tout.count('WE')
        if nb_we / len(valeurs_tout) >= 0.85:
            premiere_cle = seq_tout[0][0]
            date_dep = _cle_vers_date(premiere_cle)
            return {
                "cycle": "WE_FIXE", "cycle_depart": "WE",
                "date_depart": date_dep, "score": nb_we / len(valeurs_tout),
                "detecte": True, "note": "Cycle WE fixe détecté"
            }

    # ── 2. Séquence sans WE ──────────────────────────────────────
    seq = _sequence_reelle(semaines_dict)

    if len(seq) < MIN_SEMAINES:
        return {
            "cycle": "", "cycle_depart": "", "date_depart": "",
            "score": 0.0, "detecte": False,
            "note": f"Données insuffisantes ({len(seq)} semaines réelles < {MIN_SEMAINES})"
        }

    valeurs   = [v.get('cycle') for k, v in seq]
    premiere_cle = seq[0][0]
    date_dep  = _cle_vers_date(premiere_cle)

    cycles_presents = set(valeurs)

    # ── 3. Cycle fixe J ─────────────────────────────────────────
    if cycles_presents == {'J'}:
        return {
            "cycle": "J", "cycle_depart": "J",
            "date_depart": date_dep, "score": 1.0,
            "detecte": True, "note": "Journée fixe"
        }

    # ── 4. Cycle fixe M ou AM ────────────────────────────────────
    for poste_fixe, nom_cycle in [('M', 'M_FIXE'), ('AM', 'AM_FIXE'), ('N', 'N_FIXE')]:
        nb = valeurs.count(poste_fixe)
        if nb / len(valeurs) >= 0.90:
            return {
                "cycle": nom_cycle, "cycle_depart": poste_fixe,
                "date_depart": date_dep, "score": nb / len(valeurs),
                "detecte": True, "note": f"Poste fixe {poste_fixe} ({nb/len(valeurs)*100:.0f}%)"
            }

    # ── 5. Rotation M/AM/N — tester périodes ─────────────────────
    if not cycles_presents <= CYCLES_ROTATIFS | {'WE', 'R'}:
        # Contient des cycles inattendus
        return {
            "cycle": "", "cycle_depart": "", "date_depart": "",
            "score": 0.0, "detecte": False,
            "note": f"Cycles mixtes non reconnus : {cycles_presents}"
        }

    meilleur = {"score": 0.0, "periode": 0, "motif": []}

    for periode in PERIODES_TESTEES:
        if periode == 1:
            continue  # géré par les cas fixes ci-dessus
        motif, score = _tester_periode(valeurs, periode)
        if score > meilleur["score"]:
            meilleur = {"score": score, "periode": periode, "motif": motif}

    if meilleur["score"] >= SEUIL_COHERENCE:
        periode  = meilleur["periode"]
        motif    = meilleur["motif"]
        score    = meilleur["score"]
        depart   = motif[0]

        # Nommer le cycle
        cycles_uniques = set(motif)
        if cycles_uniques == {'M', 'AM', 'N'}:
            nom = "3x8"
        elif cycles_uniques == {'M', 'AM'}:
            nom = "2x8"
        elif cycles_uniques == {'M', 'N'} or cycles_uniques == {'AM', 'N'}:
            nom = "2x8"
        else:
            nom = f"ROTATION_{periode}"

        return {
            "cycle": nom, "cycle_depart": depart,
            "date_depart": date_dep, "score": score,
            "motif": motif,
            "detecte": True,
            "note": f"{nom} période/{periode}sem, motif: {'>'.join(motif)} ({score*100:.0f}%)"
        }

    # ── 6. Non détecté ───────────────────────────────────────────
    return {
        "cycle": "", "cycle_depart": "", "date_depart": "",
        "score": meilleur["score"], "detecte": False,
        "note": f"Motif irrégulier (meilleur score: {meilleur['score']*100:.0f}%)"
    }


def _cle_vers_date(cle_semaine):
    """
    Convertit 'S03_2023' en date JJ-MM-AAAA du lundi de cette semaine ISO.
    Ex: 'S03_2023' → '16-01-2023'
    """
    try:
        parts = cle_semaine.split('_')
        annee = int(parts[1])
        sem   = int(parts[0][1:])
        # Lundi de la semaine ISO
        d = datetime.strptime(f"{annee}-W{sem:02d}-1", "%G-W%V-%u")
        return d.strftime("%d-%m-%Y")
    except Exception:
        return ""


# ═══════════════════════════════════════════════════════════════
# FONCTION PRINCIPALE
# ═══════════════════════════════════════════════════════════════

def analyser_sans_ecrire(
    chemin_planning,
    chemin_cycles_employes,
    ecraser_manuel=False,
    callback_log=None,
):
    """
    Analyse planning_historique.json et retourne les résultats détectés
    SANS écrire dans cycles_employes.json.

    Retourne une liste de dicts par employé avec cycle détecté, cycle actuel,
    flag conflit (cycle_actuel != cycle détecté si cycle_actuel non vide).
    """
    def log(msg):
        if callback_log:
            callback_log(msg)
        else:
            print(msg)

    if not os.path.exists(chemin_planning):
        log("\u274c planning_historique.json introuvable")
        return []

    with open(chemin_planning, 'r', encoding='utf-8') as f:
        planning = json.load(f)

    cycles_employes = {}
    if os.path.exists(chemin_cycles_employes):
        with open(chemin_cycles_employes, 'r', encoding='utf-8') as f:
            cycles_employes = json.load(f)

    log(f"\n\U0001f50d Analyse de {len(planning)} employ\u00e9(s)...\n")

    resultats = []

    for cle_emp, data in sorted(planning.items()):
        if cle_emp == "COMMENTAIRE":
            continue

        semaines_brutes = data.get('semaines', {})
        jours           = data.get('jours', {})
        semaines        = _fusionner_semaines(semaines_brutes, jours)
        existant = cycles_employes.get(cle_emp, {})
        cycle_actuel = existant.get('cycle_depart', '').strip()
        deja_rempli = cycle_actuel != ''

        if deja_rempli and not ecraser_manuel:
            log(f"   \u23ed\ufe0f  {cle_emp.split('|')[0]} \u2014 conserv\u00e9 (saisi manuellement : {cycle_actuel})")
            continue

        resultat = detecter_cycle_employe(semaines)
        nom_affiche = cle_emp.split('|')[0]

        if not resultat['detecte']:
            log(f"   \u26a0\ufe0f  {nom_affiche} \u2192 {resultat['note']}")
            continue

        cycle_detecte = resultat['cycle_depart']
        conflit = deja_rempli and cycle_actuel != cycle_detecte

        suffixe = f" [\u26a0\ufe0f CONFLIT: actuel={cycle_actuel}]" if conflit else ""
        log(f"   \u2705 {nom_affiche} \u2192 {resultat['note']}{suffixe}")

        resultats.append({
            "cle_emp":      cle_emp,
            "nom":          nom_affiche,
            "cycle_depart": cycle_detecte,
            "cycle_type":   resultat['cycle'],
            "date_depart":  resultat['date_depart'],
            "motif":        resultat.get('motif', []),
            "cycle_actuel": cycle_actuel,
            "conflit":      conflit,
            "score":        resultat['score'],
            "note":         resultat['note'],
        })

    log(f"\n{'='*50}")
    log(f"\U0001f4ca {len(resultats)} cycle(s) \u00e0 valider")

    return resultats


def detecter_tous_cycles(
    chemin_planning,
    chemin_cycles_employes,
    ecraser_manuel=False,
    callback_log=None,
):
    """
    Analyse planning_historique.json et met à jour cycles_employes.json.

    Paramètres :
      chemin_planning        : chemin vers planning_historique.json
      chemin_cycles_employes : chemin vers cycles_employes.json
      ecraser_manuel         : si False (défaut), ne touche pas aux entrées
                               déjà remplies manuellement
      callback_log           : fonction(str) pour les logs UI

    Retourne un dict de résumé.
    """
    def log(msg):
        if callback_log:
            callback_log(msg)
        else:
            print(msg)

    # ── Chargement ───────────────────────────────────────────────
    if not os.path.exists(chemin_planning):
        log("❌ planning_historique.json introuvable")
        return {"erreur": "planning introuvable"}

    with open(chemin_planning, 'r', encoding='utf-8') as f:
        planning = json.load(f)

    cycles_employes = {}
    if os.path.exists(chemin_cycles_employes):
        with open(chemin_cycles_employes, 'r', encoding='utf-8') as f:
            cycles_employes = json.load(f)

    # ── Backup ───────────────────────────────────────────────────
    if os.path.exists(chemin_cycles_employes):
        try:
            from main import faire_backup
            chemin_bak = faire_backup(chemin_cycles_employes)
            if chemin_bak:
                log(f"💾 Backup : {os.path.basename(chemin_bak)}")
        except ImportError:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            bak = chemin_cycles_employes.replace(".json", f".backup_{ts}.json")
            shutil.copy2(chemin_cycles_employes, bak)
            log(f"💾 Backup : {os.path.basename(bak)}")

    # ── Détection ────────────────────────────────────────────────
    log(f"\n🔍 Analyse de {len(planning)} employé(s)...\n")

    nb_detectes   = 0
    nb_ignores    = 0
    nb_echecs     = 0
    nb_nouveaux   = 0

    for cle_emp, data in sorted(planning.items()):
        semaines_brutes = data.get('semaines', {})
        jours           = data.get('jours', {})
        semaines        = _fusionner_semaines(semaines_brutes, jours)

        # Vérifier si entrée déjà remplie manuellement
        existant = cycles_employes.get(cle_emp, {})
        deja_rempli = (
            existant.get('cycle', '').strip() != '' and
            existant.get('cycle_depart', '').strip() != ''
        )

        if deja_rempli and not ecraser_manuel:
            log(f"   ⏭️  {cle_emp.split('|')[0]} — conservé (saisi manuellement)")
            nb_ignores += 1
            continue

        # Lancer la détection
        resultat = detecter_cycle_employe(semaines)

        nom_affiche = cle_emp.split('|')[0]

        if resultat['detecte']:
            nb_detectes += 1
            if cle_emp not in cycles_employes:
                nb_nouveaux += 1
            # Mettre à jour l'entrée (sans le champ 'score' et 'note' internes)
            cycles_employes[cle_emp] = {
                "cycle":        resultat['cycle'],
                "cycle_depart": resultat['cycle_depart'],
                "date_depart":  resultat['date_depart'],
                "cycle_type":   resultat['cycle'],
                "motif":        resultat.get('motif', []),
            }
            log(f"   ✅ {nom_affiche} → {resultat['note']}")
        else:
            nb_echecs += 1
            # Créer l'entrée vide si elle n'existe pas
            if cle_emp not in cycles_employes:
                nb_nouveaux += 1
                cycles_employes[cle_emp] = {
                    "cycle": "", "cycle_depart": "", "date_depart": ""
                }
            log(f"   ⚠️  {nom_affiche} → {resultat['note']}")

    # ── Sauvegarde ───────────────────────────────────────────────
    # Trier alphabétiquement, COMMENTAIRE en premier
    commentaire = cycles_employes.pop("COMMENTAIRE", None)
    cycles_tries = dict(sorted(cycles_employes.items()))
    if commentaire:
        cycles_tries = {"COMMENTAIRE": commentaire, **cycles_tries}

    with open(chemin_cycles_employes, 'w', encoding='utf-8') as f:
        json.dump(cycles_tries, f, ensure_ascii=False, indent=2)

    # ── Résumé ───────────────────────────────────────────────────
    log(f"\n{'='*50}")
    log(f"📊 Résumé détection cycles :")
    log(f"   ✅ {nb_detectes} cycle(s) détecté(s)")
    log(f"   ⚠️  {nb_echecs} non détecté(s) (données insuffisantes ou irrégulières)")
    log(f"   ⏭️  {nb_ignores} conservé(s) (saisi(s) manuellement)")
    log(f"   🆕 {nb_nouveaux} nouvelle(s) entrée(s) créée(s)")
    log(f"   💾 Sauvegardé : {os.path.basename(chemin_cycles_employes)}")

    return {
        "detectes":  nb_detectes,
        "echecs":    nb_echecs,
        "ignores":   nb_ignores,
        "nouveaux":  nb_nouveaux,
    }


# ═══════════════════════════════════════════════════════════════
# POINT D'ENTRÉE CLI
# ═══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    import sys

    base = os.path.dirname(os.path.abspath(__file__))
    chemin_planning  = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, 'planning_historique.json')
    chemin_cycles    = sys.argv[2] if len(sys.argv) > 2 else os.path.join(base, 'cycles_employes.json')
    ecraser          = '--ecraser' in sys.argv

    print(f"📂 Planning  : {chemin_planning}")
    print(f"📂 Cycles    : {chemin_cycles}")
    if ecraser:
        print("⚠️  Mode --ecraser : les entrées manuelles seront écrasées")

    detecter_tous_cycles(chemin_planning, chemin_cycles, ecraser_manuel=ecraser)