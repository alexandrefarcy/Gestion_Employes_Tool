"""
migrer_employes.py
==================
Migration de employes_contrats.json vers le nouveau format.

Ancien format :
  { "NOM PRÉNOM|ID": "JJ-MM-AAAA" }

Nouveau format :
  {
    "NOM PRÉNOM|ID": {
      "date_debut": "JJ-MM-AAAA",
      "actif": true,
      "departements": []
    }
  }

Usage :
  python migrer_employes.py
  python migrer_employes.py --chemin /mon/dossier/employes_contrats.json
"""

import json
import os
import sys
import shutil
from datetime import datetime


def migrer(chemin: str) -> None:

    # ── 1. Vérification du fichier source ──────────────────────────────
    if not os.path.exists(chemin):
        print(f"❌ Fichier introuvable : {chemin}")
        sys.exit(1)

    with open(chemin, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # ── 2. Détecter si déjà migré ──────────────────────────────────────
    for cle, valeur in data.items():
        if cle == "COMMENTAIRE":
            continue
        if isinstance(valeur, dict):
            print("✅ Le fichier est déjà dans le nouveau format. Rien à faire.")
            return
        break  # on a vu la première entrée réelle, c'est suffisant

    # ── 3. Backup automatique ──────────────────────────────────────────
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    chemin_backup = chemin.replace(".json", f".backup_{timestamp}.json")
    shutil.copy2(chemin, chemin_backup)
    print(f"💾 Backup créé : {os.path.basename(chemin_backup)}")

    # ── 4. Migration ───────────────────────────────────────────────────
    nouveau = {}
    nb_migres = 0
    nb_ignores = 0

    for cle, valeur in data.items():
        # Supprimer la clé COMMENTAIRE
        if cle == "COMMENTAIRE":
            nb_ignores += 1
            continue

        # Ignorer les entrées sans pipe (mal formées)
        if "|" not in cle:
            print(f"   ⚠️  Ignoré (format inattendu) : {cle}")
            nb_ignores += 1
            continue

        # Convertir
        nouveau[cle] = {
            "date_debut":   str(valeur).strip(),
            "actif":        True,
            "departements": []
        }
        nb_migres += 1

    # ── 5. Sauvegarde du nouveau fichier ───────────────────────────────
    with open(chemin, 'w', encoding='utf-8') as f:
        json.dump(nouveau, f, ensure_ascii=False, indent=2)

    # ── 6. Rapport ─────────────────────────────────────────────────────
    print(f"✅ Migration terminée :")
    print(f"   {nb_migres} employé(s) migré(s)")
    if nb_ignores:
        print(f"   {nb_ignores} entrée(s) ignorée(s) (COMMENTAIRE ou format invalide)")
    print(f"   Fichier mis à jour : {os.path.basename(chemin)}")


if __name__ == "__main__":
    # Chemin par défaut : même dossier que le script
    chemin_defaut = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "employes_contrats.json"
    )

    # Permettre de passer un chemin en argument
    if "--chemin" in sys.argv:
        idx = sys.argv.index("--chemin")
        chemin = sys.argv[idx + 1] if idx + 1 < len(sys.argv) else chemin_defaut
    else:
        chemin = chemin_defaut

    print(f"📂 Fichier source : {chemin}")
    migrer(chemin)
