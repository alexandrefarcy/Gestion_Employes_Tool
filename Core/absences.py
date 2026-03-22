import os
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    print("Erreur: pdfplumber n'est pas installé. Installation en cours...")
    os.system("pip install pdfplumber")
    import pdfplumber


# Codes d'absence à rechercher
ABSENCE_CODES = [
    "3221", "3251", "3210", "31C3", "31E3", "3053",
    "3043", "3243", "3261", "3033", "3001", "31B3",
    "3111", "3021", "31G3"
]

# Codes présents dans le tableau "INFORMATIONS JOURNALIERES" (2024-2025)
TABLE_CODES = [
    "AA", "AM","AT", "CP", "CT", "JF", "JM", "JS","JT", "MA", "NN", "NP","PA", "RJ", "RN", "RP", "TH", "ZV"
]
# Chemin racine des fiches de paye
PDF_ROOT_PATH = r"I:\Dpt_Drh\Prive\ALEXANDRE\20260205_HeuresNuit\Cycle\Fiche de paye"

# Chemin de sortie du JSON
OUTPUT_JSON = r"I:\Dpt_Drh\Prive\ALEXANDRE\20260205_HeuresNuit\Cycle\RH_Tool\Absence_Python.json"

# Chemin de debug
DEBUG_TXT_PATH = r"I:\Dpt_Drh\Prive\ALEXANDRE\20260205_HeuresNuit\Cycle\debug_tableau.txt"

# Chemin du fichier employés
EMPLOYES_JSON = r"I:\Dpt_Drh\Prive\ALEXANDRE\20260205_HeuresNuit\Cycle\RH_Tool\employes_contrats.json"


def normalize_name(s):
    """Normalise un nom en majuscules sans accents et sans espaces superflus."""
    if not s:
        return ""
    s = s.replace('\u00A0', ' ')
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s.strip().upper()


# Cache des employés
EMPLOYES_CACHE = None


def load_employes():
    """
    Charge la liste des employés depuis le JSON
    Retourne un dict {nom_upper: {"id": "...", "date": "..."}}G
    """
    global EMPLOYES_CACHE
    
    if EMPLOYES_CACHE is not None:
        return EMPLOYES_CACHE
    
    EMPLOYES_CACHE = {}
    
    try:
        if not os.path.exists(EMPLOYES_JSON):
            print(f"⚠️  Fichier employés non trouvé: {EMPLOYES_JSON}")
            return EMPLOYES_CACHE
        
        with open(EMPLOYES_JSON, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Traiter chaque employé
        for key, info in data.items():
            if key == "COMMENTAIRE":
                continue
            if '|' in key:
                nom, emp_id = key.split('|')
                nom = nom.strip()
                norm = normalize_name(nom)
                # employes_contrats.json v3 : valeur = dict
                if isinstance(info, dict):
                    date_debut = info.get("date_debut", "")
                else:
                    date_debut = info
                EMPLOYES_CACHE[norm] = {
                    "id": emp_id.strip(),
                    "date_debut": date_debut,
                    "nom_original": nom
                }
        
        return EMPLOYES_CACHE
    
    except Exception as e:
        print(f"❌ Erreur lors de la lecture du fichier employés: {e}")
        return {}


def verify_employee(search_name):
    """
    Vérifie si un employé existe dans la liste
    Retourne (existe, info_employé) ou (False, None) si n'existe pas
    """
    employes = load_employes()
    
    # Normaliser le nom recherché
    search_norm = normalize_name(search_name)
    
    # Recherche exacte
    if search_norm in employes:
        return True, employes[search_norm]
    
    # Recherche partielle (contient)
    matches = []
    for nom, info in employes.items():
        if search_norm in nom or nom in search_norm:
            matches.append((nom, info))
    
    if len(matches) == 1:
        return True, matches[0][1]
    elif len(matches) > 1:
        # Plusieurs correspondances partielles
        print(f"\n⚠️  Plusieurs employés correspondent à '{search_name}':")
        for nom, info in matches:
            print(f"   • {nom} (ID: {info['id']})")
        return False, None
    
    return False, None


def export_table_debug(pdf_path, year_folder, output_file):
    """
    Exporte le contenu complet du tableau dans un fichier texte pour debug
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"DEBUG - Extraction tableau PDF\n")
                f.write(f"{'='*80}\n")
                f.write(f"Fichier: {os.path.basename(pdf_path)}\n")
                f.write(f"Année: {year_folder}\n")
                f.write(f"Chemin complet: {pdf_path}\n")
                f.write(f"{'='*80}\n\n")
                
                # Parcourir chaque page
                for page_num, page in enumerate(pdf.pages, 1):
                    text = page.extract_text()
                    if not text:
                        f.write(f"\n--- PAGE {page_num} ---\n")
                        f.write("⚠️ Pas de texte trouvé\n")
                        continue
                    
                    lines = text.split('\n')
                    
                    f.write(f"\n--- PAGE {page_num} ({len(lines)} lignes) ---\n")
                    f.write(f"{'='*80}\n")
                    
                    # Chercher la plage DU...AU
                    text_joined = "\n".join(lines)
                    duau_match = re.search(
                        r"\bDU\b\s*(\d{1,2}[\-/]\d{1,2}[\-/]\d{2,4})\s*\bAU\b\s*(\d{1,2}[\-/]\d{1,2}[\-/]\d{2,4})",
                        text_joined,
                        re.IGNORECASE
                    )
                    
                    if duau_match:
                        f.write(f"✓ Plage détectée: DU {duau_match.group(1)} AU {duau_match.group(2)}\n")
                    else:
                        f.write(f"⚠️ Aucune plage DU/AU détectée\n")
                    
                    f.write(f"\nCodes recherchés: {', '.join(TABLE_CODES)}\n")
                    f.write(f"{'='*80}\n\n")
                    
                    # Afficher chaque ligne avec analyse
                    day_pattern = re.compile(r"^\s*[A-Za-z]?\s*(\d{1,2})\s+")
                    
                    for i, line in enumerate(lines):
                        if not line.strip():
                            continue
                        
                        # Analyzer la ligne
                        m = day_pattern.search(line)
                        day_match = ""
                        if m:
                            day_num = int(m.group(1))
                            day_match = f"[JOUR {day_num}]"
                        
                        # Chercher codes
                        codes_found = []
                        for code in TABLE_CODES:
                            if re.search(rf"\b{re.escape(code)}\b", line, re.IGNORECASE):
                                codes_found.append(code)
                        
                        codes_str = f"CODES: {', '.join(codes_found)}" if codes_found else ""
                        
                        f.write(f"[{i:3d}] {day_match:12} {codes_str:30} | {line}\n")
        
        print(f"✅ Debug exporté: {output_file}")
        return True
    except Exception as e:
        print(f"❌ Erreur lors de l'export debug: {e}")
        return False


def export_person_debug(pdf_path, search_name, page_num, lines):
    """
    Exporte les lignes d'une personne trouvée dans un PDF pour debug
    """
    try:
        filename = os.path.basename(pdf_path)
        # Crée un fichier debug spécifique pour cette personne et ce PDF
        clean_name = search_name.upper().replace(" ", "_")
        clean_filename = filename.replace('.pdf', '').replace(' ', '_')
        output_file = os.path.join(
            os.path.dirname(DEBUG_TXT_PATH),
            f"debug_{clean_filename}_{clean_name}.txt"
        )
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"DEBUG - Tableau de jours\n")
            f.write(f"{'='*100}\n")
            f.write(f"Fichier: {filename}\n")
            f.write(f"Page: {page_num}\n")
            f.write(f"Collaborateur: {search_name}\n")
            f.write(f"Codes recherchés: {', '.join(TABLE_CODES)}\n")
            f.write(f"{'='*100}\n\n")
            
            # Afficher toutes les lignes de la page avec analyse
            day_pattern = re.compile(r"^\s*[A-Za-z]?\s*(\d{1,2})\s+")
            
            f.write("ANALYSE DU TABLEAU LIGNE PAR LIGNE:\n")
            f.write(f"{'-'*100}\n")
            f.write(f"{'Ligne':>5} {'Jour':>6} {'Codes trouvés':>30} | Contenu\n")
            f.write(f"{'-'*100}\n")
            
            for i, line in enumerate(lines):
                if not line.strip():
                    continue
                
                # Analyser la ligne
                m = day_pattern.search(line)
                day_match = ""
                if m:
                    day_num = int(m.group(1))
                    day_match = f"[J{day_num:2d}]"
                
                # Chercher codes
                codes_found = []
                for code in TABLE_CODES:
                    if re.search(rf"\b{re.escape(code)}\b", line, re.IGNORECASE):
                        codes_found.append(code)
                
                codes_str = f"{', '.join(codes_found)}" if codes_found else ""
                
                f.write(f"{i:5d} {day_match:>6} {codes_str:>30} | {line[:70]}\n")
        
        print(f"      ✅ Debug généré: {output_file}")
        return True
    
    except Exception as e:
        print(f"      ❌ Erreur lors de l'export debug: {e}")
        return False


def convert_date_format(date_str):
    """
    Convertit une date du format JJ/MM/YY au format JJ-MM-AAAA
    Exemple: 23/12/22 -> 23-12-2022
    """
    try:
        # Essayer différents formats de date
        for fmt in ["%d/%m/%y", "%d/%m/%Y", "%d-%m-%y", "%d-%m-%Y"]:
            try:
                dt = datetime.strptime(date_str.strip(), fmt)
                return dt.strftime("%d-%m-%Y")
            except ValueError:
                continue
        return date_str
    except:
        return date_str


def extract_dates_from_line(line):
    """
    Extrait les dates d'une ligne d'absence
    Format attendu: "code Autorisée non 23/12/22 au 23/12/22"
    Retourne tuple (date_debut, date_fin)
    """
    # Chercher le pattern de date JJ/MM/YY
    date_pattern = r'(\d{1,2}/\d{1,2}/\d{2,4})'
    matches = re.findall(date_pattern, line)
    
    if len(matches) >= 2:
        # Deux dates trouvées (début et fin)
        debut = convert_date_format(matches[0])
        fin = convert_date_format(matches[1])
        return debut, fin
    elif len(matches) == 1:
        # Une seule date, utiliser la même pour début et fin
        date = convert_date_format(matches[0])
        return date, date
    
    return None, None


def group_dates_to_ranges(dates):
    """
    Regroupe une liste de dates (datetime.date) contiguës en plages.
    Retourne une liste de dicts {"debut": "dd-mm-YYYY", "fin": "dd-mm-YYYY"}.
    """
    if not dates:
        return []
    dates = sorted(dates)
    ranges = []
    start = dates[0]
    end = dates[0]

    for d in dates[1:]:
        if (d - end).days == 1:
            end = d
        else:
            ranges.append({
                "debut": start.strftime("%d-%m-%Y"),
                "fin": end.strftime("%d-%m-%Y")
            })
            start = d
            end = d

    ranges.append({
        "debut": start.strftime("%d-%m-%Y"),
        "fin": end.strftime("%d-%m-%Y")
    })
    return ranges


def parse_table_absences(lines, start_dt, end_dt, table_codes):
    """
    Parse un tableau 'INFORMATIONS JOURNALIERES' donné sous forme de lignes de texte.
    start_dt et end_dt sont des datetime.date indiquant la plage affichée dans le tableau.
    table_codes: liste de codes (ex: JF, CP, AM...).
    Format 2024+: les infos sont à la fin: "... L 15 7 00 AM" où L=jour, 15=jour du mois, AM=code
    Retourne une liste de dicts {"debut":..., "fin":...} correspondant aux plages d'absence.
    """
    # Construire mapping jour->date pour la plage
    day_to_date = {}
    cur = start_dt
    from datetime import timedelta
    while cur <= end_dt:
        day_to_date[cur.day] = cur
        cur = cur + timedelta(days=1)

    absence_days = []

    # Pattern pour détecter les infos journalières à la fin des lignes
    # Format: [lettre] [chiffres] [heure] [CODE]
    # Exemple: L 15 7 00 AM ou M 16 7 00 AM ou J 25 7 00 CP
    jour_pattern = re.compile(r'[A-Z]\s+(\d{1,2})\s+\d+\s+\d+')  # Extrait le jour
    
    for line in lines:
        # Chercher un code d'absence à la fin de la ligne
        code_found = None
        for code in table_codes:
            # Chercher le code avec limite de mot à la fin
            if re.search(rf'\b{re.escape(code)}\s*$', line, re.IGNORECASE):
                code_found = code
                break
        
        # Si un code d'absence est trouvé, extraire le jour du mois
        if code_found:
            m = jour_pattern.search(line)
            if m:
                day_num = int(m.group(1))
                if day_num in day_to_date:
                    absence_days.append(day_to_date[day_num])

    # Regrouper jours consécutifs en plages
    ranges = group_dates_to_ranges(absence_days)
    return ranges



def _extraire_nom_page(pdf, page_idx, normalize_fn):
    """
    Extrait le nom de l'employé sur une page donnée (index 0-based).
    Gère deux formats :
      - 2021 : "SIRET :... M. NOM Prenom"  ou  "SIRET :... Mme NOM Prenom"
      - 2025 : "CONVENTION DE L'INDUSTRIE NOM PRENOM"
    Retourne le nom normalisé (ex: "ALARY SANDRINE") ou None.
    """
    import re
    try:
        text = pdf.pages[page_idx].extract_text()
        if not text:
            return None
        for line in text.split('\n')[:15]:
            norm = normalize_fn(line)
            # Format 2025 : "CONVENTION DE L'INDUSTRIE NOM PRENOM"
            m = re.search(r"CONVENTION DE L.INDUSTRIE\s+([A-Z][A-Z\s\-]{2,})", norm)
            if m:
                return m.group(1).strip()
            # Format 2021 : "SIRET :... M. NOM Prenom" ou "... Mme NOM Prenom"
            m = re.search(r"\bM(?:ME?)?\b\.?\s+([A-Z][A-Z\s\-]{2,})", norm)
            if m:
                return normalize_fn(m.group(1).strip())
    except Exception:
        pass
    return None


def _binary_search_page(pdf, norm_search, normalize_fn):
    """
    Recherche dichotomique de la page contenant norm_search dans un PDF trié
    alphabétiquement (plusieurs pages par fiche possible).
    Retourne l'index de la PREMIÈRE page de la fiche (0-based) ou None.
    """
    n = len(pdf.pages)
    lo, hi = 0, n - 1

    while lo <= hi:
        mid = (lo + hi) // 2
        nom_mid = _extraire_nom_page(pdf, mid, normalize_fn)

        if nom_mid is None:
            hi = mid - 1
            continue

        if norm_search in nom_mid or nom_mid in norm_search:
            # Trouvé — remonter à la première page de cette fiche
            while mid > 0:
                nom_prev = _extraire_nom_page(pdf, mid - 1, normalize_fn)
                if nom_prev and (norm_search in nom_prev or nom_prev in norm_search):
                    mid -= 1
                else:
                    break
            return mid

        if nom_mid < norm_search:
            lo = mid + 1
        else:
            hi = mid - 1

    # Zone de sécurité ±3 pages autour du point d'atterrissage
    for idx in range(max(0, lo - 3), min(n - 1, lo + 3) + 1):
        nom = _extraire_nom_page(pdf, idx, normalize_fn)
        if nom and (norm_search in nom or nom in norm_search):
            return idx

    return None


def search_absences_in_pdf(pdf_path, search_name, year_folder):
    """
    Recherche les absences dans un PDF pour une personne donnée
    Les absences sont cherchées UNIQUEMENT sur la première page contenant le nom
    year_folder: année (ex. '2024') pour déterminer quel système appliquer
    Retourne un dictionnaire {nom: [{"debut": "...", "fin": "..."}, ...]}
    """
    absences = {}
    current_name = None
    lines_checked = 0
    absences_found_in_pdf = 0
    
    try:
        year_num = int(year_folder)
    except:
        year_num = 2023  # Par défaut, ancien système
    
    def normalize_text(s):
        if not s:
            return ""
        s = s.replace('\u00A0', ' ')
        s = unicodedata.normalize('NFD', s)
        s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
        return s.upper()

    norm_search = normalize_text(search_name)

    try:
        print(f"    📂 Ouverture du PDF...", end=" ")
        with pdfplumber.open(pdf_path) as pdf:
            nb_pages = len(pdf.pages)
            print(f"✓ ({nb_pages} pages)")

            # ── Recherche dichotomique ────────────────────────────────────
            print(f"    🔍 Recherche binaire pour '{norm_search}'...", end=" ")
            page_idx = _binary_search_page(pdf, norm_search, normalize_text)

            person_found_on_page = False
            if page_idx is None:
                print(f"non trouvé (≈{int(nb_pages**0.5)+2} pages lues)")
            else:
                print(f"trouvé page {page_idx + 1}")
                text = pdf.pages[page_idx].extract_text() or ""
                lines = text.split('\n')
                lines_checked += len(lines)
                normalized_lines = [normalize_text(l) for l in lines]

                for nline in normalized_lines:
                    if norm_search in nline:
                        person_found_on_page = True
                        current_name = search_name
                        if current_name not in absences:
                            absences[current_name] = []
                        print(f"    ✓ Personne confirmée à la page {page_idx + 1}")
                        if "02 - Février 2024" in os.path.basename(pdf_path) or "fevrier" in os.path.basename(pdf_path).lower():
                            export_person_debug(pdf_path, search_name, page_idx + 1, lines)
                        break

                # Si on a trouvé le nom sur cette page, chercher les absences sur cette même page
                if person_found_on_page:
                    # Déterminer le système à appliquer selon l'année
                    if year_num >= 2024:
                        # Nouveau système (2024+): tableau INFORMATIONS JOURNALIERES uniquement
                        page_text_joined = "\n".join(lines)
                        
                        # Pattern plus flexible pour "DU ... AU" (accepte espaces autour des tirets)
                        duau_patterns = [
                            r"\bDU\b\s*(\d{1,2}[\-/]\d{1,2}[\-/]\d{2,4})\s*\bAU\b\s*(\d{1,2}[\-/]\d{1,2}[\-/]\d{2,4})",
                            r"\bDU\b\s*(\d{1,2}\s*[\-/]\s*\d{1,2}\s*[\-/]\s*\d{2,4})\s*\bAU\b\s*(\d{1,2}\s*[\-/]\s*\d{1,2}\s*[\-/]\s*\d{2,4})"
                        ]
                        
                        duau_match = None
                        for pattern in duau_patterns:
                            duau_match = re.search(pattern, page_text_joined, re.IGNORECASE)
                            if duau_match:
                                print(f"      🔍 Pattern trouvé (pattern {duau_patterns.index(pattern)+1})")
                                break

                        if duau_match:
                            start_s = duau_match.group(1).replace(" ", "")  # Nettoyer les espaces
                            end_s = duau_match.group(2).replace(" ", "")
                            start_dt = None
                            end_dt = None
                            for fmt in ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"):
                                try:
                                    start_dt = datetime.strptime(start_s, fmt).date()
                                    end_dt = datetime.strptime(end_s, fmt).date()
                                    break
                                except Exception:
                                    continue
                            if start_dt and end_dt:
                                print(f"      🧾 Tableau détecté: DU {start_dt.strftime('%d-%m-%Y')} AU {end_dt.strftime('%d-%m-%Y')}")
                                ranges = parse_table_absences(lines, start_dt, end_dt, TABLE_CODES)
                                if ranges:
                                    absences[current_name].extend(ranges)
                                    for r in ranges:
                                        print(f"      🔍 Tableau -> Du {r['debut']} au {r['fin']}")
                                else:
                                    print(f"      ⚠️  Aucun code d'absence détecté dans le tableau")
                            else:
                                print(f"      ⚠️  Impossible d'analyser les dates DU/AU du tableau")
                        else:
                            print(f"      ⚠️  Aucun tableau DU/AU trouvé sur cette page")

                    else:
                        # Ancien système (2021-2023): codes numériques uniquement
                        for line in lines:
                            for code in ABSENCE_CODES:
                                pattern = rf"{re.escape(code)}\s*[.\-]?\s*\d+"
                                if re.search(pattern, line, re.IGNORECASE):
                                    debut, fin = extract_dates_from_line(line)
                                    if debut and fin:
                                        absences[current_name].append({
                                            "debut": debut,
                                            "fin": fin
                                        })
                                        absences_found_in_pdf += 1
                                        print(f"      🔍 (Ancien format) Code {code} trouvé: {debut} au {fin}")
                                        print(f"         Ligne: {line.strip()[:120]}...")
                                    else:
                                        print(f"      ⚠️  (Ancien format) Code {code} trouvé mais dates non détectées")
                                        print(f"         Ligne: {line.strip()[:120]}...")
                                    break

                    page_num = page_idx + 1
                    print(f"    📊 Résumé page {page_num}: {len(absences.get(current_name, []))} absence(s) trouvée(s)")

        if not person_found_on_page:
            if page_idx is None:
                print(f"    ⚠️  Personne non trouvée dans ce PDF")
    
    except FileNotFoundError:
        print(f"\n    ❌ ERREUR: Le fichier PDF n'existe pas ou n'est pas accessible")
    except Exception as e:
        print(f"\n    ❌ Erreur lors de la lecture du PDF: {e}")
        return {}
    
    return absences


def process_all_pdfs(search_name, date_debut=None):
    """
    Traite tous les PDFs des sous-dossiers (2021, 2022, 2023)
    pour rechercher les absences de la personne
    date_debut: date de début du contrat "JJ-MM-YYYY" pour filtrer les années
    """
    all_absences = {}
    total_pdfs = 0
    pdfs_processed = 0
    
    if not os.path.exists(PDF_ROOT_PATH):
        print(f"❌ Erreur: Le chemin n'existe pas")
        print(f"   Chemin attendu: {PDF_ROOT_PATH}")
        return all_absences
    
    print(f"\n📁 Chemin source: {PDF_ROOT_PATH}")
    print(f"✓ Le répertoire existe et est accessible\n")
    
    # Parcourir les dossiers par année
    year_folders = sorted([f for f in os.listdir(PDF_ROOT_PATH) 
                          if os.path.isdir(os.path.join(PDF_ROOT_PATH, f))])
    
    # Filtrer les années si date de début spécifiée
    if date_debut:
        try:
            year_debut = int(date_debut.split('-')[-1])  # Extrait l'année de "JJ-MM-YYYY"
            year_folders = [y for y in year_folders if int(y) >= year_debut]
            print(f"📅 Date de début du contrat: {date_debut}")
            print(f"   Recherche à partir de: {year_debut}\n")
        except:
            pass  # Si parsing échoue, traiter toutes les années
    
    if not year_folders:
        print("⚠️  Aucun sous-dossier trouvé dans le répertoire racine")
        return all_absences
    
    print(f"📂 Dossiers à traiter: {', '.join(year_folders)}\n")
    
    for year_folder in year_folders:
        year_path = os.path.join(PDF_ROOT_PATH, year_folder)
        
        print(f"\n{'='*70}")
        print(f"📅 Traitement du dossier: {year_folder}")
        print(f"{'='*70}")
        
        # Lister tous les PDFs dans le dossier
        pdfs_in_folder = [f for f in os.listdir(year_path) 
                         if f.lower().endswith('.pdf')]
        
        if not pdfs_in_folder:
            print(f"  ⚠️  Aucun fichier PDF trouvé dans ce dossier")
            continue
        
        print(f"  📄 PDFs trouvés: {len(pdfs_in_folder)}")
        for pdf_file in pdfs_in_folder:
            print(f"     - {pdf_file}")
        
        print(f"\n  Traitement en cours...")
        
        # Parcourir tous les PDFs dans le dossier
        for filename in pdfs_in_folder:
            pdf_full_path = os.path.join(year_path, filename)
            total_pdfs += 1
            
            print(f"\n  [{total_pdfs}] 📖 {filename}")
            
            absences = search_absences_in_pdf(pdf_full_path, search_name, year_folder)
            pdfs_processed += 1
            
            # Fusionner les absences trouvées
            for name, absences_list in absences.items():
                if name not in all_absences:
                    all_absences[name] = []
                all_absences[name].extend(absences_list)
    
    print(f"\n\n{'='*70}")
    print(f"📊 RÉSUMÉ DE LA RECHERCHE")
    print(f"{'='*70}")
    print(f"  PDFs traités: {pdfs_processed}/{total_pdfs}")
    print(f"  Dossiers analysés: {len(year_folders)}")
    
    return all_absences


def remove_duplicates(absences_dict):
    """
    Supprime les doublons d'absences
    """
    for name in absences_dict:
        # Supprimer les doublons en utilisant un set
        unique_absences = []
        seen = set()
        
        for absence in absences_dict[name]:
            key = (absence["debut"], absence["fin"])
            if key not in seen:
                seen.add(key)
                unique_absences.append(absence)
        
        # Trier par date de début
        unique_absences.sort(
            key=lambda x: datetime.strptime(x["debut"], "%d-%m-%Y")
        )
        absences_dict[name] = unique_absences
    
    return absences_dict


def save_to_json(absences, output_path):
    """
    Sauvegarde les absences au format JSON
    """
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(absences, f, ensure_ascii=False, indent=2)
        print(f"\n✓ JSON sauvegardé: {output_path}")
        return True
    except Exception as e:
        print(f"Erreur lors de la sauvegarde du JSON: {e}")
        return False


def main():
    """
    Fonction principale
    """
    print("\n" + "="*70)
    print("🔍 SCRIPT DE RECHERCHE D'ABSENCES DANS LES FICHES DE PAYE")
    print("="*70)
    
    print("\n📋 Menu:")
    print("  1. Rechercher les absences d'une personne")
    print("  2. Générer un debug des tableaux (Février 2024)")
    
    choice = input("\nChoisissez une option (1 ou 2) [défaut: 1]: ").strip()
    if not choice:
        choice = "1"
    
    if choice == "2":
        # Mode DEBUG
        print("\n" + "="*70)
        print("🔧 MODE DEBUG - Export tableau Février 2024")
        print("="*70)
        
        feb_2024_path = os.path.join(PDF_ROOT_PATH, "2024", "Février")
        
        if not os.path.exists(feb_2024_path):
            print(f"❌ Dossier non trouvé: {feb_2024_path}")
            # Lister les dossiers disponibles
            if os.path.exists(os.path.join(PDF_ROOT_PATH, "2024")):
                available = os.listdir(os.path.join(PDF_ROOT_PATH, "2024"))
                print(f"\n📁 Dossiers disponibles en 2024:")
                for folder in sorted(available):
                    print(f"  - {folder}")
            return
        
        # Lister les PDFs du mois
        pdfs = [f for f in os.listdir(feb_2024_path) if f.lower().endswith('.pdf')]
        
        if not pdfs:
            print(f"❌ Aucun PDF trouvé dans {feb_2024_path}")
            return
        
        print(f"\n📄 PDFs trouvés ({len(pdfs)}):")
        for i, pdf in enumerate(pdfs, 1):
            print(f"  {i}. {pdf}")
        
        # Choisir un PDF
        choice_pdf = input(f"\nChoisissez un PDF (1-{len(pdfs)}) [défaut: 1]: ").strip()
        if not choice_pdf:
            choice_pdf = "1"
        
        try:
            pdf_idx = int(choice_pdf) - 1
            if pdf_idx < 0 or pdf_idx >= len(pdfs):
                print("❌ Choix invalide")
                return
        except ValueError:
            print("❌ Entrée invalide")
            return
        
        selected_pdf = pdfs[pdf_idx]
        pdf_full_path = os.path.join(feb_2024_path, selected_pdf)
        
        print(f"\n🔍 Analyse de: {selected_pdf}")
        print(f"   Chemin: {pdf_full_path}\n")
        
        # Exporter le debug
        output_debug = DEBUG_TXT_PATH.replace(".txt", f"_{selected_pdf.replace('.pdf', '')}.txt")
        if export_table_debug(pdf_full_path, "2024", output_debug):
            print(f"\n✅ Fichier debug créé!")
            print(f"   Emplacement: {output_debug}")
            print(f"\n💡 Ouvrez ce fichier pour voir comment le PDF est parsé")
        return
    
    # Mode NORMAL (choice == "1")
    # Demander le nom et prénom
    search_name = input("\n👤 Nom: ").strip()
    
    if not search_name:
        print("❌ Erreur: Vous devez entrer un nom")
        return
    
    # Vérifier que l'employé existe
    print(f"\n🔍 Vérification de l'employé '{search_name}'...")
    employe_existe, employe_info = verify_employee(search_name)
    
    if not employe_existe:
        print(f"\n❌ Employé '{search_name}' NOT FOUND")
        print(f"\n💡 Astuces:")
        print(f"   • Vérifiez l'orthographe du nom")
        print(f"   • Les noms doivent être EN MAJUSCULES (ex: ALVES JOEL)")
        print(f"\n📋 Chargement de la liste des employés disponibles...")
        
        employes = load_employes()
        if employes:
            print(f"\n✓ {len(employes)} employés disponibles:")
            for nom in sorted(employes.keys())[:20]:  # Afficher les 20 premiers
                info = employes[nom]
                print(f"   • {nom} (ID: {info['id']})")
            if len(employes) > 20:
                print(f"   ... et {len(employes) - 20} autres")
        return
    
    # Employé trouvé
    print(f"✅ Employé trouvé!")
    print(f"   Nom: {employe_info['nom_original']}")
    print(f"   ID: {employe_info['id']}")
    print(f"   Date de début: {employe_info['date_debut']}")
    
    print(f"\n🔎 Recherche en cours pour: {search_name.upper()}")
    print(f"   (La recherche est insensible à la casse)")
    
    # Traiter tous les PDFs (avec filtrage par date de début)
    all_absences = process_all_pdfs(search_name, employe_info['date_debut'])
    
    # Supprimer les doublons et trier
    all_absences = remove_duplicates(all_absences)
    
    # Vérifier les résultats
    print(f"\n\n{'='*70}")
    print(f"✅ RÉSULTATS FINAUX")
    print(f"{'='*70}\n")
    
    if all_absences and any(all_absences.values()):
        for name, absences_list in all_absences.items():
            if absences_list:
                print(f"📋 Personne: {name}")
                print(f"   Total: {len(absences_list)} absence(s) trouvée(s)\n")
                for idx, absence in enumerate(absences_list, 1):
                    print(f"   {idx}. Du {absence['debut']} au {absence['fin']}")
                print()
        
        # Sauvegarder en JSON
        print(f"{'='*70}")
        print(f"💾 Sauvegarde du JSON...")
        print(f"{'='*70}")
        if save_to_json(all_absences, OUTPUT_JSON):
            print(f"✅ Fichier créé avec succès!")
            print(f"   Emplacement: {OUTPUT_JSON}")
        else:
            print(f"❌ Erreur lors de la sauvegarde du JSON")
    else:
        print(f"⚠️  Aucune absence trouvée pour: {search_name.upper()}")
        print(f"\nVérifiez que:")
        print(f"  • Le nom saisi est correct")
        print(f"  • Les PDFs sont bien présents dans les dossiers")
        print(f"  • Les fiches de paye contiennent ce nom")
    
    print(f"\n{'='*70}\n")


if __name__ == "__main__":
    main()