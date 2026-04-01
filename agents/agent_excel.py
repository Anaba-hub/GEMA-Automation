# ============================================================
# agents/agent_excel.py — Lecture et écriture des fichiers Excel
# ============================================================
# Source unique : gema.xlsx
# Colonnes gema : id | Prénom | Nom | Moyenne /4 | Année | École |
#                 Recherche LinkedIn | Poste 1 | Poste 2 | Vérifié GEMA
#
# Sortie : output_final.xlsx (14 colonnes)
# Colonnes sortie : id | Prénom | Nom | Année | École |
#                   Lien retenu | Lien source |
#                   Poste 1 | Société 1 | Période 1 |
#                   Poste 2 | Société 2 | Période 2 | Statut
#
# Lien retenu  = URL /in/ utilisée pour l'extraction
# Lien source  = URL d'origine dans gema (peut être /search/ ou /in/)
# ============================================================

import os
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from utils.logger import get_logger

log = get_logger("EXCEL")

# En-têtes du fichier de sortie (14 colonnes)
ENTETES_OUTPUT = [
    "id", "Prénom", "Nom", "Année", "École",
    "Lien retenu", "Lien source",
    "Poste 1", "Société 1", "Période 1",
    "Poste 2", "Société 2", "Période 2",
    "Statut",
]

# Valeurs considérées comme vides/invalides dans les cellules
_VIDES = {"", "nan", "#n/a", "#n/a!", "non trouvé", "none", "n/a", "-", "#value!"}


# ----------------------------------------------------------
# Fonctions publiques
# ----------------------------------------------------------

def est_lien_valide(valeur) -> bool:
    """
    Retourne True si la valeur est une URL LinkedIn utilisable.
    Rejette : None, "", "nan", "#N/A", "Non trouvé", toute chaîne
    sans "linkedin.com" ou ne commençant pas par "http".
    """
    if not valeur:
        return False
    s = str(valeur).strip()
    if s.lower() in _VIDES:
        return False
    if not s.startswith("http"):
        return False
    if "linkedin.com" not in s.lower():
        return False
    return True


def load_gema(chemin: str = "gema.xlsx") -> list:
    """
    Lit gema.xlsx et retourne une liste de dicts.

    Chaque dict contient :
      id, prenom, nom, annee, ecole, lien,
      poste1_gema, poste2_gema, deja_traite

    deja_traite = True si Poste 1 ET Poste 2 sont déjà remplis dans gema.
    Gère les cellules vides, "nan", "#N/A" proprement.
    """
    if not Path(chemin).exists():
        log.error(f"Fichier introuvable : {chemin}")
        return []

    wb = openpyxl.load_workbook(chemin)
    ws = wb.active

    # Mapping entête → index colonne (0-based)
    entetes = {}
    for idx, cell in enumerate(ws[1]):
        if cell.value:
            entetes[str(cell.value).strip()] = idx

    def val(row, key):
        """Valeur texte d'une colonne, ou '' si absente/invalide."""
        idx = entetes.get(key, -1)
        if idx < 0 or idx >= len(row):
            return ""
        v = row[idx].value
        if v is None:
            return ""
        s = str(v).strip()
        return "" if s.lower() in _VIDES else s

    personnes = []
    for row in ws.iter_rows(min_row=2):
        nom    = val(row, "Nom")
        prenom = val(row, "Prénom")
        if not nom and not prenom:
            continue  # ligne vide

        poste1 = val(row, "Poste 1")
        poste2 = val(row, "Poste 2")

        personnes.append({
            "id":          val(row, "id"),
            "prenom":      prenom,
            "nom":         nom,
            "annee":       val(row, "Année"),
            "ecole":       val(row, "École"),
            "lien":        val(row, "Recherche LinkedIn"),
            "poste1_gema": poste1,
            "poste2_gema": poste2,
            # deja_traite : les deux postes sont déjà connus dans gema → skip extraction
            "deja_traite": bool(poste1 and poste2),
        })

    log.info(f"{len(personnes)} personne(s) chargée(s) depuis '{chemin}'.")
    return personnes


def save_row(chemin: str, data: dict) -> None:
    """
    Ajoute ou met à jour une ligne dans output_final.xlsx.

    Recherche une ligne existante par 'id' (si non vide),
    sinon par 'nom' + 'prenom'. Met à jour si trouvée, ajoute sinon.
    Crée le fichier avec en-têtes stylisés si inexistant.
    Sauvegarde atomique via .tmp + os.replace.

    data doit contenir les clés :
      id, prenom, nom, annee, ecole,
      lien_retenu, lien_source,
      poste1, societe1, periode1,
      poste2, societe2, periode2,
      statut
    """
    if not Path(chemin).exists():
        _creer_output(chemin)

    wb = openpyxl.load_workbook(chemin)
    ws = wb.active

    # Mapping entête → numéro de colonne (1-based)
    idx_map = {
        str(cell.value or "").strip(): cell.column
        for cell in ws[1]
    }

    id_data     = str(data.get("id", "")).strip()
    nom_data    = str(data.get("nom", "")).strip()
    prenom_data = str(data.get("prenom", "")).strip()

    # Chercher une ligne existante à mettre à jour
    ligne_cible = None
    col_id     = idx_map.get("id", 1)
    col_nom    = idx_map.get("Nom", 3)
    col_prenom = idx_map.get("Prénom", 2)

    for row in ws.iter_rows(min_row=2):
        id_cell     = str(row[col_id - 1].value or "").strip()
        nom_cell    = str(row[col_nom - 1].value or "").strip()
        prenom_cell = str(row[col_prenom - 1].value or "").strip()

        if id_data and id_cell == id_data:
            ligne_cible = row[0].row
            break
        if not id_data and nom_cell == nom_data and prenom_cell == prenom_data:
            ligne_cible = row[0].row
            break

    # Valeurs à écrire dans l'ordre de ENTETES_OUTPUT
    valeurs = [
        data.get("id", ""),
        data.get("prenom", ""),
        data.get("nom", ""),
        data.get("annee", ""),
        data.get("ecole", ""),
        data.get("lien_retenu", ""),   # URL /in/ résolue
        data.get("lien_source", ""),   # URL d'origine gema (/search/ ou /in/)
        data.get("poste1", ""),
        data.get("societe1", ""),
        data.get("periode1", ""),
        data.get("poste2", ""),
        data.get("societe2", ""),
        data.get("periode2", ""),
        data.get("statut", ""),
    ]

    if ligne_cible:
        for col_idx, valeur in enumerate(valeurs, start=1):
            ws.cell(row=ligne_cible, column=col_idx, value=valeur)
        log.info(
            f"Mise à jour : {prenom_data} {nom_data} → {data.get('statut', '')}"
        )
    else:
        ws.append(valeurs)
        log.info(
            f"Nouvelle ligne : {prenom_data} {nom_data} → {data.get('statut', '')}"
        )

    # Ajustement automatique de la largeur des colonnes
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 70)

    # Sauvegarde atomique
    chemin_tmp = chemin + ".tmp"
    wb.save(chemin_tmp)
    os.replace(chemin_tmp, chemin)


# ----------------------------------------------------------
# Helper interne
# ----------------------------------------------------------

def _creer_output(chemin: str) -> None:
    """Crée output_final.xlsx avec les en-têtes stylisés."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Résultats"
    ws.append(ENTETES_OUTPUT)

    fill  = PatternFill(fill_type="solid", fgColor="1F4E79")
    font  = Font(bold=True, color="FFFFFF", size=11)
    align = Alignment(horizontal="center", vertical="center")
    for cell in ws[1]:
        cell.fill      = fill
        cell.font      = font
        cell.alignment = align
    ws.row_dimensions[1].height = 20

    wb.save(chemin)
    log.info(f"Fichier créé : {chemin}")
