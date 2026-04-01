# ============================================================
# utils/helpers.py — Fonctions utilitaires partagées
# ============================================================

import json
import time
import random
from datetime import datetime
from pathlib import Path


def delai_aleatoire(min_s: float, max_s: float) -> float:
    """Attend un nombre aléatoire de secondes entre min_s et max_s. Retourne la durée."""
    duree = random.uniform(min_s, max_s)
    time.sleep(duree)
    return duree


# ----------------------------------------------------------
# Progression (format robuste pour 1 600 profils)
# ----------------------------------------------------------

def charger_progression(chemin: str) -> dict:
    """
    Charge le fichier progress.json.
    Structure :
    {
      "dernier_index":   45,
      "total":           1600,
      "timestamp":       "2026-04-01 23:45:12",
      "traites_phase1":  ["Dupont_Marie", ...],
      "traites_phase2":  ["Dupont_Marie", ...],
      "statuts":         {"Dupont_Marie": "terminé", ...},
      "resultats_phase1": {"Dupont_Marie": {...}, ...}
    }
    Retourne un dict vide si le fichier n'existe pas.
    """
    path = Path(chemin)
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            # Fichier corrompu → on repart de zéro mais on sauvegarde l'original
            import shutil
            shutil.copy(chemin, chemin + ".backup")
            return {}
    return {}


def sauvegarder_progression(chemin: str, progression: dict) -> None:
    """
    Écrit la progression dans progress.json avec timestamp.
    Utilise une écriture atomique (fichier temp puis rename) pour
    éviter la corruption en cas d'interruption brutale.
    """
    progression["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    chemin_tmp = chemin + ".tmp"
    with open(chemin_tmp, "w", encoding="utf-8") as f:
        json.dump(progression, f, ensure_ascii=False, indent=2)

    # Rename atomique : si le script est tué pendant l'écriture,
    # l'ancien fichier est intact
    import os
    os.replace(chemin_tmp, chemin)


def mettre_a_jour_statut(progression: dict, cle: str, statut: str) -> None:
    """Met à jour le statut d'une personne dans le dict de progression."""
    if "statuts" not in progression:
        progression["statuts"] = {}
    progression["statuts"][cle] = statut


def cle_personne(nom: str, prenom: str) -> str:
    """Génère une clé unique pour identifier une personne dans progress.json."""
    return f"{nom.strip()}_{prenom.strip()}"


def afficher_barre_progression(actuel: int, total: int, largeur: int = 40, suffixe: str = "") -> None:
    """
    Affiche une barre de progression dans le terminal.
    Exemple : [████████████░░░░░░░░░░░░░░░░░░] 12/1600 — Marie Dupont
    """
    pct    = actuel / total if total else 0
    rempli = int(largeur * pct)
    barre  = "█" * rempli + "░" * (largeur - rempli)
    pct_str = f"{pct * 100:.1f}%"
    print(f"\r  [{barre}] {actuel}/{total} ({pct_str}) {suffixe}", end="", flush=True)
    if actuel == total:
        print()


def nettoyer_url_linkedin(url: str) -> str:
    """Nettoie une URL LinkedIn extraite depuis Google."""
    if not url:
        return ""
    return url.split("?")[0].split("&")[0].split("#")[0].rstrip("/")


# ----------------------------------------------------------
# Rapport final
# ----------------------------------------------------------

def generer_rapport(
    chemin_rapport: str,
    resultats_phase1: list,
    resultats_phase2: list,
    heure_debut: datetime,
) -> None:
    """
    Génère rapport.txt à la fin du traitement ou sur interruption.
    Affiche aussi le rapport dans le terminal.
    """
    heure_fin = datetime.now()
    duree_totale = heure_fin - heure_debut
    heures, reste = divmod(int(duree_totale.total_seconds()), 3600)
    minutes, secondes = divmod(reste, 60)

    # Comptages phase 1
    nb_trouves    = sum(1 for r in resultats_phase1 if r.get("statut") == "trouvé")
    nb_non_verif  = sum(1 for r in resultats_phase1 if r.get("statut") == "non_vérifié")
    nb_non_trouve = sum(1 for r in resultats_phase1 if r.get("statut") == "non_trouvé")
    nb_bloques    = sum(1 for r in resultats_phase1 if r.get("statut") == "bloqué")
    nb_erreurs_p1 = sum(1 for r in resultats_phase1 if r.get("statut") == "erreur")

    # Comptages phase 2
    nb_dom    = sum(1 for r in resultats_phase2 if r.get("source") == "DOM")
    nb_vision = sum(1 for r in resultats_phase2 if r.get("source") == "Claude Vision")
    nb_err_p2 = sum(1 for r in resultats_phase2 if "erreur" in r.get("source", ""))

    lignes_erreurs_p1 = [
        f"  - {r['prenom']} {r['nom']} : {r.get('statut','?')}"
        for r in resultats_phase1
        if r.get("statut") in ("erreur", "bloqué")
    ]
    detail_erreurs = "\n".join(lignes_erreurs_p1) if lignes_erreurs_p1 else "  Aucune erreur."
    sep = "─" * 60
    sep2 = "=" * 60

    rapport = (
        f"\n{sep2}\n"
        f"  RAPPORT FINAL — LinkedIn Scraper\n"
        f"{sep2}\n\n"
        f"  Démarré    : {heure_debut.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"  Terminé    : {heure_fin.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"  Durée      : {heures}h {minutes}m {secondes}s\n\n"
        f"{sep}\n"
        f"  PHASE 1 — Recherche LinkedIn ({len(resultats_phase1)} personnes traitées)\n"
        f"{sep}\n"
        f"  Trouvés (vérifiés)    : {nb_trouves}\n"
        f"  Trouvés (non vérif.)  : {nb_non_verif}\n"
        f"  Non trouvés           : {nb_non_trouve}\n"
        f"  Bloqués Google        : {nb_bloques}\n"
        f"  Erreurs               : {nb_erreurs_p1}\n\n"
        f"{sep}\n"
        f"  PHASE 2 — Extraction postes ({len(resultats_phase2)} profils traités)\n"
        f"{sep}\n"
        f"  Via DOM               : {nb_dom}\n"
        f"  Via Claude Vision     : {nb_vision}\n"
        f"  Erreurs               : {nb_err_p2}\n\n"
        f"{sep}\n"
        f"  ERREURS PHASE 1 (détail)\n"
        f"{sep}\n"
        f"{detail_erreurs}\n\n"
        f"  Voir errors.log pour le détail complet.\n"
        f"{sep2}\n"
    )

    # Écriture dans le fichier
    with open(chemin_rapport, "w", encoding="utf-8") as f:
        f.write(rapport)

    # Affichage terminal
    print(rapport)
