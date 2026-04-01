# ============================================================
# main.py — Orchestrateur LinkedIn Scraper
# ============================================================
# Source  : gema.xlsx (fichier unique, NE PAS ÉCRASER)
# Sortie  : output_final.xlsx (ligne par ligne)
#
# Flux par personne :
#   1. Poste 1 + Poste 2 déjà remplis dans gema → "déjà_traité" → skip
#   2. Lien LinkedIn manquant/invalide           → "sans_lien"   → skip
#   3. Résolution du lien :
#      └─ Lien /search/results → extract_profile_from_search()
#           - Aucun résultat    → "profil_non_trouvé"
#           - Accès refusé      → "accès_refusé"
#      └─ Lien /in/ (direct)  → utiliser directement
#   4. Ouvrir le profil /in/ dans Firefox
#      └─ Accès refusé (/login)                 → "accès_refusé"
#      └─ Extraction DOM                         → "ok"
#      └─ DOM vide → screenshot → Claude Vision → "ok" ou "profil_sans_exp"
#
# Colonnes output_final.xlsx :
#   Lien retenu = URL /in/ résolue
#   Lien source = URL d'origine dans gema (/search/ ou /in/)
#
# Protections nuit :
#   - caffeinate    : empêche la veille Mac
#   - progress.json : reprise après interruption
#   - Pauses auto   : 3 min / 50, 10 min / 200, 20 min / 500 personnes
#   - Délai         : 4–8 s entre profils
#   - errors.log    : toutes les erreurs non fatales
#   - rapport.txt   : généré à la fin ET sur Ctrl+C
# ============================================================

import os
import sys
import time
import random
import subprocess
import traceback
from datetime import datetime
from pathlib import Path

import config
from agents.agent_excel   import load_gema, est_lien_valide, save_row
from agents.agent_browser import AgentBrowser
from agents.agent_vision  import AgentVision
from utils.logger         import get_logger
from utils.helpers        import (
    charger_progression,
    sauvegarder_progression,
    afficher_barre_progression,
)

log = get_logger("ORCHESTRATEUR")

_heure_debut = datetime.now()
_compteurs   = {"ok": 0, "deja_traite": 0, "sans_lien": 0,
                "profil_non_trouve": 0, "acces_refuse": 0,
                "profil_sans_exp": 0, "erreur": 0}


# ----------------------------------------------------------
# Utilitaires
# ----------------------------------------------------------

def _activer_caffeinate() -> None:
    """Lance caffeinate pour empêcher la veille Mac pendant le run."""
    try:
        subprocess.Popen(
            ["caffeinate", "-i", "-w", str(os.getpid())],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        log.info("caffeinate activé — le Mac ne se mettra pas en veille.")
    except FileNotFoundError:
        log.warning("caffeinate introuvable (macOS seulement).")
    except Exception as e:
        log.warning(f"caffeinate non disponible : {e}")


def _logger_erreur(contexte: str, e: Exception) -> None:
    """Enregistre une erreur non fatale dans errors.log sans tuer le script."""
    with open("errors.log", "a", encoding="utf-8") as f:
        f.write(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {contexte}\n")
        f.write(traceback.format_exc())
        f.write("\n")
    log.error(f"Erreur loggée → errors.log : {contexte} — {e}")


def _cle(personne: dict) -> str:
    """Clé unique pour progress.json : id si disponible, sinon nom_prenom."""
    return personne["id"] if personne.get("id") else f"{personne['nom']}_{personne['prenom']}"


def _pause(compteur: int) -> None:
    """Pause automatique à des jalons critiques."""
    if compteur > 0 and compteur % 500 == 0:
        log.info(f"PAUSE 20 min ({compteur} personnes traitées)…")
        time.sleep(20 * 60)
    elif compteur > 0 and compteur % 200 == 0:
        log.info(f"Pause 10 min ({compteur} personnes traitées)…")
        time.sleep(10 * 60)
    elif compteur > 0 and compteur % 50 == 0:
        log.info(f"Pause 3 min ({compteur} personnes traitées)…")
        time.sleep(3 * 60)


# ----------------------------------------------------------
# Résumé avant lancement
# ----------------------------------------------------------

def _afficher_resume(personnes: list, progression: dict) -> bool:
    """
    Affiche le résumé des catégories + estimation de temps.
    Retourne True si l'utilisateur confirme le lancement.
    """
    traites_set = set(progression.get("traites", []))

    nb_deja     = 0  # Poste 1 + 2 déjà remplis dans gema (skip immédiat)
    nb_avec     = 0  # Lien valide → à traiter
    nb_sans     = 0  # Pas de lien → ignoré
    nb_reprendre = 0  # Déjà traité dans progress.json (skip silencieux)

    for p in personnes:
        cle = _cle(p)
        if cle in traites_set:
            nb_reprendre += 1
        elif p["deja_traite"]:
            nb_deja += 1
        elif est_lien_valide(p["lien"]):
            nb_avec += 1
        else:
            nb_sans += 1

    # Estimation : ~10s par profil avec lien + 15% pauses → ×1.5 réaliste
    t_opt  = int(nb_avec * 10 * 1.15)
    t_real = int(t_opt * 1.5)

    def fmt(s):
        h, r = divmod(s, 3600)
        m, _ = divmod(r, 60)
        return f"{h}h {m:02d}min" if h else f"{m}min"

    sep  = "=" * 50
    sep2 = "-" * 50
    print(f"\n{sep}")
    print(f"  RESUME AVANT LANCEMENT")
    print(sep)
    if nb_reprendre:
        print(f"  Deja dans progress.json          : {nb_reprendre:>5}  -> skippés silencieusement")
    print(f"  Déjà traités (Poste 1+2 connus)  : {nb_deja:>5}  -> skippés")
    print(f"  Avec lien LinkedIn               : {nb_avec:>5}  -> à traiter")
    print(f"  Sans lien LinkedIn               : {nb_sans:>5}  -> ignorés")
    print(f"  {sep2}")
    print(f"  Total à traiter                  : {nb_avec:>5} profils")
    print(f"  Optimiste                        : {fmt(t_opt)}")
    print(f"  Réaliste                         : {fmt(t_real)}")
    print(sep)

    try:
        reponse = input("  Lancer ? (o/n) : ").strip().lower()
        return reponse == "o"
    except (KeyboardInterrupt, EOFError):
        return False


# Sentinelle interne pour sortir proprement du bloc try sans crash
class _SkipPerson(Exception):
    pass


# ----------------------------------------------------------
# Boucle principale
# ----------------------------------------------------------

def _traiter(personnes: list, progression: dict, browser: AgentBrowser,
             vision: AgentVision) -> None:
    """
    Boucle unique sur toutes les personnes.
    Sauvegarde output_final.xlsx + progress.json après chaque ligne.
    """
    traites_set = set(progression.get("traites", []))
    compteur    = 0  # pour les pauses

    for i, p in enumerate(personnes, start=1):
        cle    = _cle(p)
        label  = f"{p['prenom']} {p['nom']}"
        afficher_barre_progression(i, len(personnes), suffixe=label)

        # --- Reprise : déjà traité dans cette session ou une précédente ---
        if cle in traites_set:
            continue

        log.info(f"[{i}/{len(personnes)}] {label}")

        # Résultat à construire et sauvegarder
        data = {
            "id":           p["id"],
            "prenom":       p["prenom"],
            "nom":          p["nom"],
            "annee":        p["annee"],
            "ecole":        p["ecole"],
            "lien_retenu":  "",   # URL /in/ résolue (remplie plus bas)
            "lien_source":  p["lien"],  # URL d'origine dans gema
            "poste1":  "", "societe1": "", "periode1": "",
            "poste2":  "", "societe2": "", "periode2": "",
            "statut":  "",
        }

        try:
            # CAS 1 — Poste 1 + Poste 2 déjà remplis dans gema → skip extraction
            if p["deja_traite"]:
                data["poste1"]       = p["poste1_gema"]
                data["poste2"]       = p["poste2_gema"]
                data["lien_retenu"]  = p["lien"]
                data["statut"]       = "déjà_traité"
                _compteurs["deja_traite"] += 1
                log.info(f"  → déjà_traité (gema)")

            # CAS 2 — Pas de lien LinkedIn valide → ignorer
            elif not est_lien_valide(p["lien"]):
                data["statut"] = "sans_lien"
                _compteurs["sans_lien"] += 1
                log.info(f"  → sans_lien")

            # CAS 3 — Lien disponible → résoudre si /search/, puis extraire
            else:
                lien_source = p["lien"]

                # Étape 3a — Résolution du lien si c'est une page de recherche
                if "/search/results" in lien_source or "/search/" in lien_source:
                    log.info(f"  Lien /search/ → résolution en cours…")
                    lien_profil = browser.extract_profile_from_search(lien_source)

                    if lien_profil is None:
                        # Vérifier si c'est un accès refusé ou vraiment aucun résultat
                        current = browser.driver.current_url if browser.driver else ""
                        if "authwall" in current or "/login" in current:
                            data["statut"] = "accès_refusé"
                            _compteurs["acces_refuse"] += 1
                            log.warning(f"  → accès_refusé (recherche) : {lien_source}")
                        else:
                            data["statut"] = "profil_non_trouvé"
                            _compteurs["profil_non_trouve"] += 1
                            log.warning(f"  → profil_non_trouvé : {lien_source}")
                        # On passe à la personne suivante (pas d'extraction possible)
                        raise _SkipPerson()
                    else:
                        log.info(f"  /search/ → /in/ résolu : {lien_profil}")
                else:
                    # Lien déjà direct /in/
                    lien_profil = lien_source

                data["lien_retenu"] = lien_profil

                # Étape 3b — Ouvrir le profil /in/ et extraire les postes
                browser.ouvrir(lien_profil)

                if not browser.est_profil_valide():
                    data["statut"] = "accès_refusé"
                    _compteurs["acces_refuse"] += 1
                    log.warning(f"  → accès_refusé : {lien_profil}")
                else:
                    postes = browser.get_experience_dom()

                    # Fallback Claude Vision si DOM vide
                    if not postes:
                        log.info("  DOM vide → fallback Claude Vision")
                        screenshot = browser.take_screenshot()
                        if screenshot:
                            postes = vision.analyser_et_extraire(screenshot)

                    if postes:
                        p1 = postes[0] if len(postes) > 0 else {}
                        p2 = postes[1] if len(postes) > 1 else {}
                        data["poste1"]    = p1.get("titre", "")
                        data["societe1"]  = p1.get("societe", "")
                        data["periode1"]  = p1.get("periode", "")
                        data["poste2"]    = p2.get("titre", "")
                        data["societe2"]  = p2.get("societe", "")
                        data["periode2"]  = p2.get("periode", "")
                        data["statut"]    = "ok"
                        _compteurs["ok"] += 1
                        log.info(f"  → ok : {data['poste1']} @ {data['societe1']}")
                    else:
                        data["statut"] = "profil_sans_exp"
                        _compteurs["profil_sans_exp"] += 1
                        log.warning(f"  → profil_sans_exp : {lien_profil}")

        except _SkipPerson:
            pass  # statut déjà positionné avant le raise

        except Exception as e:
            _logger_erreur(f"Traitement {label}", e)
            data["statut"] = "erreur"
            _compteurs["erreur"] += 1

        # --- Sauvegarde Excel ---
        try:
            save_row(config.FICHIER_OUTPUT_FINAL, data)
        except Exception as e:
            _logger_erreur(f"save_row {label}", e)

        # --- Mise à jour progress.json ---
        traites_set.add(cle)
        progression["traites"] = list(traites_set)
        sauvegarder_progression(config.FICHIER_PROGRESSION, progression)

        compteur += 1

        # --- Pauses automatiques ---
        _pause(compteur)

        # --- Délai entre profils (seulement si visite Firefox effectuée) ---
        if data["statut"] not in ("déjà_traité", "sans_lien"):
            time.sleep(random.uniform(config.DELAI_MIN, config.DELAI_MAX))


# ----------------------------------------------------------
# Rapport final
# ----------------------------------------------------------

def _generer_rapport() -> None:
    """Génère rapport.txt avec les statistiques de la session."""
    heure_fin   = datetime.now()
    duree       = heure_fin - _heure_debut
    h, reste    = divmod(int(duree.total_seconds()), 3600)
    m, s        = divmod(reste, 60)

    total = sum(_compteurs.values())
    sep   = "=" * 60
    sep2  = "-" * 60

    rapport = (
        f"\n{sep}\n"
        f"  RAPPORT FINAL — LinkedIn Scraper\n"
        f"{sep}\n\n"
        f"  Démarré  : {_heure_debut.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"  Terminé  : {heure_fin.strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"  Durée    : {h}h {m}m {s}s\n\n"
        f"{sep2}\n"
        f"  RÉSULTATS ({total} lignes traitées)\n"
        f"{sep2}\n"
        f"  ok                 : {_compteurs['ok']}\n"
        f"  déjà_traité        : {_compteurs['deja_traite']}\n"
        f"  sans_lien          : {_compteurs['sans_lien']}\n"
        f"  profil_non_trouvé  : {_compteurs['profil_non_trouve']}\n"
        f"  accès_refusé       : {_compteurs['acces_refuse']}\n"
        f"  profil_sans_exp    : {_compteurs['profil_sans_exp']}\n"
        f"  erreur             : {_compteurs['erreur']}\n\n"
        f"  Fichier résultat   : {config.FICHIER_OUTPUT_FINAL}\n"
        f"  Voir errors.log pour le détail des erreurs.\n"
        f"{sep}\n"
    )

    with open("rapport.txt", "w", encoding="utf-8") as f:
        f.write(rapport)
    print(rapport)
    log.info("Rapport généré : rapport.txt")


# ----------------------------------------------------------
# Point d'entrée
# ----------------------------------------------------------

def main():
    global _heure_debut
    _heure_debut = datetime.now()

    log.info("=" * 60)
    log.info("  LinkedIn Scraper — Source : gema.xlsx")
    log.info(f"  Démarrage : {_heure_debut.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("=" * 60)

    # 1. Empêcher la veille Mac
    _activer_caffeinate()

    # 2. Chargement des personnes depuis gema.xlsx
    personnes = load_gema(config.FICHIER_GEMA)
    if not personnes:
        log.error(f"Aucune personne chargée depuis {config.FICHIER_GEMA}. Arrêt.")
        sys.exit(1)
    log.info(f"{len(personnes)} personne(s) chargée(s) depuis {config.FICHIER_GEMA}.")

    # 3. Chargement de la progression existante (reprise)
    progression = charger_progression(config.FICHIER_PROGRESSION)
    progression.setdefault("traites", [])

    # 4. Résumé + confirmation utilisateur
    if not _afficher_resume(personnes, progression):
        log.info("Annulé par l'utilisateur.")
        sys.exit(0)

    # 5. Initialisation des agents
    browser = AgentBrowser(firefox_profile_path=config.FIREFOX_PROFILE_PATH)
    vision  = AgentVision(api_key=config.ANTHROPIC_API_KEY, model=config.CLAUDE_MODEL)
    browser.demarrer()

    # 6. Traitement
    try:
        _traiter(personnes, progression, browser, vision)

    except KeyboardInterrupt:
        log.warning("")
        log.warning("Interruption (Ctrl+C). Progression sauvegardée dans progress.json.")
        log.warning("Relancez 'python main.py' pour reprendre là où vous vous êtes arrêté.")

    except Exception as e:
        _logger_erreur("Erreur fatale dans main()", e)
        log.error("Erreur fatale — voir errors.log.")

    finally:
        browser.fermer()
        try:
            _generer_rapport()
        except Exception as e:
            log.error(f"Impossible de générer le rapport : {e}")

    log.info("")
    log.info("=" * 60)
    log.info("  Terminé.")
    log.info(f"  Résultats : {config.FICHIER_OUTPUT_FINAL}")
    log.info(f"  Rapport   : rapport.txt")
    log.info(f"  Erreurs   : errors.log")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
