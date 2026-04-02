#!/usr/bin/env python3
"""
GEMA Automation — Extraction LinkedIn
======================================
Script unique : lance le scraping ou relance les erreurs.

  python main.py          Traitement normal
  python main.py --retry  Relance les profils en erreur
"""

import os
import sys
import time
import random
import subprocess
import traceback
from datetime import datetime

import config
from lib.excel   import load_gema, est_lien_valide, OutputWriter
from lib.browser import Browser
from lib.logger  import get_logger
from lib.helpers import (
    charger_progression,
    sauvegarder_progression,
    afficher_barre,
    nettoyer_erreurs,
)

log = get_logger("MAIN")

_heure_debut = datetime.now()
_compteurs   = {
    "ok": 0, "deja_traite": 0, "sans_lien": 0,
    "profil_non_trouve": 0, "acces_refuse": 0,
    "profil_sans_exp": 0, "erreur": 0,
}


# ----------------------------------------------------------
# Utilitaires
# ----------------------------------------------------------

def _empecher_veille():
    try:
        if sys.platform == "darwin":
            subprocess.Popen(
                ["caffeinate", "-i", "-w", str(os.getpid())],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        elif sys.platform == "win32":
            import ctypes
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000003)
    except Exception:
        pass


def _logger_erreur(contexte, e):
    with open("errors.log", "a", encoding="utf-8") as f:
        f.write(f"\n[{datetime.now():%Y-%m-%d %H:%M:%S}] {contexte}\n")
        f.write(traceback.format_exc() + "\n")
    log.error(f"Erreur loguée : {contexte} — {e}")


def _cle(p):
    if p.get("id"):
        return p["id"]
    return f"{p['nom']}_{p['prenom']}_{p.get('annee', '')}_{p.get('ecole', '')}"


def _pause(compteur):
    if compteur > 0 and compteur % 500 == 0:
        log.info(f"Pause 20 min ({compteur} profils)…")
        time.sleep(20 * 60)
    elif compteur > 0 and compteur % 200 == 0:
        log.info(f"Pause 10 min ({compteur} profils)…")
        time.sleep(10 * 60)
    elif compteur > 0 and compteur % 50 == 0:
        log.info(f"Pause 3 min ({compteur} profils)…")
        time.sleep(3 * 60)


# ----------------------------------------------------------
# Résumé avant lancement
# ----------------------------------------------------------

def _afficher_resume(personnes, progression):
    traites_set = set(progression.get("traites", []))

    nb_deja = nb_avec = nb_sans = nb_reprendre = 0
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

    t_opt  = int(nb_avec * 10 * 1.15)
    t_real = int(t_opt * 1.5)

    def fmt(s):
        h, r = divmod(s, 3600)
        m, _ = divmod(r, 60)
        return f"{h}h {m:02d}min" if h else f"{m}min"

    sep = "=" * 50
    print(f"\n{sep}")
    print("  RESUME AVANT LANCEMENT")
    print(sep)
    if nb_reprendre:
        print(f"  Déjà dans progress.json   : {nb_reprendre:>5}  (skippés)")
    print(f"  Déjà traités (gema)       : {nb_deja:>5}  (skippés)")
    print(f"  Avec lien LinkedIn        : {nb_avec:>5}  (à traiter)")
    print(f"  Sans lien LinkedIn        : {nb_sans:>5}  (ignorés)")
    print(f"  {'─'*50}")
    print(f"  Estimation : {fmt(t_opt)} — {fmt(t_real)}")
    print(sep)

    try:
        return input("  Lancer ? (o/n) : ").strip().lower() == "o"
    except (KeyboardInterrupt, EOFError):
        return False


class _SkipPerson(Exception):
    pass


# ----------------------------------------------------------
# Boucle principale
# ----------------------------------------------------------

def _traiter(personnes, progression, browser, output):
    traites_set = set(progression.get("traites", []))
    compteur = 0
    acces_refuse_consecutifs = 0

    for i, p in enumerate(personnes, start=1):
        cle   = _cle(p)
        label = f"{p['prenom']} {p['nom']}"
        afficher_barre(i, len(personnes), suffixe=label)

        if cle in traites_set:
            continue

        log.info(f"[{i}/{len(personnes)}] {label}")

        data = {
            "id": p["id"], "prenom": p["prenom"], "nom": p["nom"],
            "annee": p["annee"], "ecole": p["ecole"],
            "lien_retenu": "", "lien_source": p["lien"],
            "poste1": "", "societe1": "", "periode1": "",
            "poste2": "", "societe2": "", "periode2": "",
            "statut": "",
        }

        try:
            if p["deja_traite"]:
                data["poste1"]      = p["poste1_gema"]
                data["poste2"]      = p["poste2_gema"]
                data["lien_retenu"] = p["lien"]
                data["statut"]      = "deja_traite"
                _compteurs["deja_traite"] += 1

            elif not est_lien_valide(p["lien"]):
                data["statut"] = "sans_lien"
                _compteurs["sans_lien"] += 1

            else:
                lien_source = p["lien"]

                # Résolution search → /in/
                if "/search/" in lien_source:
                    lien_profil = browser.extract_profile_from_search(lien_source)
                    if lien_profil is None:
                        current = browser.driver.current_url if browser.driver else ""
                        if "authwall" in current or "/login" in current:
                            data["statut"] = "acces_refuse"
                            _compteurs["acces_refuse"] += 1
                        else:
                            data["statut"] = "profil_non_trouve"
                            _compteurs["profil_non_trouve"] += 1
                        raise _SkipPerson()
                else:
                    lien_profil = lien_source

                data["lien_retenu"] = lien_profil

                # Ouvrir et extraire
                browser.ouvrir(lien_profil)

                if not browser.est_profil_valide():
                    data["statut"] = "acces_refuse"
                    _compteurs["acces_refuse"] += 1
                else:
                    postes = browser.get_experience_dom()

                    if postes:
                        p1 = postes[0] if len(postes) > 0 else {}
                        p2 = postes[1] if len(postes) > 1 else {}
                        data["poste1"]   = p1.get("titre", "")
                        data["societe1"] = p1.get("societe", "")
                        data["periode1"] = p1.get("periode", "")
                        data["poste2"]   = p2.get("titre", "")
                        data["societe2"] = p2.get("societe", "")
                        data["periode2"] = p2.get("periode", "")
                        data["statut"]   = "ok"
                        _compteurs["ok"] += 1
                        log.info(f"  → ok : {data['poste1']} @ {data['societe1']}")
                    else:
                        data["statut"] = "profil_sans_exp"
                        _compteurs["profil_sans_exp"] += 1
                        log.warning(f"  → profil_sans_exp : {lien_profil}")

        except _SkipPerson:
            pass
        except Exception as e:
            _logger_erreur(f"Traitement {label}", e)
            data["statut"] = "erreur"
            _compteurs["erreur"] += 1

        # Détection déconnexion
        if data["statut"] == "acces_refuse":
            acces_refuse_consecutifs += 1
            if acces_refuse_consecutifs >= 3:
                log.error("3 accès refusés consécutifs — reconnectez-vous dans Firefox.")
                try:
                    input("  Appuyez sur Entrée pour reprendre… ")
                    acces_refuse_consecutifs = 0
                except (KeyboardInterrupt, EOFError):
                    raise KeyboardInterrupt
        else:
            acces_refuse_consecutifs = 0

        # Sauvegarde
        try:
            output.save_row(data)
        except Exception as e:
            _logger_erreur(f"save_row {label}", e)

        traites_set.add(cle)
        progression["traites"] = list(traites_set)
        sauvegarder_progression(config.FICHIER_PROGRESSION, progression)

        compteur += 1
        _pause(compteur)

        if data["statut"] not in ("deja_traite", "sans_lien"):
            time.sleep(random.uniform(config.DELAI_MIN, config.DELAI_MAX))


# ----------------------------------------------------------
# Rapport
# ----------------------------------------------------------

def _generer_rapport():
    heure_fin = datetime.now()
    duree     = heure_fin - _heure_debut
    h, reste  = divmod(int(duree.total_seconds()), 3600)
    m, s      = divmod(reste, 60)
    total     = sum(_compteurs.values())

    sep = "=" * 60
    rapport = (
        f"\n{sep}\n"
        f"  RAPPORT — LinkedIn Scraper\n{sep}\n\n"
        f"  Début  : {_heure_debut:%Y-%m-%d %H:%M:%S}\n"
        f"  Fin    : {heure_fin:%Y-%m-%d %H:%M:%S}\n"
        f"  Durée  : {h}h {m}m {s}s\n\n"
        f"  RÉSULTATS ({total} lignes)\n"
        f"  {'─'*40}\n"
        f"  ok                : {_compteurs['ok']}\n"
        f"  deja_traite       : {_compteurs['deja_traite']}\n"
        f"  sans_lien         : {_compteurs['sans_lien']}\n"
        f"  profil_non_trouve : {_compteurs['profil_non_trouve']}\n"
        f"  acces_refuse      : {_compteurs['acces_refuse']}\n"
        f"  profil_sans_exp   : {_compteurs['profil_sans_exp']}\n"
        f"  erreur            : {_compteurs['erreur']}\n\n"
        f"  Sortie : {config.FICHIER_OUTPUT_FINAL}\n{sep}\n"
    )

    with open("rapport.txt", "w", encoding="utf-8") as f:
        f.write(rapport)
    print(rapport)


# ----------------------------------------------------------
# Point d'entrée
# ----------------------------------------------------------

def main():
    global _heure_debut

    # Mode --retry
    if "--retry" in sys.argv:
        log.info("Mode retry : nettoyage des erreurs…")
        nb = nettoyer_erreurs(config.FICHIER_OUTPUT_FINAL, config.FICHIER_PROGRESSION)
        if nb == 0:
            log.info("Aucune erreur à relancer.")
            return
        log.info(f"{nb} ligne(s) supprimée(s). Relancement…")
        print()

    _heure_debut = datetime.now()
    log.info("=" * 50)
    log.info(f"  LinkedIn Scraper — {_heure_debut:%Y-%m-%d %H:%M:%S}")
    log.info("=" * 50)

    _empecher_veille()

    personnes = load_gema(config.FICHIER_GEMA)
    if not personnes:
        log.error(f"Aucune personne dans {config.FICHIER_GEMA}.")
        sys.exit(1)

    progression = charger_progression(config.FICHIER_PROGRESSION)
    progression.setdefault("traites", [])

    if not _afficher_resume(personnes, progression):
        log.info("Annulé.")
        sys.exit(0)

    browser = Browser(firefox_profile_path=config.FIREFOX_PROFILE_PATH)
    output  = OutputWriter(config.FICHIER_OUTPUT_FINAL)
    browser.demarrer()
    output.ouvrir()

    try:
        _traiter(personnes, progression, browser, output)
    except KeyboardInterrupt:
        log.warning("Interruption (Ctrl+C). Progression sauvegardée.")
    except Exception as e:
        _logger_erreur("Erreur fatale", e)
    finally:
        output.fermer()
        browser.fermer()
        try:
            _generer_rapport()
        except Exception:
            pass

    log.info("Terminé.")
    log.info(f"Résultats : {config.FICHIER_OUTPUT_FINAL}")


if __name__ == "__main__":
    main()
