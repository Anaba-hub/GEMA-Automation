# ============================================================
# agents/agent_browser.py — Contrôle Firefox via Selenium
# ============================================================
# Responsabilités :
#   - Démarrer Firefox avec le profil connecté à LinkedIn (visible, jamais headless)
#   - Résoudre un lien /search/ en lien /in/ : extract_profile_from_search(url)
#   - Ouvrir un profil LinkedIn : ouvrir(url)
#   - Vérifier que la page est bien un profil : est_profil_valide()
#   - Extraire les postes via le DOM : get_experience_dom()
#   - Prendre un screenshot si le DOM échoue : take_screenshot()
#
# Structure DOM LinkedIn — deux cas à gérer :
#   CAS A (poste simple) : une seule expérience chez une entreprise
#     li > .t-bold (titre) + .t-normal (société·type) + .t-black--light (période)
#   CAS B (postes groupés) : plusieurs rôles chez la même entreprise
#     li > .t-bold (NOM ENTREPRISE) > ul > li > .t-bold (titre rôle)
# ============================================================

import glob
import os
import re
import time
from io import BytesIO
from pathlib import Path
from typing import List, Optional
from urllib.parse import urlparse, urlunparse

from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from utils.logger import get_logger

log = get_logger("BROWSER")


class AgentBrowser:
    """
    Agent de contrôle du navigateur Firefox.
    Firefox est toujours lancé en mode visible (jamais headless).
    """

    def __init__(self, firefox_profile_path: Optional[str] = None):
        """
        Args:
            firefox_profile_path : chemin vers le profil Firefox déjà connecté à LinkedIn.
                                   Si None, tente une auto-détection macOS.
        """
        self.firefox_profile_path = firefox_profile_path
        self.driver               = None

    # ----------------------------------------------------------
    # Démarrage / Arrêt
    # ----------------------------------------------------------

    def demarrer(self) -> webdriver.Firefox:
        """
        Démarre Firefox en mode visible et retourne le driver Selenium.
        Charge le profil utilisateur existant (cookies LinkedIn inclus).
        """
        options = Options()
        # Firefox toujours visible — jamais headless
        profil = self.firefox_profile_path or self._detecter_profil_macos()

        if profil:
            # -profile charge le profil existant (session LinkedIn active)
            # -no-remote permet une 2e instance si Firefox est déjà ouvert
            options.add_argument("-profile")
            options.add_argument(profil)
            options.add_argument("-no-remote")
            log.info(f"Profil Firefox chargé : {profil}")
        else:
            log.warning(
                "Aucun profil Firefox trouvé. "
                "Vous devrez vous connecter manuellement à LinkedIn."
            )

        try:
            self.driver = webdriver.Firefox(options=options)
            log.info("Firefox démarré.")
            return self.driver
        except Exception as e:
            log.error(f"Impossible de démarrer Firefox : {e}")
            log.error("Vérifiez que geckodriver est installé (brew install geckodriver).")
            raise

    def fermer(self) -> None:
        """Ferme Firefox proprement."""
        if self.driver:
            self.driver.quit()
            self.driver = None
            log.info("Firefox fermé.")

    def _detecter_profil_macos(self) -> Optional[str]:
        """
        Détecte automatiquement le profil Firefox principal sur macOS.
        Priorité : *.default-release > *.default

        Si la détection échoue, ouvrir Firefox → about:profiles
        pour copier le chemin du profil actif et le coller dans
        FIREFOX_PROFILE_PATH dans config.py.
        """
        base = os.path.expanduser(
            "~/Library/Application Support/Firefox/Profiles/"
        )
        for pattern in ["*.default-release", "*.default"]:
            correspondances = glob.glob(os.path.join(base, pattern))
            if correspondances:
                profil = correspondances[0]
                log.info(f"Profil Firefox auto-détecté : {profil}")
                return profil
        return None

    # ----------------------------------------------------------
    # Navigation
    # ----------------------------------------------------------

    def ouvrir(self, url: str) -> None:
        """
        Navigue vers l'URL dans Firefox.
        Attend 3 secondes pour laisser le JS LinkedIn se charger.
        """
        if not self.driver:
            raise RuntimeError("Le driver n'est pas démarré. Appelez demarrer() d'abord.")
        self.driver.get(url)
        time.sleep(3)
        log.info(f"Page ouverte : {url}")

    def extract_profile_from_search(self, url_recherche: str) -> Optional[str]:
        """
        Ouvre une page de recherche LinkedIn (/search/results/…),
        lit le DOM pour trouver le premier résultat de personne,
        et retourne l'URL propre du profil (/in/slug) ou None.

        Règles :
          - Attend 3 s après chargement pour laisser le JS s'exécuter
          - Prend uniquement le 1er lien href contenant "/in/"
          - Nettoie l'URL : supprime tous les paramètres (?miniProfileUrn=…)
          - Retourne None si :
              • redirection vers /login ou authwall
              • aucun lien /in/ trouvé dans le DOM
          - Zéro appel Claude Vision ici (DOM uniquement)
        """
        if not self.driver:
            raise RuntimeError("Le driver n'est pas démarré. Appelez demarrer() d'abord.")

        self.driver.get(url_recherche)
        time.sleep(3)  # laisser le JS LinkedIn charger les résultats

        current = self.driver.current_url
        if "authwall" in current or "/login" in current:
            log.warning(f"Accès refusé sur la page de recherche : {current}")
            return None

        # Chercher tous les liens href contenant "/in/"
        try:
            liens = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='/in/']")
        except Exception as e:
            log.error(f"Erreur lors de la recherche des liens /in/ : {e}")
            return None

        for lien in liens:
            href = lien.get_attribute("href") or ""
            if "/in/" not in href:
                continue
            # Garder uniquement le chemin linkedin.com/in/slug (sans paramètres)
            url_propre = self._nettoyer_url_profil(href)
            if url_propre:
                log.info(f"Profil trouvé via recherche : {url_propre}")
                return url_propre

        log.warning(f"Aucun profil /in/ trouvé sur : {url_recherche}")
        return None

    def _nettoyer_url_profil(self, href: str) -> Optional[str]:
        """
        Extrait et retourne uniquement la partie https://www.linkedin.com/in/slug
        d'un href LinkedIn, sans les paramètres d'URL.

        Exemples :
          https://www.linkedin.com/in/marie-dupont-123?miniProfileUrn=…
            → https://www.linkedin.com/in/marie-dupont-123
          /in/marie-dupont-123/
            → https://www.linkedin.com/in/marie-dupont-123
        """
        try:
            parsed = urlparse(href)
            # Reconstituer sans les paramètres ni le fragment
            chemin = parsed.path.rstrip("/")
            if not chemin.startswith("/in/"):
                return None
            propre = urlunparse((
                "https",
                "www.linkedin.com",
                chemin,
                "", "", "",  # params, query, fragment → vides
            ))
            return propre
        except Exception:
            return None

    def est_profil_valide(self) -> bool:
        """
        Vérifie que la page courante est bien un profil LinkedIn.
        Retourne False si :
          - l'URL contient "authwall" ou "/login" (accès refusé)
          - l'URL contient "404" ou "page-not-found"
          - l'URL ne contient pas "linkedin.com/in/"
        Retourne True sinon.
        """
        if not self.driver:
            return False
        current = self.driver.current_url
        if "authwall" in current or "/login" in current:
            log.warning(f"Accès refusé (redirection auth) : {current}")
            return False
        if "404" in current or "not-found" in current or "page-not-found" in current:
            log.warning(f"Page non trouvée : {current}")
            return False
        if "linkedin.com/in/" not in current:
            log.warning(f"URL inattendue (pas un profil /in/) : {current}")
            return False
        return True

    # ----------------------------------------------------------
    # Extraction DOM — méthode principale
    # ----------------------------------------------------------

    def get_experience_dom(self) -> List[dict]:
        """
        Extrait les 2 dernières expériences depuis le DOM LinkedIn.

        Gère les deux structures LinkedIn :
          CAS A — Poste simple :
            li
              └─ .t-bold span          → titre du poste
              └─ .t-normal span        → société · type contrat
              └─ .t-black--light span  → période

          CAS B — Postes groupés sous une même entreprise :
            li
              └─ .t-bold span          → NOM DE L'ENTREPRISE (pas le poste !)
              └─ ul > li
                   └─ .t-bold span     → titre du rôle
                   └─ .t-black--light  → période

        Retourne :
            Liste de 0, 1 ou 2 dicts avec les clés : titre, societe, periode.
            Liste vide → recours au screenshot + Claude Vision.
        """
        self._scroll_vers_experience()
        postes = []

        blocs = self._trouver_blocs_experience()
        if not blocs:
            log.warning("Aucun bloc d'expérience trouvé dans le DOM.")
            return []

        for bloc in blocs:
            if self._est_bloc_groupe(bloc):
                # CAS B : extraire les sous-postes (ul > li)
                sous_postes = self._extraire_postes_groupes(bloc)
                postes.extend(sous_postes)
            else:
                # CAS A : extraire le poste simple
                poste = self._extraire_poste_simple(bloc)
                if poste["titre"] or poste["societe"]:
                    postes.append(poste)

            if len(postes) >= 2:
                break

        postes = postes[:2]
        log.info(
            f"DOM : {len(postes)} poste(s) extrait(s) — "
            + " | ".join(f"{p['titre']} @ {p['societe']}" for p in postes)
        )
        return postes

    # ----------------------------------------------------------
    # Helpers DOM internes
    # ----------------------------------------------------------

    def _trouver_blocs_experience(self):
        """Retourne la liste des éléments li de la section Expérience."""
        selecteurs = [
            # LinkedIn 2024 (nouvelle interface)
            "div#experience ~ div .pvs-list__item--line-separated",
            # Interface intermédiaire
            "section[data-section='experience'] li.artdeco-list__item",
            # Ancienne interface
            "#experience ~ div ul > li",
            "section#experience li",
        ]
        for sel in selecteurs:
            blocs = self.driver.find_elements(By.CSS_SELECTOR, sel)
            if blocs:
                log.info(f"Sélecteur DOM : '{sel}' ({len(blocs)} bloc(s))")
                return blocs
        return []

    def _est_bloc_groupe(self, bloc) -> bool:
        """
        Retourne True si le bloc contient des sous-listes
        (CAS B : plusieurs rôles chez la même entreprise).
        """
        try:
            sous_listes = bloc.find_elements(By.CSS_SELECTOR, "ul li")
            return len(sous_listes) > 0
        except Exception:
            return False

    def _extraire_poste_simple(self, bloc) -> dict:
        """CAS A : extrait titre, société, période depuis un bloc de poste unique."""
        titre = self._premier_texte(bloc, [
            ".t-bold span[aria-hidden='true']",
            "div.t-bold span",
            "h3 span",
        ])
        societe_brut = self._premier_texte(bloc, [
            "span.t-14.t-normal span[aria-hidden='true']",
            "span.t-14.t-normal:not(.t-black--light) span",
            "p.t-14 span",
        ])
        # Supprimer "· Temps plein", "· CDI", etc.
        societe = societe_brut.split("·")[0].strip() if societe_brut else ""

        periode = self._premier_texte(bloc, [
            "span.t-14.t-normal.t-black--light span[aria-hidden='true']",
            "span.pvs-entity__caption-wrapper",
            "span.t-black--light span",
        ])
        # Garder seulement la partie dates (avant " · 2 ans 3 mois")
        periode = periode.split("·")[0].strip() if periode else ""

        return {"titre": titre, "societe": societe, "periode": periode}

    def _extraire_postes_groupes(self, bloc) -> List[dict]:
        """
        CAS B : le bloc regroupe plusieurs rôles sous une même entreprise.
        Retourne jusqu'à 2 postes.
        """
        postes = []

        # Nom de l'entreprise (dans le header du bloc groupé)
        societe_brut = self._premier_texte(bloc, [
            ".t-bold span[aria-hidden='true']",
            "div.t-bold span",
        ])
        societe = societe_brut.split("·")[0].strip() if societe_brut else ""

        # Sous-postes (rôles individuels)
        sous_items = bloc.find_elements(By.CSS_SELECTOR, "ul li")
        for sous in sous_items[:2]:
            titre = self._premier_texte(sous, [
                ".t-bold span[aria-hidden='true']",
                "div.t-bold span",
                "span[aria-hidden='true']",
            ])
            periode_brut = self._premier_texte(sous, [
                "span.t-14.t-normal.t-black--light span[aria-hidden='true']",
                "span.t-black--light span",
                "span.pvs-entity__caption-wrapper",
            ])
            periode = periode_brut.split("·")[0].strip() if periode_brut else ""

            if titre:
                postes.append({"titre": titre, "societe": societe, "periode": periode})

        return postes

    def _premier_texte(self, element, selecteurs: list) -> str:
        """
        Essaie chaque sélecteur CSS dans l'ordre et retourne le texte
        du premier élément trouvé. Retourne "" si aucun ne correspond.
        """
        for sel in selecteurs:
            try:
                valeur = element.find_element(By.CSS_SELECTOR, sel).text.strip()
                if valeur:
                    return valeur
            except NoSuchElementException:
                continue
        return ""

    def _scroll_vers_experience(self) -> None:
        """Fait défiler la page vers la section Expérience pour forcer son chargement."""
        try:
            wait = WebDriverWait(self.driver, 10)
            section = wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR,
                 "div#experience, section[data-section='experience'], section#experience")
            ))
            self.driver.execute_script("arguments[0].scrollIntoView();", section)
            time.sleep(2)
        except TimeoutException:
            log.warning("Section Expérience non trouvée lors du scroll — on tente quand même.")

    # ----------------------------------------------------------
    # Screenshot (fallback si DOM vide)
    # ----------------------------------------------------------

    def take_screenshot(self, max_largeur: int = 1280) -> Optional[bytes]:
        """
        Prend un screenshot de la page courante.
        Retourne les octets JPEG compressés, ou None en cas d'erreur.
        """
        if not self.driver:
            log.error("Impossible de prendre un screenshot : driver non démarré.")
            return None

        try:
            self._scroll_vers_experience()
            time.sleep(1)

            png_bytes = self.driver.get_screenshot_as_png()
            img = Image.open(BytesIO(png_bytes))

            if img.width > max_largeur:
                ratio           = max_largeur / img.width
                nouvelle_taille = (max_largeur, int(img.height * ratio))
                img = img.resize(nouvelle_taille, Image.LANCZOS)

            buffer = BytesIO()
            img.convert("RGB").save(buffer, format="JPEG", quality=85)
            jpeg_bytes = buffer.getvalue()

            log.info(
                f"Screenshot capturé : {len(jpeg_bytes) // 1024} Ko "
                f"({img.width}×{img.height}px)"
            )
            return jpeg_bytes

        except Exception as e:
            log.error(f"Erreur lors du screenshot : {e}")
            return None
