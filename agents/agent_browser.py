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
# LinkedIn 2026 : les classes CSS sont obfusquées (hachées).
# L'extraction se fait par structure DOM + JavaScript,
# pas par sélecteurs CSS sémantiques.
# Le lazy-loading nécessite un scroll clavier (Page Down).
# ============================================================

import glob
import os
import re
import subprocess
import time
from io import BytesIO
from pathlib import Path
from typing import List, Optional
from urllib.parse import urlparse, urlunparse

from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from utils.logger import get_logger

log = get_logger("BROWSER")


# JavaScript d'extraction des expériences.
# LinkedIn 2026 : classes CSS obfusquées, lazy-loading.
# On se base sur la structure : <section> contenant "Experience",
# blocs séparés par <hr>, liens /company/ pour les noms d'entreprise,
# <p> pour titre et période, <ul>/<li> pour les postes groupés.
_JS_EXTRACT_EXPERIENCE = """(function() {
try {
    // 1. Trouver le heading "Experience" / "Expérience"
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    let expHeading = null;
    while (walker.nextNode()) {
        const txt = walker.currentNode.textContent.trim();
        if (txt === 'Experience' || txt === 'Expérience') {
            expHeading = walker.currentNode.parentElement;
            break;
        }
    }
    if (!expHeading) return JSON.stringify([]);

    // 2. Remonter à la <section>
    let section = expHeading;
    while (section && section.tagName !== 'SECTION') {
        section = section.parentElement;
    }
    if (!section) return JSON.stringify([]);

    const datePattern = /\\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|janv|f[eé]v|mars|avr|mai|juin|juil|ao[uû]t|sept|oct|nov|d[eé]c|\\d{4}|present|pr[eé]sent)\\b/i;

    // 3. Collecter les liens <a href="/company/..."> qui ont des <p>.
    //    Les séparer en top-level (hors <li>) et sub-roles (dans <li>).
    //    - Top-level sans sub-roles = poste simple
    //    - Top-level avec sub-roles = header de groupe (entreprise)
    const allLinks = section.querySelectorAll('a[href*="/company/"]');
    const topLinks = [];
    const subLinks = [];

    for (const link of allLinks) {
        if (link.querySelectorAll('p').length === 0) continue;
        if (link.closest('li')) {
            subLinks.push(link);
        } else {
            topLinks.push(link);
        }
    }

    // Index sub-roles par href pour lookup rapide
    const subByHref = {};
    for (const sl of subLinks) {
        const h = sl.getAttribute('href');
        if (!subByHref[h]) subByHref[h] = [];
        subByHref[h].push(sl);
    }

    const postes = [];

    for (const link of topLinks) {
        const ps = link.querySelectorAll('p');
        const pTexts = Array.from(ps).map(p => p.textContent.trim());
        const href = link.getAttribute('href');
        const subs = subByHref[href] || [];

        if (subs.length > 0) {
            // GROUPE : pTexts[0] = société, sous-postes dans les sub-links
            const societe = pTexts[0];
            for (const sub of subs) {
                const subPs = sub.querySelectorAll('p');
                if (subPs.length < 1) continue;
                const titre = subPs[0].textContent.trim();
                let periode = '';
                for (const p of subPs) {
                    const t = p.textContent.trim();
                    if (datePattern.test(t)) {
                        periode = t.split('\\u00b7')[0].trim();
                        break;
                    }
                }
                postes.push({titre, societe, periode});
                if (postes.length >= 2) break;
            }
        } else {
            // POSTE SIMPLE : pTexts[0] = titre, pTexts[1] = société
            if (pTexts.length >= 2) {
                const titre = pTexts[0];
                const societe = pTexts[1];
                let periode = '';
                for (const t of pTexts) {
                    if (datePattern.test(t)) {
                        periode = t.split('\\u00b7')[0].trim();
                        break;
                    }
                }
                postes.push({titre, societe, periode});
            }
        }

        if (postes.length >= 2) break;
    }

    return JSON.stringify(postes);
} catch(e) { return JSON.stringify({error: e.message, stack: String(e.stack || '')}); }
})();
"""


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
        Vérifie d'abord que Firefox n'est pas déjà ouvert.
        """
        self._verifier_firefox_ferme()

        options = Options()
        profil = self.firefox_profile_path or self._detecter_profil_macos()

        if profil:
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

    def _verifier_firefox_ferme(self) -> None:
        """
        Vérifie que Firefox n'est pas en cours d'exécution.
        Si oui, demande à l'utilisateur de le fermer avant de continuer.
        """
        while self._firefox_est_ouvert():
            log.warning("Firefox est déjà ouvert.")
            log.warning("Fermez Firefox pour que le script puisse utiliser votre profil LinkedIn.")
            try:
                input("  Appuyez sur Entrée une fois Firefox fermé… ")
            except (KeyboardInterrupt, EOFError):
                raise RuntimeError("Lancement annulé — Firefox doit être fermé.")

        log.info("Firefox est fermé — lancement possible.")

    @staticmethod
    def _firefox_est_ouvert() -> bool:
        """Retourne True si un processus Firefox est actif."""
        try:
            result = subprocess.run(
                ["pgrep", "-x", "firefox"],
                capture_output=True,
            )
            return result.returncode == 0
        except FileNotFoundError:
            return False

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
        """
        if not self.driver:
            raise RuntimeError("Le driver n'est pas démarré. Appelez demarrer() d'abord.")

        self.driver.get(url_recherche)
        time.sleep(3)

        current = self.driver.current_url
        if "authwall" in current or "/login" in current:
            log.warning(f"Accès refusé sur la page de recherche : {current}")
            return None

        try:
            liens = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='/in/']")
        except Exception as e:
            log.error(f"Erreur lors de la recherche des liens /in/ : {e}")
            return None

        for lien in liens:
            href = lien.get_attribute("href") or ""
            if "/in/" not in href:
                continue
            url_propre = self._nettoyer_url_profil(href)
            if url_propre:
                log.info(f"Profil trouvé via recherche : {url_propre}")
                return url_propre

        log.warning(f"Aucun profil /in/ trouvé sur : {url_recherche}")
        return None

    def _nettoyer_url_profil(self, href: str) -> Optional[str]:
        """
        Extrait https://www.linkedin.com/in/slug sans les paramètres d'URL.
        """
        try:
            parsed = urlparse(href)
            chemin = parsed.path.rstrip("/")
            if not chemin.startswith("/in/"):
                return None
            propre = urlunparse((
                "https", "www.linkedin.com", chemin, "", "", "",
            ))
            return propre
        except Exception:
            return None

    def est_profil_valide(self) -> bool:
        """
        Vérifie que la page courante est bien un profil LinkedIn.
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

        Utilise un script JavaScript qui traverse le DOM par structure
        (pas par classes CSS, car LinkedIn les obfusque).

        Étapes :
          1. Scroll clavier (Page Down) pour déclencher le lazy-loading
          2. Trouver le heading "Experience"/"Expérience"
          3. Remonter à la <section> conteneur
          4. Extraire titre, société, période depuis les <p> dans les <li>

        Retourne :
            Liste de 0, 1 ou 2 dicts avec les clés : titre, societe, periode.
            Liste vide → recours au screenshot + Claude Vision.
        """
        self._scroll_vers_experience()

        try:
            import json
            result_json = self.driver.execute_script("return " + _JS_EXTRACT_EXPERIENCE)
            if not result_json:
                log.warning("Le script JS n'a rien retourné.")
                return []
            postes = json.loads(result_json) if isinstance(result_json, str) else result_json
            # Vérifier si le JS a retourné une erreur
            if isinstance(postes, dict) and "error" in postes:
                log.error(f"Erreur JavaScript : {postes['error']}")
                log.debug(f"Stack JS : {postes.get('stack', '')}")
                return []
        except Exception as e:
            log.error(f"Erreur lors de l'extraction JavaScript : {e}")
            return []

        postes = postes[:2]

        if postes:
            log.info(
                f"DOM : {len(postes)} poste(s) extrait(s) — "
                + " | ".join(f"{p['titre']} @ {p['societe']}" for p in postes)
            )
        else:
            log.warning("Aucun poste extrait du DOM.")

        return postes

    # ----------------------------------------------------------
    # Scroll vers la section Expérience
    # ----------------------------------------------------------

    def _scroll_vers_experience(self) -> None:
        """
        Fait défiler la page avec des touches clavier (Page Down)
        pour déclencher le lazy-loading de LinkedIn.
        Attend que le texte "Experience"/"Expérience" apparaisse dans le DOM.
        """
        if not self.driver:
            return

        # Donner le focus à la page
        try:
            self.driver.find_element(By.TAG_NAME, "body").click()
        except Exception:
            pass
        time.sleep(0.5)

        actions = ActionChains(self.driver)

        for i in range(15):
            actions.send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(1)

            found = self.driver.execute_script("""
                return document.body.innerText.includes('Experience')
                    || document.body.innerText.includes('Expérience');
            """)
            if found:
                log.info(f"Section Experience trouvée après {i+1} Page Down.")
                # Continuer un peu pour charger tous les postes
                for _ in range(2):
                    actions.send_keys(Keys.PAGE_DOWN).perform()
                    time.sleep(1)
                # Scroller vers le heading Experience
                self.driver.execute_script("""
                    const walker = document.createTreeWalker(
                        document.body, NodeFilter.SHOW_TEXT);
                    while (walker.nextNode()) {
                        const txt = walker.currentNode.textContent.trim();
                        if (txt === 'Experience' || txt === 'Expérience') {
                            walker.currentNode.parentElement.scrollIntoView(
                                {block: 'start'});
                            break;
                        }
                    }
                """)
                time.sleep(2)
                return

        log.warning("Section Experience non trouvée après scroll complet.")

    # ----------------------------------------------------------
    # Screenshot (fallback si DOM vide)
    # ----------------------------------------------------------

    def take_screenshot(self, max_largeur: int = 1280) -> Optional[bytes]:
        """
        Prend un screenshot ciblé sur la section Expérience si possible,
        sinon capture la page entière en fallback.
        Retourne les octets JPEG compressés, ou None en cas d'erreur.
        """
        if not self.driver:
            log.error("Impossible de prendre un screenshot : driver non démarré.")
            return None

        try:
            # La section est déjà scrollée par get_experience_dom()
            time.sleep(1)

            # Tenter un screenshot ciblé via JavaScript
            png_bytes = None
            try:
                section_found = self.driver.execute_script("""
                    const walker = document.createTreeWalker(
                        document.body, NodeFilter.SHOW_TEXT);
                    while (walker.nextNode()) {
                        const txt = walker.currentNode.textContent.trim();
                        if (txt === 'Experience' || txt === 'Expérience') {
                            let section = walker.currentNode.parentElement;
                            while (section && section.tagName !== 'SECTION') {
                                section = section.parentElement;
                            }
                            if (section) {
                                arguments[0](section);
                                return true;
                            }
                        }
                    }
                    return false;
                """)
                # Si trouvé, screenshot de la section
                if section_found:
                    section_el = self.driver.execute_script("""
                        const walker = document.createTreeWalker(
                            document.body, NodeFilter.SHOW_TEXT);
                        while (walker.nextNode()) {
                            const txt = walker.currentNode.textContent.trim();
                            if (txt === 'Experience' || txt === 'Expérience') {
                                let section = walker.currentNode.parentElement;
                                while (section && section.tagName !== 'SECTION') {
                                    section = section.parentElement;
                                }
                                return section;
                            }
                        }
                        return null;
                    """)
                    if section_el:
                        png_bytes = section_el.screenshot_as_png
                        log.info("Screenshot ciblé section Experience.")
            except Exception:
                pass

            # Fallback : screenshot pleine page
            if not png_bytes:
                png_bytes = self.driver.get_screenshot_as_png()
                log.info("Screenshot pleine page (section Experience non ciblée)")

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
