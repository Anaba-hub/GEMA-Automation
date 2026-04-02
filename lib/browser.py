import glob
import json
import os
import subprocess
import time
from urllib.parse import urlparse, urlunparse

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.options import Options

from lib.logger import get_logger

log = get_logger("BROWSER")


_JS_EXTRACT_EXPERIENCE = """(function() {
try {
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

    let section = expHeading;
    while (section && section.tagName !== 'SECTION') section = section.parentElement;
    if (!section) return JSON.stringify([]);

    const datePattern = /\\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|janv|f[eé]v|mars|avr|mai|juin|juil|ao[uû]t|sept|oct|nov|d[eé]c|\\d{4}|present|pr[eé]sent)\\b/i;

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


class Browser:

    def __init__(self, firefox_profile_path=None):
        self.firefox_profile_path = firefox_profile_path
        self.driver = None

    def demarrer(self):
        self._verifier_firefox_ferme()

        options = Options()
        profil = self.firefox_profile_path or self._detecter_profil_macos()

        if profil:
            options.add_argument("-profile")
            options.add_argument(profil)
            options.add_argument("-no-remote")
            log.info(f"Profil Firefox : {profil}")
        else:
            log.warning("Aucun profil Firefox trouvé.")

        try:
            self.driver = webdriver.Firefox(options=options)
            log.info("Firefox demarré.")
            return self.driver
        except Exception as e:
            log.error(f"Impossible de demarrer Firefox : {e}")
            log.error("Verifiez que geckodriver est installé (brew install geckodriver).")
            raise

    def fermer(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
            log.info("Firefox fermé.")

    def ouvrir(self, url):
        if not self.driver:
            raise RuntimeError("Driver non démarré.")
        self.driver.get(url)
        time.sleep(3)
        log.info(f"Page ouverte : {url}")

    def extract_profile_from_search(self, url_recherche):
        if not self.driver:
            raise RuntimeError("Driver non démarré.")

        self.driver.get(url_recherche)
        time.sleep(3)

        current = self.driver.current_url
        if "authwall" in current or "/login" in current:
            log.warning(f"Accès refusé : {current}")
            return None

        try:
            liens = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='/in/']")
        except Exception as e:
            log.error(f"Erreur recherche liens /in/ : {e}")
            return None

        for lien in liens:
            href = lien.get_attribute("href") or ""
            if "/in/" not in href:
                continue
            url_propre = self._nettoyer_url(href)
            if url_propre:
                log.info(f"Profil trouvé : {url_propre}")
                return url_propre

        log.warning(f"Aucun profil trouvé : {url_recherche}")
        return None

    def est_profil_valide(self):
        if not self.driver:
            return False
        current = self.driver.current_url
        if "authwall" in current or "/login" in current:
            return False
        if "404" in current or "not-found" in current:
            return False
        if "linkedin.com/in/" not in current:
            return False
        return True

    def get_experience_dom(self):
        self._scroll_vers_experience()

        try:
            result_json = self.driver.execute_script("return " + _JS_EXTRACT_EXPERIENCE)
            if not result_json:
                log.warning("Script JS : aucun résultat.")
                return []
            postes = json.loads(result_json) if isinstance(result_json, str) else result_json
            if isinstance(postes, dict) and "error" in postes:
                log.error(f"Erreur JS : {postes['error']}")
                return []
        except Exception as e:
            log.error(f"Erreur extraction JS : {e}")
            return []

        postes = postes[:2]
        if postes:
            log.info(
                f"DOM : {len(postes)} poste(s) — "
                + " | ".join(f"{p['titre']} @ {p['societe']}" for p in postes)
            )
        else:
            log.warning("Aucun poste extrait du DOM.")
        return postes

    # --- Privé ---

    def _scroll_vers_experience(self):
        if not self.driver:
            return
        try:
            self.driver.find_element(By.TAG_NAME, "body").click()
        except Exception:
            pass
        time.sleep(0.5)

        actions = ActionChains(self.driver)
        for i in range(15):
            actions.send_keys(Keys.PAGE_DOWN).perform()
            time.sleep(1)

            found = self.driver.execute_script(
                "return document.body.innerText.includes('Experience')"
                " || document.body.innerText.includes('Expérience');"
            )
            if found:
                log.info(f"Section Experience trouvée après {i+1} Page Down.")
                for _ in range(2):
                    actions.send_keys(Keys.PAGE_DOWN).perform()
                    time.sleep(1)
                self.driver.execute_script("""
                    const walker = document.createTreeWalker(
                        document.body, NodeFilter.SHOW_TEXT);
                    while (walker.nextNode()) {
                        const txt = walker.currentNode.textContent.trim();
                        if (txt === 'Experience' || txt === 'Expérience') {
                            walker.currentNode.parentElement.scrollIntoView({block: 'start'});
                            break;
                        }
                    }
                """)
                time.sleep(2)
                return

        log.warning("Section Experience non trouvée après scroll complet.")

    def _verifier_firefox_ferme(self):
        while self._firefox_est_ouvert():
            log.warning("Firefox est déjà ouvert. Fermez-le pour continuer.")
            try:
                input("  Appuyez sur Entrée une fois Firefox fermé… ")
            except (KeyboardInterrupt, EOFError):
                raise RuntimeError("Annulé — Firefox doit être fermé.")
        log.info("Firefox fermé — lancement possible.")

    @staticmethod
    def _firefox_est_ouvert():
        try:
            return subprocess.run(["pgrep", "-x", "firefox"], capture_output=True).returncode == 0
        except FileNotFoundError:
            return False

    @staticmethod
    def _detecter_profil_macos():
        base = os.path.expanduser("~/Library/Application Support/Firefox/Profiles/")
        for pattern in ["*.default-release", "*.default"]:
            matches = glob.glob(os.path.join(base, pattern))
            if matches:
                return matches[0]
        return None

    @staticmethod
    def _nettoyer_url(href):
        try:
            parsed = urlparse(href)
            chemin = parsed.path.rstrip("/")
            if not chemin.startswith("/in/"):
                return None
            return urlunparse(("https", "www.linkedin.com", chemin, "", "", ""))
        except Exception:
            return None
