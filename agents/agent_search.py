# ============================================================
# agents/agent_search.py — Recherche Google → URL LinkedIn (vérifiée)
# ============================================================
# Protections pour run longue durée (1 600 profils) :
#   - Délais progressifs selon le nombre de recherches effectuées
#   - Pause automatique toutes les 30 recherches (5-10 min)
#   - Pause automatique toutes les 100 recherches (20-30 min)
#   - CAPTCHA → pause 45 min puis retry
#   - Vérification nom+prénom dans l'URL retournée
# ============================================================

import re
import time
import random
import unicodedata

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from utils.logger import get_logger
from utils.helpers import nettoyer_url_linkedin

log = get_logger("SEARCH")


# ----------------------------------------------------------
# Délais progressifs : plus on avance, plus on ralentit
# ----------------------------------------------------------
_PALIERS_DELAI = [
    (50,   3,  7),   # profils   1- 50 : 3-7s
    (200,  5, 10),   # profils  51-200 : 5-10s
    (500,  8, 15),   # profils 201-500 : 8-15s
    (None, 10, 20),  # profils 501+    : 10-20s
]

def _delai_pour_compteur(compteur: int) -> float:
    """Retourne un délai aléatoire en secondes selon le nombre de recherches déjà faites."""
    for seuil, dmin, dmax in _paliers_delai():
        if seuil is None or compteur <= seuil:
            return random.uniform(dmin, dmax)
    return random.uniform(10, 20)

def _paliers_delai():
    return _PALIERS_DELAI


# ----------------------------------------------------------
# Vérification nom/prénom dans l'URL
# ----------------------------------------------------------

def _normaliser(texte: str) -> str:
    """Supprime les accents, met en minuscules, garde lettres/chiffres/espaces/tirets."""
    nfkd = unicodedata.normalize("NFKD", texte)
    sans_accents = "".join(c for c in nfkd if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9\s\-]", "", sans_accents.lower())


def _score_correspondance(nom: str, prenom: str, slug_url: str, texte_lien: str) -> int:
    """
    Score de correspondance entre la personne cherchée et un résultat Google.
    2 = nom+prénom dans le slug URL  (très fiable)
    1 = nom+prénom dans le texte du résultat (fiable)
    0 = aucune correspondance (profil rejeté)
    """
    slug_norm   = _normaliser(slug_url)
    texte_norm  = _normaliser(texte_lien)
    nom_norm    = _normaliser(nom)
    prenom_norm = _normaliser(prenom)

    seuil = 2  # longueur minimale d'un token pour être significatif
    tokens_nom    = [t for t in nom_norm.split()    if len(t) >= seuil] or nom_norm.split()
    tokens_prenom = [t for t in prenom_norm.split() if len(t) >= seuil] or prenom_norm.split()

    def tous_presents(tokens, texte):
        return all(t in texte for t in tokens)

    if tous_presents(tokens_nom, slug_norm) and tous_presents(tokens_prenom, slug_norm):
        return 2
    if tous_presents(tokens_nom, texte_norm) and tous_presents(tokens_prenom, texte_norm):
        return 1
    return 0


class AgentSearch:
    """
    Agent de recherche Google.
    Conçu pour tenir 1 600 recherches consécutives avec protection anti-blocage.
    """

    MOTS_BLOCAGE = ["unusual traffic", "captcha", "i'm not a robot", "sorry"]
    SCORE_MIN    = 1   # score minimum pour accepter un profil

    def __init__(self, driver, delai_min: float, delai_max: float, max_retries: int = 3):
        self.driver          = driver
        self.delai_min       = delai_min   # conservé pour compatibilité (non utilisé directement)
        self.delai_max       = delai_max
        self.max_retries     = max_retries
        self.compteur        = 0           # nombre de recherches effectuées depuis le démarrage

    def search_profile(self, nom: str, prenom: str, ecole1: str, ecole2: str) -> tuple:
        """
        Recherche + vérification du profil LinkedIn.
        Retourne (url, statut) :
          "trouvé"      → URL validée (nom+prénom confirmés)
          "non_vérifié" → URL trouvée mais correspondance incertaine
          "non_trouvé"  → Aucun résultat
          "bloqué"      → Google a bloqué, même après les retries
          "erreur"      → Erreur technique
        """
        requete = (
            f'site:linkedin.com/in "{prenom} {nom}" '
            f'"{ecole1}" OR "{ecole2}"'
        )
        log.info(f"Requête [{self.compteur + 1}] : {requete}")

        # Pauses longues automatiques selon le compteur
        self._pause_longue_si_necessaire()

        for tentative in range(1, self.max_retries + 1):
            try:
                url, statut = self._executer_recherche(requete, nom, prenom)

                if statut == "bloqué":
                    if tentative < self.max_retries:
                        self._pause_captcha(tentative)
                        continue
                    else:
                        log.error(f"Toutes les tentatives épuisées pour {prenom} {nom}.")
                        return "", "bloqué"

                self.compteur += 1
                return url, statut

            except Exception as e:
                log.error(f"Erreur inattendue (tentative {tentative}/{self.max_retries}) : {e}")
                if tentative == self.max_retries:
                    return "", "erreur"
                time.sleep(5)

        return "", "bloqué"

    def _executer_recherche(self, requete: str, nom: str, prenom: str) -> tuple:
        """Effectue une recherche Google et retourne (url, statut)."""
        self.driver.get("https://www.google.com")
        wait = WebDriverWait(self.driver, 15)

        # Accepter les cookies si bannière présente
        try:
            bouton = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Tout accepter') or contains(., 'Accept all')]")
            ))
            bouton.click()
            time.sleep(1)
        except TimeoutException:
            pass

        # Saisir la requête
        barre = wait.until(EC.presence_of_element_located((By.NAME, "q")))
        barre.clear()
        barre.send_keys(requete)
        barre.send_keys(Keys.RETURN)

        try:
            wait.until(EC.presence_of_element_located((By.ID, "search")))
        except TimeoutException:
            log.warning("Timeout lors du chargement des résultats Google.")
            return "", "erreur"

        # Détection CAPTCHA / blocage
        page = self.driver.page_source.lower()
        if any(mot in page for mot in self.MOTS_BLOCAGE):
            log.warning("Blocage Google détecté (CAPTCHA ou trafic inhabituel).")
            return "", "bloqué"

        # Collecter les candidats LinkedIn
        candidats = self._collecter_candidats_linkedin()
        if not candidats:
            log.info(f"Aucun profil LinkedIn dans les résultats pour {prenom} {nom}.")
            return "", "non_trouvé"

        # Évaluer chaque candidat
        meilleur_url, meilleur_score = "", -1
        for url, texte in candidats:
            slug  = url.rstrip("/").split("/in/")[-1]
            score = _score_correspondance(nom, prenom, slug, texte)
            log.info(f"  Candidat score={score} : {url} | '{texte[:60]}'")
            if score > meilleur_score:
                meilleur_score, meilleur_url = score, url

        if meilleur_score >= self.SCORE_MIN:
            via = "slug URL" if meilleur_score == 2 else "texte résultat"
            log.info(f"Profil validé (score={meilleur_score} via {via}) : {meilleur_url}")
            return meilleur_url, "trouvé"

        log.warning(
            f"Profil(s) trouvé(s) mais non vérifiable(s) pour {prenom} {nom} "
            f"(meilleur score={meilleur_score})"
        )
        return meilleur_url, "non_vérifié"

    def _collecter_candidats_linkedin(self) -> list:
        """Retourne les tuples (url, texte_titre) de tous les résultats LinkedIn."""
        candidats = []
        blocs = self.driver.find_elements(By.CSS_SELECTOR, "#search .g, #rso > div")
        for bloc in blocs:
            try:
                lien = bloc.find_element(By.CSS_SELECTOR, "a[href]")
                href = lien.get_attribute("href") or ""
                if "linkedin.com/in/" not in href:
                    continue
                url_propre = nettoyer_url_linkedin(href)
                try:
                    texte = bloc.find_element(By.CSS_SELECTOR, "h3").text.strip()
                except Exception:
                    texte = lien.text.strip()
                if url_propre:
                    candidats.append((url_propre, texte))
            except Exception:
                continue
        return candidats

    # ----------------------------------------------------------
    # Gestion des pauses anti-blocage
    # ----------------------------------------------------------

    def pause(self) -> None:
        """
        Pause standard entre deux recherches.
        Le délai augmente progressivement selon self.compteur.
        """
        duree = _delai_pour_compteur(self.compteur)
        palier = self._nom_palier(self.compteur)
        log.info(f"Pause [{palier}] : {duree:.1f}s (compteur={self.compteur})")
        time.sleep(duree)

    def _pause_longue_si_necessaire(self) -> None:
        """
        Insère une pause longue automatique à des jalons critiques.
        Appelé AVANT chaque recherche.
        """
        c = self.compteur
        if c > 0 and c % 100 == 0:
            duree = random.uniform(20 * 60, 30 * 60)
            log.info(
                f"PAUSE LONGUE ({c} recherches effectuées) : "
                f"{duree / 60:.0f} minutes — Google cooling down…"
            )
            time.sleep(duree)
        elif c > 0 and c % 30 == 0:
            duree = random.uniform(5 * 60, 10 * 60)
            log.info(
                f"Pause intermédiaire ({c} recherches) : "
                f"{duree / 60:.0f} minutes…"
            )
            time.sleep(duree)

    def _pause_captcha(self, tentative: int) -> None:
        """Pause de 45 minutes après détection d'un CAPTCHA."""
        duree = 45 * 60
        log.warning(
            f"CAPTCHA détecté (tentative {tentative}/{self.max_retries}). "
            f"Pause de 45 minutes. Résolvez le CAPTCHA dans Firefox si possible."
        )
        # Compte à rebours toutes les 5 minutes pour informer l'utilisateur
        for restant in range(int(duree / 60), 0, -5):
            time.sleep(5 * 60)
            log.info(f"  Reprise dans ~{restant - 5} minutes…")

    @staticmethod
    def _nom_palier(compteur: int) -> str:
        """Retourne un label lisible pour le palier actuel."""
        if compteur <= 50:   return "palier 1-50"
        if compteur <= 200:  return "palier 51-200"
        if compteur <= 500:  return "palier 201-500"
        return "palier 501+"
