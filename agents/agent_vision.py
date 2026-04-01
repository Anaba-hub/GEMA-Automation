# ============================================================
# agents/agent_vision.py — Extraction des postes via Claude Vision
# ============================================================
# Responsabilités :
#   - Recevoir un screenshot JPEG (bytes)
#   - L'envoyer à l'API Claude Vision (claude-sonnet-4-20250514)
#   - Parser la réponse pour extraire les 2 derniers postes
#   - Retourner [{titre, entreprise, dates}, {titre, entreprise, dates}]
# ============================================================

import base64
import time

import anthropic

from utils.logger import get_logger

log = get_logger("VISION")

# Prompt envoyé à Claude Vision avec le screenshot
_PROMPT_VISION = """\
Voici un screenshot d'un profil LinkedIn. Regarde uniquement la section "Expérience".

Identifie les 2 DERNIÈRES expériences (les plus récentes, en haut de la section).
Pour chaque expérience, extrais séparément :
- le TITRE DU POSTE (ex: "Directeur Commercial", "Ingénieure Data")
- la SOCIÉTÉ (nom de l'entreprise employeur, ex: "BNP Paribas", "Accenture")
- la PÉRIODE (dates de début et fin, ex: "janv. 2022 - présent", "mars 2019 - déc. 2021")

IMPORTANT : ne confonds pas le titre du poste et le nom de la société.
Le titre du poste est ce que fait la personne. La société est l'employeur.

Réponds UNIQUEMENT dans ce format exact, sans rien ajouter d'autre :

POSTE 1
Titre: <intitulé exact du poste>
Société: <nom exact de l'entreprise>
Période: <dates de début - fin>

POSTE 2
Titre: <intitulé exact du poste>
Société: <nom exact de l'entreprise>
Période: <dates de début - fin>

Si un champ est introuvable sur le screenshot, écris "Non disponible".
"""


class AgentVision:
    """
    Agent d'analyse d'images via l'API Claude Vision.
    Utilisé en fallback quand le DOM LinkedIn est illisible.
    """

    def __init__(self, api_key: str, model: str):
        """
        Args:
            api_key : clé API Anthropic
            model   : identifiant du modèle Claude Vision (ex. claude-sonnet-4-20250514)
        """
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model  = model

    def analyze_screenshot(self, image_bytes: bytes, max_retries: int = 3) -> str:
        """
        Envoie le screenshot à l'API Claude Vision et retourne la réponse brute.
        Retry automatique avec backoff exponentiel en cas d'erreur transitoire.

        Args:
            image_bytes : bytes de l'image JPEG
            max_retries : nombre max de tentatives (défaut 3)

        Retourne :
            Texte brut de la réponse Claude, ou chaîne vide en cas d'erreur.
        """
        if not image_bytes:
            log.error("Screenshot vide — impossible d'appeler Claude Vision.")
            return ""

        image_b64 = base64.b64encode(image_bytes).decode("utf-8")
        log.info(
            f"Envoi du screenshot à Claude Vision "
            f"({len(image_bytes) // 1024} Ko, modèle {self.model})…"
        )

        for tentative in range(1, max_retries + 1):
            try:
                reponse = self.client.messages.create(
                    model=self.model,
                    max_tokens=600,
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type":   "image",
                                    "source": {
                                        "type":       "base64",
                                        "media_type": "image/jpeg",
                                        "data":       image_b64,
                                    },
                                },
                                {"type": "text", "text": _PROMPT_VISION},
                            ],
                        }
                    ],
                )
                texte = reponse.content[0].text
                log.info("Réponse Claude Vision reçue.")
                log.debug(f"Réponse brute :\n{texte}")
                return texte

            except anthropic.AuthenticationError:
                log.error("Clé API Anthropic invalide. Vérifiez ANTHROPIC_API_KEY dans config.py.")
                return ""  # pas de retry, erreur permanente

            except (anthropic.RateLimitError, anthropic.APITimeoutError, anthropic.APIConnectionError) as e:
                delai = 2 ** tentative  # 2s, 4s, 8s
                if tentative < max_retries:
                    log.warning(
                        f"Erreur transitoire (tentative {tentative}/{max_retries}) : {e}. "
                        f"Retry dans {delai}s…"
                    )
                    time.sleep(delai)
                else:
                    log.error(f"Échec après {max_retries} tentatives : {e}")
                    return ""

            except Exception as e:
                log.error(f"Erreur inattendue lors de l'appel Claude Vision : {e}")
                return ""  # pas de retry, erreur inconnue

        return ""

    def extract_positions(self, reponse_brute: str) -> list[dict]:
        """
        Parse la réponse textuelle de Claude Vision pour extraire
        les informations structurées des 2 postes.

        Retourne :
            Liste de 0, 1 ou 2 dicts avec les clés : titre, entreprise, dates.
        """
        if not reponse_brute.strip():
            log.warning("Réponse Claude Vision vide — aucun poste extrait.")
            return []

        postes = []
        # On découpe la réponse à chaque "POSTE "
        blocs = reponse_brute.strip().split("POSTE ")

        for bloc in blocs[1:3]:  # On ne traite que POSTE 1 et POSTE 2
            poste = self._parser_bloc(bloc)
            postes.append(poste)
            log.info(
                f"Poste extrait : {poste['titre']} @ {poste['societe']} ({poste['periode']})"
            )

        log.info(f"Vision : {len(postes)} poste(s) parsé(s).")
        return postes

    def _parser_bloc(self, bloc: str) -> dict:
        """
        Parse un bloc de texte du type :
            1
            Titre: Ingénieur logiciel
            Entreprise: Acme Corp
            Dates: janv. 2022 - présent

        Retourne un dict {titre, entreprise, dates}.
        """
        mapping = {
            "Titre":   "titre",
            "Société": "societe",
            "Période": "periode",
            # Variantes au cas où Claude répond légèrement différemment
            "Societe":    "societe",
            "Periode":    "periode",
            "Entreprise": "societe",   # fallback si ancien format
            "Dates":      "periode",   # fallback si ancien format
        }
        resultat = {"titre": "Non disponible", "societe": "Non disponible", "periode": "Non disponible"}

        for ligne in bloc.strip().splitlines():
            if ":" not in ligne:
                continue
            cle_brute, _, valeur = ligne.partition(":")
            cle_brute = cle_brute.strip()
            valeur    = valeur.strip()
            if cle_brute in mapping:
                resultat[mapping[cle_brute]] = valeur or "Non disponible"

        return resultat

    def analyser_et_extraire(self, image_bytes: bytes) -> list[dict]:
        """
        Méthode de commodité : appelle analyze_screenshot puis extract_positions.
        Retourne directement la liste des postes.
        """
        reponse = self.analyze_screenshot(image_bytes)
        return self.extract_positions(reponse)
