# ============================================================
# utils/logger.py — Logs colorés avec préfixes par agent
# ============================================================

import logging
import sys

# Codes couleurs ANSI
_RESET  = "\033[0m"
_BOLD   = "\033[1m"
_COLORS = {
    "EXCEL":         "\033[94m",   # Bleu clair
    "SEARCH":        "\033[93m",   # Jaune
    "BROWSER":       "\033[96m",   # Cyan
    "VISION":        "\033[95m",   # Magenta
    "ORCHESTRATEUR": "\033[92m",   # Vert
}
_LEVEL_COLORS = {
    "DEBUG":    "\033[37m",    # Gris
    "INFO":     "\033[0m",     # Normal
    "WARNING":  "\033[33m",    # Jaune vif
    "ERROR":    "\033[31m",    # Rouge
    "CRITICAL": "\033[41m",    # Fond rouge
}


class _ColorFormatter(logging.Formatter):
    """Formateur qui colore le préfixe agent et le niveau."""

    def format(self, record: logging.LogRecord) -> str:
        agent  = getattr(record, "agent", "")
        prefix = f"{_BOLD}{_COLORS.get(agent, '')}" \
                 f"[{agent}]{_RESET} " if agent else ""
        level_color = _LEVEL_COLORS.get(record.levelname, "")
        message = super().format(record)
        return f"{prefix}{level_color}{message}{_RESET}"


def get_logger(nom_agent: str) -> logging.Logger:
    """
    Retourne un logger configuré pour un agent donné.

    Usage :
        log = get_logger("SEARCH")
        log.info("Recherche en cours…")   # affiche [SEARCH] Recherche en cours…
    """
    logger = logging.getLogger(nom_agent)

    # Évite les doublons si le logger est déjà configuré
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    # --- Handler console (couleurs) ---
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    fmt = _ColorFormatter("%(asctime)s %(message)s", datefmt="%H:%M:%S")
    console.setFormatter(fmt)

    # --- Handler fichier (sans couleurs, tous niveaux) ---
    fichier = logging.FileHandler("linkedin_scraper.log", encoding="utf-8")
    fichier.setLevel(logging.DEBUG)
    fmt_fichier = logging.Formatter(
        "%(asctime)s [%(name)s] %(levelname)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )
    fichier.setFormatter(fmt_fichier)

    logger.addHandler(console)
    logger.addHandler(fichier)

    # Injecte automatiquement l'attribut "agent" dans chaque enregistrement
    old_info    = logger.info
    old_warning = logger.warning
    old_error   = logger.error
    old_debug   = logger.debug

    def _inject(fn):
        def wrapper(msg, *args, **kwargs):
            extra = kwargs.pop("extra", {})
            extra["agent"] = nom_agent
            return fn(msg, *args, extra=extra, **kwargs)
        return wrapper

    logger.info    = _inject(old_info)
    logger.warning = _inject(old_warning)
    logger.error   = _inject(old_error)
    logger.debug   = _inject(old_debug)

    return logger
