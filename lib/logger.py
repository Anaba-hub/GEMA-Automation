import logging
import os
import sys

# Activer les couleurs ANSI sur Windows 10+
if sys.platform == "win32":
    os.system("")  # active le mode VT100 dans CMD

_RESET = "\033[0m"
_BOLD  = "\033[1m"
_COLORS = {
    "EXCEL":   "\033[94m",
    "BROWSER": "\033[96m",
    "MAIN":    "\033[92m",
}
_LEVEL_COLORS = {
    "DEBUG":    "\033[37m",
    "INFO":     "\033[0m",
    "WARNING":  "\033[33m",
    "ERROR":    "\033[31m",
    "CRITICAL": "\033[41m",
}


class _ColorFormatter(logging.Formatter):
    def format(self, record):
        agent = getattr(record, "agent", "")
        prefix = f"{_BOLD}{_COLORS.get(agent, '')}[{agent}]{_RESET} " if agent else ""
        level_color = _LEVEL_COLORS.get(record.levelname, "")
        message = super().format(record)
        return f"{prefix}{level_color}{message}{_RESET}"


def get_logger(nom_agent):
    logger = logging.getLogger(nom_agent)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.INFO)
    console.setFormatter(_ColorFormatter("%(asctime)s %(message)s", datefmt="%H:%M:%S"))

    fichier = logging.FileHandler("linkedin_scraper.log", encoding="utf-8")
    fichier.setLevel(logging.DEBUG)
    fichier.setFormatter(logging.Formatter(
        "%(asctime)s [%(name)s] %(levelname)s — %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))

    logger.addHandler(console)
    logger.addHandler(fichier)

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
