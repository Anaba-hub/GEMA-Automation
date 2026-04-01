# ============================================================
# estimate_time.py — Estimation du temps avant lancement
# ============================================================
# Lit gema.xlsx + progress.json et affiche une estimation
# du temps de traitement sans lancer main.py.
#
# Usage :
#   source .venv/bin/activate
#   python estimate_time.py
# ============================================================

import config
from agents.agent_excel import load_gema, est_lien_valide
from utils.helpers      import charger_progression


def fmt_duree(secondes: int) -> str:
    """Convertit des secondes en chaîne lisible : '2h 35min' ou '45min'."""
    h, r = divmod(secondes, 3600)
    m, _ = divmod(r, 60)
    return f"{h}h {m:02d}min" if h else f"{m}min"


def main():
    # Chargement des personnes depuis gema.xlsx
    personnes = load_gema(config.FICHIER_GEMA)
    if not personnes:
        print(f"[ERREUR] Fichier introuvable ou vide : {config.FICHIER_GEMA}")
        return

    # Chargement de la progression existante
    progression = charger_progression(config.FICHIER_PROGRESSION)
    traites_set = set(progression.get("traites", []))

    # Classification
    nb_reprendre = 0  # déjà dans progress.json (session précédente)
    nb_deja      = 0  # Poste 1 + 2 remplis dans gema → skip immédiat
    nb_avec      = 0  # lien valide → visite Firefox nécessaire
    nb_sans      = 0  # pas de lien → ignoré

    for p in personnes:
        cle = p["id"] if p.get("id") else f"{p['nom']}_{p['prenom']}"
        if cle in traites_set:
            nb_reprendre += 1
        elif p["deja_traite"]:
            nb_deja += 1
        elif est_lien_valide(p["lien"]):
            nb_avec += 1
        else:
            nb_sans += 1

    # Calcul des temps (~10s par profil + 15% pauses, ×1.5 réaliste)
    t_base   = nb_avec * 10
    t_pauses = int(t_base * 0.15)
    t_opt    = t_base + t_pauses
    t_real   = int(t_opt * 1.5)

    sep  = "=" * 52
    sep2 = "-" * 50

    print(f"\n{sep}")
    print(f"  ESTIMATION — {config.FICHIER_GEMA}")
    print(sep)
    print(f"  Total dans gema.xlsx             : {len(personnes):>6}")
    print(f"  {sep2}")
    if nb_reprendre:
        print(f"  Déjà traités (progress.json)     : {nb_reprendre:>6}  →   0s")
    print(f"  Déjà traités (gema Poste 1+2)    : {nb_deja:>6}  →   0s")
    print(f"  Avec lien LinkedIn               : {nb_avec:>6}  → ~10s/profil")
    print(f"  Sans lien LinkedIn               : {nb_sans:>6}  → ignorés")
    print(f"  {sep2}")
    print(f"  Profils à traiter effectivement  : {nb_avec:>6}")
    print(f"  Pauses automatiques estimées     : {fmt_duree(t_pauses)}")
    print(f"  {sep2}")
    print(f"  Temps optimiste  (+15% pauses)   : {fmt_duree(t_opt)}")
    print(f"  Temps réaliste   (×1.5)          : {fmt_duree(t_real)}")
    print(sep)
    print()

    if nb_avec == 0:
        print("  Rien à traiter : tous les profils sont déjà traités ou sans lien.")
    else:
        print(f"  Pour lancer : source .venv/bin/activate && python main.py")
    print()


if __name__ == "__main__":
    main()
