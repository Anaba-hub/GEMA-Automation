# GEMA Automation

Scraper LinkedIn automatise pour extraire les 2 derniers postes professionnels d'etudiants eligibles au programme GEMA. Le systeme traite en masse des profils LinkedIn via Firefox (Selenium), avec un fallback Claude Vision pour les profils dont le DOM est illisible.

## Architecture

```
main.py              # Orchestrateur principal
config.py            # Configuration (cle API, delais, chemins)
estimate_time.py     # Estimation du temps sans lancer le scraper
retry.py             # Relance les profils en erreur

agents/
  agent_excel.py     # Lecture gema.xlsx + ecriture output_final.xlsx
  agent_browser.py   # Automatisation Firefox via Selenium
  agent_vision.py    # Fallback : analyse screenshot via Claude Vision

utils/
  logger.py          # Logs colores par agent
  helpers.py         # Progression, barre de progression, rapport
```

## Prerequis

- Python 3.10+
- Firefox installe avec une session LinkedIn active
- [geckodriver](https://github.com/mozilla/geckodriver/releases) installe (`brew install geckodriver`)
- Une cle API [Anthropic](https://console.anthropic.com/) (pour le fallback Vision)

## Installation

```bash
# Cloner le repo
git clone https://github.com/Anaba-hub/GEMA-Automation.git
cd GEMA-Automation

# Environnement virtuel
python3 -m venv .venv
source .venv/bin/activate

# Dependances
pip install -r requirements.txt

# Configuration
cp config.py.example config.py
# Editer config.py : renseigner ANTHROPIC_API_KEY
```

## Fichier d'entree

Placer un fichier `gema.xlsx` a la racine du projet avec les colonnes suivantes :

| Colonne | Description |
|---------|-------------|
| id | Identifiant unique (optionnel) |
| Prenom | Prenom de l'etudiant |
| Nom | Nom de l'etudiant |
| Annee | Annee / cohorte |
| Ecole | Nom de l'ecole |
| Recherche LinkedIn | URL LinkedIn (`/in/...` ou `/search/...`) |
| Poste 1 | Poste deja connu (optionnel, skip si rempli avec Poste 2) |
| Poste 2 | Poste deja connu (optionnel) |

## Utilisation

### Estimer le temps de traitement

```bash
python estimate_time.py
```

### Lancer le scraper

```bash
python main.py
```

Le script :
1. Charge `gema.xlsx` et reprend depuis `progress.json` si interruption precedente
2. Ouvre Firefox avec la session LinkedIn existante
3. Pour chaque etudiant : resout le lien, ouvre le profil, extrait les 2 derniers postes
4. Si le DOM est vide, prend un screenshot et utilise Claude Vision en fallback
5. Sauvegarde les resultats dans `output_final.xlsx`

### Relancer les profils en erreur

```bash
python retry.py
```

Supprime les lignes en erreur de `output_final.xlsx` et `progress.json`, puis relancez `main.py`.

## Fichiers generes

| Fichier | Description |
|---------|-------------|
| `output_final.xlsx` | Resultats (14 colonnes : identite, liens, 2 postes, statut) |
| `progress.json` | Etat de progression (reprise apres interruption) |
| `rapport.txt` | Statistiques de la session |
| `errors.log` | Detail des erreurs non fatales |
| `linkedin_scraper.log` | Log complet de debug |

## Statuts de traitement

| Statut | Signification |
|--------|---------------|
| `ok` | Postes extraits avec succes |
| `deja_traite` | Poste 1 et 2 deja remplis dans gema.xlsx |
| `sans_lien` | Pas d'URL LinkedIn valide |
| `acces_refuse` | Mur de connexion LinkedIn (session expiree) |
| `profil_non_trouve` | Lien /search/ sans resultat |
| `profil_sans_exp` | Profil charge mais aucune experience |
| `erreur` | Erreur technique (voir errors.log) |

## Protections integrees

- **Delais aleatoires** : 4-8s entre chaque profil
- **Pauses automatiques** : 3 min tous les 50 profils, 10 min tous les 200, 20 min tous les 500
- **Detection de deconnexion** : apres 3 acces refuses consecutifs, le script pause et demande de se reconnecter
- **Reprise** : `progress.json` permet de reprendre apres Ctrl+C ou crash
- **caffeinate** : empeche la mise en veille macOS pendant les longs runs
- **Sauvegarde atomique** : ecriture via fichier temporaire + rename pour eviter la corruption
