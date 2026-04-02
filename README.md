# GEMA Automation

Outil qui extrait automatiquement les 2 derniers postes professionnels depuis les profils LinkedIn d'anciens etudiants.

## Comment ca marche

```
gema.xlsx  ──>  main.py  ──>  output_final.xlsx
 (entree)      (Firefox)        (resultats)
```

Le script ouvre Firefox, cherche chaque profil LinkedIn, lit ses experiences et ecrit le resultat dans un fichier Excel.

## Structure

```
main.py          # Script principal (lancement + relance des erreurs)
config.py        # Configuration (delais, chemin Firefox)
lib/
  browser.py     # Pilotage de Firefox
  excel.py       # Lecture/ecriture Excel
  helpers.py     # Progression et utilitaires
  logger.py      # Journalisation
```

## Prerequis

- **Python 3.10+**
- **Firefox** avec une session LinkedIn active (connectez-vous manuellement une fois)
- **geckodriver** : `brew install geckodriver`

## Installation

```bash
git clone https://github.com/Anaba-hub/GEMA-Automation.git
cd GEMA-Automation

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Creez un fichier `config.py` a la racine (un modele existe dans le repo).

## Fichier d'entree (gema.xlsx)

Placez le fichier `gema.xlsx` dans le dossier du projet. Colonnes attendues :

| Colonne | Obligatoire | Description |
|---------|:-----------:|-------------|
| Prenom | oui | Prenom |
| Nom | oui | Nom |
| Recherche LinkedIn | oui | URL LinkedIn (`/search/...` ou `/in/...`) |
| Annee | non | Promotion |
| Ecole | non | Etablissement |
| Poste 1 / Poste 2 | non | Si les deux sont remplis, le profil est ignore |

## Utilisation

**Fermez Firefox** avant chaque lancement.

### Lancer le traitement

```bash
source .venv/bin/activate
python main.py
```

Le script affiche un resume et demande confirmation. Repondez `o` pour lancer.

### Relancer les erreurs

```bash
python main.py --retry
```

Les profils en erreur sont supprimes du fichier de sortie et retraites automatiquement.

### Interruption

`Ctrl+C` arrete proprement le script. La progression est sauvegardee. Relancez `python main.py` pour reprendre.

## Resultats

| Fichier | Contenu |
|---------|---------|
| `output_final.xlsx` | Resultats : identite, liens, 2 postes, statut |
| `rapport.txt` | Statistiques de la session |
| `errors.log` | Detail des erreurs |

### Statuts

| Statut | Signification |
|--------|---------------|
| `ok` | Postes extraits |
| `deja_traite` | Postes deja connus dans gema.xlsx |
| `sans_lien` | Pas de lien LinkedIn |
| `profil_non_trouve` | Recherche sans resultat |
| `acces_refuse` | Session LinkedIn expiree |
| `profil_sans_exp` | Profil sans section Experience |
| `erreur` | Erreur technique (voir errors.log) |

## En cas de probleme

| Probleme | Solution |
|----------|----------|
| "Firefox est deja ouvert" | Fermez Firefox puis appuyez sur Entree |
| 3 acces refuses d'affilee | Reconnectez-vous a LinkedIn dans Firefox, puis Entree |
| Le script s'arrete | Relancez `python main.py`, il reprend ou il en etait |
| Beaucoup de `profil_sans_exp` | Ces profils n'ont pas de section Experience visible |
