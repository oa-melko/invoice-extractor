---

## 🚀 Démarrage rapide — UNE seule commande

Ce projet est un ensemble de **scripts Python autonomes** (pas de serveur web).

Après `git clone`, depuis la racine du projet :

**Windows (PowerShell)**
```powershell
.\go.ps1 install                      # installe les deps
.\go.ps1 list                         # liste les scripts
.\go.ps1 run extract_factures.py      # execute un script
```

**Linux / macOS / Git Bash**
```bash
bash go.sh install
bash go.sh list
bash go.sh run extract_factures.py
```

Raccourci : `.\go.ps1 mon_script.py` équivaut à `.\go.ps1 run mon_script.py`.

### Sous-commandes
| But | PowerShell | Bash |
|---|---|---|
| Installer les deps | `.\go.ps1 install` | `bash go.sh install` |
| Lister les scripts | `.\go.ps1 list` | `bash go.sh list` |
| Lancer un script | `.\go.ps1 run <s>` | `bash go.sh run <s>` |

> Pré-requis : `python` (3.11+), `git` accessibles dans le PATH.

### 🐳 Alternative : Docker (aucune install Python locale requise)

L'outil etant un CLI (pas de serveur), on utilise `docker compose run` :

```bash
docker compose build                                                # 1ere fois
docker compose run --rm extractor python extract_factures.py        # variante OCR pur
docker compose run --rm extractor python extract_factures_llm.py    # variante OCR + LLM
```

Les dossiers d'entree (`donnee a analyser/`) et les Excel de sortie restent sur l'hote (bind-mount).

Pour la **variante LLM** (Ollama) : Ollama doit tourner sur l'hote. Le compose
passe deja `OLLAMA_URL=http://host.docker.internal:11434/api/generate` pour que
le script reach l'Ollama host depuis le container.

---
# Extraction Automatique de Factures - GIP DU TREGOR-GOELO

Outil d'extraction automatique des donnees de factures scannees (PDF) vers un fichier Excel structure, utilisant l'OCR et un LLM local.


---

## 🚀 Démarrage rapide — UNE seule commande

Ce projet est un ensemble de **scripts Python autonomes** (pas de serveur web).

Après `git clone`, depuis la racine du projet :

**Windows (PowerShell)**
```powershell
.go.ps1 install                      # installe les deps
.go.ps1 list                         # liste les scripts
.go.ps1 run extract_factures.py      # execute un script
```

**Linux / macOS / Git Bash**
```bash
bash go.sh install
bash go.sh list
bash go.sh run extract_factures.py
```

Raccourci : `.go.ps1 mon_script.py` équivaut à `.go.ps1 run mon_script.py`.

### Sous-commandes
| But | PowerShell | Bash |
|---|---|---|
| Installer les deps | `.go.ps1 install` | `bash go.sh install` |
| Lister les scripts | `.go.ps1 list` | `bash go.sh list` |
| Lancer un script | `.go.ps1 run <s>` | `bash go.sh run <s>` |

> Pré-requis : `python` (3.11+), `git` accessibles dans le PATH.
> Si `requirements.txt` n'existe pas, il sera créé vide au premier `install`. Ajoute-y tes dépendances pip.

---

## Fonctionnalites

- OCR de factures PDF scannees (images sans texte extractible) via **docTR**
- Extraction intelligente des donnees par **LLM local** (Ollama/gemma2) avec prompts specifiques par fournisseur
- Export Excel avec **surlignage jaune** des cellules a faible confiance OCR
- Support de **6 fournisseurs** : Cloro'fil, CTM Style, Halbout, Mulliez-Flory, Poyet-Motte, Tissus Gisele
- Tracabilite : chaque ligne indique le fichier PDF source

## Prerequis

- Python 3.10+
- [Ollama](https://ollama.ai/) installe et lance (`ollama serve`)
- Modele gemma2 telecharge (`ollama pull gemma2`)

### Dependances Python

```bash
pip install python-doctr[torch] openpyxl requests
```

## Utilisation rapide

```bash
# Lancer sur le dossier par defaut (donnee a analyser/)
python extract_factures_llm.py

# Lancer sur un autre dossier
python extract_factures_llm.py "C:\chemin\vers\mes_factures"
```

Le fichier `Saisie Achat Linge - RESULTAT.xlsx` est genere dans le dossier du script.

## Structure du projet

```
analyse/
├── extract_factures_llm.py          # Script principal (OCR + LLM)
├── extract_factures.py              # Version regex (sans LLM)
├── prompts_extraction_factures.md   # Prompts detailles par fournisseur
├── documentation.md                 # Documentation technique complete
├── README.md                        # Ce fichier
├── donnee a analyser/               # Dossier des factures PDF
│   ├── Clorofil_Facture du *.pdf
│   ├── CTM_Facture du *.pdf
│   ├── Halbout_Facture du *.pdf
│   ├── Mulliez_Facture du *.pdf
│   ├── Poyet Motte_Facture du *.pdf
│   ├── Tissus Gisèle_Facture du *.pdf
│   └── Saisie Achat Linge - GIP DU TREGOR-GOËLO 2025.xlsx  # Modele original
└── Saisie Achat Linge - RESULTAT.xlsx  # Fichier genere
```

## Sortie Excel

| Colonne | Contenu |
|---------|---------|
| A - Fournisseur | Nom du fournisseur (1 seule fois par groupe) |
| B - Date | Date de la facture |
| C - Reference Article | Code produit / EAN |
| D - Article | Description du produit |
| E - Composants | Specifications techniques (taille, couleur, matiere) |
| F - Quantite | Nombre commande |
| G - Fichier | Nom du PDF source |

Les cellules en **jaune** signalent une confiance OCR inferieure au seuil (80% par defaut).

## Configuration

En haut de `extract_factures_llm.py` :

```python
CONFIDENCE_THRESHOLD = 0.80          # Seuil de surlignage (0.0 a 1.0)
INPUT_DIR = "donnee a analyser"      # Dossier d'entree
OUTPUT_FILE = "Saisie Achat Linge - RESULTAT.xlsx"  # Fichier de sortie
OLLAMA_MODEL = "gemma2:latest"       # Modele Ollama a utiliser
```

## Deux versions du script

| | `extract_factures_llm.py` | `extract_factures.py` |
|---|---|---|
| Methode d'extraction | LLM (Ollama/gemma2) | Regex |
| Precision | Elevee | Moyenne |
| Composants | Bien separes | Non separes |
| Tolerance OCR | Haute (comprend les erreurs) | Faible |
| Vitesse | ~30-60 sec/facture | ~5-10 sec/facture |
| Dependance | Ollama requis | Aucune |