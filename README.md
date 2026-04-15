# Extraction Automatique de Factures - GIP DU TREGOR-GOELO

Outil d'extraction automatique des donnees de factures scannees (PDF) vers un fichier Excel structure, utilisant l'OCR et un LLM local.

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
