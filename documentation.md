# Documentation Technique - Extraction Automatique de Factures

## 1. Vue d'ensemble

Ce projet automatise l'extraction de donnees depuis des factures fournisseurs scannees (PDF images) vers un fichier Excel structure. Il combine deux technologies :

- **docTR** : moteur OCR (Optical Character Recognition) qui convertit les images en texte
- **Ollama + gemma2** : LLM local qui interprete le texte OCR et en extrait les donnees structurees

### Pipeline de traitement

```
┌─────────────┐     ┌─────────────┐     ┌──────────────────┐     ┌─────────────┐
│  Facture PDF │────>│  docTR OCR  │────>│  Ollama/gemma2   │────>│  Excel .xlsx│
│  (image)     │     │  (texte)    │     │  (JSON structure)│     │  (resultat) │
└─────────────┘     └─────────────┘     └──────────────────┘     └─────────────┘
                          │                      │
                     confiance OCR          prompt specifique
                     par mot/ligne          par fournisseur
```

---

## 2. Architecture du script

### 2.1 Fichier principal : `extract_factures_llm.py`

Le script est organise en 5 sections :

#### Section 1 : Configuration (lignes 13-17)

```python
CONFIDENCE_THRESHOLD = 0.80    # Seuil en dessous duquel une cellule est surlignee
INPUT_DIR = "donnee a analyser"
OUTPUT_FILE = "Saisie Achat Linge - RESULTAT.xlsx"
OLLAMA_MODEL = "gemma2:latest"
OLLAMA_URL = "http://localhost:11434/api/generate"
```

#### Section 2 : Mapping fournisseurs

Trois dictionnaires lient les fichiers PDF aux fournisseurs :

| Dictionnaire | Role | Exemple |
|---|---|---|
| `SUPPLIER_PREFIX_MAP` | Prefixe fichier → cle interne | `"Clorofil" → "clorofil"` |
| `SUPPLIER_DISPLAY` | Cle interne → nom affiche | `"clorofil" → "Cloro'fil Concept"` |
| `SUPPLIER_PROMPTS` | Cle interne → prompt LLM | `"clorofil" → "Rôle: Extracteur..."` |

La detection du fournisseur se fait par le **prefixe du nom de fichier** :
- `Clorofil_Facture du *.pdf` → Cloro'fil Concept
- `CTM_Facture du *.pdf` → CTM STYLE
- `Halbout_Facture du *.pdf` → HALBOUT SAS
- `Mulliez_Facture du *.pdf` → MULLIEZ-FLORY
- `Poyet Motte_Facture du *.pdf` → POYET-MOTTE
- `Tissus Gis*_Facture du *.pdf` → TISSUS GISELE

#### Section 3 : Fonctions OCR

**`ocr_pdf(filepath, model)`**
- Charge le PDF via `DocumentFile.from_pdf()`
- Execute l'OCR avec le modele docTR pre-entraine
- Retourne le dictionnaire d'export docTR (pages > blocks > lines > words)

**`extract_text_with_confidence(ocr_export)`**
- Parcourt la structure pages/blocks/lines/words
- Concatene le texte de toutes les lignes
- Calcule la confiance moyenne par ligne (moyenne des confiances de chaque mot)
- Retourne le texte complet + la confiance moyenne globale du document

#### Section 4 : Fonctions LLM

**`call_ollama(prompt, ocr_text)`**
- Construit le prompt complet : prompt fournisseur + texte OCR + instructions JSON
- Tronque le texte OCR a 6000 caracteres pour eviter les debordements de contexte
- Appelle l'API Ollama (`POST /api/generate`) avec temperature 0.1 (reponses deterministes)
- Timeout de 120 secondes par appel

**`parse_llm_response(response_text)`**
- Nettoie la reponse (retire les balises markdown ```json```)
- Cherche un tableau JSON valide dans la reponse
- Gere les erreurs courantes (virgules en trop, JSON mal forme)
- Retourne une liste de dictionnaires ou une liste vide en cas d'echec

**`extract_items_llm(ocr_text, supplier_key, filename)`**
- Orchestre l'appel LLM et le parsing de la reponse
- Normalise les cles JSON (gere les variantes avec/sans accents)
- Valide chaque item : quantite > 0, reference ou article non vide
- Normalise les quantites (supprime separateurs de milliers)
- Parse les dates en objet datetime
- Retourne la liste des items valides

**Normalisation des cles JSON :**

Le LLM peut retourner des cles en francais avec accents. Le script gere ces variantes :

| Cle LLM possible | Cle normalisee |
|---|---|
| `Fournisseur` | `fournisseur` |
| `Référence Article` | `reference_article` |
| `Quantité` / `Quantite` | `quantite` |
| `Composants` | `composants` |

#### Section 5 : Excel Writer

**`write_excel(all_items, output_path, threshold)`**
- Cree un classeur avec une feuille "2025"
- En-tetes en gras : Fournisseur, Date, Reference Article, Article, Composants, Quantite, Fichier
- Groupe les items par fournisseur (dans l'ordre de `SUPPLIER_ORDER`)
- Le nom du fournisseur n'apparait qu'une fois par groupe (colonne A, premiere ligne)
- Surligne en jaune (`FFFF00`) les cellules dont la confiance OCR < seuil
- Derniere ligne : formule `=SUM()` pour le total des quantites
- Largeurs de colonnes ajustees automatiquement

---

## 3. Module OCR : docTR

### 3.1 Pourquoi docTR

Les factures sont des **scans** (images dans le PDF, aucun texte extractible). Les librairies classiques comme PyMuPDF retournent 0 caractere. docTR utilise un modele de deep learning pour :

1. **Detection** : localiser les zones de texte dans l'image
2. **Reconnaissance** : convertir chaque zone en caracteres

### 3.2 Structure de sortie docTR

```
export['pages']              # Liste de pages
  └── page['blocks']         # Blocs de texte detectes
       └── block['lines']    # Lignes dans chaque bloc
            └── line['words'] # Mots dans chaque ligne
                 ├── word['value']      # Texte du mot ("FACTURE")
                 └── word['confidence'] # Confiance 0.0 a 1.0
```

### 3.3 Confiance OCR

Chaque mot a un score de confiance entre 0 et 1 :
- **> 0.90** : lecture fiable
- **0.70 - 0.90** : lecture probable, a verifier
- **< 0.70** : lecture incertaine

La confiance moyenne du document est propagee a toutes les cellules Excel de cette facture. Si elle est inferieure au seuil (0.80), les cellules sont surlignees en jaune.

---

## 4. Module LLM : Ollama + gemma2

### 4.1 Pourquoi un LLM local

L'approche precedente (regex) presentait des limites :
- Fragile face aux variations OCR ("REVERSIBLEI" au lieu de "REVERSIBLE")
- Un parseur different par fournisseur a maintenir
- Pas de separation article/composants
- Echecs frequents sur les formats complexes (Mulliez multi-tailles)

Le LLM comprend le contexte et tolere les erreurs OCR. Chaque fournisseur a un **prompt dedie** qui decrit :
- Les champs a extraire
- Le format attendu de la facture
- Les regles specifiques (quoi ignorer, comment combiner les champs)
- Un exemple de sortie JSON

### 4.2 Parametres de l'appel LLM

```python
{
    "model": "gemma2:latest",    # Modele 9B parametres
    "stream": False,              # Reponse complete (pas de streaming)
    "options": {
        "temperature": 0.1,       # Quasi-deterministe
        "num_predict": 4096       # Tokens max en sortie
    }
}
```

- **temperature 0.1** : minimise la creativite, maximise la precision factuelle
- **num_predict 4096** : suffisant pour ~50 items JSON
- **timeout 120s** : gemma2 9B peut etre lent sur CPU

### 4.3 Prompts par fournisseur

Les prompts sont stockes dans `SUPPLIER_PROMPTS` (dans le script) et documentes en detail dans `prompts_extraction_factures.md`.

Chaque prompt suit la meme structure :
1. **Role** : contexte metier pour le LLM
2. **Champs a extraire** : liste numerotee avec exemples
3. **Regles specifiques** : points d'attention marques `[CRITIQUE]`
4. **Exemple de sortie** : JSON attendu

### 4.4 Gestion des erreurs LLM

| Probleme | Gestion |
|---|---|
| Ollama non demarre | Message d'erreur explicite + arret |
| Timeout (>120s) | Item marque en erreur, continue |
| JSON invalide | Tentative de reparation (virgules, balises) |
| Cles avec accents | Normalisation automatique |
| Quantite en string | Conversion avec nettoyage des separateurs |
| Pas d'items retournes | Log WARNING, continue |

---

## 5. Formats de factures par fournisseur

### 5.1 Cloro'fil Concept
- **Format** : Facture classique A4, tableau unique
- **Particularite** : Reference sur 2 lignes (code interne + code produit)
- **Champs** : Reference combinee, description, composants separes, quantite

### 5.2 CTM STYLE (Chorus Pro)
- **Format** : 2 pages, page 1 = recapitulatif, page 2 = detail articles
- **Particularite** : References EAN13, format Chorus avec "Denomination de l'article"
- **Champs** : EAN, denomination, quantite facturee

### 5.3 HALBOUT SAS
- **Format** : Tableau simple 6 colonnes, descriptions techniques longues
- **Particularite** : Descriptions de 50 a 200 caracteres a separer
- **Champs** : Code article, nom produit (court), composants (technique)

### 5.4 MULLIEZ-FLORY
- **Format** : Tableau dense multi-colonnes, multi-tailles par article
- **Particularite** : Reference 6 chiffres + coloris, quantites par taille a sommer
- **Champs** : Reference-coloris, article, composants (taille/couleur), quantite totale

### 5.5 POYET-MOTTE (Chorus Pro)
- **Format** : Identique a CTM (Chorus Pro)
- **Particularite** : Parfois format direct (ref NEGOC-xxxxx) au lieu de Chorus
- **Champs** : EAN ou ref directe, denomination, quantite

### 5.6 TISSUS GISELE
- **Format** : Tableau 6 colonnes ou Chorus Pro selon la facture
- **Particularite** : Quantites au format francais (2.000,00 = 2000), codes avec dimensions
- **Champs** : Code dimensionnel, description, composition textile, quantite

---

## 6. Surlignage et confiance

### 6.1 Principe

Le surlignage jaune dans l'Excel signale les donnees potentiellement incorrectes :

```
Confiance OCR du document >= 0.80  →  cellules normales
Confiance OCR du document <  0.80  →  cellules surlignees en jaune
```

### 6.2 Interpretation

- **Cellule normale** : l'OCR a lu le document avec une bonne confiance. Le LLM a probablement extrait des donnees fiables.
- **Cellule jaune** : l'OCR a eu des difficultes (scan de mauvaise qualite, texte flou). Les donnees extraites par le LLM peuvent contenir des erreurs. **Verification manuelle recommandee.**

### 6.3 Ajuster le seuil

```python
CONFIDENCE_THRESHOLD = 0.80  # Modifier cette valeur
```

- **0.90** : strict, beaucoup de cellules surlignees
- **0.80** : equilibre (valeur par defaut)
- **0.70** : permissif, peu de surlignage

---

## 7. Performance et limites

### 7.1 Temps de traitement

| Etape | Temps par facture | Total (~71 PDFs) |
|---|---|---|
| OCR docTR | 5-15 sec | ~10 min |
| LLM gemma2 (CPU) | 20-60 sec | ~30-45 min |
| LLM gemma2 (GPU) | 5-15 sec | ~10-15 min |
| Ecriture Excel | < 1 sec | < 1 sec |

### 7.2 Limites connues

- **Scans de tres mauvaise qualite** : l'OCR peut echouer completement (confiance < 0.50)
- **Factures multi-pages (>3 pages)** : le texte OCR peut depasser la limite de 6000 caracteres
- **Nouveau fournisseur** : necessite d'ajouter un prompt dans `SUPPLIER_PROMPTS` et un prefixe dans `SUPPLIER_PREFIX_MAP`
- **Coherence LLM** : gemma2 peut occasionnellement retourner un JSON mal forme ou omettre un item. Relancer sur le fichier individuel corrige generalement le probleme.

### 7.3 Version regex (fallback)

Le script `extract_factures.py` (version regex, sans LLM) est conserve comme alternative :
- Ne necessite pas Ollama
- Plus rapide (~5 sec/facture)
- Moins precis (pas de separation article/composants, sensible aux erreurs OCR)

---

## 8. Ajout d'un nouveau fournisseur

1. **Identifier le prefixe du nom de fichier** (ex: `"NouveauFournisseur"`)

2. **Ajouter dans les dictionnaires** du script :

```python
SUPPLIER_PREFIX_MAP["NouveauFournisseur"] = "nouveau_fournisseur"
SUPPLIER_DISPLAY["nouveau_fournisseur"] = "NOUVEAU FOURNISSEUR"
SUPPLIER_ORDER.append("NOUVEAU FOURNISSEUR")
```

3. **Ecrire le prompt** dans `SUPPLIER_PROMPTS["nouveau_fournisseur"]` en suivant la structure :
   - Role et tache
   - 6 champs a extraire avec exemples
   - Regles specifiques du format de facture
   - Exemple de sortie JSON

4. **Tester** sur une facture :

```python
python -c "
from extract_factures_llm import *
from doctr.models import ocr_predictor
model = ocr_predictor(pretrained=True)
export = ocr_pdf('chemin/vers/facture.pdf', model)
text, conf = extract_text_with_confidence(export)
items = extract_items_llm(text, 'nouveau_fournisseur', 'facture.pdf')
for it in items: print(it)
"
```

---

## 9. Depannage

| Symptome | Cause probable | Solution |
|---|---|---|
| `PermissionError` sur le fichier Excel | Fichier ouvert dans Excel | Fermer le fichier et relancer |
| `Cannot connect to Ollama` | Ollama non demarre | Lancer `ollama serve` |
| `500 Server Error` d'Ollama | Texte trop long ou modele non charge | Verifier `ollama list`, reduire `num_predict` |
| 0 items extraits | LLM n'a pas retourne de JSON valide | Verifier le texte OCR manuellement |
| Cellules toutes jaunes | Scan de mauvaise qualite | Rescanner la facture en meilleure resolution |
| Quantites incorrectes | Separateurs de milliers mal interpretes | Verifier le prompt du fournisseur |
