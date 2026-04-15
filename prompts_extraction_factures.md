# Prompts d'Extraction Automatique - Factures Fournisseurs

## Contexte
Ce document contient les prompts spécifiques pour l'extraction automatique des champs de factures selon la structure unique de chaque fournisseur.

---

## Règles Générales (Applicable à tous)

```json
{
  "output_format": {
    "fournisseur": "string",
    "date": "JJ/MM/AAAA",
    "reference_article": "string",
    "article": "string", 
    "composants": "string",
    "quantite": "integer",
    "fichier": "string"
  },
  "conversions": {
    "date_excel_serial": "Convertir en date lisible (ex: 45905 → 05/09/2025)",
    "quantite_format_fr": "2.000,00 → 2000 (point = séparateur milliers)",
    "quantite_format_en": "2,000.00 → 2000 (virgule = séparateur milliers)"
  },
  "validation": [
    "Vérifier que la somme des quantités × prix unitaire ≈ Total HT",
    "Ignorer les lignes sans référence article",
    "Ignorer les totaux et sous-totaux",
    "Conserver le nom exact du fichier PDF source"
  ]
}
```

---

## 1. CLORO'FIL CONCEPT

### Identification
- **Mots-clés détection**: "CLORO'FIL", "Cloro'fil Concept", "clorofilconcept.com"
- **Structure**: Facture classique A4, tableau unique centré, mentions "Prix au 100"

### Prompt d'extraction

```text
Rôle: Extracteur de facture textile professionnel
Tâche: Extraire les lignes articles de la facture Cloro'fil Concept

CHAMPS À EXTRAIRE:
1. Fournisseur: "Cloro'fil Concept"
2. Date: Date de facture (format: 05/09/2025)
3. Référence Article: Combiner les 2 lignes de référence 
   - Ligne 1: Code numérique (ex: "000025")
   - Ligne 2: Code produit (ex: "GT350JNE")
   → Format final: "000025 GT350JNE"
4. Article: Première ligne de la désignation (ex: "GANT DE TOILETTE")
5. Composants: Spécifications techniques (ex: "350 g/m2 Jaune")
6. Quantité: Nombre entier de la colonne "Quantité" (ex: 13000)
7. Fichier: Nom du fichier PDF

RÈGLES SPÉCIFIQUES:
- [CRITIQUE] La référence est toujours sur 2 niveaux dans le tableau
- Ignorer impérativement les lignes contenant "Prix au 100" (indicateur de prix, pas d'article)
- Si plusieurs articles: créer une entrée par ligne article
- Date Excel: convertir le nombre sériel (ex: 45905) en date JJ/MM/AAAA
- Le fournisseur peut être écrit "Cloro'fil Concept" ou "CLORO'FIL Concept"

EXEMPLE DE SORTIE ATTENDUE:
```json
{
  "fournisseur": "Cloro'fil Concept",
  "date": "05/09/2025",
  "reference_article": "000025 GT350JNE",
  "article": "GANT DE TOILETTE",
  "composants": "350 g/m2 Jaune",
  "quantite": 13000,
  "fichier": "Clorofil_Facture du 5 09 25.pdf"
}
```

---

## 2. CTM STYLE (Format Chorus Pro)

### Identification
- **Mots-clés détection**: "CTM STYLE", "CTM STYLE SASU", "Chorus", "Emetteur/Client"
- **Structure**: 2 pages obligatoires, entête gauche/droite, page 2 = "Articles rattachés"

### Prompt d'extraction

```text
Rôle: Extracteur facture électronique Chorus Pro
Tâche: Extraire les données de facture CTM Style (format Chorus)

CHAMPS À EXTRAIRE:
1. Fournisseur: "CTM Style" (depuis bloc "Entité juridique")
2. Date: Date facture page 1 (ex: "14/11/2025")
3. Référence Article: "Référence produit" page 2 (EAN13 à 13 chiffres, ex: 3617540000462)
4. Article: "Dénomination de l'article" (ex: "ROMEO 80JF 313/40")
5. Composants: Laisser vide ou extraire suffixe après le nom si pertinent
6. Quantité: Colonne "Quantité facturée" page 2 (ex: 5,00 → 5)
7. Fichier: Nom du fichier PDF

RÈGLES SPÉCIFIQUES:
- [CRITIQUE] Aller systématiquement sur la page 2 intitulée "Articles rattachés au compte client"
- La page 1 ne contient que le récapitulatif global (à ignorer pour les lignes)
- Un même article peut apparaître sur plusieurs lignes avec quantités différentes → SOMMER les quantités si même référence EAN
- Format des prix: utiliser la virgule comme séparateur décimal (format FR)
- Ignorer les lignes "Totaux du site de livraison" et "Brut HT/Net HT"

POST-TRAITEMENT:
- Agréger les quantités par référence EAN unique
- Vérifier que la somme correspond au "Total HT" page 1

EXEMPLE DE SORTIE ATTENDUE:
```json
{
  "fournisseur": "CTM Style",
  "date": "14/11/2025", 
  "reference_article": "3617540000462",
  "article": "ROMEO 80JF 313/40",
  "composants": "",
  "quantite": 11,
  "fichier": "CTM_Facture du 14 11 25.pdf"
}
```

---

## 3. HALBOUT SAS

### Identification
- **Mots-clés détection**: "HALBOUT", "ALM", "TEXTILES FOR LAUNDRY", "Docelles"
- **Structure**: Logo "AM" circulaire en haut, tableau simple 6 colonnes, mentions techniques longues

### Prompt d'extraction

```text
Rôle: Extracteur facture équipement hospitalier et textile
Tâche: Extraire les lignes articles de facture Halbout SAS

CHAMPS À EXTRAIRE:
1. Fournisseur: "HALBOUT SAS"
2. Date: Date facture (ex: "03/02/2025")
3. Référence Article: Colonne "Article" (code alphanumérique, ex: ORE06006025)
4. Article: Première partie de la désignation (nom du produit uniquement)
5. Composants: Toute la description technique suivante (dimensions, matière, normes)
6. Quantité: Colonne "Quantité" (entier, ex: 100)
7. Fichier: Nom du fichier PDF

RÈGLES SPÉCIFIQUES:
- [CRITIQUE] Les descriptions sont très longues et techniques (50-200 caractères)
- Couper la description au premier espace significatif pour "Article"
- Conserver l'intégralité du texte technique dans "Composants"
- Format: "OREILLER MICRONEW 60x60..." → Article: "OREILLER MICRONEW", Composants: "60x60 BLC IGNIFUGE..."
- Vérifier le taux TVA (20,00) pour validation
- Ignorer les lignes de livraison "Livré à: CENTRE HOSPITALIER..."

POST-TRAITEMENT:
- Nettoyer les espaces multiples dans les descriptions
- Standardiser les unités (UN, PIECES, etc.)

EXEMPLE DE SORTIE ATTENDUE:
```json
{
  "fournisseur": "HALBOUT SAS",
  "date": "03/02/2025",
  "reference_article": "ORE06006025",
  "article": "OREILLER MICRONEW",
  "composants": "60x60 BLC IGNIFUGE OREILLER IMPERMEABLE - 60x60 - BLANC ENDUIT POLYURETHANE - ENVELOPPE JERSEY POLYESTER - POIDS 110 g/m2",
  "quantite": 100,
  "fichier": "Halbout_Facture du 3 02 25.pdf"
}
```

---

## 4. MULLIEZ-FLORY

### Identification
- **Mots-clés détection**: "Mulliez-Flory", "Groupe Mulliez-Flory", "SELFIA", "Dress for business"
- **Structure**: En-tête avec logo, tableau grid dense multi-colonnes, 2 pages (page 2 = tampon)

### Prompt d'extraction

```text
Rôle: Extracteur facture textile professionnel (format industriel dense)
Tâche: Extraire les données du tableau complexe Mulliez-Flory

CHAMPS À EXTRAIRE:
1. Fournisseur: "MULLIEZ-FLORY"
2. Date: Date facture (ex: "04/06/2025")
3. Référence Article: Concaténer "REFERENCE" + "COLORIS" (ex: "036484" + "M01ZZ" → "036484-M01ZZ")
4. Article: Colonne "ARTICLE" (ex: "BAVOIR CONFORBEL CROCUS BP")
5. Composants: Colonne "COLORIS" + "TAILLE" (ex: "BEIGE MARRON 045*090")
6. Quantité: Colonne "QTES" (quantité unitaire par ligne, ex: 1000) 
   ⚠️ NE PAS prendre la colonne "TOTAL" ni "TOTAL COMMANDE"
7. Fichier: Nom du fichier PDF

RÈGLES SPÉCIFIQUES:
- [CRITIQUE] Structure tabulaire complexe avec colonnes: ARTICLE, COLORIS, DOUANE, MODE, COMPO, CODE TAXE, REFERENCE, GENCOD, TAILLE, QTES, PRIX UNITAIRE, MONTANT
- [CRITIQUE] Un même article peut avoir plusieurs coloris (M01ZZ, R01ZZ) → Créer 2 lignes distinctes
- La référence est numérique (6 chiffres), le coloris est alphanumérique (4-5 caractères)
- Page 2 ne contient que le tampon "GIP SITG" et les totaux → Ignorer pour l'extraction articles
- Vérifier que la somme des lignes = TOTAL HT BRUT (et non TOTAL HT NET si remises)

POST-TRAITEMENT:
- Séparer les entrées par coloris
- Vérifier cohérence: Quantité × Prix Unitaire = Montant ligne

EXEMPLE DE SORTIE ATTENDUE:
```json
{
  "fournisseur": "MULLIEZ-FLORY",
  "date": "04/06/2025",
  "reference_article": "036484-M01ZZ",
  "article": "BAVOIR CONFORBEL CROCUS BP",
  "composants": "BEIGE MARRON 045*090",
  "quantite": 1000,
  "fichier": "Mulliez_Facture du 4 06 25.pdf"
}
```

---

## 5. POYET MOTTE (Format Chorus Pro)

### Identification
- **Mots-clés détection**: "POYET MOTTE", "POYET MOTTE SAS", "Cours (FR)"
- **Structure**: Identique à CTM Style (Format Chorus), 2 pages

### Prompt d'extraction

```text
Rôle: Extracteur facture électronique Chorus Pro
Tâche: Extraire les données de facture Poyet Motte

CHAMPS À EXTRAIRE:
1. Fournisseur: "POYET-MOTTE" ou "Poyet Motte"
2. Date: Date facture page 1 (ex: "07/11/2025")
3. Référence Article: "Référence produit" page 2 (EAN13, ex: 3120760082768)
4. Article: "Dénomination de l'article" (ex: "POLECO OUR SV")
5. Composants: Laisser vide sauf si détails pertinents après le nom
6. Quantité: "Quantité facturée" page 2 (ex: 300,00 → 300)
7. Fichier: Nom du fichier PDF

RÈGLES SPÉCIFIQUES:
- [CRITIQUE] Même structure exacte que CTM Style (Chorus Pro) → se référer au prompt CTM
- Page 2: "Articles rattachés au compte client" contient les données
- Vérifier la présence d'un numéro d'engagement (ex: E2025000802) pour validation
- Une seule ligne article typiquement, mais prévoir multi-lignes
- Date de livraison parfois présente (14/11/2025) mais ne pas confondre avec date facture

VALIDATION:
- Vérifier Total HT page 1 = somme (Quantité × Prix Unitaire) page 2

EXEMPLE DE SORTIE ATTENDUE:
```json
{
  "fournisseur": "POYET-MOTTE",
  "date": "07/11/2025",
  "reference_article": "3120760082768",
  "article": "POLECO OUR SV",
  "composants": "",
  "quantite": 300,
  "fichier": "Poyet Motte_Factue du 7 11 25.pdf"
}
```

---

## 6. TISSUS GISÈLE

### Identification
- **Mots-clés détection**: "TISSUS GISÈLE", "TISSUS GISELE", "TGL", "La Bresse"
- **Structure**: Logo TGL rond, tableau 6 colonnes, mention "Commande SOLDEE", en-tête texte riche

### Prompt d'extraction

```text
Rôle: Extracteur facture textile linge de maison
Tâche: Extraire les données de facture Tissus Gisèle

CHAMPS À EXTRAIRE:
1. Fournisseur: "TISSUS GISELE"
2. Date: Date facture (ex: "16/01/2025")
3. Référence Article: Colonne "Code" (ex: "DR C4 180X320 PC BLANCL2E")
4. Article: Colonne "Désignation" - première partie (ex: "DRAPS PLATS OURLETS PIQUES DE 4CM AU FIL")
5. Composants: Suite de la désignation (dimensions, composition: "180,0 x 320,0 BLANC UNI POLYESTER-COTON 50/50")
6. Quantité: Colonne "Quantité" (format: 2.000,00 → convertir en 2000)
7. Fichier: Nom du fichier PDF

RÈGLES SPÉCIFIQUES:
- [CRITIQUE] Format quantité spécifique: utilise le point comme séparateur de milliers (2.000,00 = deux mille)
- Le code article contient des dimensions (DR C4 180X320...)
- "Commande SOLDEE" indique la fin du tableau (arrêter l'extraction après)
- Vérifier le montant total HT (8.160,00) pour validation croisée
- Présence d'un numéro de commande client: E2025000063

POST-TRAITEMENT:
- Convertir 2.000,00 → 2000 (supprimer points, virgule = décimale si présente mais généralement entier)
- Nettoyer les espaces en début/fin de désignation

EXEMPLE DE SORTIE ATTENDUE:
```json
{
  "fournisseur": "TISSUS GISELE",
  "date": "16/01/2025",
  "reference_article": "DR C4 180X320 PC BLANCL2E",
  "article": "DRAPS PLATS OURLETS PIQUES DE 4CM AU FIL",
  "composants": "BLANC AU POINT DE CHAINETTE. 180,0 x 320,0 BLANC UNI POLYESTER-COTON 50/50 LIS 2FILS JAUNE",
  "quantite": 2000,
  "fichier": "Tissus Gisèle_Facture du 16 01 25.pdf"
}
```

---

## Matrice de Décision Automatique

Pour sélectionner le bon prompt selon le fichier:

```python
def detect_supplier(text_content):
    text_upper = text_content.upper()
    
    if "CLORO" in text_upper or "CLOROFIL" in text_upper:
        return "clorofil"
    elif "CTM STYLE" in text_upper:
        return "ctm_style"
    elif "HALBOUT" in text_upper:
        return "halbout"
    elif "MULLIEZ" in text_upper:
        return "mulliez_flory"
    elif "POYET" in text_upper:
        return "poyet_motte"
    elif "TISSUS GISELE" in text_upper or "TISSUS GISÈLE" in text_upper:
        return "tissus_gisele"
    else:
        return "unknown"
```

---

## Checklist de Validation Finale

Avant de valider l'extraction d'une facture:
- [ ] Le fournisseur est correctement identifié
- [ ] La date est au format JJ/MM/AAAA
- [ ] Les quantités sont des entiers (pas de décimales sauf si pertinent)
- [ ] Aucune ligne "Total", "Sous-total", "Prix au 100" n'est présente
- [ ] La somme des quantités × PU ≈ Total HT facture (marge 1%)
- [ ] Le nom du fichier source est conservé exactement
- [ ] Les références articles ne contiennent pas d'espaces superflus
- [ ] Pour Chorus (CTM/Poyet): la page 2 a été scannée
- [ ] Pour Mulliez: les coloris sont séparés en lignes distinctes
