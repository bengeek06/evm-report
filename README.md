# Analyse EVM (Earned Value Management)

Outil d'analyse de la performance de projets selon la méthodologie EVM (Earned Value Management).

## Fonctionnalités principales

### 📊 Calcul automatique des indicateurs EVM
- **AC (Actual Cost)** : Dépenses réelles issues de l'export SAP
- **PV (Planned Value)** : Budget prévu avec interpolation linéaire entre jalons
- **EV (Earned Value)** : Valeur acquise calculée depuis les pourcentages d'avancement
- **CV (Cost Variance)** : Écart de coût (EV - AC)
- **SV (Schedule Variance)** : Écart de délai (EV - PV)
- **CPI (Cost Performance Index)** : Indice de performance des coûts (EV / AC)
- **SPI (Schedule Performance Index)** : Indice de performance des délais (EV / PV)

### 🔮 Projections automatiques à terminaison
Le script calcule **3 scénarios de projection** automatiquement :

1. **Méthode CPI (Optimiste)** : `BAC / CPI`
   - La performance des coûts actuelle se poursuit

2. **Méthode CPI×SPI (Réaliste)** : `AC + [(BAC-EV) / (CPI×SPI)]`
   - Performance coûts ET délais se poursuit

3. **Méthode Reste à Plan (Pessimiste)** : `AC + (BAC-EV)`
   - Le reste se déroule comme prévu initialement

Un 4ème scénario peut être ajouté via un fichier `forecast.xlsx` (projection manuelle).

### 📈 Visualisation en deux graphiques

**Graphique 1 : Réalisé à date**
- Courbes AC, PV, EV
- Variances CV et SV
- Jalons planifiés

**Graphique 2 : Projections**
- Historique AC et EV
- 3-4 scénarios de projection EAC
- Dates de fin estimées

### 📄 Rapport Word complet

Génération automatique d'un rapport professionnel contenant :
1. **Définitions** : Tous les termes et formules EVM
2. **Réalisé** : Tableau de valeurs, graphique, indicateurs actuels avec interprétation
3. **Projections** : Tableau comparatif des scénarios, graphique, analyse de l'incertitude
4. **Conclusion** : Synthèse de la performance et recommandations automatiques

## Installation

```bash
# Cloner le projet
cd /chemin/vers/le/projet

# Créer un environnement virtuel
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# ou
.venv\Scripts\activate     # Windows

# Installer le projet et ses dépendances
pip install -e .

# Ou installer uniquement les dépendances de développement
pip install -e ".[dev]"
```

## Structure des fichiers d'entrée

### 1. Export SAP (EXPORT.XLSX)
Fichier Excel avec au minimum les colonnes :
- `Date de la pièce` : Date de la dépense
- `Val./Devise objet` : Montant de la dépense en euros

### 2. Planned Value (pv.xlsx)
Fichier Excel avec les colonnes :
- `Jalon` : Nom du jalon (ex: RCD, J1, J2, ...)
- `Date` : Date prévue du jalon
- `Montant planifié` : Budget alloué au jalon (€)
- `Cumul planifié` : Budget cumulé (€)

**Note** : Le script interpole linéairement la PV entre les jalons pour avoir une valeur pour chaque mois.

### 3. Valeur Acquise (va.xlsx)
Fichier Excel avec les colonnes :
- `Jalon` : Nom du jalon
- `Date` : Date prévue
- `Montant planifié` : Même que pv.xlsx
- Colonnes mensuelles avec dates : Pourcentages d'avancement (0.0 à 1.0)

**Important** : Les pourcentages doivent être en décimal (ex: 0.5 pour 50%, pas 50).

### 4. Forecast (forecast.xlsx) - OPTIONNEL
Fichier Excel avec les colonnes :
- `Jalon` : Nom du jalon
- `Date projetée` : Date estimée de réalisation
- `EAC (€)` : Coût estimé pour ce jalon
- `ETC (€)` : Reste à dépenser
- `Commentaire` : Notes

Ce fichier est **optionnel**. S'il n'est pas fourni, seules les 3 méthodes automatiques seront utilisées.

## Utilisation

### Utilisation basique

```bash
# Analyse complète avec rapport Word (utilise tous les fichiers par défaut)
python analyse.py --word rapport_evm.docx

# Avec fichiers spécifiques
python analyse.py \
  --sap mes_depenses.xlsx \
  --pv mon_pv.xlsx \
  --va mon_va.xlsx \
  --word rapport_complet.docx
```

### Sans fichier forecast (3 scénarios automatiques)

```bash
python analyse.py \
  --sap EXPORT.XLSX \
  --pv pv.xlsx \
  --va va.xlsx \
  --forecast non_existant.xlsx \
  --word rapport_auto.docx
```

### Génération des graphiques seulement (sans Word)

```bash
python analyse.py \
  --sap EXPORT.XLSX \
  --pv pv.xlsx \
  --va va.xlsx \
  --output mes_graphiques.png
```

Cela génère :
- `mes_graphiques_realise.png` : Graphique du réalisé
- `mes_graphiques_projections.png` : Graphique des projections
- `tableau_evm.csv` et `tableau_evm.xlsx` : Tableaux de données

### Options disponibles

```
--sap SAP            Fichier Excel export SAP (défaut: EXPORT.XLSX)
--pv PV              Fichier Planned Value (défaut: pv.xlsx)
--va VA              Fichier Valeur Acquise (défaut: va.xlsx)
--forecast FORECAST  Fichier projections (défaut: forecast.xlsx)
--output OUTPUT      Nom de base pour graphiques (défaut: analyse_evm.png)
--tableau TABLEAU    Nom de base pour tableaux (défaut: tableau_evm)
--word WORD          Générer rapport Word (ex: rapport_evm.docx)
--help               Afficher l'aide
--version            Afficher la version
```

## Fichiers de sortie

### Avec option --word

Le script génère uniquement le rapport Word complet :
- `rapport_evm.docx` : Rapport Word avec tout (définitions, tableaux, 2 graphiques, analyses, recommandations)

Les fichiers intermédiaires (PNG, CSV, XLSX) sont automatiquement supprimés.

### Sans option --word

Le script génère :
- `analyse_evm_realise.png` : Graphique du réalisé (AC, PV, EV, CV, SV)
- `analyse_evm_projections.png` : Graphique des projections (3-4 scénarios EAC)
- `tableau_evm.csv` : Tableau des valeurs au format CSV
- `tableau_evm.xlsx` : Tableau des valeurs au format Excel

## Interprétation des résultats

### Indicateurs de performance actuels

| Indicateur | Formule | Interprétation |
|------------|---------|----------------|
| **CPI < 1** | EV / AC | ⚠️ Dépassement de coût (inefficacité) |
| **CPI = 1** | EV / AC | ✓ Performance des coûts conforme |
| **CPI > 1** | EV / AC | ✓ Sous budget (efficace) |
| **SPI < 1** | EV / PV | ⚠️ Retard sur le planning |
| **SPI = 1** | EV / PV | ✓ Performance des délais conforme |
| **SPI > 1** | EV / PV | ✓ En avance sur le planning |

### Scénarios de projection

Les 3-4 scénarios permettent d'avoir une **fourchette d'incertitude** :

- **Optimiste (CPI)** : Si seule la performance des coûts continue
- **Réaliste (CPI×SPI)** : Si coûts ET délais continuent (recommandé)
- **Pessimiste (Reste à Plan)** : Si le reste suit le plan initial malgré les problèmes actuels
- **Manuel (Forecast)** : Basé sur l'expertise métier et ajustements spécifiques

**Analyse** : Plus l'écart entre optimiste et pessimiste est grand, plus l'incertitude est élevée.

### Code couleur dans le rapport

- 🟢 **Vert** : Indicateur positif (économies, avance)
- 🔴 **Rouge** : Indicateur négatif (dépassement, retard)

## Exemples de situations

### Projet sain
```
CPI = 1.05  (5% sous budget)
SPI = 0.98  (2% de retard)
EAC Optimiste  = 1 450 k€
EAC Réaliste   = 1 480 k€
EAC Pessimiste = 1 510 k€
```
→ Performance bonne, faible incertitude

### Projet en difficulté
```
CPI = 0.55  (45% de dépassement)
SPI = 0.74  (26% de retard)
EAC Optimiste  = 2 741 k€
EAC Réaliste   = 3 385 k€
EAC Pessimiste = 1 916 k€
```
→ Performance préoccupante, forte incertitude, actions correctives urgentes

## Dépendances

Voir [pyproject.toml](pyproject.toml) :

**Dépendances principales :**
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- matplotlib >= 3.7.0
- python-docx >= 0.8.11

**Dépendances de développement :**
- pytest >= 7.4.0 (tests)
- pytest-cov >= 4.1.0 (tests)

## Tests

Le projet inclut une suite de tests unitaires complète avec pytest.

### Exécuter les tests

```bash
# Tous les tests
pytest -v

# Avec couverture de code
pytest --cov=analyse --cov-report=term-missing

# Rapport HTML de couverture
pytest --cov=analyse --cov-report=html
```

### Résultats des tests

- ✅ **27 tests unitaires** - Tous passent
- 📊 **Couverture** : 34% du code principal
- ⏱️ **Temps d'exécution** : ~1.5 secondes

Pour plus de détails, consultez :
- [tests/README.md](tests/README.md) - Documentation des tests
- [TESTS_SUMMARY.md](TESTS_SUMMARY.md) - Résumé complet des tests

## Changelog

Voir [CHANGELOG.md](CHANGELOG.md) pour l'historique des versions.

## Méthodologie EVM

L'Earned Value Management (EVM) est une méthodologie de gestion de projet reconnue internationalement qui permet de :
- Mesurer objectivement la performance d'un projet
- Identifier précocement les déviations
- Projeter le coût final et la date de fin
- Prendre des décisions basées sur des données factuelles

### Formules principales

```
CV (Cost Variance) = EV - AC
SV (Schedule Variance) = EV - PV
CPI (Cost Performance Index) = EV / AC
SPI (Schedule Performance Index) = EV / PV

EAC CPI = BAC / CPI
EAC CPI×SPI = AC + [(BAC - EV) / (CPI × SPI)]
EAC Reste à Plan = AC + (BAC - EV)
```

## Licence

Ce projet est sous licence MIT. Voir le fichier [LICENSE](LICENSE) pour plus de détails.
