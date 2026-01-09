# Résumé des Tests - Analyse EVM

## Vue d'ensemble

✅ **27 tests unitaires** - Tous passent
📊 **Couverture** : 34% du code (fonctions principales testées)
⏱️ **Temps d'exécution** : ~1.5 secondes

## Structure des tests

### 1. test_lecture_fichiers.py (9 tests)
Tests pour la lecture des fichiers Excel d'entrée :
- ✅ Lecture export SAP avec fichier valide
- ✅ Gestion des fichiers inexistants
- ✅ Validation des colonnes manquantes
- ✅ Lecture PV (Planned Value)
- ✅ Lecture VA (Valeur Acquise)
- ✅ Lecture forecast (projections manuelles)

### 2. test_calculs_evm.py (8 tests)
Tests pour les calculs EVM principaux :
- ✅ Calcul des dépenses cumulées (AC)
- ✅ Gestion des dépenses du même mois
- ✅ Traitement des DataFrames vides
- ✅ Interpolation linéaire de la PV entre jalons
- ✅ Calcul de l'EV avec pourcentages d'avancement
- ✅ Validation des pourcentages invalides
- ✅ Gestion des données manquantes

### 3. test_projections.py (6 tests)
Tests pour les calculs de projections :
- ✅ Projections automatiques (3 méthodes : CPI, CPI×SPI, Reste à Plan)
- ✅ Scénarios de bonnes performances (CPI>1, SPI>1)
- ✅ Scénarios de mauvaises performances (CPI<1, SPI<1)
- ✅ Génération des séries temporelles futures
- ✅ Calcul EAC avec forecast manuel
- ✅ Gestion des données manquantes

### 4. test_integration.py (4 tests)
Tests d'intégration du workflow complet :
- ✅ Workflow complet avec tous les fichiers
- ✅ Workflow minimal (SAP + PV + VA, sans forecast)
- ✅ Calcul des indicateurs CPI et SPI
- ✅ Calcul des variances CV et SV

## Fixtures pytest disponibles

7 fixtures réutilisables dans `conftest.py` :
- `sample_export_sap` : Données SAP avec 5 transactions
- `sample_pv` : Planned Value avec 3 jalons
- `sample_va` : Valeur Acquise avec 3 jalons et pourcentages
- `sample_forecast` : Forecast manuel avec 3 scénarios
- `sample_depenses_cumulees` : Série AC de 5 mois
- `sample_ev_cumulee` : Série EV de 5 mois
- `sample_pv_cumulee` : Série PV de 5 mois

## Couverture de code

### Zones testées (34% du code)
- ✅ Lecture des fichiers Excel
- ✅ Validation des colonnes
- ✅ Calcul AC (dépenses cumulées)
- ✅ Calcul PV avec interpolation
- ✅ Calcul EV (earned value)
- ✅ Calcul des projections (3 méthodes EAC)
- ✅ Workflow complet d'intégration

### Zones non testées (à améliorer)
- ❌ Génération des graphiques matplotlib
- ❌ Génération du rapport Word
- ❌ Fonctions de formatage et export
- ❌ Interface CLI (argparse)

## Commandes utiles

### Exécuter tous les tests
```bash
pytest -v
```

### Exécuter avec couverture
```bash
pytest --cov=analyse --cov-report=term-missing
```

### Exécuter avec rapport HTML
```bash
pytest --cov=analyse --cov-report=html
```

### Exécuter un fichier spécifique
```bash
pytest tests/test_calculs_evm.py -v
```

## Prochaines étapes

Pour augmenter la couverture de code, considérer :

1. **Tests des graphiques** :
   - Utiliser `unittest.mock` pour mocker matplotlib
   - Vérifier que les données sont correctement passées aux fonctions de plotting

2. **Tests du rapport Word** :
   - Créer des fichiers temporaires avec `tmp_path`
   - Vérifier la structure du document généré
   - Valider le contenu des sections

3. **Tests de l'interface CLI** :
   - Tester les arguments argparse
   - Valider les chemins de fichiers
   - Tester les différentes combinaisons d'options

4. **Tests de performance** :
   - Mesurer le temps d'exécution sur de gros fichiers
   - Tester avec des datasets de tailles variées

## Notes techniques

- Python 3.13.5
- pytest 9.0.2
- pytest-cov 7.0.0
- Tous les tests utilisent des fichiers temporaires
- Les fixtures sont partagées via `conftest.py`
- Pattern AAA (Arrange-Act-Assert) utilisé dans tous les tests
