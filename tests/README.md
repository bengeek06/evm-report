# Tests Unitaires

Ce répertoire contient les tests unitaires pour le projet d'analyse EVM (Earned Value Management).

## Structure

- `conftest.py` : Fixtures pytest réutilisables pour les données de test
- `test_lecture_fichiers.py` : Tests pour les fonctions de lecture des fichiers Excel
- `test_calculs_evm.py` : Tests pour les calculs EVM (AC, PV, EV)
- `test_projections.py` : Tests pour les calculs de projections et EAC
- `test_integration.py` : Tests d'intégration du workflow complet

## Exécuter les tests

### Tous les tests
```bash
pytest -v
```

### Tests d'un fichier spécifique
```bash
pytest tests/test_calculs_evm.py -v
```

### Tests d'une classe spécifique
```bash
pytest tests/test_calculs_evm.py::TestCalculDepensesCumulees -v
```

### Test unique
```bash
pytest tests/test_calculs_evm.py::TestCalculDepensesCumulees::test_depenses_cumulees_simple -v
```

## Couverture de code

### Rapport dans le terminal
```bash
pytest --cov=analyse --cov-report=term-missing
```

### Rapport HTML interactif
```bash
pytest --cov=analyse --cov-report=html
```
Le rapport sera généré dans `htmlcov/index.html`

## Fixtures disponibles

Les fixtures suivantes sont disponibles dans `conftest.py` :

- `sample_export_sap` : Données SAP d'exemple
- `sample_pv` : Données Planned Value d'exemple
- `sample_va` : Données Valeur Acquise d'exemple
- `sample_forecast` : Données de forecast d'exemple
- `sample_depenses_cumulees` : Série temporelle AC d'exemple
- `sample_ev_cumulee` : Série temporelle EV d'exemple
- `sample_pv_cumulee` : Série temporelle PV d'exemple

## Résultats actuels

- **27 tests** : Tous passent ✅
- **Couverture** : 34% (focus sur les fonctions de calcul)
- Les fonctions de génération de graphiques et rapports Word ne sont pas encore testées

## Ajouter de nouveaux tests

1. Créer un nouveau fichier `test_*.py` dans ce répertoire
2. Importer les fixtures nécessaires depuis `conftest.py`
3. Utiliser les classes de test avec pytest :

```python
import pytest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class TestMaFonction:
    """Tests pour ma_fonction"""
    
    def test_cas_nominal(self):
        """Test du cas nominal"""
        # Arrange
        donnees = ...
        
        # Act
        resultat = ma_fonction(donnees)
        
        # Assert
        assert resultat == valeur_attendue
```

## Notes

- Les tests utilisent `tmp_path` de pytest pour créer des fichiers temporaires
- Les fichiers temporaires sont automatiquement nettoyés après chaque test
- Les fixtures sont réutilisables entre les tests pour éviter la duplication
