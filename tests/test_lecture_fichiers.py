"""
Tests pour les fonctions de lecture des fichiers Excel
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# Ajouter le répertoire parent au path pour importer analyse
sys.path.insert(0, str(Path(__file__).parent.parent))
from analyse import lire_export_sap, lire_forecast, lire_planned_value, lire_valeur_acquise


class TestLectureExport:
    """Tests pour la lecture du fichier export SAP"""

    def test_lire_export_sap_fichier_existant(self, tmp_path):
        """Test de lecture d'un fichier SAP valide"""
        # Créer un fichier Excel temporaire
        fichier = tmp_path / "export_test.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15), datetime(2025, 2, 20)],
                "Val./Devise objet": [15000.50, 25000.75],
                "Autre colonne": ["A", "B"],
            }
        )
        df.to_excel(fichier, index=False)

        # Tester la lecture
        result = lire_export_sap(str(fichier))

        assert result is not None
        assert len(result) == 2
        assert "Date de la pièce" in result.columns
        assert "Val./Devise objet" in result.columns
        assert result["Val./Devise objet"].sum() == 40001.25

    def test_lire_export_sap_fichier_inexistant(self):
        """Test avec un fichier qui n'existe pas"""
        result = lire_export_sap("fichier_inexistant.xlsx")
        assert result is None

    def test_lire_export_sap_colonnes_manquantes(self, tmp_path):
        """Test avec des colonnes manquantes"""
        fichier = tmp_path / "export_invalide.xlsx"
        df = pd.DataFrame({"Mauvaise colonne": [1, 2, 3], "Autre colonne": ["A", "B", "C"]})
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        # Le fichier est lu mais les colonnes seront vérifiées plus tard dans main()
        assert result is not None

    def test_lire_export_sap_exception_generale(self):
        """Test d'une exception générale lors de la lecture"""
        # Passer un chemin invalide (pas un fichier Excel)
        result = lire_export_sap("/dev/null")
        assert result is None


class TestLecturePlannedValue:
    """Tests pour la lecture du fichier Planned Value"""

    def test_lire_pv_valide(self, tmp_path):
        """Test de lecture d'un fichier PV valide"""
        fichier = tmp_path / "pv_test.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["RCD", "J1", "J2"],
                "Durée (mois)": [3, 5, 4],
                "Date": [datetime(2025, 3, 31), datetime(2025, 8, 31), datetime(2025, 12, 31)],
                "Montant planifié": [100000, 200000, 250000],
                "Cumul planifié": [100000, 300000, 550000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_planned_value(str(fichier))

        assert result is not None
        assert len(result) == 3
        assert list(result["Jalon"]) == ["RCD", "J1", "J2"]
        assert result["Montant planifié"].sum() == 550000

    def test_lire_pv_fichier_inexistant(self):
        """Test avec un fichier PV inexistant"""
        result = lire_planned_value("pv_inexistant.xlsx")
        assert result is None

    def test_lire_pv_exception_generale(self):
        """Test d'une exception générale lors de la lecture PV"""
        result = lire_planned_value("/dev/null")
        assert result is None


class TestLectureValeurAcquise:
    """Tests pour la lecture du fichier Valeur Acquise"""

    def test_lire_va_valide(self, tmp_path):
        """Test de lecture d'un fichier VA valide"""
        fichier = tmp_path / "va_test.xlsx"
        data = {
            "Jalon": ["RCD", "J1"],
            "Date": [datetime(2025, 3, 31), datetime(2025, 8, 31)],
            "Montant planifié": [100000, 200000],
            datetime(2025, 1, 1): [0.2, 0.0],
            datetime(2025, 2, 1): [0.5, 0.0],
        }
        df = pd.DataFrame(data)
        df.to_excel(fichier, index=False)

        result = lire_valeur_acquise(str(fichier))

        assert result is not None
        assert len(result) == 2
        assert "Jalon" in result.columns

    def test_lire_va_fichier_inexistant(self):
        """Test avec un fichier VA inexistant"""
        result = lire_valeur_acquise("va_inexistant.xlsx")
        assert result is None

    def test_lire_va_exception_generale(self):
        """Test d'une exception générale lors de la lecture VA"""
        result = lire_valeur_acquise("/dev/null")
        assert result is None


class TestLectureForecast:
    """Tests pour la lecture du fichier Forecast"""

    def test_lire_forecast_valide(self, tmp_path):
        """Test de lecture d'un fichier forecast valide"""
        fichier = tmp_path / "forecast_test.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J3", "J4"],
                "Date projetée": [datetime(2026, 7, 31), datetime(2026, 10, 31)],
                "EAC (€)": [420000, 230000],
                "ETC (€)": [420000, 230000],
                "Commentaire": ["Test 1", "Test 2"],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_forecast(str(fichier))

        assert result is not None
        assert len(result) == 2
        assert "Jalon" in result.columns
        assert "EAC (€)" in result.columns

    def test_lire_forecast_fichier_inexistant(self):
        """Test avec un fichier forecast inexistant"""
        result = lire_forecast("forecast_inexistant.xlsx")
        assert result is None

    def test_lire_forecast_exception_generale(self):
        """Test d'une exception générale lors de la lecture forecast"""
        result = lire_forecast("/dev/null")
        assert result is None
