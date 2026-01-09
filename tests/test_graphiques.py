"""
Tests pour les fonctions de génération de graphiques
"""

import sys
from collections.abc import Hashable
from pathlib import Path
from unittest.mock import patch

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent))
from analyse import tracer_courbe, tracer_courbe_projections


class TestTracerCourbe:
    """Tests pour la fonction tracer_courbe"""

    def test_tracer_courbe_ac_seulement(self, tmp_path: Path) -> None:
        """Test avec seulement les dépenses réelles (AC)"""
        fichier = tmp_path / "graphique.png"

        # Données de test
        depenses = pd.Series(
            [50000, 100000, 150000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        with patch("matplotlib.pyplot.savefig") as mock_save, patch("matplotlib.pyplot.close"):
            tracer_courbe(depenses, fichier_sortie=str(fichier))
            mock_save.assert_called_once()

    def test_tracer_courbe_ac_pv(self, tmp_path: Path) -> None:
        """Test avec AC et PV"""
        fichier = tmp_path / "graphique.png"

        depenses = pd.Series(
            [50000, 100000, 150000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        pv = pd.Series(
            [60000, 120000, 180000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        jalons: dict[Hashable, list[str]] = {
            pd.Period("2025-01", freq="M"): ["J1"],
            pd.Period("2025-03", freq="M"): ["J2"],
        }

        with patch("matplotlib.pyplot.savefig") as mock_save, patch("matplotlib.pyplot.close"):
            tracer_courbe(depenses, pv, jalons, fichier_sortie=str(fichier))
            mock_save.assert_called_once()

    def test_tracer_courbe_ac_pv_ev(self, tmp_path: Path) -> None:
        """Test avec AC, PV et EV"""
        fichier = tmp_path / "graphique.png"

        depenses = pd.Series(
            [50000, 100000, 150000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        pv = pd.Series(
            [60000, 120000, 180000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        ev = pd.Series(
            [55000, 110000, 165000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        with patch("matplotlib.pyplot.savefig") as mock_save, patch("matplotlib.pyplot.close"):
            tracer_courbe(depenses, pv, None, ev, fichier_sortie=str(fichier))
            mock_save.assert_called_once()

    def test_tracer_courbe_toutes_donnees(self, tmp_path: Path) -> None:
        """Test avec toutes les données (AC, PV, EV, jalons, EAC)"""
        fichier = tmp_path / "graphique_complet.png"

        depenses = pd.Series(
            [50000, 100000, 150000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        pv = pd.Series(
            [60000, 120000, 180000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        ev = pd.Series(
            [55000, 110000, 165000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        eac = pd.Series(
            [200000, 200000, 200000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        jalons: dict[Hashable, list[str]] = {
            pd.Period("2025-01", freq="M"): ["Jalon 1", "Jalon 2"],
            pd.Period("2025-03", freq="M"): ["Jalon 3"],
        }

        with patch("matplotlib.pyplot.savefig") as mock_save, patch("matplotlib.pyplot.close"):
            tracer_courbe(depenses, pv, jalons, ev, eac_projete=eac, fichier_sortie=str(fichier))
            mock_save.assert_called_once()


class TestTracerProjections:
    """Tests pour la fonction tracer_courbe_projections"""

    def test_tracer_projections_base(self, tmp_path: Path) -> None:
        """Test de base avec des projections"""
        fichier = tmp_path / "projections.png"

        depenses = pd.Series(
            [50000, 100000, 150000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        ev = pd.Series(
            [55000, 110000, 165000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        projections = {
            "cpi": {
                "series": pd.Series(
                    [200000, 250000],
                    index=pd.period_range(start="2025-04", periods=2, freq="M"),
                ),
                "eac": 250000,
                "date": pd.Period("2025-05", freq="M"),
            }
        }

        with patch("matplotlib.pyplot.savefig") as mock_save, patch("matplotlib.pyplot.close"):
            tracer_courbe_projections(depenses, ev, projections, str(fichier))
            mock_save.assert_called_once()

    def test_tracer_projections_toutes_donnees(self, tmp_path: Path) -> None:
        """Test avec tous les scénarios"""
        fichier = tmp_path / "projections_complet.png"

        depenses = pd.Series(
            [50000, 100000, 150000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        ev = pd.Series(
            [55000, 110000, 165000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        projections = {
            "cpi": {
                "series": pd.Series(
                    [200000, 250000, 300000],
                    index=pd.period_range(start="2025-04", periods=3, freq="M"),
                ),
                "eac": 300000,
                "date": pd.Period("2025-06", freq="M"),
            },
            "cpi_spi": {
                "series": pd.Series(
                    [210000, 260000, 310000],
                    index=pd.period_range(start="2025-04", periods=3, freq="M"),
                ),
                "eac": 310000,
                "date": pd.Period("2025-06", freq="M"),
            },
            "reste_plan": {
                "series": pd.Series(
                    [190000, 240000, 290000],
                    index=pd.period_range(start="2025-04", periods=3, freq="M"),
                ),
                "eac": 290000,
                "date": pd.Period("2025-06", freq="M"),
            },
            "forecast": {
                "series": pd.Series(
                    [220000, 270000, 320000],
                    index=pd.period_range(start="2025-04", periods=3, freq="M"),
                ),
                "eac": 320000,
                "date": pd.Period("2025-06", freq="M"),
            },
        }

        with patch("matplotlib.pyplot.savefig") as mock_save, patch("matplotlib.pyplot.close"):
            tracer_courbe_projections(depenses, ev, projections, str(fichier))
            mock_save.assert_called_once()


class TestGraphiquesSimplifies:
    """Tests simplifiés pour couvrir les fonctions de graphiques"""

    def test_tracer_courbe_minimal(self, tmp_path):
        """Test minimal avec mock complet"""
        fichier = tmp_path / "minimal.png"
        depenses = pd.Series([50000], index=pd.period_range(start="2025-01", periods=1, freq="M"))

        with (
            patch("matplotlib.pyplot.savefig"),
            patch("matplotlib.pyplot.close"),
            patch("matplotlib.pyplot.figure"),
            patch("matplotlib.pyplot.plot"),
            patch("matplotlib.pyplot.xlabel"),
            patch("matplotlib.pyplot.ylabel"),
            patch("matplotlib.pyplot.title"),
            patch("matplotlib.pyplot.legend"),
            patch("matplotlib.pyplot.grid"),
            patch("matplotlib.pyplot.tight_layout"),
        ):
            tracer_courbe(depenses, fichier_sortie=str(fichier))

    def test_tracer_projections_minimal(self, tmp_path):
        """Test minimal des projections avec mock complet"""
        fichier = tmp_path / "projections_minimal.png"
        depenses = pd.Series([50000], index=pd.period_range(start="2025-01", periods=1, freq="M"))
        projections = {}

        with (
            patch("matplotlib.pyplot.savefig"),
            patch("matplotlib.pyplot.close"),
            patch("matplotlib.pyplot.figure"),
            patch("matplotlib.pyplot.plot"),
            patch("matplotlib.pyplot.axhline"),
            patch("matplotlib.pyplot.xlabel"),
            patch("matplotlib.pyplot.ylabel"),
            patch("matplotlib.pyplot.title"),
            patch("matplotlib.pyplot.legend"),
            patch("matplotlib.pyplot.grid"),
            patch("matplotlib.pyplot.tight_layout"),
        ):
            tracer_courbe_projections(depenses, None, projections, str(fichier))
