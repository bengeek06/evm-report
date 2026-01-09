"""
Tests pour les calculs de projections EAC
"""
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))
from analyse import calculer_eac_projete, calculer_projections_automatiques


class TestProjectionsAutomatiques:
    """Tests pour les projections automatiques (CPI, CPI×SPI, Reste à Plan)"""

    def test_projections_avec_bonnes_performances(self):
        """Test avec CPI et SPI > 1 (bonnes performances)"""
        # AC: 100k dépensés
        depenses = pd.Series(
            [50000, 100000],
            index=pd.period_range('2025-01', periods=2, freq='M')
        )

        # EV: 120k de valeur acquise (bon rendement)
        ev = pd.Series(
            [60000, 120000],
            index=pd.period_range('2025-01', periods=2, freq='M')
        )

        # PV: 110k prévu, BAC = 200k
        pv = pd.Series(
            [55000, 110000, 150000, 200000],
            index=pd.period_range('2025-01', periods=4, freq='M')
        )

        df_pv = pd.DataFrame({
            "Jalon": ["J1", "J2"],
            "Date": [datetime(2025, 3, 31), datetime(2025, 4, 30)],
            "Montant planifié": [150000, 50000],
            "Cumul planifié": [150000, 200000]
        })

        result = calculer_projections_automatiques(depenses, ev, pv, df_pv)

        assert result is not None
        projections_data, series_projections, date_fin_proj = result

        # Vérifier que les 3 méthodes sont présentes
        assert 'CPI' in projections_data
        assert 'CPI_SPI' in projections_data
        assert 'RESTE_PLAN' in projections_data

        # CPI = EV/AC = 120k/100k = 1.2 > 1 (bon)
        # EAC_CPI = BAC/CPI = 200k/1.2 = 166.67k (moins que BAC)
        eac_cpi = projections_data['CPI']['eac']
        assert eac_cpi < 200000  # Sous budget
        assert eac_cpi > 100000  # Mais plus que déjà dépensé

    def test_projections_avec_mauvaises_performances(self):
        """Test avec CPI et SPI < 1 (mauvaises performances)"""
        # AC: 200k dépensés
        depenses = pd.Series(
            [100000, 200000],
            index=pd.period_range('2025-01', periods=2, freq='M')
        )

        # EV: 100k de valeur acquise (mauvais rendement)
        ev = pd.Series(
            [50000, 100000],
            index=pd.period_range('2025-01', periods=2, freq='M')
        )

        # PV: 150k prévu, BAC = 300k
        pv = pd.Series(
            [75000, 150000, 225000, 300000],
            index=pd.period_range('2025-01', periods=4, freq='M')
        )

        df_pv = pd.DataFrame({
            "Jalon": ["J1", "J2"],
            "Date": [datetime(2025, 3, 31), datetime(2025, 4, 30)],
            "Montant planifié": [225000, 75000],
            "Cumul planifié": [225000, 300000]
        })

        result = calculer_projections_automatiques(depenses, ev, pv, df_pv)

        assert result is not None
        projections_data, series_projections, date_fin_proj = result

        # CPI = EV/AC = 100k/200k = 0.5 < 1 (mauvais)
        # EAC_CPI = BAC/CPI = 300k/0.5 = 600k (dépassement)
        eac_cpi = projections_data['CPI']['eac']
        assert eac_cpi > 300000  # Dépassement de budget

        # Reste à Plan = AC + (BAC - EV) = 200k + (300k - 100k) = 400k
        eac_reste = projections_data['RESTE_PLAN']['eac']
        assert eac_reste == pytest.approx(400000, rel=0.01)

    def test_projections_series_temporelles(self):
        """Test que les séries temporelles sont générées"""
        depenses = pd.Series(
            [100000],
            index=pd.period_range('2025-01', periods=1, freq='M')
        )

        ev = pd.Series(
            [50000],
            index=pd.period_range('2025-01', periods=1, freq='M')
        )

        pv = pd.Series(
            [100000, 200000],
            index=pd.period_range('2025-01', periods=2, freq='M')
        )

        df_pv = pd.DataFrame({
            "Jalon": ["J1"],
            "Date": [datetime(2025, 2, 28)],
            "Montant planifié": [200000],
            "Cumul planifié": [200000]
        })

        result = calculer_projections_automatiques(depenses, ev, pv, df_pv)

        assert result is not None
        projections_data, series_projections, date_fin_proj = result

        # Vérifier que les séries temporelles existent
        assert 'CPI' in series_projections
        assert 'CPI_SPI' in series_projections
        assert 'RESTE_PLAN' in series_projections

        # Chaque série doit avoir au moins la valeur actuelle
        for serie in series_projections.values():
            assert len(serie) > 0

    def test_projections_sans_ev(self):
        """Test sans Earned Value"""
        depenses = pd.Series([100000], index=pd.period_range('2025-01', periods=1, freq='M'))
        pv = pd.Series([100000], index=pd.period_range('2025-01', periods=1, freq='M'))
        df_pv = pd.DataFrame({"Jalon": ["J1"], "Date": [datetime(2025, 2, 28)]})

        result = calculer_projections_automatiques(depenses, None, pv, df_pv)
        assert result is None


class TestCalculEACProjete:
    """Tests pour le calcul de l'EAC avec forecast manuel"""

    def test_eac_projete_simple(self):
        """Test avec un fichier forecast simple"""
        ev = pd.Series(
            [50000, 100000],
            index=pd.period_range('2025-01', periods=2, freq='M')
        )

        df_forecast = pd.DataFrame({
            "Jalon": ["J3", "J4"],
            "Date projetée": [datetime(2026, 7, 31), datetime(2026, 10, 31)],
            "EAC (€)": [420000, 230000]
        })

        df_pv = pd.DataFrame({
            "Jalon": ["J1", "J2", "J3", "J4"],
            "Date": [
                datetime(2025, 3, 31),
                datetime(2025, 6, 30),
                datetime(2026, 7, 31),
                datetime(2026, 10, 31)
            ],
            "Montant planifié": [100000, 200000, 300000, 150000]
        })

        result = calculer_eac_projete(ev, df_forecast, df_pv)

        assert result is not None
        eac_series, jalons_forecast = result

        # Vérifier que la série EAC est créée
        assert len(eac_series) > 0

        # Vérifier que les jalons forecast sont identifiés
        assert "J3" in jalons_forecast
        assert "J4" in jalons_forecast

    def test_eac_projete_sans_forecast(self):
        """Test sans fichier forecast"""
        ev = pd.Series([100000], index=pd.period_range('2025-01', periods=1, freq='M'))
        df_pv = pd.DataFrame({"Jalon": ["J1"], "Date": [datetime(2025, 3, 31)]})

        result = calculer_eac_projete(ev, None, df_pv)
        assert result is None
