"""
Configuration pytest et fixtures communes
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

# Ajouter le répertoire parent au path
sys.path.insert(0, str(Path(__file__).parent.parent))


@pytest.fixture
def sample_export_sap() -> pd.DataFrame:
    """Fixture pour un export SAP exemple"""
    return pd.DataFrame(
        {
            "Date de la pièce": [datetime(2025, 1, 15), datetime(2025, 2, 20), datetime(2025, 3, 10)],
            "Val./Devise objet": [15000.50, 25000.75, 18500.00],
            "Document d'achat": [24184688.0, 24309176.0, 24346974.0],
        }
    )


@pytest.fixture
def sample_pv() -> pd.DataFrame:
    """Fixture pour un fichier PV exemple"""
    return pd.DataFrame(
        {
            "Jalon": ["RCD", "J1", "J2"],
            "Durée (mois)": [3, 5, 4],
            "Date": [datetime(2025, 3, 31), datetime(2025, 8, 31), datetime(2025, 12, 31)],
            "Montant planifié": [100000, 200000, 250000],
            "Cumul planifié": [100000, 300000, 550000],
        }
    )


@pytest.fixture
def sample_va() -> pd.DataFrame:
    """Fixture pour un fichier VA exemple"""
    data: dict[str | datetime, list] = {
        "Jalon": ["RCD", "J1", "J2"],
        "Date": [datetime(2025, 3, 31), datetime(2025, 8, 31), datetime(2025, 12, 31)],
        "Montant planifié": [100000, 200000, 250000],
        "Cumul planifié": [100000, 300000, 550000],
    }

    # Ajouter des colonnes mensuelles avec pourcentages
    for i in range(1, 13):
        mois_date = datetime(2025, i, 1)
        if i <= 3:
            data[mois_date] = [0.3 * i, 0.0, 0.0]  # RCD progresse
        elif i <= 8:
            data[mois_date] = [1.0, 0.2 * (i - 3), 0.0]  # RCD fini, J1 progresse
        else:
            data[mois_date] = [1.0, 1.0, 0.25 * (i - 8)]  # RCD et J1 finis, J2 progresse

    return pd.DataFrame(data)


@pytest.fixture
def sample_forecast() -> pd.DataFrame:
    """Fixture pour un fichier forecast exemple"""
    return pd.DataFrame(
        {
            "Jalon": ["J3", "J4", "POV"],
            "Date projetée": [datetime(2026, 7, 31), datetime(2026, 10, 31), datetime(2027, 6, 30)],
            "EAC (€)": [420000, 230000, 20000],
            "ETC (€)": [420000, 230000, 20000],
            "Commentaire": ["Test 1", "Test 2", "Test 3"],
        }
    )


@pytest.fixture
def sample_depenses_cumulees() -> pd.Series:
    """Fixture pour des dépenses cumulées exemple"""
    return pd.Series([50000, 100000, 180000, 250000], index=pd.period_range("2025-01", periods=4, freq="M"))


@pytest.fixture
def sample_ev_cumulee() -> pd.Series:
    """Fixture pour une EV cumulée exemple"""
    return pd.Series([30000, 70000, 120000, 180000], index=pd.period_range("2025-01", periods=4, freq="M"))


@pytest.fixture
def sample_pv_cumulee() -> pd.Series:
    """Fixture pour une PV cumulée exemple"""
    return pd.Series([40000, 90000, 140000, 200000], index=pd.period_range("2025-01", periods=4, freq="M"))
