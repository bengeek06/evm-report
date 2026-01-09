"""
Tests pour les calculs EVM (AC, PV, EV)
"""
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))
from analyse import calculer_depenses_cumulees, calculer_earned_value, traiter_planned_value


class TestCalculDepensesCumulees:
    """Tests pour le calcul des dépenses cumulées (AC)"""

    def test_depenses_cumulees_simple(self):
        """Test avec des dépenses simples"""
        df = pd.DataFrame({
            "Date de la pièce": [
                datetime(2025, 1, 15),
                datetime(2025, 1, 20),
                datetime(2025, 2, 10)
            ],
            "Val./Devise objet": [10000, 5000, 8000]
        })

        result = calculer_depenses_cumulees(df, "Date de la pièce", "Val./Devise objet")

        assert result is not None
        assert len(result) == 2  # 2 mois
        assert result.iloc[0] == 15000  # Janvier
        assert result.iloc[1] == 23000  # Février (cumul)

    def test_depenses_cumulees_meme_mois(self):
        """Test avec plusieurs dépenses le même mois"""
        df = pd.DataFrame({
            "Date de la pièce": [
                datetime(2025, 3, 5),
                datetime(2025, 3, 15),
                datetime(2025, 3, 25)
            ],
            "Val./Devise objet": [1000, 2000, 3000]
        })

        result = calculer_depenses_cumulees(df, "Date de la pièce", "Val./Devise objet")

        assert len(result) == 1
        assert result.iloc[0] == 6000

    def test_depenses_cumulees_vide(self):
        """Test avec un DataFrame vide"""
        df = pd.DataFrame({
            "Date de la pièce": [],
            "Val./Devise objet": []
        })

        result = calculer_depenses_cumulees(df, "Date de la pièce", "Val./Devise objet")

        assert result is not None
        assert len(result) == 0


class TestTraiterPlannedValue:
    """Tests pour le traitement de la Planned Value avec interpolation"""

    def test_pv_avec_interpolation(self):
        """Test de l'interpolation linéaire de la PV"""
        df_pv = pd.DataFrame({
            "Jalon": ["RCD", "J1"],
            "Date": [datetime(2025, 3, 31), datetime(2025, 6, 30)],
            "Montant planifié": [100000, 200000],
            "Cumul planifié": [100000, 300000]
        })

        result = traiter_planned_value(df_pv)

        assert result is not None
        pv_cumulee, jalons = result

        # Vérifier que l'interpolation a créé des valeurs pour les mois intermédiaires
        assert len(pv_cumulee) > 2
        # La PV doit être croissante (utiliser iloc pour éviter FutureWarning)
        assert all(pv_cumulee.iloc[i] <= pv_cumulee.iloc[i+1] for i in range(len(pv_cumulee)-1))
        # Les jalons doivent être présents
        assert jalons is not None

    def test_pv_sans_donnees(self):
        """Test avec un DataFrame PV vide"""
        df_pv = pd.DataFrame({
            "Jalon": [],
            "Date": [],
            "Montant planifié": [],
            "Cumul planifié": []
        })

        result = traiter_planned_value(df_pv)
        # La fonction retourne un tuple vide avec des Series/dict vides
        assert result is not None
        pv, jalons = result
        assert len(pv) == 0
        assert len(jalons) == 0


class TestCalculerEarnedValue:
    """Tests pour le calcul de l'Earned Value"""

    def test_ev_simple(self):
        """Test du calcul de l'EV avec des pourcentages simples"""
        df_pv = pd.DataFrame({
            "Jalon": ["RCD", "J1"],
            "Date": [datetime(2025, 3, 31), datetime(2025, 8, 31)],
            "Montant planifié": [100000, 200000],
            "Cumul planifié": [100000, 300000]
        })

        df_va = pd.DataFrame({
            "Jalon": ["RCD", "J1"],
            "Date": [datetime(2025, 3, 31), datetime(2025, 8, 31)],
            "Montant planifié": [100000, 200000],
            datetime(2025, 1, 1): [0.0, 0.0],
            datetime(2025, 2, 1): [0.5, 0.0],  # RCD à 50%
            datetime(2025, 3, 1): [1.0, 0.0],  # RCD à 100%
            datetime(2025, 4, 1): [1.0, 0.2],  # RCD 100%, J1 20%
        })

        result = calculer_earned_value(df_pv, df_va)

        assert result is not None
        assert len(result) > 0

        # Vérifier que les valeurs sont cohérentes
        # EV de février devrait être 50% de 100000 = 50000
        ev_fev = result[result.index.month == 2].iloc[0] if any(result.index.month == 2) else 0
        assert ev_fev == pytest.approx(50000, rel=0.01)

        # EV de mars devrait être 100% de 100000 = 100000
        ev_mars = result[result.index.month == 3].iloc[0] if any(result.index.month == 3) else 0
        assert ev_mars == pytest.approx(100000, rel=0.01)

    def test_ev_pourcentages_invalides(self):
        """Test avec des pourcentages > 1.0"""
        df_pv = pd.DataFrame({
            "Jalon": ["RCD"],
            "Date": [datetime(2025, 3, 31)],
            "Montant planifié": [100000],
            "Cumul planifié": [100000]
        })

        df_va = pd.DataFrame({
            "Jalon": ["RCD"],
            "Date": [datetime(2025, 3, 31)],
            "Montant planifié": [100000],
            datetime(2025, 1, 1): [0.5],
            datetime(2025, 2, 1): [1.5],  # Plus de 100% - invalide mais géré
        })

        result = calculer_earned_value(df_pv, df_va)

        # Le code devrait gérer les valeurs > 1.0 en les limitant
        assert result is not None

    def test_ev_sans_va(self):
        """Test sans fichier VA"""
        df_pv = pd.DataFrame({
            "Jalon": ["RCD"],
            "Date": [datetime(2025, 3, 31)],
            "Montant planifié": [100000],
            "Cumul planifié": [100000]
        })

        result = calculer_earned_value(df_pv, None)
        assert result is None
