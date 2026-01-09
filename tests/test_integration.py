"""
Tests d'intégration pour le workflow complet
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from analyse import (
    calculer_depenses_cumulees,
    calculer_earned_value,
    calculer_projections_automatiques,
    lire_export_sap,
    lire_planned_value,
    lire_valeur_acquise,
    traiter_planned_value,
)


class TestWorkflowComplet:
    """Tests d'intégration pour le workflow complet"""

    def test_workflow_avec_tous_fichiers(self, tmp_path, sample_export_sap, sample_pv, sample_va):
        """Test du workflow complet avec tous les fichiers"""

        # Créer les fichiers temporaires
        fichier_sap = tmp_path / "export.xlsx"
        fichier_pv = tmp_path / "pv.xlsx"
        fichier_va = tmp_path / "va.xlsx"

        sample_export_sap.to_excel(fichier_sap, index=False)
        sample_pv.to_excel(fichier_pv, index=False)
        sample_va.to_excel(fichier_va, index=False)

        # Workflow complet
        # 1. Lire l'export SAP
        df_sap = lire_export_sap(str(fichier_sap))
        assert df_sap is not None

        # 2. Calculer les dépenses cumulées
        ac = calculer_depenses_cumulees(df_sap, "Date de la pièce", "Val./Devise objet")
        assert ac is not None
        assert len(ac) > 0

        # 3. Lire et traiter la PV
        df_pv = lire_planned_value(str(fichier_pv))
        assert df_pv is not None

        result_pv = traiter_planned_value(df_pv)
        assert result_pv is not None
        pv, jalons = result_pv
        assert pv is not None
        assert jalons is not None

        # 4. Lire et calculer l'EV
        df_va = lire_valeur_acquise(str(fichier_va))
        assert df_va is not None

        ev = calculer_earned_value(df_pv, df_va)
        assert ev is not None
        assert len(ev) > 0

        # 5. Calculer les projections
        projections = calculer_projections_automatiques(ac, ev, pv, df_pv)
        assert projections is not None

        projections_data, _, _ = projections
        assert len(projections_data) == 3  # CPI, CPI_SPI, RESTE_PLAN
        assert all(k in projections_data for k in ['CPI', 'CPI_SPI', 'RESTE_PLAN'])

    def test_workflow_minimal_sans_forecast(self, tmp_path, sample_export_sap, sample_pv, sample_va):
        """Test du workflow minimal sans fichier forecast"""

        # Créer les fichiers temporaires
        fichier_sap = tmp_path / "export.xlsx"
        fichier_pv = tmp_path / "pv.xlsx"
        fichier_va = tmp_path / "va.xlsx"

        sample_export_sap.to_excel(fichier_sap, index=False)
        sample_pv.to_excel(fichier_pv, index=False)
        sample_va.to_excel(fichier_va, index=False)

        # Workflow sans forecast
        df_sap = lire_export_sap(str(fichier_sap))
        ac = calculer_depenses_cumulees(df_sap, "Date de la pièce", "Val./Devise objet")

        df_pv = lire_planned_value(str(fichier_pv))
        result_pv = traiter_planned_value(df_pv)
        assert result_pv is not None
        pv, _ = result_pv

        df_va = lire_valeur_acquise(str(fichier_va))
        ev = calculer_earned_value(df_pv, df_va)

        # Toutes les étapes doivent réussir
        assert ac is not None
        assert pv is not None
        assert ev is not None

        # Les valeurs doivent être cohérentes
        assert ac.sum() > 0
        assert pv.iloc[-1] > 0  # BAC doit être positif
        assert ev.iloc[-1] > 0  # EV doit être positif


class TestIndicateursEVM:
    """Tests pour le calcul des indicateurs EVM"""

    def test_calcul_cpi_spi(self, sample_depenses_cumulees, sample_ev_cumulee, sample_pv_cumulee):
        """Test du calcul de CPI et SPI"""
        # Valeurs actuelles (dernier mois)
        ac_actuel = sample_depenses_cumulees.iloc[-1]
        ev_actuel = sample_ev_cumulee.iloc[-1]
        pv_actuel = sample_pv_cumulee.iloc[-1]

        # Calculate Cost Performance Index (earned value divided by actual cost)
        cpi = ev_actuel / ac_actuel
        assert cpi > 0
        # Dans cet exemple: EV=180k, AC=250k => CPI=0.72 < 1 (mauvais)
        assert cpi < 1

        # Calculate Schedule Performance Index (earned value divided by planned value)
        spi = ev_actuel / pv_actuel
        assert spi > 0
        # Dans cet exemple: EV=180k, PV=200k => SPI=0.9 < 1 (retard)
        assert spi < 1

    def test_calcul_variances(self, sample_depenses_cumulees, sample_ev_cumulee, sample_pv_cumulee):
        """Test du calcul de CV et SV"""
        ac_actuel = sample_depenses_cumulees.iloc[-1]
        ev_actuel = sample_ev_cumulee.iloc[-1]
        pv_actuel = sample_pv_cumulee.iloc[-1]

        # Calculate Cost Variance (earned value minus actual cost)
        cv = ev_actuel - ac_actuel
        # Dans cet exemple: CV = 180k - 250k = -70k (dépassement)
        assert cv < 0

        # Calculate Schedule Variance (earned value minus planned value)
        sv = ev_actuel - pv_actuel
        # Dans cet exemple: SV = 180k - 200k = -20k (retard)
        assert sv < 0
