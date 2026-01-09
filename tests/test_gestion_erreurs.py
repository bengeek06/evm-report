"""
Tests de gestion des erreurs et des exceptions
"""

import sys
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))
from analyse import (
    calculer_depenses_cumulees,
    calculer_earned_value,
    lire_export_sap,
    lire_planned_value,
    lire_valeur_acquise,
    traiter_planned_value,
)


class TestErreursFichiers:
    """Tests des erreurs liées aux fichiers"""

    def test_fichier_inexistant(self):
        """Test avec un fichier qui n'existe pas"""
        result = lire_export_sap("/chemin/inexistant/fichier.xlsx")
        assert result is None

    def test_chemin_est_un_repertoire(self, tmp_path):
        """Test avec un chemin qui pointe vers un répertoire"""
        result = lire_export_sap(str(tmp_path))
        assert result is None

    def test_permissions_lecture_refusees(self, tmp_path):
        """Test avec un fichier sans permissions de lecture"""
        fichier = tmp_path / "sans_permission.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier, index=False)

        # Retirer les permissions de lecture
        Path(fichier).chmod(0o000)

        try:
            result = lire_export_sap(str(fichier))
            assert result is None
        finally:
            # Restaurer les permissions pour le nettoyage
            Path(fichier).chmod(0o644)

    def test_chemin_vide(self):
        """Test avec un chemin vide"""
        result = lire_export_sap("")
        assert result is None

    def test_chemin_none(self):
        """Test avec un chemin None"""
        # Le code gère maintenant None et retourne None au lieu de lever une exception
        result = lire_export_sap(None)  # type: ignore[arg-type]
        assert result is None

    def test_disque_plein_simulation(self, tmp_path):
        """Simulation d'un disque plein lors de l'écriture"""
        fichier = tmp_path / "test.xlsx"

        with patch("pandas.read_excel", side_effect=OSError("Disque plein")):
            result = lire_export_sap(str(fichier))
            assert result is None


class TestErreursDonnees:
    """Tests des erreurs liées aux données"""

    def test_dataframe_none_depenses_cumulees(self):
        """Test calculer_depenses_cumulees avec None"""
        # Le code gère maintenant None et retourne une série vide
        result = calculer_depenses_cumulees(None)
        assert result is not None
        assert len(result) == 0

    def test_dataframe_vide_depenses_cumulees(self):
        """Test calculer_depenses_cumulees avec DataFrame vide"""
        df = pd.DataFrame()
        result = calculer_depenses_cumulees(df)
        # Devrait retourner une série vide
        assert result is not None
        assert len(result) == 0

    def test_colonnes_manquantes_depenses_cumulees(self):
        """Test calculer_depenses_cumulees avec colonnes manquantes"""
        df = pd.DataFrame({"Colonne_Inconnue": [1, 2, 3]})
        with pytest.raises(KeyError):
            calculer_depenses_cumulees(df)

    def test_dataframe_none_earned_value(self):
        """Test calculer_earned_value avec None"""
        result = calculer_earned_value(None, None)
        assert result is None

    def test_dates_none_earned_value(self):
        """Test calculer_earned_value avec DataFrame None"""
        df = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date": [datetime(2025, 3, 31)],
                "Montant planifié": [100000],
            }
        )
        result = calculer_earned_value(df, None)
        # Devrait gérer le cas gracieusement
        assert result is None

    def test_montants_negatifs_dans_calculs(self):
        """Test des calculs avec des montants négatifs"""
        df = pd.DataFrame(
            {
                "Jalon": ["J1", "J2"],
                "Date": [datetime(2025, 1, 15), datetime(2025, 2, 15)],
                "Montant planifié": [-50000, 100000],
                "Cumul planifié": [-50000, 50000],
            }
        )
        result = traiter_planned_value(df)
        assert result is not None
        # Les montants négatifs devraient être traités

    def test_donnees_nan_dans_calculs(self):
        """Test des calculs avec des valeurs NaN"""
        df = pd.DataFrame(
            {
                "Date": [datetime(2025, 1, 15), datetime(2025, 2, 15)],
                "Montant": [10000, float("nan")],
            }
        )
        result = calculer_depenses_cumulees(df, "Date", "Montant")
        assert result is not None
        # Les NaN devraient être gérés (supprimés ou remplacés par 0)

    def test_donnees_infinies_dans_calculs(self):
        """Test des calculs avec des valeurs infinies"""
        df = pd.DataFrame(
            {
                "Date": [datetime(2025, 1, 15), datetime(2025, 2, 15)],
                "Montant": [10000, float("inf")],
            }
        )
        result = calculer_depenses_cumulees(df, "Date", "Montant")
        assert result is not None


class TestCoherenceDonnees:
    """Tests de cohérence des données entre fichiers"""

    def test_jalons_manquants_entre_fichiers(self, tmp_path):
        """Test quand un jalon existe dans PV mais pas dans VA"""
        fichier_pv = tmp_path / "pv.xlsx"
        df_pv = pd.DataFrame(
            {
                "Jalon": ["J1", "J2"],
                "Date": [datetime(2025, 3, 31), datetime(2025, 6, 30)],
                "Montant planifié": [100000, 200000],
                "Cumul planifié": [100000, 300000],
            }
        )
        df_pv.to_excel(fichier_pv, index=False)

        fichier_va = tmp_path / "va.xlsx"
        df_va = pd.DataFrame(
            {
                "Jalon": ["J1"],  # J2 manquant
                "Date": [datetime(2025, 3, 31)],
                "Montant planifié": [100000],
                datetime(2025, 1, 1): [0.5],
            }
        )
        df_va.to_excel(fichier_va, index=False)

        pv = lire_planned_value(str(fichier_pv))
        va = lire_valeur_acquise(str(fichier_va))

        assert pv is not None
        assert va is not None
        assert len(pv) == 2
        assert len(va) == 1

    def test_montants_differents_entre_fichiers(self, tmp_path):
        """Test quand le montant planifié diffère entre PV et VA"""
        fichier_pv = tmp_path / "pv.xlsx"
        df_pv = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date": [datetime(2025, 3, 31)],
                "Montant planifié": [100000],
                "Cumul planifié": [100000],
            }
        )
        df_pv.to_excel(fichier_pv, index=False)

        fichier_va = tmp_path / "va.xlsx"
        df_va = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date": [datetime(2025, 3, 31)],
                "Montant planifié": [150000],  # Montant différent
                datetime(2025, 1, 1): [0.5],
            }
        )
        df_va.to_excel(fichier_va, index=False)

        pv = lire_planned_value(str(fichier_pv))
        va = lire_valeur_acquise(str(fichier_va))

        assert pv is not None
        assert va is not None
        # Les montants différents devraient être détectés ou gérés

    def test_dates_differentes_entre_fichiers(self, tmp_path):
        """Test quand la date d'un jalon diffère entre PV et VA"""
        fichier_pv = tmp_path / "pv.xlsx"
        df_pv = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date": [datetime(2025, 3, 31)],
                "Montant planifié": [100000],
                "Cumul planifié": [100000],
            }
        )
        df_pv.to_excel(fichier_pv, index=False)

        fichier_va = tmp_path / "va.xlsx"
        df_va = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date": [datetime(2025, 6, 30)],  # Date différente
                "Montant planifié": [100000],
                datetime(2025, 1, 1): [0.5],
            }
        )
        df_va.to_excel(fichier_va, index=False)

        pv = lire_planned_value(str(fichier_pv))
        va = lire_valeur_acquise(str(fichier_va))

        assert pv is not None
        assert va is not None


class TestMemoire:
    """Tests liés à la gestion de la mémoire"""

    def test_fichier_tres_volumineux(self, tmp_path):
        """Test avec un fichier Excel très volumineux"""
        fichier = tmp_path / "volumineux.xlsx"

        # Créer un DataFrame avec beaucoup de lignes
        dates = pd.date_range(start="2020-01-01", end="2025-12-31", freq="D")
        df = pd.DataFrame(
            {
                "Date de la pièce": dates,
                "Val./Devise objet": [1000.0] * len(dates),
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        assert result is not None
        assert len(result) > 1000  # Devrait avoir beaucoup de lignes

    def test_nombreux_jalons(self, tmp_path):
        """Test avec un très grand nombre de jalons"""
        fichier = tmp_path / "nombreux_jalons.xlsx"

        # Créer un DataFrame avec beaucoup de jalons
        n_jalons = 1000
        df = pd.DataFrame(
            {
                "Jalon": [f"J{i}" for i in range(n_jalons)],
                "Date": [datetime(2025, 1, 1) + pd.Timedelta(days=i) for i in range(n_jalons)],
                "Montant planifié": [10000] * n_jalons,
                "Cumul planifié": [10000 * (i + 1) for i in range(n_jalons)],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_planned_value(str(fichier))
        assert result is not None
        assert len(result) == n_jalons


class TestEncodage:
    """Tests liés à l'encodage des fichiers"""

    def test_noms_jalons_caracteres_speciaux(self, tmp_path):
        """Test avec des noms de jalons contenant des caractères spéciaux"""
        fichier = tmp_path / "caracteres_speciaux.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["Jalon É", "Jalon ç", "Jalon à"],
                "Date": [
                    datetime(2025, 3, 31),
                    datetime(2025, 6, 30),
                    datetime(2025, 12, 31),
                ],
                "Montant planifié": [100000, 200000, 150000],
                "Cumul planifié": [100000, 300000, 450000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_planned_value(str(fichier))
        assert result is not None
        assert len(result) == 3

    def test_noms_jalons_unicode(self, tmp_path):
        """Test avec des noms de jalons en Unicode"""
        fichier = tmp_path / "unicode.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["日本語", "中文", "한국어"],
                "Date": [
                    datetime(2025, 3, 31),
                    datetime(2025, 6, 30),
                    datetime(2025, 12, 31),
                ],
                "Montant planifié": [100000, 200000, 150000],
                "Cumul planifié": [100000, 300000, 450000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_planned_value(str(fichier))
        assert result is not None
        assert len(result) == 3

    def test_emojis_dans_jalons(self, tmp_path):
        """Test avec des emojis dans les noms de jalons"""
        fichier = tmp_path / "emojis.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["Jalon 🚀", "Jalon ✅", "Jalon 📊"],
                "Date": [
                    datetime(2025, 3, 31),
                    datetime(2025, 6, 30),
                    datetime(2025, 12, 31),
                ],
                "Montant planifié": [100000, 200000, 150000],
                "Cumul planifié": [100000, 300000, 450000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_planned_value(str(fichier))
        assert result is not None
        assert len(result) == 3
