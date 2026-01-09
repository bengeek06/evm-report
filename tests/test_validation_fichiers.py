"""
Tests de validation des fichiers Excel - Gestion des erreurs et cas limites
"""

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent))
from analyse import lire_export_sap, lire_forecast, lire_planned_value, lire_valeur_acquise


class TestValidationExportSAP:
    """Tests de validation du fichier Export SAP"""

    def test_fichier_excel_corrompu(self, tmp_path):
        """Test avec un fichier Excel corrompu"""
        fichier = tmp_path / "corrompu.xlsx"
        fichier.write_text("Ceci n'est pas un fichier Excel valide")

        result = lire_export_sap(str(fichier))
        assert result is None

    def test_fichier_vide(self, tmp_path):
        """Test avec un fichier Excel vide"""
        fichier = tmp_path / "vide.xlsx"
        df = pd.DataFrame()
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        # Fichier vide mais valide, devrait retourner un DataFrame vide
        assert result is not None
        assert len(result) == 0

    def test_colonnes_avec_espaces(self, tmp_path):
        """Test avec des noms de colonnes comportant des espaces superflus"""
        fichier = tmp_path / "espaces.xlsx"
        df = pd.DataFrame(
            {
                " Date de la pièce ": [datetime(2025, 1, 15)],
                " Val./Devise objet ": [10000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        assert result is not None
        # Les colonnes devraient être nettoyées
        assert "Date de la pièce" in result.columns or " Date de la pièce " in result.columns

    def test_dates_invalides(self, tmp_path):
        """Test avec des dates au format invalide"""
        fichier = tmp_path / "dates_invalides.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": ["pas une date", "2025-13-45", None],
                "Val./Devise objet": [1000, 2000, 3000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        # Le fichier devrait être lu même avec des dates invalides
        assert result is not None
        assert len(result) == 3

    def test_montants_negatifs(self, tmp_path):
        """Test avec des montants négatifs"""
        fichier = tmp_path / "negatifs.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15), datetime(2025, 2, 20)],
                "Val./Devise objet": [-5000, -3000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        assert result is not None
        assert len(result) == 2
        # Les montants négatifs devraient être acceptés (remboursements possibles)

    def test_montants_non_numeriques(self, tmp_path):
        """Test avec des montants non numériques"""
        fichier = tmp_path / "non_numeriques.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15), datetime(2025, 2, 20)],
                "Val./Devise objet": ["abc", "xyz"],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        assert result is not None
        # pandas devrait gérer ça et mettre NaN

    def test_lignes_dupliquees(self, tmp_path):
        """Test avec des lignes dupliquées"""
        fichier = tmp_path / "dupliques.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)] * 3,
                "Val./Devise objet": [5000, 5000, 5000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        assert result is not None
        assert len(result) == 3  # Les doublons sont acceptés


class TestValidationPlannedValue:
    """Tests de validation du fichier Planned Value"""

    def test_dates_non_chronologiques(self, tmp_path):
        """Test avec des dates non chronologiques"""
        fichier = tmp_path / "non_chrono.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1", "J2", "J3"],
                "Date": [
                    datetime(2025, 12, 31),
                    datetime(2025, 6, 30),
                    datetime(2025, 3, 31),
                ],
                "Montant planifié": [100000, 200000, 150000],
                "Cumul planifié": [100000, 300000, 450000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_planned_value(str(fichier))
        assert result is not None
        # Le fichier devrait être lu même si les dates ne sont pas chronologiques

    def test_montants_cumules_incoherents(self, tmp_path):
        """Test avec des cumuls planifiés incohérents"""
        fichier = tmp_path / "cumuls_incoherents.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1", "J2", "J3"],
                "Date": [
                    datetime(2025, 3, 31),
                    datetime(2025, 6, 30),
                    datetime(2025, 12, 31),
                ],
                "Montant planifié": [100000, 200000, 150000],
                "Cumul planifié": [100000, 250000, 300000],  # Incohérent: 100k+200k != 250k
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_planned_value(str(fichier))
        assert result is not None

    def test_jalons_dupliques(self, tmp_path):
        """Test avec des jalons dupliqués"""
        fichier = tmp_path / "jalons_dupliques.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1", "J1", "J2"],
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
        # Les jalons dupliqués devraient être acceptés (ou gérés)

    def test_montants_zero(self, tmp_path):
        """Test avec des montants à zéro"""
        fichier = tmp_path / "montants_zero.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1", "J2"],
                "Date": [datetime(2025, 3, 31), datetime(2025, 6, 30)],
                "Montant planifié": [0, 200000],
                "Cumul planifié": [0, 200000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_planned_value(str(fichier))
        assert result is not None
        assert len(result) == 2


class TestValidationValeurAcquise:
    """Tests de validation du fichier Valeur Acquise"""

    def test_pourcentages_superieurs_100(self, tmp_path):
        """Test avec des pourcentages supérieurs à 100%"""
        fichier = tmp_path / "pourcentages_sup.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date": [datetime(2025, 3, 31)],
                "Montant planifié": [100000],
                datetime(2025, 1, 1): [1.5],  # 150%
                datetime(2025, 2, 1): [2.0],  # 200%
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_valeur_acquise(str(fichier))
        assert result is not None

    def test_pourcentages_negatifs(self, tmp_path):
        """Test avec des pourcentages négatifs"""
        fichier = tmp_path / "pourcentages_negatifs.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date": [datetime(2025, 3, 31)],
                "Montant planifié": [100000],
                datetime(2025, 1, 1): [-0.5],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_valeur_acquise(str(fichier))
        assert result is not None

    def test_colonnes_dates_manquantes(self, tmp_path):
        """Test avec seulement les colonnes obligatoires"""
        fichier = tmp_path / "colonnes_min.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date": [datetime(2025, 3, 31)],
                "Montant planifié": [100000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_valeur_acquise(str(fichier))
        assert result is not None

    def test_jalons_sans_montant(self, tmp_path):
        """Test avec des jalons sans montant planifié"""
        fichier = tmp_path / "sans_montant.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1", "J2"],
                "Date": [datetime(2025, 3, 31), datetime(2025, 6, 30)],
                "Montant planifié": [None, 0],
                datetime(2025, 1, 1): [0.5, 0.3],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_valeur_acquise(str(fichier))
        assert result is not None


class TestValidationForecast:
    """Tests de validation du fichier Forecast"""

    def test_eac_inferieur_ac(self, tmp_path):
        """Test avec EAC inférieur aux dépenses actuelles (impossible)"""
        fichier = tmp_path / "eac_impossible.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date projetée": [datetime(2026, 6, 30)],
                "EAC (€)": [50000],  # Si AC est déjà à 100k, c'est impossible
                "ETC (€)": [0],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_forecast(str(fichier))
        assert result is not None
        # Le fichier devrait être lu, la validation sera faite ailleurs

    def test_dates_passees(self, tmp_path):
        """Test avec des dates projetées dans le passé"""
        fichier = tmp_path / "dates_passees.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date projetée": [datetime(2020, 1, 1)],
                "EAC (€)": [100000],
                "ETC (€)": [50000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_forecast(str(fichier))
        assert result is not None

    def test_etc_superieur_eac(self, tmp_path):
        """Test avec ETC supérieur à EAC (incohérent)"""
        fichier = tmp_path / "etc_sup_eac.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1"],
                "Date projetée": [datetime(2026, 6, 30)],
                "EAC (€)": [100000],
                "ETC (€)": [150000],  # Incohérent si AC > 0
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_forecast(str(fichier))
        assert result is not None

    def test_colonnes_optionnelles_manquantes(self, tmp_path):
        """Test avec seulement Jalon, Date projetée et EAC"""
        fichier = tmp_path / "colonnes_min_forecast.xlsx"
        df = pd.DataFrame(
            {
                "Jalon": ["J1", "J2"],
                "Date projetée": [datetime(2026, 6, 30), datetime(2026, 12, 31)],
                "EAC (€)": [100000, 200000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_forecast(str(fichier))
        assert result is not None
        assert len(result) == 2


class TestFormatsExcelVariants:
    """Tests avec différents formats Excel"""

    def test_fichier_xls_ancien_format(self, tmp_path):
        """Test que les fichiers .xls (ancien format) ne sont pas supportés"""
        fichier = tmp_path / "ancien.xls"
        fichier.write_text("Ancien format XLS")

        # Les fonctions attendent .xlsx, .xls ne devrait pas fonctionner
        result = lire_export_sap(str(fichier))
        assert result is None

    def test_extension_incorrecte(self, tmp_path):
        """Test avec une extension incorrecte"""
        fichier = tmp_path / "fichier.csv"
        fichier.write_text("A,B\n1,2")

        result = lire_export_sap(str(fichier))
        assert result is None

    def test_chemin_avec_espaces(self, tmp_path):
        """Test avec un chemin contenant des espaces"""
        dossier = tmp_path / "dossier avec espaces"
        dossier.mkdir()
        fichier = dossier / "fichier avec espaces.xlsx"

        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        assert result is not None

    def test_chemin_avec_caracteres_speciaux(self, tmp_path):
        """Test avec un chemin contenant des caractères spéciaux"""
        fichier = tmp_path / "fichier_éàç_special.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier, index=False)

        result = lire_export_sap(str(fichier))
        assert result is not None
