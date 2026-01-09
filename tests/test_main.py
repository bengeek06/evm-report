"""
Tests pour l'argument parsing et la fonction main
"""

import sys
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent))
from src.cli import main, parser_arguments


class TestParserArguments:
    """Tests pour le parser d'arguments"""

    def test_parser_arguments_defaut(self):
        """Test avec les arguments par défaut"""
        with patch("sys.argv", ["analyse.py"]):
            args = parser_arguments()
            assert args.sap == "EXPORT.XLSX"
            assert args.pv == "pv.xlsx"
            assert args.va == "va.xlsx"
            assert args.forecast == "forecast.xlsx"
            assert args.output == "analyse_evm.png"
            assert args.tableau == "tableau_evm"
            assert args.word is None

    def test_parser_arguments_personnalises(self):
        """Test avec des arguments personnalisés"""
        with patch(
            "sys.argv",
            [
                "analyse.py",
                "--sap",
                "mes_depenses.xlsx",
                "--pv",
                "mon_pv.xlsx",
                "--va",
                "ma_va.xlsx",
                "--forecast",
                "mon_forecast.xlsx",
                "--output",
                "mon_graphique.png",
                "--tableau",
                "mon_tableau",
                "--word",
                "mon_rapport.docx",
            ],
        ):
            args = parser_arguments()
            assert args.sap == "mes_depenses.xlsx"
            assert args.pv == "mon_pv.xlsx"
            assert args.va == "ma_va.xlsx"
            assert args.forecast == "mon_forecast.xlsx"
            assert args.output == "mon_graphique.png"
            assert args.tableau == "mon_tableau"
            assert args.word == "mon_rapport.docx"

    def test_parser_arguments_chemins_avec_espaces(self):
        """Test avec des chemins contenant des espaces"""
        with patch(
            "sys.argv",
            [
                "analyse.py",
                "--sap",
                "mes fichiers/export SAP.xlsx",
                "--output",
                "mes rapports/graphique EVM.png",
            ],
        ):
            args = parser_arguments()
            assert args.sap == "mes fichiers/export SAP.xlsx"
            assert args.output == "mes rapports/graphique EVM.png"

    def test_parser_arguments_chemins_absolus(self):
        """Test avec des chemins absolus"""
        with patch(
            "sys.argv",
            [
                "analyse.py",
                "--sap",
                "/home/user/data/export.xlsx",
                "--output",
                "/tmp/graphique.png",
            ],
        ):
            args = parser_arguments()
            assert args.sap == "/home/user/data/export.xlsx"
            assert args.output == "/tmp/graphique.png"

    def test_parser_help_affiche(self):
        """Test que l'aide s'affiche correctement"""
        with patch("sys.argv", ["analyse.py", "--help"]):
            with pytest.raises(SystemExit) as exc_info:
                parser_arguments()
            assert exc_info.value.code == 0


class TestMain:
    """Tests pour la fonction main"""

    def test_main_fichier_sap_manquant(self, capsys):
        """Test quand le fichier SAP n'existe pas"""
        with patch("sys.argv", ["analyse.py", "--sap", "fichier_inexistant.xlsx"]):
            main()
            captured = capsys.readouterr()
            assert "Erreur" in captured.out or "pas été trouvé" in captured.out

    def test_main_colonnes_manquantes(self, tmp_path, capsys):
        """Test quand les colonnes requises sont absentes"""
        # Créer un fichier Excel avec de mauvaises colonnes
        fichier_sap = tmp_path / "export_mauvais.xlsx"
        df = pd.DataFrame(
            {
                "Mauvaise_Colonne_1": [datetime(2025, 1, 15)],
                "Mauvaise_Colonne_2": [10000],
            }
        )
        df.to_excel(fichier_sap, index=False)

        with patch("sys.argv", ["analyse.py", "--sap", str(fichier_sap)]):
            main()
            captured = capsys.readouterr()
            assert "Colonnes requises non trouvées" in captured.out

    def test_main_workflow_minimal(self, tmp_path):
        """Test du workflow minimal avec fichier SAP valide"""
        # Créer un fichier SAP valide
        fichier_sap = tmp_path / "export.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [
                    datetime(2025, 1, 15),
                    datetime(2025, 2, 20),
                    datetime(2025, 3, 10),
                ],
                "Val./Devise objet": [10000, 15000, 20000],
            }
        )
        df.to_excel(fichier_sap, index=False)

        fichier_output = tmp_path / "graphique.png"

        with (
            patch("sys.argv", ["analyse.py", "--sap", str(fichier_sap), "--output", str(fichier_output)]),
            patch("src.cli.tracer_courbe_realise") as mock_tracer,
            patch("src.visualisation.graphiques.tracer_courbe_realise"),
            patch("src.visualisation.graphiques.tracer_courbe_projections"),
        ):
            main()
            # Vérifier que tracer_courbe_realise a été appelé
            mock_tracer.assert_called()

    def test_main_avec_tous_fichiers(self, tmp_path):
        """Test avec tous les fichiers d'entrée"""
        # Créer tous les fichiers nécessaires
        fichier_sap = tmp_path / "export.xlsx"
        df_sap = pd.DataFrame(
            {
                "Date de la pièce": [
                    datetime(2025, 1, 15),
                    datetime(2025, 2, 20),
                ],
                "Val./Devise objet": [10000, 15000],
            }
        )
        df_sap.to_excel(fichier_sap, index=False)

        fichier_pv = tmp_path / "pv.xlsx"
        df_pv = pd.DataFrame(
            {
                "Jalon": ["J1", "J2"],
                "Date": [datetime(2025, 3, 31), datetime(2025, 6, 30)],
                "Montant planifié": [50000, 100000],
                "Cumul planifié": [50000, 150000],
            }
        )
        df_pv.to_excel(fichier_pv, index=False)

        fichier_va = tmp_path / "va.xlsx"
        df_va = pd.DataFrame(
            {
                "Jalon": ["J1", "J2"],
                "Date": [datetime(2025, 3, 31), datetime(2025, 6, 30)],
                "Montant planifié": [50000, 100000],
                datetime(2025, 1, 1): [0.3, 0.2],
                datetime(2025, 2, 1): [0.6, 0.4],
            }
        )
        df_va.to_excel(fichier_va, index=False)

        fichier_forecast = tmp_path / "forecast.xlsx"
        df_forecast = pd.DataFrame(
            {
                "Jalon": ["J1", "J2"],
                "Date projetée": [datetime(2025, 4, 30), datetime(2025, 7, 31)],
                "EAC (€)": [55000, 110000],
                "ETC (€)": [5000, 10000],
            }
        )
        df_forecast.to_excel(fichier_forecast, index=False)

        fichier_output = tmp_path / "graphique.png"

        with (
            patch(
                "sys.argv",
                [
                    "analyse.py",
                    "--sap",
                    str(fichier_sap),
                    "--pv",
                    str(fichier_pv),
                    "--va",
                    str(fichier_va),
                    "--forecast",
                    str(fichier_forecast),
                    "--output",
                    str(fichier_output),
                ],
            ),
            patch("src.visualisation.graphiques.tracer_courbe_realise"),
            patch("src.visualisation.graphiques.tracer_courbe_projections"),
        ):
            main()

    def test_main_avec_rapport_word(self, tmp_path):
        """Test de génération de rapport Word"""
        # Créer un fichier SAP valide
        fichier_sap = tmp_path / "export.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier_sap, index=False)

        fichier_word = tmp_path / "rapport.docx"
        fichier_output = tmp_path / "graphique.png"

        with (
            patch(
                "sys.argv",
                [
                    "analyse.py",
                    "--sap",
                    str(fichier_sap),
                    "--output",
                    str(fichier_output),
                    "--word",
                    str(fichier_word),
                ],
            ),
            patch("src.visualisation.graphiques.tracer_courbe_realise"),
            patch("src.visualisation.graphiques.tracer_courbe_projections"),
            patch("src.cli.generer_rapport_word") as mock_word,
        ):
            main()
            # Vérifier que la génération du rapport Word a été appelée
            mock_word.assert_called_once()

    def test_main_chemin_sortie_personnalise(self, tmp_path):
        """Test avec des chemins de sortie personnalisés"""
        fichier_sap = tmp_path / "export.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier_sap, index=False)

        dossier_sortie = tmp_path / "resultats"
        dossier_sortie.mkdir()
        fichier_output = dossier_sortie / "mon_graphique_evm.png"
        fichier_tableau = dossier_sortie / "mon_tableau_evm"

        with (
            patch(
                "sys.argv",
                [
                    "analyse.py",
                    "--sap",
                    str(fichier_sap),
                    "--output",
                    str(fichier_output),
                    "--tableau",
                    str(fichier_tableau),
                ],
            ),
            patch("src.visualisation.graphiques.tracer_courbe_realise"),
            patch("src.visualisation.graphiques.tracer_courbe_projections"),
        ):
            main()


class TestMainRobustesse:
    """Tests de robustesse pour la fonction main"""

    def test_main_fichier_sap_corrompu(self, tmp_path, capsys):
        """Test avec un fichier SAP corrompu"""
        fichier_sap = tmp_path / "corrompu.xlsx"
        fichier_sap.write_text("Pas un fichier Excel valide")

        with patch("sys.argv", ["analyse.py", "--sap", str(fichier_sap)]):
            main()
            captured = capsys.readouterr()
            assert "Erreur" in captured.out

    def test_main_fichiers_optionnels_manquants(self, tmp_path):
        """Test quand les fichiers PV/VA/Forecast n'existent pas"""
        fichier_sap = tmp_path / "export.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier_sap, index=False)

        with (
            patch(
                "sys.argv",
                [
                    "analyse.py",
                    "--sap",
                    str(fichier_sap),
                    "--pv",
                    "pv_inexistant.xlsx",
                    "--va",
                    "va_inexistant.xlsx",
                ],
            ),
            patch("src.visualisation.graphiques.tracer_courbe_realise"),
            patch("src.visualisation.graphiques.tracer_courbe_projections"),
        ):
            # Ne devrait pas crasher
            main()

    def test_main_espace_disque_insuffisant(self, tmp_path):
        """Test de simulation d'espace disque insuffisant"""
        fichier_sap = tmp_path / "export.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier_sap, index=False)

        with (
            patch("sys.argv", ["analyse.py", "--sap", str(fichier_sap)]),
            patch("src.cli.tracer_courbe_realise", side_effect=OSError("No space left on device")),
            pytest.raises(OSError, match="No space left on device"),
        ):
            main()

    def test_main_multiples_executions_consecutives(self, tmp_path):
        """Test d'exécutions multiples consécutives"""
        fichier_sap = tmp_path / "export.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier_sap, index=False)

        fichier_output = tmp_path / "graphique.png"

        with (
            patch("sys.argv", ["analyse.py", "--sap", str(fichier_sap), "--output", str(fichier_output)]),
            patch("src.visualisation.graphiques.tracer_courbe_realise"),
            patch("src.visualisation.graphiques.tracer_courbe_projections"),
        ):
            # Exécution 1
            main()
            # Exécution 2 (devrait écraser les fichiers existants)
            main()
            # Exécution 3
            main()


class TestMainAffichage:
    """Tests pour l'affichage dans main"""

    def test_main_affiche_fichiers_entree(self, tmp_path, capsys):
        """Test que les fichiers d'entrée sont affichés"""
        fichier_sap = tmp_path / "export.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier_sap, index=False)

        with (
            patch("sys.argv", ["analyse.py", "--sap", str(fichier_sap)]),
            patch("src.visualisation.graphiques.tracer_courbe_realise"),
            patch("src.visualisation.graphiques.tracer_courbe_projections"),
        ):
            main()
            captured = capsys.readouterr()
            assert "Fichiers d'entrée" in captured.out
            assert "SAP:" in captured.out
            assert "PV:" in captured.out

    def test_main_affiche_colonnes_utilisees(self, tmp_path, capsys):
        """Test que les colonnes utilisées sont affichées"""
        fichier_sap = tmp_path / "export.xlsx"
        df = pd.DataFrame(
            {
                "Date de la pièce": [datetime(2025, 1, 15)],
                "Val./Devise objet": [10000],
            }
        )
        df.to_excel(fichier_sap, index=False)

        with (
            patch("sys.argv", ["analyse.py", "--sap", str(fichier_sap)]),
            patch("src.visualisation.graphiques.tracer_courbe_realise"),
            patch("src.visualisation.graphiques.tracer_courbe_projections"),
        ):
            main()
            captured = capsys.readouterr()
            assert "Colonne date utilisée" in captured.out
            assert "Colonne montant utilisée" in captured.out
