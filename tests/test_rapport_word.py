"""
Tests pour la génération de rapports Word
"""

import struct
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent))
from analyse import generer_rapport_word


class TestGenerationRapportWord:
    """Tests pour la génération de rapports Word"""

    def test_generer_rapport_word_complet(self, tmp_path):
        """Test de génération d'un rapport Word complet"""
        fichier_word = tmp_path / "rapport.docx"
        fichier_graphique = tmp_path / "graphique_realise.png"
        fichier_projections = tmp_path / "graphique_projections.png"

        # Créer des fichiers graphiques PNG valides (1x1 pixel)
        def create_minimal_png(path):
            # Créer un PNG minimal 1x1 transparent
            png_header = b"\x89PNG\r\n\x1a\n"
            ihdr = b"IHDR" + struct.pack(">2I5B", 1, 1, 8, 6, 0, 0, 0)
            ihdr_crc = struct.pack(">I", 0xD951E134)  # CRC pré-calculé
            idat = b"IDAT" + b"\x08\x1d\x01\x05\x00\xfa\xff\x00\x00\x00\x00\x01"
            idat_crc = struct.pack(">I", 0xDE09EC4C)  # CRC pré-calculé
            iend = b"IEND" + struct.pack(">I", 0xAE426082)  # CRC pré-calculé

            png_data = png_header
            png_data += struct.pack(">I", 13) + ihdr + ihdr_crc
            png_data += struct.pack(">I", 12) + idat + idat_crc
            png_data += struct.pack(">I", 0) + iend

            path.write_bytes(png_data)

        create_minimal_png(fichier_graphique)
        create_minimal_png(fichier_projections)

        # Données de test
        df_tableau = pd.DataFrame(
            {
                "Mois": ["2025-01", "2025-02", "2025-03"],
                "AC (Dépenses réelles)": [10000, 25000, 45000],
                "PV (Budget prévu)": [12000, 24000, 48000],
                "EV (Valeur acquise)": [11000, 23000, 46000],
            }
        )

        depenses = pd.Series(
            [10000, 25000, 45000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        pv = pd.Series(
            [12000, 24000, 48000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        ev = pd.Series(
            [11000, 23000, 46000],
            index=pd.period_range(start="2025-01", periods=3, freq="M"),
        )

        projections = {
            "cpi": {
                "series": pd.Series([48000], index=pd.period_range(start="2025-04", periods=1, freq="M")),
                "eac": 48000,
                "date": datetime(2025, 4, 30),
            },
            "cpi_spi": {
                "series": pd.Series([50000], index=pd.period_range(start="2025-04", periods=1, freq="M")),
                "eac": 50000,
                "date": datetime(2025, 4, 30),
            },
        }

        # Générer le rapport
        generer_rapport_word(
            str(fichier_word),
            df_tableau,
            str(fichier_graphique),
            depenses,
            pv,
            ev,
            projections,
            str(fichier_projections),
        )

        # Vérifier que le fichier a été créé
        assert fichier_word.exists()

    def test_generer_rapport_word_avec_toutes_projections(self, tmp_path):
        """Test avec toutes les méthodes de projection"""
        fichier_word = tmp_path / "rapport_projections.docx"
        fichier_graphique = tmp_path / "graphique_realise.png"
        fichier_projections = tmp_path / "graphique_projections.png"

        # Créer des fichiers graphiques PNG valides (1x1 pixel)
        def create_minimal_png(path):
            # Créer un PNG minimal 1x1 transparent
            png_header = b"\x89PNG\r\n\x1a\n"
            ihdr = b"IHDR" + struct.pack(">2I5B", 1, 1, 8, 6, 0, 0, 0)
            ihdr_crc = struct.pack(">I", 0xD951E134)  # CRC pré-calculé
            idat = b"IDAT" + b"\x08\x1d\x01\x05\x00\xfa\xff\x00\x00\x00\x00\x01"
            idat_crc = struct.pack(">I", 0xDE09EC4C)  # CRC pré-calculé
            iend = b"IEND" + struct.pack(">I", 0xAE426082)  # CRC pré-calculé

            png_data = png_header
            png_data += struct.pack(">I", 13) + ihdr + ihdr_crc
            png_data += struct.pack(">I", 12) + idat + idat_crc
            png_data += struct.pack(">I", 0) + iend

            path.write_bytes(png_data)

        create_minimal_png(fichier_graphique)
        create_minimal_png(fichier_projections)

        df_tableau = pd.DataFrame(
            {
                "Mois": ["2025-01"],
                "AC (Dépenses réelles)": [10000],
                "PV (Budget prévu)": [12000],
                "EV (Valeur acquise)": [11000],
            }
        )

        depenses = pd.Series([10000], index=pd.period_range(start="2025-01", periods=1, freq="M"))
        pv = pd.Series([12000], index=pd.period_range(start="2025-01", periods=1, freq="M"))
        ev = pd.Series([11000], index=pd.period_range(start="2025-01", periods=1, freq="M"))

        projections = {
            "cpi": {
                "series": pd.Series([48000], index=pd.period_range(start="2025-04", periods=1, freq="M")),
                "eac": 48000,
                "date": datetime(2025, 4, 30),
            },
            "cpi_spi": {
                "series": pd.Series([50000], index=pd.period_range(start="2025-04", periods=1, freq="M")),
                "eac": 50000,
                "date": datetime(2025, 4, 30),
            },
            "reste_plan": {
                "series": pd.Series([45000], index=pd.period_range(start="2025-04", periods=1, freq="M")),
                "eac": 45000,
                "date": datetime(2025, 4, 30),
            },
            "forecast": {
                "series": pd.Series([52000], index=pd.period_range(start="2025-04", periods=1, freq="M")),
                "eac": 52000,
                "date": datetime(2025, 4, 30),
            },
        }

        generer_rapport_word(
            str(fichier_word),
            df_tableau,
            str(fichier_graphique),
            depenses,
            pv,
            ev,
            projections,
            str(fichier_projections),
        )

        assert fichier_word.exists()
