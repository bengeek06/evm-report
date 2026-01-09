"""
Module de lecture des fichiers Excel pour l'analyse EVM
"""

import pandas as pd


def lire_export_sap(fichier="EXPORT.XLSX"):
    """
    Lit le fichier Excel d'export SAP des dépenses
    """
    try:
        df = pd.read_excel(fichier)
        print(f"Fichier chargé avec succès: {len(df)} lignes")
        print(f"Colonnes disponibles: {df.columns.tolist()}")
        return df
    except FileNotFoundError:
        print(f"Erreur: Le fichier {fichier} n'a pas été trouvé")
        return None
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier: {e}")
        return None


def lire_planned_value(fichier="pv.xlsx"):
    """
    Lit le fichier Excel contenant la Planned Value (PV)
    """
    try:
        df_pv = pd.read_excel(fichier)
        print(f"\nFichier PV chargé avec succès: {len(df_pv)} lignes")
        print(f"Colonnes PV: {df_pv.columns.tolist()}")
        return df_pv
    except FileNotFoundError:
        print(f"\n⚠️  Le fichier {fichier} n'a pas été trouvé")
        print("La Planned Value ne sera pas affichée.")
        return None
    except Exception as e:
        print(f"\n⚠️  Erreur lors de la lecture du fichier PV: {e}")
        return None


def lire_valeur_acquise(fichier="va.xlsx"):
    """
    Lit le fichier Excel contenant les pourcentages d'avancement (Valeur Acquise)
    """
    try:
        df_va = pd.read_excel(fichier)
        print(f"\nFichier VA chargé avec succès: {len(df_va)} lignes")
        print(f"Colonnes VA: {df_va.columns.tolist()}")
        return df_va
    except FileNotFoundError:
        print(f"\n⚠️  Le fichier {fichier} n'a pas été trouvé")
        print("La Valeur Acquise ne sera pas affichée.")
        return None
    except Exception as e:
        print(f"\n⚠️  Erreur lors de la lecture du fichier VA: {e}")
        return None


def lire_forecast(fichier="forecast.xlsx"):
    """
    Lit le fichier Excel contenant les projections (forecast)
    """
    try:
        df_forecast = pd.read_excel(fichier)
        print(f"\nFichier Forecast chargé avec succès: {len(df_forecast)} lignes")
        print(f"Colonnes Forecast: {df_forecast.columns.tolist()}")
        return df_forecast
    except FileNotFoundError:
        print(f"\n⚠️  Le fichier {fichier} n'a pas été trouvé")
        print("Les projections ne seront pas affichées.")
        return None
    except Exception as e:
        print(f"\n⚠️  Erreur lors de la lecture du fichier Forecast: {e}")
        return None
