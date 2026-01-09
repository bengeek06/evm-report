"""
Module de calcul des dépenses cumulées
"""

import pandas as pd


def calculer_depenses_cumulees(df, colonne_date="Date", colonne_montant="Montant"):
    """
    Calcule les dépenses cumulées mois par mois

    Args:
        df: DataFrame contenant les données
        colonne_date: nom de la colonne contenant les dates
        colonne_montant: nom de la colonne contenant les montants
    """
    # Vérifications des données
    if df is None or df.empty:
        return pd.Series(dtype=float)

    if colonne_date not in df.columns:
        raise KeyError(f"La colonne '{colonne_date}' est absente du DataFrame")

    if colonne_montant not in df.columns:
        raise KeyError(f"La colonne '{colonne_montant}' est absente du DataFrame")

    # Copie du dataframe pour éviter les modifications
    df_work = df.copy()

    # Conversion de la colonne date en datetime si nécessaire
    if not pd.api.types.is_datetime64_any_dtype(df_work[colonne_date]):
        df_work[colonne_date] = pd.to_datetime(df_work[colonne_date])

    # Extraction du mois et année
    df_work["Mois"] = df_work[colonne_date].dt.to_period("M")

    # Agrégation par mois
    depenses_mensuelles = df_work.groupby("Mois")[colonne_montant].sum().sort_index()

    # Calcul des dépenses cumulées
    return depenses_mensuelles.cumsum()
