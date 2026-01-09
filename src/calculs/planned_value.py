"""
Module de traitement de la Planned Value
"""

import pandas as pd


def traiter_planned_value(df_pv):
    """
    Traite les données de Planned Value et calcule les montants cumulés
    """
    if df_pv is None:
        return None

    # Détection automatique des colonnes
    colonnes_date = ["Date", "date", "Date prévisionnelle", "Date prév.", "Mois"]
    colonnes_montant = [
        "Montant planifié",
        "Cumul planifié",
        "Montant",
        "montant",
        "Montant prévisionnel",
        "Montant prév.",
        "PV",
        "Planned Value",
    ]

    colonne_date = None
    colonne_montant = None

    for col in colonnes_date:
        if col in df_pv.columns:
            colonne_date = col
            break

    for col in colonnes_montant:
        if col in df_pv.columns:
            colonne_montant = col
            break

    if colonne_date is None or colonne_montant is None:
        print(f"\n⚠️  Colonnes PV non détectées. Colonnes disponibles: {df_pv.columns.tolist()}")
        return None

    print(f"✓ PV - Colonne date: {colonne_date}")
    print(f"✓ PV - Colonne montant: {colonne_montant}")

    df_work = df_pv.copy()

    # Conversion en datetime
    if not pd.api.types.is_datetime64_any_dtype(df_work[colonne_date]):
        df_work[colonne_date] = pd.to_datetime(df_work[colonne_date])

    # Extraction du mois
    df_work["Mois"] = df_work[colonne_date].dt.to_period("M")

    # Agrégation par mois
    pv_mensuelle = df_work.groupby("Mois")[colonne_montant].sum().sort_index()

    # Calcul cumulé
    pv_cumulee = pv_mensuelle.cumsum()

    # Interpolation linéaire pour remplir les mois manquants
    if len(pv_cumulee) > 1:
        # Créer une série avec tous les mois entre le premier et le dernier jalon
        premier_mois = pv_cumulee.index[0]
        dernier_mois = pv_cumulee.index[-1]

        # Générer tous les mois dans cet intervalle
        tous_les_mois = pd.period_range(start=premier_mois, end=dernier_mois, freq="M")

        # Réindexer avec tous les mois
        pv_cumulee_complete = pv_cumulee.reindex(tous_les_mois)

        # Interpolation linéaire pour les valeurs manquantes
        # Convertir en série temporelle pour l'interpolation
        pv_temp = pd.Series(pv_cumulee_complete.values, index=[p.to_timestamp() for p in pv_cumulee_complete.index])
        pv_temp = pv_temp.interpolate(method="time")

        # Reconvertir en Period index
        pv_cumulee = pd.Series(pv_temp.values, index=[pd.Period(d, freq="M") for d in pv_temp.index])

        print(f"✓ PV interpolée sur {len(pv_cumulee)} mois")

    # Récupération des jalons pour chaque mois
    jalons = {}
    if "Jalon" in df_pv.columns:
        for _, row in df_pv.iterrows():
            mois = pd.Period(row[colonne_date], freq="M")
            if mois not in jalons:
                jalons[mois] = []
            jalons[mois].append(row["Jalon"])

    return pv_cumulee, jalons
