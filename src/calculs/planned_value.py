"""
Module de traitement de la Planned Value
"""

import pandas as pd


def _find_first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def _ensure_datetime(series: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    return pd.to_datetime(series)


def _interpolate_missing_months(pv_cumulee: pd.Series) -> pd.Series:
    if len(pv_cumulee) <= 1:
        return pv_cumulee

    premier_mois = pv_cumulee.index[0]
    dernier_mois = pv_cumulee.index[-1]
    tous_les_mois = pd.period_range(start=premier_mois, end=dernier_mois, freq="M")
    pv_cumulee_complete = pv_cumulee.reindex(tous_les_mois)

    pv_temp = pd.Series(
        pv_cumulee_complete.values,
        index=[p.to_timestamp() for p in pv_cumulee_complete.index],
    ).interpolate(method="time")

    return pd.Series(
        pv_temp.values,
        index=[pd.Period(d, freq="M") for d in pv_temp.index],
    )


def _extract_jalons_by_month(df_pv: pd.DataFrame, colonne_date: str) -> dict[pd.Period, list[str]]:
    jalons: dict[pd.Period, list[str]] = {}
    if "Jalon" not in df_pv.columns:
        return jalons

    for _, row in df_pv.iterrows():
        mois = pd.Period(row[colonne_date], freq="M")
        jalons.setdefault(mois, []).append(row["Jalon"])
    return jalons


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

    colonne_date = _find_first_existing_column(df_pv, colonnes_date)
    colonne_montant = _find_first_existing_column(df_pv, colonnes_montant)

    if colonne_date is None or colonne_montant is None:
        print(f"\n⚠️  Colonnes PV non détectées. Colonnes disponibles: {df_pv.columns.tolist()}")
        return None

    print(f"✓ PV - Colonne date: {colonne_date}")
    print(f"✓ PV - Colonne montant: {colonne_montant}")

    df_work = df_pv.copy()

    # Conversion en datetime
    df_work[colonne_date] = _ensure_datetime(df_work[colonne_date])

    # Extraction du mois
    df_work["Mois"] = df_work[colonne_date].dt.to_period("M")

    # Agrégation par mois
    pv_mensuelle = df_work.groupby("Mois")[colonne_montant].sum().sort_index()

    # Calcul cumulé
    pv_cumulee = pv_mensuelle.cumsum()

    # Interpolation linéaire pour remplir les mois manquants
    pv_cumulee_interp = _interpolate_missing_months(pv_cumulee)
    if len(pv_cumulee_interp) != len(pv_cumulee):
        pv_cumulee = pv_cumulee_interp
        print(f"✓ PV interpolée sur {len(pv_cumulee)} mois")

    # Récupération des jalons pour chaque mois
    jalons = _extract_jalons_by_month(df_pv, colonne_date)

    return pv_cumulee, jalons
