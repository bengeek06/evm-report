"""
Module de calcul de la Earned Value
"""

import pandas as pd


def _find_first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def _extract_va_month_columns(df_va: pd.DataFrame) -> dict[pd.Period, list[str]]:
    month_to_cols: dict[pd.Period, list[str]] = {}
    for col in df_va.columns:
        if col == "Jalon":
            continue
        try:
            date_col = pd.to_datetime(col)
        except Exception:
            continue
        month = date_col.to_period("M")
        month_to_cols.setdefault(month, []).append(col)
    return month_to_cols


def _build_va_percent_table(df_va: pd.DataFrame) -> pd.DataFrame:
    """Retourne un tableau (index=Jalon, colonnes=mois Period) des % cumulés par jalon."""
    month_to_cols = _extract_va_month_columns(df_va)
    if not month_to_cols:
        return pd.DataFrame(dtype=float)

    va_by_jalon = df_va.set_index("Jalon")
    month_frames: list[pd.Series] = []
    for month, cols in month_to_cols.items():
        # max si plusieurs colonnes pointent le même mois
        values = va_by_jalon[cols].apply(pd.to_numeric, errors="coerce").max(axis=1)
        values.name = month
        month_frames.append(values)

    table = pd.concat(month_frames, axis=1)
    table = table.reindex(sorted(table.columns), axis=1)
    # Cumulatif : pour chaque mois, garder le max jusqu'ici
    return table.cummax(axis=1).fillna(0).astype(float)


def _print_ev_debug(ev_by_month: pd.Series, df_pv: pd.DataFrame, montant_col: str, va_table: pd.DataFrame) -> None:
    if ev_by_month.empty:
        return

    months = list(ev_by_month.index)
    for month in months:
        print(f"\nMois {month}: EV totale = {ev_by_month[month]:.2f}€")

        for _idx, row_pv in df_pv.iterrows():
            jalon = row_pv["Jalon"]
            montant_jalon = row_pv[montant_col]

            if jalon not in va_table.index or month not in va_table.columns:
                continue

            pct = float(va_table.loc[jalon, month])
            if pct <= 0:
                continue

            ev_jalon = montant_jalon * pct
            print(f"  {jalon}: {pct * 100:.1f}% -> {ev_jalon:.2f}€")


def calculer_earned_value(df_pv, df_va):
    """
    Calcule la Earned Value (EV) en combinant les montants planifiés et les pourcentages d'avancement
    """
    if df_pv is None or df_va is None:
        return None

    # Identification des colonnes de dates et montants dans PV
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

    colonne_date_pv = _find_first_existing_column(df_pv, colonnes_date)
    colonne_montant_pv = _find_first_existing_column(df_pv, colonnes_montant)

    if colonne_date_pv is None or colonne_montant_pv is None:
        print("\n⚠️  Impossible de calculer l'EV: colonnes PV manquantes")
        return None

    # Identification de la colonne Jalon
    if "Jalon" not in df_pv.columns or "Jalon" not in df_va.columns:
        print("\n⚠️  Colonne 'Jalon' manquante dans PV ou VA")
        return None

    # Calcul de l'EV mois par mois
    # Les pourcentages dans VA sont cumulatifs par jalon
    # Pour chaque mois, on doit sommer l'EV de tous les jalons

    print("\n--- Debug calcul EV ---")

    va_table = _build_va_percent_table(df_va)
    pv_by_jalon = df_pv.set_index("Jalon")[colonne_montant_pv].apply(pd.to_numeric, errors="coerce").fillna(0)
    common_jalons = pv_by_jalon.index.intersection(va_table.index)
    if va_table.empty or common_jalons.empty:
        if va_table.empty:
            print("\n⚠️  Aucune colonne de dates détectée dans VA")
        else:
            print("\n⚠️  Aucun jalon commun entre PV et VA")
        return None

    pv_by_jalon = pv_by_jalon.loc[common_jalons]
    va_table = va_table.loc[common_jalons]

    ev_by_month = va_table.mul(pv_by_jalon, axis=0).sum(axis=0)
    ev_by_month = ev_by_month[ev_by_month > 0].sort_index()

    _print_ev_debug(ev_by_month, df_pv, colonne_montant_pv, va_table)

    if ev_by_month.empty:
        print("\n⚠️  Aucune donnée EV calculée")
        return None

    print(f"\n✓ Earned Value calculée pour {len(ev_by_month)} mois")
    return ev_by_month
