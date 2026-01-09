"""
Module de calcul des projections EVM
"""

import pandas as pd


def _compute_indices(*, ac_actuel: float, ev_actuel: float, pv_actuel: float) -> tuple[float, float]:
    cpi = ev_actuel / ac_actuel if ac_actuel > 0 else 1
    spi = ev_actuel / pv_actuel if pv_actuel > 0 else 1
    return cpi, spi


def _estimate_finish_date(dernier_mois: pd.Period, pv_cumulee: pd.Series | None, spi: float) -> pd.Period:
    if pv_cumulee is not None and spi > 0:
        dernier_mois_pv = pv_cumulee.index[-1]
        mois_restants_plan = (dernier_mois_pv.to_timestamp().year - dernier_mois.to_timestamp().year) * 12 + (
            dernier_mois_pv.to_timestamp().month - dernier_mois.to_timestamp().month
        )
        mois_restants_proj = int(mois_restants_plan / spi)
        return dernier_mois + mois_restants_proj
    return dernier_mois + 12


def _build_projection_series(
    depenses_cumulees: pd.Series,
    *,
    dernier_mois: pd.Period,
    ac_actuel: float,
    reste: float,
    date_fin_proj: pd.Period,
) -> pd.Series:
    mois_projection = pd.period_range(start=dernier_mois, end=date_fin_proj, freq="M")

    serie = pd.Series(dtype=float)
    for mois in depenses_cumulees.index:
        serie[mois] = depenses_cumulees[mois]

    if len(mois_projection) > 1:
        nb_mois = len(mois_projection) - 1
        increment = reste / nb_mois if nb_mois > 0 else 0
        valeur_actuelle = ac_actuel
        for mois in mois_projection:
            if mois > dernier_mois:
                valeur_actuelle += increment
                serie[mois] = valeur_actuelle  # type: ignore[index]

    return serie.sort_index()


def _parse_forecast_eac_par_jalon(df_forecast: pd.DataFrame) -> dict:
    eac_par_jalon: dict = {}
    for _, row in df_forecast.iterrows():
        jalon = row["Jalon"]
        date_projetee = row["Date projetée"]
        eac = row["EAC (€)"]

        if isinstance(date_projetee, pd.Timestamp):
            mois_proj = date_projetee.to_period("M")
        else:
            mois_proj = pd.to_datetime(date_projetee).to_period("M")

        eac_par_jalon[jalon] = {"date": mois_proj, "eac": eac}
    return eac_par_jalon


def _build_eac_series(
    ev_cumulee: pd.Series,
    eac_par_jalon: dict,
) -> pd.Series | None:
    dernier_mois_ev = ev_cumulee.index[-1]
    ev_finale = ev_cumulee.iloc[-1]

    mois_projection = sorted({info["date"] for info in eac_par_jalon.values()})
    if not mois_projection:
        return None

    eac_par_mois: dict[pd.Period, float] = {}
    for mois in mois_projection:
        if mois <= dernier_mois_ev:
            if mois in ev_cumulee.index:
                eac_par_mois[mois] = float(ev_cumulee[mois])
            continue

        eac_cumul = float(ev_finale)
        for _jalon, info in eac_par_jalon.items():
            if info["date"] <= mois:
                eac_cumul += float(info["eac"])
        eac_par_mois[mois] = eac_cumul
        print(f"Mois {mois}: EAC projeté = {eac_cumul:.2f}€")

    if not eac_par_mois:
        return None

    return pd.Series(eac_par_mois).sort_index()


def calculer_projections_automatiques(depenses_cumulees, ev_cumulee, pv_cumulee, _df_pv):
    """
    Calcule automatiquement les projections EAC selon différentes méthodes EVM
    """
    if ev_cumulee is None or len(ev_cumulee) == 0:
        return None

    print("\n--- Calcul des projections automatiques ---")

    # Valeurs actuelles
    ac_actuel = depenses_cumulees.iloc[-1]
    ev_actuel = ev_cumulee.iloc[-1]
    dernier_mois = depenses_cumulees.index[-1]

    # Budget à complétion (BAC)
    bac = pv_cumulee.iloc[-1] if pv_cumulee is not None else 0

    pv_actuel = pv_cumulee[dernier_mois] if pv_cumulee is not None and dernier_mois in pv_cumulee.index else 0
    cpi, spi = _compute_indices(ac_actuel=ac_actuel, ev_actuel=ev_actuel, pv_actuel=pv_actuel)

    # Calcul des EAC selon différentes méthodes
    projections = {}

    # Méthode 1 : Performance actuelle continue (CPI uniquement)
    eac_cpi = bac / cpi if cpi > 0 else bac
    projections["CPI"] = {
        "eac": eac_cpi,
        "formule": "BAC / CPI",
        "description": "Performance des coûts actuelle se poursuit",
        "reste": eac_cpi - ac_actuel,
    }

    # Méthode 2 : Performance coûts ET délais (CPI et SPI)
    eac_cpi_spi = ac_actuel + ((bac - ev_actuel) / (cpi * spi)) if (cpi * spi) > 0 else bac
    projections["CPI_SPI"] = {
        "eac": eac_cpi_spi,
        "formule": "AC + [(BAC-EV) / (CPI × SPI)]",
        "description": "Performance coûts ET délais se poursuit",
        "reste": eac_cpi_spi - ac_actuel,
    }

    # Méthode 3 : Reste comme prévu
    eac_reste_plan = ac_actuel + (bac - ev_actuel)
    projections["RESTE_PLAN"] = {
        "eac": eac_reste_plan,
        "formule": "AC + (BAC - EV)",
        "description": "Le reste se déroule comme prévu initialement",
        "reste": bac - ev_actuel,
    }

    date_fin_proj = _estimate_finish_date(dernier_mois, pv_cumulee, spi)

    series_projections = {}
    for methode, data in projections.items():
        series_projections[methode] = _build_projection_series(
            depenses_cumulees,
            dernier_mois=dernier_mois,
            ac_actuel=float(ac_actuel),
            reste=float(data["reste"]),
            date_fin_proj=date_fin_proj,
        )

        print(f"Méthode {methode}: EAC = {data['eac']:,.2f} € (Reste: {data['reste']:,.2f} €)")

    return projections, series_projections, date_fin_proj


def calculer_eac_projete(ev_cumulee, df_forecast, df_pv):
    """
    Calcule l'EAC projeté en combinant l'EV actuelle et les projections
    """
    if ev_cumulee is None or df_forecast is None or df_pv is None:
        return None

    print("\n--- Calcul EAC projeté ---")

    # Identifier les colonnes
    if "Date projetée" not in df_forecast.columns or "EAC (€)" not in df_forecast.columns:
        print("⚠️  Colonnes manquantes dans forecast.xlsx")
        return None

    eac_par_jalon = _parse_forecast_eac_par_jalon(df_forecast)
    eac_series = _build_eac_series(ev_cumulee, eac_par_jalon)
    if eac_series is None:
        return None

    print(f"✓ EAC projeté calculé pour {len(eac_series)} mois")

    return eac_series, eac_par_jalon
