"""
Module de calcul des projections EVM
"""

import pandas as pd


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

    # Calcul des indices de performance
    cpi = ev_actuel / ac_actuel if ac_actuel > 0 else 1

    pv_actuel = pv_cumulee[dernier_mois] if pv_cumulee is not None and dernier_mois in pv_cumulee.index else 0
    spi = ev_actuel / pv_actuel if pv_actuel > 0 else 1

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

    # Estimation de la date de fin basée sur SPI
    if pv_cumulee is not None and spi > 0:
        dernier_mois_pv = pv_cumulee.index[-1]
        mois_restants_plan = (dernier_mois_pv.to_timestamp().year - dernier_mois.to_timestamp().year) * 12 + (
            dernier_mois_pv.to_timestamp().month - dernier_mois.to_timestamp().month
        )
        mois_restants_proj = int(mois_restants_plan / spi)
        date_fin_proj = dernier_mois + mois_restants_proj
    else:
        date_fin_proj = dernier_mois + 12  # Par défaut +12 mois

    # Générer les séries temporelles pour chaque projection
    mois_projection = pd.period_range(start=dernier_mois, end=date_fin_proj, freq="M")

    series_projections = {}
    for methode, data in projections.items():
        # Série avec AC historique + projection linéaire jusqu'à EAC
        serie = pd.Series(dtype=float)

        # Ajouter l'historique AC
        for mois in depenses_cumulees.index:
            serie[mois] = depenses_cumulees[mois]

        # Ajouter la projection linéaire
        if len(mois_projection) > 1:
            reste = data["reste"]
            nb_mois = len(mois_projection) - 1
            increment = reste / nb_mois if nb_mois > 0 else 0

            valeur_actuelle = ac_actuel
            for _i, mois in enumerate(mois_projection):
                if mois > dernier_mois:
                    valeur_actuelle += increment
                    serie[mois] = valeur_actuelle  # type: ignore[index]

        serie = serie.sort_index()
        series_projections[methode] = serie

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

    # Créer un dictionnaire des montants EAC par jalon
    eac_par_jalon = {}
    for _, row in df_forecast.iterrows():
        jalon = row["Jalon"]
        date_projetee = row["Date projetée"]
        eac = row["EAC (€)"]

        # Convertir en period si nécessaire
        if isinstance(date_projetee, pd.Timestamp):
            mois_proj = date_projetee.to_period("M")
        else:
            mois_proj = pd.to_datetime(date_projetee).to_period("M")

        eac_par_jalon[jalon] = {"date": mois_proj, "eac": eac}

    # Récupérer tous les jalons avec leur montant planifié
    jalons_pv = {}
    for _, row in df_pv.iterrows():
        jalon = row["Jalon"]
        montant = row["Montant planifié"] if "Montant planifié" in df_pv.columns else 0
        jalons_pv[jalon] = montant

    # Créer une série temporelle pour l'EAC projeté
    # On commence avec l'EV actuelle
    dernier_mois_ev = ev_cumulee.index[-1]
    ev_finale = ev_cumulee.iloc[-1]

    # Collecter tous les mois de projection
    mois_projection = set()
    for info in eac_par_jalon.values():
        mois_projection.add(info["date"])

    mois_projection = sorted(mois_projection)

    # Calculer l'EAC cumulé mois par mois
    eac_par_mois = {}

    for mois in mois_projection:
        if mois <= dernier_mois_ev:
            # Pour les mois passés, on garde l'EV
            if mois in ev_cumulee.index:
                eac_par_mois[mois] = ev_cumulee[mois]
        else:
            # Pour les mois futurs, on accumule l'EAC
            eac_cumul = ev_finale
            for _jalon, info in eac_par_jalon.items():
                if info["date"] <= mois:
                    eac_cumul += info["eac"]
            eac_par_mois[mois] = eac_cumul
            print(f"Mois {mois}: EAC projeté = {eac_cumul:.2f}€")

    if not eac_par_mois:
        return None

    eac_series = pd.Series(eac_par_mois).sort_index()
    print(f"✓ EAC projeté calculé pour {len(eac_series)} mois")

    return eac_series, eac_par_jalon
