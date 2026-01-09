"""
Module de calcul de la Earned Value
"""

import pandas as pd


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

    colonne_date_pv = None
    colonne_montant_pv = None

    for col in colonnes_date:
        if col in df_pv.columns:
            colonne_date_pv = col
            break

    for col in colonnes_montant:
        if col in df_pv.columns:
            colonne_montant_pv = col
            break

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

    # D'abord, collecter tous les mois présents
    tous_les_mois = set()
    for col in df_va.columns:
        try:
            date_col = pd.to_datetime(col)
            mois = date_col.to_period("M")
            tous_les_mois.add(mois)
        except Exception:
            continue

    tous_les_mois = sorted(tous_les_mois)

    # Pour chaque mois, calculer l'EV totale
    ev_par_mois = {}

    for mois_cible in tous_les_mois:
        ev_mois = 0

        # Parcourir chaque jalon
        for _idx, row_pv in df_pv.iterrows():
            jalon = row_pv["Jalon"]
            montant_jalon = row_pv[colonne_montant_pv]

            # Trouver le jalon correspondant dans VA
            row_va = df_va[df_va["Jalon"] == jalon]
            if row_va.empty:
                continue

            row_va = row_va.iloc[0]

            # Trouver le pourcentage d'avancement le plus récent pour ce mois
            pourcentage_actuel = 0
            for col in df_va.columns:
                try:
                    date_col = pd.to_datetime(col)
                    mois = date_col.to_period("M")

                    # Si ce mois est <= au mois cible, on prend le pourcentage
                    if mois <= mois_cible:
                        pourcentage = row_va[col]
                        if pd.notna(pourcentage):
                            pourcentage_actuel = max(pourcentage_actuel, pourcentage)
                except Exception:
                    continue

            # Calculer l'EV pour ce jalon (pourcentage est déjà en décimal: 1.0 = 100%)
            if pourcentage_actuel > 0:
                ev_jalon = montant_jalon * pourcentage_actuel
                ev_mois += ev_jalon

        if ev_mois > 0:
            ev_par_mois[mois_cible] = ev_mois

    # Afficher le détail pour debug
    for mois_cible in tous_les_mois:
        if mois_cible in ev_par_mois:
            print(f"\nMois {mois_cible}: EV totale = {ev_par_mois[mois_cible]:.2f}€")

            # Détail par jalon
            for _idx, row_pv in df_pv.iterrows():
                jalon = row_pv["Jalon"]
                montant_jalon = row_pv[colonne_montant_pv]

                row_va = df_va[df_va["Jalon"] == jalon]
                if row_va.empty:
                    continue

                row_va = row_va.iloc[0]

                pourcentage_actuel = 0
                for col in df_va.columns:
                    try:
                        date_col = pd.to_datetime(col)
                        mois = date_col.to_period("M")

                        if mois <= mois_cible:
                            pourcentage = row_va[col]
                            if pd.notna(pourcentage):
                                pourcentage_actuel = max(pourcentage_actuel, pourcentage)
                    except Exception:
                        continue

                if pourcentage_actuel > 0:
                    ev_jalon = montant_jalon * pourcentage_actuel
                    print(f"  {jalon}: {pourcentage_actuel * 100:.1f}% -> {ev_jalon:.2f}€")

    if not ev_par_mois:
        print("\n⚠️  Aucune donnée EV calculée")
        return None

    # Convertir en Series et trier
    ev_series = pd.Series(ev_par_mois).sort_index()

    print(f"\n✓ Earned Value calculée pour {len(ev_series)} mois")

    return ev_series
