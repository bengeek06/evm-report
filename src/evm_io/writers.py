"""
Module d'écriture de fichiers (CSV, Excel)
"""

import pandas as pd

# Constantes pour les noms de colonnes
COL_EV = "EV (Valeur acquise)"
COL_PV = "PV (Budget prévu)"
COL_AC = "AC (Dépenses réelles)"


def generer_tableau_comparatif(depenses_cumulees, pv_cumulee=None, ev_cumulee=None, nom_base="tableau_evm"):
    """
    Génère un tableau comparatif des valeurs PV, AC et EV
    """
    # Créer un DataFrame avec toutes les périodes
    toutes_periodes = set(depenses_cumulees.index)
    if pv_cumulee is not None:
        toutes_periodes.update(pv_cumulee.index)
    if ev_cumulee is not None:
        toutes_periodes.update(ev_cumulee.index)

    toutes_periodes = sorted(toutes_periodes)

    # Créer le tableau
    data = {
        "Mois": [str(p) for p in toutes_periodes],
        COL_AC: [depenses_cumulees.get(p, 0) for p in toutes_periodes],
    }

    if pv_cumulee is not None:
        data[COL_PV] = [pv_cumulee.get(p, 0) for p in toutes_periodes]

    if ev_cumulee is not None:
        data[COL_EV] = [ev_cumulee.get(p, 0) for p in toutes_periodes]

    df_tableau = pd.DataFrame(data)

    # Calculer les écarts si toutes les valeurs sont disponibles
    if pv_cumulee is not None and ev_cumulee is not None:
        df_tableau["SV (Schedule Variance)"] = df_tableau.get(COL_EV, 0) - df_tableau.get(COL_PV, 0)
        df_tableau["CV (Cost Variance)"] = df_tableau.get(COL_EV, 0) - df_tableau[COL_AC]

    # Sauvegarder en CSV et Excel
    csv_file = f"{nom_base}.csv"
    xlsx_file = f"{nom_base}.xlsx"
    df_tableau.to_csv(csv_file, index=False)
    df_tableau.to_excel(xlsx_file, index=False)

    print("\n=== TABLEAU COMPARATIF EVM ===")
    print(df_tableau.to_string(index=False))
    print(f"\n✓ Tableau sauvegardé: {csv_file} et {xlsx_file}")

    return df_tableau
