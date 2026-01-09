"""
Module CLI pour l'analyse EVM
"""

import argparse
from pathlib import Path

from src.calculs import (
    calculer_depenses_cumulees,
    calculer_eac_projete,
    calculer_earned_value,
    calculer_projections_automatiques,
    traiter_planned_value,
)
from src.io import generer_tableau_comparatif, lire_export_sap, lire_forecast, lire_planned_value, lire_valeur_acquise
from src.visualisation import generer_rapport_word, tracer_courbe_projections, tracer_courbe_realise


def parser_arguments():
    """
    Parse les arguments de la ligne de commande
    """
    parser = argparse.ArgumentParser(
        description="Analyse EVM (Earned Value Management) - Analyse des dépenses et projections",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemples d'utilisation:
  %(prog)s
  %(prog)s --sap EXPORT.xlsx --pv pv.xlsx --va va.xlsx --forecast forecast.xlsx
  %(prog)s --sap mes_depenses.xlsx --output graphique_evm.png
  %(prog)s --word rapport_evm.docx

Fichiers requis:
  - Fichier SAP : Export SAP des dépenses avec colonnes "Date de la pièce" et "Val./Devise objet"
  - Fichier PV  : Planned Value avec colonnes "Jalon", "Date", "Montant planifié"
  - Fichier VA  : Valeur Acquise avec colonnes "Jalon" et pourcentages mensuels
  - Fichier Forecast : Projections avec colonnes "Jalon", "Date projetée", "EAC (€)"

Rapport Word:
  L'option --word génère un rapport complet contenant définitions EVM, tableau, graphique
  et conclusion. Les fichiers intermédiaires (PNG, CSV, XLSX) sont automatiquement supprimés.
        """,
    )

    parser.add_argument(
        "--sap", default="EXPORT.XLSX", help="Fichier Excel d'export SAP des dépenses (défaut: EXPORT.XLSX)"
    )

    parser.add_argument("--pv", default="pv.xlsx", help="Fichier Excel de la Planned Value (défaut: pv.xlsx)")

    parser.add_argument("--va", default="va.xlsx", help="Fichier Excel de la Valeur Acquise (défaut: va.xlsx)")

    parser.add_argument(
        "--forecast", default="forecast.xlsx", help="Fichier Excel des projections (défaut: forecast.xlsx)"
    )

    parser.add_argument(
        "--output", default="analyse_evm.png", help="Fichier de sortie pour le graphique (défaut: analyse_evm.png)"
    )

    parser.add_argument(
        "--tableau", default="tableau_evm", help="Nom de base pour les fichiers tableau (défaut: tableau_evm)"
    )

    parser.add_argument(
        "--word", help="Génère un rapport Word complet (ex: rapport_evm.docx) et supprime les fichiers intermédiaires"
    )

    return parser.parse_args()


def main():
    """
    Fonction principale
    """
    # Parser les arguments
    args = parser_arguments()

    print("=== Analyse des dépenses SAP ===\n")
    print("Fichiers d'entrée:")
    print(f"  - SAP:      {args.sap}")
    print(f"  - PV:       {args.pv}")
    print(f"  - VA:       {args.va}")
    print(f"  - Forecast: {args.forecast}")
    print("\nFichiers de sortie:")
    print(f"  - Graphique: {args.output}")
    print(f"  - Tableaux:  {args.tableau}.csv / {args.tableau}.xlsx\n")

    # Lecture du fichier Excel
    df = lire_export_sap(args.sap)

    if df is None:
        return

    # Affichage d'un échantillon des données
    print("\nAperçu des données:")
    print(df.head())

    # Colonnes spécifiques pour cet export SAP
    colonne_date = "Date de la pièce"
    colonne_montant = "Val./Devise objet"

    # Vérification que les colonnes existent
    if colonne_date not in df.columns or colonne_montant not in df.columns:
        print("\n⚠️  Colonnes requises non trouvées.")
        print(f"Colonnes disponibles: {df.columns.tolist()}")
        return

    print(f"\n✓ Colonne date utilisée: {colonne_date}")
    print(f"✓ Colonne montant utilisée: {colonne_montant}")

    # Calcul des dépenses cumulées
    depenses_cumulees = calculer_depenses_cumulees(df, colonne_date, colonne_montant)

    print("\nDépenses cumulées par mois (AC):")
    print(depenses_cumulees)

    # Lecture et traitement de la Planned Value
    df_pv = lire_planned_value(args.pv)
    result = traiter_planned_value(df_pv)

    pv_cumulee = None
    jalons = None
    if result is not None:
        if isinstance(result, tuple):
            pv_cumulee, jalons = result
        else:
            pv_cumulee = result
        print("\nPlanned Value cumulée par mois (PV):")
        print(pv_cumulee)

    # Lecture et calcul de l'Earned Value
    df_va = lire_valeur_acquise(args.va)
    ev_cumulee = calculer_earned_value(df_pv, df_va)

    if ev_cumulee is not None:
        print("\nEarned Value cumulée par mois (EV):")
        print(ev_cumulee)

    # Calcul automatique des projections EAC avec différentes méthodes
    result_proj = calculer_projections_automatiques(depenses_cumulees, ev_cumulee, pv_cumulee, df_pv)

    if result_proj is not None:
        projections_data, series_projections, date_fin_proj = result_proj

        # Convertir au format attendu par les fonctions de tracé
        projections = {}
        for methode in ["CPI", "CPI_SPI", "RESTE_PLAN"]:
            if methode in projections_data and methode in series_projections:
                projections[methode.lower()] = {
                    "series": series_projections[methode],
                    "eac": projections_data[methode]["eac"],
                    "date": date_fin_proj.to_timestamp() if hasattr(date_fin_proj, "to_timestamp") else date_fin_proj,
                }
    else:
        projections = {}

    # Si un fichier forecast est fourni, l'ajouter aux projections
    df_forecast = lire_forecast(args.forecast)
    if df_forecast is not None and ev_cumulee is not None:
        result = calculer_eac_projete(ev_cumulee, df_forecast, df_pv)
        if result is not None:
            eac_projete, _ = result
            # Ajouter la projection forecast au dictionnaire
            projections["forecast"] = {
                "series": eac_projete,
                "eac": eac_projete.iloc[-1] if len(eac_projete) > 0 else 0,
                "date": eac_projete.index[-1].to_timestamp() if len(eac_projete) > 0 else None,
            }
            print("\nEAC projeté (forecast manuel) par mois:")
            print(eac_projete)

    # Générer le tableau comparatif
    df_tableau = generer_tableau_comparatif(depenses_cumulees, pv_cumulee, ev_cumulee, args.tableau)

    # Tracer les deux graphiques séparés
    fichier_realise = args.output.replace(".png", "_realise.png")
    fichier_projections = args.output.replace(".png", "_projections.png")

    tracer_courbe_realise(depenses_cumulees, pv_cumulee, jalons, ev_cumulee, fichier_realise)
    tracer_courbe_projections(depenses_cumulees, ev_cumulee, projections, fichier_projections)

    # Générer le rapport Word si demandé
    if args.word:
        generer_rapport_word(
            args.word,
            df_tableau,
            fichier_realise,
            depenses_cumulees,
            pv_cumulee,
            ev_cumulee,
            projections,
            fichier_projections,
        )

        # Supprimer les fichiers intermédiaires
        print("\n=== Nettoyage des fichiers intermédiaires ===")
        fichiers_a_supprimer = [fichier_realise, fichier_projections, f"{args.tableau}.csv", f"{args.tableau}.xlsx"]

        for fichier in fichiers_a_supprimer:
            fichier_path = Path(fichier)
            if fichier_path.exists():
                fichier_path.unlink()
                print(f"✓ {fichier} supprimé")

    print("\n✓ Analyse terminée avec succès!")


if __name__ == "__main__":
    main()
