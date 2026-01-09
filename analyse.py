#!/usr/bin/env python3
"""
Wrapper de compatibilité pour analyse.py
Ce fichier maintient la compatibilité avec l'ancien code en important depuis les nouveaux modules.
"""

# Imports depuis les nouveaux modules
from src.calculs import (
    calculer_depenses_cumulees,
    calculer_eac_projete,
    calculer_earned_value,
    calculer_projections_automatiques,
    traiter_planned_value,
)
from src.cli import main, parser_arguments
from src.evm_io import (
    generer_tableau_comparatif,
    lire_export_sap,
    lire_forecast,
    lire_planned_value,
    lire_valeur_acquise,
)
from src.visualisation import generer_rapport_word, tracer_courbe, tracer_courbe_projections, tracer_courbe_realise

# Exposer toutes les fonctions pour la compatibilité
__all__ = [
    # CLI
    "parser_arguments",
    "main",
    # I/O
    "lire_export_sap",
    "lire_planned_value",
    "lire_valeur_acquise",
    "lire_forecast",
    "generer_tableau_comparatif",
    # Calculs
    "calculer_depenses_cumulees",
    "traiter_planned_value",
    "calculer_earned_value",
    "calculer_projections_automatiques",
    "calculer_eac_projete",
    # Visualisation
    "tracer_courbe",
    "tracer_courbe_realise",
    "tracer_courbe_projections",
    "generer_rapport_word",
]

if __name__ == "__main__":
    main()
