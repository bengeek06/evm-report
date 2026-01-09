"""
Module de calculs EVM (dépenses, earned value, projections)
"""

from .depenses import calculer_depenses_cumulees
from .earned_value import calculer_earned_value
from .planned_value import traiter_planned_value
from .projections import calculer_eac_projete, calculer_projections_automatiques

__all__ = [
    "calculer_depenses_cumulees",
    "traiter_planned_value",
    "calculer_earned_value",
    "calculer_projections_automatiques",
    "calculer_eac_projete",
]
