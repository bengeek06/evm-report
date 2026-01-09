"""
Module de visualisation (graphiques et rapports Word)
"""

from .graphiques import tracer_courbe, tracer_courbe_projections, tracer_courbe_realise
from .rapport_word import generer_rapport_word

__all__ = [
    "tracer_courbe",
    "tracer_courbe_realise",
    "tracer_courbe_projections",
    "generer_rapport_word",
]
