"""Slide Engine package.

Production-grade automated PowerPoint generation engine based on
Python 3.12, Pydantic V2, and fail-soft architecture.
"""

from slide_engine.controller import generate_presentation
from slide_engine.models import PresentationDeck

__all__ = ["generate_presentation", "PresentationDeck"]