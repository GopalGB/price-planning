"""Calculation adapters and process-pool primitives."""

from .formualizer import FormualizerEngine
from .protocol import Engine

FormualizerAdapter = FormualizerEngine

__all__ = ["Engine", "FormualizerEngine", "FormualizerAdapter"]
