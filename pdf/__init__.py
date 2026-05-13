"""PDF extraction helpers.

Submodules are intentionally imported directly by callers so lightweight model
imports do not eagerly load Camelot/pdfminer.
"""

__all__ = ["extract", "fields", "recognize"]
