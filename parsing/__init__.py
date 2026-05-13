"""Parsing layer."""

from importlib import import_module

from .utils import convDate, convDateTo1CFormat, convSum

__all__ = [
    "convDate",
    "convDateTo1CFormat",
    "convSum",
    "extractBitrixDocInfo",
    "extractServerDocInfo",
]


def __getattr__(name):
    if name not in {"extractBitrixDocInfo", "extractServerDocInfo"}:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    module = import_module(".documents", __name__)
    value = getattr(module, name)
    globals()[name] = value
    return value
