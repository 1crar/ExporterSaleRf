"""1C integration layer."""

from importlib import import_module

__all__ = [
    "check1Cdocdeal",
    "correctord",
    "createInvoice",
    "createPTU",
    "createRTU",
    "operations1C",
    "print_docs",
    "testupdprint",
]

_EXPORTS = {
    "print_docs": ".printing",
    "check1Cdocdeal": ".documents",
    "correctord": ".documents",
    "createInvoice": ".documents",
    "createPTU": ".documents",
    "createRTU": ".documents",
    "operations1C": ".documents",
    "testupdprint": ".documents",
}


def __getattr__(name):
    if name not in _EXPORTS:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    module = import_module(_EXPORTS[name], __name__)
    value = getattr(module, name)
    globals()[name] = value
    return value
