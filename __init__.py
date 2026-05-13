"""AutoDoc Accounting package.

The package keeps the original automation logic intact while separating
Bitrix, PDF parsing, and 1C operations into explicit modules.
"""

__all__ = ["main"]

def main() -> None:
    from .app import main as run

    run()
