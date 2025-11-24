"""Deprecated formula module.

The table/cell formula feature was rolled back. This stub remains only to avoid
ImportError for any lingering imports until all references are purged.
"""

def has_formulas(html: str) -> bool:  # pragma: no cover
    return False

def recompute_formulas(html: str) -> str:  # pragma: no cover
    return html

def set_cell_formula(html: str, *args, **kwargs) -> str:  # pragma: no cover
    return html
