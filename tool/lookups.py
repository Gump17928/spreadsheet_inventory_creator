"""Reusable formula fragments: read a value from the Catalog, People or Role Defaults sheets.

Each function returns a formula *expression* (no leading '='), so fragments can be nested.
"""
import refs
from layout import Ppl, Sheet
from vocab import Applicability, Status


def text_or_blank(expr):
    """`expr` as text, or "" when it fails or is empty."""
    return f'IFERROR({expr}&"","")'


def person_at(col, people_row):
    """A People value by row; Inventory and Matrix rows are tied to People rows by position."""
    return refs.at(Sheet.PEOPLE, col, people_row)


def catalog_at(col, catalog_row):
    """A Catalog value by row; Inventory rows are tied to Catalog rows by position."""
    return refs.at(Sheet.CATALOG, col, catalog_row)


def person_by_name(col, person):
    """A People column value for the person named in the cell `person` (used where a name is picked)."""
    return f"INDEX({refs.ppl(col)},MATCH({person},{refs.ppl(Ppl.NAME)},0))"


def role_default_by_slot(role, slot):
    """Role Defaults entry for catalog slot `slot` (0-based) and the role in cell `role`. Blank means Optional."""
    optional = Applicability.OPTIONAL
    cell = f"INDEX({refs.ROLE_GRID},MATCH({role},{refs.ROLE_NAMES},0),{slot + 1})"
    return f'IFERROR(IF({cell}&""="","{optional}",{cell}),"{optional}")'


def not_installed_status(applicability):
    """Status of a title that is not installed, from an applicability expression."""
    return (f'IF({applicability}="{Applicability.NOT_APPLICABLE}","{Status.NOT_APPLICABLE}",'
            f'IF({applicability}="{Applicability.REQUIRED}","{Status.MISSING}","{Status.NOT_INSTALLED}"))')
