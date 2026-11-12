"""Reusable formula fragments: look a value up in the Catalog, People or Role Defaults sheets.

Each function returns a formula *expression* (no leading '='), so fragments can be nested.
"""
import refs
from layout import Cat, Ppl
from vocab import Applicability, Status


def text_or_blank(expr):
    """`expr` as text, or "" when it fails or is empty."""
    return f'IFERROR({expr}&"","")'


def person_value(col, person):
    """A People column value for the person named in the cell `person`."""
    return f"INDEX({refs.ppl(col)},MATCH({person},{refs.ppl(Ppl.NAME)},0))"


def title_match(software):
    """Row position of a title within the Catalog."""
    return f"MATCH({software},{refs.cat(Cat.NAME)},0)"


def title_value(col, software):
    """A Catalog column value for the title named in the cell `software`."""
    return f"INDEX({refs.cat(col)},{title_match(software)})"


def _blank_is_optional(cell):
    optional = Applicability.OPTIONAL
    return f'IFERROR(IF({cell}&""="","{optional}",{cell}),"{optional}")'


def role_default_by_slot(role, slot):
    """Role Defaults entry for catalog slot `slot` (0-based) and the role in cell `role`. Blank means Optional."""
    cell = f"INDEX({refs.ROLE_GRID},MATCH({role},{refs.ROLE_NAMES},0),{slot + 1})"
    return _blank_is_optional(cell)


def role_default_by_name(role, software):
    """Role Defaults entry for the title named in cell `software`. Blank means Optional."""
    cell = f"INDEX({refs.ROLE_GRID},MATCH({role},{refs.ROLE_NAMES},0),MATCH({software},{refs.ROLE_TITLES},0))"
    return _blank_is_optional(cell)


def not_installed_status(applicability):
    """Status of a title that is not installed, from an applicability expression."""
    return (f'IF({applicability}="{Applicability.NOT_APPLICABLE}","{Status.NOT_APPLICABLE}",'
            f'IF({applicability}="{Applicability.REQUIRED}","{Status.MISSING}","{Status.NOT_INSTALLED}"))')
