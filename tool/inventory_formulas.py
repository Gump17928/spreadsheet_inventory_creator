"""Formulas for each Inventory column. Every function takes a row number and returns '=...'.

An Inventory row belongs to one person and one catalog title by position (see `layout.inventory_row`), so these
formulas point straight at that person's People row and that title's Catalog row.
"""
from layout import STALE_DAYS, WARN_DAYS, Cat, Inv, Ppl, catalog_row_for, people_row_for, slot_for
from lookups import catalog_at, not_installed_status, person_at, role_default_by_slot, text_or_blank
from vocab import Active, Applicability, InstallState, LicenseAssign, LicenseCheck, Status, VersionCheck

LIVE_MARK = "Y"   # in the hidden LIVE column when the row is in use


def _c(col, row):
    """Same-row reference with an absolute column, e.g. '$E6'."""
    return f"${col}{row}"


def _not_live(row):
    return f'{_c(Inv.LIVE, row)}=""'


def _unless_installed(row, body):
    """Blank unless this row is an installed record."""
    return f'IF(OR({_not_live(row)},{_c(Inv.STATE, row)}<>"{InstallState.INSTALLED}"),"",{body})'


def _mirror(cell):
    return f'=IF({cell}="","",{cell})'


def person(row):
    """The person's name, shown only on rows whose title slot is in use (spare slots stay blank)."""
    name = person_at(Ppl.NAME, people_row_for(row))
    title = catalog_at(Cat.NAME, catalog_row_for(row))
    return f'=IF(OR({name}="",{title}=""),"",{name})'


def software(row):
    return _mirror(catalog_at(Cat.NAME, catalog_row_for(row)))


def computer(row):
    name = _c(Inv.PERSON, row)
    return f'=IF({name}="","",{text_or_blank(person_at(Ppl.COMPUTER, people_row_for(row)))})'


def group(row):
    name = _c(Inv.SOFTWARE, row)
    return f'=IF({name}="","",{text_or_blank(catalog_at(Cat.GROUP, catalog_row_for(row)))})'


def live(row):
    """In use when both the person and the title exist and the person is not retired."""
    retired = f'{person_at(Ppl.ACTIVE, people_row_for(row))}="{Active.NO}"'
    return f'=IF(OR({_c(Inv.PERSON, row)}="",{_c(Inv.SOFTWARE, row)}="",{retired}),"","{LIVE_MARK}")'


def applicability(row):
    override = _c(Inv.APPL_OVERRIDE, row)
    role = person_at(Ppl.ROLE, people_row_for(row))
    return f'=IF({_not_live(row)},"",IF({override}<>"",{override},{role_default_by_slot(role, slot_for(row))}))'


def expiry(row):
    """The row's own expiry if given, otherwise the catalog's (blank if neither)."""
    override = _c(Inv.EXPIRY_OVERRIDE, row)
    catalog = catalog_at(Cat.EXPIRY, catalog_row_for(row))
    body = f'IF({override}<>"",{override},IFERROR(IF({catalog}="","",{catalog}),""))'
    return "=" + _unless_installed(row, body)


def version_check(row):
    """Numbers compare numerically (2026 vs 2025); text such as '2025 R2' must match exactly."""
    version = _c(Inv.VERSION, row)
    approved = text_or_blank(catalog_at(Cat.APPROVED, catalog_row_for(row)))
    comparison = (f'IFERROR(IF(VALUE({version})>=VALUE({approved}),"{VersionCheck.CURRENT}","{VersionCheck.OUTDATED}"),'
                  f'IF({version}&""={approved},"{VersionCheck.CURRENT}","{VersionCheck.CHECK}"))')
    body = (f'IF({approved}="","{VersionCheck.NOT_SET}",'
            f'IF({version}="","{VersionCheck.NO_VERSION}",{comparison}))')
    return "=" + _unless_installed(row, body)


def license_check(row):
    assignment, expires = _c(Inv.ASSIGNMENT, row), _c(Inv.EXPIRY, row)
    dated = (f'IF({expires}="","{LicenseCheck.VALID}",IF({expires}<TODAY(),"{LicenseCheck.EXPIRED}",'
             f'IF({expires}<=TODAY()+{WARN_DAYS},"{LicenseCheck.EXPIRING}","{LicenseCheck.VALID}")))')
    body = (f'IF({assignment}="{LicenseAssign.UNLICENSED}","{LicenseCheck.UNLICENSED}",'
            f'IF({assignment}="{LicenseAssign.FREE}","{LicenseCheck.NOT_REQUIRED}",'
            f'IF(OR({assignment}="",{assignment}="{LicenseAssign.UNKNOWN}"),"{LicenseCheck.UNVERIFIED}",{dated})))')
    return "=" + _unless_installed(row, body)


def days_since_verified(row):
    verified = _c(Inv.VERIFIED_ON, row)
    return f'=IF(OR({_not_live(row)},{verified}=""),"",TODAY()-{verified})'


def _installed_status(row):
    """Status of an installed row, worst problem first."""
    appl, lic, ver = _c(Inv.APPLICABILITY, row), _c(Inv.LICENSE_CHECK, row), _c(Inv.VERSION_CHECK, row)
    verified, days = _c(Inv.VERIFIED_ON, row), _c(Inv.DAYS_SINCE, row)
    unverified = (f'OR({lic}="{LicenseCheck.UNVERIFIED}",{ver}="{VersionCheck.CHECK}",'
                  f'{ver}="{VersionCheck.NO_VERSION}",{verified}="",{days}>{STALE_DAYS})')
    return (f'IF({appl}="{Applicability.NOT_APPLICABLE}","{Status.REVIEW}",'
            f'IF(OR({lic}="{LicenseCheck.EXPIRED}",{lic}="{LicenseCheck.UNLICENSED}"),"{Status.LICENSE_ISSUE}",'
            f'IF({ver}="{VersionCheck.OUTDATED}","{Status.OUTDATED}",'
            f'IF({lic}="{LicenseCheck.EXPIRING}","{Status.EXPIRING}",'
            f'IF({unverified},"{Status.VERIFY}","{Status.INSTALLED}")))))')


def status(row):
    """The one-word verdict. A blank or 'Not Installed' state follows the person's applicability for the title."""
    state = _c(Inv.STATE, row)
    return (f'=IF({_not_live(row)},"",'
            f'IF({state}="{InstallState.UNKNOWN}","{Status.UNKNOWN}",'
            f'IF({state}="{InstallState.INSTALLED}",{_installed_status(row)},'
            f'IF({state}="{InstallState.PENDING}","{Status.PENDING}",'
            f'{not_installed_status(_c(Inv.APPLICABILITY, row))}))))')
