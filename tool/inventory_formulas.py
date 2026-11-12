"""Formulas for each calculated Inventory column. Every function takes a row number and returns '=...'."""
from layout import FIRST_ROW, INV_LAST, STALE_DAYS, WARN_DAYS, Cat, Inv, Ppl
from lookups import not_installed_status, person_value, role_default_by_name, text_or_blank, title_value
from vocab import Applicability, InstallState, LicenseAssign, LicenseCheck, Status, VersionCheck


def _c(col, row):
    """Same-row reference with an absolute column, e.g. '$E6'."""
    return f"${col}{row}"


def _no_key(row):
    return f'{_c(Inv.KEY, row)}=""'


def _unless_installed(row, body):
    """Blank unless this row is an installed record."""
    return f'IF(OR({_no_key(row)},{_c(Inv.STATE, row)}<>"{InstallState.INSTALLED}"),"",{body})'


def key(row):
    person, software = _c(Inv.PERSON, row), _c(Inv.SOFTWARE, row)
    return f'=IF(OR({person}="",{software}=""),"",{person}&"|"&{software})'


def computer(row):
    person = _c(Inv.PERSON, row)
    return f'=IF({person}="","",{text_or_blank(person_value(Ppl.COMPUTER, person))})'


def group(row):
    software = _c(Inv.SOFTWARE, row)
    return f'=IF({software}="","",{text_or_blank(title_value(Cat.GROUP, software))})'


def applicability(row):
    person, software = _c(Inv.PERSON, row), _c(Inv.SOFTWARE, row)
    override = _c(Inv.APPL_OVERRIDE, row)
    default = role_default_by_name(person_value(Ppl.ROLE, person), software)
    return f'=IF({_no_key(row)},"",IF({override}<>"",{override},{default}))'


def expiry(row):
    """The row's own expiry if given, otherwise the catalog's (blank if neither)."""
    software, override = _c(Inv.SOFTWARE, row), _c(Inv.EXPIRY_OVERRIDE, row)
    catalog = title_value(Cat.EXPIRY, software)
    body = f'IF({override}<>"",{override},IFERROR(IF({catalog}="","",{catalog}),""))'
    return "=" + _unless_installed(row, body)


def version_check(row):
    """Numbers compare numerically (2026 vs 2025); text such as '2025 R2' must match exactly."""
    version, software = _c(Inv.VERSION, row), _c(Inv.SOFTWARE, row)
    approved = text_or_blank(title_value(Cat.APPROVED, software))
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
    return f'=IF(OR({_no_key(row)},{verified}=""),"",TODAY()-{verified})'


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


def _not_installed_or_pending(row):
    state = _c(Inv.STATE, row)
    return (f'IF({state}="{InstallState.PENDING}","{Status.PENDING}",'
            f'{not_installed_status(_c(Inv.APPLICABILITY, row))})')


def status(row):
    """The one-word verdict for the row. Duplicates and unknown states win over everything else."""
    state, row_key = _c(Inv.STATE, row), _c(Inv.KEY, row)
    keys = f"${Inv.KEY}${FIRST_ROW}:${Inv.KEY}${INV_LAST}"
    return (f'=IF({_no_key(row)},"",IF(COUNTIF({keys},{row_key})>1,"{Status.DUPLICATE}",'
            f'IF(OR({state}="",{state}="{InstallState.UNKNOWN}"),"{Status.UNKNOWN}",'
            f'IF({state}="{InstallState.INSTALLED}",{_installed_status(row)},{_not_installed_or_pending(row)}))))')
