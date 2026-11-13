"""The fixed words the workbook's formulas and dropdowns agree on, and how each status looks."""
from dataclasses import dataclass
from enum import StrEnum


class InstallState(StrEnum):
    INSTALLED = "Installed"
    NOT_INSTALLED = "Not Installed"
    PENDING = "Pending"
    UNKNOWN = "Unknown"


class Applicability(StrEnum):
    REQUIRED = "Required"
    OPTIONAL = "Optional"
    NOT_APPLICABLE = "Not Applicable"


class LicenseAssign(StrEnum):
    NAMED = "Named User Seat"
    FLOATING = "Network (Floating)"
    STANDALONE = "Standalone Key"
    SITE = "Site / Enterprise"
    FREE = "Free / Not Required"
    UNLICENSED = "Unlicensed"
    UNKNOWN = "Unknown"


class LicenseModel(StrEnum):
    NAMED = "Named User (Subscription)"
    STANDALONE = "Standalone (Perpetual)"
    FLOATING = "Network / Floating"
    SITE = "Site / Enterprise"
    FREE = "Free / Open Source"
    OTHER = "Other"


class VersionCheck(StrEnum):
    CURRENT = "Current"
    OUTDATED = "Outdated"
    CHECK = "Check"
    NO_VERSION = "No version"
    NOT_SET = "n/a"


class LicenseCheck(StrEnum):
    VALID = "Valid"
    EXPIRING = "Expiring"
    EXPIRED = "Expired"
    UNLICENSED = "Unlicensed"
    UNVERIFIED = "Unverified"
    NOT_REQUIRED = "Not Required"
    NO_EXPIRY = "No expiry set"


class SeatStatus(StrEnum):
    OK = "OK"
    OVER = "Over-allocated"
    MONITOR = "Installs > seats (monitor)"
    NOT_TRACKED = "Not tracked"


class Status(StrEnum):
    INSTALLED = "Installed"
    OUTDATED = "Outdated"
    EXPIRING = "Expiring"
    LICENSE_ISSUE = "License Issue"
    VERIFY = "Verify"
    PENDING = "Pending"
    MISSING = "Missing"
    REVIEW = "Review"
    UNKNOWN = "Unknown"
    NOT_INSTALLED = "Not Installed"
    NOT_APPLICABLE = "N/A"


class Active(StrEnum):
    YES = "Yes"
    NO = "No"


@dataclass(frozen=True)
class StatusStyle:
    fill: str
    font: str
    bold: bool
    symbol: str      # shown before the status text by the conditional format, so status reads without colour
    meaning: str
    action: str


_COLOURS = {   # status -> (cell fill, font colour, bold, symbol)
    Status.INSTALLED: ("DCE4C6", "2F3D12", False, "●"),
    Status.OUTDATED: ("EFDCA0", "4F3A05", False, "▲"),
    Status.EXPIRING: ("EBC7A0", "5E2E08", False, "▲"),
    Status.LICENSE_ISSUE: ("8E1B1B", "FFFFFF", True, "✕"),
    Status.VERIFY: ("D3DAE2", "26323F", False, "◆"),
    Status.PENDING: ("DCD6E6", "3F3558", False, "…"),
    Status.MISSING: ("B3202A", "FFFFFF", True, "✕"),
    Status.REVIEW: ("E6D0DA", "5A2342", False, "▲"),
    Status.UNKNOWN: ("EFE8D2", "4D4524", False, "?"),
    Status.NOT_INSTALLED: ("F6F3EE", "6B6254", False, "○"),
    Status.NOT_APPLICABLE: ("DAD6CF", "6E6A63", False, "–"),
}
_TEXT = {      # status -> (meaning, what to do)
    Status.INSTALLED: ("Installed, current, licensed and recently verified.", "Nothing."),
    Status.OUTDATED: ("Installed version is older than the approved version.", "Schedule an upgrade."),
    Status.EXPIRING: ("License expires within the warning window (see Lists).", "Renew or plan a replacement."),
    Status.LICENSE_ISSUE: ("Installed, but the license is expired or the install is unlicensed.", "Fix immediately."),
    Status.VERIFY: ("Installed, but the data is stale or the version/license is unverified.",
                    "Re-check the machine and update the row."),
    Status.PENDING: ("Install requested or in progress.", "Follow up until installed."),
    Status.MISSING: ("Required for this person's role and not installed (the default for a blank row).", "Install it."),
    Status.REVIEW: ("Installed although it is marked Not Applicable.", "Remove it or change the applicability."),
    Status.UNKNOWN: ("Install state was set to Unknown.", "Check the machine."),
    Status.NOT_INSTALLED: ("Not installed and optional for this person (the default for a blank row).", "Nothing."),
    Status.NOT_APPLICABLE: ("Not applicable to this person's role.", "Nothing."),
}
STATUS_STYLE = {status: StatusStyle(*_COLOURS[status], *_TEXT[status]) for status in Status}

# "Have" = the software is on the machine, whatever its condition.
HAVE = [Status.INSTALLED, Status.OUTDATED, Status.EXPIRING, Status.LICENSE_ISSUE, Status.VERIFY]
ATTENTION = [Status.OUTDATED, Status.EXPIRING, Status.LICENSE_ISSUE, Status.VERIFY,
             Status.MISSING, Status.REVIEW, Status.UNKNOWN]
DASHBOARD_STATUSES = HAVE + [Status.PENDING, Status.MISSING, Status.REVIEW, Status.UNKNOWN,
                             Status.NOT_INSTALLED, Status.NOT_APPLICABLE]


def values(enum_class):
    """The members of a StrEnum as plain strings, in definition order."""
    return [member.value for member in enum_class]
