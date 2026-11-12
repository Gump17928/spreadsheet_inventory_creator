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
    DUPLICATE = "Duplicate"


@dataclass(frozen=True)
class StatusStyle:
    fill: str
    font: str
    bold: bool
    meaning: str
    action: str


_COLOURS = {   # status -> (cell fill, font colour, bold)
    Status.INSTALLED: ("D9EAD3", "1E5631", False),
    Status.OUTDATED: ("FFE699", "7F6000", False),
    Status.EXPIRING: ("FCE4D6", "9C4A00", False),
    Status.LICENSE_ISSUE: ("C0504D", "FFFFFF", True),
    Status.VERIFY: ("DDEBF7", "1F4E78", False),
    Status.PENDING: ("E4DFEC", "5B3F8C", False),
    Status.MISSING: ("FFC7CE", "9C0006", True),
    Status.REVIEW: ("F8D7F0", "8E2A78", False),
    Status.UNKNOWN: ("FFF9DB", "7F6000", False),
    Status.NOT_INSTALLED: ("F2F2F2", "595959", False),
    Status.NOT_APPLICABLE: ("D9D9D9", "7F7F7F", False),
    Status.DUPLICATE: ("404040", "FFFFFF", True),
}
_TEXT = {      # status -> (meaning, what to do)
    Status.INSTALLED: ("Installed, current, licensed and recently verified.", "Nothing."),
    Status.OUTDATED: ("Installed version is older than the approved version.", "Schedule an upgrade."),
    Status.EXPIRING: ("License expires within the warning window (see Lists).", "Renew or plan a replacement."),
    Status.LICENSE_ISSUE: ("Installed, but the license is expired or the install is unlicensed.", "Fix immediately."),
    Status.VERIFY: ("Installed, but the data is stale or the version/license is unverified.",
                    "Re-check the machine and update the row."),
    Status.PENDING: ("Install requested or in progress.", "Follow up until installed."),
    Status.MISSING: ("Required for this person's role but not installed.", "Install it."),
    Status.REVIEW: ("Installed although it is marked Not Applicable.", "Remove it or change the applicability."),
    Status.UNKNOWN: ("A record exists but the install state is unknown or blank.", "Check the machine."),
    Status.NOT_INSTALLED: ("Not installed and optional for this person.", "Nothing."),
    Status.NOT_APPLICABLE: ("Not applicable to this person's role.", "Nothing."),
    Status.DUPLICATE: ("The same person + software appears on more than one row.", "Delete the extra row."),
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
