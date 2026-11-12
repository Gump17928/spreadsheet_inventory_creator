"""Generates the fake installation records for the example workbook. Pure: no workbook, no I/O."""
import random
from dataclasses import dataclass
from datetime import date, timedelta

from demo_data import PEOPLE, DemoPerson
from starter_data import TITLES, SoftwareTitle, role_default
from vocab import Applicability, InstallState, LicenseAssign, LicenseModel

SEED = 7
SKIP_OPTIONAL = 0.55        # optional titles most people simply don't have
SKIP_REQUIRED = 0.12        # a few required titles stay uninstalled so "Missing" shows up
OUTDATED_CHANCE = 0.15
PENDING_CHANCE = 0.07
UNKNOWN_CHANCE = 0.04
UNLICENSED_CHANCE = 0.02
UNVERIFIED_CHANCE = 0.04
INSTALL_AGE_DAYS = (30, 700)
VERIFIED_AGE_CHOICES = (10, 20, 30, 45, 60, 75, 90, 120, 150, 260)   # only 260 is past the 180-day stale limit
VERIFIERS = ("IT Admin", "S. Patel", "Self-reported")

ASSIGNMENT_FOR_MODEL = {
    LicenseModel.NAMED: LicenseAssign.NAMED,
    LicenseModel.STANDALONE: LicenseAssign.STANDALONE,
    LicenseModel.FLOATING: LicenseAssign.FLOATING,
    LicenseModel.SITE: LicenseAssign.SITE,
    LicenseModel.FREE: LicenseAssign.FREE,
}
PREVIOUS_VERSION = {"2026": "2025", "2025": "2024", "2025 R2": "2024 R2", "R2025b": "R2024b",
                    "3.12": "3.11", "11": "10", "21": "20"}


@dataclass
class InstallRecord:
    person: str
    software: str
    state: str
    version: str = ""
    installed_on: date | None = None
    assignment: str = ""
    applicability_override: str = ""
    verified_on: date | None = None
    verified_by: str = ""
    notes: str = ""


def _wanted(rnd, applicability):
    """Whether this person gets a record for a title of this applicability."""
    if applicability == Applicability.NOT_APPLICABLE:
        return False
    skip = SKIP_OPTIONAL if applicability == Applicability.OPTIONAL else SKIP_REQUIRED
    return rnd.random() >= skip


def _state_and_version(rnd, title):
    """Mostly installed at the approved version; sometimes pending, unknown or one release behind."""
    if rnd.random() < PENDING_CHANCE:
        return InstallState.PENDING, ""
    if rnd.random() < UNKNOWN_CHANCE:
        return InstallState.UNKNOWN, ""
    behind = title.approved_version and rnd.random() < OUTDATED_CHANCE
    version = PREVIOUS_VERSION.get(title.approved_version, title.approved_version) if behind else title.approved_version
    return InstallState.INSTALLED, version


def _fill_installed(rnd, record, title, today):
    """Dates, license and verifier for an installed record."""
    record.installed_on = today - timedelta(days=rnd.randint(*INSTALL_AGE_DAYS))
    record.assignment = ASSIGNMENT_FOR_MODEL[title.model]
    record.verified_on = today - timedelta(days=rnd.choice(VERIFIED_AGE_CHOICES))
    record.verified_by = rnd.choice(VERIFIERS)
    if rnd.random() < UNLICENSED_CHANCE:
        record.assignment = LicenseAssign.UNLICENSED
        record.notes = "Unknown key found during audit - remove or license."
    elif rnd.random() < UNVERIFIED_CHANCE:
        record.assignment = LicenseAssign.UNKNOWN


def _make_record(rnd, person: DemoPerson, title: SoftwareTitle, today):
    state, version = _state_and_version(rnd, title)
    record = InstallRecord(person.name, title.name, state, version)
    if state == InstallState.INSTALLED:
        _fill_installed(rnd, record, title, today)
    elif state == InstallState.PENDING:
        record.notes = "Ticket opened with IT; awaiting seat assignment."
    return record


def _scenarios(today):
    """Hand-made rows that demonstrate Review, License Issue and an applicability override."""
    verified = today - timedelta(days=20)
    return [
        InstallRecord("Hannah Fischer", "CATIA", InstallState.INSTALLED, "2025", None, LicenseAssign.FLOATING,
                      verified_on=verified, verified_by="IT Admin",
                      notes="Not applicable for this role - why is it here?"),
        InstallRecord("Marcus Webb", "MATLAB", InstallState.INSTALLED, "R2024b", None, LicenseAssign.FLOATING,
                      verified_on=verified, verified_by="IT Admin", notes="Server license expired last cycle."),
        InstallRecord("Riley Chen", "AutoCAD", InstallState.NOT_INSTALLED,
                      applicability_override=Applicability.REQUIRED,
                      notes="Override: Riley reviews drawings and needs AutoCAD."),
    ]


def make_records(today=None, seed=SEED):
    """All demo records, reproducible for a given date and seed."""
    today = today or date.today()
    rnd = random.Random(seed)
    records = []
    for person in PEOPLE:
        for title in TITLES:
            if _wanted(rnd, role_default(person.role, title.name)):
                records.append(_make_record(rnd, person, title, today))
    scenarios = _scenarios(today)
    taken = {(s.person, s.software) for s in scenarios}
    return [r for r in records if (r.person, r.software) not in taken] + scenarios
