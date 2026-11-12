"""Starter content that ships in the blank template: lists, a sample catalog and role profiles.

Pure data plus one lookup; nothing here touches a workbook.
"""
from dataclasses import dataclass

from vocab import Applicability, LicenseModel

GROUPS = ["Autodesk", "Dassault Systèmes", "ANSYS", "Microsoft", "Adobe", "Bentley", "MathWorks",
          "PTC", "Open Source", "Other"]
CATEGORIES = ["CAD / Design", "Simulation / Analysis", "Engineering Calculation", "Documentation / PDF",
              "Productivity", "Programming / Scripting", "Project Management", "Utilities"]
DEPARTMENTS = ["Mechanical", "Electrical", "Controls", "Structural", "Projects", "Business Admin"]


@dataclass(frozen=True)
class SoftwareTitle:
    name: str
    group: str
    category: str
    vendor: str
    approved_version: str
    model: LicenseModel


_NAMED, _FLOAT, _SITE = LicenseModel.NAMED, LicenseModel.FLOATING, LicenseModel.SITE
CAD, SIM, CALC, DOCS, OFFICE, CODE, PM = CATEGORIES[:7]

TITLES = [
    SoftwareTitle("AutoCAD", "Autodesk", CAD, "Autodesk", "2026", _NAMED),
    SoftwareTitle("Revit", "Autodesk", CAD, "Autodesk", "2026", _NAMED),
    SoftwareTitle("Inventor", "Autodesk", CAD, "Autodesk", "2026", _NAMED),
    SoftwareTitle("Navisworks Manage", "Autodesk", PM, "Autodesk", "2026", _NAMED),
    SoftwareTitle("SOLIDWORKS", "Dassault Systèmes", CAD, "Dassault Systèmes", "2026", _FLOAT),
    SoftwareTitle("CATIA", "Dassault Systèmes", CAD, "Dassault Systèmes", "2025", _FLOAT),
    SoftwareTitle("ANSYS Mechanical", "ANSYS", SIM, "ANSYS", "2025 R2", _FLOAT),
    SoftwareTitle("ANSYS Fluent", "ANSYS", SIM, "ANSYS", "2025 R2", _FLOAT),
    SoftwareTitle("PTC Mathcad Prime", "PTC", CALC, "PTC", "11", LicenseModel.STANDALONE),
    SoftwareTitle("MATLAB", "MathWorks", CODE, "MathWorks", "R2025b", _FLOAT),
    SoftwareTitle("Microsoft 365 Apps", "Microsoft", OFFICE, "Microsoft", "", _SITE),
    SoftwareTitle("Microsoft Visio", "Microsoft", DOCS, "Microsoft", "", _NAMED),
    SoftwareTitle("Adobe Acrobat Pro", "Adobe", DOCS, "Adobe", "2025", _NAMED),
    SoftwareTitle("Bluebeam Revu", "Other", DOCS, "Bluebeam", "21", _NAMED),
    SoftwareTitle("Python", "Open Source", CODE, "PSF", "3.12", LicenseModel.FREE),
]
TITLE_NAMES = [title.name for title in TITLES]


@dataclass(frozen=True)
class RoleProfile:
    """Titles a role must have or may have; every other title is Not Applicable."""
    required: tuple[str, ...]
    optional: tuple[str, ...] = ()


ROLE_PROFILES = {
    "Mechanical Engineer": RoleProfile(
        ("Inventor", "SOLIDWORKS", "ANSYS Mechanical", "PTC Mathcad Prime", "Microsoft 365 Apps", "Adobe Acrobat Pro"),
        ("AutoCAD", "CATIA", "ANSYS Fluent", "MATLAB", "Microsoft Visio", "Bluebeam Revu", "Python")),
    "Electrical Engineer": RoleProfile(
        ("AutoCAD", "Revit", "PTC Mathcad Prime", "Microsoft 365 Apps", "Microsoft Visio", "Adobe Acrobat Pro",
         "Bluebeam Revu"),
        ("Navisworks Manage", "MATLAB", "Python")),
    "Controls Engineer": RoleProfile(
        ("AutoCAD", "MATLAB", "Microsoft 365 Apps", "Microsoft Visio", "Adobe Acrobat Pro", "Python"),
        ("PTC Mathcad Prime", "Bluebeam Revu")),
    "Structural Engineer": RoleProfile(
        ("AutoCAD", "Revit", "PTC Mathcad Prime", "Microsoft 365 Apps", "Adobe Acrobat Pro", "Bluebeam Revu"),
        ("Navisworks Manage", "ANSYS Mechanical", "Python", "Microsoft Visio")),
    "Project Manager": RoleProfile(
        ("Microsoft 365 Apps", "Microsoft Visio", "Adobe Acrobat Pro", "Bluebeam Revu"),
        ("Navisworks Manage",)),
    "Business Admin": RoleProfile(
        ("Microsoft 365 Apps", "Adobe Acrobat Pro"),
        ("Microsoft Visio",)),
}
ROLES = list(ROLE_PROFILES)


def role_default(role, title_name):
    """Applicability of a title for a role according to the starter profiles."""
    profile = ROLE_PROFILES[role]
    if title_name in profile.required:
        return Applicability.REQUIRED
    if title_name in profile.optional:
        return Applicability.OPTIONAL
    return Applicability.NOT_APPLICABLE
