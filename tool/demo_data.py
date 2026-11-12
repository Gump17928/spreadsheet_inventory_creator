"""Fake data used only by the example workbook: license details, people and a few hand-made scenarios."""
from dataclasses import dataclass


@dataclass(frozen=True)
class LicenseDetail:
    server: str
    seats: int | None
    expires_in_days: int | None    # relative to the build date; negative = already expired
    admin: str


LICENSE_DETAILS = {
    "AutoCAD": LicenseDetail("", 25, 160, "IT Admin"),
    "Revit": LicenseDetail("", 12, 160, "IT Admin"),
    "Inventor": LicenseDetail("", 10, 160, "IT Admin"),
    "Navisworks Manage": LicenseDetail("", 6, 160, "IT Admin"),
    "SOLIDWORKS": LicenseDetail("lic-sw01:25734", 8, 45, "J. Alvarez"),
    "CATIA": LicenseDetail("lic-dsls01:4085", 3, 20, "J. Alvarez"),
    "ANSYS Mechanical": LicenseDetail("lic-ansys01:1055", 4, 90, "K. Brandt"),
    "ANSYS Fluent": LicenseDetail("lic-ansys01:1055", 2, 90, "K. Brandt"),
    "PTC Mathcad Prime": LicenseDetail("", 20, None, "IT Admin"),
    "MATLAB": LicenseDetail("lic-mw01:27000", 5, -12, "K. Brandt"),
    "Microsoft 365 Apps": LicenseDetail("", 60, 300, "IT Admin"),
    "Microsoft Visio": LicenseDetail("", 15, 300, "IT Admin"),
    "Adobe Acrobat Pro": LicenseDetail("", 40, 200, "IT Admin"),
    "Bluebeam Revu": LicenseDetail("", 15, 120, "IT Admin"),
    "Python": LicenseDetail("", None, None, "IT Admin"),
}


@dataclass(frozen=True)
class DemoPerson:
    name: str
    department: str
    role: str


PEOPLE = [
    DemoPerson("Alex Rivera", "Mechanical", "Mechanical Engineer"),
    DemoPerson("Priya Shah", "Mechanical", "Mechanical Engineer"),
    DemoPerson("Marcus Webb", "Mechanical", "Mechanical Engineer"),
    DemoPerson("Dana Kowalski", "Electrical", "Electrical Engineer"),
    DemoPerson("Tomás Ortega", "Electrical", "Electrical Engineer"),
    DemoPerson("Mei Lin", "Electrical", "Electrical Engineer"),
    DemoPerson("Sam Okafor", "Controls", "Controls Engineer"),
    DemoPerson("Jordan Blake", "Controls", "Controls Engineer"),
    DemoPerson("Hannah Fischer", "Structural", "Structural Engineer"),
    DemoPerson("Luca Moretti", "Structural", "Structural Engineer"),
    DemoPerson("Riley Chen", "Projects", "Project Manager"),
    DemoPerson("Noor Haddad", "Projects", "Project Manager"),
    DemoPerson("Casey Morgan", "Business Admin", "Business Admin"),
    DemoPerson("Taylor Brooks", "Business Admin", "Business Admin"),
]
LOCATION = "HQ - Floor 2"
MANAGER_ENGINEERING, MANAGER_BUSINESS = "Pat Delgado", "Lee Tran"
BUSINESS_DEPARTMENT = "Business Admin"
