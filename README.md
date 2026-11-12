# Software Inventory

An Excel workbook that shows, for a group of people, exactly what software is installed on each computer, what is
missing, what is out of date or unlicensed, and what simply doesn't apply to them.

**Just want the spreadsheet?** Open `Software_Inventory.xlsx` (blank template) or `Software_Inventory_Example.xlsx`
(the same workbook filled with fake data). Start on the *Start Here* sheet. Desktop Excel 2013 or later, no macros.

## What you get

| Sheet | Purpose |
|---|---|
| Start Here / Legend | How to use it; every status and colour explained. |
| Inventory | **Data entry.** One row per person + software, with version, license assignment, expiry and verification date. |
| Matrix | Everyone x every title, colour-coded, grouped under suites (Autodesk, Dassault, ANSYS, ...). |
| Person View | Pick a name to see what they have, are missing, or don't need. |
| Dashboard | Coverage and "needs attention" by group, department and software. |
| Catalog | Software titles: group, approved version, license model, **license server, seats, expiry**, seat usage. |
| People / Role Defaults / Lists | Roster, what each role must/may have, dropdown lists and warning thresholds. |

Capacity: 150 people, 40 titles, 30 roles, 1,500 inventory rows.

## Rebuilding (optional)

The `.xlsx` files are standalone. To regenerate them after changing the code:

1. Double-click `setup.bat` once (creates `.venv` and installs the requirements).
2. Double-click `build.bat`.
3. `run_tests.bat` runs the tests and the style check.

Developers: read `CODING_RULES.md` and `DEVELOPER_NOTES.md`.
