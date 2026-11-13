# Software Inventory

An Excel workbook that shows, for a group of people, exactly what software is installed on each computer, what is
missing, what is out of date or unlicensed, and what simply doesn't apply to them.

**Just want the spreadsheet?** Open `Software_Inventory.xlsx` (blank template) or `Software_Inventory_Example.xlsx`
(the same workbook filled with fake data). Start on the *About* sheet. Desktop Excel 2013 or later, no macros.

## How it works

Fill in **Catalog** (software), **People** and **Role Defaults** (what each role must, may or need not have). The
**Inventory** sheet then lays itself out: every person gets a row for every title, already marked Required, Optional or
Not Applicable. You only type what is installed; a blank row means "not installed" and shows as Missing, Not Installed
or N/A depending on the role.

| Sheet | Purpose |
|---|---|
| About / Key | How to use it; every status and colour explained. |
| Inventory | **Data entry.** Install state, version, license assignment, expiry and verification date. |
| Matrix | Everyone against every title, colour-coded and grouped under suites (Autodesk, Dassault, ANSYS, ...). |
| Person View | Pick a name to see what they have, are missing, or don't need. |
| Dashboard | Coverage and "needs attention" by group, department and software. |
| Catalog | Titles: group, approved version, license model, **license server, seats, expiry**, seat usage. |
| People / Role Defaults / Lists | Roster (with an Active flag), role profiles, dropdown lists and warning thresholds. |

Adding someone or something: type in the next empty row of People or Catalog. Never insert, delete or sort rows
there (they are protected, without a password, to stop it). Retire a leaver with Active = No.

Capacity: 100 people, 30 titles, 30 roles, 3,000 Inventory rows.

## Rebuilding (optional)

The `.xlsx` files are standalone. To regenerate them after changing the code:

1. Double-click `setup.bat` once (creates `.venv` and installs the requirements).
2. Double-click `build.bat`.
3. `run_tests.bat` runs the tests and the style check.

Developers: read `CODING_RULES.md` and `DEVELOPER_NOTES.md`. The look (fonts, colours) is the theme block at the top of
`tool/styles.py`.
