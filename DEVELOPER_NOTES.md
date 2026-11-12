# Developer notes

Read `CODING_RULES.md` first. Run `run_tests.bat` before and after every change.

## Layout

| File (in `tool/`) | Role |
|---|---|
| `build_workbook.py` | Driver: `create_structure` (all sheets), `add_starter_content`, `add_demo_content`; saves the template and the example. |
| `layout.py` | Sheet names, capacities (`N_PEOPLE`, ...), first data row, and the column letters of every sheet (`Inv`, `Cat`, `Ppl`, `Lst`, `Mat`). The one place that says "where things are". |
| `vocab.py` | The fixed words (`Status`, `InstallState`, `Applicability`, ...) as `StrEnum`s, plus how each status is coloured and explained. |
| `styles.py` | Palette, `Col`, header/body styling, conditional-format helpers, page setup, sheet protection. |
| `validation.py` | Data-validation rules (dropdown, date, count, text length, unique). |
| `cell_values.py` | `number_if_plain`: `"2026"` -> `2026` so Excel shows no "number stored as text" warning. |
| `refs.py` | Absolute range references (`inv(Inv.STATUS)` -> `Inventory!$S$6:$S$1505`) and `growing_list`. |
| `lookups.py` | Formula fragments: look up a Catalog / People / Role Defaults value; `not_installed_status`. |
| `inventory_formulas.py` | One function per calculated Inventory column. Each takes a row number and returns `=...`. |
| `starter_data.py` | Lists, sample catalog titles and role profiles that ship in the blank template. |
| `demo_data.py`, `demo_inventory.py` | Fake license details, people and installs for the example workbook (seeded, reproducible). |
| `dynamic_names.py` | Defined names for dropdowns that follow other sheets (`PeopleList`, `RoleList`, `SoftwareList`). |
| `sheet_*.py` | One module per sheet: `create(wb)` builds structure; `fill_starter` / `fill_demo` add content. |
| `dashboard_tables.py` | The three Dashboard roll-up tables from one generic `Table` + `write_table`. |
| `excel_com.py`, `verify_in_excel.py` | Dev only: drive real Excel over COM (pywin32) to recalculate a workbook and read values back. |

Dependency order: `layout`, `vocab` -> `styles`, `validation`, `cell_values`, `refs` -> `lookups` -> `inventory_formulas`;
`starter_data` -> `demo_data`, `demo_inventory`; `dynamic_names`; `sheet_*` -> `dashboard_tables` -> `sheet_dashboard`
-> `build_workbook`.

## Tests

`run_tests.bat` (or `python -m unittest discover -s tests`) runs everything; about 10 seconds.

- `tests/test_data.py` - starter and demo data agree with each other and fit the workbook capacity.
- `tests/test_formulas.py` - formula text is balanced and inside Excel's limits (8,192 characters, 64 nesting levels);
  the built workbooks have the right sheets, names, dropdowns and protection.
- `tests/test_cell_values.py` - the number/text conversion.
- `tests/test_excel_status.py` - builds a small workbook with one row per status, lets **real Excel** calculate it and
  checks every status, the Matrix fallbacks (Missing / N/A / Not Installed with no row) and that no cell is an error.
  Skipped automatically when Excel or pywin32 is missing.

If you change a formula, add or change a case in `CASES` in `test_excel_status.py` first.

## How the workbook works

| Idea | How it is built |
|---|---|
| Single source of truth | `Inventory` holds one row per person + software. Matrix, Person View and Dashboard are formulas over it. |
| "No row = not installed" | The Matrix looks up `Person|Software` in `Inventory[Key]`; if absent it falls back to the role default (`not_installed_status`). |
| Applicability | Role Defaults grid (columns follow Catalog order, by *position*) unless the Inventory row has an override. Blank grid cell = Optional. |
| Fixed capacity | Formulas, validation and formatting are pre-filled to `N_*` rows. To grow, change the constants in `layout.py` and rebuild. |
| Invisible spare rows | `clear_when_blank` conditional formats remove fill and borders where the key cell is empty. |
| Dropdowns that grow | `growing_list` builds `OFFSET(first,0,0,COUNTA(...),1)` defined names. |
| Report sheets | Protected without a password (`protect_report`); only the Person View picker is unlocked. |
| Recalculation | `fullCalcOnLoad` is set; openpyxl stores no cached results, so Excel computes everything on open. |

Facts learned the hard way (each is also commented next to the code that depends on it):

- `INDEX(...)` of an empty cell is `0` in arithmetic but `INDEX(...)&""=""` is true, which is how "blank" is tested.
- Numbers typed into a text-formatted cell, or written as strings by code, get a green "number stored as text" triangle; write numbers as numbers.
- Excel's `SpecialCells` raises COM error `-2146827284` when nothing matches; `excel_com.error_cells` treats that as "no errors".
- A version like `3.9` is the number 3.9, which is greater than `3.12`. Only year-style versions compare reliably as numbers.
- `Status` etc. are `StrEnum`, so they format as their plain text inside f-strings; a test checks no enum name leaks into a formula.

## Not yet done

- No charts: the spare capacity rows would show as empty categories. Data bars on Coverage are used instead.
- Not tested in Google Sheets or LibreOffice (data validation, defined `OFFSET` names and conditional formats may differ).
- Previewers that do not calculate formulas (GitHub, some mail clients) show blank report cells until opened in Excel.
- Dotted-version comparison (see above).
