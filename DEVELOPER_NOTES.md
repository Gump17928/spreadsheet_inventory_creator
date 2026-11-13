# Developer notes

Read `CODING_RULES.md` first. Run `run_tests.bat` before and after every change.

## Layout

| File (in `tool/`) | Role |
|---|---|
| `build_workbook.py` | Driver: `create_structure` (all sheets), `add_starter_content`, `add_demo_content`; saves the template and the example. |
| `layout.py` | Sheet names, capacities (`N_PEOPLE`, `N_TITLES`, ...), first data row, every sheet's column letters (`Inv`, `Cat`, `Ppl`, `Lst`, `Mat`), and the row-position helpers (`inventory_row`, `people_row_for`, `catalog_row_for`, `slot_for`). The one place that says "where things are". |
| `vocab.py` | The fixed words (`Status`, `InstallState`, `Applicability`, `Active`, ...) as `StrEnum`s, plus each status's colours, symbol and explanation. |
| `styles.py` | **The theme** (fonts, sizes, palette at the top), `Col`, header/body styling, `status_colours` (adds the symbol), `unlock`, `protect_sheet`, page setup. |
| `validation.py` | Data-validation rules (dropdown, date, count, text length, unique). |
| `cell_values.py` | `number_if_plain`: `"2026"` -> `2026` so Excel shows no "number stored as text" warning. |
| `refs.py` | Absolute references: ranges (`inv(Inv.STATUS)`), single cells (`at(sheet, col, row)`), `growing_list`. |
| `lookups.py` | Formula fragments: read a People / Catalog cell by row, a person by name, a role default; `not_installed_status`. |
| `inventory_formulas.py` | One function per Inventory column; each takes a row number and returns `=...`. |
| `starter_data.py` | Lists, sample catalog titles and role profiles that ship in the blank template. |
| `demo_data.py`, `demo_inventory.py` | Fake license details, people and installs for the example workbook (seeded, reproducible); `PERSON_INDEX` / `SLOT_INDEX`. |
| `dynamic_names.py` | Defined names for dropdowns that follow other sheets (`PeopleList`, `RoleList`). |
| `sheet_*.py` | One module per sheet: `create(wb)` builds structure; `fill_starter` / `fill_demo` add content. |
| `dashboard_tables.py` | The three Dashboard roll-up tables from one generic `Table` + `write_table`. |
| `excel_com.py`, `verify_in_excel.py` | Dev only: drive real Excel over COM (pywin32) to recalculate a workbook and read values back. |

Dependency order: `layout`, `vocab` -> `styles`, `validation`, `cell_values`, `refs` -> `lookups` -> `inventory_formulas`;
`starter_data` -> `demo_data`, `demo_inventory`; `dynamic_names`; `sheet_*` -> `dashboard_tables` -> `sheet_dashboard`
-> `build_workbook`.

## Tests

`run_tests.bat` (or `python -m unittest discover -s tests`) runs everything; about 15 seconds.

- `tests/test_data.py` - starter and demo data agree with each other and fit the workbook capacity.
- `tests/test_formulas.py` - the row mapping, formula text (balanced, inside Excel's 8,192-character / 64-level limits,
  pointing at the right People and Catalog rows) and the built workbooks' sheets, names, dropdowns and protection.
- `tests/test_cell_values.py` - the number/text conversion.
- `tests/test_excel_status.py` - builds a workbook with one typed row per status, lets **real Excel** calculate it and
  checks every status, blank-row defaults, a person and a title added later, a retired person, spare slots, and that no
  cell is an error. Skipped automatically when Excel or pywin32 is missing.

If you change a formula, add or change a case in `CASES` / `BLANK_CASES` in `test_excel_status.py` first.

## How the workbook works

| Idea | How it is built |
|---|---|
| Generated Inventory | One row for every (person, catalog slot) pair: row = `FIRST_ROW + person_index * N_TITLES + slot`. Person and Software are formulas pointing at `People!$A$<row>` and `Catalog!$A$<row>`; only install details are typed. |
| Nothing shifts | A row's person and title are fixed by position, so typed data never moves. This is why People and Catalog are protected against inserting, deleting and sorting rows. |
| Blank = not installed | The Status formula treats a blank (or "Not Installed") state by applicability: Required -> Missing, Optional -> Not Installed, Not Applicable -> N/A. |
| Applicability | The Role Defaults grid (columns follow Catalog order, by *position*) unless the Inventory row has an override. A blank grid cell is Optional. |
| Matrix and Person View | Each Matrix cell is `=Inventory!$S$<row>`; Person View indexes into that person's block with `MATCH(name)`. No searching of the whole Inventory. |
| Retiring | `People!Active = "No"` blanks the hidden Live column, so the Inventory status, Matrix row and Dashboard counts drop out; the data stays. |
| Spare capacity | Slots with no title or person show blank; `clear_when_blank` conditional formats remove fill and borders. |
| Theme | Everything visual is the block at the top of `styles.py` plus `_COLOURS` in `vocab.py`. |
| Status symbols | `status_colours` adds a number format (`"● "@`) to each status's conditional format, so the symbol is display-only and `COUNTIF` still sees the plain word. |
| Protection | No password. Typed columns are unlocked (`unlock`); formulas are locked. Reports, People, Catalog and Role Defaults are protected. |
| Recalculation | `fullCalcOnLoad` is set; openpyxl stores no cached results, so Excel computes everything on open. |

Facts learned the hard way (each is also commented next to the code that depends on it):

- `INDEX(...)` of an empty cell is `0` in arithmetic but `INDEX(...)&""=""` is true, which is how "blank" is tested.
- Numbers written as strings get a green "number stored as text" triangle; write numbers as numbers.
- Excel's `SpecialCells` raises COM error `-2146827284` when nothing matches; `excel_com.error_cells` treats that as "no errors".
- A version like `3.9` is the number 3.9, which is greater than `3.12`. Only year-style versions compare reliably as numbers.
- `StrEnum` members format as plain text in f-strings; a test checks no enum name leaks into a formula.
- Conditional-format number formats need unique ids of 164 or more (`SYMBOL_FORMAT_ID_BASE`).
- A protected sheet blocks sorting, and Excel only allows filtering on a protected sheet if an autofilter is already set (it is).

## Not yet done

- Charts: spare capacity rows would show as empty categories. Data bars on Coverage are used instead.
- Not tested in Google Sheets or LibreOffice (validation, `OFFSET` names, conditional-format number formats may differ).
- Previewers that do not calculate formulas (GitHub, some mail clients) show blank report cells until opened in Excel.
- The Person View picker lists retired people too; their view is blank.
- Dotted-version comparison (see above).
