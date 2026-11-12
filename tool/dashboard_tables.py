"""Dashboard tables: the by-group, by-department and by-software roll-ups, built from one generic writer."""
from collections.abc import Callable
from dataclasses import dataclass

from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Font

import refs
from layout import (
    FIRST_ROW,
    N_SUMMARY_ROWS,
    N_TITLES,
    Cat,
    Lst,
    Ppl,
    Sheet,
    col_letter,
)
from styles import NAVY, ORANGE_TEXT, clear_when_blank, fill, font, header_cell, style_cells
from vocab import ATTENTION, DASHBOARD_STATUSES, HAVE, STATUS_STYLE, Status

FIRST_STATUS_COL = 3
COVERAGE_COL = FIRST_STATUS_COL + len(DASHBOARD_STATUSES)
ATTENTION_COL = COVERAGE_COL + 1
NAME_WIDTH, SECOND_WIDTH, STATUS_WIDTH, SUMMARY_WIDTH = 28, 20, 11, 12
HEADER_HEIGHT = 30
LOW_COVERAGE = 0.75       # coverage below this is shown in red
BAR_COLOUR = "8FB4E3"
PERCENT_FORMAT = "0%"

STATUS_LETTER = {status: col_letter(FIRST_STATUS_COL + i) for i, status in enumerate(DASHBOARD_STATUSES)}
COVERAGE_LETTER, ATTENTION_LETTER = col_letter(COVERAGE_COL), col_letter(ATTENTION_COL)


@dataclass(frozen=True)
class Table:
    """One roll-up table. Each *_formula takes (row, index-in-table) except status_formula(row, status_letter)."""
    title: str
    name_header: str
    second_header: str
    header_row: int
    rows: int
    name_formula: Callable[[int, int], str]
    second_formula: Callable[[int, int], str]
    status_formula: Callable[[int, str], str]

    @property
    def first(self):
        return self.header_row + 1

    @property
    def last(self):
        return self.header_row + self.rows


def _sum(statuses, row):
    return "+".join(f"{STATUS_LETTER[status]}{row}" for status in statuses)


def _coverage(row):
    have = _sum(HAVE, row)
    missing, pending = (f"{STATUS_LETTER[status]}{row}" for status in (Status.MISSING, Status.PENDING))
    return f'=IF($A{row}="","",IFERROR(({have})/(({have})+{missing}+{pending}),"—"))'


def _attention(row):
    return f'=IF($A{row}="","",{_sum(ATTENTION, row)})'


def _write_head(ws, table):
    ws[f"A{table.header_row - 1}"] = table.title
    ws[f"A{table.header_row - 1}"].font = font(12, True, NAVY)
    header_cell(ws, table.header_row, 1, table.name_header, width=NAME_WIDTH)
    header_cell(ws, table.header_row, 2, table.second_header, width=SECOND_WIDTH)
    for index, status in enumerate(DASHBOARD_STATUSES):
        cell = header_cell(ws, table.header_row, FIRST_STATUS_COL + index, status.value, width=STATUS_WIDTH)
        style = STATUS_STYLE[status]
        cell.fill, cell.font = fill(style.fill), Font(name="Calibri", bold=True, size=9, color=style.font)
    header_cell(ws, table.header_row, COVERAGE_COL, "Coverage", calc=True, width=STATUS_WIDTH)
    header_cell(ws, table.header_row, ATTENTION_COL, "Needs Attention", calc=True, width=SUMMARY_WIDTH)
    ws.row_dimensions[table.header_row].height = HEADER_HEIGHT


def _write_rows(ws, table):
    for index in range(table.rows):
        row = table.first + index
        ws[f"A{row}"] = table.name_formula(row, index)
        ws[f"B{row}"] = table.second_formula(row, index)
        for letter in STATUS_LETTER.values():
            ws[f"{letter}{row}"] = table.status_formula(row, letter)
        ws[f"{COVERAGE_LETTER}{row}"] = _coverage(row)
        ws[f"{ATTENTION_LETTER}{row}"] = _attention(row)


def _format_rows(ws, table):
    style_cells(ws, f"A{table.first}:B{table.last}")
    style_cells(ws, f"C{table.first}:{ATTENTION_LETTER}{table.last}", align="center")
    style_cells(ws, f"{COVERAGE_LETTER}{table.first}:{COVERAGE_LETTER}{table.last}", fmt=PERCENT_FORMAT, align="center")
    coverage = f"{COVERAGE_LETTER}{table.first}:{COVERAGE_LETTER}{table.last}"
    ws.conditional_formatting.add(coverage, FormulaRule(
        formula=[f"AND(ISNUMBER({COVERAGE_LETTER}{table.first}),{COVERAGE_LETTER}{table.first}<{LOW_COVERAGE})"],
        font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(coverage, DataBarRule(
        start_type="num", start_value=0, end_type="num", end_value=1, color=BAR_COLOUR, showValue=True))
    ws.conditional_formatting.add(f"{ATTENTION_LETTER}{table.first}:{ATTENTION_LETTER}{table.last}", CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color=ORANGE_TEXT, bold=True)))
    clear_when_blank(ws, f"A{table.first}:{ATTENTION_LETTER}{table.last}", f"$A{table.first}")


def write_table(ws, table):
    _write_head(ws, table)
    _write_rows(ws, table)
    _format_rows(ws, table)


def _list_name(list_col):
    """Name formula that mirrors an editable list on the Lists sheet."""
    def formula(row, index):
        cell = f"{Sheet.LISTS}!${list_col}${FIRST_ROW + index}"
        return f'=IF({cell}="","",{cell})'
    return formula


def group_table(header_row, software):
    """By group: sums the by-software table."""
    def second(row, index):
        return f'=IF($A{row}="","",COUNTIF($B${software.first}:$B${software.last},$A{row}))'

    def status(row, letter):
        return (f'=IF($A{row}="","",SUMIF($B${software.first}:$B${software.last},$A{row},'
                f'{letter}${software.first}:{letter}${software.last}))')
    return Table("By Group / Suite", "Group / Suite", "Titles", header_row, N_SUMMARY_ROWS,
                 _list_name(Lst.GROUP), second, status)


def department_table(header_row):
    """By department: counts statuses straight from the Matrix."""
    def second(row, index):
        return f'=IF($A{row}="","",COUNTIF({refs.ppl(Ppl.DEPARTMENT)},$A{row}))'

    def status(row, letter):
        departments = refs.matrix_col("B")
        return (f'=IF($A{row}="","",SUMPRODUCT(({departments}=$A{row})*({refs.matrix_body()}={letter}${header_row})))')
    return Table("By Department", "Department", "People", header_row, N_SUMMARY_ROWS,
                 _list_name(Lst.DEPARTMENT), second, status)


def software_table(header_row):
    """By software: counts statuses down one Matrix column per catalog slot."""
    def name(row, index):
        cell = f"{Sheet.CATALOG}!${Cat.NAME}${FIRST_ROW + index}"
        return f'=IF({cell}="","",{cell})'

    def second(row, index):
        return f'=IF($A{row}="","",{Sheet.CATALOG}!${Cat.GROUP}${FIRST_ROW + index}&"")'

    def status(row, letter):
        index = row - header_row - 1
        return f'=IF($A{row}="","",COUNTIF(INDEX({refs.matrix_body()},0,{index + 1}),{letter}${header_row}))'
    return Table("By Software", "Software", "Group / Suite", header_row, N_TITLES, name, second, status)
