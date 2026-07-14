"""
Business logic for turning a raw Ventus composition export into the
formatted "Reranged" sales report.

This module knows nothing about Tkinter or file dialogs — it just takes
an input path (+ optional top-seller codes) and produces an output file.
That separation is what makes it testable and reusable (e.g. from a
future CLI or batch job) without dragging a GUI along.
"""

from __future__ import annotations

import os
from collections.abc import Callable, Iterable
from ctypes import alignment
from dataclasses import dataclass, field
from datetime import datetime
from tkinter import CENTER

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

import constants as C

ProgressCallback = Callable[[str], None]


class ProcessingError(Exception):
    """Raised for anything that stops processing with a user-facing message."""


@dataclass
class ProcessingResult:
    output_path: str
    row_count: int
    bestseller_matches: int
    unmatched_codes: list[str] = field(default_factory=list)


def _noop(_msg: str) -> None:
    pass


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def process_file(
    input_path: str,
    top_codes: Iterable[str] | None = None,
    progress: ProgressCallback = _noop,
) -> ProcessingResult:
    """
    Run the full pipeline on `input_path` and save the result next to it
    as "<name>_processed.xlsx". Returns a ProcessingResult summary.

    Raises ProcessingError with a readable message on any expected failure
    (missing file, missing required columns, etc.).
    """
    if not input_path or not os.path.isfile(input_path):
        raise ProcessingError(f'Soubor nenalezen: {input_path}')

    top_codes = _normalize_codes(top_codes or [])

    progress('Otevírám soubor…')
    try:
        wb = load_workbook(input_path)
    except Exception as exc:  # noqa: BLE001 - surface as a clean error
        raise ProcessingError(f'Nepodařilo se otevřít soubor: {exc}') from exc

    ws = wb.active
    if not isinstance(ws, Worksheet):
        raise ProcessingError('Nepodařilo se zjistit aktivní list v souboru.')

    progress('Čistím zdrojová data…')
    _clean_source_sheet(ws)

    column_map = _build_column_map(ws)
    _require_columns(column_map, [C.SORT_COLUMN, C.CODE_COLUMN])

    progress('Řadím podle prodejní ceny 2025…')
    data_rows = list(ws.iter_rows(min_row=2, max_row=ws.max_row))
    sorted_rows = _sort_rows(data_rows, column_map)

    progress('Sestavuji nový list…')
    new_ws = wb.create_sheet(title=C.SHEET_NAME_OUTPUT)
    _write_header(new_ws)
    _write_data_rows(new_ws, ws.title, column_map, sorted_rows)

    progress('Přidávám sloupce dodávek…')
    _append_delivery_columns(new_ws, column_map, sorted_rows)

    progress('Počítám celkové CBM…')
    _add_total_cbm(new_ws, len(sorted_rows))

    progress('Formátuji vzhled…')
    _apply_styling(new_ws)

    bestseller_matches, unmatched = 0, list(top_codes)
    if top_codes:
        progress('Zvýrazňuji top prodejní položky…')
        bestseller_matches, unmatched = _mark_bestsellers(new_ws, top_codes)

    output_path = _build_output_path(input_path)
    progress('Ukládám výsledek…')
    try:
        wb.save(output_path)
    except Exception as exc:  # noqa: BLE001 - surface as a clean error
        raise ProcessingError(
            'Nepodařilo se uložit výsledek. Zdrojový soubor pravděpodobně '
            f'obsahuje formátování, které se nepodařilo zpracovat ({exc}).'
        ) from exc

    return ProcessingResult(
        output_path=output_path,
        row_count=len(sorted_rows),
        bestseller_matches=bestseller_matches,
        unmatched_codes=unmatched,
    )


# ---------------------------------------------------------------------------
# Cleaning
# ---------------------------------------------------------------------------


def _clean_source_sheet(ws: Worksheet) -> None:
    _strip_volatile_structures(ws)

    if 'A3:C3' in ws.merged_cells:
        ws.unmerge_cells('A3:C3')
    if 'A1:C1' in ws.merged_cells:
        ws.unmerge_cells('A1:C1')

    ws.delete_rows(2, 2)
    ws.delete_rows(ws.max_row)
    ws['B1'] = 'Nazev'
    ws['C1'] = 'Skupina'


def _strip_volatile_structures(ws: Worksheet) -> None:
    """Remove Excel Tables, autofilters and data validations from the
    source sheet before any rows are deleted.

    openpyxl does not update these when rows/columns are removed, so a
    Table (or autofilter) that pointed at e.g. "A1:D50" silently keeps
    pointing there even after we shrink the sheet underneath it. That
    mismatch corrupts the file and surfaces as a cryptic error - often
    something like "'<column name>' is not a valid column name" - when
    the workbook is saved. We only ever read raw values from this sheet,
    so none of this formatting needs to survive anyway.
    """
    for name in list(ws.tables.keys()):
        del ws.tables[name]
    ws.auto_filter.ref = None
    ws.data_validations.dataValidation = []


def _build_column_map(ws: Worksheet, header_row: int = 1) -> dict[str, list[int]]:
    """Map cleaned header name -> list of 1-based column indices (a name
    can appear more than once, e.g. multiple delivery batches)."""
    column_map: dict[str, list[int]] = {}
    for cell in ws[header_row]:
        if cell.value is None:
            continue
        name = ' '.join(str(cell.value).split()).strip()
        column_map.setdefault(name, []).append(cell.column)
    return column_map


def _require_columns(column_map: dict[str, list[int]], required: Iterable[str]) -> None:
    missing = [name for name in required if name not in column_map]
    if missing:
        raise ProcessingError('Ve zdrojovém souboru chybí očekávané sloupce: ' + ', '.join(missing))


# ---------------------------------------------------------------------------
# Sorting
# ---------------------------------------------------------------------------


def _sort_rows(data_rows: list[tuple], column_map: dict[str, list[int]]) -> list[tuple]:
    sort_col_idx = column_map.get(C.SORT_COLUMN, [None])[0]
    if sort_col_idx is None:
        return data_rows

    def sort_key(row: tuple) -> float:
        value = row[sort_col_idx - 1].value
        try:
            return float(value) if value is not None else 0.0
        except (TypeError, ValueError):
            return 0.0

    return sorted(data_rows, key=sort_key, reverse=True)


# ---------------------------------------------------------------------------
# Building the "Reranged" sheet
# ---------------------------------------------------------------------------


def _col(name: str) -> int:
    """1-based column index of `name` in the output layout."""
    return C.DESIRED_ORDER.index(name) + 1


def _col_letter(name: str) -> str:
    return get_column_letter(_col(name))


def _write_header(new_ws: Worksheet) -> None:
    for idx, name in enumerate(C.DESIRED_ORDER, start=1):
        new_ws.cell(row=C.HEADER_ROW, column=idx, value=name)


def _write_data_rows(
    new_ws: Worksheet,
    source_sheet_name: str,
    column_map: dict[str, list[int]],
    sorted_rows: list[tuple],
) -> None:
    delivery_indices = column_map.get(C.DELIVERY_QTY_COLUMN, [])

    for row_offset, row_cells in enumerate(sorted_rows, start=C.FIRST_DATA_ROW):
        source_row_num = row_cells[0].row

        for col_idx, col_name in enumerate(C.DESIRED_ORDER, start=1):
            if col_name == 'Order ctn':
                new_ws.cell(row=row_offset, column=col_idx, value=0)
                continue

            if col_name == 'Suma zbývá dodat':
                if delivery_indices:
                    parts = [
                        f"'{source_sheet_name}'!{get_column_letter(i)}{source_row_num}"
                        for i in delivery_indices
                    ]
                    new_ws.cell(row=row_offset, column=col_idx, value='=' + '+'.join(parts))
                else:
                    new_ws.cell(row=row_offset, column=col_idx, value=0)
                continue

            source_indices = column_map.get(col_name, [])
            value = None
            if source_indices:
                value = row_cells[source_indices[0] - 1].value
                if col_name in C.PERCENT_COLUMNS and isinstance(value, (int, float)):
                    value = value / 100
            new_ws.cell(row=row_offset, column=col_idx, value=value)

        _write_row_formulas(new_ws, row_offset)


def _write_row_formulas(new_ws: Worksheet, row: int) -> None:
    stock_free = _col_letter('Skladem volné')
    remaining = _col_letter('Suma zbývá dodat')
    qty_2025 = _col_letter('Množství 2025-Y25')
    order_ctn = _col_letter('Order ctn')
    package_3 = _col_letter('Balení 3')
    volume = _col_letter('Objem')

    new_ws.cell(
        row=row,
        column=_col('(Skladem + zbývá dodat) / prodej 2025'),
        value=f'=IFERROR(({stock_free}{row}+{remaining}{row})/{qty_2025}{row}, 0)',
    )
    new_ws.cell(
        row=row,
        column=_col('Order pcs'),
        value=f'={order_ctn}{row}*{package_3}{row}',
    )
    new_ws.cell(
        row=row,
        column=_col('Objem celkem'),
        value=f'={volume}{row}*{order_ctn}{row}',
    )


def _append_delivery_columns(
    new_ws: Worksheet,
    column_map: dict[str, list[int]],
    sorted_rows: list[tuple],
) -> None:
    """Each extra 'Zbývá dodat ks' / 'Datum dodání z.' pair (multiple
    deliveries) gets appended as its own qty+date column pair after the
    fixed layout."""
    qty_indices = column_map.get(C.DELIVERY_QTY_COLUMN, [])
    date_indices = column_map.get(C.DELIVERY_DATE_COLUMN, [])
    if not qty_indices:
        return

    start_col = len(C.DESIRED_ORDER) + 1

    for batch, raw_qty_col in enumerate(qty_indices):
        qty_col = start_col + batch * 2
        date_col = qty_col + 1
        raw_date_col = date_indices[batch] if batch < len(date_indices) else raw_qty_col + 1

        new_ws.cell(row=C.HEADER_ROW, column=qty_col, value=f'Zbyva dodat {batch + 1}')
        new_ws.cell(row=C.HEADER_ROW, column=date_col, value=f'Datum dodáni {batch + 1}')

        for row_offset, row_cells in enumerate(sorted_rows, start=C.FIRST_DATA_ROW):
            qty_val = row_cells[raw_qty_col - 1].value
            if not qty_val:
                continue
            new_ws.cell(row=row_offset, column=qty_col, value=qty_val)
            date_val = row_cells[raw_date_col - 1].value
            new_ws.cell(row=row_offset, column=date_col, value=date_val)


def _add_total_cbm(new_ws: Worksheet, row_count: int) -> None:
    volume_letter = _col_letter('Objem celkem')
    last_row = row_count + C.FIRST_DATA_ROW - 1
    new_ws.cell(row=1, column=28, value='Total CBM')
    new_ws.cell(row=1, column=29, value=f'=SUM({volume_letter}3:{volume_letter}{last_row})')


# ---------------------------------------------------------------------------
# Bestseller highlighting (new feature)
# ---------------------------------------------------------------------------


def _normalize_codes(codes: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result = []
    for raw in codes:
        code = str(raw).strip()
        if not code or code.upper() in seen:
            continue
        seen.add(code.upper())
        result.append(code)
    return result


def _mark_bestsellers(new_ws: Worksheet, top_codes: list[str]) -> tuple[int, list[str]]:
    """Highlight every row whose product code (Sortiment) is in `top_codes`.
    Returns (number of rows matched, codes that were never found)."""
    code_col = _col(C.CODE_COLUMN)
    fill = PatternFill(
        start_color=C.COLOR_HEX['bestseller'],
        end_color=C.COLOR_HEX['bestseller'],
        fill_type='solid',
    )
    wanted = {c.upper() for c in top_codes}
    found: set[str] = set()

    for row in new_ws.iter_rows(min_row=C.FIRST_DATA_ROW, max_row=new_ws.max_row):
        cell = row[code_col - 1]
        value = str(cell.value).strip().upper() if cell.value is not None else ''
        if value in wanted:
            found.add(value)
            for c in row[:4]:
                c.fill = fill

    unmatched = [c for c in top_codes if c.upper() not in found]
    return len(found), unmatched


# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------


def _fill(name: str) -> PatternFill:
    hex_color = C.COLOR_HEX[name]
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')


def _apply_styling(new_ws: Worksheet) -> None:
    base_font = Font(name=C.FONT_NAME, size=C.BASE_FONT_SIZE)
    header_font = Font(name=C.FONT_NAME, size=C.HEADER_FONT_SIZE, bold=True)

    new_ws.row_dimensions[1].height = 25
    new_ws.row_dimensions[2].height = 40
    for letter in C.HIDDEN_COLUMNS:
        new_ws.column_dimensions[letter].hidden = True
    for row_idx in range(3, new_ws.max_row + 1):
        new_ws.row_dimensions[row_idx].height = 18

    _autosize_columns(new_ws)
    _color_by_column_band(new_ws, base_font)
    _apply_number_formats_and_borders(new_ws, header_font)
    _style_threshold_row(new_ws)
    _style_total_cbm_cells(new_ws)


def _autosize_columns(new_ws: Worksheet) -> None:
    for col in new_ws.columns:
        # col[0].column can be None in some cases (type checker warns). Skip such columns.
        if col[0].column is None:
            continue
        letter = get_column_letter(col[0].column)
        max_len = max((len(str(c.value)) for c in col[2:] if c.value is not None), default=0)
        width = max_len + 3
        new_ws.column_dimensions[letter].width = min(max(width, 11), 35)
    new_ws.column_dimensions['O'].width = 15
    new_ws.column_dimensions['P'].width = 15
    new_ws.column_dimensions['Q'].width = 15
    new_ws.column_dimensions['AL'].width = 15
    new_ws.column_dimensions['Z'].width = 15
    new_ws.column_dimensions['AA'].width = 15


def _color_by_column_band(new_ws: Worksheet, base_font: Font) -> None:
    bands = [
        (range(6, 9), 'light_blue', None),
        (range(12, 18), 'soft_orange', None),
        (range(18, 21), 'light_yellow', None),
        (range(21, 24), 'soft_gray', None),
        (range(27, 28), 'soft_green', None),
        (
            range(28, 30),
            'yellow',
            Font(name=C.FONT_NAME, size=14, bold=True, color='ff3838'),
        ),
        (
            range(30, 31),
            'violet',
            Font(name=C.FONT_NAME, size=10, bold=True, color='fbff1f'),
        ),
        (
            range(31, 32),
            'blue',
            Font(name=C.FONT_NAME, size=10, bold=True, color='fbff1f'),
        ),
    ]
    for row_idx in range(1, new_ws.max_row + 1):
        for col_idx in range(1, new_ws.max_column + 1):
            cell = new_ws.cell(row=row_idx, column=col_idx)
            cell.font = base_font
            for col_range, color_name, override_font in bands:
                if col_idx in col_range:
                    cell.fill = _fill(color_name)
                    if override_font:
                        cell.font = override_font


def _apply_number_formats_and_borders(new_ws: Worksheet, header_font: Font) -> None:
    thin = Side(border_style='thin', color='242424')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alignment = Alignment(horizontal='center', vertical='center')  # noqa: F811

    for row_idx in range(1, new_ws.max_row + 1):
        for col_idx in range(1, new_ws.max_column + 1):
            cell = new_ws.cell(row=row_idx, column=col_idx)
            cell.number_format = '###,##0.00'
            cell.border = border

            if 5 <= col_idx <= 35:
                cell.alignment = alignment

            if row_idx == 2:
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            if row_idx == 1:
                cell.fill = _fill('light_green')
                cell.font = header_font
                cell.number_format = '0.00'

            if 6 <= col_idx <= 14:
                cell.number_format = '#,###,##0'

            if 16 <= col_idx <= 33 or col_idx in (10, 11):
                cell.number_format = '#,##0'
            if 15 <= col_idx <= 17 or col_idx == 27:
                cell.number_format = '0.0%'
            if col_idx in (39, 41):
                cell.number_format = '#,##0'

    title_cell = new_ws.cell(row=1, column=1, value='Nazev sestavu')
    new_ws.merge_cells('A1:D1')
    title_cell.font = Font(name=C.FONT_NAME, size=16, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')

    sort_flag = new_ws.cell(row=1, column=7, value='RAZENO!')
    sort_flag.fill = _fill('yellow')
    sort_flag.font = Font(name=C.FONT_NAME, size=12, bold=True, color='ff3838')
    sort_flag.alignment = Alignment(horizontal='center', vertical='center')

    date_cell = new_ws.cell(row=1, column=8, value=datetime.now().date())
    date_cell.number_format = 'dd.mm.yyyy'
    date_cell.font = Font(name=C.FONT_NAME, size=10, bold=True)
    date_cell.alignment = Alignment(horizontal='center', vertical='center')


def _style_threshold_row(new_ws: Worksheet) -> None:
    highlight_fill = PatternFill(start_color='a6d5ff', end_color='a6d5ff', fill_type='solid')
    highlight_font = Font(name=C.FONT_NAME, size=12, bold=True)

    for cell in new_ws['E1:F1'][0]:
        cell.fill = highlight_fill
        cell.font = highlight_font
        if cell.column == 5:
            cell.value = 'Vic než:'
        if cell.column == 6:
            cell.value = 'napiš tady'
            cell.number_format = '#,##0'
            new_ws.conditional_formatting.add(
                f'F3:H{new_ws.max_row}',
                CellIsRule(
                    operator='greaterThan',
                    formula=['=$F$1'],
                    stopIfTrue=True,
                    fill=highlight_fill,
                    font=highlight_font,
                ),
            )


def _style_total_cbm_cells(new_ws: Worksheet) -> None:
    for cell in new_ws['AB1:AC1'][0]:
        cell.fill = _fill('blue')
        if cell.column == 29:
            cell.font = Font(name=C.FONT_NAME, size=16, bold=True, color='fbff1f')
            cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
            cell.number_format = '0.00'
        else:
            cell.font = Font(name=C.FONT_NAME, size=11, bold=True, color='fbff1f')
            cell.alignment = Alignment(horizontal='right', vertical='center')


# ---------------------------------------------------------------------------
# Misc
# ---------------------------------------------------------------------------


def _build_output_path(input_path: str) -> str:
    folder = os.path.dirname(input_path)
    name = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(folder, f'{name}_processed.xlsx')
