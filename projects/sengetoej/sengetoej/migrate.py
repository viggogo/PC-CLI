"""One-time restructuring of the Sengetøj sheet.

Two jobs, one save:

  1. Move the free-text comments one column right, D -> E. Dropping the new
     `cli` column into C would otherwise consume the blank column that
     separates the table from the comments, leaving the table butting
     straight up against them.

  2. Add the `cli` column to Table2 -- header, ref, autofilter and
     tableColumns, all of which must agree or Excel calls the file corrupt.

Run once, deliberately, never from the CLI. It takes a backup first.
"""

import datetime as dt
import shutil
from copy import copy
from pathlib import Path

from openpyxl.worksheet.formula import ArrayFormula

from . import sheet

COMMENT_COL_BEFORE = 4  # D
COMMENT_COL_AFTER = 5   # E


def comment_rows(ws, col: int) -> list[int]:
    """Populated rows in `col`, searched to a bound derived from the data.

    The comments include headings above the data and remarks beside it, and
    they are sparse, so this walks a fixed window rather than stopping at the
    first blank.
    """
    _, next_row = sheet.dates_and_next_row(ws)
    return [r for r in range(1, next_row + sheet.BLANK_RUN)
            if ws.cell(r, col).value is not None]


def move_comments(ws) -> int:
    """Move every populated comment one column right. Returns the count.

    Deliberately cell-by-cell rather than ws.move_range: move_range with
    translate=True would rewrite the counter's INDEX(A:A,COUNTA(A:A)) into
    INDEX(B:B,COUNTA(B:B)), and even with translate=False it would not
    rebuild the ArrayFormula's own ref.
    """
    rows = comment_rows(ws, COMMENT_COL_BEFORE)

    for row in rows:
        src = ws.cell(row, COMMENT_COL_BEFORE)
        dst = ws.cell(row, COMMENT_COL_AFTER)

        value = src.value
        if isinstance(value, ArrayFormula):
            # The ref names the cell the formula lives in, so it moves too.
            value = ArrayFormula(ref=f"{dst.column_letter}{row}", text=value.text)

        dst.value = value
        dst._style = copy(src._style)

        src.value = None
        # Reset rather than leave the old style behind: the counter cell is
        # bold on a yellow fill, and an emptied cell that kept that fill
        # would show as a stray coloured block in the spacer column.
        # "Normal" is openpyxl's built-in default and clears both.
        src.style = "Normal"

    before = ws.cell(1, COMMENT_COL_BEFORE).column_letter
    after = ws.cell(1, COMMENT_COL_AFTER).column_letter
    if before in ws.column_dimensions:
        ws.column_dimensions[after].width = ws.column_dimensions[before].width
        del ws.column_dimensions[before]

    return len(rows)
