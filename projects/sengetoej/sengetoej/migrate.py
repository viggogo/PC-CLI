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
    """Every populated row in `col`, exactly.

    Enumerates ws._cells rather than walking a row window: the comments are
    sparse and sit both above and beside the data, so any window is a guess,
    and a comment outside it would be silently left behind by a migration
    that only runs once. This also avoids ws.cell(), which CREATES empty
    cells as a side effect of reading them.
    """
    return sorted(row for (row, column), cell in ws._cells.items()
                  if column == col and cell.value is not None)


def move_comments(ws) -> int:
    """Move every populated comment one column right. Returns the count.

    Deliberately cell-by-cell rather than ws.move_range: move_range with
    translate=True would rewrite the counter's INDEX(A:A,COUNTA(A:A)) into
    INDEX(B:B,COUNTA(B:B)), and even with translate=False it would not
    rebuild the ArrayFormula's own ref.

    Two guards, because this runs once against an irreplaceable file and
    must fail loudly rather than silently mangle it:
      - a populated destination cell stops the move before it overwrites
        anything;
      - a non-empty source column after the loop means comment_rows missed
        something, and that is raised rather than swallowed.
    """
    rows = comment_rows(ws, COMMENT_COL_BEFORE)

    for row in rows:
        src = ws.cell(row, COMMENT_COL_BEFORE)
        dst = ws.cell(row, COMMENT_COL_AFTER)

        if dst.value is not None:
            raise RuntimeError(
                f"kolonne E har allerede en værdi i række {row}: "
                f"{dst.value!r} -- flytning afbrudt for ikke at overskrive")

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

    leftover = comment_rows(ws, COMMENT_COL_BEFORE)
    if leftover:
        raise RuntimeError(
            f"kolonne D er ikke tom efter flytningen: rækker {leftover}")

    before_letter = ws.cell(1, COMMENT_COL_BEFORE).column_letter
    after_letter = ws.cell(1, COMMENT_COL_AFTER).column_letter
    if before_letter in ws.column_dimensions:
        src_dim = ws.column_dimensions[before_letter]
        dst_dim = ws.column_dimensions[after_letter]
        # Width, visibility and outline grouping travel with the comments.
        # Not `style`: that is an index into the workbook's shared style
        # table, and copying it across columns is riskier than it is worth.
        dst_dim.width = src_dim.width
        dst_dim.hidden = src_dim.hidden
        dst_dim.outlineLevel = src_dim.outlineLevel
        dst_dim.bestFit = src_dim.bestFit
        del ws.column_dimensions[before_letter]

    return len(rows)
