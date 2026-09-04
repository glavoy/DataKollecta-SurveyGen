"""Reading a spreadsheet cell as text.

One module because there were two near-identical copies of this: ExcelReader's
`_to_str`/`_get_cell_trim` and CrfReader's `_cell_trim`. They agreed by
coincidence rather than by construction, and a divergence between them is not
hypothetical -- the `crfs` sheet and the `_dd` sheets are read by different
code paths, so the same cell could be turned into two different strings
depending on which one looked at it.

The line splitter lives here for a sharper reason: the VALIDATOR
(excel_reader) and the EMITTER (xml_generator) each had their own copy of the
same regex, and they have to agree about how many response lines a cell holds.
C4 was exactly what happens when a validator and a generator disagree about
the same cell -- the validator passed and the generator dropped the element.

`skip_parser.split_skip_lines` is deliberately NOT folded in: it filters on
`line.strip()` rather than `line`, so a whitespace-only line is dropped there
and kept here. That difference matters. A blank-looking response line has to
survive as far as the static-response check, which reports "please remove
leading spaces" -- useful feedback that silently discarding the line would
lose.
"""

from __future__ import annotations

import re
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

_LINE_BREAK_RE = re.compile(r"\r\n|\n|\r")


def to_str(value: Any) -> str:
    """A cell value as the string a data dictionary author meant to type.

    The float case is the one that matters: openpyxl reads a whole number in a
    General-formatted cell as `1.0`, and `str(1.0)` is `"1.0"` -- which then
    becomes a response code of "1.0" instead of "1", or a MaxCharacters of
    "3.0" that no downstream check can parse.

    Still open, and deliberately not changed here: a date-formatted cell comes
    back as a `datetime`, so `str()` gives "2026-07-13 00:00:00" rather than
    "2026-07-13". Fixing it would change generated output for any dictionary
    that formats a range cell as a date, so it needs its own verification pass
    against the real dictionaries rather than riding along with a
    consolidation.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def cell_raw(worksheet: Worksheet, row: int, col: int) -> str:
    """A cell as text, with surrounding whitespace left alone.

    Used where the whitespace is part of what is being validated -- the
    Responses column, where a leading space on an option is an error worth
    reporting rather than trimming away.
    """
    return to_str(worksheet.cell(row=row, column=col).value)


def cell_trim(worksheet: Worksheet, row: int, col: int) -> str:
    """A cell as text, trimmed. The default for every ordinary column."""
    return cell_raw(worksheet, row, col).strip()


def split_cell_lines(text: str) -> list[str]:
    """One cell, split into its lines, keeping every non-empty one.

    Handles all three line endings, because a dictionary edited on Windows and
    one edited on a Mac must read identically.
    """
    return [line for line in _LINE_BREAK_RE.split(text) if line]
