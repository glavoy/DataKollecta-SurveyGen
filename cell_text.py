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
from datetime import date, datetime, time
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

_LINE_BREAK_RE = re.compile(r"\r\n|\n|\r")


def to_str(value: Any) -> str:
    """A cell value as the string a data dictionary author meant to type.

    The float case is the one that matters: openpyxl reads a whole number in a
    General-formatted cell as `1.0`, and `str(1.0)` is `"1.0"` -- which then
    becomes a response code of "1.0" instead of "1", or a MaxCharacters of
    "3.0" that no downstream check can parse.

    Dates are the other case, and the reason to handle them here rather than
    at each call site. openpyxl reads a date-formatted cell as a `datetime`, so
    `str()` gave "2026-07-13 00:00:00" where the author typed a date. The
    documented dictionary format is `yyyy-mm-dd` (README "Hard-coded Date
    Format"), which is exactly what `ExcelReader.HARDCODED_DATE_RE` matches, so
    a midnight datetime is rendered date-only.

    A previous version of this docstring said the problem was that such a cell
    "fails with a misleading message" in a range. That is the *loud* path, and
    it was the wrong one to worry about: a range cell that does not match the
    regex produces a validation error, which is confusing but visible. The
    silent path is the one that mattered -- the same value flows into
    QuestionText, into response codes, and into a constant calculation's value,
    none of which check the format. A date typed as a range bound got an error;
    a date typed anywhere else shipped "2026-07-13 00:00:00" into the field.

    A datetime carrying an actual time is left as date-and-time. No dictionary
    column is documented to hold one, so there is no author intent to honour;
    rendering it in full keeps the value visible to whatever check reads it
    next instead of quietly discarding the time. Microseconds are dropped,
    which `str()` would have kept, so the output is stable.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    # datetime before date -- datetime is a subclass of date, so the order is
    # what decides which branch a datetime takes.
    if isinstance(value, datetime):
        if value.time() == time(0, 0):
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
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
