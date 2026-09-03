from __future__ import annotations

import json

from openpyxl.worksheet.worksheet import Worksheet

from models import Crf, IdConfig, IdConfigField

# The order is load-bearing: every field below is read by column *position*,
# so inserting a column in the sheet shifts everything after it. Before the
# header check in `read_crfs_worksheet`, adding a column before `idconfig`
# silently landed `displayname` in `primarykey` and the idconfig JSON in
# `isbase`, and the resulting manifest was well-formed JSON holding wrong
# values for every form. Documented in README "The CRFS Worksheet".
CRFS_COLUMN_NAMES = [
    "display_order",
    "tablename",
    "displayname",
    "primarykey",
    "idconfig",
    "isbase",
    "linkingfield",
    "parenttable",
    "incrementfield",
    "requireslink",
    "repeat_count_field",
    "auto_start_repeat",
    "repeat_enforce_count",
    "display_fields",
    "entry_condition",
]
NUMBER_OF_CRFS_COLUMNS = len(CRFS_COLUMN_NAMES)


class CrfReader:
    @staticmethod
    def read_crfs_worksheet(worksheet: Worksheet) -> tuple[list[Crf], list[str]]:
        """Read the `crfs` sheet into models, plus any errors found.

        Returns `(crfs, errors)`. A non-empty `errors` must stop the build:
        every problem it reports produces a manifest that is valid JSON and
        wrong, which no later stage can detect.
        """
        errors: list[str] = []

        headers = [_cell_trim(worksheet, 1, i + 1).lower() for i in range(NUMBER_OF_CRFS_COLUMNS)]
        if headers != CRFS_COLUMN_NAMES:
            errors.append(
                "ERROR - crfs: The header names in the 'crfs' worksheet are incorrect. "
                f"Expected, in this order: {', '.join(CRFS_COLUMN_NAMES)}. "
                f"Found: {', '.join(h or '(blank)' for h in headers)}. "
                "Every column is read by position, so an inserted or reordered "
                "column silently writes the wrong value into every form."
            )
            # Positions cannot be trusted, so reading rows would invent data.
            return [], errors

        crfs: list[Crf] = []
        for row_idx in range(2, worksheet.max_row + 1):
            cells = [_cell_trim(worksheet, row_idx, i + 1) for i in range(NUMBER_OF_CRFS_COLUMNS)]

            # openpyxl counts a row that carries only formatting -- a border or
            # a fill, common at the bottom of a hand-maintained sheet -- toward
            # max_row. Such a row used to become an all-None Crf, which
            # `json_generator` faithfully serialized as a `{}` entry in the
            # manifest's crfs array.
            if not any(cells):
                continue

            crf = Crf(
                display_order=_nullable_int(cells[0]),
                tablename=_null_if_empty(cells[1]),
                displayname=_null_if_empty(cells[2]),
                primarykey=_null_if_empty(cells[3]),
                isbase=_nullable_int(cells[5]),
                linkingfield=_null_if_empty(cells[6]),
                parenttable=_null_if_empty(cells[7]),
                incrementfield=_null_if_empty(cells[8]),
                requireslink=_nullable_int(cells[9]),
                repeat_count_field=_null_if_empty(cells[10]),
                auto_start_repeat=_nullable_int(cells[11]),
                repeat_enforce_count=_nullable_int(cells[12]),
                display_fields=_null_if_empty(cells[13]),
                entry_condition=_null_if_empty(cells[14]),
            )

            idconfig_json = cells[4]
            if idconfig_json:
                try:
                    crf.idconfig = _parse_idconfig(idconfig_json)
                except Exception as ex:
                    # Previously `except Exception: crf.idconfig = None`, with
                    # no log line at all. idconfig is the highest-consequence
                    # cell in the file -- it builds the study's primary keys --
                    # so one smart quote or missing brace made the manifest
                    # omit it and the app generate a different ID scheme than
                    # intended, silently.
                    errors.append(
                        f"ERROR - crfs: Row {row_idx} (tablename "
                        f"'{crf.tablename or '(blank)'}') has an idconfig that is not valid "
                        f"JSON: {ex}. Value: {idconfig_json}"
                    )

            crfs.append(crf)

        return crfs, errors


def _parse_idconfig(idconfig_json: str) -> IdConfig:
    raw = json.loads(idconfig_json)
    if not isinstance(raw, dict):
        raise ValueError("expected a JSON object")

    fields = raw.get("fields")
    parsed_fields = None
    if isinstance(fields, list):
        parsed_fields = [
            IdConfigField(name=f.get("name", ""), length=int(f.get("length", 0))) for f in fields
        ]

    return IdConfig(
        prefix=raw.get("prefix"),
        fields=parsed_fields,
        incrementLength=raw.get("incrementLength"),
    )


def _cell_trim(worksheet: Worksheet, row: int, col: int) -> str:
    value = worksheet.cell(row=row, column=col).value
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _nullable_int(value: str) -> int | None:
    try:
        return int(value)
    except Exception:
        return None


def _null_if_empty(value: str) -> str | None:
    return value if value else None
