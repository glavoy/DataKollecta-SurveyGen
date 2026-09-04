"""Parsing of a `source:` block from the Responses column.

Split out of `excel_reader.py`. This is the dynamic-response cluster: a single
entry point (`_parse_dynamic_responses`, called from
`ExcelReader.create_question_list` when a Responses cell starts with `source:`)
plus the two filter-operator helpers it uses and nothing else does.

A mixin, for the same reason `CalculationParsingMixin` is one -- the methods
keep resolving at their original attribute path on `ExcelReader`, so nothing
about how they are reached, tested or introspected changes.

`_parse_operator` raises `ValueError` rather than recording a validation error,
and that is deliberate: it is reached only for an operator `FILTER_MATCH_RE`
already matched, so a spelling it cannot translate means the two have drifted
apart in code. That is a programming error, not a dictionary author's mistake,
and it should stop the run loudly rather than land in the log next to things a
user can fix.

Five of the ninety-four `_error` call sites live here; this module is listed in
`SCANNED_MODULES` in `tests/test_error_message_shape.py`.
"""

from __future__ import annotations

import re

from models import (
    Filter,
    Question,
    ResponseSourceType,
)


class ResponseParsingMixin:
    """The `source:` half of `ExcelReader`. Not usable on its own.

    Expects the host class to provide `_error`, `_split_lines` and `logstring`.
    """

    # Alternatives are ordered longest-first so '>=' is not consumed as '>',
    # which would leave the '=' stranded at the front of the filter value.
    # The word operators are matched case-insensitively.
    FILTER_MATCH_RE = re.compile(
        r"^(\w+)\s*(?:((?i:not\s+in|in)|>=|<=|!=|<>|=|>|<)\s*)?(.+)$"
    )
    # Operator spellings a dictionary author plausibly writes that this
    # filter grammar does NOT support. They matter because FILTER_MATCH_RE's
    # operator group is optional: with no recognised operator the whole
    # remainder becomes the *value*, so `filter:region like North` parsed as
    # `region = 'like North'` and came back from the device as an empty
    # response list, with nothing said at generation time. Naming them is the
    # difference between an error the author can fix and a question that is
    # simply blank in the field.
    UNSUPPORTED_FILTER_OPERATORS = {
        "like",
        "not like",
        "ilike",
        "is",
        "is not",
        "between",
        "contains",
        "does not contain",
        "starts with",
        "ends with",
        "matches",
        "regexp",
        "glob",
    }
    # A value that begins with one of these is a mistyped operator rather than
    # data -- `=<`, `==`, `!`, `~` and friends all land here, because the
    # regex consumed only the part of them it recognised.
    OPERATOR_LEAD_CHARACTERS = "=<>!~"

    def _unsupported_filter_operator(self, value: str, matched_operator: str | None) -> str | None:
        """The operator-shaped text left stranded in a filter's value, if any.

        Only ever inspects the value, never rewrites it: a filter value may
        legitimately contain spaces (`filter:district North West` means
        `district = 'North West'`), so the test is on the *leading* token
        alone.
        """
        if not value:
            return None

        if value[0] in self.OPERATOR_LEAD_CHARACTERS:
            # e.g. `x =< 5`, where the regex matched `=` and left `< 5`.
            stranded = value[: len(value) - len(value.lstrip(self.OPERATOR_LEAD_CHARACTERS))]
            return f"{matched_operator.strip()}{stranded}" if matched_operator else stranded

        # Word operators are only a possibility when the regex found no
        # operator at all -- `in`/`not in` would already have matched.
        if matched_operator:
            return None

        tokens = value.split()
        for word_count in (3, 2, 1):
            if len(tokens) > word_count:
                candidate = " ".join(tokens[:word_count]).lower()
                if candidate in self.UNSUPPORTED_FILTER_OPERATORS:
                    return " ".join(tokens[:word_count])
        return None

    def _parse_operator(self, op: str) -> str:
        # Collapse "NOT  IN" and similar spellings to a single canonical form
        op = " ".join(op.split()).lower() if op else "="
        if op == ">":
            return "&gt;"
        if op == "<":
            return "&lt;"
        if op == ">=":
            return "&gt;="
        if op == "<=":
            return "&lt;="
        if op == "<>":
            # '<' is not legal inside an XML attribute value
            return "&lt;&gt;"
        if op in {"=", "!=", "in", "not in"}:
            return op
        raise ValueError(
            f"Unsupported filter operator {op!r}. FILTER_MATCH_RE only matches "
            "the operators handled above, so reaching here means the regex and "
            "this function have drifted apart."
        )

    def _parse_dynamic_responses(self, responses: str, question: Question, worksheet: str, fieldname: str) -> None:
        for line in self._split_lines(responses):
            trimmed = line.strip()
            if not trimmed:
                continue
            parts = trimmed.split(":", 1)
            if len(parts) != 2:
                self._error(
                    f"ERROR - Responses: Invalid dynamic response line format for FieldName '{fieldname}' "
                    f"in worksheet '{worksheet}': '{trimmed}'"
                )
                continue
            key = parts[0].strip().lower()
            value = parts[1].strip()

            if key == "source":
                lowered = value.lower()
                if lowered == "csv":
                    question.responseSourceType = ResponseSourceType.CSV
                elif lowered == "database":
                    question.responseSourceType = ResponseSourceType.DATABASE
                else:
                    self._error(
                        f"ERROR - Responses: Invalid source type '{value}' for FieldName '{fieldname}' in worksheet '{worksheet}'. "
                        "Must be 'csv' or 'database'."
                    )
            elif key == "file":
                question.responseSourceFile = value
            elif key == "table":
                question.responseSourceTable = value
            elif key == "filter":
                match = self.FILTER_MATCH_RE.match(value)
                if match:
                    filter_value = match.group(3).strip()
                    bad_operator = self._unsupported_filter_operator(
                        filter_value, matched_operator=match.group(2)
                    )
                    if bad_operator:
                        self._error(
                            f"ERROR - Responses: Unsupported filter operator '{bad_operator}' for FieldName "
                            f"'{fieldname}' in worksheet '{worksheet}': '{value}'. "
                            "Supported operators are =, !=, <>, <, <=, >, >=, in and not in. "
                            "Without one of those the whole phrase is taken as the value to match, "
                            "so the response list comes back empty at interview time."
                        )
                    else:
                        question.responseFilters.append(
                            Filter(
                                column=match.group(1).strip(),
                                operator=self._parse_operator(match.group(2) if match.group(2) else "="),
                                value=filter_value,
                            )
                        )
                else:
                    self._error(
                        f"ERROR - Responses: Invalid filter format for FieldName '{fieldname}' in worksheet '{worksheet}': "
                        f"'{value}'. Expected 'column [operator] value'."
                    )
            elif key == "display":
                question.responseDisplayColumn = value
            elif key == "value":
                question.responseValueColumn = value
            elif key == "distinct":
                lowered = value.lower()
                if lowered == "true":
                    question.responseDistinct = True
                elif lowered == "false":
                    question.responseDistinct = False
                else:
                    self._error(
                        f"ERROR - Responses: Invalid boolean value for 'distinct' for FieldName '{fieldname}' in worksheet '{worksheet}'. "
                        "Must be 'true' or 'false'."
                    )
            elif key == "empty_message":
                question.responseEmptyMessage = value
            elif key == "dont_know":
                parts = value.split(",", 1)
                question.responseDontKnowValue = parts[0].strip()
                if len(parts) > 1:
                    question.responseDontKnowLabel = parts[1].strip()
            elif key == "not_in_list":
                parts = value.split(",", 1)
                question.responseNotInListValue = parts[0].strip()
                if len(parts) > 1:
                    question.responseNotInListLabel = parts[1].strip()
            else:
                self.logstring.append(
                    f"WARNING - Responses: Unknown dynamic response key '{key}' for FieldName '{fieldname}' in worksheet '{worksheet}'."
                )
