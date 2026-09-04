"""Parsing and validation of a `calc:` block from the Responses column.

Split out of `excel_reader.py`, which was 1,480 lines. This is the whole
calculation cluster: one entry point (`_parse_automatic_calculation`, called
from `ExcelReader.create_question_list` when a Responses cell starts with
`calc:`), and one exit (`_validate_calculation_fields`, tail-called at the end
of it and from nowhere else).

**This is a mixin, not a module of free functions, and that is load-bearing.**
`tests/test_calculation_registry.py` asserts coverage by reading source text:

    inspect.getsource(excel_reader.ExcelReader._validate_calculation_fields)

and checking that the body names `CalculationType.<NAME>` for all twelve
calculable types. Inheriting the method keeps it reachable at that exact
attribute path and keeps `inspect.getsource` working, so the test needs no
change. It also means two things must not be "tidied up" here:

  * `_validate_calculation_fields` stays **one method**. Splitting it into
    per-type helpers, or replacing the `elif` chain with a dispatch table,
    breaks that test while changing no behaviour -- and the chain is
    deliberately explicit anyway, because the checks are genuinely
    heterogeneous per type rather than a table's worth of the same check.
  * The error messages stay in this file and this file is listed in
    `SCANNED_MODULES` in `tests/test_error_message_shape.py`, which asserts
    that more than 80 `_error` call sites are found across the reader modules.
    Twenty-nine of the ninety-four live here. (Note for anyone editing this
    docstring: that test scrapes source text, so spelling the call out in full
    here would be counted as a site with an empty message.)

`DATE_RANGE_RE` is deliberately *not* moved. `_validate_calculation_fields`
uses it to check a `date_offset` value, and `ExcelReader._check_date_range`
uses it for the LowerRange/UpperRange columns -- one regex, two callers in
different clusters. It stays on `ExcelReader` and resolves through `self` here,
the same object it resolved through before the split.
"""

from __future__ import annotations

import re

from models import (
    CALCULATION_TYPE_BY_ALIAS,
    calculation_alias_list,
    CalculationParameter,
    CalculationPart,
    CalculationType,
    CaseCondition,
    Question,
)


class CalculationParsingMixin:
    """The `calc:` half of `ExcelReader`. Not usable on its own.

    Expects the host class to provide `_error`, `_split_lines`, `logstring`
    and `DATE_RANGE_RE`.
    """

    PARAMETER_RE = re.compile(r"^(@?\w+)\s*=\s*(\w+)$")
    # Alternatives are ordered longest-first for the same reason as
    # FILTER_MATCH_RE in response_parser; "does not contain" embeds its own \s+
    # since it is three words, not a single token like the other operators.
    WHEN_CONDITION_RE = re.compile(
        r"^(\w+)\s+(=|!=|<>|>=|<=|>|<|(?i:does\s+not\s+contain|contains))\s+(.+?)\s*=>\s*(.+)$"
    )

    def _parse_automatic_calculation(self, responses: str, question: Question, worksheet: str, fieldname: str) -> None:
        current_calc = ""
        when_lines: list[str] = []
        part_lines: list[str] = []
        for line in self._split_lines(responses):
            trimmed = line.strip()
            if not trimmed:
                continue
            parts = trimmed.split(":", 1)
            if len(parts) != 2:
                self._error(
                    f"ERROR - Calculation: Invalid line format for FieldName '{fieldname}' in worksheet '{worksheet}': '{trimmed}'"
                )
                continue
            key = parts[0].strip().lower()
            value = parts[1].strip()

            if key == "calc":
                current_calc = value.lower()
                # One shared table (models.CALCULATION_ALIASES), and the valid
                # list in the message derived from it -- the dict and the
                # hand-written list of the same twelve words used to be two
                # separate places to keep in step.
                if current_calc in CALCULATION_TYPE_BY_ALIAS:
                    question.calculationType = CALCULATION_TYPE_BY_ALIAS[current_calc]
                else:
                    self._error(
                        f"ERROR - Calculation: Invalid calculation type '{value}' for FieldName '{fieldname}' in worksheet '{worksheet}'. "
                        f"Must be {calculation_alias_list()}."
                    )
            elif key == "sql":
                question.calculationQuerySql = value
            elif key == "param":
                self._parse_parameter(value, question, worksheet, fieldname)
            elif key == "when":
                when_lines.append(value)
            elif key == "else":
                if current_calc == "case":
                    question.calculationCaseElse = self._parse_result_value(value)
            elif key == "value":
                if current_calc in {"constant", "age_from_date", "age_at_date", "date_offset", "date_diff"}:
                    question.calculationConstantValue = value
            elif key == "field":
                if current_calc in {"lookup", "age_from_date", "age_at_date", "date_offset", "date_diff", "date_part"}:
                    question.calculationLookupField = value
            elif key == "unit":
                if current_calc == "date_diff":
                    question.calculationUnit = value
                elif current_calc == "date_part":
                    normalized = value.strip().lower()
                    allowed = {"yyyy", "yy", "mm", "dd", "doy"}
                    if normalized in allowed:
                        question.calculationUnit = normalized
                    else:
                        self._error(
                            f"ERROR - Calculation: Invalid date_part unit '{value}' for FieldName '{fieldname}' in worksheet '{worksheet}'. "
                            "Must be 'yyyy', 'yy', 'mm', 'dd', or 'doy'."
                        )
            elif key == "operator":
                if current_calc == "math":
                    if value in {"+", "-", "*", "/"}:
                        question.calculationMathOperator = value
                    else:
                        self._error(
                            f"ERROR - Calculation: Invalid math operator '{value}' for FieldName '{fieldname}' in worksheet '{worksheet}'. Must be +, -, *, or /."
                        )
            elif key == "separator":
                if current_calc in {"concat", "age_at_date"}:
                    question.calculationConcatSeparator = value
            elif key == "part":
                part_lines.append(value)
            else:
                self.logstring.append(
                    f"WARNING - Calculation: Unknown calculation key '{key}' for FieldName '{fieldname}' in worksheet '{worksheet}'."
                )

        if current_calc == "case":
            for when_line in when_lines:
                self._parse_when_condition(when_line, question, worksheet, fieldname)

        if current_calc in {"math", "concat"}:
            for part_line in part_lines:
                part = self._parse_part_line(part_line, worksheet, fieldname)
                if not part:
                    continue
                if current_calc == "math":
                    question.calculationMathParts.append(part)
                else:
                    question.calculationConcatParts.append(part)

        self._validate_calculation_fields(question, worksheet, fieldname)

    def _parse_parameter(self, param_str: str, question: Question, worksheet: str, fieldname: str) -> None:
        match = self.PARAMETER_RE.match(param_str)
        if not match:
            self._error(
                f"ERROR - Calculation: Invalid parameter format '{param_str}' for FieldName '{fieldname}' in worksheet '{worksheet}'. "
                "Expected format: '@paramName = fieldName'."
            )
            return
        name = match.group(1).strip()
        if not name.startswith("@"):
            name = "@" + name
        question.calculationQueryParameters.append(
            CalculationParameter(name=name, fieldName=match.group(2).strip())
        )

    def _parse_when_condition(self, when_str: str, question: Question, worksheet: str, fieldname: str) -> None:
        match = self.WHEN_CONDITION_RE.match(when_str)
        if not match:
            self._error(
                f"ERROR - Calculation: Invalid when condition format '{when_str}' for FieldName '{fieldname}' in worksheet '{worksheet}'. "
                "Expected format: 'field operator value => result'."
            )
            return
        question.calculationCaseConditions.append(
            CaseCondition(
                field=match.group(1).strip(),
                operator=match.group(2).strip(),
                value=match.group(3).strip(),
                result=self._parse_result_value(match.group(4).strip()),
            )
        )

    @staticmethod
    def _parse_result_value(value: str) -> CalculationPart:
        return CalculationPart(type=CalculationType.CONSTANT, constantValue=value)

    def _parse_part_line(self, part_line: str, worksheet: str, fieldname: str) -> CalculationPart | None:
        words = part_line.split(" ", 1)
        if len(words) < 2:
            self._error(
                f"ERROR - Calculation: Invalid part format '{part_line}' for FieldName '{fieldname}' in worksheet '{worksheet}'. "
                "Expected 'type value'."
            )
            return None
        part_type = words[0].strip().lower()
        part_value = words[1].strip()
        if part_type == "constant":
            return CalculationPart(type=CalculationType.CONSTANT, constantValue=part_value)
        if part_type == "lookup":
            return CalculationPart(type=CalculationType.LOOKUP, lookupField=part_value)
        if part_type == "query":
            return CalculationPart(type=CalculationType.QUERY, querySql=part_value)

        self._error(
            f"ERROR - Calculation: Invalid part type '{part_type}' for FieldName '{fieldname}' in worksheet '{worksheet}'. "
            "Must be 'constant', 'lookup', or 'query'."
        )
        return None

    def _validate_calculation_fields(self, question: Question, worksheet: str, fieldname: str) -> None:
        ctype = question.calculationType
        if ctype == CalculationType.QUERY and not question.calculationQuerySql:
            self._error(
                f"ERROR - Calculation: Query calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                "is missing required 'sql' field."
            )
        elif ctype == CalculationType.CASE and len(question.calculationCaseConditions) == 0:
            self._error(
                f"ERROR - Calculation: Case calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                "is missing 'when' conditions."
            )
        elif ctype == CalculationType.CONSTANT and not question.calculationConstantValue:
            self._error(
                f"ERROR - Calculation: Constant calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                "is missing required 'value' field."
            )
        elif ctype == CalculationType.LOOKUP and not question.calculationLookupField:
            self._error(
                f"ERROR - Calculation: Lookup calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                "is missing required 'field' field."
            )
        elif ctype == CalculationType.MATH:
            if not question.calculationMathOperator:
                self._error(
                    f"ERROR - Calculation: Math calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'operator' field."
                )
            if len(question.calculationMathParts) < 2:
                self._error(
                    f"ERROR - Calculation: Math calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "must have at least 2 parts."
                )
        elif ctype == CalculationType.CONCAT and len(question.calculationConcatParts) == 0:
            self._error(
                f"ERROR - Calculation: Concat calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                "must have at least 1 part."
            )
        elif ctype == CalculationType.AGE_FROM_DATE:
            if not question.calculationLookupField:
                self._error(
                    f"ERROR - Calculation: AgeFromDate calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'field' field."
                )
            if not question.calculationConstantValue:
                self._error(
                    f"ERROR - Calculation: AgeFromDate calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'value' field."
                )
        elif ctype == CalculationType.AGE_AT_DATE:
            if not question.calculationLookupField:
                self._error(
                    f"ERROR - Calculation: AgeAtDate calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'field' field."
                )
            if not question.calculationConstantValue:
                self._error(
                    f"ERROR - Calculation: AgeAtDate calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'value' field."
                )
        elif ctype == CalculationType.DATE_OFFSET:
            if not question.calculationLookupField:
                self._error(
                    f"ERROR - Calculation: DateOffset calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'field' field."
                )
            if not question.calculationConstantValue:
                self._error(
                    f"ERROR - Calculation: DateOffset calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'value' field."
                )
            elif not self.DATE_RANGE_RE.fullmatch(question.calculationConstantValue):
                self._error(
                    f"ERROR - Calculation: DateOffset calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    f"has invalid 'value' format: {question.calculationConstantValue}. Expected format like '+28d', '-1y', etc."
                )
        elif ctype == CalculationType.DATE_DIFF:
            if not question.calculationLookupField:
                self._error(
                    f"ERROR - Calculation: DateDiff calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'field' field (start date)."
                )
            if not question.calculationConstantValue:
                self._error(
                    f"ERROR - Calculation: DateDiff calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'value' field (end date)."
                )
            if not question.calculationUnit:
                self._error(
                    f"ERROR - Calculation: DateDiff calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'unit' field."
                )
            elif question.calculationUnit.lower() not in {"d", "w", "m", "y"}:
                self._error(
                    f"ERROR - Calculation: DateDiff calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    f"has invalid 'unit': {question.calculationUnit}. Must be 'd', 'w', 'm', or 'y'."
                )
        elif ctype == CalculationType.DATE_PART:
            # An invalid unit value is already rejected where it's read (the
            # "unit" key handler above), the same way an invalid math
            # operator is -- this only catches the key being absent entirely.
            if not question.calculationLookupField:
                self._error(
                    f"ERROR - Calculation: DatePart calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'field' field."
                )
            if not question.calculationUnit:
                self._error(
                    f"ERROR - Calculation: DatePart calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    "is missing required 'unit' field."
                )
        elif ctype == CalculationType.TIMESTAMP:
            if question.fieldType.strip().lower() != "datetime":
                self._error(
                    f"ERROR - Calculation: Timestamp calculation for FieldName '{fieldname}' in worksheet '{worksheet}' "
                    f"requires FieldType 'datetime', got '{question.fieldType}'."
                )
