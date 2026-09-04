from __future__ import annotations

import re

from openpyxl.worksheet.worksheet import Worksheet

from models import (
    KNOWN_AUTOMATIC_FIELDS,
    RESERVED_SYSTEM_FIELDS,
    TRAILING_SYSTEM_FIELD_NAMES,
    CalculationPart,
    CalculationType,
    Filter,
    Question,
    ResponseSourceType,
)
from calculation_parser import CalculationParsingMixin
from cell_text import cell_raw, cell_trim, split_cell_lines, to_str
from skip_parser import parse_skip, split_skip_lines


class ExcelReader(CalculationParsingMixin):
    """Reads a `_dd` worksheet into a list of `Question`, validating as it goes.

    The `calc:` block of the Responses column is parsed by
    `CalculationParsingMixin` (`calculation_parser.py`), inherited rather than
    delegated to so that `_validate_calculation_fields` stays reachable at
    `ExcelReader._validate_calculation_fields` -- which is where
    `tests/test_calculation_registry.py` reads its source to check that every
    `CalculationType` has a branch.
    """

    NUMBER_OF_COLUMNS = 14
    COLUMN_NAMES = [
        "FieldName",
        "QuestionType",
        "FieldType",
        "QuestionText",
        "MaxCharacters",
        "Responses",
        "LowerRange",
        "UpperRange",
        "LogicCheck",
        "DontKnow",
        "Refuse",
        "Optional",
        "Skip",
        "Comments",
    ]
    # Index of the Optional/NA column within COLUMN_NAMES -- a dictionary
    # written before this column was repurposed still has "NA" there, and is
    # still accepted: see the header check in create_question_list.
    OPTIONAL_COLUMN_INDEX = 11
    LEGACY_NA_COLUMN_NAME = "NA"

    NUMERIC_ONLY_RE = re.compile(r"^\d+$")
    DECIMAL_RE = re.compile(r"^\d+(\.\d+)?$")
    DATE_RANGE_RE = re.compile(r"^([+-])(\d+)([dwmy])$")
    HARDCODED_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
    FIELD_NAME_RE = re.compile(r"\b[a-z_][a-z0-9_]*\b", re.IGNORECASE)
    QUOTED_STRING_RE = re.compile(r"'[^']*'")
    PLACEHOLDER_RE = re.compile(r"\[\[(\w+)\]\]")
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

    # Spellings of "automatic". A question's behaviour is decided by its
    # fieldname (reserved variables), by whether it has a calculation, or by
    # the CRF's idconfig — never by which of these words was used. They are
    # normalized to "automatic" on the way in, so the generated XML always uses
    # the one spelling every app build understands.
    AUTOMATIC_TYPE_ALIASES = {"calc", "calculation", "calculated"}

    VALID_QUESTION_TYPES = {
        "radio",
        "combobox",
        "checkbox",
        "text",
        "date",
        "information",
        "automatic",
        "button",
    }
    VALID_FIELD_TYPES = {
        "text",
        "datetime",
        "date",
        "integer",
        "text_integer",
        "text_decimal",
        "n/a",
        "hourmin",
    }
    # What a typed answer may be. `integer` is deliberately absent: it is the
    # FieldType a radio question stores, and allowing it here would leave the
    # app unable to tell a whole number from a decimal.
    TYPED_FIELD_TYPES = {"text", "text_integer", "text_decimal", "hourmin"}
    # Field types whose answer is a number, and so can carry a range.
    NUMERIC_FIELD_TYPES = {"text_integer", "text_decimal"}
    # Question types answered by picking from a list rather than by typing.
    # Their answer is a code, so a numeric range means nothing for them.
    SELECTION_QUESTION_TYPES = {"radio", "checkbox", "combobox"}
    # A time is always hh:mm.
    HOURMIN_MAX_CHARACTERS = "=5"
    # Reserved system variables. The generator writes these itself, in the
    # positions that make them correct, so they never need a calculation and a
    # declared row is dropped rather than passed through. See models.py.
    BUILT_IN_AUTO_FIELDS = RESERVED_SYSTEM_FIELDS
    # Word operators are excluded from field-reference validation.
    LOGIC_KEYWORDS = {"and", "or", "not", "contains", "does", "contain"}

    def __init__(self, supplied_auto_fields: set[str] | None = None) -> None:
        self.logstring: list[str] = []
        self.errorsEncountered = False
        self.worksheetErrorsEncountered = False
        self.questionList: list[Question] = []
        # Reserved variables whose calculation was dropped, so the
        # reserved-variable check can report it.
        self.reservedFieldsWithCalculation: set[str] = set()
        # Automatic fields the manifest fills in (linking field, increment
        # field, primary key, idconfig parts). They legitimately have no
        # calculation, so they are exempt from the calculation check.
        self.suppliedAutoFields = {f.lower() for f in (supplied_auto_fields or set())}
        # Set from the header row: True once column 12 reads "NA" (an
        # old-style dictionary, written before the column was repurposed).
        # Its contents are then ignored entirely rather than parsed as the
        # Optional flag -- an old sheet's NA cells were never meant to make
        # anything optional, so a stray value there must not start doing so.
        self.optionalColumnIsLegacyNa = False

    def create_question_list(self, worksheet: Worksheet) -> list[Question]:
        self.worksheetErrorsEncountered = False
        self.logstring.append(f"\rChecking worksheet: '{worksheet.title}'")
        self.questionList = []
        self.reservedFieldsWithCalculation = set()
        self.optionalColumnIsLegacyNa = False

        for row_idx in range(1, worksheet.max_row + 1):
            try:
                if row_idx == 1:
                    current_headers = [self._get_cell_trim(worksheet, 1, i + 1) for i in range(self.NUMBER_OF_COLUMNS)]
                    expected_legacy = list(self.COLUMN_NAMES)
                    expected_legacy[self.OPTIONAL_COLUMN_INDEX] = self.LEGACY_NA_COLUMN_NAME
                    if current_headers == expected_legacy:
                        self.optionalColumnIsLegacyNa = True
                    elif current_headers != self.COLUMN_NAMES:
                        self._error(
                            "ERROR - Header: The header names in worksheet "
                            f"'{worksheet.title}' are incorrect. Header names should be: "
                            "FieldName, QuestionType, FieldType, QuestionText, MaxCharacters, "
                            "Responses, LowerRange, UpperRange, LogicCheck, DontKnow, Refuse, Optional, Skip, Comments "
                            "(a dictionary written before the Optional column was added may still say 'NA' there instead)"
                        )
                    continue

                comments_merge = self._merged_range_at(
                    worksheet, row_idx, self.NUMBER_OF_COLUMNS
                )
                if comments_merge is not None:
                    # A merged Comments cell skips the whole row. That is right
                    # for a section-heading banner and wrong for a question,
                    # and the discriminator is whether the same merged range
                    # also covers column 1 (FieldName).
                    #
                    # A banner is one range spanning the full row, so column 1
                    # is inside it and there is no FieldName to lose. Every
                    # merge in both live dictionaries is exactly that -- 28 of
                    # them, all single-row and all full width -- which is why
                    # this has to stay a silent skip rather than becoming an
                    # error.
                    #
                    # A merge that covers Comments but *not* column 1 is the
                    # hazard: the row still has its FieldName and its question,
                    # and skipping it drops that question from the generated
                    # XML with no log line anywhere. The author ships a survey
                    # missing a question and is told nothing. The common way in
                    # is merging a Comments note down across two question rows.
                    if comments_merge.min_col == 1:
                        continue
                    self._error(
                        f"ERROR - Merged cell: Row {row_idx} in worksheet "
                        f"'{worksheet.title}' has its Comments cell merged across "
                        f"{comments_merge.coord}, which does not include the FieldName "
                        "column. The whole row is skipped, so its question would be "
                        "dropped from the generated XML. Unmerge it, or merge the full "
                        "row if it is meant to be a section heading."
                    )
                    continue

                q = Question()
                q.fieldName = self._get_cell_trim(worksheet, row_idx, 1)
                if not q.fieldName:
                    self._error(
                        f"ERROR - FieldName: Row {row_idx} in worksheet '{worksheet.title}' has a blank FieldName."
                    )
                    continue

                self._check_field_name(worksheet.title, q.fieldName)
                q.questionType = self._normalize_question_type(
                    self._get_cell_trim(worksheet, row_idx, 2))
                q.fieldType = self._get_cell_trim(worksheet, row_idx, 3)
                q.questionText = self._get_cell_trim(worksheet, row_idx, 4)

                if q.questionText == "" and q.questionType != "automatic":
                    self._error(
                        f"ERROR - QuestionText: FieldName '{q.fieldName}' in worksheet '{worksheet.title}' has blank QuestionText."
                    )

                max_chars = self._get_cell_trim(worksheet, row_idx, 5)
                q.maxCharacters = max_chars if max_chars else "-9"
                if q.maxCharacters != "-9":
                    self._check_max_chars_value(worksheet.title, q.maxCharacters, q.fieldName)

                raw_responses = self._get_cell_raw(worksheet, row_idx, 6)
                raw_stripped = raw_responses.strip()
                if raw_stripped.lower().startswith("source:"):
                    self._parse_dynamic_responses(raw_responses, q, worksheet.title, q.fieldName)
                elif raw_stripped.lower().startswith("calc:"):
                    if q.questionType == "automatic":
                        if q.fieldName.lower() not in self.BUILT_IN_AUTO_FIELDS:
                            self._parse_automatic_calculation(raw_responses, q, worksheet.title, q.fieldName)
                        else:
                            # The app has its own handler for these, so the
                            # calculation is dropped rather than written out.
                            # Record it so the reserved-variable check can say so.
                            self.reservedFieldsWithCalculation.add(
                                q.fieldName.lower())
                    else:
                        self._error(
                            f"ERROR - Calculation: FieldName '{q.fieldName}' in worksheet '{worksheet.title}' "
                            "has calculation syntax but QuestionType is not 'automatic'."
                        )
                elif raw_stripped.lower().startswith("mask:"):
                    if q.questionType == "text":
                        q.mask = raw_stripped[5:].strip()
                    else:
                        self._error(
                            f"ERROR - Mask: FieldName '{q.fieldName}' in worksheet '{worksheet.title}' "
                            "has mask syntax but QuestionType is not 'text'."
                        )
                else:
                    q.responses = raw_responses

                self._check_question_field_type(q, worksheet.title)

                lower_val = self._get_cell_trim(worksheet, row_idx, 7)
                upper_val = self._get_cell_trim(worksheet, row_idx, 8)
                q.lowerRange = lower_val if lower_val else "-9"
                q.upperRange = upper_val if upper_val else "-9"

                if q.questionType == "date":
                    self._check_date_range(worksheet.title, q.lowerRange, q.fieldName, "LowerRange")
                    self._check_date_range(worksheet.title, q.upperRange, q.fieldName, "UpperRange")
                else:
                    if q.lowerRange != "-9":
                        self._check_numeric_range(worksheet.title, q.lowerRange, q.fieldName, "LowerRange")
                    if q.upperRange != "-9":
                        self._check_numeric_range(worksheet.title, q.upperRange, q.fieldName, "UpperRange")

                logic_raw = self._get_cell_trim(worksheet, row_idx, 9)
                if logic_raw:
                    for check in self._split_lines(logic_raw):
                        trimmed = check.strip()
                        if trimmed.startswith("unique;"):
                            parts = trimmed.split(";", 1)
                            if len(parts) == 2:
                                message = parts[1].strip()
                                if message.startswith("'") and message.endswith("'"):
                                    q.uniqueCheckMessage = message.strip("'")
                                else:
                                    self._error(
                                        f"ERROR - LogicCheck: FieldName '{q.fieldName}' in worksheet '{worksheet.title}' "
                                        "has invalid syntax for unique check message (must be in single quotes): "
                                        f"{trimmed}"
                                    )
                            else:
                                self._error(
                                    f"ERROR - LogicCheck: FieldName '{q.fieldName}' in worksheet '{worksheet.title}' "
                                    f"has invalid syntax for unique check (missing message): {trimmed}"
                                )
                        else:
                            q.logicChecks.append(trimmed)
                            self._check_logic_check_syntax(worksheet.title, trimmed, q.fieldName)

                q.dontKnow = self._get_cell_trim(worksheet, row_idx, 10) or "-9"
                if q.dontKnow != "-9":
                    self._check_special_button(worksheet.title, q.dontKnow, q.fieldName, "DontKnow")

                q.refuse = self._get_cell_trim(worksheet, row_idx, 11) or "-9"
                if q.refuse != "-9":
                    self._check_special_button(worksheet.title, q.refuse, q.fieldName, "Refuse")

                if self.optionalColumnIsLegacyNa:
                    # Legacy dictionary: this column is still labelled "NA".
                    # Its contents are never parsed as Optional -- see the
                    # comment on optionalColumnIsLegacyNa in __init__.
                    q.optional = "-9"
                else:
                    q.optional = self._get_cell_trim(worksheet, row_idx, 12) or "-9"
                    if q.optional != "-9":
                        self._check_special_button(worksheet.title, q.optional, q.fieldName, "Optional")
                        if q.questionType != "text":
                            self._error(
                                f"ERROR - Optional: FieldName '{q.fieldName}' in worksheet '{worksheet.title}' "
                                "sets Optional, but Optional is only meaningful for a 'text' QuestionType "
                                f"(this question is '{q.questionType}')."
                            )

                q.skip = self._get_cell_trim(worksheet, row_idx, 13)
                if q.skip:
                    self._check_skip_syntax(worksheet.title, q.skip, q.fieldName)

                self.questionList.append(q)
            except Exception as ex:
                self._error(
                    f"ERROR - Row: An unexpected error occurred while processing row {row_idx} "
                    f"in worksheet '{worksheet.title}'. The error was: {ex}"
                )

        # Run unconditionally. These used to be skipped entirely whenever any
        # single row had errored, which bought nothing: a row that failed
        # validation is still in `questionList` (only a blank FieldName is
        # dropped), so the cross-row checks see the same data either way. What
        # the gate did cost was a second wave -- an author fixing row errors
        # got a clean-looking run, then hit a fresh wall of structural errors
        # on the run where the last row error disappeared.
        self._check_logic_field_names(worksheet.title)
        self._check_skip_to_field_names(worksheet.title)
        self._check_reserved_variable_reads(worksheet.title)
        self._check_message_placeholders(worksheet.title)
        self._check_required_max_characters(worksheet.title)
        self._check_ranges(worksheet.title)
        self._check_duplicate_columns(worksheet.title)
        self._check_automatic_has_calculation(worksheet.title)
        self._check_responses_are_answerable(worksheet.title)
        self._check_preskip_does_not_test_itself(worksheet.title)
        self._check_max_characters_is_meaningful(worksheet.title)
        self._check_reserved_automatic_fields(worksheet.title)
        self._check_comments_field_is_optional(worksheet.title)
        if not self.worksheetErrorsEncountered:
            self.logstring.append(f"No errors found in '{worksheet.title}'")

        return self.questionList

    # Thin wrappers over cell_text, kept as methods because they are called
    # ~200 times in this file and the shorter name reads better at the call
    # site. The implementations are shared with crf_reader and xml_generator
    # so the same cell cannot be read two different ways.
    _split_lines = staticmethod(split_cell_lines)
    _to_str = staticmethod(to_str)

    def _get_cell_trim(self, ws: Worksheet, row: int, col: int) -> str:
        return cell_trim(ws, row, col)

    def _get_cell_raw(self, ws: Worksheet, row: int, col: int) -> str:
        return cell_raw(ws, row, col)

    @staticmethod
    def _merged_range_at(ws: Worksheet, row: int, col: int):
        """The merged range covering this cell, or `None` if it is not merged.

        Returns the range rather than a bool because the caller has to know
        *how far* the merge reaches to tell a section heading from a swallowed
        question -- see the call site.
        """
        coord = ws.cell(row=row, column=col).coordinate
        for merged_range in ws.merged_cells.ranges:
            if coord in merged_range:
                return merged_range
        return None

    def _error(self, message: str) -> None:
        self.errorsEncountered = True
        self.worksheetErrorsEncountered = True
        self.logstring.append(message)

    def _check_field_name(self, worksheet: str, fieldname: str) -> None:
        if not fieldname:
            self._error(f"ERROR - FieldName: A row in worksheet '{worksheet}' has an empty FieldName")
            return
        if fieldname[0].isdigit():
            self._error(f"ERROR - FieldName: FieldName '{fieldname}' in worksheet '{worksheet}' starts with a number")
        elif " " in fieldname:
            self._error(f"ERROR - FieldName: FieldName '{fieldname}' in worksheet '{worksheet}' contains a space")
        elif not re.fullmatch(r"[A-Za-z0-9_]+", fieldname):
            # Deliberately an ASCII character class rather than `str.isalnum`,
            # which is Unicode-aware: `prenom` with an accent, and every other
            # accented spelling the French dictionaries naturally reach for,
            # passed every check here and then became an XML attribute and a
            # SQLite column name.
            self._error(
                f"ERROR - FieldName: FieldName '{fieldname}' in worksheet '{worksheet}' is invalid. "
                "Only letters, digits, and underscores are allowed."
            )
        elif fieldname != fieldname.lower():
            self._error(f"ERROR - FieldName: FieldName '{fieldname}' in worksheet '{worksheet}' is not all lowercase")
        elif fieldname[0] == "_":
            self._error(f"ERROR - FieldName: FieldName '{fieldname}' in worksheet '{worksheet}' starts with an underscore")
        elif fieldname.lower() == "end":
            # Reserved as the Skip target sentinel meaning "end of form" (see
            # `_check_skip_to_field_names`). Not part of RESERVED_SYSTEM_FIELDS
            # -- that set is used elsewhere to silently drop a declared row
            # for the name, which is wrong here: a row named 'end' should be
            # rejected outright, not dropped, since the app would then be
            # unable to skip to it by name (it never gets that far -- the
            # sentinel is checked before any fieldname lookup).
            self._error(
                f"ERROR - FieldName: FieldName '{fieldname}' in worksheet '{worksheet}' uses 'end', "
                "which is reserved as the Skip target meaning 'end of form'. "
                "Choose a different FieldName."
            )

    def _check_max_chars_value(self, worksheet: str, max_chars: str, fieldname: str) -> None:
        numeric = max_chars[1:] if max_chars.startswith("=") else max_chars
        if not self.NUMERIC_ONLY_RE.fullmatch(numeric):
            self._error(
                f"ERROR - MaxCharacters: FieldName '{fieldname}' in worksheet '{worksheet}' "
                f"has a non-numeric value for MaxCharacters: {max_chars}"
            )
            return
        num = int(numeric)
        if num < 1 or num > 2000:
            self._error(
                f"ERROR - MaxCharacters: FieldName '{fieldname}' in worksheet '{worksheet}' "
                f"has a MaxCharacters value that is out of range (1 to 2000): {max_chars}"
            )

    def _check_question_field_type(self, q: Question, worksheet: str) -> None:
        questiontype = q.questionType
        fieldtype = q.fieldType
        fieldname = q.fieldName

        if questiontype not in self.VALID_QUESTION_TYPES:
            self._error(
                f"ERROR - QuestionType: The QuestionType {questiontype} for FieldName '{fieldname}' "
                f"in worksheet '{worksheet}' is not among the predefined list."
            )

        if fieldtype not in self.VALID_FIELD_TYPES:
            self._error(
                f"ERROR - FieldType: The FieldType '{fieldtype}' for FieldName '{fieldname}' in worksheet '{worksheet}' "
                "is not among the predefined list."
            )

        if questiontype == "text" and fieldtype not in self.TYPED_FIELD_TYPES:
            self._error(
                f"ERROR - FieldType: The FieldType '{fieldtype}' for FieldName '{fieldname}' "
                f"in worksheet '{worksheet}' cannot be used with the QuestionType 'text'. Use one of: "
                f"{', '.join(sorted(self.TYPED_FIELD_TYPES))}."
            )

        if questiontype == "radio" and fieldtype != "integer":
            self._error(
                f"ERROR - FieldType: The FieldType for FieldName '{fieldname}' in worksheet '{worksheet}' must be integer "
                "when the QuestionType is 'radio'."
            )

        if questiontype == "checkbox" and fieldtype != "text":
            self._error(
                f"ERROR - FieldType: The FieldType for FieldName '{fieldname}' in worksheet '{worksheet}' must be text "
                "when the QuestionType is 'checkbox'."
            )

        if questiontype == "date" and fieldtype not in {"date", "datetime"}:
            self._error(
                f"ERROR - FieldType: The FieldType for FieldName '{fieldname}' in worksheet '{worksheet}' "
                "must be date when the QuestionType is 'date' or 'datetime'."
            )

        # combobox belongs here as much as radio and checkbox do: the XML
        # generator emits all three through the same `response.find(":")`
        # split, so a combobox option written as `Yes` rather than `1:Yes`
        # silently stored the value "Ye" -- valid XML, wrong code, on every
        # answer to that question.
        if (
            questiontype in {"radio", "checkbox", "combobox"}
            and q.responseSourceType == ResponseSourceType.STATIC
            and q.responses
        ):
            responses = self._split_lines(q.responses)
            seen: list[str] = []
            for response in responses:
                index = response.find(":")
                if index == -1:
                    self._error(
                        f"ERROR - Responses: Invalid static {questiontype} options for '{fieldname}' in worksheet '{worksheet}'. "
                        f"Expected format 'number:Statement', found '{response}'."
                    )
                    return
                if len(response.split(":")) != 2:
                    self._error(
                        f"ERROR - Responses: Invalid static {questiontype} options for '{fieldname}' in worksheet '{worksheet}'. "
                        f"Expected format 'number:Statement', found '{response}'."
                    )
                    return
                key = response[:index]
                seen.append(key)
                duplicates = sorted({k for k in seen if seen.count(k) > 1})
                if len(set(seen)) != len(seen):
                    self._error(
                        f"ERROR - Responses: The Responses for FieldName '{fieldname}' in worksheet '{worksheet}' "
                        f"has duplicates {','.join(duplicates)}"
                    )
                    return
                if response.startswith(" "):
                    self._error(
                        f"ERROR - Responses: Invalid static {questiontype} options for '{fieldname}' in worksheet '{worksheet}'. "
                        "Please remove leading spaces."
                    )
                    return
                if ": " in response:
                    self._error(
                        f"ERROR - Responses: Invalid static {questiontype} options for '{fieldname}' in worksheet '{worksheet}'. "
                        "Please remove space after the colon (:) for static responses."
                    )
                    return

    def _check_numeric_range(self, worksheet: str, value: str, fieldname: str, range_name: str) -> None:
        if not self.DECIMAL_RE.fullmatch(value):
            self._error(
                f"ERROR - {range_name}: FieldName '{fieldname}' in worksheet '{worksheet}' "
                f"has a non-numeric value for {range_name}: {value}"
            )

    def _check_date_range(self, worksheet: str, value: str, fieldname: str, range_name: str) -> None:
        if value == "-9":
            self._error(
                f"ERROR - {range_name}: FieldName '{fieldname}' in worksheet '{worksheet}' has a missing value for {range_name}"
            )
            return
        if value in {"0", "+0d", "-0d"}:
            return
        if self.DATE_RANGE_RE.fullmatch(value):
            return
        if self.HARDCODED_DATE_RE.fullmatch(value):
            try:
                from datetime import datetime

                datetime.strptime(value, "%Y-%m-%d")
                return
            except ValueError:
                self._error(
                    f"ERROR - {range_name}: FieldName '{fieldname}' in worksheet '{worksheet}' has an invalid date value: {value}"
                )
                return
        self._error(
            f"ERROR - {range_name}: FieldName '{fieldname}' in worksheet '{worksheet}' has an invalid format for {range_name}: {value}"
        )

    def _check_logic_check_syntax(self, worksheet: str, logic_check: str, fieldname: str) -> None:
        if ";" not in logic_check:
            self._error(
                f"ERROR - LogicCheck: FieldName '{fieldname}' in worksheet '{worksheet}' has invalid syntax "
                f"for LogicCheck (missing semicolon): {logic_check}"
            )
            return
        parts = logic_check.split(";", 1)
        if len(parts) != 2:
            self._error(
                f"ERROR - LogicCheck: FieldName '{fieldname}' in worksheet '{worksheet}' has invalid syntax for LogicCheck: {logic_check}"
            )
            return
        expression = parts[0].strip()
        message = parts[1].strip()
        if not (message.startswith("'") and message.endswith("'")):
            self._error(
                f"ERROR - LogicCheck: FieldName '{fieldname}' in worksheet '{worksheet}' has invalid syntax for LogicCheck "
                f"(message must be in single quotes): {logic_check}"
            )
            return
        operators = ["=", "!=", "<>", ">", ">=", "<", "<=", "and", "or", "contains", "does not contain"]
        if not any(op in expression for op in operators):
            self._error(
                f"ERROR - LogicCheck: FieldName '{fieldname}' in worksheet '{worksheet}' has invalid syntax for LogicCheck "
                f"(no operator found): {logic_check}"
            )

    def _check_special_button(self, worksheet: str, value: str, fieldname: str, button_name: str) -> None:
        # Matches what the generator itself accepts as truthy
        # ({"TRUE", "True"} -- see xml_generator.py); this used to require
        # exactly "True"/"False" and reject "TRUE", which the generator
        # would otherwise have treated as set.
        if value not in {"True", "TRUE", "False", "FALSE"}:
            self._error(
                f"ERROR - {button_name}: FieldName '{fieldname}' in worksheet '{worksheet}' "
                f"has an invalid value for '{button_name}': {value}"
            )

    def _check_skip_syntax(self, worksheet: str, skip_text: str, fieldname: str) -> None:
        """Reject anything `skip_parser` cannot read.

        This used to hand-roll its own parse -- token counts, space offsets and
        a `len_skip = 13 if postskip else 12` slice -- while the generator
        hand-rolled a different one. Going through the shared parser is what
        makes it impossible for a string this accepts to be dropped later (see
        `skip_parser`'s module docstring for the `Preskip:` case that motivated
        it).
        """
        for skip in split_skip_lines(skip_text):
            if parse_skip(skip) is None:
                self._error(
                    f"ERROR - Skip: FieldName '{fieldname}' in worksheet '{worksheet}' "
                    f"has invalid syntax for Skip: {skip}. Expected "
                    "'preskip: if <field> <operator> <value>, skip to <target>' "
                    "(operators: =, >, >=, <, <=, <>, 'contains', 'does not contain'; "
                    "the value must be a single word with no spaces)."
                )

    def _check_logic_field_names(self, worksheet: str) -> None:
        field_index = {q.fieldName: i for i, q in enumerate(self.questionList)}
        for question in self.questionList:
            for logic_check in question.logicChecks:
                cur_field = question.fieldName
                expression = logic_check.split(";", 1)[0].strip()
                clean_expression = self.QUOTED_STRING_RE.sub("", expression)
                matches = self.FIELD_NAME_RE.findall(clean_expression)
                referenced = {m for m in matches if m.lower() not in self.LOGIC_KEYWORDS}
                for ref in referenced:
                    if ref.lower() in TRAILING_SYSTEM_FIELD_NAMES:
                        self._error(
                            f"ERROR - LogicCheck: In worksheet '{worksheet}', the LogicCheck for FieldName '{cur_field}' "
                            f"uses the reserved variable '{ref}', which is still empty while the questionnaire is "
                            "being answered. The check would never fire, so it would look like a validation that "
                            "always passes. Use 'startdate' if the interview date is what was meant."
                        )
                    elif ref in field_index:
                        if field_index[ref] > field_index[cur_field]:
                            self._error(
                                f"ERROR - LogicCheck: In worksheet '{worksheet}', the LogicCheck for FieldName '{cur_field}' "
                                f"uses a FieldName AFTER the current question: {ref}"
                            )
                    else:
                        self._error(
                            f"ERROR - LogicCheck: In worksheet '{worksheet}', the LogicCheck for FieldName '{cur_field}' "
                            f"uses a nonexistent FieldName: {ref}"
                        )

    @classmethod
    def _part_field_refs(cls, part: CalculationPart | None) -> set[str]:
        """Field names one calculation part reads, following nested parts."""
        if part is None:
            return set()
        refs: set[str] = set()
        if part.lookupField:
            refs.add(part.lookupField)
        refs |= set(cls.PLACEHOLDER_RE.findall(part.querySql))
        refs |= set(cls.PLACEHOLDER_RE.findall(part.constantValue))
        for nested in part.parts:
            refs |= cls._part_field_refs(nested)
        return refs

    @classmethod
    def _question_field_refs(cls, question: Question) -> set[tuple[str, str]]:
        """Every field a question reads, paired with where it reads it.

        Three places a field name can appear: a calculation's structured slots,
        a response filter narrowing a list against an earlier answer, and the
        question text a respondent is shown. The raw Responses text is consumed
        by whichever block it held, so the parsed slots are read instead of it.
        """
        refs: set[tuple[str, str]] = set()

        def add(where: str, names) -> None:
            refs.update((name.strip(), where) for name in names if name and name.strip())

        add("its calculation", [question.calculationLookupField])
        add("its calculation", [p.fieldName for p in question.calculationQueryParameters])
        for condition in question.calculationCaseConditions:
            add("its calculation", [condition.field])
            add("its calculation", cls._part_field_refs(condition.result))
        add("its calculation", cls._part_field_refs(question.calculationCaseElse))
        for part in question.calculationMathParts + question.calculationConcatParts:
            add("its calculation", cls._part_field_refs(part))
        # `separator` is where an age-at-date calculation puts its reference date.
        for text in (question.calculationConcatSeparator,
                     question.calculationQuerySql,
                     question.calculationConstantValue):
            add("its calculation", cls.PLACEHOLDER_RE.findall(text))
        for response_filter in question.responseFilters:
            add("a response filter", cls.PLACEHOLDER_RE.findall(response_filter.value))
        add("its question text", cls.PLACEHOLDER_RE.findall(question.questionText))
        return refs

    def _check_comments_field_is_optional(self, worksheet: str) -> None:
        """A field named 'comments' used to be hardcoded as always-optional in
        the app. That hardcode is gone -- Optional is now what makes a text
        field skippable, and 'comments' is not special anymore. This is a
        warning, not an auto-tick: silently setting it back would reintroduce
        the exact magic this column exists to remove.
        """
        for question in self.questionList:
            if question.fieldName.lower() != "comments":
                continue
            if question.questionType != "text" or question.optional in {"", "-9"}:
                self.logstring.append(
                    f"WARNING - Optional: In worksheet '{worksheet}', FieldName 'comments' does not "
                    "have Optional set to TRUE. A field named 'comments' used to be always-optional "
                    "automatically; it no longer is, so this interview cannot be finished without "
                    "answering it unless Optional is set."
                )

    def _check_message_placeholders(self, worksheet: str) -> None:
        """A validation message is shown exactly as it was written.

        Question text goes through `SurveyLoader.expandPlaceholders`, but the
        message on a logic or unique check does not -- the app renders the
        string straight into the error banner. So a placeholder there is not an
        empty value, it is the literal brackets, on screen, in front of the
        interviewer. That is true of every field name, not just a reserved one,
        which is why this warns about all of them.
        """
        for question in self.questionList:
            messages = [check.partition(";")[2] for check in question.logicChecks]
            messages.append(question.uniqueCheckMessage)
            for message in messages:
                for name in sorted(set(self.PLACEHOLDER_RE.findall(message))):
                    self.logstring.append(
                        f"WARNING - Message: In worksheet '{worksheet}', the validation message "
                        f"for FieldName '{question.fieldName}' contains '[[{name}]]'. Messages "
                        "are shown exactly as written, so the interviewer would see the brackets "
                        "rather than a value. Word the message without it."
                    )

    def _check_reserved_variable_reads(self, worksheet: str) -> None:
        """Nothing in a question may read a trailing variable.

        `starttime` and `startdate` are deliberately allowed: they hold a value
        from the first question onward, and an age-at-date calculation reading
        `[[startdate]]` is the intended use.
        """
        for question in self.questionList:
            if question.fieldName.lower() in RESERVED_SYSTEM_FIELDS:
                continue  # the row is dropped and already warned about
            for ref, where in sorted(self._question_field_refs(question)):
                if ref.lower() in TRAILING_SYSTEM_FIELD_NAMES:
                    self._error(
                        f"ERROR - Reserved variable: In worksheet '{worksheet}', FieldName "
                        f"'{question.fieldName}' reads '{ref}' in {where}. That variable is "
                        "written after the last question, so it is still empty while the "
                        "questionnaire is being answered and would be read as nothing at all. "
                        "Use 'startdate' if the interview date is what was meant."
                    )

    def _check_skip_to_field_names(self, worksheet: str) -> None:
        field_index = {q.fieldName: i for i, q in enumerate(self.questionList)}
        for question in self.questionList:
            if not question.skip:
                continue
            cur_field = question.fieldName
            for skip in split_skip_lines(question.skip):
                parsed = parse_skip(skip)
                if parsed is None:
                    # `_check_skip_syntax` already reported it, and the
                    # worksheet-level gate means this check normally does not
                    # run at all once that fired.
                    continue
                fieldname_to_check = parsed.field
                fieldname_to_skip_to = parsed.target
                cur_index = field_index[cur_field]

                if fieldname_to_check.lower() in RESERVED_SYSTEM_FIELDS:
                    self._error(
                        f"ERROR - Skip: In worksheet '{worksheet}', the skip for FieldName '{cur_field}' "
                        f"tests the reserved variable '{fieldname_to_check}'. Which questions get asked "
                        "must not depend on a value the generator supplies: the trailing variables are "
                        "still empty while the questionnaire is being answered, so the skip would never "
                        "fire, and branching on 'starttime' or 'startdate' would make one package ask "
                        "different questions on different days. Regenerate the questionnaire instead."
                    )
                elif fieldname_to_check in field_index:
                    if field_index[fieldname_to_check] > cur_index:
                        self._error(
                            f"ERROR - Skip: In worksheet '{worksheet}', the skip for FieldName '{cur_field}' "
                            f"checks skip for a FieldName AFTER the current question: {fieldname_to_check}"
                        )
                else:
                    self._error(
                        f"ERROR - Skip: In worksheet '{worksheet}', the skip for FieldName '{cur_field}' "
                        f"checks skip of a nonexistent FieldName: {fieldname_to_check}"
                    )

                if fieldname_to_skip_to.lower() == "end":
                    # The app's own end-of-form sentinel (see
                    # SurveyNavigationService.endOfFormSkipTarget in the app
                    # repo) -- not a fieldname, so none of the checks below
                    # apply to it.
                    pass
                elif fieldname_to_skip_to.lower() in RESERVED_SYSTEM_FIELDS:
                    self._error(
                        f"ERROR - Skip: In worksheet '{worksheet}', the skip for FieldName '{cur_field}' "
                        f"skips to the reserved variable '{fieldname_to_skip_to}'. The generator places "
                        "reserved variables itself ('starttime' and 'startdate' before the first question, "
                        "the rest after the last), so a skip to one cannot land where the dictionary "
                        "intends. Skip to a real question instead."
                    )
                elif fieldname_to_skip_to in field_index:
                    target_index = field_index[fieldname_to_skip_to]
                    if target_index < cur_index:
                        self._error(
                            f"ERROR - Skip: In worksheet '{worksheet}', the skip for FieldName '{cur_field}' "
                            f"skips to a FieldName BEFORE the current question: {fieldname_to_skip_to}"
                        )
                    elif target_index == cur_index:
                        self._error(
                            f"ERROR - Skip: In worksheet '{worksheet}', the skip for FieldName '{cur_field}' "
                            f"skips to the current question: {fieldname_to_skip_to}"
                        )
                else:
                    self._error(
                        f"ERROR - Skip: In worksheet '{worksheet}', the skip for FieldName '{cur_field}' "
                        f"skips to a nonexistent FieldName: {fieldname_to_skip_to}"
                    )

    def _check_required_max_characters(self, worksheet: str) -> None:
        """Every typed answer needs a length limit.

        Keyed on the QuestionType rather than a list of FieldTypes: the list
        used to name `text_integer` but not `integer`, so moving a field to the
        newer spelling silently dropped its length limit.
        """
        for question in self.questionList:
            if question.questionType != "text":
                continue

            if question.fieldType == "hourmin":
                # A time is hh:mm, so the length is not the author's to choose.
                if question.maxCharacters != self.HOURMIN_MAX_CHARACTERS:
                    self._error(
                        f"ERROR - MaxCharacters: In worksheet '{worksheet}', MaxCharacters for FieldName "
                        f"'{question.fieldName}' must be '{self.HOURMIN_MAX_CHARACTERS}' when the FieldType "
                        "is 'hourmin' — a time is always hh:mm."
                    )
                continue

            if question.maxCharacters == "-9":
                self._error(
                    f"ERROR - MaxCharacters: In worksheet '{worksheet}', MaxCharacters for FieldName '{question.fieldName}' needs a value"
                )

    def _check_ranges(self, worksheet: str) -> None:
        """LowerRange and UpperRange come as a pair, and only on numbers.

        A lower bound with no upper bound is the dangerous case: the generator
        writes the blank as `maxvalue='-9'`, so every value above -9 fails and
        the question cannot be answered at all.
        """
        for question in self.questionList:
            # A range on a selection question is always a mistake, and a
            # silent one: the generator writes <numeric_check> for any
            # question whose LowerRange is set (xml_generator only excludes
            # 'date'), so a LowerRange left behind on a copy-pasted checkbox
            # row makes every option fail validation and the question
            # unanswerable in the field.
            if question.questionType in self.SELECTION_QUESTION_TYPES:
                if question.lowerRange != "-9" or question.upperRange != "-9":
                    self._error(
                        f"ERROR - Range: In worksheet '{worksheet}', FieldName "
                        f"'{question.fieldName}' is a '{question.questionType}' question with a "
                        "LowerRange or UpperRange. A range applies to a typed number, not to a "
                        "list of options, and leaves the question unanswerable. Clear both."
                    )
                continue

            if question.questionType not in {"text", "date"}:
                continue

            has_lower = question.lowerRange != "-9"
            has_upper = question.upperRange != "-9"

            if question.fieldType == "hourmin":
                if has_lower or has_upper:
                    self._error(
                        f"ERROR - Range: In worksheet '{worksheet}', FieldName '{question.fieldName}' has a "
                        "LowerRange or UpperRange, which cannot be applied to a 'hourmin' field. Leave both blank."
                    )
                continue

            if has_lower != has_upper:
                missing = "UpperRange" if has_lower else "LowerRange"
                self._error(
                    f"ERROR - Range: In worksheet '{worksheet}', FieldName '{question.fieldName}' sets only one "
                    f"end of its range. Set {missing} as well, or clear both — a half-set range rejects every answer."
                )
                continue

            # A fixed-length number is an identifier (a household ID, a phone
            # number), where a range means nothing. A variable-length one is a
            # quantity, and a quantity without a range accepts any value that
            # fits.
            if (
                question.fieldType in self.NUMERIC_FIELD_TYPES
                and not has_lower
                and not question.maxCharacters.startswith("=")
            ):
                self.logstring.append(
                    f"WARNING - Range: In worksheet '{worksheet}', FieldName '{question.fieldName}' is numeric "
                    "with no LowerRange or UpperRange, so any value that fits MaxCharacters is accepted."
                )

    @classmethod
    def _normalize_question_type(cls, question_type: str) -> str:
        """Collapse the spellings of "automatic" to the canonical one.

        A dictionary may say `calc`, `calculation` or `calculated`; they all
        mean the same thing. Normalizing here rather than in the XML keeps
        every downstream check on one word, and keeps the generated file
        readable by app builds that only recognise `automatic`.
        """
        if question_type.strip().lower() in cls.AUTOMATIC_TYPE_ALIASES:
            return "automatic"
        return question_type

    def _check_reserved_automatic_fields(self, worksheet: str) -> None:
        """Point out reserved variables, without ever failing the build.

        Declaring one is perfectly valid — many dictionaries do, so the
        variable is visible to analysts reading the spreadsheet. The row itself
        is dropped and the generator writes the variable in the position that
        makes it correct, so this only ever warns and every dictionary written
        before these names were reserved keeps working.

        Giving one a calculation is worth calling out separately: the
        calculation is dropped rather than written to the XML, so an author who
        writes one gets no effect at all.
        """
        for question in self.questionList:
            name = question.fieldName.lower()
            if name not in self.BUILT_IN_AUTO_FIELDS:
                continue

            if name in self.reservedFieldsWithCalculation:
                self.logstring.append(
                    f"WARNING - Reserved variable: In worksheet '{worksheet}', "
                    f"'{question.fieldName}' is reserved and its calculation is "
                    "ignored. Use a different FieldName for your own value."
                )
                continue

            self.logstring.append(
                f"WARNING - Reserved variable: In worksheet '{worksheet}', "
                f"'{question.fieldName}' is reserved. The generator writes it "
                "itself, in the correct position; this row is ignored."
            )

    def _check_automatic_has_calculation(self, worksheet: str) -> None:
        """An automatic field with no calculation is never populated.

        It stays null for every record, and — worse — any skip that tests it
        silently fails open, because a skip whose field is unanswered never
        fires. The question it was meant to guard is then asked of everyone.

        There is no FieldType-based exemption: a mid-questionnaire timestamp
        (previously a blank-Responses `datetime` field left to an implicit
        runtime fallback) must now use `calc:timestamp` explicitly. See
        "Custom Timestamp Fields" in README.md.
        """
        for question in self.questionList:
            if question.questionType != "automatic":
                continue
            name = question.fieldName.lower()
            if (
                name in self.BUILT_IN_AUTO_FIELDS
                or name in self.suppliedAutoFields
                or name in KNOWN_AUTOMATIC_FIELDS
            ):
                continue
            if question.calculationType != CalculationType.NONE:
                continue
            detail = (
                "the Responses column is blank"
                if not question.responses.strip()
                else "the Responses column does not start with 'calc:'"
            )
            self._error(
                f"ERROR - Calculation: In worksheet '{worksheet}', automatic FieldName "
                f"'{question.fieldName}' has no calculation ({detail}), so it is never "
                "given a value. Add a 'calc:' block, or remove the field."
            )

    def _check_responses_are_answerable(self, worksheet: str) -> None:
        """A selection question with no options cannot be answered at all."""
        for question in self.questionList:
            if question.questionType not in {"radio", "checkbox", "combobox"}:
                continue
            if question.responseSourceType != ResponseSourceType.STATIC:
                continue
            if not question.responses.strip():
                self._error(
                    f"ERROR - Responses: In worksheet '{worksheet}', FieldName "
                    f"'{question.fieldName}' is a {question.questionType} question with no "
                    "responses, so it cannot be answered. Add response options, or a "
                    "'source:csv'/'source:database' block."
                )

    def _check_preskip_does_not_test_itself(self, worksheet: str) -> None:
        """A preskip that tests its own field can never behave as intended.

        Preskips run before the question is shown, so on a new record the field
        is still unanswered and the skip never fires. On an existing record the
        stored value does fire it, and the jump clears every answer it passes
        over — including this one. Such a rule is nearly always meant to be a
        postskip.
        """
        for question in self.questionList:
            if not question.skip:
                continue
            for skip in split_skip_lines(question.skip):
                parsed = parse_skip(skip)
                if parsed is None or parsed.kind != "preskip":
                    continue
                if parsed.field == question.fieldName:
                    self._error(
                        f"ERROR - Skip: In worksheet '{worksheet}', the preskip for FieldName "
                        f"'{question.fieldName}' tests its own field. It cannot fire on a new "
                        "record, and on an existing record it erases the answer. Use a "
                        f"postskip instead: {skip.strip()}"
                    )

    def _check_max_characters_is_meaningful(self, worksheet: str) -> None:
        """MaxCharacters only affects typed input; on a selection it is ignored."""
        for question in self.questionList:
            if question.maxCharacters == "-9":
                continue
            if question.questionType in {"radio", "checkbox", "combobox"}:
                self.logstring.append(
                    f"WARNING - MaxCharacters: In worksheet '{worksheet}', FieldName "
                    f"'{question.fieldName}' is a {question.questionType} question, so "
                    "MaxCharacters is ignored. Remove it, or change the QuestionType to 'text'."
                )

    def _check_duplicate_columns(self, worksheet: str) -> None:
        fields = [q.fieldName for q in self.questionList if q.questionType != "information"]
        duplicates = sorted({f for f in fields if fields.count(f) > 1})
        if len(set(fields)) != len(fields):
            self._error(
                f"ERROR - FieldName: Duplicate FieldNames in worksheet '{worksheet}': "
                f"{','.join(duplicates)}. "
                "Check for empty rows at the end of the spreadsheet and delete them."
            )

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
