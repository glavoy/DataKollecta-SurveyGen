from __future__ import annotations

import re

from openpyxl.worksheet.worksheet import Worksheet

from models import (
    RESERVED_SYSTEM_FIELDS,
    Question,
    ResponseSourceType,
)
from calculation_parser import CalculationParsingMixin
from cell_text import cell_raw, cell_trim, split_cell_lines
from dd_validators import WorksheetValidationMixin
from response_parser import ResponseParsingMixin
from skip_parser import parse_skip, split_skip_lines


class ExcelReader(
    CalculationParsingMixin, ResponseParsingMixin, WorksheetValidationMixin
):
    """Reads a `_dd` worksheet into a list of `Question`, validating as it goes.

    What is left here is the shape of the job: the column layout, the row loop
    that reads each cell, the checks that judge a cell on its own, and the error
    sink everything reports through. Three mixins carry the rest --

      * `CalculationParsingMixin` (`calculation_parser.py`) -- a `calc:` block
      * `ResponseParsingMixin` (`response_parser.py`) -- a `source:` block
      * `WorksheetValidationMixin` (`dd_validators.py`) -- the checks that need
        the finished `questionList` rather than one cell

    -- and they are inherited rather than delegated to so every method stays
    reachable at the attribute path it had before the split. That matters for
    `_validate_calculation_fields` in particular, since
    `tests/test_calculation_registry.py` reads its source off `ExcelReader` to
    check that every `CalculationType` has a branch.

    None of the mixins carry state. They read `_error`, `logstring`,
    `questionList` and the class constants below off `self`, the same object
    they read them off before. The constants stay here on purpose: several are
    shared across the split, and separating one from a caller is how the two
    drift apart.
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
                q.rowIndex = row_idx
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

    # Thin wrappers over cell_text, kept as methods because the shorter name
    # reads better at the ~20 call sites across this class and its mixins.
    # The implementations are shared with crf_reader and xml_generator so the
    # same cell cannot be read two different ways.
    #
    # There is deliberately no `_to_str` alias: nothing ever called it, and
    # the two callers that want that function are `cell_trim`/`cell_raw`
    # inside cell_text itself.
    _split_lines = staticmethod(split_cell_lines)

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
