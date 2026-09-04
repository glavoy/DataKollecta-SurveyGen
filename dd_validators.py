"""The checks that need the whole worksheet, not one cell.

Split out of `excel_reader.py`. These are the thirteen validations
`ExcelReader.create_question_list` runs in one block *after* the row loop has
finished, once `self.questionList` is complete -- cross-references between
questions, ordering rules, and the reserved-variable rules that can only be
judged against every row at once.

That is a real seam in the control flow, not just a naming one. The per-cell
checks (`_check_field_name`, `_check_question_field_type`, `_check_ranges`'
scalar cousins and so on) take scalar arguments and run *mid-row*, interleaved
with reading the cells; these take a worksheet name and read `self.questionList`.
Leaving the per-cell checks behind keeps `excel_reader.py` readable as the one
thing it is -- read the sheet, check each cell as you go -- and puts everything
that needs the finished list here.

`_part_field_refs` and `_question_field_refs` come too. They are used only by
`_check_reserved_variable_reads` and exist to walk a calculation's operands
looking for field references.

**No cell reader may be reimplemented in here.** `cell_text.py` exists because
the validator and the emitter once had divergent copies of that logic, and the
worst bug the review found was a validator that passed a cell the generator then
dropped. Anything in this module that needs a cell goes through
`self._get_cell_trim` / `self._split_lines`, which resolve to `cell_text`.
Nothing here touches a worksheet cell directly today, and nothing should start.

Twenty of the ninety-four `_error` call sites live here; this module is listed
in `SCANNED_MODULES` in `tests/test_error_message_shape.py`.
"""

from __future__ import annotations

from models import (
    KNOWN_AUTOMATIC_FIELDS,
    RESERVED_SYSTEM_FIELDS,
    TRAILING_SYSTEM_FIELD_NAMES,
    CalculationPart,
    CalculationType,
    Question,
    ResponseSourceType,
)
from skip_parser import parse_skip, split_skip_lines


class WorksheetValidationMixin:
    """The whole-worksheet checks of `ExcelReader`. Not usable on its own.

    Reads `_error`, `logstring`, `questionList`, `suppliedAutoFields` and
    `reservedFieldsWithCalculation` off the host, plus the class constants
    `PLACEHOLDER_RE`, `FIELD_NAME_RE`, `QUOTED_STRING_RE`, `LOGIC_KEYWORDS`,
    `BUILT_IN_AUTO_FIELDS`, `TYPED_FIELD_TYPES`, `NUMERIC_FIELD_TYPES`,
    `SELECTION_QUESTION_TYPES` and `HOURMIN_MAX_CHARACTERS`.

    Those constants deliberately stay on `ExcelReader` rather than moving here.
    Several are shared with checks that did not move, and splitting a constant
    from one of its two callers is how the two drift apart -- the same failure
    `cell_text.py` was created to end.
    """

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
