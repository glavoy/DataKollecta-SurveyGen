import unittest

from openpyxl import Workbook

from excel_reader import ExcelReader


HEADERS = [
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

# A dictionary written before the Optional column was added still has "NA"
# there -- still accepted, but its contents are ignored entirely.
LEGACY_NA_HEADERS = [h if h != "Optional" else "NA" for h in HEADERS]


def row(fieldname, qtype, ftype, text="Question text", maxchars="", responses="",
        skip="", lower="", upper="", optional=""):
    return [fieldname, qtype, ftype, text, maxchars, responses, lower, upper,
            "", "", "", optional, skip, ""]


def numeric_row(fieldname="age", ftype="text_integer", maxchars="3", **kwargs):
    """A numeric question with nothing for the validations to complain about."""
    kwargs.setdefault("lower", "0")
    kwargs.setdefault("upper", "120")
    return row(fieldname, "text", ftype, maxchars=maxchars, **kwargs)


def read(rows, supplied=None, headers=HEADERS):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "demo_dd"
    worksheet.append(headers)
    for r in rows:
        worksheet.append(r)
    reader = ExcelReader(supplied_auto_fields=supplied)
    reader.create_question_list(worksheet)
    return reader


def errors(reader):
    return [line for line in reader.logstring if line.startswith("ERROR")]


def warnings(reader):
    return [line for line in reader.logstring if line.startswith("WARNING")]


class AutomaticNeedsCalculationTests(unittest.TestCase):
    """An automatic field with no calculation is never given a value, and any
    skip that tests it then fails open."""

    def test_blank_responses_is_an_error(self):
        reader = read([row("region", "automatic", "integer")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("has no calculation (the Responses column is blank)", errors(reader)[0])

    def test_response_options_instead_of_a_calc_block_is_an_error(self):
        # The Responses column on an automatic question must hold a calculation;
        # an option list there is silently ignored at runtime.
        reader = read([row("feverortemp", "automatic", "integer", responses="1:Yes\n0:No")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("does not start with 'calc:'", errors(reader)[0])

    def test_a_calc_block_is_accepted(self):
        reader = read(
            [row("total", "automatic", "integer", responses="calc:constant\nvalue:1")]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_built_in_automatic_fields_are_exempt(self):
        reader = read(
            [
                row("starttime", "automatic", "datetime"),
                row("startdate", "automatic", "date"),
                row("uniqueid", "automatic", "text"),
                row("lastmod", "automatic", "datetime"),
                row("swver", "automatic", "text"),
                row("survey_id", "automatic", "text"),
                row("stoptime", "automatic", "datetime"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_computed_automatic_variables_are_exempt_regardless_of_idconfig(self):
        # The bug this guards against: yyyy/yy/mm/dd/doy are exempt via
        # KNOWN_AUTOMATIC_FIELDS, not via `supplied` (idconfig.fields) --
        # they must not error on a worksheet whose table has no idconfig at
        # all, e.g. a repeating child form linked by its parent's key rather
        # than ID-generated. No `supplied` set is passed here on purpose.
        reader = read(
            [
                row("yyyy", "automatic", "text"),
                row("yy", "automatic", "text"),
                row("mm", "automatic", "text"),
                row("dd", "automatic", "text"),
                row("doy", "automatic", "text"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_fields_supplied_by_the_manifest_are_exempt(self):
        # hhid and linenum come from the crfs linking/increment fields.
        reader = read(
            [
                row("hhid", "automatic", "integer"),
                row("linenum", "automatic", "integer"),
            ],
            supplied={"hhid", "linenum"},
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_an_unsupplied_field_is_still_reported(self):
        reader = read(
            [
                row("hhid", "automatic", "integer"),
                row("region", "automatic", "integer"),
            ],
            supplied={"hhid"},
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertEqual(len(errors(reader)), 1)
        self.assertIn("'region'", errors(reader)[0])

    def test_a_bare_datetime_field_is_no_longer_exempt(self):
        # There used to be a FieldType-based exemption here: a non-reserved,
        # non-KNOWN_AUTOMATIC_FIELDS name with FieldType datetime and a blank
        # Responses column was treated as a deliberate mid-questionnaire
        # timestamp. That implicit fallback is gone -- use calc:timestamp
        # instead (see tests/test_timestamp_calc.py and "Custom Timestamp
        # Fields" in README.md).
        reader = read([row("time_eligible", "automatic", "datetime")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("has no calculation (the Responses column is blank)", errors(reader)[0])

    def test_a_custom_non_datetime_field_is_still_reported(self):
        reader = read([row("region", "automatic", "date")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("has no calculation (the Responses column is blank)", errors(reader)[0])


class AnswerableResponsesTests(unittest.TestCase):
    """A selection question with no options cannot be answered."""

    def test_radio_with_no_responses_is_an_error(self):
        reader = read([row("memberstructure", "radio", "integer")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("cannot be answered", errors(reader)[0])

    def test_checkbox_with_no_responses_is_an_error(self):
        reader = read([row("symptoms", "checkbox", "text")])

        self.assertTrue(reader.errorsEncountered)

    def test_static_options_are_accepted(self):
        reader = read([row("sex", "radio", "integer", responses="1:Male\n2:Female")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_a_database_source_is_accepted(self):
        reader = read(
            [
                row(
                    "memberstructure",
                    "radio",
                    "integer",
                    responses="source:database\ntable:sleeping_structure\n"
                    "display:structurelabel\nvalue:structurenum",
                )
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_a_text_question_needs_no_responses(self):
        reader = read([row("notes", "text", "text", maxchars="80")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))


class SelfReferencingPreskipTests(unittest.TestCase):
    """A preskip cannot test the answer to its own question."""

    def test_preskip_on_its_own_field_is_an_error(self):
        reader = read(
            [
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="preskip: if everhung = 0, skip to netshape"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("tests its own field", errors(reader)[0])

    def test_postskip_on_its_own_field_is_correct(self):
        reader = read(
            [
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if everhung = 0, skip to netshape"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_preskip_on_an_earlier_field_is_correct(self):
        reader = read(
            [
                row("obs", "radio", "integer", responses="1:Yes\n0:No"),
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="preskip: if obs = 1, skip to netshape"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))


class MaxCharactersWarningTests(unittest.TestCase):
    """MaxCharacters only affects typed input."""

    def test_max_characters_on_a_radio_warns(self):
        reader = read(
            [row("mthfcty", "radio", "integer", maxchars="10", responses="1:Car\n2:Bus")]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(len(warnings(reader)), 1)
        self.assertIn("MaxCharacters is ignored", warnings(reader)[0])

    def test_max_characters_on_a_text_question_is_silent(self):
        reader = read([row("notes", "text", "text", maxchars="80")])

        self.assertEqual(warnings(reader), [])


class AutomaticTypeAliasTests(unittest.TestCase):
    """`calc`, `calculation` and `calculated` all mean `automatic`."""

    def _question_type_for(self, declared):
        reader = read(
            [row("total", declared, "integer", responses="calc:constant\nvalue:1")]
        )
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        return reader.questionList[0].questionType

    def test_every_spelling_normalizes_to_automatic(self):
        for declared in ("automatic", "calc", "calculation", "calculated"):
            with self.subTest(declared=declared):
                self.assertEqual(self._question_type_for(declared), "automatic")

    def test_spelling_is_case_insensitive(self):
        self.assertEqual(self._question_type_for("Calculated"), "automatic")

    def test_an_unrelated_type_is_untouched(self):
        reader = read([row("sex", "radio", "integer", responses="1:Male\n2:Female")])
        self.assertEqual(reader.questionList[0].questionType, "radio")

    def test_a_real_typo_is_still_rejected(self):
        reader = read([row("total", "calculatd", "integer")])
        self.assertTrue(reader.errorsEncountered)


class ReservedAutomaticFieldTests(unittest.TestCase):
    """Reserved names are pointed out but never fail the build."""

    def test_declaring_one_warns_but_does_not_error(self):
        reader = read([row("starttime", "automatic", "datetime")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        warning = "\n".join(warnings(reader))
        self.assertIn("starttime", warning)
        self.assertIn("is reserved", warning)
        self.assertIn("this row is ignored", warning)

    def test_every_reserved_name_is_covered(self):
        names = ["starttime", "startdate", "stoptime", "lastmod",
                 "uniqueid", "swver", "survey_id"]
        reader = read([row(n, "automatic", "text") for n in names])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(len(warnings(reader)), len(names))

    def test_a_calculation_on_a_reserved_name_warns_it_is_ignored(self):
        # The generator drops the calculation rather than writing it out, so
        # an author who writes one gets no effect at all.
        reader = read(
            [row("starttime", "automatic", "datetime",
                 responses="calc:constant\nvalue:x")]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertIn("calculation is ignored", "\n".join(warnings(reader)))
        self.assertEqual(len(warnings(reader)), 1)

    def test_an_ordinary_field_is_not_flagged(self):
        reader = read([numeric_row()])

        self.assertEqual(warnings(reader), [])

    def test_an_uppercase_field_name_is_rejected_before_this_check(self):
        # FieldNames must be lowercase, so a reserved name can only ever
        # arrive in lowercase.
        reader = read([row("StartTime", "automatic", "datetime")])

        self.assertTrue(reader.errorsEncountered)


class SkipToReservedFieldTests(unittest.TestCase):
    """A skip can never target a reserved variable.

    The generator decides where those go, so the target is never in the place
    the dictionary imagined: the leading pair sits before the first question
    and the rest after the last.
    """

    def test_skipping_to_a_reserved_name_is_an_error(self):
        reader = read(
            [
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if everhung = 0, skip to uniqueid"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        message = "\n".join(errors(reader))
        self.assertIn("uniqueid", message)
        self.assertIn("reserved variable", message)
        # The old message called it nonexistent, which was misleading: the
        # generator does write it, just somewhere else.
        self.assertNotIn("nonexistent", message)

    def test_declaring_the_target_does_not_make_it_legal(self):
        # The case that used to pass silently. Declaring the row put it in
        # field_index, so the target resolved -- but the generator drops the
        # declared row and writes its own at the end of the questionnaire, so
        # the skip jumped past every remaining question.
        reader = read(
            [
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if everhung = 0, skip to stoptime"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
                row("stoptime", "automatic", "datetime"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("reserved variable", "\n".join(errors(reader)))

    def test_every_reserved_name_is_rejected(self):
        names = ["starttime", "startdate", "stoptime", "lastmod",
                 "uniqueid", "swver", "survey_id"]
        for name in names:
            with self.subTest(name=name):
                reader = read(
                    [
                        row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                            skip=f"postskip: if everhung = 0, skip to {name}"),
                        row("netshape", "radio", "integer",
                            responses="1:Round\n2:Square"),
                    ]
                )

                self.assertTrue(reader.errorsEncountered)
                self.assertIn(name, "\n".join(errors(reader)))

    def test_skipping_to_a_real_question_is_still_correct(self):
        reader = read(
            [
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if everhung = 0, skip to netshape"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_a_genuinely_missing_target_keeps_its_own_message(self):
        reader = read(
            [
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if everhung = 0, skip to nowhere"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("nonexistent FieldName", "\n".join(errors(reader)))


class OptionalColumnTests(unittest.TestCase):
    """Optional replaces the old NA column: a text question with Optional set
    to TRUE may be left blank -- the Next button stays enabled even with no
    answer. Restricted to QuestionType 'text'; choice questions already have
    DontKnow/Refuse for an explicit non-answer."""

    def test_optional_true_on_a_text_question_is_accepted(self):
        reader = read([row("notes", "text", "text", maxchars="80", optional="True")])
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_all_caps_true_is_also_accepted(self):
        # The generator itself already treats 'TRUE' as set
        # (xml_generator.py checks {"TRUE", "True"}); the validator used to
        # reject it, which was an inconsistency, not a real rule.
        reader = read([row("notes", "text", "text", maxchars="80", optional="TRUE")])
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_an_unset_optional_column_is_silent(self):
        reader = read([row("notes", "text", "text", maxchars="80")])
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_an_invalid_value_is_an_error(self):
        reader = read([row("notes", "text", "text", maxchars="80", optional="yes")])
        self.assertTrue(reader.errorsEncountered)
        self.assertIn("Optional", "\n".join(errors(reader)))

    def test_optional_on_a_non_text_question_is_an_error(self):
        reader = read(
            [row("consent", "radio", "integer", responses="1:Yes\n0:No", optional="True")]
        )
        self.assertTrue(reader.errorsEncountered)
        message = "\n".join(errors(reader))
        self.assertIn("Optional", message)
        self.assertIn("text", message)

    def test_a_comments_field_without_optional_warns(self):
        reader = read([row("comments", "text", "text", maxchars="80")])
        self.assertFalse(reader.errorsEncountered)
        self.assertIn("comments", "\n".join(warnings(reader)))

    def test_a_comments_field_with_optional_does_not_warn(self):
        reader = read([row("comments", "text", "text", maxchars="80", optional="True")])
        self.assertEqual(warnings(reader), [])

    def test_an_unrelated_field_without_optional_does_not_warn(self):
        reader = read([row("notes", "text", "text", maxchars="80")])
        self.assertEqual(warnings(reader), [])

    def test_a_legacy_na_header_is_still_accepted(self):
        reader = read([row("notes", "text", "text", maxchars="80")], headers=LEGACY_NA_HEADERS)
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_a_legacy_na_columns_contents_are_ignored_entirely(self):
        # Whatever an old sheet has in the NA cell -- valid, garbage, or a
        # value that would have set Optional -- must never be parsed. A
        # stray value there was never meant to make anything optional.
        reader = read(
            [row("notes", "text", "text", maxchars="80", optional="garbage")],
            headers=LEGACY_NA_HEADERS,
        )
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_a_header_that_is_neither_na_nor_optional_is_rejected(self):
        bad_headers = [h if h != "Optional" else "Something Else" for h in HEADERS]
        reader = read([row("notes", "text", "text", maxchars="80")], headers=bad_headers)
        self.assertTrue(reader.errorsEncountered)
        self.assertIn("header", "\n".join(errors(reader)).lower())


class SkipToEndOfFormTests(unittest.TestCase):
    """`skip to end` is the app's own end-of-form sentinel, not a fieldname --
    see SurveyNavigationService.endOfFormSkipTarget in the app repo."""

    def test_skipping_to_end_is_accepted(self):
        reader = read(
            [
                row("consent", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if consent = 0, skip to end"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_end_is_case_insensitive(self):
        reader = read(
            [
                row("consent", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if consent = 0, skip to END"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_a_field_named_end_is_rejected(self):
        reader = read(
            [
                row("end", "text", "text"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("reserved", "\n".join(errors(reader)))

    def test_a_genuinely_missing_target_still_errors(self):
        # 'end' must not become a loophole that silently accepts any target.
        reader = read(
            [
                row("consent", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if consent = 0, skip to endzone"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("nonexistent FieldName", "\n".join(errors(reader)))


class SkipTestingAReservedFieldTests(unittest.TestCase):
    """Which questions get asked must not depend on a generator-supplied value."""

    def test_testing_a_trailing_variable_is_an_error(self):
        # Empty while the questionnaire is being answered, so the skip would
        # never fire and the guarded question would be asked of everyone.
        reader = read(
            [
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if survey_id = 2, skip to netshape"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("reserved variable", "\n".join(errors(reader)))

    def test_testing_a_leading_variable_is_also_an_error(self):
        # This one would work at runtime, but it makes a single package ask
        # different questions on different days. Regenerate instead.
        reader = read(
            [
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="postskip: if startdate = 2026-01-01, skip to netshape"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("reserved variable", "\n".join(errors(reader)))

    def test_testing_an_ordinary_field_is_still_correct(self):
        reader = read(
            [
                row("obs", "radio", "integer", responses="1:Yes\n0:No"),
                row("everhung", "radio", "integer", responses="1:Yes\n0:No",
                    skip="preskip: if obs = 1, skip to netshape"),
                row("netshape", "radio", "integer", responses="1:Round\n2:Square"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))


class TrailingVariableInExpressionTests(unittest.TestCase):
    """A trailing variable is empty during the interview, so reading one is
    always a silent failure -- in a logic check or in the Responses column."""

    def test_logic_check_referencing_one_is_an_error(self):
        # `row` has no logic-check argument, so set the column directly.
        workbook_rows = [numeric_row("age"), numeric_row("confirm_age")]
        workbook_rows[1][8] = "lastmod < startdate; 'Bad date!'"
        reader = read(workbook_rows)

        self.assertTrue(reader.errorsEncountered)
        message = "\n".join(errors(reader))
        self.assertIn("lastmod", message)
        self.assertIn("reserved variable", message)

    def test_logic_check_referencing_a_leading_variable_is_allowed(self):
        workbook_rows = [numeric_row("age"), numeric_row("confirm_age")]
        workbook_rows[1][8] = "confirm_age <> age; 'That does not match!'"
        reader = read(workbook_rows)

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_a_calculation_placeholder_on_a_leading_variable_is_allowed(self):
        # The real AVERT age calculation. This must keep working.
        reader = read(
            [
                row("dob", "date", "date", lower="1900-01-01", upper="2050-12-31"),
                row("age_calculated", "automatic", "integer",
                    responses="calc:age_at_date\nfield:dob\nvalue:years\n"
                              "separator:[[startdate]]"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_a_calculation_placeholder_on_a_trailing_variable_is_an_error(self):
        # The realistic slip: `lastmod` reads like "the interview date".
        reader = read(
            [
                row("dob", "date", "date", lower="1900-01-01", upper="2050-12-31"),
                row("age_calculated", "automatic", "integer",
                    responses="calc:age_at_date\nfield:dob\nvalue:years\n"
                              "separator:[[lastmod]]"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("lastmod", "\n".join(errors(reader)))

    def test_a_case_condition_on_a_trailing_variable_is_an_error(self):
        reader = read(
            [
                row("need_visit", "automatic", "integer",
                    responses="calc:case\nwhen:uniqueid = 1 => 1\nelse:0"),
            ]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("uniqueid", "\n".join(errors(reader)))

    def test_a_lookup_on_a_trailing_variable_is_an_error(self):
        reader = read(
            [row("copy", "automatic", "text", responses="calc:lookup\nfield:swver")]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("swver", "\n".join(errors(reader)))

    def test_question_text_referencing_one_is_an_error(self):
        # Same emptiness, so the respondent would be shown a gap in the
        # sentence. Nothing about being display-only makes it work.
        reader = read(
            [row("q1", "text", "text", text="Recorded at [[lastmod]]?", maxchars="20")]
        )

        self.assertTrue(reader.errorsEncountered)
        message = "\n".join(errors(reader))
        self.assertIn("lastmod", message)
        self.assertIn("question text", message)

    def test_question_text_referencing_an_ordinary_field_is_allowed(self):
        reader = read(
            [
                row("participantsname", "text", "text", maxchars="40"),
                row("q1", "text", "text", text="Is [[participantsname]] well?",
                    maxchars="20"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_question_text_referencing_startdate_is_allowed(self):
        reader = read(
            [row("q1", "text", "text", text="On [[startdate]], was the child well?",
                 maxchars="20")]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_the_message_names_where_the_reference_was_found(self):
        reader = read(
            [row("copy", "automatic", "text", responses="calc:lookup\nfield:swver")]
        )

        self.assertIn("its calculation", "\n".join(errors(reader)))

    def test_a_placeholder_in_a_logic_message_warns(self):
        # Messages are not expanded, so this is not an empty value -- it is the
        # literal brackets on screen. Warned about for every field name, not
        # only reserved ones, because none of them work there.
        workbook_rows = [numeric_row("age"), numeric_row("confirm_age")]
        workbook_rows[1][8] = "confirm_age <> age; 'Does not match [[age]]!'"
        reader = read(workbook_rows)

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        warning = "\n".join(warnings(reader))
        self.assertIn("[[age]]", warning)
        self.assertIn("shown exactly as written", warning)

    def test_a_reserved_placeholder_in_a_logic_message_warns_too(self):
        workbook_rows = [numeric_row("age"), numeric_row("confirm_age")]
        workbook_rows[1][8] = "confirm_age <> age; 'Recorded [[lastmod]]!'"
        reader = read(workbook_rows)

        self.assertIn("[[lastmod]]", "\n".join(warnings(reader)))

    def test_a_message_without_a_placeholder_is_silent(self):
        workbook_rows = [numeric_row("age"), numeric_row("confirm_age")]
        workbook_rows[1][8] = "confirm_age <> age; 'That does not match!'"
        reader = read(workbook_rows)

        self.assertEqual(warnings(reader), [])

    def test_a_declared_reserved_row_is_not_double_reported(self):
        # Declaring one already warns; its calculation is dropped, so the
        # Responses check must stay quiet rather than adding an error.
        reader = read(
            [row("stoptime", "automatic", "datetime",
                 responses="calc:lookup\nfield:lastmod")]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

class FieldNameMessageTests(unittest.TestCase):
    """The message an author gets should name what is actually wrong."""

    def test_a_field_name_with_a_space_says_so(self):
        # A space is neither alnum nor "_", so the generic
        # "only letters, digits and underscores" branch used to fire first and
        # this specific message was unreachable.
        reader = read([numeric_row("has space")])
        self.assertTrue(reader.errorsEncountered)
        self.assertIn("contains a space", "\n".join(errors(reader)))

    def test_other_invalid_characters_still_get_the_generic_message(self):
        reader = read([numeric_row("has-dash")])
        self.assertTrue(reader.errorsEncountered)
        joined = "\n".join(errors(reader))
        self.assertIn("Only letters, digits, and underscores", joined)
        self.assertNotIn("contains a space", joined)

    def test_a_leading_underscore_is_still_reported(self):
        reader = read([numeric_row("_leading")])
        self.assertTrue(reader.errorsEncountered)
        self.assertIn("starts with an underscore", "\n".join(errors(reader)))


if __name__ == "__main__":
    unittest.main()


class ComboboxStaticResponsesTests(unittest.TestCase):
    """A combobox's static options are split on ':' exactly as a radio's are.

    xml_generator emits radio, checkbox and combobox through the same
    `response.find(":")` split, but only radio and checkbox were format-checked
    -- so `Yes` on a combobox stored the value "Ye" (find returns -1, and
    `response[:-1]` is the whole string minus its last character). Valid XML,
    wrong code, on every answer to that question.
    """

    def test_a_combobox_option_without_a_colon_is_rejected(self):
        reader = read([row("consent", "combobox", "integer", responses="Yes\nNo")])

        self.assertTrue(
            any("Invalid static combobox options" in e for e in errors(reader)),
            errors(reader),
        )

    def test_a_correctly_formatted_combobox_is_accepted(self):
        reader = read([row("consent", "combobox", "integer", responses="1:Yes\n2:No")])

        self.assertEqual(errors(reader), [])

    def test_the_message_names_the_question_type_it_is_talking_about(self):
        # It used to say "radio button options" whatever the question was.
        reader = read([row("consent", "checkbox", "text", responses="Yes")])

        self.assertTrue(
            any("Invalid static checkbox options" in e for e in errors(reader)),
            errors(reader),
        )

    def test_duplicate_codes_are_caught_on_a_combobox_too(self):
        reader = read([row("consent", "combobox", "integer", responses="1:Yes\n1:No")])

        self.assertTrue(any("has duplicates" in e for e in errors(reader)), errors(reader))


class RangeOnSelectionQuestionTests(unittest.TestCase):
    """A LowerRange left on a selection question makes it unanswerable.

    xml_generator writes <numeric_check> for any question whose LowerRange is
    set (it only excludes 'date'), so a range copied onto a radio row is
    emitted and then rejects every option code the interviewer can pick.
    _check_ranges only ever looked at 'text' and 'date' questions.
    """

    def test_a_range_on_a_radio_is_rejected(self):
        reader = read([
            row("sex", "radio", "integer", responses="1:Male\n2:Female", lower="1", upper="2")
        ])

        self.assertTrue(
            any("is a 'radio' question with a" in e for e in errors(reader)),
            errors(reader),
        )

    def test_a_half_set_range_on_a_checkbox_is_rejected_too(self):
        # The worse half: the generator writes the blank UpperRange as
        # maxvalue='-9', so every value above -9 fails.
        reader = read([
            row("symptoms", "checkbox", "text", responses="1:Fever\n2:Cough", lower="1")
        ])

        self.assertTrue(
            any("LowerRange or UpperRange" in e for e in errors(reader)),
            errors(reader),
        )

    def test_a_combobox_is_covered(self):
        reader = read([
            row("village", "combobox", "integer", responses="1:Kirembe", upper="99")
        ])

        self.assertTrue(
            any("is a 'combobox' question with a" in e for e in errors(reader)),
            errors(reader),
        )

    def test_a_selection_question_with_no_range_is_untouched(self):
        reader = read([row("sex", "radio", "integer", responses="1:Male\n2:Female")])

        self.assertEqual(errors(reader), [])

    def test_a_typed_number_still_carries_its_range(self):
        reader = read([numeric_row()])

        self.assertEqual(errors(reader), [])


class UnsupportedFilterOperatorTests(unittest.TestCase):
    """An operator the filter grammar does not know must not become '='.

    FILTER_MATCH_RE's operator group is optional, so with nothing recognised
    the entire remainder becomes the value: `filter:region like North` parsed
    as `region = 'like North'`, generated clean XML, and came back from the
    device as an empty response list with nothing said at generation time.
    """

    @staticmethod
    def combobox_with(filter_text):
        return row(
            "village",
            "combobox",
            "integer",
            responses=(
                "source:database\ntable:villages\n"
                f"filter:{filter_text}\ndisplay:name\nvalue:code"
            ),
        )

    def filter_errors(self, filter_text):
        reader = read([self.combobox_with(filter_text)])
        return [e for e in errors(reader) if "filter operator" in e]

    def test_a_word_operator_is_named_rather_than_swallowed(self):
        found = self.filter_errors("region like North")

        self.assertEqual(len(found), 1, found)
        self.assertIn("'like'", found[0])

    def test_contains_is_caught_even_though_skips_accept_it(self):
        # 'contains' is a legal skip operator, which is exactly why an author
        # reaches for it here.
        self.assertTrue(self.filter_errors("region contains North"))

    def test_a_multi_word_operator_is_caught(self):
        found = self.filter_errors("region is not null")

        self.assertTrue(found)
        self.assertIn("'is not'", found[0])

    def test_a_mistyped_symbolic_operator_is_caught(self):
        # The regex matches the '=' and strands the '<' at the front of the
        # value, so the filter silently became `district = '< 5'`.
        found = self.filter_errors("district =< 5")

        self.assertTrue(found)
        self.assertIn("=<", found[0])

    def test_the_message_lists_the_operators_that_do_work(self):
        found = self.filter_errors("region like North")

        self.assertIn("in and not in", found[0])

    def test_a_legitimate_multi_word_value_is_still_accepted(self):
        # `district North West` means district = 'North West'. The operator is
        # optional by design, so only the leading token may be inspected.
        self.assertEqual(self.filter_errors("district North West"), [])

    def test_every_supported_operator_still_parses(self):
        for supported in ["= North", "!= North", "<> North", "in 1,2,3",
                          "not in 1,2", "> 5", ">= 5", "< 5", "<= 5"]:
            with self.subTest(operator=supported):
                self.assertEqual(self.filter_errors(f"region {supported}"), [])


class FieldNameCharacterTests(unittest.TestCase):
    """FieldName is ASCII, because it becomes an XML attribute and a column.

    `str.isalnum()` is Unicode-aware, so every accented spelling the French
    dictionaries naturally reach for passed silently.
    """

    def test_an_accented_field_name_is_rejected(self):
        reader = read([row("prénom", "text", "text")])

        self.assertTrue(
            any("Only letters, digits, and underscores" in e for e in errors(reader)),
            errors(reader),
        )

    def test_a_field_name_with_a_cedilla_is_rejected(self):
        reader = read([row("français", "text", "text")])

        self.assertTrue(errors(reader))

    def test_a_plain_ascii_field_name_is_still_accepted(self):
        reader = read([row("prenom", "text", "text", maxchars="30")])

        self.assertEqual(errors(reader), [])

    def test_digits_and_underscores_are_still_accepted(self):
        reader = read([row("net_brand_2", "text", "text", maxchars="30")])

        self.assertEqual(errors(reader), [])

    def test_a_space_still_gets_its_own_more_specific_message(self):
        reader = read([row("first name", "text", "text")])

        self.assertTrue(any("contains a space" in e for e in errors(reader)), errors(reader))


class CrossRowChecksRunRegardlessTests(unittest.TestCase):
    """A row-level error no longer suppresses the structural checks.

    The offending row is still in questionList (only a blank FieldName is
    dropped), so the cross-row checks see the same data either way -- the gate
    only delayed their output to a later run.
    """

    def test_a_row_error_and_a_structural_error_are_reported_together(self):
        reader = read([
            # Row error: blank QuestionText.
            row("age", "text", "text_integer", text="", maxchars="3", lower="0", upper="120"),
            # Structural error: a skip pointing at a field that does not exist.
            row("sex", "radio", "integer", responses="1:Male\n2:Female",
                skip="postskip: if sex = 1, skip to nowhere"),
        ])

        found = errors(reader)
        self.assertTrue(any("blank QuestionText" in e for e in found), found)
        self.assertTrue(any("nowhere" in e for e in found), found)

    def test_a_clean_worksheet_still_says_so(self):
        reader = read([numeric_row()])

        self.assertIn("No errors found in 'demo_dd'", reader.logstring)
