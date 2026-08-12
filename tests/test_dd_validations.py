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
    "NA",
    "Skip",
    "Comments",
]


def row(fieldname, qtype, ftype, text="Question text", maxchars="", responses="",
        skip="", lower="", upper=""):
    return [fieldname, qtype, ftype, text, maxchars, responses, lower, upper,
            "", "", "", "", skip, ""]


def numeric_row(fieldname="age", ftype="text_integer", maxchars="3", **kwargs):
    """A numeric question with nothing for the validations to complain about."""
    kwargs.setdefault("lower", "0")
    kwargs.setdefault("upper", "120")
    return row(fieldname, "text", ftype, maxchars=maxchars, **kwargs)


def read(rows, supplied=None):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "demo_dd"
    worksheet.append(HEADERS)
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
        reader = read([row("comments", "text", "text", maxchars="80")])

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
        reader = read([row("comments", "text", "text", maxchars="80")])

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


if __name__ == "__main__":
    unittest.main()
