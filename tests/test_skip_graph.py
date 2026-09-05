"""What the skips do to a form, taken as a whole.

Each rule gets a case that fires, a case that does not, and the boundary
between them. The does-not-fire cases carry as much weight as the others: this
module has to stay silent on two real studies with 182 skip cells between them,
and a rule that cannot be quiet is a rule that gets the whole linter switched
off.
"""

from __future__ import annotations

import unittest

import skip_graph
from skip_graph import Severity, lint_form
from tests.test_dd_validations import HEADERS, numeric_row, read, row

_DONT_KNOW_COLUMN = 9
_REFUSE_COLUMN = 10


def specials(base_row, dont_know=False, refuse=False):
    """A row with the Don't know / Refuse buttons switched on.

    `row()` has no parameter for them, and adding one would touch every test
    that imports it. Both cells are truthy-by-string in the reader, so setting
    them positionally here is the same thing the reader will read.
    """
    copy = list(base_row)
    if dont_know:
        copy[_DONT_KNOW_COLUMN] = "TRUE"
    if refuse:
        copy[_REFUSE_COLUMN] = "TRUE"
    return copy


def findings(rows, rule_id=None):
    reader = read(rows)
    found = lint_form("demo_dd", reader.questionList)
    if rule_id is None:
        return found
    return [f for f in found if f.rule_id == rule_id]


def messages(rows, rule_id=None):
    return [f.message for f in findings(rows, rule_id)]


class DeadBranchTests(unittest.TestCase):
    RULE = "skip.graph.deadBranch"

    def test_a_value_outside_the_response_codes_can_never_fire(self):
        found = findings(
            [
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
                row("pregnant", "radio", "integer", responses="1:Yes\n0:No",
                    skip="preskip: if sex = 3, skip to occupation"),
                row("occupation", "text", "text"),
            ],
            self.RULE,
        )
        self.assertEqual(len(found), 1)
        self.assertIn("can hold only 1, 2", found[0].message)
        self.assertIs(found[0].severity, Severity.ERROR)

    def test_a_value_that_is_a_response_code_is_silent(self):
        self.assertEqual(
            findings(
                [
                    row("sex", "radio", "integer", responses="1:Male\n2:Female"),
                    row("pregnant", "radio", "integer", responses="1:Yes\n0:No",
                        skip="preskip: if sex = 2, skip to occupation"),
                    row("occupation", "text", "text"),
                ],
                self.RULE,
            ),
            [],
        )

    def test_the_dont_know_code_is_answerable_when_the_button_is_on(self):
        # -7 is never in the Responses cell, but it is a value the field can
        # hold. Calling a skip that tests it a dead branch would be wrong, and
        # would be wrong on a construct dictionaries actually use.
        rows = [
            specials(
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
                dont_know=True,
            ),
            row("pregnant", "radio", "integer", responses="1:Yes\n0:No",
                skip="preskip: if sex = -7, skip to occupation"),
            row("occupation", "text", "text"),
        ]
        self.assertEqual(findings(rows, self.RULE), [])

    def test_the_dont_know_code_is_dead_when_the_button_is_off(self):
        rows = [
            row("sex", "radio", "integer", responses="1:Male\n2:Female"),
            row("pregnant", "radio", "integer", responses="1:Yes\n0:No",
                skip="preskip: if sex = -7, skip to occupation"),
            row("occupation", "text", "text"),
        ]
        self.assertEqual(len(findings(rows, self.RULE)), 1)

    def test_a_value_outside_a_fully_declared_range_can_never_fire(self):
        found = findings(
            [
                numeric_row("age", lower="0", upper="120"),
                row("retired", "radio", "integer", responses="1:Yes\n0:No",
                    skip="preskip: if age > 200, skip to occupation"),
                row("occupation", "text", "text"),
            ],
            self.RULE,
        )
        self.assertEqual(len(found), 1)
        self.assertIn("0 to 120", found[0].message)

    def test_a_half_declared_range_makes_no_claim(self):
        # An interval open at one end holds values this module cannot
        # enumerate, so nothing is decidable and nothing is said.
        rows = [
            row("age", "text", "text_integer", maxchars="3", lower="0"),
            row("retired", "radio", "integer", responses="1:Yes\n0:No",
                skip="preskip: if age > 200, skip to occupation"),
            row("occupation", "text", "text"),
        ]
        self.assertEqual(findings(rows, self.RULE), [])

    def test_a_csv_backed_question_makes_no_claim(self):
        rows = [
            row("village", "radio", "integer",
                responses="source: csv\nfile: villages.csv\ndisplay: name\nvalue: code"),
            row("distance", "text", "text_integer", maxchars="3", lower="0", upper="99",
                skip="preskip: if village = 9999, skip to occupation"),
            row("occupation", "text", "text"),
        ]
        self.assertEqual(findings(rows, self.RULE), [])

    def test_a_checkbox_contains_rule_makes_no_claim(self):
        # A checkbox stores a comma-joined list, so neither its domain nor the
        # word operators over it are modelled.
        rows = [
            row("symptoms", "checkbox", "integer", responses="1:Fever\n2:Cough"),
            row("details", "text", "text",
                skip="preskip: if symptoms contains 9, skip to occupation"),
            row("occupation", "text", "text"),
        ]
        self.assertEqual(findings(rows, self.RULE), [])

    def test_the_message_names_both_rows(self):
        found = findings(
            [
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
                row("pregnant", "radio", "integer", responses="1:Yes\n0:No",
                    skip="preskip: if sex = 3, skip to occupation"),
                row("occupation", "text", "text"),
            ],
            self.RULE,
        )
        # Row 2 is `sex` (row 1 is the header), row 3 is `pregnant`.
        self.assertIn("row 3", found[0].message)
        self.assertIn("row 2", found[0].message)


class TotalConditionTests(unittest.TestCase):
    RULE = "skip.graph.totalCondition"

    def test_a_condition_true_for_every_code_is_unconditional(self):
        found = findings(
            [
                row("consent", "radio", "integer", responses="0:No\n1:Yes"),
                row("roof", "radio", "integer", responses="1:Thatch\n2:Iron",
                    skip="preskip: if consent >= 0, skip to water"),
                row("water", "text", "text"),
            ],
            self.RULE,
        )
        self.assertEqual(len(found), 1)
        self.assertIn("true for every value", found[0].message)
        self.assertIn("'roof'", found[0].message)

    def test_a_condition_true_for_only_some_codes_is_silent(self):
        self.assertEqual(
            findings(
                [
                    row("consent", "radio", "integer", responses="0:No\n1:Yes"),
                    row("roof", "radio", "integer", responses="1:Thatch\n2:Iron",
                        skip="preskip: if consent > 0, skip to water"),
                    row("water", "text", "text"),
                ],
                self.RULE,
            ),
            [],
        )

    def test_a_not_equals_against_an_absent_code_is_unconditional(self):
        # `<> 9` where 9 is not a response code is true for everything -- the
        # mirror image of the dead branch, and just as silent in the field.
        found = findings(
            [
                row("consent", "radio", "integer", responses="0:No\n1:Yes"),
                row("roof", "radio", "integer", responses="1:Thatch\n2:Iron",
                    skip="preskip: if consent <> 9, skip to water"),
                row("water", "text", "text"),
            ],
            self.RULE,
        )
        self.assertEqual(len(found), 1)

    def test_a_special_code_keeps_a_condition_from_being_total(self):
        # `>= 0` covers both response codes but not -7, so with Don't know on
        # the rule is no longer unconditional. The boundary case for specials.
        rows = [
            specials(
                row("consent", "radio", "integer", responses="0:No\n1:Yes"),
                dont_know=True,
            ),
            row("roof", "radio", "integer", responses="1:Thatch\n2:Iron",
                skip="preskip: if consent >= 0, skip to water"),
            row("water", "text", "text"),
        ]
        self.assertEqual(findings(rows, self.RULE), [])


class ShadowedRuleTests(unittest.TestCase):
    RULE = "skip.graph.shadowedRule"

    def test_a_wider_earlier_rule_hides_a_narrower_later_one(self):
        found = findings(
            [
                numeric_row("age", lower="0", upper="120"),
                row("school", "radio", "integer", responses="1:Yes\n0:No",
                    skip=("postskip: if age < 18, skip to guardian\n"
                          "postskip: if age < 12, skip to guardian")),
                row("guardian", "text", "text"),
            ],
            self.RULE,
        )
        self.assertEqual(len(found), 1)
        self.assertIn("already fires for every value", found[0].message)
        self.assertIn("first match wins", found[0].message)

    def test_the_narrower_rule_first_is_correct(self):
        self.assertEqual(
            findings(
                [
                    numeric_row("age", lower="0", upper="120"),
                    row("school", "radio", "integer", responses="1:Yes\n0:No",
                        skip=("postskip: if age < 12, skip to guardian\n"
                              "postskip: if age < 18, skip to guardian")),
                    row("guardian", "text", "text"),
                ],
                self.RULE,
            ),
            [],
        )

    def test_the_same_condition_twice_with_different_targets_is_reported(self):
        found = findings(
            [
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
                row("pregnant", "radio", "integer", responses="1:Yes\n0:No",
                    skip=("postskip: if sex = 1, skip to occupation\n"
                          "postskip: if sex = 1, skip to income")),
                row("occupation", "text", "text"),
                row("income", "text", "text"),
            ],
            self.RULE,
        )
        self.assertEqual(len(found), 1)
        self.assertIn("skipping to 'occupation' instead", found[0].message)

    def test_an_earlier_unconditional_rule_hides_everything_after_it(self):
        found = findings(
            [
                row("consent", "radio", "integer", responses="0:No\n1:Yes"),
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
                row("roof", "radio", "integer", responses="1:Thatch\n2:Iron",
                    skip=("postskip: if consent >= 0, skip to water\n"
                          "postskip: if sex = 1, skip to water")),
                row("water", "text", "text"),
            ],
            self.RULE,
        )
        self.assertEqual(len(found), 1)
        self.assertIn("is true for every value its field can hold", found[0].message)

    def test_two_rules_on_different_fields_are_not_compared(self):
        # Deciding this would need to know which paths reach the question,
        # which is not decidable here. Staying quiet is the whole reason this
        # rule can be an error.
        self.assertEqual(
            findings(
                [
                    row("sex", "radio", "integer", responses="1:Male\n2:Female"),
                    numeric_row("age", lower="0", upper="120"),
                    row("school", "radio", "integer", responses="1:Yes\n0:No",
                        skip=("postskip: if sex = 1, skip to guardian\n"
                              "postskip: if age < 12, skip to guardian")),
                    row("guardian", "text", "text"),
                ],
                self.RULE,
            ),
            [],
        )

    def test_a_preskip_does_not_shadow_a_postskip(self):
        # They are evaluated at two different moments -- one when navigation
        # lands on the question, the other after it is answered -- so the first
        # cannot pre-empt the second however wide it is.
        self.assertEqual(
            findings(
                [
                    row("consent", "radio", "integer", responses="0:No\n1:Yes"),
                    row("roof", "radio", "integer", responses="1:Thatch\n2:Iron",
                        skip=("preskip: if consent >= 0, skip to water\n"
                              "postskip: if consent = 1, skip to water")),
                    row("water", "text", "text"),
                ],
                self.RULE,
            ),
            [],
        )


class TestedFieldAvailabilityTests(unittest.TestCase):
    """The fail-open trap: an unanswered tested field makes a skip do nothing."""

    NEVER = "skip.graph.testsNeverAnsweredField"
    MAYBE = "skip.graph.testedFieldNotGuaranteed"

    def _fail_open_form(self, guard_cell=""):
        # `gate` postskips over `middle`, and `later` then tests `middle`.
        # On the path the postskip takes, `middle` is blank.
        return [
            row("gate", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if gate = 0, skip to later"),
            row("middle", "radio", "integer", responses="1:Yes\n0:No"),
            row("later", "radio", "integer", responses="1:Yes\n0:No",
                skip=(guard_cell + "preskip: if middle = 1, skip to tail")),
            row("tail", "text", "text"),
        ]

    def test_a_field_blank_on_some_paths_warns(self):
        found = findings(self._fail_open_form(), self.MAYBE)
        self.assertEqual(len(found), 1)
        self.assertIs(found[0].severity, Severity.WARNING)
        self.assertIn("fails open", found[0].message)

    def test_a_field_answered_on_every_path_is_silent(self):
        rows = [
            row("gate", "radio", "integer", responses="1:Yes\n0:No"),
            row("middle", "radio", "integer", responses="1:Yes\n0:No"),
            row("later", "radio", "integer", responses="1:Yes\n0:No",
                skip="preskip: if middle = 1, skip to tail"),
            row("tail", "text", "text"),
        ]
        self.assertEqual(findings(rows, self.MAYBE), [])

    def test_an_earlier_rule_carrying_the_clearing_condition_suppresses_it(self):
        # The author already handled the blank case: on the path where `middle`
        # was cleared, the first rule fires and the second never runs.
        found = findings(
            self._fail_open_form("preskip: if gate = 0, skip to tail\n"), self.MAYBE
        )
        self.assertEqual(found, [])

    def test_a_preskip_suppresses_a_postskip_on_the_same_question(self):
        # A preskip fires before the question is displayed, so when it fires
        # the question's postskips never run at all -- whatever their order.
        rows = [
            row("gate", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if gate = 0, skip to later"),
            row("middle", "radio", "integer", responses="1:Yes\n0:No"),
            row("later", "radio", "integer", responses="1:Yes\n0:No",
                skip=("preskip: if gate = 0, skip to tail\n"
                      "postskip: if middle = 1, skip to tail")),
            row("tail", "text", "text"),
        ]
        self.assertEqual(findings(rows, self.MAYBE), [])

    def test_a_postskip_does_not_suppress_a_preskip(self):
        # The reverse does not hold: a postskip runs only after the question
        # was displayed, which means no preskip fired.
        rows = [
            row("gate", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if gate = 0, skip to later"),
            row("middle", "radio", "integer", responses="1:Yes\n0:No"),
            row("later", "radio", "integer", responses="1:Yes\n0:No",
                skip=("preskip: if middle = 1, skip to tail\n"
                      "postskip: if gate = 0, skip to tail")),
            row("tail", "text", "text"),
        ]
        self.assertEqual(len(findings(rows, self.MAYBE)), 1)

    def test_one_question_testing_one_field_is_reported_once(self):
        rows = [
            row("gate", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if gate = 0, skip to later"),
            row("middle", "radio", "integer", responses="1:Yes\n0:No"),
            row("later", "radio", "integer", responses="1:Yes\n0:No",
                skip=("preskip: if middle = 1, skip to tail\n"
                      "postskip: if middle = 0, skip to tail")),
            row("tail", "text", "text"),
        ]
        self.assertEqual(len(findings(rows, self.MAYBE)), 1)

    def test_an_optional_question_is_never_guaranteed(self):
        # Optional means the interviewer may leave it blank, so a skip testing
        # one fails open for anyone who does.
        rows = [
            row("note", "text", "text", optional="TRUE"),
            row("later", "radio", "integer", responses="1:Yes\n0:No",
                skip="preskip: if note = 1, skip to tail"),
            row("tail", "text", "text"),
        ]
        self.assertEqual(len(findings(rows, self.MAYBE)), 1)

    def test_a_postskip_may_test_its_own_field(self):
        # A postskip runs after its own question is answered, so it always has
        # a value to read -- the one shape that is guaranteed by construction.
        rows = [
            row("gate", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if gate = 0, skip to tail"),
            row("middle", "text", "text"),
            row("tail", "text", "text"),
        ]
        self.assertEqual(findings(rows, self.MAYBE), [])


class SkipsThatNullCalculationsTests(unittest.TestCase):
    NULLS = "skip.graph.skipNullsCalculation"
    ENDS = "skip.graph.endNullsCalculation"

    def test_a_jump_over_a_computable_calculation_warns(self):
        # `derived` reads `price`, which is answered before the jump -- so it
        # could have been computed, and is nulled instead.
        rows = [
            numeric_row("price", lower="0", upper="100"),
            row("gate", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if gate = 0, skip to tail"),
            row("derived", "automatic", "integer", responses="calc:lookup\nfield:price"),
            row("tail", "text", "text"),
        ]
        found = findings(rows, self.NULLS)
        self.assertEqual(len(found), 1)
        self.assertIn("'derived'", found[0].message)

    def test_a_jump_over_a_calculation_whose_inputs_were_also_skipped_is_silent(self):
        # The section was not asked, so its derived value being empty is
        # correct. This is what keeps the rule from firing on every long jump.
        rows = [
            row("gate", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if gate = 0, skip to tail"),
            numeric_row("price", lower="0", upper="100"),
            row("derived", "automatic", "integer", responses="calc:lookup\nfield:price"),
            row("tail", "text", "text"),
        ]
        self.assertEqual(findings(rows, self.NULLS), [])

    def test_ending_early_over_a_computable_calculation_warns_separately(self):
        rows = [
            numeric_row("price", lower="0", upper="100"),
            row("gate", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if gate = 0, skip to end"),
            row("derived", "automatic", "integer", responses="calc:lookup\nfield:price"),
            row("tail", "text", "text"),
        ]
        found = findings(rows, self.ENDS)
        self.assertEqual(len(found), 1)
        self.assertIn("skip to end", found[0].message)
        self.assertIn("trailing system variables still compute", found[0].message)
        self.assertEqual(findings(rows, self.NULLS), [])


class AlreadyReportedElsewhereTests(unittest.TestCase):
    """Cases `dd_validators` owns, which this module must not repeat.

    Two checks reporting the same cell is the failure `skip_parser`'s docstring
    exists to memorialise, in its other form: not two parsers disagreeing, but
    two messages for one mistake, so fixing the one the reader sees leaves the
    other still firing.
    """

    def test_a_backward_target_is_not_reported_here(self):
        rows = [
            row("first", "text", "text"),
            row("sex", "radio", "integer", responses="1:Male\n2:Female"),
            row("later", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if sex = 1, skip to first"),
        ]
        self.assertEqual(findings(rows), [])

    def test_a_nonexistent_target_is_not_reported_here(self):
        rows = [
            row("sex", "radio", "integer", responses="1:Male\n2:Female"),
            row("later", "radio", "integer", responses="1:Yes\n0:No",
                skip="postskip: if sex = 1, skip to nowhere"),
        ]
        self.assertEqual(findings(rows), [])

    def test_an_unparseable_cell_is_not_reported_here(self):
        rows = [
            row("sex", "radio", "integer", responses="1:Male\n2:Female"),
            row("later", "radio", "integer", responses="1:Yes\n0:No",
                skip="this is not a skip at all"),
        ]
        self.assertEqual(findings(rows), [])


class RuleCatalogueTests(unittest.TestCase):
    def test_every_rule_id_is_declared_with_a_severity_and_a_summary(self):
        self.assertTrue(skip_graph.RULES)
        for rule_id, rule in skip_graph.RULES.items():
            with self.subTest(rule=rule_id):
                self.assertEqual(rule.id, rule_id)
                self.assertIn(rule.severity, (Severity.ERROR, Severity.WARNING))
                self.assertTrue(rule.summary.strip())
                self.assertTrue(rule.summary.endswith("."))

    def test_a_finding_formats_as_this_tool_writes_every_other_message(self):
        # Against the same regex `test_error_message_shape` holds the rest of
        # the tool to. This module formats in one place rather than at every
        # call site, which is why it is not in that file's SCANNED_MODULES --
        # but the shape it produces still has to match.
        from tests.test_error_message_shape import PREFIX_RE

        found = findings(
            [
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
                row("pregnant", "radio", "integer", responses="1:Yes\n0:No",
                    skip="preskip: if sex = 3, skip to occupation"),
                row("occupation", "text", "text"),
            ]
        )
        self.assertTrue(found)
        self.assertTrue(found[0].format().startswith("ERROR - Skip graph: "))
        self.assertTrue(PREFIX_RE.match(found[0].format()))

    def test_no_message_says_in_table(self):
        # Same wording rule the rest of the tool follows: the dictionary's tabs
        # are worksheets, and `tablename` is a crfs column meaning something
        # else entirely.
        for message in messages(
            [
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
                row("pregnant", "radio", "integer", responses="1:Yes\n0:No",
                    skip="preskip: if sex = 3, skip to occupation"),
                row("occupation", "text", "text"),
            ]
        ):
            self.assertNotIn("in table", message)


class RowIndexTests(unittest.TestCase):
    def test_a_question_remembers_the_row_it_came_from(self):
        reader = read([row("sex", "radio", "integer", responses="1:M\n2:F")])
        # Row 1 is the header.
        self.assertEqual(reader.questionList[0].rowIndex, 2)


class BuildGateTests(unittest.TestCase):
    """That a finding reaches the build, not just `lint_form`.

    Every other test here calls the module directly, which proves the analysis
    and proves nothing about the wiring. A rule that fires perfectly into a log
    nobody reads, on a build that still ships a zip, is the same as no rule.
    """

    def _dictionary(self, directory, skip_cell):
        from openpyxl import Workbook

        from crf_reader import CRFS_COLUMN_NAMES

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "enrollee_dd"
        worksheet.append(HEADERS)
        worksheet.append(row("subjid", "automatic", "text", "Subject ID"))
        worksheet.append(row("sex", "radio", "integer", responses="1:Male\n2:Female"))
        worksheet.append(
            row("pregnant", "radio", "integer", responses="1:Yes\n0:No", skip=skip_cell)
        )
        worksheet.append(numeric_row("age"))

        crfs = workbook.create_sheet("crfs")
        crfs.append(list(CRFS_COLUMN_NAMES))
        crfs.append(
            [10, "enrollee", "Enrollee", "subjid", "", 1, "", "", "", "", "", "", "", "", ""]
        )

        path = directory / "dictionary.xlsx"
        workbook.save(path)
        workbook.close()
        return path

    def _run(self, skip_cell):
        import io
        from contextlib import redirect_stdout
        from pathlib import Path
        from tempfile import TemporaryDirectory

        from models import AppConfig
        from processor import SurveyGenProcessor

        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            processor = SurveyGenProcessor(
                AppConfig(
                    excelFile=str(self._dictionary(Path(tmp), skip_cell)),
                    csvFiles="",
                    outputPath=str(out),
                    surveyName="Test",
                    surveyId="test_survey",
                    databaseName="test.sqlite",
                )
            )
            with redirect_stdout(io.StringIO()):
                code = processor.run()
            return code, processor, sorted(p.name for p in out.iterdir())

    def test_a_graph_error_leaves_no_zip_and_no_loose_xml(self):
        code, processor, names = self._run("preskip: if sex = 3, skip to age")
        self.assertEqual(code, 1)
        self.assertEqual(names, ["gistlogfile.txt"])
        self.assertTrue(
            any("Skip graph" in line for line in processor.logstring),
            "\n".join(processor.logstring),
        )

    def test_a_clean_dictionary_still_builds(self):
        code, processor, names = self._run("preskip: if sex = 2, skip to age")
        self.assertEqual(code, 0, "\n".join(processor.logstring))
        self.assertEqual(names, ["gistlogfile.txt", "test_survey.zip"])
        self.assertFalse(any("Skip graph" in line for line in processor.logstring))

    def test_the_finding_is_logged_under_its_own_worksheet_heading(self):
        _, processor, _ = self._run("preskip: if sex = 3, skip to age")
        heading = "\rChecking worksheet: 'enrollee_dd'"
        self.assertIn(heading, processor.logstring)
        finding = next(i for i, l in enumerate(processor.logstring) if "Skip graph" in l)
        self.assertLess(processor.logstring.index(heading), finding)


if __name__ == "__main__":
    unittest.main()
