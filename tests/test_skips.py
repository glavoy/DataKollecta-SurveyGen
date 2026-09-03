"""Skip parsing and generation.

`_generate_skip` had no tests at all before this file, which is how a
capitalized `Preskip:` came to pass every validation check and then be dropped
from the XML entirely -- the validator and the generator each parsed the cell
their own way and disagreed about the prefix.
"""

import re
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from excel_reader import ExcelReader
from skip_parser import parse_skip
from xml_generator import XmlGenerator

from tests.test_dd_validations import HEADERS, numeric_row, read

from openpyxl import Workbook


def skip_elements(xml):
    """Every generated <skip>, as (section, fieldname, condition, response, target)."""
    found = []
    for section in ("preskip", "postskip"):
        block = re.search(rf"<{section}>(.*?)</{section}>", xml, re.DOTALL)
        if not block:
            continue
        for m in re.finditer(
            r"<skip fieldname='([^']*)' condition = '([^']*)' response='([^']*)' "
            r"response_type='fixed' skiptofieldname ='([^']*)'>",
            block.group(1),
        ):
            found.append((section, *m.groups()))
    return found


def generate(rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "demo_dd"
    worksheet.append(HEADERS)
    for r in rows:
        worksheet.append(r)

    reader = ExcelReader()
    questions = reader.create_question_list(worksheet)
    with TemporaryDirectory() as tmp:
        path = XmlGenerator().write_xml("demo_dd", questions, Path(tmp))
        return path.read_text(encoding="utf-8"), reader


def two_questions(skip):
    """A dictionary carrying `skip` on `age`, testing `obs` and targeting `occupation`.

    The tested field has to be an EARLIER one: a preskip on its own field is
    itself an error (`_check_preskip_does_not_test_itself`), which would mask
    what these tests are checking.
    """
    return [
        numeric_row("obs"),
        numeric_row("age", skip=skip),
        numeric_row("occupation"),
    ]


class SkipPrefixTests(unittest.TestCase):
    """The prefix decides which section the rule lands in, so it must survive."""

    def test_a_lowercase_preskip_is_generated(self):
        xml, reader = generate(two_questions("preskip: if obs = 0, skip to occupation"))
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(
            [("preskip", "obs", "=", "0", "occupation")], skip_elements(xml)
        )

    def test_a_capitalized_preskip_is_still_generated(self):
        """Excel's AutoCorrect capitalizes the first letter of a cell by default.

        This used to pass validation and then match neither the generator's
        `preskip` nor its `postskip` test, so no element was written at all and
        the log still said "No errors found".
        """
        xml, reader = generate(two_questions("Preskip: if obs = 0, skip to occupation"))
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(
            [("preskip", "obs", "=", "0", "occupation")], skip_elements(xml)
        )

    def test_a_capitalized_postskip_is_still_generated(self):
        xml, reader = generate(two_questions("PostSkip: if obs = 0, skip to occupation"))
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(
            [("postskip", "obs", "=", "0", "occupation")], skip_elements(xml)
        )

    def test_a_misspelled_prefix_is_a_reported_error(self):
        """It must never be silently dropped, which is what used to happen.

        Only the reader runs here: the processor gates generation on
        `errorsEncountered`, and the generator now raises rather than emit a
        questionnaire missing a rule -- see `test_the_generator_refuses_...`.
        """
        reader = read(two_questions("postkip: if obs = 0, skip to occupation"))
        self.assertTrue(reader.errorsEncountered)
        self.assertIn("invalid syntax for Skip", "\n".join(reader.logstring))

    def test_the_generator_refuses_an_unparseable_skip(self):
        """Belt and braces: if validation is ever bypassed, fail loudly."""
        reader = ExcelReader()
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "demo_dd"
        worksheet.append(HEADERS)
        for r in two_questions("postkip: if obs = 0, skip to occupation"):
            worksheet.append(r)
        questions = reader.create_question_list(worksheet)
        with TemporaryDirectory() as tmp:
            with self.assertRaises(ValueError):
                XmlGenerator().write_xml("demo_dd", questions, Path(tmp))


class SkipTargetSpellingTests(unittest.TestCase):
    """Three spellings of the target clause are in live use; all must work.

    The old generator took everything after the last space, so all three
    parsed. Rejecting any of them breaks a real dictionary.
    """

    def test_skip_to_target(self):
        xml, _ = generate(two_questions("preskip: if obs = 0, skip to occupation"))
        self.assertEqual("occupation", skip_elements(xml)[0][4])

    def test_bare_target(self):
        """AVERT (fr) writes `, vx_dose_summary` with no `skip to`."""
        xml, _ = generate(two_questions("preskip: if obs = 0, occupation"))
        self.assertEqual("occupation", skip_elements(xml)[0][4])

    def test_then_skip_to_target(self):
        """PRISM CSS writes `, then skip to everreceivesmc`."""
        xml, _ = generate(two_questions("preskip: if obs = 0, then skip to occupation"))
        self.assertEqual("occupation", skip_elements(xml)[0][4])

    def test_unknown_filler_before_the_target_is_an_error(self):
        reader = read(two_questions("preskip: if obs = 0, goto occupation"))
        self.assertTrue(reader.errorsEncountered)


class SkipOperatorTests(unittest.TestCase):
    def test_comparison_operators_are_entity_encoded(self):
        for written, expected in [("<", "&lt;"), (">", "&gt;"), ("<>", "&lt;&gt;")]:
            with self.subTest(written):
                xml, reader = generate(
                    two_questions(f"preskip: if obs {written} 5, skip to occupation")
                )
                self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
                self.assertEqual(expected, skip_elements(xml)[0][2])

    def test_quoted_and_bare_contains_land_on_one_spelling(self):
        for written in ["'contains'", "contains"]:
            with self.subTest(written):
                xml, _ = generate(
                    two_questions(f"postskip: if obs {written} 5, skip to occupation")
                )
                self.assertEqual("contains", skip_elements(xml)[0][2])

    def test_does_not_contain(self):
        xml, _ = generate(
            two_questions("postskip: if obs 'does not contain' 5, skip to occupation")
        )
        self.assertEqual("does not contain", skip_elements(xml)[0][2])


class SkipParserUnitTests(unittest.TestCase):
    """The parser alone, with no workbook."""

    def test_a_multi_word_value_is_rejected(self):
        self.assertIsNone(parse_skip("preskip: if region = North West, skip to x"))

    def test_a_missing_target_is_rejected(self):
        self.assertIsNone(parse_skip("preskip: if obs = 0, skip to"))

    def test_a_missing_comma_is_rejected(self):
        self.assertIsNone(parse_skip("preskip: if obs = 0 skip to occupation"))

    def test_the_operator_is_stored_unescaped(self):
        """Escaping belongs at emission, so the model cannot be double-escaped."""
        self.assertEqual("<", parse_skip("preskip: if obs < 5, skip to x").operator)

    def test_end_is_accepted_as_a_target(self):
        self.assertEqual("end", parse_skip("postskip: if obs = 0, skip to end").target)

    def test_extra_internal_whitespace_is_tolerated(self):
        parsed = parse_skip("preskip:  if   obs  =  0 ,  skip  to  occupation")
        self.assertIsNotNone(parsed)
        self.assertEqual(("obs", "=", "0", "occupation"),
                         (parsed.field, parsed.operator, parsed.value, parsed.target))


class MultipleSkipsTests(unittest.TestCase):
    def test_a_cell_can_hold_a_preskip_and_a_postskip(self):
        xml, reader = generate(
            two_questions(
                "preskip: if obs = 0, skip to occupation\n"
                "postskip: if obs = 1, skip to occupation"
            )
        )
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(
            [
                ("preskip", "obs", "=", "0", "occupation"),
                ("postskip", "obs", "=", "1", "occupation"),
            ],
            skip_elements(xml),
        )


if __name__ == "__main__":
    unittest.main()
