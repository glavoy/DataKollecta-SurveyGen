import re
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from excel_reader import ExcelReader
from models import (
    LEADING_SYSTEM_FIELDS,
    PARENT_LINK_FIELD,
    RESERVED_SYSTEM_FIELDS,
    TRAILING_SYSTEM_FIELDS,
)
from xml_generator import XmlGenerator

from tests.test_dd_validations import HEADERS, numeric_row, row

from openpyxl import Workbook


LEADING = [name for name, _ in LEADING_SYSTEM_FIELDS]
TRAILING = [name for name, _ in TRAILING_SYSTEM_FIELDS]


def generate(rows, supplied=None, has_parent=False):
    """Run a set of dictionary rows through the reader and the generator.

    `supplied` mirrors the crfs sheet's linking/increment/primary-key fields,
    which are automatic but filled in from the manifest rather than by a
    calculation.
    """
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "demo_dd"
    worksheet.append(HEADERS)
    for r in rows:
        worksheet.append(r)

    reader = ExcelReader(supplied_auto_fields=supplied)
    questions = reader.create_question_list(worksheet)
    assert not reader.errorsEncountered, "\n".join(reader.logstring)

    with TemporaryDirectory() as tmp:
        path = XmlGenerator().write_xml(
            "demo_dd", questions, Path(tmp), has_parent=has_parent
        )
        return path.read_text(encoding="utf-8"), reader


def field_order(xml):
    return re.findall(r"fieldname = '([a-z_0-9]+)'", xml)


class SystemFieldsAreWrittenTests(unittest.TestCase):
    """The generator supplies the system variables, so a dictionary never has to."""

    def test_they_are_written_even_when_not_declared(self):
        xml, _ = generate([numeric_row()])
        names = field_order(xml)

        for name in LEADING + TRAILING:
            self.assertIn(name, names)

    def test_leading_fields_come_before_the_first_real_question(self):
        xml, _ = generate([numeric_row()])
        names = field_order(xml)

        self.assertEqual(names[: len(LEADING)], LEADING)
        self.assertLess(names.index("startdate"), names.index("age"))

    def test_trailing_fields_sit_between_the_last_question_and_the_final_screen(self):
        # Navigation stops on the end-of-survey screen, so anything after it is
        # never computed and would be saved empty.
        xml, _ = generate([numeric_row()])
        names = field_order(xml)

        self.assertEqual(names[-1], "end_of_questions")
        self.assertEqual(names[-1 - len(TRAILING) : -1], TRAILING)
        self.assertLess(names.index("age"), names.index("uniqueid"))

    def test_they_are_written_with_the_expected_types(self):
        xml, _ = generate([numeric_row()])

        for name, fieldtype in LEADING_SYSTEM_FIELDS + TRAILING_SYSTEM_FIELDS:
            self.assertIn(
                f"<question type = 'automatic' fieldname = '{name}' "
                f"fieldtype = '{fieldtype}'>",
                xml,
            )


class DeclaredSystemFieldsTests(unittest.TestCase):
    """A declared row is dropped, so the questionnaire cannot get two."""

    def test_a_declared_field_appears_exactly_once(self):
        xml, _ = generate(
            [
                row("starttime", "automatic", "datetime"),
                numeric_row(),
                row("stoptime", "automatic", "datetime"),
            ]
        )
        names = field_order(xml)

        self.assertEqual(names.count("starttime"), 1)
        self.assertEqual(names.count("stoptime"), 1)

    def test_a_misplaced_field_is_written_in_the_right_place(self):
        # stoptime declared in the middle would otherwise record a time partway
        # through the interview.
        xml, _ = generate(
            [
                numeric_row(),
                row("stoptime", "automatic", "datetime"),
                row("weight", "text", "text_integer", maxchars="3"),
            ]
        )
        names = field_order(xml)

        self.assertGreater(names.index("stoptime"), names.index("weight"))
        self.assertEqual(names[-1], "end_of_questions")

    def test_declaring_them_all_gives_the_same_file_as_declaring_none(self):
        declared, _ = generate(
            [row(name, "automatic", "text") for name in LEADING]
            + [numeric_row()]
            + [row(name, "automatic", "text") for name in TRAILING]
        )
        omitted, _ = generate([numeric_row()])

        self.assertEqual(declared, omitted)

    def test_declaring_one_still_only_warns(self):
        _, reader = generate(
            [
                row("starttime", "automatic", "datetime"),
                numeric_row(),
            ]
        )

        self.assertFalse(reader.errorsEncountered)
        warnings = [l for l in reader.logstring if l.startswith("WARNING")]
        self.assertEqual(len(warnings), 1)
        self.assertIn("this row is ignored", warnings[0])


class OtherAutomaticFieldsAreUntouchedTests(unittest.TestCase):
    """Linking and increment fields are not reserved and must keep their place."""

    def test_crf_fields_keep_their_declared_position(self):
        # hhid and netnum come from the crfs linking/increment columns and are
        # declared in the dictionary; they must not be moved or dropped.
        xml, _ = generate(
            [
                row("hhid", "automatic", "integer"),
                row("netnum", "automatic", "integer"),
                numeric_row(),
            ],
            supplied={"hhid", "netnum"},
        )
        names = field_order(xml)

        self.assertEqual(
            names,
            LEADING + ["hhid", "netnum", "age"] + TRAILING + ["end_of_questions"],
        )

    def test_a_calculated_field_keeps_its_position(self):
        xml, _ = generate(
            [
                numeric_row(),
                row("age_group", "calculated", "integer",
                    responses="calc:constant\nvalue:1"),
            ]
        )
        names = field_order(xml)

        self.assertEqual(
            names,
            LEADING + ["age", "age_group"] + TRAILING + ["end_of_questions"],
        )


if __name__ == "__main__":
    unittest.main()


class ParentLinkFieldTests(unittest.TestCase):
    """`parent_uniqueid` ties a child to its parent by a value nothing edits.

    The business key (hhid) is built from typed answers, so an interviewer
    correcting a mistyped household number changes it. A UUID cannot be
    retyped, so a join on it cannot drift.
    """

    def test_it_is_written_on_a_form_with_a_parent(self):
        xml, _ = generate([numeric_row()], has_parent=True)

        self.assertIn(PARENT_LINK_FIELD[0], field_order(xml))

    def test_it_is_absent_from_a_form_with_no_parent(self):
        # A base form has nothing to point at, so the column would only ever
        # be null -- and a null join key is worse than none, because analysis
        # then has to fall back to the business key anyway.
        xml, _ = generate([numeric_row()], has_parent=False)

        self.assertNotIn(PARENT_LINK_FIELD[0], field_order(xml))

    def test_it_comes_after_uniqueid(self):
        # Nothing consumes it, so position carries no meaning -- but sitting
        # with the record's own uniqueid is what makes the pair legible.
        names = field_order(generate([numeric_row()], has_parent=True)[0])

        self.assertGreater(
            names.index(PARENT_LINK_FIELD[0]), names.index("uniqueid")
        )

    def test_it_stays_ahead_of_the_end_of_survey_screen(self):
        # Navigation stops on that screen, so anything after it is never
        # reached and would be saved empty.
        names = field_order(generate([numeric_row()], has_parent=True)[0])

        self.assertLess(
            names.index(PARENT_LINK_FIELD[0]), names.index("end_of_questions")
        )

    def test_a_declared_row_is_dropped_rather_than_duplicated(self):
        # It is reserved, so a dictionary that declares it anyway must not end
        # up with two of them -- the same treatment uniqueid and swver get.
        self.assertIn(PARENT_LINK_FIELD[0], RESERVED_SYSTEM_FIELDS)

        xml, reader = generate(
            [row(PARENT_LINK_FIELD[0], "automatic", "text", "Parent"), numeric_row()],
            has_parent=True,
        )

        self.assertEqual(field_order(xml).count(PARENT_LINK_FIELD[0]), 1)
        self.assertTrue(
            any("reserved" in line.lower() for line in reader.logstring),
            "\n".join(reader.logstring),
        )

    def test_a_declared_row_needs_no_calculation(self):
        # Being reserved exempts it from "an automatic field with no
        # calculation is never given a value" -- the app supplies it.
        _, reader = generate(
            [row(PARENT_LINK_FIELD[0], "automatic", "text", "Parent"), numeric_row()],
            has_parent=True,
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
