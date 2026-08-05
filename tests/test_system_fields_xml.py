import re
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from excel_reader import ExcelReader
from models import LEADING_SYSTEM_FIELDS, TRAILING_SYSTEM_FIELDS
from xml_generator import XmlGenerator

from tests.test_dd_validations import HEADERS, row

from openpyxl import Workbook


LEADING = [name for name, _ in LEADING_SYSTEM_FIELDS]
TRAILING = [name for name, _ in TRAILING_SYSTEM_FIELDS]


def generate(rows, supplied=None):
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
        path = XmlGenerator().write_xml("demo_dd", questions, Path(tmp))
        return path.read_text(encoding="utf-8"), reader


def field_order(xml):
    return re.findall(r"fieldname = '([a-z_0-9]+)'", xml)


class SystemFieldsAreWrittenTests(unittest.TestCase):
    """The generator supplies the system variables, so a dictionary never has to."""

    def test_they_are_written_even_when_not_declared(self):
        xml, _ = generate([row("age", "text", "text_integer", maxchars="3")])
        names = field_order(xml)

        for name in LEADING + TRAILING:
            self.assertIn(name, names)

    def test_leading_fields_come_before_the_first_real_question(self):
        xml, _ = generate([row("age", "text", "text_integer", maxchars="3")])
        names = field_order(xml)

        self.assertEqual(names[: len(LEADING)], LEADING)
        self.assertLess(names.index("startdate"), names.index("age"))

    def test_trailing_fields_sit_between_the_last_question_and_the_final_screen(self):
        # Navigation stops on the end-of-survey screen, so anything after it is
        # never computed and would be saved empty.
        xml, _ = generate([row("age", "text", "text_integer", maxchars="3")])
        names = field_order(xml)

        self.assertEqual(names[-1], "end_of_questions")
        self.assertEqual(names[-1 - len(TRAILING) : -1], TRAILING)
        self.assertLess(names.index("age"), names.index("uniqueid"))

    def test_they_are_written_with_the_expected_types(self):
        xml, _ = generate([row("age", "text", "text_integer", maxchars="3")])

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
                row("age", "text", "text_integer", maxchars="3"),
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
                row("age", "text", "text_integer", maxchars="3"),
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
            + [row("age", "text", "text_integer", maxchars="3")]
            + [row(name, "automatic", "text") for name in TRAILING]
        )
        omitted, _ = generate([row("age", "text", "text_integer", maxchars="3")])

        self.assertEqual(declared, omitted)

    def test_declaring_one_still_only_warns(self):
        _, reader = generate(
            [
                row("starttime", "automatic", "datetime"),
                row("age", "text", "text_integer", maxchars="3"),
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
                row("age", "text", "text_integer", maxchars="3"),
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
                row("age", "text", "text_integer", maxchars="3"),
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
