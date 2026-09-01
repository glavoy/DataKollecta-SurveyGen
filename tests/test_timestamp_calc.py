import unittest

from openpyxl import Workbook

from excel_reader import ExcelReader
from tests.test_system_fields_xml import generate


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


def read(rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "demo_dd"
    worksheet.append(HEADERS)
    for r in rows:
        worksheet.append(r)
    reader = ExcelReader()
    reader.create_question_list(worksheet)
    return reader


def errors(reader):
    return [line for line in reader.logstring if line.startswith("ERROR")]


def timestamp_row(fieldname):
    return row(fieldname, "automatic", "datetime", responses="calc:timestamp")


class TimestampCalculationTests(unittest.TestCase):
    """`calc:timestamp` stamps the current date-and-time the moment its row
    is reached -- the explicit, always-required replacement for the old
    "leave Responses blank on a datetime automatic field" convention."""

    def test_a_valid_declaration_parses_without_error(self):
        reader = read([timestamp_row("time_eligible")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_generated_xml_carries_the_preserve_flag(self):
        xml, _ = generate([timestamp_row("time_eligible")])

        self.assertIn("<calculation type='timestamp' preserve='true' />", xml)

    def test_a_non_datetime_field_type_is_an_error(self):
        reader = read([row("time_eligible", "automatic", "text", responses="calc:timestamp")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("requires FieldType 'datetime'", errors(reader)[0])

    def test_a_bare_datetime_automatic_field_with_no_calc_is_now_an_error(self):
        # The old implicit fallback (blank Responses + FieldType=datetime)
        # is gone -- calc:timestamp is the only way to get this behavior now.
        reader = read([row("time_eligible", "automatic", "datetime")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("has no calculation", errors(reader)[0])


if __name__ == "__main__":
    unittest.main()
