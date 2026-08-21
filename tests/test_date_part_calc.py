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


def date_part_row(fieldname, field, unit):
    return row(
        fieldname,
        "automatic",
        "text",
        responses=f"calc:date_part\nfield:{field}\nunit:{unit}",
    )


class DatePartCalculationTests(unittest.TestCase):
    """`calc:date_part` extracts a component from a named date field --
    distinct from the yyyy/yy/mm/dd/doy Computed Automatic Variables, which
    only ever mean "today" and take no calc: block at all."""

    def test_a_valid_declaration_parses_without_error(self):
        reader = read([
            row("dob", "date", "date", lower="1900-01-01", upper="2050-12-31"),
            date_part_row("birthmonth", "dob", "mm"),
        ])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_every_unit_is_accepted(self):
        for unit in ("yyyy", "yy", "mm", "dd", "doy"):
            with self.subTest(unit=unit):
                reader = read([
                    row("dob", "date", "date", lower="1900-01-01", upper="2050-12-31"),
                    date_part_row("result", "dob", unit),
                ])
                self.assertFalse(
                    reader.errorsEncountered, "\n".join(reader.logstring)
                )

    def test_unit_is_accepted_case_insensitively(self):
        reader = read([
            row("dob", "date", "date", lower="1900-01-01", upper="2050-12-31"),
            date_part_row("result", "dob", "MM"),
        ])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_field_today_is_accepted(self):
        reader = read([date_part_row("current_month", "today", "mm")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_missing_field_is_an_error(self):
        reader = read([
            row("result", "automatic", "text", responses="calc:date_part\nunit:mm"),
        ])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("missing required 'field'", errors(reader)[0])

    def test_missing_unit_is_an_error(self):
        reader = read([
            row("dob", "date", "date", lower="1900-01-01", upper="2050-12-31"),
            row("result", "automatic", "text", responses="calc:date_part\nfield:dob"),
        ])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("missing required 'unit'", errors(reader)[0])

    def test_an_invalid_unit_is_an_error_naming_the_allowed_set(self):
        reader = read([
            row("dob", "date", "date", lower="1900-01-01", upper="2050-12-31"),
            date_part_row("result", "dob", "hh"),
        ])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("Invalid date_part unit", errors(reader)[0])
        self.assertIn("'yyyy', 'yy', 'mm', 'dd', or 'doy'", errors(reader)[0])


if __name__ == "__main__":
    unittest.main()
