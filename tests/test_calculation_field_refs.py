"""A calculation may only read fields the form declares.

The app reads a nonexistent field as empty text and a `case` comparing empty
text just does not match, so a typo (or a field that was never added) does not
fail -- the rule silently never fires. A Burkina Faso eligibility rule tested
`age_at_sep2023`, which no row defined, and shipped doing nothing.
"""

import unittest

from tests.test_dd_validations import errors, read, row


def auto(name, responses, ftype="integer"):
    return row(name, "automatic", ftype, responses=responses)


DOB = row("dob", "date", "date", lower="2020-01-01", upper="-5m")


class CalculationFieldExistsTests(unittest.TestCase):
    def assertOneMissing(self, reader, name):
        found = [e for e in errors(reader) if "nonexistent FieldName" in e]
        self.assertEqual(len(found), 1, "\n".join(reader.logstring))
        self.assertIn(f": {name}.", found[0])

    def test_case_when_field_that_does_not_exist_is_an_error(self):
        reader = read([
            DOB,
            auto("months", "calc:age_at_date\nfield:dob\nvalue:months\nseparator:2025-03-31"),
            auto("in_range", "calc:case\nwhen:months < 5 => 0\nwhen:age_at_sep2023 >= 12 => 0\nelse:1"),
        ])
        self.assertOneMissing(reader, "age_at_sep2023")

    def test_case_when_field_that_exists_is_fine(self):
        reader = read([
            DOB,
            auto("months", "calc:age_at_date\nfield:dob\nvalue:months\nseparator:2025-03-31"),
            auto("in_range", "calc:case\nwhen:months < 5 => 0\nelse:1"),
        ])
        self.assertEqual(errors(reader), [], "\n".join(reader.logstring))

    def test_age_at_date_field_that_does_not_exist_is_an_error(self):
        reader = read([auto("months", "calc:age_at_date\nfield:birthdate\nvalue:months\nseparator:2025-03-31")])
        self.assertOneMissing(reader, "birthdate")

    def test_startdate_placeholder_and_today_are_allowed(self):
        reader = read([
            DOB,
            auto("age_y", "calc:age_at_date\nfield:dob\nvalue:years\nseparator:[[startdate]]"),
            auto("days", "calc:date_diff\nfield:dob\nvalue:today\nunit:d"),
            auto("yr", "calc:date_part\nfield:today\nunit:yyyy"),
        ])
        self.assertEqual(errors(reader), [], "\n".join(reader.logstring))

    def test_date_diff_end_field_that_does_not_exist_is_an_error(self):
        reader = read([DOB, auto("days", "calc:date_diff\nfield:dob\nvalue:visitdate\nunit:d")])
        self.assertOneMissing(reader, "visitdate")

    def test_placeholder_in_a_constant_that_does_not_exist_is_an_error(self):
        reader = read([auto("label", "calc:constant\nvalue:[[nosuchfield]]", ftype="text")])
        self.assertOneMissing(reader, "nosuchfield")

    def test_query_param_field_that_does_not_exist_is_an_error(self):
        reader = read([
            auto("total", "calc:query\nsql:SELECT count(*) FROM t WHERE k = @k\nparam:@k = missingkey"),
        ])
        self.assertOneMissing(reader, "missingkey")

    def test_lookup_of_a_field_that_does_not_exist_is_an_error(self):
        reader = read([auto("copy", "calc:lookup\nfield:origin", ftype="text")])
        self.assertOneMissing(reader, "origin")

    def test_a_field_that_comes_after_the_calculation_is_an_error(self):
        reader = read([
            auto("months", "calc:age_at_date\nfield:dob\nvalue:months\nseparator:2025-03-31"),
            DOB,
        ])
        found = [e for e in errors(reader) if "AFTER the current question: dob" in e]
        self.assertEqual(len(found), 1, "\n".join(reader.logstring))

    def test_a_calculation_reading_itself_is_an_error(self):
        reader = read([auto("total", "calc:case\nwhen:total = 1 => 1\nelse:0")])
        found = [e for e in errors(reader) if "uses itself: total" in e]
        self.assertEqual(len(found), 1, "\n".join(reader.logstring))

    def test_a_field_defined_just_above_is_fine(self):
        reader = read([DOB, auto("months", "calc:age_at_date\nfield:dob\nvalue:months\nseparator:2025-03-31")])
        self.assertEqual(errors(reader), [], "\n".join(reader.logstring))


if __name__ == "__main__":
    unittest.main()
