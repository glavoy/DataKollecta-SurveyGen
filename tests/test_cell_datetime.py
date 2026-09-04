"""A date-formatted cell must read as the date the author typed.

openpyxl hands back a `datetime` for any date-formatted cell, so `str()` gave
"2026-07-13 00:00:00" where the dictionary says `yyyy-mm-dd`. In a range that
produced a confusing but *visible* validation error; everywhere else --
QuestionText, a response code, a constant calculation value -- nothing checks
the format and the padded value shipped to the device.

**Neither real dictionary contains a datetime cell**, so the byte-diff harness
cannot see this change at all. These synthetic fixtures are the only evidence
there is, which is why they pin the chosen format explicitly: whatever they
assert becomes generated output the first time a dictionary date-formats a
cell.
"""

import unittest
from datetime import date, datetime, time

from cell_text import to_str
from tests.test_dd_validations import HEADERS, read, row, errors


class ToStrDateTests(unittest.TestCase):
    def test_a_date_formatted_cell_reads_as_a_plain_date(self):
        # The documented format, and exactly what HARDCODED_DATE_RE matches.
        self.assertEqual(to_str(datetime(2026, 7, 13)), "2026-07-13")

    def test_a_bare_date_reads_the_same_way(self):
        self.assertEqual(to_str(date(2026, 7, 13)), "2026-07-13")

    def test_a_datetime_with_a_time_keeps_the_time(self):
        # No dictionary column is documented to hold one, so there is no intent
        # to honour -- keep the value visible rather than silently truncating.
        self.assertEqual(
            to_str(datetime(2026, 7, 13, 14, 30, 5)), "2026-07-13 14:30:05"
        )

    def test_microseconds_are_dropped_so_the_output_is_stable(self):
        self.assertEqual(
            to_str(datetime(2026, 7, 13, 14, 30, 5, 123456)), "2026-07-13 14:30:05"
        )

    def test_a_time_only_cell_is_left_alone(self):
        self.assertEqual(to_str(time(14, 30)), "14:30:00")

    def test_the_float_case_still_works(self):
        # The behaviour this function existed for before dates were handled.
        self.assertEqual(to_str(1.0), "1")
        self.assertEqual(to_str(3.5), "3.5")
        self.assertEqual(to_str(None), "")


class DateInARangeTests(unittest.TestCase):
    """The loud path: a date range bound used to fail its own format check."""

    def test_a_date_formatted_range_bound_is_now_accepted(self):
        reader = read([
            row("visit", "date", "date",
                lower=datetime(2026, 1, 1), upper=datetime(2026, 12, 31)),
        ])

        self.assertEqual(errors(reader), [])

    def test_the_bounds_reach_the_question_as_dates(self):
        reader = read([
            row("visit", "date", "date",
                lower=datetime(2026, 1, 1), upper=datetime(2026, 12, 31)),
        ])
        question = next(q for q in reader.questionList if q.fieldName == "visit")

        self.assertEqual(question.lowerRange, "2026-01-01")
        self.assertEqual(question.upperRange, "2026-12-31")


class DateInQuestionTextTests(unittest.TestCase):
    """The silent path, and the reason this belongs in `to_str`."""

    def test_a_date_formatted_question_text_loses_its_midnight(self):
        reader = read([
            row("visit", "date", "date", text=datetime(2026, 7, 13)),
        ])
        question = next(q for q in reader.questionList if q.fieldName == "visit")

        # Nothing validates QuestionText's format, so before this change the
        # interviewer was shown "2026-07-13 00:00:00".
        self.assertEqual(question.questionText, "2026-07-13")


if __name__ == "__main__":
    unittest.main()
