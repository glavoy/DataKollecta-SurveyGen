"""A merged Comments cell skips the whole row, so it has to be told apart.

`create_question_list` skips any row whose Comments cell is part of a merged
range. That is right for a section-heading banner and wrong for a question, and
it used to do both without a word in the log.

The discriminator is whether the merged range also covers column 1, FieldName:

  * A banner is one range spanning the full row, so column 1 is inside it and
    there is no FieldName to lose. **Every** merge in both live dictionaries is
    exactly that -- 28 of them, all single-row, all full width -- which is why
    this has to stay a silent skip. Turning merged cells into an error would
    refuse both real packages.
  * A range covering Comments but not column 1 leaves the row's FieldName and
    question intact and then drops them from the generated XML. That is the one
    worth reporting, and merging a Comments note down across two question rows
    is the ordinary way to cause it.
"""

import unittest

from openpyxl import Workbook

from excel_reader import ExcelReader
from tests.test_dd_validations import HEADERS, row


COMMENTS_COL = "N"   # NUMBER_OF_COLUMNS = 14
LAST_COL = "N"


def read_with_merges(rows, merges):
    """Rows appended under the standard header, then `merges` applied.

    Row 1 is the header, so dictionary row *n* is worksheet row *n + 1*.
    """
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "demo_dd"
    worksheet.append(HEADERS)
    for r in rows:
        worksheet.append(r)
    for merge in merges:
        worksheet.merge_cells(merge)
    reader = ExcelReader()
    questions = reader.create_question_list(worksheet)
    return reader, questions


def field_names(questions):
    return [q.fieldName for q in questions]


def merge_errors(reader):
    return [line for line in reader.logstring if "Merged cell" in line]


class FullWidthBannerTests(unittest.TestCase):
    """The shape both real dictionaries actually use. Must stay silent."""

    def test_a_full_width_banner_row_is_skipped_without_an_error(self):
        reader, questions = read_with_merges(
            [
                row("age", "text", "text_integer", maxchars="3", lower="0", upper="120"),
                ["SECTION B: HOUSEHOLD"] + [""] * 13,
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
            ],
            [f"A3:{LAST_COL}3"],
        )

        self.assertEqual(merge_errors(reader), [])
        self.assertEqual(field_names(questions), ["age", "sex"])

    def test_a_multi_row_full_width_banner_is_also_skipped_silently(self):
        # Not present in either dictionary today, but it is still a banner --
        # column 1 is inside the merge, so no question is being lost. Reporting
        # it would be a false positive.
        reader, questions = read_with_merges(
            [
                row("age", "text", "text_integer", maxchars="3", lower="0", upper="120"),
                ["SECTION B: HOUSEHOLD"] + [""] * 13,
                [""] * 14,
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
            ],
            [f"A3:{LAST_COL}4"],
        )

        self.assertEqual(merge_errors(reader), [])
        self.assertEqual(field_names(questions), ["age", "sex"])


class SwallowedQuestionTests(unittest.TestCase):
    """The hazard: a merge that keeps the FieldName but loses the question."""

    def test_a_comments_note_merged_across_two_question_rows_is_reported(self):
        reader, questions = read_with_merges(
            [
                row("age", "text", "text_integer", maxchars="3", lower="0", upper="120"),
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
            ],
            [f"{COMMENTS_COL}2:{COMMENTS_COL}3"],
        )

        errors = merge_errors(reader)
        self.assertEqual(len(errors), 2, reader.logstring)
        self.assertTrue(reader.errorsEncountered)
        # Both rows name themselves, so the author knows where to look.
        self.assertIn("Row 2", errors[0])
        self.assertIn("Row 3", errors[1])
        self.assertIn("demo_dd", errors[0])
        self.assertIn("N2:N3", errors[0])

    def test_the_swallowed_questions_are_absent_from_the_question_list(self):
        # The point of the error: without it, these two simply vanish and the
        # run still reports success.
        _, questions = read_with_merges(
            [
                row("age", "text", "text_integer", maxchars="3", lower="0", upper="120"),
                row("sex", "radio", "integer", responses="1:Male\n2:Female"),
            ],
            [f"{COMMENTS_COL}2:{COMMENTS_COL}3"],
        )

        self.assertEqual(field_names(questions), [])

    def test_a_partial_merge_that_stops_short_of_column_one_is_reported(self):
        # D:N on a single row -- QuestionText through Comments merged away,
        # FieldName still present. The question is unusable, not a heading.
        reader, _ = read_with_merges(
            [row("age", "text", "text_integer", maxchars="3", lower="0", upper="120")],
            [f"D2:{LAST_COL}2"],
        )

        self.assertEqual(len(merge_errors(reader)), 1, reader.logstring)


class UnmergedRowsTests(unittest.TestCase):
    def test_a_sheet_with_no_merges_reports_nothing(self):
        reader, questions = read_with_merges(
            [row("age", "text", "text_integer", maxchars="3", lower="0", upper="120")],
            [],
        )

        self.assertEqual(merge_errors(reader), [])
        self.assertEqual(field_names(questions), ["age"])

    def test_a_merge_that_misses_the_comments_column_does_not_skip_the_row(self):
        # Only a merge over the Comments column triggers the skip at all.
        reader, questions = read_with_merges(
            [row("age", "text", "text_integer", maxchars="3", lower="0", upper="120")],
            ["I2:J2"],
        )

        self.assertEqual(merge_errors(reader), [])
        self.assertEqual(field_names(questions), ["age"])


if __name__ == "__main__":
    unittest.main()
