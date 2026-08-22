"""What a typed (QuestionType 'text') question is allowed to declare.

Three rules, each closing a hole that let a plausible-looking dictionary build
a package that misbehaves only in the field.
"""

import unittest

from tests.test_dd_validations import errors, numeric_row, read, row, warnings


class TypedFieldTypeTests(unittest.TestCase):
    """Rule 1 — a text question stores text, a whole number, a decimal or a time."""

    def test_the_four_typed_field_types_are_accepted(self):
        reader = read(
            [
                row("notes", "text", "text", maxchars="80"),
                numeric_row("age", "text_integer"),
                numeric_row("weight", "text_decimal", maxchars="5"),
                row("timeseen", "text", "hourmin", maxchars="=5"),
            ]
        )

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_integer_on_a_text_question_is_an_error(self):
        # 'integer' is what a radio stores. Allowed here it also slipped past
        # the MaxCharacters requirement, which named 'text_integer' only.
        reader = read([row("agemonths", "text", "integer", maxchars="2")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("cannot be used with the QuestionType 'text'", errors(reader)[0])

    def test_a_date_field_type_on_a_text_question_is_an_error(self):
        reader = read([row("visitdate", "text", "date", maxchars="10")])

        self.assertTrue(reader.errorsEncountered)

    def test_radio_still_requires_integer(self):
        reader = read([row("sex", "radio", "integer", responses="1:Male\n2:Female")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_retired_field_types_are_rejected(self):
        # phone_num and text_id behaved exactly as 'text' in the app and were
        # used by no dictionary.
        for retired in ["phone_num", "text_id"]:
            reader = read([row("f", "text", retired, maxchars="10")])

            self.assertTrue(reader.errorsEncountered, retired)
            self.assertIn("not among the predefined list", errors(reader)[0])


class RequiredMaxCharactersTests(unittest.TestCase):
    """Rule 2 — every typed answer needs a length limit."""

    def test_a_missing_value_is_an_error_for_every_typed_field_type(self):
        for fieldtype in ["text", "text_integer", "text_decimal"]:
            reader = read([row("f", "text", fieldtype, lower="0", upper="9")])

            self.assertTrue(reader.errorsEncountered, fieldtype)
            self.assertIn("needs a value", errors(reader)[0])

    def test_a_selection_question_still_needs_nothing(self):
        reader = read([row("sex", "radio", "integer", responses="1:Male\n2:Female")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))

    def test_hourmin_must_be_five_characters(self):
        reader = read([row("timeseen", "text", "hourmin", maxchars="5")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("a time is always hh:mm", errors(reader)[0])

    def test_hourmin_without_a_length_is_an_error(self):
        reader = read([row("timeseen", "text", "hourmin")])

        self.assertTrue(reader.errorsEncountered)


class RangePairTests(unittest.TestCase):
    """Rule 3 — LowerRange and UpperRange come as a pair, and only on numbers."""

    def test_a_lower_bound_alone_is_an_error(self):
        # The blank upper bound is written as maxvalue='-9', so every answer
        # above -9 fails and the question cannot be answered at all.
        reader = read([row("age", "text", "text_integer", maxchars="3", lower="0")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("only one", errors(reader)[0])
        self.assertIn("UpperRange", errors(reader)[0])

    def test_an_upper_bound_alone_is_an_error(self):
        reader = read([row("age", "text", "text_integer", maxchars="3", upper="120")])

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("LowerRange", errors(reader)[0])

    def test_both_bounds_are_accepted(self):
        reader = read([numeric_row()])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(warnings(reader), [])

    def test_a_range_on_an_hourmin_field_is_an_error(self):
        reader = read(
            [row("timeseen", "text", "hourmin", maxchars="=5", lower="0", upper="23")]
        )

        self.assertTrue(reader.errorsEncountered)
        self.assertIn("cannot be applied to a 'hourmin' field", errors(reader)[0])


class MissingRangeWarningTests(unittest.TestCase):
    """Rule 3 — a quantity with no range warns; an identifier does not."""

    def test_a_variable_length_number_with_no_range_warns(self):
        reader = read([row("timestook", "text", "text_integer", maxchars="2")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(len(warnings(reader)), 1)
        self.assertIn("any value that fits MaxCharacters is accepted", warnings(reader)[0])

    def test_a_fixed_length_number_with_no_range_is_silent(self):
        # '=10' means an identifier — a phone number, a household ID — where a
        # numeric range means nothing.
        reader = read([row("phonenumber", "text", "text_integer", maxchars="=10")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(warnings(reader), [])

    def test_a_decimal_with_no_range_warns_too(self):
        reader = read([row("weight", "text", "text_decimal", maxchars="5")])

        self.assertEqual(len(warnings(reader)), 1)

    def test_a_plain_text_field_is_never_asked_for_a_range(self):
        reader = read([row("notes", "text", "text", maxchars="80")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(warnings(reader), [])

    def test_an_hourmin_field_is_never_asked_for_a_range(self):
        reader = read([row("timeseen", "text", "hourmin", maxchars="=5")])

        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertEqual(warnings(reader), [])


if __name__ == "__main__":
    unittest.main()
