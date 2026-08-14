"""Operators in a `when` condition must reach the app in a form it evaluates.

The app's shared comparator accepts `=`, `!=`, `<>`, `>`, `<`, `>=`, `<=`,
`contains`, and `does not contain`, treating `<>` as a synonym for `!=` on
its own. This module still normalizes `<>` to `!=` on the way out, so only
one canonical spelling of "not equal" reaches the app either way.
"""

import re
import unittest

from tests.test_dd_validations import numeric_row, row
from tests.test_system_fields_xml import generate
from xml_generator import XmlGenerator


def case_row(fieldname, when_lines, else_value="0"):
    responses = "calc:case\n" + "\n".join(when_lines) + f"\nelse:{else_value}"
    return row(fieldname, "calc", "integer", responses=responses)


def when_attributes(xml):
    return re.findall(r"<when field='([a-z_0-9]+)' operator='([^']*)' value='([^']*)'>", xml)


class OperatorConversionTests(unittest.TestCase):
    def test_every_supported_operator_survives_the_round_trip(self):
        expected = {
            "=": "=",
            "!=": "!=",
            "<>": "!=",
            ">": "&gt;",
            "<": "&lt;",
            ">=": "&gt;=",
            "<=": "&lt;=",
            "contains": "contains",
            "does not contain": "does not contain",
        }

        for written, in_xml in expected.items():
            self.assertEqual(
                XmlGenerator._convert_operator_to_xml(written), in_xml, written
            )

    def test_an_unknown_operator_is_not_silently_turned_into_equals(self):
        # The regex that reads a when condition should make this unreachable;
        # failing loudly means a widened regex cannot introduce a silent bug.
        with self.assertRaises(ValueError):
            XmlGenerator._convert_operator_to_xml("~=")


class GeneratedCaseTests(unittest.TestCase):
    def test_not_equals_is_written_so_the_app_can_evaluate_it(self):
        xml, _ = generate(
            [
                numeric_row(),
                case_row("adult", ["when:age <> 18 => 1"]),
            ]
        )

        self.assertIn(("age", "!=", "18"), when_attributes(xml))
        self.assertNotIn("&lt;&gt;", xml)

    def test_both_spellings_produce_the_same_file(self):
        angle, _ = generate([numeric_row(), case_row("adult", ["when:age <> 18 => 1"])])
        bang, _ = generate([numeric_row(), case_row("adult", ["when:age != 18 => 1"])])

        self.assertEqual(angle, bang)

    def test_comparisons_are_still_escaped(self):
        xml, _ = generate(
            [
                numeric_row(),
                case_row(
                    "band",
                    ["when:age >= 18 => 1", "when:age < 5 => 2"],
                ),
            ]
        )

        self.assertIn(("age", "&gt;=", "18"), when_attributes(xml))
        self.assertIn(("age", "&lt;", "5"), when_attributes(xml))

    def test_contains_is_written_for_the_app_to_evaluate(self):
        xml, _ = generate(
            [
                numeric_row(fieldname="screen_cab_drug2"),
                case_row(
                    "exclude", ["when:screen_cab_drug2 contains 99 => 0"], else_value="1"
                ),
            ]
        )

        self.assertIn(("screen_cab_drug2", "contains", "99"), when_attributes(xml))

    def test_does_not_contain_is_written_for_the_app_to_evaluate(self):
        xml, _ = generate(
            [
                numeric_row(fieldname="screen_cab_drug2"),
                case_row(
                    "exclude",
                    ["when:screen_cab_drug2 does not contain 99 => 0"],
                    else_value="1",
                ),
            ]
        )

        self.assertIn(
            ("screen_cab_drug2", "does not contain", "99"), when_attributes(xml)
        )

    def test_contains_is_case_insensitive_and_normalizes_to_lowercase(self):
        xml, _ = generate(
            [
                numeric_row(fieldname="screen_cab_drug2"),
                case_row(
                    "exclude", ["when:screen_cab_drug2 Contains 99 => 0"], else_value="1"
                ),
            ]
        )

        self.assertIn(("screen_cab_drug2", "contains", "99"), when_attributes(xml))


if __name__ == "__main__":
    unittest.main()
