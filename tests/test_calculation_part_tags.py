"""A compound calculation's operands must use the element name the app reads.

`survey_loader.dart` picks the element by **context**, not by what kind of
operand it is:

    math / concat   ->  node.findElements('part')      (line 452)
    when            ->  whenNode.getElement('result')  (line 468)
    else            ->  elseNode.getElement('result')  (line 484)

The generator used to pick by operand *type* -- `<result>` for a constant,
`<part>` for everything else -- so `part:constant 10` inside a `calc:math`
emitted `<result>` into a context that only looks for `<part>`. The XML stayed
well-formed, every validation passed, the run reported success, and the
calculation quietly ran with one fewer operand: `price + 10` computed `price`.

Nothing caught it because no test asserted either tag, and neither live
dictionary uses math or concat. These tests assert the tags with the same
lookups the app performs, so the next person to touch the emitter finds out
here.
"""

import unittest
import xml.etree.ElementTree as ET

from tests.test_system_fields_xml import generate


def row(fieldname, qtype, ftype, responses="", maxchars="", lower="", upper=""):
    return [fieldname, qtype, ftype, "Q text", maxchars, responses, lower,
            upper, "", "", "", "", "", ""]


def price_row():
    """A plain numeric question for the calculations to look up."""
    return row("price", "text", "text_integer", maxchars="3", lower="0", upper="999")


def calculation_for(xml: str, fieldname: str) -> ET.Element:
    """The `<calculation>` element of one question, found as the app would."""
    root = ET.fromstring(xml)
    for question in root.findall("question"):
        if question.get("fieldname") == fieldname:
            calc = question.find("calculation")
            assert calc is not None, f"no <calculation> on '{fieldname}'"
            return calc
    raise AssertionError(f"no question '{fieldname}' in the generated XML")


class MathOperandTagTests(unittest.TestCase):
    def test_every_math_operand_is_a_part_element(self):
        xml, _ = generate([
            price_row(),
            row("total", "automatic", "integer",
                "calc:math\noperator:+\npart:lookup price\npart:constant 10"),
        ])

        calc = calculation_for(xml, "total")

        # findElements('part') is literally what the app runs, so this count
        # is the number of operands the arithmetic will actually use.
        parts = calc.findall("part")
        self.assertEqual(len(parts), 2, ET.tostring(calc, encoding="unicode"))
        self.assertEqual(
            [(p.get("type"), p.get("field") or p.get("value")) for p in parts],
            [("lookup", "price"), ("constant", "10")],
        )

    def test_a_constant_operand_is_not_emitted_as_result(self):
        # The regression itself: <result> here is invisible to the app.
        xml, _ = generate([
            price_row(),
            row("total", "automatic", "integer",
                "calc:math\noperator:+\npart:lookup price\npart:constant 10"),
        ])

        calc = calculation_for(xml, "total")

        self.assertEqual(calc.findall("result"), [])

    def test_a_math_of_two_constants_keeps_both_operands(self):
        # Nothing but constants, so if the tag were type-driven the app would
        # find no operands at all and the calculation would be empty.
        xml, _ = generate([
            price_row(),
            row("total", "automatic", "integer",
                "calc:math\noperator:*\npart:constant 2\npart:constant 3"),
        ])

        self.assertEqual(len(calculation_for(xml, "total").findall("part")), 2)


class ConcatOperandTagTests(unittest.TestCase):
    def test_every_concat_operand_is_a_part_element(self):
        xml, _ = generate([
            price_row(),
            row("label", "automatic", "text",
                "calc:concat\nseparator:-\npart:lookup price\npart:constant USD"),
        ])

        calc = calculation_for(xml, "label")

        parts = calc.findall("part")
        self.assertEqual(len(parts), 2, ET.tostring(calc, encoding="unicode"))
        self.assertEqual(calc.get("separator"), "-")
        self.assertEqual(calc.findall("result"), [])


class CaseResultTagTests(unittest.TestCase):
    """The other half of the same rule: inside a case, `<result>` is correct.

    This path was always right -- a case result is always a constant, so the
    old type-driven tag happened to agree -- and these tests exist to stop the
    fix swinging the other way and renaming them all to `<part>`.
    """

    def test_a_when_result_is_a_result_element(self):
        xml, _ = generate([
            price_row(),
            row("band", "automatic", "text",
                "calc:case\nwhen:price < 18 => Minor\nelse:Adult"),
        ])

        calc = calculation_for(xml, "band")
        when = calc.find("when")

        self.assertIsNotNone(when.find("result"))
        self.assertEqual(when.find("result").get("value"), "Minor")
        self.assertEqual(when.findall("part"), [])

    def test_an_else_result_is_a_result_element(self):
        xml, _ = generate([
            price_row(),
            row("band", "automatic", "text",
                "calc:case\nwhen:price < 18 => Minor\nelse:Adult"),
        ])

        else_node = calculation_for(xml, "band").find("else")

        self.assertIsNotNone(else_node.find("result"))
        self.assertEqual(else_node.find("result").get("value"), "Adult")
        self.assertEqual(else_node.findall("part"), [])


class UnknownPartTypeTests(unittest.TestCase):
    def test_an_unhandled_part_type_raises_instead_of_vanishing(self):
        # The MATH/CONCAT branches removed from _generate_calculation_part were
        # unreachable (a dictionary can only write constant/lookup/query), and
        # the function had no final else -- so a new part type would have
        # emitted nothing and reported success. Same shape as C4.
        from models import CalculationPart, CalculationType
        from xml_generator import XmlGenerator

        part = CalculationPart(type=CalculationType.TIMESTAMP)

        with self.assertRaises(ValueError):
            XmlGenerator()._generate_calculation_part(
                lambda line: None, part, 3, "part"
            )


if __name__ == "__main__":
    unittest.main()
