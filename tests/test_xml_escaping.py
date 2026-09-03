"""Everything an author types must survive into well-formed XML.

Nothing tested escaping before this file. Every value was interpolated raw, so
an `&` in a response label or an apostrophe in a French `dont_know` label
produced XML that `ET.parse` rejects -- reported as "not well-formed: line 412,
column 30", with no field name.

The escaping is deliberately *idempotent*: the real dictionaries already
hand-escape some cells (`&lt;14 ou &gt;35` appears verbatim in an AVERT logic
check message), and escaping those again would ship literal "&lt;14" to the
device.
"""

import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from tempfile import TemporaryDirectory

from excel_reader import ExcelReader
from xml_generator import XmlGenerator, _esc_attr, _esc_text

from tests.test_dd_validations import HEADERS, numeric_row, row

from openpyxl import Workbook


def generate(rows):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "demo_dd"
    worksheet.append(HEADERS)
    for r in rows:
        worksheet.append(r)
    reader = ExcelReader()
    questions = reader.create_question_list(worksheet)
    with TemporaryDirectory() as tmp:
        path = XmlGenerator().write_xml("demo_dd", questions, Path(tmp))
        return path.read_text(encoding="utf-8"), reader


def parses(xml):
    """The whole document, through the same parser `_validate_xml_syntax` uses."""
    ET.fromstring(xml)
    return True


class HelperTests(unittest.TestCase):
    def test_text_escapes_the_three_markup_characters(self):
        self.assertEqual("a &amp; b &lt;c&gt;", _esc_text("a & b <c>"))

    def test_attributes_also_escape_both_quotes(self):
        self.assertEqual("d&apos;accord &quot;x&quot;", _esc_attr('d\'accord "x"'))

    def test_an_existing_character_reference_is_left_alone(self):
        """Idempotence -- the property that keeps hand-escaped cells working."""
        for already in ["&lt;14", "&gt;35", "&amp;", "&apos;", "&#233;", "&#x2019;"]:
            with self.subTest(already):
                self.assertEqual(already, _esc_text(already))

    def test_escaping_twice_changes_nothing(self):
        once = _esc_text("a & b <c> d'e")
        self.assertEqual(once, _esc_text(once))

    def test_a_bare_ampersand_next_to_a_reference_is_still_escaped(self):
        self.assertEqual("&amp; &lt;", _esc_text("& &lt;"))

    def test_something_that_only_looks_like_a_reference_is_escaped(self):
        self.assertEqual("&amp;not a ref", _esc_text("&not a ref"))


class QuestionTextTests(unittest.TestCase):
    def test_an_ampersand_in_question_text(self):
        xml, reader = generate([numeric_row("age", text="Bed nets & curtains?")])
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertTrue(parses(xml))
        self.assertIn("<text>Bed nets &amp; curtains?</text>", xml)

    def test_a_less_than_in_question_text(self):
        xml, _ = generate([numeric_row("age", text="Enfants <5 ans?")])
        self.assertTrue(parses(xml))
        self.assertIn("<text>Enfants &lt;5 ans?</text>", xml)

    def test_an_apostrophe_in_question_text_needs_no_escaping_in_element_text(self):
        xml, _ = generate([numeric_row("age", text="Qu'est-ce que c'est?")])
        self.assertTrue(parses(xml))
        self.assertEqual(
            "Qu'est-ce que c'est?", ET.fromstring(xml).find(".//question/text").text
        )


class ResponseLabelTests(unittest.TestCase):
    def radio(self, responses):
        return row("choice", "radio", "integer", text="Pick", responses=responses)

    def test_an_ampersand_in_a_response_label(self):
        xml, reader = generate([self.radio("1:Oui & non\n2:Autre")])
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertTrue(parses(xml))
        labels = [e.text for e in ET.fromstring(xml).findall(".//response")]
        self.assertIn("Oui & non", labels)

    def test_an_apostrophe_in_a_dont_know_label(self):
        """`label='...'` is single-quoted, so this was the likeliest break.

        `dont_know:` is a dynamic-source key, so the question has to draw its
        responses from a CSV for the label to exist at all.
        """
        xml, reader = generate(
            [
                self.radio(
                    "source:csv\nfile:villages.csv\ndisplay:name\nvalue:id\n"
                    "dont_know:88,Ne sait pas d'accord"
                )
            ]
        )
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertTrue(parses(xml))
        dont_know = ET.fromstring(xml).find(".//dont_know")
        self.assertEqual("Ne sait pas d'accord", dont_know.get("label"))

    def test_an_apostrophe_in_a_not_in_list_label(self):
        xml, reader = generate(
            [
                self.radio(
                    "source:csv\nfile:villages.csv\ndisplay:name\nvalue:id\n"
                    "not_in_list:96,N'est pas dans la liste"
                )
            ]
        )
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertTrue(parses(xml))
        self.assertEqual(
            "N'est pas dans la liste",
            ET.fromstring(xml).find(".//not_in_list").get("label"),
        )


class MaskAndMessageTests(unittest.TestCase):
    def test_a_double_quote_in_a_mask(self):
        """`mask value="..."` is the one double-quoted attribute in the format."""
        xml, reader = generate(
            [row("code", "text", "text", text="Code?", maxchars="6",
                 responses='mask:AA-"99"')]
        )
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertTrue(parses(xml))
        self.assertEqual('AA-"99"', ET.fromstring(xml).find(".//mask").get("value"))

    def logic_row(self, message):
        rows = [numeric_row("age"), numeric_row("weight", lower="0", upper="200")]
        rows[1][8] = f"weight > 0; '{message}'"
        return rows

    def test_an_ampersand_in_a_logic_check_message(self):
        xml, reader = generate(self.logic_row("Must be positive & realistic"))
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertTrue(parses(xml))
        self.assertIn("&amp; realistic", xml)

    def test_a_hand_escaped_logic_check_message_is_not_double_escaped(self):
        """Exactly what the AVERT French dictionary does."""
        xml, reader = generate(
            self.logic_row("doses espacees de &lt;14 ou &gt;35 jours")
        )
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertTrue(parses(xml))
        self.assertIn("&lt;14 ou &gt;35", xml)
        self.assertNotIn("&amp;lt;", xml)


class WholeDocumentTests(unittest.TestCase):
    def test_a_dictionary_full_of_special_characters_still_parses(self):
        rows = [
            numeric_row("age", text="Age & <5? \"quoted\" 'single'"),
            row(
                "choice",
                "radio",
                "integer",
                text="R&D <tag> \"q\"",
                responses="1:A & B\n2:C <D>\ndont_know:88,N'sait pas",
            ),
        ]
        xml, reader = generate(rows)
        self.assertFalse(reader.errorsEncountered, "\n".join(reader.logstring))
        self.assertTrue(parses(xml))


if __name__ == "__main__":
    unittest.main()
