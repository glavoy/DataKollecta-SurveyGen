"""Every calculation type must be reachable, validated, and emitted.

Adding a CalculationType means touching three places that have no compile-time
relationship: the alias table (so a dictionary can name it), the validator (so
a half-written one is rejected), and the emitter (so it reaches the XML). None
of them fails loudly when a new type is missed -- a type with no alias is
simply unreachable, and one the emitter skips is silently absent from the
generated file, which is the C4 failure shape all over again.

These are coverage tests rather than behaviour tests: they assert that the
three enumerations agree, so the next person to add a type finds out here
instead of in the field.
"""

import inspect
import unittest

from models import (
    CALCULATION_ALIASES,
    CALCULATION_TYPE_BY_ALIAS,
    CalculationType,
    calculation_alias_list,
)
import excel_reader
import xml_generator


CALCULABLE_TYPES = [t for t in CalculationType if t != CalculationType.NONE]


class AliasCoverageTests(unittest.TestCase):
    def test_every_calculation_type_has_a_dictionary_word(self):
        missing = [t.name for t in CALCULABLE_TYPES if t not in CALCULATION_ALIASES]

        self.assertEqual(missing, [], f"unreachable from a dictionary: {missing}")

    def test_none_is_deliberately_not_addressable(self):
        # NONE means "this question has no calculation". A `calc:none` would be
        # a way to write a calculation that is not one.
        self.assertNotIn(CalculationType.NONE, CALCULATION_ALIASES)

    def test_aliases_are_unique(self):
        self.assertEqual(len(CALCULATION_TYPE_BY_ALIAS), len(CALCULATION_ALIASES))

    def test_aliases_are_lowercase_with_no_spaces(self):
        # The reader lowercases the value before looking it up, so an alias
        # with a capital or a space could never match.
        for alias in CALCULATION_ALIASES.values():
            with self.subTest(alias=alias):
                self.assertEqual(alias, alias.lower().strip())
                self.assertNotIn(" ", alias)

    def test_the_error_message_lists_every_word(self):
        # It used to be a hand-typed list of the same twelve words.
        listed = calculation_alias_list()
        for alias in CALCULATION_ALIASES.values():
            with self.subTest(alias=alias):
                self.assertIn(f"'{alias}'", listed)


class ValidatorCoverageTests(unittest.TestCase):
    def test_the_validator_names_every_calculation_type(self):
        source = inspect.getsource(excel_reader.ExcelReader._validate_calculation_fields)
        unhandled = [t.name for t in CALCULABLE_TYPES if f"CalculationType.{t.name}" not in source]

        self.assertEqual(
            unhandled,
            [],
            f"_validate_calculation_fields has no branch for: {unhandled} -- a "
            "half-written calculation of that type would pass validation and "
            "ship.",
        )


class EmitterCoverageTests(unittest.TestCase):
    def test_the_emitter_names_every_calculation_type(self):
        source = inspect.getsource(xml_generator.XmlGenerator._generate_calculation_xml)
        unhandled = [t.name for t in CALCULABLE_TYPES if f"CalculationType.{t.name}" not in source]

        self.assertEqual(
            unhandled,
            [],
            f"_generate_calculation_xml has no branch for: {unhandled} -- a "
            "question of that type would validate clean and then be written "
            "with no calculation at all.",
        )

    def test_an_unhandled_type_would_not_pass_silently(self):
        # The guard that makes the coverage test above more than a formality:
        # if a type ever does reach the emitter unhandled, it must be loud.
        # This is the same lesson as the skip parser -- the old two-parser
        # split dropped a rule silently, and the replacement raises.
        source = inspect.getsource(xml_generator.XmlGenerator._generate_calculation_xml)

        self.assertIn(
            "raise",
            source,
            "_generate_calculation_xml has no final `raise`, so an unhandled "
            "calculation type would emit nothing and report success.",
        )
