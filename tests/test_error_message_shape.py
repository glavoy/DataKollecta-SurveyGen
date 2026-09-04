"""Every validation error names a category, so the log is scannable.

The messages are written by hand at 94 call sites across the reader modules,
and the review flagged that as drift waiting to happen. Measuring it turned up
something narrower than expected, which is worth recording because it decided
the fix:

  * The **location clause is already uniform** -- 63 of 63 messages say
    `in worksheet '...'`. Ten used to say `in table`, which is the one real
    inconsistency and is now gone.
  * The **prefix is uniform too**, after three messages were fixed: two had no
    category at all (`ERROR:` alone) and one used the category slot to hold the
    message (`ERROR - Duplicate fieldnames found in worksheet:`).
  * A `_field_error(category, worksheet, fieldname, ...)` helper -- the
    review's suggested fix -- would have fitted only **7** of the 94 sites.
    FieldName is the sentence's subject in 7 messages and appears
    mid-sentence in 64, because what the message is usually about is the
    offending *value*, not the field. Threading a helper through those would
    have flattened the specific wording that makes them useful.

So the invariant is asserted here instead of enforced by a constructor. It
costs one test rather than 94 rewrites, and it fails on the next message that
forgets its category.
"""

import inspect
import re
import unittest

import calculation_parser
import dd_validators
import excel_reader
import response_parser


# Every module that raises validation errors through `self._error(...)`.
#
# `excel_reader` was the only one when this test was written. The reader is
# being split into mixins, and each new module is added here in the same commit
# that moves code into it -- so `test_the_scan_finds_the_call_sites` below
# cannot start passing vacuously just because the call sites moved to a file
# nobody scans. That guard is the whole reason this list is explicit rather
# than a glob: a module missing from it fails loudly, a glob would quietly
# absorb whatever it found.
SCANNED_MODULES = (
    excel_reader,
    calculation_parser,
    response_parser,
    dd_validators,
)


# `ERROR - Category: `, where the category may be an f-string placeholder --
# LowerRange/UpperRange and the DontKnow/Refuse button names are chosen at
# runtime, which is deliberate and not drift.
PREFIX_RE = re.compile(r"ERROR - (\{[a-z_]+\}|[A-Za-z][A-Za-z ]*): ")

ERROR_CALL_RE = re.compile(r"self\._error\(\s*(.*?)\s*\)\n", re.S)
STRING_RE = re.compile(r'f?"((?:[^"\\]|\\.)*)"')


def error_messages() -> list[str]:
    """Every `self._error(...)` message, with its f-string pieces joined."""
    source = "\n".join(inspect.getsource(m) for m in SCANNED_MODULES)
    # Implicit concatenation means the pieces join with nothing between them;
    # collapsing runs of whitespace afterwards keeps a line break inside a
    # message from reading as a different wording.
    return [
        re.sub(r"\s+", " ", "".join(STRING_RE.findall(call)))
        for call in ERROR_CALL_RE.findall(source)
    ]


class ErrorPrefixTests(unittest.TestCase):
    def test_the_scan_finds_the_call_sites(self):
        # Guards the test itself: a refactor that changes how errors are
        # raised would otherwise make this suite vacuously pass.
        self.assertGreater(len(error_messages()), 80)

    def test_every_error_names_a_category(self):
        missing = [m[:70] for m in error_messages() if not PREFIX_RE.match(m)]

        self.assertEqual(missing, [], f"no `ERROR - Category: ` prefix: {missing}")

    def test_no_message_says_in_table(self):
        # The dictionary's tabs are worksheets; `tablename` is a crfs column
        # and means something else entirely, so "in table 'demo_dd'" pointed
        # the author at the wrong concept.
        offenders = [m[:70] for m in error_messages() if "in table '" in m]

        self.assertEqual(offenders, [])

    def test_the_location_clause_has_one_wording(self):
        located = [m for m in error_messages() if "worksheet" in m]
        odd = [
            m[:70] for m in located
            if "in worksheet '" not in m and "In worksheet '" not in m
        ]

        self.assertEqual(odd, [], f"unexpected location wording: {odd}")


if __name__ == "__main__":
    unittest.main()
