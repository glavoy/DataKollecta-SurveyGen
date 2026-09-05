"""The matrix variants, and that they still build.

Two halves with different requirements, deliberately kept in separate classes.

`MatrixDeclarationTests` is pure and always runs: it pins the precedence in
`matrix.expected_outcome`, which is a restatement of the app's
`RepeatCountService.evaluate` and is the thing the simulator compares against.
Getting that wrong would make the simulator agree with itself and with nothing
else.

`VariantBuildTests` needs a real dictionary, which is not in the repo -- no
`.xlsx` is tracked here, deliberately -- so it skips with a clear message when
one cannot be found. When it can, it is the closest thing this repo has to an
end-to-end test: twelve real packages built from a real study, asserted to
produce no errors and no warning nobody has seen before.

That last part is the point of enumerating warnings rather than ignoring them.
A new check that fires on a legitimate construct is exactly the kind of thing
that gets waved through when the log is already noisy; here it fails a test.
"""

from __future__ import annotations

import io
import json
import os
import re
import unittest
import zipfile
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory

from models import AppConfig
from processor import SurveyGenProcessor
from test_kit import matrix as matrix_module
from test_kit.make_variants import VariantError, build

REPO_ROOT = Path(__file__).resolve().parent.parent

# `WARNING - <category>: <message>`. Only the category is asserted on: the
# messages name worksheets and values, so pinning them whole would make this
# test fail every time a dictionary is revised, which is not what it is for.
_WARNING_CATEGORY = re.compile(r"^WARNING - ([^:]+):")

# Every warning a clean variant is allowed to produce.
#
# `Reserved variable` -- PRISM declares `starttime`/`stoptime` rows that the
# generator writes itself and therefore ignores. A property of the source
# dictionary, not of the variants.
#
# `databaseName` -- `_check_database_name_stability` warns when databaseName is
# the surveyId plus `.sqlite`, because in a real study that restarts the
# subject-ID counter on every revision. Here that is the intent: each cell is a
# throwaway database that must not share a counter with the other eleven.
EXPECTED_WARNING_CATEGORIES = frozenset({"Reserved variable", "databaseName"})


def _source_dictionary() -> tuple[Path, Path] | None:
    """The dictionary to build variants from, and where its CSVs live.

    Two ways to point at one, both outside the repo because the dictionaries
    are real studies and are deliberately gitignored:

    1. `SURVEYGEN_TEST_DICTIONARY`, for anyone who keeps theirs elsewhere.
    2. `config_prism_css.json`, which already names both paths for the study
       this kit was built around. Reusing it means there is nothing extra to
       configure on the machine that already builds PRISM.
    """
    override = os.environ.get("SURVEYGEN_TEST_DICTIONARY")
    if override:
        source = Path(override).expanduser()
        csv_dir = Path(
            os.environ.get("SURVEYGEN_TEST_CSV_DIR", str(source.parent))
        ).expanduser()
        return (source, csv_dir) if source.is_file() else None

    config_path = REPO_ROOT / "config_prism_css.json"
    if not config_path.is_file():
        return None
    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    source = Path(data.get("excelFile", "")).expanduser()
    csv_dir = Path(data.get("csvFiles", "") or source.parent).expanduser()
    return (source, csv_dir) if source.is_file() else None


def _run_quietly(processor: SurveyGenProcessor) -> int:
    """`run` prints the Windows app's SUCCESS/ERRORS message box equivalent."""
    with redirect_stdout(io.StringIO()):
        return processor.run()


class MatrixDeclarationTests(unittest.TestCase):
    def test_every_cell_is_covered_exactly_once(self):
        cells = matrix_module.cells()
        self.assertEqual(len(cells), 12)
        self.assertEqual(len({c.key for c in cells}), 12)
        self.assertEqual(
            {(c.auto_start_repeat, c.repeat_enforce_count) for c in cells},
            {
                (a, e)
                for a in matrix_module.AUTO_START_VALUES
                for e in matrix_module.ENFORCE_VALUES
            },
        )

    def test_only_modes_one_and_two_start_a_loop(self):
        for cell in matrix_module.cells():
            with self.subTest(cell=cell.key):
                self.assertEqual(cell.starts_loop, cell.auto_start_repeat in (1, 2))
                self.assertEqual(cell.prompts_first, cell.auto_start_repeat == 1)
                self.assertEqual(cell.blocks_exit, cell.repeat_enforce_count == 2)

    def test_a_skipped_count_is_left_alone_whatever_the_mode(self):
        # The highest-precedence rule: NULL means the count question was
        # skipped, and filling it in would undo the skip logic that emptied it.
        for cell in matrix_module.cells():
            with self.subTest(cell=cell.key):
                self.assertEqual(
                    matrix_module.expected_outcome(cell, declared=None, actual=3),
                    matrix_module.OUTCOME_COUNT_NOT_DECLARED,
                )

    def test_a_matching_count_outranks_the_range_gate(self):
        # `inSync` is checked before LowerRange/UpperRange, so a declared count
        # equal to the actual one needs no decision even when both sit outside
        # the range the question declares. Pinning it because the reverse order
        # would look just as reasonable and would be wrong.
        cell = matrix_module.MatrixCell(auto_start_repeat=2, repeat_enforce_count=3)
        self.assertEqual(
            matrix_module.expected_outcome(
                cell, declared=40, actual=40, lower=1, upper=30
            ),
            matrix_module.OUTCOME_IN_SYNC,
        )

    def test_the_range_gate_outranks_the_mode(self):
        for cell in matrix_module.cells():
            with self.subTest(cell=cell.key):
                self.assertEqual(
                    matrix_module.expected_outcome(
                        cell, declared=6, actual=0, lower=1, upper=30
                    ),
                    matrix_module.OUTCOME_BELOW_MINIMUM,
                )
                self.assertEqual(
                    matrix_module.expected_outcome(
                        cell, declared=6, actual=31, lower=1, upper=30
                    ),
                    matrix_module.OUTCOME_ABOVE_MAXIMUM,
                )

    def test_a_count_question_with_no_range_reconciles_unconditionally(self):
        cell = matrix_module.MatrixCell(auto_start_repeat=2, repeat_enforce_count=3)
        self.assertEqual(
            matrix_module.expected_outcome(cell, declared=6, actual=0),
            matrix_module.OUTCOME_UPDATE_SILENTLY,
        )

    def test_each_enforce_mode_reaches_its_own_outcome(self):
        expected = {
            0: matrix_module.OUTCOME_NO_ENFORCEMENT,
            1: matrix_module.OUTCOME_ASK_TO_UPDATE,
            2: matrix_module.OUTCOME_FORCE_MODE_HANDLED_IN_LOOP,
            3: matrix_module.OUTCOME_UPDATE_SILENTLY,
        }
        for cell in matrix_module.cells():
            with self.subTest(cell=cell.key):
                self.assertEqual(
                    matrix_module.expected_outcome(
                        cell, declared=6, actual=4, lower=1, upper=30
                    ),
                    expected[cell.repeat_enforce_count],
                )

    def test_ask_to_update_is_not_claimed_to_write(self):
        # Whether mode 1 writes depends on which button the interviewer
        # pressed, so the simulator has to record its own answer. Claiming it
        # here would make the simulator assert a write that may not happen.
        self.assertNotIn(
            matrix_module.OUTCOME_ASK_TO_UPDATE, matrix_module.OUTCOMES_THAT_WRITE
        )
        self.assertIn(
            matrix_module.OUTCOME_UPDATE_SILENTLY, matrix_module.OUTCOMES_THAT_WRITE
        )

    def test_the_json_the_simulator_reads_names_every_cell_and_outcome(self):
        payload = matrix_module.to_json_dict()
        self.assertEqual(len(payload["cells"]), 12)
        self.assertEqual(set(payload["outcomes"]), set(matrix_module.ALL_OUTCOMES))
        self.assertEqual(payload["matrixChild"], matrix_module.MATRIX_CHILD)
        # Serialisable, since its whole purpose is to be read by another
        # language.
        json.dumps(payload)


class VariantBuildTests(unittest.TestCase):
    """Twelve real packages, built from a real study."""

    @classmethod
    def setUpClass(cls):
        found = _source_dictionary()
        if found is None:
            raise unittest.SkipTest(
                "No source dictionary. Set SURVEYGEN_TEST_DICTIONARY to a .xlsx, "
                "or keep config_prism_css.json pointing at one. Dictionaries are "
                "deliberately not tracked in this repo."
            )
        cls.source, cls.csv_dir = found
        cls._tmp = TemporaryDirectory()
        cls.out_dir = Path(cls._tmp.name) / "variants"
        cls.output_path = Path(cls._tmp.name) / "packages"
        cls.output_path.mkdir(parents=True)
        cls.variants = build(
            cls.source,
            cls.out_dir,
            csv_dir=cls.csv_dir,
            output_path=cls.output_path,
        )

        # Built once, here, rather than per test method. Each build is a full
        # run over a four-form study, so rebuilding for every assertion turned
        # twelve builds into thirty-six and the suite from half a second into
        # four.
        cls.results = {}
        for variant in cls.variants:
            config = json.loads(variant.config.read_text(encoding="utf-8"))
            processor = SurveyGenProcessor(AppConfig(**config))
            code = _run_quietly(processor)
            cls.results[variant.cell.key] = (code, processor)

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "_tmp"):
            cls._tmp.cleanup()

    def test_every_variant_builds_with_no_errors(self):
        for variant in self.variants:
            with self.subTest(cell=variant.cell.key):
                code, processor = self.results[variant.cell.key]
                errors = [
                    line for line in processor.logstring if line.startswith("ERROR")
                ]
                self.assertEqual(errors, [], f"{variant.cell.key} produced errors")
                self.assertEqual(code, 0)
                self.assertTrue(
                    (self.output_path / f"{variant.survey_id}.zip").is_file(),
                    f"{variant.cell.key} produced no zip",
                )

    def test_no_variant_produces_an_unexpected_warning(self):
        seen: set[str] = set()
        for _, processor in self.results.values():
            for line in processor.logstring:
                match = _WARNING_CATEGORY.match(line)
                if match:
                    seen.add(match.group(1).strip())

        unexpected = seen - EXPECTED_WARNING_CATEGORIES
        self.assertEqual(
            unexpected,
            set(),
            "A new warning category fired on a dictionary that is known to be "
            "correct. Either the check is wrong, or the list in this file needs "
            "the new category added with a note saying why it is acceptable.",
        )

    def test_the_override_reaches_the_manifest_and_nothing_else_moves(self):
        for variant in self.variants:
            with self.subTest(cell=variant.cell.key):
                archive = zipfile.ZipFile(self.output_path / f"{variant.survey_id}.zip")
                manifest = json.loads(
                    archive.read("survey_manifest.gistx").decode("utf-8")
                )
                crfs = {c["tablename"]: c for c in manifest["crfs"]}

                child = crfs[matrix_module.MATRIX_CHILD]
                self.assertEqual(
                    child.get("auto_start_repeat"), variant.cell.auto_start_repeat
                )
                self.assertEqual(
                    child.get("repeat_enforce_count"),
                    variant.cell.repeat_enforce_count,
                )

                # The siblings keep the study's own configuration, so every
                # variant still has several repeating children in
                # display_order -- which is what exercises the sequencing in
                # the app's auto-repeat loop.
                siblings = [
                    name
                    for name, crf in crfs.items()
                    if crf.get("parenttable") and name != matrix_module.MATRIX_CHILD
                ]
                self.assertTrue(siblings, "the fixture should have sibling children")
                for name in siblings:
                    self.assertEqual(crfs[name].get("auto_start_repeat"), 2)
                    self.assertEqual(crfs[name].get("repeat_enforce_count"), 3)

    def test_each_variant_gets_its_own_survey_id_and_database(self):
        # Both are device-global storage keys: surveyId is the extraction
        # folder, databaseName is the SQLite filename. Two variants sharing
        # either would share a subject-ID counter.
        self.assertEqual(
            len({v.survey_id for v in self.variants}), len(self.variants)
        )
        self.assertEqual(
            len({v.database_name for v in self.variants}), len(self.variants)
        )
        for variant in self.variants:
            self.assertTrue(variant.survey_id.startswith("zz_"))

    def test_a_dictionary_without_the_named_child_is_refused(self):
        with TemporaryDirectory() as tmp:
            with self.assertRaises(VariantError) as caught:
                build(
                    self.source,
                    Path(tmp),
                    csv_dir=self.csv_dir,
                    output_path=Path(tmp),
                    child="no_such_form",
                )
            self.assertIn("no_such_form", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
