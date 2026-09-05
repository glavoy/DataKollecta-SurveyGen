"""The repeat matrix, declared once.

`crfs.auto_start_repeat` (0/1/2) and `crfs.repeat_enforce_count` (0/1/2/3) are
independent, so together they describe twelve configurations. **No real
dictionary covers more than one of them.** Both shipped dictionaries leave all
four repeat columns blank; PRISM CSS sets `(2, 3)` and sets it identically on
all three of its children. So eleven of the twelve have never been built, let
alone run.

This module is the single declaration of what those twelve are and what each
one should do. `make_variants.py` reads it to write the workbooks, the Python
test kit reads it to assert they build, and the simulator reads the JSON it
emits to know what to expect from the records. One source, three readers --
the alternative is the same table transcribed into a Dart file, and a matrix
that disagrees with the fixtures it describes is worse than no matrix.

Two blank-cell rules from the README are deliberately *not* modelled here, and
are covered by the linter instead (`form.repeat.valueOutOfDomain`): a blank
`auto_start_repeat` reads as 0, and a blank `repeat_enforce_count` reads as 1.
Every cell below names its value explicitly, so nothing here depends on a
default.
"""

from __future__ import annotations

from dataclasses import dataclass

# ── The axes ─────────────────────────────────────────────────────────────────
# 0 off, 1 prompt, 2 force. Decides only whether the loop starts, and whether
# the interviewer is asked first.
AUTO_START_VALUES: tuple[int, ...] = (0, 1, 2)

# 0 flexible, 1 warn, 2 force, 3 auto-sync. Decides the exit rule inside the
# loop and what happens to the parent's count afterwards.
ENFORCE_VALUES: tuple[int, ...] = (0, 1, 2, 3)

# The child form each variant reconfigures.
#
# `hh_members` rather than `nets`, for one reason: its count question
# (`nmembers`, 1..30) is answered on every path, so a variant that changes only
# the two repeat cells changes only the thing under test. `nnets` carries
# `preskip: if havenets <> 1, skip to totals`, so a variant built on `nets`
# would conflate "this enforce mode does X" with "the count was NULL", which
# is a different row of the table.
#
# The other two children keep PRISM's real `(2, 3)` in every variant. That is
# not laziness: it means each variant still has three repeating children in
# `display_order`, so the sequencing in `_checkAndStartAutoRepeat` is exercised
# every run, and `nets` still supplies the skippable-count case alongside
# whatever cell is under test.
MATRIX_CHILD = "hh_members"

# The child whose count question can be skipped, so the NULL-count row of the
# table has somewhere to happen. Named here rather than in the simulator so
# the fact lives next to the matrix it qualifies.
SKIPPABLE_COUNT_CHILD = "nets"


@dataclass(frozen=True)
class MatrixCell:
    """One (auto_start_repeat, repeat_enforce_count) configuration."""

    auto_start_repeat: int
    repeat_enforce_count: int

    @property
    def key(self) -> str:
        """Short, sortable, and safe in a filename, a surveyId and a table name."""
        return f"a{self.auto_start_repeat}e{self.repeat_enforce_count}"

    @property
    def starts_loop(self) -> bool:
        """Whether saving the parent starts the child loop at all.

        The negative case is the one worth having a name for: with
        `auto_start_repeat = 0` there is no loop, so `repeat_enforce_count` has
        nothing to enforce and the count is only ever reconciled when a child
        is saved on its own later.
        """
        return self.auto_start_repeat in (1, 2)

    @property
    def prompts_first(self) -> bool:
        """Mode 1 asks "you indicated X records -- add them now?" before looping."""
        return self.auto_start_repeat == 1

    @property
    def blocks_exit(self) -> bool:
        """Mode 2 has no "Exit Anyway": the loop cannot be left short."""
        return self.repeat_enforce_count == 2

    def describe(self) -> str:
        starts = {
            0: "no loop",
            1: "loop after a prompt",
            2: "loop immediately",
        }[self.auto_start_repeat]
        enforces = {
            0: "any count accepted",
            1: "offer to update the count",
            2: "block exit until the count matches",
            3: "update the count silently",
        }[self.repeat_enforce_count]
        return f"{starts}; {enforces}"


def cells() -> list[MatrixCell]:
    """All twelve, in a stable order."""
    return [
        MatrixCell(auto_start_repeat=a, repeat_enforce_count=e)
        for a in AUTO_START_VALUES
        for e in ENFORCE_VALUES
    ]


# ── What the app should decide ───────────────────────────────────────────────
# These names are `RepeatCountOutcome` in
# `DataKollecta/lib/services/repeat_count_service.dart`. They are spelled out
# here rather than imported (there is nothing to import across languages) and
# asserted equal by the simulator, so a rename on either side fails loudly
# instead of quietly comparing two different vocabularies.
OUTCOME_IN_SYNC = "inSync"
OUTCOME_NO_ENFORCEMENT = "noEnforcement"
OUTCOME_COUNT_NOT_DECLARED = "countNotDeclared"
OUTCOME_BELOW_MINIMUM = "belowMinimum"
OUTCOME_ABOVE_MAXIMUM = "aboveMaximum"
OUTCOME_FORCE_MODE_HANDLED_IN_LOOP = "forceModeHandledInLoop"
OUTCOME_ASK_TO_UPDATE = "askToUpdate"
OUTCOME_UPDATE_SILENTLY = "updateSilently"

ALL_OUTCOMES: tuple[str, ...] = (
    OUTCOME_IN_SYNC,
    OUTCOME_NO_ENFORCEMENT,
    OUTCOME_COUNT_NOT_DECLARED,
    OUTCOME_BELOW_MINIMUM,
    OUTCOME_ABOVE_MAXIMUM,
    OUTCOME_FORCE_MODE_HANDLED_IN_LOOP,
    OUTCOME_ASK_TO_UPDATE,
    OUTCOME_UPDATE_SILENTLY,
)

# Outcomes after which the parent's count column holds the number of children
# actually entered rather than the number the interviewer declared.
#
# `askToUpdate` is deliberately absent: whether it writes depends on which
# button the interviewer pressed, so the simulator has to record its own answer
# and cannot infer it from the outcome alone.
OUTCOMES_THAT_WRITE: frozenset[str] = frozenset({OUTCOME_UPDATE_SILENTLY})


def expected_outcome(
    cell: MatrixCell,
    *,
    declared: int | None,
    actual: int,
    lower: int | None = None,
    upper: int | None = None,
) -> str:
    """What reconciliation should decide, given a parent and its children.

    A restatement of `RepeatCountService.evaluate`, written from its
    documented behaviour rather than from its code, so that the simulator
    comparing the two is a real check and not a tautology. **If this and the
    Dart ever disagree, that disagreement is the finding** -- do not "fix" this
    to match without first working out which one is right.

    The precedence is the part worth getting exactly right, because three of
    the four gates sit *before* the mode is even looked at:

    1. A count question that was skipped stores NULL, and NULL is left alone --
       filling it in would undo the skip logic that emptied it.
    2. A count that already matches needs no decision. Note this is checked
       **before** the range gate, so a declared count equal to the actual one
       is `inSync` even when both are outside the question's declared range.
    3. The count question's own `LowerRange`/`UpperRange` gate every write. The
       app must not store a number the interviewer could not have typed.
    4. Only then does `repeat_enforce_count` choose.
    """
    if declared is None:
        return OUTCOME_COUNT_NOT_DECLARED

    if declared == actual:
        return OUTCOME_IN_SYNC

    if lower is not None and actual < lower:
        return OUTCOME_BELOW_MINIMUM
    if upper is not None and actual > upper:
        return OUTCOME_ABOVE_MAXIMUM

    return {
        0: OUTCOME_NO_ENFORCEMENT,
        1: OUTCOME_ASK_TO_UPDATE,
        2: OUTCOME_FORCE_MODE_HANDLED_IN_LOOP,
        3: OUTCOME_UPDATE_SILENTLY,
    }[cell.repeat_enforce_count]


# ── The scenarios each cell is run through ───────────────────────────────────
# `children` is how many child records the simulated interviewer completes,
# relative to the count the parent declared. `declared_is_null` covers the
# count question having been skipped, which needs the child named in
# SKIPPABLE_COUNT_CHILD rather than MATRIX_CHILD.
SCENARIOS: tuple[dict, ...] = (
    {"name": "equal", "delta": 0, "declared_is_null": False},
    {"name": "fewer", "delta": -2, "declared_is_null": False},
    {"name": "more", "delta": +2, "declared_is_null": False},
    {"name": "null_count", "delta": 0, "declared_is_null": True},
)


def to_json_dict() -> dict:
    """The whole declaration, for the readers that are not Python.

    `make_variants.py` writes this next to the variants so the simulator reads
    the same table that produced the workbooks it is running, rather than a
    copy that was accurate when someone last edited it.
    """
    return {
        "matrixChild": MATRIX_CHILD,
        "skippableCountChild": SKIPPABLE_COUNT_CHILD,
        "outcomes": list(ALL_OUTCOMES),
        "outcomesThatWrite": sorted(OUTCOMES_THAT_WRITE),
        "scenarios": [dict(s) for s in SCENARIOS],
        "cells": [
            {
                "key": cell.key,
                "autoStartRepeat": cell.auto_start_repeat,
                "repeatEnforceCount": cell.repeat_enforce_count,
                "startsLoop": cell.starts_loop,
                "promptsFirst": cell.prompts_first,
                "blocksExit": cell.blocks_exit,
                "description": cell.describe(),
            }
            for cell in cells()
        ],
    }
