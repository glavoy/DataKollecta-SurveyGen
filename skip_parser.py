"""The one parser for a Skip cell.

Both the validator (`excel_reader._check_skip_syntax`) and the generator
(`xml_generator._generate_skip`) go through this module. That is the whole
point of it: they used to parse the same string independently -- the validator
by counting tokens, the generator by counting spaces and slicing at hardcoded
offsets (`len_skip = 13 if postSkip else 12`, the length of `"postskip: if "`).

The two disagreed. `Preskip: if has_children = 0, skip to occupation` -- which
is what Excel produces, because AutoCorrect capitalizes the first letter of a
cell by default -- passed every validation check, then matched neither the
generator's `preskip` nor its `postskip` test, so **no skip element was written
at all**. The log said "No errors found" and the branching logic simply did not
exist in the field.

With one parser they cannot disagree: whatever the validator accepted is
exactly what the generator emits, and a string this module rejects never
reaches the generator.

Grammar (README "Preskip" / "Postskip"):

    preskip:  if <field> <operator> <value>, skip to <target>
    postskip: if <field> <operator> <value>, skip to <target>

`skip to end` is a reserved target rather than a real field name. Field names
cannot contain spaces (enforced in `_check_field_name`), and a value is a
single token -- the previous token-count check required exactly 5 space
separated tokens before the comma, so a multi-word value was already a hard
error, and staying single-token keeps that contract.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# Longest-first: `does not contain` before `contains`, `<>`/`>=`/`<=` before
# the single-character operators they start with. Quoted spellings are what the
# README documents for the two word operators ('contains'), but the bare form
# has always parsed too, so both stay accepted.
_OPERATOR_PATTERN = (
    r"'does not contain'|'contains'|does not contain|contains"
    r"|<>|!=|>=|<=|=|>|<"
)

_IDENTIFIER = r"[A-Za-z_][A-Za-z0-9_]*"

SKIP_RE = re.compile(
    r"^\s*(?P<kind>preskip|postskip)\s*:\s*if\s+"
    rf"(?P<field>{_IDENTIFIER})\s+"
    rf"(?P<operator>{_OPERATOR_PATTERN})\s+"
    r"(?P<value>\S+?)\s*,\s*"
    r"(?P<target_clause>.+?)\s*$",
    re.IGNORECASE,
)

# The words that may sit between the comma and the target. Three spellings are
# in live use across the real dictionaries, and all three worked before,
# because the old generator simply took everything after the LAST space:
#
#     ..., skip to everreceivesmc        <- README's examples
#     ..., vx_dose_summary               <- AVERT (fr); the README's format
#                                           line reads `skip_to_target`, which
#                                           genuinely looks like just a target
#     ..., then skip to everreceivesmc   <- PRISM CSS
#
# So the target is the last token, and anything before it must be filler.
# Being stricter than this is not a fix -- it silently rejected 40+ working
# skips across two of the two dictionaries available when this was written.
_TARGET_FILLER_WORDS = frozenset({"then", "skip", "to"})

_TARGET_RE = re.compile(rf"^{_IDENTIFIER}$")


@dataclass(frozen=True)
class ParsedSkip:
    """One skip rule, normalized.

    `operator` is held **unescaped** (`<`, `<>`, `contains`). Escaping happens
    at the point of emission, so the model never carries XML-encoded data and
    cannot be double-escaped.
    """

    kind: str  # "preskip" | "postskip"
    field: str
    operator: str
    value: str
    target: str


def parse_skip(text: str) -> ParsedSkip | None:
    """Parse one skip line, or return None if it does not match the grammar.

    Callers decide what a None means: the validator reports it against the
    offending field, and the generator treats it as a bug (it should never see
    one, because validation runs first and blocks generation).
    """
    match = SKIP_RE.match(text)
    if match is None:
        return None

    tokens = match.group("target_clause").split()
    target = tokens[-1]
    if not _TARGET_RE.match(target):
        return None
    if any(word.lower() not in _TARGET_FILLER_WORDS for word in tokens[:-1]):
        return None
    # A dangling `, skip to` would otherwise leave "to" as the target, since
    # structurally it is indistinguishable from a target that happens to be
    # named "to". Rejecting is the safe reading: no real field is called
    # to/then/skip, and `_check_skip_to_field_names` would reject one anyway.
    if target.lower() in _TARGET_FILLER_WORDS:
        return None

    # `'does not contain'` and `does  not  contain` both land on the single
    # spelling the app matches against.
    operator = " ".join(match.group("operator").strip("'").lower().split())

    return ParsedSkip(
        kind=match.group("kind").lower(),
        field=match.group("field"),
        operator=operator,
        value=match.group("value"),
        target=target,
    )


def split_skip_lines(text: str) -> list[str]:
    """One Skip cell can hold several rules, one per line."""
    return [line for line in re.split(r"\r\n|\n|\r", text) if line.strip()]
