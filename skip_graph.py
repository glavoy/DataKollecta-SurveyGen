"""What the skips do to a form, taken as a whole.

Every existing check reads one row, or one cell against the list of field
names. That catches a skip that names something misspelled, and it is why
`_check_skip_to_field_names` exists. What no per-row check can see is the
*shape* the skips give the form: whether a condition can ever be true, whether
one is true so often that the question it guards is never asked, whether a rule
sits behind another that always fires first.

Those produce a package that is valid XML, builds without a word, and then
quietly asks the wrong questions in the field. That is the class of defect this
module is for.

## What it does not do

It does not re-derive anything `dd_validators` already reports. A target that
does not exist, is a reserved variable, or resolves to an index at or before
the current question is **already an error** there, and a backward target is
additionally one the app silently ignores at runtime. Such an edge is
classified `IGNORED` here and no finding is raised for it.

That restraint is the point. Two checks that disagree about the same cell is
exactly the failure `skip_parser`'s docstring exists to memorialise -- the
validator accepted a string the generator then dropped, and the log said "No
errors found" while the branching logic did not exist. One parser, and one
place per rule.

## The model

Nodes are the authored questions, `0..n-1`, plus a virtual END. Out of node `i`:

* **preskip** edges, evaluated when navigation *lands* on `i`, before it is
  displayed, in cell order. These clear `[i, target)` -- note the range starts
  at `i`, so a preskip clears its own question's answer.
* if no preskip fires, the **display** case, then **postskip** edges, which
  clear `[i+1, target)`, and the **fall-through** edge to `i+1`.

Both ranges come from the app: `SurveyNavigationService.findNextDisplayedQuestion`
calls `clearAnswersInRange(startIndex: index)` for a preskip and
`advanceFromQuestion` calls it with `currentIndex + 1` for a postskip. `skip to
end` walks and clears `[i, n)` rather than jumping, so trailing system fields
still compute.

Rules are tried in cell order and the first match wins
(`SkipService.evaluateSkips` returns on the first hit), and `xml_generator`
partitions a cell into preskip and postskip elements while preserving that
order -- so shadowing is real, and it is scoped *within* a section: a preskip
cannot shadow a postskip.

## Why the known-domain set is so small

A rule can only be called impossible if the set of values its field can hold is
known exactly. Three cases qualify, and everything else is treated as unknown
and never reported:

* a `radio`/`combobox` whose responses are static, giving the `value:` codes;
* a `text_integer`/`text_decimal` declaring **both** ends of its range;
* neither of the above -- so no claim.

`checkbox` is excluded because its stored value is a comma-joined list, which
is why `contains` is excluded with it. A `source:csv` or `source:database`
question is excluded because its options come from a file this tool does not
read into memory. A half-declared range is excluded because an interval open at
one end contains values this module cannot enumerate.

That narrowness is what lets these be errors rather than warnings. A wider
domain model would find more, and would be wrong often enough that the first
false positive on a real study would get the whole thing switched off.
"""

from __future__ import annotations

from dataclasses import dataclass, field as dataclass_field
from enum import Enum

from cell_text import split_cell_lines
from models import Question, ResponseSourceType
from skip_parser import ParsedSkip, parse_skip, split_skip_lines

# ── The rule catalogue ───────────────────────────────────────────────────────
# Ids are dot-namespaced so they read the same way as the web designer's
# `RULE` catalogue, which makes a later comparison between the two a matter of
# set arithmetic rather than translation.


class Severity(str, Enum):
    ERROR = "ERROR"
    WARNING = "WARNING"


@dataclass(frozen=True)
class Rule:
    id: str
    severity: Severity
    summary: str


RULES: dict[str, Rule] = {
    rule.id: rule
    for rule in (
        Rule(
            "skip.graph.deadBranch",
            Severity.ERROR,
            "The tested field can never hold the value the rule compares against.",
        ),
        Rule(
            "skip.graph.totalCondition",
            Severity.ERROR,
            "The condition is true for every value the tested field can hold, "
            "so the rule is unconditional.",
        ),
        Rule(
            "skip.graph.shadowedRule",
            Severity.ERROR,
            "An earlier rule in the same cell always fires first, so this one "
            "can never be reached.",
        ),
    )
}


@dataclass(frozen=True)
class Finding:
    rule_id: str
    worksheet: str
    field_name: str
    row_index: int
    message: str

    @property
    def severity(self) -> Severity:
        return RULES[self.rule_id].severity

    def format(self) -> str:
        """The log line, in the shape every other check in this tool writes."""
        return f"{self.severity.value} - Skip graph: {self.message}"


# ── Domains ──────────────────────────────────────────────────────────────────
# The special-response codes the app stores when the interviewer presses
# "Don't know" or "Refuse". They are answers a field can hold that its
# Responses cell never lists, so a skip testing one is legitimate and must not
# be called a dead branch.
DONT_KNOW_CODE = "-7"
REFUSE_CODE = "-8"

_TRUTHY = frozenset({"True", "TRUE"})

# Blank cells arrive as this sentinel rather than an empty string.
_UNSET = "-9"

_SELECTION_TYPES = frozenset({"radio", "combobox"})
_NUMERIC_FIELD_TYPES = frozenset({"text_integer", "text_decimal"})


@dataclass(frozen=True)
class Domain:
    """The values a field can hold, when that is knowable.

    Exactly one of `values` (a finite set) or `interval` (inclusive bounds) is
    populated. `None` from `domain_of` means unknown, and no rule fires.
    """

    values: frozenset[str] | None = None
    interval: tuple[float, float] | None = None
    # Present in both shapes: the special codes sit outside a numeric range but
    # are still storable answers.
    specials: frozenset[str] = dataclass_field(default_factory=frozenset)

    def satisfying(self, operator: str, value: str) -> frozenset[str] | None:
        """Which of this domain's values make `<field> <operator> <value>` true.

        `None` means the comparison cannot be evaluated over this domain --
        a numeric operator against a non-numeric literal, say -- which is
        treated exactly like an unknown domain: no finding.
        """
        candidates = self._enumerate()
        if candidates is None:
            return None

        satisfied: set[str] = set()
        for candidate in candidates:
            verdict = _compare(candidate, operator, value)
            if verdict is None:
                return None
            if verdict:
                satisfied.add(candidate)
        return frozenset(satisfied)

    def _enumerate(self) -> frozenset[str] | None:
        if self.values is not None:
            return self.values | self.specials
        if self.interval is not None:
            low, high = self.interval
            # Enumerating an interval is only sound when it is small and
            # integral. A decimal range holds values no loop can list, and a
            # wide one (AVERT declares 0..999999) would cost more than the
            # finding is worth -- so both fall back to interval reasoning in
            # `_compare_interval` instead. See `satisfying_over_interval`.
            if low != int(low) or high != int(high) or high - low > 1000:
                return None
            return frozenset(str(v) for v in range(int(low), int(high) + 1)) | self.specials
        return None

    def describe(self) -> str:
        if self.values is not None:
            listed = ", ".join(sorted(self.values, key=_sort_key))
            extra = _describe_specials(self.specials)
            return f"only {listed}{extra}"
        if self.interval is not None:
            low, high = self.interval
            extra = _describe_specials(self.specials)
            return f"only {_num(low)} to {_num(high)}{extra}"
        return "an unknown set of values"

    def is_empty_after(self, satisfied: frozenset[str]) -> bool:
        """True when nothing in the domain satisfies the condition."""
        return not satisfied

    def is_total(self, satisfied: frozenset[str]) -> bool:
        """True when everything in the domain satisfies the condition."""
        enumerated = self._enumerate()
        return enumerated is not None and satisfied == enumerated


def _describe_specials(specials: frozenset[str]) -> str:
    if not specials:
        return ""
    return " (plus " + ", ".join(sorted(specials, key=_sort_key)) + ")"


def _sort_key(value: str):
    try:
        return (0, float(value), "")
    except ValueError:
        return (1, 0.0, value)


def _num(value: float) -> str:
    return str(int(value)) if value == int(value) else str(value)


def _specials_for(question: Question) -> frozenset[str]:
    codes: set[str] = set()
    if question.dontKnow in _TRUTHY:
        codes.add(DONT_KNOW_CODE)
    if question.refuse in _TRUTHY:
        codes.add(REFUSE_CODE)
    return frozenset(codes)


def static_response_values(question: Question) -> list[str]:
    """The `value:` codes of a static Responses cell, parsed as the emitter does.

    `xml_generator` splits each line at the **first** colon and takes what is
    before it, so a label containing a colon is not a problem and a value
    containing one is impossible. Mirroring that exactly is what stops this
    module deciding a value exists that never reaches the XML.
    """
    values: list[str] = []
    for line in split_cell_lines(question.responses):
        index = line.find(":")
        if index < 0:
            continue
        values.append(line[:index].strip())
    return values


def domain_of(question: Question) -> Domain | None:
    """The values `question` can hold, or None when that is not knowable."""
    if question.questionType in _SELECTION_TYPES:
        if question.responseSourceType != ResponseSourceType.STATIC:
            return None
        values = [v for v in static_response_values(question) if v != ""]
        if not values:
            return None
        return Domain(values=frozenset(values), specials=_specials_for(question))

    if question.questionType == "text" and question.fieldType in _NUMERIC_FIELD_TYPES:
        if question.lowerRange in ("", _UNSET) or question.upperRange in ("", _UNSET):
            return None
        try:
            low = float(question.lowerRange)
            high = float(question.upperRange)
        except ValueError:
            return None
        if low > high:
            # A range the wrong way round is `_check_ranges`' finding, not
            # this module's; reasoning over it would produce a second, more
            # confusing report of the same cell.
            return None
        return Domain(interval=(low, high), specials=_specials_for(question))

    return None


# ── Comparison ───────────────────────────────────────────────────────────────
# Mirrors `FieldComparator.compare` in the app: numeric when both sides parse
# as numbers, string otherwise. The word operators are excluded before we get
# here, because they only apply to checkbox values whose domain is unknown.
_NUMERIC_OPERATORS = frozenset({">", ">=", "<", "<="})
_EQUALITY_OPERATORS = frozenset({"=", "<>", "!="})
_WORD_OPERATORS = frozenset({"contains", "does not contain"})


def _compare(actual: str, operator: str, expected: str) -> bool | None:
    """Whether `actual <operator> expected` holds. None when undecidable."""
    if operator in _EQUALITY_OPERATORS:
        equal = _values_equal(actual, expected)
        if equal is None:
            return None
        return equal if operator == "=" else not equal

    if operator in _NUMERIC_OPERATORS:
        try:
            left = float(actual)
            right = float(expected)
        except ValueError:
            # The app compares as strings here rather than refusing, but a
            # string ordering over response codes is not something to raise an
            # error from -- so this stays undecidable and nothing is reported.
            return None
        if operator == ">":
            return left > right
        if operator == ">=":
            return left >= right
        if operator == "<":
            return left < right
        return left <= right

    return None


def _values_equal(actual: str, expected: str) -> bool | None:
    try:
        return float(actual) == float(expected)
    except ValueError:
        return actual == expected


# ── Edges ────────────────────────────────────────────────────────────────────


class EdgeClass(str, Enum):
    LIVE = "live"
    DEAD = "dead"
    TOTAL = "total"
    IGNORED = "ignored"


@dataclass(frozen=True)
class SkipEdge:
    source: int
    parsed: ParsedSkip
    order: int
    """Position within its own section of the cell, so shadowing can be scoped."""
    target: int | None
    """Index of the target question; None for `skip to end`; -1 when unresolved."""
    edge_class: EdgeClass
    satisfying: frozenset[str] | None

    @property
    def is_preskip(self) -> bool:
        return self.parsed.kind == "preskip"

    @property
    def is_end(self) -> bool:
        return self.target is None


END_TARGET = "end"


def _classify(
    parsed: ParsedSkip,
    source_index: int,
    target: int | None,
    tested: Question | None,
) -> tuple[EdgeClass, frozenset[str] | None]:
    if target is not None and target <= source_index:
        # Already an error in `_check_skip_to_field_names`, and silently
        # dropped by the app at runtime (`advanceFromQuestion` only jumps when
        # targetIndex > currentIndex). Nothing more to say about it here.
        return EdgeClass.IGNORED, None

    if tested is None:
        return EdgeClass.IGNORED, None

    if parsed.operator in _WORD_OPERATORS:
        return EdgeClass.LIVE, None

    domain = domain_of(tested)
    if domain is None:
        return EdgeClass.LIVE, None

    satisfied = domain.satisfying(parsed.operator, parsed.value)
    if satisfied is None:
        return EdgeClass.LIVE, None

    if domain.is_empty_after(satisfied):
        return EdgeClass.DEAD, satisfied
    if domain.is_total(satisfied):
        return EdgeClass.TOTAL, satisfied
    return EdgeClass.LIVE, satisfied


@dataclass
class FormGraph:
    """One questionnaire's questions and the edges its skips create."""

    worksheet: str
    questions: list[Question]
    edges: list[SkipEdge]

    @property
    def index_of(self) -> dict[str, int]:
        return {q.fieldName: i for i, q in enumerate(self.questions)}


def build_graph(worksheet: str, questions: list[Question]) -> FormGraph:
    index_of = {q.fieldName: i for i, q in enumerate(questions)}
    edges: list[SkipEdge] = []

    for source_index, question in enumerate(questions):
        if not question.skip:
            continue

        # Order is counted per section, because the app evaluates preskips and
        # postskips at two different moments and the first match wins within
        # each. A preskip listed after a postskip in the same cell is still the
        # first preskip.
        order_in_section = {"preskip": 0, "postskip": 0}

        for line in split_skip_lines(question.skip):
            parsed = parse_skip(line)
            if parsed is None:
                # `_check_skip_syntax` reported it and the worksheet gate means
                # generation is already blocked.
                continue

            if parsed.target.lower() == END_TARGET:
                target: int | None = None
            else:
                target = index_of.get(parsed.target, -1)
                if target == -1:
                    # Nonexistent target: already an error elsewhere.
                    edges.append(
                        SkipEdge(
                            source=source_index,
                            parsed=parsed,
                            order=order_in_section[parsed.kind],
                            target=-1,
                            edge_class=EdgeClass.IGNORED,
                            satisfying=None,
                        )
                    )
                    order_in_section[parsed.kind] += 1
                    continue

            tested = questions[index_of[parsed.field]] if parsed.field in index_of else None
            edge_class, satisfied = _classify(parsed, source_index, target, tested)

            edges.append(
                SkipEdge(
                    source=source_index,
                    parsed=parsed,
                    order=order_in_section[parsed.kind],
                    target=target,
                    edge_class=edge_class,
                    satisfying=satisfied,
                )
            )
            order_in_section[parsed.kind] += 1

    return FormGraph(worksheet=worksheet, questions=questions, edges=edges)


# ── Rules ────────────────────────────────────────────────────────────────────


def _where(graph: FormGraph, index: int) -> str:
    q = graph.questions[index]
    return f"row {q.rowIndex}" if q.rowIndex else f"FieldName '{q.fieldName}'"


def _guarded_range(graph: FormGraph, edge: SkipEdge) -> str:
    """The questions a firing edge jumps over, for a message."""
    start = edge.source if edge.is_preskip else edge.source + 1
    end = len(graph.questions) if edge.is_end else edge.target
    if end is None or end <= start:
        return ""
    names = [graph.questions[i].fieldName for i in range(start, end)]
    if len(names) == 1:
        return f"'{names[0]}'"
    return f"'{names[0]}' through '{names[-1]}' ({len(names)} questions)"


def _check_dead_branches(graph: FormGraph) -> list[Finding]:
    findings: list[Finding] = []
    for edge in graph.edges:
        if edge.edge_class is not EdgeClass.DEAD:
            continue
        source = graph.questions[edge.source]
        tested = graph.questions[graph.index_of[edge.parsed.field]]
        domain = domain_of(tested)
        skipped = _guarded_range(graph, edge)
        consequence = (
            f" so {skipped} is asked of everyone" if skipped else " so it never fires"
        )
        findings.append(
            Finding(
                rule_id="skip.graph.deadBranch",
                worksheet=graph.worksheet,
                field_name=source.fieldName,
                row_index=source.rowIndex,
                message=(
                    f"In worksheet '{graph.worksheet}', the {edge.parsed.kind} for "
                    f"FieldName '{source.fieldName}' ({_where(graph, edge.source)}) tests "
                    f"'{edge.parsed.field} {edge.parsed.operator} {edge.parsed.value}', but "
                    f"'{edge.parsed.field}' ({_where(graph, graph.index_of[edge.parsed.field])}) "
                    f"can hold {domain.describe() if domain else 'other values'}. "
                    f"The condition can never be true,{consequence}."
                ),
            )
        )
    return findings


def _check_total_conditions(graph: FormGraph) -> list[Finding]:
    findings: list[Finding] = []
    for edge in graph.edges:
        if edge.edge_class is not EdgeClass.TOTAL:
            continue
        source = graph.questions[edge.source]
        tested = graph.questions[graph.index_of[edge.parsed.field]]
        domain = domain_of(tested)
        skipped = _guarded_range(graph, edge)
        consequence = (
            f" {skipped} is never asked" if skipped else " the jump always happens"
        )
        # A preskip on the tested question's own field is impossible (checked
        # elsewhere), so the only self-referential shape reaching here is a
        # postskip, where "unconditional" means the question is answered and
        # then the jump always happens.
        findings.append(
            Finding(
                rule_id="skip.graph.totalCondition",
                worksheet=graph.worksheet,
                field_name=source.fieldName,
                row_index=source.rowIndex,
                message=(
                    f"In worksheet '{graph.worksheet}', the {edge.parsed.kind} for "
                    f"FieldName '{source.fieldName}' ({_where(graph, edge.source)}) tests "
                    f"'{edge.parsed.field} {edge.parsed.operator} {edge.parsed.value}', which is "
                    f"true for every value '{edge.parsed.field}' can hold "
                    f"({domain.describe() if domain else 'its whole range'}). "
                    f"The rule is unconditional, so{consequence}."
                ),
            )
        )
    return findings


def _check_shadowed_rules(graph: FormGraph) -> list[Finding]:
    """A rule an earlier one in the same section always pre-empts.

    Three decidable shapes, and nothing else. Two rules on *different* fields
    can only be compared by reasoning about which paths reach the question,
    which is not decidable here and would produce false positives on exactly
    the dictionaries this has to stay quiet on.
    """
    findings: list[Finding] = []
    by_source_section: dict[tuple[int, str], list[SkipEdge]] = {}
    for edge in graph.edges:
        if edge.edge_class is EdgeClass.IGNORED:
            continue
        by_source_section.setdefault((edge.source, edge.parsed.kind), []).append(edge)

    for (source_index, kind), edges in by_source_section.items():
        edges = sorted(edges, key=lambda e: e.order)
        source = graph.questions[source_index]
        for position, edge in enumerate(edges):
            for earlier in edges[:position]:
                reason = _shadow_reason(earlier, edge)
                if reason is None:
                    continue
                findings.append(
                    Finding(
                        rule_id="skip.graph.shadowedRule",
                        worksheet=graph.worksheet,
                        field_name=source.fieldName,
                        row_index=source.rowIndex,
                        message=(
                            f"In worksheet '{graph.worksheet}', FieldName "
                            f"'{source.fieldName}' ({_where(graph, source_index)}) has a "
                            f"{kind} rule that can never be reached: "
                            f"'if {edge.parsed.field} {edge.parsed.operator} "
                            f"{edge.parsed.value}' sits after "
                            f"'if {earlier.parsed.field} {earlier.parsed.operator} "
                            f"{earlier.parsed.value}', which {reason}. Rules are tried in "
                            "cell order and the first match wins -- put the narrower rule "
                            "first."
                        ),
                    )
                )
                break
    return findings


def _shadow_reason(earlier: SkipEdge, later: SkipEdge) -> str | None:
    if earlier.edge_class is EdgeClass.TOTAL:
        return "is true for every value its field can hold"

    if earlier.parsed.field != later.parsed.field:
        return None

    if (
        earlier.parsed.operator == later.parsed.operator
        and earlier.parsed.value == later.parsed.value
    ):
        if earlier.parsed.target == later.parsed.target:
            return "is the same rule"
        return f"is the same condition, skipping to '{earlier.parsed.target}' instead"

    if earlier.satisfying is None or later.satisfying is None:
        return None
    if later.satisfying and later.satisfying <= earlier.satisfying:
        return "already fires for every value that would satisfy it"
    return None


def lint_form(worksheet: str, questions: list[Question]) -> list[Finding]:
    """Every finding for one questionnaire."""
    graph = build_graph(worksheet, questions)
    findings: list[Finding] = []
    findings.extend(_check_dead_branches(graph))
    findings.extend(_check_total_conditions(graph))
    findings.extend(_check_shadowed_rules(graph))
    return sorted(findings, key=lambda f: (f.row_index, f.rule_id))
