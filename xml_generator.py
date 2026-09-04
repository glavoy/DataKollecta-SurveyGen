from __future__ import annotations

import re
from pathlib import Path

from models import (
    LEADING_SYSTEM_FIELDS,
    RESERVED_SYSTEM_FIELDS,
    TRAILING_SYSTEM_FIELDS,
    CalculationPart,
    CalculationType,
    Question,
    ResponseSourceType,
)
from skip_parser import ParsedSkip, parse_skip, split_skip_lines
from cell_text import split_cell_lines

# The Windows app writes files with .NET StreamWriter on Windows: every WriteLine
# ends with CRLF, and WriteLine("\n") emits "\n\r\n". NEWLINE/BLANK_LINE reproduce
# those exact byte sequences so output diffs cleanly against the original app.
NEWLINE = "\r\n"
BLANK_LINE = "\n\r\n"


# A character reference that is already written out: `&amp;`, `&lt;`, `&#233;`,
# `&#x2019;`. Authors have been hand-escaping their cells to work around the
# missing escaping -- the AVERT French dictionary writes `&lt;14 ou &gt;35`
# directly into a logic-check message -- so escaping has to leave an existing
# reference alone or it becomes `&amp;lt;14`, which then renders as literal
# "&lt;14" on the device.
_CHARACTER_REFERENCE_RE = re.compile(r"&(?:#[0-9]+|#[xX][0-9A-Fa-f]+|[A-Za-z][A-Za-z0-9]*);")


def _escape_ampersands(text: str) -> str:
    """Escape every `&` that does not already begin a character reference."""
    out: list[str] = []
    position = 0
    while True:
        found = text.find("&", position)
        if found < 0:
            out.append(text[position:])
            return "".join(out)
        out.append(text[position:found])
        existing = _CHARACTER_REFERENCE_RE.match(text, found)
        if existing:
            out.append(existing.group(0))
            position = existing.end()
        else:
            out.append("&amp;")
            position = found + 1


def _esc_text(value: object) -> str:
    """Escape a value written as element text.

    Everything a data dictionary author types goes through this or
    `_esc_attr`. Before they existed every value was interpolated raw, so an
    `&` in a response label or a `<` in a question text produced XML that
    `ET.parse` rejects -- and the reported error named a line and column, not
    the field, leaving the author to find it by hand in a 400-row sheet.

    Idempotent: text that is already escaped survives unchanged, so this can
    be applied to existing dictionaries without rewriting their cells.
    """
    text = _escape_ampersands(str(value))
    return text.replace("<", "&lt;").replace(">", "&gt;")


def _esc_attr(value: object) -> str:
    """Escape a value written inside an attribute.

    Escapes both quote characters, so it is safe regardless of whether the
    attribute is single-quoted (nearly all of them) or double-quoted (`mask`).
    An apostrophe was the most likely break in practice: the French
    dictionaries put `d'accord` and `n'a pas` in `dont_know`/`not_in_list`
    labels, which are emitted as `label='...'`.
    """
    return _esc_text(value).replace("'", "&apos;").replace('"', "&quot;")


class XmlGenerator:
    def __init__(self) -> None:
        self.logstring: list[str] = []

    def write_xml(self, worksheet_name: str, question_list: list[Question], xml_path: Path) -> Path:
        if worksheet_name.endswith("_dd"):
            xml_name = worksheet_name[:-3]
        else:
            xml_name = worksheet_name[:-4]

        out_file = xml_path / f"{xml_name}.xml"
        with out_file.open("w", encoding="utf-8", newline="") as f:

            def wl(line: str = "") -> None:
                f.write(line + NEWLINE)

            def write_system_question(fieldname: str, fieldtype: str) -> None:
                wl(
                    f"\t<question type = 'automatic' fieldname = '{fieldname}' "
                    f"fieldtype = '{fieldtype}'>"
                )
                wl("\t</question>")
                f.write(BLANK_LINE)

            wl("<?xml version = '1.0' encoding = 'utf-8'?>")
            wl("<survey>")
            f.write(BLANK_LINE)

            # Written here rather than wherever the dictionary happened to
            # declare them, so they always record the right moment.
            for fieldname, fieldtype in LEADING_SYSTEM_FIELDS:
                write_system_question(fieldname, fieldtype)

            for q in question_list:
                # Supplied above and below; a declared row is dropped so the
                # questionnaire cannot end up with two of them.
                if q.fieldName.lower() in RESERVED_SYSTEM_FIELDS:
                    continue

                wl(
                    f"\t<question type = '{_esc_attr(q.questionType)}' "
                    f"fieldname = '{_esc_attr(q.fieldName)}' fieldtype = '{_esc_attr(q.fieldType)}'>"
                )

                if q.questionType != "automatic":
                    wl(f"\t\t<text>{_esc_text(q.questionText)}</text>")

                if q.questionType == "automatic" and q.calculationType != CalculationType.NONE:
                    self._generate_calculation_xml(wl, q)

                if q.maxCharacters != "-9":
                    wl(f"\t\t<maxCharacters>{_esc_text(q.maxCharacters)}</maxCharacters>")

                if q.mask:
                    wl(f"\t\t<mask value=\"{_esc_attr(q.mask)}\" />")

                if q.uniqueCheckMessage:
                    wl("\t\t<unique_check>")
                    wl(f"\t\t\t<message>{_esc_text(q.uniqueCheckMessage)}</message>")
                    wl("\t\t</unique_check>")

                if q.questionType != "date" and q.lowerRange != "-9":
                    lower = _esc_attr(q.lowerRange)
                    upper = _esc_attr(q.upperRange)
                    wl("\t\t<numeric_check>")
                    wl(
                        f"\t\t\t<values minvalue ='{lower}' maxvalue='{upper}' other_values = '{lower}' "
                        f"message = 'Number must be between {lower} and {upper}!'></values>"
                    )
                    wl("\t\t</numeric_check>")

                if q.questionType == "date":
                    wl("\t\t<date_range>")
                    wl(f"\t\t\t<min_date>{_esc_text(q.lowerRange)}</min_date>")
                    wl(f"\t\t\t<max_date>{_esc_text(q.upperRange)}</max_date>")
                    wl("\t\t</date_range>")

                for logic_check in q.logicChecks:
                    wl("\t\t<logic_check>")
                    wl(self._generate_logic_check(logic_check))
                    wl("\t\t</logic_check>")

                if q.questionType in {"radio", "checkbox", "combobox"}:
                    attrs = ""
                    if q.responseSourceType == ResponseSourceType.CSV:
                        attrs += f" source='csv' file='{_esc_attr(q.responseSourceFile)}'"
                    elif q.responseSourceType == ResponseSourceType.DATABASE:
                        attrs += f" source='database' table='{_esc_attr(q.responseSourceTable)}'"
                    wl(f"\t\t<responses{attrs}>")

                    for flt in q.responseFilters:
                        # `flt.operator` is deliberately NOT escaped: the
                        # dictionary's operators are already encoded into
                        # entities by `_parse_operator` when they are read, so
                        # escaping again would produce `&amp;lt;`.
                        wl(
                            f"\t\t\t<filter column='{_esc_attr(flt.column)}' "
                            f"operator='{flt.operator}' value='{_esc_attr(flt.value)}'/>"
                        )
                    if q.responseDisplayColumn:
                        wl(f"\t\t\t<display column='{_esc_attr(q.responseDisplayColumn)}'/>")
                    if q.responseValueColumn:
                        wl(f"\t\t\t<value column='{_esc_attr(q.responseValueColumn)}'/>")
                    if q.responseDistinct is not None:
                        wl(f"\t\t\t<distinct>{str(q.responseDistinct).lower()}</distinct>")
                    if q.responseEmptyMessage:
                        wl(f"\t\t\t<empty_message>{_esc_text(q.responseEmptyMessage)}</empty_message>")
                    if q.responseDontKnowValue:
                        label_attr = (
                            f" label='{_esc_attr(q.responseDontKnowLabel)}'" if q.responseDontKnowLabel else ""
                        )
                        wl(f"\t\t\t<dont_know value='{_esc_attr(q.responseDontKnowValue)}'{label_attr}/>")
                    if q.responseNotInListValue:
                        label_attr = (
                            f" label='{_esc_attr(q.responseNotInListLabel)}'" if q.responseNotInListLabel else ""
                        )
                        wl(f"\t\t\t<not_in_list value='{_esc_attr(q.responseNotInListValue)}'{label_attr}/>")

                    if q.responseSourceType == ResponseSourceType.STATIC:
                        responses = split_cell_lines(q.responses)
                        if len(responses) == 0:
                            wl("\t\t\t<response></response>")
                        else:
                            for response in responses:
                                index = response.find(":")
                                value = response[:index]
                                label = response[index + 1 :].strip()
                                wl(
                                    f"\t\t\t<response value = '{_esc_attr(value)}'>"
                                    f"{_esc_text(label)}</response>"
                                )
                    wl("\t\t</responses>")

                if q.skip:
                    parsed_skips: list[ParsedSkip] = []
                    for line in split_skip_lines(q.skip):
                        parsed = parse_skip(line)
                        if parsed is None:
                            # Validation rejects these before generation ever
                            # runs, so reaching here means the validator and
                            # this module have drifted apart again. Fail loudly
                            # rather than silently dropping the rule, which is
                            # exactly what the old two-parser split did.
                            raise ValueError(
                                f"Unparseable Skip on field '{q.fieldName}' "
                                f"in worksheet '{worksheet_name}': {line!r}"
                            )
                        parsed_skips.append(parsed)

                    pre = [s for s in parsed_skips if s.kind == "preskip"]
                    post = [s for s in parsed_skips if s.kind == "postskip"]
                    if pre:
                        wl("\t\t<preskip>")
                        for s in pre:
                            wl(self._generate_skip(s))
                        wl("\t\t</preskip>")
                    if post:
                        wl("\t\t<postskip>")
                        for s in post:
                            wl(self._generate_skip(s))
                        wl("\t\t</postskip>")

                if q.dontKnow in {"TRUE", "True"}:
                    wl("\t\t<dont_know>-7</dont_know>")
                if q.refuse in {"TRUE", "True"}:
                    wl("\t\t<refuse>-8</refuse>")
                if q.optional in {"TRUE", "True"}:
                    wl("\t\t<optional>1</optional>")

                wl("\t</question>")
                f.write(BLANK_LINE)

            # Ahead of the end-of-survey screen: navigation stops on that
            # screen, so anything after it is never computed.
            for fieldname, fieldtype in TRAILING_SYSTEM_FIELDS:
                write_system_question(fieldname, fieldtype)

            wl("\t<question type = 'information' fieldname = 'end_of_questions' fieldtype = 'n/a'>")
            wl("\t\t<text>Press the 'Finish' button to save the data.</text >")
            wl("\t</question>")
            f.write(BLANK_LINE)
            wl("</survey>")

        return out_file

    def _generate_skip(self, skip: ParsedSkip) -> str:
        # The operator is the only pre-escaped-looking value here, and it comes
        # from the parser's closed vocabulary rather than free text, so `<`/`>`
        # are the only characters that can need encoding. Everything else on
        # the line is escaped as an ordinary attribute.
        condition = skip.operator.replace("<", "&lt;").replace(">", "&gt;")
        return (
            f"\t\t\t<skip fieldname='{_esc_attr(skip.field)}' condition = '{condition}' "
            f"response='{_esc_attr(skip.value)}' "
            f"response_type='fixed' skiptofieldname ='{_esc_attr(skip.target)}'></skip>"
        )

    def _generate_logic_check(self, logic_check: str) -> str:
        expression, message = [p.strip() for p in logic_check.split(";", 1)]
        # The expression's own comparison operators are encoded below, so it
        # must not go through `_esc_text` as well. The message is free text and
        # must.
        message = _esc_text(message)
        expression = expression.replace("!=", "&lt;&gt;")
        expression = expression.replace("<>", "&lt;&gt;")
        expression = expression.replace("<=", "&lt;=")
        expression = expression.replace(">=", "&gt;=")
        expression = re.sub(r"(?<!&lt;)(?<!&gt;)<(?!=)", "&lt;", expression)
        expression = re.sub(r"(?<!&lt;=)(?<!&gt;=)>(?!=)", "&gt;", expression)

        if " or " in expression:
            # Mirrors the C# StringBuilder layout: each 'or' clause on its own line,
            # ';' directly after the last clause, message on the following line.
            parts = expression.split(" or ")
            result = ""
            for i, part in enumerate(parts):
                result += "\t\t\t" + part.strip()
                if i < len(parts) - 1:
                    result += " or" + NEWLINE
            result += ";" + NEWLINE
            result += "\t\t\t" + message
            return result

        return f"\t\t\t{expression}; {message}"

    def _generate_calculation_xml(self, wl, q: Question) -> None:
        # `_convert_operator_to_xml` already returns entity-encoded operators,
        # so `op` below is the one value here that must not be escaped again.
        field = _esc_attr(q.calculationLookupField)
        constant = _esc_attr(q.calculationConstantValue)
        unit = _esc_attr(q.calculationUnit)

        if q.calculationType == CalculationType.QUERY:
            wl("\t\t<calculation type='query'>")
            wl(f"\t\t\t<sql>{_esc_text(q.calculationQuerySql)}</sql>")
            for param in q.calculationQueryParameters:
                wl(
                    f"\t\t\t<parameter name='{_esc_attr(param.name)}' "
                    f"field='{_esc_attr(param.fieldName)}' />"
                )
            wl("\t\t</calculation>")
        elif q.calculationType == CalculationType.CASE:
            wl("\t\t<calculation type='case'>")
            for cond in q.calculationCaseConditions:
                op = self._convert_operator_to_xml(cond.operator)
                wl(
                    f"\t\t\t<when field='{_esc_attr(cond.field)}' operator='{op}' "
                    f"value='{_esc_attr(cond.value)}'>"
                )
                if cond.result:
                    self._generate_calculation_part(wl, cond.result, 4)
                wl("\t\t\t</when>")
            if q.calculationCaseElse:
                wl("\t\t\t<else>")
                self._generate_calculation_part(wl, q.calculationCaseElse, 4)
                wl("\t\t\t</else>")
            wl("\t\t</calculation>")
        elif q.calculationType == CalculationType.CONSTANT:
            wl(f"\t\t<calculation type='constant' value='{constant}' />")
        elif q.calculationType == CalculationType.LOOKUP:
            wl(f"\t\t<calculation type='lookup' field='{field}' />")
        elif q.calculationType == CalculationType.MATH:
            wl(f"\t\t<calculation type='math' operator='{_esc_attr(q.calculationMathOperator)}'>")
            for part in q.calculationMathParts:
                self._generate_calculation_part(wl, part, 3)
            wl("\t\t</calculation>")
        elif q.calculationType == CalculationType.CONCAT:
            separator_attr = (
                f" separator='{_esc_attr(q.calculationConcatSeparator)}'" if q.calculationConcatSeparator else ""
            )
            wl(f"\t\t<calculation type='concat'{separator_attr}>")
            for part in q.calculationConcatParts:
                self._generate_calculation_part(wl, part, 3)
            wl("\t\t</calculation>")
        elif q.calculationType == CalculationType.AGE_FROM_DATE:
            wl(f"\t\t<calculation type='age_from_date' field='{field}' value='{constant}'/>")
        elif q.calculationType == CalculationType.AGE_AT_DATE:
            separator_attr = (
                f" separator='{_esc_attr(q.calculationConcatSeparator)}'" if q.calculationConcatSeparator else ""
            )
            wl(
                f"\t\t<calculation type='age_at_date' field='{field}' value='{constant}'{separator_attr}/>"
            )
        elif q.calculationType == CalculationType.DATE_OFFSET:
            wl(f"\t\t<calculation type='date_offset' field='{field}' value='{constant}' />")
        elif q.calculationType == CalculationType.DATE_DIFF:
            wl(
                f"\t\t<calculation type='date_diff' field='{field}' value='{constant}' unit='{unit}' />"
            )
        elif q.calculationType == CalculationType.DATE_PART:
            wl(f"\t\t<calculation type='date_part' field='{field}' unit='{unit}' />")
        elif q.calculationType == CalculationType.TIMESTAMP:
            wl("\t\t<calculation type='timestamp' preserve='true' />")
        else:
            # There was no final branch here at all, so a CalculationType this
            # chain does not know about produced NO <calculation> element and
            # reported success -- the question would validate clean, ship, and
            # simply never compute anything in the field. That is the same
            # failure shape as C4's vanished skip element, and the same answer:
            # state the invariant rather than imply it. Reaching here means a
            # type was added to the enum and to the validator but not to this
            # emitter (see tests/test_calculation_registry.py).
            raise ValueError(
                f"No XML emitter for calculation type {q.calculationType!r} on field "
                f"'{q.fieldName}'. Add a branch here, or the calculation is silently dropped."
            )

    def _generate_calculation_part(self, wl, part: CalculationPart, indent_level: int) -> None:
        indent = "\t" * indent_level
        if part.type == CalculationType.CONSTANT:
            wl(f"{indent}<result type='constant' value='{_esc_attr(part.constantValue)}' />")
        elif part.type == CalculationType.LOOKUP:
            wl(f"{indent}<part type='lookup' field='{_esc_attr(part.lookupField)}' />")
        elif part.type == CalculationType.QUERY:
            wl(f"{indent}<part type='query'>")
            wl(f"{indent}\t<sql>{_esc_text(part.querySql)}</sql>")
            for param in part.queryParameters:
                wl(
                    f"{indent}\t<parameter name='{_esc_attr(param.name)}' "
                    f"field='{_esc_attr(param.fieldName)}' />"
                )
            wl(f"{indent}</part>")
        elif part.type == CalculationType.MATH:
            wl(f"{indent}<part type='math' operator='{_esc_attr(part.mathOperator)}'>")
            for nested in part.parts:
                self._generate_calculation_part(wl, nested, indent_level + 1)
            wl(f"{indent}</part>")
        elif part.type == CalculationType.CONCAT:
            separator_attr = f" separator='{_esc_attr(part.concatSeparator)}'" if part.concatSeparator else ""
            wl(f"{indent}<part type='concat'{separator_attr}>")
            for nested in part.parts:
                self._generate_calculation_part(wl, nested, indent_level + 1)
            wl(f"{indent}</part>")

    @staticmethod
    def _convert_operator_to_xml(op: str) -> str:
        """Escape a `when` condition's operator for the XML attribute.

        `<>` becomes `!=` rather than `&lt;&gt;`, so only one canonical
        spelling of "not equal" ever reaches the app, even though the app's
        comparator accepts both as synonyms on its own. `contains` / `does
        not contain` pass straight through, in whatever case and internal
        spacing WHEN_CONDITION_RE let through -- normalized to lowercase with
        single spaces first, so "Contains" and "does  not contain" both land
        on the one spelling the app matches against.
        """
        op = " ".join(op.strip().lower().split())
        converted = {
            "=": "=",
            "!=": "!=",
            "<>": "!=",
            ">": "&gt;",
            "<": "&lt;",
            ">=": "&gt;=",
            "<=": "&lt;=",
            "contains": "contains",
            "does not contain": "does not contain",
        }.get(op)
        if converted is None:
            raise ValueError(f"Unsupported operator in a when condition: {op!r}")
        return converted
