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

# The Windows app writes files with .NET StreamWriter on Windows: every WriteLine
# ends with CRLF, and WriteLine("\n") emits "\n\r\n". NEWLINE/BLANK_LINE reproduce
# those exact byte sequences so output diffs cleanly against the original app.
NEWLINE = "\r\n"
BLANK_LINE = "\n\r\n"


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

                wl(f"\t<question type = '{q.questionType}' fieldname = '{q.fieldName}' fieldtype = '{q.fieldType}'>")

                if q.questionType != "automatic":
                    wl(f"\t\t<text>{q.questionText}</text>")

                if q.questionType == "automatic" and q.calculationType != CalculationType.NONE:
                    self._generate_calculation_xml(wl, q)

                if q.maxCharacters != "-9":
                    wl(f"\t\t<maxCharacters>{q.maxCharacters}</maxCharacters>")

                if q.mask:
                    wl(f"\t\t<mask value=\"{q.mask}\" />")

                if q.uniqueCheckMessage:
                    wl("\t\t<unique_check>")
                    wl(f"\t\t\t<message>{q.uniqueCheckMessage}</message>")
                    wl("\t\t</unique_check>")

                if q.questionType != "date" and q.lowerRange != "-9":
                    wl("\t\t<numeric_check>")
                    wl(
                        f"\t\t\t<values minvalue ='{q.lowerRange}' maxvalue='{q.upperRange}' other_values = '{q.lowerRange}' "
                        f"message = 'Number must be between {q.lowerRange} and {q.upperRange}!'></values>"
                    )
                    wl("\t\t</numeric_check>")

                if q.questionType == "date":
                    wl("\t\t<date_range>")
                    wl(f"\t\t\t<min_date>{q.lowerRange}</min_date>")
                    wl(f"\t\t\t<max_date>{q.upperRange}</max_date>")
                    wl("\t\t</date_range>")

                for logic_check in q.logicChecks:
                    wl("\t\t<logic_check>")
                    wl(self._generate_logic_check(logic_check))
                    wl("\t\t</logic_check>")

                if q.questionType in {"radio", "checkbox", "combobox"}:
                    attrs = ""
                    if q.responseSourceType == ResponseSourceType.CSV:
                        attrs += f" source='csv' file='{q.responseSourceFile}'"
                    elif q.responseSourceType == ResponseSourceType.DATABASE:
                        attrs += f" source='database' table='{q.responseSourceTable}'"
                    wl(f"\t\t<responses{attrs}>")

                    for flt in q.responseFilters:
                        wl(f"\t\t\t<filter column='{flt.column}' operator='{flt.operator}' value='{flt.value}'/>")
                    if q.responseDisplayColumn:
                        wl(f"\t\t\t<display column='{q.responseDisplayColumn}'/>")
                    if q.responseValueColumn:
                        wl(f"\t\t\t<value column='{q.responseValueColumn}'/>")
                    if q.responseDistinct is not None:
                        wl(f"\t\t\t<distinct>{str(q.responseDistinct).lower()}</distinct>")
                    if q.responseEmptyMessage:
                        wl(f"\t\t\t<empty_message>{q.responseEmptyMessage}</empty_message>")
                    if q.responseDontKnowValue:
                        label_attr = f" label='{q.responseDontKnowLabel}'" if q.responseDontKnowLabel else ""
                        wl(f"\t\t\t<dont_know value='{q.responseDontKnowValue}'{label_attr}/>")
                    if q.responseNotInListValue:
                        label_attr = f" label='{q.responseNotInListLabel}'" if q.responseNotInListLabel else ""
                        wl(f"\t\t\t<not_in_list value='{q.responseNotInListValue}'{label_attr}/>")

                    if q.responseSourceType == ResponseSourceType.STATIC:
                        responses = [r for r in re.split(r"\r\n|\n|\r", q.responses) if r]
                        if len(responses) == 0:
                            wl("\t\t\t<response></response>")
                        else:
                            for response in responses:
                                index = response.find(":")
                                value = response[:index]
                                label = response[index + 1 :].strip()
                                wl(f"\t\t\t<response value = '{value}'>{label}</response>")
                    wl("\t\t</responses>")

                if q.skip:
                    skips = [s for s in re.split(r"\r\n|\n|\r", q.skip) if s]
                    pre = [s for s in skips if s[: s.find(":")] == "preskip"]
                    post = [s for s in skips if s[: s.find(":")] == "postskip"]
                    if pre:
                        wl("\t\t<preskip>")
                        for s in pre:
                            wl(self._generate_skip(s, "preSkip"))
                        wl("\t\t</preskip>")
                    if post:
                        wl("\t\t<postskip>")
                        for s in post:
                            wl(self._generate_skip(s, "postSkip"))
                        wl("\t\t</postskip>")

                if q.dontKnow in {"TRUE", "True"}:
                    wl("\t\t<dont_know>-7</dont_know>")
                if q.refuse in {"TRUE", "True"}:
                    wl("\t\t<refuse>-8</refuse>")
                if q.na in {"TRUE", "True"}:
                    wl("\t\t<na>-6</na>")

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

    def _generate_skip(self, skip: str, skip_type: str) -> str:
        len_skip = 13 if skip_type == "postSkip" else 12
        space_indices = [i for i, ch in enumerate(skip) if ch == " "]
        fieldname_to_check = skip[len_skip : space_indices[2]]

        if len(space_indices) == 9:
            condition = "does not contain"
            value = skip[space_indices[5] + 1 : space_indices[6] - 1]
        elif "contains" in skip:
            condition = "contains"
            value = skip[space_indices[3] + 1 : space_indices[4] - 1]
        else:
            condition = skip[space_indices[2] + 1 : space_indices[3]]
            condition = condition.replace("<", "&lt;").replace(">", "&gt;")
            value = skip[space_indices[3] + 1 : space_indices[4] - 1]

        fieldname_to_skip_to = skip[space_indices[-1] + 1 :]
        return (
            f"\t\t\t<skip fieldname='{fieldname_to_check}' condition = '{condition}' response='{value}' "
            f"response_type='fixed' skiptofieldname ='{fieldname_to_skip_to}'></skip>"
        )

    def _generate_logic_check(self, logic_check: str) -> str:
        expression, message = [p.strip() for p in logic_check.split(";", 1)]
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
        if q.calculationType == CalculationType.QUERY:
            wl("\t\t<calculation type='query'>")
            wl(f"\t\t\t<sql>{q.calculationQuerySql}</sql>")
            for param in q.calculationQueryParameters:
                wl(f"\t\t\t<parameter name='{param.name}' field='{param.fieldName}' />")
            wl("\t\t</calculation>")
        elif q.calculationType == CalculationType.CASE:
            wl("\t\t<calculation type='case'>")
            for cond in q.calculationCaseConditions:
                op = self._convert_operator_to_xml(cond.operator)
                wl(f"\t\t\t<when field='{cond.field}' operator='{op}' value='{cond.value}'>")
                if cond.result:
                    self._generate_calculation_part(wl, cond.result, 4)
                wl("\t\t\t</when>")
            if q.calculationCaseElse:
                wl("\t\t\t<else>")
                self._generate_calculation_part(wl, q.calculationCaseElse, 4)
                wl("\t\t\t</else>")
            wl("\t\t</calculation>")
        elif q.calculationType == CalculationType.CONSTANT:
            wl(f"\t\t<calculation type='constant' value='{q.calculationConstantValue}' />")
        elif q.calculationType == CalculationType.LOOKUP:
            wl(f"\t\t<calculation type='lookup' field='{q.calculationLookupField}' />")
        elif q.calculationType == CalculationType.MATH:
            wl(f"\t\t<calculation type='math' operator='{q.calculationMathOperator}'>")
            for part in q.calculationMathParts:
                self._generate_calculation_part(wl, part, 3)
            wl("\t\t</calculation>")
        elif q.calculationType == CalculationType.CONCAT:
            separator_attr = f" separator='{q.calculationConcatSeparator}'" if q.calculationConcatSeparator else ""
            wl(f"\t\t<calculation type='concat'{separator_attr}>")
            for part in q.calculationConcatParts:
                self._generate_calculation_part(wl, part, 3)
            wl("\t\t</calculation>")
        elif q.calculationType == CalculationType.AGE_FROM_DATE:
            wl(
                f"\t\t<calculation type='age_from_date' field='{q.calculationLookupField}' value='{q.calculationConstantValue}'/>"
            )
        elif q.calculationType == CalculationType.AGE_AT_DATE:
            separator_attr = f" separator='{q.calculationConcatSeparator}'" if q.calculationConcatSeparator else ""
            wl(
                f"\t\t<calculation type='age_at_date' field='{q.calculationLookupField}' value='{q.calculationConstantValue}'{separator_attr}/>"
            )
        elif q.calculationType == CalculationType.DATE_OFFSET:
            wl(
                f"\t\t<calculation type='date_offset' field='{q.calculationLookupField}' value='{q.calculationConstantValue}' />"
            )
        elif q.calculationType == CalculationType.DATE_DIFF:
            wl(
                f"\t\t<calculation type='date_diff' field='{q.calculationLookupField}' value='{q.calculationConstantValue}' unit='{q.calculationUnit}' />"
            )

    def _generate_calculation_part(self, wl, part: CalculationPart, indent_level: int) -> None:
        indent = "\t" * indent_level
        if part.type == CalculationType.CONSTANT:
            wl(f"{indent}<result type='constant' value='{part.constantValue}' />")
        elif part.type == CalculationType.LOOKUP:
            wl(f"{indent}<part type='lookup' field='{part.lookupField}' />")
        elif part.type == CalculationType.QUERY:
            wl(f"{indent}<part type='query'>")
            wl(f"{indent}\t<sql>{part.querySql}</sql>")
            for param in part.queryParameters:
                wl(f"{indent}\t<parameter name='{param.name}' field='{param.fieldName}' />")
            wl(f"{indent}</part>")
        elif part.type == CalculationType.MATH:
            wl(f"{indent}<part type='math' operator='{part.mathOperator}'>")
            for nested in part.parts:
                self._generate_calculation_part(wl, nested, indent_level + 1)
            wl(f"{indent}</part>")
        elif part.type == CalculationType.CONCAT:
            separator_attr = f" separator='{part.concatSeparator}'" if part.concatSeparator else ""
            wl(f"{indent}<part type='concat'{separator_attr}>")
            for nested in part.parts:
                self._generate_calculation_part(wl, nested, indent_level + 1)
            wl(f"{indent}</part>")

    @staticmethod
    def _convert_operator_to_xml(op: str) -> str:
        """Escape a `when` condition's operator for the XML attribute.

        `<>` becomes `!=` rather than `&lt;&gt;`. Every other evaluator in the
        app accepts both spellings, but the one that runs case calculations
        knows only `!=`, so a `<>` written here would never match and the
        calculation would quietly return the wrong value. Normalizing keeps
        `<>` meaning what it means everywhere else in a data dictionary.
        """
        op = op.strip()
        converted = {
            "=": "=",
            "!=": "!=",
            "<>": "!=",
            ">": "&gt;",
            "<": "&lt;",
            ">=": "&gt;=",
            "<=": "&lt;=",
        }.get(op)
        if converted is None:
            raise ValueError(f"Unsupported operator in a when condition: {op!r}")
        return converted
