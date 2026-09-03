from __future__ import annotations

import json
from pathlib import Path

from models import Crf, IdConfig, SurveyManifest

# The Windows app serializes the manifest with Newtonsoft Json.NET
# (Formatting.Indented) on Windows: 2-space indent, CRLF newlines, no trailing
# newline, null properties omitted, and the idconfig "fields" array written on
# a single compact line (via a custom converter). This module reproduces that
# output byte for byte.
NEWLINE = "\r\n"

# Crf property order must match the C# class declaration order.
CRF_FIELD_ORDER = [
    "display_order",
    "tablename",
    "displayname",
    "isbase",
    "primarykey",
    "linkingfield",
    "idconfig",
    "parenttable",
    "incrementfield",
    "requireslink",
    # NOTE: the C# class this order mirrors also declares
    # `repeat_count_source` here. It is omitted deliberately: nothing on
    # the Python side ever set it, so it was always absent from the
    # output, and only fields that are actually emitted affect ordering.
    "repeat_count_field",
    "auto_start_repeat",
    "repeat_enforce_count",
    "display_fields",
    "entry_condition",
]


def _jstr(value: str) -> str:
    return json.dumps(value, ensure_ascii=False)


def _jval(value) -> str:
    if isinstance(value, str):
        return _jstr(value)
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


def _idconfig_lines(idconfig: IdConfig, indent: str) -> list[str]:
    inner = indent + "  "
    lines = [indent + '"idconfig": {']
    parts = []
    if idconfig.prefix is not None:
        parts.append(f'{inner}"prefix": {_jstr(idconfig.prefix)}')
    if idconfig.fields is not None:
        fields_json = ",".join(
            "{" + f'"name":{_jstr(f.name)},"length":{f.length}' + "}" for f in idconfig.fields
        )
        parts.append(f'{inner}"fields": [{fields_json}]')
    else:
        parts.append(f'{inner}"fields": null')
    if idconfig.incrementLength is not None:
        parts.append(f'{inner}"incrementLength": {idconfig.incrementLength}')
    for i, part in enumerate(parts):
        lines.append(part + ("," if i < len(parts) - 1 else ""))
    lines.append(indent + "}")
    return lines


def _crf_lines(crf: Crf, indent: str) -> list[str]:
    inner = indent + "  "
    entries: list[list[str]] = []
    for name in CRF_FIELD_ORDER:
        value = getattr(crf, name)
        if value is None:
            continue
        if name == "idconfig":
            entries.append(_idconfig_lines(value, inner))
        else:
            entries.append([f'{inner}"{name}": {_jval(value)}'])
    if not entries:
        return [indent + "{}"]
    lines = [indent + "{"]
    for i, entry in enumerate(entries):
        entry = list(entry)
        if i < len(entries) - 1:
            entry[-1] += ","
        lines.extend(entry)
    lines.append(indent + "}")
    return lines


class JsonGenerator:
    @staticmethod
    def write_manifest(path: Path, manifest: SurveyManifest) -> None:
        lines: list[str] = ["{"]
        lines.append(f'  "surveyName": {_jstr(manifest.surveyName)},')
        lines.append(f'  "surveyId": {_jstr(manifest.surveyId)},')
        lines.append(f'  "databaseName": {_jstr(manifest.databaseName)},')

        if manifest.xmlFiles:
            lines.append('  "xmlFiles": [')
            for i, xml_file in enumerate(manifest.xmlFiles):
                comma = "," if i < len(manifest.xmlFiles) - 1 else ""
                lines.append(f"    {_jstr(xml_file)}{comma}")
            lines.append("  ],")
        else:
            lines.append('  "xmlFiles": [],')

        if manifest.crfs:
            lines.append('  "crfs": [')
            for i, crf in enumerate(manifest.crfs):
                crf_lines = _crf_lines(crf, "    ")
                if i < len(manifest.crfs) - 1:
                    crf_lines[-1] += ","
                lines.extend(crf_lines)
            lines.append("  ]")
        else:
            lines.append('  "crfs": []')

        lines.append("}")
        with path.open("w", encoding="utf-8", newline="") as f:
            f.write(NEWLINE.join(lines))
