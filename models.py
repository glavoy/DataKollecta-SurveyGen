from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class ResponseSourceType(str, Enum):
    STATIC = "Static"
    CSV = "Csv"
    DATABASE = "Database"


class CalculationType(str, Enum):
    NONE = "None"
    QUERY = "Query"
    CASE = "Case"
    CONSTANT = "Constant"
    LOOKUP = "Lookup"
    MATH = "Math"
    CONCAT = "Concat"
    AGE_FROM_DATE = "AgeFromDate"
    AGE_AT_DATE = "AgeAtDate"
    DATE_OFFSET = "DateOffset"
    DATE_DIFF = "DateDiff"


@dataclass
class Filter:
    column: str
    value: str
    operator: str = "="


@dataclass
class CalculationParameter:
    name: str
    fieldName: str


@dataclass
class CalculationPart:
    type: CalculationType = CalculationType.NONE
    constantValue: str = ""
    lookupField: str = ""
    querySql: str = ""
    queryParameters: list[CalculationParameter] = field(default_factory=list)
    mathOperator: str = ""
    parts: list["CalculationPart"] = field(default_factory=list)
    concatSeparator: str = ""


@dataclass
class CaseCondition:
    field: str
    operator: str
    value: str
    result: CalculationPart | None = None


@dataclass
class Question:
    fieldName: str = ""
    questionType: str = ""
    fieldType: str = ""
    questionText: str = ""
    maxCharacters: str = ""
    responses: str = ""

    responseSourceType: ResponseSourceType = ResponseSourceType.STATIC
    responseSourceFile: str = ""
    responseSourceTable: str = ""
    responseFilters: list[Filter] = field(default_factory=list)
    responseDisplayColumn: str = ""
    responseValueColumn: str = ""
    responseDistinct: bool | None = None
    responseEmptyMessage: str = ""
    responseDontKnowValue: str = ""
    responseDontKnowLabel: str = ""
    responseNotInListValue: str = ""
    responseNotInListLabel: str = ""

    calculationType: CalculationType = CalculationType.NONE
    calculationQuerySql: str = ""
    calculationQueryParameters: list[CalculationParameter] = field(default_factory=list)
    calculationCaseConditions: list[CaseCondition] = field(default_factory=list)
    calculationCaseElse: CalculationPart | None = None
    calculationConstantValue: str = ""
    calculationLookupField: str = ""
    calculationMathOperator: str = ""
    calculationMathParts: list[CalculationPart] = field(default_factory=list)
    calculationConcatSeparator: str = ""
    calculationConcatParts: list[CalculationPart] = field(default_factory=list)
    calculationUnit: str = ""

    lowerRange: str = ""
    upperRange: str = ""
    logicChecks: list[str] = field(default_factory=list)
    uniqueCheckMessage: str = ""
    dontKnow: str = ""
    refuse: str = ""
    na: str = ""
    skip: str = ""
    mask: str = ""


@dataclass
class AppConfig:
    excelFile: str
    csvFiles: str = ""
    outputPath: str = ""
    surveyName: str = ""
    surveyId: str = ""
    databaseName: str = ""


@dataclass
class IdConfigField:
    name: str
    length: int


@dataclass
class IdConfig:
    prefix: str | None = None
    fields: list[IdConfigField] | None = None
    incrementLength: int | None = None


@dataclass
class Crf:
    display_order: int | None = None
    tablename: str | None = None
    displayname: str | None = None
    isbase: int | None = None
    primarykey: str | None = None
    linkingfield: str | None = None
    idconfig: IdConfig | None = None
    parenttable: str | None = None
    incrementfield: str | None = None
    requireslink: int | None = None
    repeat_count_source: str | None = None
    repeat_count_field: str | None = None
    auto_start_repeat: int | None = None
    repeat_enforce_count: int | None = None
    display_fields: str | None = None
    entry_condition: str | None = None


@dataclass
class SurveyManifest:
    surveyName: str
    surveyId: str
    databaseName: str
    xmlFiles: list[str]
    crfs: list[Crf]
