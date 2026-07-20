# GiSTXConfig_Python - Survey Configuration Generator

Cross-platform Python implementation of GiSTConfigX.
It reads an Excel data dictionary and generates:
- XML file(s) for questionnaire worksheets
- `survey_manifest.gistx`
- `<surveyId>.zip` package (XML + manifest + optional CSV files)
- `gistlogfile.txt` validation/processing log

## Table of Contents

- [Overview](#overview)
- [How It Works](#how-it-works)
- [Installation](#installation)
- [Configuration](#configuration)
- [Creating the Excel Data Dictionary](#creating-the-excel-data-dictionary)
  - [Worksheet Naming Convention](#worksheet-naming-convention)
  - [Required Column Structure](#required-column-structure)
  - [Column Specifications](#column-specifications)
    - [FieldName](#fieldname)
    - [QuestionType](#questiontype)
    - [FieldType](#fieldtype)
    - [QuestionText](#questiontext)
    - [MaxCharacters](#maxcharacters)
    - [Responses](#responses)
    - [Input Masking (Text Fields)](#input-masking-text-fields)
    - [Automatic Calculations](#automatic-calculations)
    - [LowerRange](#lowerrange)
    - [UpperRange](#upperrange)
    - [LogicCheck](#logiccheck)
    - [DontKnow](#dontknow)
    - [Refuse](#refuse)
    - [NA](#na)
    - [Skip](#skip)
    - [Comments](#comments)
- [The CRFS Worksheet](#the-crfs-worksheet)
- [Output Files](#output-files)
- [Error Handling and Validation](#error-handling-and-validation)

---

## Overview

The processor reads worksheets ending in `_dd` or `_xml`, validates each question row, and generates XML definitions plus a survey manifest.

### Inputs

- Excel workbook with questionnaire worksheets (`*_dd`/`*_xml`)
- `crfs` worksheet (form relationships and metadata)
- `config.json`

### Outputs

- One XML file per questionnaire worksheet
- `survey_manifest.gistx`
- `<surveyId>.zip`
- `gistlogfile.txt`

---

## How It Works

1. Load `config.json`
2. Open workbook
3. Validate questionnaire worksheets
4. If no errors, generate XML files
5. If no errors, build `survey_manifest.gistx` from `crfs`
6. If no errors, create `<surveyId>.zip`
7. Always write `gistlogfile.txt`

If validation errors are found, XML/manifest/zip generation is skipped.

---

## Installation

### Requirements

- Python 3.10+
- `pip`

### Install

```bash
pip install -r requirements.txt
```

### Run

```bash
python3 main.py --config config.json
```

Exit code:
- `0` success
- `1` error(s)

---

## Configuration

Copy `config.sample.json` to `config.json` in the project root and edit it
(`config.json` is gitignored, so your local paths stay out of the repo):

```json
{
  "excelFile": "/absolute/path/to/data_dictionary.xlsx",
  "csvFiles": "/absolute/path/to/csv/folder",
  "outputPath": "/absolute/path/to/output",
  "surveyName": "PRISM CSS 2026-01-05",
  "surveyId": "prism_css_2026_01_05",
  "databaseName": "prism_css_2026_01_05.sqlite"
}
```

### Configuration Parameters

| Parameter | Required | Description |
|---|---|---|
| `excelFile` | Yes | Full path to Excel data dictionary |
| `csvFiles` | No | Folder containing CSV files for dynamic responses; all `*.csv` added to zip |
| `outputPath` | Yes | Directory for zip and log output |
| `surveyName` | Yes | Human-readable survey name |
| `surveyId` | Yes | Survey identifier, used for the zip filename (`<surveyId>.zip`) |
| `databaseName` | Yes | SQLite database filename written verbatim into `survey_manifest.gistx` |

Notes:
- `databaseName` in manifest is auto-generated as `<surveyId>.sqlite`
- Generated XML + manifest are temporary files and are deleted after a successful zip
- `gistlogfile.txt` remains in `outputPath`

---

## Creating the Excel Data Dictionary

### Worksheet Naming Convention

- Questionnaire worksheets: names ending in `_dd` or `_xml`
- CRFS worksheet: exactly `crfs`
- Other worksheets: ignored

### Required Column Structure

Each questionnaire worksheet must have these 14 headers in row 1, exact order and spelling:

1. `FieldName`
2. `QuestionType`
3. `FieldType`
4. `QuestionText`
5. `MaxCharacters`
6. `Responses`
7. `LowerRange`
8. `UpperRange`
9. `LogicCheck`
10. `DontKnow`
11. `Refuse`
12. `NA`
13. `Skip`
14. `Comments`

Important:
- Merged rows are ignored as non-question rows
- Non-merged rows after header are parsed as questions

### Column Specifications

#### FieldName

Rules:
- Lowercase only
- Must not start with number
- Must not start with underscore
- Allowed chars: letters, numbers, underscore
- No spaces
- Unique in worksheet (except information rows are excluded from duplicate check)

Examples:
- Valid: `age`, `participant_id`, `hh_member_count`
- Invalid: `Age`, `_id`, `2ndvisit`, `first name`

#### QuestionType

Valid values:
- `radio`
- `combobox`
- `checkbox`
- `text`
- `date`
- `information`
- `automatic`
- `button`

Compatibility rules:
- `radio` => `FieldType` must be `integer`
- `checkbox` => `FieldType` must be `text`
- `date` => `FieldType` must be `date` or `datetime`

#### FieldType

Valid values:
- `text`
- `datetime`
- `date`
- `phone_num`
- `integer`
- `text_integer`
- `text_decimal`
- `text_id`
- `n/a`
- `hourmin`

#### QuestionText

Rules:
- Required unless `QuestionType=automatic`
- Any text allowed
- Placeholder references like `[[fieldname]]` are preserved into XML as text

#### MaxCharacters

Rules:
- Optional generally, but required for:
- `FieldType` in `text`, `text_integer`, `phone_num`
- except `QuestionType` in `automatic`, `checkbox`, `combobox`
- Numeric only (or `=number` format)
- Range: `1..2000`

Examples:
- `80`
- `10`
- `=3`

#### Responses

Used by `radio`, `checkbox`, `combobox`, and `automatic` (calc syntax).

##### Static Responses

Format: one response per line, `value:text`

Example:
```text
1:Yes
2:No
3:Don't Know
```

Applies to: `radio`, `checkbox`, `combobox`

Validation rules (enforced for `radio` and `checkbox`):
- Must contain exactly one `:`
- No leading space before value
- No space immediately after `:` (`1: Yes` is invalid)
- Values must be unique

`combobox` also uses this format but does not apply strict format validation.

##### Dynamic Responses

Start with `source:` and provide key-value lines.

CSV example:
```text
source:csv
file:mrcvillage.csv
filter:region = [[region]]
filter:mrccode = [[mrccode]]
display:villagename
value:vcode
distinct:true
empty_message:No villages found
dont_know:-7, Don't know
not_in_list:-99, Not in list
```

Database example:
```text
source:database
table:hh_members
filter:hhid = [[hhid]]
filter:sex = 1
display:participantsname
value:uniqueid
empty_message:No records found
```

Supported dynamic keys:
- `source` (`csv` or `database`)
- `file` (CSV source)
- `table` (DB source)
- `filter`
- `display`
- `value`
- `distinct` (`true`/`false`)
- `empty_message`
- `dont_know` (`value,label`)
- `not_in_list` (`value,label`)

Supported filter operators:
- `=`
- `!=`
- `<>`
- `>`
- `<`
- `>=`
- `<=`

Filter values support dynamic field references using `[[fieldname]]` notation. At runtime, the GiSTX application substitutes the current value of the named field, allowing responses to be filtered based on earlier answers.

Example: `filter: region = [[region_field]]` filters where the `region` column equals the current value of the `region_field` question.

##### Response XML Behavior

- Static => `<responses><response .../></responses>`
- CSV => `<responses source='csv' file='...'>...`
- Database => `<responses source='database' table='...'>...`

#### Input Masking (Text Fields)

Mask syntax in `Responses` column:

```text
mask:R21-[0-9][0-9][0-9]-[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]
```

Rules:
- Only valid for `QuestionType=text`
- Stored as `<mask value="..." />`

#### Automatic Calculations

Calculation syntax in `Responses` column, only for `QuestionType=automatic`.

Built-in automatic field names that are exempt from calc parsing:
- `starttime`, `stoptime`, `uniqueid`, `swver`, `survey_id`, `lastmod`

Supported calc types:

1. `constant`
```text
calc:constant
value:1
```

2. `lookup`
```text
calc:lookup
field:participant_name
```

3. `query`
```text
calc:query
sql:SELECT count(*) FROM members WHERE hhid = @hhid
param:@hhid = hhid
```

4. `math`
```text
calc:math
operator:+
part:lookup age1
part:lookup age2
```

5. `concat`
```text
calc:concat
separator:, 
part:lookup first_name
part:lookup last_name
```

6. `case`
```text
calc:case
when:age < 18 => Minor
when:age >= 18 => Adult
else:Unknown
```

7. `age_from_date`
```text
calc:age_from_date
field:dob
value:today
```

8. `age_at_date`
```text
calc:age_at_date
field:dob
value:visit_date
separator:-
```

9. `date_offset`
```text
calc:date_offset
field:vx_dose1_date
value:+28d
```

10. `date_diff`
```text
calc:date_diff
field:admission_date
value:today
unit:d
```

Validation enforced per type:
- Required keys (`sql`, `field`, `value`, etc.)
- Math operator must be one of `+ - * /`
- `math` needs at least 2 parts
- `concat` needs at least 1 part
- `date_offset` value must match `[+/-][number][dwmy]`
- `date_diff` unit must be one of `d,w,m,y`

#### LowerRange

For non-date questions:
- Numeric only (integer or decimal)

For date questions:
- Required
- Allowed values:
- `0`, `+0d`, `-0d`
- relative format `[+/-][number][d|w|m|y]`
- fixed date `yyyy-mm-dd`

Examples:
- `0`
- `-1y`
- `+6m`
- `2023-01-01`

#### UpperRange

Same validation as `LowerRange`.

#### LogicCheck

One or more lines in format:

```text
expression; 'error message'
```

Examples:
```text
age >= 18; 'Participant must be 18 or older'
end_date > start_date; 'End date must be after start date'
```

Multiple checks: one line per check.

Unique check format:
```text
unique; 'This ID has already been used'
```

Checkbox `contains` operators:

Use these operators in a LogicCheck expression when testing a checkbox or other multi-select field. Unlike Skip rules, the operator words are not quoted:

```text
symptoms contains '99'; 'Example message'
symptoms does not contain '99'; 'Example message'
```

For example, define a `symptoms` checkbox field with these static responses:

```text
1:Fever
2:Cough
3:Symptom 3
4:Symptom 4
5:Symptom 5
6:Symptom 6
7:Symptom 7
8:Symptom 8
9:Symptom 9
10:Symptom 10
11:Symptom 11
12:Symptom 12
99:Child does not have any symptoms
```

Use this complete value in the `LogicCheck` cell to prevent selecting `99` together with any other symptom:

```text
symptoms contains '99' and (symptoms contains '1' or symptoms contains '2' or symptoms contains '3' or symptoms contains '4' or symptoms contains '5' or symptoms contains '6' or symptoms contains '7' or symptoms contains '8' or symptoms contains '9' or symptoms contains '10' or symptoms contains '11' or symptoms contains '12'); 'Cannot select "Child does not have any symptoms" together with other symptoms'
```

Validation rules:
- Must include semicolon
- Message must be in single quotes
- Expression must include comparison, logical, `contains`, or `does not contain` operators
- Referenced fields must exist
- Referenced fields must appear before current row

#### DontKnow

Valid values:
- `True`
- `False`
- blank (treated as not set)

If `True`, XML includes `<dont_know>-7</dont_know>`.

#### Refuse

Valid values:
- `True`
- `False`
- blank

If `True`, XML includes `<refuse>-8</refuse>`.

#### NA

Valid values:
- `True`
- `False`
- blank

If `True`, XML includes `<na>-6</na>`.

#### Skip

One or more lines (one skip rule per line).

Supported prefixes:
- `preskip:` — evaluated before the question is rendered
- `postskip:` — evaluated after the user enters a response

Required format:
```
[preskip|postskip]: if [checkfield] [condition] [value], skip to [targetfield]
```

- `checkfield` — field whose value is tested; must be defined before the current question
- `condition` — comparison operator; must be one of: `=`, `>`, `>=`, `<`, `<=`, `<>`
- `value` — value to compare against (no quotes needed for numbers)
- `targetfield` — last word on the line; must be a field defined after the current question

Examples:
```text
preskip: if age < 18, skip to comments
postskip: if pregnant = 2, skip to next_section
```

Checkbox `contains` operators:
```text
postskip: if symptoms 'contains' 1, skip to fever
postskip: if symptoms 'does not contain' 9, skip to cough
```

- Use `'contains'` (with single quotes) as the condition for checkbox fields
- `does not contain` is written without quotes and expands to 7 tokens in the condition section

Validation rules:
- Must start with `preskip:` or `postskip:`
- Must have exactly one comma separating the condition section from the skip-to section
- `checkfield` must exist and appear before current row
- `targetfield` must exist and appear after current row
- Cannot skip to the current question itself

#### Comments

Ignored by processor. Use for notes/documentation.

---

## The CRFS Worksheet

Worksheet name must be exactly `crfs`. Columns are read positionally (left to right); column headers in the sheet are ignored.

| # | Column | Type | Description |
|---|--------|------|-------------|
| 1 | `display_order` | integer | Order in which this form appears in the navigation menu |
| 2 | `tablename` | string | Database table name where this form's data is stored |
| 3 | `displayname` | string | Human-readable form name shown to the user |
| 4 | `primarykey` | string | Name of the primary key field for this form's table |
| 5 | `idconfig` | JSON | Configuration for auto-generated ID values (see below) |
| 6 | `isbase` | integer | `1` = root/base form, `0` = related form |
| 7 | `linkingfield` | string | Field in this form that links it to the parent record |
| 8 | `parenttable` | string | Table name of the parent form (for hierarchical forms) |
| 9 | `incrementfield` | string | Field used for auto-incrementing part of a composite ID |
| 10 | `requireslink` | integer | `1` = a parent record must exist before this form can be entered |
| 11 | `repeat_count_field` | string | Field whose value controls how many repeat instances are created |
| 12 | `auto_start_repeat` | integer | `1` = automatically create repeat instances without user action |
| 13 | `repeat_enforce_count` | integer | `1` = enforce that exactly `repeat_count_field` instances are completed |
| 14 | `display_fields` | string | Comma-separated field names to display in the form summary/list view |
| 15 | `entry_condition` | string | Expression that must evaluate to true before this form can be entered |

All columns are optional (null values are omitted from the manifest).

### `idconfig` JSON structure

Used when the primary key is a composite ID built from field values plus an auto-increment counter:

```json
{
  "prefix": "GL",
  "fields": [
    {"name": "community", "length": 2},
    {"name": "village", "length": 2}
  ],
  "incrementLength": 3
}
```

- `prefix` — fixed string prepended to the ID
- `fields` — array of field names and their character lengths contributing to the ID
- `incrementLength` — number of digits for the auto-increment portion

`crfs` rows are written into `survey_manifest.gistx` under `crfs`.

---

## Output Files

### XML files

Generated for each worksheet ending in `_dd` or `_xml`.

Filename mapping:
- `hh_info_dd` -> `hh_info.xml`
- `followup_xml` -> `followup.xml`

Each XML ends with synthetic information question:
- `fieldname='end_of_questions'`

### `survey_manifest.gistx`

Contains:
- `surveyName`
- `surveyId`
- `databaseName` (`<surveyId>.sqlite`)
- `xmlFiles` list
- `crfs` entries

### `<surveyId>.zip`

Contains:
- generated XML files
- `survey_manifest.gistx`
- all `*.csv` from `csvFiles` folder (if configured)

### `gistlogfile.txt`

Always written. Includes:
- per-worksheet checks
- warnings/errors
- packaging messages
- end-of-log marker

---

## Error Handling and Validation

### Common Validation Errors

- Invalid/missing headers
- FieldName format violations
- Invalid QuestionType/FieldType values
- QuestionType/FieldType mismatch (`radio`/`checkbox`/`date`)
- Malformed static responses
- Duplicate static response keys
- Invalid dynamic response keys/values
- Invalid calc syntax/required fields missing
- Invalid numeric/date range values
- LogicCheck syntax/reference errors
- Skip syntax/reference/ordering errors
- Duplicate non-information field names
- Invalid `DontKnow`/`Refuse`/`NA` values (must be `True`/`False` when set)
- XML syntax validation failure after generation

### Validation Process Order (per worksheet)

1. Header validation
2. Row parsing + field-level validations
3. Logic field-reference validation
4. Skip field-reference validation
5. Required MaxCharacters validation
6. Duplicate FieldName validation

If any errors are found, generation is halted.

### Troubleshooting

- `Excel file not found`: fix `excelFile` path in config
- Missing outputs: check `gistlogfile.txt` for first validation error
- No CSV in zip: verify `csvFiles` exists and contains `*.csv`
- Manifest not generated: ensure `crfs` sheet exists and workbook has no prior validation errors

---

## File Layout

All Python source files are in the project root:
- `main.py` — CLI entry point
- `processor.py` — orchestration
- `excel_reader.py` — worksheet parsing and validation
- `xml_generator.py` — XML generation
- `crf_reader.py` — CRFS worksheet parsing
- `json_generator.py` — manifest writing
- `models.py` — data models and enums
- `config.json` — runtime configuration
