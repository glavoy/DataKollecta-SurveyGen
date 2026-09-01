# DataKollecta-SurveyGen - Survey Configuration Generator

A tool for generating XML configuration files and survey manifests from Excel-based data dictionaries. This application streamlines the creation of survey forms for data collection instruments by automating the generation of XML configuration files and survey manifest files.

## Table of Contents

- [Overview](#overview)
- [How It Works](#how-it-works)
- [Installation](#installation)
- [Configuration](#configuration)
- [Versioning a survey](#versioning-a-survey)
- [Creating the Excel Data Dictionary](#creating-the-excel-data-dictionary)
  - [Worksheet Naming Convention](#worksheet-naming-convention)
  - [Column Specifications](#column-specifications)
    - [FieldName](#fieldname)
    - [QuestionType](#questiontype)
    - [Reserved Automatic Variables](#reserved-automatic-variables)
    - [Computed Automatic Variables (yyyy, yy, mm, dd, doy)](#computed-automatic-variables-yyyy-yy-mm-dd-doy)
    - [Custom Timestamp Fields](#custom-timestamp-fields)
    - [FieldType](#fieldtype)
    - [QuestionText](#questiontext)
    - [MaxCharacters](#maxcharacters)
    - [Responses](#responses)
    - [LowerRange](#lowerrange)
    - [UpperRange](#upperrange)
    - [LogicCheck](#logiccheck)
    - [DontKnow](#dontknow)
    - [Refuse](#refuse)
    - [Optional](#optional)
    - [Skip](#skip)
    - [Comments](#comments)
  - [The CRFS Worksheet](#the-crfs-worksheet)
- [Examples](#examples)
- [Output Files](#output-files)
- [Error Handling and Validation](#error-handling-and-validation)

---

## Overview

**DataKollecta-SurveyGen** reads Excel workbooks containing structured data dictionaries and automatically generates:
- **XML configuration files** for each questionnaire/form definition
- **survey_manifest.gistx** configuration file containing survey metadata and form relationships

### What the App Does

The application processes an Excel spreadsheet (the "data dictionary") that defines your survey questionnaires. Each worksheet ending in `_dd` represents a separate questionnaire. The app:

1. Validates the data dictionary structure and syntax
2. Generates XML files for each questionnaire with question definitions, validation rules, skip logic, and response options
3. Creates a survey_manifest.gistx configuration file that defines the survey structure, form hierarchy, and relationships
4. Generates comprehensive error logs to help you fix any issues

### Inputs and Outputs

**Inputs:**
- Excel spreadsheet (data dictionary) with questionnaires defined in worksheets ending in `_dd`
- `config.json` configuration file with paths and survey metadata

**Outputs:**
- XML files (one per questionnaire) in the specified output path
- `survey_manifest.gistx` configuration file
- `gistlogfile.txt` log file with validation results

---

## How It Works

1. **Configure Paths**: Set up your `config.json` file with paths to your Excel file, output directories, and survey metadata
2. **Create Data Dictionary**: Build your Excel workbook with questionnaires in worksheets ending in `_dd` and a `crfs` worksheet defining form metadata
3. **Run the Application**: Click the "Generate XML" button
4. **Review Output**: Check the log file for any errors or warnings
5. **Use Generated Files**: The XML files and survey_manifest.gistx file are ready to use in your survey application

The application validates:
- Field name syntax and uniqueness
- Question type and field type compatibility
- Response format and syntax
- Logic check syntax and field references
- Skip logic syntax and field references
- Required fields and values

---

## Installation

This is a Python command-line tool — no IDE or build step required.

1. Clone this repository.
2. Install [Python 3.9 or later](https://www.python.org/downloads/) if you don't already have it.
3. From the repository folder, create a virtual environment and install dependencies:

   **macOS / Linux:**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

   **Windows:**
   ```powershell
   py -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```

4. Run it:

   ```bash
   python main.py --config config.json
   ```

   `--config` defaults to `config.json` in the current folder, so `python main.py` alone works too if you're using the default filename.

A virtual environment isn't portable — if you move or re-clone this repository, recreate `.venv` rather than reuse one from another location.

### System Requirements

- Python 3.9 or later — macOS or Windows

### Dependencies

- **openpyxl** (`>=3.1.0`) — reads the Excel data dictionary. Installed automatically via `requirements.txt`.

---

## Configuration

Create or edit the `config.json` file in the application directory with the following settings:

```json
{
  "excelFile": "C:\\PRISM\\Excel\\prismcss.xlsx",
  "outputPath": "C:\\temp\\",
  "surveyName": "PRISM CSS 2025-12-01",
  "surveyId": "prism_css_2025_12_01",
  "databaseName": "prism_css.sqlite"
}
```

Note that `databaseName` carries **no date**, while `surveyName` and `surveyId` do. That
difference is deliberate and load-bearing — see [Versioning a survey](#versioning-a-survey).

On macOS/Linux, use forward-slash paths instead, e.g. `"excelFile": "/Users/you/PRISM/Excel/prismcss.xlsx"` and `"outputPath": "/tmp/"`.

### Configuration Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| `excelFile` | Full path to your Excel data dictionary | `"C:\\PRISM\\Excel\\prismcss.xlsx"` |
| `outputPath` | Directory where the final zip file and log file will be saved | `"C:\\temp\\"` |
| `surveyName` | Human-readable survey name displayed in the app. **Change every revision** | `"PRISM CSS 2025-12-01"` |
| `surveyId` | Unique survey identifier (lowercase, no spaces). **Change every revision** | `"prism_css_2025_12_01"` |
| `databaseName` | SQLite database filename written to `survey_manifest.gistx`. **Never change it once data has been collected** | `"prism_css.sqlite"` |

**Notes:**
- The application creates a zip file containing all XML files, the survey_manifest.gistx file, and any CSV files used for dynamic responses
- The log file (`gistlogfile.txt`) is written to the `outputPath` directory
- The database name is read from `databaseName` in `config.json`

---

## Versioning a survey

A survey is revised many times over a study's life — a question is reworded, a response option
is added, a `repeat_enforce_count` is changed. Each revision produces a new zip that is
installed over the previous one on devices that may already hold collected data. Two of the
three identifiers in `config.json` exist to describe *this revision*; the third exists to
describe *the study*, and confusing them destroys data.

| Field | Role | On a new revision |
|---|---|---|
| `surveyName` | What the interviewer sees in the app's survey list | **Change it** — carry the revision date |
| `surveyId` | Identifies the revision; the app stores each one in its own folder | **Change it** — carry the revision date |
| `databaseName` | Names the SQLite file holding **all collected data for the study** | **Never change it** |

### `databaseName` must never change

This is the one rule that cannot be broken. The app creates its database from `databaseName`,
so a new name means a **new, empty database** — and generated IDs are derived by taking the
highest existing value in that database and adding one. An empty database restarts the
counter:

```
prism_css.sqlite  already holds  hhid 3010001, 3010002, 3010003
   |
   |  new zip deployed with databaseName "prism_css_2026_08_20.sqlite"
   v
prism_css_2026_08_20.sqlite is created empty
   next household enrolled gets hhid 3010001   <-- DUPLICATE
```

The duplicates are not detected at entry — each device is internally consistent — and only
surface when data is pooled, by which time the interviews have happened. The old database is
also left behind on the device, so previously collected records stop syncing.

So `databaseName` names the **study**, not the revision: `prism_css.sqlite` stays put for the
life of the study, no matter how many times the dictionary is regenerated. Only add a year (
`prism_css_2026.sqlite`) if a genuinely new round is starting that is *meant* to begin its ID
counters afresh.

See `DataKollecta/docs/DATABASE_VERSIONING_DECISIONS.md` in the app repository for the full
analysis.

### Revising a survey

1. Edit the Excel data dictionary.
2. In `config.json`, set `surveyName` and `surveyId` to the new revision date. Leave
   `databaseName` **exactly as it was**.
3. `python main.py`
4. Check `gistlogfile.txt` — a survey that failed validation still writes a zip for the
   worksheets that passed, so a clean log is the only proof the package is complete.
5. Deploy the zip and record what went out: the revision date, what changed, and which devices
   received it. The app records `survey_id` on every collected record, so this is what lets you
   tell later which revision a given record was collected under.

Keep `config.json` per study rather than editing one shared file — `config_prism_css.json` and
`config_r21.json` in this repository are that pattern. It keeps each study's `databaseName`
where it cannot be edited by accident while working on another study.

### Choosing a revision identifier

Dates, not semantic versions. A survey revision is not software: there is no meaningful
distinction between a "minor" and a "patch" change to a questionnaire, and the question anyone
actually asks about a collected record is *when* was this version in the field.

```
prism_css_2026_08_20     ->  prism_css_2026_09_14     ->  prism_css_2026_11_02
```

Use `YYYY_MM_DD` in `surveyId` (lowercase, underscores — it becomes a folder name) and
`YYYY-MM-DD` in `surveyName`, matching the existing packages. If two revisions go out on one
day, add a suffix: `prism_css_2026_08_20b`.

---

## Creating the Excel Data Dictionary

### Worksheet Naming Convention

- **Questionnaires**: Any worksheet ending in `_dd` will be processed as a questionnaire
  - Example: `hh_info_dd`, `enrollment_dd`, `followup_dd`
  - The `_dd` suffix is removed when creating the XML filename

- **CRFS Worksheet**: A worksheet named exactly `crfs` must be included to define form metadata (see [The CRFS Worksheet](#the-crfs-worksheet))

- **Other Worksheets**: Any worksheet not ending in `_dd` and not named `crfs` will be ignored (can be used for reference data, documentation, etc.)

### Required Column Structure

Each questionnaire worksheet (`_dd` worksheets) must have exactly 14 columns with these specific headers (in this exact order):

| Column | Header Name |
|--------|-------------|
| 1 | FieldName |
| 2 | QuestionType |
| 3 | FieldType |
| 4 | QuestionText |
| 5 | MaxCharacters |
| 6 | Responses |
| 7 | LowerRange |
| 8 | UpperRange |
| 9 | LogicCheck |
| 10 | DontKnow |
| 11 | Refuse |
| 12 | Optional |
| 13 | Skip |
| 14 | Comments |

**Important Notes:**
- The first row of each worksheet must contain these exact column headers
- Rows that are merged will be ignored (useful for section headers or notes)
- Each non-merged row after the header represents one question/field
- Column 12 was named `NA` before the Optional column existed -- a worksheet with `NA` there
  instead of `Optional` is still accepted, but that column's contents are ignored entirely. See
  the [Optional](#optional) section below.

---

## Column Specifications

### FieldName

The variable/field name that will be used in the database and XML.

**Requirements:**
- Must be lowercase
- Must start with a letter (not a number or underscore)
- Can only contain letters, numbers, and underscores
- No spaces allowed
- Must be unique within the worksheet

**Examples:**
- ✅ `age`
- ✅ `participant_name`
- ✅ `hh_member_count`
- ❌ `Age` (not lowercase)
- ❌ `_fieldname` (starts with underscore)
- ❌ `2ndvisit` (starts with number)
- ❌ `first name` (contains space)

---

### QuestionType

The type of question/input control.

**Valid Values:**

| QuestionType | Description | Use Case |
|--------------|-------------|----------|
| `radio` | Single selection (radio buttons) | Select one option from a list |
| `checkbox` | Multiple selection | Select multiple options from a list |
| `combobox` | Dropdown selection | Select one option from a dropdown |
| `text` | Text entry field | Free text input |
| `date` | Date picker | Date selection |
| `information` | Display-only text | Show information without collecting data |
| `automatic` | Auto-calculated/system field | Field populated by code (not shown to user) |
| `button` | Action button | Trigger an action |

**Requirements:**
- `radio` must have `fieldtype` = `integer`
- `checkbox` must have `fieldtype` = `text`
- `date` must have `fieldtype` = `date` or `datetime`
- `radio`, `checkbox`, and `combobox` must have responses defined

**Spellings of `automatic`:** `calc`, `calculation` and `calculated` all mean the same
thing as `automatic` and can be used interchangeably. They are normalised to `automatic`
in the generated XML, so a dictionary written either way produces a file that every
version of the app can read.

The spelling never decides what the field does. That is determined by:

| What the app finds | What it does |
|---|---|
| the FieldName is a reserved automatic variable (see below) | records the built-in value |
| the question has a `calc:` block | runs the calculation |
| neither, and the form's `idconfig` is set | generates the ID (e.g. `subjid`, `hhid`) |

Pick whichever spelling reads best — `calculation` for formula fields and `automatic`
for system and ID fields is a reasonable convention, but nothing enforces it.

---

### Reserved Automatic Variables

These FieldNames have built-in meaning. The app records their values itself, so they
never need a `calc:` block, responses or validation:

| FieldName | What it records |
|-----------|-----------------|
| `starttime` | the date and time the interview was started |
| `startdate` | the date the interview was started |
| `stoptime` | the date and time the interview was saved |
| `lastmod` | the date and time the record was last modified |
| `uniqueid` | a unique identifier generated for the record |
| `swver` | the version of the app that collected the record |
| `survey_id` | the identifier of the survey the record belongs to |

**The generator writes these into every questionnaire itself**, in the positions where
they record the right moment: `starttime` and `startdate` before the first question, and
the rest after the last one but ahead of the end-of-survey screen. You never have to
declare them, and forgetting one can no longer produce a questionnaire that is missing it.

**Declaring them is never an error.** The generator logs a warning, drops the declared
row and writes its own, so it does not matter where in the spreadsheet you put them and a
questionnaire can never end up with two. Keep the rows if you want the variables visible
to whoever analyses the data — that is a perfectly good reason to leave them in.

The one thing to avoid is giving a reserved variable a `calc:` block: the calculation is
**dropped**, not applied, because the app supplies the value itself. The generator warns
when this happens. If you need a value of your own, use a different FieldName.

**Not the same thing as `yyyy`/`yy`/`mm`/`dd`/`doy`**, below — those are automatic too, and
the app computes them the same way, but you still have to declare them yourself; nothing
writes them in for you.

---

### Computed Automatic Variables (yyyy, yy, mm, dd, doy)

Like the reserved variables above, these `FieldName`s have built-in meaning and take no
`calc:` block or validation — but unlike them, **the generator does not write these in for
you.** You have to declare a row for each one, exactly where you want it in the
questionnaire. Each is computed once, from *today's* date, and never recomputed — the
"today" is fixed the moment the row is first reached, not re-read every time the record is
opened.

| FieldName | What it computes | Format |
|-----------|-------------------|--------|
| `yyyy` | the current four-digit year | zero-padded to 4 — `"2026"` |
| `yy` | the current two-digit year | `year % 100`, zero-padded to 2 — `2026` → `"26"` |
| `mm` | the current month | zero-padded to 2 — `"01"`–`"12"` |
| `dd` | the current day of the month | zero-padded to 2 — `"01"`–`"31"` |
| `doy` | the ordinal day within the current calendar year | zero-padded to 3 — `"001"` (Jan 1) through `"365"`/`"366"` (Dec 31) |

**Worked examples:** `2001` → `yy` `"01"`; Feb 1 → `doy` `"032"`.

**A given `yy` repeats every century** — `2000` and `2100` are both `"00"`. A real ambiguity,
not a bug: acceptable because no study spans a century, but worth knowing before reading
`yy` back out of old data years later.

**Why they aren't reserved.** The generator's auto-injection only knows two fixed positions —
before the first question, or after the last one (see above) — which fits a value every
record needs, computed at a fixed moment. These fields don't fit that: they're useful
anywhere a survey wants "today," and one common use is sitting immediately before a specific
`idconfig`-generated ID field, a position the generator has no way to infer on its own. So
they're declared like an ordinary field instead — a plain `FieldType='automatic'` row with a
**blank** `Responses` column, no `calc:` block, on **any** worksheet, not only a table that
happens to use one of them in its own `idconfig`.

**Automatic vs. `calc:`.** These five fields work like `startdate`: set once and never
recomputed, even if the record is later edited. If you instead want a component pulled from
some *other* date the survey collects — `dob`, an appointment date, anything but today — use
the `date_part` calculation type in [Automatic Calculations](#automatic-calculations) below.
It shares the same five unit tokens, but reads a named field and recomputes it whenever that
field's value changes, since it's meant to track its source rather than freeze at first use.

**Example — Computed Automatic Variable**, feeding a reinstall-resilient subject ID (full
version in [§7](#7-idconfig-reference-id-generation)):

| FieldName | QuestionType | FieldType | Responses | Purpose |
|-----------|--------------|-----------|-----------|---------|
| `doy` | `automatic` | `text` | *(blank)* | Records today's day-of-year once, the first time the row is reached, and never recomputes — including on a later edit. That's what lets `idconfig.fields` use it as a stable piece of the subject ID. |

**Example — `date_part` calc field**, reporting the year a participant enrolled:

| FieldName | QuestionType | FieldType | Responses | Purpose |
|-----------|--------------|-----------|-----------|---------|
| `enroll_year` | `automatic` | `text` | `calc:date_part`<br>`field:enroll_date`<br>`unit:yyyy` | Extracts the year from the survey's own `enroll_date` field. If `enroll_date` is later corrected, `enroll_year` recomputes to match — it tracks its source rather than freezing, the opposite of the `doy` example above. |

---

### Custom Timestamp Fields

`starttime` and `stoptime` aren't special magic — they're `automatic` questions with a
`datetime` `FieldType`, and the app stamps the current date-and-time into any question that
shape the moment its row is reached. What makes `starttime`/`stoptime` different is only that
their *FieldNames* are reserved, so the generator auto-injects them at two fixed positions
(see [Reserved Automatic Variables](#reserved-automatic-variables) above). Give an ordinary,
non-reserved FieldName the same `automatic` + `datetime` shape, with a `calc:timestamp` block,
and you get a real timestamp at **whatever point in the questionnaire you place the row** — no
reserved position, no offset math from `starttime` needed.

| FieldName | QuestionType | FieldType | Responses | Purpose |
|-----------|--------------|-----------|-----------|---------|
| `section3_time` | `automatic` | `datetime` | `calc:timestamp` | Stamped with the date and time the moment the interview reaches this row — wherever you place it. |

**Position decides what it records** — put the row immediately after whatever section you want
timed, the same way `starttime`'s fixed position (before the first question) is what makes it
mean "when the interview started" rather than anything else. Add as many of these as you like
at different points in one questionnaire, each with its own FieldName.

**Frozen once the record is saved and later reopened for editing** — `calc:timestamp` always
generates with `preserve: true` (see [Timestamp](#12-timestamp) under Automatic Calculations),
so editing an existing record does not restamp it to the edit time. Within a single, not-yet-saved
interview, it *does* still recompute if the interviewer backs up past this row and comes forward
again — like every other `calc:` type, and unlike `starttime`/`stoptime` themselves, which are
reserved system fields rather than `calc:` fields and freeze on first reach unconditionally. If
you instead want a value that tracks another date field and recomputes whenever that field
changes, use the `date_part` calculation type ([Automatic Calculations](#automatic-calculations)
below) — that's the same distinction drawn for `yyyy`/`yy`/`mm`/`dd`/`doy` above.

**`FieldType` must be `datetime`, not `date` or `text`** — enforced by the generator; anything
else is a validation error.

**Required, not optional.** Every `automatic` field needs a `calc:` block (or one of the handful
of documented exemptions — see [Silent-Failure Checks](#silent-failure-checks)); a blank
Responses column on a `datetime` automatic field is a build-blocking error now, not an implicit
custom-timestamp fallback.

**Why not a `date_offset` calculation from `starttime`?** `date_offset` computes a *guess*
forward or backward from another date, in whole days/weeks/months/years — it has no
hours/minutes granularity and isn't a real observation. A custom `automatic`/`datetime` field
records what actually happened, which is what a mid-interview timestamp is for.

---

### FieldType

The data type that determines how the value is stored and validated.

**Valid Values:**

| FieldType | Description | Storage Type |
|-----------|-------------|--------------|
| `text` | Text string | Text |
| `integer` | Whole numbers — the value a `radio` stores | Integer |
| `text_integer` | Typed field accepting only digits | Text (validated) |
| `text_decimal` | Typed field accepting digits and one decimal point | Text (validated) |
| `hourmin` | 24-hour time, `hh:mm` — the `:` is inserted by the app | Text (validated) |
| `date` | Date only | Date |
| `datetime` | Date and time | DateTime |
| `n/a` | Not applicable | None (for information questions) |

**A `text` question must use `text`, `text_integer`, `text_decimal` or
`hourmin`.** `integer` is what a `radio` stores; using it on a typed question
leaves the app unable to tell a whole number from a decimal, and it used to
skip the MaxCharacters requirement as well.

`phone_num` and `text_id` have been removed. Both behaved exactly as `text`,
and no dictionary used them — record a phone number as `text_integer` with a
fixed length (`=10`).

---

### QuestionText

The actual question text shown to users.

**Requirements:**
- Required for all question types except `automatic`
- Can contain any text, including special characters
- Can include placeholder variables using `[[fieldname]]` syntax
  - Example: `"What is [[child_name]]'s date of birth?"`

#### High-Visibility Warning Theme

If the `QuestionText` starts with the word **"Warning"** (case-insensitive), the survey application automatically triggers a high-visibility theme:

- **Visual Alert**: The question text is wrapped in an amber-colored box with an orange border and a warning icon.
- **Smart Titles**: For `information` question types, the header title automatically changes from "Information" to "Warning".
- **Broad Support**: This works across all visible question types (radio, checkbox, text, etc.).

**Example:**
`Warning: Please ensure you have obtained written consent before proceeding.`

**Examples of Question Text:**
```
What is your age?
How many people live in this household?
Select the mother of [[child_name]]
```

---

### MaxCharacters

Maximum character length for text fields.

**Requirements:**
- **Required for every question whose QuestionType is `text`**, whatever its FieldType
- Must be a number between 1 and 2000
- Leave blank for questions that are not typed into (`radio`, `checkbox`, `date`, `automatic`)
- Use `=` to force length
- Must be `=5` when the FieldType is `hourmin` — a time is always `hh:mm`

**Examples:**
- `80` for a name field
- `=10` for a phone number
- `255` for a comments field
- `=3` user must enter 3 characters and 3 characters will always be saved in the database
- `5` for a `text_decimal` such as `120.5` — **the decimal point counts toward the limit**

---

### Responses

Defines the response options for radio, checkbox, and combobox questions.

#### Static Responses (Traditional)

For hardcoded response options, use the format:
```
value1:Display Text 1
value2:Display Text 2
value3:Display Text 3
```

**Important Rules:**
- Each response on a new line
- Format: `value:text`
- No spaces before the value
- No space immediately after the colon
- Values must be unique
- For radio buttons, values are typically integers: `1`, `2`, `3`, etc.
- For checkboxes, values can be any unique identifier

**Example (Radio Button):**
```
1:Yes
2:No
3:Don't Know
```

**Example (Checkbox):**
```
A:Mosquito net
B:Bed
C:Blanket
D:Pillow
```

#### Dynamic Responses (CSV or Database)

For response options loaded from CSV files or database tables, use a multi-line format:

**CSV Example:**
```
source:csv
file:mrcvillage.csv
filter:region = [[region]]
filter:mrccode = [[mrccode]]
display:villagename
value:vcode
distinct:true
empty_message:No villages found for this region and MRC
dont_know:-7, Don't know which village
not_in_list:-99, Village not in this list
```

**Database Example:**
```
source:database
table:hh_members
filter:hhid = [[hhid]]
filter:sex = 1
filter:census_age >= 15
display:participantsname
value:uniqueid
empty_message:No eligible mothers found in this household
```

**Dynamic Response Parameters:**

| Parameter | Description | Example |
|-----------|-------------|---------|
| `source` | Source type: `csv` or `database` | `source:csv` |
| `file` | CSV filename (for CSV source) | `file:villages.csv` |
| `table` | Table name (for database source) | `table:hh_members` |
| `filter` | Filter condition (can have multiple) | `filter:region = [[region]]` |
| `display` | Column to show to user | `display:villagename` |
| `value` | Column value to save | `value:vcode` |
| `distinct` | Remove duplicates (default: true) | `distinct:true` |
| `empty_message` | Message when no options found | `empty_message:No options available` |
| `dont_know` | Add "Don't know" option | `dont_know:-7, Don't know` |
| `not_in_list` | Add "Not in list" option | `not_in_list:-99, Other` |

**Filter Operators:**
- `=` (equals)
- `!=` or `<>` (not equals)
- `>` (greater than)
- `<` (less than)
- `>=` (greater than or equal)
- `<=` (less than or equal)
- `in` (value is in a comma-separated list)
- `not in` (value is not in a comma-separated list)

**Excluding options already used elsewhere (`in` / `not in`):**

These operators treat the filter value as a **comma-separated list**, so a single
placeholder can include or exclude several rows at once. The list normally comes from
an `automatic` field that queries the database.

A worked example — offering only the household members who have not already been
recorded as sleeping under an earlier net:

```
FieldName: used_linenums   (QuestionType: automatic, FieldType: text)
Responses:
calc:query
sql:SELECT group_concat(sleptunder) FROM nets WHERE hhid = @hhid AND netnum != @netnum AND sleptunder IS NOT NULL
param:@hhid = hhid
param:@netnum = netnum
```

```
FieldName: sleptunder   (QuestionType: checkbox, FieldType: text)
Responses:
source:database
table:hh_members
filter:hhid = [[hhid]]
filter:linenum not in [[used_linenums]]
display:participantsname
value:linenum
```

**Notes:**
- The field supplying the list must appear **before** the question that filters on it.
- An **empty list is safe**: `not in` with nothing to exclude offers every row, and `in`
  with an empty list offers none. This is the normal case for the first record of a
  repeating section.
- Checkbox answers are stored comma-separated, so `group_concat` over several records
  flattens correctly (`"2"` + `"3,4"` → `"2,3,4"`).
- Use `!=` rather than `<>` inside `sql:` — the SQL is written to the XML **without
  escaping**, and `<` would produce an invalid XML file. Keep the SQL on one line.

**Filter Value Placeholders:**
Use `[[fieldname]]` to reference values from previous questions:
- `filter:region = [[region]]` - filters where region equals the value selected in the region question
- `filter:hhid = [[hhid]]` - filters where hhid matches the current household ID

---

### Input Masking (Text Fields)

You can apply input masks to `text` type questions using the `Responses` column. This helps surveyors follow a specific format (like barcodes or IDs) and automatically inserts fixed characters like dashes.

**Syntax:**
```
mask:PATTERN
```

#### Mask Pattern Syntax

The syntax uses a "regex-style" approach to avoid ambiguity with literal text.

- **Placeholders**: Wrap any valid regular expression character class in square brackets `[]`. Each pair of brackets represents **exactly one character**.
  - `[0-9]` : Exactly one digit.
  - `[A-Z]` : Exactly one letter.
  - `[A-Z0-9]` : Exactly one alphanumeric character.
- **Literals**: Anything outside of square brackets is treated as literal text.

#### Features

1. **Explicit Literals**: You can safely use any character as literal text. For example, `Part A: [0-9]` will auto-populate `Part A: ` and then wait for a digit.
2. **Auto-population**: If a mask starts with literal characters (like `R21-`), these are automatically filled in when the question loads.
3. **Auto-insertion**: As the user types, literals in the middle (like the second `-`) are automatically inserted.
4. **Uppercase Enforcement**: All input is automatically converted to uppercase.

**Example:**
To validate a format like `R21-123-A1B2`:

```
mask:R21-[0-9][0-9][0-9]-[A-Z0-9][0-9A-Z][A-Z0-9][A-Z0-9]
```

**XML Output:**
```xml
<question type='text' fieldname='barcode' fieldtype='text'>
    <text>Enter R21 STUDY barcode</text>
    <maxCharacters>=12</maxCharacters>
    <mask value="R21-[0-9][0-9][0-9]-[A-Z0-9][0-9A-Z][A-Z0-9][A-Z0-9]" />
</question>
```

---

### Automatic Calculations

For questions with `QuestionType: automatic`, use the `Responses` column to define the calculation logic.

#### 1. Constant Value
Assigns a static value to the field.

```
calc:constant
value:1
```

#### 2. Lookup Value
Copies a value from another field.

```
calc:lookup
field:participant_name
```

Copying a `checkbox` field copies every selected value, comma-joined
(`"1,3"`) — the same form used everywhere else a calculation reads one.

#### 3. SQL Query
Executes a SQL query against the local database.

```
calc:query
sql:SELECT count(*) FROM members WHERE hhid = @hhid
param:@hhid = hhid
```

#### 4. Math Calculation
Performs basic arithmetic (+, -, *, /) on two or more values.

```
calc:math
operator:+
part:lookup price
part:constant 10
```

#### 5. Concatenation
Joins multiple text values.

```
calc:concat
separator:, 
part:lookup first_name
part:lookup last_name
```

#### 6. Case Logic
Conditional logic (like a switch/case statement).

```
calc:case
when:age < 18 => Minor
when:age >= 18 => Adult
else:Unknown
```

**A `when` is one comparison** — `field operator value => result`. There is no
`and`, `or` or bracketing. Conditions are tested in order and the first match
wins, which is how you build an AND: test for each *failure* and let `else`
carry the success.

```
calc:case
when:sex != 1 => 0
when:consented != 1 => 0
else:1
```

For an OR, give the group its own automatic field and test that field instead.

**Notes**

- Operators: `=`, `!=`, `<>`, `>`, `<`, `>=`, `<=`, `contains`, `does not
  contain`. `<>` and `!=` mean the same thing.
- Put spaces around the operator — `when:age>=18 => 1` is rejected.
- Do not quote values. `when:sex = "1"` compares against a three-character
  string and never matches.
- A field with no answer does not match a numeric comparison, so an inverted
  condition such as `when:sex != 1 => 0` fires — the calculation fails safe.
- The field is computed when navigation reaches it, so it must sit **below**
  every field it reads.
- For a `checkbox` field, `=`/`!=` compare against every selected value
  joined with commas (`"1,3"`), the same form a skip condition uses — **not**
  membership in the list. This is a plain equality test, so it only does what
  you want when the value is a choice that excludes the others, such as a
  "none of these" option enforced by a `logic_check` against the rest of the
  responses:

  ```
  calc:case
  when:screen_cab_drug2 != 99 => 0    # 99 = "not receiving any of these drugs"
  else:1
  ```

  For genuine membership — "was option 5 selected, regardless of what else
  was" — use `contains` / `does not contain` directly instead:

  ```
  calc:case
  when:symptoms contains 5 => 1    # 5 = Vomiting, selected alongside anything else
  else:0
  ```

  Unlike skip's `'contains'` syntax, the operator here is not quoted — it
  follows the same bare, unquoted convention as every other `when` operator.

#### 7. Age From Date
Calculates age in years based on a date field.

```
calc:age_from_date
field:dob
value:today
```

#### 8. Age At Date
Calculates age in years at a specific reference date.

```
calc:age_at_date
field:dob
value:visit_date
```

#### 9. Date Offset
Creates a new date by adding or subtracting time from a source date.

**Format:** `[+/-][number][unit]`
- Units: `d` (days), `w` (weeks), `m` (months), `y` (years)

```
calc:date_offset
field:vx_dose1_date
value:+28d
```

#### 10. Date Difference (Duration)
Calculates the time elapsed between two dates in specific units.

**Parameters:**
- `field`: Start date
- `value`: End date (or `today`)
- `unit`: Unit of time (`d`=days, `w`=weeks, `m`=months, `y`=years)

```
calc:date_diff
field:admission_date
value:today
unit:d
```

#### 11. Date Part (Extract)
Extracts a single component from a date field.

**Parameters:**
- `field`: The date field to extract from (or `today`)
- `unit`: Which component (`yyyy`=4-digit year, `yy`=2-digit year, `mm`=month, `dd`=day,
  `doy`=day of year)

```
calc:date_part
field:dob
unit:mm
```

Not the same thing as the [Computed Automatic Variables](#computed-automatic-variables-yyyy-yy-mm-dd-doy)
`yyyy`/`yy`/`mm`/`dd`/`doy` fields, which take no `calc:` block at all and always mean
"today," fixed forever once set. This extracts from *any* named date field and, like every
`calc:` field in this generator, recomputes on every edit — there is currently no way to
author `preserve: true` from the Excel dictionary at all (a pre-existing gap in the
generator, not specific to this type). If a value needs to survive an edit unchanged, a
Computed Automatic Variable is the only option that does that today; use `date_part` when
tracking a live source field is exactly what you want.

#### 12. Timestamp
Stamps the current date-and-time the moment this row is reached in the questionnaire — a
mid-questionnaire timestamp, at whatever position you place the row. No parameters.

**Requirements:** `FieldType` must be `datetime`.

```
calc:timestamp
```

This is the one calc type that always generates with `preserve='true'` baked in, regardless
of the pre-existing "no `preserve: true` from Excel" gap noted above for `date_part` — a
"time this section was reached" stamp should freeze on first capture, the same way
`starttime`/`stoptime` do, not silently jump to the edit time if the record is opened again
later. See [Custom Timestamp Fields](#custom-timestamp-fields) for the full explanation of
why this exists as its own type rather than a `date_part`/`constant` variant.

---

### LowerRange

Minimum value for numeric validation or minimum date for date questions.

#### For Numeric Fields

**Requirements:**
- Must be a number (integer or decimal)
- Used with `UpperRange` to create a validation range
- **Set both or neither.** A lower bound with a blank upper bound is written as
  `maxvalue='-9'`, which rejects every answer and makes the question
  impossible to complete — so this is an error.
- Not allowed on `hourmin`, where the format already fixes the valid values
- A `text_integer` or `text_decimal` field with **no** range warns, unless its
  MaxCharacters is fixed (`=10`) — a fixed length means an identifier, such as
  a household ID or phone number, where a numeric range is meaningless

**Example:**
- `LowerRange: 0`
- `UpperRange: 120`
- Result: Value must be between 0 and 120

#### For Date Fields

**Requirements:**
- Must be in special date offset format OR a hard-coded date
- Offset Format: `[+/-][number][unit]` where unit is `d` (days), `w` (weeks), `m` (months), or `y` (years)
- Hard-coded Date Format: `yyyy-mm-dd`
- Special value: `0` means today's date

**Examples:**
- `0` - Today
- `-1y` - One year ago
- `+6m` - Six months from now
- `-30d` - 30 days ago
- `+2w` - Two weeks from now
- `2023-01-01` - Specific date

---

### UpperRange

Maximum value for numeric validation or maximum date for date questions.

Same format and requirements as `LowerRange`.

**Example (Numeric):**
- `LowerRange: 18`
- `UpperRange: 65`
- Validation: Age must be between 18 and 65

**Example (Date):**
- `LowerRange: -5y`
- `UpperRange: 0`
- Validation: Date must be between 5 years ago and today

---

### LogicCheck

Defines validation rules that compare field values and show error messages if conditions are not met.

#### Simple Logic Check

**Format:**
```
expression; 'error message'
```

**Examples:**
```
age >= 18; 'Participant must be 18 or older'
end_date > start_date; 'End date must be after start date'
password = confirm_password; 'Passwords must match'
```

#### Compound Logic Check

Use `and` / `or` operators for complex conditions:

**Examples:**
```
(age >= 18 and age <= 65); 'Age must be between 18 and 65'
(month = 2 and day = 29); 'February only has 28 days in non-leap years'
(status = 1 or status = 2); 'Invalid status value'
```

#### Multiple Logic Checks

It is possible to have more than one logic check per question. Just ensure each logic check is on a separate line.

**Example:**
```
vx_dose3_date < vx_dose2_date; 'Date of dose 3 cannot be before date of dose 2'
vx_dose3_date < dob; 'Date of vaccination cannot be before date of birth!'
```


#### Unique Check

Ensures the value is unique in the database table:

**Format:**
```
unique; 'error message'
```

**Example:**
```
unique; 'This ID has already been used'
```

**Operators:**
- `=` (equals)
- `!=` or `<>` (not equals)
- `>` (greater than)
- `<` (less than)
- `>=` (greater than or equal)
- `<=` (less than or equal)
- `and` (logical AND)
- `or` (logical OR)

**Important Notes:**
- Referenced field names must exist in the same worksheet
- Referenced fields must appear **before** the current question
- Error message must be enclosed in single quotes
- Expression and message must be separated by a semicolon

---

### DontKnow

Adds a "Don't Know" response option to the question.

**Valid Values:**
- `True` / `TRUE` - Show the "Don't Know" response
- `False` / `FALSE` - Don't show the response
- Leave blank if not needed

**Example:**
When set to `True`, a "Don't Know" response appears that allows the user to skip the question without providing an answer.

---

### Refuse

Adds a "Refuse to Answer" response option to the question.

**Valid Values:**
- `True` / `TRUE` - Show the "Refuse to Answer" response
- `False` / `FALSE` - Don't show the response
- Leave blank if not needed

**Example:**
When set to `True`, a "Refuse to Answer" response appears for sensitive questions.

---

### Optional

Marks a **text** question as skippable: the user can press Next with the field left blank.
Replaces the old NA column, which added a "Not Applicable" response option that the app never
actually read -- it had no effect on any survey.

**Valid Values:**
- `True` / `TRUE` - The question may be left blank
- `False` / `FALSE`, or blank - The question must be answered (the default)
- Only meaningful for `QuestionType = text`. Choice questions (`radio`/`checkbox`/`combobox`)
  already have DontKnow/Refuse above for an explicit non-answer.

**Example:**
```
FieldName: notes | QuestionType: text | FieldType: text | Optional: TRUE
```

**Note on `comments`:** a field named `comments` used to be always-optional automatically,
regardless of this column. That is no longer true -- if your dictionary has a `comments` field
that should stay skippable, set `Optional` to `TRUE` on it explicitly. The generator warns (but
does not error) if it finds a `comments` field without Optional set.

**Backward compatibility:** a dictionary written before this column existed still has `NA` as
the header there instead of `Optional`. That is still accepted -- the column's contents are
simply ignored, exactly as before.

---

### Skip

Defines skip patterns (branching logic) that control question flow based on previous responses.

#### Preskip

Evaluated **before** showing the question. If the condition is true, skip to the target question without showing this question.

**Format:**
```
preskip: if_condition, skip_to_target
```

**Examples:**
```
preskip: if has_children = 0, skip to occupation
preskip: if eligible = 0, skip to end
```

**How it works:**
- If `has_children` equals `0`, then skip to `occupation`
- The question is never shown if the condition is true

#### Postskip

Evaluated **after** answering the question. If the condition is true, skip to the target question.

**Format:**
```
postskip: if_condition, skip_to_target
```

**Examples:**
```
postskip: if pregnant = 1, skip to pregnant_date
postskip: if owns_car = 0, skip to owns_house
```

**How it works:**
- After the user answers the current question, the condition is evaluated
- If `pregnant` equals `1`, then skip to `pregnant_date`


#### Multiple Skip Conditions

You can have multiple skip conditions (one per line):

**Example:**
```
preskip: if age < 18, skip to end
postskip: if consent = =1, skip to age
```

**Operators:**
- `=` (equals)
- `>` (greater than)
- `>=` (greater than or equal)
- `<` (less than)
- `<=` (less than or equal)
- `<>` (not equals)
- `'contains'` (string contains - checks if value is present in a comma-separated list)
- `'does not contain'` (string does not contain - checks if value is NOT present in a comma-separated list)

**Using 'contains' and 'does not contain':**

These operators are used for checkbox questions where multiple values are stored as a comma-separated string.

**Example:**
```
postskip: if symptoms 'contains' 1, skip to fever
postskip: if symptoms 'does not contain' 9, skip to cough
```

If the `symptoms` checkbox question has values stored as `"1,2,3"` (user selected Fever, Headache, and Fatigue):
- `symptoms 'contains' 1` → **true** (1 is in the list) → skip to `fever`
- `symptoms 'does not contain' 9` → **true** (9 is NOT in the list) → skip to `cough`

**Important Notes:**
- The field to check must exist and appear **before** the current question
- The target field to skip to must exist and appear **after** the current question
- Cannot skip to the same question
- For checkbox questions, values are stored as comma-separated strings (e.g., "1,2,3")

#### Skipping to the end of the questionnaire

`skip to end` is a reserved target, not a FieldName -- use it to end the interview early once a
screening question rules the respondent out, without needing a real question to land on:

```
postskip: if eligible = 0, skip to end
```

Unlike an ordinary skip, `end` doesn't jump past the remaining questions -- it walks through them,
computing any calculated fields and the reserved trailing variables (`uniqueid`, `swver`,
`survey_id`, `lastmod`, `stoptime`) exactly as if navigation had reached them normally, so the
record is saved complete. Because `end` is reserved for this, it cannot be used as a FieldName.

---

### Comments

Developer comments and notes. This column is not processed by the application.

**Use for:**
- Notes about the question
- References to documentation
- Implementation notes
- Reminders

**Example:**
```
TODO: Verify age range with client
See specification document section 3.2
This field is auto-calculated by the system
```

---

## The CRFS Worksheet


This section explains how to set up the **`crfs`** worksheet in the Excel data dictionary so that the
DataKollecta-SurveyGen app can generate a correct `survey_manifest.gistx`, and so the DataKollecta survey app
administers your forms the way you expect.

It covers **three distinct form scenarios**:

- **A. Base (top-level) form** — e.g. an enrollment or household form.
- **B. Repeating child form** — a form entered *N* times per parent (e.g. household members).
- **C. One-off "sister" form** — a form linked to a parent by a shared field but entered **once**,
  often only in certain scenarios (e.g. a vaccine-coverage form done once per enrollee, and only
  when the enrollee was flagged for it).

---

### 1. Overview — how it all flows

```
Excel data dictionary
   ├── <table>_dd worksheets   →  one XML form file per table  (e.g. enrollee.xml)
   └── crfs worksheet          →  the "crfs" array in survey_manifest.gistx
                                       │
                                       ▼
                         DataKollecta app reads survey_manifest.gistx
                         and uses the crfs entries to drive
                         navigation, linking, ID generation,
                         and auto-repeat behavior.
```

Each **row** of the `crfs` worksheet describes **one form** (one table). The columns tell the app:

- where the form sits in the hierarchy (base vs. child),
- how its records link to a parent,
- how its unique ID is built,
- whether/how it auto-repeats,
- and (optionally) the condition under which it is offered.

A form only appears in the app if its `<tablename>.xml` is included in the manifest's `xmlFiles`
list (DataKollecta-SurveyGen handles this automatically when the `_dd` worksheet exists).

---

### 2. Column reference

Columns are listed in the order used in the data dictionary. **"Required"** means the app/config
needs it for that row to work; blank cells are fine for any column marked Optional.

| Column | Required? | Default | Purpose |
|--------|-----------|---------|---------|
| `display_order` | Recommended | `0` | Integer controlling the order forms appear in the app menu (10, 20, 30…). Also controls the order auto-repeat children are processed. |
| `tablename` | **Required** | — | Unique table id. Must match the `_dd` worksheet / XML filename (without `.xml`). |
| `displayname` | **Required** | — | Human-readable name shown to the user. |
| `primarykey` | **Required** | — | Primary key field(s). Comma-separated for a composite key (e.g. `hhid,linenum`). |
| `idconfig` | Optional | (none) | JSON object describing how to build this form's own unique ID. See §7. Usually only on **base** forms. |
| `isbase` | Optional | `0` | `1` = top-level form; `0` = child/linked form. |
| `linkingfield` | Optional* | (none) | The field that links a child record to its parent. *Required for child forms.* |
| `parenttable` | Optional* | (none) | The parent form's `tablename`. *Required for child forms.* |
| `incrementfield` | Optional | (none) | Field that auto-increments per parent (e.g. `linenum`). **Repeating children only.** |
| `requireslink` | Optional | `0` | `1` = before opening this form, the user must first pick a parent record. *Required (=1) for any child/sister form.* |
| `repeat_count_field` | Optional | (none) | Name of the field **in the parent form** holding how many child records to create. **Repeating children only.** |
| `auto_start_repeat` | Optional | `0` | How the repeat loop starts: `0` off, `1` prompt, `2` force. See §9. **Repeating children only.** |
| `repeat_enforce_count` | Optional | `1` | How strictly the child count is enforced: `0`/`1`/`2`/`3`. See §9. **Repeating children only.** |
| `display_fields` | Optional | (none) | Comma-separated fields used to summarize each record in selection lists (e.g. `subjid,participantsname`). |
| `entry_condition` | Optional | (none) | `field=value` rule limiting which parent records this form can be attached to. See §8. **Only used when `requireslink=1`.** |

> The app stores every column verbatim into a SQLite `crfs` table and reads it back at runtime; only
> `tablename`, `displayname`, and `primarykey` are practically mandatory. Everything else defaults to
> "off/none" when blank, so **leave columns blank when a pattern doesn't use them.**

---

### 3. Which pattern do I need?

```
Is this a top-level form (no parent)?
   └── YES → Scenario A: Base form  (isbase=1)

Is it linked to a parent table by a shared field?
   ├── Is it entered MANY times per parent, based on a count? → Scenario B: Repeating child
   └── Is it entered just ONCE per parent (maybe only sometimes)? → Scenario C: One-off "sister"
```

The difference between B and C is **only** whether the repeat/increment columns are filled in.
Both use `requireslink=1`, `linkingfield`, and `parenttable`.

---

### 4. Scenario A — Base (top-level) form

A base form has no parent. The user opens it directly and a new unique ID is generated for it.

**Set these columns:**
- `isbase` = `1`
- `primarykey` = the form's own key (e.g. `subjid` or `hhid`)
- `idconfig` = JSON to build the ID (see §7)
- `display_fields` = fields to show in record lists
- `display_order` = e.g. `10`

**Leave blank:** `parenttable`, `incrementfield`, `requireslink`, all `repeat_*` columns,
`entry_condition`. (`linkingfield` may be set to the key other forms link on — see note below.)

**Example — `enrollee` (R21):**

| Column | Value |
|--------|-------|
| display_order | 10 |
| tablename | enrollee |
| displayname | Enrollee |
| primarykey | subjid |
| idconfig | `{"prefix":"","fields":[{"name":"deviceid","length":3},{"name":"mrc","length":3}],"incrementLength":4}` |
| isbase | 1 |
| linkingfield | fpbarcode1_r21 |
| display_fields | startdate,participantsname |

> **Note on `linkingfield` for a base form:** It is harmless to set it. It documents which field
> children link on, and it is *not* used to require a parent (that only happens when
> `requireslink=1`). In the R21 example the enrollee's `fpbarcode1_r21` is the value the sister form
> links to.

---

### 5. Scenario B — Repeating child form (the classic case)

Use this when one parent record produces a **known number of child records**, e.g. a household with
`nmembers` members. The parent form asks for the count; the app then loops to collect that many
child records.

**Set these columns:**
- `isbase` = `0`
- `primarykey` = **composite**: parent key + the increment field (e.g. `hhid,linenum`)
- `linkingfield` = the shared field inherited from the parent (e.g. `hhid`)
- `parenttable` = the parent's `tablename` (e.g. `hh_info`)
- `incrementfield` = the per-parent counter field (e.g. `linenum`)
- `requireslink` = `1`
- `repeat_count_field` = the **parent's** count field (e.g. `nmembers`)
- `auto_start_repeat` = `0`/`1`/`2` (see §9)
- `repeat_enforce_count` = `0`/`1`/`2`/`3` (see §9)
- `display_fields`, `display_order`

**Leave blank:** `idconfig` (the key comes from the parent link + increment, not a generated ID),
`entry_condition` unless you also want to restrict which parents are eligible.

**Example — `hh_members` (PRISM):**

| Column | Value |
|--------|-------|
| display_order | 30 |
| tablename | hh_members |
| displayname | Household Members |
| primarykey | hhid,linenum |
| isbase | 0 |
| linkingfield | hhid |
| parenttable | hh_info |
| incrementfield | linenum |
| requireslink | 1 |
| repeat_count_field | nmembers |
| auto_start_repeat | 2 |
| repeat_enforce_count | 2 |
| display_fields | participantsname |

**User flow:** Fill `hh_info` → enter `nmembers = 3` → on save, the app launches "Household
Member 1 of 3", then 2 of 3, then 3 of 3, each pre-filled with the parent `hhid` and the next
`linenum`.

---

### 6. Scenario C — One-off "sister" form (linked, not repeated)

Use this when a form is linked to a parent by a shared field, but is entered **only once** per
parent — and possibly only in certain scenarios. This is the R21 `vaccine_coverage` case: it shares
`fpbarcode1_r21` with the `enrollee`, is done once, and only for enrollees flagged with `vac_cov=1`.

**This pattern is fully supported by the current app — no special handling or code change is
needed.** It is simply a linked child with all the repeat/increment columns left blank.

**Set these columns:**
- `isbase` = `0`
- `primarykey` = the **single** shared field (the linking field itself — there is no per-parent
  counter, so no composite key)
- `linkingfield` = the shared field (e.g. `fpbarcode1_r21`)
- `parenttable` = the parent's `tablename` (e.g. `enrollee`)
- `requireslink` = `1`
- `display_fields`, `display_order`
- `entry_condition` = optional `field=value` rule restricting which parents are eligible (see §8)

**Leave blank:** `incrementfield`, `repeat_count_field`, `auto_start_repeat`,
`repeat_enforce_count`, and usually `idconfig` (the record's key *is* the inherited link value).

**Example — `vaccine_coverage` (R21):**

| Column | Value |
|--------|-------|
| display_order | 20 |
| tablename | vaccine_coverage |
| displayname | Vaccine Coverage |
| primarykey | fpbarcode1_r21 |
| isbase | 0 |
| linkingfield | fpbarcode1_r21 |
| parenttable | enrollee |
| requireslink | 1 |
| display_fields | fpbarcode1_r21 |
| entry_condition | vac_cov=1 |

**User flow:**
1. User taps **Vaccine Coverage** in the form menu.
2. Because `requireslink=1`, the app shows a **parent selection list** of `enrollee` records —
   filtered to only those where `vac_cov = 1` (the `entry_condition`).
3. User picks an eligible enrollee. The chosen `fpbarcode1_r21` is **pre-filled** into the sister
   form.
4. User completes the form once and saves. No repeat loop runs (the repeat columns are blank).

> **"Only once" is procedural, not enforced.** The app does not prevent a second Vaccine Coverage
> record for the same enrollee — eligible parents stay in the selection list even after one has been
> entered. If strict one-per-parent is important, handle it by training/SOP, or by setting an
> `entry_condition` on a flag your workflow clears after entry.

---

### 7. `idconfig` reference (ID generation)

The `idconfig` field contains a JSON object that defines how a record's unique ID is built. The ID
is assembled left-to-right from three parts: a static **prefix**, one or more **field** values taken
from the survey answers, and an auto-incrementing **sequence number**. It is typically only used on
**base** forms.

**Structure:**
```json
{
  "prefix": "STRING",
  "fields": [
    {"name": "field_name", "length": INT},
    ...
  ],
  "incrementLength": INT
}
```

**Parameters:**
- `prefix`: Static string prepended to every ID (e.g., `"GL"`, `"HH"`, `"3"`, or `""` for none).
- `fields`: Array of field objects defining which survey fields to embed in the ID, in order.
  - `name`: Field name from the questionnaire. Its **value is read from the current survey's
    answers**, so the field must be answered before the ID is generated.
  - `length`: Fixed length for this part. Shorter values are **padded with leading zeros**
    (e.g. `5` → `05`); longer values are **truncated** to fit.
- `incrementLength`: Number of digits for the auto-incrementing sequence appended at the end. The app
  finds the **next available number** for that prefix+fields combination and zero-pads it to this
  length (e.g. `3` → `001`, `002`, … `999`). This is a **required** key, but set it to `0` when the
  ID has no incrementing part.

> **Order matters.** The final ID is exactly `prefix` + each `fields` value (in array order) +
> the increment. Lay the `fields` array out in the same order you want the digits to appear.

**Example 1: With Increment**
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

This generates IDs like: `GL0105001`, `GL0105002`, etc.
- Prefix: `GL`
- Community: `01` (padded to 2 digits)
- Village: `05` (padded to 2 digits)
- Increment: `001`, `002`, `003` (3 digits, auto-assigned)

**Example 2: Without Increment**
```json
{
  "prefix": "3",
  "fields": [
    {"name": "mrccode", "length": 2},
    {"name": "vcode", "length": 2},
    {"name": "hhnum", "length": 4}
  ],
  "incrementLength": 0
}
```

This generates IDs like: `301050001` (no auto-increment — every part comes from survey answers).
- Prefix: `3`
- MRC code: `01` (padded to 2 digits)
- Village code: `05` (padded to 2 digits)
- Household number: `0001` (padded to 4 digits)
- No increment part (`incrementLength: 0`)

> **Tip — choosing `incrementLength`:** Use `0` only when the prefix + fields are guaranteed to be
> unique on their own (as with the manually entered `hhnum` above). If two records could otherwise
> produce the same ID, give `incrementLength` enough digits to cover the largest expected count
> (e.g. `3` allows up to 999 records per prefix+fields combination).

> **One-off children (Scenario C) normally omit `idconfig`** — their key is the inherited
> `linkingfield` value, not a freshly generated ID.

**Example 3: A reinstall-resilient ID, using `yy` / `doy`**

`incrementLength`'s counter is a **local** count — the app looks at what's already on *this
device*. Reinstalling the app deletes its local data, so the counter silently restarts at 1,
and a freshly generated ID can collide with one already generated (and possibly already
synced) before the reinstall. Folding
[`yy`/`doy`](#computed-automatic-variables-yyyy-yy-mm-dd-doy) into `idconfig.fields` alongside
an interviewer/device code shrinks that risk: the counter only has to stay collision-free
**within one interviewer's one calendar day**, since a collision now requires reusing the
exact same day *and* interviewer, not just the same device ever — a real, if partial, risk
reduction, not a guarantee.

```json
{
  "prefix": "GX",
  "fields": [
    {"name": "nn", "length": 2},
    {"name": "yy", "length": 2},
    {"name": "doy", "length": 3}
  ],
  "incrementLength": 2
}
```
Generates IDs like `GX07260451` — interviewer `07`, year `20`**`26`**, day-of-year
`045`, then a 2-digit increment that resets every day since `fields` (and therefore the
counter's base) changes daily.

See [Computed Automatic Variables](#computed-automatic-variables-yyyy-yy-mm-dd-doy) above for how to
declare `yy`/`doy`, their exact format, and why they must not be given a `calc:` block. This
is unrelated to `databaseName` (see [Versioning a survey](#versioning-a-survey)) — this
changes what an ID is built from, not which database the record lands in.

---

### 8. `entry_condition` reference (when is a form offered?)

`entry_condition` restricts **which parent records** a linked form can be attached to. It is the
mechanism for "this form is only done in certain scenarios."

**Syntax:** a single equality, `field=value`.

```
vac_cov=1
enrolled=1
```

**Semantics (exactly how the app evaluates it):**
- It is read only on forms where **`requireslink=1`** (it filters the parent-selection list). On a
  base form, or any form with `requireslink=0`, it is **ignored**.
- The app loads all records from `parenttable`, then keeps only those where
  `parentRecord[field] == value`. The field name is matched **case-insensitively**; the value is
  compared as a **string** (so a numeric `1` must be stored as `1`).
- The surviving records' `linkingfield` values become the selectable list.
- If **no** parent record matches, the user sees a "no eligible IDs" message and cannot start the
  form — which is exactly the desired behavior when the scenario doesn't apply.

**Limitations (by design — keep conditions simple):**
- Only a **single** `field=value` test. No `!=`, `<`, `>`, ranges, `AND`/`OR`, or multiple
  conditions.
- Only equality. If you need richer logic, model it as a single derived flag field in the parent
  form and test that flag.

---

### 9. Auto-repeat reference (Scenario B only)

These two columns only matter when `repeat_count_field` is set and points at a valid count in the
parent form. For one-off sister forms, leave them blank.

### `auto_start_repeat` — how the repeat loop starts

| Value | Behavior |
|-------|----------|
| `0` | **Off.** No automatic loop. (Blank = off.) |
| `1` | **Prompt.** After the parent is saved, ask "You indicated X records — add them now?" with Add Now / Add Later. |
| `2` | **Force.** Immediately launch the child loop after the parent is saved, without asking. |

### `repeat_enforce_count` — what happens if the count doesn't match

| Value | Behavior |
|-------|----------|
| `0` | **Flexible.** Any number of children allowed; no check. |
| `1` | **Warn** *(default if blank)*. On mismatch, warn and offer to update the parent's count to the actual number entered. |
| `2` | **Force.** Block exit until exactly the stated number of children are entered. There is **no** "Exit Anyway" escape in this mode. |
| `3` | **Auto-sync.** Update the parent's count field to match the number actually entered, with no choice offered -- then show a single acknowledgement dialog telling the interviewer the count was changed and to what. Unlike mode `1`, there is no "Exit Anyway" to leave the mismatch unresolved. |

**Stopping the loop early, in modes `1` and `3`.** The child form's own Cancel/X dialog
already tells the interviewer the consequence before they confirm: "You've entered X of Y
&lt;records&gt;. If you stop now, ..." followed by mode `1`'s "you'll be asked whether to update
the count" or mode `3`'s "the count will be automatically updated to X." Mode `2` never shows
this, since the loop blocks completion instead.

**Typical repeating setup:** `repeat_count_field=nmembers`, `auto_start_repeat=2`,
`repeat_enforce_count=2`.

#### When the count is reconciled

Modes `1` and `3` reconcile the parent's count in two places:

1. **When the auto-repeat loop ends** — whether the interviewer completed every child or
   left early.
2. **Whenever a child record is saved outside that loop** — one added later through
   *New Survey → \<child form\> → pick the parent ID*, or an existing one edited through
   *Modify Existing Survey*. This is what keeps `nmembers` correct when a seventh household
   member turns up weeks after enrolment.

Mode `2` is enforced only inside the loop, because there is nothing to block outside it.

#### The count question's own range is never violated

Before writing, the app checks the actual number of children against the **`LowerRange`** and
**`UpperRange`** the count question declares in its `_dd` worksheet — the same check the
interviewer's typed answer has to pass, `Responses`-style exceptions included. The count is
**not** written when:

- **Fewer children than `LowerRange`.** `nmembers` declared as `1..30` can never be
  auto-synced to `0`. Inside the repeat loop the interviewer is told they must enter at
  least that many and cannot leave until they have; outside it, nothing is written.
- **More children than `UpperRange`.** The interviewer is warned and the data left for
  correction, rather than storing an impossible count.
- **The count question was skipped**, so its value is NULL. A skip stores NULL, and NULL is
  left alone — a household that answered "no" to `havenets` keeps `nnets` empty even if a
  net record is later added by hand.

A count question with no `numeric_check` (no `LowerRange`/`UpperRange` in its `_dd` row) has
no floor or ceiling, and reconciles unconditionally.

**Design note:** to require at least *N* children of a form, set `LowerRange` on its count
question. There is no separate "minimum entries" column on `crfs`, deliberately — the
minimum belongs next to the question it constrains, and one declaration cannot drift out of
step with another.

---

### 10. Caveats & gotchas

- **Blank ≠ zero only where noted.** Most columns treat blank as "off/none". `display_order` blank
  is treated as `0`; `repeat_enforce_count` blank defaults to `1` (warn).
- **`repeat_enforce_count` does nothing unless `auto_start_repeat` is `1` or `2`.** The loop is
  what triggers reconciliation after the parent is saved; with `auto_start_repeat=0` there is no
  loop, and the count is only reconciled if a child is later saved on its own.
- **`repeat_enforce_count` never writes a count outside the count question's declared
  `LowerRange`/`UpperRange`**, and never fills in a count question that was skipped. See §9.
- **`entry_condition` needs `requireslink=1`** to have any effect.
- **`entry_condition` is equality-only** and string-compared — store the compared value exactly
  (`1`, not `1.0`).
- **One-off forms aren't capped at one record** by the app; enforce "once" via SOP or a clearing
  flag if it matters.
- **Composite vs. single primary key:** repeating children need `parentkey,incrementfield`; one-off
  sister forms use just the single linking field.
- **`tablename` must match** the `_dd` worksheet and the generated `.xml`, or the form won't appear.

---

### 11. Worked examples

#### A. Base (top-level)

**`crfs` worksheet rows:**

| display_order | tablename | displayname | primarykey | idconfig | isbase | linkingfield | parenttable | incrementfield | requireslink | repeat_count_field | auto_start_repeat | repeat_enforce_count | display_fields | entry_condition |
|--|--|--|--|--|--|--|--|--|--|--|--|--|--|--|
| 10 | enrollee | Enrollee | subjid | `{...}` | 1 | fpbarcode1_r21 | | | | | | | startdate,participantsname | |

**Resulting manifest (`survey_manifest.gistx`):**
```json
"crfs": [
  {
    "display_order": 10, "tablename": "enrollee", "displayname": "Enrollee",
    "isbase": 1, "primarykey": "subjid", "linkingfield": "fpbarcode1_r21",
    "idconfig": {"prefix":"","fields":[{"name":"deviceid","length":3},{"name":"mrc","length":3}],"incrementLength":4},
    "display_fields": "startdate,participantsname"
  }]
```

#### B. One-off sister survey

**`crfs` worksheet rows:**

| display_order | tablename | displayname | primarykey | idconfig | isbase | linkingfield | parenttable | incrementfield | requireslink | repeat_count_field | auto_start_repeat | repeat_enforce_count | display_fields | entry_condition |
|--|--|--|--|--|--|--|--|--|--|--|--|--|--|--|
| 10 | enrollee | Enrollee | subjid | `{...}` | 1 | fpbarcode1_r21 | | | | | | | startdate,participantsname | |
| 20 | vaccine_coverage | Vaccine Coverage | fpbarcode1_r21 | | 0 | fpbarcode1_r21 | enrollee | | 1 | | | | fpbarcode1_r21 | vac_cov=1 |

**Resulting manifest (`survey_manifest.gistx`):**
```json
"crfs": [
  {
    "display_order": 10, "tablename": "enrollee", "displayname": "Enrollee",
    "isbase": 1, "primarykey": "subjid", "linkingfield": "fpbarcode1_r21",
    "idconfig": {"prefix":"","fields":[{"name":"deviceid","length":3},{"name":"mrc","length":3}],"incrementLength":4},
    "display_fields": "startdate,participantsname"
  },
  {
    "display_order": 20, "tablename": "vaccine_coverage", "displayname": "Vaccine Coverage",
    "isbase": 0, "primarykey": "fpbarcode1_r21", "linkingfield": "fpbarcode1_r21",
    "parenttable": "enrollee", "requireslink": 1,
    "display_fields": "fpbarcode1_r21", "entry_condition": "vac_cov=1"
  }
]
```

#### C. Repeating children

**Resulting manifest (`survey_manifest.gistx`):**
```json
"crfs": [
  {
    "display_order": 10, "tablename": "hh_info", "displayname": "Household Survey",
    "isbase": 1, "primarykey": "hhid", "linkingfield": "hhid",
    "idconfig": {"prefix":"3","fields":[{"name":"mrccode","length":2},{"name":"vcode","length":2},{"name":"hhnum","length":4}],"incrementLength":0}
  },
  {
    "display_order": 30, "tablename": "hh_members", "displayname": "Household Members",
    "isbase": 0, "primarykey": "hhid,linenum", "linkingfield": "hhid",
    "parenttable": "hh_info", "incrementfield": "linenum", "requireslink": 1,
    "repeat_count_field": "nmembers", "auto_start_repeat": 2, "repeat_enforce_count": 2,
    "display_fields": "participantsname", "entry_condition": "enrolled=1"
  }
]
```



























## Examples

### Example 1: Simple Radio Button Question

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | Optional | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| gender | radio | integer | What is your gender? | | 1:Male<br>2:Female | | | | False | False | False | | |

### Example 2: Text Field with Validation

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | Optional | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| age | text | text_integer | What is your age? | 3 | | 0 | 120 | | False | False | False | | Valid ages 0-120 |

### Example 3: Date Field with Range

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | Optional | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| birth_date | date | date | What is your date of birth? | | | -100y | 0 | | False | False | False | | |

### Example 4: Checkbox Question

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | Optional | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| symptoms | checkbox | text | What symptoms do you have? | | A:Fever<br>B:Cough<br>C:Headache<br>D:Fatigue | | | | False | False | False | | Multiple selection |

### Example 5: Dynamic Responses from CSV

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | Optional | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| village | combobox | text | Select your village | | source:csv<br>file:villages.csv<br>filter:region = [[region]]<br>display:villagename<br>value:vcode | | | | False | False | False | | Cascading dropdown |

### Example 6: Logic Check

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | Optional | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| confirm_age | text | text_integer | Please confirm your age | 3 | | 0 | 120 | confirm_age = age; 'Age does not match!' | False | False | False | | Verification field |

### Example 7: Skip Logic

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | Optional | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| pregnant | radio | integer | Are you pregnant? | | 1:Yes<br>2:No | | | | False | False | False | postskip: pregnant = 2, if, next_section | Skip if not pregnant |

### Example 8: Unique Check

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | Optional | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| participant_id | text | text | Enter participant ID | 20 | | | | unique; 'This ID has already been used in the database!' | False | False | False | | Must be unique |

---

## Output Files

### XML Files

For each worksheet ending in `_dd`, the application generates an XML file (e.g., `hh_info_dd` → `hh_info.xml`).

**Example XML Output:**

```xml
<?xml version='1.0' encoding='utf-8'?>
<survey>

    <question type='radio' fieldname='gender' fieldtype='integer'>
        <text>What is your gender?</text>
        <responses>
            <response value='1'>Male</response>
            <response value='2'>Female</response>
        </responses>
    </question>

    <question type='text' fieldname='age' fieldtype='text_integer'>
        <text>What is your age?</text>
        <maxCharacters>3</maxCharacters>
        <numeric_check>
            <values minvalue='0' maxvalue='120' other_values='0' message='Number must be between 0 and 120!'></values>
        </numeric_check>
    </question>

    <question type='date' fieldname='birth_date' fieldtype='date'>
        <text>What is your date of birth?</text>
        <date_range>
            <min_date>-100y</min_date>
            <max_date>0</max_date>
        </date_range>
    </question>

</survey>
```

**Dynamic Response XML Example:**

```xml
<question type='combobox' fieldname='village' fieldtype='text'>
    <text>Select your village</text>
    <responses source='csv' file='villages.csv'>
        <filter column='region' operator='=' value='[[region]]'/>
        <display column='villagename'/>
        <value column='vcode'/>
    </responses>
</question>
```

### survey_manifest.gistx

The survey manifest file contains metadata about the survey and all forms.

**Example:**

```json
{
  "surveyName": "PRISM CSS 2025-12-01",
  "surveyId": "prism_css_2025_12_01",
  "databaseName": "prism_css.sqlite",
  "xmlFiles": [
    "hh_info.xml",
    "hh_members.xml",
    "nets.xml"
  ],
  "crfs": [
    {
      "display_order": 10,
      "tablename": "hh_info",
      "displayname": "Household Survey",
      "isbase": 1,
      "primarykey": "hhid",
      "linkingfield": "hhid",
      "idconfig": {
        "prefix": "3",
        "fields": [
          {"name": "mrccode", "length": 2},
          {"name": "vcode", "length": 2}
        ],
        "incrementLength": 4
      },
      "requireslink": 0,
      "auto_start_repeat": 0,
      "repeat_enforce_count": 0
    },
    {
      "display_order": 20,
      "tablename": "hh_members",
      "displayname": "Household Members",
      "isbase": 0,
      "primarykey": "hhid,linenum",
      "linkingfield": "hhid",
      "parenttable": "hh_info",
      "incrementfield": "linenum",
      "requireslink": 1,
      "repeat_count_field": "nmembers",
      "auto_start_repeat": 1,
      "repeat_enforce_count": 2,
      "display_fields": "participantsname"
    }
  ]
}
```

### Log File

The `gistlogfile.txt` contains detailed validation results:

```
Log file for: C:\PRISM\Excel\prismcss.xlsx

Checking worksheet: 'hh_info_dd'
No errors found in 'hh_info_dd'

Checking worksheet: 'hh_members_dd'
Be sure to write code for each automatic variable: hhid, linenum
No errors found in 'hh_members_dd'

Successfully generated survey_manifest.gistx

--------------------------------------------------------------------------------
End of log file
--------------------------------------------------------------------------------
```

**Error Example:**

```
Checking worksheet: 'enrollment_dd'
ERROR - FieldName: enrollment_dd has a FieldName that starts with a number: 2ndvisit
ERROR - QuestionText: FieldName 'age' in worksheet 'enrollment_dd' has blank QuestionText.
ERROR - Responses: Invalid static radio button options for 'gender' in table 'enrollment_dd'. Expected format 'number:Statement', found '1: Male'.
ERROR - LogicCheck: FieldName 'confirm_age' in worksheet 'enrollment_dd' has invalid syntax for LogicCheck (missing semicolon): confirm_age = age
```

---

## Error Handling and Validation

### Common Validation Errors

#### Silent-Failure Checks

These catch mistakes that produce a valid-looking XML file but broken data collection:

- **Automatic field with no calculation**: an `automatic` question whose `Responses`
  column is blank, or holds response options instead of a `calc:` block, is never given a
  value. Any skip that tests it then **fails open** — a skip whose field is unanswered
  never fires, so the question it was meant to guard is asked of everyone. Exempt:
  built-in fields (`starttime`, `startdate`, `stoptime`, `uniqueid`, `swver`, `survey_id`,
  `lastmod`) and fields the manifest supplies (`linkingfield`, `incrementfield`,
  `primarykey`, and `idconfig` parts).
- **Selection question with no responses**: a `radio`, `checkbox`, or `combobox` with an
  empty `Responses` column cannot be answered at all — the interviewer reaches a question
  with nothing to select. Add options, or a `source:csv` / `source:database` block.
- **Preskip that tests its own field**: preskips run *before* the question is shown, so on
  a new record the field is still unanswered and the skip never fires. On an existing
  record the stored value *does* fire it, and the jump clears every answer it passes over
  — including that question's own. Use a `postskip` instead.
- **MaxCharacters on a selection question** *(warning)*: only typed input is length-limited,
  so `MaxCharacters` on a `radio`/`checkbox`/`combobox` is ignored. Usually a sign the
  QuestionType should have been `text`.
- **Reserved automatic variable declared** *(warning)*: a FieldName such as `starttime` has
  built-in meaning. The generator writes it into the questionnaire itself, in the correct
  position, and drops the declared row — so the placement in the spreadsheet does not
  matter and there can never be two. Never an error; the row can stay. A `calc:` block on
  one is called out separately, because the calculation is dropped rather than applied.
- **Skip that tests a reserved variable**: which questions get asked must not depend on a
  value the generator supplies. The trailing five are still empty while the questionnaire
  is being answered, so the skip would never fire and the question it guards would be
  asked of everyone; and branching on `starttime` or `startdate` would make one deployed
  package ask different questions on different days. Regenerate the questionnaire instead.
- **Trailing reserved variable read anywhere in a question**: `uniqueid`, `swver`,
  `survey_id`, `lastmod` and `stoptime` are written *after* the last question, so
  anything reading one during the interview sees an empty value — in a logic check, a
  calculation, a response filter, or the question text itself. A logic check never fires
  and looks like a validation that always passes; a calculation or filter computes from
  nothing; question text shows the respondent a gap in the sentence. The usual cause is
  reaching for `lastmod` when the interview date was meant — use `startdate`. The error
  names which of the four it found.

  `starttime` and `startdate` are deliberately **allowed** here. They hold a value from
  the first question onward, and `separator:[[startdate]]` on an `age_at_date` calculation
  is the intended way to get age at interview.
- **Placeholder in a validation message** *(warning)*: question text is expanded, but the
  message on a logic or unique check is not — the app puts that string straight into the
  error banner. So `'Does not match [[age]]'` shows the interviewer the brackets, not the
  answer. This warns for *every* field name, not just reserved ones, because none of them
  work there. Word the value into the sentence instead.
- **Skip to a reserved variable**: because the generator chooses where those go —
  `starttime` and `startdate` before the first question, the rest after the last — a skip
  aimed at one never lands where the spreadsheet suggests. It jumps either back past the
  start of the questionnaire or forward past every question that remains, clearing the
  answers on the way. This is an error rather than a warning: unlike a declared row, there
  is no arrangement in which it does something useful. Skip to a real question instead.

#### FieldName Errors
- **Starts with number**: Field names must start with a letter
- **Not lowercase**: All field names must be lowercase
- **Contains invalid characters**: Only letters, numbers, and underscores allowed
- **Duplicate field names**: Each field name must be unique within a worksheet

#### QuestionType/FieldType Errors
- **Invalid QuestionType**: Must be one of: radio, checkbox, combobox, text, date, information, automatic, button
- **Invalid FieldType**: Must be one of: text, integer, text_integer, text_decimal, date, datetime, hourmin, n/a
- **Mismatched types**:
  - Radio must use integer fieldtype
  - Checkbox must use text fieldtype
  - Date must use date or datetime fieldtype
  - Text must use text, text_integer, text_decimal or hourmin fieldtype

#### Response Errors
- **Missing colon**: Static responses must be in format `value:text`
- **Space after colon**: No space allowed after the colon
- **Leading spaces**: No spaces before the value
- **Duplicate values**: Response values must be unique
- **Missing responses**: Radio and checkbox questions must have responses defined

#### Logic Check Errors
- **Missing semicolon**: Logic checks must have format `expression; 'message'`
- **Message not in quotes**: Error message must be enclosed in single quotes
- **Invalid operator**: Must use valid operators (=, !=, <>, >, <, >=, <=, and, or)
- **Nonexistent field**: Referenced field must exist in the worksheet
- **Field appears after**: Referenced field must appear before the current question

#### Skip Logic Errors
- **Missing colon**: Skip must have format `preskip: field operator value, target` or `postskip: field operator value, target`
- **Invalid skip type**: Must start with `preskip` or `postskip`
- **Missing comma**: Must have comma separating condition from target
- **Nonexistent field to check**: Field to check must exist and appear before current question
- **Nonexistent target field**: Field to skip to must exist and appear after current question
- **Skip to same question**: Cannot skip to the current question

#### Range Errors
- **Non-numeric range**: LowerRange and UpperRange must be numbers for numeric fields
- **Invalid date format**: Date ranges must use format like `+1y`, `-30d`, `0`
- **Missing date range**: Date questions must have both LowerRange and UpperRange defined
- **Half-set range**: Setting one of LowerRange/UpperRange and leaving the other blank writes `maxvalue='-9'`, which rejects every answer — set both or neither
- **Range on a time**: `hourmin` fields cannot have a range
- **No range (warning)**: A `text_integer` or `text_decimal` field with no range accepts any value that fits MaxCharacters; not warned when the length is fixed (`=10`), which means an identifier

#### MaxCharacters Errors
- **Non-numeric value**: MaxCharacters must be a number
- **Out of range**: MaxCharacters must be between 1 and 2000
- **Missing for text fields**: Required for every question whose QuestionType is `text`
- **Wrong length for a time**: `hourmin` must use `=5`

### Validation Process

1. **Column Header Check**: Validates that all 14 required column headers are present and correct
2. **Field Name Validation**: Checks each field name for syntax and uniqueness
3. **Question/Field Type Validation**: Verifies valid types and correct combinations
4. **Response Validation**: Checks format and syntax of response options
5. **Range Validation**: Validates numeric and date ranges
6. **Logic Check Validation**: Verifies logic check syntax and field references
7. **Skip Logic Validation**: Verifies skip syntax and field references
8. **Cross-Field Validation**: Checks that referenced fields exist and appear in correct order
9. **Duplicate Detection**: Identifies duplicate field names

### Best Practices

1. **Use lowercase field names**: Always use lowercase for field names (e.g., `participant_name`, not `ParticipantName`)
2. **No spaces in static responses**: Format as `1:Yes` not `1: Yes`
3. **Reference fields in order**: Logic checks and skips can only reference fields that appear earlier in the worksheet
4. **Use meaningful field names**: Use descriptive names like `birth_date` instead of `bd`
5. **Test incrementally**: Add questions gradually and test the generation frequently
6. **Check the log file**: Always review the log file for warnings and errors
7. **Use comments column**: Document complex logic and reminders in the Comments column
8. **Use merged rows for section headers**: Merge all 14 columns to create section dividers
9. **Delete empty rows**: Remove any empty rows at the end of your worksheets to avoid duplicate field name errors

### Troubleshooting

**Application won't start:**
- Ensure .NET Framework 4.7.2 or higher is installed
- Check that all dependencies are installed
- Verify Excel is installed

**"Column headers are incorrect" error:**
- Ensure the first row has exactly 14 columns with the correct header names
- Check for extra spaces in header names
- Verify column order is correct

**"Data Dictionary contains errors" message:**
- Open the log file (`gistlogfile.txt`) in the log path
- Review all ERROR messages
- Fix issues in the Excel file
- Run the application again

**XML files not generated:**
- Check that worksheets end in `_dd`
- Verify there are no validation errors in the log file
- Ensure the XML output path exists and is writable

**survey_manifest.gistx not created:**
- Verify you have a `crfs` worksheet
- Check that the `crfs` worksheet has the correct structure
- Review log file for CRFS-related errors

---

## Support

For issues or questions:
1. Check the generated log file for detailed error information
2. Verify Excel data dictionary format matches specifications
3. Ensure all file paths in config.json are correct
4. Review this README and reference documents
5. Check the sample Excel files for examples
