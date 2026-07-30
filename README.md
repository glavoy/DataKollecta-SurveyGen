# GiSTConfigX - Survey Configuration Generator

A tool for generating XML configuration files and survey manifests from Excel-based data dictionaries. This application streamlines the creation of survey forms for data collection instruments by automating the generation of XML configuration files and survey manifest files.

## Table of Contents

- [Overview](#overview)
- [How It Works](#how-it-works)
- [Installation](#installation)
- [Configuration](#configuration)
- [Creating the Excel Data Dictionary](#creating-the-excel-data-dictionary)
  - [Worksheet Naming Convention](#worksheet-naming-convention)
  - [Column Specifications](#column-specifications)
    - [FieldName](#fieldname)
    - [QuestionType](#questiontype)
    - [FieldType](#fieldtype)
    - [QuestionText](#questiontext)
    - [MaxCharacters](#maxcharacters)
    - [Responses](#responses)
    - [LowerRange](#lowerrange)
    - [UpperRange](#upperrange)
    - [LogicCheck](#logiccheck)
    - [DontKnow](#dontknow)
    - [Refuse](#refuse)
    - [NA](#na)
    - [Skip](#skip)
    - [Comments](#comments)
  - [The CRFS Worksheet](#the-crfs-worksheet)
- [Examples](#examples)
- [Output Files](#output-files)
- [Error Handling and Validation](#error-handling-and-validation)

---

## Overview

**GiSTConfigX** reads Excel workbooks containing structured data dictionaries and automatically generates:
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

1. Clone or download this repository
2. Open `GistConfigX.sln` in Visual Studio
3. Restore NuGet packages:
   - Right-click on the project in Solution Explorer
   - Select "Manage NuGet Packages"
   - Install `System.Data.SQLite` if not already installed
4. Build the solution (Build > Build Solution)
5. Run the application

### System Requirements

- Windows OS (tested on Windows 10/11)
- .NET Framework 4.7.2 or higher
- Microsoft Excel (for Excel Interop functionality)

### Dependencies

- **System.Data.SQLite** (v1.0.119.0)
- **Microsoft.Office.Interop.Excel**
- **Newtonsoft.Json**
- **Windows Forms**

---

## Configuration

Create or edit the `config.json` file in the application directory with the following settings:

```json
{
  "excelFile": "C:\\GeoffOffline\\GiSTConfigX\\Excel\\prismcss.xlsx",
  "outputPath": "C:\\temp\\",
  "surveyName": "PRISM CSS 2025-12-01",
  "surveyId": "prism_css_2025_12_01",
  "databaseName": "prism_css_2025_12_01.sqlite"
}
```

### Configuration Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| `excelFile` | Full path to your Excel data dictionary | `"C:\\GeoffOffline\\GiSTConfigX\\Excel\\prismcss.xlsx"` |
| `outputPath` | Directory where the final zip file and log file will be saved | `"C:\\temp\\"` |
| `surveyName` | Human-readable survey name displayed in the app | `"PRISM CSS 2025-12-01"` |
| `surveyId` | Unique survey identifier (lowercase, no spaces) | `"prism_css_2025_12_01"` |
| `databaseName` | SQLite database filename written to `survey_manifest.gistx` | `"prism_css_2025_12_01.sqlite"` |

**Notes:**
- The application creates a zip file containing all XML files, the survey_manifest.gistx file, and any CSV files used for dynamic responses
- The log file (`gistlogfile.txt`) is written to the `outputPath` directory
- The database name is read from `databaseName` in `config.json`

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
| 12 | NA |
| 13 | Skip |
| 14 | Comments |

**Important Notes:**
- The first row of each worksheet must contain these exact column headers
- Rows that are merged will be ignored (useful for section headers or notes)
- Each non-merged row after the header represents one question/field

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

---

### FieldType

The data type that determines how the value is stored and validated.

**Valid Values:**

| FieldType | Description | Storage Type |
|-----------|-------------|--------------|
| `text` | Text string | Text |
| `integer` | Whole numbers | Integer |
| `text_integer` | Text field accepting only integers | Text (validated) |
| `text_decimal` | Text field accepting decimal numbers | Text (validated) |
| `text_id` | Text identifier | Text |
| `phone_num` | Phone number | Text |
| `date` | Date only | Date |
| `datetime` | Date and time | DateTime |
| `hourmin` | Hour:minute format | Text |
| `n/a` | Not applicable | None (for information questions) |

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
- Required for `text`, `text_integer`, and `phone_num` field types
- Must be a number between 1 and 2000
- Leave blank for non-text fields
- Use `=` to force length

**Examples:**
- `80` for a name field
- `10` for a phone number
- `255` for a comments field
- `=3` user must enter 3 characters and 3 characters will always be saved in the database

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

---

### LowerRange

Minimum value for numeric validation or minimum date for date questions.

#### For Numeric Fields

**Requirements:**
- Must be a number (integer or decimal)
- Used with `UpperRange` to create a validation range

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
- `True` - Show the "Don't Know" response
- `False` - Don't show the response
- Leave blank if not needed

**Example:**
When set to `True`, a "Don't Know" response appears that allows the user to skip the question without providing an answer.

---

### Refuse

Adds a "Refuse to Answer" response option to the question.

**Valid Values:**
- `True` - Show the "Refuse to Answer" response
- `False` - Don't show the response
- Leave blank if not needed

**Example:**
When set to `True`, a "Refuse to Answer" response appears for sensitive questions.

---

### NA

Adds a "Not Applicable" response option to the question.

**Valid Values:**
- `True` - Show the "Not Applicable" response
- `False` - Don't show the response
- Leave blank if not needed

**Example:**
When set to `True`, a "N/A" response appears when the question may not apply to all respondents.

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
preskip: if age < 18, skip to comments
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
preskip: if age < 18, skip to commetns
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
GiSTXConfig app can generate a correct `survey_manifest.gistx`, and so the GiSTX survey app
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
                         GiSTX app reads survey_manifest.gistx
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
list (GiSTXConfig handles this automatically when the `_dd` worksheet exists).

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
| `2` | **Force.** Block exit until exactly the stated number of children are entered (an "Exit Anyway" escape may still be shown). |
| `3` | **Auto-sync.** Silently update the parent's count field to match the number actually entered. |

**Typical repeating setup:** `repeat_count_field=nmembers`, `auto_start_repeat=2`,
`repeat_enforce_count=2`.

---

### 10. Caveats & gotchas

- **Blank ≠ zero only where noted.** Most columns treat blank as "off/none". `display_order` blank
  is treated as `0`; `repeat_enforce_count` blank defaults to `1` (warn).
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

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | NA | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| gender | radio | integer | What is your gender? | | 1:Male<br>2:Female | | | | False | False | False | | |

### Example 2: Text Field with Validation

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | NA | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| age | text | text_integer | What is your age? | 3 | | 0 | 120 | | False | False | False | | Valid ages 0-120 |

### Example 3: Date Field with Range

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | NA | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| birth_date | date | date | What is your date of birth? | | | -100y | 0 | | False | False | False | | |

### Example 4: Checkbox Question

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | NA | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| symptoms | checkbox | text | What symptoms do you have? | | A:Fever<br>B:Cough<br>C:Headache<br>D:Fatigue | | | | False | False | False | | Multiple selection |

### Example 5: Dynamic Responses from CSV

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | NA | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| village | combobox | text | Select your village | | source:csv<br>file:villages.csv<br>filter:region = [[region]]<br>display:villagename<br>value:vcode | | | | False | False | False | | Cascading dropdown |

### Example 6: Logic Check

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | NA | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| confirm_age | text | text_integer | Please confirm your age | 3 | | 0 | 120 | confirm_age = age; 'Age does not match!' | False | False | False | | Verification field |

### Example 7: Skip Logic

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | NA | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| pregnant | radio | integer | Are you pregnant? | | 1:Yes<br>2:No | | | | False | False | False | postskip: pregnant = 2, if, next_section | Skip if not pregnant |

### Example 8: Unique Check

| FieldName | QuestionType | FieldType | QuestionText | MaxCharacters | Responses | LowerRange | UpperRange | LogicCheck | DontKnow | Refuse | NA | Skip | Comments |
|-----------|-------------|-----------|--------------|---------------|-----------|------------|------------|------------|----------|--------|----|----|----------|
| participant_id | text | text_id | Enter participant ID | 20 | | | | unique; 'This ID has already been used in the database!' | False | False | False | | Must be unique |

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
  "databaseName": "prism_css_2025_12_01.sqlite",
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
Log file for: C:\GeoffOffline\GiSTConfigX\Excel\prismcss.xlsx

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

#### FieldName Errors
- **Starts with number**: Field names must start with a letter
- **Not lowercase**: All field names must be lowercase
- **Contains invalid characters**: Only letters, numbers, and underscores allowed
- **Duplicate field names**: Each field name must be unique within a worksheet

#### QuestionType/FieldType Errors
- **Invalid QuestionType**: Must be one of: radio, checkbox, combobox, text, date, information, automatic, button
- **Invalid FieldType**: Must be one of: text, integer, text_integer, text_decimal, text_id, phone_num, date, datetime, hourmin, n/a
- **Mismatched types**:
  - Radio must use integer fieldtype
  - Checkbox must use text fieldtype
  - Date must use date or datetime fieldtype

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

#### MaxCharacters Errors
- **Non-numeric value**: MaxCharacters must be a number
- **Out of range**: MaxCharacters must be between 1 and 2000
- **Missing for text fields**: Required for text, text_integer, and phone_num field types

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
