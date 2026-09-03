"""What reaches disk when a build fails.

The gate used to be evaluated once, before the per-file XML validation inside
the loop could set `errorsEncountered` -- so a validation failure skipped only
the zip. Every .xml and survey_manifest.gistx stayed in outputPath while the
console printed "The XML files and manifest HAVE NOT been created", and
hand-zipping those files ships a package containing malformed XML.
"""

import io
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from crf_reader import CRFS_COLUMN_NAMES
from models import AppConfig
from processor import SurveyGenProcessor

from tests.test_dd_validations import HEADERS, numeric_row


def workbook_file(directory, crfs_headers=None, crfs_rows=None):
    """A minimal but complete dictionary: one `_dd` sheet plus a `crfs` sheet."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "enrollee_dd"
    worksheet.append(HEADERS)
    worksheet.append(numeric_row("age"))

    crfs = workbook.create_sheet("crfs")
    crfs.append(list(CRFS_COLUMN_NAMES) if crfs_headers is None else crfs_headers)
    for r in crfs_rows if crfs_rows is not None else [
        [10, "enrollee", "Enrollee", "subjid", "", 1, "", "", "", "", "", "", "", "", ""]
    ]:
        crfs.append(r)

    path = Path(directory) / "dictionary.xlsx"
    workbook.save(path)
    workbook.close()
    return path


def run_quietly(processor):
    """`run` prints the Windows app's SUCCESS/ERRORS message box equivalent."""
    with redirect_stdout(io.StringIO()):
        return processor.run()


def config_for(excel_path, output_path):
    return AppConfig(
        excelFile=str(excel_path),
        csvFiles="",
        outputPath=str(output_path),
        surveyName="Test",
        surveyId="test_survey",
        databaseName="test.sqlite",
    )


class SuccessfulBuildTests(unittest.TestCase):
    def test_a_clean_dictionary_produces_only_a_zip_and_a_log(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            processor = SurveyGenProcessor(config_for(workbook_file(tmp), out))
            self.assertEqual(0, run_quietly(processor), "\n".join(processor.logstring))

            names = sorted(p.name for p in out.iterdir())
            self.assertEqual(["gistlogfile.txt", "test_survey.zip"], names)

    def test_the_zip_holds_the_xml_and_the_manifest(self):
        from zipfile import ZipFile

        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            run_quietly(SurveyGenProcessor(config_for(workbook_file(tmp), out)))
            with ZipFile(out / "test_survey.zip") as archive:
                self.assertEqual(
                    ["enrollee.xml", "survey_manifest.gistx"],
                    sorted(archive.namelist()),
                )

    def test_the_manifest_names_the_file_that_is_actually_in_the_zip(self):
        import json
        from zipfile import ZipFile

        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            run_quietly(SurveyGenProcessor(config_for(workbook_file(tmp), out)))
            with ZipFile(out / "test_survey.zip") as archive:
                manifest = json.loads(archive.read("survey_manifest.gistx"))
                members = set(archive.namelist())
            for xml_name in manifest["xmlFiles"]:
                self.assertIn(xml_name, members)

    def test_the_log_banner_comes_after_the_zip_lines(self):
        """`_create_zip_file` appends its own lines, so it must run first."""
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            processor = SurveyGenProcessor(config_for(workbook_file(tmp), out))
            run_quietly(processor)
            log = processor.logstring
            self.assertLess(
                max(i for i, line in enumerate(log) if "Added to zip" in line),
                next(i for i, line in enumerate(log) if line == "End of log file"),
            )


class FailedValidationTests(unittest.TestCase):
    def test_invalid_xml_leaves_no_manifest_and_no_loose_files(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            processor = SurveyGenProcessor(config_for(workbook_file(tmp), out))
            # Stand in for XML that ET rejects. Escaping now makes malformed
            # output very hard to produce on purpose, which is the point -- but
            # the gate still has to hold if anything ever slips through.
            processor._validate_xml_syntax = lambda path: False

            self.assertEqual(1, run_quietly(processor))

            names = sorted(p.name for p in out.iterdir())
            self.assertEqual(["gistlogfile.txt"], names)
            self.assertFalse((out / "survey_manifest.gistx").exists())
            self.assertFalse((out / "test_survey.zip").exists())

    def test_the_log_says_the_output_was_discarded(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            processor = SurveyGenProcessor(config_for(workbook_file(tmp), out))
            processor._validate_xml_syntax = lambda path: False
            run_quietly(processor)
            self.assertTrue(
                any("Discarded incomplete output" in line for line in processor.logstring),
                "\n".join(processor.logstring),
            )


class CrfsErrorTests(unittest.TestCase):
    def test_a_bad_crfs_header_stops_the_build(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            headers = list(CRFS_COLUMN_NAMES)
            headers.insert(4, "notes")
            excel = workbook_file(tmp, crfs_headers=headers)

            processor = SurveyGenProcessor(config_for(excel, out))
            self.assertEqual(1, run_quietly(processor))
            self.assertEqual(["gistlogfile.txt"], [p.name for p in out.iterdir()])
            self.assertTrue(
                any("header names" in line for line in processor.logstring),
                "\n".join(processor.logstring),
            )

    def test_a_bad_idconfig_stops_the_build(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out"
            rows = [[
                10, "enrollee", "Enrollee", "subjid", "{bad json", 1,
                "", "", "", "", "", "", "", "", "",
            ]]
            excel = workbook_file(tmp, crfs_rows=rows)

            processor = SurveyGenProcessor(config_for(excel, out))
            self.assertEqual(1, run_quietly(processor))
            self.assertFalse((out / "test_survey.zip").exists())


if __name__ == "__main__":
    unittest.main()
