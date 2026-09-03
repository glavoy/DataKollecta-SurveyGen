from __future__ import annotations

import json
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook

from crf_reader import CrfReader
from excel_reader import ExcelReader
from json_generator import JsonGenerator
from models import AppConfig, Question, SurveyManifest
from xml_generator import XmlGenerator


class SurveyGenProcessor:
    def __init__(self, config: AppConfig):
        self.config = config
        self.errorsEncountered = False
        self.logstring: list[str] = []
        self.generated_files: list[Path] = []
        self.question_list_cache: dict[str, list[Question]] = {}

    def run(self) -> int:
        self.logstring = [f"Log file for: {self.config.excelFile}"]

        excel_path = Path(self.config.excelFile)
        output_path = Path(self.config.outputPath)
        output_path.mkdir(parents=True, exist_ok=True)

        if not excel_path.exists():
            self.logstring.append("ERROR: Excel file not found!")
            self.logstring.append(str(excel_path))
            self.logstring.extend(
                [
                    "\r--------------------------------------------------------------------------------",
                    "End of log file",
                    "--------------------------------------------------------------------------------",
                ]
            )
            self._write_logfile()
            print(f"ERROR: Excel file not found: {excel_path}")
            print("Please check the 'excelFile' path in config.json.")
            return 1

        workbook = load_workbook(filename=excel_path, data_only=False)

        try:
            worksheets = [ws for ws in workbook.worksheets if ws.title.endswith("_dd") or ws.title.endswith("_xml")]

            # Read the crfs sheet first: it tells us which automatic fields the
            # manifest fills in (linking field, increment field, primary key,
            # idconfig parts), so they are not reported as missing a calculation.
            crfs_ws = workbook["crfs"] if "crfs" in workbook.sheetnames else None
            if crfs_ws is not None:
                crfs, crf_errors = CrfReader.read_crfs_worksheet(crfs_ws)
                if crf_errors:
                    self.logstring.append("\rChecking worksheet: 'crfs'")
                    self.logstring.extend(crf_errors)
                    self.errorsEncountered = True
            else:
                crfs = []
            supplied_by_table = self._supplied_auto_fields(crfs)

            for ws in worksheets:
                table = ws.title.replace("_dd", "").replace("_xml", "")
                reader = ExcelReader(supplied_auto_fields=supplied_by_table.get(table, set()))
                qlist = reader.create_question_list(ws)
                if reader.errorsEncountered:
                    self.errorsEncountered = True
                self.logstring.extend(reader.logstring)
                self.question_list_cache[ws.title] = qlist

            xml_files: list[str] = []

            if not self.errorsEncountered:
                for ws_name, qlist in self.question_list_cache.items():
                    xml_generator = XmlGenerator()
                    xml_path = xml_generator.write_xml(ws_name, qlist, output_path)
                    self.logstring.extend(xml_generator.logstring)
                    self.generated_files.append(xml_path)

                    # Taken from the file that was actually written rather than
                    # re-derived from the worksheet name. The two derivations
                    # used to disagree (this one did
                    # `ws_name.replace("_dd", ".xml").replace("_xml", ".xml")`,
                    # `write_xml` sliced the suffix), so a sheet named
                    # `hh_xml_dd` put `hh.xml.xml` in the manifest while the
                    # zip held `hh_xml.xml` -- a manifest naming a file that
                    # was not in the package.
                    xml_files.append(xml_path.name)

                    if not self._validate_xml_syntax(xml_path):
                        self.errorsEncountered = True

            # Deliberately a second, separate check rather than part of the
            # block above. The XML loop can set `errorsEncountered` itself, and
            # when it did, the enclosing `if` had already been entered -- so
            # the manifest was written anyway and only the zip was skipped.
            # That left every .xml and the .gistx sitting in outputPath while
            # the console said "HAVE NOT been created", and hand-zipping those
            # files ships a package containing malformed XML.
            if not self.errorsEncountered:
                manifest = SurveyManifest(
                    surveyName=self.config.surveyName,
                    surveyId=self.config.surveyId,
                    databaseName=self.config.databaseName,
                    xmlFiles=xml_files,
                    crfs=crfs,
                )
                manifest_path = output_path / "survey_manifest.gistx"
                JsonGenerator.write_manifest(manifest_path, manifest)
                self.logstring.append("")
                self.logstring.append("Successfully generated survey_manifest.gistx")
                self.generated_files.append(manifest_path)

                self._create_zip_file()
            else:
                self._discard_generated_files()

            # After the zip, so `_create_zip_file`'s own "Added to zip" and
            # "Deleted temporary file" lines land before the banner instead of
            # after the end of the file.
            self.logstring.extend(
                [
                    "\r--------------------------------------------------------------------------------",
                    "End of log file",
                    "--------------------------------------------------------------------------------",
                ]
            )

            self._write_logfile()

            # Console equivalent of the Windows app's SUCCESS/ERRORS FOUND message boxes
            logfile_path = output_path / "gistlogfile.txt"
            if self.errorsEncountered:
                print("ERRORS FOUND: The Data Dictionary contains errors!")
                print("No package was created: the XML files and manifest have been discarded.")
                print(f"Please refer to the log file and rectify all errors: {logfile_path}")
            else:
                print("SUCCESS: Built the XML file(s) and the manifest. No errors were found.")
                print(f"All files have been packaged in: {output_path / (self.config.surveyId + '.zip')}")
                print(f"Log file: {logfile_path}")
            return 1 if self.errorsEncountered else 0
        finally:
            workbook.close()

    @staticmethod
    def _supplied_auto_fields(crfs: list) -> dict[str, set[str]]:
        """Automatic fields the app fills in from the manifest, per table.

        These are populated at runtime from the crfs entry rather than by a
        calculation, so they must not be reported as missing one.
        """
        supplied: dict[str, set[str]] = {}
        for crf in crfs:
            if not crf.tablename:
                continue
            fields: set[str] = set()
            for key in (crf.linkingfield, crf.incrementfield):
                if key:
                    fields.add(key.strip())
            for key in (crf.primarykey or "").split(","):
                if key.strip():
                    fields.add(key.strip())
            if crf.idconfig and crf.idconfig.fields:
                for f in crf.idconfig.fields:
                    if f.name:
                        fields.add(f.name.strip())
            supplied[crf.tablename] = fields
        return supplied

    def _validate_xml_syntax(self, file_path: Path) -> bool:
        try:
            ET.parse(file_path)
            return True
        except ET.ParseError as ex:
            self.logstring.append(f"CRITICAL ERROR: XML Syntax Error in file '{file_path.name}'")
            self.logstring.append(f"Details: {ex}")
            return False
        except Exception as ex:
            self.logstring.append(f"CRITICAL ERROR: Could not validate XML file '{file_path.name}'")
            self.logstring.append(f"Details: {ex}")
            return False

    def _write_logfile(self) -> None:
        # Matches the C# StreamWriter output: CRLF after every line, and a
        # final WriteLine("\n") that emits "\n\r\n".
        logfile = Path(self.config.outputPath) / "gistlogfile.txt"
        with logfile.open("w", encoding="utf-8", newline="") as f:
            for line in self.logstring:
                f.write(line + "\r\n")
            f.write("\n\r\n")

    def _create_zip_file(self) -> None:
        zip_file_path = Path(self.config.outputPath) / f"{self.config.surveyId}.zip"
        if zip_file_path.exists():
            zip_file_path.unlink()

        with ZipFile(zip_file_path, "w", compression=ZIP_DEFLATED) as archive:
            for file_path in self.generated_files:
                if file_path.exists():
                    archive.write(file_path, arcname=file_path.name)
                    self.logstring.append(f"Added to zip: {file_path.name}")

            if self.config.csvFiles:
                csv_dir = Path(self.config.csvFiles.rstrip("\\/"))
                if csv_dir.exists() and csv_dir.is_dir():
                    csv_files = sorted(csv_dir.glob("*.csv"))
                    if csv_files:
                        self.logstring.append("")
                        self.logstring.append("Adding CSV files to package:")
                        for csv in csv_files:
                            archive.write(csv, arcname=csv.name)
                            self.logstring.append(f"Added to zip: {csv.name}")
                    else:
                        self.logstring.append(f"WARNING: No CSV files found in {csv_dir}")
                else:
                    self.logstring.append(f"WARNING: CSV files directory not found: {csv_dir}")

        self.logstring.append("")
        self.logstring.append(f"Successfully created zip file: {zip_file_path}")

        for file_path in self.generated_files:
            if file_path.exists():
                file_path.unlink()
                self.logstring.append(f"Deleted temporary file: {file_path.name}")

    def _discard_generated_files(self) -> None:
        """Remove part-built output after a failure.

        Cleanup used to live only inside `_create_zip_file`, which is skipped
        on failure -- so a failed run left loose .xml files (and, before the
        gate above was split in two, survey_manifest.gistx) behind for someone
        to zip by hand.
        """
        for file_path in self.generated_files:
            if file_path.exists():
                file_path.unlink()
                self.logstring.append(f"Discarded incomplete output: {file_path.name}")


def run_from_config_file(config_file: str | Path) -> int:
    cfg_path = Path(config_file)
    data = json.loads(cfg_path.read_text(encoding="utf-8"))
    config = AppConfig(
        excelFile=data.get("excelFile", ""),
        csvFiles=data.get("csvFiles", ""),
        outputPath=data.get("outputPath", ""),
        surveyName=data.get("surveyName", ""),
        surveyId=data.get("surveyId", ""),
        databaseName=data.get("databaseName", ""),
    )
    processor = SurveyGenProcessor(config)
    return processor.run()
