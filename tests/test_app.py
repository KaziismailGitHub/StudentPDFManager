import csv
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

import app


class FakeGoogleError(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(message)
        self.resp = SimpleNamespace(status=status)


class FakeRequest:
    def __init__(self, result):
        self.result = result

    def execute(self):
        if isinstance(self.result, Exception):
            raise self.result
        return self.result


class FakeFiles:
    def __init__(self):
        self.metadata = {
            "valid-a": {"id": "valid-a", "name": "answer.pdf", "mimeType": "application/pdf", "capabilities": {"canDownload": True}},
            "valid-b": {"id": "valid-b", "name": "second.pdf", "mimeType": "application/pdf", "capabilities": {"canDownload": True}},
            "valid-c": {"id": "valid-c", "name": "third.pdf", "mimeType": "application/pdf", "capabilities": {"canDownload": True}},
            "doc": {"id": "doc", "name": "document.docx", "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
        }
        self.get_calls = []

    def get(self, fileId, **kwargs):
        self.get_calls.append(fileId)
        if fileId == "denied-api":
            return FakeRequest(FakeGoogleError(403, "Access denied for this file"))
        if fileId not in self.metadata:
            return FakeRequest(FakeGoogleError(404, "File not found"))
        return FakeRequest(self.metadata[fileId])

    def get_media(self, fileId, **kwargs):
        return FakeRequest({"file_id": fileId})


class FakeDrive:
    def __init__(self):
        self.file_api = FakeFiles()

    def files(self):
        return self.file_api


class FakeDownloader:
    def __init__(self, destination, request, chunksize=0):
        self.destination = destination
        self.file_id = request.result["file_id"]

    def next_chunk(self):
        self.destination.write(b"%PDF-1.7\nMock PDF content\n")
        self.destination.flush()
        return (None, True)


class DownloaderAppTests(unittest.TestCase):
    def create_workbook(self, path: Path) -> None:
        workbook = Workbook()
        sheet_a = workbook.active
        sheet_a.title = "50A"
        sheet_a.append(["Student ID", "Student Name", "ignored", 1])
        sheet_a.append(["S001", "Alice", "ignored", "Open PDF"])
        sheet_a["D2"].hyperlink = "https://drive.google.com/file/d/valid-a/view"
        sheet_a.append(["S002", "Bob", "ignored", "Submitted manually"])
        sheet_a.append(["S003", "Carol", "ignored", "Access Denied"])
        sheet_a.append(["S001", "Alice", "ignored", "Duplicate"])
        sheet_a["D5"].hyperlink = "https://drive.google.com/file/d/valid-a/view"
        sheet_a.append(["S004", "Dan", "ignored", "Not a PDF"])
        sheet_a["D6"].hyperlink = "https://drive.google.com/file/d/doc/view"
        sheet_a.append(["S005", "Eve", "ignored", None])
        sheet_a.append(["S006", "Farah", "ignored", "Access link"])
        sheet_a["D8"].hyperlink = "https://drive.google.com/file/d/denied-api/view"

        sheet_b = workbook.create_sheet("50B")
        sheet_b.append(["Student ID", "Student Name", "ignored", "Assignment 1"])
        sheet_b.append(["S101", "Grace", "ignored", "https://drive.google.com/open?id=valid-b"])

        sheet_c = workbook.create_sheet("50c")
        sheet_c.append(["Student ID", "Student Name", "ignored", "Assignment 1"])
        sheet_c.append(["S201", "Hana", "ignored", "https://drive.google.com/file/d/valid-c/view"])

        other = workbook.create_sheet("Extra")
        other.append(["ID", "Name", "ignored", "Assignment 1"])
        other.append(["SX", "Ignored", "ignored", "https://drive.google.com/file/d/valid-a/view"])
        workbook.save(path)

    def test_only_requested_sheets_and_expected_outcomes(self):
        with tempfile.TemporaryDirectory() as temporary:
            temp = Path(temporary)
            source = temp / "submissions.xlsx"
            output = temp / "output"
            self.create_workbook(source)
            self.assertEqual(app.get_workbook_sheets(source), ["50A", "50B", "50C"])

            drive = FakeDrive()
            with patch.object(app, "get_drive_service", return_value=drive), patch(
                "googleapiclient.http.MediaIoBaseDownload", FakeDownloader
            ):
                result = app.run_download(source, Path("client.json"), output, None, lambda event: None)

            self.assertEqual((result["success"], result["failed"], result["skipped"]), (3, 2, 4))
            self.assertEqual(result["done"], result["total"])
            self.assertEqual(result["report_rows"], 9)
            self.assertTrue((output / "S001" / "answer.pdf").is_file())
            self.assertTrue((output / "S101" / "second.pdf").is_file())
            self.assertTrue((output / "S201" / "third.pdf").is_file())
            self.assertEqual(drive.file_api.get_calls.count("valid-a"), 1)

            with result["report"].open(encoding="utf-8-sig", newline="") as report_file:
                records = list(csv.DictReader(report_file))
            self.assertEqual(len(records), 9)
            statuses = [record["Status"] for record in records]
            self.assertEqual(statuses.count("Downloaded"), 3)
            self.assertEqual(statuses.count("Failed"), 2)
            self.assertEqual(statuses.count("Skipped"), 4)
            self.assertTrue(any("Access Denied" in record["Details"] for record in records))
            self.assertTrue(all(record["Batch"] != "Extra" for record in records))

    def test_existing_filename_is_never_overwritten_and_manifest_skips_repeat(self):
        with tempfile.TemporaryDirectory() as temporary:
            temp = Path(temporary)
            source = temp / "submissions.xlsx"
            output = temp / "output"
            self.create_workbook(source)
            existing = output / "S001" / "answer.pdf"
            existing.parent.mkdir(parents=True)
            existing.write_bytes(b"keep this file")
            drive = FakeDrive()

            with patch.object(app, "get_drive_service", return_value=drive), patch(
                "googleapiclient.http.MediaIoBaseDownload", FakeDownloader
            ):
                first = app.run_download(source, Path("client.json"), output, ("50A",), lambda event: None)
                second = app.run_download(source, Path("client.json"), output, ("50A",), lambda event: None)

            self.assertEqual(existing.read_bytes(), b"keep this file")
            self.assertTrue((output / "S001" / "answer (2).pdf").is_file())
            self.assertGreater(first["success"], 0)
            self.assertEqual(second["success"], 0)
            self.assertIn("valid-a", drive.file_api.get_calls)

    def test_hyperlink_formula_and_drive_link_parsing(self):
        self.assertEqual(app.extract_drive_file_id("https://drive.google.com/file/d/abc_123/view"), "abc_123")
        self.assertEqual(app.extract_drive_file_id("https://drive.google.com/open?id=abc_123"), "abc_123")
        self.assertIsNone(app.extract_drive_file_id("ordinary text"))

        class FormulaCell:
            value = '=HYPERLINK("https://drive.google.com/file/d/formula123/view","PDF")'
            hyperlink = None

        self.assertEqual(app.extract_drive_file_id(app.excel_cell_link(FormulaCell())), "formula123")


if __name__ == "__main__":
    unittest.main()

