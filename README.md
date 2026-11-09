# Student PDF Manager

A Windows desktop app that reads student IDs and names from columns A and B, reads Google Drive links from column D, and downloads accessible PDFs from worksheets 50A, 50B, and 50C. Drive access is read-only.

## Local storage

- PDFs: `output\<Student ID>\<Drive filename>.pdf`
- CSV reports: `output\download_report_*.csv`
- Activity log: `logs\app.log`
- App settings and OAuth token: `config\settings.json` and `config\token.json`
- Python environment: `.venv\` (created by `setup.bat`)

The app always writes PDFs and CSV reports to the project `output` folder; the destination is fixed in the app. Existing PDFs are not overwritten. A local manifest prevents repeat downloads. The project-local config, logs, downloaded output, and virtual environment are excluded from Git.

## Requirements and setup

1. Install Python 3.15.2 for Windows, including `pip` and Tcl/Tk.
2. Enable Google Drive API in Google Cloud Console and create a Desktop OAuth client JSON file.
3. Add your Google account as a test user if the OAuth consent screen is in Testing mode.
4. Double-click `setup.bat` to create `.venv` and install packages.
5. Double-click `run_app.bat` to launch the app.
6. In the app, choose the source workbook and OAuth client JSON, then connect to Drive.

The source workbook and OAuth client JSON are selected as input files; this project does not copy them automatically. Keep the OAuth client JSON private. The app saves its authorization token under this project’s ignored `config` folder.

## Workbook rules

- Reads only worksheets 50A, 50B, and 50C, matched without regard to case.
- Uses column A for Student ID, column B for Student Name, and column D for the Drive link; row 1 is treated as the header.
- Skips ordinary text, inaccessible files, non-PDF files, and content without a PDF signature.
- Creates timestamped CSV reports for processed rows.

## Files

- `app.py` — application UI, workbook parsing, read-only Drive OAuth, PDF downloads, duplicate tracking, and report generation.
- `requirements.txt` — Python package requirements.
- `setup.bat` — creates the local virtual environment and installs dependencies.
- `run_app.bat` — starts the app.
- `tests\test_app.py` — offline mocked tests.

