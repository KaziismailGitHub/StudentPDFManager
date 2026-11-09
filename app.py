"""Download student PDFs linked from column D of an Excel workbook."""

from __future__ import annotations

import csv
import json
import os
import re
import tempfile
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Callable
from urllib.parse import parse_qs, urlparse


APP_ROOT = Path(__file__).resolve().parent
APP_NAME = "Student PDF Manager"
TARGET_BATCHES = ("50A", "50B", "50C")
DEFAULT_WORKBOOK = APP_ROOT / "sample_data" / "students.xlsx"
DEFAULT_OUTPUT = APP_ROOT / "output"
LOG_PATH = APP_ROOT / "logs" / "app.log"
SCOPES = ("https://www.googleapis.com/auth/drive.readonly",)
CONFIG_DIR = APP_ROOT / "config"
CONFIG_PATH = CONFIG_DIR / "settings.json"
TOKEN_PATH = CONFIG_DIR / "token.json"
MANIFEST_NAME = ".student_pdf_downloads.json"
REPORT_FIELDS = ("Batch", "Source Row", "Student ID", "Student Name", "Column D value", "Drive Link", "File Name", "Status", "Details")

EventCallback = Callable[[dict[str, Any]], None]


def user_message(exc: Exception) -> str:
    return str(exc).strip().replace("\r", " ").replace("\n", " ") or exc.__class__.__name__


def extract_drive_file_id(value: str) -> str | None:
    """Accept common Google Drive sharing URLs; reject ordinary text and other hosts."""
    value = value.strip()
    if not value:
        return None
    candidate = value if "://" in value else "https://" + value
    parsed = urlparse(candidate)
    host = (parsed.hostname or "").lower()
    if host not in {"drive.google.com", "docs.google.com"} and not host.endswith(".drive.google.com"):
        return None
    parts = [part for part in parsed.path.split("/") if part]
    for marker in ("file", "document", "spreadsheets", "presentation", "drawings"):
        if marker in parts:
            index = parts.index(marker)
            if index + 2 < len(parts) and parts[index + 1] == "d":
                file_id = parts[index + 2]
                if re.fullmatch(r"[A-Za-z0-9_-]+", file_id):
                    return file_id
    query_id = parse_qs(parsed.query).get("id", [None])[0]
    if query_id and re.fullmatch(r"[A-Za-z0-9_-]+", query_id):
        return query_id
    return None


def excel_value_text(cell: Any) -> str:
    value = cell.value
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        number_format = str(getattr(cell, "number_format", ""))
        if re.fullmatch(r"0+", number_format):
            return str(int(value)).zfill(len(number_format))
        return str(int(value))
    return str(value).strip()


def excel_cell_link(cell: Any) -> str:
    hyperlink = getattr(cell, "hyperlink", None)
    target = getattr(hyperlink, "target", None) or ""
    if target:
        return str(target).strip()
    value = cell.value
    if isinstance(value, str):
        match = re.match(r'^\s*=\s*HYPERLINK\s*\(\s*"((?:[^"]|"")*)"', value, re.IGNORECASE)
        if match:
            return match.group(1).replace('""', '"').strip()
        if extract_drive_file_id(value):
            return value.strip()
    return ""


def safe_component(value: str, fallback: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value.strip()).rstrip(" .")
    if value in {"", ".", ".."}:
        value = fallback
    reserved = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
    if value.split(".")[0].upper() in reserved:
        value = "_" + value
    return value[:120]


def load_json(path: Path, fallback: Any) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return fallback


def save_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + f".{os.getpid()}.tmp")
    temp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, path)


def get_workbook_sheets(workbook_path: Path) -> list[str]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError("The Excel reader is missing. Run setup.bat, then start the app again.") from exc
    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    try:
        available = {name.casefold() for name in workbook.sheetnames}
        return [batch for batch in TARGET_BATCHES if batch.casefold() in available]
    finally:
        workbook.close()


def read_workbook_rows(workbook_path: Path, selected_sheets: tuple[str, ...] | None = None) -> tuple[list[str], dict[str, list[dict[str, Any]]]]:
    """Read only A, B and D from the three requested batch worksheets."""
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError("The Excel reader is missing. Run setup.bat, then start the app again.") from exc

    workbook = load_workbook(workbook_path, read_only=False, data_only=False)
    actual_names = {name.casefold(): name for name in workbook.sheetnames}
    requested = list(TARGET_BATCHES) if selected_sheets is None else [
        next((batch for batch in TARGET_BATCHES if batch.casefold() == selected.casefold()), selected)
        for selected in selected_sheets
    ]
    rows_by_sheet: dict[str, list[dict[str, Any]]] = {}
    try:
        for batch in requested:
            actual_title = actual_names.get(batch.casefold())
            if actual_title is None:
                continue
            worksheet = workbook[actual_title]
            rows: list[dict[str, Any]] = []
            for row_number in range(2, worksheet.max_row + 1):
                id_cell = worksheet.cell(row_number, 1)
                name_cell = worksheet.cell(row_number, 2)
                link_cell = worksheet.cell(row_number, 4)
                student_id = excel_value_text(id_cell)
                student_name = excel_value_text(name_cell)
                display = excel_value_text(link_cell)
                link = excel_cell_link(link_cell)
                all_d_text = " ".join((display, link)).lower()
                access_denied = "access denied" in all_d_text or "permission denied" in all_d_text
                if student_id or student_name or display or link:
                    rows.append({
                        "row": row_number,
                        "student_id": student_id,
                        "student_name": student_name,
                        "display": display,
                        "link": link,
                        "access_denied": access_denied,
                    })
            rows_by_sheet[batch] = rows
    finally:
        workbook.close()
    present_batches = [batch for batch in TARGET_BATCHES if batch.casefold() in actual_names]
    return present_batches, rows_by_sheet


def get_drive_service(client_secret: Path, notify: EventCallback):
    try:
        from google.auth.transport.requests import Request
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError("Google API packages are missing. Run setup.bat, then start the app again.") from exc

    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    credentials = None
    if TOKEN_PATH.exists():
        try:
            from google.oauth2.credentials import Credentials

            credentials = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
            if not set(SCOPES).issubset(set(credentials.scopes or ())):
                credentials = None
        except Exception:
            credentials = None

    if credentials and credentials.expired and credentials.refresh_token:
        try:
            credentials.refresh(Request())
        except Exception:
            credentials = None

    if not credentials or not credentials.valid:
        if not client_secret.is_file():
            raise FileNotFoundError("Choose the OAuth desktop client JSON file from Google Cloud Console.")
        notify({"type": "status", "text": "Opening Google Drive sign-in in your browser…"})
        flow = InstalledAppFlow.from_client_secrets_file(str(client_secret), SCOPES)
        credentials = flow.run_local_server(port=0, open_browser=True)

    TOKEN_PATH.write_text(credentials.to_json(), encoding="utf-8")
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def file_error_is_access_denied(exc: Exception) -> bool:
    status = getattr(getattr(exc, "resp", None), "status", None)
    message = user_message(exc).lower()
    return status == 403 or any(text in message for text in ("access denied", "permission denied", "insufficient permissions", "not authorized", "forbidden"))


def record_report(report: list[dict[str, str]], batch: str, row: int | str, student_id: str, student_name: str, link_value: str, drive_link: str, file_name: str, status: str, details: str) -> None:
    report.append({
        "Batch": batch,
        "Source Row": str(row),
        "Student ID": student_id,
        "Student Name": student_name,
        "Column D value": link_value,
        "Drive Link": drive_link,
        "File Name": file_name,
        "Status": status,
        "Details": details,
    })


def save_download_report(output_root: Path, report: list[dict[str, str]]) -> Path:
    output_root.mkdir(parents=True, exist_ok=True)
    report_path = output_root / f"download_report_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.csv"
    with report_path.open("w", encoding="utf-8-sig", newline="") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=REPORT_FIELDS)
        writer.writeheader()
        writer.writerows(report)
    return report_path


def unique_target(folder: Path, file_name: str) -> Path:
    stem = Path(file_name).stem
    suffix = Path(file_name).suffix or ".pdf"
    candidate = folder / (stem + suffix)
    index = 2
    while candidate.exists():
        candidate = folder / f"{stem} ({index}){suffix}"
        index += 1
    return candidate


def run_download(workbook_path: Path, client_secret: Path, output_root: Path, selected_sheets: tuple[str, ...] | None, notify: EventCallback) -> dict[str, Any]:
    from googleapiclient.http import MediaIoBaseDownload

    sheet_names, rows_by_sheet = read_workbook_rows(workbook_path, selected_sheets)
    drive_service = get_drive_service(client_secret, notify)
    requested_sheets = list(TARGET_BATCHES) if selected_sheets is None else list(selected_sheets)
    missing_sheets = [name for name in requested_sheets if name not in rows_by_sheet]
    rows = [(sheet, record) for sheet in requested_sheets if sheet in rows_by_sheet for record in rows_by_sheet[sheet]]
    total = len(rows) + len(missing_sheets)
    if total == 0:
        total = 1
    report: list[dict[str, str]] = []
    counts = {"success": 0, "failed": 0, "skipped": 0, "done": 0, "total": total}
    for sheet in missing_sheets:
        record_report(report, sheet, "", "", "", "", "", "", "Failed", "Worksheet was not found in the selected workbook.")
        counts["failed"] += 1
        counts["done"] += 1
    if not rows and not missing_sheets:
        counts["done"] = 1
    notify({"type": "counts", **counts})

    manifest_path = output_root / MANIFEST_NAME
    manifest = load_json(manifest_path, {})
    if not isinstance(manifest, dict):
        manifest = {}
    seen: set[str] = set()

    def advance() -> None:
        counts["done"] += 1
        notify({"type": "counts", **counts})

    for batch, entry in rows:
        row_number = entry["row"]
        student_id = entry["student_id"]
        student_name = entry["student_name"]
        display = entry["display"]
        link = entry["link"]
        notify({"type": "file_progress", "value": 0, "text": ""})

        def finish(status: str, details: str, file_name: str = "") -> None:
            record_report(report, batch, row_number, student_id, student_name, display, link, file_name, status, details)
            if status == "Downloaded":
                counts["success"] += 1
                notify({"type": "log", "text": f"Downloaded: {batch} / {student_id} / {file_name}"})
            elif status == "Skipped":
                counts["skipped"] += 1
                notify({"type": "log", "text": f"Skipped: {batch} / {student_id or '(no ID)'} — {details}"})
            else:
                counts["failed"] += 1
                notify({"type": "log", "text": f"Failed: {batch} / {student_id or '(no ID)'} — {details}"})

        if entry["access_denied"]:
            finish("Skipped", "Access Denied text in Column D.")
            advance()
            continue
        if not student_id:
            finish("Failed", "Student ID is missing in Column A; file was not downloaded.")
            advance()
            continue
        if not link:
            if display:
                finish("Skipped", "Column D contains ordinary text, not a Google Drive link.")
            else:
                finish("Failed", "Missing assignment link in Column D.")
            advance()
            continue
        file_id = extract_drive_file_id(link)
        if not file_id:
            finish("Skipped", "Column D does not contain a supported Google Drive link.")
            advance()
            continue

        identity = json.dumps([batch, student_id, file_id], ensure_ascii=False)
        if identity in seen:
            finish("Skipped", "Duplicate link for this batch and student in this run.")
            advance()
            continue
        seen.add(identity)
        previous = manifest.get(identity)
        if previous and (output_root / previous).is_file():
            finish("Skipped", "This Drive file was downloaded previously.", Path(previous).name)
            advance()
            continue

        student_folder = output_root / safe_component(student_id, "Missing ID")
        temp_path: Path | None = None
        target: Path | None = None
        created_target = False
        try:
            student_folder.mkdir(parents=True, exist_ok=True)
            metadata = drive_service.files().get(
                fileId=file_id,
                fields="id,name,mimeType,trashed,capabilities(canDownload)",
                supportsAllDrives=True,
            ).execute()
            if metadata.get("trashed"):
                raise ValueError("The Drive file is in the trash.")
            if metadata.get("mimeType") != "application/pdf":
                raise ValueError(f"The Drive file is not a PDF (type: {metadata.get('mimeType', 'unknown')}).")
            if (metadata.get("capabilities") or {}).get("canDownload") is False:
                finish("Skipped", "Access Denied: Google Drive does not allow this file to be downloaded.", metadata.get("name", ""))
                advance()
                continue
            file_name = safe_component(metadata.get("name", "assignment.pdf"), "assignment.pdf")
            if not file_name.lower().endswith(".pdf"):
                file_name += ".pdf"
            with tempfile.NamedTemporaryFile(mode="w+b", prefix=".student-pdf-", suffix=".part", dir=student_folder, delete=False) as destination:
                temp_path = Path(destination.name)
                request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
                downloader = MediaIoBaseDownload(destination, request, chunksize=1024 * 1024)
                finished = False
                while not finished:
                    status, finished = downloader.next_chunk()
                    if status:
                        progress = int(status.progress() * 100)
                        notify({"type": "file_progress", "value": progress, "text": f"{progress}% — {metadata.get('name', file_name)}"})
            with temp_path.open("rb") as downloaded:
                if b"%PDF-" not in downloaded.read(1024):
                    raise ValueError("The downloaded content does not have a PDF signature.")
            # Windows os.rename fails if the destination exists, preventing overwrite.
            while True:
                target = unique_target(student_folder, file_name)
                try:
                    os.rename(temp_path, target)
                    created_target = True
                    temp_path = None
                    break
                except FileExistsError:
                    continue
            manifest[identity] = str(target.relative_to(output_root))
            save_json(manifest_path, manifest)
            finish("Downloaded", "PDF saved to the local Windows folder.", target.name)
        except Exception as exc:
            if temp_path and temp_path.exists():
                try:
                    temp_path.unlink()
                except OSError:
                    pass
            if created_target and target and target.exists():
                try:
                    target.unlink()
                except OSError:
                    pass
            message = user_message(exc)
            if file_error_is_access_denied(exc):
                finish("Skipped", f"Access Denied: {message}")
            else:
                finish("Failed", message)
        advance()

    report_path = save_download_report(output_root, report)
    return {**counts, "report": report_path, "report_rows": len(report)}


class DownloaderApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("850x680")
        self.root.minsize(740, 580)
        self.busy = False
        self.config = load_json(CONFIG_PATH, {})
        if not isinstance(self.config, dict):
            self.config = {}
        workbook_default = self.config.get("workbook", str(DEFAULT_WORKBOOK))
        client_default = self.config.get("client_secret", "")
        output_default = str(DEFAULT_OUTPUT)
        self.workbook_var = tk.StringVar(value=workbook_default)
        self.client_var = tk.StringVar(value=client_default)
        self.output_var = tk.StringVar(value=output_default)
        self.batch_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Choose the Excel workbook and Google Drive OAuth client file to begin.")
        self.counts_var = tk.StringVar(value="Downloaded: 0    Failed: 0    Skipped: 0")
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self._append_log("Application opened.")
        if Path(workbook_default).is_file():
            self.refresh_batches(silent=True)

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)
        ttk.Label(main, text=APP_NAME, font=("Segoe UI", 17, "bold")).pack(anchor="w")
        ttk.Label(
            main,
            text="Reads only columns A, B, and D from worksheets 50A, 50B, and 50C. Downloads stay on this PC.",
            wraplength=790,
        ).pack(anchor="w", pady=(2, 14))

        form = ttk.Frame(main)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text="Excel workbook (.xlsx)").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(form, textvariable=self.workbook_var).grid(row=0, column=1, sticky="ew", pady=5)
        self.workbook_browse_button = ttk.Button(form, text="Browse…", command=self.choose_workbook)
        self.workbook_browse_button.grid(row=0, column=2, padx=(8, 0), pady=5)
        self.load_workbook_button = ttk.Button(form, text="Load Workbook", command=self.refresh_batches)
        self.load_workbook_button.grid(row=0, column=3, padx=(8, 0), pady=5)

        ttk.Label(form, text="Google Drive OAuth desktop client JSON").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(form, textvariable=self.client_var).grid(row=1, column=1, columnspan=2, sticky="ew", pady=5)
        self.client_browse_button = ttk.Button(form, text="Browse…", command=self.choose_client)
        self.client_browse_button.grid(row=1, column=3, padx=(8, 0), pady=5)
        ttk.Label(form, text="Download folder (fixed)").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=5)
        ttk.Entry(form, textvariable=self.output_var, state="readonly").grid(row=2, column=1, columnspan=3, sticky="ew", pady=5)

        buttons = ttk.Frame(main)
        buttons.pack(fill="x", pady=(14, 8))
        self.connect_button = ttk.Button(buttons, text="Connect Google Drive", command=self.connect)
        self.connect_button.pack(side="left")
        self.download_all_button = ttk.Button(buttons, text="Download All", command=self.download_all)
        self.download_all_button.pack(side="left", padx=(8, 0))
        self.batch_combo = ttk.Combobox(buttons, textvariable=self.batch_var, values=(), width=16, state="readonly")
        self.batch_combo.pack(side="left", padx=(14, 6))
        self.select_button = ttk.Button(buttons, text="Select Batch", command=self.download_selected)
        self.select_button.pack(side="left")
        self.open_folder_button = ttk.Button(buttons, text="Open Folder", command=self.open_folder)
        self.open_folder_button.pack(side="right")

        ttk.Label(main, textvariable=self.status_var, wraplength=790).pack(anchor="w", pady=(8, 4))
        ttk.Label(main, text="Overall progress").pack(anchor="w")
        self.overall_progress = ttk.Progressbar(main, mode="determinate", maximum=1)
        self.overall_progress.pack(fill="x", pady=(3, 8))
        ttk.Label(main, text="Current PDF").pack(anchor="w")
        self.file_progress = ttk.Progressbar(main, mode="determinate", maximum=100)
        self.file_progress.pack(fill="x", pady=(3, 8))
        ttk.Label(main, textvariable=self.counts_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(2, 8))

        log_frame = ttk.LabelFrame(main, text="Activity", padding=6)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, wrap="word", height=12, state="disabled", font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def choose_workbook(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose the student submissions workbook",
            filetypes=(("Excel workbook", "*.xlsx"), ("All files", "*.*")),
        )
        if path:
            self.workbook_var.set(path)
            self.refresh_batches(silent=True)

    def choose_client(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose OAuth desktop client JSON",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
        )
        if path:
            self.client_var.set(path)

    def refresh_batches(self, silent: bool = False) -> None:
        workbook_path = Path(self.workbook_var.get().strip()).expanduser()
        if not workbook_path.is_file():
            if not silent:
                messagebox.showerror(APP_NAME, "Choose an existing .xlsx workbook.", parent=self.root)
            return
        try:
            sheet_names = get_workbook_sheets(workbook_path)
        except Exception as exc:
            if not silent:
                messagebox.showerror(APP_NAME, user_message(exc), parent=self.root)
            return
        self.batch_combo.configure(values=sheet_names)
        if sheet_names and self.batch_var.get() not in sheet_names:
            self.batch_var.set(sheet_names[0])
        missing = [batch for batch in TARGET_BATCHES if batch not in sheet_names]
        self.status_var.set(f"Loaded {len(sheet_names)} of 3 requested worksheets: {', '.join(sheet_names)}")
        self._append_log(f"Workbook loaded. Worksheets: {', '.join(sheet_names)}")
        if missing:
            self._append_log(f"Required worksheets not found: {', '.join(missing)}")

    def save_settings(self) -> None:
        save_json(CONFIG_PATH, {
            "workbook": self.workbook_var.get().strip(),
            "client_secret": self.client_var.get().strip(),
            "output_folder": str(DEFAULT_OUTPUT),
        })

    def validate_inputs(self, require_workbook: bool = True) -> tuple[Path | None, Path, Path] | None:
        workbook_path: Path | None = None
        if require_workbook:
            workbook_path = Path(self.workbook_var.get().strip()).expanduser()
            if not workbook_path.is_file() or workbook_path.suffix.lower() != ".xlsx":
                messagebox.showerror(APP_NAME, "Choose an existing .xlsx workbook.", parent=self.root)
                return None
        client_secret = Path(self.client_var.get().strip())
        if not client_secret.is_file():
            messagebox.showerror(APP_NAME, "Choose your Google Drive OAuth desktop client JSON file.", parent=self.root)
            return None
        output_root = DEFAULT_OUTPUT
        return workbook_path, client_secret, output_root

    def connect(self) -> None:
        values = self.validate_inputs(require_workbook=False)
        if not values:
            return
        _, client_secret, _ = values
        self._start_worker("connect", None, client_secret, None)

    def download_all(self) -> None:
        self._download(None)

    def download_selected(self) -> None:
        selected = self.batch_var.get().strip()
        if not selected:
            messagebox.showerror(APP_NAME, "Load the workbook and select a worksheet first.", parent=self.root)
            return
        self._download((selected,))

    def _download(self, selected_sheets: tuple[str, ...] | None) -> None:
        values = self.validate_inputs(require_workbook=True)
        if not values:
            return
        workbook_path, client_secret, output_root = values
        assert workbook_path is not None
        self._start_worker("download", workbook_path, client_secret, output_root, selected_sheets)

    def _start_worker(
        self,
        mode: str,
        workbook_path: Path | None,
        client_secret: Path,
        output_root: Path | None,
        selected_sheets: tuple[str, ...] | None = None,
    ) -> None:
        if self.busy:
            return
        self.save_settings()
        self.busy = True
        self._set_controls(False)
        self.overall_progress.configure(value=0, maximum=1)
        self.file_progress.configure(value=0)
        self.status_var.set("Connecting to Google Drive…" if mode == "connect" else "Preparing local download…")
        self._append_log(f"Started {mode} operation.")

        def worker() -> None:
            def notify(event: dict[str, Any]) -> None:
                self.root.after(0, self._handle_event, event)

            try:
                if mode == "connect":
                    drive_service = get_drive_service(client_secret, notify)
                    drive_service.about().get(fields="user(displayName,emailAddress)").execute()
                    notify({"type": "done", "result": {"mode": "connect"}})
                    return
                assert workbook_path is not None and output_root is not None
                output_root.mkdir(parents=True, exist_ok=True)
                result = run_download(workbook_path, client_secret, output_root, selected_sheets, notify)
                result["mode"] = "download"
                notify({"type": "done", "result": result})
            except Exception as exc:
                notify({"type": "fatal", "message": user_message(exc)})

        threading.Thread(target=worker, name="StudentPDFDownloader", daemon=True).start()

    def _handle_event(self, event: dict[str, Any]) -> None:
        kind = event.get("type")
        if kind == "status":
            self.status_var.set(event.get("text", ""))
        elif kind == "log":
            self._append_log(event.get("text", ""))
        elif kind == "counts":
            done = int(event.get("done", 0))
            total = max(int(event.get("total", 1)), 1)
            self.overall_progress.configure(maximum=total, value=min(done, total))
            self.counts_var.set(f"Downloaded: {event.get('success', 0)}    Failed: {event.get('failed', 0)}    Skipped: {event.get('skipped', 0)}")
            self.status_var.set(f"Processing row {min(done, total)} of {total}…")
        elif kind == "file_progress":
            self.file_progress.configure(value=event.get("value", 0))
            if event.get("text"):
                self.status_var.set(event["text"])
        elif kind == "done":
            self.busy = False
            self._set_controls(True)
            result = event.get("result", {})
            if result.get("mode") == "connect":
                self.status_var.set("Connected to Google Drive. Download permissions are read-only.")
                messagebox.showinfo(APP_NAME, "Google Drive sign-in worked.", parent=self.root)
            else:
                report = result.get("report")
                self.status_var.set(f"Finished. Download report: {report}")
                self._append_log(f"Finished. Download report saved to {report}")
                messagebox.showinfo(
                    APP_NAME,
                    f"Download finished.\n\nDownloaded: {result.get('success', 0)}\nFailed: {result.get('failed', 0)}\nSkipped: {result.get('skipped', 0)}\n\nReport: {report}",
                    parent=self.root,
                )
        elif kind == "fatal":
            self.busy = False
            self._set_controls(True)
            self.status_var.set("Operation stopped because of an error.")
            self._append_log("ERROR: " + event.get("message", "Unknown error"))
            messagebox.showerror(APP_NAME, event.get("message", "Unknown error"), parent=self.root)

    def _append_log(self, text: str) -> None:
        timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
        try:
            LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
            with LOG_PATH.open("a", encoding="utf-8") as log_file:
                log_file.write(f"{timestamp} {text}\n")
        except OSError:
            pass
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_controls(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for control in (
            self.workbook_browse_button,
            self.load_workbook_button,
            self.client_browse_button,
            self.connect_button,
            self.download_all_button,
            self.select_button,
            self.open_folder_button,
        ):
            control.configure(state=state)
        self.batch_combo.configure(state="readonly" if enabled else "disabled")

    def open_folder(self) -> None:
        folder = DEFAULT_OUTPUT
        try:
            folder.mkdir(parents=True, exist_ok=True)
            os.startfile(str(folder))
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Could not open the folder:\n{user_message(exc)}", parent=self.root)

    def close(self) -> None:
        if self.busy:
            messagebox.showinfo(APP_NAME, "Wait for the current Google Drive operation to finish before closing the app.", parent=self.root)
            return
        self.save_settings()
        self.root.destroy()


def main() -> None:
    root = tk.Tk()
    DownloaderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()


