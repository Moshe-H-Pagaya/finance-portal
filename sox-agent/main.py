#!/usr/bin/env python3
"""
SOX Controls Testing Agent
──────────────────────────
Reads the SOX Dashboard Launcher Google Sheet, analyzes evidence files
stored in Google Drive via Gemini (Vertex AI), and writes results back.

Column-name detection is dynamic — the script reads row 1 to find each
column by its header label, so reordering columns in the sheet is safe.
The only hard-coded expectation is that row 1 contains the headers.

Expected sheet columns (case-insensitive partial match):
  - Control ID         — unique identifier, used as the Drive subfolder name
  - Control Name       — short display name
  - Control Objective  — what the control should verify (fed to Gemini)
  - Run?               — "Yes" / "No"; only "Yes" rows are processed
  - Result / Pass      — verdict written back: "Passed" or "Failed"
  - Gap / Notes        — gap details written back
  - Last Run           — timestamp written back

Environment variables
  SOX_SPREADSHEET_ID    (required) Google Sheet ID
  SOX_SHEET_TAB         Tab name (default: Sheet1)
  SOX_DRIVE_FOLDER_ID   (required) Google Drive root folder for SOX evidence
  GOOGLE_CLOUD_PROJECT  GCP project (default: finance-ai-497313)
  GOOGLE_CLOUD_LOCATION Vertex AI region (default: us-east1)
  GEMINI_MODEL          Model ID (default: gemini-2.0-flash-001)
  DRY_RUN               Set to "true" to skip writing results back to sheet
"""

import io
import json
import logging
import mimetypes
import os
import sys
import time
from datetime import datetime, timezone
from typing import Any, Optional

import vertexai
from google.auth.transport.requests import Request
from google.oauth2 import service_account
import google.auth
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from vertexai.generative_models import GenerativeModel, Part

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%dT%H:%M:%SZ",
)
log = logging.getLogger("sox-agent")

# ── Config ───────────────────────────────────────────────────────────────────
SPREADSHEET_ID = os.environ.get("SOX_SPREADSHEET_ID", "")
SHEET_TAB = os.environ.get("SOX_SHEET_TAB", "Sheet1")
SOX_DRIVE_FOLDER_ID = os.environ.get("SOX_DRIVE_FOLDER_ID", "")
GCP_PROJECT = os.environ.get("GOOGLE_CLOUD_PROJECT", "finance-ai-497313")
GCP_LOCATION = os.environ.get("GOOGLE_CLOUD_LOCATION", "us-east1")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.0-flash-001")
DRY_RUN = os.environ.get("DRY_RUN", "").lower() == "true"

# Supported evidence file extensions
SUPPORTED_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".csv", ".txt", ".docx", ".doc", ".png", ".jpg", ".jpeg"}

# Max file size for a single evidence file (10 MB)
MAX_FILE_BYTES = 10 * 1024 * 1024

# ── Auth ─────────────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]


def get_credentials():
    """Return Application Default Credentials with the required scopes."""
    creds, _ = google.auth.default(scopes=SCOPES)
    return creds


# ── Google Sheets helpers ─────────────────────────────────────────────────────
def sheets_service(creds):
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def read_sheet(svc, spreadsheet_id: str, tab: str) -> list[list[str]]:
    result = (
        svc.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=tab)
        .execute()
    )
    return result.get("values", [])


def write_cells(svc, spreadsheet_id: str, updates: list[dict]) -> None:
    """
    updates: list of {"range": "Sheet1!V2", "value": "Passed"}
    """
    if DRY_RUN:
        for u in updates:
            log.info("[DRY_RUN] would write %r to %s", u["value"], u["range"])
        return

    data = [
        {"range": u["range"], "values": [[u["value"]]]}
        for u in updates
    ]
    body = {"valueInputOption": "USER_ENTERED", "data": data}
    svc.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id, body=body
    ).execute()


# ── Column detection ──────────────────────────────────────────────────────────
def find_col(headers: list[str], *keywords: str) -> Optional[int]:
    """
    Return the 0-based column index whose header contains ALL keywords
    (case-insensitive). Returns the first match. None if not found.
    """
    for i, h in enumerate(headers):
        h_lower = h.lower()
        if all(kw.lower() in h_lower for kw in keywords):
            return i
    return None


def col_letter(zero_based: int) -> str:
    """Convert a 0-based column index to a spreadsheet letter (A, B, ..., Z, AA, ...)."""
    result = ""
    n = zero_based + 1
    while n:
        n, r = divmod(n - 1, 26)
        result = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[r] + result
    return result


# ── Google Drive helpers ──────────────────────────────────────────────────────
def drive_service(creds):
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def find_subfolder(drv, parent_id: str, name: str) -> Optional[str]:
    """Find a folder by name inside a parent folder. Returns folder ID or None."""
    q = (
        f"'{parent_id}' in parents"
        f" and name = '{name}'"
        f" and mimeType = 'application/vnd.google-apps.folder'"
        f" and trashed = false"
    )
    res = drv.files().list(q=q, fields="files(id, name)", pageSize=10).execute()
    files = res.get("files", [])
    return files[0]["id"] if files else None


def list_files_in_folder(drv, folder_id: str) -> list[dict]:
    """List all non-folder files in a Drive folder (non-recursive)."""
    q = f"'{folder_id}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed = false"
    res = drv.files().list(q=q, fields="files(id, name, mimeType, size)", pageSize=100).execute()
    return res.get("files", [])


def download_file(drv, file_id: str, mime_type: str) -> Optional[bytes]:
    """
    Download a Drive file. Google Workspace files (Docs/Sheets/Slides) are
    exported to a compatible format. Binary files are downloaded directly.
    """
    # Export map for Google Workspace files
    export_map = {
        "application/vnd.google-apps.document": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/vnd.google-apps.spreadsheet": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "application/vnd.google-apps.presentation": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    }

    buf = io.BytesIO()
    try:
        if mime_type in export_map:
            req = drv.files().export(fileId=file_id, mimeType=export_map[mime_type])
        else:
            req = drv.files().get_media(fileId=file_id)
        downloader = MediaIoBaseDownload(buf, req, chunksize=4 * 1024 * 1024)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    except HttpError as e:
        log.error("Drive download failed for %s: %s", file_id, e)
        return None

    data = buf.getvalue()
    if len(data) > MAX_FILE_BYTES:
        log.warning("File %s is %d bytes, truncating to %d", file_id, len(data), MAX_FILE_BYTES)
        data = data[:MAX_FILE_BYTES]
    return data


def get_evidence_files(drv, control_id: str) -> list[dict]:
    """
    Navigate: SOX_DRIVE_FOLDER_ID / <control_id> / input /
    Returns list of {name, mime_type, data} dicts.
    """
    if not SOX_DRIVE_FOLDER_ID:
        log.warning("SOX_DRIVE_FOLDER_ID not set; skipping Drive lookup for %s", control_id)
        return []

    # Find the control subfolder
    ctrl_folder_id = find_subfolder(drv, SOX_DRIVE_FOLDER_ID, control_id)
    if not ctrl_folder_id:
        log.warning("No Drive folder found for control '%s' under %s", control_id, SOX_DRIVE_FOLDER_ID)
        return []

    # Find the input subfolder inside it
    input_folder_id = find_subfolder(drv, ctrl_folder_id, "input")
    if not input_folder_id:
        log.warning("No 'input' subfolder for control '%s' (folder %s)", control_id, ctrl_folder_id)
        return []

    files = list_files_in_folder(drv, input_folder_id)
    if not files:
        log.warning("No files found in input folder for control '%s'", control_id)
        return []

    evidence = []
    for f in files:
        name = f["name"]
        mime = f.get("mimeType", "application/octet-stream")
        size = int(f.get("size", 0)) if f.get("size") else 0

        _, ext = os.path.splitext(name.lower())
        if ext not in SUPPORTED_EXTENSIONS and not mime.startswith("application/vnd.google-apps"):
            log.info("Skipping unsupported file: %s (%s)", name, mime)
            continue

        log.info("Downloading evidence file: %s (%d bytes)", name, size)
        data = download_file(drv, f["id"], mime)
        if data:
            evidence.append({"name": name, "mime_type": mime, "data": data})

    return evidence


# ── Gemini analysis ───────────────────────────────────────────────────────────
def effective_mime(filename: str, drive_mime: str) -> str:
    """Map Drive MIME types and filenames to a MIME Gemini accepts."""
    # Exported Google Docs → docx
    if drive_mime == "application/vnd.google-apps.document":
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if drive_mime == "application/vnd.google-apps.spreadsheet":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    # Use mimetypes for well-known extensions
    guessed, _ = mimetypes.guess_type(filename)
    if guessed:
        return guessed
    if drive_mime and not drive_mime.startswith("application/vnd.google-apps"):
        return drive_mime
    return "application/octet-stream"


ANALYSIS_PROMPT_TEMPLATE = """
You are a SOX (Sarbanes-Oxley) compliance auditor. Your task is to review evidence files
for an internal control and determine whether the control is operating effectively.

CONTROL DETAILS
Control ID:        {control_id}
Control Name:      {control_name}
Control Objective: {control_objective}

INSTRUCTIONS
1. Carefully review all attached evidence files.
2. Determine whether the evidence demonstrates that the control is operating effectively.
3. If the control is NOT passing, identify specific gaps — concrete, actionable findings.
4. Be concise but precise. Avoid vague statements like "documentation could be improved".

RESPONSE FORMAT (respond with valid JSON only, no markdown fences):
{{
  "verdict": "Passed" or "Failed",
  "gaps": ["gap description 1", "gap description 2"],
  "summary": "One or two sentence summary of your finding"
}}

If no evidence files were provided, respond with:
{{
  "verdict": "Failed",
  "gaps": ["No evidence files found in the input folder"],
  "summary": "Unable to test control — no evidence uploaded"
}}
"""


def analyze_control(
    control_id: str,
    control_name: str,
    control_objective: str,
    evidence_files: list[dict],
) -> dict:
    """
    Call Gemini to analyze SOX evidence.
    Returns {"verdict": "Passed"|"Failed", "gaps": [...], "summary": "..."}.
    """
    vertexai.init(project=GCP_PROJECT, location=GCP_LOCATION)
    model = GenerativeModel(GEMINI_MODEL)

    prompt_text = ANALYSIS_PROMPT_TEMPLATE.format(
        control_id=control_id,
        control_name=control_name,
        control_objective=control_objective or "(not specified)",
    )

    parts: list[Any] = [prompt_text]

    for ef in evidence_files:
        mime = effective_mime(ef["name"], ef["mime_type"])
        log.info("Attaching evidence: %s (%s, %d bytes)", ef["name"], mime, len(ef["data"]))
        parts.append(Part.from_data(data=ef["data"], mime_type=mime))

    try:
        response = model.generate_content(parts)
        raw = response.text.strip()
        # Strip markdown code fences if Gemini adds them
        if raw.startswith("```"):
            raw = raw.split("```", 2)[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.rsplit("```", 1)[0].strip()
        return json.loads(raw)
    except json.JSONDecodeError as e:
        log.error("Gemini returned non-JSON: %s", response.text[:500])
        return {"verdict": "Failed", "gaps": [f"Agent error: could not parse Gemini response"], "summary": str(e)}
    except Exception as e:
        log.error("Gemini call failed: %s", e)
        return {"verdict": "Failed", "gaps": [f"Agent error: {e}"], "summary": str(e)}


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if not SPREADSHEET_ID:
        log.error("SOX_SPREADSHEET_ID is required")
        sys.exit(1)

    log.info("Starting SOX agent (project=%s, model=%s, dry_run=%s)", GCP_PROJECT, GEMINI_MODEL, DRY_RUN)

    creds = get_credentials()
    svc = sheets_service(creds)
    drv = drive_service(creds)

    # Read sheet
    all_rows = read_sheet(svc, SPREADSHEET_ID, SHEET_TAB)
    if not all_rows:
        log.warning("Sheet is empty or unreadable")
        return

    headers = [str(h).strip() for h in all_rows[0]]
    log.info("Sheet headers: %s", headers)

    # Detect columns by header content (case-insensitive)
    col_control_id  = find_col(headers, "control", "id")
    col_name        = find_col(headers, "control", "name") or find_col(headers, "name")
    col_objective   = find_col(headers, "objective") or find_col(headers, "description")
    col_run         = find_col(headers, "run")
    col_result      = find_col(headers, "result") or find_col(headers, "pass") or find_col(headers, "verdict")
    col_gaps        = find_col(headers, "gap") or find_col(headers, "notes") or find_col(headers, "finding")
    col_last_run    = find_col(headers, "last run") or find_col(headers, "timestamp") or find_col(headers, "date")

    log.info(
        "Column map: control_id=%s name=%s objective=%s run=%s result=%s gaps=%s last_run=%s",
        col_control_id, col_name, col_objective, col_run, col_result, col_gaps, col_last_run,
    )

    if col_control_id is None:
        log.error(
            "Could not find 'Control ID' column in headers: %s\n"
            "Please ensure the header row contains a column with 'Control' and 'ID' in its name.",
            headers,
        )
        sys.exit(1)

    if col_run is None:
        log.error("Could not find 'Run?' column. Headers: %s", headers)
        sys.exit(1)

    processed = failed = 0

    for row_idx, row in enumerate(all_rows[1:], start=2):  # 1-indexed, skip header
        def cell(col: Optional[int]) -> str:
            if col is None or col >= len(row):
                return ""
            return str(row[col]).strip()

        run_flag = cell(col_run).lower()
        if run_flag not in ("yes", "y", "true", "1"):
            continue

        control_id = cell(col_control_id)
        if not control_id:
            log.warning("Row %d has Run=Yes but no Control ID; skipping", row_idx)
            continue

        control_name = cell(col_name)
        control_objective = cell(col_objective)

        log.info("── Processing row %d: %s (%s)", row_idx, control_id, control_name)

        # Fetch evidence files from Drive
        evidence = get_evidence_files(drv, control_id)
        log.info("Found %d evidence file(s) for %s", len(evidence), control_id)

        # Analyze via Gemini
        result = analyze_control(control_id, control_name, control_objective, evidence)
        verdict = result.get("verdict", "Failed")
        gaps = result.get("gaps", [])
        summary = result.get("summary", "")

        log.info("Result for %s: %s | gaps=%d", control_id, verdict, len(gaps))
        if gaps:
            for g in gaps:
                log.info("  GAP: %s", g)

        # Build gap text: bullet list of gaps + summary
        gap_text = "\n".join(f"• {g}" for g in gaps)
        if summary and summary not in gap_text:
            gap_text = (gap_text + "\n\n" + summary).strip()

        # Timestamp
        now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

        # Write back to sheet
        updates = []
        if col_result is not None:
            updates.append({"range": f"{SHEET_TAB}!{col_letter(col_result)}{row_idx}", "value": verdict})
        if col_gaps is not None:
            updates.append({"range": f"{SHEET_TAB}!{col_letter(col_gaps)}{row_idx}", "value": gap_text})
        if col_last_run is not None:
            updates.append({"range": f"{SHEET_TAB}!{col_letter(col_last_run)}{row_idx}", "value": now})

        if updates:
            try:
                write_cells(svc, SPREADSHEET_ID, updates)
                log.info("Wrote results to sheet for row %d", row_idx)
            except HttpError as e:
                log.error("Failed to write row %d: %s", row_idx, e)
                failed += 1
                continue

        processed += 1
        if verdict == "Failed":
            failed += 1

        # Brief pause to avoid quota limits
        time.sleep(1)

    log.info(
        "Done. Processed %d control(s), %d failed (or had gaps).",
        processed, failed,
    )

    if failed > 0 and processed > 0:
        sys.exit(2)  # Partial failure — Cloud Run Job shows as failed


if __name__ == "__main__":
    main()
