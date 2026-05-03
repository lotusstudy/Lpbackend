"""
Lotus Academy Management — Backend API
Minimal role: fetch Excel files from Google Drive → extract → return JSON
Deploy to: Render / Railway / Wispbyte
"""

from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import io
import re
import os
from datetime import datetime

# Google Drive (optional — enable if you have credentials)
try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
    GDRIVE_ENABLED = True
except ImportError:
    GDRIVE_ENABLED = False

app = Flask(__name__)
CORS(app, origins=["*"])

# ─── CONFIG ──────────────────────────────────────────────
# Set these in your deployment environment variables
GDRIVE_FOLDER_ID = os.environ.get("GDRIVE_FOLDER_ID", "")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")


# ─── GOOGLE DRIVE HELPERS ────────────────────────────────
def get_drive_service():
    if not GDRIVE_ENABLED or not GOOGLE_SERVICE_ACCOUNT_JSON:
        return None
    import json
    info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    return build("drive", "v3", credentials=creds)


def list_excel_files(service):
    """List all Excel files in the configured Drive folder."""
    query = f"'{GDRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
    results = service.files().list(q=query, fields="files(id,name,modifiedTime)").execute()
    return results.get("files", [])


def download_file(service, file_id):
    """Download file bytes from Drive."""
    from googleapiclient.http import MediaIoBaseDownload
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


# ─── EXCEL PARSERS ───────────────────────────────────────
def parse_date_from_filename(fname: str) -> str:
    """Extract YYYY-MM-DD from filename."""
    m = re.search(r"(\d{4}-\d{2}-\d{2})", fname)
    return m.group(1) if m else datetime.today().strftime("%Y-%m-%d")


def parse_class_from_filename(fname: str) -> str:
    """Extract class label e.g. 'Class 12' from filename."""
    m = re.search(r"Class_(.+?)_\d{4}-\d{2}-\d{2}", fname, re.IGNORECASE)
    return m.group(1).strip() if m else "Unknown"


def parse_student_excel(buf, filename: str) -> dict:
    """
    Parse a student attendance Excel file.
    Expected format:
      Row 0: 🏫 LOTUS Academy | Date | Generated On ...
      Row 1: Report title    | YYYY-MM-DD | ...
      Row 2: blank
      Row 3: column headers (0,1,2,3,4,5,6)
      Row 4+: student data (Name | Class | Board | Stream | Date | Time | Status)
    """
    date = parse_date_from_filename(filename)
    class_label = parse_class_from_filename(filename)

    df_raw = pd.read_excel(buf, header=None)

    # Find the data start row — look for row where col[6] is PRESENT/ABSENT
    data_start = 4
    for i in range(2, min(15, len(df_raw))):
        val = str(df_raw.iloc[i, 6] if df_raw.shape[1] > 6 else "").upper()
        if val in ("PRESENT", "ABSENT"):
            data_start = i
            break

    df = df_raw.iloc[data_start:].reset_index(drop=True)
    df.columns = range(df.shape[1])

    students = []
    for _, row in df.iterrows():
        name = str(row.get(0, "")).strip()
        if not name or len(name) < 2:
            continue
        status = str(row.get(6, "ABSENT")).strip().upper()
        if status not in ("PRESENT", "ABSENT"):
            status = "ABSENT"
        time_val = str(row.get(5, "")).strip()
        if time_val in ("00:00:00", "nan", "NaT", ""):
            time_val = ""
        students.append({
            "name": name,
            "class": str(row.get(1, class_label)).strip() or class_label,
            "board": str(row.get(2, "")).strip(),
            "stream": str(row.get(3, "")).strip(),
            "time": time_val,
            "status": status,
        })

    return {
        "date": date,
        "class": class_label,
        "type": "student",
        "students": students,
        "teachers": [],
    }


def parse_teacher_excel(buf, filename: str) -> dict:
    """
    Parse a teacher attendance Excel file.
    Expected format: Name | Subject | ... | Time | Status
    """
    date = parse_date_from_filename(filename)
    df_raw = pd.read_excel(buf, header=None)

    data_start = 4
    for i in range(2, min(15, len(df_raw))):
        val = str(df_raw.iloc[i, -1] if df_raw.shape[1] > 0 else "").upper()
        if val in ("PRESENT", "ABSENT"):
            data_start = i
            break

    df = df_raw.iloc[data_start:].reset_index(drop=True)
    df.columns = range(df.shape[1])

    teachers = []
    for _, row in df.iterrows():
        name = str(row.get(0, "")).strip()
        if not name or len(name) < 2:
            continue
        ncols = df.shape[1]
        status = str(row.get(ncols - 1, "ABSENT")).strip().upper()
        if status not in ("PRESENT", "ABSENT"):
            status = "ABSENT"
        time_val = str(row.get(ncols - 2, "")).strip()
        if time_val in ("00:00:00", "nan", "NaT", ""):
            time_val = ""
        teachers.append({
            "name": name,
            "subject": str(row.get(1, "")).strip(),
            "time": time_val,
            "status": status,
        })

    return {
        "date": date,
        "class": "teachers",
        "type": "teacher",
        "students": [],
        "teachers": teachers,
    }


# ─── ROUTES ──────────────────────────────────────────────

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "gdrive": GDRIVE_ENABLED and bool(GDRIVE_FOLDER_ID)})


@app.route("/sync-latest", methods=["GET"])
def sync_latest():
    """
    Fetch latest Excel files from Google Drive, parse them, return JSON array.
    Each item: { date, class, students: [...], teachers: [...] }
    Frontend decides whether to store (deduplication handled client-side).
    """
    service = get_drive_service()
    if not service:
        return jsonify([]), 200  # No Drive config — frontend uses manual import

    files = list_excel_files(service)
    records = []

    for f in files:
        fname = f["name"]
        try:
            buf = download_file(service, f["id"])
            is_teacher = re.match(r"^Teachers?_\d{4}-\d{2}-\d{2}\.xlsx$", fname, re.IGNORECASE)
            is_student = re.match(r"^Class_.+_\d{4}-\d{2}-\d{2}\.xlsx$", fname, re.IGNORECASE)

            if is_teacher:
                record = parse_teacher_excel(buf, fname)
                records.append(record)
            elif is_student:
                record = parse_student_excel(buf, fname)
                records.append(record)
        except Exception as e:
            print(f"[WARN] Failed to parse {fname}: {e}")

    return jsonify(records)


@app.route("/parse-upload", methods=["POST"])
def parse_upload():
    """
    Accepts a multipart file upload, parses it, returns JSON.
    Alternative to Drive sync for direct uploads.
    """
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    f = request.files["file"]
    fname = f.filename
    buf = io.BytesIO(f.read())

    try:
        is_teacher = re.match(r"^Teachers?_\d{4}-\d{2}-\d{2}\.xlsx$", fname, re.IGNORECASE)
        if is_teacher:
            record = parse_teacher_excel(buf, fname)
        else:
            record = parse_student_excel(buf, fname)
        return jsonify(record)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ─── ENTRY POINT ─────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
