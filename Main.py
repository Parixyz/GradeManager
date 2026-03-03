# -*- coding: utf-8 -*-
"""
Java Grading Gallery — v2.2
(Pastel + per-student EXACTLY 2 questions + code comments/highlight + PDF + Summary tab + LabID + class stats + histogram)

Key features kept / improved:
- Submissions DB: scan/commit folders, store LabID per student, store file content optionally.
- Robust student ID/name detection:
  - From Java file headers (comment scanning)
  - Optional folder-name patterns (regex) + path fallback
- Flexible scan patterns:
  - Provide a list of glob patterns (e.g., ["*.java", "*.txt"]) AND/OR an include-regex for filenames.
- Grading DB: rubric CSV loading, per-question rubric scoring, rationale, per-student assessed questions (exactly two).
- Code comments tool (like “PDF comment/highlight” concept):
  - Highlight in app
  - In PDF: (a) annotated code snapshot with injected comment blocks AND (b) “highlighted lines” rendering (background for lines with comments).
- Summary tab: full class summary table including LabID + Q1/Q2 IDs + totals + overall.
- Status: class avg/min/max + suggested curve factor + histogram chart and curve overlay.
  - Includes a curve factor control (scale) to preview adjusted distribution.
- Export:
  - Excel (all)
  - Summary PDF
  - Student PDF (single)
  - Batch export PDFs for ALL students into a folder
- AI auto grade is separated into a separate class (optional, can be left empty).
  - No references/branding in UI text. The button is “Auto Grade (optional)” and it safely no-ops if disabled.

Dependencies:
- tkinter, sqlite3, openpyxl
- reportlab (optional for PDF)
- matplotlib (optional for histogram)
"""

import csv
import os
import re
import sqlite3
import hashlib
import math
import random
from pathlib import Path
from datetime import datetime
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from openpyxl import Workbook

# PDF (optional)
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        Preformatted, PageBreak
    )
    from reportlab.lib import colors
except Exception:
    SimpleDocTemplate = None

# Matplotlib (optional, for histogram UI)
try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
except Exception:
    FigureCanvasTkAgg = None
    Figure = None


# =============================================================================
# Config / Constants
# =============================================================================

SUBMISSIONS_DB = "submissions.sqlite"

DEFAULT_THEME = "Grade strictly. If unclear, give 0 and explain why."

# Per your request: keep model/AI plumbing separate; leave default empty/off.
DEFAULT_AI_ENABLED = True

ID_DIGITS_RE = re.compile(r"\b\d{5,12}\b")


# =============================================================================
# Utility
# =============================================================================

def now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def sha256_text(s: str) -> str:
    return hashlib.sha256(s.encode("utf-8", errors="ignore")).hexdigest()

def read_file_text(p: Path) -> str:
    return p.read_text(encoding="utf-8", errors="ignore")


def extract_numeric_id(raw: str) -> str:
    """
    Keep only the numeric student ID portion.
    """
    text = (raw or "").strip()
    if not text:
        return ""
    m = ID_DIGITS_RE.search(text)
    return m.group(0) if m else ""

def is_full_student(student_id: str) -> bool:
    return (student_id or "").strip().lower() == "full"


# =============================================================================
# 1) Student detection from Java header comments
# =============================================================================

def clean_comment_line(line: str) -> str:
    s = line.strip()
    s = s.replace("/*", "").replace("*/", "")
    s = re.sub(r"^\s*//\s*", "", s)
    s = re.sub(r"^\s*\*\s*", "", s)
    return s.strip()

def extract_value(line: str, want: str):
    s = clean_comment_line(line)

    if want == "id":
        s2 = re.sub(r"stud\w*\s*(number|id|no|num)\s*[:=\-]?\s*", "", s, flags=re.IGNORECASE)
        m = ID_DIGITS_RE.search(s2)
        return m.group(0) if m else None

    if want == "name":
        s2 = re.sub(r"stud\w*\s*name\s*[:=\-]?\s*", "", s, flags=re.IGNORECASE).strip()
        if not s2:
            return None
        if ID_DIGITS_RE.search(s2):
            return None
        return s2

    return None

def extract_student_info_from_file(file_path: Path, max_lines: int = 200):
    """
    Tries to parse student id/name from top-of-file comments.
    """
    student_id = None
    student_name = None
    expect_next = None  # "id" or "name"

    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            for i, raw in enumerate(f):
                if i > max_lines:
                    break

                stripped = raw.strip()
                if not stripped:
                    continue

                if not (stripped.startswith("//") or stripped.startswith("/*") or stripped.startswith("*")):
                    continue

                line = clean_comment_line(stripped)
                if not line:
                    continue

                low = line.lower()

                if expect_next == "id" and student_id is None:
                    m = ID_DIGITS_RE.search(line)
                    if m:
                        student_id = m.group(0)
                        expect_next = None
                        continue

                if expect_next == "name" and student_name is None:
                    if "student" not in low and "name" not in low and "number" not in low and "id" not in low:
                        student_name = line.strip()
                        expect_next = None
                        continue

                if student_id is None and ("student" in low or "stud" in low) and (
                    "number" in low or "id" in low or "num" in low or "no" in low
                ):
                    v = extract_value(line, "id")
                    if v:
                        student_id = v
                    else:
                        expect_next = "id"

                if student_name is None and ("student" in low or "stud" in low) and "name" in low:
                    v = extract_value(line, "name")
                    if v:
                        student_name = v
                    else:
                        expect_next = "name"

                if student_id and student_name:
                    break
    except Exception:
        pass

    return student_id, student_name


# =============================================================================
# 2) Folder-name / path-based detection (flexibility you asked for)
# =============================================================================

def try_extract_from_folder_name(folder: Path, id_regex: str, name_regex: str):
    """
    id_regex / name_regex are optional regex strings with a capturing group.
    Examples:
      id_regex   = r"(\\d{7,10})"
      name_regex = r"([A-Za-z]+\\s+[A-Za-z]+)"
    """
    fid = None
    fname = None
    text = str(folder)

    try:
        if id_regex:
            m = re.search(id_regex, text)
            if m:
                fid = m.group(1) if m.groups() else m.group(0)
    except Exception:
        pass

    try:
        if name_regex:
            m = re.search(name_regex, text)
            if m:
                fname = m.group(1) if m.groups() else m.group(0)
    except Exception:
        pass

    return fid, fname

def infer_student_for_folder(
    folder: Path,
    file_globs: list[str],
    include_filename_regex: str,
    folder_id_regex: str,
    folder_name_regex: str,
):
    """
    - Finds candidate files via file_globs (glob patterns) AND optional include_filename_regex.
    - Extracts student info from files (primary)
    - Falls back to folder-based regex extraction if needed.
    """
    detected_id = None
    detected_name = None

    # Collect files by globs
    files = []
    for g in (file_globs or ["*.java"]):
        files.extend(list(folder.rglob(g)))

    # Optional filename regex filter
    if include_filename_regex:
        try:
            rr = re.compile(include_filename_regex)
            files = [p for p in files if rr.search(p.name)]
        except Exception:
            # if regex broken, ignore filter
            pass

    # Prefer .java content for header scan, but allow other files for listing
    java_like = [p for p in files if p.suffix.lower() == ".java"]
    scan_candidates = java_like if java_like else files

    for jf in scan_candidates:
        sid, name = extract_student_info_from_file(jf)
        if sid and not detected_id:
            detected_id = sid
        if name and not detected_name:
            detected_name = name
        if detected_id and detected_name:
            break

    # Optional folder regex fallback (disabled by default; folder names are often machine IDs).
    folder_id = ""
    folder_name = ""
    if folder_id_regex or folder_name_regex:
        folder_id, folder_name = try_extract_from_folder_name(folder, folder_id_regex, folder_name_regex)
        if not detected_id and folder_id:
            detected_id = folder_id
        if not detected_name and folder_name:
            detected_name = folder_name

    final_id = None
    final_name = None

    if detected_id or detected_name:
        if detected_id:
            final_name = (detected_name or "").strip()
            final_id = build_student_key(detected_id, final_name)
        else:
            final_name = (detected_name or "").strip()
            final_id = f"NAME:{final_name}" if final_name else ""

    return final_id, final_name, (detected_id or ""), (detected_name or ""), [str(p) for p in files]


def build_student_key(student_id: str, student_name: str) -> str:
    """
    Build a stable student key.
    Numeric IDs are always stored as numeric-only keys.
    """
    raw_sid = (student_id or "").strip()
    if raw_sid.startswith("NAME:"):
        return raw_sid
    if raw_sid.lower() == "full":
        return raw_sid

    sid = extract_numeric_id(raw_sid)
    sname = re.sub(r"\s+", " ", (student_name or "").strip())

    if not sid:
        return f"NAME:{sname}" if sname else ""
    return sid


def has_required_student_fields(student_id: str, student_name: str) -> bool:
    """
    Treat rows as valid student-like records when:
    - numeric student ID + non-empty student name, OR
    - the special FULL aggregate row.
    """
    raw_sid = (student_id or "").strip()
    sname = re.sub(r"\s+", " ", (student_name or "").strip())

    if raw_sid.lower() == "full":
        return True

    sid = extract_numeric_id(raw_sid)
    if not sid or not sname:
        return False
    if sname.lower() in {"unknown student", "unknown"}:
        return False
    return True


class FolderScannerBase:
    """
    Base scanner so folder scanning can be customized via inheritance later.
    """
    def __init__(self, root_folder: Path, file_globs: list[str], include_filename_regex: str,
                 folder_id_regex: str, folder_name_regex: str):
        self.root_folder = root_folder
        self.file_globs = file_globs or ["*.java"]
        self.include_filename_regex = include_filename_regex
        self.folder_id_regex = folder_id_regex
        self.folder_name_regex = folder_name_regex

    def collect_folders(self) -> list[Path]:
        folders = [p for p in self.root_folder.iterdir() if p.is_dir()]
        root_has_any = any(any(self.root_folder.glob(g)) for g in self.file_globs)
        if root_has_any:
            folders = [self.root_folder] + folders
        return folders

    def detect_folder(self, folder: Path):
        return infer_student_for_folder(
            folder,
            file_globs=self.file_globs,
            include_filename_regex=self.include_filename_regex,
            folder_id_regex=self.folder_id_regex,
            folder_name_regex=self.folder_name_regex,
        )


class DefaultFolderScanner(FolderScannerBase):
    pass


# =============================================================================
# 3) Submissions DB
# =============================================================================

def db_connect(db_path: Path) -> sqlite3.Connection:
    con = sqlite3.connect(db_path)
    con.execute("PRAGMA foreign_keys = ON;")
    return con

def submissions_db_init(con: sqlite3.Connection):
    con.executescript("""
    CREATE TABLE IF NOT EXISTS app_meta (
        meta_key TEXT PRIMARY KEY,
        meta_value TEXT
    );

    CREATE TABLE IF NOT EXISTS students (
        student_id TEXT PRIMARY KEY,
        student_name TEXT NOT NULL,
        lab_id TEXT,
        folder_path TEXT
    );

    CREATE TABLE IF NOT EXISTS files (
        file_path TEXT PRIMARY KEY,
        student_id TEXT,
        source_folder TEXT,
        detected_id TEXT,
        detected_name TEXT,
        file_hash TEXT,
        file_content TEXT,
        last_seen TEXT,
        FOREIGN KEY(student_id) REFERENCES students(student_id) ON DELETE SET NULL
    );
    """)
    # Safe migration
    try:
        cols = [r[1] for r in con.execute("PRAGMA table_info(students)").fetchall()]
        if "lab_id" not in cols:
            con.execute("ALTER TABLE students ADD COLUMN lab_id TEXT;")
    except Exception:
        pass
    con.commit()

def sub_meta_set(con: sqlite3.Connection, key: str, value: str):
    con.execute("""
      INSERT INTO app_meta(meta_key, meta_value)
      VALUES(?, ?)
      ON CONFLICT(meta_key) DO UPDATE SET meta_value=excluded.meta_value
    """, (key, value))
    con.commit()

def sub_meta_get(con: sqlite3.Connection, key: str, default: str = "") -> str:
    row = con.execute("SELECT meta_value FROM app_meta WHERE meta_key=?", (key,)).fetchone()
    return row[0] if row and row[0] is not None else default

def upsert_student(con: sqlite3.Connection, student_id: str, student_name: str, lab_id: str | None, folder_path: str | None):
    con.execute("""
    INSERT INTO students(student_id, student_name, lab_id, folder_path)
    VALUES(?, ?, ?, ?)
    ON CONFLICT(student_id) DO UPDATE SET
      student_name=excluded.student_name,
      lab_id=COALESCE(excluded.lab_id, students.lab_id),
      folder_path=COALESCE(excluded.folder_path, students.folder_path)
    """, (student_id, student_name, lab_id, folder_path))
    con.commit()

def upsert_file(con: sqlite3.Connection, file_path: str, student_id: str | None,
                source_folder: str | None,
                detected_id: str | None, detected_name: str | None,
                file_hash: str | None, file_content: str | None):
    con.execute("""
    INSERT INTO files(file_path, student_id, source_folder, detected_id, detected_name, file_hash, file_content, last_seen)
    VALUES(?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(file_path) DO UPDATE SET
      student_id=excluded.student_id,
      source_folder=excluded.source_folder,
      detected_id=excluded.detected_id,
      detected_name=excluded.detected_name,
      file_hash=excluded.file_hash,
      file_content=excluded.file_content,
      last_seen=excluded.last_seen
    """, (file_path, student_id, source_folder, detected_id, detected_name, file_hash, file_content, now_ts()))
    con.commit()

def get_students(con: sqlite3.Connection):
    cur = con.execute("""
    SELECT s.student_id, s.student_name, COALESCE(s.lab_id,''),
           (SELECT COUNT(*) FROM files f WHERE f.student_id = s.student_id) AS file_count
    FROM students s
    ORDER BY CASE WHEN LOWER(s.student_id)='full' OR LOWER(s.student_name)='full' THEN 0 ELSE 1 END,
             s.student_id
    """)
    return cur.fetchall()

def get_student_files(con: sqlite3.Connection, student_id: str):
    cur = con.execute("""
    SELECT file_path FROM files
    WHERE student_id = ?
    ORDER BY file_path
    """, (student_id,))
    return [r[0] for r in cur.fetchall()]

def get_file_content(con: sqlite3.Connection, file_path: str) -> str | None:
    row = con.execute("SELECT file_content FROM files WHERE file_path=?", (file_path,)).fetchone()
    return row[0] if row else None

def merge_student_code(sub_con: sqlite3.Connection, student_id: str) -> str:
    files = get_student_files(sub_con, student_id)
    parts = []
    for fp in files:
        content = get_file_content(sub_con, fp)
        if content is None:
            try:
                content = Path(fp).read_text(encoding="utf-8", errors="ignore")
            except Exception:
                content = ""
        parts.append(f"\n\n// ===== FILE: {Path(fp).name} =====\n{content}")
    return "\n".join(parts)


# =============================================================================
# 4) Grading DB (rubric + assignments + notes + code comments)
# =============================================================================

def grading_db_init(con: sqlite3.Connection):
    con.executescript("""
    PRAGMA foreign_keys = ON;

    CREATE TABLE IF NOT EXISTS meta(
      meta_key TEXT PRIMARY KEY,
      meta_value TEXT
    );

    CREATE TABLE IF NOT EXISTS rubric_questions(
      question_id TEXT PRIMARY KEY,
      question_title TEXT,
      sub_id TEXT
    );

    CREATE TABLE IF NOT EXISTS rubric_columns(
      question_id TEXT NOT NULL,
      col_key TEXT NOT NULL,
      col_group TEXT,
      col_text TEXT NOT NULL,
      col_max REAL NOT NULL,
      col_order INTEGER NOT NULL,
      PRIMARY KEY(question_id, col_key),
      FOREIGN KEY(question_id) REFERENCES rubric_questions(question_id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS rubric_scores(
      student_id TEXT NOT NULL,
      question_id TEXT NOT NULL,
      col_key TEXT NOT NULL,
      points REAL,
      note TEXT,
      updated_at TEXT,
      PRIMARY KEY(student_id, question_id, col_key)
    );

    CREATE TABLE IF NOT EXISTS student_notes(
      student_id TEXT NOT NULL,
      question_id TEXT NOT NULL,
      rationale TEXT,
      overall_grade REAL,
      updated_at TEXT,
      PRIMARY KEY(student_id, question_id)
    );

    CREATE TABLE IF NOT EXISTS student_assignments(
      student_id TEXT PRIMARY KEY,
      q1 TEXT,
      q2 TEXT,
      updated_at TEXT
    );

    CREATE TABLE IF NOT EXISTS code_comments(
      comment_id INTEGER PRIMARY KEY AUTOINCREMENT,
      student_id TEXT NOT NULL,
      file_path TEXT NOT NULL,
      start_index TEXT NOT NULL,
      end_index TEXT NOT NULL,
      comment_text TEXT NOT NULL,
      color TEXT,
      created_at TEXT
    );
    """)
    try:
        cols = [r[1] for r in con.execute("PRAGMA table_info(rubric_questions)").fetchall()]
        if "sub_id" not in cols:
            con.execute("ALTER TABLE rubric_questions ADD COLUMN sub_id TEXT;")
    except Exception:
        pass
    con.commit()

def meta_set(con: sqlite3.Connection, key: str, value: str):
    con.execute("""
      INSERT INTO meta(meta_key, meta_value)
      VALUES(?, ?)
      ON CONFLICT(meta_key) DO UPDATE SET meta_value=excluded.meta_value
    """, (key, value))
    con.commit()

def meta_get(con: sqlite3.Connection, key: str, default: str = "") -> str:
    row = con.execute("SELECT meta_value FROM meta WHERE meta_key=?", (key,)).fetchone()
    return row[0] if row and row[0] is not None else default

def wipe_rubric(con: sqlite3.Connection):
    con.execute("DELETE FROM rubric_scores")
    con.execute("DELETE FROM student_notes")
    con.execute("DELETE FROM rubric_columns")
    con.execute("DELETE FROM rubric_questions")
    con.commit()

def load_scheme_csv_into_db(con: sqlite3.Connection, csv_path: Path):
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        required = {"question_id", "question_title", "group", "col_key", "col_text", "col_max", "col_order"}
        if not reader.fieldnames:
            raise ValueError("CSV has no header row.")
        missing = required - set([h.strip() for h in reader.fieldnames])
        if missing:
            raise ValueError(f"CSV missing columns: {', '.join(sorted(missing))}")

        wipe_rubric(con)

        has_sub_id = any((h or "").strip() == "sub_id" for h in reader.fieldnames)

        questions = {}
        rows = []
        for r in reader:
            qid = (r.get("question_id") or "").strip()
            qtitle = (r.get("question_title") or "").strip()
            sub_id = (r.get("sub_id") or "").strip() if has_sub_id else ""
            group = (r.get("group") or "").strip()
            col_key = (r.get("col_key") or "").strip()
            col_text = (r.get("col_text") or "").strip()

            if not qid or not col_key or not col_text:
                continue

            try:
                col_max = float((r.get("col_max") or "").strip())
            except Exception:
                col_max = 0.0

            try:
                col_order = int(float((r.get("col_order") or "").strip()))
            except Exception:
                col_order = 0

            questions[qid] = (qtitle or qid, sub_id)
            rows.append((qid, col_key, group, col_text, col_max, col_order))

        for qid, payload in questions.items():
            title, sub_id = payload
            con.execute("""
              INSERT OR REPLACE INTO rubric_questions(question_id, question_title, sub_id)
              VALUES(?, ?, ?)
            """, (qid, title, sub_id))

        for qid, col_key, group, col_text, col_max, col_order in rows:
            con.execute("""
              INSERT OR REPLACE INTO rubric_columns(question_id, col_key, col_group, col_text, col_max, col_order)
              VALUES(?,?,?,?,?,?)
            """, (qid, col_key, group, col_text, col_max, col_order))

        con.commit()
        meta_set(con, "scheme_csv_path", str(csv_path))

def fetch_questions(con: sqlite3.Connection):
    return con.execute("""
      SELECT question_id, question_title, COALESCE(sub_id, '')
      FROM rubric_questions
      ORDER BY question_id
    """).fetchall()

def fetch_columns_for_question(con: sqlite3.Connection, question_id: str):
    return con.execute("""
      SELECT col_key, col_group, col_text, col_max
      FROM rubric_columns
      WHERE question_id=?
      ORDER BY col_order, col_key
    """, (question_id,)).fetchall()

def load_student_scores(con: sqlite3.Connection, student_id: str, question_id: str):
    rows = con.execute("""
      SELECT col_key, points, note
      FROM rubric_scores
      WHERE student_id=? AND question_id=?
    """, (student_id, question_id)).fetchall()
    score_map, note_map = {}, {}
    for k, p, n in rows:
        score_map[k] = p
        note_map[k] = n
    return score_map, note_map

def upsert_score(con: sqlite3.Connection, student_id: str, question_id: str, col_key: str, points: float | None, note: str):
    con.execute("""
      INSERT INTO rubric_scores(student_id, question_id, col_key, points, note, updated_at)
      VALUES(?,?,?,?,?,?)
      ON CONFLICT(student_id, question_id, col_key) DO UPDATE SET
        points=excluded.points,
        note=excluded.note,
        updated_at=excluded.updated_at
    """, (student_id, question_id, col_key, points, note, now_ts()))
    con.commit()

def compute_total(con: sqlite3.Connection, student_id: str, question_id: str | None) -> float:
    if not question_id:
        return 0.0
    row = con.execute("""
      SELECT COALESCE(SUM(COALESCE(points,0)),0)
      FROM rubric_scores
      WHERE student_id=? AND question_id=?
    """, (student_id, question_id)).fetchone()
    return float(row[0] if row else 0.0)

def compute_overall_total(con: sqlite3.Connection, student_id: str) -> float:
    row = con.execute("""
      SELECT COALESCE(SUM(COALESCE(points,0)),0)
      FROM rubric_scores
      WHERE student_id=?
    """, (student_id,)).fetchone()
    return float(row[0] if row else 0.0)
def upsert_student_note(con: sqlite3.Connection, student_id: str, question_id: str, rationale: str, overall_grade: float | None):
    con.execute("""
      INSERT INTO student_notes(student_id, question_id, rationale, overall_grade, updated_at)
      VALUES(?,?,?,?,?)
      ON CONFLICT(student_id, question_id) DO UPDATE SET
        rationale=excluded.rationale,
        overall_grade=excluded.overall_grade,
        updated_at=excluded.updated_at
    """, (student_id, question_id, rationale, overall_grade, now_ts()))
    con.commit()

def load_student_note(con: sqlite3.Connection, student_id: str, question_id: str):
    return con.execute("""
      SELECT rationale, overall_grade
      FROM student_notes
      WHERE student_id=? AND question_id=?
    """, (student_id, question_id)).fetchone()

def upsert_student_assignment(con: sqlite3.Connection, student_id: str, q1: str | None, q2: str | None):
    con.execute("""
      INSERT INTO student_assignments(student_id, q1, q2, updated_at)
      VALUES(?,?,?,?)
      ON CONFLICT(student_id) DO UPDATE SET
        q1=excluded.q1, q2=excluded.q2, updated_at=excluded.updated_at
    """, (student_id, q1, q2, now_ts()))
    con.commit()

def load_student_assignment(con: sqlite3.Connection, student_id: str) -> tuple[str | None, str | None]:
    row = con.execute("SELECT q1, q2 FROM student_assignments WHERE student_id=?", (student_id,)).fetchone()
    if not row:
        return None, None
    return (row[0] or None), (row[1] or None)

def add_code_comment(con: sqlite3.Connection, student_id: str, file_path: str, start_index: str, end_index: str,
                     comment_text: str, color: str = "#FFF2B2"):
    con.execute("""
      INSERT INTO code_comments(student_id, file_path, start_index, end_index, comment_text, color, created_at)
      VALUES(?,?,?,?,?,?,?)
    """, (student_id, file_path, start_index, end_index, comment_text, color, now_ts()))
    con.commit()

def delete_code_comments_in_range(con: sqlite3.Connection, student_id: str, file_path: str, start_index: str, end_index: str):
    con.execute("""
      DELETE FROM code_comments
      WHERE student_id=? AND file_path=?
        AND NOT (end_index<=? OR start_index>=?)
    """, (student_id, file_path, start_index, end_index))
    con.commit()

def fetch_code_comments_for_file(con: sqlite3.Connection, student_id: str, file_path: str):
    return con.execute("""
      SELECT comment_id, start_index, end_index, comment_text, color, created_at
      FROM code_comments
      WHERE student_id=? AND file_path=?
      ORDER BY comment_id
    """, (student_id, file_path)).fetchall()

def fetch_code_comments_for_student(con: sqlite3.Connection, student_id: str):
    return con.execute("""
      SELECT file_path, start_index, end_index, comment_text, color, created_at
      FROM code_comments
      WHERE student_id=?
      ORDER BY file_path, comment_id
    """, (student_id,)).fetchall()


# =============================================================================
# 5) Assignment detection heuristic (kept)
# =============================================================================

def detect_assigned_questions(merged_code: str) -> list[str]:
    s = (merged_code or "").lower()

    tier2_signals = [
        "toofar(", "too far", "pixel", "getpixel", "setpixel", "getred", "getgreen", "getblue",
        "distance", "color distance", "colour distance", "swap", "one pixel"
    ]
    tier1_signals = [
        "drawstar(", "drawshape(", "random", "getsrandom", "getrandomints", "getrandomint",
        "color", "colour", "polygon", "star"
    ]

    tier2_score = sum(1 for k in tier2_signals if k in s)
    tier1_score = sum(1 for k in tier1_signals if k in s)

    if tier2_score >= 2 and tier2_score >= tier1_score:
        return ["P21", "P22"]
    return ["P11", "P12"]


# =============================================================================
# 6) Export: Excel
# =============================================================================

def fetch_all_question_ids(con: sqlite3.Connection):
    rows = con.execute("SELECT question_id FROM rubric_questions ORDER BY question_id").fetchall()
    return [r[0] for r in rows]

def compute_question_max(con: sqlite3.Connection, question_id: str) -> float:
    row = con.execute("""
      SELECT COALESCE(SUM(col_max),0)
      FROM rubric_columns
      WHERE question_id=?
    """, (question_id,)).fetchone()
    return float(row[0] if row else 0.0)

def export_all_to_excel(sub_con: sqlite3.Connection, grade_con: sqlite3.Connection, out_path: Path):
    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Brightspace_Summary"

    question_ids = fetch_all_question_ids(grade_con)
    qmax = {qid: compute_question_max(grade_con, qid) for qid in question_ids}

    ws_sum.append([
        "Student ID", "Student Name", "LabID",
        "Questions graded", "Overall raw"
    ])

    students = sub_con.execute("""
      SELECT student_id, student_name, COALESCE(lab_id,'')
      FROM students
      ORDER BY student_id
    """).fetchall()

    assessed_students = [r for r in students if not is_full_student(r[0])]

    for sid, sname, lab in assessed_students:
        graded_count = grade_con.execute("SELECT COUNT(DISTINCT question_id) FROM rubric_scores WHERE student_id=?", (sid,)).fetchone()[0]
        overall = compute_overall_total(grade_con, sid)
        ws_sum.append([sid, sname, lab, graded_count, overall])

    # Per-question sheets
    for qid in question_ids:
        title_row = grade_con.execute(
            "SELECT question_title FROM rubric_questions WHERE question_id=?",
            (qid,)
        ).fetchone()
        qtitle = title_row[0] if title_row else qid

        ws = wb.create_sheet(title=f"{qid}")
        ws.append([f"{qid} — {qtitle}"])
        ws.append([])

        cols = fetch_columns_for_question(grade_con, qid)

        header = ["Student ID", "Student Name", "LabID", "Total", "Rationale"]
        for col_key, group, text, mx in cols:
            header.append(f"{(group or '').strip()} | {text} (/ {mx:g})".strip(" |"))
            header.append("Note")
        ws.append(header)

        for sid, sname, lab in assessed_students:
            total = compute_total(grade_con, sid, qid)
            note_row = load_student_note(grade_con, sid, qid)
            rationale = note_row[0] if note_row and note_row[0] else ""

            score_map, note_map = load_student_scores(grade_con, sid, qid)

            row = [sid, sname, lab, total, rationale]
            for col_key, _group, _text, _mx in cols:
                row.append("" if score_map.get(col_key) is None else score_map.get(col_key))
                row.append(note_map.get(col_key, "") or "")
            ws.append(row)

    # Code comment sheet (highlighted ranges + comment text)
    ws_comments = wb.create_sheet(title="Code_Comments")
    ws_comments.append(["Student ID", "Student Name", "Code (file)", "Highlighted part", "Comment"])
    for sid, sname, _lab in assessed_students:
        for fp, sidx, eidx, txt, _color, _ts in fetch_code_comments_for_student(grade_con, sid):
            ws_comments.append([sid, sname, Path(fp).name, f"{sidx}–{eidx}", txt or ""])

    # Reference model row if FULL exists
    full_row = next((r for r in students if is_full_student(r[0])), None)
    if full_row:
        sid, sname, lab = full_row
        ws_full = wb.create_sheet(title="FULL_Model")
        ws_full.append(["Student ID", "Student Name", "LabID", "Overall raw"])
        ws_full.append([sid, sname, lab, compute_overall_total(grade_con, sid)])

        ws_full.append([])
        ws_full.append(["Question", "Total", "Rationale"])
        for qid in question_ids:
            qtotal = compute_total(grade_con, sid, qid)
            nrow = load_student_note(grade_con, sid, qid)
            ws_full.append([qid, qtotal, nrow[0] if nrow and nrow[0] else ""])

    # light formatting
    for ws in wb.worksheets:
        try:
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(48, max(12, len(str(col[0].value or "")) + 2))
        except Exception:
            pass

    wb.save(out_path)


# =============================================================================
# 7) PDF Export: student + summary + batch
# =============================================================================

class PDFExporter:
    def __init__(self, sub_con: sqlite3.Connection, grade_con: sqlite3.Connection, question_map: dict[str, str]):
        self.sub_con = sub_con
        self.grade_con = grade_con
        self.question_map = question_map

    def _build_annotated_code_injected(self, sid: str) -> str:
        """
        Annotated code dump: injects comment blocks right before the line where comment starts.
        Uses start_index's line number from Tk indices like '12.0'.
        """
        out = []
        files = get_student_files(self.sub_con, sid)

        for fp in files:
            code = get_file_content(self.sub_con, fp)
            if code is None:
                try:
                    code = Path(fp).read_text(encoding="utf-8", errors="ignore")
                except Exception:
                    code = ""

            comments = self.grade_con.execute("""
              SELECT comment_id, start_index, end_index, comment_text, created_at
              FROM code_comments
              WHERE student_id=? AND file_path=?
              ORDER BY comment_id
            """, (sid, fp)).fetchall()

            by_line = {}
            for cid, sidx, eidx, txt, ts in comments:
                try:
                    line = int(str(sidx).split(".", 1)[0])
                except Exception:
                    line = 1
                by_line.setdefault(line, []).append((cid, sidx, eidx, txt, ts))

            lines = code.splitlines()
            out.append(f"\n\n// ===== FILE: {Path(fp).name} =====\n")

            for i, line_text in enumerate(lines, start=1):
                if i in by_line:
                    for (cid, sidx, eidx, txt, ts) in by_line[i]:
                        out.append(f"// >>> COMMENT #{cid} [{sidx}–{eidx}] ({ts}):")
                        for c_line in (txt or "").splitlines():
                            out.append(f"//     {c_line}")
                        out.append("// >>> END COMMENT\n")
                out.append(line_text)

        return "\n".join(out)

    def _escape_pdf_text(self, text: str) -> str:
        return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _parse_tk_index(self, idx: str) -> tuple[int, int]:
        try:
            line, col = str(idx).split(".", 1)
            return max(1, int(line)), max(0, int(col))
        except Exception:
            return 1, 0

    def _tk_index_to_offset(self, line_starts: list[int], code_len: int, idx: str) -> int:
        line_no, col = self._parse_tk_index(idx)
        if not line_starts:
            return 0
        line_no = min(max(1, line_no), len(line_starts))
        base = line_starts[line_no - 1]
        return min(code_len, max(0, base + col))

    def _render_line_with_highlights(self, line_text: str, line_start: int, line_end: int,
                                     ranges: list[tuple[int, int]], line_no: int) -> str:
        if line_start >= line_end:
            return f"{line_no:04d} |"

        overlaps: list[tuple[int, int]] = []
        for r0, r1 in ranges:
            if r1 <= line_start or r0 >= line_end:
                continue
            overlaps.append((max(r0, line_start), min(r1, line_end)))

        if not overlaps:
            return f"{line_no:04d} | {self._escape_pdf_text(line_text)}"

        overlaps.sort()
        merged: list[tuple[int, int]] = []
        for s, e in overlaps:
            if not merged or s > merged[-1][1]:
                merged.append((s, e))
            else:
                merged[-1] = (merged[-1][0], max(merged[-1][1], e))

        parts = [f"{line_no:04d} | "]
        cursor = line_start
        for s, e in merged:
            if cursor < s:
                parts.append(self._escape_pdf_text(line_text[cursor - line_start:s - line_start]))
            parts.append(f"<font backColor=\"#FFF9A6\">{self._escape_pdf_text(line_text[s - line_start:e - line_start])}</font>")
            cursor = e
        if cursor < line_end:
            parts.append(self._escape_pdf_text(line_text[cursor - line_start:line_end - line_start]))
        return "".join(parts)

    def _build_highlighted_code_blocks(self, sid: str):
        """
        Returns per-file rows rendered with exact character-level highlights,
        matching the Tk text selection ranges used by graders.
        """
        blocks: list[tuple[str, list[str]]] = []
        files = get_student_files(self.sub_con, sid)

        rows = fetch_code_comments_for_student(self.grade_con, sid)
        file_ranges: dict[str, list[tuple[str, str]]] = {}
        for fp, sidx, eidx, _txt, _color, _ts in rows:
            file_ranges.setdefault(fp, []).append((sidx, eidx))

        for fp in files:
            code = get_file_content(self.sub_con, fp)
            if code is None:
                try:
                    code = Path(fp).read_text(encoding="utf-8", errors="ignore")
                except Exception:
                    code = ""
            lines = code.splitlines(keepends=True)

            line_starts: list[int] = []
            pos = 0
            for line in lines:
                line_starts.append(pos)
                pos += len(line)
            code_len = len(code)

            normalized_ranges: list[tuple[int, int]] = []
            for sidx, eidx in file_ranges.get(fp, []):
                start_off = self._tk_index_to_offset(line_starts, code_len, sidx)
                end_off = self._tk_index_to_offset(line_starts, code_len, eidx)
                if end_off < start_off:
                    start_off, end_off = end_off, start_off
                if start_off == end_off:
                    continue
                normalized_ranges.append((start_off, end_off))

            rendered_lines: list[str] = []
            for i, raw_line in enumerate(lines, start=1):
                clean_line = raw_line.rstrip("\r\n")
                line_start = line_starts[i - 1]
                line_end = line_start + len(clean_line)
                rendered_lines.append(self._render_line_with_highlights(clean_line, line_start, line_end, normalized_ranges, i))
            blocks.append((Path(fp).name, rendered_lines))

        return blocks

    def _extract_code_snippet(self, file_path: str, start_index: str, end_index: str, max_chars: int = 420) -> str:
        code = get_file_content(self.sub_con, file_path)
        if code is None:
            try:
                code = Path(file_path).read_text(encoding="utf-8", errors="ignore")
            except Exception:
                code = ""

        lines = code.splitlines(keepends=True)
        if not lines:
            return ""

        def parse_idx(idx: str):
            try:
                a, b = str(idx).split(".", 1)
                return max(1, int(a)), max(0, int(b))
            except Exception:
                return 1, 0

        s_line, s_col = parse_idx(start_index)
        e_line, e_col = parse_idx(end_index)
        s_line = min(max(1, s_line), len(lines))
        e_line = min(max(1, e_line), len(lines))
        if (e_line, e_col) < (s_line, s_col):
            s_line, e_line = e_line, s_line
            s_col, e_col = e_col, s_col

        if s_line == e_line:
            snippet = lines[s_line - 1][s_col:e_col]
        else:
            parts = [lines[s_line - 1][s_col:]]
            for ln in range(s_line, e_line - 1):
                parts.append(lines[ln])
            parts.append(lines[e_line - 1][:e_col])
            snippet = "".join(parts)

        snippet = snippet.strip("\n\r ")
        if len(snippet) > max_chars:
            snippet = snippet[:max_chars].rstrip() + " …"
        return snippet

    def export_student_pdf(self, sid: str, out_path: Path):
        if SimpleDocTemplate is None:
            raise RuntimeError("reportlab not installed. Install: pip install reportlab")

        srow = self.sub_con.execute(
            "SELECT student_name, COALESCE(lab_id,'') FROM students WHERE student_id=?",
            (sid,)
        ).fetchone()
        sname = srow[0] if srow else ""
        lab = srow[1] if srow else ""

        styles = getSampleStyleSheet()
        doc = SimpleDocTemplate(
            str(out_path),
            pagesize=letter,
            title=f"{sid} grading report",
            leftMargin=28,
            rightMargin=28,
            topMargin=28,
            bottomMargin=28,
        )

        story = []
        story.append(Paragraph("<b>Grading Report</b>", styles["Title"]))
        story.append(Paragraph(f"<b>Student:</b> {sid} — {sname}", styles["Normal"]))
        if lab:
            story.append(Paragraph(f"<b>LabID:</b> {lab}", styles["Normal"]))
        overall = compute_overall_total(self.grade_con, sid)
        story.append(Paragraph(f"<b>Overall total:</b> {overall:g}", styles["Normal"]))
        story.append(Spacer(1, 12))

        # Per-question tables
        for qid in fetch_all_question_ids(self.grade_con):
            qtitle = self.question_map.get(qid, qid)
            total = compute_total(self.grade_con, sid, qid)
            if total <= 0 and not fetch_columns_for_question(self.grade_con, qid):
                continue
            story.append(Paragraph(f"<b>{qid}</b> — {qtitle} (Total: {total:g})", styles["Heading2"]))

            cols = fetch_columns_for_question(self.grade_con, qid)
            score_map, note_map = load_student_scores(self.grade_con, sid, qid)

            cell_style = styles["BodyText"].clone("rubric_cell_style")
            cell_style.fontSize = 8
            cell_style.leading = 10

            table_data = [["Group", "Criterion", "Max", "Pts", "Note"]]
            for col_key, group, text, mx in cols:
                pts = score_map.get(col_key, 0.0) or 0.0
                note = (note_map.get(col_key, "") or "")
                esc_group = (group or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                esc_text = (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                esc_note = note[:240].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                table_data.append([
                    Paragraph(esc_group, cell_style),
                    Paragraph(esc_text, cell_style),
                    f"{mx:g}",
                    f"{pts:g}",
                    Paragraph(esc_note, cell_style),
                ])

            tbl = Table(table_data, colWidths=[80, 180, 40, 40, 216], repeatRows=1)
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EFE5FF")),
                ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#3B2D5C")),
                ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#B7B7B7")),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("FONTSIZE", (0,0), (-1,-1), 8),
                ("LEFTPADDING", (0,0), (-1,-1), 4),
                ("RIGHTPADDING", (0,0), (-1,-1), 4),
                ("TOPPADDING", (0,0), (-1,-1), 3),
                ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 10))

            note_row = load_student_note(self.grade_con, sid, qid)
            rationale = note_row[0] if note_row and note_row[0] else ""
            if rationale:
                story.append(Paragraph("<b>Rationale</b>", styles["Heading3"]))
                story.append(Paragraph(rationale.replace("\n", "<br/>"), styles["Normal"]))
                story.append(Spacer(1, 10))

        # Highlighted code + comments table (compact, printable)
        story.append(PageBreak())
        story.append(Paragraph("<b>Highlighted Code Review</b>", styles["Heading2"]))
        rows = fetch_code_comments_for_student(self.grade_con, sid)
        if not rows:
            story.append(Paragraph("(No code comments.)", styles["Normal"]))
        else:
            cell_style = styles["BodyText"].clone("comment_cell_style")
            cell_style.fontSize = 8
            cell_style.leading = 10

            td = [["File", "Range", "Code (highlighted part)", "Comment"]]
            for fp, sidx, eidx, txt, _color, _ts in rows:
                snippet = self._extract_code_snippet(fp, sidx, eidx)
                esc_file = Path(fp).name.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                esc_snippet = (snippet or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                esc_comment = (txt or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                td.append([
                    Paragraph(esc_file, cell_style),
                    Paragraph(f"{sidx}–{eidx}", cell_style),
                    Paragraph(esc_snippet or "(empty selection)", cell_style),
                    Paragraph(esc_comment[:320], cell_style),
                ])

            tbl2 = Table(td, colWidths=[85, 90, 200, 181], repeatRows=1)
            tbl2.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EFE5FF")),
                ("TEXTCOLOR", (0,0), (-1,0), colors.HexColor("#FF4FA3")),
                ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#B7B7B7")),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                ("BACKGROUND", (2,1), (2,-1), colors.HexColor("#FFF9A6")),
                ("FONTSIZE", (0,0), (-1,-1), 8),
                ("LEFTPADDING", (0,0), (-1,-1), 4),
                ("RIGHTPADDING", (0,0), (-1,-1), 4),
                ("TOPPADDING", (0,0), (-1,-1), 3),
                ("BOTTOMPADDING", (0,0), (-1,-1), 3),
            ]))
            story.append(tbl2)

        # Full code listing with highlighted markers
        story.append(PageBreak())
        story.append(Paragraph("<b>Full Code Listing</b>", styles["Heading2"]))
        blocks = self._build_highlighted_code_blocks(sid)
        line_style = styles["Code"].clone("code_line_style")
        line_style.fontSize = 7
        line_style.leading = 8
        for bi, (fname, rendered_lines) in enumerate(blocks):
            story.append(Paragraph(f"<b>{fname}</b>", styles["Heading3"]))
            if not rendered_lines:
                story.append(Paragraph("(empty file)", styles["Normal"]))
            else:
                table_data = [[Paragraph(line or " ", line_style)] for line in rendered_lines]
                tbl_code = Table(table_data, colWidths=[516], repeatRows=0)
                tbl_code.setStyle(TableStyle([
                    ("GRID", (0,0), (-1,-1), 0.2, colors.HexColor("#E0E0E0")),
                    ("LEFTPADDING", (0,0), (-1,-1), 4),
                    ("RIGHTPADDING", (0,0), (-1,-1), 4),
                    ("TOPPADDING", (0,0), (-1,-1), 1),
                    ("BOTTOMPADDING", (0,0), (-1,-1), 1),
                ]))
                story.append(tbl_code)
            if bi < len(blocks) - 1:
                story.append(PageBreak())

        doc.build(story)

    def export_summary_pdf(self, out_path: Path, class_stats_text: str):
        if SimpleDocTemplate is None:
            raise RuntimeError("reportlab not installed. Install: pip install reportlab")

        styles = getSampleStyleSheet()
        doc = SimpleDocTemplate(str(out_path), pagesize=letter, title="Summary")
        story = []
        story.append(Paragraph("<b>Grading Summary</b>", styles["Title"]))
        story.append(Spacer(1, 8))
        story.append(Paragraph(class_stats_text, styles["Normal"]))
        story.append(Spacer(1, 12))

        students = self.sub_con.execute("""
          SELECT student_id, student_name, COALESCE(lab_id,'')
          FROM students
        """).fetchall()

        def key(r):
            sid = (r[0] or "")
            return (0 if sid.lower() == "full" else 1, sid)
        students = sorted(students, key=key)

        td = [["Student ID", "Name", "LabID", "Questions graded", "Overall"]]
        for sid, sname, lab in students:
            graded_count = self.grade_con.execute("SELECT COUNT(DISTINCT question_id) FROM rubric_scores WHERE student_id=?", (sid,)).fetchone()[0]
            overall = compute_overall_total(self.grade_con, sid)
            td.append([sid, sname, lab, str(graded_count), f"{overall:g}"])

        tbl = Table(td, colWidths=[90, 170, 80, 110, 80])
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EFE5FF")),
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ]))
        story.append(tbl)

        doc.build(story)

    def export_all_students_pdfs(self, out_dir: Path, report_tag: str = "Midterm", progress_cb=None):
        out_dir.mkdir(parents=True, exist_ok=True)
        students = self.sub_con.execute("""
          SELECT student_id, COALESCE(lab_id,'')
          FROM students
          WHERE LOWER(student_id) <> 'full'
          ORDER BY student_id
        """).fetchall()

        safe_tag = re.sub(r"[^A-Za-z0-9_-]+", "", (report_tag or "Midterm").strip()) or "Midterm"

        for i, (sid, lab) in enumerate(students, start=1):
            safe_lab = re.sub(r"[^A-Za-z0-9_-]+", "", (lab or "").strip()) or "LabX"
            out_name = f"{sid}_Report_{safe_lab}_{safe_tag}.pdf"
            out_path = out_dir / out_name
            try:
                self.export_student_pdf(sid, out_path)
            except Exception as e:
                # keep going, but record failures
                if progress_cb:
                    progress_cb(i, len(students), sid, False, str(e))
                continue
            if progress_cb:
                progress_cb(i, len(students), sid, True, "")

    def export_compare_to_full_pdfs(self, out_dir: Path, report_tag: str = "Midterm", progress_cb=None):
        """
        Alias for batch export naming used from the "Compare to FULL" workflow.
        """
        self.export_all_students_pdfs(out_dir=out_dir, report_tag=report_tag, progress_cb=progress_cb)


# =============================================================================
# 8) Optional Auto Grader (separated; can be empty/void)
# =============================================================================

class AutoGrader:
    """
    Optional component.
    - Keep disabled by default.
    - Implement your own logic later (unit tests, static checks, etc.).
    """
    def __init__(self, enabled: bool = False):
        self.enabled = enabled

    def auto_grade(self, *, merged_code: str, rubric_items: list[dict], theme_text: str) -> dict:
        """
        Return structure:
          {
            "scores": [{"col_key": "...", "points": float, "note": "..."}, ...],
            "rationale": "..."
          }
        If disabled, raise or return empty.
        """
        if not self.enabled:
            raise RuntimeError("AutoGrader is disabled (optional component).")

        scores = []
        for item in (rubric_items or []):
            col_key = item.get("col_key")
            if not col_key:
                continue
            max_points = float(item.get("max_points", 0.0) or 0.0)
            pts = random.uniform(0.0, max_points) if max_points > 0 else 0.0
            pts = round(pts, 2)
            scores.append({
                "col_key": col_key,
                "points": pts,
                "note": "Auto draft (random in range).",
            })

        return {
            "scores": scores,
            "rationale": "Auto-generated draft scores (randomized within each rubric max). Please review.",
        }


# =============================================================================
# 9) UI widgets
# =============================================================================

class ScrollableRubricGrid(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.inner = ttk.Frame(self.canvas)
        self.inner_id = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")

        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.inner.bind("<Configure>", lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda _e: self.canvas.itemconfigure(self.inner_id))

        self.columns = []
        self.score_vars = {}
        self.note_vars = {}

    def build(self, columns):
        for w in self.inner.winfo_children():
            w.destroy()

        self.columns = columns
        self.score_vars.clear()
        self.note_vars.clear()

        ttk.Label(self.inner, text="Criterion", style="Pastel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(self.inner, text="Score", style="Pastel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(self.inner, text="Note", style="Pastel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=2, sticky="w", padx=6, pady=4)

        r = 1
        current_group = None
        for col_key, group, text, maxp in columns:
            group = group or ""

            if group != current_group:
                current_group = group
                ttk.Separator(self.inner, orient="horizontal").grid(row=r, column=0, columnspan=3, sticky="ew", pady=(10, 6))
                r += 1
                if current_group.strip():
                    ttk.Label(self.inner, text=current_group, style="Pastel.TLabel", font=("Segoe UI", 10, "bold")).grid(row=r, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 4))
                    r += 1

            header = f"{text}\n/{maxp:g}"
            ttk.Label(self.inner, text=header, style="Pastel.TLabel", justify="left", wraplength=560).grid(row=r, column=0, sticky="w", padx=6, pady=4)

            sv = tk.StringVar(value="")
            nv = tk.StringVar(value="")
            ttk.Entry(self.inner, textvariable=sv, width=8).grid(row=r, column=1, sticky="w", padx=6, pady=4)
            ttk.Entry(self.inner, textvariable=nv, width=60).grid(row=r, column=2, sticky="ew", padx=6, pady=4)

            self.score_vars[col_key] = sv
            self.note_vars[col_key] = nv
            r += 1

        self.inner.columnconfigure(2, weight=1)

    def set_values(self, score_map, note_map):
        for k, sv in self.score_vars.items():
            v = score_map.get(k, None)
            sv.set("" if v is None else str(v))
        for k, nv in self.note_vars.items():
            nv.set(note_map.get(k, "") or "")

    def get_values(self):
        scores = {k: v.get().strip() for k, v in self.score_vars.items()}
        notes = {k: v.get().strip() for k, v in self.note_vars.items()}
        return scores, notes


# =============================================================================
# 10) Pastel theme
# =============================================================================

def pastel_style(root: tk.Tk):
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    bg = "#F7F2FF"        # very light purple
    panel = "#FFFDE8"     # lemon-cream yellow
    accent = "#DCCBFF"    # light purple
    accent2 = "#FF4FA3"   # bright pink (used sparingly: active states)
    text = "#2A2440"
    select = "#EFE5FF"

    root.configure(bg=bg)

    style.configure(".", font=("Segoe UI", 10))
    style.configure("Pastel.TFrame", background=bg)
    style.configure("PastelCard.TFrame", background=panel, relief="flat")
    style.configure("Pastel.TLabel", background=bg, foreground=text)
    style.configure("PastelCard.TLabel", background=panel, foreground=text)

    style.configure("TButton", padding=6)
    style.map("TButton", background=[("active", accent2)], foreground=[("disabled", "#888")])

    style.configure("TNotebook", background=bg, borderwidth=0)
    style.configure("TNotebook.Tab", padding=(12, 6))
    style.map("TNotebook.Tab",
              background=[("selected", panel), ("!selected", bg)],
              foreground=[("selected", text), ("!selected", text)])

    style.configure("Treeview",
                    background=panel,
                    fieldbackground=panel,
                    foreground=text,
                    rowheight=24,
                    borderwidth=0)
    style.map("Treeview", background=[("selected", select)], foreground=[("selected", text)])

    style.configure("Treeview.Heading",
                    font=("Segoe UI", 10, "bold"),
                    background=accent,
                    foreground=text,
                    relief="flat")
    style.map("Treeview.Heading", background=[("active", accent2)])

    return {"bg": bg, "panel": panel, "accent": accent, "accent2": accent2, "text": text, "select": select}


# =============================================================================
# 11) Scan Window (folder-by-folder review + LabID + scan flexibility)
# =============================================================================

class ScanWindow(tk.Toplevel):
    def __init__(self, parent_app):
        super().__init__(parent_app.root)
        self.app = parent_app
        self.con = parent_app.sub_con

        self.title("Scan/Edit Submissions")
        self.geometry("1650x920")

        self.root_folder: Path | None = None
        # Always store file content so previews are consistently available after reload.
        self.store_file_content = tk.BooleanVar(value=True)

        # Flex settings
        self.file_globs_var = tk.StringVar(value="*.java")
        self.filename_regex_var = tk.StringVar(value="")  # optional
        self.folder_id_regex_var = tk.StringVar(value="")
        self.folder_name_regex_var = tk.StringVar(value="")
        self.only_new_files_var = tk.BooleanVar(value=False)
        self.global_lab_id_var = tk.StringVar(value="")
        self.find_var = tk.StringVar(value="")
        self._find_from = "1.0"

        self.rows: dict[str, dict] = {}
        self.folder_order: list[str] = []

        # Skimming mode (quick review through files/students)
        self.skim_running = False
        self.skim_delay_ms_var = tk.IntVar(value=300)
        self._skim_folder_idx = 0
        self._skim_file_idx = 0
        self._skimmable_folder_keys: list[str] = []
        self.selected_folder_key: str | None = None

        self._last_folder_selection: tuple[str, ...] = ()
        self._last_scan_file_selection: tuple[str, ...] = ()

        self._build()
        self.load_existing_rows_from_db()
        self.after(120, self._poll_scan_selections)

    def _build(self):
        top = ttk.Frame(self, padding=10, style="Pastel.TFrame")
        top.pack(fill=tk.X)

        actions = ttk.Frame(top, style="Pastel.TFrame")
        actions.pack(fill=tk.X)

        ttk.Button(actions, text="Choose ROOT Folder", command=self.choose_root).pack(side=tk.LEFT)
        ttk.Button(actions, text="Scan / Rescan", command=self.scan).pack(side=tk.LEFT, padx=8)

        ttk.Button(actions, text="Save Scan to DB", command=self.save_to_db).pack(side=tk.RIGHT)

        self.status_lbl = ttk.Label(actions, text="No folder selected.", style="Pastel.TLabel")
        self.status_lbl.pack(side=tk.RIGHT, padx=10)

        opts_nb = ttk.Notebook(top)
        opts_nb.pack(fill=tk.X, pady=(8, 0))

        filters_tab = ttk.Frame(opts_nb, style="Pastel.TFrame", padding=8)
        regex_tab = ttk.Frame(opts_nb, style="Pastel.TFrame", padding=8)
        skim_tab = ttk.Frame(opts_nb, style="Pastel.TFrame", padding=8)
        opts_nb.add(filters_tab, text="Scan Filters")
        opts_nb.add(regex_tab, text="Regex")
        opts_nb.add(skim_tab, text="Skimming")

        ttk.Label(filters_tab, text="File globs (comma-separated)", style="Pastel.TLabel").pack(side=tk.LEFT)
        ttk.Entry(filters_tab, textvariable=self.file_globs_var, width=30).pack(side=tk.LEFT, padx=(6, 12))
        ttk.Label(filters_tab, text="Filename include-regex (optional)", style="Pastel.TLabel").pack(side=tk.LEFT)
        ttk.Entry(filters_tab, textvariable=self.filename_regex_var, width=34).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Checkbutton(filters_tab, text="Only show files not already registered", variable=self.only_new_files_var).pack(side=tk.LEFT, padx=(12, 0))

        ttk.Label(regex_tab, text="Folder ID regex (cap group)", style="Pastel.TLabel").pack(side=tk.LEFT)
        ttk.Entry(regex_tab, textvariable=self.folder_id_regex_var, width=24).pack(side=tk.LEFT, padx=(6, 12))
        ttk.Label(regex_tab, text="Folder Name regex (cap group)", style="Pastel.TLabel").pack(side=tk.LEFT)
        ttk.Entry(regex_tab, textvariable=self.folder_name_regex_var, width=32).pack(side=tk.LEFT, padx=(6, 12))
        ttk.Button(regex_tab, text="Reset Regex", command=self.reset_regex_defaults).pack(side=tk.LEFT)
        ttk.Button(regex_tab, text="Save Regex Copy", command=self.save_regex_copy).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(regex_tab, text="Load Regex Copy", command=self.load_regex_copy).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Label(skim_tab, text="Skim delay (ms)", style="Pastel.TLabel").pack(side=tk.LEFT)
        ttk.Entry(skim_tab, textvariable=self.skim_delay_ms_var, width=7).pack(side=tk.LEFT, padx=(6, 10))
        ttk.Button(skim_tab, text="Start Skimming (from selected)", command=self.start_skimming).pack(side=tk.LEFT)
        ttk.Button(skim_tab, text="Stop Skimming", command=self.stop_skimming).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(skim_tab, text="Stop keys: Q / Esc", style="Pastel.TLabel").pack(side=tk.LEFT, padx=(12, 0))

        main = ttk.Frame(self, padding=10, style="Pastel.TFrame")
        main.pack(fill=tk.BOTH, expand=True)

        main.columnconfigure(0, weight=2)
        main.columnconfigure(1, weight=2)
        main.columnconfigure(2, weight=3)
        main.rowconfigure(0, weight=1)

        left = ttk.Frame(main, style="Pastel.TFrame")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.rowconfigure(1, weight=3)
        left.rowconfigure(3, weight=2)
        ttk.Label(left, text="Full list (all scanned folders)", style="Pastel.TLabel").grid(row=0, column=0, sticky="w")

        cols = ("include", "folder", "det_id", "det_name", "final_id", "final_name", "lab_id", "nfiles")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse")
        for c, w in [("include", 70), ("folder", 330), ("det_id", 120), ("det_name", 160),
                     ("final_id", 160), ("final_name", 170), ("lab_id", 90), ("nfiles", 70)]:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="w")
        self.tree.grid(row=1, column=0, sticky="nsew")

        ttk.Label(left, text="Student list (valid ID + name)", style="Pastel.TLabel").grid(row=2, column=0, sticky="w", pady=(10, 0))
        student_cols = ("student_id", "student_name", "lab_id", "folder", "nfiles")
        self.student_tree = ttk.Treeview(left, columns=student_cols, show="headings", selectmode="none", height=8)
        for c, w in [("student_id", 130), ("student_name", 170), ("lab_id", 90), ("folder", 220), ("nfiles", 60)]:
            self.student_tree.heading(c, text=c)
            self.student_tree.column(c, width=w, anchor="w")
        self.student_tree.grid(row=3, column=0, sticky="nsew")

        mid = ttk.Frame(main, style="Pastel.TFrame")
        mid.grid(row=0, column=1, sticky="nsew", padx=(0, 10))
        mid.columnconfigure(0, weight=1)

        ttk.Label(mid, text="Selected Folder Edit", style="Pastel.TLabel").grid(row=0, column=0, sticky="w")
        self.sel_folder_var = tk.StringVar()
        ttk.Entry(mid, textvariable=self.sel_folder_var, state="readonly").grid(row=1, column=0, sticky="ew", pady=(2, 8))

        self.include_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(mid, text="Include this folder in commit", variable=self.include_var,
                        command=self.apply_include_toggle).grid(row=2, column=0, sticky="w")

        ttk.Label(mid, text="Final Student ID", style="Pastel.TLabel").grid(row=3, column=0, sticky="w")
        self.final_id_var = tk.StringVar()
        ttk.Entry(mid, textvariable=self.final_id_var).grid(row=4, column=0, sticky="ew")

        ttk.Label(mid, text="Final Student Name", style="Pastel.TLabel").grid(row=5, column=0, sticky="w", pady=(8, 0))
        self.final_name_var = tk.StringVar()
        ttk.Entry(mid, textvariable=self.final_name_var).grid(row=6, column=0, sticky="ew")

        ttk.Label(mid, text="LabID (section/group)", style="Pastel.TLabel").grid(row=7, column=0, sticky="w", pady=(8, 0))
        self.lab_id_var = tk.StringVar()
        ttk.Entry(mid, textvariable=self.lab_id_var).grid(row=8, column=0, sticky="ew")

        ttk.Label(mid, text="Apply LabID to included folders", style="Pastel.TLabel").grid(row=9, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(mid, textvariable=self.global_lab_id_var).grid(row=10, column=0, sticky="ew")
        ttk.Button(mid, text="Apply LabID to All Included", command=self.apply_global_lab_id).grid(row=11, column=0, sticky="ew", pady=(4, 0))

        ttk.Label(mid, text="Selection shortcuts", style="Pastel.TLabel").grid(row=12, column=0, sticky="w", pady=(10, 0))
        quick = ttk.Frame(mid, style="Pastel.TFrame")
        quick.grid(row=13, column=0, sticky="ew")
        ttk.Button(quick, text="Use selected text as ID (I)", command=self.use_selection_as_id).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(quick, text="Use selected text as Name (N)", command=self.use_selection_as_name).pack(side=tk.LEFT)

        self._suspend_auto_apply = False
        self._bind_auto_apply()

        right = ttk.Frame(main, style="Pastel.TFrame")
        right.grid(row=0, column=2, sticky="nsew")
        right.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)

        ttk.Label(right, text="Files in Selected Folder", style="Pastel.TLabel").grid(row=0, column=0, sticky="w")
        self.files_tree = ttk.Treeview(right, columns=("file",), show="headings", height=10, selectmode="browse")
        self.files_tree.heading("file", text="file path")
        self.files_tree.column("file", width=520, anchor="w")
        self.files_tree.grid(row=1, column=0, sticky="nsew")
        self.files_tree.bind("<<TreeviewSelect>>", self.on_scan_file_select)

        preview_frame = ttk.Frame(right, style="Pastel.TFrame")
        preview_frame.grid(row=2, column=0, sticky="nsew")
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)

        self.preview = tk.Text(preview_frame, wrap="none")
        self.preview.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.preview.configure(yscrollcommand=sb.set)
        self.preview.tag_configure("find_hit", background="#ffe082")

        find_row = ttk.Frame(right, style="Pastel.TFrame")
        find_row.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(find_row, text="Find in file", style="Pastel.TLabel").pack(side=tk.LEFT)
        self.find_entry = ttk.Entry(find_row, textvariable=self.find_var, width=26)
        self.find_entry.pack(side=tk.LEFT, padx=(6, 6))
        ttk.Button(find_row, text="Find Next", command=self.find_next).pack(side=tk.LEFT)

        self.preview.bind("<KeyPress-i>", self._hotkey_use_id)
        self.preview.bind("<KeyPress-I>", self._hotkey_use_id)
        self.preview.bind("<KeyPress-n>", self._hotkey_use_name)
        self.preview.bind("<KeyPress-N>", self._hotkey_use_name)
        self.preview.bind("<Control-f>", lambda _e: self.focus_find_entry())
        self.bind("<KeyPress-q>", lambda _e: self.stop_skimming())
        self.bind("<Escape>", lambda _e: self.stop_skimming())

    def _poll_scan_selections(self):
        folder_sel = tuple(self.tree.selection())
        if folder_sel != self._last_folder_selection:
            self._last_folder_selection = folder_sel
            if folder_sel:
                self.on_folder_select()

        file_sel = tuple(self.files_tree.selection())
        if file_sel != self._last_scan_file_selection:
            self._last_scan_file_selection = file_sel
            if file_sel:
                self.on_scan_file_select()

        self.after(120, self._poll_scan_selections)

    def choose_root(self):
        folder = filedialog.askdirectory(title="Select ROOT submissions folder")
        if not folder:
            return
        self.root_folder = Path(folder)
        self._set_scan_status(prefix=f"Root: {self.root_folder}")
        self.rows.clear()
        self.folder_order.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
        self.preview.delete("1.0", tk.END)

    def _parse_globs(self) -> list[str]:
        raw = (self.file_globs_var.get() or "").strip()
        if not raw:
            return ["*.java"]
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        return parts if parts else ["*.java"]

    def _scan_counts(self) -> dict:
        total_folders = len(self.folder_order)
        total_files = 0
        blank_folders = 0
        total_students = 0
        included_rows = 0
        included_students = 0

        for folder_key in self.folder_order:
            r = self.rows.get(folder_key, {})
            files = r.get("files") or []
            nfiles = len(files)
            total_files += nfiles
            if nfiles == 0:
                blank_folders += 1

            is_student = has_required_student_fields(r.get("final_id", ""), r.get("final_name", ""))
            if is_student:
                total_students += 1

            if r.get("include"):
                included_rows += 1
                if is_student:
                    included_students += 1

        return {
            "total_computers": total_folders,
            "folders": total_folders,
            "files": total_files,
            "blank_folders": blank_folders,
            "total_students": total_students,
            "included_rows": included_rows,
            "included_students": included_students,
        }

    def _set_scan_status(self, prefix: str = ""):
        c = self._scan_counts()
        parts = [
            f"Total Computers: {c['total_computers']}",
            f"Files: {c['files']}",
            f"Total students: {c['total_students']}",
            f"Included rows: {c['included_rows']}",
            f"Included students: {c['included_students']}",
            f"Blank folders: {c['blank_folders']}",
        ]
        msg = " | ".join(parts)
        if prefix:
            msg = f"{prefix} | {msg}"
        self.status_lbl.config(text=msg)

    def reset_regex_defaults(self):
        self.file_globs_var.set("*.java")
        self.filename_regex_var.set("")
        self.folder_id_regex_var.set("")
        self.folder_name_regex_var.set("")


    def _bind_auto_apply(self):
        self.final_id_var.trace_add("write", self._on_manual_edit)
        self.final_name_var.trace_add("write", self._on_manual_edit)
        self.lab_id_var.trace_add("write", self._on_manual_edit)

    def _on_manual_edit(self, *_args):
        if self._suspend_auto_apply:
            return
        self.apply_edits()

    @staticmethod
    def _normalized_name(name: str, folder_label: str, detected_name: str = "") -> str:
        """
        Keep blank names blank.
        Do not auto-store folder names as student names when nothing was detected.
        """
        cleaned = re.sub(r"\s+", " ", (name or "").strip())
        if not cleaned:
            return ""

        folder_base = (Path(folder_label).name if folder_label else "").strip()
        detected = re.sub(r"\s+", " ", (detected_name or "").strip())
        if folder_base and cleaned == folder_base and not detected:
            return ""
        return cleaned

    def _scan_settings_payload(self) -> dict:
        return {
            "file_globs": (self.file_globs_var.get() or "").strip() or "*.java",
            "filename_regex": (self.filename_regex_var.get() or "").strip(),
            "folder_id_regex": (self.folder_id_regex_var.get() or "").strip(),
            "folder_name_regex": (self.folder_name_regex_var.get() or "").strip(),
            "only_new_files": bool(self.only_new_files_var.get()),
        }

    def save_regex_copy(self):
        path = filedialog.asksaveasfilename(
            title="Save regex copy",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        payload = self._scan_settings_payload()
        try:
            Path(path).write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("Save failed", str(e))
            return
        messagebox.showinfo("Saved", f"Regex copy saved:\n{path}")

    def load_regex_copy(self):
        path = filedialog.askopenfilename(
            title="Load regex copy",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as e:
            messagebox.showerror("Load failed", str(e))
            return

        self.file_globs_var.set((payload.get("file_globs") or "*.java").strip() or "*.java")
        self.filename_regex_var.set((payload.get("filename_regex") or "").strip())
        self.folder_id_regex_var.set((payload.get("folder_id_regex") or "").strip())
        self.folder_name_regex_var.set((payload.get("folder_name_regex") or "").strip())
        self.only_new_files_var.set(bool(payload.get("only_new_files", False)))
        messagebox.showinfo("Loaded", f"Regex copy loaded:\n{path}")

    def load_existing_rows_from_db(self):
        """
        Populate Scan/Edit with current submissions DB rows.
        Includes both saved students and any file groups that are not currently
        attached to a valid student record so they can be edited/included later.
        """
        self.rows.clear()
        self.folder_order.clear()

        students = self.con.execute("""
            SELECT student_id, student_name, COALESCE(lab_id,''), COALESCE(folder_path,'')
            FROM students
            ORDER BY student_id
        """).fetchall()
        student_map = {sid: {"name": sname or "", "lab": lab or "", "folder_path": folder_path or ""}
                       for sid, sname, lab, folder_path in students}

        used_keys = set()
        known_students = {sid for sid, *_rest in students}

        def _unique_key(base_key: str) -> str:
            folder_key = base_key
            n = 2
            while folder_key in used_keys:
                folder_key = f"{base_key}#{n}"
                n += 1
            used_keys.add(folder_key)
            return folder_key

        folder_groups: dict[str, dict] = {}

        for sid, sname, lab, folder_path in students:
            base_key = (folder_path or "").strip() or f"DB:{sid}"
            grp = folder_groups.setdefault(base_key, {
                "files": [],
                "det_id": "",
                "det_name": "",
                "student_id": "",
                "lab_id": "",
            })
            grp["student_id"] = sid or grp.get("student_id", "")
            grp["det_id"] = sid or grp.get("det_id", "")
            grp["det_name"] = sname or grp.get("det_name", "")
            grp["lab_id"] = lab or grp.get("lab_id", "")

        file_rows = self.con.execute("""
            SELECT
              file_path,
              COALESCE(student_id, ''),
              COALESCE(source_folder, ''),
              COALESCE(detected_id, ''),
              COALESCE(detected_name, '')
            FROM files
            ORDER BY source_folder, file_path
        """).fetchall()

        for fp, sid, source_folder, det_id, det_name in file_rows:
            sid = (sid or "").strip()
            folder_path = (source_folder or "").strip() or str(Path(fp).parent)
            if not folder_path:
                folder_path = f"UNASSIGNED:{fp}"

            grp = folder_groups.setdefault(folder_path, {
                "files": [],
                "det_id": "",
                "det_name": "",
                "student_id": "",
                "lab_id": "",
            })
            grp["files"].append(fp)
            if sid and not grp["student_id"]:
                grp["student_id"] = sid
            if sid and not grp["lab_id"]:
                grp["lab_id"] = (student_map.get(sid) or {}).get("lab", "")
            if det_id and not grp["det_id"]:
                grp["det_id"] = det_id
            if det_name and not grp["det_name"]:
                grp["det_name"] = det_name

        represented_students = set()

        for folder_path in sorted(folder_groups.keys()):
            grp = folder_groups[folder_path]
            files = grp.get("files") or []
            sid = (grp.get("student_id") or "").strip()
            student_row = student_map.get(sid)
            if student_row:
                represented_students.add(sid)

            det_id = (grp.get("det_id") or "").strip()
            det_name = (grp.get("det_name") or "").strip()
            final_id = sid if student_row else det_id
            final_name = self._normalized_name(
                student_row["name"] if student_row else det_name,
                folder_path,
                det_name,
            )
            lab = student_row["lab"] if student_row else (grp.get("lab_id") or "")

            include_row = bool(files) and has_required_student_fields(final_id, final_name)
            folder_key = _unique_key(folder_path)
            self.folder_order.append(folder_key)
            self.rows[folder_key] = {
                "include": include_row,
                "manual_include_override": None,
                "folder": folder_path,
                "det_id": final_id or det_id,
                "det_name": final_name or det_name,
                "final_id": final_id,
                "final_name": final_name,
                "lab_id": lab,
                "files": files,
            }

        # Keep students without files editable too (legacy or pre-setup rows).
        for sid, sname, lab, folder_path in students:
            if sid in represented_students:
                continue
            base_key = (folder_path or "").strip() or f"DB:{sid}"
            folder_key = _unique_key(base_key)
            self.folder_order.append(folder_key)
            self.rows[folder_key] = {
                "include": False,
                "manual_include_override": None,
                "folder": base_key,
                "det_id": sid or "",
                "det_name": sname or "",
                "final_id": sid or "",
                "final_name": self._normalized_name(sname or "", base_key, sname or ""),
                "lab_id": lab or "",
                "files": [],
            }

        self._reload_tree_rows()

        self._set_scan_status(prefix="Loaded from DB")

    def scan(self):
        if not self.root_folder:
            messagebox.showinfo("Pick folder", "Choose ROOT folder first.")
            return

        existing_files = set()
        if self.only_new_files_var.get():
            existing_files = {r[0] for r in self.con.execute("SELECT file_path FROM files").fetchall()}

        file_globs = self._parse_globs()
        filename_regex = (self.filename_regex_var.get() or "").strip()
        folder_id_regex = (self.folder_id_regex_var.get() or "").strip()
        folder_name_regex = (self.folder_name_regex_var.get() or "").strip()

        scanner = DefaultFolderScanner(
            self.root_folder,
            file_globs=file_globs,
            include_filename_regex=filename_regex,
            folder_id_regex=folder_id_regex,
            folder_name_regex=folder_name_regex,
        )
        folders = scanner.collect_folders()

        existing_rows = {
            k: {
                "include": bool(v.get("include")),
                "manual_include_override": v.get("manual_include_override"),
                "folder": v.get("folder", k),
                "det_id": v.get("det_id", ""),
                "det_name": v.get("det_name", ""),
                "final_id": v.get("final_id", ""),
                "final_name": v.get("final_name", ""),
                "lab_id": v.get("lab_id", ""),
                "files": list(v.get("files") or []),
            }
            for k, v in self.rows.items()
        }
        existing_order = list(self.folder_order)

        self.rows.clear()
        self.folder_order.clear()

        for sub in folders:
            final_id, final_name, det_id, det_name, files = scanner.detect_folder(sub)
            if existing_files:
                files = [fp for fp in files if fp not in existing_files]
            folder_key = str(sub)
            self.folder_order.append(folder_key)
            is_student = has_required_student_fields(final_id or "", final_name or "")
            include_row = bool(files) and is_student
            self.rows[folder_key] = {
                "include": include_row,
                "manual_include_override": None,
                "folder": folder_key,
                "det_id": det_id or "",
                "det_name": det_name or "",
                "final_id": final_id or "",
                "final_name": self._normalized_name(final_name or "", str(sub), det_name or ""),
                "lab_id": "",
                "files": files,
            }

            prior = existing_rows.get(folder_key)
            if prior:
                self.rows[folder_key]["final_id"] = prior.get("final_id", self.rows[folder_key]["final_id"])
                self.rows[folder_key]["final_name"] = self._normalized_name(
                    prior.get("final_name", self.rows[folder_key]["final_name"]),
                    folder_key,
                    self.rows[folder_key].get("det_name", ""),
                )
                self.rows[folder_key]["lab_id"] = prior.get("lab_id", "")
                self.rows[folder_key]["manual_include_override"] = prior.get("manual_include_override")

                is_student = has_required_student_fields(
                    self.rows[folder_key]["final_id"],
                    self.rows[folder_key]["final_name"],
                )
                has_files = bool(self.rows[folder_key].get("files"))
                manual = self.rows[folder_key]["manual_include_override"]
                if not (is_student and has_files):
                    self.rows[folder_key]["include"] = False
                elif manual is False:
                    self.rows[folder_key]["include"] = False
                elif manual is True:
                    self.rows[folder_key]["include"] = True
                else:
                    self.rows[folder_key]["include"] = True

        for folder_key in existing_order:
            if folder_key in self.rows:
                continue
            self.folder_order.append(folder_key)
            self.rows[folder_key] = existing_rows[folder_key]

        self._reload_tree_rows()

        self._set_scan_status(prefix="Scan done")

    def on_folder_select(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        folder_key = sel[0]
        self.selected_folder_key = folder_key
        r = self.rows.get(folder_key)
        if not r:
            return

        self.sel_folder_var.set(folder_key)
        self.include_var.set(bool(r["include"]))
        self._suspend_auto_apply = True
        self.final_id_var.set(r["final_id"])
        self.final_name_var.set(r["final_name"])
        self.lab_id_var.set(r.get("lab_id",""))
        self._suspend_auto_apply = False

        for item in self.files_tree.get_children():
            self.files_tree.delete(item)
        for idx, fp in enumerate(r["files"]):
            self.files_tree.insert("", "end", iid=f"file-{idx}", values=(fp,))

        self.preview.delete("1.0", tk.END)

    def on_scan_file_select(self, _evt=None, file_iid: str | None = None):
        if file_iid is not None:
            sel = (file_iid,)
        else:
            sel = self.files_tree.selection()
        if not sel:
            return
        fp = self.files_tree.item(sel[0], "values")[0]
        try:
            content = Path(fp).read_text(encoding="utf-8", errors="ignore")
        except Exception as e:
            content = f"Error reading file:\n{e}"
        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", content)
        self.preview.tag_remove("find_hit", "1.0", tk.END)
        self._find_from = "1.0"

    def apply_include_toggle(self):
        folder_key = self.selected_folder_key or self.sel_folder_var.get().strip()
        if not folder_key or folder_key not in self.rows:
            return
        row = self.rows[folder_key]
        files = row.get("files") or []
        if not files:
            row["include"] = False
            row["manual_include_override"] = False
            self.include_var.set(False)
            self._refresh_tree_row(folder_key)
            self._reload_student_rows()
            self._set_scan_status(prefix="Scan info")
            return

        row["include"] = bool(self.include_var.get())
        row["manual_include_override"] = row["include"]
        self._refresh_tree_row(folder_key)
        self._reload_student_rows()
        self._set_scan_status(prefix="Scan info")

    def apply_edits(self):
        folder_key = self.selected_folder_key or self.sel_folder_var.get().strip()
        if not folder_key or folder_key not in self.rows:
            return

        raw_fid = self.final_id_var.get().strip()
        fid = "FULL" if raw_fid.lower() == "full" else extract_numeric_id(raw_fid)
        fname = self._normalized_name(self.final_name_var.get(), folder_key, self.rows[folder_key].get("det_name", ""))
        lab = self.lab_id_var.get().strip()

        self.rows[folder_key]["final_id"] = fid
        self.rows[folder_key]["final_name"] = fname
        self.rows[folder_key]["lab_id"] = lab

        has_files = bool(self.rows[folder_key].get("files"))
        manual_override = self.rows[folder_key].get("manual_include_override")

        if not has_files:
            self.rows[folder_key]["include"] = False
        elif manual_override is False:
            self.rows[folder_key]["include"] = False
        else:
            self.rows[folder_key]["include"] = True

        self._suspend_auto_apply = True
        self.include_var.set(bool(self.rows[folder_key]["include"]))
        self.final_name_var.set(fname)
        self._suspend_auto_apply = False

        self._refresh_tree_row(folder_key)
        self._reload_student_rows()
        self._set_scan_status(prefix="Scan info")

    def _refresh_tree_row(self, folder_key: str):
        r = self.rows[folder_key]
        self.tree.item(folder_key, values=(
            "YES" if r["include"] else "NO",
            Path(r["folder"]).name,
            r["det_id"],
            r["det_name"],
            r["final_id"],
            r["final_name"],
            r.get("lab_id",""),
            str(len(r["files"]))
        ))

    def apply_global_lab_id(self):
        lab = (self.global_lab_id_var.get() or "").strip()
        for folder_key in self.folder_order:
            r = self.rows.get(folder_key)
            if not r or not r.get("include"):
                continue
            r["lab_id"] = lab
            self._refresh_tree_row(folder_key)

        current = self.selected_folder_key or self.sel_folder_var.get().strip()
        if current and current in self.rows:
            self._suspend_auto_apply = True
            self.lab_id_var.set(self.rows[current].get("lab_id", ""))
            self._suspend_auto_apply = False

        self._reload_student_rows()
        self._set_scan_status(prefix="Scan info")

    def _selected_preview_text(self) -> str:
        try:
            return self.preview.get("sel.first", "sel.last").strip()
        except tk.TclError:
            return ""

    def _hotkey_use_id(self, _evt=None):
        self.use_selection_as_id()
        return "break"

    def _hotkey_use_name(self, _evt=None):
        self.use_selection_as_name()
        return "break"

    def use_selection_as_id(self):
        selected = self._selected_preview_text()
        if not selected:
            return
        # Prefer numeric extraction, but still allow full raw selection so
        # quick-paste IDs such as NAME:... or custom keys can be used.
        only_id = extract_numeric_id(selected)
        self.final_id_var.set(only_id if only_id else selected)

    def _selected_folder_from_tree(self) -> str | None:
        sel = self.tree.selection()
        if not sel:
            return None
        folder_key = sel[0]
        return folder_key if folder_key in self.rows else None

    def _build_skimmable_sequence(self, start_folder_key: str | None) -> list[str]:
        folders = self._get_skimmable_folders()
        if not folders:
            return []
        if not start_folder_key or start_folder_key not in folders:
            return folders

        start_idx = folders.index(start_folder_key)
        return folders[start_idx:] + folders[:start_idx]

    def use_selection_as_name(self):
        selected = self._selected_preview_text()
        if not selected:
            return
        self.final_name_var.set(selected)

    def focus_find_entry(self):
        self.find_entry.focus_set()
        return "break"

    def find_next(self):
        query = (self.find_var.get() or "").strip()
        self.preview.tag_remove("find_hit", "1.0", tk.END)
        if not query:
            return

        start = self.preview.search(query, self._find_from, stopindex=tk.END, nocase=True)
        if not start:
            start = self.preview.search(query, "1.0", stopindex=tk.END, nocase=True)
            if not start:
                return

        end = f"{start}+{len(query)}c"
        self.preview.tag_add("find_hit", start, end)
        self.preview.mark_set(tk.INSERT, end)
        self.preview.see(start)
        self.preview.focus_set()
        self._find_from = end

    def _get_skimmable_folders(self) -> list[str]:
        """
        Skim every folder that has at least one file so names/IDs can be verified quickly.
        """
        out = []
        for folder_key in self.folder_order:
            row = self.rows.get(folder_key) or {}
            if not (row.get("files") or []):
                continue
            out.append(folder_key)
        return out

    def _select_folder_by_index(self, idx: int):
        if idx < 0 or idx >= len(self.folder_order):
            return None
        folder_key = self.folder_order[idx]
        self.tree.selection_set(folder_key)
        self.tree.focus(folder_key)
        self.tree.see(folder_key)
        self.on_folder_select()
        return folder_key

    def _reload_tree_rows(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for folder_key in self.folder_order:
            r = self.rows[folder_key]
            self.tree.insert("", "end", iid=folder_key, values=(
                "YES" if r["include"] else "NO",
                Path(r["folder"]).name,
                r["det_id"],
                r["det_name"],
                r["final_id"],
                r["final_name"],
                r.get("lab_id", ""),
                str(len(r["files"]))
            ))

        self._reload_student_rows()

    def _reload_student_rows(self):
        for item in self.student_tree.get_children():
            self.student_tree.delete(item)

        idx = 0
        for folder_key in self.folder_order:
            r = self.rows.get(folder_key) or {}
            if not has_required_student_fields(r.get("final_id", ""), r.get("final_name", "")):
                continue
            idx += 1
            self.student_tree.insert("", "end", iid=f"student-{idx}", values=(
                r.get("final_id", ""),
                r.get("final_name", ""),
                r.get("lab_id", ""),
                Path(r.get("folder", folder_key)).name,
                str(len(r.get("files") or [])),
            ))

    def _select_file_in_current_folder(self, file_idx: int) -> bool:
        self.update_idletasks()
        children = self.files_tree.get_children()
        size = len(children)
        if file_idx < 0 or file_idx >= size:
            return False
        iid = children[file_idx]
        self.files_tree.selection_set(iid)
        self.files_tree.focus(iid)
        self.files_tree.see(iid)
        self.files_tree.focus_set()
        self.update_idletasks()
        self.files_tree.event_generate("<<TreeviewSelect>>")
        self.on_scan_file_select(file_iid=iid)
        return True

    def start_skimming(self):
        if not self.folder_order:
            if self.root_folder:
                self.scan()
            else:
                self.load_existing_rows_from_db()
        if not self.folder_order:
            messagebox.showinfo("Skimming", "No rows loaded. Scan folders or load from DB first.")
            return

        start_folder_key = self._selected_folder_from_tree()
        self._skimmable_folder_keys = self._build_skimmable_sequence(start_folder_key)
        if not self._skimmable_folder_keys:
            messagebox.showinfo("Skimming", "No folders with files to skim.")
            return

        try:
            delay = int(self.skim_delay_ms_var.get())
        except Exception:
            delay = 300
        if delay < 100:
            delay = 100
        self.skim_delay_ms_var.set(delay)

        self._skim_folder_idx = 0
        self._skim_file_idx = 0
        self.skim_running = True

        first_folder_key = self._skimmable_folder_keys[0]
        try:
            first_folder_idx = self.folder_order.index(first_folder_key)
            self._select_folder_by_index(first_folder_idx)
        except Exception:
            pass

        self._set_scan_status(prefix="Skimming started")
        self._skim_step()

    def stop_skimming(self):
        if self.skim_running:
            self.skim_running = False
            self._set_scan_status(prefix="Skimming stopped")
        return "break"

    def _skim_step(self):
        if not self.skim_running:
            return

        if self._skim_folder_idx >= len(self._skimmable_folder_keys):
            self.skim_running = False
            self._set_scan_status(prefix="Skimming done")
            return

        folder_key = self._skimmable_folder_keys[self._skim_folder_idx]
        files = self.rows.get(folder_key, {}).get("files") or []
        if not files:
            self._skim_folder_idx += 1
            self._skim_file_idx = 0
            self.after(int(self.skim_delay_ms_var.get()), self._skim_step)
            return

        try:
            folder_idx = self.folder_order.index(folder_key)
        except ValueError:
            self._skim_folder_idx += 1
            self._skim_file_idx = 0
            self.after(int(self.skim_delay_ms_var.get()), self._skim_step)
            return

        current_selected = self._selected_folder_from_tree()
        if current_selected != folder_key:
            if not self._select_folder_by_index(folder_idx):
                self._skim_folder_idx += 1
                self._skim_file_idx = 0
                self.after(int(self.skim_delay_ms_var.get()), self._skim_step)
                return

        if self._skim_file_idx >= len(files):
            self._skim_folder_idx += 1
            self._skim_file_idx = 0
            self.after(int(self.skim_delay_ms_var.get()), self._skim_step)
            return

        selected = self._select_file_in_current_folder(self._skim_file_idx)
        if selected:
            self._set_scan_status(prefix=f"Skimming: student {self._skim_folder_idx + 1}/{len(self._skimmable_folder_keys)}, file {self._skim_file_idx + 1}/{len(files)}")

        self._skim_file_idx += 1
        self.after(int(self.skim_delay_ms_var.get()), self._skim_step)

    def save_to_db(self):
        if not self.rows:
            self.load_existing_rows_from_db()
        if not self.rows:
            messagebox.showinfo("Nothing", "No rows loaded. Scan folders or load from DB first.")
            return

        store_content = self.store_file_content.get()

        committed_folders = 0
        committed_files = 0
        created_students = 0
        skipped_folders = 0

        for folder_key in self.folder_order:
            r = self.rows[folder_key]
            if not r.get("files"):
                skipped_folders += 1
                continue

            include_row = bool(r.get("include"))
            fid = (r["final_id"] or "").strip()
            fname = (r["final_name"] or "").strip()
            lab = (r.get("lab_id") or "").strip() or None

            student_ok = has_required_student_fields(fid, fname)
            normalized_sid = (fid or "").strip()
            if normalized_sid.lower() != "full":
                normalized_sid = extract_numeric_id(normalized_sid)

            if include_row and student_ok and normalized_sid:
                upsert_student(self.con, normalized_sid, fname, lab, r.get("folder") or folder_key)
                created_students += 1
            elif include_row and student_ok and (fid or "").strip().lower() == "full":
                upsert_student(self.con, "FULL", fname or "FULL", lab, r.get("folder") or folder_key)
                normalized_sid = "FULL"
                created_students += 1
            elif not include_row:
                skipped_folders += 1

            for fp in r["files"]:
                p = Path(fp)
                file_text = None
                file_hash = None
                try:
                    if store_content:
                        file_text = read_file_text(p)
                        file_hash = sha256_text(file_text)
                except Exception:
                    file_text = None
                    file_hash = None

                upsert_file(
                    self.con,
                    file_path=fp,
                    student_id=normalized_sid if include_row and student_ok and normalized_sid else None,
                    source_folder=r.get("folder") or folder_key,
                    detected_id=fid or r["det_id"] or None,
                    detected_name=fname or r["det_name"] or None,
                    file_hash=file_hash,
                    file_content=file_text
                )
                committed_files += 1

            committed_folders += 1

        self.app.refresh_students(keep_selected=False)
        messagebox.showinfo(
            "Saved",
            f"Folders committed: {committed_folders}\nFolders skipped: {skipped_folders}\nFiles saved: {committed_files}\nStudents created/updated: {created_students}"
        )


# =============================================================================
# 12) Main App
# =============================================================================

def clamp_points(val: float | None, mx: float) -> float | None:
    if val is None:
        return None
    if val < 0:
        return 0.0
    if val > mx:
        return mx
    return val

class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Java Grading Gallery — v2.2 (Assignments + Comments + PDF + Pastel + LabID + Histogram)")
        self.root.geometry("1920x1040")

        self.palette = pastel_style(root)

        self.sub_db_path = Path(SUBMISSIONS_DB)
        self.sub_con = db_connect(self.sub_db_path)
        submissions_db_init(self.sub_con)
        grading_db_init(self.sub_con)

        # Grading data is stored in the same submissions DB.
        self.grade_db_path: Path | None = self.sub_db_path
        self.grade_con: sqlite3.Connection | None = self.sub_con

        self.student_ids: list[str] = []
        self.selected_student_id: str | None = None
        self.selected_file_path: str | None = None

        self.question_map: dict[str, str] = {}
        self.selected_question_id: str | None = None

        # curve preview factor (histogram overlay)
        self.curve_preview_var = tk.DoubleVar(value=1.0)

        # Optional auto-grader
        self.auto_grader = AutoGrader(enabled=DEFAULT_AI_ENABLED)

        self._grade_last_student_selection: tuple[int, ...] = ()
        self._grade_last_file_selection: tuple[int, ...] = ()

        self._build_ui()
        self.refresh_students(keep_selected=False)
        self.root.after(120, self._poll_grade_selections)

    def _poll_grade_selections(self):
        student_sel = tuple(self.student_list.curselection())
        if student_sel != self._grade_last_student_selection:
            self._grade_last_student_selection = student_sel
            if student_sel:
                self.on_student_select()

        file_sel = tuple(self.file_list.curselection())
        if file_sel != self._grade_last_file_selection:
            self._grade_last_file_selection = file_sel
            if file_sel:
                self.on_grade_file_select()

        self.root.after(120, self._poll_grade_selections)

    def require_grading_db(self) -> bool:
        if self.grade_con is None:
            self.grade_con = self.sub_con
            self.grade_db_path = self.sub_db_path
            grading_db_init(self.grade_con)
        return True

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=10, style="Pastel.TFrame")
        top.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(top, text="Open Scan / Edit", command=self.open_scan_window).pack(side=tk.LEFT)

        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(top, text="New Submissions DB...", command=self.new_submissions_db).pack(side=tk.LEFT)
        ttk.Button(top, text="Open Submissions DB...", command=self.open_submissions_db).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Save Submissions DB As...", command=self.save_submissions_db_as).pack(side=tk.LEFT, padx=6)

        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(top, text="Load Scheme CSV...", command=self.load_scheme_csv).pack(side=tk.LEFT)
        ttk.Button(top, text="Reload Last Scheme", command=self.reload_last_scheme).pack(side=tk.LEFT, padx=6)

        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        self.sub_db_status_lbl = ttk.Label(top, text=f"Submissions DB: {self.sub_db_path.name}", style="Pastel.TLabel")
        self.sub_db_status_lbl.pack(side=tk.LEFT, padx=10)

        self.db_status_lbl = ttk.Label(top, text=f"Grading: integrated with {self.sub_db_path.name}", style="Pastel.TLabel")
        self.db_status_lbl.pack(side=tk.LEFT, padx=10)

        self.student_count_lbl = ttk.Label(top, text="Students: 0", style="Pastel.TLabel")
        self.student_count_lbl.pack(side=tk.RIGHT)

        # Histogram + stats area on top right
        stats_frame = ttk.Frame(top, style="Pastel.TFrame")
        stats_frame.pack(side=tk.RIGHT, padx=14)

        self.class_stats_lbl = ttk.Label(stats_frame, text="Class: avg - | min - | max - | curve -", style="Pastel.TLabel")
        self.class_stats_lbl.pack(side=tk.TOP, anchor="e")

        curve_row = ttk.Frame(stats_frame, style="Pastel.TFrame")
        curve_row.pack(side=tk.TOP, anchor="e", pady=(2, 0))
        ttk.Label(curve_row, text="Curve preview ×", style="Pastel.TLabel").pack(side=tk.LEFT)
        self.curve_scale = ttk.Scale(curve_row, from_=0.8, to=1.3, variable=self.curve_preview_var, command=lambda _v: self.refresh_summary())
        self.curve_scale.pack(side=tk.LEFT, padx=(6, 0))

        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill=tk.BOTH, expand=True)

        self.tab_grade = ttk.Frame(self.nb, style="Pastel.TFrame", padding=10)
        self.tab_summary = ttk.Frame(self.nb, style="Pastel.TFrame", padding=10)
        self.tab_stats = ttk.Frame(self.nb, style="Pastel.TFrame", padding=10)

        self.nb.add(self.tab_grade, text="Grade")
        self.nb.add(self.tab_summary, text="Summary")
        self.nb.add(self.tab_stats, text="Stats")

        self._build_grade_tab()
        self._build_summary_tab()
        self._build_stats_tab()

    def _build_grade_tab(self):
        main = ttk.Frame(self.tab_grade, style="Pastel.TFrame")
        main.pack(fill=tk.BOTH, expand=True)

        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=2)
        main.columnconfigure(2, weight=2)
        main.rowconfigure(0, weight=1)

        # LEFT
        left = ttk.Frame(main, style="PastelCard.TFrame", padding=10)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.rowconfigure(3, weight=1)
        left.columnconfigure(0, weight=1)

        ttk.Label(left, text="Students", style="PastelCard.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.student_list = tk.Listbox(left, bg=self.palette["panel"], fg=self.palette["text"], highlightthickness=0, selectbackground=self.palette["select"])
        self.student_list.grid(row=1, column=0, sticky="nsew")

        ttk.Label(left, text="Files", style="PastelCard.TLabel", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.file_list = tk.Listbox(left, height=8, bg=self.palette["panel"], fg=self.palette["text"], highlightthickness=0, selectbackground=self.palette["select"])
        self.file_list.grid(row=3, column=0, sticky="nsew")
        self.file_list.bind("<<ListboxSelect>>", self.on_grade_file_select)

        # MIDDLE preview + comments
        mid = ttk.Frame(main, style="PastelCard.TFrame", padding=10)
        mid.grid(row=0, column=1, sticky="nsew", padx=(0, 10))
        mid.rowconfigure(2, weight=1)
        mid.columnconfigure(0, weight=1)

        self.student_header = ttk.Label(mid, text="No student selected", style="PastelCard.TLabel", font=("Segoe UI", 12, "bold"))
        self.student_header.grid(row=0, column=0, sticky="w")

        codebar = ttk.Frame(mid, style="PastelCard.TFrame")
        codebar.grid(row=1, column=0, sticky="ew", pady=(6, 6))
        ttk.Button(codebar, text="Add comment to selection", command=self.add_comment_to_selection).pack(side=tk.LEFT)
        ttk.Button(codebar, text="Clear comments in selection", command=self.clear_comments_in_selection).pack(side=tk.LEFT, padx=6)
        ttk.Button(codebar, text="Compare to FULL", command=self.compare_to_full).pack(side=tk.LEFT, padx=6)
        ttk.Button(codebar, text="Export All (Compare-to-FULL naming)", command=self.export_compare_to_full_pdfs).pack(side=tk.LEFT, padx=6)
        ttk.Button(codebar, text="Export PDF (this student)", command=self.export_student_pdf).pack(side=tk.RIGHT)

        preview_frame = ttk.Frame(mid, style="PastelCard.TFrame")
        preview_frame.grid(row=2, column=0, sticky="nsew")
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)

        self.preview = tk.Text(preview_frame, wrap="none",
                               bg="#FFFDF7", fg=self.palette["text"],
                               insertbackground=self.palette["text"],
                               selectbackground=self.palette["select"],
                               highlightthickness=1, highlightbackground="#E8E1FF")
        self.preview.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.preview.configure(yscrollcommand=sb.set)

        self.preview.tag_configure("comment_highlight", background="#FFF9A6")

        # RIGHT grading
        right = ttk.Frame(main, style="PastelCard.TFrame", padding=10)
        right.grid(row=0, column=2, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(14, weight=1)

        ttk.Label(right, text="Question (from loaded scheme)", style="PastelCard.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.question_var = tk.StringVar(value="")
        self.question_combo = ttk.Combobox(right, textvariable=self.question_var, state="readonly")
        self.question_combo.grid(row=1, column=0, sticky="ew", pady=(4,8))
        self.question_combo.bind("<<ComboboxSelected>>", self.on_question_select)

        ttk.Separator(right, orient="horizontal").grid(row=2, column=0, sticky="ew", pady=(8,10))

        ttk.Label(right, text="Theme / Instructions (saved in DB)", style="PastelCard.TLabel").grid(row=5, column=0, sticky="w", pady=(10, 0))
        self.theme_text = tk.Text(right, height=4,
                                  bg="#FFFDF7", fg=self.palette["text"],
                                  insertbackground=self.palette["text"],
                                  highlightthickness=1, highlightbackground="#E8E1FF")
        self.theme_text.grid(row=6, column=0, sticky="nsew")
        self.theme_text.insert("1.0", DEFAULT_THEME)

        btns = ttk.Frame(right, style="PastelCard.TFrame")
        btns.grid(row=7, column=0, sticky="ew", pady=(8, 8))

        ttk.Button(btns, text="Save Theme", command=self.save_theme).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(btns, text="Auto Grade (optional)", command=self.auto_grade_optional).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        ttk.Button(btns, text="Save (this question)", command=self.save_scores_and_rationale).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        ttk.Button(btns, text="Autofill/Test (assigned × all students)", command=self.make_all_assigned).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        ttk.Button(btns, text="Export Excel (all)", command=self.save_all_excel).pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)

        ttk.Label(right, text="Rationale (per student, per question)", style="PastelCard.TLabel").grid(row=8, column=0, sticky="w")
        self.rationale_text = tk.Text(right, height=6,
                                      bg="#FFFDF7", fg=self.palette["text"],
                                      insertbackground=self.palette["text"],
                                      highlightthickness=1, highlightbackground="#E8E1FF")
        self.rationale_text.grid(row=9, column=0, sticky="nsew")

        self.total_lbl = ttk.Label(right, text="Total: -", style="PastelCard.TLabel", font=("Segoe UI", 11, "bold"))
        self.total_lbl.grid(row=10, column=0, sticky="w", pady=(6, 6))

        ttk.Label(right, text="Rubric Table (selected question)", style="PastelCard.TLabel").grid(row=11, column=0, sticky="w")
        self.rubric_grid = ScrollableRubricGrid(right)
        self.rubric_grid.grid(row=12, column=0, sticky="nsew")

        ttk.Label(right, text="Code comments (this file)", style="PastelCard.TLabel").grid(row=13, column=0, sticky="w", pady=(10,0))
        self.comment_list = tk.Listbox(right, height=7, bg=self.palette["panel"], fg=self.palette["text"],
                                       highlightthickness=0, selectbackground=self.palette["select"])
        self.comment_list.grid(row=14, column=0, sticky="nsew")

    def _build_summary_tab(self):
        top = ttk.Frame(self.tab_summary, style="Pastel.TFrame")
        top.pack(fill=tk.X)

        ttk.Button(top, text="Refresh Summary", command=self.refresh_summary).pack(side=tk.LEFT)
        ttk.Button(top, text="Export Grade (selected)", command=self.export_selected_excel).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Export Grades (all students)", command=self.save_all_excel).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Export PDF (selected)", command=self.export_student_pdf).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Export PDFs (all students)", command=self.export_all_students_pdfs).pack(side=tk.LEFT, padx=6)

        cols = ("student_id", "student_name", "lab_id", "graded_questions", "overall", "overall_curved")
        self.sum_tree = ttk.Treeview(self.tab_summary, columns=cols, show="headings", height=25)
        for c, w in [("student_id", 140), ("student_name", 220), ("lab_id", 80),
                     ("graded_questions", 130), ("overall", 90), ("overall_curved", 120)]:
            self.sum_tree.heading(c, text=c)
            self.sum_tree.column(c, width=w, anchor="w")
        self.sum_tree.pack(fill=tk.BOTH, expand=True, pady=(10,0))

        note = ttk.Label(self.tab_summary,
                         text="Tip: the special FULL student is used as a model answer reference and excluded from student grading lists.",
                         style="Pastel.TLabel")
        note.pack(anchor="w", pady=(8,0))

    def _build_stats_tab(self):
        ttk.Label(self.tab_stats, text="Class Histogram (raw + curved preview)", style="Pastel.TLabel",
                  font=("Segoe UI", 12, "bold")).pack(anchor="w")

        self.hist_frame = ttk.Frame(self.tab_stats, style="Pastel.TFrame")
        self.hist_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        if FigureCanvasTkAgg is None or Figure is None:
            ttk.Label(self.hist_frame, text="Matplotlib not available. Install: pip install matplotlib",
                      style="Pastel.TLabel").pack(anchor="w")
            self.hist_canvas = None
            self.hist_fig = None
            self.hist_ax = None
            return

        self.hist_fig = Figure(figsize=(8, 4.2), dpi=110)
        self.hist_ax = self.hist_fig.add_subplot(111)
        self.hist_canvas = FigureCanvasTkAgg(self.hist_fig, master=self.hist_frame)
        self.hist_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        ttk.Label(self.tab_stats,
                  text="Note: Curve preview uses multiplication factor (clipped at 0). It does not change stored grades unless you export/transform them externally.",
                  style="Pastel.TLabel").pack(anchor="w", pady=(8, 0))

    # ---- scan window ----
    def open_scan_window(self):
        return ScanWindow(self)

    # ---- submissions DB open/create/save ----
    def new_submissions_db(self):
        path = filedialog.asksaveasfilename(
            title="Create submissions DB",
            defaultextension=".sqlite",
            filetypes=[("SQLite DB", "*.sqlite"), ("All files", "*.*")]
        )
        if not path:
            return
        self._open_submissions_db(Path(path))

    def open_submissions_db(self):
        path = filedialog.askopenfilename(
            title="Open submissions DB",
            filetypes=[("SQLite DB", "*.sqlite"), ("All files", "*.*")]
        )
        if not path:
            return
        self._open_submissions_db(Path(path))

    def save_submissions_db_as(self):
        path = filedialog.asksaveasfilename(
            title="Save submissions DB as",
            defaultextension=".sqlite",
            filetypes=[("SQLite DB", "*.sqlite"), ("All files", "*.*")]
        )
        if not path:
            return

        target = Path(path)
        try:
            out_con = db_connect(target)
            self.sub_con.backup(out_con)
            out_con.close()
        except Exception as e:
            messagebox.showerror("Save failed", str(e))
            return

        self._open_submissions_db(target)
        messagebox.showinfo("Saved", f"Submissions DB saved to:\n{target}")

    def _open_submissions_db(self, path: Path):
        if self.sub_con is not None:
            try:
                self.sub_con.close()
            except Exception:
                pass

        self.sub_db_path = path
        self.sub_con = db_connect(path)
        submissions_db_init(self.sub_con)
        grading_db_init(self.sub_con)
        self.grade_con = self.sub_con
        self.grade_db_path = self.sub_db_path

        self.sub_db_status_lbl.config(text=f"Submissions DB: {path.name}")
        self.db_status_lbl.config(text=f"Grading: integrated with {path.name}")

        theme = meta_get(self.grade_con, "theme", DEFAULT_THEME)
        self.theme_text.delete("1.0", tk.END)
        self.theme_text.insert("1.0", theme)
        self.refresh_question_lists()

        self.selected_student_id = None
        self.selected_file_path = None
        self.refresh_students(keep_selected=False)
        self.refresh_summary()
        self.preview.delete("1.0", tk.END)
        self.file_list.delete(0, tk.END)
        self.comment_list.delete(0, tk.END)

    # ---- grading DB open/create ----
    def new_grading_db(self):
        messagebox.showinfo("Integrated DB", "Grading data is stored in the current submissions DB.")

    def open_grading_db(self):
        messagebox.showinfo("Integrated DB", "Grading data is stored in the current submissions DB.")

    def _open_grade_db(self, path: Path):
        messagebox.showinfo("Integrated DB", "Grading data is stored in the current submissions DB.")

    # ---- Scheme CSV loading ----
    def load_scheme_csv(self):
        if not self.require_grading_db():
            return
        path = filedialog.askopenfilename(
            title="Load scheme CSV",
            filetypes=[("CSV", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            load_scheme_csv_into_db(self.grade_con, Path(path))
        except Exception as e:
            messagebox.showerror("CSV load failed", str(e))
            return

        self.refresh_question_lists()
        messagebox.showinfo("Loaded", f"Scheme loaded from:\n{path}")

    def reload_last_scheme(self):
        if not self.require_grading_db():
            return
        p = meta_get(self.grade_con, "scheme_csv_path", "")
        if not p or not Path(p).exists():
            messagebox.showinfo("No saved scheme", "No last scheme path saved (or file missing). Load Scheme CSV first.")
            return
        try:
            load_scheme_csv_into_db(self.grade_con, Path(p))
        except Exception as e:
            messagebox.showerror("Reload failed", str(e))
            return
        self.refresh_question_lists()
        messagebox.showinfo("Reloaded", f"Reloaded:\n{p}")

    def refresh_question_lists(self):
        if self.grade_con is None:
            return
        qs = fetch_questions(self.grade_con)
        self.question_map = {qid: (f"{title} [{sub_id}]" if sub_id else title) for qid, title, sub_id in qs}
        self.refresh_question_picker_for_student()

    def refresh_question_picker_for_student(self):
        if self.grade_con is None:
            return

        allowed = list(self.question_map.keys())
        items = [f"{qid} — {self.question_map[qid]}" for qid in allowed if qid in self.question_map]
        self.question_combo["values"] = items

        if not items:
            self.selected_question_id = None
            self.question_var.set("")
            return

        if not self.selected_question_id or self.selected_question_id not in self.question_map:
            self.selected_question_id = items[0].split(" — ", 1)[0].strip()
            self.question_var.set(items[0])
            try:
                self.question_combo.current(0)
            except Exception:
                pass

    def on_question_select(self, _evt=None):
        v = self.question_var.get().strip()
        if " — " in v:
            self.selected_question_id = v.split(" — ", 1)[0].strip()
        self.load_student_question_view()

    # ---- theme ----
    def save_theme(self):
        if not self.require_grading_db():
            return
        theme = self.theme_text.get("1.0", tk.END).strip()
        meta_set(self.grade_con, "theme", theme)
        messagebox.showinfo("Saved", "Theme saved.")

    # ---- students ----
    def refresh_students(self, keep_selected: bool = True):
        selected_id = self.selected_student_id if keep_selected else None

        rows = get_students(self.sub_con)
        self.student_list.delete(0, tk.END)
        self.student_ids = []

        for sid, name, lab, file_count in rows:
            if (sid or "").strip().lower() == "full":
                continue
            if not has_required_student_fields(sid, name):
                continue
            lab_txt = f" | Lab:{lab}" if lab else ""
            self.student_list.insert(tk.END, f"{sid} — {name}{lab_txt} | files: {file_count}")
            self.student_ids.append(sid)

        self.student_count_lbl.config(text=f"Students: {len(self.student_ids)}")

        if keep_selected and selected_id and selected_id in self.student_ids:
            idx = self.student_ids.index(selected_id)
            self.student_list.selection_clear(0, tk.END)
            self.student_list.selection_set(idx)
            self.student_list.activate(idx)
            self.student_list.see(idx)

    def on_student_select(self, _evt=None):
        sel = self.student_list.curselection()
        if not sel:
            return
        sid = self.student_ids[sel[0]]
        self.selected_student_id = sid

        row = self.sub_con.execute("SELECT student_name, COALESCE(lab_id,'') FROM students WHERE student_id=?", (sid,)).fetchone()
        name = row[0] if row else ""
        lab = row[1] if row else ""
        lab_txt = f" | Lab:{lab}" if lab else ""
        self.student_header.config(text=f"{sid} — {name}{lab_txt}")

        files = get_student_files(self.sub_con, sid)
        self.file_list.delete(0, tk.END)
        for f in files:
            self.file_list.insert(tk.END, f)

        self.preview.delete("1.0", tk.END)
        self.selected_file_path = None
        self.comment_list.delete(0, tk.END)

        if self.grade_con is not None:
            self.refresh_question_picker_for_student()

        self.load_student_question_view()
        self.refresh_summary()

    def on_grade_file_select(self, _evt=None):
        sel = self.file_list.curselection()
        if not sel:
            return
        fp = self.file_list.get(sel[0])
        self.selected_file_path = fp

        content = get_file_content(self.sub_con, fp)
        if content is None:
            try:
                content = Path(fp).read_text(encoding="utf-8", errors="ignore")
            except Exception as e:
                content = f"Error reading file:\n{e}"

        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", content)
        self._apply_comments_highlights()

    # ---- code comments ----
    def _apply_comments_highlights(self):
        self.preview.tag_remove("comment_highlight", "1.0", tk.END)
        self.comment_list.delete(0, tk.END)

        if self.grade_con is None:
            return
        if not self.selected_student_id or not self.selected_file_path:
            return

        comments = fetch_code_comments_for_file(self.grade_con, self.selected_student_id, self.selected_file_path)
        for cid, sidx, eidx, text, color, _created_at in comments:
            tag = "comment_highlight"
            if color:
                try:
                    self.preview.tag_configure(tag, background=color)
                except Exception:
                    pass
            try:
                self.preview.tag_add(tag, sidx, eidx)
            except Exception:
                pass
            self.comment_list.insert(tk.END, f"#{cid} {sidx}–{eidx}: {text}")

    def add_comment_to_selection(self):
        if not self.require_grading_db():
            return
        if not self.selected_student_id or not self.selected_file_path:
            messagebox.showinfo("Select", "Select a student and a file first.")
            return
        try:
            sidx = self.preview.index("sel.first")
            eidx = self.preview.index("sel.last")
        except Exception:
            messagebox.showinfo("Select", "Highlight a region of code first.")
            return

        txt = simpledialog.askstring("Add comment", "Comment for highlighted code:")
        if not txt:
            return
        add_code_comment(self.grade_con, self.selected_student_id, self.selected_file_path, sidx, eidx, txt, color="#FFF9A6")
        self._apply_comments_highlights()

    def clear_comments_in_selection(self):
        if not self.require_grading_db():
            return
        if not self.selected_student_id or not self.selected_file_path:
            return
        try:
            sidx = self.preview.index("sel.first")
            eidx = self.preview.index("sel.last")
        except Exception:
            messagebox.showinfo("Select", "Highlight a region first (selection).")
            return
        delete_code_comments_in_range(self.grade_con, self.selected_student_id, self.selected_file_path, sidx, eidx)
        self._apply_comments_highlights()

    def compare_to_full(self):
        if not self.selected_student_id:
            return
        full_id = "FULL"
        full_exists = self.sub_con.execute(
            "SELECT 1 FROM students WHERE LOWER(student_id)='full' LIMIT 1"
        ).fetchone()
        if not full_exists:
            messagebox.showinfo("No FULL", "No student with ID 'FULL' exists in submissions DB.")
            return

        import difflib
        a = merge_student_code(self.sub_con, full_id).splitlines()
        b = merge_student_code(self.sub_con, self.selected_student_id).splitlines()
        diff = difflib.unified_diff(a, b, fromfile="FULL", tofile=self.selected_student_id, lineterm="")
        text = "\n".join(diff) or "(No differences detected.)"

        win = tk.Toplevel(self.root)
        win.title(f"Diff: FULL vs {self.selected_student_id}")
        win.geometry("1400x800")
        t = tk.Text(win, wrap="none", bg="#0f111a", fg="#e6e6e6", insertbackground="#e6e6e6")
        t.pack(fill=tk.BOTH, expand=True)
        t.insert("1.0", text)

    # ---- load rubric table for selected student + question ----
    def load_student_question_view(self):
        self.rationale_text.delete("1.0", tk.END)
        self.total_lbl.config(text="Total: -")
        self.rubric_grid.build([])

        if self.grade_con is None:
            return
        if not self.selected_student_id:
            return
        if not self.selected_question_id:
            self.refresh_question_picker_for_student()
            return
        if self.selected_question_id not in self.question_map:
            return

        cols = fetch_columns_for_question(self.grade_con, self.selected_question_id)
        self.rubric_grid.build(cols)

        score_map, note_map = load_student_scores(self.grade_con, self.selected_student_id, self.selected_question_id)
        self.rubric_grid.set_values(score_map, note_map)

        note_row = load_student_note(self.grade_con, self.selected_student_id, self.selected_question_id)
        if note_row and note_row[0]:
            self.rationale_text.insert("1.0", note_row[0])

        total = compute_total(self.grade_con, self.selected_student_id, self.selected_question_id)
        self.total_lbl.config(text=f"Total: {total:g}")

    # ---- save ----
    def save_scores_and_rationale(self):
        if not self.require_grading_db():
            return
        if not self.selected_student_id or not self.selected_question_id:
            messagebox.showinfo("Missing", "Select a student and a question first.")
            return

        cols = fetch_columns_for_question(self.grade_con, self.selected_question_id)
        max_map = {k: float(mx) for (k, _g, _t, mx) in cols}

        score_raw, note_raw = self.rubric_grid.get_values()

        for col_key, raw in score_raw.items():
            raw = (raw or "").strip()
            if raw == "":
                points = None
            else:
                try:
                    points = float(raw)
                except ValueError:
                    messagebox.showerror("Invalid score", f"Score must be numeric or blank.\nBad value for:\n{col_key}")
                    return
                points = clamp_points(points, max_map.get(col_key, points))

            upsert_score(self.grade_con, self.selected_student_id, self.selected_question_id, col_key, points, note_raw.get(col_key, ""))

        rationale = self.rationale_text.get("1.0", tk.END).strip()
        total = compute_total(self.grade_con, self.selected_student_id, self.selected_question_id)
        upsert_student_note(self.grade_con, self.selected_student_id, self.selected_question_id, rationale, overall_grade=total)

        self.total_lbl.config(text=f"Total: {total:g}")
        self.refresh_summary()
        messagebox.showinfo("Saved", "Saved scores + rationale for this question.")

    # ---- Optional auto grade (separate component) ----
    def auto_grade_optional(self):
        if not self.require_grading_db():
            return
        if not self.selected_student_id or not self.selected_question_id:
            messagebox.showinfo("Missing", "Select a student and a question first.")
            return
        if not self.auto_grader.enabled:
            messagebox.showinfo("Disabled", "Auto grading is disabled (optional component).")
            return

        cols = fetch_columns_for_question(self.grade_con, self.selected_question_id)
        rubric_items = [{"col_key": col_key, "group": (group or ""), "criterion": text, "max_points": float(mx)}
                        for col_key, group, text, mx in cols]
        merged_code = merge_student_code(self.sub_con, self.selected_student_id)
        theme = self.theme_text.get("1.0", tk.END).strip()

        try:
            res = self.auto_grader.auto_grade(merged_code=merged_code, rubric_items=rubric_items, theme_text=theme)
        except Exception as e:
            messagebox.showerror("Auto grade failed", str(e))
            return

        # Apply
        score_map = {x["col_key"]: x.get("points", 0.0) for x in res.get("scores", []) if "col_key" in x}
        note_map = {x["col_key"]: x.get("note", "") for x in res.get("scores", []) if "col_key" in x}

        max_map = {k: float(mx) for (k, _g, _t, mx) in cols}
        for col_key in max_map.keys():
            pts = float(score_map.get(col_key, 0.0))
            pts = clamp_points(pts, max_map[col_key])
            upsert_score(self.grade_con, self.selected_student_id, self.selected_question_id, col_key, pts, note_map.get(col_key, ""))

        rationale = (res.get("rationale") or "").strip()
        total = compute_total(self.grade_con, self.selected_student_id, self.selected_question_id)
        upsert_student_note(self.grade_con, self.selected_student_id, self.selected_question_id, rationale, overall_grade=total)

        self.load_student_question_view()
        self.refresh_summary()
        messagebox.showinfo("Auto grade", "Draft applied and saved. Review/edit if needed.")

    # ---- Autofill/Test: iterate all students × assigned questions (kept; now purely a loop) ----
    def make_all_assigned(self):
        if not self.require_grading_db():
            return
        if not self.question_map:
            messagebox.showinfo("No scheme", "Load Scheme CSV first.")
            return

        students = list(self.student_ids)
        if not students:
            messagebox.showinfo("No students", "Scan/Commit first so students exist.")
            return

        # Here: intentionally does NOT call any external/AI.
        # Placeholder pass that can be extended for validation checks.
        self.refresh_summary()
        messagebox.showinfo("Done", f"Checked {len(students)} students. Scheme-driven grading is active.")

    # ---- Export Excel (all) ----
    def save_all_excel(self):
        if not self.require_grading_db():
            return
        out = filedialog.asksaveasfilename(
            title="Export all grades (Excel)",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not out:
            return
        export_all_to_excel(self.sub_con, self.grade_con, Path(out))
        messagebox.showinfo("Exported", f"Saved:\n{out}")

    def export_selected_excel(self):
        if not self.require_grading_db():
            return
        if not self.selected_student_id:
            messagebox.showinfo("Select", "Select a student first.")
            return

        out = filedialog.asksaveasfilename(
            title="Export selected student's grades (Excel)",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not out:
            return

        export_all_to_excel(self.sub_con, self.grade_con, Path(out))
        messagebox.showinfo("Exported", f"Saved:\n{out}\n\nTip: use sheet filters to keep only {self.selected_student_id} rows.")

    # ---- Summary + class stats + histogram ----
    def compute_class_values(self):
        """
        Returns list of overall raw totals for non-FULL students.
        """
        if self.grade_con is None:
            return []
        rows = self.sub_con.execute("""
          SELECT student_id FROM students
          WHERE LOWER(student_id) <> 'full'
        """).fetchall()
        vals = []
        for (sid,) in rows:
            vals.append(compute_overall_total(self.grade_con, sid))
        return vals

    def compute_class_stats_text(self):
        vals = self.compute_class_values()
        if not vals:
            return "Class Stats: (not available)"
        avg = sum(vals) / len(vals)
        mn = min(vals)
        mx = max(vals)
        med = sorted(vals)[len(vals)//2]
        variance = sum((v - avg) ** 2 for v in vals) / len(vals)
        std = math.sqrt(variance)
        target_avg = 75.0
        curve_factor = (target_avg / avg) if avg > 0 else 1.0
        return f"Class Stats: avg {avg:.2f} | median {med:.2f} | std {std:.2f} | min {mn:.2f} | max {mx:.2f} | suggested curve× {curve_factor:.3f}"

    def refresh_histogram(self):
        if self.hist_canvas is None or self.hist_ax is None:
            return
        vals = self.compute_class_values()
        self.hist_ax.clear()
        if not vals:
            self.hist_ax.set_title("No class data")
            self.hist_canvas.draw()
            return

        k = float(self.curve_preview_var.get())
        curved = [max(0.0, v * k) for v in vals]

        # No explicit colors (per your preference earlier for charts); matplotlib defaults
        self.hist_ax.hist(vals, bins=12, alpha=0.6, label="raw")
        self.hist_ax.hist(curved, bins=12, alpha=0.6, label=f"curved×{k:.2f}")
        mean_val = sum(vals) / len(vals)
        med_val = sorted(vals)[len(vals)//2]
        self.hist_ax.axvline(mean_val, linestyle="--", linewidth=1.5, label=f"mean {mean_val:.2f}")
        self.hist_ax.axvline(med_val, linestyle=":", linewidth=1.5, label=f"median {med_val:.2f}")
        self.hist_ax.set_title("Overall totals distribution")
        self.hist_ax.set_xlabel("Overall total")
        self.hist_ax.set_ylabel("Count")
        self.hist_ax.legend()
        self.hist_canvas.draw()

    def refresh_summary(self):
        if self.grade_con is None:
            for item in self.sum_tree.get_children():
                self.sum_tree.delete(item)
            self.class_stats_lbl.config(text="Class: avg - | min - | max - | curve -")
            self.refresh_histogram()
            return

        for item in self.sum_tree.get_children():
            self.sum_tree.delete(item)

        students = self.sub_con.execute("""
          SELECT student_id, student_name, COALESCE(lab_id,'')
          FROM students
          WHERE LOWER(student_id) <> 'full'
          ORDER BY student_id
        """).fetchall()

        curve_k = float(self.curve_preview_var.get())

        for sid, sname, lab in students:
            graded_count = self.grade_con.execute("SELECT COUNT(DISTINCT question_id) FROM rubric_scores WHERE student_id=?", (sid,)).fetchone()[0]
            overall = compute_overall_total(self.grade_con, sid)
            curved = max(0.0, overall * curve_k)
            self.sum_tree.insert(
                "", "end",
                values=(sid, sname, lab, graded_count, f"{overall:g}", f"{curved:.2f}")
            )

        stats_text = self.compute_class_stats_text()
        # Also keep your original suggested curve factor in the top label (separate from preview)
        self.class_stats_lbl.config(text=stats_text.replace("Class Stats:", "Class:"))

        self.refresh_histogram()

    # ---- PDF Exports ----
    def export_student_pdf(self):
        if not self.require_grading_db():
            return
        if SimpleDocTemplate is None:
            messagebox.showinfo("PDF missing", "reportlab not installed. Install: pip install reportlab")
            return
        if not self.selected_student_id:
            messagebox.showinfo("Select", "Select a student first.")
            return

        out = filedialog.asksaveasfilename(
            title="Export student PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")]
        )
        if not out:
            return

        exporter = PDFExporter(self.sub_con, self.grade_con, self.question_map)
        try:
            exporter.export_student_pdf(self.selected_student_id, Path(out))
        except Exception as e:
            messagebox.showerror("PDF export failed", str(e))
            return
        messagebox.showinfo("PDF exported", f"Saved:\n{out}")

    def export_summary_pdf(self):
        if not self.require_grading_db():
            return
        if SimpleDocTemplate is None:
            messagebox.showinfo("PDF missing", "reportlab not installed. Install: pip install reportlab")
            return

        out = filedialog.asksaveasfilename(
            title="Export summary PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")]
        )
        if not out:
            return

        exporter = PDFExporter(self.sub_con, self.grade_con, self.question_map)
        try:
            exporter.export_summary_pdf(Path(out), self.compute_class_stats_text())
        except Exception as e:
            messagebox.showerror("PDF export failed", str(e))
            return
        messagebox.showinfo("PDF exported", f"Saved:\n{out}")

    def export_all_students_pdfs(self):
        if not self.require_grading_db():
            return
        if SimpleDocTemplate is None:
            messagebox.showinfo("PDF missing", "reportlab not installed. Install: pip install reportlab")
            return

        out_dir = filedialog.askdirectory(title="Choose output folder for all student PDFs")
        if not out_dir:
            return

        report_tag = simpledialog.askstring(
            "Report tag",
            "Enter report tag to include in filename (example: Midterm):",
            initialvalue="Midterm"
        )
        if report_tag is None:
            return

        exporter = PDFExporter(self.sub_con, self.grade_con, self.question_map)

        prog = tk.Toplevel(self.root)
        prog.title("Exporting PDFs...")
        prog.geometry("520x140")
        lbl = ttk.Label(prog, text="Starting...", style="Pastel.TLabel")
        lbl.pack(pady=10)
        pb = ttk.Progressbar(prog, orient="horizontal", length=480, mode="determinate")
        pb.pack(pady=10)

        def progress_cb(i, n, sid, ok, err):
            pb["maximum"] = n
            pb["value"] = i
            if ok:
                lbl.config(text=f"[{i}/{n}] Exported: {sid}")
            else:
                lbl.config(text=f"[{i}/{n}] Failed: {sid} ({err})")
            prog.update_idletasks()

        try:
            exporter.export_all_students_pdfs(Path(out_dir), report_tag=report_tag, progress_cb=progress_cb)
        except Exception as e:
            messagebox.showerror("Batch export failed", str(e))
            prog.destroy()
            return

        prog.destroy()
        messagebox.showinfo("Done", f"Saved PDFs to:\n{out_dir}")

    def export_compare_to_full_pdfs(self):
        if not self.require_grading_db():
            return
        if SimpleDocTemplate is None:
            messagebox.showinfo("PDF missing", "reportlab not installed. Install: pip install reportlab")
            return

        out_dir = filedialog.askdirectory(title="Choose output folder (Compare-to-FULL naming)")
        if not out_dir:
            return

        report_tag = simpledialog.askstring(
            "Report tag",
            "Enter report tag (example: Midterm):",
            initialvalue="Midterm"
        )
        if report_tag is None:
            return

        exporter = PDFExporter(self.sub_con, self.grade_con, self.question_map)

        prog = tk.Toplevel(self.root)
        prog.title("Exporting Compare-to-FULL PDFs...")
        prog.geometry("520x140")
        lbl = ttk.Label(prog, text="Starting...", style="Pastel.TLabel")
        lbl.pack(pady=10)
        pb = ttk.Progressbar(prog, orient="horizontal", length=480, mode="determinate")
        pb.pack(pady=10)

        def progress_cb(i, n, sid, ok, err):
            pb["maximum"] = n
            pb["value"] = i
            if ok:
                lbl.config(text=f"[{i}/{n}] Exported: {sid}")
            else:
                lbl.config(text=f"[{i}/{n}] Failed: {sid} ({err})")
            prog.update_idletasks()

        try:
            exporter.export_compare_to_full_pdfs(Path(out_dir), report_tag=report_tag, progress_cb=progress_cb)
        except Exception as e:
            messagebox.showerror("Batch export failed", str(e))
            prog.destroy()
            return

        prog.destroy()
        messagebox.showinfo("Done", f"Saved PDFs to:\n{out_dir}")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
