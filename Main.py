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
            final_name = detected_name or "Unknown Student"
            final_id = build_student_key(detected_id, final_name)
        else:
            final_name = (detected_name or "Unknown Student").strip()
            final_id = f"NAME:{final_name}"

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
    Only treat rows as valid students when BOTH are present:
    - numeric student ID
    - non-empty student name (not placeholder text)
    """
    sid = extract_numeric_id(student_id)
    sname = re.sub(r"\s+", " ", (student_name or "").strip())
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
        parts.append(f"\n\n// ===== FILE: {fp} =====\n{content}")
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

    for sid, sname, lab in students:
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

        for sid, sname, lab in students:
            total = compute_total(grade_con, sid, qid)
            note_row = load_student_note(grade_con, sid, qid)
            rationale = note_row[0] if note_row and note_row[0] else ""

            score_map, note_map = load_student_scores(grade_con, sid, qid)

            row = [sid, sname, lab, total, rationale]
            for col_key, _group, _text, _mx in cols:
                row.append("" if score_map.get(col_key) is None else score_map.get(col_key))
                row.append(note_map.get(col_key, "") or "")
            ws.append(row)

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
            out.append(f"\n\n// ===== FILE: {fp} =====\n")

            for i, line_text in enumerate(lines, start=1):
                if i in by_line:
                    for (cid, sidx, eidx, txt, ts) in by_line[i]:
                        out.append(f"// >>> COMMENT #{cid} [{sidx}–{eidx}] ({ts}):")
                        for c_line in (txt or "").splitlines():
                            out.append(f"//     {c_line}")
                        out.append("// >>> END COMMENT\n")
                out.append(line_text)

        return "\n".join(out)

    def _build_highlighted_code_blocks(self, sid: str, max_chars_per_block: int = 90000):
        """
        Returns list of Preformatted blocks with "line highlights" effect simulated by
        prefixing lines that have comments. True PDF highlight is not supported by reportlab
        in the Acrobat-annotation sense, so we simulate with markers and (optionally) background
        in a separate table view.
        """
        blocks = []
        files = get_student_files(self.sub_con, sid)

        all_rows = fetch_code_comments_for_student(self.grade_con, sid)
        # Map: file -> set(lines_with_comment)
        file_lines = {}
        for fp, sidx, _eidx, _txt, _color, _ts in all_rows:
            try:
                line = int(str(sidx).split(".", 1)[0])
            except Exception:
                line = 1
            file_lines.setdefault(fp, set()).add(line)

        big_text = []
        for fp in files:
            code = get_file_content(self.sub_con, fp)
            if code is None:
                try:
                    code = Path(fp).read_text(encoding="utf-8", errors="ignore")
                except Exception:
                    code = ""
            lines = code.splitlines()
            marks = file_lines.get(fp, set())

            big_text.append(f"\n\n// ===== FILE: {fp} =====\n")
            for i, line in enumerate(lines, start=1):
                # marker for “highlight”
                prefix = ">> " if i in marks else "   "
                big_text.append(f"{prefix}{i:04d} | {line}")

        txt = "\n".join(big_text)
        # Split into manageable chunks
        for i in range(0, len(txt), max_chars_per_block):
            blocks.append(txt[i:i+max_chars_per_block])
        return blocks

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
        doc = SimpleDocTemplate(str(out_path), pagesize=letter, title=f"{sid} grading report")

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

            table_data = [["Group", "Criterion", "Max", "Pts", "Note"]]
            for col_key, group, text, mx in cols:
                pts = score_map.get(col_key, 0.0) or 0.0
                note = (note_map.get(col_key, "") or "")
                table_data.append([group or "", text, f"{mx:g}", f"{pts:g}", note[:180]])

            tbl = Table(table_data, colWidths=[70, 245, 40, 40, 160])
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EFE7FF")),
                ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 10))

            note_row = load_student_note(self.grade_con, sid, qid)
            rationale = note_row[0] if note_row and note_row[0] else ""
            if rationale:
                story.append(Paragraph("<b>Rationale</b>", styles["Heading3"]))
                story.append(Paragraph(rationale.replace("\n", "<br/>"), styles["Normal"]))
                story.append(Spacer(1, 10))

        # Code section: injected comments
        story.append(PageBreak())
        story.append(Paragraph("<b>Annotated Code Snapshot (Injected Comments)</b>", styles["Heading2"]))
        injected = self._build_annotated_code_injected(sid)
        story.append(Preformatted(injected[:120000], styles["Code"]))

        # Code section: simulated highlight listing with markers
        story.append(PageBreak())
        story.append(Paragraph("<b>Code Listing (Highlighted Lines Marked with '>>')</b>", styles["Heading2"]))
        blocks = self._build_highlighted_code_blocks(sid)
        for bi, btxt in enumerate(blocks):
            story.append(Preformatted(btxt, styles["Code"]))
            if bi < len(blocks) - 1:
                story.append(PageBreak())

        # Comments table (all files)
        story.append(PageBreak())
        story.append(Paragraph("<b>Code Comments (All Files)</b>", styles["Heading2"]))
        rows = fetch_code_comments_for_student(self.grade_con, sid)
        if not rows:
            story.append(Paragraph("(No code comments.)", styles["Normal"]))
        else:
            td = [["File", "Range", "Comment", "Time"]]
            for fp, sidx, eidx, txt, _color, ts in rows:
                td.append([Path(fp).name, f"{sidx}–{eidx}", (txt or "")[:260], ts or ""])
            tbl2 = Table(td, colWidths=[120, 90, 280, 60])
            tbl2.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#DFF6FF")),
                ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
                ("VALIGN", (0,0), (-1,-1), "TOP"),
                ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ]))
            story.append(tbl2)

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
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#EFE7FF")),
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ]))
        story.append(tbl)

        doc.build(story)

    def export_all_students_pdfs(self, out_dir: Path, progress_cb=None):
        out_dir.mkdir(parents=True, exist_ok=True)
        students = self.sub_con.execute("""
          SELECT student_id
          FROM students
          ORDER BY student_id
        """).fetchall()
        for i, (sid,) in enumerate(students, start=1):
            out_path = out_dir / f"{sid}_report.pdf"
            try:
                self.export_student_pdf(sid, out_path)
            except Exception as e:
                # keep going, but record failures
                if progress_cb:
                    progress_cb(i, len(students), sid, False, str(e))
                continue
            if progress_cb:
                progress_cb(i, len(students), sid, True, "")


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

    bg = "#F7F4FF"        # very light lavender
    panel = "#FFF7E8"     # light warm cream
    accent = "#C9B6FF"    # lavender
    accent2 = "#B7E3FF"   # baby blue
    text = "#2A2440"
    select = "#E6DCFF"

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

        self.title("Scan Submissions — Review + Edit before saving")
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
        self.last_scan_snapshot_path: str = ""
        self._find_from = "1.0"

        self.rows: dict[str, dict] = {}
        self.folder_order: list[str] = []

        self._build()

    def _build(self):
        top = ttk.Frame(self, padding=10, style="Pastel.TFrame")
        top.pack(fill=tk.X)

        actions = ttk.Frame(top, style="Pastel.TFrame")
        actions.pack(fill=tk.X)

        ttk.Button(actions, text="Choose ROOT Folder", command=self.choose_root).pack(side=tk.LEFT)
        ttk.Button(actions, text="Scan Now", command=self.scan).pack(side=tk.LEFT, padx=8)
        ttk.Button(actions, text="Save Scan Snapshot", command=self.save_scan_snapshot).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="Load Scan Snapshot", command=self.load_scan_snapshot).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="Load Current Snapshot", command=self.load_current_scan_snapshot).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="Set Loaded Snapshot as Current", command=self.set_loaded_snapshot_as_current).pack(side=tk.LEFT, padx=(8, 0))

        ttk.Button(actions, text="Save Scan to DB", command=self.save_to_db).pack(side=tk.RIGHT)

        self.status_lbl = ttk.Label(actions, text="No folder selected.", style="Pastel.TLabel")
        self.status_lbl.pack(side=tk.RIGHT, padx=10)

        opts_nb = ttk.Notebook(top)
        opts_nb.pack(fill=tk.X, pady=(8, 0))

        filters_tab = ttk.Frame(opts_nb, style="Pastel.TFrame", padding=8)
        regex_tab = ttk.Frame(opts_nb, style="Pastel.TFrame", padding=8)
        opts_nb.add(filters_tab, text="Scan Filters")
        opts_nb.add(regex_tab, text="Regex")

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

        main = ttk.Frame(self, padding=10, style="Pastel.TFrame")
        main.pack(fill=tk.BOTH, expand=True)

        main.columnconfigure(0, weight=2)
        main.columnconfigure(1, weight=2)
        main.columnconfigure(2, weight=3)
        main.rowconfigure(0, weight=1)

        left = ttk.Frame(main, style="Pastel.TFrame")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.rowconfigure(1, weight=1)
        ttk.Label(left, text="Folders (edit final ID/Name/LabID here before saving)", style="Pastel.TLabel").grid(row=0, column=0, sticky="w")

        cols = ("include", "folder", "det_id", "det_name", "final_id", "final_name", "lab_id", "nfiles")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse")
        for c, w in [("include", 70), ("folder", 330), ("det_id", 120), ("det_name", 160),
                     ("final_id", 160), ("final_name", 170), ("lab_id", 90), ("nfiles", 70)]:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="w")
        self.tree.grid(row=1, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self.on_folder_select)

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
        self.files_list = tk.Listbox(right, height=10)
        self.files_list.grid(row=1, column=0, sticky="nsew")
        self.files_list.bind("<<ListboxSelect>>", self.on_file_select)

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
        self.preview.bind("<KeyPress-n>", self._hotkey_use_name)
        self.preview.bind("<Control-f>", lambda _e: self.focus_find_entry())

    def choose_root(self):
        folder = filedialog.askdirectory(title="Select ROOT submissions folder")
        if not folder:
            return
        self.root_folder = Path(folder)
        self.status_lbl.config(text=f"Root: {self.root_folder}")
        self.rows.clear()
        self.folder_order.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.files_list.delete(0, tk.END)
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
            f"Folders: {c['folders']}",
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

    def save_scan_snapshot(self):
        if not self.rows:
            messagebox.showinfo("Nothing", "Scan first, then save a snapshot.")
            return
        path = filedialog.asksaveasfilename(
            title="Save scan snapshot",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return

        payload = {
            "root_folder": str(self.root_folder) if self.root_folder else "",
            "settings": self._scan_settings_payload(),
            "folder_order": self.folder_order,
            "rows": self.rows,
        }
        try:
            Path(path).write_text(json.dumps(payload, indent=2), encoding="utf-8")
            self.last_scan_snapshot_path = str(path)
        except Exception as e:
            messagebox.showerror("Save failed", str(e))
            return
        messagebox.showinfo("Saved", f"Scan snapshot saved:\n{path}")

    def _load_scan_snapshot_from_path(self, path: str):
        try:
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
        except Exception as e:
            messagebox.showerror("Load failed", str(e))
            return

        settings = payload.get("settings") or {}
        self.file_globs_var.set((settings.get("file_globs") or "*.java").strip() or "*.java")
        self.filename_regex_var.set((settings.get("filename_regex") or "").strip())
        self.folder_id_regex_var.set((settings.get("folder_id_regex") or "").strip())
        self.folder_name_regex_var.set((settings.get("folder_name_regex") or "").strip())
        self.only_new_files_var.set(bool(settings.get("only_new_files", False)))

        loaded_rows = payload.get("rows") or {}
        loaded_order = payload.get("folder_order") or list(loaded_rows.keys())
        self.rows = {k: v for k, v in loaded_rows.items() if isinstance(v, dict)}
        self.folder_order = [k for k in loaded_order if k in self.rows]

        root_raw = (payload.get("root_folder") or "").strip()
        self.root_folder = Path(root_raw) if root_raw else None

        for item in self.tree.get_children():
            self.tree.delete(item)

        for folder_key in self.folder_order:
            r = self.rows[folder_key]
            files = r.get("files") or []
            is_student = has_required_student_fields(r.get("final_id", ""), r.get("final_name", ""))
            # Backward-compatible snapshots: auto-exclude non-students or empty-file rows by default.
            base_include = bool(r["include"]) if "include" in r else True
            r["include"] = base_include and bool(files) and is_student
            self.tree.insert("", "end", iid=folder_key, values=(
                "YES" if r.get("include") else "NO",
                Path(r.get("folder") or folder_key).name,
                r.get("det_id", ""),
                r.get("det_name", ""),
                r.get("final_id", ""),
                r.get("final_name", ""),
                r.get("lab_id", ""),
                str(len(files))
            ))

        self.last_scan_snapshot_path = str(path)
        root_text = f"Root: {self.root_folder}" if self.root_folder else "Root: (from snapshot)"
        self._set_scan_status(prefix=f"{root_text} | Loaded snapshot")

    def load_scan_snapshot(self):
        path = filedialog.askopenfilename(
            title="Load scan snapshot",
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        self._load_scan_snapshot_from_path(path)

    def load_current_scan_snapshot(self):
        path = sub_meta_get(self.con, "current_scan_snapshot", "").strip()
        if not path:
            messagebox.showinfo("No current snapshot", "No current scan snapshot is set.")
            return
        if not Path(path).exists():
            messagebox.showerror("Missing file", f"Current snapshot file does not exist:\n{path}")
            return
        self._load_scan_snapshot_from_path(path)

    def set_loaded_snapshot_as_current(self):
        path = (self.last_scan_snapshot_path or "").strip()
        if not path:
            messagebox.showinfo("No snapshot", "Save or load a scan snapshot first.")
            return
        sub_meta_set(self.con, "current_scan_snapshot", path)
        messagebox.showinfo("Saved", f"Current scan snapshot set to:\n{path}")

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
                "folder": folder_key,
                "det_id": det_id or "",
                "det_name": det_name or "",
                "final_id": final_id or "",
                "final_name": final_name or (Path(sub).name if not files else ""),
                "lab_id": "",
                "files": files,
            }

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
                r.get("lab_id",""),
                str(len(r["files"]))
            ))

        self._set_scan_status(prefix="Scan done")

    def on_folder_select(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        folder_key = sel[0]
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

        self.files_list.delete(0, tk.END)
        for fp in r["files"]:
            self.files_list.insert(tk.END, fp)

        self.preview.delete("1.0", tk.END)

    def on_file_select(self, _evt=None):
        sel = self.files_list.curselection()
        if not sel:
            return
        fp = self.files_list.get(sel[0])
        try:
            content = Path(fp).read_text(encoding="utf-8", errors="ignore")
        except Exception as e:
            content = f"Error reading file:\n{e}"
        self.preview.delete("1.0", tk.END)
        self.preview.insert("1.0", content)
        self.preview.tag_remove("find_hit", "1.0", tk.END)
        self._find_from = "1.0"

    def apply_include_toggle(self):
        folder_key = self.sel_folder_var.get().strip()
        if not folder_key or folder_key not in self.rows:
            return
        row = self.rows[folder_key]
        files = row.get("files") or []
        is_student = has_required_student_fields(row.get("final_id", ""), row.get("final_name", ""))
        if not files or not is_student:
            row["include"] = False
            self.include_var.set(False)
            self._refresh_tree_row(folder_key)
            self._set_scan_status(prefix="Scan info")
            return
        row["include"] = bool(self.include_var.get())
        self._refresh_tree_row(folder_key)
        self._set_scan_status(prefix="Scan info")

    def apply_edits(self):
        folder_key = self.sel_folder_var.get().strip()
        if not folder_key or folder_key not in self.rows:
            return

        raw_fid = self.final_id_var.get().strip()
        fid = extract_numeric_id(raw_fid)
        fname = self.final_name_var.get().strip()
        lab = self.lab_id_var.get().strip()

        self.rows[folder_key]["final_id"] = fid
        self.rows[folder_key]["final_name"] = fname
        self.rows[folder_key]["lab_id"] = lab

        is_student = has_required_student_fields(fid, fname)
        has_files = bool(self.rows[folder_key].get("files"))
        self.rows[folder_key]["include"] = bool(self.rows[folder_key].get("include")) and has_files and is_student
        if is_student and has_files:
            # If user just completed ID + Name, include it by default.
            self.rows[folder_key]["include"] = True

        self._suspend_auto_apply = True
        self.include_var.set(bool(self.rows[folder_key]["include"]))
        self._suspend_auto_apply = False

        self._refresh_tree_row(folder_key)
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

        current = self.sel_folder_var.get().strip()
        if current and current in self.rows:
            self._suspend_auto_apply = True
            self.lab_id_var.set(self.rows[current].get("lab_id", ""))
            self._suspend_auto_apply = False

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
        only_id = extract_numeric_id(selected)
        if not only_id:
            return
        self.final_id_var.set(only_id)

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

    def save_to_db(self):
        if not self.rows:
            messagebox.showinfo("Nothing", "Scan first.")
            return

        store_content = self.store_file_content.get()

        committed_folders = 0
        committed_files = 0
        created_students = 0
        skipped_folders = 0

        for folder_key in self.folder_order:
            r = self.rows[folder_key]
            if not r["include"]:
                skipped_folders += 1
                continue
            if not r.get("files"):
                skipped_folders += 1
                continue

            fid = (r["final_id"] or "").strip()
            fname = (r["final_name"] or "").strip()
            lab = (r.get("lab_id") or "").strip() or None

            # Save only complete student records (numeric ID + name).
            if not has_required_student_fields(fid, fname):
                skipped_folders += 1
                continue

            fid = extract_numeric_id(fid)
            student_ok = bool(fid)

            upsert_student(self.con, fid, fname, lab, folder_key)
            created_students += 1

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
                    student_id=fid if student_ok else None,
                    source_folder=folder_key,
                    detected_id=r["det_id"] or None,
                    detected_name=r["det_name"] or None,
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

        self.grade_db_path: Path | None = None
        self.grade_con: sqlite3.Connection | None = None

        self.student_ids: list[str] = []
        self.selected_student_id: str | None = None
        self.selected_file_path: str | None = None

        self.question_map: dict[str, str] = {}
        self.selected_question_id: str | None = None

        # curve preview factor (histogram overlay)
        self.curve_preview_var = tk.DoubleVar(value=1.0)

        # Optional auto-grader
        self.auto_grader = AutoGrader(enabled=DEFAULT_AI_ENABLED)

        self._build_ui()
        self.refresh_students(keep_selected=False)

    def require_grading_db(self) -> bool:
        if self.grade_con is None:
            messagebox.showinfo("No grading DB", "Create/Open a grading DB first.")
            return False
        return True

    def _build_ui(self):
        top = ttk.Frame(self.root, padding=10, style="Pastel.TFrame")
        top.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(top, text="Open Scan Window", command=self.open_scan_window).pack(side=tk.LEFT)
        ttk.Button(top, text="Open Current Scan Snapshot", command=self.open_current_scan_snapshot).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(top, text="New Grading DB...", command=self.new_grading_db).pack(side=tk.LEFT)
        ttk.Button(top, text="Open Grading DB...", command=self.open_grading_db).pack(side=tk.LEFT, padx=6)

        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(top, text="Load Scheme CSV...", command=self.load_scheme_csv).pack(side=tk.LEFT)
        ttk.Button(top, text="Reload Last Scheme", command=self.reload_last_scheme).pack(side=tk.LEFT, padx=6)

        ttk.Separator(top, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)

        self.db_status_lbl = ttk.Label(top, text="Grading DB: (none)", style="Pastel.TLabel")
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
        self.student_list.bind("<<ListboxSelect>>", self.on_student_select)

        ttk.Label(left, text="Files", style="PastelCard.TLabel", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.file_list = tk.Listbox(left, height=8, bg=self.palette["panel"], fg=self.palette["text"], highlightthickness=0, selectbackground=self.palette["select"])
        self.file_list.grid(row=3, column=0, sticky="nsew")
        self.file_list.bind("<<ListboxSelect>>", self.on_file_select)

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

        self.preview.tag_configure("comment_highlight", background="#FFF2B2")
        self.preview.tag_configure("comment_highlight2", background="#DFF6FF")

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
        ttk.Button(top, text="Export PDF (summary)", command=self.export_summary_pdf).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Export PDFs (ALL students)", command=self.export_all_students_pdfs).pack(side=tk.LEFT, padx=6)

        cols = ("student_id", "student_name", "lab_id", "graded_questions", "overall", "overall_curved")
        self.sum_tree = ttk.Treeview(self.tab_summary, columns=cols, show="headings", height=25)
        for c, w in [("student_id", 140), ("student_name", 220), ("lab_id", 80),
                     ("graded_questions", 130), ("overall", 90), ("overall_curved", 120)]:
            self.sum_tree.heading(c, text=c)
            self.sum_tree.column(c, width=w, anchor="w")
        self.sum_tree.pack(fill=tk.BOTH, expand=True, pady=(10,0))

        note = ttk.Label(self.tab_summary,
                         text="Tip: create a student named/ID 'FULL' as your reference solution; it’s pinned to the top.",
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

    def open_current_scan_snapshot(self):
        win = ScanWindow(self)
        win.load_current_scan_snapshot()

    # ---- grading DB open/create ----
    def new_grading_db(self):
        path = filedialog.asksaveasfilename(
            title="Create grading DB",
            defaultextension=".sqlite",
            filetypes=[("SQLite DB", "*.sqlite"), ("All files", "*.*")]
        )
        if not path:
            return
        self._open_grade_db(Path(path))

    def open_grading_db(self):
        path = filedialog.askopenfilename(
            title="Open grading DB",
            filetypes=[("SQLite DB", "*.sqlite"), ("All files", "*.*")]
        )
        if not path:
            return
        self._open_grade_db(Path(path))

    def _open_grade_db(self, path: Path):
        if self.grade_con is not None:
            try:
                self.grade_con.close()
            except Exception:
                pass

        self.grade_db_path = path
        self.grade_con = db_connect(path)
        grading_db_init(self.grade_con)

        theme = meta_get(self.grade_con, "theme", DEFAULT_THEME)
        self.theme_text.delete("1.0", tk.END)
        self.theme_text.insert("1.0", theme)

        self.refresh_question_lists()
        self.db_status_lbl.config(text=f"Grading DB: {path.name}")

        self.load_student_question_view()
        self.refresh_summary()

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

    def on_file_select(self, _evt=None):
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
        self.preview.tag_remove("comment_highlight2", "1.0", tk.END)
        self.comment_list.delete(0, tk.END)

        if self.grade_con is None:
            return
        if not self.selected_student_id or not self.selected_file_path:
            return

        comments = fetch_code_comments_for_file(self.grade_con, self.selected_student_id, self.selected_file_path)
        for i, (cid, sidx, eidx, text, color, _created_at) in enumerate(comments):
            tag = "comment_highlight" if (i % 2 == 0) else "comment_highlight2"
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
        add_code_comment(self.grade_con, self.selected_student_id, self.selected_file_path, sidx, eidx, txt, color="#FFF2B2")
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
        if full_id not in self.student_ids:
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
          ORDER BY student_id
        """).fetchall()

        def key(r):
            sid = (r[0] or "")
            return (0 if sid.lower() == "full" else 1, sid)
        students = sorted(students, key=key)

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
            exporter.export_all_students_pdfs(Path(out_dir), progress_cb=progress_cb)
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
