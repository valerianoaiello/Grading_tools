"""
Grading of open-ended questions in PDF exams.

- Reads the student list from an Excel file (sheet "Scores") or a CSV file. Column
  names are recognised in English or Italian (Exam/Compito, Student ID/Matricola,
  Surname/Cognome, Name/Nome, Notes/Note).
  Excel: every column with a numeric value in the "max" row is a question to grade.
  CSV: the questions come from the grading rubric and their columns are added.
- Each PDF is matched to a student through the exam number, taken from the
  last group of digits in the file name (e.g. "exam_12.pdf" -> 12).
- Lets you draw on the pages (mouse or pen tablet) and write comments; on save it
  writes the scores to the Excel/CSV file and rebuilds the PDF (original + drawings
  + evaluation page). Originals are kept in the "originals" subfolder, so an exam can
  be regraded without duplicating anything.
- Grading rubric (rubric.txt in the PDF folder): for each exercise, the points
  given for each part of the answer.
- Proposals (proposals.json in the PDF folder, e.g. written by Claude in a Claude Code
  session after reading the scans, the rubric and the solutions): proposed scores and
  comment, shown highlighted, to be checked and edited before saving.
- Solutions: a zip with one file per exam (.tex compiled with pdflatex, or .pdf) or a
  single PDF for everybody. Shown in a second window and optionally appended to the
  student's PDF.
- True/false page (CSV only): a first page with the answers given, the correct answers
  and the points of a true/false part.
- Student-facing pages are in Italian or English (per exam, with a folder default).

Folders created by older versions (griglia.txt, proposte.json, soluzioni.zip,
originali/, .correzione_stato.json) keep working with their old file names.
"""
import io
import os
import re
import csv
import json
import shutil
import tempfile
import threading
import time
import subprocess
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext, ttk
from datetime import datetime

from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pdf2image import convert_from_path
from PIL import Image, ImageTk
from openpyxl import load_workbook

CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".examinatrix_grader.json")
RENDER_DPI = 200
PEN_COLOR = "red"
PEN_WIDTH_SCREEN = 3
PEN_WIDTH_PDF = 1.5  # points
ERASER_RADIUS = 10   # screen pixels: strokes closer than this to the eraser are removed
# Questions that accept negative scores (down to -maximum), unless the Excel
# file has a "min" row giving the minimum of each question
NEGATIVE_QUESTIONS = {"G1", "G2"}
PROPOSAL_BG = "#fff3b0"                 # background of proposed (unchecked) scores
SOLUTIONS_DPI = 150

# Files kept in the PDF folder: (name, name used by older versions or None).
# The old name is used when only the old file exists.
FOLDER_FILES = {
    "state": (".grading_state.json", ".correzione_stato.json"),
    "originals": ("originals", "originali"),
    "rubric": ("rubric.txt", "griglia.txt"),
    "proposals": ("proposals.json", "proposte.json"),
    "solutions_zip": ("solutions.zip", "soluzioni.zip"),
    "solutions_pdf": ("solutions.pdf", None),
    "compiled": ("solutions_pdf", "soluzioni_pdf"),   # solutions compiled from the zip
}

# Column names accepted in the Excel/CSV file (compared lower case)
COLUMN_NAMES = {
    "exam": ("exam", "exam number", "exam no", "compito", "numero compito"),
    "student_id": ("student id", "id", "matricola"),
    "surname": ("surname", "last name", "family name", "cognome"),
    "name": ("name", "first name", "given name", "nome"),
    "notes": ("notes", "note"),
}
ITALIAN_EXAM_COLUMNS = {"compito", "numero compito"}
SHEET_NAMES = ("Scores", "Punteggi")
MAX_ROW_LABELS = ("max", "maximum", "massimo")
MIN_ROW_LABELS = ("min", "minimum", "minimo")
# Column with the total of the open questions, added to a CSV file
CSV_TOTAL_COLUMN = {False: "Open questions total", True: "Totale domande aperte"}  # key: Italian file

# Texts of the evaluation page and of the quick comment buttons (key: English)
LABELS = {
    False: {"title": "VALUTAZIONE", "exam": "Compito", "id": "Matricola",
            "scores": "PUNTEGGI", "total": "Totale domande aperte",
            "comments": "COMMENTI", "none": "Nessun commento",
            "quick": ["Risposta incompleta"]},
    True: {"title": "EVALUATION", "exam": "Exam", "id": "Student ID",
           "scores": "SCORES", "total": "Total (open questions)",
           "comments": "COMMENTS", "none": "No comments",
           "quick": ["Incomplete answer"]},
}

RUBRIC_HELP = (
    "One section per exercise: [name] on its own line (the name of the score column).\n"
    "Below it, one line per criterion: points and description, e.g.  1  correct characteristic equation\n"
    "Negative points = penalties (e.g.  -0.25  calculation error). The maximum of the exercise\n"
    "is the sum of its positive points. Other lines are notes for the grader; # = comment.")


# ==================================================
# Utilities
# ==================================================

def exam_number(filename):
    """Exam number: last group of digits in the file name."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    nums = re.findall(r"\d+", stem)
    return int(nums[-1]) if nums else None


def fmt_num(x, comma=False):
    """Number without needless decimals; decimal comma if requested (Italian pages)."""
    if x is None:
        return ""
    x = float(x)
    s = f"{x:.0f}" if x == int(x) else f"{x:.2f}".rstrip("0")
    return s.replace(".", ",") if comma else s


def parse_num(text):
    """Number typed with a decimal point or comma."""
    return float(str(text).strip().replace(",", "."))


def to_float(x):
    try:
        return parse_num(x)
    except ValueError:
        return None


def folder_path(folder, key):
    """Path of one of the program's files in the PDF folder (see FOLDER_FILES)."""
    new, old = FOLDER_FILES[key]
    if old and not os.path.exists(os.path.join(folder, new)) \
            and os.path.exists(os.path.join(folder, old)):
        return os.path.join(folder, old)
    return os.path.join(folder, new)


def find_column(lower, key):
    """Index of the column for a field (see COLUMN_NAMES), or None.
    lower: {lower-case header: index}."""
    return next((lower[a] for a in COLUMN_NAMES[key] if a in lower), None)


def _register_font():
    """Unicode-capable font for the generated pages (falls back to Helvetica)."""
    candidates = [
        ("/System/Library/Fonts/Supplemental/Arial Unicode.ttf", None),
        ("/System/Library/Fonts/Supplemental/Arial.ttf",
         "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ]
    for regular, bold in candidates:
        if os.path.exists(regular):
            try:
                pdfmetrics.registerFont(TTFont("EvalFont", regular))
                if bold and os.path.exists(bold):
                    pdfmetrics.registerFont(TTFont("EvalFont-Bold", bold))
                    return "EvalFont", "EvalFont-Bold"
                return "EvalFont", "EvalFont"
            except Exception:
                continue
    return "Helvetica", "Helvetica-Bold"


def load_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
    except Exception:
        pass


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ==================================================
# Grading rubric
# ==================================================

_CRITERION = re.compile(r"^([+-]?\d+(?:[.,]\d+)?)\s+(.+)$")


def parse_rubric(text):
    """Rubric text -> list of exercises {"name", "criteria": [(points, text)],
    "notes": [str], "max": float}. Raises ValueError on malformed input."""
    exercises = []
    for lineno, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("["):
            end = line.find("]")
            name = line[1:end].strip() if end > 0 else ""
            if not name:
                raise ValueError(f"Line {lineno}: missing exercise name: {raw!r}")
            if any(e["name"] == name for e in exercises):
                raise ValueError(f"Line {lineno}: exercise {name} appears twice")
            exercises.append({"name": name, "criteria": [], "notes": []})
            continue
        if not exercises:
            raise ValueError(f"Line {lineno}: write an [exercise name] first")
        m = _CRITERION.match(line)
        if m:
            exercises[-1]["criteria"].append((parse_num(m.group(1)), m.group(2).strip()))
        else:
            exercises[-1]["notes"].append(line)
    for e in exercises:
        e["max"] = sum(p for p, _ in e["criteria"] if p > 0)
    return exercises


def rubric_template(questions):
    """Empty rubric with one section per question (name, maximum)."""
    lines = ["# Grading rubric", ""]
    for name, qmax in questions:
        lines.append(f"[{name}]" + (f"   # maximum {fmt_num(qmax)}" if qmax else ""))
        lines.append("")
    return "\n".join(lines)


# ==================================================
# Score files (Excel or CSV)
# ==================================================

class ExcelGrades:
    """Reads/writes scores in the Excel file, by exam number."""
    kind = "Excel"
    questions_from_rubric = False

    def __init__(self, path):
        self.path = path
        self.wb = load_workbook(path)
        sheet = next((s for s in SHEET_NAMES if s in self.wb.sheetnames), None)
        self.ws = self.wb[sheet] if sheet else self.wb.active
        ws = self.ws

        headers = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if v is not None and str(v).strip():
                headers[c] = str(v).strip()
        lower = {h.lower(): c for c, h in headers.items()}

        self.col_exam = find_column(lower, "exam")
        if self.col_exam is None:
            raise ValueError("No 'Exam' (or 'Compito') column in row 1 of the sheet.")
        self.col_student_id = find_column(lower, "student_id")
        self.col_surname = find_column(lower, "surname")
        self.col_name = find_column(lower, "name")
        self.col_notes = find_column(lower, "notes")

        # Row of maximums and student rows
        self.max_row = None
        self.min_row = None
        self.rows = {}
        self.duplicates = []
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=self.col_exam).value
            label = v.strip().lower() if isinstance(v, str) else None
            if label in MAX_ROW_LABELS:
                self.max_row = r
                continue
            if label in MIN_ROW_LABELS:
                self.min_row = r
                continue
            try:
                n = int(v)
            except (TypeError, ValueError):
                continue
            if n in self.rows:
                self.duplicates.append(n)
            else:
                self.rows[n] = r
        if self.max_row is None:
            raise ValueError("No 'max' (or 'massimo') row in the exam column.")

        # Questions = columns with a numeric maximum
        self.questions = []  # list of (name, column, maximum)
        for c, h in sorted(headers.items()):
            if c == self.col_exam:
                continue
            m = ws.cell(row=self.max_row, column=c).value
            if _is_number(m):
                self.questions.append((h, c, float(m)))
        if not self.questions:
            raise ValueError("No questions found (no numeric value in the 'max' row).")

        # Minimum score: from the "min" row if present, otherwise -maximum for
        # questions in NEGATIVE_QUESTIONS and 0 for the others
        self.minimums = {}
        for qname, c, qmax in self.questions:
            m = ws.cell(row=self.min_row, column=c).value if self.min_row else None
            if _is_number(m):
                self.minimums[qname] = float(m)
            elif qname.upper() in NEGATIVE_QUESTIONS:
                self.minimums[qname] = -qmax
            else:
                self.minimums[qname] = 0.0

        self.backup_path = None

    @property
    def max_total(self):
        return sum(q[2] for q in self.questions)

    def lock_file_exists(self):
        """Excel creates '~$filename.xlsx' while the file is open."""
        d, name = os.path.split(self.path)
        return os.path.exists(os.path.join(d, "~$" + name))

    def student(self, n):
        r = self.rows.get(n)
        if r is None:
            return None

        def val(col):
            if col is None:
                return ""
            v = self.ws.cell(row=r, column=col).value
            return "" if v is None else str(v)

        return {"exam": n, "student_id": val(self.col_student_id),
                "surname": val(self.col_surname), "name": val(self.col_name),
                "notes": val(self.col_notes)}

    def get_scores(self, n):
        r = self.rows[n]
        out = []
        for _, c, _ in self.questions:
            v = self.ws.cell(row=r, column=c).value
            out.append(v if _is_number(v) else None)
        return out

    def is_graded(self, n):
        return all(v is not None for v in self.get_scores(n))

    def set_scores(self, n, scores):
        r = self.rows[n]
        for (_, c, _), s in zip(self.questions, scores):
            self.ws.cell(row=r, column=c).value = int(s) if float(s).is_integer() else s

    def save(self):
        if self.backup_path is None:
            stem, ext = os.path.splitext(self.path)
            self.backup_path = f"{stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{ext}"
            shutil.copy2(self.path, self.backup_path)
        tmp = self.path + ".tmp"
        self.wb.save(tmp)
        os.replace(tmp, self.path)


class CsvGrades:
    """Reads/writes scores in a CSV file with an exam number column (e.g. the CSV of
    "Correggi VF"). The questions are not in the file: they come from the rubric
    (set_questions), and their columns are added if missing."""
    kind = "CSV"
    questions_from_rubric = True

    def __init__(self, path):
        self.path = path
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            text = f.read()
        try:
            self.dialect = csv.Sniffer().sniff(text.split("\n", 1)[0], delimiters=",;\t")
        except csv.Error:
            self.dialect = csv.excel
        self.data = [r for r in csv.reader(io.StringIO(text), self.dialect)]
        if not self.data:
            raise ValueError("The CSV file is empty.")
        self.header = [h.strip() for h in self.data[0]]
        lower = {h.lower(): i for i, h in enumerate(self.header) if h}

        self.col_exam = find_column(lower, "exam")
        if self.col_exam is None:
            raise ValueError("No 'Exam' (or 'Numero compito') column in the first row.")
        italian = self.header[self.col_exam].lower() in ITALIAN_EXAM_COLUMNS
        # Total of the open questions: an existing column, or one in the file's language
        self.total_column = next((t for t in CSV_TOTAL_COLUMN.values() if t in self.header),
                                 CSV_TOTAL_COLUMN[italian])
        self.col_student_id = find_column(lower, "student_id")
        self.col_surname = find_column(lower, "surname")
        self.col_name = find_column(lower, "name")
        self.col_notes = find_column(lower, "notes")

        self.rows = {}
        self.duplicates = []
        for i, row in enumerate(self.data[1:], start=1):
            try:
                n = int(str(row[self.col_exam]).strip())
            except (IndexError, ValueError):
                continue
            if n in self.rows:
                self.duplicates.append(n)
            else:
                self.rows[n] = i
        self.questions = []
        self.minimums = {}
        self.backup_path = None

    def _column(self, name):
        if name not in self.header:
            self.header.append(name)
        return self.header.index(name)

    def set_questions(self, exercises):
        """exercises: list of (name, maximum)."""
        self.questions = [(name, self._column(name), float(qmax)) for name, qmax in exercises]
        self.minimums = {name: (-qmax if name.upper() in NEGATIVE_QUESTIONS else 0.0)
                         for name, _, qmax in self.questions}

    @property
    def max_total(self):
        return sum(q[2] for q in self.questions)

    def lock_file_exists(self):
        d, name = os.path.split(self.path)
        return os.path.exists(os.path.join(d, "~$" + name))

    def _cell(self, r, c):
        row = self.data[r]
        return row[c].strip() if c is not None and c < len(row) else ""

    def student(self, n):
        r = self.rows.get(n)
        if r is None:
            return None
        return {"exam": n, "student_id": self._cell(r, self.col_student_id),
                "surname": self._cell(r, self.col_surname), "name": self._cell(r, self.col_name),
                "notes": self._cell(r, self.col_notes)}

    def get_scores(self, n):
        r = self.rows[n]
        return [to_float(self._cell(r, c)) for _, c, _ in self.questions]

    def is_graded(self, n):
        return bool(self.questions) and all(v is not None for v in self.get_scores(n))

    def set_scores(self, n, scores):
        row = self.data[self.rows[n]]
        values = [(c, f"{s:g}") for (_, c, _), s in zip(self.questions, scores)]
        values.append((self._column(self.total_column), f"{sum(scores):g}"))
        for c, v in values:
            row.extend([""] * (c + 1 - len(row)))
            row[c] = v

    def save(self):
        if self.backup_path is None:
            stem, ext = os.path.splitext(self.path)
            self.backup_path = f"{stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{ext}"
            shutil.copy2(self.path, self.backup_path)
        width = len(self.header)
        rows = [self.header] + [r + [""] * (width - len(r)) for r in self.data[1:]]
        fd, tmp = tempfile.mkstemp(suffix=".csv", dir=os.path.dirname(self.path))
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as f:
            csv.writer(f, delimiter=self.dialect.delimiter).writerows(rows)
        shutil.copymode(self.path, tmp)
        os.replace(tmp, self.path)


# ==================================================
# PDF
# ==================================================

def _display_to_pdf(u, v, page):
    """Normalized coordinates on the displayed page (u to the right, v downwards)
    -> PDF coordinates of the page, taking cropbox and /Rotate into account."""
    box = page.cropbox
    llx, lly = float(box.left), float(box.bottom)
    W, H = float(box.width), float(box.height)
    rot = int(page.get("/Rotate", 0) or 0) % 360
    if rot == 90:
        x, y = v * W, u * H
    elif rot == 180:
        x, y = (1 - u) * W, v * H
    elif rot == 270:
        x, y = (1 - v) * W, (1 - u) * H
    else:
        x, y = u * W, (1 - v) * H
    return llx + x, lly + y


def _overlay_page(page, strokes):
    """Transparent PDF page with the (vector) drawings."""
    box = page.mediabox
    buf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    buf.close()
    c = rl_canvas.Canvas(buf.name, pagesize=(float(box.right), float(box.top)))
    c.setStrokeColorRGB(1, 0, 0)
    c.setLineWidth(PEN_WIDTH_PDF)
    c.setLineCap(1)
    c.setLineJoin(1)
    for stroke in strokes:
        if len(stroke) < 2:
            continue
        p = c.beginPath()
        x, y = _display_to_pdf(*stroke[0], page)
        p.moveTo(x, y)
        for u, v in stroke[1:]:
            x, y = _display_to_pdf(u, v, page)
            p.lineTo(x, y)
        c.drawPath(p, stroke=1, fill=0)
    c.showPage()
    c.save()
    with open(buf.name, "rb") as f:
        data = f.read()
    os.unlink(buf.name)
    return PdfReader(io.BytesIO(data)).pages[0]


def _evaluation_pages(student, questions, scores, comments, fonts, english=False):
    regular, bold = fonts
    T = LABELS[bool(english)]
    num = lambda x: fmt_num(x, comma=not english)
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    left, right = 50, W - 50

    c.setFont(bold, 18)
    c.drawCentredString(W / 2, H - 50, T["title"])
    c.line(left, H - 62, right, H - 62)

    y = H - 85
    c.setFont(regular, 11)
    who = f"{T['exam']} {student['exam']}"
    name = f"{student.get('surname', '')} {student.get('name', '')}".strip()
    if name:
        who += f" - {name}"
    if student.get("student_id"):
        who += f" - {T['id']} {student['student_id']}"
    c.drawString(left, y, who)
    y -= 35

    c.setFont(bold, 14)
    c.drawString(left, y, T["scores"])
    y -= 25
    c.setFont(regular, 12)
    total = 0.0
    for (qname, _, qmax), s in zip(questions, scores):
        c.drawString(70, y, f"{qname}: {num(s)} / {num(qmax)}")
        total += s
        y -= 22
    y -= 8
    c.setFont(bold, 13)
    c.drawString(70, y, f"{T['total']}: {num(total)} / {num(sum(q[2] for q in questions))}")
    y -= 40

    c.setFont(bold, 14)
    c.drawString(left, y, T["comments"])
    y -= 25
    c.setFont(regular, 11)
    comments = comments.strip()
    if not comments:
        c.drawString(70, y, T["none"])
    else:
        for paragraph in comments.split("\n"):
            if not paragraph.strip():
                y -= 10
                continue
            for line in simpleSplit(paragraph, regular, 11, right - 70):
                if y < 50:
                    c.showPage()
                    c.setFont(regular, 11)
                    y = H - 50
                c.drawString(70, y, line)
                y -= 16
            y -= 4
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages


def build_graded_pdf(original, output, strokes_by_page, student, questions,
                     scores, comments, fonts, english=False, first_page=None, solutions=None):
    """Builds the graded PDF from the original (which is left untouched):
    [first_page] + exam with drawings + evaluation page(s) + [solutions PDF]."""
    reader = PdfReader(original)
    writer = PdfWriter()
    if first_page is not None:
        writer.add_page(first_page)
    for i, page in enumerate(reader.pages, start=1):
        strokes = strokes_by_page.get(str(i)) or strokes_by_page.get(i)
        if strokes:
            page.merge_page(_overlay_page(page, strokes))
        writer.add_page(page)
    for p in _evaluation_pages(student, questions, scores, comments, fonts, english):
        writer.add_page(p)
    if solutions:
        for p in PdfReader(solutions).pages:
            writer.add_page(p)

    fd, tmp = tempfile.mkstemp(suffix=".pdf", dir=os.path.dirname(output))
    with os.fdopen(fd, "wb") as f:
        writer.write(f)
    os.chmod(tmp, 0o644)              # mkstemp makes it private to you
    os.replace(tmp, output)


# ==================================================
# True/false page (from a CSV with one answer column per question)
# ==================================================

TF_TEXT = {
    False: {"title": "Domande vero/falso", "exam": "Compito n.", "student": "Studente",
            "id": "Matricola", "question": "Domanda", "given": "Risposta data",
            "correct": "Risposta corretta", "outcome": "Esito", "points": "Punti",
            "total": "Punteggio vero/falso", "right": "giusta", "wrong": "sbagliata",
            "blank": "non data", "special": "DSA", "answer": {"T": "Vero", "F": "Falso"}},
    True: {"title": "True/false questions", "exam": "Exam no.", "student": "Student",
           "id": "Student ID", "question": "Question", "given": "Your answer",
           "correct": "Correct answer", "outcome": "Outcome", "points": "Points",
           "total": "True/false score", "right": "correct", "wrong": "wrong",
           "blank": "not given", "special": "Special arrangements",
           "answer": {"T": "True", "F": "False"}},
}
GREEN, RED, GREY = (0.0, 0.5, 0.15), (0.8, 0.1, 0.1), (0.45, 0.45, 0.45)

# Fields of the column mapping: (key, label shown in the window, required)
TF_FIELDS = [("exam", "Exam number", True), ("surname", "Surname", True),
             ("name", "First name", True), ("student_id", "Student ID", False),
             ("special", "Special arrangements (e.g. DSA)", False),
             ("total", "True/false score (total)", True),
             ("key", "Correct answers (all in one column)", True)]
TF_COLUMN_NAMES = {
    **{k: COLUMN_NAMES[k] for k in ("exam", "surname", "name", "student_id")},
    "special": ("special arrangements", "dsa"),
    "total": ("total", "totale", "true/false total"),
    "key": ("correct answers", "answer key", "key", "risposte giuste"),
}
ANSWER_PREFIXES = ("answer ", "risposta ")    # e.g. "Answer 1A", "Risposta 1A"
POINTS_PREFIXES = ("points ", "punti ")       # e.g. "Points 1A", "Punti 1A"
# Keys of the mapping saved by older versions
_OLD_TF_KEYS = {"num": "exam", "cognome": "surname", "nome": "name",
                "matricola": "student_id", "dsa": "special"}


def tf_value(x):
    """'T' (true), 'F' (false) or '' for an answer written as T/F, V/F, True/Vero..."""
    x = str(x).strip().upper()
    if x in ("T", "V", "TRUE", "VERO"):
        return "T"
    if x in ("F", "FALSE", "FALSO"):
        return "F"
    return ""


def _fmt_text(x, comma):
    """A number read from the CSV, formatted (unchanged if not a number)."""
    v = to_float(x)
    return fmt_num(v, comma) if v is not None else str(x)


def _strip_prefix(header, prefixes):
    """Rest of the header after one of the prefixes, or None."""
    h = header.lower()
    return next((header[len(p):].strip() for p in prefixes if h.startswith(p)), None)


def guess_tf_columns(header):
    """Guess of the columns, by name: {field: index or None, "answers": [...], "points": [...]}"""
    low = [h.lower() for h in header]
    lower = {h: i for i, h in reversed(list(enumerate(low))) if h}
    m = {key: next((lower[a] for a in names if a in lower), None)
         for key, names in TF_COLUMN_NAMES.items()}
    key_names = set(TF_COLUMN_NAMES["key"])
    m["answers"] = [i for i, h in enumerate(header)
                    if _strip_prefix(h, ANSWER_PREFIXES) and low[i] not in key_names]
    points = {_strip_prefix(h, POINTS_PREFIXES): i for i, h in enumerate(header)
              if _strip_prefix(h, POINTS_PREFIXES)}
    m["points"] = [points[q] for q in (_strip_prefix(header[i], ANSWER_PREFIXES)
                                       for i in m["answers"]) if q in points]
    if len(m["points"]) != len(m["answers"]):
        m["points"] = []
    return m


def check_tf_columns(m):
    """Error message for an incomplete mapping, or None."""
    for key, label, required in TF_FIELDS:
        if required and m.get(key) is None:
            return f"Choose the «{label}» column."
    if not m.get("answers"):
        return "Choose the columns with the student's answers."
    if m.get("points") and len(m["points"]) != len(m["answers"]):
        return (f"You chose {len(m['answers'])} answer columns and {len(m['points'])} "
                "points columns: they must be as many (or no points column at all).")
    return None


def tf_question_label(name):
    return _strip_prefix(name, ANSWER_PREFIXES) or name


def tf_students(header, rows, m):
    """Returns (students {n: dict}, questions [label], max_points [float|None])."""
    err = check_tf_columns(m)
    if err:
        raise ValueError(err)
    questions = [tf_question_label(header[i]) for i in m["answers"]]
    students = {}
    for r in rows:
        r = r + [""] * (len(header) - len(r))
        try:
            n = int(float(r[m["exam"]]))
        except ValueError:
            continue                      # e.g. a "max" row
        get = lambda k: r[m[k]].strip() if m.get(k) is not None else ""
        special = get("special")
        students[n] = {
            "surname": get("surname"), "name": get("name"), "student_id": get("student_id"),
            "special": bool(special) and special.lower() not in ("no", "0", "false"),
            "total": get("total"),
            "key": [tf_value(ch) for ch in re.sub(r"[\s,;]", "", get("key"))],
            "given": [tf_value(r[i]) for i in m["answers"]],
            "points": [r[i].strip() for i in m["points"]] or [""] * len(questions),
        }
    # Maximum of each question: the best score obtained by someone (unknown if nobody got it)
    max_points = []
    for k in range(len(questions)):
        vals = [to_float(s["points"][k]) for s in students.values()]
        best = max((v for v in vals if v is not None), default=0)
        max_points.append(best if best > 0 else None)
    return students, questions, max_points


def true_false_page(n, st, questions, max_points, fonts, english=False):
    """First page with the true/false part: answers given, correct answers, points."""
    regular, bold = fonts
    T = TF_TEXT[bool(english)]
    comma = not english
    W, H = size = A4
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=size)
    left, right = 60, W - 60
    y = H - 70
    c.setFont(bold, 18)
    c.drawString(left, y, T["title"])
    c.setFont(regular, 12)
    c.drawRightString(right, y, f"{T['exam']} {n:05d}")
    y -= 28
    c.drawString(left, y, f"{T['student']}: {st['surname']} {st['name']}".rstrip())
    if st["student_id"]:
        c.drawRightString(right, y, f"{T['id']}: {st['student_id']}")
    if st["special"]:
        y -= 18
        c.setFont(regular, 11)
        c.drawString(left, y, T["special"])
    y -= 36

    cols = [left, left + 85, left + 200, left + 330, right]   # question, given, correct, outcome, points
    c.setFont(bold, 11)
    for x, t in zip(cols[:4], (T["question"], T["given"], T["correct"], T["outcome"])):
        c.drawString(x, y, t)
    c.drawRightString(cols[4], y, T["points"])
    y -= 8
    c.setLineWidth(0.8)
    c.line(left, y, right, y)
    y -= 18

    words = T["answer"]
    for k, q in enumerate(questions):
        given = st["given"][k] if k < len(st["given"]) else ""
        key = st["key"][k] if k < len(st["key"]) else ""
        if not given:
            outcome, colour = T["blank"], GREY
        elif given == key:
            outcome, colour = T["right"], GREEN
        else:
            outcome, colour = T["wrong"], RED
        c.setFillColorRGB(0, 0, 0)
        c.setFont(bold, 12)
        c.drawString(cols[0] + 10, y, q)
        c.setFont(regular, 12)
        c.drawString(cols[1], y, words.get(given, "—"))
        c.drawString(cols[2], y, words.get(key, "?"))
        c.setFillColorRGB(*colour)
        c.drawString(cols[3], y, outcome)
        c.setFillColorRGB(0, 0, 0)
        pts = st["points"][k] if k < len(st["points"]) else ""
        mx = max_points[k]
        c.drawRightString(cols[4], y, _fmt_text(pts, comma)
                          + (f" / {fmt_num(mx, comma)}" if mx and pts != "" else ""))
        y -= 22
        if k + 1 < len(questions) and questions[k + 1][:1] != q[:1]:   # new exercise
            c.setStrokeColorRGB(0.8, 0.8, 0.8)
            c.line(left, y + 12, right, y + 12)
            c.setStrokeColorRGB(0, 0, 0)
            y -= 6

    y -= 4
    c.setLineWidth(0.8)
    c.line(left, y + 12, right, y + 12)
    y -= 14
    c.setFont(bold, 14)
    # Students with special arrangements may answer fewer questions: no overall maximum
    total_max = sum(max_points) if all(max_points) and not st["special"] else None
    c.drawString(left, y, T["total"])
    c.drawRightString(right, y, _fmt_text(st["total"], comma)
                      + (f" / {fmt_num(total_max, comma)}" if total_max else ""))
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def choose_tf_columns(root, header, rows, guess):
    """Window to choose the columns of the CSV. Returns the mapping, or None if cancelled."""
    NONE = "(none)"
    names = [f"{i}: {h}" for i, h in enumerate(header)]
    example = next((r for r in rows if r and any(c.strip() for c in r)), [])

    def sample(i):
        v = example[i].strip() if i is not None and i < len(example) else ""
        return f"e.g. {v[:40]}" if v else ""

    win = tk.Toplevel(root)
    win.title("True/false columns")
    win.grab_set()
    tk.Label(win, text="Check which CSV columns to use (the first data row is shown "
                       "as an example).", font=("Arial", 11)).grid(
        row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 6))
    combos = {}
    for r, (key, label, required) in enumerate(TF_FIELDS, start=1):
        tk.Label(win, text=label + ("" if required else " (optional)")).grid(
            row=r, column=0, sticky="w", padx=10, pady=2)
        cb = ttk.Combobox(win, values=([] if required else [NONE]) + names,
                          state="readonly", width=38)
        cb.set(names[guess[key]] if guess.get(key) is not None else ("" if required else NONE))
        cb.grid(row=r, column=1, sticky="w", padx=5, pady=2)
        ex = tk.Label(win, text=sample(guess.get(key)), fg="gray")
        ex.grid(row=r, column=2, sticky="w", padx=5)
        cb.bind("<<ComboboxSelected>>", lambda e, cb=cb, ex=ex: ex.config(
            text=sample(names.index(cb.get()) if cb.get() in names else None)))
        combos[key] = cb

    r = len(TF_FIELDS) + 1
    lists = {}
    for c, (key, label) in enumerate((("answers", "Student's answers\n(one column per question)"),
                                      ("points", "Points of each question (optional,\n"
                                                 "in the same order as the answers)"))):
        f = tk.Frame(win)
        f.grid(row=r, column=c, columnspan=1 if c == 0 else 2, sticky="nsew", padx=10, pady=(10, 0))
        tk.Label(f, text=label, justify="left").pack(anchor="w")
        lb = tk.Listbox(f, selectmode="multiple", exportselection=False, height=14, width=34)
        sb = tk.Scrollbar(f, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side=tk.LEFT, fill="both", expand=True)
        sb.pack(side=tk.LEFT, fill="y")
        for n in names:
            lb.insert(tk.END, n)
        for i in guess.get(key, []):
            lb.selection_set(i)
        if guess.get(key):
            lb.see(guess[key][0])
        lists[key] = lb
    count = tk.Label(win, fg="gray")
    count.grid(row=r + 1, column=0, columnspan=3, sticky="w", padx=10)

    def update_count(e=None):
        a, p = len(lists["answers"].curselection()), len(lists["points"].curselection())
        count.config(text=f"Questions chosen: {a}   —   points columns: {p}",
                     fg="gray" if p in (0, a) else "red")
    for lb in lists.values():
        lb.bind("<<ListboxSelect>>", update_count)
    update_count()

    result = {}

    def ok():
        m = {}
        for key, cb in combos.items():
            m[key] = names.index(cb.get()) if cb.get() in names else None
        m["answers"] = list(lists["answers"].curselection())
        m["points"] = list(lists["points"].curselection())
        err = check_tf_columns(m)
        if err:
            messagebox.showwarning("Columns", err, parent=win)
            return
        result["m"] = m
        win.destroy()

    bf = tk.Frame(win)
    bf.grid(row=r + 2, column=0, columnspan=3, pady=10)
    tk.Button(bf, text="Cancel", width=12, command=win.destroy).pack(side=tk.LEFT, padx=5)
    tk.Button(bf, text="Continue", width=12, command=ok).pack(side=tk.LEFT, padx=5)
    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.wait_window()
    return result.get("m")


# ==================================================
# Solutions: a zip with one file per exam (compiled to PDF), or one PDF for everybody
# ==================================================

def _find_pdflatex():
    for c in (shutil.which("pdflatex"), "/Library/TeX/texbin/pdflatex",
              "/usr/local/bin/pdflatex", "/opt/homebrew/bin/pdflatex"):
        if c and os.path.exists(c):
            return c
    return None


class SolutionsCompiler:
    """Gives the solutions PDF of each exam.

    The source (from get_source()) is either one PDF for everybody, or a zip holding one
    file per exam (.tex or .pdf) whose name ends with the exam number (00012.tex -> 12).
    The .tex files are compiled in a background thread into out_dir and reused until the
    zip changes.
    """

    def __init__(self, get_source, out_dir):
        self.get_source = get_source
        self.out_dir = out_dir
        self.pdflatex = _find_pdflatex()
        self.cond = threading.Condition()
        self.queue = []
        self.busy = None
        self.errors = {}         # {n: message}
        self.entries = {}        # {n: name in the zip}
        self.source = None       # path of the zip or of the single PDF
        self.single = False      # one PDF for everybody
        self.source_mtime = None
        self.work_dir = None
        self.thread = None

    def refresh(self):
        """Re-reads the source if it changed. Returns False if there are no solutions."""
        src = self.get_source()
        try:
            mtime = os.path.getmtime(src) if src else None
        except OSError:
            mtime = None
        if mtime is None:
            with self.cond:
                self.entries, self.source, self.single = {}, None, False
            return False
        if src != self.source or mtime != self.source_mtime:
            single = src.lower().endswith(".pdf")
            entries = {}
            if not single:
                with zipfile.ZipFile(src) as z:
                    for name in z.namelist():
                        base = os.path.basename(name)
                        stem, ext = os.path.splitext(base)
                        nums = re.findall(r"\d+", stem)
                        if ext.lower() in (".tex", ".pdf") and nums and not base.startswith("."):
                            n = int(nums[-1])
                            # a ready-made PDF wins over the LaTeX source
                            if n not in entries or ext.lower() == ".pdf":
                                entries[n] = name
            with self.cond:
                self.cleanup()
                self.queue = []
                self.entries = entries
                self.source, self.single, self.source_mtime = src, single, mtime
                self.errors = {}
        return True

    def has(self, n):
        return self.single or n in self.entries

    def cleanup(self):
        """Removes the extracted zip (it is extracted again when needed)."""
        if self.work_dir and self.busy is None:
            shutil.rmtree(self.work_dir, ignore_errors=True)
        self.work_dir = None

    def pdf_path(self, n):
        if self.single:
            return self.source
        return os.path.join(self.out_dir, f"{n:05d}.pdf")

    def ready(self, n):
        if self.single:
            return self.source is not None and os.path.exists(self.source)
        p = self.pdf_path(n)
        return os.path.exists(p) and self.source_mtime is not None \
            and os.path.getmtime(p) >= self.source_mtime

    def request(self, n, first=False):
        if self.single or n not in self.entries or self.ready(n) or n in self.errors:
            return
        with self.cond:
            if n in self.queue:
                if not first:
                    return
                self.queue.remove(n)
            if n != self.busy:
                self.queue.insert(0, n) if first else self.queue.append(n)
            self.cond.notify()
            if self.thread is None:
                self.thread = threading.Thread(target=self._run, daemon=True)
                self.thread.start()

    def _run(self):
        while True:
            with self.cond:
                while not self.queue:
                    self.cond.wait()
                n = self.queue.pop(0)
                self.busy = n
            try:
                if not self.ready(n):
                    self._compile(n)
            except Exception as e:
                self.errors[n] = str(e)
            finally:
                self.busy = None

    def _compile(self, n):
        name = self.entries[n]
        zip_path = self.source
        os.makedirs(self.out_dir, exist_ok=True)
        work = self.work_dir
        if work is None or not os.path.isdir(work):
            work = tempfile.mkdtemp(prefix="solutions_")
            with zipfile.ZipFile(zip_path) as z:
                z.extractall(work)     # the other files of the zip may be needed (images...)
            self.work_dir = work
        src = os.path.join(work, name)
        out = self.pdf_path(n)
        if name.lower().endswith(".pdf"):
            self._publish(src, out)
            return
        if not self.pdflatex:
            raise RuntimeError("pdflatex not found: install a LaTeX distribution "
                               "(e.g. MacTeX or TeX Live) to compile the solutions.")
        d, base = os.path.split(src)
        job = os.path.splitext(base)[0]
        r = subprocess.run([self.pdflatex, "-interaction=nonstopmode", "-halt-on-error",
                            "-no-shell-escape", base],
                           cwd=d, capture_output=True, text=True, errors="replace", timeout=180)
        pdf = os.path.join(d, job + ".pdf")
        if r.returncode != 0 or not os.path.exists(pdf):
            err = [l for l in r.stdout.splitlines() if l.startswith("!")]
            raise RuntimeError(f"LaTeX error in {base}:\n" + ("\n".join(err[:3]) or r.stdout[-500:]))
        self._publish(pdf, out)

    @staticmethod
    def _publish(src, out):
        """Copies src to out atomically: a PDF that exists is always complete."""
        tmp = out + ".tmp"
        shutil.copyfile(src, tmp)
        os.chmod(tmp, 0o644)
        os.replace(tmp, out)


class SolutionsWindow:
    """Second window with the solutions of the exam being graded (follows the main window)."""

    def __init__(self, app):
        self.app = app
        self.compiler = app.sol_compiler
        self.n = None
        self.images = []          # pages shown (PIL)
        self.images_for = None
        self.tk_imgs = []
        self.zoom = 1.0           # relative to the window width
        self.prefetched = False

        self.win = tk.Toplevel(app.root)
        self.win.title("Solutions")
        self.win.geometry(app.config.get("solutions_geometry", "800x900+60+60"))
        self.win.protocol("WM_DELETE_WINDOW", self.hide)
        tb = tk.Frame(self.win)
        tb.pack(fill="x", padx=5, pady=5)
        self.label = tk.Label(tb, font=("Arial", 13, "bold"), anchor="w")
        self.label.pack(side=tk.LEFT)
        tk.Button(tb, text="Fit", command=self.zoom_reset).pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="Zoom −", command=lambda: self.set_zoom(self.zoom / 1.2)).pack(side=tk.RIGHT, padx=2)
        tk.Button(tb, text="Zoom +", command=lambda: self.set_zoom(self.zoom * 1.2)).pack(side=tk.RIGHT, padx=2)
        cf = tk.Frame(self.win)
        cf.pack(fill="both", expand=True)
        cf.rowconfigure(0, weight=1)
        cf.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(cf, bg="gray", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vs = tk.Scrollbar(cf, orient="vertical", command=self.canvas.yview)
        vs.grid(row=0, column=1, sticky="ns")
        hs = tk.Scrollbar(cf, orient="horizontal", command=self.canvas.xview)
        hs.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(3, "units"))
        self.canvas.bind("<Configure>", self.on_resize)
        self.resize_job = None
        self.poll_job = None
        for mod in ("Control", "Command"):
            self.win.bind(f"<{mod}-plus>", lambda e: self.set_zoom(self.zoom * 1.2))
            self.win.bind(f"<{mod}-equal>", lambda e: self.set_zoom(self.zoom * 1.2))
            self.win.bind(f"<{mod}-minus>", lambda e: self.set_zoom(self.zoom / 1.2))

    # ---- window ----

    def remember_geometry(self):
        try:
            if self.win.winfo_viewable():
                self.app.config["solutions_geometry"] = self.win.geometry()
                save_config(self.app.config)
        except tk.TclError:
            pass

    def hide(self):
        self.remember_geometry()
        self.win.withdraw()

    def visible(self):
        return self.win.winfo_viewable()

    def open(self, n):
        self.win.deiconify()
        self.win.lift()
        self.show(n)

    # ---- content ----

    def message(self, text):
        self.images, self.images_for, self.tk_imgs = [], None, []
        self.canvas.delete("all")
        self.canvas.update_idletasks()
        w = max(self.canvas.winfo_width(), 300)
        self.canvas.create_text(w / 2, 60, text=text, width=w - 40, fill="white",
                                font=("Arial", 13), justify="center")
        self.canvas.config(scrollregion=(0, 0, w, 200))

    def show(self, n):
        self.n = n
        if self.poll_job:
            self.win.after_cancel(self.poll_job)
            self.poll_job = None
        self.label.config(text=f"Solutions — exam {n}")
        c = self.compiler
        before = (c.source, c.source_mtime)
        if not c.refresh():
            self.message("No solutions in this folder.\n"
                         "Use the «Solutions…» button to choose them.")
            return
        if (c.source, c.source_mtime) != before:
            self.prefetched = False      # new solutions: compile everything again
        if not c.has(n):
            self.message(f"The zip has no file for exam {n}.")
            return
        if c.ready(n):
            key = (c.pdf_path(n), c.source_mtime)
            if self.images_for != key:
                try:
                    self.images = convert_from_path(c.pdf_path(n), dpi=SOLUTIONS_DPI)
                    self.images_for = key
                except Exception as e:
                    self.message(f"Cannot show the solutions:\n{e}")
                    return
                self.canvas.yview_moveto(0)
            self.draw()
            self.prefetch()
            return
        if n in c.errors:
            self.message(c.errors[n])
            return
        self.message("Compiling the solutions…")
        c.request(n, first=True)
        self.poll_job = self.win.after(300, lambda: self.show(n) if self.n == n else None)

    def prefetch(self):
        """Compiles the other exams in the background, in grading order."""
        if self.prefetched or self.compiler.single:
            return
        self.prefetched = True
        start = self.app.current_index
        order = self.app.pdf_files[start:] + self.app.pdf_files[:start]
        for k, _ in order:
            self.compiler.request(k)

    def draw(self):
        if not self.images:
            return
        cw = max(self.canvas.winfo_width() - 20, 200)
        self.canvas.delete("all")
        self.tk_imgs = []
        y, maxw = 10, 0
        for img in self.images:
            w = max(1, int(cw * self.zoom))
            h = max(1, int(img.height * w / img.width))
            t = ImageTk.PhotoImage(img.resize((w, h), Image.LANCZOS))
            self.tk_imgs.append(t)
            self.canvas.create_image(10, y, anchor="nw", image=t)
            y += h + 10
            maxw = max(maxw, w)
        self.canvas.config(scrollregion=(0, 0, maxw + 20, y))

    def set_zoom(self, z):
        self.zoom = min(max(z, 0.3), 4.0)
        self.draw()

    def zoom_reset(self):
        self.set_zoom(1.0)

    def on_resize(self, e):
        if self.resize_job:
            self.win.after_cancel(self.resize_job)
        self.resize_job = self.win.after(150, self.draw)

    def on_wheel(self, e):
        step = -e.delta if abs(e.delta) < 10 else -int(e.delta / 120)
        self.canvas.yview_scroll(step, "units")


# ==================================================
# Application
# ==================================================

class PDFAnnotatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF exam grader")
        self.root.geometry("1400x900")
        self.ok = False

        self.config = load_config()
        self.fonts = _register_font()

        self.grades = None
        self.folder = ""
        self.paths = {}          # program files in the PDF folder (see FOLDER_FILES)
        self.pdf_files = []      # list of (exam_number, path)
        self.current_index = 0
        self.current_page = 1
        self.total_pages = 1
        self.zoom = None
        self.page_image = None   # rendered page (PIL, at RENDER_DPI)
        self.disp_size = (1, 1)
        self.state = {"files": {}}
        self.strokes = {}        # {"page": [[(u, v), ...], ...]} normalized 0..1
        self.current_stroke = None
        self.current_stroke_ids = []
        self.undo_stack = {}     # {"page": [snapshot of the page strokes, ...]}
        self.erasing = False
        self.solutions = None    # SolutionsWindow
        self.sol_compiler = None # SolutionsCompiler (solutions PDFs, also appended on save)
        self.tf = None           # true/false data from the CSV (see setup_tf)

        self.rubric = []         # parsed rubric (see parse_rubric)
        self.proposals = {}      # {"12": {...}} from proposals.json
        self.proposals_mtime = None
        self.proposal_shown = False   # the comment box holds a proposal

        if not self.open_grades() or not self.open_folder() or not self.setup_rubric():
            return
        if not self.solutions_path():
            self.choose_solutions(startup=True)
        self.sol_compiler = SolutionsCompiler(self.solutions_path, self.paths["compiled"])
        self.setup_tf()
        self.build_ui()
        self.current_index = self.first_ungraded_index()
        if self.solutions_path():
            self.solutions = SolutionsWindow(self)
        self.load_pdf()
        self.ok = True

    # --------------------------------------------------
    # Setup
    # --------------------------------------------------

    def open_grades(self):
        last = self.config.get("grades_file", self.config.get("excel"))
        path = filedialog.askopenfilename(
            title="Choose the file with the student list (Excel or CSV)",
            initialdir=os.path.dirname(last) if last else None,
            filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")])
        if not path:
            return False
        try:
            if path.lower().endswith(".csv"):
                self.grades = CsvGrades(path)
            else:
                self.grades = ExcelGrades(path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot read the file:\n{e}")
            return False

        if self.grades.questions_from_rubric:
            msg = "The exercises and their maximum scores are defined in the grading rubric."
        else:
            qs = ", ".join(
                f"{q[0]} (max {fmt_num(q[2])}"
                + (f", min {fmt_num(self.grades.minimums[q[0]])})"
                   if self.grades.minimums[q[0]] else ")")
                for q in self.grades.questions)
            msg = f"Questions to grade:\n{qs}"
        msg += f"\n\nStudents found: {len(self.grades.rows)}"
        if self.grades.duplicates:
            msg += f"\n\nWARNING: duplicate exam numbers: {self.grades.duplicates}"
        if self.grades.lock_file_exists():
            msg += ("\n\nWARNING: the file seems to be open in Excel. Close it before "
                    "saving, otherwise Excel may overwrite the scores.")
        if not messagebox.askokcancel(f"{self.grades.kind} file", msg):
            return False
        self.config["grades_file"] = path
        self.config.pop("excel", None)
        save_config(self.config)
        return True

    def open_folder(self):
        folder = filedialog.askdirectory(
            title="Choose the folder with the exam PDFs",
            initialdir=self.config.get("folder"))
        if not folder:
            return False
        files, no_number, not_in_list, seen = [], [], [], {}
        for f in os.listdir(folder):
            if not f.lower().endswith(".pdf") or f.startswith("."):
                continue
            if f == FOLDER_FILES["solutions_pdf"][0]:
                continue                          # the solutions for everybody, not an exam
            n = exam_number(f)
            if n is None:
                no_number.append(f)
            elif n not in self.grades.rows:
                not_in_list.append(f)
            elif n in seen:
                messagebox.showerror("Error", f"Two PDFs for exam {n}:\n{seen[n]}\n{f}")
                return False
            else:
                seen[n] = f
                files.append((n, os.path.join(folder, f)))
        if not files:
            messagebox.showerror("Error", "No PDF matches an exam of the list.")
            return False
        files.sort()

        warn = []
        if no_number:
            warn.append("PDFs without a number in the name (ignored):\n  " + "\n  ".join(no_number))
        if not_in_list:
            warn.append("PDFs whose number is not in the list (ignored):\n  "
                        + "\n  ".join(not_in_list))
        missing = sorted(set(self.grades.rows) - set(seen))
        if missing:
            warn.append(f"Exams in the list without a PDF: {missing}")
        if warn:
            messagebox.showwarning("Warning", "\n\n".join(warn))

        self.folder = folder
        self.paths = {k: folder_path(folder, k) for k in FOLDER_FILES}
        self.pdf_files = files
        self.config["folder"] = folder
        save_config(self.config)
        self.load_state()
        return True

    def first_ungraded_index(self):
        for i, (n, _) in enumerate(self.pdf_files):
            if not self.grades.is_graded(n):
                return i
        return 0

    # --------------------------------------------------
    # Rubric
    # --------------------------------------------------

    def read_rubric_text(self):
        try:
            with open(self.paths["rubric"], "r", encoding="utf-8") as f:
                return f.read()
        except FileNotFoundError:
            return None

    def setup_rubric(self):
        """Loads the rubric; if missing, asks for it (required for a CSV file)."""
        text = self.read_rubric_text()
        if text is not None:
            try:
                rubric = parse_rubric(text)
                problem = self.check_rubric(rubric)
            except ValueError as e:
                problem = str(e)
            if not problem:
                self.apply_rubric(rubric)
                return True
            messagebox.showwarning("Grading rubric",
                                   f"{os.path.basename(self.paths['rubric'])} needs fixing:\n{problem}")
        elif not self.grades.questions_from_rubric:
            if not messagebox.askyesno(
                    "Grading rubric",
                    "There is no grading rubric in the folder yet.\n\n"
                    "Write it now? (it is used for the grading proposals; "
                    "you can also do it later with the «Rubric…» button)"):
                return True
        return self.edit_rubric(startup=True)

    def check_rubric(self, rubric):
        """Error message, or "" if the rubric can be used with the score file."""
        if self.grades.questions_from_rubric:
            if not rubric:
                return "Write at least one exercise, e.g. [Exercise 1]."
            empty = [e["name"] for e in rubric if e["max"] <= 0]
            if empty:
                return f"Exercises without positive points: {', '.join(empty)}"
            return ""
        known = {q[0] for q in self.grades.questions}
        unknown = [e["name"] for e in rubric if e["name"] not in known]
        if unknown:
            return (f"Exercises that are not columns of the Excel file: {', '.join(unknown)}\n"
                    f"Columns: {', '.join(q[0] for q in self.grades.questions)}")
        return ""

    def rubric_mismatches(self, rubric):
        """Exercises whose rubric total differs from the Excel maximum."""
        if self.grades.questions_from_rubric:
            return []
        qmax = {q[0]: q[2] for q in self.grades.questions}
        return [f"{e['name']}: rubric {fmt_num(e['max'])}, Excel {fmt_num(qmax[e['name']])}"
                for e in rubric if e["criteria"] and abs(e["max"] - qmax[e["name"]]) > 1e-9]

    def apply_rubric(self, rubric):
        self.rubric = rubric
        if self.grades.questions_from_rubric:
            self.grades.set_questions([(e["name"], e["max"]) for e in rubric])

    def edit_rubric(self, startup=False):
        """Rubric editor. At startup returns False if cancelled and required."""
        text = self.read_rubric_text()
        if text is None:
            if self.grades.questions_from_rubric:
                text = rubric_template([("Exercise 1", 0), ("Exercise 2", 0)])
            else:
                text = rubric_template([(q[0], q[2]) for q in self.grades.questions])

        win = tk.Toplevel(self.root)
        win.title("Grading rubric")
        win.geometry("760x600")
        win.transient(self.root)
        tk.Label(win, text=RUBRIC_HELP, justify="left", anchor="w",
                 font=("Arial", 10)).pack(fill="x", padx=10, pady=(10, 5))
        box = scrolledtext.ScrolledText(win, font=("Menlo", 12), wrap="word", undo=True)
        box.pack(fill="both", expand=True, padx=10, pady=5)
        box.insert("1.0", text)
        result = {"ok": False}

        def save():
            try:
                rubric = parse_rubric(box.get("1.0", tk.END))
            except ValueError as e:
                messagebox.showerror("Rubric", str(e), parent=win)
                return
            problem = self.check_rubric(rubric)
            if problem:
                messagebox.showerror("Rubric", problem, parent=win)
                return
            mism = self.rubric_mismatches(rubric)
            if mism and not messagebox.askyesno(
                    "Rubric", "The sum of the points differs from the maximum:\n"
                    + "\n".join(mism) + "\n\nSave anyway?", parent=win):
                return
            with open(self.paths["rubric"], "w", encoding="utf-8") as f:
                f.write(box.get("1.0", "end-1c").rstrip() + "\n")
            values = self.current_entry_values() if hasattr(self, "score_frame") else None
            self.apply_rubric(rubric)
            if values is not None:
                self.build_score_rows(values)
            result["ok"] = True
            win.destroy()

        bf = tk.Frame(win)
        bf.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(bf, text="Save", width=12, bg="lightgreen", command=save).pack(side=tk.RIGHT)
        tk.Button(bf, text="Cancel", width=12, command=win.destroy).pack(side=tk.RIGHT, padx=5)
        win.grab_set()
        box.focus_set()
        self.root.wait_window(win)
        if startup and not result["ok"] and self.grades.questions_from_rubric:
            return False
        return True

    # --------------------------------------------------
    # Solutions (zip with one file per exam, or one PDF for everybody)
    # --------------------------------------------------

    def solutions_path(self):
        """The most recently chosen solutions (zip or PDF) in the folder, or None."""
        found = [p for p in (self.paths["solutions_zip"], self.paths["solutions_pdf"])
                 if os.path.exists(p)]
        return max(found, key=os.path.getmtime) if found else None

    def choose_solutions(self, startup=False):
        """Copies the solutions (zip or PDF) into the PDF folder."""
        if startup and not messagebox.askyesno(
                "Solutions",
                "Do you want to choose the solutions now?\n\n"
                "Either a zip with one file per exam (.tex or .pdf, named after the exam "
                "number) or a single PDF for everybody. They are shown in a second window, "
                "can be appended to each student's PDF and are read by Claude for the "
                "grading proposals.\n\n(You can also do it later with «Solutions…».)"):
            return
        path = filedialog.askopenfilename(
            title="Choose the solutions (zip or PDF)",
            initialdir=self.folder,
            filetypes=[("Zip or PDF", "*.zip *.pdf"), ("All files", "*.*")])
        if not path:
            return
        dest = self.paths["solutions_pdf" if path.lower().endswith(".pdf") else "solutions_zip"]
        if os.path.abspath(path) != os.path.abspath(dest):
            shutil.copy2(path, dest)
        os.utime(dest)   # newest: wins over the other kind, and over PDFs compiled from an old zip
        messagebox.showinfo(
            "Solutions",
            f"Solutions copied to:\n{dest}\n\n"
            "If you use Claude Code for the grading proposals, ask Claude to prepare "
            f"them for this folder:\n{self.folder}\nThey appear here as soon as they are ready.")
        if hasattr(self, "label_proposal") and self.pdf_files:
            self.load_pdf()
            self.show_solutions()

    def show_solutions(self):
        if not self.pdf_files:
            return
        if self.solutions is None:
            self.solutions = SolutionsWindow(self)
        self.solutions.open(self.current_number())

    # --------------------------------------------------
    # Extra pages of the PDF: true/false page and solutions
    # --------------------------------------------------

    def setup_tf(self):
        """True/false data from the CSV: columns saved for this folder, otherwise
        guessed from the names. Nothing for an Excel file."""
        self.tf = None
        if not isinstance(self.grades, CsvGrades):
            return
        header = self.grades.header
        m = self._tf_columns_from_state(header) or guess_tf_columns(header)
        if check_tf_columns(m) is None:
            self._set_tf(m)

    def _set_tf(self, m):
        students, questions, max_points = tf_students(self.grades.header, self.grades.data[1:], m)
        self.tf = {"columns": m, "students": students, "questions": questions, "max": max_points}

    def _tf_columns_from_state(self, header):
        saved = self.state.get("tf_columns")
        if not saved and self.state.get("vf_columns"):     # saved by an older version
            saved = {_OLD_TF_KEYS.get(k, k): v for k, v in self.state["vf_columns"].items()}
        if not saved:
            return None
        try:
            idx = lambda name: None if name is None else header.index(name)
            m = {k: idx(saved.get(k)) for k, _, _ in TF_FIELDS}
            m["answers"] = [header.index(x) for x in saved.get("answers", [])]
            m["points"] = [header.index(x) for x in saved.get("points", [])]
        except ValueError:            # a saved column is no longer in the file
            return None
        return m

    def choose_tf(self):
        header = self.grades.header
        guess = self.tf["columns"] if self.tf else guess_tf_columns(header)
        m = choose_tf_columns(self.root, header, self.grades.data[1:], guess)
        if m is None:
            return
        try:
            self._set_tf(m)
        except Exception as e:
            messagebox.showerror("True/false columns", str(e))
            return
        name = lambda i: None if i is None else header[i]
        saved = {k: name(m.get(k)) for k, _, _ in TF_FIELDS}
        saved["answers"] = [header[i] for i in m["answers"]]
        saved["points"] = [header[i] for i in m["points"]]
        self.state["tf_columns"] = saved
        self.state.pop("vf_columns", None)
        self.opt_tf.set(True)
        self.save_options()

    def options(self):
        """Folder options saved in the state (older versions used the key "vf")."""
        o = self.state.get("options", {})
        return {"true_false": o.get("true_false", o.get("vf", True)),
                "solutions": o.get("solutions", True),
                "english_default": o.get("english_default", False)}

    def save_options(self):
        self.state["options"] = {"true_false": self.opt_tf.get(),
                                 "solutions": self.opt_solutions.get(),
                                 "english_default": self.english_default.get()}
        self.save_state()

    def solution_pdf(self, n):
        """(path of the solutions PDF of exam n, None) or (None, reason). Compiles it now if needed."""
        c = self.sol_compiler
        if not c.refresh():
            return None, "no solutions chosen («Solutions…»)"
        if not c.has(n):
            return None, f"the zip has no solution for exam {n}"
        if not c.ready(n) and n not in c.errors:
            c.request(n, first=True)
            self.root.config(cursor="watch")
            self.root.update_idletasks()
            t = time.time()
            while not c.ready(n) and n not in c.errors and time.time() - t < 180:
                time.sleep(0.1)
            self.root.config(cursor="")
        if c.ready(n):
            return c.pdf_path(n), None
        return None, c.errors.get(n, "compilation took too long")

    def extra_pages(self, n, english):
        """(true/false page or None, solutions PDF or None, notes on what is missing)."""
        first = solutions = None
        notes = []
        if self.opt_tf.get():
            if self.tf is None:
                notes.append("true/false page: choose the columns with «T/F columns…»")
            elif n not in self.tf["students"]:
                notes.append(f"true/false page: exam {n} is not in the CSV")
            else:
                first = true_false_page(n, self.tf["students"][n], self.tf["questions"],
                                        self.tf["max"], self.fonts, english)
        if self.opt_solutions.get():
            solutions, why = self.solution_pdf(n)
            if why:
                notes.append("solutions: " + why)
        return first, solutions, notes

    def exam_english(self, saved):
        """Language of the student pages of an exam: as saved, else the folder default."""
        return bool(saved.get("english", self.english_default.get()))

    def regenerate_all(self):
        """Rebuilds the PDF of every graded exam with the current options."""
        self.remember_current()
        todo = [(k, p) for k, p in self.pdf_files
                if self.grades.is_graded(k) and os.path.exists(self.original_path(p))]
        if not todo:
            messagebox.showinfo("Rebuild PDFs", "No exam graded yet.")
            return
        yes_no = lambda v: "yes" if v else "no"
        if not messagebox.askyesno(
                "Rebuild PDFs",
                f"Rebuild the PDF of the {len(todo)} graded exams with the current options:\n"
                f"  true/false page: {yes_no(self.opt_tf.get())}\n"
                f"  solutions appended: {yes_no(self.opt_solutions.get())}\n\n"
                "Scores, comments, drawings and language stay as saved."):
            return
        problems = []
        for i, (k, p) in enumerate(todo, start=1):
            self.label_status.config(text=f"Rebuilding the PDFs: {i}/{len(todo)}…")
            self.root.update_idletasks()
            saved = self.state["files"].get(os.path.basename(p), {})
            english = self.exam_english(saved)
            first, solutions, notes = self.extra_pages(k, english)
            problems += [f"{k}: {x}" for x in notes]
            try:
                build_graded_pdf(self.original_path(p), p, saved.get("strokes", {}),
                                 self.grades.student(k), self.grades.questions,
                                 self.grades.get_scores(k), saved.get("comments", ""),
                                 self.fonts, english=english, first_page=first,
                                 solutions=solutions)
            except Exception as e:
                problems.append(f"{k}: {e}")
        self.load_pdf()
        report = f"PDFs rebuilt: {len(todo)}."
        if problems:
            report += "\n\nProblems:\n" + "\n".join(problems[:30])
            if len(problems) > 30:
                report += f"\n… and {len(problems) - 30} more"
        (messagebox.showwarning if problems else messagebox.showinfo)("Rebuild PDFs", report)

    # --------------------------------------------------
    # Proposals (proposals.json, e.g. written by Claude)
    # --------------------------------------------------

    def reload_proposals(self):
        """Re-reads the proposals if they changed (they may be written while grading)."""
        path = self.paths["proposals"]
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            self.proposals, self.proposals_mtime = {}, None
            return
        if mtime == self.proposals_mtime:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.proposals = {str(k): v for k, v in data.items() if isinstance(v, dict)}
            self.proposals_mtime = mtime
        except Exception as e:
            messagebox.showwarning("Proposals", f"Cannot read {os.path.basename(path)}:\n{e}")
            self.proposals, self.proposals_mtime = {}, mtime

    def proposal(self, n):
        return self.proposals.get(str(n))

    @staticmethod
    def proposal_comment(prop, english):
        """Proposed comment for the English or Italian page ("comment" fits both;
        "commento" is the key used by older versions)."""
        if not prop:
            return ""
        keys = ("comment_en", "comment") if english else ("comment_it", "commento", "comment")
        return str(next((prop[k] for k in keys if prop.get(k)), "")).strip()

    @staticmethod
    def proposal_scores(prop):
        return (prop or {}).get("scores") or (prop or {}).get("punteggi") or {}

    def proposal_score(self, prop, qname):
        """Proposed score for a question, if valid."""
        v = self.proposal_scores(prop).get(qname)
        if isinstance(v, str):
            v = to_float(v)
        if not _is_number(v):
            return None
        q = next(q for q in self.grades.questions if q[0] == qname)
        return float(v) if self.grades.minimums[qname] <= v <= q[2] else None

    # --------------------------------------------------
    # UI
    # --------------------------------------------------

    def build_ui(self):
        top = tk.Frame(self.root)
        top.pack(fill="x", padx=5, pady=5)
        self.label_file = tk.Label(top, font=("Arial", 13, "bold"), anchor="w")
        self.label_file.pack(side=tk.LEFT)
        self.label_status = tk.Label(top, font=("Arial", 10), fg="blue")
        self.label_status.pack(side=tk.RIGHT)

        main = tk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=5, pady=5)

        # ---- PDF ----
        left = tk.Frame(main)
        left.pack(side=tk.LEFT, fill="both", expand=True)
        tb = tk.Frame(left)
        tb.pack(fill="x", pady=(0, 5))
        tk.Button(tb, text="◀ Page", command=self.prev_page).pack(side=tk.LEFT, padx=2)
        self.page_label = tk.Label(tb, text="1/1", width=8)
        self.page_label.pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Page ▶", command=self.next_page).pack(side=tk.LEFT, padx=2)
        tk.Label(tb, text="|").pack(side=tk.LEFT, padx=5)
        tk.Button(tb, text="Zoom +", command=self.zoom_in).pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Zoom −", command=self.zoom_out).pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Fit", command=self.zoom_reset).pack(side=tk.LEFT, padx=2)
        tk.Label(tb, text="|").pack(side=tk.LEFT, padx=5)
        self.tool_var = tk.StringVar(value="pen")
        tk.Radiobutton(tb, text="✎ Pen", variable=self.tool_var, value="pen",
                       indicatoron=False, padx=6).pack(side=tk.LEFT, padx=2)
        tk.Radiobutton(tb, text="⌫ Eraser", variable=self.tool_var, value="eraser",
                       indicatoron=False, padx=6).pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Undo", command=self.undo).pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Clear page", command=self.clear_annotations).pack(side=tk.LEFT, padx=2)
        tk.Label(tb, text="|").pack(side=tk.LEFT, padx=5)
        tk.Button(tb, text="Solutions ⧉", command=self.show_solutions).pack(side=tk.LEFT, padx=2)

        cf = tk.Frame(left, relief="sunken", borderwidth=2)
        cf.pack(fill="both", expand=True)
        cf.rowconfigure(0, weight=1)
        cf.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(cf, bg="gray", cursor="cross", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vs = tk.Scrollbar(cf, orient="vertical", command=self.canvas.yview)
        vs.grid(row=0, column=1, sticky="ns")
        hs = tk.Scrollbar(cf, orient="horizontal", command=self.canvas.xview)
        hs.grid(row=1, column=0, sticky="ew")
        self.canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)

        # Pen tip / left button: draw (or erase with the eraser tool).
        # Pen side button / right button: erase.
        self.canvas.bind("<ButtonPress-1>", self.start_draw)
        self.canvas.bind("<B1-Motion>", self.draw_motion)
        self.canvas.bind("<ButtonRelease-1>", self.end_draw)
        for b in (2, 3):
            self.canvas.bind(f"<ButtonPress-{b}>", self.start_erase)
            self.canvas.bind(f"<B{b}-Motion>", self.erase_at)
            self.canvas.bind(f"<ButtonRelease-{b}>", self.end_erase)
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self.on_wheel_h)
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(3, "units"))

        # ---- Grading ----
        right = tk.Frame(main, width=420)
        right.pack(side=tk.RIGHT, fill="y", padx=(10, 0))
        right.pack_propagate(False)

        self.label_student = tk.Label(right, font=("Arial", 11), justify="left",
                                      anchor="w", wraplength=400)
        self.label_student.pack(fill="x", padx=5, pady=(0, 5))

        options = self.options()
        opts = tk.Frame(right)
        opts.pack(fill="x", padx=5)
        tk.Button(opts, text="Rubric…", command=self.edit_rubric).pack(side=tk.LEFT, padx=2)
        tk.Button(opts, text="Solutions…", command=self.choose_solutions).pack(side=tk.LEFT, padx=2)
        self.english_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opts, text="English", variable=self.english_var,
                       command=self.toggle_english,
                       font=("Arial", 11, "bold")).pack(side=tk.RIGHT, padx=2)
        self.english_default = tk.BooleanVar(value=options["english_default"])
        tk.Checkbutton(opts, text="English by default", variable=self.english_default,
                       command=self.toggle_english_default).pack(side=tk.RIGHT)

        # What goes into the PDF besides the exam and the evaluation page
        opts2 = tk.Frame(right)
        opts2.pack(fill="x", padx=5, pady=(3, 0))
        is_csv = isinstance(self.grades, CsvGrades)
        self.opt_tf = tk.BooleanVar(value=is_csv and options["true_false"])
        self.opt_solutions = tk.BooleanVar(value=options["solutions"])
        tk.Checkbutton(opts2, text="T/F page", variable=self.opt_tf, command=self.save_options,
                       state="normal" if is_csv else "disabled").pack(side=tk.LEFT)
        tk.Button(opts2, text="T/F columns…", command=self.choose_tf,
                  state="normal" if is_csv else "disabled").pack(side=tk.LEFT, padx=(0, 6))
        tk.Checkbutton(opts2, text="Append solutions", variable=self.opt_solutions,
                       command=self.save_options).pack(side=tk.LEFT)
        tk.Button(opts2, text="Rebuild PDFs…", command=self.regenerate_all).pack(side=tk.RIGHT)

        self.score_frame = tk.LabelFrame(right, text="Scores", font=("Arial", 10, "bold"))
        self.score_frame.pack(fill="x", padx=5, pady=5)
        self.score_entries = []
        self.total_label = None
        self.build_score_rows()

        # Note from the proposal (doubts, unreadable parts...): not written anywhere
        self.label_proposal = tk.Label(right, font=("Arial", 10), fg="#a15c00",
                                       justify="left", anchor="w", wraplength=400)
        self.label_proposal.pack(fill="x", padx=5)

        nav = tk.Frame(right)
        nav.pack(fill="x", padx=5, pady=5)
        tk.Button(nav, text="◀ Previous", command=self.previous_student,
                  width=14).pack(side=tk.LEFT, padx=2)
        tk.Button(nav, text="Save and continue ▶", command=self.save_and_next,
                  width=18, bg="lightblue").pack(side=tk.RIGHT, padx=2)
        nav2 = tk.Frame(right)
        nav2.pack(fill="x", padx=5)
        tk.Button(nav2, text="Go to exam…", command=self.goto_exam,
                  width=14).pack(side=tk.LEFT, padx=2)
        tk.Button(nav2, text="Next (without saving) ▶", command=self.skip_student,
                  width=22).pack(side=tk.RIGHT, padx=2)
        tk.Button(right, text="Save and exit", command=self.save_and_exit,
                  width=18, bg="lightgreen").pack(pady=5)

        tk.Label(right, text="Comments:", font=("Arial", 10, "bold")).pack(anchor="w", padx=5, pady=(10, 0))
        self.text_box = scrolledtext.ScrolledText(right, width=45, height=9,
                                                  font=("Arial", 11), wrap="word", undo=True)
        self.text_box.pack(fill="both", expand=True, padx=5, pady=5)

        cbf = tk.Frame(right)
        cbf.pack(fill="x", padx=5, pady=(0, 5))
        for i in range(len(LABELS[True]["quick"])):
            tk.Button(cbf, text=LABELS[True]["quick"][i][:20],
                      command=lambda i=i: self.insert_comment(
                          LABELS[self.english_var.get()]["quick"][i])).pack(side=tk.LEFT, padx=2)
        tk.Button(cbf, text="Copy previous comment",
                  command=self.copy_previous_comment).pack(side=tk.LEFT, padx=2)

        # Phrases already used in other exams (double-click to insert)
        tk.Label(right, text="Phrases already used (double-click to insert):",
                 font=("Arial", 10, "bold")).pack(anchor="w", padx=5, pady=(5, 0))
        self.phrase_filter = tk.Entry(right, font=("Arial", 10))
        self.phrase_filter.pack(fill="x", padx=5, pady=(2, 2))
        self.phrase_filter.bind("<KeyRelease>", lambda e: self.refresh_phrases())
        lf = tk.Frame(right)
        lf.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        self.phrase_list = tk.Listbox(lf, font=("Arial", 10), height=8, activestyle="none")
        sb = tk.Scrollbar(lf, orient="vertical", command=self.phrase_list.yview)
        self.phrase_list.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill="y")
        self.phrase_list.pack(side=tk.LEFT, fill="both", expand=True)
        self.phrase_list.bind("<Double-Button-1>", self.insert_selected_phrase)
        self.phrase_list.bind("<Return>", self.insert_selected_phrase)
        self.phrases = []
        self.previous_comment = ""

    def build_score_rows(self, values=None):
        """(Re)builds the score boxes; values = {question: (text, is_proposal)} to keep."""
        for w in self.score_frame.winfo_children():
            w.destroy()
        self.score_entries = []
        for i, (qname, _, qmax) in enumerate(self.grades.questions):
            row = tk.Frame(self.score_frame)
            row.pack(fill="x", padx=10, pady=2)
            qmin = self.grades.minimums[qname]
            rng = f"max {fmt_num(qmax)}" if qmin == 0 else f"{fmt_num(qmin)} to {fmt_num(qmax)}"
            tk.Label(row, text=f"{qname}  ({rng}):", width=18,
                     anchor="w").pack(side=tk.LEFT)
            e = tk.Entry(row, width=8, font=("Arial", 11))
            e.pack(side=tk.LEFT)
            e.default_bg = e.cget("background")
            e.bind("<Return>", lambda ev, idx=i: self.focus_next_entry(idx))
            e.bind("<KeyRelease>", self.on_score_key)
            if values and values.get(qname):
                e.insert(0, values[qname][0])
                if values[qname][1]:
                    e.config(background=PROPOSAL_BG)
            self.score_entries.append(e)
        trow = tk.Frame(self.score_frame)
        trow.pack(fill="x", padx=10, pady=(6, 4))
        tk.Label(trow, text="Total:", width=18, anchor="w",
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.total_label = tk.Label(trow, text="0", font=("Arial", 11, "bold"), fg="green")
        self.total_label.pack(side=tk.LEFT)
        tk.Label(self.score_frame, text="Yellow: proposal to be checked",
                 bg=PROPOSAL_BG, font=("Arial", 9)).pack(anchor="w", padx=10, pady=(0, 4))
        self.calculate_total()

    def current_entry_values(self):
        """{question: (text, is_proposal)} of the score boxes."""
        return {q[0]: (e.get(), e.cget("background") == PROPOSAL_BG)
                for q, e in zip(self.grades.questions, self.score_entries)}

    def on_score_key(self, event):
        # A score typed or edited by hand is no longer a proposal
        if event.keysym not in ("Return", "Tab", "Shift_L", "Shift_R"):
            event.widget.config(background=event.widget.default_bg)
        self.calculate_total()

    def toggle_english(self):
        """Switches the language of the student pages and of the proposed comment."""
        english = self.english_var.get()
        prop = self.proposal(self.current_number())
        new = self.proposal_comment(prop, english)
        old = self.proposal_comment(prop, not english)
        current = self.text_box.get("1.0", tk.END).strip()
        language = "English" if english else "Italian"
        if new and current != new:
            if (not current or current == old or messagebox.askyesno(
                    language, f"Replace the comment with the {language} proposal?\n"
                              "(your changes to the comment will be lost)")):
                self.text_box.delete("1.0", tk.END)
                self.text_box.insert("1.0", new)
                self.proposal_shown = True
        elif english and not new and prop:
            messagebox.showinfo("English", "There is no English comment proposal for this "
                                "exam: please write it yourself.")
        name = os.path.basename(self.current_path())
        self.state["files"].setdefault(name, {})["english"] = english
        self.save_state()

    def toggle_english_default(self):
        """Language of the exams not opened yet; also the current one if never set."""
        self.save_options()
        name = os.path.basename(self.current_path())
        if "english" not in self.state["files"].get(name, {}) \
                and self.english_var.get() != self.english_default.get():
            self.english_var.set(self.english_default.get())
            self.toggle_english()

    # --------------------------------------------------
    # State (drawings and comments, per file, in the PDF folder)
    # --------------------------------------------------

    def load_state(self):
        try:
            with open(self.paths["state"], "r", encoding="utf-8") as f:
                self.state = json.load(f)
            self.state.setdefault("files", {})
        except Exception:
            self.state = {"files": {}}

    def save_state(self):
        try:
            fd, tmp = tempfile.mkstemp(suffix=".json", dir=self.folder)
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(self.state, f)
            os.replace(tmp, self.paths["state"])
        except Exception as e:
            print(f"Cannot save the state: {e}")

    def remember_current(self):
        """Stores drawings and comments of the current exam (not the scores)."""
        if not self.pdf_files:
            return
        name = os.path.basename(self.current_path())
        entry = self.state["files"].setdefault(name, {})
        entry["strokes"] = self.strokes
        entry["comments"] = self.text_box.get("1.0", tk.END).strip()
        entry["english"] = self.english_var.get()
        if self.proposal_shown:
            entry["proposal_seen"] = True   # from now on keep your comment, even if empty
        self.save_state()

    # --------------------------------------------------
    # Current files
    # --------------------------------------------------

    def current_number(self):
        return self.pdf_files[self.current_index][0]

    def current_path(self):
        return self.pdf_files[self.current_index][1]

    def original_path(self, path):
        return os.path.join(self.paths["originals"], os.path.basename(path))

    def source_path(self, path):
        """PDF to display/process: the original if already saved, otherwise the file itself."""
        orig = self.original_path(path)
        return orig if os.path.exists(orig) else path

    def load_pdf(self):
        path = self.current_path()
        n = self.current_number()
        name = os.path.basename(path)
        saved = self.state["files"].get(name, {})
        self.strokes = {k: [list(map(tuple, s)) for s in v]
                        for k, v in saved.get("strokes", {}).items()}
        self.undo_stack = {}

        english = self.exam_english(saved)
        self.english_var.set(english)
        self.reload_proposals()
        graded = self.grades.is_graded(n)
        # Proposals are offered only for exams not graded yet
        prop = None if graded else self.proposal(n)

        self.text_box.delete("1.0", tk.END)
        self.proposal_shown = False
        if saved.get("comments") or saved.get("proposal_seen"):
            self.text_box.insert("1.0", saved.get("comments", ""))
        elif self.proposal_comment(prop, english):
            self.text_box.insert("1.0", self.proposal_comment(prop, english))
            self.proposal_shown = True
        self.text_box.edit_reset()
        self.phrase_filter.delete(0, tk.END)
        self.refresh_phrases()

        scores = self.grades.get_scores(n)
        for e, s, q in zip(self.score_entries, scores, self.grades.questions):
            e.delete(0, tk.END)
            e.config(background=e.default_bg)
            if s is not None:
                e.insert(0, fmt_num(s))
            elif prop is not None and self.proposal_score(prop, q[0]) is not None:
                e.insert(0, fmt_num(self.proposal_score(prop, q[0])))
                e.config(background=PROPOSAL_BG)
        self.calculate_total()

        note = ""
        if prop is not None:
            note = "Proposal: check the scores and the comment."
            proposed = self.proposal_scores(prop)
            bad = [q[0] for q in self.grades.questions
                   if q[0] in proposed and self.proposal_score(prop, q[0]) is None]
            if bad:
                note += f"\nProposed scores out of range (ignored): {', '.join(bad)}"
            doubts = prop.get("doubts") or prop.get("dubbi")
            if doubts:
                note += f"\n⚠ {doubts}"
        elif not graded and (self.proposals or os.path.exists(self.paths["proposals"])):
            note = "No proposal for this exam yet."
        self.label_proposal.config(text=note)

        st = self.grades.student(n)
        self.label_file.config(text=f"Exam {n}  —  {st['surname']} {st['name']}")
        info = f"Student ID: {st['student_id']}\nFile: {name}"
        if st["notes"]:
            info += f"\nNotes: {st['notes']}"
        if graded:
            info += f"\n(already graded: the scores come from the {self.grades.kind} file)"
        self.label_student.config(text=info)
        graded = sum(1 for k, _ in self.pdf_files if self.grades.is_graded(k))
        self.label_status.config(
            text=f"{self.current_index + 1}/{len(self.pdf_files)}  —  graded: {graded}")

        try:
            self.total_pages = len(PdfReader(self.source_path(path)).pages)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open the PDF:\n{e}")
            self.total_pages = 1
        self.current_page = 1
        self.zoom = None
        self.render_page()
        if self.score_entries:
            self.score_entries[0].focus_set()
        if self.solutions and self.solutions.visible():
            self.solutions.show(n)

    # --------------------------------------------------
    # Display
    # --------------------------------------------------

    def render_page(self):
        try:
            self.page_image = convert_from_path(
                self.source_path(self.current_path()),
                first_page=self.current_page, last_page=self.current_page,
                dpi=RENDER_DPI)[0]
        except Exception as e:
            messagebox.showerror("Error", f"Cannot show the PDF:\n{e}")
            self.page_image = Image.new("RGB", (800, 1100), "white")
        self.show_page()

    def show_page(self):
        img = self.page_image
        if img is None:
            return
        if self.zoom is None:
            self.root.update_idletasks()
            cw = self.canvas.winfo_width() - 10
            ch = self.canvas.winfo_height() - 10
            if cw < 100 or ch < 100:
                self.zoom = 0.5
            else:
                self.zoom = min(cw / img.width, ch / img.height)
        self.zoom = min(max(self.zoom, 0.1), 3.0)

        w, h = max(1, int(img.width * self.zoom)), max(1, int(img.height * self.zoom))
        self.disp_size = (w, h)
        self.tk_img = ImageTk.PhotoImage(img.resize((w, h), Image.LANCZOS))
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        self.canvas.config(scrollregion=(0, 0, w, h))
        for stroke in self.strokes.get(str(self.current_page), []):
            self._draw_stroke(stroke)
        self.page_label.config(text=f"{self.current_page}/{self.total_pages}")

    def _draw_stroke(self, stroke):
        w, h = self.disp_size
        pts = [coord for (u, v) in stroke for coord in (u * w, v * h)]
        if len(pts) < 4:
            return
        if all(p == stroke[0] for p in stroke):     # a dot
            x, y, r = pts[0], pts[1], PEN_WIDTH_SCREEN / 2 + 0.5
            self.canvas.create_oval(x - r, y - r, x + r, y + r, fill=PEN_COLOR, outline=PEN_COLOR)
        else:
            self.canvas.create_line(*pts, fill=PEN_COLOR, width=PEN_WIDTH_SCREEN,
                                    capstyle="round", joinstyle="round", smooth=False)

    def next_page(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.render_page()

    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.render_page()

    def zoom_in(self):
        self.zoom *= 1.2
        self.show_page()

    def zoom_out(self):
        self.zoom /= 1.2
        self.show_page()

    def zoom_reset(self):
        self.zoom = None
        self.show_page()

    def on_wheel(self, e):
        step = -e.delta if abs(e.delta) < 10 else -int(e.delta / 120)
        self.canvas.yview_scroll(step, "units")

    def on_wheel_h(self, e):
        step = -e.delta if abs(e.delta) < 10 else -int(e.delta / 120)
        self.canvas.xview_scroll(step, "units")

    # --------------------------------------------------
    # Drawing (coordinates stored normalized 0..1: independent of zoom)
    # --------------------------------------------------

    def _norm(self, e):
        w, h = self.disp_size
        x, y = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        return (min(max(x / w, 0.0), 1.0), min(max(y / h, 0.0), 1.0))

    def _push_undo(self):
        page = str(self.current_page)
        self.undo_stack.setdefault(page, []).append(list(self.strokes.get(page, [])))

    def start_draw(self, e):
        if self.tool_var.get() == "eraser":
            self.start_erase(e)
            return
        self.current_stroke = [self._norm(e)]
        self.current_stroke_ids = []

    def draw_motion(self, e):
        if self.erasing:
            self.erase_at(e)
            return
        if self.current_stroke is None:
            return
        w, h = self.disp_size
        u, v = self._norm(e)
        pu, pv = self.current_stroke[-1]
        if abs(u - pu) * w < 1 and abs(v - pv) * h < 1:   # tablets send many tiny moves
            return
        self.current_stroke.append((u, v))
        # Only the new segment is drawn: fast even for long strokes
        self.current_stroke_ids.append(self.canvas.create_line(
            pu * w, pv * h, u * w, v * h, fill="blue", width=PEN_WIDTH_SCREEN,
            capstyle="round"))

    def end_draw(self, e):
        if self.erasing:
            self.end_erase(e)
            return
        stroke = self.current_stroke
        self.current_stroke = None
        for i in self.current_stroke_ids:
            self.canvas.delete(i)
        self.current_stroke_ids = []
        if stroke:
            if len(stroke) == 1:          # a tap: keep it as a dot
                stroke = [stroke[0], stroke[0]]
            self._push_undo()
            self.strokes.setdefault(str(self.current_page), []).append(stroke)
            self._draw_stroke(stroke)

    def start_erase(self, e):
        self.erasing = True
        self.erase_pushed = False
        self.canvas.config(cursor="circle")
        self.erase_at(e)

    def erase_at(self, e):
        page = self.strokes.get(str(self.current_page))
        if not self.erasing or not page:
            return
        w, h = self.disp_size
        x, y = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)

        def near(stroke):
            pts = [(u * w, v * h) for u, v in stroke]
            for (x1, y1), (x2, y2) in zip(pts, pts[1:] or pts):
                dx, dy = x2 - x1, y2 - y1
                L = dx * dx + dy * dy
                t = 0 if L == 0 else max(0, min(1, ((x - x1) * dx + (y - y1) * dy) / L))
                if (x1 + t * dx - x) ** 2 + (y1 + t * dy - y) ** 2 <= ERASER_RADIUS ** 2:
                    return True
            return False

        keep = [s for s in page if not near(s)]
        if len(keep) != len(page):
            if not self.erase_pushed:
                self._push_undo()
                self.erase_pushed = True
            self.strokes[str(self.current_page)] = keep
            self.show_page()

    def end_erase(self, e):
        self.erasing = False
        self.canvas.config(cursor="cross")

    def undo(self):
        stack = self.undo_stack.get(str(self.current_page))
        if stack:
            self.strokes[str(self.current_page)] = stack.pop()
            self.show_page()

    def clear_annotations(self):
        if self.strokes.get(str(self.current_page)) and \
                messagebox.askyesno("Confirm", "Delete all drawings on this page?\n"
                                    "(you can bring them back with «Undo»)"):
            self._push_undo()
            self.strokes[str(self.current_page)] = []
            self.show_page()

    # --------------------------------------------------
    # Scores and comments
    # --------------------------------------------------

    def calculate_total(self, event=None):
        if self.total_label is None:
            return
        total = 0.0
        for e in self.score_entries:
            total += to_float(e.get() or "0") or 0.0
        self.total_label.config(text=f"{fmt_num(total)} / {fmt_num(self.grades.max_total)}")

    def focus_next_entry(self, idx):
        if idx + 1 < len(self.score_entries):
            self.score_entries[idx + 1].focus_set()
            self.score_entries[idx + 1].select_range(0, tk.END)

    def insert_comment(self, comment):
        # Start a new line if the cursor is not at the beginning of a line
        if self.text_box.index(tk.INSERT).split(".")[1] != "0":
            comment = "\n" + comment
        self.text_box.insert(tk.INSERT, comment + "\n")
        self.text_box.see(tk.INSERT)

    def copy_previous_comment(self):
        if self.previous_comment:
            self.insert_comment(self.previous_comment)
        else:
            messagebox.showinfo("Previous comment", "The previous exam has no comment.")

    def refresh_phrases(self):
        """Lists the comment lines used in other exams, most frequent first."""
        current = os.path.basename(self.current_path()) if self.pdf_files else None
        counts = {}
        for name, entry in self.state["files"].items():
            if name == current:
                continue
            for line in entry.get("comments", "").split("\n"):
                line = line.strip()
                if line:
                    counts[line] = counts.get(line, 0) + 1
        flt = self.phrase_filter.get().strip().lower()
        self.phrases = sorted((p for p in counts if flt in p.lower()),
                              key=lambda p: (-counts[p], p.lower()))
        self.phrase_list.delete(0, tk.END)
        for p in self.phrases:
            self.phrase_list.insert(tk.END, f"({counts[p]}×) {p}" if counts[p] > 1 else p)

    def insert_selected_phrase(self, event=None):
        sel = self.phrase_list.curselection()
        if sel:
            self.insert_comment(self.phrases[sel[0]])
            self.text_box.focus_set()

    def read_scores(self):
        """Validated scores, or None (after showing the error)."""
        scores = []
        for e, (qname, _, qmax) in zip(self.score_entries, self.grades.questions):
            txt = e.get().strip()
            if not txt:
                messagebox.showwarning("Error", f"Enter the score for {qname}")
                e.focus_set()
                return None
            s = to_float(txt)
            if s is None:
                messagebox.showwarning("Error", f"Invalid score for {qname}: '{txt}'")
                e.focus_set()
                return None
            qmin = self.grades.minimums[qname]
            if s < qmin or s > qmax:
                messagebox.showwarning("Error", f"{qname}: the score must be between "
                                                f"{fmt_num(qmin)} and {fmt_num(qmax)}")
                e.focus_set()
                return None
            scores.append(s)
        return scores

    # --------------------------------------------------
    # Saving
    # --------------------------------------------------

    def save_current(self):
        """Writes the scores to the Excel/CSV file and rebuilds the PDF. True on success."""
        scores = self.read_scores()
        if scores is None:
            return False
        n = self.current_number()
        path = self.current_path()
        kind = self.grades.kind

        while self.grades.lock_file_exists():
            if not messagebox.askretrycancel(
                    "File open",
                    f"The {kind} file is open in Excel. Close it (saving your own "
                    "changes, if any) and press Retry."):
                return False

        self.remember_current()
        try:
            self.grades.set_scores(n, scores)
            self.grades.save()
        except Exception as e:
            messagebox.showerror("Error", f"Cannot save the {kind} file:\n{e}")
            return False
        for e in self.score_entries:
            e.config(background=e.default_bg)

        try:
            orig = self.original_path(path)
            if not os.path.exists(orig):
                os.makedirs(os.path.dirname(orig), exist_ok=True)
                shutil.copy2(path, orig)
            english = self.english_var.get()
            first, solutions, notes = self.extra_pages(n, english)
            build_graded_pdf(orig, path, self.strokes, self.grades.student(n),
                             self.grades.questions, scores,
                             self.text_box.get("1.0", tk.END), self.fonts,
                             english=english, first_page=first, solutions=solutions)
        except Exception as e:
            messagebox.showerror("Error",
                                 f"Scores saved in the {kind} file, but the PDF could not be saved:\n{e}")
            return False
        if notes:
            messagebox.showwarning("PDF saved", "PDF saved without:\n- " + "\n- ".join(notes))
        return True

    def go_to(self, index):
        self.remember_current()
        self.previous_comment = self.text_box.get("1.0", tk.END).strip()
        self.current_index = index
        self.load_pdf()

    def save_and_next(self):
        if not self.save_current():
            return
        if self.current_index + 1 >= len(self.pdf_files):
            missing = [k for k, _ in self.pdf_files if not self.grades.is_graded(k)]
            if missing:
                messagebox.showinfo("End of the list",
                                    f"Last exam saved.\nStill to grade: {missing}")
            else:
                messagebox.showinfo("Done", "All exams have been graded!")
            self.load_pdf()
            return
        self.go_to(self.current_index + 1)

    def skip_student(self):
        if self.current_index + 1 < len(self.pdf_files):
            self.go_to(self.current_index + 1)

    def previous_student(self):
        if self.current_index > 0:
            self.go_to(self.current_index - 1)

    def goto_exam(self):
        n = simpledialog.askinteger("Go to exam", "Exam number:", parent=self.root)
        if n is None:
            return
        for i, (k, _) in enumerate(self.pdf_files):
            if k == n:
                self.go_to(i)
                return
        messagebox.showwarning("Warning", f"No PDF for exam {n}")

    def save_and_exit(self):
        ans = messagebox.askyesnocancel("Exit", "Save the current exam before exiting?")
        if ans is None:
            return
        if ans and not self.save_current():
            return
        self.remember_current()
        if self.solutions:
            self.solutions.remember_geometry()
        if self.sol_compiler:
            self.sol_compiler.cleanup()
        self.root.destroy()

    def on_closing(self):
        self.save_and_exit()


# ==================================================
# Startup
# ==================================================

def _not_typing(root):
    w = root.focus_get()
    return not isinstance(w, (tk.Entry, tk.Text))


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFAnnotatorApp(root)
    if not app.ok:
        root.destroy()
    else:
        root.protocol("WM_DELETE_WINDOW", app.on_closing)
        for mod in ("Control", "Command"):
            try:
                root.bind(f"<{mod}-s>", lambda e: app.save_and_next())
                root.bind(f"<{mod}-z>", lambda e: app.undo() if _not_typing(root) else None)
                root.bind(f"<{mod}-plus>", lambda e: app.zoom_in())
                root.bind(f"<{mod}-equal>", lambda e: app.zoom_in())
                root.bind(f"<{mod}-minus>", lambda e: app.zoom_out())
            except tk.TclError:
                pass
        root.mainloop()
