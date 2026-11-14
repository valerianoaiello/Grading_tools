"""
Grading of open-ended questions in PDF exams.

- Reads the student list and the questions from an Excel file (sheet "Punteggi"):
  row 1 = headers (Compito, Matricola, Cognome, Nome, ..., G1, G2, 7A, ...),
  row "massimo" = maximum score of each question. Every column with a numeric
  value in the "massimo" row is treated as a question to grade.
- Each PDF is matched to a student through the exam number, taken from the
  last group of digits in the file name (e.g. "compito_12.pdf" -> 12).
- Lets you draw on the pages and write comments; on save it writes the scores
  to the Excel file and regenerates the PDF (original + drawings + evaluation
  page). Originals are kept in the "originali" subfolder, so an exam can be
  regraded without duplicating anything.
- Also accepts the CSV written by "Correggi VF" (risposte_totali.csv, column
  "Numero compito"): the questions then come from the grading rubric and their
  columns are added to the CSV.
- Grading rubric (griglia.txt in the PDF folder): for each exercise, the points
  given for each part of the answer. Asked at startup if missing.
- Proposals (proposte.json in the PDF folder, written by Claude in a Claude Code
  session after reading the scans, the rubric and the LaTeX solutions): proposed
  scores and comment (Italian and English) shown highlighted, to be checked and
  edited before saving. Nothing is written until you save.
- "Inglese" switch: comment proposal and evaluation page in English.
- Drawing with mouse or pen tablet (side button / right click erases) and a second
  window with the solutions of the current exam (LaTeX compiled with pdflatex).
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
from tkinter import filedialog, messagebox, simpledialog, scrolledtext
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

STATE_NAME = ".correzione_stato.json"   # in the PDF folder
BACKUP_DIR = "originali"                # subfolder holding the original PDFs
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".examinatrix_grader.json")
RENDER_DPI = 200
PEN_COLOR = "red"
PEN_WIDTH_SCREEN = 3
PEN_WIDTH_PDF = 1.5  # points
ERASER_RADIUS = 10   # screen pixels: strokes closer than this to the eraser are removed
# Questions that accept negative scores (down to -maximum), unless the Excel
# file has a "minimo" row giving the minimum of each question
NEGATIVE_QUESTIONS = {"G1", "G2"}
RUBRIC_NAME = "griglia.txt"             # grading rubric, in the PDF folder
PROPOSALS_NAME = "proposte.json"        # Claude's proposals, in the PDF folder
SOLUTIONS_NAME = "soluzioni.zip"        # LaTeX solutions, copied into the PDF folder
CSV_TOTAL_COLUMN = "Totale domande aperte"
PROPOSAL_BG = "#fff3b0"                 # background of proposed (unchecked) scores
SOLUTIONS_PDF_DIR = "soluzioni_pdf"     # compiled solutions, in the PDF folder
SOLUTIONS_DPI = 150

# Texts of the evaluation page and of the quick comment buttons
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
    "Una sezione per esercizio: [nome] su una riga (il nome è quello della colonna).\n"
    "Sotto, una riga per criterio: punti e descrizione, es.  1  equazione caratteristica corretta\n"
    "Punti negativi = penalità (es.  -0,25  errore di calcolo). Il massimo dell'esercizio\n"
    "è la somma dei punti positivi. Le altre righe sono note per chi corregge; # = commento.")


# ==================================================
# Utilities
# ==================================================

def compito_number(filename):
    """Exam number: last group of digits in the file name."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    nums = re.findall(r"\d+", stem)
    return int(nums[-1]) if nums else None


def fmt_num(x):
    """Number in Italian format (decimal comma), without needless decimals."""
    if x is None:
        return ""
    x = float(x)
    s = f"{x:.0f}" if x == int(x) else f"{x:.2f}".rstrip("0")
    return s.replace(".", ",")


def parse_num(text):
    return float(text.strip().replace(",", "."))


def _register_font():
    """Unicode-capable font for the evaluation page (falls back to Helvetica)."""
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
                raise ValueError(f"Riga {lineno}: nome dell'esercizio mancante: {raw!r}")
            if any(e["name"] == name for e in exercises):
                raise ValueError(f"Riga {lineno}: esercizio {name} ripetuto")
            exercises.append({"name": name, "criteria": [], "notes": []})
            continue
        if not exercises:
            raise ValueError(f"Riga {lineno}: manca prima [nome dell'esercizio]")
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
    lines = ["# Griglia di valutazione", ""]
    for name, qmax in questions:
        lines.append(f"[{name}]" + (f"   # massimo {fmt_num(qmax)}" if qmax else ""))
        lines.append("")
    return "\n".join(lines)


# ==================================================
# Excel
# ==================================================

class ExcelGrades:
    """Reads/writes scores in the Excel file, by exam number."""
    kind = "Excel"
    questions_from_rubric = False

    def __init__(self, path):
        self.path = path
        self.wb = load_workbook(path)
        self.ws = self.wb["Punteggi"] if "Punteggi" in self.wb.sheetnames else self.wb.active
        ws = self.ws

        headers = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if v is not None and str(v).strip():
                headers[c] = str(v).strip()
        lower = {h.lower(): c for c, h in headers.items()}

        if "compito" not in lower:
            raise ValueError("Colonna 'Compito' non trovata nella riga 1 del foglio.")
        self.col_compito = lower["compito"]
        self.col_matricola = lower.get("matricola")
        self.col_cognome = lower.get("cognome")
        self.col_nome = lower.get("nome")
        self.col_note = lower.get("note")

        # Row of maximums and student rows
        self.max_row = None
        self.min_row = None
        self.rows = {}
        self.duplicates = []
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=self.col_compito).value
            if isinstance(v, str) and v.strip().lower() == "massimo":
                self.max_row = r
                continue
            if isinstance(v, str) and v.strip().lower() == "minimo":
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
            raise ValueError("Riga 'massimo' non trovata nella colonna 'Compito'.")

        # Questions = columns with a numeric maximum
        self.questions = []  # list of (name, column, maximum)
        for c, h in sorted(headers.items()):
            if c == self.col_compito:
                continue
            m = ws.cell(row=self.max_row, column=c).value
            if isinstance(m, (int, float)) and not isinstance(m, bool):
                self.questions.append((h, c, float(m)))
        if not self.questions:
            raise ValueError("Nessuna domanda trovata (nessun valore numerico nella riga 'massimo').")

        # Minimum score: from the "minimo" row if present, otherwise -maximum for
        # questions in NEGATIVE_QUESTIONS and 0 for the others
        self.minimums = {}
        for qname, c, qmax in self.questions:
            m = ws.cell(row=self.min_row, column=c).value if self.min_row else None
            if isinstance(m, (int, float)) and not isinstance(m, bool):
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

        return {"compito": n, "matricola": val(self.col_matricola),
                "cognome": val(self.col_cognome), "nome": val(self.col_nome),
                "note": val(self.col_note)}

    def get_scores(self, n):
        r = self.rows[n]
        out = []
        for _, c, _ in self.questions:
            v = self.ws.cell(row=r, column=c).value
            out.append(v if isinstance(v, (int, float)) and not isinstance(v, bool) else None)
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
    """Reads/writes scores in the CSV of "Correggi VF" (or any CSV with a
    "Numero compito"/"Compito" column). The questions are not in the file: they
    come from the rubric (set_questions), and their columns are added if missing."""
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
            raise ValueError("Il file CSV è vuoto.")
        self.header = [h.strip() for h in self.data[0]]
        lower = {h.lower(): i for i, h in enumerate(self.header) if h}

        col = lower.get("numero compito", lower.get("compito"))
        if col is None:
            raise ValueError("Colonna 'Numero compito' (o 'Compito') non trovata nella prima riga.")
        self.col_compito = col
        self.col_matricola = lower.get("matricola")
        self.col_cognome = lower.get("cognome")
        self.col_nome = lower.get("nome")
        self.col_note = lower.get("note")

        self.rows = {}
        self.duplicates = []
        for i, row in enumerate(self.data[1:], start=1):
            try:
                n = int(str(row[col]).strip())
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
        return {"compito": n, "matricola": self._cell(r, self.col_matricola),
                "cognome": self._cell(r, self.col_cognome), "nome": self._cell(r, self.col_nome),
                "note": self._cell(r, self.col_note)}

    def get_scores(self, n):
        r = self.rows[n]
        out = []
        for _, c, _ in self.questions:
            try:
                out.append(parse_num(self._cell(r, c)))
            except ValueError:
                out.append(None)
        return out

    def is_graded(self, n):
        return bool(self.questions) and all(v is not None for v in self.get_scores(n))

    def set_scores(self, n, scores):
        row = self.data[self.rows[n]]
        values = [(c, f"{s:g}") for (_, c, _), s in zip(self.questions, scores)]
        values.append((self._column(CSV_TOTAL_COLUMN), f"{sum(scores):g}"))
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
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    left, right = 50, W - 50

    c.setFont(bold, 18)
    c.drawCentredString(W / 2, H - 50, T["title"])
    c.line(left, H - 62, right, H - 62)

    y = H - 85
    c.setFont(regular, 11)
    who = f"{T['exam']} {student['compito']}"
    name = f"{student.get('cognome', '')} {student.get('nome', '')}".strip()
    if name:
        who += f" - {name}"
    if student.get("matricola"):
        who += f" - {T['id']} {student['matricola']}"
    c.drawString(left, y, who)
    y -= 35

    c.setFont(bold, 14)
    c.drawString(left, y, T["scores"])
    y -= 25
    c.setFont(regular, 12)
    total = 0.0
    for (qname, _, qmax), s in zip(questions, scores):
        c.drawString(70, y, f"{qname}: {fmt_num(s)} / {fmt_num(qmax)}")
        total += s
        y -= 22
    y -= 8
    c.setFont(bold, 13)
    c.drawString(70, y, f"{T['total']}: {fmt_num(total)} / "
                        f"{fmt_num(sum(q[2] for q in questions))}")
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
# True/false page (from the CSV of "Correggi VF")
# ==================================================

VF_TEXT = {
    False: {"title": "Domande vero/falso", "exam": "Compito n.", "student": "Studente",
            "id": "Matricola", "question": "Domanda", "given": "Risposta data",
            "correct": "Risposta corretta", "outcome": "Esito", "points": "Punti",
            "total": "Punteggio vero/falso", "right": "giusta", "wrong": "sbagliata",
            "blank": "non data", "answer": {"V": "Vero", "F": "Falso"}},
    True: {"title": "True/false questions", "exam": "Exam no.", "student": "Student",
           "id": "Student ID", "question": "Question", "given": "Your answer",
           "correct": "Correct answer", "outcome": "Outcome", "points": "Points",
           "total": "True/false score", "right": "correct", "wrong": "wrong",
           "blank": "not given", "answer": {"V": "True", "F": "False"}},
}
GREEN, RED, GREY = (0.0, 0.5, 0.15), (0.8, 0.1, 0.1), (0.45, 0.45, 0.45)



def to_float(x):
    try:
        return float(str(x).strip().replace(",", "."))
    except ValueError:
        return None


def _fmt_text(x):
    """A number read from the CSV, in Italian format (unchanged if not a number)."""
    v = to_float(x)
    return fmt_num(v) if v is not None else str(x)


# Fields of the mapping: (key, label shown in the window, required)
VF_FIELDS = [("num", "Numero del compito", True), ("cognome", "Cognome", True),
          ("nome", "Nome", True), ("matricola", "Matricola", False), ("dsa", "DSA", False),
          ("total", "Punteggio vero/falso (totale)", True),
          ("key", "Risposte giuste (tutte in una colonna)", True)]


def guess_vf_columns(header):
    """Guess of the columns, by name: {field: index or None, "answers": [...], "points": [...]}"""
    low = [h.lower() for h in header]

    def col(*names):
        return next((low.index(n.lower()) for n in names if n.lower() in low), None)

    m = {"num": col("Numero compito", "Compito"), "cognome": col("Cognome"),
         "nome": col("Nome"), "matricola": col("Matricola"), "dsa": col("DSA"),
         "total": col("Totale"), "key": col("Risposte giuste")}
    m["answers"] = [i for i, h in enumerate(low)
                    if h.startswith("risposta ") and h != "risposte giuste"]
    points = {h[len("punti "):].strip(): i for i, h in enumerate(low) if h.startswith("punti ")}
    m["points"] = [points[low[i][len("risposta "):].strip()] for i in m["answers"]
                   if low[i][len("risposta "):].strip() in points]
    if len(m["points"]) != len(m["answers"]):
        m["points"] = []
    return m


def check_vf_columns(m):
    """Error message for an incomplete mapping, or None."""
    for key, label, required in VF_FIELDS:
        if required and m.get(key) is None:
            return f"Scegli la colonna «{label}»."
    if not m.get("answers"):
        return "Scegli le colonne con le risposte date dallo studente."
    if m.get("points") and len(m["points"]) != len(m["answers"]):
        return (f"Hai scelto {len(m['answers'])} colonne di risposte e {len(m['points'])} "
                "di punti: devono essere tante quante (o nessuna colonna di punti).")
    return None


def vf_question_label(name):
    return name[len("Risposta "):].strip() if name.lower().startswith("risposta ") else name


def vf_students(header, rows, m):
    """Returns (students {n: dict}, questions [label], max_points [float|None])."""
    err = check_vf_columns(m)
    if err:
        raise ValueError(err)
    questions = [vf_question_label(header[i]) for i in m["answers"]]
    students = {}
    for r in rows:
        r = r + [""] * (len(header) - len(r))
        try:
            n = int(float(r[m["num"]]))
        except ValueError:
            continue                      # e.g. a "massimo" row
        get = lambda k: r[m[k]].strip() if m.get(k) is not None else ""
        students[n] = {
            "cognome": get("cognome"), "nome": get("nome"), "matricola": get("matricola"),
            "dsa": bool(get("dsa")) and get("dsa").lower() not in ("no", "0", "false"),
            "total": get("total"), "key": get("key").upper().replace(" ", ""),
            "given": [r[i].strip().upper() for i in m["answers"]],
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
    T = VF_TEXT[bool(english)]
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
    c.drawString(left, y, f"{T['student']}: {st['cognome']} {st['nome']}".rstrip())
    if st["matricola"]:
        c.drawRightString(right, y, f"{T['id']}: {st['matricola']}")
    if st["dsa"]:
        y -= 18
        c.setFont(regular, 11)
        c.drawString(left, y, "DSA")
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
        if given not in ("V", "F"):
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
        c.drawRightString(cols[4], y, _fmt_text(pts) + (f" / {fmt_num(mx)}" if mx and pts != "" else ""))
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
    # DSA students answer fewer questions: no overall maximum for them
    total_max = sum(max_points) if all(max_points) and not st["dsa"] else None
    c.drawString(left, y, T["total"])
    c.drawRightString(right, y, _fmt_text(st["total"]) + (f" / {fmt_num(total_max)}" if total_max else ""))
    c.showPage()
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]



def choose_vf_columns(root, header, rows, guess):
    """Window to choose the columns of the CSV. Returns the mapping, or None if cancelled."""
    import tkinter as tk
    from tkinter import ttk, messagebox

    NONE = "(nessuna)"
    names = [f"{i}: {h}" for i, h in enumerate(header)]
    example = next((r for r in rows if r and any(c.strip() for c in r)), [])

    def sample(i):
        v = example[i].strip() if i is not None and i < len(example) else ""
        return f"es.: {v[:40]}" if v else ""

    win = tk.Toplevel(root)
    win.title("Colonne del vero/falso")
    win.grab_set()
    tk.Label(win, text="Controlla quali colonne del CSV usare (la prima riga di dati è "
                       "mostrata come esempio).", font=("Arial", 11)).grid(
        row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 6))
    combos = {}
    for r, (key, label, required) in enumerate(VF_FIELDS, start=1):
        tk.Label(win, text=label + ("" if required else " (facoltativa)")).grid(
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

    r = len(VF_FIELDS) + 1
    lists = {}
    for c, (key, label) in enumerate((("answers", "Risposte dello studente\n(una colonna per domanda)"),
                                      ("points", "Punti di ogni domanda (facoltative,\n"
                                                 "nello stesso ordine delle risposte)"))):
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
        count.config(text=f"Domande scelte: {a}   —   colonne di punti: {p}",
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
        err = check_vf_columns(m)
        if err:
            messagebox.showwarning("Colonne", err, parent=win)
            return
        result["m"] = m
        win.destroy()

    bf = tk.Frame(win)
    bf.grid(row=r + 2, column=0, columnspan=3, pady=10)
    tk.Button(bf, text="Annulla", width=12, command=win.destroy).pack(side=tk.LEFT, padx=5)
    tk.Button(bf, text="Continua", width=12, command=ok).pack(side=tk.LEFT, padx=5)
    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.wait_window()
    return result.get("m")


# ==================================================
# Solutions (zip of LaTeX files, one per exam, compiled to PDF)
# ==================================================

def _find_pdflatex():
    for c in (shutil.which("pdflatex"), "/Library/TeX/texbin/pdflatex",
              "/usr/local/bin/pdflatex", "/opt/homebrew/bin/pdflatex"):
        if c and os.path.exists(c):
            return c
    return None


class SolutionsCompiler:
    """Compiles the solutions of each exam in a background thread.

    The zip holds one file per exam (.tex or .pdf), whose name ends with the exam number
    (00012.tex -> 12). PDFs go to <folder>/soluzioni_pdf/ and are reused until the zip changes.
    """

    def __init__(self, zip_path, out_dir):
        self.zip_path = zip_path
        self.out_dir = out_dir
        self.pdflatex = _find_pdflatex()
        self.cond = threading.Condition()
        self.queue = []
        self.busy = None
        self.errors = {}         # {n: message}
        self.entries = {}        # {n: name in the zip}
        self.zip_mtime = None
        self.work_dir = None
        self.thread = None

    def _zip_mtime(self):
        try:
            return os.path.getmtime(self.zip_path)
        except OSError:
            return None

    def refresh(self):
        """Re-reads the zip if it changed. Returns False if there is no zip."""
        mtime = self._zip_mtime()
        if mtime is None:
            self.entries = {}
            return False
        if mtime != self.zip_mtime:
            entries = {}
            with zipfile.ZipFile(self.zip_path) as z:
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
                self.entries = entries
                self.zip_mtime = mtime
                self.errors = {}
        return True

    def has(self, n):
        return n in self.entries

    def cleanup(self):
        """Removes the extracted zip (it is extracted again when needed)."""
        if self.work_dir and self.busy is None:
            shutil.rmtree(self.work_dir, ignore_errors=True)
        self.work_dir = None

    def pdf_path(self, n):
        return os.path.join(self.out_dir, f"{n:05d}.pdf")

    def ready(self, n):
        p = self.pdf_path(n)
        return os.path.exists(p) and self.zip_mtime is not None \
            and os.path.getmtime(p) >= self.zip_mtime

    def request(self, n, first=False):
        if n not in self.entries or self.ready(n) or n in self.errors:
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
        os.makedirs(self.out_dir, exist_ok=True)
        work = self.work_dir
        if work is None or not os.path.isdir(work):
            work = tempfile.mkdtemp(prefix="soluzioni_")
            with zipfile.ZipFile(self.zip_path) as z:
                z.extractall(work)     # the other files of the zip may be needed (images...)
            self.work_dir = work
        src = os.path.join(work, name)
        out = self.pdf_path(n)
        if name.lower().endswith(".pdf"):
            self._publish(src, out)
            return
        if not self.pdflatex:
            raise RuntimeError("pdflatex non trovato: installa una distribuzione LaTeX "
                               "(es. MacTeX) per vedere le soluzioni.")
        d, base = os.path.split(src)
        job = os.path.splitext(base)[0]
        r = subprocess.run([self.pdflatex, "-interaction=nonstopmode", "-halt-on-error",
                            "-no-shell-escape", base],
                           cwd=d, capture_output=True, text=True, errors="replace", timeout=180)
        pdf = os.path.join(d, job + ".pdf")
        if r.returncode != 0 or not os.path.exists(pdf):
            err = [l for l in r.stdout.splitlines() if l.startswith("!")]
            raise RuntimeError(f"Errore LaTeX in {base}:\n" + ("\n".join(err[:3]) or r.stdout[-500:]))
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
        self.images = []          # pages of the exam shown (PIL)
        self.images_for = None
        self.tk_imgs = []
        self.zoom = 1.0           # relative to the window width
        self.prefetched = False

        self.win = tk.Toplevel(app.root)
        self.win.title("Soluzioni")
        self.win.geometry(app.config.get("solutions_geometry", "800x900+60+60"))
        self.win.protocol("WM_DELETE_WINDOW", self.hide)
        tb = tk.Frame(self.win)
        tb.pack(fill="x", padx=5, pady=5)
        self.label = tk.Label(tb, font=("Arial", 13, "bold"), anchor="w")
        self.label.pack(side=tk.LEFT)
        tk.Button(tb, text="Adatta", command=self.zoom_reset).pack(side=tk.RIGHT, padx=2)
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
        self.label.config(text=f"Soluzioni — compito {n}")
        c = self.compiler
        before = c.zip_mtime
        if not c.refresh():
            self.message("Nessuno zip di soluzioni in questa cartella.\n"
                         "Usa il pulsante «Soluzioni…» per sceglierlo.")
            return
        if c.zip_mtime != before:
            self.prefetched = False      # new zip: compile everything again
        if not c.has(n):
            self.message(f"Nello zip non c'è un file per il compito {n}.")
            return
        if c.ready(n):
            if self.images_for != (n, c.zip_mtime):
                try:
                    self.images = convert_from_path(c.pdf_path(n), dpi=SOLUTIONS_DPI)
                    self.images_for = (n, c.zip_mtime)
                except Exception as e:
                    self.message(f"Impossibile visualizzare le soluzioni:\n{e}")
                    return
                self.canvas.yview_moveto(0)
            self.draw()
            self.prefetch()
            return
        if n in c.errors:
            self.message(c.errors[n])
            return
        self.message("Compilazione delle soluzioni in corso…")
        c.request(n, first=True)
        self.poll_job = self.win.after(300, lambda: self.show(n) if self.n == n else None)

    def prefetch(self):
        """Compiles the other exams in the background, in grading order."""
        if self.prefetched:
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
        self.root.title("Correzione compiti PDF")
        self.root.geometry("1400x900")
        self.ok = False

        self.config = load_config()
        self.fonts = _register_font()

        self.grades = None
        self.folder = ""
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
        self.vf = None           # true/false data from the CSV (see setup_vf)

        self.rubric = []         # parsed rubric (see parse_rubric)
        self.proposals = {}      # {"12": {...}} from proposte.json
        self.proposals_mtime = None
        self.proposal_shown = False   # the comment box holds a proposal

        if not self.open_grades() or not self.open_folder() or not self.setup_rubric():
            return
        if not os.path.exists(self.solutions_path()):
            self.choose_solutions(startup=True)
        self.sol_compiler = SolutionsCompiler(self.solutions_path(),
                                              os.path.join(self.folder, SOLUTIONS_PDF_DIR))
        self.setup_vf()
        self.build_ui()
        self.current_index = self.first_ungraded_index()
        if os.path.exists(self.solutions_path()):
            self.solutions = SolutionsWindow(self)
        self.load_pdf()
        self.ok = True

    # --------------------------------------------------
    # Setup
    # --------------------------------------------------

    def open_grades(self):
        last = self.config.get("excel")
        path = filedialog.askopenfilename(
            title="Seleziona il file con l'elenco degli studenti (Excel o CSV)",
            initialdir=os.path.dirname(last) if last else None,
            filetypes=[("Excel o CSV", "*.xlsx *.xlsm *.csv"), ("Tutti i file", "*.*")])
        if not path:
            return False
        try:
            if path.lower().endswith(".csv"):
                self.grades = CsvGrades(path)
            else:
                self.grades = ExcelGrades(path)
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile leggere il file:\n{e}")
            return False

        if self.grades.questions_from_rubric:
            msg = ("Gli esercizi e i loro punteggi massimi si definiscono nella griglia "
                   "di valutazione.")
        else:
            qs = ", ".join(
                f"{q[0]} (max {fmt_num(q[2])}"
                + (f", min {fmt_num(self.grades.minimums[q[0]])})"
                   if self.grades.minimums[q[0]] else ")")
                for q in self.grades.questions)
            msg = f"Domande da correggere:\n{qs}"
        msg += f"\n\nStudenti trovati: {len(self.grades.rows)}"
        if self.grades.duplicates:
            msg += f"\n\nATTENZIONE: numeri di compito duplicati: {self.grades.duplicates}"
        if self.grades.lock_file_exists():
            msg += ("\n\nATTENZIONE: il file sembra aperto in Excel. Chiudilo prima di "
                    "salvare, altrimenti Excel potrebbe sovrascrivere i punteggi.")
        if not messagebox.askokcancel(f"File {self.grades.kind}", msg):
            return False
        self.config["excel"] = path
        save_config(self.config)
        return True

    # --------------------------------------------------
    # Rubric
    # --------------------------------------------------

    def rubric_path(self):
        return os.path.join(self.folder, RUBRIC_NAME)

    def read_rubric_text(self):
        try:
            with open(self.rubric_path(), "r", encoding="utf-8") as f:
                return f.read()
        except FileNotFoundError:
            return None

    def setup_rubric(self):
        """Loads griglia.txt; if missing, asks for it (required for a CSV file)."""
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
            messagebox.showwarning("Griglia di valutazione",
                                   f"{RUBRIC_NAME} va corretta:\n{problem}")
        elif not self.grades.questions_from_rubric:
            if not messagebox.askyesno(
                    "Griglia di valutazione",
                    "Nella cartella non c'è ancora la griglia di valutazione.\n\n"
                    "Vuoi scriverla ora? (serve per le proposte di valutazione; "
                    "puoi farlo anche dopo col pulsante «Griglia…»)"):
                return True
        return self.edit_rubric(startup=True)

    def check_rubric(self, rubric):
        """Error message, or "" if the rubric can be used with the grades file."""
        if self.grades.questions_from_rubric:
            if not rubric:
                return "Scrivi almeno un esercizio, es. [Esercizio 1]."
            empty = [e["name"] for e in rubric if e["max"] <= 0]
            if empty:
                return f"Esercizi senza punti positivi: {', '.join(empty)}"
            return ""
        known = {q[0] for q in self.grades.questions}
        unknown = [e["name"] for e in rubric if e["name"] not in known]
        if unknown:
            return (f"Esercizi che non sono colonne dell'Excel: {', '.join(unknown)}\n"
                    f"Colonne: {', '.join(q[0] for q in self.grades.questions)}")
        return ""

    def rubric_mismatches(self, rubric):
        """Exercises whose rubric total differs from the Excel maximum."""
        if self.grades.questions_from_rubric:
            return []
        qmax = {q[0]: q[2] for q in self.grades.questions}
        return [f"{e['name']}: griglia {fmt_num(e['max'])}, Excel {fmt_num(qmax[e['name']])}"
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
                text = rubric_template([("Esercizio 1", 0), ("Esercizio 2", 0)])
            else:
                text = rubric_template([(q[0], q[2]) for q in self.grades.questions])

        win = tk.Toplevel(self.root)
        win.title("Griglia di valutazione")
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
                messagebox.showerror("Griglia", str(e), parent=win)
                return
            problem = self.check_rubric(rubric)
            if problem:
                messagebox.showerror("Griglia", problem, parent=win)
                return
            mism = self.rubric_mismatches(rubric)
            if mism and not messagebox.askyesno(
                    "Griglia", "La somma dei punti non coincide con il massimo:\n"
                    + "\n".join(mism) + "\n\nSalvare comunque?", parent=win):
                return
            with open(self.rubric_path(), "w", encoding="utf-8") as f:
                f.write(box.get("1.0", "end-1c").rstrip() + "\n")
            values = self.current_entry_values() if hasattr(self, "score_frame") else None
            self.apply_rubric(rubric)
            if values is not None:
                self.build_score_rows(values)
            result["ok"] = True
            win.destroy()

        bf = tk.Frame(win)
        bf.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(bf, text="Salva", width=12, bg="lightgreen", command=save).pack(side=tk.RIGHT)
        tk.Button(bf, text="Annulla", width=12, command=win.destroy).pack(side=tk.RIGHT, padx=5)
        win.grab_set()
        box.focus_set()
        self.root.wait_window(win)
        if startup and not result["ok"] and self.grades.questions_from_rubric:
            return False
        return True

    # --------------------------------------------------
    # Solutions (zip with the LaTeX files, read by Claude for the proposals)
    # --------------------------------------------------

    def solutions_path(self):
        return os.path.join(self.folder, SOLUTIONS_NAME)

    def choose_solutions(self, startup=False):
        """Copies the solutions zip into the PDF folder as soluzioni.zip."""
        if startup and not messagebox.askyesno(
                "Soluzioni",
                "Per le proposte di valutazione servono le soluzioni (zip con i file LaTeX).\n\n"
                "Vuoi sceglierle ora? (puoi farlo anche dopo col pulsante «Soluzioni…»)"):
            return
        path = filedialog.askopenfilename(
            title="Seleziona lo zip con le soluzioni LaTeX",
            initialdir=self.folder, filetypes=[("Zip", "*.zip"), ("Tutti i file", "*.*")])
        if not path:
            return
        if os.path.abspath(path) != os.path.abspath(self.solutions_path()):
            shutil.copy2(path, self.solutions_path())
            os.utime(self.solutions_path())   # newer than the PDFs compiled from the old zip
        messagebox.showinfo(
            "Soluzioni",
            f"Soluzioni copiate in:\n{self.solutions_path()}\n\n"
            "Ora chiedi a Claude (in Claude Code) di preparare le proposte per questa "
            f"cartella:\n{self.folder}\n\nCompariranno qui man mano che sono pronte.")
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

    def setup_vf(self):
        """True/false data from the CSV of Correggi VF: columns saved for this folder,
        otherwise guessed from the names. Nothing for an Excel file."""
        self.vf = None
        if not isinstance(self.grades, CsvGrades):
            return
        header = self.grades.header
        m = self._vf_columns_from_state(header) or guess_vf_columns(header)
        if check_vf_columns(m) is None:
            self._set_vf(m)

    def _set_vf(self, m):
        students, questions, max_points = vf_students(self.grades.header, self.grades.data[1:], m)
        self.vf = {"columns": m, "students": students, "questions": questions, "max": max_points}

    def _vf_columns_from_state(self, header):
        saved = self.state.get("vf_columns")
        if not saved:
            return None
        try:
            idx = lambda name: None if name is None else header.index(name)
            m = {k: idx(saved.get(k)) for k, _, _ in VF_FIELDS}
            m["answers"] = [header.index(x) for x in saved.get("answers", [])]
            m["points"] = [header.index(x) for x in saved.get("points", [])]
        except ValueError:            # a saved column is no longer in the file
            return None
        return m

    def choose_vf(self):
        header = self.grades.header
        guess = self.vf["columns"] if self.vf else guess_vf_columns(header)
        m = choose_vf_columns(self.root, header, self.grades.data[1:], guess)
        if m is None:
            return
        try:
            self._set_vf(m)
        except Exception as e:
            messagebox.showerror("Colonne V/F", str(e))
            return
        name = lambda i: None if i is None else header[i]
        saved = {k: name(m.get(k)) for k, _, _ in VF_FIELDS}
        saved["answers"] = [header[i] for i in m["answers"]]
        saved["points"] = [header[i] for i in m["points"]]
        self.state["vf_columns"] = saved
        self.opt_vf.set(True)
        self.save_options()

    def save_options(self):
        self.state["options"] = {"vf": self.opt_vf.get(), "solutions": self.opt_solutions.get()}
        self.save_state()

    def solution_pdf(self, n):
        """(path of the solutions PDF of exam n, None) or (None, reason). Compiles it now if needed."""
        c = self.sol_compiler
        if not c.refresh():
            return None, "manca lo zip delle soluzioni («Soluzioni…»)"
        if not c.has(n):
            return None, f"nello zip non c'è la soluzione del compito {n}"
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
        return None, c.errors.get(n, "compilazione troppo lunga")

    def extra_pages(self, n, english):
        """(true/false page or None, solutions PDF or None, notes on what is missing)."""
        first = solutions = None
        notes = []
        if self.opt_vf.get():
            if self.vf is None:
                notes.append("pagina V/F: scegli le colonne con «Colonne V/F…»")
            elif n not in self.vf["students"]:
                notes.append(f"pagina V/F: il compito {n} non è nel CSV")
            else:
                first = true_false_page(n, self.vf["students"][n], self.vf["questions"],
                                        self.vf["max"], self.fonts, english)
        if self.opt_solutions.get():
            solutions, why = self.solution_pdf(n)
            if why:
                notes.append("soluzioni: " + why)
        return first, solutions, notes

    def regenerate_all(self):
        """Rebuilds the PDF of every graded exam with the current options."""
        self.remember_current()
        todo = [(k, p) for k, p in self.pdf_files
                if self.grades.is_graded(k) and os.path.exists(self.original_path(p))]
        if not todo:
            messagebox.showinfo("Rigenera PDF", "Nessun compito già corretto.")
            return
        yes_no = lambda v: "sì" if v else "no"
        if not messagebox.askyesno(
                "Rigenera PDF",
                f"Rifaccio il PDF dei {len(todo)} compiti già corretti, con le opzioni attuali:\n"
                f"  pagina V/F: {yes_no(self.opt_vf.get())}\n"
                f"  soluzioni in coda: {yes_no(self.opt_solutions.get())}\n\n"
                "Punteggi, commenti, disegni e lingua restano quelli salvati."):
            return
        problems = []
        for i, (k, p) in enumerate(todo, start=1):
            self.label_status.config(text=f"Rigenero i PDF: {i}/{len(todo)}…")
            self.root.update_idletasks()
            saved = self.state["files"].get(os.path.basename(p), {})
            english = bool(saved.get("english"))
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
        report = f"PDF rifatti: {len(todo)}."
        if problems:
            report += "\n\nProblemi:\n" + "\n".join(problems[:30])
            if len(problems) > 30:
                report += f"\n… e altri {len(problems) - 30}"
        (messagebox.showwarning if problems else messagebox.showinfo)("Rigenera PDF", report)

    # --------------------------------------------------
    # Proposals (proposte.json, written by Claude)
    # --------------------------------------------------

    def proposals_path(self):
        return os.path.join(self.folder, PROPOSALS_NAME)

    def reload_proposals(self):
        """Re-reads proposte.json if it changed (it may be written while grading)."""
        try:
            mtime = os.path.getmtime(self.proposals_path())
        except OSError:
            self.proposals, self.proposals_mtime = {}, None
            return
        if mtime == self.proposals_mtime:
            return
        try:
            with open(self.proposals_path(), "r", encoding="utf-8") as f:
                data = json.load(f)
            self.proposals = {str(k): v for k, v in data.items() if isinstance(v, dict)}
            self.proposals_mtime = mtime
        except Exception as e:
            messagebox.showwarning("Proposte", f"Impossibile leggere {PROPOSALS_NAME}:\n{e}")
            self.proposals, self.proposals_mtime = {}, mtime

    def proposal(self, n):
        return self.proposals.get(str(n))

    def proposal_comment(self, prop, english):
        if not prop:
            return ""
        return str(prop.get("comment_en" if english else "commento") or "").strip()

    def proposal_score(self, prop, qname):
        """Proposed score for a question, if valid."""
        v = (prop or {}).get("punteggi", {}).get(qname)
        if isinstance(v, str):
            try:
                v = parse_num(v)
            except ValueError:
                return None
        if not _is_number(v):
            return None
        q = next(q for q in self.grades.questions if q[0] == qname)
        return float(v) if self.grades.minimums[qname] <= v <= q[2] else None

    def open_folder(self):
        folder = filedialog.askdirectory(
            title="Seleziona la cartella con i PDF",
            initialdir=self.config.get("folder"))
        if not folder:
            return False
        files, no_number, not_in_excel, seen = [], [], [], {}
        for f in os.listdir(folder):
            if not f.lower().endswith(".pdf") or f.startswith("."):
                continue
            n = compito_number(f)
            if n is None:
                no_number.append(f)
            elif n not in self.grades.rows:
                not_in_excel.append(f)
            elif n in seen:
                messagebox.showerror("Errore", f"Due PDF per il compito {n}:\n{seen[n]}\n{f}")
                return False
            else:
                seen[n] = f
                files.append((n, os.path.join(folder, f)))
        if not files:
            messagebox.showerror("Errore", "Nessun PDF associabile a un compito dell'elenco.")
            return False
        files.sort()

        warn = []
        if no_number:
            warn.append("PDF senza numero nel nome (ignorati):\n  " + "\n  ".join(no_number))
        if not_in_excel:
            warn.append("PDF con numero non presente nell'elenco (ignorati):\n  "
                        + "\n  ".join(not_in_excel))
        missing = sorted(set(self.grades.rows) - set(seen))
        if missing:
            warn.append(f"Compiti nell'elenco senza PDF: {missing}")
        if warn:
            messagebox.showwarning("Attenzione", "\n\n".join(warn))

        self.folder = folder
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
        tk.Button(tb, text="◀ Pagina", command=self.prev_page).pack(side=tk.LEFT, padx=2)
        self.page_label = tk.Label(tb, text="1/1", width=8)
        self.page_label.pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Pagina ▶", command=self.next_page).pack(side=tk.LEFT, padx=2)
        tk.Label(tb, text="|").pack(side=tk.LEFT, padx=5)
        tk.Button(tb, text="Zoom +", command=self.zoom_in).pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Zoom −", command=self.zoom_out).pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Adatta", command=self.zoom_reset).pack(side=tk.LEFT, padx=2)
        tk.Label(tb, text="|").pack(side=tk.LEFT, padx=5)
        self.tool_var = tk.StringVar(value="pen")
        tk.Radiobutton(tb, text="✎ Penna", variable=self.tool_var, value="pen",
                       indicatoron=False, padx=6).pack(side=tk.LEFT, padx=2)
        tk.Radiobutton(tb, text="⌫ Gomma", variable=self.tool_var, value="eraser",
                       indicatoron=False, padx=6).pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Annulla", command=self.undo).pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Cancella pagina", command=self.clear_annotations).pack(side=tk.LEFT, padx=2)
        tk.Label(tb, text="|").pack(side=tk.LEFT, padx=5)
        tk.Button(tb, text="Soluzioni ⧉", command=self.show_solutions).pack(side=tk.LEFT, padx=2)

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

        opts = tk.Frame(right)
        opts.pack(fill="x", padx=5)
        tk.Button(opts, text="Griglia…", command=self.edit_rubric).pack(side=tk.LEFT, padx=2)
        tk.Button(opts, text="Soluzioni…", command=self.choose_solutions).pack(side=tk.LEFT, padx=2)
        self.english_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opts, text="Inglese", variable=self.english_var,
                       command=self.toggle_english,
                       font=("Arial", 11, "bold")).pack(side=tk.RIGHT, padx=2)

        # What goes into the PDF besides the exam and the evaluation page
        opts2 = tk.Frame(right)
        opts2.pack(fill="x", padx=5, pady=(3, 0))
        options = self.state.get("options", {})
        is_csv = isinstance(self.grades, CsvGrades)
        self.opt_vf = tk.BooleanVar(value=is_csv and options.get("vf", True))
        self.opt_solutions = tk.BooleanVar(value=options.get("solutions", True))
        tk.Checkbutton(opts2, text="Pagina V/F", variable=self.opt_vf, command=self.save_options,
                       state="normal" if is_csv else "disabled").pack(side=tk.LEFT)
        tk.Button(opts2, text="Colonne V/F…", command=self.choose_vf,
                  state="normal" if is_csv else "disabled").pack(side=tk.LEFT, padx=(0, 6))
        tk.Checkbutton(opts2, text="Soluzioni in coda", variable=self.opt_solutions,
                       command=self.save_options).pack(side=tk.LEFT)
        tk.Button(opts2, text="Rigenera PDF…", command=self.regenerate_all).pack(side=tk.RIGHT)

        self.score_frame = tk.LabelFrame(right, text="Punteggi", font=("Arial", 10, "bold"))
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
        tk.Button(nav, text="◀ Precedente", command=self.previous_student,
                  width=14).pack(side=tk.LEFT, padx=2)
        tk.Button(nav, text="Salva e continua ▶", command=self.save_and_next,
                  width=18, bg="lightblue").pack(side=tk.RIGHT, padx=2)
        nav2 = tk.Frame(right)
        nav2.pack(fill="x", padx=5)
        tk.Button(nav2, text="Vai al compito…", command=self.goto_compito,
                  width=14).pack(side=tk.LEFT, padx=2)
        tk.Button(nav2, text="Successivo (senza salvare) ▶", command=self.skip_student,
                  width=22).pack(side=tk.RIGHT, padx=2)
        tk.Button(right, text="Salva e esci", command=self.save_and_exit,
                  width=18, bg="lightgreen").pack(pady=5)

        tk.Label(right, text="Commenti:", font=("Arial", 10, "bold")).pack(anchor="w", padx=5, pady=(10, 0))
        self.text_box = scrolledtext.ScrolledText(right, width=45, height=9,
                                                  font=("Arial", 11), wrap="word", undo=True)
        self.text_box.pack(fill="both", expand=True, padx=5, pady=5)

        cbf = tk.Frame(right)
        cbf.pack(fill="x", padx=5, pady=(0, 5))
        for i in range(len(LABELS[False]["quick"])):
            tk.Button(cbf, text=LABELS[False]["quick"][i][:20],
                      command=lambda i=i: self.insert_comment(
                          LABELS[self.english_var.get()]["quick"][i])).pack(side=tk.LEFT, padx=2)
        tk.Button(cbf, text="Copia commento precedente",
                  command=self.copy_previous_comment).pack(side=tk.LEFT, padx=2)

        # Phrases already used in other exams (double-click to insert)
        tk.Label(right, text="Frasi già usate (doppio clic per inserire):",
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
        """(Re)builds the score boxes; values = {question: text} to keep."""
        for w in self.score_frame.winfo_children():
            w.destroy()
        self.score_entries = []
        for i, (qname, _, qmax) in enumerate(self.grades.questions):
            row = tk.Frame(self.score_frame)
            row.pack(fill="x", padx=10, pady=2)
            qmin = self.grades.minimums[qname]
            rng = f"max {fmt_num(qmax)}" if qmin == 0 else f"da {fmt_num(qmin)} a {fmt_num(qmax)}"
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
        tk.Label(trow, text="Totale:", width=18, anchor="w",
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.total_label = tk.Label(trow, text="0", font=("Arial", 11, "bold"), fg="green")
        self.total_label.pack(side=tk.LEFT)
        tk.Label(self.score_frame, text="In giallo: proposta da controllare",
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
        """Switches the language of the evaluation page and of the proposed comment."""
        english = self.english_var.get()
        prop = self.proposal(self.current_number())
        new = self.proposal_comment(prop, english)
        old = self.proposal_comment(prop, not english)
        current = self.text_box.get("1.0", tk.END).strip()
        if new and current != new:
            if (not current or current == old or messagebox.askyesno(
                    "Inglese" if english else "Italiano",
                    "Sostituire il commento con la proposta in "
                    + ("inglese" if english else "italiano")
                    + "?\n(le modifiche fatte al commento andranno perse)")):
                self.text_box.delete("1.0", tk.END)
                self.text_box.insert("1.0", new)
                self.proposal_shown = True
        elif english and not new and prop:
            messagebox.showinfo("Inglese", "Per questo compito non c'è una proposta di "
                                "commento in inglese: scrivilo tu.")
        name = os.path.basename(self.current_path())
        self.state["files"].setdefault(name, {})["english"] = english
        self.save_state()

    # --------------------------------------------------
    # State (drawings and comments, per file, in the PDF folder)
    # --------------------------------------------------

    def state_path(self):
        return os.path.join(self.folder, STATE_NAME)

    def load_state(self):
        try:
            with open(self.state_path(), "r", encoding="utf-8") as f:
                self.state = json.load(f)
            self.state.setdefault("files", {})
        except Exception:
            self.state = {"files": {}}

    def save_state(self):
        try:
            fd, tmp = tempfile.mkstemp(suffix=".json", dir=self.folder)
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(self.state, f)
            os.replace(tmp, self.state_path())
        except Exception as e:
            print(f"Impossibile salvare lo stato: {e}")

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
        return os.path.join(os.path.dirname(path), BACKUP_DIR, os.path.basename(path))

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

        english = bool(saved.get("english"))
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
            note = "Proposta di Claude: controlla punteggi e commento."
            bad = [q[0] for q in self.grades.questions
                   if q[0] in prop.get("punteggi", {}) and self.proposal_score(prop, q[0]) is None]
            if bad:
                note += f"\nPunteggi proposti fuori intervallo (ignorati): {', '.join(bad)}"
            if prop.get("dubbi"):
                note += f"\n⚠ {prop['dubbi']}"
        elif not graded:
            if not self.rubric:
                note = "Nessuna proposta: manca la griglia di valutazione (pulsante «Griglia…»)."
            elif not os.path.exists(self.solutions_path()):
                note = "Nessuna proposta: mancano le soluzioni (pulsante «Soluzioni…»)."
            elif not self.proposals:
                note = ("Nessuna proposta ancora: chiedi a Claude di prepararle per questa "
                        "cartella. Compaiono qui appena sono pronte.")
            else:
                note = "Proposta per questo compito non ancora pronta."
        self.label_proposal.config(text=note)

        st = self.grades.student(n)
        self.label_file.config(text=f"Compito {n}  —  {st['cognome']} {st['nome']}")
        info = f"Matricola: {st['matricola']}\nFile: {name}"
        if st["note"]:
            info += f"\nNote: {st['note']}"
        if graded:
            info += f"\n(già corretto: i punteggi vengono dal file {self.grades.kind})"
        self.label_student.config(text=info)
        graded = sum(1 for k, _ in self.pdf_files if self.grades.is_graded(k))
        self.label_status.config(
            text=f"{self.current_index + 1}/{len(self.pdf_files)}  —  corretti: {graded}")

        try:
            self.total_pages = len(PdfReader(self.source_path(path)).pages)
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire il PDF:\n{e}")
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
            messagebox.showerror("Errore", f"Impossibile visualizzare il PDF:\n{e}")
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
                messagebox.askyesno("Conferma", "Cancellare tutti i disegni di questa pagina?\n"
                                    "(si può annullare con «Annulla»)"):
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
            try:
                total += parse_num(e.get() or "0")
            except ValueError:
                pass
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
            messagebox.showinfo("Commento precedente", "Nessun commento nel compito precedente.")

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
                messagebox.showwarning("Errore", f"Inserisci il punteggio per {qname}")
                e.focus_set()
                return None
            try:
                s = parse_num(txt)
            except ValueError:
                messagebox.showwarning("Errore", f"Punteggio non valido per {qname}: '{txt}'")
                e.focus_set()
                return None
            qmin = self.grades.minimums[qname]
            if s < qmin or s > qmax:
                messagebox.showwarning("Errore",
                                       f"{qname}: il punteggio deve essere tra "
                                       f"{fmt_num(qmin)} e {fmt_num(qmax)}")
                e.focus_set()
                return None
            scores.append(s)
        return scores

    # --------------------------------------------------
    # Saving
    # --------------------------------------------------

    def save_current(self):
        """Writes the scores to the Excel/CSV file and regenerates the PDF. True on success."""
        scores = self.read_scores()
        if scores is None:
            return False
        n = self.current_number()
        path = self.current_path()
        kind = self.grades.kind

        while self.grades.lock_file_exists():
            if not messagebox.askretrycancel(
                    "File aperto",
                    f"Il file {kind} è aperto in Excel. Chiudilo (salvando eventuali tue "
                    "modifiche) e premi Riprova."):
                return False

        self.remember_current()
        try:
            self.grades.set_scores(n, scores)
            self.grades.save()
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile salvare il file {kind}:\n{e}")
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
            messagebox.showerror("Errore",
                                 f"Punteggi salvati nel file {kind}, ma impossibile salvare il PDF:\n{e}")
            return False
        if notes:
            messagebox.showwarning("PDF salvato", "PDF salvato senza:\n- " + "\n- ".join(notes))
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
                messagebox.showinfo("Fine elenco",
                                    f"Ultimo compito salvato.\nAncora da correggere: {missing}")
            else:
                messagebox.showinfo("Fine", "Tutti i compiti sono stati corretti!")
            self.load_pdf()
            return
        self.go_to(self.current_index + 1)

    def skip_student(self):
        if self.current_index + 1 < len(self.pdf_files):
            self.go_to(self.current_index + 1)

    def previous_student(self):
        if self.current_index > 0:
            self.go_to(self.current_index - 1)

    def goto_compito(self):
        n = simpledialog.askinteger("Vai al compito", "Numero del compito:", parent=self.root)
        if n is None:
            return
        for i, (k, _) in enumerate(self.pdf_files):
            if k == n:
                self.go_to(i)
                return
        messagebox.showwarning("Attenzione", f"Nessun PDF per il compito {n}")

    def save_and_exit(self):
        ans = messagebox.askyesnocancel("Esci", "Salvare il compito corrente prima di uscire?")
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
