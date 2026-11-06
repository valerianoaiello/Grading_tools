"""
Correzione delle domande aperte su compiti in PDF.

- Legge l'elenco studenti e le domande da un file Excel (foglio "Punteggi"):
  riga 1 = intestazioni (Compito, Matricola, Cognome, Nome, ..., G1, G2, 7A, ...),
  riga "massimo" = punteggio massimo di ogni domanda. Sono considerate domande
  da correggere tutte le colonne che hanno un valore numerico nella riga "massimo".
- Ogni PDF viene associato a uno studente tramite il numero del compito, preso
  dall'ultimo gruppo di cifre nel nome del file (es. "compito_12.pdf" -> 12).
- Permette di disegnare sulle pagine e di scrivere commenti; al salvataggio
  scrive i punteggi nell'Excel e rigenera il PDF (originale + disegni + pagina
  di valutazione). Gli originali vengono conservati nella sottocartella
  "originali", quindi un compito può essere ricorretto senza duplicare nulla.
"""
import io
import os
import re
import json
import shutil
import tempfile
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

STATE_NAME = ".correzione_stato.json"   # nella cartella dei PDF
BACKUP_DIR = "originali"                # sottocartella con i PDF originali
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".examinatrix_grader.json")
RENDER_DPI = 200
PEN_COLOR = "red"
PEN_WIDTH_SCREEN = 3
PEN_WIDTH_PDF = 1.5  # punti
# Domande che ammettono punteggi negativi (fino a -massimo), se l'Excel non ha
# una riga "minimo" che indica il minimo di ogni domanda
NEGATIVE_QUESTIONS = {"G1", "G2"}


# ==================================================
# Utilità
# ==================================================

def compito_number(filename):
    """Numero del compito: ultimo gruppo di cifre nel nome del file."""
    stem = os.path.splitext(os.path.basename(filename))[0]
    nums = re.findall(r"\d+", stem)
    return int(nums[-1]) if nums else None


def fmt_num(x):
    """Numero in formato italiano, senza decimali inutili."""
    if x is None:
        return ""
    x = float(x)
    s = f"{x:.0f}" if x == int(x) else f"{x:.2f}".rstrip("0")
    return s.replace(".", ",")


def parse_num(text):
    return float(text.strip().replace(",", "."))


def _register_font():
    """Font con supporto Unicode per la pagina di valutazione (fallback Helvetica)."""
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


# ==================================================
# Excel
# ==================================================

class ExcelGrades:
    """Lettura/scrittura dei punteggi nel file Excel, per numero di compito."""

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

        # Riga dei massimi e righe degli studenti
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

        # Domande = colonne con un massimo numerico
        self.questions = []  # lista di (nome, colonna, massimo)
        for c, h in sorted(headers.items()):
            if c == self.col_compito:
                continue
            m = ws.cell(row=self.max_row, column=c).value
            if isinstance(m, (int, float)) and not isinstance(m, bool):
                self.questions.append((h, c, float(m)))
        if not self.questions:
            raise ValueError("Nessuna domanda trovata (nessun valore numerico nella riga 'massimo').")

        # Punteggio minimo: dalla riga "minimo" se esiste, altrimenti -massimo per le
        # domande in NEGATIVE_QUESTIONS e 0 per le altre
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
        """Excel crea '~$nomefile.xlsx' quando il file è aperto."""
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


# ==================================================
# PDF
# ==================================================

def _display_to_pdf(u, v, page):
    """Coordinate normalizzate della pagina visualizzata (u verso destra, v verso il basso)
    -> coordinate PDF della pagina, tenendo conto di cropbox e /Rotate."""
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
    """Pagina PDF trasparente con i disegni (vettoriali)."""
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


def _evaluation_pages(student, questions, scores, comments, fonts):
    regular, bold = fonts
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    left, right = 50, W - 50

    c.setFont(bold, 18)
    c.drawCentredString(W / 2, H - 50, "VALUTAZIONE")
    c.line(left, H - 62, right, H - 62)

    y = H - 85
    c.setFont(regular, 11)
    who = f"Compito {student['compito']}"
    name = f"{student.get('cognome', '')} {student.get('nome', '')}".strip()
    if name:
        who += f" - {name}"
    if student.get("matricola"):
        who += f" - Matricola {student['matricola']}"
    c.drawString(left, y, who)
    y -= 35

    c.setFont(bold, 14)
    c.drawString(left, y, "PUNTEGGI")
    y -= 25
    c.setFont(regular, 12)
    total = 0.0
    for (qname, _, qmax), s in zip(questions, scores):
        c.drawString(70, y, f"{qname}: {fmt_num(s)} / {fmt_num(qmax)}")
        total += s
        y -= 22
    y -= 8
    c.setFont(bold, 13)
    c.drawString(70, y, f"Totale domande aperte: {fmt_num(total)} / "
                        f"{fmt_num(sum(q[2] for q in questions))}")
    y -= 40

    c.setFont(bold, 14)
    c.drawString(left, y, "COMMENTI")
    y -= 25
    c.setFont(regular, 11)
    comments = comments.strip()
    if not comments:
        c.drawString(70, y, "Nessun commento")
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
                     scores, comments, fonts):
    """Crea il PDF corretto a partire dall'originale (che non viene modificato)."""
    reader = PdfReader(original)
    writer = PdfWriter()
    for i, page in enumerate(reader.pages, start=1):
        strokes = strokes_by_page.get(str(i)) or strokes_by_page.get(i)
        if strokes:
            page.merge_page(_overlay_page(page, strokes))
        writer.add_page(page)
    for p in _evaluation_pages(student, questions, scores, comments, fonts):
        writer.add_page(p)

    fd, tmp = tempfile.mkstemp(suffix=".pdf", dir=os.path.dirname(output))
    with os.fdopen(fd, "wb") as f:
        writer.write(f)
    os.replace(tmp, output)


# ==================================================
# Applicazione
# ==================================================

class PDFAnnotatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Correzione compiti PDF")
        self.root.geometry("1400x900")
        self.ok = False

        self.config = load_config()
        self.fonts = _register_font()

        self.excel = None
        self.folder = ""
        self.pdf_files = []      # lista di (numero_compito, percorso)
        self.current_index = 0
        self.current_page = 1
        self.total_pages = 1
        self.zoom = None
        self.page_image = None   # pagina renderizzata (PIL, a RENDER_DPI)
        self.disp_size = (1, 1)
        self.state = {"files": {}}
        self.strokes = {}        # {"pagina": [[(u, v), ...], ...]} normalizzate 0..1
        self.current_stroke = None
        self.current_stroke_id = None

        if not self.open_excel() or not self.open_folder():
            return
        self.build_ui()
        self.current_index = self.first_ungraded_index()
        self.load_pdf()
        self.ok = True

    # --------------------------------------------------
    # Setup
    # --------------------------------------------------

    def open_excel(self):
        last = self.config.get("excel")
        path = filedialog.askopenfilename(
            title="Seleziona il file Excel con i punteggi",
            initialdir=os.path.dirname(last) if last else None,
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tutti i file", "*.*")])
        if not path:
            return False
        try:
            self.excel = ExcelGrades(path)
        except Exception as e:
            messagebox.showerror("Errore Excel", f"Impossibile leggere il file Excel:\n{e}")
            return False

        qs = ", ".join(
            f"{q[0]} (max {fmt_num(q[2])}"
            + (f", min {fmt_num(self.excel.minimums[q[0]])})" if self.excel.minimums[q[0]] else ")")
            for q in self.excel.questions)
        msg = f"Domande da correggere:\n{qs}\n\nStudenti trovati: {len(self.excel.rows)}"
        if self.excel.duplicates:
            msg += f"\n\nATTENZIONE: numeri di compito duplicati: {self.excel.duplicates}"
        if self.excel.lock_file_exists():
            msg += ("\n\nATTENZIONE: il file sembra aperto in Excel. Chiudilo prima di "
                    "salvare, altrimenti Excel potrebbe sovrascrivere i punteggi.")
        if not messagebox.askokcancel("File Excel", msg):
            return False
        self.config["excel"] = path
        save_config(self.config)
        return True

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
            elif n not in self.excel.rows:
                not_in_excel.append(f)
            elif n in seen:
                messagebox.showerror("Errore", f"Due PDF per il compito {n}:\n{seen[n]}\n{f}")
                return False
            else:
                seen[n] = f
                files.append((n, os.path.join(folder, f)))
        if not files:
            messagebox.showerror("Errore", "Nessun PDF associabile a un compito dell'Excel.")
            return False
        files.sort()

        warn = []
        if no_number:
            warn.append("PDF senza numero nel nome (ignorati):\n  " + "\n  ".join(no_number))
        if not_in_excel:
            warn.append("PDF con numero non presente nell'Excel (ignorati):\n  "
                        + "\n  ".join(not_in_excel))
        missing = sorted(set(self.excel.rows) - set(seen))
        if missing:
            warn.append(f"Compiti nell'Excel senza PDF: {missing}")
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
            if not self.excel.is_graded(n):
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
        tk.Button(tb, text="Annulla tratto", command=self.undo).pack(side=tk.LEFT, padx=2)
        tk.Button(tb, text="Cancella pagina", command=self.clear_annotations).pack(side=tk.LEFT, padx=2)

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

        self.canvas.bind("<ButtonPress-1>", self.start_draw)
        self.canvas.bind("<B1-Motion>", self.draw_motion)
        self.canvas.bind("<ButtonRelease-1>", self.end_draw)
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self.on_wheel_h)
        self.canvas.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Button-5>", lambda e: self.canvas.yview_scroll(3, "units"))

        # ---- Valutazioni ----
        right = tk.Frame(main, width=420)
        right.pack(side=tk.RIGHT, fill="y", padx=(10, 0))
        right.pack_propagate(False)

        self.label_student = tk.Label(right, font=("Arial", 11), justify="left",
                                      anchor="w", wraplength=400)
        self.label_student.pack(fill="x", padx=5, pady=(0, 5))

        sf = tk.LabelFrame(right, text="Punteggi", font=("Arial", 10, "bold"))
        sf.pack(fill="x", padx=5, pady=5)
        self.score_entries = []
        for i, (qname, _, qmax) in enumerate(self.excel.questions):
            row = tk.Frame(sf)
            row.pack(fill="x", padx=10, pady=2)
            qmin = self.excel.minimums[qname]
            rng = f"max {fmt_num(qmax)}" if qmin == 0 else f"da {fmt_num(qmin)} a {fmt_num(qmax)}"
            tk.Label(row, text=f"{qname}  ({rng}):", width=18,
                     anchor="w").pack(side=tk.LEFT)
            e = tk.Entry(row, width=8, font=("Arial", 11))
            e.pack(side=tk.LEFT)
            e.bind("<Return>", lambda ev, idx=i: self.focus_next_entry(idx))
            e.bind("<KeyRelease>", self.calculate_total)
            self.score_entries.append(e)
        trow = tk.Frame(sf)
        trow.pack(fill="x", padx=10, pady=(6, 4))
        tk.Label(trow, text="Totale:", width=18, anchor="w",
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        self.total_label = tk.Label(trow, text="0", font=("Arial", 11, "bold"), fg="green")
        self.total_label.pack(side=tk.LEFT)

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

        default_comments = ["Risposta incompleta"]
        cbf = tk.Frame(right)
        cbf.pack(fill="x", padx=5, pady=(0, 5))
        for comment in default_comments:
            tk.Button(cbf, text=comment[:20],
                      command=lambda c=comment: self.insert_comment(c)).pack(side=tk.LEFT, padx=2)
        tk.Button(cbf, text="Copia commento precedente",
                  command=self.copy_previous_comment).pack(side=tk.LEFT, padx=2)

        # Frasi già usate negli altri compiti (doppio clic per inserire)
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

    # --------------------------------------------------
    # Stato (disegni e commenti, per file, nella cartella dei PDF)
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
        """Memorizza disegni e commenti del compito corrente (non i punteggi)."""
        if not self.pdf_files:
            return
        name = os.path.basename(self.current_path())
        entry = self.state["files"].setdefault(name, {})
        entry["strokes"] = self.strokes
        entry["comments"] = self.text_box.get("1.0", tk.END).strip()
        self.save_state()

    # --------------------------------------------------
    # File correnti
    # --------------------------------------------------

    def current_number(self):
        return self.pdf_files[self.current_index][0]

    def current_path(self):
        return self.pdf_files[self.current_index][1]

    def original_path(self, path):
        return os.path.join(os.path.dirname(path), BACKUP_DIR, os.path.basename(path))

    def source_path(self, path):
        """Il PDF da mostrare/elaborare: l'originale se già salvato, altrimenti il file."""
        orig = self.original_path(path)
        return orig if os.path.exists(orig) else path

    def load_pdf(self):
        path = self.current_path()
        n = self.current_number()
        name = os.path.basename(path)
        saved = self.state["files"].get(name, {})
        self.strokes = {k: [list(map(tuple, s)) for s in v]
                        for k, v in saved.get("strokes", {}).items()}

        self.text_box.delete("1.0", tk.END)
        if saved.get("comments"):
            self.text_box.insert("1.0", saved["comments"])
        self.text_box.edit_reset()
        self.phrase_filter.delete(0, tk.END)
        self.refresh_phrases()

        scores = self.excel.get_scores(n)
        for e, s in zip(self.score_entries, scores):
            e.delete(0, tk.END)
            if s is not None:
                e.insert(0, fmt_num(s))
        self.calculate_total()

        st = self.excel.student(n)
        self.label_file.config(text=f"Compito {n}  —  {st['cognome']} {st['nome']}")
        info = f"Matricola: {st['matricola']}\nFile: {name}"
        if st["note"]:
            info += f"\nNote: {st['note']}"
        if self.excel.is_graded(n):
            info += "\n(già corretto: i punteggi vengono dall'Excel)"
        self.label_student.config(text=info)
        graded = sum(1 for k, _ in self.pdf_files if self.excel.is_graded(k))
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

    # --------------------------------------------------
    # Visualizzazione
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
        if len(pts) >= 4:
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
    # Disegno (coordinate salvate normalizzate 0..1: indipendenti dallo zoom)
    # --------------------------------------------------

    def _norm(self, e):
        w, h = self.disp_size
        x, y = self.canvas.canvasx(e.x), self.canvas.canvasy(e.y)
        return (min(max(x / w, 0.0), 1.0), min(max(y / h, 0.0), 1.0))

    def start_draw(self, e):
        self.current_stroke = [self._norm(e)]
        self.current_stroke_id = None

    def draw_motion(self, e):
        if self.current_stroke is None:
            return
        self.current_stroke.append(self._norm(e))
        w, h = self.disp_size
        pts = [c for (u, v) in self.current_stroke for c in (u * w, v * h)]
        if self.current_stroke_id is None:
            self.current_stroke_id = self.canvas.create_line(
                *pts, fill="blue", width=PEN_WIDTH_SCREEN, capstyle="round", joinstyle="round")
        else:
            self.canvas.coords(self.current_stroke_id, *pts)

    def end_draw(self, e):
        stroke = self.current_stroke
        self.current_stroke = None
        if self.current_stroke_id is not None:
            self.canvas.delete(self.current_stroke_id)
            self.current_stroke_id = None
        if stroke and len(stroke) > 1:
            self.strokes.setdefault(str(self.current_page), []).append(stroke)
            self._draw_stroke(stroke)

    def undo(self):
        page = self.strokes.get(str(self.current_page))
        if page:
            page.pop()
            self.show_page()

    def clear_annotations(self):
        if self.strokes.get(str(self.current_page)) and \
                messagebox.askyesno("Conferma", "Cancellare tutti i disegni di questa pagina?"):
            self.strokes[str(self.current_page)] = []
            self.show_page()

    # --------------------------------------------------
    # Punteggi e commenti
    # --------------------------------------------------

    def calculate_total(self, event=None):
        total = 0.0
        for e in self.score_entries:
            try:
                total += parse_num(e.get() or "0")
            except ValueError:
                pass
        self.total_label.config(text=f"{fmt_num(total)} / {fmt_num(self.excel.max_total)}")

    def focus_next_entry(self, idx):
        if idx + 1 < len(self.score_entries):
            self.score_entries[idx + 1].focus_set()
            self.score_entries[idx + 1].select_range(0, tk.END)

    def insert_comment(self, comment):
        # Va a capo se il cursore non è a inizio riga
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
        """Elenca le righe di commento usate negli altri compiti, le più frequenti in alto."""
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
        """Punteggi validati, oppure None (dopo aver mostrato l'errore)."""
        scores = []
        for e, (qname, _, qmax) in zip(self.score_entries, self.excel.questions):
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
            qmin = self.excel.minimums[qname]
            if s < qmin or s > qmax:
                messagebox.showwarning("Errore",
                                       f"{qname}: il punteggio deve essere tra "
                                       f"{fmt_num(qmin)} e {fmt_num(qmax)}")
                e.focus_set()
                return None
            scores.append(s)
        return scores

    # --------------------------------------------------
    # Salvataggio
    # --------------------------------------------------

    def save_current(self):
        """Scrive i punteggi nell'Excel e rigenera il PDF. True se riuscito."""
        scores = self.read_scores()
        if scores is None:
            return False
        n = self.current_number()
        path = self.current_path()

        while self.excel.lock_file_exists():
            if not messagebox.askretrycancel(
                    "File Excel aperto",
                    "Il file Excel è aperto in Excel. Chiudilo (salvando eventuali tue "
                    "modifiche) e premi Riprova."):
                return False

        self.remember_current()
        try:
            self.excel.set_scores(n, scores)
            self.excel.save()
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile salvare l'Excel:\n{e}")
            return False

        try:
            orig = self.original_path(path)
            if not os.path.exists(orig):
                os.makedirs(os.path.dirname(orig), exist_ok=True)
                shutil.copy2(path, orig)
            build_graded_pdf(orig, path, self.strokes, self.excel.student(n),
                             self.excel.questions, scores,
                             self.text_box.get("1.0", tk.END), self.fonts)
        except Exception as e:
            messagebox.showerror("Errore",
                                 f"Punteggi salvati nell'Excel, ma impossibile salvare il PDF:\n{e}")
            return False
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
            missing = [k for k, _ in self.pdf_files if not self.excel.is_graded(k)]
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
        self.root.destroy()

    def on_closing(self):
        self.save_and_exit()


# ==================================================
# Avvio
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
