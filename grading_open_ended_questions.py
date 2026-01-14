"""
Docstring per Nuovo.grading_open_ended_questions
This application allows teachers to grade open-ended questions in PDF assignments. 
It provides tools to annotate student submissions, input scores for each question, and add comments. 
The application saves the graded PDFs with annotations and a summary page containing scores and comments. 
It also maintains a CSV file with all scores for easy record-keeping. 
"""
import os
import csv
import json
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext
from datetime import datetime
import textwrap

from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

from pdf2image import convert_from_path
from PIL import Image, ImageTk, ImageDraw

STATE_FILE = "stato_lavoro.json"
CSV_FILE = "valutazioni.csv"


class PDFAnnotatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Correzione compiti PDF")
        self.root.geometry("1400x900")  # Dimensioni più adatte
        
        # Variabili di stato
        self.pdf_files = []
        self.current_index = 0
        self.current_page = 1
        self.total_pages = 1
        self.folder = ""
        
        # Per annotazioni - ora memorizziamo per pagina
        self.zoom = None
        self.lines = {}  # Dizionario: {numero_pagina: [lista_di_linee]}
        self.current_line = None
        self.undo_stack = {}  # Dizionario per undo multipli per pagina
        
        # Per valutazioni
        self.num_questions = self.ask_num_questions()
        if self.num_questions is None:
            self.root.quit()
            return
            
        self.score_entries = []
        self.comments_history = {}  # Per memorizzare commenti tra i file
        
        self.build_ui()
        self.load_state()  # Carica stato precedente se esiste
        self.select_folder()

    # --------------------------------------------------
    # Setup
    # --------------------------------------------------

    def ask_num_questions(self):
        """Chiede il numero di domande con valore di default"""
        n = simpledialog.askinteger(
            "Domande aperte",
            "Numero di domande aperte:",
            minvalue=1,
            initialvalue=4
        )
        return n

    # --------------------------------------------------
    # UI
    # --------------------------------------------------

    def build_ui(self):
        """Costruisce l'interfaccia grafica"""
        # Barra superiore
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill="x", padx=5, pady=5)
        
        self.label_file = tk.Label(top_frame, font=("Arial", 12, "bold"))
        self.label_file.pack(side=tk.LEFT)
        
        self.label_status = tk.Label(top_frame, font=("Arial", 10), fg="blue")
        self.label_status.pack(side=tk.RIGHT)
        
        # Frame principale
        main = tk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=5, pady=5)

        # -------- AREA PDF E ANNOTAZIONI --------
        left = tk.Frame(main)
        left.pack(side=tk.LEFT, fill="both", expand=True)

        # Toolbar per annotazioni
        toolbar = tk.Frame(left)
        toolbar.pack(fill="x", pady=(0, 5))
        
        tk.Button(toolbar, text="◀ Pagina", command=self.prev_page).pack(side=tk.LEFT, padx=2)
        self.page_label = tk.Label(toolbar, text="1/1", width=10)
        self.page_label.pack(side=tk.LEFT, padx=2)
        tk.Button(toolbar, text="Pagina ▶", command=self.next_page).pack(side=tk.LEFT, padx=2)
        
        tk.Label(toolbar, text="|").pack(side=tk.LEFT, padx=5)
        
        tk.Button(toolbar, text="Zoom +", command=self.zoom_in).pack(side=tk.LEFT, padx=2)
        tk.Button(toolbar, text="Zoom −", command=self.zoom_out).pack(side=tk.LEFT, padx=2)
        tk.Button(toolbar, text="Zoom Reset", command=self.zoom_reset).pack(side=tk.LEFT, padx=2)
        
        tk.Label(toolbar, text="|").pack(side=tk.LEFT, padx=5)
        
        tk.Button(toolbar, text="Annulla", command=self.undo).pack(side=tk.LEFT, padx=2)
        tk.Button(toolbar, text="Cancella Tutto", command=self.clear_annotations).pack(side=tk.LEFT, padx=2)

        # Canvas per visualizzazione PDF
        canvas_frame = tk.Frame(left, relief="sunken", borderwidth=2)
        canvas_frame.pack(fill="both", expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, bg="gray", cursor="cross")
        self.canvas.pack(fill="both", expand=True)
        
        # Bind eventi per disegno
        self.canvas.bind("<ButtonPress-1>", self.start_draw)
        self.canvas.bind("<B1-Motion>", self.draw_motion)
        self.canvas.bind("<ButtonRelease-1>", self.end_draw)
        
        # Sostituiamo le scrollbar con frecce di navigazione
        nav_frame = tk.Frame(canvas_frame)
        nav_frame.place(relx=0.5, rely=1.0, anchor="s", y=-10)
        
        self.scroll_left_btn = tk.Button(nav_frame, text="◀", width=3, 
                                        command=self.scroll_left)
        self.scroll_left_btn.pack(side=tk.LEFT, padx=2)
        
        self.scroll_up_btn = tk.Button(nav_frame, text="▲", width=3,
                                      command=self.scroll_up)
        self.scroll_up_btn.pack(side=tk.LEFT, padx=2)
        
        self.scroll_down_btn = tk.Button(nav_frame, text="▼", width=3,
                                        command=self.scroll_down)
        self.scroll_down_btn.pack(side=tk.LEFT, padx=2)
        
        self.scroll_right_btn = tk.Button(nav_frame, text="▶", width=3,
                                         command=self.scroll_right)
        self.scroll_right_btn.pack(side=tk.LEFT, padx=2)

        # -------- AREA VALUTAZIONI --------
        right = tk.Frame(main, width=400)
        right.pack(side=tk.RIGHT, fill="y", padx=(10, 0))
        right.pack_propagate(False)

        # Frame punteggi
        scores_frame = tk.LabelFrame(right, text="Punteggi", font=("Arial", 10, "bold"))
        scores_frame.pack(fill="x", padx=5, pady=5)
        
        for i in range(self.num_questions):
            score_frame = tk.Frame(scores_frame)
            score_frame.pack(fill="x", padx=10, pady=2)
            
            tk.Label(score_frame, text=f"Domanda {i+1}:", width=12, anchor="w").pack(side=tk.LEFT)
            e = tk.Entry(score_frame, width=8, font=("Arial", 10))
            e.pack(side=tk.LEFT)
            e.bind("<Return>", lambda event, idx=i: self.focus_next_entry(idx))
            self.score_entries.append(e)
            
            # Label per il punteggio totale
            if i == self.num_questions - 1:
                tk.Label(score_frame, text="Totale:", width=12, anchor="w", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(20, 0))
                self.total_label = tk.Label(score_frame, text="0", font=("Arial", 10, "bold"), fg="green")
                self.total_label.pack(side=tk.LEFT)
                
                # Bind per calcolo automatico totale
                for entry in self.score_entries:
                    entry.bind("<KeyRelease>", self.calculate_total)

        # Pulsanti di navigazione
        nav_frame = tk.Frame(right)
        nav_frame.pack(fill="x", padx=5, pady=5)
        
        tk.Button(nav_frame, text="◀ Compito precedente", 
                 command=self.previous_student, width=18).pack(side=tk.LEFT, padx=2)
        tk.Button(nav_frame, text="Salva e continua ▶", 
                 command=self.save_and_next, width=18, bg="lightblue").pack(side=tk.RIGHT, padx=2)
        
        tk.Button(nav_frame, text="Salva e esci", 
                 command=self.save_and_exit, width=18, bg="lightgreen").pack(side=tk.BOTTOM, pady=5)

        # Commenti
        tk.Label(right, text="Commenti:", font=("Arial", 10, "bold")).pack(anchor="w", padx=5, pady=(10, 0))
        
        comment_frame = tk.Frame(right)
        comment_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.text_box = scrolledtext.ScrolledText(comment_frame, width=45, height=20, 
                                                 font=("Arial", 10), wrap="word")
        self.text_box.pack(fill="both", expand=True)
        
        # Pulsanti commenti predefiniti
        default_comments = ["Risposta incompleta"]
#        default_comments = ["Ottimo lavoro!", "Da migliorare", "Risposta incompleta"]
        comment_btn_frame = tk.Frame(right)
        comment_btn_frame.pack(fill="x", padx=5, pady=5)
        
        for comment in default_comments:
            btn = tk.Button(comment_btn_frame, text=comment[:15], 
                           command=lambda c=comment: self.insert_comment(c),
                           width=15)
            btn.pack(side=tk.LEFT, padx=2)

    # --------------------------------------------------
    # Gestione stato
    # --------------------------------------------------

    def load_state(self):
        """Carica lo stato precedente se esiste"""
        if os.path.exists(STATE_FILE):
            try:
                with open(STATE_FILE, 'r') as f:
                    state = json.load(f)
                    self.current_index = state.get('current_index', 0)
                    self.num_questions = state.get('num_questions', self.num_questions)
                    self.comments_history = state.get('comments_history', {})
            except:
                pass

    def save_state(self):
        """Salva lo stato corrente"""
        state = {
            'current_index': self.current_index,
            'num_questions': self.num_questions,
            'comments_history': self.comments_history,
            'folder': self.folder,
            'timestamp': datetime.now().isoformat()
        }
        with open(STATE_FILE, 'w') as f:
            json.dump(state, f)

    # --------------------------------------------------
    # Gestione cartella e file
    # --------------------------------------------------

    def select_folder(self):
        """Seleziona la cartella con i PDF"""
        initial_dir = self.folder if self.folder else None
        folder = filedialog.askdirectory(
            title="Seleziona cartella con i PDF",
            initialdir=initial_dir
        )
        
        if not folder:
            if not self.folder:
                self.root.quit()
            return
            
        self.folder = folder
        self.pdf_files = [
            os.path.join(self.folder, f)
            for f in sorted(os.listdir(self.folder))
            if f.lower().endswith(".pdf")
        ]

        if not self.pdf_files:
            messagebox.showerror("Errore", "Nessun PDF trovato nella cartella")
            self.root.quit()
            return

        self.label_status.config(text=f"{len(self.pdf_files)} file trovati")
        self.load_pdf()

    def load_pdf(self):
        """Carica il PDF corrente"""
        if self.current_index >= len(self.pdf_files):
            return
            
        # Reset delle annotazioni per il nuovo file
        self.lines = {}
        self.undo_stack = {}
        self.zoom = None
        
        path = self.pdf_files[self.current_index]
        reader = PdfReader(path)
        self.total_pages = len(reader.pages)
        self.current_page = 1

        # Aggiorna UI
        filename = os.path.basename(path)
        self.label_file.config(text=f"{filename} ({self.current_index + 1}/{len(self.pdf_files)})")
        
        # Ripristina commenti precedenti se esistenti
        if filename in self.comments_history:
            self.text_box.delete("1.0", tk.END)
            self.text_box.insert("1.0", self.comments_history[filename])
        
        # Pulisci i punteggi
        for entry in self.score_entries:
            entry.delete(0, tk.END)
        self.calculate_total()
        
        self.show_page(initial=True)

    # --------------------------------------------------
    # Visualizzazione PDF
    # --------------------------------------------------

    def show_page(self, initial=False):
        """Mostra la pagina corrente"""
        if not self.pdf_files:
            return
            
        try:
            # Aumenta la risoluzione a 300 DPI per visualizzazione migliore
            img = convert_from_path(
                self.pdf_files[self.current_index],
                first_page=self.current_page,
                last_page=self.current_page,
                dpi=300  # Aumentata la risoluzione per meno sfocatura
            )[0]
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile caricare il PDF: {str(e)}")
            return

        if initial or self.zoom is None:
            # Calcola zoom iniziale
            self.root.update_idletasks()
            cw = self.canvas.winfo_width() - 20
            ch = self.canvas.winfo_height() - 20

            if cw < 100 or ch < 100:
                self.zoom = 0.7
            else:
                self.zoom = min(cw / img.width, ch / img.height) * 0.9

        self.zoom = max(self.zoom, 0.1)
        self.zoom = min(self.zoom, 5.0)  # Limita zoom massimo

        # Ridimensiona immagine
        new_width = int(img.width * self.zoom)
        new_height = int(img.height * self.zoom)
        img = img.resize((new_width, new_height), Image.LANCZOS)

        self.base_image = img
        self.annotated_image = img.copy()
        self.draw = ImageDraw.Draw(self.annotated_image)

        # Ridisegna solo le annotazioni della pagina corrente
        page_lines = self.lines.get(self.current_page, [])
        for line in page_lines:
            if len(line) > 1:
                self.draw.line(line, fill="red", width=3)

        # Aggiorna canvas
        self.tk_img = ImageTk.PhotoImage(self.annotated_image)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)
        
        # Aggiorna dimensioni canvas
        self.canvas.config(scrollregion=(0, 0, new_width, new_height))
        
        # Aggiorna label pagina
        self.page_label.config(text=f"{self.current_page}/{self.total_pages}")

    # --------------------------------------------------
    # Navigazione
    # --------------------------------------------------

    def next_page(self):
        if self.current_page < self.total_pages:
            self.current_page += 1
            self.show_page()

    def prev_page(self):
        if self.current_page > 1:
            self.current_page -= 1
            self.show_page()

    def previous_student(self):
        """Torna al compito precedente"""
        if self.current_index > 0:
            self.save_current_state()
            self.current_index -= 1
            self.load_pdf()

    # --------------------------------------------------
    # Scorrimento con frecce
    # --------------------------------------------------

    def scroll_left(self):
        """Scorre a sinistra"""
        self.canvas.xview_scroll(-50, "units")

    def scroll_right(self):
        """Scorre a destra"""
        self.canvas.xview_scroll(50, "units")

    def scroll_up(self):
        """Scorre in alto"""
        self.canvas.yview_scroll(-50, "units")

    def scroll_down(self):
        """Scorre in basso"""
        self.canvas.yview_scroll(50, "units")

    # --------------------------------------------------
    # Annotazioni
    # --------------------------------------------------

    def start_draw(self, e):
        """Inizia a disegnare"""
        x = self.canvas.canvasx(e.x)
        y = self.canvas.canvasy(e.y)
        self.current_line = [(x, y)]

    def draw_motion(self, e):
        """Disegna durante il movimento del mouse"""
        if self.current_line:
            x = self.canvas.canvasx(e.x)
            y = self.canvas.canvasy(e.y)
            self.current_line.append((x, y))
            
            # Disegna sulla copia temporanea
            temp_img = self.annotated_image.copy()
            temp_draw = ImageDraw.Draw(temp_img)
            
            # Disegna tutte le linee esistenti per questa pagina
            page_lines = self.lines.get(self.current_page, [])
            for line in page_lines:
                if len(line) > 1:
                    temp_draw.line(line, fill="red", width=3)
            
            # Disegna la linea corrente
            if len(self.current_line) > 1:
                temp_draw.line(self.current_line, fill="blue", width=3)
            
            # Aggiorna display
            self.tk_img = ImageTk.PhotoImage(temp_img)
            self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)

    def end_draw(self, e):
        """Termina il disegno"""
        if self.current_line and len(self.current_line) > 1:
            # Salva la linea per la pagina corrente
            if self.current_page not in self.lines:
                self.lines[self.current_page] = []
                self.undo_stack[self.current_page] = []
            
            self.lines[self.current_page].append(self.current_line)
            self.undo_stack[self.current_page].append(self.current_line.copy())
            self.current_line = None
            self.show_page()  # Ridisegna con la nuova linea permanente

    def undo(self):
        """Annulla l'ultima annotazione per la pagina corrente"""
        if self.current_page in self.lines and self.lines[self.current_page]:
            self.lines[self.current_page].pop()
            if (self.current_page in self.undo_stack and 
                self.undo_stack[self.current_page]):
                self.undo_stack[self.current_page].pop()
            self.show_page()

    def clear_annotations(self):
        """Cancella tutte le annotazioni della pagina corrente"""
        if messagebox.askyesno("Conferma", "Cancellare tutte le annotazioni di questa pagina?"):
            if self.current_page in self.lines:
                self.lines[self.current_page].clear()
            if self.current_page in self.undo_stack:
                self.undo_stack[self.current_page].clear()
            self.show_page()

    # --------------------------------------------------
    # Zoom
    # --------------------------------------------------

    def zoom_in(self):
        self.zoom *= 1.2
        self.show_page()

    def zoom_out(self):
        self.zoom /= 1.2
        self.show_page()

    def zoom_reset(self):
        self.zoom = None
        self.show_page()

    # --------------------------------------------------
    # Gestione punteggi
    # --------------------------------------------------

    def calculate_total(self, event=None):
        """Calcola il punteggio totale"""
        total = 0
        for entry in self.score_entries:
            try:
                value = float(entry.get().replace(',', '.') or 0)
                total += value
            except ValueError:
                pass
        self.total_label.config(text=f"{total:.1f}".replace('.', ','))

    def focus_next_entry(self, current_idx):
        """Passa al campo successivo con Enter"""
        if current_idx + 1 < len(self.score_entries):
            self.score_entries[current_idx + 1].focus_set()

    # --------------------------------------------------
    # Gestione commenti
    # --------------------------------------------------

    def insert_comment(self, comment):
        """Inserisce un commento predefinito"""
        self.text_box.insert(tk.INSERT, comment + "\n")

    def save_current_state(self):
        """Salva lo stato corrente per ripristino"""
        if self.current_index < len(self.pdf_files):
            filename = os.path.basename(self.pdf_files[self.current_index])
            self.comments_history[filename] = self.text_box.get("1.0", tk.END).strip()

    # --------------------------------------------------
    # Salvataggio
    # --------------------------------------------------

    def save_and_next(self):
        """Salva il compito corrente e passa al successivo"""
        # Valida punteggi
        scores = []
        for i, entry in enumerate(self.score_entries):
            value = entry.get().strip()
            if not value:
                messagebox.showwarning("Errore", f"Inserisci il punteggio per la domanda {i+1}")
                entry.focus_set()
                return
            try:
                # Sostituisci virgola con punto per la conversione
                score_value = float(value.replace(',', '.'))
                scores.append(score_value)
            except ValueError:
                messagebox.showwarning("Errore", f"Punteggio non valido per la domanda {i+1}")
                entry.focus_set()
                return

        # Salva
        if self.save_pdf(scores):
            # Salva anche nel CSV
            self.save_to_csv(scores)
            
            # Salva commenti nella history
            filename = os.path.basename(self.pdf_files[self.current_index])
            self.comments_history[filename] = self.text_box.get("1.0", tk.END).strip()
            
            # Passa al prossimo
            self.current_index += 1
            if self.current_index >= len(self.pdf_files):
                messagebox.showinfo("Fine", "Tutti i compiti sono stati corretti!")
                self.save_state()
                self.root.quit()
            else:
                self.load_pdf()
                self.save_state()

    def save_and_exit(self):
        """Salva e esce dal programma"""
        if messagebox.askyesno("Conferma", "Salvare le modifiche prima di uscire?"):
            # Salva lo stato corrente
            self.save_state()
        self.root.quit()

        """Salva il PDF con annotazioni e pagina di valutazione"""
        try:
            input_pdf = self.pdf_files[self.current_index]
            reader = PdfReader(input_pdf)
            writer = PdfWriter()

            # Processa tutte le pagine
            for page_num in range(1, self.total_pages + 1):
                # Controlla se questa pagina ha annotazioni
                if page_num in self.lines and self.lines[page_num]:
                    # Converte la pagina a immagine ad alta risoluzione
                    img = convert_from_path(
                        input_pdf,
                        first_page=page_num,
                        last_page=page_num,
                        dpi=400  # Alta risoluzione per meno sfocatura
                    )[0]

                    # Disegna le annotazioni sull'immagine
                    draw = ImageDraw.Draw(img)
                    page_lines = self.lines[page_num]
                    
                    for line in page_lines:
                        if len(line) > 1:
                            # Converti le coordinate del canvas alle coordinate dell'immagine originale
                            img_width, img_height = img.size
                            canvas_width, canvas_height = self.base_image.size
                            
                            scaled_line = []
                            for (x, y) in line:
                                # Scala le coordinate dal canvas all'immagine originale
                                x_scaled = (x / canvas_width) * img_width
                                y_scaled = (y / canvas_height) * img_height
                                scaled_line.append((x_scaled, y_scaled))
                            
                            draw.line(scaled_line, fill="red", width=8)

                    # Salva l'immagine con annotazioni come PDF
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as imgf:
                        img.save(imgf.name, "PNG", dpi=(400, 400))

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as pdff:
                        # Usa le dimensioni dell'immagine originale
                        img_width, img_height = img.size
                        # Converti da pixel a punti (72 punti per pollice)
                        page_width = img_width * 72 / 400
                        page_height = img_height * 72 / 400
                        
                        c = canvas.Canvas(pdff.name, pagesize=(page_width, page_height))
                        c.drawImage(imgf.name, 0, 0, width=page_width, height=page_height)
                        c.showPage()
                        c.save()

                    ann_page = PdfReader(pdff.name).pages[0]
                    writer.add_page(ann_page)
                    
                    # Pulisci file temporanei
                    os.unlink(imgf.name)
                    os.unlink(pdff.name)
                else:
                    # Se non ci sono annotazioni, aggiungi la pagina originale
                    writer.add_page(reader.pages[page_num - 1])

            # Aggiungi la pagina di valutazione alla FINE
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as evalf:
                c = canvas.Canvas(evalf.name, pagesize=A4)
                
                # Titolo
                c.setFont("Helvetica-Bold", 18)
                c.drawCentredString(A4[0]/2, A4[1] - 50, "VALUTAZIONE")
                
                # Linea separatrice
                c.line(50, A4[1] - 70, A4[0] - 50, A4[1] - 70)
                
                # Punteggi
                y = A4[1] - 100
                c.setFont("Helvetica-Bold", 14)
                c.drawString(50, y, "PUNTEGGI")
                y -= 30
                
                c.setFont("Helvetica", 12)
                total = 0
                for i, s in enumerate(scores):
                    c.drawString(70, y, f"Domanda {i+1}: {s:.1f}".replace('.', ','))
                    total += s
                    y -= 25
                
                # Totale
                y -= 10
                c.setFont("Helvetica-Bold", 14)
                c.drawString(70, y, f"TOTALE: {total:.1f}".replace('.', ','))
                y -= 40
                
                # Commenti
                c.drawString(50, y, "COMMENTI:")
                y -= 30
                
                c.setFont("Helvetica", 11)
                comments = self.text_box.get("1.0", tk.END).strip()
                
                # Divide i commenti in linee più gestibili con textwrap
                if comments:
                    # Usa textwrap per dividere il testo in linee di max 80 caratteri
                    wrapped_lines = textwrap.wrap(comments, width=80)
                    for line in wrapped_lines:
                        if y < 50:  # Nuova pagina se necessario
                            c.showPage()
                            y = A4[1] - 50
                            c.setFont("Helvetica", 11)
                        
                        c.drawString(70, y, line)
                        y -= 20
                else:
                    c.drawString(70, y, "Nessun commento")
                    y -= 20
                
                c.showPage()
                c.save()

            # Aggiungi la pagina di valutazione alla fine
            eval_reader = PdfReader(evalf.name)
            for page in eval_reader.pages:
                writer.add_page(page)

            # Salva su file temporaneo e poi sostituisci
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                with open(tmp.name, "wb") as f:
                    writer.write(f)
            
            # Sostituisci il file originale
            os.replace(tmp.name, input_pdf)
            
            # Pulisci file temporanei
            os.unlink(evalf.name)
            
            return True
            
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile salvare il PDF: {str(e)}")
            return False

    def save_pdf(self, scores):
        """Salva il PDF con annotazioni e pagina di valutazione"""
        try:
            input_pdf = self.pdf_files[self.current_index]
            reader = PdfReader(input_pdf)
            writer = PdfWriter()

            # Processa tutte le pagine
            for page_num in range(1, self.total_pages + 1):
                # Controlla se questa pagina ha annotazioni
                if page_num in self.lines and self.lines[page_num]:
                    # Converte la pagina a immagine ad alta risoluzione
                    img = convert_from_path(
                        input_pdf,
                        first_page=page_num,
                        last_page=page_num,
                        dpi=400  # Alta risoluzione per meno sfocatura
                    )[0]

                    # Disegna le annotazioni sull'immagine
                    draw = ImageDraw.Draw(img)
                    page_lines = self.lines[page_num]
                    
                    for line in page_lines:
                        if len(line) > 1:
                            # Converti le coordinate del canvas alle coordinate dell'immagine originale
                            img_width, img_height = img.size
                            canvas_width, canvas_height = self.base_image.size
                            
                            scaled_line = []
                            for (x, y) in line:
                                # Scala le coordinate dal canvas all'immagine originale
                                x_scaled = (x / canvas_width) * img_width
                                y_scaled = (y / canvas_height) * img_height
                                scaled_line.append((x_scaled, y_scaled))
                            
                            draw.line(scaled_line, fill="red", width=8)

                    # Salva l'immagine con annotazioni come PDF
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as imgf:
                        img.save(imgf.name, "PNG", dpi=(400, 400))

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as pdff:
                        # Usa le dimensioni dell'immagine originale
                        img_width, img_height = img.size
                        # Converti da pixel a punti (72 punti per pollice)
                        page_width = img_width * 72 / 400
                        page_height = img_height * 72 / 400
                        
                        c = canvas.Canvas(pdff.name, pagesize=(page_width, page_height))
                        c.drawImage(imgf.name, 0, 0, width=page_width, height=page_height)
                        c.showPage()
                        c.save()

                    ann_page = PdfReader(pdff.name).pages[0]
                    writer.add_page(ann_page)
                    
                    # Pulisci file temporanei
                    os.unlink(imgf.name)
                    os.unlink(pdff.name)
                else:
                    # Se non ci sono annotazioni, aggiungi la pagina originale
                    writer.add_page(reader.pages[page_num - 1])

            # Aggiungi la pagina di valutazione alla FINE
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as evalf:
                c = canvas.Canvas(evalf.name, pagesize=A4)
                
                # Titolo
                c.setFont("Helvetica-Bold", 18)
                c.drawCentredString(A4[0]/2, A4[1] - 50, "VALUTAZIONE")
                
                # Linea separatrice
                c.line(50, A4[1] - 70, A4[0] - 50, A4[1] - 70)
                
                # Punteggi
                y = A4[1] - 100
                c.setFont("Helvetica-Bold", 14)
                c.drawString(50, y, "PUNTEGGI")
                y -= 30
                
                c.setFont("Helvetica", 12)
                total = 0
                for i, s in enumerate(scores):
                    c.drawString(70, y, f"Domanda {i+1}: {s:.1f}".replace('.', ','))
                    total += s
                    y -= 25
                
                # Totale
                y -= 10
                c.setFont("Helvetica-Bold", 14)
                c.drawString(70, y, f"TOTALE: {total:.1f}".replace('.', ','))
                y -= 40
                
                # Commenti
                c.drawString(50, y, "COMMENTI:")
                y -= 30
                
                c.setFont("Helvetica", 11)
                comments = self.text_box.get("1.0", tk.END).strip()
                
                # Divide i commenti in linee più gestibili con textwrap
                if comments:
                    # Prima dividi in paragrafi (basati sui ritorni a capo esistenti)
                    paragraphs = comments.split('\n')
                    
                    for paragraph in paragraphs:
                        if not paragraph.strip():  # Se è una riga vuota
                            y -= 15  # Spazio extra tra paragrafi
                            continue
                        
                        # Usa textwrap per dividere ogni paragrafo in linee
                        wrapped_lines = textwrap.wrap(paragraph, width=80)
                        
                        for line in wrapped_lines:
                            if y < 50:  # Nuova pagina se necessario
                                c.showPage()
                                y = A4[1] - 50
                                c.setFont("Helvetica", 11)
                            
                            c.drawString(70, y, line)
                            y -= 20
                        
                        # Aggiungi uno spazio tra i paragrafi
                        y -= 5
                else:
                    c.drawString(70, y, "Nessun commento")
                    y -= 20
                
                c.showPage()
                c.save()

            # Aggiungi la pagina di valutazione alla fine
            eval_reader = PdfReader(evalf.name)
            for page in eval_reader.pages:
                writer.add_page(page)

            # Salva su file temporaneo e poi sostituisci
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                with open(tmp.name, "wb") as f:
                    writer.write(f)
            
            # Sostituisci il file originale
            os.replace(tmp.name, input_pdf)
            
            # Pulisci file temporanei
            os.unlink(evalf.name)
            
            return True
            
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile salvare il PDF: {str(e)}")
            return False
        
    def save_to_csv(self, scores):
        """Salva i punteggi in un file CSV con formato europeo (; come separatore, , come decimale)"""
        try:
            file_exists = os.path.exists(CSV_FILE)
            with open(CSV_FILE, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                
                if not file_exists:
                    # Intestazioni: solo File e le domande
                    headers = ["File"]
                    headers.extend([f"D{i+1}" for i in range(self.num_questions)])
                    writer.writerow(headers)
                
                # Dati: nome file senza estensione e punteggi
                filename = os.path.splitext(
                    os.path.basename(self.pdf_files[self.current_index])
                )[0]  # Rimuove l'estensione .pdf
                
                # Formatta i punteggi con virgola como separatore decimale
                formatted_scores = []
                for score in scores:
                    try:
                        # Formatta con virgola come separatore decimale
                        formatted_score = f"{score:.1f}".replace('.', ',')
                        formatted_scores.append(formatted_score)
                    except (ValueError, TypeError):
                        formatted_scores.append("0,0")
                
                row = [filename]
                row.extend(formatted_scores)
                
                writer.writerow(row)
                
        except Exception as e:
            print(f"Errore nel salvataggio CSV: {e}")
            messagebox.showwarning("Attenzione", 
                                 f"Impossibile salvare nel CSV: {str(e)}")

    # --------------------------------------------------
    # Cleanup
    # --------------------------------------------------

    def on_closing(self):
        """Gestisce la chiusura della finestra"""
        self.save_state()
        self.root.destroy()


# --------------------------------------------------
# Avvio programma
# --------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFAnnotatorApp(root)
    
    # Gestione chiusura finestra
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Bind scorciatoie tastiera
    root.bind("<Control-s>", lambda e: app.save_and_next())
    root.bind("<Control-z>", lambda e: app.undo())
    root.bind("<Control-plus>", lambda e: app.zoom_in())
    root.bind("<Control-minus>", lambda e: app.zoom_out())
    
    root.mainloop()