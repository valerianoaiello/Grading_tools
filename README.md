
This repository contains the first tool in a series of small Python utilities designed to assist with academic and administrative tasks.

## Grader for open ended questions

A Python desktop application for grading PDF exams with annotations, scoring, and feedback.

### Features
- Reads students and questions from an Excel file (sheet `Punteggi`) or a CSV (see below):
  row 1 holds the headers (`Compito`, `Matricola`, `Cognome`, `Nome`, …), the row whose
  `Compito` cell is `massimo` holds the maximum score of each question. Every column with a
  numeric maximum is a question to grade (e.g. `G1`, `G2`, `7A`…`7D`).
- Matches each PDF to a student by exam number: the last group of digits in the file name
  (`compito_12.pdf`, `261103_12.pdf` and `12.pdf` all map to exam 12)
- Writes each score into the right Excel cell and checks it does not exceed the maximum;
  a timestamped backup of the Excel file is made at the first save of each session
- Draw annotations on PDF pages (stored as vectors, independent of zoom), with the mouse or a
  pen tablet (Wacom…): pen tip draws, a tap leaves a dot, the side button (right click) erases
  the strokes it touches; *Gomma* turns the pen tip into an eraser; *Annulla* (Cmd/Ctrl+Z)
  undoes the last stroke, erase or page clear
- **Solutions window** (*Soluzioni ⧉*): a second window showing the solutions of the exam being
  graded, following the main window. The zip holds one `.tex` (or `.pdf`) per exam, named after
  the exam number; the `.tex` files are compiled with `pdflatex` in the background into
  `soluzioni_pdf/` in the PDF folder. The window position is remembered (e.g. on a second screen)
- Comments are appended to the PDF as an evaluation page with name, scores and total
- **Complete PDF for the student**, built on save:
  1. *Pagina V/F* (CSV of Correggi VF only): a first page with the true/false part, i.e.
     answers given, correct answers, outcome and points of each question, and the true/false
     score. The columns are guessed from their names (`Compito`, `Risposta 1A`…, `Punti 1A`…,
     `Totale`, `Risposte giuste`); *Colonne V/F…* shows them, with the first data row as an
     example, and lets you change them (the choice is remembered for the folder). The maximum
     of each question is the best score obtained in the class
  2. the exam with your drawings, then the evaluation page
  3. *Soluzioni in coda*: the solutions of that exam (from `soluzioni.zip`, see below)

  Both extra parts are in English for the exams marked *Inglese*. *Rigenera PDF…* rebuilds
  the PDF of every exam already graded with the current options (scores, comments, drawings
  and language stay as saved)
- Originals are kept in an `originali/` subfolder, so an exam can be regraded without
  duplicating the evaluation page; drawings and comments are kept in `.correzione_stato.json`
- Resumes from the first exam without scores; "Vai al compito…" jumps to any exam number
- Also reads the CSV written by *Correggi VF* (`risposte_totali.csv`, column `Numero compito`):
  the open questions are then defined in the grading rubric, and their columns (plus
  `Totale domande aperte`) are added to the CSV. A timestamped backup is made first
- **Grading rubric** (`griglia.txt` in the PDF folder, asked at startup if missing, editable
  with *Griglia…*): for each exercise, the points given for each part of the answer
- **Proposals** (`proposte.json` in the PDF folder): proposed scores (highlighted in yellow)
  and comment, written by Claude after reading the scans; check and edit them, nothing is
  written to the Excel/CSV file until you save
- **Inglese** switch: English comment proposal and English evaluation page, remembered per exam

### Grading rubric format
```
[Esercizio 1]
1     equazione caratteristica corretta
0,5   soluzione generale
1     problema di Cauchy risolto
-0,25 ogni errore di calcolo
Note for the grader: any line that does not start with a number.

[Esercizio 2]
...
```
The name in brackets is the column name (for an Excel file it must match a column with a
`massimo`). The maximum of an exercise is the sum of its positive points.

### Proposals with Claude Code
The scans are read by Claude in a Claude Code session (no API key needed). Give Claude the
PDF folder (with `griglia.txt` and `soluzioni.zip`: the grader asks for the zip at startup
and copies it there, or use *Soluzioni…*), and ask for the
proposals. Claude writes `proposte.json`, which the grader re-reads whenever it changes:
```json
{
  "12": {
    "punteggi": {"Esercizio 1": 2, "Esercizio 2": 1.5},
    "commento": "Esercizio 1: corretto. Esercizio 2: manca la verifica della condizione iniziale.",
    "comment_en": "Exercise 1: correct. Exercise 2: the initial condition is not checked.",
    "dubbi": "Pagina 3 poco leggibile: controlla il segno nel passaggio finale."
  }
}
```
Keys are exam numbers. `dubbi` is shown only to you, never written in the PDF. Proposed
scores outside the allowed range are ignored. Proposals are not used for exams already graded.

### Requirements
```bash
pip install PyPDF2 pdf2image reportlab openpyxl pillow
```
`pdf2image` also needs Poppler (`brew install poppler` on macOS).
Close the Excel file in Excel while grading: the program refuses to save while it is open.



## Email Sender 

This script allows you to send **personalized emails** to students listed in a CSV file using the **Gmail API**.  
Each email is sent **individually**, with optional attachments, and supports simple text customization.

### Features

- Select a CSV file containing student data
- Specify which CSV columns correspond to:
  - First name
  - Last name
  - Email address
- Write a common email body with placeholders:
  - `[name]`
  - `[surname]`
  - `[email]`
- Select a Gmail **credentials JSON file** (OAuth 2.0). `credential.example.json` shows the
  expected format; keep your real credentials outside the repository (`credential*.json`
  and `token*.json` are git-ignored)
- Attach one or more files that will be sent to **all recipients**
- Simple graphical user interface (Tkinter)
- Emails are sent one by one (not as a mailing list)

### Example CSV Structure

The CSV file can have **any structure**, as long as you know the column indices.

Example:

```csv
id,name,surname,email
1,Mario,Rossi,mario.rossi@example.com
2,Luisa,Bianchi,luisa.bianchi@example.com
```

## Acknowledgments

The grading tool was substantially improved with the help of Claude, made available through
the **Claude Team plan for scientists** by Anthropic. Claude helped connect the grader to the
Excel class list, fix several bugs in the original program, and add new features such as
negative scores and reusable comment suggestions.

## License

Released under the [MIT License](LICENSE).
