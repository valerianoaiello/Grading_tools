
This repository contains the first tool in a series of small Python utilities designed to assist with academic and administrative tasks.

## Grader for open ended questions

A Python desktop application for grading scanned PDF exams: draw on the pages, give
scores and comments, and build for each student a PDF with the graded exam, an evaluation
page and (optionally) the solutions and a true/false page.

### Requirements
- Python 3 with Tkinter, and the packages
  ```bash
  pip install PyPDF2 pdf2image reportlab openpyxl pillow
  ```
- **Poppler**, used by `pdf2image` to show the pages:
  `brew install poppler` (macOS), `sudo apt install poppler-utils` (Debian/Ubuntu),
  or the [Windows build](https://github.com/oschwartz10612/poppler-windows) added to `PATH`
- **LaTeX** (`pdflatex`), only if your solutions are `.tex` files: MacTeX on macOS,
  `texlive` on Linux, MiKTeX or TeX Live on Windows. Solutions given as PDF need nothing

### Try it with the demo
```bash
python3 demo/make_demo.py
python3 grading_open_ended_questions.py
```
Choose `demo/students.csv`, then the folder `demo/exams`, then `demo/solutions.zip`
(one PDF per exam) or `demo/solutions.pdf` (the same for everybody). The demo has three
made-up students, a rubric, example proposals and a true/false part.

### Input files
- **Exam PDFs**, one per student, all in one folder. The exam number is the last group of
  digits in the file name: `exam_12.pdf`, `261103_12.pdf` and `12.pdf` all map to exam 12.
- **Student list**, Excel (`.xlsx`) or CSV. Row 1 holds the column names, in English or
  Italian (case does not matter):

  | Field | Accepted names | |
  |---|---|---|
  | exam number | `Exam`, `Exam number`, `Compito`, `Numero compito` | required |
  | student ID | `Student ID`, `ID`, `Matricola` | optional |
  | surname | `Surname`, `Last name`, `Cognome` | optional |
  | first name | `Name`, `First name`, `Nome` | optional |
  | notes for you | `Notes`, `Note` | optional |

  *Excel*: the sheet `Scores` (or `Punteggi`, otherwise the first sheet). A row whose exam
  cell is `max` holds the maximum score of each question: every column with a numeric
  maximum is a question to grade (e.g. `Q1`, `Q2`, `7A`…). An optional `min` row allows
  negative scores (questions named `G1`, `G2` go down to −maximum by default; change
  `NEGATIVE_QUESTIONS` in the code). Scores are written into the cells.

  *CSV*: the questions are the exercises of the grading rubric; their columns and a total
  column (`Open questions total`, or `Totale domande aperte` for an Italian file) are added.

  A timestamped backup of the list is made at the first save of each session. Close the
  file in Excel while grading: the program refuses to save while it is open.
- **Solutions** (optional, button *Solutions…*): either a single PDF for everybody, or a zip
  with one file per exam named after the exam number (`00012.pdf` or `00012.tex`; other
  files of the zip, e.g. images, are available to LaTeX). The `.tex` files are compiled in
  the background with `pdflatex`.

### Features
- Draw on the pages (stored as vectors, independent of zoom) with the mouse or a pen tablet
  (Wacom…): the pen tip draws, a tap leaves a dot, the side button (right click) erases the
  strokes it touches; *Eraser* turns the pen tip into an eraser; *Undo* (Cmd/Ctrl+Z) undoes
  the last stroke, erase or page clear
- Scores are checked against the maximum (and minimum) of each question; the program
  resumes from the first exam without scores, and *Go to exam…* jumps to any exam
- Comments, with buttons for frequent phrases and a list of the phrases already used in
  other exams
- **Solutions window** (*Solutions ⧉*): a second window with the solutions of the exam
  being graded, following the main window; its position is remembered (e.g. on a second
  screen)
- **PDF for the student**, rebuilt at each save:
  1. *T/F page* (CSV only): a first page with a true/false part, i.e. the answers given,
     the correct answers, outcome and points of each question, and the score. The columns
     are guessed from their names (`Answer 1A`… or `Risposta 1A`…, `Points 1A`… or
     `Punti 1A`…, `Total`/`Totale`, `Correct answers`/`Risposte giuste`); *T/F columns…*
     shows them with the first data row as an example and lets you change them. Answers
     may be `T`/`F` or `V`/`F`; the correct answers are all in one column (e.g. `TFFT`).
     The maximum of each question is the best score obtained in the class
  2. the exam with your drawings, then the evaluation page with scores and comments
  3. *Append solutions*: the solutions of that exam
- **Language of the student pages**: Italian or English. *English* sets it for the current
  exam, *English by default* for all the exams not opened yet
- *Rebuild PDFs…* rebuilds the PDF of every graded exam with the current options (scores,
  comments, drawings and language stay as saved)
- **Grading rubric** (`rubric.txt` in the PDF folder, asked at startup if missing,
  editable with *Rubric…*): for each exercise, the points given for each part of the answer
- **Proposals** (`proposals.json` in the PDF folder, optional): proposed scores
  (highlighted in yellow) and comments, e.g. written by Claude after reading the scans.
  Check and edit them: nothing is written until you save

### Files written in the PDF folder
| File | Content |
|---|---|
| `originals/` | the original scans: an exam can be regraded without duplicating pages |
| `.grading_state.json` | drawings, comments, language and options of each exam |
| `rubric.txt` | the grading rubric |
| `solutions.zip` / `solutions.pdf` | a copy of the chosen solutions |
| `solutions_pdf/` | the solutions compiled from the zip |

Folders created by older versions, with the Italian names (`originali/`,
`.correzione_stato.json`, `griglia.txt`, `proposte.json`, `soluzioni.zip`,
`soluzioni_pdf/`), keep working.

### Grading rubric format
```
[Exercise 1]
1     correct characteristic equation
0.5   general solution
1     initial value problem solved
-0.25 each calculation error
Note for the grader: any line that does not start with a number.

[Exercise 2]
...
```
The name in brackets is the name of the score column (for an Excel file it must match a
column with a `max`). The maximum of an exercise is the sum of its positive points.

### Proposals format
```json
{
  "12": {
    "scores": {"Exercise 1": 2, "Exercise 2": 1.5},
    "comment_it": "Esercizio 1: corretto. Esercizio 2: manca la verifica della condizione iniziale.",
    "comment_en": "Exercise 1: correct. Exercise 2: the initial condition is not checked.",
    "doubts": "Page 3 hard to read: check the sign in the last step."
  }
}
```
Keys are exam numbers. `comment_it` and `comment_en` are used for Italian and English
pages; a plain `comment` is used for both. `doubts` is shown only to you, never written in
the PDF. Scores outside the allowed range are ignored, and proposals are not used for exams
already graded. The Italian keys of older versions (`punteggi`, `commento`, `dubbi`) are
also read.

**With Claude Code**: the scans can be read by Claude in a Claude Code session (no API key
needed). Give Claude the PDF folder with the rubric and the solutions and ask for the
proposals: Claude writes `proposals.json`, which the grader re-reads whenever it changes.

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
