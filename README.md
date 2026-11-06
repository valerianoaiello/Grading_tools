
This repository contains the first tool in a series of small Python utilities designed to assist with academic and administrative tasks.

## Grader for open ended questions

A Python desktop application for grading PDF exams with annotations, scoring, and feedback.

### Features
- Reads students and questions from an Excel file (sheet `Punteggi`):
  row 1 holds the headers (`Compito`, `Matricola`, `Cognome`, `Nome`, …), the row whose
  `Compito` cell is `massimo` holds the maximum score of each question. Every column with a
  numeric maximum is a question to grade (e.g. `G1`, `G2`, `7A`…`7D`).
- Matches each PDF to a student by exam number: the last group of digits in the file name
  (`compito_12.pdf`, `261103_12.pdf` and `12.pdf` all map to exam 12)
- Writes each score into the right Excel cell and checks it does not exceed the maximum;
  a timestamped backup of the Excel file is made at the first save of each session
- Draw annotations on PDF pages (stored as vectors, independent of zoom)
- Comments are appended to the PDF as an evaluation page with name, scores and total
- Originals are kept in an `originali/` subfolder, so an exam can be regraded without
  duplicating the evaluation page; drawings and comments are kept in `.correzione_stato.json`
- Resumes from the first exam without scores; "Vai al compito…" jumps to any exam number

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
- Select a Gmail **credentials JSON file** (OAuth 2.0)
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
