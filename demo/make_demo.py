"""
Creates a small demo for the grader, with made-up students:

  demo/students.csv          student list with a true/false part (one answer column
                             per question, points, total and the correct answers)
  demo/exams/exam_1.pdf …    three "scanned" exams (typed text standing in for handwriting)
  demo/exams/rubric.txt      grading rubric of the two open exercises
  demo/exams/proposals.json  example grading proposals (scores and comments)
  demo/solutions.pdf         one solutions PDF for everybody
  demo/solutions.zip         one solutions PDF per exam (named after the exam number)

Run it, then start grading_open_ended_questions.py and choose demo/students.csv,
the demo/exams folder, and one of the two solutions files.
"""
import csv
import io
import json
import os
import zipfile

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

HERE = os.path.dirname(os.path.abspath(__file__))
EXAMS = os.path.join(HERE, "exams")

STUDENTS = [  # exam number, student ID, surname, name, true/false answers given
    (1, "100001", "Doe", "Jane", "TFTF"),
    (2, "100002", "Smith", "John", "TTTF"),
    (3, "100003", "Rossi", "Maria", "T FF"),
]
KEY = "TFFF"
QUESTIONS = ["1A", "1B", "2A", "2B"]

ANSWERS = {
    1: ["y' = 2y,  y(0) = 3", "Separating variables: dy/y = 2 dx, so y = C e^(2x).",
        "y(0) = 3 gives C = 3:  y = 3 e^(2x).", "",
        "Exercise 2", "Integral of x e^x dx: by parts, u = x, dv = e^x dx,",
        "= x e^x - e^x + C."],
    2: ["y' = 2y,  y(0) = 3", "y = C e^(2x)", "C = 3?", "",
        "Exercise 2", "Integral of x e^x dx = x e^x + C"],
    3: ["y' = 2y,  y(0) = 3", "Characteristic: r = 2,  y = C e^(2x),  y = 3 e^(2x)", "",
        "Exercise 2", "By parts: x e^x - integral of e^x dx = x e^x - e^x + C"],
}


def write_pdf(path, pages):
    c = canvas.Canvas(path, pagesize=A4)
    W, H = A4
    for title, lines in pages:
        c.setFont("Helvetica-Bold", 16)
        c.drawString(60, H - 70, title)
        c.setFont("Courier", 13)
        y = H - 110
        for line in lines:
            c.drawString(60, y, line)
            y -= 24
        c.showPage()
    c.save()


def pdf_bytes(pages):
    buf = io.BytesIO()
    write_pdf(buf, pages)
    return buf.getvalue()


def main():
    os.makedirs(EXAMS, exist_ok=True)

    # Student list with the true/false part (1 point per correct answer)
    header = ["Exam", "Student ID", "Surname", "Name"] \
        + [f"Answer {q}" for q in QUESTIONS] + [f"Points {q}" for q in QUESTIONS] \
        + ["Total", "Correct answers"]
    with open(os.path.join(HERE, "students.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        for n, sid, surname, name, given in STUDENTS:
            given = list(given.replace(" ", "_"))
            points = [1 if g == k else 0 for g, k in zip(given, KEY)]
            w.writerow([n, sid, surname, name]
                       + [g if g in "TF" else "" for g in given]
                       + points + [sum(points), KEY])

    # Exams
    for n, *_ in STUDENTS:
        write_pdf(os.path.join(EXAMS, f"exam_{n}.pdf"),
                  [(f"Exam {n} - Exercise 1", ANSWERS[n])])

    # Rubric
    with open(os.path.join(EXAMS, "rubric.txt"), "w", encoding="utf-8") as f:
        f.write("# Grading rubric\n\n"
                "[Exercise 1]\n"
                "1     general solution y = C e^(2x)\n"
                "1     initial condition: C = 3\n"
                "-0.5  each calculation error\n\n"
                "[Exercise 2]\n"
                "1     integration by parts set up correctly\n"
                "1     correct result x e^x - e^x + C\n"
                "Check that the constant of integration is there.\n")

    # Example proposals: scores and comments to be checked by the grader
    proposals = {
        "1": {"scores": {"Exercise 1": 2, "Exercise 2": 2},
              "comment": "Esercizio 1: corretto. Esercizio 2: corretto.",
              "comment_en": "Exercise 1: correct. Exercise 2: correct."},
        "2": {"scores": {"Exercise 1": 1, "Exercise 2": 0},
              "comment": "Esercizio 1: manca la verifica di C = 3. "
                         "Esercizio 2: il risultato non è corretto.",
              "comment_en": "Exercise 1: C = 3 is not justified. "
                            "Exercise 2: the result is wrong.",
              "doubts": "Exercise 1: the question mark after C = 3 may mean the student was unsure."},
    }
    with open(os.path.join(EXAMS, "proposals.json"), "w", encoding="utf-8") as f:
        json.dump(proposals, f, indent=2, ensure_ascii=False)

    # Solutions: one PDF for everybody, and a zip with one PDF per exam
    solution = [("Solutions", ["Exercise 1:  y = 3 e^(2x)",
                               "Exercise 2:  x e^x - e^x + C"])]
    write_pdf(os.path.join(HERE, "solutions.pdf"), solution)
    with zipfile.ZipFile(os.path.join(HERE, "solutions.zip"), "w") as z:
        for n, *_ in STUDENTS:
            z.writestr(f"{n:05d}.pdf", pdf_bytes([(f"Solutions - exam {n}", solution[0][1])]))

    print(f"Demo created in {HERE}")


if __name__ == "__main__":
    main()
