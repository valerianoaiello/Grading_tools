
This repository contains the first tool in a series of small Python utilities designed to assist with academic and administrative tasks.

## Grader for open ended questions

A Python desktop application for grading PDF exams with annotations, scoring, and feedback.

### Features
- Draw annotations directly on PDF pages
- Score multiple questions per exam
- Add comments that are appended as evaluation pages
- Export scores to CSV (European format: `;` separator, `,` decimal)
- Save progress automatically
- Navigate with arrow controls


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
