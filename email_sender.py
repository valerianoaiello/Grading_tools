"""
Email Sender for Students (CSV-based)

This script sends personalized emails to students listed in a CSV file.
The user can:
- Select a CSV file
- Specify which columns contain first name, last name, and email
- Write a common email body using placeholders: [name], [surname], [email]
- Select a Gmail credentials JSON file
- Optionally attach one or more files to every email

Requirements:
- Python 3.x
- google-auth-oauthlib
- google-api-python-client
"""

import csv
import os
import base64
import tkinter as tk
from tkinter import filedialog, messagebox
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


SCOPES = ['https://www.googleapis.com/auth/gmail.send']


class StudentEmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Student Email Sender (CSV)")

        self.csv_file = ""
        self.credentials_file = ""
        self.attachments = []

        self.build_ui()

    # ---------------- UI ----------------

    def build_ui(self):
        frame = tk.Frame(self.root)
        frame.pack(padx=10, pady=10)

        tk.Button(frame, text="Select students CSV",
                  command=self.select_csv).grid(row=0, column=0, sticky="w")

        tk.Button(frame, text="Select Gmail credentials",
                  command=self.select_credentials).grid(row=0, column=1, sticky="w")

        tk.Button(frame, text="Add attachment(s)",
                  command=self.add_attachments).grid(row=0, column=2, sticky="w")

        tk.Label(frame, text="Name column index:").grid(row=1, column=0, sticky="w")
        self.entry_name = tk.Entry(frame, width=5)
        self.entry_name.grid(row=1, column=1)

        tk.Label(frame, text="Surname column index:").grid(row=2, column=0, sticky="w")
        self.entry_surname = tk.Entry(frame, width=5)
        self.entry_surname.grid(row=2, column=1)

        tk.Label(frame, text="Email column index:").grid(row=3, column=0, sticky="w")
        self.entry_email = tk.Entry(frame, width=5)
        self.entry_email.grid(row=3, column=1)

        tk.Label(frame, text="Email subject:").grid(row=4, column=0, sticky="w")
        self.entry_subject = tk.Entry(frame, width=60)
        self.entry_subject.grid(row=4, column=1, columnspan=2)

        tk.Label(
            frame,
            text="Email body (use [name], [surname], [email]):"
        ).grid(row=5, column=0, columnspan=3, sticky="w")

        self.text_body = tk.Text(frame, width=80, height=10)
        self.text_body.grid(row=6, column=0, columnspan=3, pady=5)

        tk.Button(frame, text="Send emails",
                  command=self.send_all).grid(row=7, column=0, columnspan=3, pady=10)

        tk.Label(frame, text="Log:").grid(row=8, column=0, sticky="w")
        self.log_box = tk.Text(frame, width=80, height=10)
        self.log_box.grid(row=9, column=0, columnspan=3)

    # ---------------- Utility ----------------

    def log(self, message):
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        print(message)

    def select_csv(self):
        self.csv_file = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")]
        )
        if self.csv_file:
            self.log(f"CSV selected: {self.csv_file}")

    def select_credentials(self):
        self.credentials_file = filedialog.askopenfilename(
            title="Select Gmail credentials JSON",
            filetypes=[("JSON files", "*.json")]
        )
        if self.credentials_file:
            self.log(f"Credentials selected: {self.credentials_file}")

    def add_attachments(self):
        files = filedialog.askopenfilenames(
            title="Select attachment files"
        )
        if files:
            self.attachments = list(files)
            self.log(f"Attachments added: {', '.join(os.path.basename(f) for f in files)}")

    # ---------------- Gmail ----------------

    def get_gmail_service(self):
        if not self.credentials_file:
            raise RuntimeError("No credentials file selected")

        flow = InstalledAppFlow.from_client_secrets_file(
            self.credentials_file, SCOPES
        )
        creds = flow.run_local_server(port=0)
        return build('gmail', 'v1', credentials=creds)

    def send_email(self, service, to_email, subject, body, attachments):
        message = MIMEMultipart()
        message['to'] = to_email
        message['subject'] = subject

        message.attach(MIMEText(body, 'plain', 'utf-8'))

        for filepath in attachments:
            with open(filepath, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{os.path.basename(filepath)}"'
                )
                message.attach(part)

        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        service.users().messages().send(
            userId='me',
            body={'raw': raw}
        ).execute()

    # ---------------- Main logic ----------------

    def send_all(self):
        if not self.csv_file:
            messagebox.showwarning("Error", "Please select a CSV file")
            return

        if not self.credentials_file:
            messagebox.showwarning("Error", "Please select a credentials file")
            return

        try:
            col_name = int(self.entry_name.get())
            col_surname = int(self.entry_surname.get())
            col_email = int(self.entry_email.get())
        except ValueError:
            messagebox.showerror("Error", "Column indices must be integers")
            return

        subject = self.entry_subject.get().strip()
        base_text = self.text_body.get("1.0", tk.END).strip()

        if not subject or not base_text:
            messagebox.showwarning("Error", "Subject and email body are required")
            return

        try:
            service = self.get_gmail_service()

            with open(self.csv_file, newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                headers = next(reader)
                self.log(f"CSV header: {headers}")

                for row in reader:
                    name = row[col_name].strip()
                    surname = row[col_surname].strip()
                    email = row[col_email].strip()

                    body = (
                        base_text
                        .replace("[name]", name)
                        .replace("[surname]", surname)
                        .replace("[email]", email)
                    )

                    self.send_email(
                        service,
                        email,
                        subject,
                        body,
                        self.attachments
                    )

                    self.log(f"Email sent to {name} {surname} ({email})")

            messagebox.showinfo("Done", "All emails have been sent")

        except Exception as e:
            messagebox.showerror("Error", str(e))


# ---------------- Entry point ----------------

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentEmailApp(root)
    root.mainloop()
