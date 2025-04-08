import tkinter as tk
from tkinter import filedialog, messagebox
from tkhtmlview import HTMLLabel
import pandas as pd
import os
import time
import sys
import subprocess
import win32com.client as win32
import pythoncom
import win32com.client.gencache

REQUIRED_COLUMNS = {"Name", "Company", "Email"}

class EmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook Email Sender")

        self.csv_path = tk.StringVar()
        self.html_path = tk.StringVar()

        # File selection
        tk.Label(root, text="CSV File:").grid(row=0, column=0, sticky="e")
        tk.Entry(root, textvariable=self.csv_path, width=40).grid(row=0, column=1)
        tk.Button(root, text="Browse", command=self.browse_csv).grid(row=0, column=2)

        tk.Label(root, text="HTML Template File:").grid(row=1, column=0, sticky="e")
        tk.Entry(root, textvariable=self.html_path, width=40).grid(row=1, column=1)
        tk.Button(root, text="Browse", command=self.browse_html).grid(row=1, column=2)

        # HTML preview area
        tk.Label(root, text="First Email Preview:").grid(row=2, column=0, sticky="ne")
        self.html_preview = HTMLLabel(root, html="", width=80, height=20, background="white", relief="sunken", bd=1)
        self.html_preview.grid(row=2, column=1, columnspan=2, sticky="nsew")

        # Buttons
        tk.Button(root, text="Preview First Email", command=self.load_preview).grid(row=3, column=1, sticky="e")
        tk.Button(root, text="Draft Emails", command=lambda: self.run_email_logic(send_mode=False)).grid(row=4, column=1, sticky="e")
        tk.Button(root, text="Send Emails", command=lambda: self.run_email_logic(send_mode=True)).grid(row=4, column=2, sticky="w")
        tk.Button(root, text="Exit", command=root.quit).grid(row=5, column=2, sticky="e")

    def browse_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            self.csv_path.set(path)

    def browse_html(self):
        path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
        if path:
            self.html_path.set(path)

    def load_preview(self):
        csv = self.csv_path.get()
        html = self.html_path.get()

        if not os.path.exists(csv) or not os.path.exists(html):
            messagebox.showerror("Error", "Both CSV and HTML template must be valid files.")
            return

        try:
            df = pd.read_csv(csv)
            if not REQUIRED_COLUMNS.issubset(df.columns):
                missing = REQUIRED_COLUMNS - set(df.columns)
                messagebox.showerror("Missing Columns", f"CSV is missing: {', '.join(missing)}")
                return

            first = df.iloc[0]
            with open(html, 'r', encoding='utf-8') as f:
                template = f.read()
            preview_html = template.replace("{contact_name}", first["Name"]).replace("{company_name}", first["Company"])
            self.html_preview.set_html(preview_html)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run_email_logic(self, send_mode):
        csv = self.csv_path.get()
        html = self.html_path.get()

        if not os.path.exists(csv) or not os.path.exists(html):
            messagebox.showerror("Error", "Both CSV and HTML template must be selected before sending.")
            return

        if send_mode:
            confirm = messagebox.askyesno("Confirm Send", "Are you sure you want to send the emails directly?")
            if not confirm:
                return

        try:
            pythoncom.CoInitialize()

            if is_outlook_running():
                messagebox.showwarning("Outlook Running", "Please close Outlook completely before continuing.")
                return

            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            df = pd.read_csv(csv)

            if not REQUIRED_COLUMNS.issubset(df.columns):
                missing = REQUIRED_COLUMNS - set(df.columns)
                messagebox.showerror("Missing Columns", f"CSV is missing: {', '.join(missing)}")
                return

            with open(html, 'r', encoding='utf-8') as f:
                template = f.read()

            if not send_mode:
                confirm = messagebox.askyesno("Confirm Drafting", "Proceed to draft the emails?")
                if not confirm:
                    return

            for i, (_, row) in enumerate(df.iterrows()):
                name, company, email = row["Name"], row["Company"], row["Email"]
                mail = outlook.CreateItem(0)
                mail.To = email
                mail.SentOnBehalfOfName = "hackrpi@rpi.edu"
                mail.Subject = f"HackRPI 2025 Sponsorship Invitation for {company}"

                mail.HTMLBody = template.replace("{contact_name}", name).replace("{company_name}", company)

                if send_mode:
                    mail.Send()
                else:
                    mail.Display()
                    if i == 0:
                        if not messagebox.askyesno("First Draft OK?", "Is the first draft correct?"):
                            messagebox.showinfo("Canceled", "Drafting canceled.")
                            return

                time.sleep(1)

            messagebox.showinfo("Success", "Emails processed successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            pythoncom.CoUninitialize()

def is_outlook_running():
    try:
        tasks = subprocess.check_output('tasklist', shell=True).decode()
        return "OUTLOOK.EXE" in tasks
    except Exception as e:
        print(f"Error checking running processes: {e}")
        return False

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderGUI(root)
    root.mainloop()
