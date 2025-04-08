import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkhtmlview import HTMLLabel
import pandas as pd
import os
import time
import subprocess
import win32com.client as win32
import pythoncom
import win32com.client.gencache
from PIL import Image, ImageTk

REQUIRED_COLUMNS = {"Name", "Company", "Email"}

class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Outlook Email Sender")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)

        self.csv_path = tk.StringVar()
        self.html_path = tk.StringVar()
        self.use_html_file = tk.BooleanVar(value=True)
        self.use_html_format = tk.BooleanVar(value=True)
        self.attachment_paths = []
        self.image_paths = []
        self.subject_template = tk.StringVar(value="HackRPI 2025 Sponsorship Invitation for {company_name}")

        self.df = pd.DataFrame()

        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

        self.build_tabs()

    def build_tabs(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True)

        self.build_email_tab(notebook)
        self.build_csv_viewer_tab(notebook)

    def build_email_tab(self, notebook):
        email_tab = tk.Frame(notebook)
        notebook.add(email_tab, text="Email Composer")

        for i in range(4):
            email_tab.columnconfigure(i, weight=1)
        email_tab.rowconfigure(7, weight=1)

        tk.Label(email_tab, text="CSV File:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        tk.Entry(email_tab, textvariable=self.csv_path).grid(row=0, column=1, columnspan=2, sticky="ew", padx=5, pady=5)
        tk.Button(email_tab, text="Browse", command=self.browse_csv).grid(row=0, column=3, padx=5, pady=5)

        tk.Checkbutton(email_tab, text="Send as HTML", variable=self.use_html_format).grid(row=1, column=1, sticky="w", padx=5, pady=5)

        tk.Label(email_tab, text="Email Template Source:").grid(row=2, column=0, sticky="e")
        tk.Radiobutton(email_tab, text="Import File", variable=self.use_html_file, value=True, command=self.toggle_html_source).grid(row=2, column=1, sticky="w")
        tk.Radiobutton(email_tab, text="Write in Editor", variable=self.use_html_file, value=False, command=self.toggle_html_source).grid(row=2, column=2, sticky="w")

        self.html_file_entry = tk.Entry(email_tab, textvariable=self.html_path)
        self.html_file_entry.grid(row=3, column=1, columnspan=2, sticky="ew", padx=5)
        self.browse_html_button = tk.Button(email_tab, text="Browse", command=self.browse_html)
        self.browse_html_button.grid(row=3, column=3, padx=5)

        self.html_editor = tk.Text(email_tab, wrap="word")
        html_editor_scroll = tk.Scrollbar(email_tab, command=self.html_editor.yview)
        self.html_editor.config(yscrollcommand=html_editor_scroll.set)

        self.html_editor.grid_forget()
        html_editor_scroll.grid_forget()

        tk.Label(email_tab, text="Email Subject Template:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        tk.Entry(email_tab, textvariable=self.subject_template).grid(row=4, column=1, columnspan=3, sticky="ew", padx=5, pady=5)

        tk.Button(email_tab, text="Add PDF Attachment(s)", command=self.add_attachments).grid(row=5, column=1, sticky="ew", padx=5, pady=5)
        tk.Button(email_tab, text="Add Inline Image(s)", command=self.add_images).grid(row=5, column=2, sticky="ew", padx=5, pady=5)

        self.image_preview_frame = tk.Frame(email_tab)
        self.image_preview_frame.grid(row=6, column=0, columnspan=4, sticky="ew", padx=5, pady=5)

        tk.Button(email_tab, text="Send Emails", command=self.send_emails).grid(row=7, column=2, sticky="e", padx=5, pady=10)

    def build_csv_viewer_tab(self, notebook):
        csv_tab = tk.Frame(notebook)
        notebook.add(csv_tab, text="CSV Viewer")

        self.csv_table = ttk.Treeview(csv_tab)
        self.csv_table.pack(fill="both", expand=True, side="left")

        scrollbar_y = ttk.Scrollbar(csv_tab, orient="vertical", command=self.csv_table.yview)
        scrollbar_y.pack(side="right", fill="y")
        self.csv_table.configure(yscrollcommand=scrollbar_y.set)

    def toggle_html_source(self):
        if self.use_html_file.get():
            self.html_file_entry.grid()
            self.browse_html_button.grid()
            self.html_editor.grid_forget()
        else:
            self.html_file_entry.grid_forget()
            self.browse_html_button.grid_forget()
            self.html_editor.grid(row=3, column=1, columnspan=3, rowspan=2, sticky="nsew", padx=5, pady=5)

    def browse_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if path:
            self.csv_path.set(path)
            self.load_csv(path)

    def browse_html(self):
        path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
        if path:
            self.html_path.set(path)

    def load_csv(self, path):
        try:
            self.df = pd.read_csv(path)
            self.populate_csv_table()
        except Exception as e:
            messagebox.showerror("Error", f"Could not load CSV: {e}")

    def populate_csv_table(self):
        self.csv_table.delete(*self.csv_table.get_children())

        if self.df.empty:
            return

        self.csv_table["columns"] = list(self.df.columns)
        self.csv_table["show"] = "headings"

        for col in self.df.columns:
            self.csv_table.heading(col, text=col)
            self.csv_table.column(col, anchor="w", width=150)

        for _, row in self.df.head(50).iterrows():
            self.csv_table.insert("", "end", values=list(row))

    def add_attachments(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if paths:
            self.attachment_paths.extend(paths)

    def add_images(self):
        paths = filedialog.askopenfilenames(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.gif")])
        if paths:
            self.image_paths.extend(paths)
            self.show_image_previews()

    def show_image_previews(self):
        for widget in self.image_preview_frame.winfo_children():
            widget.destroy()
        for img_path in self.image_paths:
            try:
                img = Image.open(img_path)
                img.thumbnail((100, 100))
                tk_img = ImageTk.PhotoImage(img)
                label = tk.Label(self.image_preview_frame, image=tk_img)
                label.image = tk_img
                label.pack(side="left", padx=5)
            except:
                pass

    def send_emails(self):
        if not os.path.exists(self.csv_path.get()):
            messagebox.showerror("Error", "CSV file is missing.")
            return

        try:
            df = pd.read_csv(self.csv_path.get())
            if not REQUIRED_COLUMNS.issubset(df.columns):
                missing = REQUIRED_COLUMNS - set(df.columns)
                messagebox.showerror("Missing Columns", f"CSV is missing: {', '.join(missing)}")
                return

            if self.use_html_file.get():
                if not os.path.exists(self.html_path.get()):
                    messagebox.showerror("Error", "HTML template file is missing.")
                    return
                with open(self.html_path.get(), "r", encoding="utf-8") as f:
                    html_template = f.read()
            else:
                html_template = self.html_editor.get("1.0", tk.END)

            pythoncom.CoInitialize()
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")

            for _, row in df.iterrows():
                mail = outlook.CreateItem(0)
                mail.To = row["Email"]
                subject = self.subject_template.get().replace("{company_name}", row["Company"]).replace("{contact_name}", row["Name"])
                mail.Subject = subject
                body = html_template.replace("{contact_name}", row["Name"]).replace("{company_name}", row["Company"])

                if self.use_html_format.get():
                    mail.HTMLBody = body
                else:
                    mail.Body = body

                for file in self.attachment_paths:
                    mail.Attachments.Add(file)

                for i, img_path in enumerate(self.image_paths):
                    attachment = mail.Attachments.Add(img_path)
                    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", f"img{i}")
                    mail.HTMLBody = mail.HTMLBody.replace(f"cid:image{i}", f"cid:img{i}")

                mail.Display()
                time.sleep(1)

            messagebox.showinfo("Success", "Emails drafted successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to send emails: {e}")

        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def on_exit(self):
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()