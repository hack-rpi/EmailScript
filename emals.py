import subprocess
import sys
import pandas as pd
import win32com.client as win32
import pythoncom
import win32com.client.gencache
import time
import os

REQUIRED_COLUMNS = {"Name", "Company", "Email"}

def is_outlook_running():
    try:
        tasks = subprocess.check_output('tasklist', shell=True).decode()
        return "OUTLOOK.EXE" in tasks
    except Exception as e:
        print(f"Error checking running processes: {e}")
        return False

def open_outlook():
    if is_outlook_running():
        print("Outlook is already running.")
        print("Please close Outlook in Task Manager (end task for OUTLOOK.EXE) before continuing.")
        input("Press Enter after you have closed Outlook...")

        if is_outlook_running():
            print("Outlook is still running. Exiting script.")
            sys.exit()

    try:
        print("Starting Outlook...")
        pythoncom.CoInitialize()
        outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        print("Outlook started successfully.")
        return outlook
    except Exception as e:
        print(f"Failed to start Outlook: {e}")
        return None

def get_user_input():
    csv_file = input("Enter the name of the CSV file (e.g., Mails.csv): ").strip()
    html_file = input("Enter the name of the HTML template file (e.g., email_template.html): ").strip()
    send_direct = input("Send emails directly? (y/N): ").strip().lower() == 'y'

    if not os.path.exists(csv_file):
        print(f"CSV file not found: {csv_file}")
        sys.exit(1)

    if not os.path.exists(html_file):
        print(f"HTML template file not found: {html_file}")
        sys.exit(1)

    return csv_file, html_file, send_direct

def validate_csv(df):
    if not REQUIRED_COLUMNS.issubset(df.columns):
        missing = REQUIRED_COLUMNS - set(df.columns)
        print(f"CSV is missing required columns: {', '.join(missing)}")
        sys.exit(1)

def main():
    csv_file, html_file, send_mode = get_user_input()
    outlook = open_outlook()

    if not outlook:
        print("Script terminated due to Outlook startup failure.")
        return

    try:
        contacts = pd.read_csv(csv_file)
        validate_csv(contacts)

        with open(html_file, 'r', encoding='utf-8') as f:
            html_template = f.read()

        for _, row in contacts.iterrows():
            contact_name = row["Name"]
            company_name = row["Company"]
            email = row["Email"]

            mail = outlook.CreateItem(0)
            mail.To = email
            mail.SentOnBehalfOfName = "hackrpi@rpi.edu"
            mail.Subject = f"HackRPI 2025 Sponsorship Invitation for {company_name}"

            html_body = html_template.replace("{contact_name}", contact_name).replace("{company_name}", company_name)
            mail.HTMLBody = html_body

            if send_mode:
                mail.Send()
                print(f"Sent email to {contact_name} at {company_name} ({email})", flush=True)
            else:
                mail.Display()
                print(f"Drafted email to {contact_name} at {company_name} ({email})", flush=True)

            # Optional delay to prevent Outlook lag
            # time.sleep(1)

    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
