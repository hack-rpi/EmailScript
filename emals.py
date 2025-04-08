import subprocess
import sys
import pandas as pd
import win32com.client as win32
import pythoncom
import win32com.client.gencache

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


outlook = open_outlook()

if outlook:
    contacts = pd.read_csv("Mails.csv")

    for _, row in contacts.iterrows():
        contact_name = row["Name"]
        company_name = row["Company"]
        email = row["Email"]

        mail = outlook.CreateItem(0)
        mail.To = email
        mail.SentOnBehalfOfName = "hackrpi@rpi.edu"
        mail.Subject = f"HackRPI 2025 Sponsorship Invitation for {company_name}"

        mail.HTMLBody = f"""
        <html>
        <body style="font-family:Segoe UI, sans-serif; font-size:14px; color:#333333; line-height:1.6;">
            <p>Hello {contact_name},</p>

            <p>
            Our team is thrilled to announce <strong>HackRPI 2025</strong> – the upcoming annual HackRPI event taking place in the fall. 
            We will be hosting this event on <strong>November 15–16</strong> at 
            <a href="https://www.rpi.edu" style="color:#0066cc; text-decoration:none;">Rensselaer Polytechnic Institute</a> in Troy, New York.
            </p>

            <p>
            HackRPI is a student-run organization that hosts annual hackathons at Rensselaer Polytechnic Institute. 
            We firmly believe hackathons are a leading source of innovation and ingenuity. 
            As such, we invite students with diverse programming backgrounds from across the globe to form teams, 
            generate novel ideas, and design original prototypes. 
            Last November’s hackathon had upwards of <strong>500 attendees</strong>, making us the largest hackathon in New York’s Capital District.
            </p>

            <p>
            Our team would like to invite <strong>{company_name}</strong> to sponsor HackRPI 2025 this year. 
            As a sponsor, <strong>{company_name}</strong> will receive various perks as detailed in our sponsorship booklet.
            </p>

            <p>
            We also want to thank our past sponsors for all the invaluable feedback they have provided to make these events possible. 
            With a decade of experience now behind us, we fully intend on turning HackRPI 2025 into our greatest hackathon yet.
            </p>

            <p>
            Please feel free to reach back out to me with any questions regarding a potential sponsorship, 
            and let me know if you would like to receive a copy of our sponsorship booklet.
            </p>

            <p style="margin-top: 30px;">
            Regards,<br>
            Aaryan Guatam<br>
            Director of Sponsorship<br>
            <em>HackRPI Organizing Team</em>
            </p>
        </body>
        </html>
        """

        mail.Display()
        print(f"Drafted email to {contact_name} at {company_name} ({email})")
else:
    print("Script terminated due to Outlook startup failure.")
