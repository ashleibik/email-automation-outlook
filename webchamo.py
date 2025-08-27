import win32com.client as win32
import os
import csv
from datetime import datetime

# ---------------------------
# SETTINGS
# ---------------------------
SAVE_FOLDER = r"C:\Users\shlei\Downloads\Attachments"
LOG_FILE = os.path.join(SAVE_FOLDER, "attachments_log.csv")

# ---------------------------
# OUTLOOK CONNECTION
# ---------------------------
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

# ---------------------------
# ENSURE SAVE FOLDER EXISTS
# ---------------------------
os.makedirs(SAVE_FOLDER, exist_ok=True)

# ---------------------------
# PREPARE LOG FILE
# ---------------------------
if not os.path.exists(LOG_FILE):
    with open(LOG_FILE, mode="w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["Date", "Sender", "Subject", "AttachmentName", "SavedPath"])

# ---------------------------
# LOOP THROUGH EMAILS
# ---------------------------
for email in inbox.Items:
    try:
        sender = email.SenderEmailAddress or "UnknownSender"
        subject = email.Subject or "(No Subject)"
        date_received = email.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S")

        # Create a folder for each sender
        sender_folder = os.path.join(SAVE_FOLDER, sender.replace("@", "_at_"))
        os.makedirs(sender_folder, exist_ok=True)

        # Save attachments
        for attachment in email.Attachments:
            filename = attachment.FileName
            save_path = os.path.join(sender_folder, filename)

            # Save file
            attachment.SaveAsFile(save_path)

            # Log info
            with open(LOG_FILE, mode="a", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow([date_received, sender, subject, filename, save_path])

            print(f"✔ Saved: {filename} from {sender}")

    except Exception as e:
        print(f"⚠ Error processing email: {e}")
