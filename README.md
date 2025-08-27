# 📧 Outlook Email Automation Tool

A Python script that automatically downloads and organizes attachments from Microsoft Outlook emails.  
Attachments are stored in sender-specific folders, and all activity is logged into a CSV file.

## ✨ Features
- Connects to Microsoft Outlook using `pywin32`
- Saves attachments into `Attachments/<SenderEmail>/`
- Logs each file into `attachments_log.csv` with:
  - Date received
  - Sender
  - Subject
  - Filename
  - Saved path

## 🚀 Setup
1. Clone this repo:
   ```bash
   git clone https://github.com/ashleibik/email-automation-outlook.git
   cd email-automation-outlook
   ```
2. Install dependencies:
   ```bash
   py -m pip install -r requirements.txt
   ```
3. Run the script:
   ```bash
   py webchamo.py
   ```

## 👤 Example Output
```
Attachments/
├─ sender1_at_example.com/
│  └─ invoice.pdf
├─ sender2_at_example.com/
│  └─ report.docx
```

## 📛 License
MIT
