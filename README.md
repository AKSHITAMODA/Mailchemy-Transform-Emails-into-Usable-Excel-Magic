# Mailchemy:Transform Emails into Usable Excel Magic


**email-xlsx-parser** is a lightweight Python tool that connects to your Gmail inbox, downloads `.xlsx` attachments, scans them for keyword matches, and stores results in a SQLite database. A built-in Tkinter GUI lets you browse matches and export results to Excel files per category.

---

## ✨ Features

- 🔐 Connects securely to Gmail via IMAP
- 📎 Downloads only `.xlsx` email attachments
- 🧠 Scans Excel sheets for user-defined keywords
- 🗃 Saves results to a local SQLite database
- 📊 GUI viewer to browse and filter results
- 📤 One-click Excel export (per category or all)

---

## 📦 Requirements

- Python 3.8+
- [openpyxl](https://pypi.org/project/openpyxl/)
- Tkinter (included in standard Python installations)

Install required package:

```bash
pip install -r requirements.txt


email-xlsx-parser/
├── gui.py                  # GUI and interaction logic
├── mail_parser.py          # Email fetching and Excel scanning
├── db.py                   # SQLite setup and interaction
├── requirements.txt
├── README.md
├── matches.db              # Created after run
└── xlsx_attachments_only/  # Folder for Excel attachments

