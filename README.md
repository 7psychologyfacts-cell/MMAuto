# 📨 MMAuto — Mail Merge & Bulk Document Automation Platform

**A full-stack web application that turns a spreadsheet into hundreds of personalized documents and emails — automatically, in minutes, with zero manual copy-paste.**

Built end-to-end (backend + frontend + deployment) to solve a real, everyday business problem: generating and dispatching personalized letters, case summaries, and reports at scale — the kind of workflow used daily in insurance claims, legal case management, HR communication, finance recovery teams, and corporate operations.

---

## 🎯 Why This Project Matters

Every organization has *someone* stuck doing this manually:
> "Open Excel → copy a row → paste into a Word template → save as a new file → repeat 200 times → then email each one individually."

MMAuto replaces that entire manual loop with a **self-serve web tool**: upload your spreadsheet, upload your Word template, map the fields once, and the system does the rest — document generation, formatting, and email dispatch — for one record or for thousands, individually or grouped by company/client.

This project demonstrates the ability to **identify a real operational bottleneck and engineer a complete, production-shaped solution** for it — not just a script, but a tool other people can actually use.

---

## 🚀 Key Features

### 📄 Smart Document Generation
- Upload any Word (`.docx`) template with `{{tag}}` placeholders — the app auto-detects every tag, including inside headers/footers.
- Map spreadsheet columns to template tags with a visual, point-and-click interface.
- Generate one merged document per row, bundled into a single downloadable `.zip`.
- Export as **Word (.docx)** or **PDF** — PDF rendering is done server-side without relying on MS Word/LibreOffice, rebuilding formatting (bold/italic/underline, headings, bordered tables, landscape layout) from scratch.

### 🏢 Company-Wise Grouping (Batch Intelligence)
- Automatically group hundreds of spreadsheet rows by any column (e.g., Insurance Company, Client, Department).
- For each group, generate **and email in one shot**:
  - A personalized merged **Letter**
  - An auto-formatted **Summary Table Document** (aligned, bordered, auto-landscape for wide tables)
  - A **Filtered Excel** containing only that group's records, with correct number/date formatting
  - A **.zip of individual case documents** for that group
  - *(Newest feature)* The summary table can also be embedded **directly inside the email body**, so a full case summary is visible without opening any attachment.

### 📧 Real Email Dispatch — Not Just a Mail-Merge Preview
- Direct **SMTP** integration to send real, personalized emails (individual or grouped).
- Optional **CC support**, custom sender name, and a rich-text (WYSIWYG) email body editor.
- Optional **IMAP "Save to Sent"** — because SMTP alone never saves a copy to your own mailbox.
- Live **Email Shot Status dashboard**: total / sent / failed counters, progress bar, and a per-recipient send log exportable as CSV.
- Built-in "Send Test Email" before committing to a full batch send.

### 🔐 Privacy-First, Stateless Architecture
- No database. No files persisted on the server.
- Spreadsheet data, templates, and credentials live only in the browser session and the current request — cleared the moment the tab closes.
- Designed to run on **serverless infrastructure (Vercel)**, so it scales to zero cost when idle.

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| **Backend** | Python, Flask (REST API) |
| **Data Processing** | Pandas, OpenPyXL |
| **Document Engineering** | python-docx (Word generation/parsing), ReportLab (PDF rendering engine) |
| **Email** | smtplib (SMTP), imaplib (IMAP APPEND), MIME multipart (HTML + plain-text) |
| **Frontend** | Vanilla JavaScript, HTML5, CSS3 — no framework bloat, fast and dependency-light |
| **Deployment** | Vercel (serverless Python runtime) |

---

## 💡 Skills This Project Demonstrates

- **Backend API design** — clean REST endpoints for upload, template parsing, generation, and dispatch
- **Data manipulation & ETL** — cleaning, typing (number/date detection), and reshaping spreadsheet data with Pandas
- **Document automation / templating engines** — parsing and dynamically rewriting `.docx` XML structure
- **Email systems engineering** — SMTP/IMAP protocols, MIME construction, HTML+plaintext multipart emails
- **Full-stack thinking** — a real user-facing UI wired to a real backend, not just a notebook script
- **Product sense** — features driven by an actual business workflow (grouped attachments, test-send, status dashboard, error logging)
- **Deployment & DevOps basics** — serverless deployment configuration, stateless/session-safe architecture

---

## 📂 Project Structure

```
MMAuto/
├── app.py            # Flask backend — all API routes & document/email logic
├── index.html         # Full frontend UI (upload, mapping, dispatch, dashboard)
├── requirements.txt    # Python dependencies
└── vercel.json         # Serverless deployment configuration
```

---

## ⚙️ Getting Started

```bash
# 1. Clone the repository
git clone <your-repo-url>
cd MMAuto

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run locally
python app.py
```

Then open `http://localhost:5000` in your browser.

---

## 📌 Roadmap Ideas

- OAuth-based email sending (Gmail/Outlook API) instead of raw SMTP credentials
- Scheduling / delayed batch dispatch
- Template library with reusable saved mappings
- Multi-language template support

---

## 🙋 About Me

I built this project to solve a real repetitive-work problem end-to-end — from spreadsheet parsing to document generation to live email dispatch — and to demonstrate that I can take an idea from "this is annoying and manual" to a working, deployable tool.

📧 **Email:** milanvadher2003@gmail.com
💻 **GitHub:** github.com/7psychologyfacts-cell

---

⭐ *If you found this project interesting, a star on the repo is appreciated!*
