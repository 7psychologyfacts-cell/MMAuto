"""
Mail Merge Automation - Backend
--------------------------------
Flask backend for:
  1. Parsing uploaded Excel/CSV data
  2. Parsing an uploaded Word (.docx) template and extracting {{tag}} placeholders
  3. Generating merged .docx documents (bulk, zipped)
  4. Sending merged documents as email attachments over SMTP
  5. Optionally saving a copy of every sent email into the mailbox's "Sent"
     folder via IMAP APPEND (SMTP alone does NOT do this automatically)
  6. Returning a per-recipient send log that the frontend can export as CSV

NOTE ON STATE: this app is designed to run on Vercel's serverless Python
runtime, which is stateless between requests (and may run on a different
instance/container each time). Nothing is stored on disk or in memory
across requests. The browser holds the parsed Excel data, the template
(base64) and the field mapping, and re-sends whatever a request needs.
Credentials (SMTP/IMAP) are supplied by the user in the browser each
session and are only ever held in server memory for the lifetime of a
single request - never written to disk or logged.
"""

import base64
import csv
import io
import os
import re
import smtplib
import ssl
import time
import zipfile
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, make_msgid

import pandas as pd
from docx import Document
from flask import Flask, jsonify, request, send_file, send_from_directory

app = Flask(__name__, static_folder=None)

# Max upload size: 15 MB (Excel + docx templates are small; keeps abuse in check)
app.config["MAX_CONTENT_LENGTH"] = 15 * 1024 * 1024

TAG_PATTERN = re.compile(r"\{\{\s*([A-Za-z0-9_.\-]+)\s*\}\}")


# --------------------------------------------------------------------------
# Static frontend
# --------------------------------------------------------------------------
@app.route("/")
def index():
    return send_from_directory(os.path.dirname(os.path.abspath(__file__)), "index.html")


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def df_to_records(df: pd.DataFrame):
    """Convert a dataframe to plain-JSON-safe list[dict], NaN -> ''."""
    df = df.fillna("")
    # Force everything to string-safe python types (avoids numpy int64 etc.)
    records = []
    for _, row in df.iterrows():
        records.append({str(col): ("" if pd.isna(v) else str(v)) for col, v in row.items()})
    return records


def extract_tags_from_docx(doc: Document):
    """Find every {{tag}} used anywhere in the document body, tables, and headers/footers."""
    tags = set()

    def scan(text):
        for m in TAG_PATTERN.finditer(text):
            tags.add(m.group(1))

    for p in doc.paragraphs:
        scan(p.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    scan(p.text)
    for section in doc.sections:
        for part in (section.header, section.footer):
            for p in part.paragraphs:
                scan(p.text)

    return sorted(tags)


def replace_tags_in_paragraph(paragraph, mapping):
    """Replace {{tag}} occurrences in a paragraph while keeping the paragraph's
    formatting. Word frequently splits a single {{tag}} across multiple runs,
    so we rebuild the paragraph text as a whole and re-apply it to the first
    run (simple, reliable approach for mail-merge style templates)."""
    full_text = "".join(r.text for r in paragraph.runs)
    if "{{" not in full_text:
        return

    def sub(m):
        key = m.group(1)
        return str(mapping.get(key, m.group(0)))

    new_text = TAG_PATTERN.sub(sub, full_text)
    if new_text == full_text:
        return

    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for r in paragraph.runs[1:]:
            r.text = ""
    else:
        paragraph.text = new_text


def fill_template(template_bytes: bytes, mapping: dict) -> bytes:
    """Return a new .docx (bytes) with every {{tag}} in the template replaced
    using `mapping` = {tag_name: value}."""
    doc = Document(io.BytesIO(template_bytes))

    for p in doc.paragraphs:
        replace_tags_in_paragraph(p, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_tags_in_paragraph(p, mapping)
    for section in doc.sections:
        for part in (section.header, section.footer):
            for p in part.paragraphs:
                replace_tags_in_paragraph(p, mapping)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def merge_text(template_str: str, row: dict) -> str:
    """Substitute {{ColumnName}} directly against an excel row (used for
    email subject/body, which use the raw excel column names as tags)."""
    if not template_str:
        return ""

    def sub(m):
        key = m.group(1)
        return str(row.get(key, m.group(0)))

    return TAG_PATTERN.sub(sub, template_str)


def safe_filename(name: str) -> str:
    name = re.sub(r"[^\w\-.]+", "_", str(name)).strip("_")
    return name or "document"


def open_smtp(conf):
    host = conf["host"]
    port = int(conf["port"])
    security = conf.get("security", "starttls")  # 'ssl' | 'starttls' | 'none'

    if security == "ssl":
        server = smtplib.SMTP_SSL(host, port, timeout=20, context=ssl.create_default_context())
    else:
        server = smtplib.SMTP(host, port, timeout=20)
        server.ehlo()
        if security == "starttls":
            server.starttls(context=ssl.create_default_context())
            server.ehlo()

    if conf.get("username"):
        server.login(conf["username"], conf.get("password", ""))
    return server


def open_imap(conf):
    import imaplib

    host = conf["host"]
    port = int(conf.get("port") or 993)
    if conf.get("security", "ssl") == "ssl":
        m = imaplib.IMAP4_SSL(host, port, timeout=20)
    else:
        m = imaplib.IMAP4(host, port, timeout=20)
        if conf.get("security") == "starttls":
            m.starttls(ssl_context=ssl.create_default_context())
    m.login(conf["username"], conf.get("password", ""))
    return m


def append_to_sent(imap_conf, raw_message: bytes):
    import imaplib

    m = open_imap(imap_conf)
    try:
        folder = imap_conf.get("folder") or "Sent"
        typ, _ = m.select(folder, readonly=False)
        if typ != "OK":
            # try common alternates if the given folder name doesn't exist
            for alt in ("Sent", "INBOX.Sent", "Sent Items", "INBOX.Sent Items"):
                typ, _ = m.select(alt, readonly=False)
                if typ == "OK":
                    break
        m.append(folder, "\\Seen", imaplib.Time2Internaldate(time.time()), raw_message)
    finally:
        try:
            m.logout()
        except Exception:
            pass


def build_email(from_name, from_email, to_email, subject, html_body, attachment=None, attachment_name=None):
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = formataddr((from_name or from_email, from_email))
    msg["To"] = to_email
    msg["Message-ID"] = make_msgid()

    alt = MIMEMultipart("alternative")
    plain = re.sub("<[^<]+?>", "", html_body or "")
    alt.attach(MIMEText(plain, "plain", "utf-8"))
    alt.attach(MIMEText(html_body or "", "html", "utf-8"))
    msg.attach(alt)

    if attachment is not None:
        part = MIMEApplication(attachment, Name=attachment_name)
        part["Content-Disposition"] = f'attachment; filename="{attachment_name}"'
        msg.attach(part)

    return msg


# --------------------------------------------------------------------------
# API: upload spreadsheet
# --------------------------------------------------------------------------
@app.route("/api/upload-data", methods=["POST"])
def upload_data():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400

    filename = f.filename or ""
    try:
        if filename.lower().endswith(".csv"):
            df = pd.read_csv(f)
        else:
            df = pd.read_excel(f)
    except Exception as e:
        return jsonify({"error": f"Could not read spreadsheet: {e}"}), 400

    if df.empty:
        return jsonify({"error": "Spreadsheet has no rows"}), 400

    df.columns = [str(c).strip() for c in df.columns]
    records = df_to_records(df)

    return jsonify({
        "filename": filename,
        "columns": list(df.columns),
        "rows": records,
        "row_count": len(records),
    })


# --------------------------------------------------------------------------
# API: upload word template
# --------------------------------------------------------------------------
@app.route("/api/upload-template", methods=["POST"])
def upload_template():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    if not (f.filename or "").lower().endswith(".docx"):
        return jsonify({"error": "Please upload a .docx file (old .doc format is not supported)"}), 400

    file_bytes = f.read()
    try:
        doc = Document(io.BytesIO(file_bytes))
        tags = extract_tags_from_docx(doc)
    except Exception as e:
        return jsonify({"error": f"Could not read Word template: {e}"}), 400

    if not tags:
        return jsonify({"error": "No {{tags}} found in this template. Add placeholders like {{Full_Name}} and re-upload."}), 400

    return jsonify({
        "filename": f.filename,
        "tags": tags,
        "template_b64": base64.b64encode(file_bytes).decode("ascii"),
    })


# --------------------------------------------------------------------------
# API: generate merged documents (zip)
# --------------------------------------------------------------------------
@app.route("/api/generate", methods=["POST"])
def generate_documents():
    data = request.get_json(force=True, silent=True) or {}
    template_b64 = data.get("template_b64")
    mapping = data.get("mapping") or {}       # {tag: excel_column}
    rows = data.get("rows") or []
    name_field = data.get("filename_field")   # excel column to use for output filenames

    if not template_b64 or not rows:
        return jsonify({"error": "Missing template or data rows"}), 400

    try:
        template_bytes = base64.b64decode(template_b64)
    except Exception:
        return jsonify({"error": "Invalid template data"}), 400

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        used_names = set()
        for i, row in enumerate(rows, start=1):
            row_mapping = {tag: row.get(col, "") for tag, col in mapping.items()}
            doc_bytes = fill_template(template_bytes, row_mapping)

            base_name = safe_filename(row.get(name_field, "")) if name_field else f"document_{i}"
            fname = f"{base_name}.docx"
            n = 1
            while fname in used_names:
                n += 1
                fname = f"{base_name}_{n}.docx"
            used_names.add(fname)

            zf.writestr(fname, doc_bytes)

    zip_buffer.seek(0)
    return send_file(
        zip_buffer,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"merged_documents_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
    )


# --------------------------------------------------------------------------
# API: test SMTP / IMAP credentials
# --------------------------------------------------------------------------
@app.route("/api/test-connection", methods=["POST"])
def test_connection():
    data = request.get_json(force=True, silent=True) or {}
    kind = data.get("type")  # 'smtp' | 'imap'
    conf = data.get("config") or {}

    try:
        if kind == "smtp":
            server = open_smtp(conf)
            server.quit()
        elif kind == "imap":
            m = open_imap(conf)
            m.logout()
        else:
            return jsonify({"ok": False, "error": "Unknown connection type"}), 400
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 200

    return jsonify({"ok": True})


# --------------------------------------------------------------------------
# API: send emails (called by the frontend in small batches)
# --------------------------------------------------------------------------
@app.route("/api/send-emails", methods=["POST"])
def send_emails():
    data = request.get_json(force=True, silent=True) or {}

    smtp_conf = data.get("smtp") or {}
    from_name = data.get("from_name") or ""
    from_email = data.get("from_email") or smtp_conf.get("username", "")
    subject_template = data.get("subject") or ""
    body_template = data.get("body") or ""
    recipient_field = data.get("recipient_field")
    rows = data.get("rows") or []

    save_copy = bool(data.get("save_sent_copy"))
    imap_conf = data.get("imap") or {}

    attach_doc = bool(data.get("attach_document"))
    template_b64 = data.get("template_b64")
    doc_mapping = data.get("doc_mapping") or {}
    filename_field = data.get("filename_field")

    if not recipient_field:
        return jsonify({"error": "recipient_field (which column holds the email address) is required"}), 400
    if not rows:
        return jsonify({"error": "No rows to send"}), 400

    template_bytes = None
    if attach_doc and template_b64:
        try:
            template_bytes = base64.b64decode(template_b64)
        except Exception:
            return jsonify({"error": "Invalid template data"}), 400

    log = []
    try:
        smtp_server = open_smtp(smtp_conf)
    except Exception as e:
        return jsonify({"error": f"Could not connect to SMTP server: {e}"}), 400

    for row in rows:
        to_email = str(row.get(recipient_field, "")).strip()
        entry = {
            "recipient": to_email,
            "subject": "",
            "status": "failed",
            "error": "",
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        }

        if not to_email or "@" not in to_email:
            entry["error"] = "Missing/invalid email address"
            log.append(entry)
            continue

        try:
            subject = merge_text(subject_template, row)
            body = merge_text(body_template, row)
            entry["subject"] = subject

            attachment_bytes = None
            attachment_name = None
            if attach_doc and template_bytes is not None:
                row_mapping = {tag: row.get(col, "") for tag, col in doc_mapping.items()}
                attachment_bytes = fill_template(template_bytes, row_mapping)
                base_name = safe_filename(row.get(filename_field, "")) if filename_field else safe_filename(to_email)
                attachment_name = f"{base_name}.docx"

            msg = build_email(from_name, from_email, to_email, subject, body, attachment_bytes, attachment_name)
            raw = msg.as_bytes()

            smtp_server.sendmail(from_email, [to_email], raw)
            entry["status"] = "sent"

            if save_copy:
                try:
                    append_to_sent(imap_conf, raw)
                    entry["sent_copy_saved"] = True
                except Exception as e:
                    entry["sent_copy_saved"] = False
                    entry["sent_copy_error"] = str(e)

        except Exception as e:
            entry["status"] = "failed"
            entry["error"] = str(e)

        log.append(entry)

    try:
        smtp_server.quit()
    except Exception:
        pass

    sent = sum(1 for e in log if e["status"] == "sent")
    failed = len(log) - sent
    return jsonify({"log": log, "sent": sent, "failed": failed})


# --------------------------------------------------------------------------
# API: log download as CSV (kept server-side too, in case the frontend
# wants to re-download a log it already has instead of rebuilding client-side)
# --------------------------------------------------------------------------
@app.route("/api/download-log", methods=["POST"])
def download_log():
    data = request.get_json(force=True, silent=True) or {}
    log = data.get("log") or []

    buf = io.StringIO()
    fieldnames = ["timestamp", "recipient", "subject", "status", "sent_copy_saved", "error"]
    writer = csv.DictWriter(buf, fieldnames=fieldnames, extrasaction="ignore")
    writer.writeheader()
    for row in log:
        writer.writerow(row)

    mem = io.BytesIO(buf.getvalue().encode("utf-8-sig"))
    return send_file(
        mem,
        mimetype="text/csv",
        as_attachment=True,
        download_name=f"send_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
