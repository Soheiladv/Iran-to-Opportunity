"""Email Connector — Gmail IMAP reader for MigrationHunter
Reads job-related emails, prompts for password at runtime (getpass, hidden).
Usage:  python scripts/email_connector.py [max_emails]
"""

import imaplib
import email
import email.header
import getpass
import json
import os
import sys
import re
from datetime import datetime

FA = "--fa" in sys.argv

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE, "output", "emails")
MEM_DIR = os.path.join(BASE, "memory")
TRACKER = os.path.join(MEM_DIR, "EMAIL_TRACKER.json")

GMAIL_HOST = "imap.gmail.com"
GMAIL_PORT = 993


def decode(s):
    if not s:
        return ""
    parts = email.header.decode_header(s)
    out = ""
    for text, enc in parts:
        if isinstance(text, bytes):
            out += text.decode(enc or "utf-8", errors="replace")
        else:
            out += text
    return out


def get_body(msg):
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition", ""))
            if ctype == "text/plain" and "attachment" not in disp:
                payload = part.get_payload(decode=True)
                if payload:
                    body += payload.decode(
                        part.get_content_charset() or "utf-8", errors="replace"
                    )
                    break
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            body = payload.decode(
                msg.get_content_charset() or "utf-8", errors="replace"
            )
    return body

JOB_KEYWORDS = re.compile(
    r"(job|vacanc|position|recruit|career|application|interview|offer|"
    r"hiring|opportunit|cv|resume|hiring|employer|sponsor)", re.I)

def fetch_emails(imap, mailbox="INBOX", max_emails=100, only_unseen=False):
    imap.select(mailbox, readonly=True)
    criteria = "(UNSEEN)" if only_unseen else "(ALL)"
    _, data = imap.search(None, criteria)
    ids = data[0].split()
    total = len(ids)
    print(f"[i] Mailbox '{mailbox}': {total} emails found, fetching last {min(max_emails, total)}")
    emails = []
    for eid in ids[-max_emails:]:
        _, msg_data = imap.fetch(eid, "(RFC822)")
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        date_tuple = email.utils.parsedate_to_datetime(msg.get("Date", ""))
        emails.append({
            "id": eid.decode(),
            "from": decode(msg.get("From", "")),
            "from_email": (email.utils.parseaddr(msg.get("From", ""))[1] or "").lower(),
            "subject": decode(msg.get("Subject", "")),
            "date": date_tuple.strftime("%Y-%m-%d %H:%M") if date_tuple else "",
            "body": get_body(msg)[:5000],
        })
    return emails

def is_job_related(e):
    text = e["subject"] + " " + e["from"] + " " + e["body"][:1000]
    return bool(JOB_KEYWORDS.search(text))

def save_results(all_emails, job_emails):
    os.makedirs(OUT_DIR, exist_ok=True)
    os.makedirs(MEM_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M")

    # Raw dump of job emails
    dump = os.path.join(OUT_DIR, f"gmail_job_emails_{ts}.md")
    with open(dump, "w", encoding="utf-8") as f:
        if FA:
            f.write(f'<div dir="rtl">\n\n# 📩 ایمیل‌های شغلی دریافت‌شده از Gmail — {ts}\n\n')
            f.write(f"**تعداد کل بررسی‌شده:** {len(all_emails)} | **مرتبط با شغل:** {len(job_emails)}\n\n")
            for e in job_emails:
                f.write(f"## 📧 {e['subject']}\n- **فرستنده:** {e['from']}\n- **تاریخ:** {e['date']}\n\n```\n{e['body'][:2000]}\n```\n\n")
            f.write("</div>\n")
        else:
            f.write(f"# Gmail Job Emails — {ts}\n\n")
            for e in job_emails:
                f.write(f"## {e['subject']}\n- **From:** {e['from']}\n- **Date:** {e['date']}\n\n```\n{e['body'][:2000]}\n```\n\n")

    # Update tracker JSON
    tracker = {"last_sync": ts, "total_scanned": len(all_emails), "job_related": len(job_emails), "emails": job_emails}
    if os.path.exists(TRACKER):
        try:
            old = json.load(open(TRACKER, encoding="utf-8"))
            seen = {x["id"] for x in old.get("emails", [])}
            tracker["emails"] = old.get("emails", []) + [e for e in job_emails if e["id"] not in seen]
        except Exception:
            pass
    json.dump(tracker, open(TRACKER, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"[+] Saved {len(job_emails)} job emails -> {dump}")
    print(f"[+] Tracker updated -> {TRACKER}")

def main():
    addr_input = input("Email address [t.arjmand1980@gmail.com]: ").strip()
    addr = addr_input or "t.arjmand1980@gmail.com"
    pwd = getpass.getpass("App Password (hidden input): ")

    max_emails = int(sys.argv[1]) if len(sys.argv) > 1 else 100
    only_unseen = "--unseen" in sys.argv

    print(f"[*] Connecting to {GMAIL_HOST}...")
    imap = imaplib.IMAP4_SSL(GMAIL_HOST, GMAIL_PORT)
    try:
        imap.login(addr, pwd)
    except imaplib.IMAP4.error as ex:
        print(f"[!] LOGIN FAILED: {ex}")
        print("    Make sure you use a 16-char App Password (not your Gmail password):")
        print("    https://myaccount.google.com/apppasswords")
        sys.exit(1)
    print("[+] Logged in successfully")

    all_emails = fetch_emails(imap, "INBOX", max_emails, only_unseen)
    job_emails = [e for e in all_emails if is_job_related(e)]
    print(f"[+] Job-related: {len(job_emails)} / {len(all_emails)}")
    save_results(all_emails, job_emails)
    imap.logout()

if __name__ == "__main__":
    main()
