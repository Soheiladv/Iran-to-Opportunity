"""Job Email Analyzer — extracts structured job data from EMAIL_TRACKER.json
Produces memory/EMAIL_EVIDENCE.md + output/JOB_EMAIL_ANALYSIS.md
"""

import json, os, re, sys
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TRACKER = os.path.join(BASE, "memory", "EMAIL_TRACKER.json")
EVIDENCE = os.path.join(BASE, "memory", "EMAIL_EVIDENCE.md")
REPORT = os.path.join(BASE, "output", "JOB_EMAIL_ANALYSIS.md")

STATUSES = ["Offer", "Interview", "Applied", "Rejected", "Contacted"]
STATUS_PATTERNS = {
    "Offer": r"congratulation|offer of employment|we are pleased to offer|job offer",
    "Interview": r"interview|invite you to (a )?(call|meeting)|schedule a call|video call",
    "Rejected": r"unfortunately|we regret|not (be )?(able to )?move forward|position has been filled|decided not to",
    "Applied": r"application (received|confirmed)|thank you for (your )?(applying|application)|we (have )?received your application",
    "Contacted": r"recruiter|opportunit|would you be interested|are you (still )?(available|looking)",
}
COUNTRIES = [
    "Canada",
    "Australia",
    "New Zealand",
    "Germany",
    "Austria",
    "Netherlands",
    "UK|United Kingdom",
    "USA|United Sta=tes",
    "Norway",
    "Sweden",
    "Denmark",
    "Finland",
    "Switzerland",
    "Ireland",
    "Belgium",
    "Saskatchewan",
    "Alberta",
]
DEADLINE = r"(deadline|apply by|closing date|before)[:\s]*([A-Z][a-z]+ \d{1,2}(?:st|nd|rd|th)?,? \d{4}|\d{1,2}[/.]\d{1,2}[/.]\d{2,4})"
LINK = r"https?://[^\s\"'><]+"


def detect_status(text):
    text_l = text.lower()
    for st in STATUSES:
        if re.search(STATUS_PATTERNS[st], text_l):
            return st
    return "Other"


def detect_country(text):
    for c in COUNTRIES:
        if re.search(rf"\b({c})\b", text, re.I):
            return c.replace("|United Kingdom", "").replace("|United States", "")
    return ""


def extract_company(sender):
    m = re.search(
        r"@([\w-]+)\.(com|net|org|io|co|gov|ca|nz|de|at|nl|au|uk|eu|health)", sender
    )
    if m:
        dom = m.group(1)
        if dom not in ("gmail", "yahoo", "outlook", "hotmail", "linkedin"):
            return dom.title()
    return "Unknown"


FA = "--fa" in sys.argv

FA_SUBJECTS_TITLE = "📋 فهرست ایمیل‌های شغلی دریافت‌شده — {ts}\n"

FA_STATUS_MAP = {
    "Offer": "پیشنهاد همکاری",
    "Interview": "مصاحبه",
    "Applied": "درخواست ارسال‌شده",
    "Rejected": "رد‌شده",
    "Contacted": "تماس اولیه",
    "Other": "سایر",
}

FA_HEADER = """<div dir="rtl">

# 📊 تحلیل ایمیل‌های شغلی — {ts}

"""

FA_HEADERS = {
    "pipeline": "## 🔄 وضعیت پرونده‌ها",
    "countries": "## 🌍 کشورها",
    "followup": "## ⏰ نیازمند پیگیری ({n})",
}


def analyze():
    if not os.path.exists(TRACKER):
        print("[!] Run email_connector.py first.")
        return
    data = json.load(open(TRACKER, encoding="utf-8"))
    emails = data.get("emails", [])
    records = []
    for e in emails:
        text = e["subject"] + "\n" + e["body"]
        status = detect_status(text)
        country = detect_country(text)
        deadline = re.search(DEADLINE, text, re.I)
        links = re.findall(LINK, text)
        records.append(
            {
                "id": e["id"],
                "from": e["from"],
                "date": e["date"],
                "subject": e["subject"],
                "company": extract_company(e["from"]),
                "status": status,
                "country": country,
                "deadline": deadline.group(2) if deadline else "",
                "link": links[0][:120] if links else "",
            }
        )

    # Stats
    stats = {}
    for r in records:
        stats[r["status"]] = stats.get(r["status"], 0) + 1
    countries = {}
    for r in records:
        if r["country"]:
            countries[r["country"]] = countries.get(r["country"], 0) + 1
    followups = [
        r for r in records if r["status"] in ("Contacted", "Applied", "Interview")
    ]

    ts = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Evidence (Source of Truth)
    with open(EVIDENCE, "w", encoding="utf-8") as f:
        f.write(f"# Email Evidence Registry — {ts}\n\n")
        f.write(
            "| # | Date | From | Company | Country | Status | Deadline | Subject |\n|---|---|---|---|---|---|---|---|\n"
        )
        for i, r in enumerate(records, 1):
            f.write(
                f"| {i} | {r['date'][:10]} | {r['from'][:30]} | {r['company']} | {r['country']} | {r['status']} | {r['deadline']} | {r['subject'][:60]} |\n"
            )

    # Report
    with open(REPORT, "w", encoding="utf-8") as f:
        if FA:
            f.write(FA_HEADER.format(ts=ts))
            f.write(FA_HEADERS["pipeline"] + "\n\n")
            for st in STATUSES + ["Other"]:
                if st in stats:
                    f.write(f"- **{FA_STATUS_MAP[st]}** ({st}): {stats[st]} مورد\n")
            f.write("\n" + FA_HEADERS["countries"] + "\n\n")
            for c, n in sorted(countries.items(), key=lambda x: -x[1]):
                f.write(f"- **{c}:** {n} ایمیل\n")
            f.write("\n" + FA_HEADERS["followup"].format(n=len(followups)) + "\n\n")
            for r in followups:
                f.write(
                    f"- **[{FA_STATUS_MAP[r['status']]}]** {r['subject']} — {r['from']} ({r['date'][:10]})\n"
                )
            f.write("\n</div>\n")
        else:
            f.write(f"# Job Email Analysis — {ts}\n\n## 📊 Status Pipeline\n\n")
            for st in STATUSES + ["Other"]:
                if st in stats:
                    f.write(f"- **{st}:** {stats[st]}\n")
            f.write(f"\n## 🌍 Countries\n\n")
            for c, n in sorted(countries.items(), key=lambda x: -x[1]):
                f.write(f"- **{c}:** {n}\n")
            f.write(f"\n## ⏰ Follow-up Needed ({len(followups)})\n\n")
            for r in followups:
                f.write(
                    f"- [{r['status']}] {r['subject']} — {r['from']} ({r['date'][:10]})\n"
                )

    print(f"[+] Analyzed {len(records)} emails -> {REPORT}")
    print(f"[+] Evidence registry -> {EVIDENCE}")
    print(f"[+] Status: {stats}")


if __name__ == "__main__":
    analyze()
