#!/usr/bin/env python3
"""
MigrationHunter — Email Analyzer via IMAP
تحلیل خودکار ایمیل‌های شغلی و تطبیق با فرصت‌های موجود

پشتیبانی: Gmail, Outlook, Yahoo, هر سرویس IMAP

نحوه استفاده:
  python email_analyzer.py --email you@gmail.com --password APP_PASSWORD
  
  یا با فایل .env:
  python email_analyzer.py

Gmail App Password:
  1. Google Account → Security → 2-Step Verification → ON
  2. Google Account → Security → App passwords → Create
  3. کد 16 رقمی را وارد کن

Outlook/Office 365:
  1. Account Settings → Security → App passwords → Create
  2. یا IMAP را فعال کن

Yahoo:
  1. Account Security → App passwords → Generate
"""
import os, sys, re, json, email, imaplib, email.header
from datetime import datetime, timedelta
from collections import defaultdict

BASE = os.path.dirname(os.path.abspath(__file__))
MEM = os.path.join(BASE, "memory")
OUT = os.path.join(BASE, "output")

# ═══════════════════════════════════════════════════
# EMAIL PROVIDER CONFIGS
# ═══════════════════════════════════════════════════
PROVIDERS = {
    "gmail": {
        "imap": "imap.gmail.com",
        "port": 993,
        "smtp": "smtp.gmail.com",
        "note": "Use App Password, not regular password",
    },
    "outlook": {
        "imap": "outlook.office365.com",
        "port": 993,
        "smtp": "smtp.office365.com",
        "note": "Enable IMAP in Outlook settings",
    },
    "yahoo": {
        "imap": "imap.mail.yahoo.com",
        "port": 993,
        "smtp": "smtp.mail.yahoo.com",
        "note": "Use App Password",
    },
    "icloud": {
        "imap": "imap.mail.me.com",
        "port": 993,
        "smtp": "smtp.mail.me.com",
        "note": "Use App Password",
    },
}

# ═══════════════════════════════════════════════════
# JOB EMAIL KEYWORDS
# ═══════════════════════════════════════════════════
JOB_KEYWORDS = [
    # Positive
    "job offer", "position", "vacancy", "opportunity", "hiring",
    "recruitment", "apply", "application", "resume", "cv",
    "interview", "screening", "onsite", "offer letter",
    "salary", "compensation", "relocation", "visa sponsorship",
    "congratulations", "welcome", "onboarding",
    "shortlist", "selected", "proceed", "next step",
    "registered midwife", "midwife", "it manager", "infrastructure",
    "health new zealand", "saskatchewan", "alberta", "hays",
    "rgh global", "kate cowhig", "cpl healthcare",
    "make it in germany", "talentorange", "holalemania",
    "work in austria", "ind netherlands", "kennismigrant",
    "blue card", "aewv", "green list",
    # Negative
    "unfortunately", "regret", "not selected", "not progressing",
    "no longer", "position filled", "closed",
    # Follow-up
    "follow up", "following up", "checking in", "status update",
    "thank you for applying", "received your application",
]

REJECTION_KEYWORDS = [
    "unfortunately", "regret to inform", "not selected",
    "not progressing", "not moving forward", "position filled",
    "no longer accepting", "decided not to", "other candidates",
    "thank you for your interest", "wish you the best",
]

INTERVIEW_KEYWORDS = [
    "interview", "screening call", "phone call", "video call",
    "technical interview", "assessment", "test", "trial",
    "meet the team", "site visit", "onsite",
]

OFFER_KEYWORDS = [
    "job offer", "offer letter", "congratulations",
    "pleased to offer", "welcome to", "start date",
    "salary", "compensation package", "relocation",
]

FOLLOWUP_KEYWORDS = [
    "following up", "follow up", "checking in",
    "status update", "any update", "timeline",
]

# ═══════════════════════════════════════════════════
# EMPLOYER MATCHING
# ═══════════════════════════════════════════════════
KNOWN_EMPLOYERS = {
    "health new zealand": {"name": "Health New Zealand", "country": "NZ", "type": "Government"},
    "healthnz": {"name": "Health New Zealand", "country": "NZ", "type": "Government"},
    "te whatu ora": {"name": "Health New Zealand", "country": "NZ", "type": "Government"},
    "rgh global": {"name": "RGH Global", "country": "NZ", "type": "Recruiter"},
    "saskatchewan": {"name": "Saskatchewan HA", "country": "CA", "type": "Government"},
    "alberta health": {"name": "Alberta Health Services", "country": "CA", "type": "Government"},
    "kate cowhig": {"name": "Kate Cowhig", "country": "IE", "type": "Recruiter"},
    "cpl healthcare": {"name": "CPL Healthcare", "country": "IE", "type": "Recruiter"},
    "holalemania": {"name": "Holalemania", "country": "DE", "type": "Recruiter"},
    "talentorange": {"name": "TalentOrange", "country": "DE", "type": "Recruiter"},
    "make it in germany": {"name": "Make it in Germany", "country": "DE", "type": "Government"},
    "work in austria": {"name": "Work in Austria", "country": "AT", "type": "Government"},
    "ind.nl": {"name": "IND Netherlands", "country": "NL", "type": "Government"},
    "hays": {"name": "Hays Healthcare", "country": "AU", "type": "Recruiter"},
    "ahpra": {"name": "AHPRA", "country": "AU", "type": "Regulator"},
    "finncare": {"name": "Finncare", "country": "FI", "type": "Recruiter"},
    "workindenmark": {"name": "WorkInDenmark", "country": "DK", "type": "Government"},
    "vardforbundet": {"name": "Vårdförbundet", "country": "SE", "type": "Association"},
    "medicarrera": {"name": "MediCarrera", "country": "NL", "type": "Recruiter"},
}

# ═══════════════════════════════════════════════════
# IMAP CONNECTOR
# ═══════════════════════════════════════════════════
class EmailConnector:
    def __init__(self, email_addr, password, provider="gmail"):
        self.email = email_addr
        self.password = password
        config = PROVIDERS.get(provider, PROVIDERS["gmail"])
        self.imap_server = config["imap"]
        self.imap_port = config["port"]
        self.conn = None
    
    def connect(self):
        print(f"🔌 اتصال به {self.imap_server}...")
        self.conn = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
        self.conn.login(self.email, self.password)
        print(f"✅ اتصال موفق — {self.email}")
        return True
    
    def get_folders(self):
        """لیست پوشه‌ها"""
        status, folders = self.conn.list()
        result = []
        for f in folders:
            if isinstance(f, bytes):
                name = f.decode().split('" "')[-1].strip('"')
                result.append(name)
        return result
    
    def search_emails(self, folder="INBOX", days=30, subject_filter=None, 
                      sender_filter=None, limit=100):
        """جستجوی ایمیل‌ها"""
        self.conn.select(folder, readonly=True)
        
        # Build search criteria
        criteria = []
        
        # Date filter
        since_date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%y")
        criteria.append(f'(SINCE "{since_date}")')
        
        # Subject filter
        if subject_filter:
            for kw in subject_filter:
                criteria.append(f'(OR (SUBJECT "{kw}") ')
        
        # Combine
        search_str = " ".join(criteria) if criteria else "ALL"
        
        # For Gmail, use simple search
        try:
            status, messages = self.conn.search(None, "ALL")
        except:
            status, messages = self.conn.search(None, "ALL")
        
        if status != "OK":
            return []
        
        email_ids = messages[0].split()
        
        # Limit
        if limit:
            email_ids = email_ids[-limit:]
        
        results = []
        for eid in email_ids:
            try:
                status, msg_data = self.conn.fetch(eid, "(RFC822)")
                if status != "OK":
                    continue
                
                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)
                
                parsed = self._parse_email(msg, eid.decode())
                if parsed:
                    results.append(parsed)
            except Exception as e:
                continue
        
        return results
    
    def _parse_email(self, msg, eid):
        """Parse a single email"""
        try:
            # Subject
            subject = ""
            if msg["Subject"]:
                decoded = email.header.decode_header(msg["Subject"])
                for part, enc in decoded:
                    if isinstance(part, bytes):
                        subject += part.decode(enc or "utf-8", errors="ignore")
                    else:
                        subject += part
            
            # From
            from_addr = msg.get("From", "")
            
            # Date
            date_str = msg.get("Date", "")
            try:
                date_tuple = email.utils.parsedate_to_datetime(date_str)
                date_formatted = date_tuple.strftime("%Y-%m-%d %H:%M")
            except:
                date_formatted = date_str
            
            # Body
            body = self._get_body(msg)
            
            # Links
            links = re.findall(r'https?://[^\s<>"\']+', body + " " + subject)
            
            return {
                "id": eid,
                "subject": subject,
                "from": from_addr,
                "date": date_formatted,
                "body_preview": body[:2000] if body else "",
                "links": links,
                "has_attachments": self._has_attachments(msg),
            }
        except Exception as e:
            return None
    
    def _get_body(self, msg):
        """Extract email body"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        body += payload.decode(charset, errors="ignore")
                elif part.get_content_type() == "text/html" and not body:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        html = payload.decode(charset, errors="ignore")
                        # Simple HTML to text
                        body += re.sub(r'<[^>]+>', ' ', html)
        else:
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                body = payload.decode(charset, errors="ignore")
        return body.strip()
    
    def _has_attachments(self, msg):
        """Check for attachments"""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_maintype() != "multipart":
                    if part.get("Content-Disposition") and "attachment" in part.get("Content-Disposition"):
                        return True
        return False
    
    def disconnect(self):
        if self.conn:
            try:
                self.conn.close()
            except:
                pass
            self.conn.logout()

# ═══════════════════════════════════════════════════
# EMAIL ANALYZER
# ═══════════════════════════════════════════════════
class EmailAnalyzer:
    def __init__(self):
        self.results = []
    
    def classify(self, email_data):
        """دسته‌بندی یک ایمیل"""
        subject = email_data.get("subject", "").lower()
        body = email_data.get("body_preview", "").lower()
        from_addr = email_data.get("from", "").lower()
        full_text = subject + " " + body + " " + from_addr
        
        classification = {
            "category": "unknown",
            "is_job_related": False,
            "employer_match": None,
            "applicant": None,
            "urgency": "normal",
            "keywords_found": [],
            "sentiment": "neutral",
        }
        
        # Check if job-related
        matched_keywords = []
        for kw in JOB_KEYWORDS:
            if kw.lower() in full_text:
                matched_keywords.append(kw)
        
        classification["keywords_found"] = matched_keywords
        classification["is_job_related"] = len(matched_keywords) > 0
        
        if not classification["is_job_related"]:
            return classification
        
        # Classify category
        for kw in OFFER_KEYWORDS:
            if kw in full_text:
                classification["category"] = "offer"
                classification["sentiment"] = "positive"
                classification["urgency"] = "high"
                return classification
        
        for kw in REJECTION_KEYWORDS:
            if kw in full_text:
                classification["category"] = "rejection"
                classification["sentiment"] = "negative"
                return classification
        
        for kw in INTERVIEW_KEYWORDS:
            if kw in full_text:
                classification["category"] = "interview"
                classification["sentiment"] = "positive"
                classification["urgency"] = "high"
                return classification
        
        for kw in FOLLOWUP_KEYWORDS:
            if kw in full_text:
                classification["category"] = "follow_up"
                classification["urgency"] = "medium"
                return classification
        
        # Check for application acknowledgment
        if any(w in full_text for w in ["received your", "thank you for applying", "application received"]):
            classification["category"] = "acknowledgment"
            classification["sentiment"] = "neutral"
            return classification
        
        # Default: inquiry/response
        classification["category"] = "inquiry"
        return classification
    
    def match_employer(self, email_data):
        """تطبیق با کارفرمای شناخته شده"""
        from_addr = email_data.get("from", "").lower()
        subject = email_data.get("subject", "").lower()
        body = email_data.get("body_preview", "").lower()
        full_text = from_addr + " " + subject + " " + body
        
        for keyword, info in KNOWN_EMPLOYERS.items():
            if keyword in full_text:
                return info
        
        return None
    
    def detect_applicant(self, email_data):
        """تشخیص اینکه ایمیل برای کدام متقاضی است"""
        full_text = (email_data.get("subject", "") + " " + 
                    email_data.get("body_preview", "")).lower()
        
        neda_keywords = ["midwife", "midwifery", "neda", "مامایی"]
        tohid_keywords = ["it manager", "infrastructure", "systems", "tohid", "devops"]
        
        neda_score = sum(1 for k in neda_keywords if k in full_text)
        tohid_score = sum(1 for k in tohid_keywords if k in full_text)
        
        if neda_score > tohid_score:
            return "NEDA"
        elif tohid_score > neda_score:
            return "TOHID"
        return "UNKNOWN"
    
    def extract_dates(self, email_data):
        """استخراج تاریخ‌های مهم از ایمیل"""
        body = email_data.get("body_preview", "")
        dates = []
        
        # Common date patterns
        patterns = [
            r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
            r'(\w+ \d{1,2},? \d{4})',
            r'(\d{1,2} \w+ \d{4})',
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, body)
            dates.extend(matches)
        
        # Look for "by [date]" or "before [date]"
        deadline_patterns = [
            r'(?:by|before|until|deadline[:\s]+)(\w+ \d{1,2},? \d{4})',
            r'(?:by|before|until|deadline[:\s]+)(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
        ]
        
        deadlines = []
        for pattern in deadline_patterns:
            matches = re.findall(pattern, body, re.IGNORECASE)
            deadlines.extend(matches)
        
        return {"dates": dates, "deadlines": deadlines}
    
    def extract_links(self, email_data):
        """استخراج لینک‌های مهم"""
        links = email_data.get("links", [])
        important = []
        
        for link in links:
            if any(w in link.lower() for w in [
                "apply", "career", "job", "application", "form",
                "schedule", "interview", "calendar", "meeting",
                "portal", "register", "signup",
            ]):
                important.append({"url": link, "type": "actionable"})
            else:
                important.append({"url": link, "type": "reference"})
        
        return important
    
    def analyze_all(self, emails):
        """تحلیل تمام ایمیل‌ها"""
        results = []
        
        for em in emails:
            classification = self.classify(em)
            employer = self.match_employer(em)
            applicant = self.detect_applicant(em)
            dates = self.extract_dates(em)
            links = self.extract_links(em)
            
            result = {
                **em,
                **classification,
                "employer_match": employer,
                "applicant": applicant,
                "dates": dates,
                "important_links": links,
            }
            results.append(result)
        
        self.results = results
        return results

# ═══════════════════════════════════════════════════
# REPORT GENERATOR
# ═══════════════════════════════════════════════════
def generate_report(results, email_addr):
    """تولید گزارش تحلیل ایمیل"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    # Statistics
    total = len(results)
    job_related = sum(1 for r in results if r.get("is_job_related"))
    offers = sum(1 for r in results if r.get("category") == "offer")
    rejections = sum(1 for r in results if r.get("category") == "rejection")
    interviews = sum(1 for r in results if r.get("category") == "interview")
    follow_ups = sum(1 for r in results if r.get("category") == "follow_up")
    pending = sum(1 for r in results if r.get("category") in ["inquiry", "acknowledgment"])
    
    # By applicant
    neda_count = sum(1 for r in results if r.get("applicant") == "NEDA")
    tohid_count = sum(1 for r in results if r.get("applicant") == "TOHID")
    
    # By employer
    by_employer = defaultdict(int)
    for r in results:
        emp = r.get("employer_match")
        if emp:
            by_employer[emp["name"]] += 1
    
    # Build markdown
    lines = []
    lines.append(f"# گزارش تحلیل ایمیل شغلی")
    lines.append(f"")
    lines.append(f"**تاریخ:** {now}")
    lines.append(f"**ایمیل:** {email_addr}")
    lines.append(f"**تعداد ایمیل تحلیل شده:** {total}")
    lines.append(f"")
    lines.append(f"---")
    lines.append(f"")
    lines.append(f"## خلاصه آماری")
    lines.append(f"")
    lines.append(f"| وضعیت | تعداد |")
    lines.append(f"|-------|-------|")
    lines.append(f"| کل ایمیل‌ها | {total} |")
    lines.append(f"| مرتبط با کار | {job_related} |")
    lines.append(f"| 🎉 Job Offer | {offers} |")
    lines.append(f"| 🗣️ Interview | {interviews} |")
    lines.append(f"| ❌ Rejection | {rejections} |")
    lines.append(f"| ⏰ Follow-up | {follow_ups} |")
    lines.append(f"| 📩 Pending | {pending} |")
    lines.append(f"")
    lines.append(f"### بر اساس متقاضی")
    lines.append(f"")
    lines.append(f"| متقاضی | تعداد |")
    lines.append(f"|--------|-------|")
    lines.append(f"| 👩 ندا | {neda_count} |")
    lines.append(f"| 👨 توحید | {tohid_count} |")
    lines.append(f"")
    
    if by_employer:
        lines.append(f"### بر اساس کارفرما")
        lines.append(f"")
        lines.append(f"| کارفرما | تعداد |")
        lines.append(f"|---------|-------|")
        for emp, count in sorted(by_employer.items(), key=lambda x: -x[1]):
            lines.append(f"| {emp} | {count} |")
        lines.append(f"")
    
    # Action items
    lines.append(f"---")
    lines.append(f"")
    lines.append(f"## اقدامات لازم")
    lines.append(f"")
    
    # Interviews needing response
    interview_emails = [r for r in results if r.get("category") == "interview"]
    if interview_emails:
        lines.append(f"### 🗣️ مصاحبه — نیاز به پاسخ فوری")
        lines.append(f"")
        for r in interview_emails:
            lines.append(f"- **{r['from']}** — {r['subject']} ({r['date']})")
            if r.get("employer_match"):
                lines.append(f"  - کارفرما: {r['employer_match']['name']}")
            lines.append(f"")
    
    # Offers
    offer_emails = [r for r in results if r.get("category") == "offer"]
    if offer_emails:
        lines.append(f"### 🎉 Job Offer — بررسی و پاسخ")
        lines.append(f"")
        for r in offer_emails:
            lines.append(f"- **{r['from']}** — {r['subject']} ({r['date']})")
            lines.append(f"")
    
    # Follow-ups needed
    followup_emails = [r for r in results if r.get("category") == "follow_up"]
    if followup_emails:
        lines.append(f"### ⏰ Follow-up — پیگیری")
        lines.append(f"")
        for r in followup_emails:
            lines.append(f"- **{r['from']}** — {r['subject']} ({r['date']})")
            lines.append(f"")
    
    # Pending responses
    pending_emails = [r for r in results if r.get("category") in ["inquiry", "acknowledgment"] and r.get("is_job_related")]
    if pending_emails:
        lines.append(f"### 📩 در انتظار بررسی")
        lines.append(f"")
        for r in pending_emails[:10]:  # Show top 10
            lines.append(f"- **{r['from']}** — {r['subject']} ({r['date']})")
        lines.append(f"")
    
    # Rejections
    rejection_emails = [r for r in results if r.get("category") == "rejection"]
    if rejection_emails:
        lines.append(f"### ❌ Rejected — بایگانی")
        lines.append(f"")
        for r in rejection_emails:
            lines.append(f"- **{r['from']}** — {r['subject']} ({r['date']})")
        lines.append(f"")
    
    # All job-related emails detail
    lines.append(f"---")
    lines.append(f"")
    lines.append(f"## تمام ایمیل‌های شغلی")
    lines.append(f"")
    lines.append(f"| # | تاریخ | از | موضوع | دسته | متقاضی | کارفرما |")
    lines.append(f"|---|-------|-----|--------|------|--------|---------|")
    
    for idx, r in enumerate(sorted(results, key=lambda x: x.get("date", ""), reverse=True), 1):
        if not r.get("is_job_related"):
            continue
        
        from_short = r.get("from", "")[:40]
        subject_short = r.get("subject", "")[:60]
        category_emoji = {
            "offer": "🎉", "rejection": "❌", "interview": "🗣️",
            "follow_up": "⏰", "acknowledgment": "📩", "inquiry": "💬",
        }.get(r.get("category", ""), "❓")
        
        applicant = r.get("applicant", "?")
        applicant_label = "👩 ندا" if applicant == "NEDA" else "👨 توحید" if applicant == "TOHID" else "?"
        
        employer = r.get("employer_match", {})
        employer_name = employer.get("name", "") if employer else ""
        
        lines.append(f"| {idx} | {r.get('date','')} | {from_short} | {subject_short} | {category_emoji} {r.get('category','')} | {applicant_label} | {employer_name} |")
    
    lines.append(f"")
    lines.append(f"---")
    lines.append(f"")
    lines.append(f"> **آخرین بروزرسانی:** {now}")
    
    return "\n".join(lines)

def generate_email_excel_data(results):
    """تولید داده برای اضافه کردن به Excel"""
    rows = []
    for r in results:
        if not r.get("is_job_related"):
            continue
        
        employer = r.get("employer_match", {})
        rows.append({
            "date": r.get("date", ""),
            "from": r.get("from", ""),
            "subject": r.get("subject", ""),
            "category": r.get("category", ""),
            "applicant": r.get("applicant", ""),
            "employer": employer.get("name", "") if employer else "",
            "country": employer.get("country", "") if employer else "",
            "urgency": r.get("urgency", ""),
            "links": ", ".join(r.get("links", [])[:3]),
        })
    return rows

# ═══════════════════════════════════════════════════
# SAVE TO MEMORY
# ═══════════════════════════════════════════════════
def save_to_memory(results, email_addr):
    """ذخیره نتایج در memory/"""
    os.makedirs(MEM, exist_ok=True)
    
    email_data = []
    for r in results:
        if not r.get("is_job_related"):
            continue
        employer = r.get("employer_match", {})
        email_data.append({
            "date": r.get("date", ""),
            "from": r.get("from", ""),
            "subject": r.get("subject", ""),
            "category": r.get("category", ""),
            "applicant": r.get("applicant", ""),
            "employer": employer.get("name", "") if employer else "",
            "urgency": r.get("urgency", ""),
        })
    
    filepath = os.path.join(MEM, "EMAIL_ANALYSIS.json")
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump({
            "email": email_addr,
            "analyzed_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "total_emails": len(results),
            "job_related": len(email_data),
            "emails": email_data,
        }, f, ensure_ascii=False, indent=2)
    
    print(f"  💾 ذخیره شد: {filepath}")

# ═══════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════
def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="MigrationHunter Email Analyzer")
    parser.add_argument("--email", help="Email address")
    parser.add_argument("--password", help="App Password (NOT regular password)")
    parser.add_argument("--provider", default="gmail", 
                       choices=["gmail", "outlook", "yahoo", "icloud"],
                       help="Email provider")
    parser.add_argument("--folder", default="INBOX", help="Email folder")
    parser.add_argument("--days", type=int, default=30, help="Days to look back")
    parser.add_argument("--limit", type=int, default=200, help="Max emails to fetch")
    parser.add_argument("--dry-run", action="store_true", help="Test connection only")
    args = parser.parse_args()
    
    # Load from .env if not provided
    if not args.email:
        env_file = os.path.join(BASE, ".env")
        if os.path.exists(env_file):
            with open(env_file) as f:
                for line in f:
                    line = line.strip()
                    if line.startswith("EMAIL_ADDRESS="):
                        args.email = line.split("=", 1)[1].strip().strip('"').strip("'")
                    elif line.startswith("EMAIL_PASSWORD="):
                        args.password = line.split("=", 1)[1].strip().strip('"').strip("'")
                    elif line.startswith("EMAIL_PROVIDER="):
                        args.provider = line.split("=", 1)[1].strip().strip('"').strip("'")
    
    if not args.email or not args.password:
        print("❌ لطفاً ایمیل و رمز عبور را وارد کنید:")
        print()
        print("روش ۱:")
        print("  python email_analyzer.py --email you@gmail.com --password APP_PASSWORD")
        print()
        print("روش ۲:")
        print("  فایل .env بسازید با:")
        print("  EMAIL_ADDRESS=you@gmail.com")
        print("  EMAIL_PASSWORD=your_app_password")
        print("  EMAIL_PROVIDER=gmail")
        print()
        print("Gmail App Password:")
        print("  1. myaccount.google.com → Security → 2-Step Verification → ON")
        print("  2. myaccount.google.com → Security → App passwords → Create")
        print("  3. کد 16 رقمی را به عنوان password استفاده کنید")
        sys.exit(1)
    
    print("=" * 60)
    print("MigrationHunter — Email Analyzer")
    print(f"📧 {args.email}")
    print(f"🔌 Provider: {args.provider}")
    print(f"📅 Days: {args.days}")
    print("=" * 60)
    
    # Connect
    connector = EmailConnector(args.email, args.password, args.provider)
    try:
        connector.connect()
    except imaplib.IMAP4.error as e:
        print(f"\n❌ خطا در اتصال: {e}")
        print("\npossible causes:")
        print("  - App Password اشتباه")
        print("  - IMAP فعال نیست")
        print("  - 2FA فعال نیست (Gmail)")
        sys.exit(1)
    
    if args.dry_run:
        print("\n✅ اتصال موفق! (dry-run mode)")
        folders = connector.get_folders()
        print(f"\n📁 پوشه‌ها:")
        for f in folders[:20]:
            print(f"  - {f}")
        connector.disconnect()
        return
    
    # Search emails
    print(f"\n🔍 جستجوی ایمیل‌های {args.days} روز اخیر...")
    emails = connector.search_emails(
        folder=args.folder,
        days=args.days,
        limit=args.limit,
    )
    print(f"  📩 {len(emails)} ایمیل یافت شد")
    
    connector.disconnect()
    
    if not emails:
        print("\n⚠️ ایمیلی یافت نشد")
        return
    
    # Analyze
    print("\n🔬 تحلیل ایمیل‌ها...")
    analyzer = EmailAnalyzer()
    results = analyzer.analyze_all(emails)
    
    job_related = [r for r in results if r.get("is_job_related")]
    print(f"  ✅ {len(job_related)} ایمیل شغلی شناسایی شد")
    
    # Generate report
    print("\n📊 تولید گزارش...")
    report = generate_report(results, args.email)
    
    os.makedirs(OUT, exist_ok=True)
    report_path = os.path.join(OUT, "EMAIL_ANALYSIS_REPORT.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"  📄 {report_path}")
    
    # Save to memory
    print("\n💾 ذخیره در حافظه...")
    save_to_memory(results, args.email)
    
    # Summary
    print("\n" + "=" * 60)
    print("📊 خلاصه")
    print("=" * 60)
    
    categories = defaultdict(int)
    for r in job_related:
        categories[r.get("category", "unknown")] += 1
    
    emoji_map = {"offer": "🎉", "rejection": "❌", "interview": "🗣️",
                 "follow_up": "⏰", "acknowledgment": "📩", "inquiry": "💬"}
    
    for cat, count in sorted(categories.items(), key=lambda x: -x[1]):
        emoji = emoji_map.get(cat, "❓")
        print(f"  {emoji} {cat}: {count}")
    
    print(f"\n  📄 گزارش: output/EMAIL_ANALYSIS_REPORT.md")
    print(f"  💾 حافظه: memory/EMAIL_ANALYSIS.json")
    print("=" * 60)

if __name__ == "__main__":
    main()
