#!/usr/bin/env python3
"""
MigrationHunter — یادآوری پیگیری ۷ روزه
بررسی ایمیل‌های بی‌پاسخ و تولید لیست پیگیری

اجرا: python followup_reminder.py
"""
import os, sys, json, io, re
from datetime import datetime, timedelta
from config_loader import get_applicant_label

# Fix Windows console encoding for emoji support
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

BASE = os.path.dirname(os.path.abspath(__file__))
MEM = os.path.join(BASE, "memory")
OUT = os.path.join(BASE, "output")

NOW = datetime.now()
DATE_STR = NOW.strftime("%Y-%m-%d %H:%M")
TODAY = NOW.date()

# ═══════════════════════════════════════════════════
# بارگذاری ایمیل‌ها
# ═══════════════════════════════════════════════════
def load_emails():
    fp = os.path.join(MEM, "EMAIL_ANALYSIS.json")
    if not os.path.exists(fp):
        return None
    with open(fp, "r", encoding="utf-8") as f:
        return json.load(f)

# ═══════════════════════════════════════════════════
# تشخیص خبرنامه‌ها و ایمیل‌های تبلیغاتی
# ═══════════════════════════════════════════════════
# دامنه‌های شناخته‌شده خبرنامه/جاب‌ایجنت (نه کارفرمای واقعی)
NEWSLETTER_SENDERS = [
    "hokify",          # جاب‌ایجنت اتریشی
    "stepstone",       # جاب‌ایجنت آلمانی
    "wayup",           # پلتفرم کارآموزی آمریکا
    "canadavisa",      # خبرنامه مهاجرت CIC News
    "jobagent",
    "linkedin.com",
    "indeed.com",
    "glassdoor",
    "mailchimp",
    "sendgrid",
    "campaign",
    "newsletter",
    "notify",
    "no-reply",
    "noreply",
    "donotreply",
]

# الگوهای موضوعی خبرنامه (فقط زمانی فیلتر می‌شود که فرستنده هم مطابق باشد
# یا این الگوها صریحاً تبلیغاتی باشند)
NEWSLETTER_SUBJECT_PATTERNS = [
    r"jobs?( could| might| die)? (können|passen|wären|sein)",  # hokify آلمانی
    r"diese jobs", r"ein spannender job", r"dein neuer job",
    r"dieser job", r"jobs, die zu dir passen",
    r"you have a great chance", r"your cv is a great match",
    r"closing soon, apply",
    r"\b\d+ interviews were booked",
    r"hidden\"? job market",
    r"\bnewsletter\b", r"\bdigest\b",
    r"sei dabei",
]

# خبرنامه‌های اطلاع‌رسانی مهاجرت — ارزش اطلاعاتی دارند ولی پاسخ کارفرما نیستند
NEWSLETTER_INFO_PATTERNS = [
    r"\bcanada (invites|expands|extends)\b",
    r"\bregistration deadline\b",
    r"\bexpress entry\b\b?\bdraw\b",
]


def is_newsletter(email):
    """
    تشخیص اینکه ایمیل، خبرنامه/تبلیغ جاب‌ایجنت است نه پاسخ واقعی کارفرما.
    خروجی: (bool, دلیل)
    """
    sender = (email.get("from", "") or "").lower()
    subject = (email.get("subject", "") or "").lower()

    # ۱. دامنه/نام فرستنده در لیست سیاه خبرنامه‌ها
    for name in NEWSLETTER_SENDERS:
        if name in sender:
            return True, f"فرستنده خبرنامه/جاب‌ایجنت: {name}"

    # ۲. الگوهای صریح تبلیغاتی در موضوع
    for pat in NEWSLETTER_SUBJECT_PATTERNS:
        if re.search(pat, subject, re.I):
            return True, f"موضوع تبلیغاتی"

    return False, ""


def find_needs_followup(data):
    """پیدا کردن ایمیل‌هایی که نیاز به پیگیری دارند (خبرنامه‌ها فیلتر می‌شوند)"""
    emails = data.get("emails", [])
    followup_needed = []
    newsletters = []
    
    for e in emails:
        # فیلتر خبرنامه‌ها: پاسخ کارفرما نیستند، پیگیری ندارند
        is_nl, nl_reason = is_newsletter(e)
        if is_nl:
            newsletters.append({**e, "newsletter_reason": nl_reason})
            continue
        
        cat = e.get("category", "")
        
        # فقط ایمیل‌های شغلی که نیاز به پاسخ دارند
        if cat not in ["inquiry", "acknowledgment", "interview"]:
            continue
        
        # تاریخ ایمیل
        date_str = e.get("date", "")
        if not date_str:
            continue
        
        try:
            email_date = datetime.strptime(date_str[:10], "%Y-%m-%d").date()
        except:
            continue
        
        # محاسبه روزهای سپری‌شده
        days_since = (TODAY - email_date).days
        
        # فقط ایمیل‌های بیش از ۷ روز
        if days_since < 7:
            continue
        
        # اولویت‌بندی
        if days_since >= 14:
            priority = "🔴 فوری"
        elif days_since >= 10:
            priority = "🟡 مهم"
        else:
            priority = "🟢 عادی"
        
        followup_needed.append({
            **e,
            "days_since": days_since,
            "priority": priority,
            "email_date": email_date.isoformat(),
        })
    
    # مرتب‌سازی بر اساس اولویت و تاریخ
    followup_needed.sort(key=lambda x: -x["days_since"])
    
    return followup_needed, newsletters

# ═══════════════════════════════════════════════════
# تولید گزارش
# ═══════════════════════════════════════════════════
def generate_report(followup_list, newsletters, data):
    """تولید گزارش فارسی"""
    lines = []
    lines.append(f"# یادآوری پیگیری ایمیل‌ها")
    lines.append(f"")
    lines.append(f"**تاریخ:** {DATE_STR}")
    lines.append(f"**ایمیل‌های نیاز به پیگیری:** {len(followup_list)}")
    lines.append(f"**خبرنامه‌های فیلترشده:** {len(newsletters)}")
    lines.append(f"")
    lines.append(f"---")
    lines.append(f"")
    
    if not followup_list:
        lines.append(f"✅ هیچ ایمیلی نیاز به پیگیری ندارد.")
        lines.append(f"")
    else:
        # آمار
        urgent = sum(1 for e in followup_list if "فوری" in e.get("priority", ""))
        important = sum(1 for e in followup_list if "مهم" in e.get("priority", ""))
        normal = sum(1 for e in followup_list if "عادی" in e.get("priority", ""))
        
        lines.append(f"## خلاصه")
        lines.append(f"")
        lines.append(f"| اولویت | تعداد |")
        lines.append(f"|--------|-------|")
        lines.append(f"| 🔴 فوری (۱۴+ روز) | {urgent} |")
        lines.append(f"| 🟡 مهم (۱۰-۱۳ روز) | {important} |")
        lines.append(f"| 🟢 عادی (۷-۹ روز) | {normal} |")
        lines.append(f"")
        
        # لیست
        lines.append(f"## لیست پیگیری")
        lines.append(f"")
        lines.append(f"| # | اولویت | روز | تاریخ | فرستنده | موضوع | متقاضی | اقدام |")
        lines.append(f"|---|--------|-----|-------|---------|-------|--------|-------|")
        
        for idx, e in enumerate(followup_list, 1):
            app = e.get("applicant", "?")
            app_label = get_applicant_label(app) if app in ['NEDA', 'TOHID'] else "?"
            sender = e.get("from", "").split("<")[0].strip().strip('"')[:25]
            subject = e.get("subject", "")[:45]
            
            # اقدام پیشنهادی
            if e.get("days_since", 0) >= 14:
                action = "پیگیری فوری"
            elif e.get("category") == "interview":
                action = "تایید حضور"
            else:
                action = "یادآوری"
            
            lines.append(f"| {idx} | {e.get('priority', '')} | {e.get('days_since', '')} | {e.get('email_date', '')} | {sender} | {subject} | {app_label} | {action} |")
        
        lines.append(f"")
    
    # خبرنامه‌های فیلترشده (شفافیت: چیزی مخفی نمی‌شود)
    if newsletters:
        lines.append(f"## خبرنامه‌های فیلترشده (پاسخ کارفرما نیستند)")
        lines.append(f"")
        lines.append(f"| # | تاریخ | فرستنده | موضوع | دلیل فیلتر |")
        lines.append(f"|---|-------|---------|-------|------------|")
        for idx, e in enumerate(newsletters, 1):
            sender = e.get("from", "").split("<")[0].strip().strip('"')[:28]
            subject = e.get("subject", "")[:45]
            reason = e.get("newsletter_reason", "")
            lines.append(f"| {idx} | {e.get('date', '')[:10]} | {sender} | {subject} | {reason} |")
        lines.append(f"")
    
    lines.append(f"---")
    lines.append(f"")
    lines.append(f"> **آخرین بررسی:** {DATE_STR}")
    
    return "\n".join(lines)

# ═══════════════════════════════════════════════════
# اجرا
# ═══════════════════════════════════════════════════
def main():
    print("═" * 50)
    print("MigrationHunter — یادآوری پیگیری")
    print(f"📅 {DATE_STR}")
    print("═" * 50)
    
    data = load_emails()
    if not data:
        print(f"\n⚠️ ابتدا email_analyzer.py را اجرا کنید")
        return
    
    print(f"\n🔍 بررسی ایمیل‌های بی‌پاسخ...")
    followup_list, newsletters = find_needs_followup(data)
    
    print(f"  📩 {len(followup_list)} ایمیل نیاز به پیگیری")
    print(f"  🗑️ {len(newsletters)} خبرنامه فیلتر شد")
    
    # گزارش
    print(f"\n📄 تولید گزارش...")
    report = generate_report(followup_list, newsletters, data)
    
    os.makedirs(OUT, exist_ok=True)
    fp = os.path.join(OUT, "FOLLOWUP_REMINDER.md")
    with open(fp, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"  {OK} {fp}")
    
    # خلاصه
    print("\n" + "═" * 50)
    print("📊 خلاصه")
    print("═" * 50)
    
    if followup_list:
        for e in followup_list[:10]:
            app = get_applicant_label(e.get("applicant", "").lower()) if e.get("applicant") in ['NEDA', 'TOHID'] else "?"
            sender = e.get("from", "").split("<")[0].strip().strip('"')[:30]
            print(f"  {e.get('priority', '')} {e.get('days_since', 0)} روز — {sender} — {app}")
    else:
        print(f"  ✅ همه ایمیل‌ها پاسخ داده شده")
    
    print("═" * 50)

OK = "✅"

if __name__ == "__main__":
    main()
