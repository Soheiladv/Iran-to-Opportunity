#!/usr/bin/env python3
"""
MigrationHunter — یادآوری پیگیری ۷ روزه
بررسی ایمیل‌های بی‌پاسخ و تولید لیست پیگیری

اجرا: python followup_reminder.py
"""
import os, sys, json, io
from datetime import datetime, timedelta

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
# تشخیص ایمیل‌های نیاز به پیگیری
# ═══════════════════════════════════════════════════
def find_needs_followup(data):
    """پیدا کردن ایمیل‌هایی که نیاز به پیگیری دارند"""
    emails = data.get("emails", [])
    followup_needed = []
    
    for e in emails:
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
    
    return followup_needed

# ═══════════════════════════════════════════════════
# تولید گزارش
# ═══════════════════════════════════════════════════
def generate_report(followup_list, data):
    """تولید گزارش فارسی"""
    lines = []
    lines.append(f"# یادآوری پیگیری ایمیل‌ها")
    lines.append(f"")
    lines.append(f"**تاریخ:** {DATE_STR}")
    lines.append(f"**ایمیل‌های نیاز به پیگیری:** {len(followup_list)}")
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
            app_label = "👩 ندا" if app == "NEDA" else "👨 توحید" if app == "TOHID" else "❓"
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
    followup_list = find_needs_followup(data)
    
    print(f"  📩 {len(followup_list)} ایمیل نیاز به پیگیری")
    
    # گزارش
    print(f"\n📄 تولید گزارش...")
    report = generate_report(followup_list, data)
    
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
            app = "👩 ندا" if e.get("applicant") == "NEDA" else "👨 توحید" if e.get("applicant") == "TOHID" else "❓"
            sender = e.get("from", "").split("<")[0].strip().strip('"')[:30]
            print(f"  {e.get('priority', '')} {e.get('days_since', 0)} روز — {sender} — {app}")
    else:
        print(f"  ✅ همه ایمیل‌ها پاسخ داده شده")
    
    print("═" * 50)

OK = "✅"

if __name__ == "__main__":
    main()
