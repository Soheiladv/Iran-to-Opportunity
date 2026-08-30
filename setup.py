#!/usr/bin/env python3
"""
MigrationHunter — راه‌اندازی داینامیک
تنها منبع حقیقت: config.json + .env
هیچ hardcode وجود ندارد
"""
import os, sys, json, io
from datetime import datetime

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

BASE = os.path.dirname(os.path.abspath(__file__))

def banner():
    print("=" * 60)
    print("  MigrationHunter — راه‌اندازی داینامیک")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    val = input(f"  {prompt}{suffix}: ").strip()
    return val if val else default

def ask_int(prompt, default=1):
    while True:
        val = ask(prompt, str(default))
        try:
            return int(val)
        except ValueError:
            print("  ⚠️ عدد وارد کنید")

def main():
    banner()

    # ─── تعداد متقاضیان ───
    print("\n  ─── تنظیمات کلی ───")
    num_applicants = ask_int("چند متقاضی دارید؟", 1)

    applicants = []

    for i in range(1, num_applicants + 1):
        print(f"\n  ─── متقاضی {i} ───")
        
        name = ask("نام کامل (English)")
        name_fa = ask("نام فارسی")
        profession = ask("حرفه (مثلاً English Teacher, Midwife, IT Manager)")
        
        # جنسیت - پرسیده می‌شود نه فرض
        gender = ask("جنسیت (male/female)", "female").lower()
        if gender not in ("male", "female"):
            gender = "female"
        emoji = "👨" if gender == "male" else "👩"
        
        english = ask("سطح انگلیسی", "A2")
        german = ask("سطح آلمانی", "A1")

        # ─── ایمیل‌ها ───
        print(f"\n  📧 ایمیل‌های {name_fa or name}:")
        num_emails = ask_int("  چند ایمیل می‌خواهید ثبت کنید؟", 1)
        emails = []
        for e in range(1, num_emails + 1):
            email = ask(f"  ایمیل {e} (Gmail)")
            if email:
                emails.append(email)
        
        # ─── لینکدین‌ها ───
        print(f"\n  🔗 لینکدین {name_fa or name}:")
        num_linkedins = ask_int("  چند پروفایل لینکدین دارید؟", 1)
        linkedins = []
        for l in range(1, num_linkedins + 1):
            li = ask(f"  لینکدین {l} (URL کامل)")
            if li:
                linkedins.append(li)

        if name or emails:
            # id از نام فارسی یا انگلیسی
            name_id = (name_fa or name or f"person{i}").lower().replace(" ", "_")
            name_id = "".join(c for c in name_id if c.isalnum() or c == "_")
            
            # کلمات کلیدی
            keywords = []
            if profession:
                keywords.extend(profession.lower().split())
            if name_fa:
                keywords.append(name_fa.lower())
            if name:
                keywords.extend(name.lower().split())

            applicants.append({
                "id": name_id,
                "name": name,
                "name_fa": name_fa,
                "gender": gender,
                "emoji": emoji,
                "profession": profession,
                "keywords": keywords,
                "emails": emails,
                "linkedins": linkedins,
                "english": english,
                "german": german,
            })
            print(f"  ✅ ذخیره شد: {name_fa or name_id} ({len(emails)} ایمیل، {len(linkedins)} لینکدین)\n")
        else:
            print(f"  ⏭️ خالی ماند\n")

    if not applicants:
        print("\n  ❌ هیچ متقاضی ثبت نشد!")
        return

    # ─── نوشتن config.json ───
    config = {
        "project": "MigrationHunter",
        "version": "3.0",
        "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "applicants": applicants,
    }
    config_path = os.path.join(BASE, "config.json")
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"  ✅ ذخیره شد: config.json")

    # ─── نوشتن .env ───
    print("\n  ─── رمزهای App Password ───")
    print("  (راهنما: https://myaccount.google.com/apppasswords)")
    print("  برای هر ایمیل یک رمز 16 رقمی نیاز است\n")

    env_lines = [
        "# MigrationHunter — Passwords",
        f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
    ]

    for a in applicants:
        app_id = a["id"].upper()
        print(f"  {a['emoji']} {a.get('name_fa', a['id'])} — {len(a['emails'])} ایمیل")
        
        for idx, email in enumerate(a["emails"], 1):
            pw = ask(f"    App Password برای {email}", "REPLACE_WITH_APP_PASSWORD")
            env_lines.append(f"# {a.get('name_fa', a['id'])} — {email}")
            env_lines.append(f"EMAIL_{app_id}_{idx}={email}")
            env_lines.append(f"EMAIL_PASSWORD_{app_id}_{idx}={pw}")
            env_lines.append(f"EMAIL_PROVIDER_{app_id}_{idx}=gmail")
            env_lines.append("")

    env_path = os.path.join(BASE, ".env")
    with open(env_path, "w", encoding="utf-8") as f:
        f.write("\n".join(env_lines))
    print(f"\n  ✅ ذخیره شد: .env")

    # ─── ساخت پوشه‌ها ───
    for d in ["memory", "profiles", "output", "dashboard", "dashboard/archive", "input"]:
        os.makedirs(os.path.join(BASE, d), exist_ok=True)

    # ─── خلاصه ───
    print("\n" + "=" * 60)
    print("  راه‌اندازی کامل شد!")
    print("=" * 60)
    print(f"\n  فایل‌های ساخته شده:")
    print(f"    config.json    ← اطلاعات متقاضیان (داینامیک)")
    print(f"    .env           ← رمزهای ایمیل (چندین ایمیل برای هر نفر)")
    print(f"\n  برای شروع:")
    print(f"    python run.py")
    print(f"\n  برای تست اتصال ایمیل:")
    print(f"    python email_analyzer.py --dry-run")
    print("=" * 60)

if __name__ == "__main__":
    main()