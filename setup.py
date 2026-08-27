#!/usr/bin/env python3
"""
MigrationHunter — راه‌اندازی سیستم جدید
هر آنچه برای شروع لازم است را می‌پرسد و تنظیم می‌کند

اجرا: python setup.py
"""
import os, sys, json, shutil, io
from datetime import datetime

# Fix Windows console encoding for emoji support
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

BASE = os.path.dirname(os.path.abspath(__file__))

# ═══════════════════════════════════════════════════
# رنگ‌ها و نمادها
# ═══════════════════════════════════════════════════
OK = "✅"
WARN = "⚠️"
FAIL = "❌"
INFO = "ℹ️"
STEP = "▶"

def banner():
    print("═" * 60)
    print("  MigrationHunter — راه‌اندازی سیستم جدید")
    print(f"  📅 {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("═" * 60)

def ask(prompt, default="", secret=False):
    """پرسیدن سوال از کاربر"""
    suffix = f" [{default}]" if default else ""
    if secret:
        val = input(f"  {prompt}{suffix}: ").strip()
    else:
        val = input(f"  {prompt}{suffix}: ").strip()
    return val if val else default

def check_mark(condition, msg):
    if condition:
        print(f"  {OK} {msg}")
    else:
        print(f"  {WARN} {msg} — خالی ماند")

# ═══════════════════════════════════════════════════
# مرحله ۱: بررسی محیط
# ═══════════════════════════════════════════════════
def check_environment():
    print(f"\n{STEP} مرحله ۱: بررسی محیط")
    print("─" * 40)
    
    # Python
    import sys
    py_ver = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    print(f"  {OK} Python: {py_ver}")
    
    # Required packages
    required = ["openpyxl"]
    missing = []
    for pkg in required:
        try:
            __import__(pkg)
            print(f"  {OK} {pkg}: نصب است")
        except ImportError:
            print(f"  {FAIL} {pkg}: نصب نیست")
            missing.append(pkg)
    
    if missing:
        print(f"\n  {INFO} نصب پکیج‌ها...")
        for pkg in missing:
            os.system(f"pip install {pkg}")
    
    # Directories
    dirs = ["memory", "profiles", "output", "dashboard", "dashboard/archive", "input"]
    for d in dirs:
        path = os.path.join(BASE, d)
        if not os.path.isdir(path):
            os.makedirs(path, exist_ok=True)
            print(f"  {OK} پوشه ساخته شد: {d}")
        else:
            print(f"  {OK} پوشه موجود: {d}")

# ═══════════════════════════════════════════════════
# مرحله ۲: پروفایل متقاضیان
# ═══════════════════════════════════════════════════
def setup_profiles():
    print(f"\n{STEP} مرحله ۲: پروفایل متقاضیان")
    print("─" * 40)
    
    profiles = {}
    
    # TOHID
    print(f"\n  👨 اطلاعات توحید:")
    profiles["tohid"] = {
        "name": ask("نام کامل", "Tohid Arjmand"),
        "age": ask("سن", "46"),
        "country": ask("کشور", "Iran"),
        "profession": ask("حرفه اصلی", "IT Operations Manager"),
        "education": ask("تحصیلات", "Bachelor of IT"),
        "experience": ask("سابقه کاری (سال)", "19"),
        "english": ask("سطح انگلیسی", "A2"),
        "german": ask("سطح آلمانی", "A1"),
        "linkedin": ask("آدرس LinkedIn", "https://www.linkedin.com/in/tohid-arjmand"),
        "email_primary": ask("ایمیل اصلی"),
    }
    
    # NEDA
    print(f"\n  👩 اطلاعات ندا:")
    profiles["neda"] = {
        "name": ask("نام کامل", "Neda Arjmand"),
        "age": ask("سن", "38"),
        "country": ask("کشور", "Iran"),
        "profession": ask("حرفه اصلی", "Midwife"),
        "current_job": ask("شغل فعلی", "Milad Hospital, Tehran"),
        "education": ask("تحصیلات", "Bachelor of Midwifery"),
        "english": ask("سطح انگلیسی", "A2"),
        "german": ask("سطح آلمانی", "A1"),
        "linkedin": ask("آدرس LinkedIn", "https://www.linkedin.com/in/neda-arjmand"),
        "email_primary": ask("ایمیل اصلی"),
    }
    
    # ذخیره
    for person, data in profiles.items():
        fp = os.path.join(BASE, "profiles", f"{person.upper()}_PROFILE.md")
        with open(fp, "w", encoding="utf-8") as f:
            f.write(f"# {data['name'].upper()} — Profile\n\n")
            for k, v in data.items():
                if v:
                    f.write(f"| {k} | {v} |\n")
            f.write(f"\n**Last Updated:** {datetime.now().strftime('%Y-%m-%d')}\n")
        print(f"  {OK} ذخیره شد: {fp}")
    
    return profiles

# ═══════════════════════════════════════════════════
# مرحله ۳: ایمیل‌ها
# ═══════════════════════════════════════════════════
def setup_emails():
    print(f"\n{STEP} مرحله ۳: تنظیم ایمیل‌ها")
    print("─" * 40)
    print(f"  {INFO} برای هر ایمیل به App Password نیاز دارید")
    print(f"  {INFO} راهنما: EMAIL_SETUP_GUIDE_FA.md")
    
    accounts = []
    
    for person, label in [("TOHID", "توحید"), ("NEDA", "ندا")]:
        print(f"\n  👤 ایمیل {label}:")
        email = ask(f"ایمیل {label}")
        password = ask(f"App Password {label} (16 رقمی)", secret=True)
        provider = ask("سرویس‌دهنده", "gmail")
        linkedin = ask(f"لینکدین {label}")
        
        if email:
            accounts.append({
                "id": f"{person.lower()}_1",
                "person": person,
                "person_fa": label,
                "email": email,
                "provider": provider,
                "label": f"ایمیل اصلی {label}",
                "linkedin": linkedin,
                "active": True,
            })
            
            # نوشتن در .env
            env_line = f"EMAIL_PASSWORD_{person.upper()}_1={password}"
            print(f"  {OK} در .env ذخیره شد")
        else:
            print(f"  {WARN} ایمیل {label} خالی ماند — بعداً تنظیم کنید")
    
    # ذخیره email_accounts.json
    if accounts:
        fp = os.path.join(BASE, "email_accounts.json")
        with open(fp, "w", encoding="utf-8") as f:
            json.dump({"accounts": accounts}, f, ensure_ascii=False, indent=2)
        print(f"  {OK} ذخیره شد: {fp}")
    
    return accounts

# ═══════════════════════════════════════════════════
# مرحله ۴: نوشتن .env
# ═══════════════════════════════════════════════════
def write_env(accounts):
    print(f"\n{STEP} مرحله ۴: نوشتن .env")
    print("─" * 40)
    
    env_path = os.path.join(BASE, ".env")
    lines = [
        "# MigrationHunter — Environment Config",
        f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
    ]
    
    for acc in accounts:
        person = acc["person"]
        lines.append(f"# {acc.get('person_fa', person)}")
        lines.append(f"EMAIL_{person}_1={acc['email']}")
        lines.append(f"EMAIL_PASSWORD_{person}_1=REPLACE_WITH_APP_PASSWORD")
        lines.append(f"EMAIL_PROVIDER_{person}_1={acc.get('provider', 'gmail')}")
        lines.append("")
    
    lines.append("# OpenAI API (اختیاری)")
    lines.append("OPENAI_API_KEY=your_key_here")
    
    with open(env_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    
    print(f"  {OK} نوشته شد: .env")
    print(f"  {WARN} App Passwordها را در .env وارد کنید!")

# ═══════════════════════════════════════════════════
# مرحله ۵: تأیید نهایی
# ═══════════════════════════════════════════════════
def final_check():
    print(f"\n{STEP} مرحله ۵: تأیید نهایی")
    print("─" * 40)
    
    checks = [
        ("فایل .env", os.path.exists(os.path.join(BASE, ".env"))),
        ("فایل email_accounts.json", os.path.exists(os.path.join(BASE, "email_accounts.json"))),
        ("پوشه memory", os.path.isdir(os.path.join(BASE, "memory"))),
        ("پوشه profiles", os.path.isdir(os.path.join(BASE, "profiles"))),
        ("پوشه dashboard", os.path.isdir(os.path.join(BASE, "dashboard"))),
        ("پوشه output", os.path.isdir(os.path.join(BASE, "output"))),
        ("اسکریپت run.py", os.path.exists(os.path.join(BASE, "run.py"))),
        ("اسکریپت build_dashboard.py", os.path.exists(os.path.join(BASE, "build_dashboard.py"))),
        ("اسکریپت email_analyzer.py", os.path.exists(os.path.join(BASE, "email_analyzer.py"))),
    ]
    
    for name, ok in checks:
        check_mark(ok, name)
    
    all_ok = all(ok for _, ok in checks)
    
    print("\n" + "═" * 60)
    if all_ok:
        print(f"  {OK} راه‌اندازی کامل شد!")
        print(f"\n  برای شروع:")
        print(f"    python run.py")
    else:
        print(f"  {WARN} بعضی موارد خالی ماند — بعداً تکمیل کنید")
        print(f"\n  برای شروع (با اطلاعات ناقص):")
        print(f"    python run.py")
    print("═" * 60)

# ═══════════════════════════════════════════════════
# اجرا
# ═══════════════════════════════════════════════════
def main():
    banner()
    
    check_environment()
    profiles = setup_profiles()
    accounts = setup_emails()
    write_env(accounts)
    final_check()

if __name__ == "__main__":
    main()
