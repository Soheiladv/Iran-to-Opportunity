#!/usr/bin/env python3
"""
MigrationHunter — راه‌اندازی ساده
فقط یکبار اجرا کن → config.json + .env ساخته می‌شود

اجرا: python setup.py
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
    print("  MigrationHunter — راه‌اندازی ساده")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    val = input(f"  {prompt}{suffix}: ").strip()
    return val if val else default

def main():
    banner()

    # ─── پرسیدن اطلاعات ───
    print("\n  اطلاعات هر نفر را وارد کن.")
    print("  اگر خالی بگذاری، بعداً قابل تغییر است.\n")

    applicants = []

    for i in range(1, 3):
        print(f"  ─── متقاضی {i} ───")
        name = ask("نام کامل (English)")
        name_fa = ask("نام فارسی")
        profession = ask("حرفه (مثلاً Midwife یا IT Manager)")
        email = ask("ایمیل Gmail")
        linkedin = ask("لینکدین (URL)")
        english = ask("سطح انگلیسی", "A2")
        german = ask("سطح آلمانی", "A1")

        if name or email:
            # تشخیص خودکار
            name_id = name.lower().split()[0] if name else f"person{i}"
            gender = "male" if i == 1 else "female"
            emoji = "👨" if gender == "male" else "👩"

            # کلمات کلیدی
            keywords = []
            if profession:
                keywords.extend(profession.lower().split())
            if name_fa:
                keywords.append(name_fa.lower())

            applicants.append({
                "id": name_id,
                "name": name,
                "name_fa": name_fa,
                "gender": gender,
                "emoji": emoji,
                "profession": profession,
                "keywords": keywords,
                "email": email,
                "linkedin": linkedin,
                "english": english,
                "german": german,
            })
            print(f"  OK ذخیره شد: {name_fa or name_id}\n")
        else:
            print(f"  SKIP خالی ماند\n")

    if not applicants:
        print("\n  ERROR هيچ متقاضی ثبت نشد!")
        return

    # ─── نوشتن config.json ───
    config = {
        "project": "MigrationHunter",
        "version": "2.0",
        "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "applicants": applicants,
    }
    config_path = os.path.join(BASE, "config.json")
    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"  OK ذخیره شد: config.json")

    # ─── نوشتن .env ───
    print("\n  حالا رمزهای App Password را وارد کن:")
    print("  (راهنما: https://myaccount.google.com/apppasswords)\n")

    env_lines = [
        "# MigrationHunter — Passwords",
        f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "",
    ]

    for a in applicants:
        app_id = a["id"].upper()
        print(f"  رمز App Password {a.get('name_fa', a['id'])}:")
        pw = input(f"  Password (16 رقمی) [{a['email']}]: ").strip()

        env_lines.append(f"# {a.get('name_fa', a['id'])}")
        env_lines.append(f"EMAIL_{app_id}_1={a['email']}")
        env_lines.append(f"EMAIL_PASSWORD_{app_id}_1={pw or 'REPLACE_WITH_APP_PASSWORD'}")
        env_lines.append(f"EMAIL_PROVIDER_{app_id}_1=gmail")
        env_lines.append("")

    env_path = os.path.join(BASE, ".env")
    with open(env_path, "w", encoding="utf-8") as f:
        f.write("\n".join(env_lines))
    print(f"\n  OK ذخیره شد: .env")

    # ─── ساخت پوشه‌ها ───
    for d in ["memory", "profiles", "output", "dashboard", "dashboard/archive", "input"]:
        os.makedirs(os.path.join(BASE, d), exist_ok=True)

    # ─── خلاصه ───
    print("\n" + "=" * 60)
    print("  راه‌اندازی کامل شد!")
    print("=" * 60)
    print(f"\n  فایل‌های ساخته شده:")
    print(f"    config.json    ← اطلاعات متقاضیان")
    print(f"    .env           ← رمزهای ایمیل")
    print(f"\n  برای شروع:")
    print(f"    python run.py")
    print(f"\n  برای تست اتصال ایمیل:")
    print(f"    python email_analyzer.py --dry-run")
    print("=" * 60)

if __name__ == "__main__":
    main()
