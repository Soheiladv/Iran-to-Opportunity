#!/usr/bin/env python3
"""
MigrationHunter — راه‌اندازی داینامیک
تنها منبع حقیقت: config.json + .env
هیچ hardcode وجود ندارد

اجرا:
    python setup.py
"""
import getpass
import io
import json
import os
import subprocess
import sys
from datetime import datetime

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE, "config.json")
ENV_PATH = os.path.join(BASE, ".env")
GITIGNORE_PATH = os.path.join(BASE, ".gitignore")

# کلماتی که در انگلیسی/فارسی معنی خاصی به‌عنوان کلیدواژه‌ی شغلی ندارند —
# اگر بدون فیلتر باشند، تشخیص ایمیل/آگهی را با نویز پر می‌کنند
STOPWORDS = {
    "and", "or", "the", "a", "an", "of", "for", "in", "to", "at",
    "و", "در", "به", "از", "با", "برای", "را",
}


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


def ask_yes_no(prompt, default=True):
    d = "y" if default else "n"
    val = ask(f"{prompt} (y/n)", d).strip().lower()
    return val in ("y", "yes", "بله", "آره", "1")


def ask_secret(prompt):
    """
    رمز را بدون نمایش روی صفحه می‌گیرد (مثل رمز لینوکس/گیت).
    اگر ترمینال از حالت مخفی پشتیبانی نکند (بعضی IDEها)، به input عادی
    برمی‌گردد و هشدار می‌دهد که روی صفحه دیده می‌شود.
    """
    try:
        val = getpass.getpass(f"    {prompt}: ").strip()
        return val
    except Exception:
        print("    ⚠️ این ترمینال حالت مخفی رمز را پشتیبانی نمی‌کند — رمز روی صفحه دیده می‌شود.")
        return input(f"    {prompt}: ").strip()


# ────────────────────────────────────────────────────────
# بررسی وابستگی‌ها
# ────────────────────────────────────────────────────────
def check_dependencies():
    print("\n  ─── بررسی پکیج‌های لازم ───")
    required = {"openpyxl": "openpyxl"}
    missing = []
    for module_name, pip_name in required.items():
        try:
            __import__(module_name)
        except ImportError:
            missing.append(pip_name)

    if not missing:
        print("  ✅ همه‌ی پکیج‌های لازم نصب هستند")
        return

    print(f"  ⚠️ این پکیج‌ها نصب نیستند: {', '.join(missing)}")
    if ask_yes_no("  الان نصب کنم؟", True):
        try:
            subprocess.run([sys.executable, "-m", "pip", "install"] + missing, check=True)
            print("  ✅ نصب شد")
        except subprocess.CalledProcessError:
            print("  ❌ نصب خودکار شکست خورد — دستی نصب کن: pip install -r requirements.txt")
    else:
        print("  بعداً دستی نصب کن: pip install -r requirements.txt")


# ────────────────────────────────────────────────────────
# متقاضی‌های موجود (برای جلوگیری از پاک‌شدن دیتای قبلی)
# ────────────────────────────────────────────────────────
def load_existing_applicants():
    if not os.path.exists(CONFIG_PATH):
        return []
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f).get("applicants", [])
    except Exception:
        return []


def make_keywords(profession, name, name_fa):
    words = []
    if profession:
        words.extend(profession.lower().split())
    if name:
        words.extend(name.lower().split())
    if name_fa:
        words.append(name_fa.lower())
    # حذف تکراری‌ها و کلمات خیلی کوتاه/بی‌معنی، با حفظ ترتیب
    seen = set()
    cleaned = []
    for w in words:
        w = w.strip(",.")
        if not w or len(w) < 2 or w in STOPWORDS or w in seen:
            continue
        seen.add(w)
        cleaned.append(w)
    return cleaned


def unique_id(base_id, existing_ids):
    """اگر id تکراری بود (مثلاً دو نفر هم‌نام)، عدد اضافه می‌کند."""
    if base_id not in existing_ids:
        return base_id
    i = 2
    while f"{base_id}{i}" in existing_ids:
        i += 1
    return f"{base_id}{i}"


def collect_applicants(existing):
    existing_ids = {a["id"] for a in existing}
    applicants = list(existing)

    print("\n  ─── تنظیمات کلی ───")
    num_applicants = ask_int("چند متقاضی جدید می‌خواهید اضافه کنید؟", 1)

    for i in range(1, num_applicants + 1):
        print(f"\n  ─── متقاضی جدید {i} ───")

        name = ask("نام کامل (English)")
        name_fa = ask("نام فارسی")
        profession = ask("حرفه (مثلاً English Teacher, Midwife, IT Manager)")

        gender = ask("جنسیت (male/female)", "female").lower()
        if gender not in ("male", "female"):
            gender = "female"
        emoji = "👨" if gender == "male" else "👩"

        english = ask("سطح انگلیسی", "A2")
        german = ask("سطح آلمانی", "A1")

        print(f"\n  📧 ایمیل‌های {name_fa or name}:")
        num_emails = ask_int("  چند ایمیل می‌خواهید ثبت کنید؟", 1)
        emails = []
        for e in range(1, num_emails + 1):
            email = ask(f"  ایمیل {e} (Gmail)")
            if email:
                emails.append(email)

        print(f"\n  🔗 لینکدین {name_fa or name}:")
        num_linkedins = ask_int("  چند پروفایل لینکدین دارید؟", 1)
        linkedins = []
        for l in range(1, num_linkedins + 1):
            li = ask(f"  لینکدین {l} (URL کامل)")
            if li:
                linkedins.append(li)

        if not (name or emails):
            print("  ⏭️ خالی ماند\n")
            continue

        base_id = (name or name_fa or f"person{i}").lower().replace(" ", "_")
        # فقط حروف/عدد لاتین — چون این id توی .env هم به‌عنوان کلید محیطی استفاده می‌شود
        # و کاراکترهای فارسی/غیرلاتین در اسم متغیر محیطی روی بعضی سیستم‌ها مشکل‌ساز است
        base_id = "".join(c for c in base_id if (c.isascii() and c.isalnum()) or c == "_")
        if not any(c.isalnum() for c in base_id):
            base_id = f"person{i}"
        app_id = unique_id(base_id, existing_ids)
        existing_ids.add(app_id)

        applicants.append({
            "id": app_id,
            "name": name,
            "name_fa": name_fa,
            "gender": gender,
            "emoji": emoji,
            "profession": profession,
            "keywords": make_keywords(profession, name, name_fa),
            "emails": emails,
            "linkedins": linkedins,
            "english": english,
            "german": german,
        })
        print(f"  ✅ ذخیره شد: {name_fa or app_id} ({len(emails)} ایمیل، {len(linkedins)} لینکدین)\n")

    return applicants


def write_config(applicants):
    config = {
        "project": "MigrationHunter",
        "version": "3.0",
        "created": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "applicants": applicants,
    }
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    try:
        os.chmod(CONFIG_PATH, 0o600)
    except OSError:
        pass
    print("  ✅ ذخیره شد: config.json")


def collect_passwords(applicants, newly_added_ids):
    """
    فقط برای متقاضی‌های تازه‌اضافه‌شده (یا اگر کاربر بخواهد، همه) رمز می‌پرسد،
    تا رمزهای قبلاً واردشده در .env دست‌نخورده بمانند.
    """
    print("\n  ─── رمزهای App Password ───")
    print("  (راهنما: https://myaccount.google.com/apppasswords)")
    print("  رمز روی صفحه نمایش داده نمی‌شود.\n")

    targets = [a for a in applicants if a["id"] in newly_added_ids]
    if not targets and applicants:
        if ask_yes_no("  متقاضی جدیدی نیست — می‌خوای رمز متقاضی‌های قبلی رو هم دوباره وارد/عوض کنی؟", False):
            targets = applicants

    env_lines = []
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            env_lines = [l.rstrip("\n") for l in f]
    else:
        env_lines = [
            "# MigrationHunter — Passwords",
            f"# Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "",
        ]

    def upsert_env_line(key, value):
        prefix = f"{key}="
        for idx, line in enumerate(env_lines):
            if line.startswith(prefix):
                env_lines[idx] = f"{key}={value}"
                return
        env_lines.append(f"{key}={value}")

    for a in targets:
        app_id = a["id"].upper()
        print(f"  {a['emoji']} {a.get('name_fa', a['id'])} — {len(a.get('emails', []))} ایمیل")
        for idx, email in enumerate(a.get("emails", []), 1):
            want = ask_yes_no(f"    رمز {email} را الان وارد کنم؟", True)
            pw = ask_secret(f"App Password برای {email}") if want else "REPLACE_WITH_APP_PASSWORD"
            upsert_env_line(f"EMAIL_{app_id}_{idx}", email)
            upsert_env_line(f"EMAIL_PASSWORD_{app_id}_{idx}", pw)
            upsert_env_line(f"EMAIL_PROVIDER_{app_id}_{idx}", "gmail")

    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(env_lines) + "\n")
    try:
        os.chmod(ENV_PATH, 0o600)  # فقط خودت بتوانی بخوانی/بنویسی
    except OSError:
        pass
    print("\n  ✅ ذخیره شد: .env")


def setup_ai_key():
    print("\n  ─── (اختیاری) کلید هوش‌مصنوعی برای کاور لتر/ایمیل خودکار ───")
    if not ask_yes_no("  می‌خوای الان یک کلید AI (OpenAI یا Gemini) تنظیم کنی؟", False):
        print("  رد شد — بعداً می‌تونی از تب «⚙️ تنظیمات» توی web_ui.py یا مستقیم توی .env اضافه کنی.")
        return

    provider = ask("  کدوم؟ (openai / gemini)", "openai").strip().lower()
    if provider not in ("openai", "gemini"):
        provider = "openai"
    key = ask_secret(f"{provider.upper()} API Key")
    if not key:
        print("  چیزی وارد نشد — رد شد.")
        return

    env_lines = []
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            env_lines = [l.rstrip("\n") for l in f]

    def upsert(key_name, value):
        prefix = f"{key_name}="
        for idx, line in enumerate(env_lines):
            if line.startswith(prefix):
                env_lines[idx] = f"{key_name}={value}"
                return
        env_lines.append(f"{key_name}={value}")

    upsert("AI_PROVIDER", provider)
    upsert("AI_API_KEY", key)
    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(env_lines) + "\n")
    try:
        os.chmod(ENV_PATH, 0o600)
    except OSError:
        pass
    print("  ✅ کلید AI ذخیره شد")


def init_sources():
    """اگر sources.json نبود، با لیست پیش‌فرض job_crawler.py ساخته می‌شود."""
    sources_path = os.path.join(BASE, "sources.json")
    if os.path.exists(sources_path):
        return
    try:
        sys.path.insert(0, BASE)
        import job_crawler
        job_crawler.load_sources()  # خودش می‌سازد اگر نبود
        print("  ✅ ساخته شد: sources.json (۱۸ منبع پیش‌فرض — از تب تنظیمات قابل مدیریت است)")
    except Exception:
        pass  # مهم نیست — job_crawler.py موقع اجرای واقعی خودش می‌سازدش


def check_gitignore_safety():
    print("\n  ─── بررسی امنیتی .gitignore ───")
    if not os.path.exists(GITIGNORE_PATH):
        print("  ⚠️ فایل .gitignore پیدا نشد! اگر این پوشه را git init کرده‌ای،")
        print("     .env و config.json ممکن است به‌اشتباه commit شوند.")
        return
    with open(GITIGNORE_PATH, "r", encoding="utf-8") as f:
        patterns = {line.strip() for line in f if line.strip() and not line.strip().startswith("#")}
    missing = [p for p in (".env", "config.json") if p not in patterns]
    if missing:
        print(f"  ⚠️ این فایل‌ها در .gitignore نیستند: {', '.join(missing)}")
        print("     قبل از هر git commit حتماً اضافه‌شون کن، وگرنه رمزها/اطلاعات شخصی لو می‌رود.")
    else:
        print("  ✅ .env و config.json در .gitignore هستند")


def main():
    banner()
    check_dependencies()

    existing = load_existing_applicants()
    if existing:
        print(f"\n  ℹ️ {len(existing)} متقاضی از قبل در config.json ثبت شده:")
        for a in existing:
            print(f"     {a.get('emoji','👤')} {a.get('name_fa') or a.get('name') or a['id']}")
        print("\n  متقاضی‌های جدیدی که الان وارد کنی به همین لیست اضافه می‌شوند")
        print("  (چیزی از قبلی‌ها پاک/بازنویسی نمی‌شود).")

    before_ids = {a["id"] for a in existing}
    applicants = collect_applicants(existing)

    if not applicants:
        print("\n  ❌ هیچ متقاضی‌ای (نه قبلی، نه جدید) ثبت نشد — چیزی ذخیره نشد.")
        return

    newly_added_ids = {a["id"] for a in applicants} - before_ids
    write_config(applicants)
    collect_passwords(applicants, newly_added_ids)
    setup_ai_key()

    for d in ["memory", "profiles", "output", "dashboard", "dashboard/archive", "input"]:
        os.makedirs(os.path.join(BASE, d), exist_ok=True)

    init_sources()
    check_gitignore_safety()

    print("\n" + "=" * 60)
    print("  راه‌اندازی کامل شد!")
    print("=" * 60)
    print(f"\n  فایل‌های آماده:")
    print(f"    config.json    ← {len(applicants)} متقاضی")
    print(f"    .env           ← رمزهای ایمیل (و کلید AI در صورت تنظیم)")
    print(f"    sources.json   ← منابع جستجوی آگهی")
    print(f"\n  برای شروع (پیشنهادی — رابط وب با دکمه و گزارش):")
    print(f"    python web_ui.py")
    print(f"    بعد مرورگر را باز کن: http://127.0.0.1:8877")
    print(f"\n  یا خط‌فرمان مستقیم:")
    print(f"    python run.py")
    print(f"\n  برای تست اتصال ایمیل (بدون آنالیز کامل):")
    print(f"    python email_analyzer.py --dry-run")
    print("=" * 60)


if __name__ == "__main__":
    main()
