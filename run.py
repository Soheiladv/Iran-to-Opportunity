#!/usr/bin/env python3
"""
MigrationHunter — Pipeline Runner with Progress Bar
"""
import os, sys, subprocess, io
from datetime import datetime

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

NOW = datetime.now()
DATE_STR = NOW.strftime("%Y-%m-%d %H:%M")

TOTAL_STEPS = 5
BAR_WIDTH = 30

STEP_INFO = [
    {"key": "email_analyze", "name": "تحلیل ایمیل شغلی",  "script": "email_analyzer.py",      "emoji": "📧"},
    {"key": "email_excel",   "name": "ساخت Excel ایمیل",  "script": "email_dashboard.py",     "emoji": "📊"},
    {"key": "job_search",    "name": "جستجوی خودکار کار", "script": "job_crawler.py",         "emoji": "🔍"},
    {"key": "followup",      "name": "یادآوری پیگیری",    "script": "followup_reminder.py",   "emoji": "⏰"},
    {"key": "dashboard",     "name": "ساخت داشبورد اصلی", "script": "build_dashboard.py",     "emoji": "📈"},
]


def progress_bar(current, total, width=BAR_WIDTH):
    pct = current / total if total else 0
    filled = int(width * pct)
    bar = "█" * filled + "░" * (width - filled)
    return f"  [{bar}] {current}/{total} ({int(pct*100)}%)"


def print_header():
    print()
    print("  " + "=" * 56)
    print("  MigrationHunter — Pipeline Runner")
    print(f"  {DATE_STR}")
    print("  " + "=" * 56)
    print()


def print_status_line(results, running_idx=None):
    for i, step in enumerate(STEP_INFO):
        key = step["key"]
        if key in results:
            icon = "  ✅" if results[key] else "  ❌"
        elif running_idx == i:
            icon = "  🔄"
        else:
            icon = "  ⬜"
        print(f"  {icon} {step['emoji']} {step['name']}")
    print()


def run_step(step_info, step_num):
    name = step_info["name"]
    script = step_info["script"]
    emoji = step_info["emoji"]

    if not os.path.exists(script):
        print(f"  ⏭️  {emoji} {script} پیدا نشد — رد شد")
        return False

    print(f"  ▶ مرحله {step_num}/{TOTAL_STEPS}: {emoji} {name}")
    print(f"  {'─' * 50}")

    cmd = [sys.executable, script]
    try:
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=600,
            encoding='utf-8', errors='replace', env=env
        )
        if result.stdout:
            for line in result.stdout.strip().split("\n"):
                print(f"    {line}")

        if result.stderr:
            for line in result.stderr.split("\n"):
                if any(w in line for w in ["Error", "error", "Traceback", "NameError", "SyntaxError"]):
                    print(f"    ⚠️  {line}")

        ok = result.returncode == 0
        if ok:
            print(f"  ✅ {name} — انجام شد\n")
        else:
            print(f"  ❌ {name} — خطا (code {result.returncode})\n")
        return ok

    except subprocess.TimeoutExpired:
        print(f"  ⏰ {name} — timeout\n")
        return False
    except Exception as e:
        print(f"  ❌ {name} — {e}\n")
        return False


def main():
    config_path = os.path.join(BASE, "config.json")
    if not os.path.exists(config_path):
        print_header()
        print("  ⚠️  config.json پیدا نشد!")
        print("  ابتدا اجرا کنید: python setup.py")
        print("  سپس دوباره اجرا کنید: python run.py")
        return

    results = {}
    print_header()
    print_status_line(results, running_idx=0)

    for i, step in enumerate(STEP_INFO):
        # Show which step is running
        for j, s in enumerate(STEP_INFO):
            key = s["key"]
            if key in results:
                icon = "  ✅" if results[key] else "  ❌"
            elif j == i:
                icon = "  ▶️ "
            else:
                icon = "  ⬜"
            print(f"  {icon} {s['emoji']} {s['name']}")
        print()

        # Progress bar
        completed = sum(1 for v in results.values() if v is not None)
        print(progress_bar(completed, TOTAL_STEPS))
        print()

        # Run step
        success = run_step(step, i + 1)
        results[step["key"]] = success

    # Final
    completed = sum(1 for v in results.values() if v)
    failed = sum(1 for v in results.values() if not v)

    print("=" * 56)
    print("  📊 نتیجه نهایی")
    print("=" * 56)
    for step in STEP_INFO:
        ok = results.get(step["key"], False)
        icon = "✅" if ok else "❌"
        print(f"  {icon} {step['emoji']} {step['name']}")
    print()
    print(f"  {completed} موفق | {failed} خطا از {TOTAL_STEPS}")
    print(progress_bar(completed + failed, TOTAL_STEPS))
    print()

    # Latest Excel
    dash_dir = os.path.join(BASE, "dashboard")
    if os.path.isdir(dash_dir):
        xlsx_files = [f for f in os.listdir(dash_dir) if f.endswith(".xlsx") and "~$" not in f]
        if xlsx_files:
            latest = sorted(xlsx_files)[-1]
            print(f"  📊 آخرین Excel: dashboard/{latest}")
    print()


if __name__ == "__main__":
    main()