#!/usr/bin/env python3
"""
MigrationHunter — اسکریپت اصلی اجرا
یک دستور همه چیز را اجرا می‌کند

اجرا: python run.py
"""
import os, sys, subprocess
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE)

NOW = datetime.now()
DATE_STR = NOW.strftime("%Y-%m-%d %H:%M")

def run_step(name, script, args=None):
    """اجرای یک مرحله"""
    print(f"\n{'─'*50}")
    print(f"▶ {name}")
    print(f"{'─'*50}")
    
    cmd = [sys.executable, script]
    if args:
        cmd.extend(args)
    
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        print(result.stdout)
        if result.stderr:
            # Only show real errors, not warnings
            for line in result.stderr.split("\n"):
                if "Error" in line or "error" in line or "Traceback" in line:
                    print(f"  ⚠️ {line}")
        return result.returncode == 0
    except subprocess.TimeoutExpired:
        print(f"  ⏰ Timeout — {name}")
        return False
    except Exception as e:
        print(f"  ❌ خطا: {e}")
        return False

def main():
    print("═" * 60)
    print("  MigrationHunter — اجرای کامل")
    print(f"  📅 {DATE_STR}")
    print("═" * 60)
    
    steps = [
        ("تحلیل ایمیل شغلی", "email_analyzer.py"),
        ("ساخت Excel ایمیل", "email_dashboard.py"),
        ("جستجوی خودکار کار", "job_crawler.py"),
        ("یادآوری پیگیری", "followup_reminder.py"),
        ("ساخت داشبورد اصلی", "build_dashboard.py"),
    ]
    
    results = {}
    for name, script in steps:
        if not os.path.exists(script):
            print(f"\n⚠️ {script} پیدا نشد — رد شد")
            results[name] = False
            continue
        
        success = run_step(name, script)
        results[name] = success
    
    # Summary
    print("\n" + "═" * 60)
    print("  📊 خلاصه نهایی")
    print("═" * 60)
    
    for name, success in results.items():
        status = "✅" if success else "❌"
        print(f"  {status} {name}")
    
    # Show latest Excel
    dash_dir = os.path.join(BASE, "dashboard")
    if os.path.isdir(dash_dir):
        xlsx_files = [f for f in os.listdir(dash_dir) if f.endswith(".xlsx")]
        if xlsx_files:
            latest = sorted(xlsx_files)[-1]
            print(f"\n  📊 آخرین Excel: dashboard/{latest}")
    
    print("\n" + "═" * 60)
    print("  ✅ اجرا تمام شد")
    print("═" * 60)

if __name__ == "__main__":
    main()
