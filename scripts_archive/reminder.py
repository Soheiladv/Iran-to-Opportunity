#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Migration Hunter — Reminder System
سیستم یادآوری خودکار پیگیری ایمیل‌ها

نحوه اجرا:
    python reminder.py                    # بررسی یادآوری‌ها
    python reminder.py --add              # اضافه کردن یادآوری
    python reminder.py --list             # لیست یادآوری‌ها
    python reminder.py --complete ID      # تکمیل یادآوری
    python reminder.py --export           # خروجی به فایل
"""

import os
import json
from datetime import datetime, timedelta
from pathlib import Path

BASE_DIR = Path(__file__).parent
REMINDER_FILE = BASE_DIR / "memory" / "REMINDERS.json"
TRACKER_FILE = BASE_DIR / "memory" / "EMAIL_TRACKER.json"

# ==========================================
# REMINDER SYSTEM
# ==========================================

class ReminderSystem:
    """سیستم یادآوری خودکار"""
    
    def __init__(self):
        self.reminders = self._load_reminders()
        self.emails = self._load_emails()
    
    def _load_reminders(self):
        """بارگذاری یادآوری‌ها"""
        if REMINDER_FILE.exists():
            with open(REMINDER_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {"reminders": [], "templates": []}
    
    def _load_emails(self):
        """بارگذاری ایمیل‌ها"""
        if TRACKER_FILE.exists():
            with open(TRACKER_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {"emails": []}
    
    def _save(self):
        """ذخیره یادآوری‌ها"""
        REMINDER_FILE.parent.mkdir(exist_ok=True)
        with open(REMINDER_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.reminders, f, ensure_ascii=False, indent=2)
    
    def add_reminder(self, title, due_date, reminder_type="follow_up", priority="medium", notes="", email_id=None):
        """اضافه کردن یادآوری"""
        now = datetime.now()
        reminder = {
            "id": f"REM-{len(self.reminders['reminders'])+1:03d}",
            "title": title,
            "type": reminder_type,  # follow_up, application, interview, deadline
            "priority": priority,   # low, medium, high, urgent
            "due_date": due_date,
            "due_time": "09:00:00",
            "status": "pending",    # pending, completed, cancelled
            "email_id": email_id,
            "notes": notes,
            "created_at": now.isoformat(),
            "completed_at": None
        }
        
        self.reminders["reminders"].append(reminder)
        self._save()
        
        print(f"✅ یادآوری ثبت شد: {reminder['id']}")
        print(f"   عنوان: {title}")
        print(f"   تاریخ سررسید: {due_date}")
        print(f"   اولویت: {priority}")
        
        return reminder
    
    def check_due_reminders(self):
        """بررسی یادآوری‌های سررسید شده"""
        today = datetime.now().strftime("%Y-%m-%d")
        due = []
        
        for reminder in self.reminders["reminders"]:
            if reminder["status"] == "pending" and reminder["due_date"] <= today:
                due.append(reminder)
        
        return due
    
    def check_upcoming_reminders(self, days=3):
        """بررسی یادآوری‌های آینده"""
        today = datetime.now()
        upcoming = []
        
        for reminder in self.reminders["reminders"]:
            if reminder["status"] == "pending":
                due = datetime.strptime(reminder["due_date"], "%Y-%m-%d")
                diff = (due - today).days
                if 0 <= diff <= days:
                    upcoming.append(reminder)
        
        return upcoming
    
    def complete_reminder(self, reminder_id):
        """تکمیل یادآوری"""
        for reminder in self.reminders["reminders"]:
            if reminder["id"] == reminder_id:
                reminder["status"] = "completed"
                reminder["completed_at"] = datetime.now().isoformat()
                self._save()
                print(f"✅ یادآوری تکمیل شد: {reminder_id}")
                return reminder
        
        print(f"❌ یادآوری یافت نشد: {reminder_id}")
        return None
    
    def cancel_reminder(self, reminder_id):
        """لغو یادآوری"""
        for reminder in self.reminders["reminders"]:
            if reminder["id"] == reminder_id:
                reminder["status"] = "cancelled"
                self._save()
                print(f"❌ یادآوری لغو شد: {reminder_id}")
                return reminder
        
        print(f"❌ یادآوری یافت نشد: {reminder_id}")
        return None
    
    def auto_create_follow_ups(self):
        """ایجاد خودکار یادآوری پیگیری برای ایمیل‌ها"""
        created = 0
        
        for email in self.emails.get("emails", []):
            # بررسی آیا یادآوری قبلاً ایجاد شده
            exists = any(r.get("email_id") == email["id"] for r in self.reminders["reminders"])
            
            if not exists and email["status"] == "sent":
                # ایجاد یادآوری پیگیری ۷ روزه
                follow_up_date = email.get("follow_up_date")
                if follow_up_date:
                    self.add_reminder(
                        title=f"پیگیری ایمیل به {email['employer']}",
                        due_date=follow_up_date,
                        reminder_type="follow_up",
                        priority="high",
                        notes=f"ایمیل {email['id']} ارسال شده در {email['sent_date']}",
                        email_id=email["id"]
                    )
                    created += 1
        
        if created > 0:
            print(f"\n✅ {created} یادآوری خودکار ایجاد شد")
        else:
            print("\n✅ هیچ یادآوری جدیدی نیاز نیست")
        
        return created
    
    def generate_report(self):
        """تولید گزارش یادآوری‌ها"""
        now = datetime.now()
        today = now.strftime("%Y-%m-%d")
        
        due = self.check_due_reminders()
        upcoming = self.check_upcoming_reminders(7)
        completed = [r for r in self.reminders["reminders"] if r["status"] == "completed"]
        pending = [r for r in self.reminders["reminders"] if r["status"] == "pending"]
        
        report = [
            f"# گزارش یادآوری‌ها — {today} {now.strftime('%H:%M')}",
            "",
            "---",
            "",
        ]
        
        # یادآوری‌های سررسید شده
        if due:
            report.append("## ⚠️ یادآوری‌های سررسید شده (نیاز به اقدام فوری)")
            report.append("")
            for r in due:
                report.append(f"### {r['id']}: {r['title']}")
                report.append(f"| فیلد | مقدار |")
                report.append(f"|------|-------|")
                report.append(f"| نوع | {r['type']} |")
                report.append(f"| اولویت | {r['priority']} |")
                report.append(f"| تاریخ سررسید | {r['due_date']} |")
                report.append(f"| یادداشت | {r['notes']} |")
                report.append(f"| وضعیت | ⚠️ نیاز به اقدام |")
                report.append("")
        else:
            report.append("## ✅ هیچ یادآوری سررسید شده‌ای نیست")
            report.append("")
        
        # یادآوری‌های آینده
        if upcoming:
            report.append("## 📅 یادآوری‌های آینده (۷ روز آینده)")
            report.append("")
            for r in upcoming:
                days_left = (datetime.strptime(r['due_date'], "%Y-%m-%d") - now).days
                report.append(f"- **{r['id']}**: {r['title']} — {r['due_date']} ({days_left} روز دیگر)")
            report.append("")
        
        # آمار کلی
        report.append("## 📊 آمار کلی")
        report.append("")
        report.append("| شاخص | مقدار |")
        report.append("|------|-------|")
        report.append(f"| کل یادآوری‌ها | {len(self.reminders['reminders'])} |")
        report.append(f"| در انتظار | {len(pending)} |")
        report.append(f"| سررسید شده | {len(due)} |")
        report.append(f"| تکمیل شده | {len(completed)} |")
        report.append(f"| آینده (۷ روز) | {len(upcoming)} |")
        report.append("")
        
        # لیست کامل
        report.append("## 📋 لیست کامل یادآوری‌ها")
        report.append("")
        report.append("| ID | عنوان | نوع | اولویت | تاریخ | وضعیت |")
        report.append("|-----|-------|-----|--------|-------|-------|")
        for r in self.reminders["reminders"]:
            status_emoji = "✅" if r["status"] == "completed" else "❌" if r["status"] == "cancelled" else "⏳"
            report.append(f"| {r['id']} | {r['title'][:30]} | {r['type']} | {r['priority']} | {r['due_date']} | {status_emoji} {r['status']} |")
        report.append("")
        
        content = "\n".join(report)
        
        # ذخیره
        output_file = BASE_DIR / "output" / "REMINDER_REPORT.md"
        output_file.write_text(content, encoding='utf-8')
        
        return content
    
    def list_reminders(self):
        """لیست تمام یادآوری‌ها"""
        return self.reminders["reminders"]
    
    def get_statistics(self):
        """آمار کلی"""
        reminders = self.reminders["reminders"]
        due = self.check_due_reminders()
        return {
            "total": len(reminders),
            "pending": len([r for r in reminders if r["status"] == "pending"]),
            "completed": len([r for r in reminders if r["status"] == "completed"]),
            "cancelled": len([r for r in reminders if r["status"] == "cancelled"]),
            "due_now": len(due),
            "urgent": len([r for r in reminders if r["priority"] == "urgent" and r["status"] == "pending"]),
            "high": len([r for r in reminders if r["priority"] == "high" and r["status"] == "pending"]),
        }

# ==========================================
# MAIN
# ==========================================

def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Reminder System')
    parser.add_argument('--add', action='store_true', help='Add reminder')
    parser.add_argument('--list', action='store_true', help='List all reminders')
    parser.add_argument('--check', action='store_true', help='Check due reminders')
    parser.add_argument('--complete', type=str, help='Complete reminder by ID')
    parser.add_argument('--cancel', type=str, help='Cancel reminder by ID')
    parser.add_argument('--auto', action='store_true', help='Auto create follow-ups')
    parser.add_argument('--export', action='store_true', help='Export report')
    parser.add_argument('--stats', action='store_true', help='Show statistics')
    
    # For --add
    parser.add_argument('--title', type=str, help='Reminder title')
    parser.add_argument('--date', type=str, help='Due date (YYYY-MM-DD)')
    parser.add_argument('--type', type=str, default='follow_up', help='Reminder type')
    parser.add_argument('--priority', type=str, default='medium', help='Priority')
    parser.add_argument('--notes', type=str, default='', help='Notes')
    
    args = parser.parse_args()
    
    system = ReminderSystem()
    
    if args.add:
        if not all([args.title, args.date]):
            print("❌ لطفاً عنوان و تاریخ را مشخص کنید")
            print("   --title 'عنوان' --date '2026-08-26'")
            return
        system.add_reminder(args.title, args.date, args.type, args.priority, args.notes)
    
    elif args.list:
        reminders = system.list_reminders()
        if reminders:
            print("\n📋 لیست یادآوری‌ها:")
            for r in reminders:
                status = "✅" if r["status"] == "completed" else "❌" if r["status"] == "cancelled" else "⏳"
                print(f"  {status} {r['id']}: {r['title']} ({r['due_date']})")
        else:
            print("هیچ یادآوری ثبت نشده")
    
    elif args.check:
        due = system.check_due_reminders()
        upcoming = system.check_upcoming_reminders(3)
        
        if due:
            print(f"\n⚠️ {len(due)} یادآوری سررسید شده:")
            for r in due:
                print(f"  {r['id']}: {r['title']} ({r['due_date']})")
        else:
            print("\n✅ هیچ یادآوری سررسید شده‌ای نیست")
        
        if upcoming:
            print(f"\n📅 {len(upcoming)} یادآوری در ۳ روز آینده:")
            for r in upcoming:
                print(f"  {r['id']}: {r['title']} ({r['due_date']})")
    
    elif args.complete:
        system.complete_reminder(args.complete)
    
    elif args.cancel:
        system.cancel_reminder(args.cancel)
    
    elif args.auto:
        system.auto_create_follow_ups()
    
    elif args.export:
        report = system.generate_report()
        print(report)
    
    elif args.stats:
        stats = system.get_statistics()
        print("\n📊 آمار یادآوری‌ها:")
        print(f"   کل: {stats['total']}")
        print(f"   در انتظار: {stats['pending']}")
        print(f"   تکمیل شده: {stats['completed']}")
        print(f"   لغو شده: {stats['cancelled']}")
        print(f"   سررسید شده: {stats['due_now']}")
        print(f"   فوری: {stats['urgent']}")
        print(f"   مهم: {stats['high']}")
    
    else:
        # نمایش وضعیت
        due = system.check_due_reminders()
        upcoming = system.check_upcoming_reminders(7)
        stats = system.get_statistics()
        
        print("\n" + "=" * 60)
        print("🔔 Reminder System — سیستم یادآوری")
        print("=" * 60)
        
        if due:
            print(f"\n⚠️ {len(due)} یادآوری سررسید شده:")
            for r in due:
                print(f"  {r['id']}: {r['title']}")
        
        if upcoming:
            print(f"\n📅 {len(upcoming)} یادآوری آینده:")
            for r in upcoming:
                print(f"  {r['id']}: {r['title']} ({r['due_date']})")
        
        print(f"\n📊 آمار: {stats['total']} کل | {stats['pending']} در انتظار | {stats['due_now']} سررسید")
        print("=" * 60)

if __name__ == "__main__":
    main()
