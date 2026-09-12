#!/usr/bin/env python3
"""
MigrationHunter — Application Tracker (بانک درخواست‌ها)
========================================================
Single source of truth for real job applications:
  - هر ایمیل ارسال‌شده با کارفرما، تاریخ، و مهلت پاسخ منتظره ثبت می‌شود
  - چرخه وضعیت: DRAFT → SENT → FOLLOW_UP → REPLIED → INTERVIEW → OFFER
  - خروجی: memory/APPLICATION_BANK.md (گزارش فارسی)

اجرا:
  python application_tracker.py                     → گزارش کامل + هشدار مهلت‌ها
  python application_tracker.py --add               → ثبت درخواست جدید (تعاملی)
  python application_tracker.py --status MH-001 replied --date 2026-09-09
                                                    → تغییر وضعیت یک رکورد
  python application_tracker.py --seed              → درون‌ریزی از EMAIL_TRACKER.json
"""
import os, sys, io, json, argparse
from datetime import datetime, timedelta

# Fix Windows console encoding for emoji support
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    os.environ["PYTHONIOENCODING"] = "utf-8"

BASE = os.path.dirname(os.path.abspath(__file__))
MEM = os.path.join(BASE, "memory")
BANK_JSON = os.path.join(MEM, "APPLICATION_BANK.json")
BANK_MD = os.path.join(MEM, "APPLICATION_BANK.md")
TRACKER_JSON = os.path.join(MEM, "EMAIL_TRACKER.json")

NOW = datetime.now()
DATE_STR = NOW.strftime("%Y-%m-%d %H:%M")
TODAY = NOW.date()

# ═══════════════════════════════════════════════════
# چرخه وضعیت و قوانین پیگیری
# ═══════════════════════════════════════════════════
# هر وضعیت: (وضعیت بعدی، روزهای مهلت پاسخ، برچسب فارسی)
STATUS_FLOW = {
    "DRAFT":     (None,        None, "📝 پیش‌نویس"),
    "SENT":      ("FOLLOW_UP", 7,    "📤 ارسال شده"),
    "FOLLOW_UP": ("REPLIED",   7,    "🔁 پیگیری شده"),
    "REPLIED":   ("INTERVIEW", 5,    "📬 پاسخ دریافت شد"),
    "INTERVIEW": ("OFFER",     10,   "🎤 مصاحبه"),
    "OFFER":     (None,        None, "🎉 پیشنهاد شغلی"),
    "REJECTED":  (None,        None, "❌ رد شده"),
    "WITHDRAWN": (None,        None, "🚪 انصراف"),
}

# رنگ مهلت پاسخ
DEADLINE_URGENT = 3      # ≤ ۳ روز مانده: قرمز
DEADLINE_SOON = 7        # ≤ ۷ روز مانده: زرد


def load_bank():
    if not os.path.exists(BANK_JSON):
        return {"applications": []}
    with open(BANK_JSON, "r", encoding="utf-8") as f:
        return json.load(f)


def save_bank(bank):
    os.makedirs(MEM, exist_ok=True)
    with open(BANK_JSON, "w", encoding="utf-8") as f:
        json.dump(bank, f, ensure_ascii=False, indent=2)


def next_id(bank):
    ids = [a["id"] for a in bank["applications"]]
    nums = []
    for i in ids:
        try:
            nums.append(int(i.split("-")[-1]))
        except ValueError:
            pass
    n = (max(nums) + 1) if nums else 1
    return f"MH-{NOW.year}-{n:03d}"


def reply_deadline(sent_date_str, status):
    """مهلت پاسخ: ارسال → ۷ روز، پیگیری → ۷ روز بعد از پیگیری."""
    days = STATUS_FLOW.get(status, (None, None, ""))[1]
    if not days:
        return None
    try:
        d = datetime.strptime(sent_date_str[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None
    return (d + timedelta(days=days)).isoformat()


def status_fa(status):
    return STATUS_FLOW.get(status, (None, None, status))[2]


# ═══════════════════════════════════════════════════
# ثبت درخواست جدید
# ═══════════════════════════════════════════════════
def add_application(bank, applicant, employer, job, country, email_to, subject,
                    sent_date, status="SENT", notes=""):
    app = {
        "id": next_id(bank),
        "applicant": applicant,
        "employer": employer,
        "job": job,
        "country": country,
        "email_to": email_to,
        "subject": subject,
        "sent_date": sent_date,
        "status": status.upper(),
        "reply_deadline": reply_deadline(sent_date, status.upper()),
        "reply_date": None,
        "reply_summary": "",
        "interview_date": None,
        "offer_date": None,
        "notes": notes,
        "created_at": NOW.isoformat(timespec="seconds"),
        "updated_at": NOW.isoformat(timespec="seconds"),
    }
    bank["applications"].append(app)
    return app


def record_followup(bank, app_id, date_str=None, note=""):
    """
    ثبت ارسال ایمیل پیگیری: وضعیت → FOLLOW_UP و مهلت پاسخ جدید ۷ روزه.
    تاریخ مبنای مهلت، تاریخ خودِ پیگیری است (نه ارسال اولیه).
    """
    for a in bank["applications"]:
        if a["id"].lower() == app_id.lower():
            fu_date = date_str or TODAY.isoformat()
            old = a["status"]
            a["status"] = "FOLLOW_UP"
            a["followup_date"] = fu_date
            a["followup_count"] = int(a.get("followup_count") or 0) + 1
            # مهلت جدید: ۷ روز از تاریخ پیگیری
            a["reply_deadline"] = reply_deadline(fu_date, "FOLLOW_UP")
            a["updated_at"] = NOW.isoformat(timespec="seconds")
            if note:
                prev = a.get("notes") or ""
                a["notes"] = (prev + " | " if prev else "") + f"پیگیری #{a['followup_count']} در {fu_date}: {note}"
            return a, old
    return None, None


def set_status(bank, app_id, new_status, date_str=None, reply_summary=""):
    for a in bank["applications"]:
        if a["id"].lower() == app_id.lower():
            old = a["status"]
            a["status"] = new_status.upper()
            a["updated_at"] = NOW.isoformat(timespec="seconds")
            if new_status.upper() == "REPLIED":
                a["reply_date"] = date_str or TODAY.isoformat()
                if reply_summary:
                    a["reply_summary"] = reply_summary
            elif new_status.upper() == "INTERVIEW":
                a["interview_date"] = date_str or ""
            elif new_status.upper() == "OFFER":
                a["offer_date"] = date_str or TODAY.isoformat()
            # مهلت جدید بر اساس وضعیت جدید (از تاریخ آخرین رویداد)
            base_date = date_str or a.get("sent_date", TODAY.isoformat())
            a["reply_deadline"] = reply_deadline(base_date, new_status.upper())
            return a, old
    return None, None


# ═══════════════════════════════════════════════════
# درون‌ریزی از EMAIL_TRACKER.json (داده‌های قدیمی)
# ═══════════════════════════════════════════════════
def seed_from_tracker(bank):
    if not os.path.exists(TRACKER_JSON):
        print("⚠️ EMAIL_TRACKER.json یافت نشد")
        return bank, 0
    with open(TRACKER_JSON, "r", encoding="utf-8") as f:
        old = json.load(f)
    existing = {(a["employer"], a["sent_date"]) for a in bank["applications"]}
    added = 0
    for e in old.get("emails", []):
        key = (e.get("employer", ""), e.get("sent_date", ""))
        if key in existing:
            continue
        status = "REPLIED" if e.get("status") == "replied" else "SENT"
        app = add_application(
            bank,
            applicant=e.get("applicant", ""),
            employer=e.get("employer", ""),
            job=e.get("subject", ""),
            country=(e.get("country") or "").upper(),
            email_to=e.get("recipient_email", ""),
            subject=e.get("subject", ""),
            sent_date=e.get("sent_date", TODAY.isoformat()),
            status=status,
            notes=e.get("notes", ""),
        )
        if status == "REPLIED":
            app["reply_date"] = e.get("reply_date") or e.get("sent_date")
        existing.add(key)
        added += 1
    return bank, added


# ═══════════════════════════════════════════════════
# گزارش فارسی → memory/APPLICATION_BANK.md
# ═══════════════════════════════════════════════════
def build_report(bank):
    apps = bank["applications"]
    L = []
    L.append("# 📋 بانک درخواست‌های واقعی — APPLICATION BANK")
    L.append("")
    L.append(f"**آخرین بروزرسانی:** {DATE_STR}")
    L.append(f"**تعداد کل درخواست‌ها:** {len(apps)}")
    L.append("")

    # ── هشدار مهلت‌ها ──
    alerts = []
    for a in apps:
        if a["status"] not in ("SENT", "FOLLOW_UP"):
            continue
        dl = a.get("reply_deadline")
        if not dl:
            continue
        try:
            remain = (datetime.strptime(dl, "%Y-%m-%d").date() - TODAY).days
        except ValueError:
            continue
        if remain < 0:
            alerts.append((a, remain, "⏰ مهلت گذشته → پیگیری فوری (FOLLOW_UP)"))
        elif remain <= DEADLINE_URGENT:
            alerts.append((a, remain, "🔴 فوری"))
        elif remain <= DEADLINE_SOON:
            alerts.append((a, remain, "🟡 نزدیک"))

    if alerts:
        L.append("## ⚠️ هشدار مهلت پاسخ")
        L.append("")
        L.append("| # | ID | کارفرما | وضعیت | مهلت | روز مانده | اقدام |")
        L.append("|---|----|---------|-------|------|-----------|-------|")
        for i, (a, remain, label) in enumerate(sorted(alerts, key=lambda x: x[1]), 1):
            action = "پیگیری فوری" if remain < 0 else ("بازبینی امروز" if remain <= DEADLINE_URGENT else "در انتظار")
            L.append(f"| {i} | {a['id']} | {a['employer'][:28]} | {status_fa(a['status'])} | {dl} | {remain} | {label} → {action} |")
        L.append("")

    # ── جدول اصلی ──
    L.append("## 📊 همه درخواست‌ها")
    L.append("")
    L.append("| ID | متقاضی | کارفرما | شغل | کشور | ایمیل فرستاده به | تاریخ ارسال | مهلت پاسخ | وضعیت | پاسخ |")
    L.append("|----|--------|---------|-----|------|------------------|-------------|-----------|-------|------|")
    for a in sorted(apps, key=lambda x: x.get("sent_date", ""), reverse=True):
        reply = a.get("reply_date") or "—"
        dl = a.get("reply_deadline") or "—"
        L.append(
            f"| {a['id']} | {a.get('applicant','')} | {a.get('employer','')[:30]} | "
            f"{(a.get('job') or '—')[:30]} | {a.get('country','')} | {a.get('email_to') or '—'} | "
            f"{a.get('sent_date','')} | {dl} | {status_fa(a.get('status',''))} | {reply} |"
        )
    L.append("")

    # ── جزئیات هر رکورد ──
    L.append("## 📁 جزئیات رکوردها")
    for a in sorted(apps, key=lambda x: x.get("sent_date", ""), reverse=True):
        L.append("")
        L.append(f"### {a['id']} — {a.get('employer','')} ({a.get('applicant','')})")
        L.append("")
        L.append("| فیلد | مقدار |")
        L.append("|------|-------|")
        rows = [
            ("کارفرما", a.get("employer", "")),
            ("شغل / موضوع", a.get("job", "")),
            ("کشور", a.get("country", "")),
            ("ایمیل گیرنده", a.get("email_to", "") or "—"),
            ("موضوع ایمیل", a.get("subject", "")),
            ("تاریخ ارسال", a.get("sent_date", "")),
            ("مهلت پاسخ منتظره", a.get("reply_deadline", "") or "—"),
            ("تعداد پیگیری‌ها", a.get("followup_count", 0)),
            ("تاریخ آخرین پیگیری", a.get("followup_date", "") or "—"),
            ("وضعیت", status_fa(a.get("status", ""))),
            ("تاریخ پاسخ", a.get("reply_date", "") or "—"),
            ("خلاصه پاسخ", a.get("reply_summary", "") or "—"),
            ("تاریخ مصاحبه", a.get("interview_date", "") or "—"),
            ("تاریخ پیشنهاد شغلی", a.get("offer_date", "") or "—"),
            ("یادداشت", a.get("notes", "") or "—"),
        ]
        for k, v in rows:
            L.append(f"| **{k}** | {v} |")

    # ── خلاصه آماری ──
    L.append("")
    L.append("## 📈 خلاصه آماری")
    L.append("")
    L.append("| وضعیت | تعداد |")
    L.append("|--------|-------|")
    for st, (_, _, fa) in STATUS_FLOW.items():
        n = sum(1 for a in apps if a["status"] == st)
        if n:
            L.append(f"| {fa} | {n} |")
    L.append(f"| **جمع کل** | **{len(apps)}** |")
    L.append("")
    real = sum(1 for a in apps if a["status"] != "DRAFT")
    replied = sum(1 for a in apps if a["status"] in ("REPLIED", "INTERVIEW", "OFFER"))
    L.append(f"- درخواست‌های واقعی ارسال‌شده: **{real}**")
    L.append(f"- پاسخ دریافت‌شده: **{replied}**")
    overdue = sum(1 for a, r, _ in alerts if r < 0)
    L.append(f"- مهلت‌های گذشته (نیاز به پیگیری): **{overdue}**")
    L.append("")
    L.append("---")
    L.append("")
    L.append("## 🔄 چرخه وضعیت (راهنما)")
    L.append("")
    L.append("```")
    L.append("DRAFT → SENT (۷ روز مهلت) → FOLLOW_UP (۷ روز مهلت) → REPLIED (۵ روز)")
    L.append("      → INTERVIEW (۱۰ روز) → OFFER")
    L.append("هر مرحله می‌تواند به REJECTED یا WITHDRAWN ختم شود.")
    L.append("```")
    L.append("")
    L.append("**قوانین بانک:**")
    L.append("")
    L.append("1. هر درخواست فقط با `--add` یا `--seed` ثبت می‌شود — هیچ رکورد دستی بی‌منبع.")
    L.append("2. هر رکورد باید کارفرما، ایمیل گیرنده، تاریخ ارسال و مهلت پاسخ داشته باشد.")
    L.append("3. تغییر وضعیت فقط با `--status` انجام می‌شود تا تاریخ‌ها و مهلت‌ها خودکار به‌روز شوند.")
    L.append("4. اگر مهلت پاسخ گذشت و پاسخی نیامد → ارسال ایمیل پیگیری و ثبت آن با `--followup MH-XXXX --note \"…\"` — این دستور خودکار وضعیت `FOLLOW_UP` و مهلت پاسخ جدید ۷ روزه ثبت می‌کند.")
    L.append("5. این فایل (Markdown) هر بار توسط اسکریپت از `APPLICATION_BANK.json` بازتولید می‌شود — ویرایش دستی نکنید.")
    L.append("")
    L.append(f"> **آخرین بررسی خودکار:** {DATE_STR}")
    return "\n".join(L), alerts


# ═══════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser(description="MigrationHunter Application Tracker")
    ap.add_argument("--add", action="store_true", help="ثبت درخواست جدید (تعاملی)")
    ap.add_argument("--status", nargs=2, metavar=("ID", "STATUS"),
                    help="تغییر وضعیت رکورد: MH-2026-001 REPLIED")
    ap.add_argument("--date", help="تاریخ رویداد (با --status)", default=None)
    ap.add_argument("--summary", help="خلاصه پاسخ (با --status REPLIED)", default="")
    ap.add_argument("--followup", nargs=1, metavar="ID",
                    help="ثبت ارسال ایمیل پیگیری: وضعیت FOLLOW_UP + مهلت جدید ۷ روزه")
    ap.add_argument("--note", help="یادداشت پیگیری (با --followup)", default="")
    ap.add_argument("--seed", action="store_true", help="درون‌ریزی از EMAIL_TRACKER.json")
    ap.add_argument("--json", action="store_true", help="نمایش JSON به جای گزارش")
    args = ap.parse_args()

    print("═" * 56)
    print("MigrationHunter — بانک درخواست‌ها")
    print(f"📅 {DATE_STR}")
    print("═" * 56)

    bank = load_bank()
    changed = False

    if args.add:
        print("\n📝 ثبت درخواست جدید:")
        fields = {}
        for key, label, default in [
            ("applicant", "متقاضی (ندا/توحید)", ""),
            ("employer", "کارفرما", ""),
            ("job", "شغل", ""),
            ("country", "کشور (کد دوحرفی)", ""),
            ("email_to", "ایمیل گیرنده", ""),
            ("subject", "موضوع ایمیل", ""),
            ("sent_date", "تاریخ ارسال (YYYY-MM-DD)", TODAY.isoformat()),
        ]:
            val = input(f"  {label}" + (f" [{default}]" if default else "") + ": ").strip()
            fields[key] = val or default
        app = add_application(bank, status="SENT", **fields)
        print(f"\n  ✅ ثبت شد: {app['id']} — مهلت پاسخ: {app['reply_deadline']}")
        changed = True

    if args.followup:
        app, old = record_followup(bank, args.followup[0], args.date, args.note)
        if app:
            print(f"\n  🔁 {app['id']}: پیگیری #{app.get('followup_count', 1)} ثبت شد")
            print(f"     وضعیت: {status_fa(old)} → {status_fa(app['status'])}")
            print(f"     مهلت پاسخ جدید: {app['reply_deadline']} (۷ روز از تاریخ پیگیری)")
            changed = True
        else:
            print(f"\n  ❌ رکورد {args.followup[0]} یافت نشد")

    if args.status:
        app_id, new_status = args.status
        app, old = set_status(bank, app_id, new_status, args.date, args.summary)
        if app:
            print(f"\n  ✅ {app_id}: {status_fa(old)} → {status_fa(app['status'])}"
                  f" | مهلت جدید: {app.get('reply_deadline') or '—'}")
            changed = True
        else:
            print(f"\n  ❌ رکورد {app_id} یافت نشد")

    if args.seed:
        bank, added = seed_from_tracker(bank)
        print(f"\n  📥 {added} رکورد از EMAIL_TRACKER.json درون‌ریزی شد")
        changed = True

    if changed:
        save_bank(bank)
        print(f"  💾 ذخیره شد: {BANK_JSON}")

    if args.json:
        print(json.dumps(bank, ensure_ascii=False, indent=2))
        return

    report, alerts = build_report(bank)
    with open(BANK_MD, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"\n  ✅ گزارش فارسی: {BANK_MD}")

    # خلاصه
    apps = bank["applications"]
    overdue = sum(1 for a, r, _ in alerts if r < 0)
    urgent = sum(1 for a, r, _ in alerts if 0 <= r <= DEADLINE_URGENT)
    print("\n" + "═" * 56)
    print("📊 خلاصه")
    print("═" * 56)
    print(f"  📋 کل درخواست‌ها: {len(apps)}")
    print(f"  📤 ارسال‌شده: {sum(1 for a in apps if a['status'] in ('SENT','FOLLOW_UP'))}")
    print(f"  📬 پاسخ: {sum(1 for a in apps if a['status'] in ('REPLIED','INTERVIEW','OFFER'))}")
    if overdue:
        print(f"  ⏰ مهلت گذشته: {overdue} → پیگیری فوری!")
    if urgent:
        print(f"  🔴 مهلت فوری (≤{DEADLINE_URGENT} روز): {urgent}")
    if not overdue and not urgent:
        print("  ✅ هیچ مهلت فوری‌ای وجود ندارد")
    print("═" * 56)


if __name__ == "__main__":
    main()
