#!/usr/bin/env python3
"""
MigrationHunter — مسیر تحصیل: کشف برنامه‌های آموزشی

این کراولر دقیقاً مثل job_crawler.py کار می‌کند ولی به‌جای «آگهی شغلی»
دنبال «برنامهٔ تحصیلی» می‌گردد: کارشناسی، کارشناسی ارشد، دورهٔ زبان،
بورسیه. منابعش هم جدا در sources.json با track="education" مشخص شده‌اند —
هیچ آدرسی در این فایل هاردکد نیست.

هر دو کراولر زیرساخت مشترک job_crawler را قرض می‌گیرند (fetch، لاگ زنده،
progress، اکسل، dedupe) تا رفتارشان یکی باشد.

اجرا:
    python education_crawler.py                # همهٔ کشورها
    python education_crawler.py --country FI   # فقط فنلاند
    python job_crawler.py --track education    # همین کار، از درون کراولر کار
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import job_crawler as jc

BASE = jc.BASE
MEM = jc.MEM
DASH = jc.DASH
DATE_STR = jc.DATE_STR
FILE_DATE = jc.FILE_DATE

EDU_RESULTS_PATH = os.path.join(MEM, "EDUCATION_RESULTS.json")
EDU_BANK_PATH = os.path.join(MEM, "EDUCATION_BANK.json")


# ══════════════════════════════════════════════════════════════
# استخراج — JSON-LD
#
# پورتال‌های رسمی تحصیل (Studyinfo، DAAD، University of Helsinki و …)
# برنامه‌ها را با schema.org منتشر می‌کنند. سه نوع مهم:
#   Course                            — یک برنامهٔ درسی
#   EducationalOccupationalProgram   — برنامه با مدرک/شرایط
#   CourseInstance / Event            — یک نوبت مشخص با مهلت اپلای
# ══════════════════════════════════════════════════════════════
EDU_TYPES = (
    "Course", "EducationalOccupationalProgram", "CourseInstance",
    "EducationalOccupationalCredential", "Event",
)

# عنوان برنامه‌ها معمولاً با اینها شروع می‌شود — کمک به فیلتر نویز
PROGRAM_PREFIX_RE = re.compile(
    r"^(bachelor|master|masters|msc|ma\b|ms\b|ba\b|bs\b|phd|doctor|mba|"
    r"diploma|certificate|professional|year programme|year program|"
    r"kandidaatti|maisteri|opintokurssi|suuntautuvat|"
    r"laurea|täydentävät|jatkokurssi)",
    re.I,
)


def _edu_url_of(item):
    """آدرس برنامه را از چند جای احتمالی JSON-LD بیرون می‌کشد."""
    for key in ("url", "sameAs"):
        v = item.get(key)
        if isinstance(v, str) and v.startswith("http"):
            return v
    for holder in ("offers", "provider", "educationalProgram"):
        sub = item.get(holder)
        if isinstance(sub, dict):
            v = sub.get("url") or sub.get("sameAs")
            if isinstance(v, str) and v.startswith("http"):
                return v
        if isinstance(sub, list):
            for s in sub:
                if isinstance(s, dict):
                    v = s.get("url")
                    if isinstance(v, str) and v.startswith("http"):
                        return v
    return ""


def _edu_provider(item):
    org = item.get("provider") or item.get("publisher") or item.get("author") or {}
    if isinstance(org, dict):
        return (org.get("name") or "").strip()
    if isinstance(org, str):
        return org.strip()
    return ""


def _edu_degree(item):
    """مدرک: Bachelor / Master / … — مستقیم یا از نام برنامه."""
    for key in ("educationalCredentialAwarded", "degree", "award", "credentialCategory"):
        v = item.get(key)
        if isinstance(v, str) and v.strip():
            return v.strip()[:60]
        if isinstance(v, dict):
            n = v.get("name") or v.get("credentialCategory")
            if isinstance(n, str) and n.strip():
                return n.strip()[:60]
        if isinstance(v, list) and v:
            first = v[0]
            if isinstance(first, str):
                return first.strip()[:60]
            if isinstance(first, dict) and first.get("name"):
                return str(first["name"]).strip()[:60]
    title = item.get("name") or item.get("title") or ""
    m = re.match(r"^(bachelor|master|msc|mba|phd|doctor|ma\b|ms\b|ba\b|bs\b)", title, re.I)
    return m.group(1).upper() if m else ""


def _edu_language(item):
    lang = item.get("inLanguage") or item.get("teaches") or ""
    if isinstance(lang, dict):
        lang = lang.get("name") or lang.get("alternateName") or ""
    if isinstance(lang, list):
        lang = ", ".join(str(x) for x in lang)
    return str(lang)[:60]


def _edu_deadline(item):
    """مهلت اپلای — رایج‌ترین چیزی که برنامهٔ تحصیلی را قابل‌اقدام می‌کند."""
    for key in ("applicationDeadline", "validThrough", "deadline"):
        v = item.get(key)
        if isinstance(v, str) and v.strip():
            return v.strip()[:40]
        if isinstance(v, dict):
            for k in ("deadline", "dateTime", "value"):
                if isinstance(v.get(k), str) and v[k].strip():
                    return v[k].strip()[:40]
    return ""


def _edu_fee(item):
    offers = item.get("offers") or {}
    if isinstance(offers, list) and offers:
        offers = offers[0]
    if isinstance(offers, dict):
        for k in ("price", "lowPrice"):
            v = offers.get(k)
            if isinstance(v, (int, float)):
                cur = offers.get("priceCurrency") or ""
                return f"{cur} {v:,.0f}".strip()
        if isinstance(offers.get("price"), str):
            return offers["price"][:40]
    return ""


def extract_programs_jsonld(html, source_info):
    """برنامه‌های تحصیلی از JSON-LD (Course و مشتقاتش)."""
    programs = []
    for m in re.finditer(
        r'<script[^>]*type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        html, re.S | re.I
    ):
        try:
            data = json.loads(m.group(1).strip())
        except Exception:
            continue
        items = (data if isinstance(data, list)
                 else data.get("@graph", [data]) if isinstance(data, dict) else [data])
        for item in items:
            if not isinstance(item, dict):
                continue
            raw_type = item.get("@type")
            types = raw_type if isinstance(raw_type, list) else [raw_type]
            if not any(t in EDU_TYPES for t in types if isinstance(t, str)):
                continue
            title = str(item.get("name") or item.get("title") or "").strip()
            if not title or len(title) < 6:
                continue
            programs.append({
                "title": title[:150],
                "provider": _edu_provider(item)[:100] or "نامشخص",
                "degree": _edu_degree(item),
                "language": _edu_language(item),
                "deadline": _edu_deadline(item),
                "fee": _edu_fee(item),
                "country": source_info.get("country", "؟"),
                "source": source_info.get("name", "نامشخص"),
                "url": (_edu_url_of(item) or source_info.get("url", ""))[:300],
                "found_at": DATE_STR,
                "via": "json-ld",
            })
    return programs


# ══════════════════════════════════════════════════════════════
# استخراج — لینک‌های برنامه (پشتیبان JSON-LD)
# ══════════════════════════════════════════════════════════════
PROGRAM_PATH_RE = re.compile(
    r'(/program|/programme|/programs|/course|/courses|/study/|/studies/|'
    r'/degree|/majors?/|study-programmes|/study-program|/koulutus|/opinto|/suuntautuvat|'
    r'/tutkinto|/utbildning|/uddannelse|/-/study|'
    r'search-programmes|programmes-and-courses|find-a-program|'
    r'/master|/bachelor|/msc|/mba|/phd|'
    # مسیرهای واقعی که در بررسی زنده پیدا شد
    r'studies\.helsinki\.fi|/study$|/studies?$)',
    re.I,
)

PROGRAM_NOISE_RE = re.compile(
    r'(apply now|read more|see more|view all|sign in|log in|register|'
    r'privacy|cookie|terms|newsletter|subscribe|contact us|about us|'
    r'why study|why [a-z]+|cost|fees?$|funding|scholarships?$|'
    r'book a call|get started|explore|discover|home|search|'
    r'همه|بیشتر|جستجو|ورود|ثبت نام|تماس|درباره ما)',
    re.I,
)

# عنوان‌هایی که «صفحهٔ فهرست» هستند نه «یک برنامه»
LIST_PAGE_RE = re.compile(
    r'(all (programmes|programs|courses)|browse|search results|'
    r'find (a|your) (course|program)|view all|programme finder|course finder|'
    r'study options|fields of study|subjects?$)',
    re.I,
)


def extract_programs_from_links(html, source_info):
    """برنامه‌های تحصیلی از لینک‌های HTML — برای سایت‌هایی که JSON-LD ندارند."""
    parser = jc.DynamicJobParser(base_url=source_info.get("url", ""))
    # پارسر شغلی الگوی مسیر آگهی را دارد؛ برای تحصیل الگوی خودمان را می‌گذاریم
    parser.JOB_PATH_RE = PROGRAM_PATH_RE
    try:
        parser.feed(html)
    except Exception:
        pass

    programs = []
    seen = set()
    for li in parser.possible_links:
        href, text = li["href"], li["text"]
        if href in seen:
            continue
        if not PROGRAM_PATH_RE.search(href):
            continue
        if PROGRAM_NOISE_RE.search(text) or LIST_PAGE_RE.search(text):
            continue
        if len(text) < 8:
            continue
        # فقط عنوان‌هایی که بوی «برنامهٔ درسی» می‌دهند
        if not (PROGRAM_PREFIX_RE.search(text)
                or re.search(r"\b(programme|program|degree|course|bachelor|master|msc|ma\b|ms\b|phd)\b", text, re.I)):
            continue
        seen.add(href)
        programs.append({
            "title": text[:150],
            "provider": (li.get("near_company") or "")[:100] or "نامشخص",
            "degree": "", "language": "", "deadline": "", "fee": "",
            "country": source_info.get("country", "؟"),
            "source": source_info.get("name", "نامشخص"),
            "url": href[:300],
            "found_at": DATE_STR,
            "via": "html-link",
        })
        if len(programs) > 120:
            break
    return programs


def extract_programs_from_html(html, source_info):
    """استخراج برنامهٔ تحصیلی — اول JSON-LD، بعد لینک‌ها."""
    programs = []
    try:
        programs.extend(extract_programs_jsonld(html, source_info))
    except Exception:
        pass
    seen = {p["url"] for p in programs if p.get("url")}
    for p in extract_programs_from_links(html, source_info):
        if p["url"] in seen:
            continue
        seen.add(p["url"])
        programs.append(p)
    return programs


# ══════════════════════════════════════════════════════════════
# اکسل
# ══════════════════════════════════════════════════════════════
EDU_HEADERS = ["#", "نام برنامه", "دانشگاه/مؤسسه", "مدرک", "زبان",
               "مهلت اپلای", "شهریه", "کشور", "منبع", "لینک برنامه",
               "تاریخ یافت", "تطبیق با پروفایل", "متقاضی پیشنهادی"]

EDU_WIDTHS = [5, 48, 28, 12, 12, 14, 14, 8, 22, 48, 16, 14, 20]


def _row(n, p, applicants_config):
    r = jc.detect_job_realness(p.get("title", ""), p.get("provider", ""), applicants_config)
    suggestion = "—"
    t = (p.get("title") or "").lower()
    for a in applicants_config:
        if any((k or "").lower() in t for k in a.get("keywords", [])):
            suggestion = a.get("name_fa") or a.get("name", a.get("id", ""))
            break
    return [
        n, p.get("title", ""), p.get("provider", "نامشخص"), p.get("degree", "") or "—",
        p.get("language", "") or "—", p.get("deadline", "") or "—", p.get("fee", "") or "—",
        p.get("country", "") or "؟", p.get("source", ""), p.get("url", ""),
        p.get("found_at", ""), r["match_score"], suggestion,
    ]


def build_programs_excel(programs, applicants_config):
    return jc.build_excel(programs, applicants_config, None, "برنامه‌های تحصیلی",
                          EDU_HEADERS, EDU_WIDTHS, _row, 9)


# ══════════════════════════════════════════════════════════════
# بانک اپلیکیشن تحصیلی — پیگیری هر برنامه تا پذیرش
# ══════════════════════════════════════════════════════════════
APP_STATUSES = [
    ("PLANNED", "📌", "در نظر گرفته‌شده"),
    ("APPLYING", "✍️", "در حال آماده‌سازی مدارک"),
    ("SUBMITTED", "📤", "ارسال شده"),
    ("ACCEPTED", "🎉", "پذیرش گرفته شد"),
    ("ENROLLED", "🎓", "ثبت‌نام قطعی"),
    ("REJECTED", "❌", "رد شد"),
    ("WITHDRAWN", "🚪", "انصراف داده شد"),
]
APP_STATUS_FA = {k: fa for k, _e, fa in APP_STATUSES}
APP_STATUS_EMOJI = {k: e for k, e, _fa in APP_STATUSES}


def _load_bank():
    if not os.path.exists(EDU_BANK_PATH):
        return {"applications": []}
    try:
        with open(EDU_BANK_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {"applications": []}
    except Exception:
        return {"applications": []}


def _save_bank(bank):
    os.makedirs(MEM, exist_ok=True)
    with open(EDU_BANK_PATH, "w", encoding="utf-8") as f:
        json.dump(bank, f, ensure_ascii=False, indent=2)


def load_applications(applicant=None, country=None):
    apps = _load_bank().get("applications", [])
    if applicant:
        apps = [a for a in apps if a.get("applicant") == applicant]
    if country:
        cc = country.upper()
        apps = [a for a in apps if (a.get("country") or "").upper() == cc]
    return sorted(apps, key=lambda a: a.get("deadline") or "9999")


def save_applications(apps):
    _save_bank({"applications": apps})
    return apps


def add_application(applicant, title, provider="", country="", url="",
                    deadline="", status="PLANNED", notes="", degree="", language=""):
    """افزودن یا به‌روزرسانی یک اپلیکیشن تحصیلی (کلید یکتا = applicant+url|title)."""
    apps = _load_bank().get("applications", [])
    key = (url or title or "").strip().lower()
    entry = {
        "applicant": applicant, "title": title, "provider": provider,
        "country": country.upper() if country else "", "url": url,
        "deadline": deadline, "degree": degree, "language": language,
        "status": status, "notes": notes,
        "created_at": jc.DATE_STR, "updated_at": jc.DATE_STR,
        "history": [{"date": jc.DATE_STR, "status": status, "note": "ثبت اولیه"}],
    }
    for i, a in enumerate(apps):
        if a.get("applicant") == applicant and (a.get("url") or a.get("title", "")).strip().lower() == key:
            entry["created_at"] = a.get("created_at", jc.DATE_STR)
            entry["history"] = a.get("history", []) + [
                {"date": jc.DATE_STR, "status": status, "note": notes or "به‌روزرسانی"}]
            apps[i] = entry
            save_applications(apps)
            return entry
    apps.append(entry)
    save_applications(apps)
    return entry


def set_app_status(applicant, key, new_status, note=""):
    apps = _load_bank().get("applications", [])
    key = (key or "").strip().lower()
    for a in apps:
        if a.get("applicant") == applicant and (a.get("url") or a.get("title", "")).strip().lower() == key:
            a["status"] = new_status
            a["updated_at"] = jc.DATE_STR
            a.setdefault("history", []).append(
                {"date": jc.DATE_STR, "status": new_status, "note": note})
            if note:
                a["notes"] = (a.get("notes", "") + "\n" + note).strip()
            save_applications(apps)
            return a
    return None


def add_app_note(applicant, key, note):
    apps = _load_bank().get("applications", [])
    key = (key or "").strip().lower()
    for a in apps:
        if a.get("applicant") == applicant and (a.get("url") or a.get("title", "")).strip().lower() == key:
            a["notes"] = (a.get("notes", "") + "\n" + note).strip()
            a["updated_at"] = jc.DATE_STR
            a.setdefault("history", []).append(
                {"date": jc.DATE_STR, "status": a.get("status", ""), "note": note})
            save_applications(apps)
            return a
    return None


def remove_application(applicant, key):
    apps = _load_bank().get("applications", [])
    key = (key or "").strip().lower()
    keep = [a for a in apps
            if not (a.get("applicant") == applicant
                    and (a.get("url") or a.get("title", "")).strip().lower() == key)]
    save_applications(keep)
    return len(apps) - len(keep)


# وضعیت‌هایی که یعنی «فرایند شروع شده و برنامه دیگر باز نیست»
APP_ADVANCED = ("SUBMITTED", "ACCEPTED", "ENROLLED")
# وضعیت‌هایی که یعنی «هنوز باید کاری کنم»
APP_OPEN = ("PLANNED", "APPLYING")


def app_stats(applications=None):
    """خلاصهٔ وضعیت اپلیکیشن‌ها.

    by_status    → شمارش هر وضعیت
    submitted    → ارسال‌شده/پذیرش/ثبت‌نام (چیزی که واقعاً اقدامی شده)
    accepted     → پذیرش یا ثبت‌نام قطعی
    open         → هنوز باید کاری کنم
    no_deadline  → مهلتی ثبت نشده (ریسک فراموشی)
    """
    apps = applications if applications is not None else load_applications()
    by_status = {}
    for a in apps:
        st = a.get("status", "PLANNED")
        by_status[st] = by_status.get(st, 0) + 1
    return {
        "total": len(apps),
        "by_status": by_status,
        "statuses": APP_STATUSES,
        "submitted": sum(by_status.get(k, 0) for k in APP_ADVANCED),
        "accepted": by_status.get("ACCEPTED", 0) + by_status.get("ENROLLED", 0),
        "open": sum(by_status.get(k, 0) for k in APP_OPEN),
        "no_deadline": sum(1 for a in apps if not (a.get("deadline") or "").strip()),
    }


def deadline_warnings(applications=None, days=30):
    """برنامه‌هایی که مهلتشان نزدیک است یا گذشته."""
    import datetime as _dt
    apps = applications if applications is not None else load_applications()
    today = _dt.date.today()
    soon, passed = [], []
    for a in apps:
        dl = (a.get("deadline") or "").strip()
        if not dl:
            continue
        try:
            d = _dt.date.fromisoformat(dl[:10])
        except ValueError:
            continue
        if d < today:
            passed.append(a)
        elif (d - today).days <= days:
            soon.append(a)
    return soon, passed


# ══════════════════════════════════════════════════════════════
def main(argv=None):
    """همان کراولر شغلی، ولی روی منابع تحصیلی.

    --track عمداً فقط برای اطمینان خوانده می‌شود: این فایل مسیر تحصیل است و
    نباید با --track job روی منابع کاریابی اجرا شود (آن job_crawler.py است).
    اگر مسیر اشتباه داده شود خطا می‌دهیم به‌جای اینکه خاموش روی مسیر دیگری
    برود — web_ui همیشه --track می‌فرستد، پس بدون این مرحله اجرا نمی‌شد.
    """
    import argparse
    ap = argparse.ArgumentParser(description="MigrationHunter — جستجوی برنامه‌های تحصیلی")
    ap.add_argument("--track", default="education",
                    help="باید education باشد (job_crawler.py برای job است)")
    ap.add_argument("--country", default=None, help="فقط یک کشور، مثلاً FI")
    ap.add_argument("--no-excel", action="store_true", help="فایل اکسل نساز")
    args = ap.parse_args(argv)

    if (args.track or "education").lower() != "education":
        ap.error(f"education_crawler.py فقط مسیر تحصیل است (track=education)، نه «{args.track}» "
                 f"— برای کاریابی از job_crawler.py استفاده کن.")
    return jc.main(track="education", country=args.country,
                   discover=False, make_excel=not args.no_excel)


if __name__ == "__main__":
    sys.exit(main() or 0)
