# -*- coding: utf-8 -*-
"""
unified_report.py — لایهٔ یکپارچه‌سازی خروجی‌ها (v4.1)

هدف این ماژول یک چیز است: کاربر نباید برای فهمیدنِ «فلان متقاضی چه چیزی
دارد» بین ۲۰ فایل پراکنده در memory/ بگردد.

هر بانک جداگانه فقط یک تکه از پازل را نگه می‌دارد:
  SPONSOR_RESULTS.json  → ۲۱۹۱ آگهی اسپانسر (هیچ‌جا در UI نشان داده نمی‌شد)
  CRAWLER_RESULTS.json  → ۷۰ آگهی کراولر رسمی
  APPLICATION_BANK.json → درخواست‌هایی که فرستاده شده
  EMAIL_ANALYSIS.json   → ۳۱ ایمیل شغلی پیدا‌شده
  EMAIL_TRACKER.json    → ۵ ایمیل ارسالی + پیگیری
  REMINDERS.json        → ۱۲ یادآور

این ماژول همه را می‌خواند، برای هر متقاضی جدا می‌کند، تکراری‌ها را با URL
یکی می‌کند و یک «دیکشنری اقدام» می‌دهد که UI مستقیم از رویش رندر می‌شود.

هیچ داده‌ای اینجا ساخته یا حدس زده نمی‌شود — فقط خوانده، نرمال‌سازی و
ادغام می‌شود. اگر یک بانک نبود، آن بخش خالی می‌ماند و صفحه خطا نمی‌دهد.
"""

import json
import os
import re
from datetime import datetime, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
MEM_DIR = os.path.join(BASE, "memory")

# ── وضعیت‌های درخواست (هم‌نام APPLICATION_BANK تا نیمی از داده‌ها گم نشود) ──
APP_STATUS = {
    "PLANNED":    {"fa": "در نظر گرفته‌شده", "emoji": "🗒️"},
    "PREPARING":  {"fa": "در حال آماده‌سازی", "emoji": "✍️"},
    "SENT":       {"fa": "ارسال‌شده",        "emoji": "📤"},
    "FOLLOW_UP":  {"fa": "پیگیری‌شده",       "emoji": "🔁"},
    "REPLIED":    {"fa": "پاسخ گرفته",       "emoji": "📬"},
    "INTERVIEW":  {"fa": "مصاحبه",           "emoji": "🎤"},
    "OFFER":      {"fa": "پیشنهاد شغلی",     "emoji": "🎉"},
    "REJECTED":   {"fa": "رد‌شده",           "emoji": "❌"},
    "WITHDRAWN":  {"fa": "انصراف",           "emoji": "🚪"},
}

# ترتیب نمایش از «تازه‌ترین اقدام لازم» تا «بسته‌شده»
STATUS_ORDER = ["PLANNED", "PREPARING", "SENT", "FOLLOW_UP", "REPLIED",
                "INTERVIEW", "OFFER", "REJECTED", "WITHDRAWN"]

OPEN_STATUSES = ("PLANNED", "PREPARING", "SENT", "FOLLOW_UP")

URGENCY_EMOJI = {"INTERVIEW": "🎤", "OFFER": "🎉", "REPLIED": "📬"}

# مقادیر شرکت/عنوان که در واقع «آگهی» نیستند ولی از متن HTML صفحه
# نشت کرده‌اند. نمونهٔ واقعی از job_market.fi:
#   company = "There are 1109 open jobs"  یا  "Go to job"
# اگر این‌ها پاک‌سازی نشوند، کاربر در گزارش «شرکت: There are 1109 open jobs»
# می‌بیند و گزارش را بی‌اعتماد می‌کند.
JUNK_COMPANY = re.compile(
    r"(there (are|is)\s+\d+|open jobs?|go to job|view job|see job|"
    r"\d+\s+open\s+jobs?|job search|careers? page|näytä|katso työ|"
    r"all jobs|browse jobs|search jobs)", re.I)

# عنوان‌هایی که صفحهٔ راهنما/خدمات هستند نه آگهی استخدام
NOT_A_JOB = re.compile(
    r"(job search discussions?|creating a job posting|monitoring of|"
    r"information about working life|how to (find|apply)|register as|"
    r"employment services?|statistics and reports?|guidance|instructions?|"
    r"forms? and (documents?|templates)|news|blog|press|event)", re.I)


FIN_CITIES = (r"(?:Helsinki|Espoo|Tampere|Vantaa|Turku|Oulu|Jyväskylä|Kuopio|"
              r"Lahti|Pori|Rauma|Salo|Kotka|Porvoo|Loviisa|Kauniainen|Kerava|"
              r"Nurmijärvi|Somero|Hämeenlinna|Siilinjärvi|Kärsämäki|Valtimo|"
              r"Kajaani|Rovaniemi|Oulu|Tornio|Lapua|Vantaa)")

# الگوی «انتهای عنوان + نام شرکت + شهر» که HTML فشرده به هم می‌چسباند:
#   "Senior Software Engineer (AI Platform)SmartlyHelsinki"
#   "Storage Engineering ManagerVerdaHelsinki"
#   "Senior System Specialist (LUMI-AI Lead Admin)CSC - Tieteen tarkastuslaitos"
TRAILING_COMPANY_RE = re.compile(
    r"\)[ ]*(?P<c1>[A-Z][A-Za-zÀ-ÿ&\.\-]{2,30}?)" + FIN_CITIES + r"\b|"
    r"(?<=[a-z\)])(?P<c2>[A-Z][A-Za-zÀ-ÿ&\.\-]{2,30}?)" + FIN_CITIES + r"\b")


def _split_trailing_company(text):
    """
    انتهای عنوانِ چسبیده را به شرکت جدا می‌کند و (عنوان، شرکت) برمی‌گرداند.
    اگر الگو نخورد، خود متن برگردانده می‌شود و شرکت None است.
    """
    m = TRAILING_COMPANY_RE.search(text)
    if not m:
        return text, None
    cand = (m.group("c1") or m.group("c2") or "").strip()
    if len(cand) < 3 or JUNK_COMPANY.search(cand):
        return text, None
    # پرانتزِ قبل از شرکت («(AI Platform)Smartly») بخشی از عنوان است، پس
    # از دست ندهیم — فقط از محل شروعِ «نام شرکت» به بعد را حذف می‌کنیم.
    cut = m.start("c1") if m.group("c1") else m.start("c2")
    return text[:cut].strip(), cand


def clean_title(value, fallback=""):
    """
    عنوان آگهی را از متن چسبیدهٔ HTML جدا می‌کند.

    نمونهٔ واقعی از job_market.fi:
        "Senior Software Engineer (AI Platform)SmartlyHelsinkiGo to job site"
    در واقع: عنوان = "Senior Software Engineer (AI Platform)"
             شرکت = "Smartly"
    """
    t = (value or "").strip()
    # ۱) از اولین دکمه/CTA به بعد را دور می‌ریزیم
    t = re.split(r"(?:Go to job|View job|Apply now|Read more|Lue lisää|Katso työ|"
                 r"Hakemuksen|See all|Show all)", t, flags=re.I)[0]
    # ۲) اگر انتهای عنوان شرکت+شهر چسبیده، جدا شود
    t, _ = _split_trailing_company(t)
    # پرانتزها بخشی از عنوان‌اند («Senior (AI Platform)») — فقط اگر
    # واقعاً نامتوازن باشند دست می‌زنیم.
    t = t.strip(" \t–—-·|,")
    if t.count("(") > t.count(")"):
        t = t.rsplit("(", 1)[0].strip() or t
    return t[:160] or fallback


def clean_company(value, title="", fallback="نامشخص"):
    """
    نام شرکت را از نویز HTML پاک می‌کند.

    اولویت: فیلد company خودِ رکورد → اگر نویزی بود، از انتهای عنوانِ
    چسبیده بیرون بکشیم → در غیر این صورت «نامشخص».
    """
    # اول: شرکتِ قابل استفاده از خود فیلد value
    v = (value or "").strip()
    if v and not JUNK_COMPANY.search(v) and len(v) >= 2:
        # «SmartlyHelsinkiGo to job site» → «Smartly»
        v2 = re.split(r"(?=[A-Z][a-z]{2,}(?:Go to job|View job))", v)[0].strip()
        v2 = re.sub(FIN_CITIES + r"\b$", "", v2).strip()
        if v2 and not JUNK_COMPANY.search(v2):
            return v2[:80]

    # دوم: شرکت از انتهای عنوانِ چسبیده (CTA را اول حذف می‌کنیم)
    if title:
        head = re.split(r"(?:Go to job|View job|Apply now|Read more|Lue lisää|"
                        r"Katso työ|Hakemuksen)", title, flags=re.I)[0]
        _, cand = _split_trailing_company(head)
        if cand:
            return cand[:80]
    return fallback


def is_real_job(job):
    """آیا این رکورد واقعاً آگهی استخدام است یا صفحهٔ راهنما/خدمات؟"""
    title = job.get("title") or ""
    if NOT_A_JOB.search(title):
        return False
    return True


def _read_json(name, default=None):
    """خواندن امن JSON — فایل نبود یا خراب بود نباید کل صفحه را از کار بیندازد."""
    path = os.path.join(MEM_DIR, name)
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _now():
    return datetime.now()


now = _now()


def _norm(text):
    return re.sub(r"[\s_\-]+", " ", (text or "").strip().lower())


def _parse_date(value):
    """تاریخ را به datetime برمی‌گرداند یا None. چند قالب را می‌پذیرد."""
    if not value:
        return None
    s = str(value).strip().replace("/", "-")
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                "%d-%m-%Y", "%Y-%m-%dT%H:%M"):
        try:
            return datetime.strptime(s[:len(fmt) + 2].strip(), fmt)
        except Exception:
            continue
    return None


# ══════════════════════════════════════════════════════════════════
# تطبیق آگهی با متقاضی — فراتر از شمارش سادهٔ کلیدواژه
# ══════════════════════════════════════════════════════════════════

# کلمات بی‌معنا برای تطبیق. توجه: «it» را حذف نمی‌کنیم چون رشتهٔ
# Information Technology یک مسیر شغلی واقعی است؛ به‌جایش در متن با
# مرزبندی \b تطبیق داده می‌شود تا داخل «with» یا «site» اشتباه گرفته نشود.
STOPWORDS = {"and", "or", "the", "a", "an", "for", "of", "in", "on",
              "to", "with", "at", "job", "work", "jobs"}

# معادل‌های شغلی — کلیدواژهٔ انگلیسی متقاضی ↔ عنوان فنلandsی/سوئدی آگهی
# بدون این جدول، «midwife» هرگز با «sairaanhoitaja» تطبیق نمی‌خورد و ندا
# عملاً هیچ آگهی‌ای نمی‌بیند.
TITLE_SYNONYMS = {
    "midwife": ["sairaanhoitaja", "lähihoitaja", "lahihoitaja", "sjukskötare",
                "hoitaja", "terveydenhoitaja", "midwife", "nurse", "carer"],
    "nurse": ["sairaanhoitaja", "lähihoitaja", "hoitaja", "nurse", "registered nurse"],
    "it": ["ohjelmistokehittäjä", "ohjelmoija", "software", "developer",
           "system", "tietoturva", "data", "analyst", "engineer", "helpdesk",
           "tekninen", "sovellus", "ohjelmisto"],
    "software": ["ohjelmistokehittäjä", "ohjelmoija", "software", "developer"],
    "engineering": ["insinööri", "engineer", "tekniikka", "engineering",
                    "teknologi", "automaatio", "sähkö"],
    "enginering": ["insinööri", "engineer", "tekniikka", "engineering",
                   "teknologi", "automaatio", "sähkö"],  # غلط املایی پرتکرار
    "data": ["data", "analyst", "analyytikko", "tieto"],
}

# کشور → نام نمایشی
COUNTRY_FA = {
    "FI": "فنلاند", "SE": "سوئد", "NO": "نروژ", "DK": "دانمارک",
    "DE": "آلمان", "NL": "هلند", "CA": "کانادا", "AU": "استرالیا",
    "NZ": "نیوزیلند", "GB": "انگلستان", "IE": "ایرلند", "US": "آمریکا",
    "IT": "ایتالیا", "ES": "اسپانیا", "FR": "فرانسه", "AT": "اتریش",
    "CH": "سوئیس", "PL": "لهستان", "PT": "پرتغال", "BE": "بلژیک",
}


def applicant_terms(applicant):
    """
    کلیدواژه‌های واقعاً مرتبط یک متقاضی را برمی‌گرداند.

    نام و نام خانوادگی کنار گذاشته می‌شوند: هیچ آگهی شغلی نام متقاضی را
    نمی‌نویسد، پس وجودشان فقط امتیاز الکی تولید می‌کند. (برای ندا فقط
    «midwife» می‌ماند، برای توحید «enginering» و «it».)
    """
    # نام‌هایی که متعلق به خودِ متقاضی‌اند و ربطی به شغل ندارند
    personal = set()
    for field in ("name", "name_fa", "id"):
        v = _norm(applicant.get(field))
        if not v:
            continue
        personal.add(v)
        # «توحید ارجمند» → «ارجمند» هم باید حذف شود
        for part in v.split():
            if len(part) >= 3:
                personal.add(part)

    out = []
    for k in (applicant.get("keywords") or []):
        k = (k or "").strip()
        if not k:
            continue
        n = _norm(k)
        if not n or n in STOPWORDS or n in personal:
            continue
        if n not in [x[0] for x in out]:
            out.append((n, k))
    return out


def score_job(job, applicant):
    """
    امتیاز تطبیق آگهی با متقاضی: ۰ تا ۱۰۰.

    برخلاف نسخهٔ قبلی (شمارش سادهٔ کلیدواژه) این نسخه:
      • نام متقاضی را امتیاز نمی‌دهد (آگهی شغلی نام متقاضی ندارد)
      • معادل‌های فنلاندی/سوئدی را می‌شناسد (midwife ↔ sairaanhoitaja)
      • به متن توضیحات آگهی هم نگاه می‌کند، نه فقط عنوان
    """
    title = _norm(job.get("title"))
    desc = _norm(job.get("description") or job.get("summary") or "")
    company = _norm(job.get("company"))
    hay_title = title + " " + company

    score = 0
    hits = []
    for norm_kw, original in applicant_terms(applicant):
        variants = TITLE_SYNONYMS.get(norm_kw, [norm_kw])
        for v in variants:
            if re.search(r"\b" + re.escape(v.lower()) + r"\b", hay_title):
                score += 30
                hits.append((original, "عنوان"))
                break
        else:
            for v in variants:
                if desc and re.search(r"\b" + re.escape(v.lower()) + r"\b", desc):
                    score += 12
                    hits.append((original, "توضیحات"))
                    break

    return min(score, 100), hits


# ══════════════════════════════════════════════════════════════════
# جمع‌آوری آگهی‌ها از همهٔ بانک‌ها
# ══════════════════════════════════════════════════════════════════

def _load_sponsor_jobs():
    sp = _read_json("SPONSOR_RESULTS.json", {}) or {}
    jobs = sp.get("jobs") or []
    out = []
    for j in jobs:
        if not isinstance(j, dict) or not j.get("title"):
            continue
        if not is_real_job(j):
            continue
        raw_title = j["title"]
        title = clean_title(raw_title)
        out.append({
            "title": title,
            "company": clean_company(j.get("company"), raw_title),
            "country": (j.get("country") or "").strip().upper()[:2],
            "location": j.get("location") or "",
            "url": (j.get("url") or "").strip(),
            "description": (j.get("description") or "")[:1200],
            "posted": j.get("posted") or "",
            "source": j.get("source") or "Sponsor",
            "via": "sponsor",
        })
    return out


def _load_crawler_jobs():
    cr = _read_json("CRAWLER_RESULTS.json", {}) or {}
    out = []
    for j in cr.get("jobs") or []:
        if not isinstance(j, dict) or not j.get("title"):
            continue
        if not is_real_job(j):
            continue
        raw_title = j["title"]
        title = clean_title(raw_title)
        out.append({
            "title": title,
            "company": clean_company(j.get("company"), raw_title),
            "country": (j.get("country") or "").strip().upper()[:2],
            "location": j.get("location") or "",
            "url": (j.get("url") or "").strip(),
            "description": "",
            "posted": j.get("date_posted") or "",
            "source": j.get("source") or "Crawler",
            "via": "crawler",
        })
    return out


def all_jobs():
    """همهٔ آگهی‌ها از همهٔ بانک‌ها، بدون تکرار بر اساس URL."""
    merged = _load_crawler_jobs() + _load_sponsor_jobs()
    seen = set()
    out = []
    for j in merged:
        key = _norm(j.get("url")) or (j.get("title", "").lower() + j.get("company", "").lower())
        if key in seen:
            continue
        seen.add(key)
        out.append(j)
    return out


# ══════════════════════════════════════════════════════════════════
# وضعیت اقدام برای هر آگهی (کاربر قبلاً چه کرده؟)
# ══════════════════════════════════════════════════════════════════

def _load_linkedin(applicant):
    """
    پروفایل‌های لینکدین این متقاضی از LINKEDIN_DB.json.

    تا قبل از این، هیچ کدی این فایل را نمی‌خواند — پروفایل‌های لینکدین
    هر دو نفر (neda-arjmand و tohid-arjmand) در گزارش‌ها نامرئی بودند.
    """
    db = _read_json("LINKEDIN_DB.json", {}) or {}
    out = []
    for li in (db.get("linkedins") or []):
        if not isinstance(li, dict):
            continue
        # رکورد لینکدین فیلد applicant با کد لاتین دارد («NEDA»)؛
        # _matches_applicant همان را هم می‌فهمد. URL را هم با
        # لینک‌های ثبت‌شده در config مقایسه می‌کنیم.
        if _matches_applicant(li, applicant):
            out.append({
                "name": li.get("name") or "",
                "url": li.get("url") or "",
                "profession": li.get("profession") or "",
            })
    return out


def _load_applications():
    ab = _read_json("APPLICATION_BANK.json", {}) or {}
    return ab.get("applications") or []


def application_state(applications):
    """نقشهٔ url → وضعیت درخواست، برای اینکه UI بداند کدام آگهی اقدام‌شده است."""
    state = {}
    for a in applications:
        url = _norm(a.get("job_url") or a.get("url"))
        if url:
            state[url] = a
    return state


# ══════════════════════════════════════════════════════════════════
# نمای یکپارچهٔ هر متقاضی
# ══════════════════════════════════════════════════════════════════

def _applicant_aliases(applicant):
    """
    همهٔ نام‌هایی که یک متقاضی در بانک‌های مختلف با آن‌ها ثبت شده است.

    مشکل واقعی: هر بانک با یک قرارداد اسم نوشته —
      APPLICATION_BANK / EMAIL_TRACKER : «ندا»        (اسم کوتاه)
      EMAIL_ANALYSIS                   : «ندا_ارجمند» (id کامل config)
      LINKEDIN_DB                      : «NEDA»        (کد لاتین بزرگ)
    تطبیق قبلی فقط id کامل را می‌فهمید، پس عملاً داده‌های ندا
    (۵ درخواست، ۵ ایمیل ارسالی، ۱۲ یادآور) نامرئی بودند و گزارش
    انگار فقط «یک نفر» را می‌شناخت.
    """
    aliases = set()
    for field in ("id", "name", "name_fa"):
        v = _norm(applicant.get(field))
        if v:
            aliases.add(v)
            # «ندا ارجمند» → «ندا» و «ارجمند» هم اسم همین شخص‌اند
            for part in re.split(r"[^a-z\u00c0-\u024f\u0370-\u03ff\u0400-\u04ff"
                                 r"\u0600-\u06ff0-9]+", v):
                if len(part) >= 2:
                    aliases.add(part)
    return {a for a in aliases if a}


_SHARED_TOKENS = None


def _shared_tokens():
    """
    توکن‌هایی که بین چند متقاضی مشترک‌اند (مثل نام خانوادگی «ارجمند»).

    این‌ها برای تطبیقِ «اسم داخل مقدار رکورد» ممنوع‌اند؛ وگرنه رکورد
    «توحید_ارجمند» به‌خاطر «ارجمند» به ندا هم می‌چسبد و هر دو نفر
    همهٔ ایمیل‌ها را می‌بینند. فقط یک‌بار از config خوانده و کش می‌شود.
    """
    global _SHARED_TOKENS
    if _SHARED_TOKENS is None:
        try:
            with open(os.path.join(BASE, "config.json"), encoding="utf-8") as f:
                apps = (json.load(f) or {}).get("applicants", [])
        except Exception:
            apps = []
        from collections import Counter
        c = Counter()
        for a in apps:
            c.update(set(_applicant_aliases(a)))
        _SHARED_TOKENS = {t for t, n in c.items() if n > 1}
    return _SHARED_TOKENS


def _matches_applicant(record, applicant):
    """آیا این رکورد به این متقاضی مربوط است؟ نام/ایمیل/linkedin تطبیق داده می‌شود."""
    aliases = _applicant_aliases(applicant)
    shared = _shared_tokens()
    emails = {_norm(e) for e in (applicant.get("emails") or []) if _norm(e)}
    if applicant.get("email"):
        emails.add(_norm(applicant["email"]))
    linkedins = {_norm(l) for l in (applicant.get("linkedins") or []) if _norm(l)}
    if applicant.get("linkedin"):
        linkedins.add(_norm(applicant["linkedin"]))

    for field in ("applicant", "applicant_id", "person", "account_id", "owner"):
        val = _norm(record.get(field))
        if not val or len(val) < 2:
            continue
        for a in aliases:
            if val == a:
                return True
            # مقدار رکورد داخل اسم کامل متقاضی: «ندا» در «ندا ارجمند» ✓
            if len(val) >= 2 and val in a:
                return True
            # اسم متقاضی داخل مقدار رکورد: فقط اگر متمایز باشد —
            # «ارجمند» در «توحید_ارجمند» برای ندا ممنوع است (مشترک خانوادگی)
            if len(a) >= 2 and a in val and a not in shared:
                return True
    mail = _norm(record.get("email") or record.get("to")
                 or record.get("recipient_email"))
    if mail and mail in emails:
        return True
    for f in ("linkedin", "url", "profile"):
        li = _norm(record.get(f))
        if li and any(li in x or x in li for x in linkedins if x):
            return True
    return False


def _tracker_applicant_map():
    """
    نگاشت email_id → applicant از روی EMAIL_TRACKER.

    یادآورها (REMINDERS) و اعلان‌ها فیلد متقاضی ندارند و فقط email_id
    دارند؛ بدون این نگاشت به هیچ‌کس نمی‌رسیدند و گم می‌شدند.
    """
    et = _read_json("EMAIL_TRACKER.json", {}) or {}
    return {e.get("id"): e.get("applicant")
            for e in (et.get("emails") or []) if e.get("id")}


def _job_rank(job):
    """
    ترتیب نمایش آگهی‌ها. اولویت با کاری است که کاربر کرده، بعد تازگی،
    بعد امتیاز تطبیق — نه صرفاً امتیاز، چون آگهی‌ای که برایش اقدام کرده
    از آگهی ۶۰ امتیازیِ تازه مهم‌تر است.
    """
    posted = _parse_date(job.get("posted") or job.get("found_at"))
    days_old = (now - posted).days if posted else 9999
    return (
        1 if job.get("applied") else 0,
        -min(days_old, 9999),
        job.get("score", 0),
    )


def applicant_dossier(applicant, limit=60, country=None):
    """
    همه‌چیزِ مربوط به یک متقاضی در یک دیکشنری:
      jobs        آگهی‌های تطبیق‌یافته با وضعیت اقدام و لینک مستقیم
      applications  درخواست‌های ثبت‌شده
      emails      ایمیل‌های شغلیِ پیدا‌شده برای این شخص
      sent_emails ایمیل‌های ارسالی و پیگیری‌ها
      reminders   یادآورهای سررسید
      stats       خلاصهٔ عددی برای کارت‌های بالای صفحه

    country=None یعنی «همهٔ کشورها». توجه کن که بدون این فیلتر، برای یک
    متقاضی IT مثلاً ۳۰۰ آگهی آمریکا برمی‌گردد که عملاً بی‌استفاده است؛
    بنابراین UI همیشه کشور را از کوئری می‌فرستد.
    """
    aid = applicant.get("id")
    all_app_rows = _load_applications()
    my_apps = [a for a in all_app_rows if _matches_applicant(a, applicant)]
    state = application_state(my_apps)

    # ── آگهی‌ها ──
    scored = []
    for j in all_jobs():
        if country and (j.get("country") or "").upper() != country.upper():
            continue
        s, hits = score_job(j, applicant)
        if s <= 0:
            continue
        jj = dict(j)
        jj["score"] = s
        jj["matched_on"] = hits
        key = _norm(j.get("url"))
        existing = state.get(key)
        jj["applied"] = bool(existing)
        jj["app_status"] = existing.get("status", "") if existing else ""
        jj["app_id"] = existing.get("id", "") if existing else ""
        jj["reply_deadline"] = existing.get("reply_deadline") if existing else None
        jj["_why"] = "، ".join(f"{k} ({w})" for k, w in hits[:3])
        scored.append(jj)
    scored.sort(key=_job_rank)
    jobs_out = scored[:limit]

    # ── ایمیل‌های شغلی پیدا‌شده ──
    ea = _read_json("EMAIL_ANALYSIS.json", {}) or {}
    my_emails = [e for e in (ea.get("emails") or [])
                 if _matches_applicant(e, applicant)]
    my_emails.sort(key=lambda e: e.get("date", "") or "", reverse=True)

    # ── ایمیل‌های ارسالی ──
    et = _read_json("EMAIL_TRACKER.json", {}) or {}
    my_sent = [e for e in (et.get("emails") or [])
               if _matches_applicant(e, applicant)]
    tracker_map = _tracker_applicant_map()
    my_notifs = []
    for n in (et.get("notifications") or []):
        if _matches_applicant(n, applicant):
            my_notifs.append(n)
        elif n.get("email_id") in tracker_map and _matches_applicant(
                {"applicant": tracker_map[n["email_id"]]}, applicant):
            my_notifs.append(n)

    # ── یادآورها ──
    # بیشتر یادآورها فیلد متقاضی ندارند و فقط email_id دارند؛
    # صاحبشان را از روی ایمیل متناظر در ترکر پیدا می‌کنیم.
    rm = _read_json("REMINDERS.json", {}) or {}
    my_rem = []
    for r in (rm.get("reminders") or []):
        if _matches_applicant(r, applicant):
            my_rem.append(r)
        elif r.get("email_id") in tracker_map and _matches_applicant(
                {"applicant": tracker_map[r["email_id"]]}, applicant):
            r = dict(r)
            r["_via_email"] = tracker_map[r["email_id"]]
            my_rem.append(r)

    # ── لینکدین ──
    my_linkedin = _load_linkedin(applicant)

    now = _now()
    overdue_rem = []
    for r in my_rem:
        d = _parse_date(r.get("due_date"))
        if d and d < now and r.get("status") not in ("done", "completed"):
            r["_overdue"] = round((now - d).days)
            overdue_rem.append(r)
    overdue_rem.sort(key=lambda r: -r["_overdue"])
    open_rem = [r for r in my_rem if r not in overdue_rem
                and r.get("status") not in ("done", "completed")]

    # ── آمار ──
    open_apps = [a for a in my_apps if a.get("status") in OPEN_STATUSES]
    next_steps = _next_steps(jobs_out, my_apps, overdue_rem, open_rem, my_emails)

    return {
        "applicant": applicant,
        "jobs": jobs_out,
        "applications": sorted(my_apps, key=lambda a: a.get("sent_date") or "", reverse=True),
        "emails": my_emails[:20],
        "sent_emails": my_sent,
        "notifications": my_notifs,
        "reminders": my_rem,
        "overdue": overdue_rem,
        "open_reminders": open_rem,
        "linkedin": my_linkedin,
        "stats": {
            "jobs_matched": len(scored),
            "jobs_shown": len(jobs_out),
            "applied": sum(1 for j in jobs_out if j["applied"]),
            "open_apps": len(open_apps),
            "total_apps": len(my_apps),
            "emails_found": len(my_emails),
            "emails_sent": len(my_sent),
            "reminders": len(my_rem),
            "linkedin": len(my_linkedin),
            "overdue": len(overdue_rem),
        },
        "next_steps": next_steps,
    }


def _next_steps(jobs, apps, overdue_rem, open_rem, emails):
    """
    فهرست «الان چه کار کن» — اولویت‌بندی‌شده.
    این همان چیزی است که کاربر مجبور بود خودش از فایل‌ها درمی‌آورد.
    """
    steps = []

    for r in overdue_rem[:4]:
        steps.append({
            "emoji": "🚨", "urgency": "high",
            "text": f"یادآور «{r.get('title', '؟')}» {r['_overdue']} روز از موعد گذشته",
            "href": "/reminders", "cta": "باز کردن یادآورها",
        })

    for e in emails[:3]:
        cat = (e.get("category") or "").lower()
        if any(k in cat for k in ("interview", "offer", "response")):
            steps.append({
                "emoji": URGENCY_EMOJI.get("INTERVIEW", "📬"), "urgency": "high",
                "text": f"ایمیل «{(e.get('subject') or '')[:60]}» از {e.get('from', '؟')}",
                "href": f"mailto:{e.get('from', '')}", "cta": "باز کردن ایمیل",
            })

    pending = [j for j in jobs if not j["applied"]]
    for j in pending[:3]:
        steps.append({
            "emoji": "🎯", "urgency": "medium",
            "text": f"{j['title'][:70]} — {j['company']} ({j.get('country') or '؟'})",
            "href": j.get("url") or "#", "cta": "دیدن آگهی",
            "job": j,
        })

    for r in open_rem[:3]:
        d = _parse_date(r.get("due_date"))
        when = f" — تا {d.strftime('%Y-%m-%d')}" if d else ""
        steps.append({
            "emoji": "⏰", "urgency": "medium",
            "text": f"یادآور: {r.get('title', '؟')}{when}",
            "href": "/reminders", "cta": "مشاهده",
        })

    if not steps:
        steps.append({
            "emoji": "✅", "urgency": "low",
            "text": "کار بازی برای این متقاضی نمانده — آگهی تطبیقی تازه‌ای هم نیست.",
            "href": "", "cta": "",
        })
    return steps


def overview():
    """
    نمای کلی همهٔ متقاضیان برای صفحهٔ /reports —
    همان چیزی که قبلاً فقط یک کارت خالی بود.
    """
    from config_loader import load_config  # اگر نبود، خالی برمی‌گردانیم
    try:
        applicants = (load_config() or {}).get("applicants", [])
    except Exception:
        try:
            cfg = _read_json(os.path.join("..", "config.json"), {}) or {}
            applicants = cfg.get("applicants", [])
        except Exception:
            applicants = []
    return applicants