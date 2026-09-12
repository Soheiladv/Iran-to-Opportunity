#!/usr/bin/env python3
"""
MigrationHunter — Sponsor-First Crawler (v4)
=============================================
منطق جدید: به جای اسکرپ سایت‌های JS-heavy (Seek/Indeed که جواب نمی‌دهند)،
این موتور می‌رود سراغ منابعی که واقعاً قابل دسترسی‌اند:

  1. Greenhouse / Lever public job APIs  → JSON تمیز، بدون اسکرپ، بدون JS
     (شرکت‌های بین‌المللی که همین امروز آگهی IT/Infra دارند)
  2. صفحات کاریابی با HTML server-rendered + استخراج JSON-LD (schema.org
     JobPosting) — استاندارد واقعی آگهی‌ها
  3. موتور تشخیص آگهی فیک/واقعی با پرچم‌های وزن‌دار + ستون «توضیح صحت‌سنجی»
     در همان فایل اکسل

اجرای:  python sponsor_crawler.py
"""
import os, sys, io, json, re, time, argparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin, quote_plus
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from config_loader import get_applicants, get_applicant_label

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    os.environ["PYTHONIOENCODING"] = "utf-8"

BASE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.join(BASE, "dashboard")
NOW = datetime.now()
DATE_STR = NOW.strftime("%Y-%m-%d %H:%M")
FILE_DATE = NOW.strftime("%Y%m%d_%H%M")

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")

# ═══════════════════════════════════════════════════
# Progress bar (works in cmd / Windows Terminal)
# ═══════════════════════════════════════════════════
def progress(current, total, label="", width=28):
    if total <= 0:
        return
    frac = min(current / total, 1.0)
    filled = int(frac * width)
    bar = "█" * filled + "░" * (width - filled)
    sys.stdout.write(f"\r  ▶ [{bar}] {current}/{total}  {label[:40]:<40}")
    sys.stdout.flush()
    if current >= total:
        sys.stdout.write("\n")
        sys.stdout.flush()


def fetch_json(url, timeout=10):
    req = Request(url, headers={"User-Agent": UA, "Accept": "application/json"})
    try:
        with urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8", "ignore"))
    except Exception:
        return None


def fetch_html(url, timeout=10):
    req = Request(url, headers={"User-Agent": UA, "Accept": "text/html"})
    try:
        with urlopen(req, timeout=timeout) as r:
            return r.read().decode("utf-8", "ignore")
    except Exception:
        return None


# ═══════════════════════════════════════════════════
# Greenhouse public boards (structured JSON, no scraping)
# Docs: https://developers.greenhouse.io/job-board.html
# ═══════════════════════════════════════════════════
GREENHOUSE_COMPANIES = [
    "databricks", "stripe", "shopify", "canva", "atlassian", "figma",
    "klaviyo", "deliveroo", "monzo", "revolut", "wise", "n26",
    "hellofresh", "doctolib", "babylonhealth",
]

# Lever public boards: https://api.lever.co/v0/postings/{company}?mode=json
LEVER_COMPANIES = [
    "netflix", "palantir", "spotify", "klarna", "pluckey", "missionsquare",
]


def fetch_greenhouse(company):
    """Greenhouse board API → list of job dicts."""
    data = fetch_json(f"https://boards-api.greenhouse.io/v1/boards/{company}/jobs")
    if not data or not data.get("jobs"):
        return []
    jobs = []
    for j in data["jobs"]:
        title = j.get("title", "")
        loc = j.get("location", {}).get("name", "")
        jobs.append({
            "title": title,
            "company": company.replace("-", " ").title(),
            "location": loc,
            "country": country_from_location(loc),
            "source": f"Greenhouse ({company})",
            "url": j.get("absolute_url", ""),
            "description": j.get("content", ""),  # base64-ish html
            "posted": j.get("updated_at", ""),
        })
    return jobs


def fetch_lever(company):
    data = fetch_json(f"https://api.lever.co/v0/postings/{company}?mode=json")
    if not data:
        return []
    jobs = []
    for j in data:
        loc = (j.get("categories") or {}).get("location", "")
        jobs.append({
            "title": j.get("text", ""),
            "company": company.replace("-", " ").title(),
            "location": loc,
            "country": country_from_location(loc),
            "source": f"Lever ({company})",
            "url": (j.get("hostedUrl") or j.get("applyUrl") or ""),
            "description": strip_html(j.get("descriptionPlain") or ""),
            "posted": "",
        })
    return jobs


# ═══════════════════════════════════════════════════
# Government / server-rendered boards + JSON-LD parsing
# ═══════════════════════════════════════════════════
GOV_SEARCHES = [
    # (نام, کشور, url، آیا JSON-LD دارد)
    ("Job Bank Canada", "CA",
     "https://www.jobbank.gc.ca/jobsearch/jobsearch?searchstring={kw}", True),
    ("Australian JobSearch", "AU",
     "https://jobsearch.gov.au/jobs?keywords={kw}", True),
    ("EURES (EU portal)", "EU",
     "https://ec.europa.eu/eures/eures-searchengine/page/jobs-search?keywordsEverywhere={kw}", True),
    ("Make it in Germany job board", "DE",
     "https://www.make-it-in-germany.com/en/working-in-germany/job-listings?search={kw}", True),
    ("Workforce Australia", "AU",
     "https://www.workforceaustralia.gov.au/individuals/coi/search-results?page=1&q={kw}", True),
    ("Working In NZ (gov endorsed)", "NZ",
     "https://www.workingin-newzealand.com/jobs?search={kw}", False),
    ("YAxes NZ (Accredited Employers)", "NZ",
     "https://www.yaxes.com/jobs?search={kw}", False),
]


# کشورها به چند زبان/الگوی رایج در location فیلدهای ATS
COUNTRY_PATTERNS = [
    # (کشور, الگوهای regex — اولین تطبیق برنده است)
    ("US", [r"\bunited\s*states\b", r"\bu\.?s\.?a\.?\b", r"\b(usa|us)\b",
            r"\balabama\b|\balaska\b|\barizona\b|\barkansas\b|\bcalifornia\b|\bcolorado\b|"
            r"\bconnecticut\b|\bdelaware\b|\bflorida\b|\bgeorgia\b|\bhawaii\b|\bidaho\b|"
            r"\billinois\b|\bindiana\b|\biowa\b|\bkansas\b|\bkentucky\b|\blouisiana\b|"
            r"\bmaine\b|\bmaryland\b|\bmassachusetts\b|\b michigan\b|\bminnesota\b|"
            r"\bmississippi\b|\bmissouri\b|\bmontana\b|\bnebraska\b|\bnevada\b|"
            r"\bnew hampshire\b|\bnew jersey\b|\bnew mexico\b|\bnew york\b|\bnorth carolina\b|"
            r"\bnorth dakota\b|\bohio\b|\boklahoma\b|\boregon\b|\bpennsylvania\b|"
            r"\brhode island\b|\bsouth carolina\b|\bsouth dakota\b|\btennessee\b|\btexas\b|"
            r"\butah\b|\bvermont\b|\bvirginia\b|\bwashington(?!\s+dc)|\bwest virginia\b|"
            r"\bwisconsin\b|\bwyoming\b",
            r",\s*(al|ak|az|ar|ca|co|ct|de|fl|ga|hi|id|il|in|ia|ks|ky|la|me|md|ma|mi|mn|"
            r"ms|mo|mt|ne|nv|nh|nj|nm|ny|nc|nd|oh|ok|or|pa|ri|sc|sd|tn|tx|ut|vt|va|wa|wv|wi|wy)\b(,|\s|$)",
            # شهرهای بزرگ آمریکا
            r"\b(san francisco|new york city|los angeles|chicago|boston|seattle|austin|"
            r"denver|atlanta|dallas|houston|mountain view|palo alto|menlo park|san jose|"
            r"san diego|portland|phoenix|philadelphia|detroit|minneapolis|bellevue|"
            r"newark|sunnyvale|santa clara|redmond|salt lake city)\b",
            # مخفف‌های رایج شهرها در ATS
            r"\b(nyc|sf|n\.?y\.?c\.?)\b",
            # مخفف‌های شهری دوحرفی
            r"\b(chi|atl)\b", r"\busca\b"]),
    ("CA", [r"\bcanada\b", r"\b(alberta|british columbia|manitoba|new brunswick|"
            r"newfoundland|nova scotia|ontario|prince edward island|quebec|saskatchewan)\b",
            r",\s*(ab|bc|mb|nb|nl|ns|on|pe|qc|sk)\b(,|\s|$)",
            r"\b(toronto|vancouver|montreal|calgary|ottawa|edmonton|winnipeg)\b"]),
    ("UK", [r"\b(united\s*kingdom|england|scotland|wales|northern ireland)\b",
            r"\b(london|manchester|birmingham|edinburgh|glasgow|leeds|bristol|"
            r"cambridge|oxford|reading|milton keynes)\b"]),
    ("DE", [r"\b(germany|deutschland)\b", r"\b(berlin|munich|münchen|hamburg|frankfurt|"
            r"cologne|köln|stuttgart|düsseldorf|leipzig|dresden|nuremberg|hannover|"
            r"münster|osnabrück|mainz|würzburg|bielefeld)\b"]),
    ("AU", [r"\baustralia\b", r"\b(new south wales|victoria|queensland|western australia|"
            r"south australia|tasmania|australian capital territory)\b",
            r",\s*(nsw|vic|qld|wa|sa|tas|act|nt)\b(,|\s|$)",
            r"\b(sydney|melbourne|brisbane|perth|adelaide|canberra)\b"]),
    ("NZ", [r"\bnew\s*zealand\b", r"\b(auckland|wellington|christchurch|hamilton)\b"]),
    ("TH", [r"\bthailand\b", r"\bbangkok\b"]),
    ("IE", [r"\bireland\b", r"\b(dublin|cork|galway|limerick)\b", r"\bdub\b"]),
    ("NL", [r"\b(netherlands|holland)\b", r"\b(amsterdam|rotterdam|the hague|den haag|"
            r"utrecht|eindhoven|groningen)\b"]),
    ("FR", [r"\bfrance\b", r"\b(paris|lyon|marseille|toulouse|bordeaux|nantes|nice|"
            r"le mans|poitiers|tours|annecy|lille)\b"]),
    ("ES", [r"\bspain\b", r"\b(madrid|barcelona|valencia|seville|sevilla|malaga)\b"]),
    ("IT", [r"\bitaly\b|\bitalia\b", r"\b(rome|roma|milan|milano|turin|torino|naples|napoli)\b"]),
    ("SE", [r"\bsweden\b", r"\b(stockholm|gothenburg|göteborg|malmo|malmö)\b"]),
    ("NO", [r"\bnorway\b", r"\b(oslo|bergen|trondheim)\b"]),
    ("DK", [r"\bdenmark\b|\bdanmark\b", r"\b(copenhagen|københavn|aarhus|odense)\b"]),
    ("FI", [r"\bfinland\b|\bsuomi\b", r"\b(helsinki|espoo|tampere|turku)\b"]),
    ("CH", [r"\bswitzerland\b|\bschweiz\b|\bsuisse\b", r"\b(zurich|zürich|geneva|genève|"
            r"basel|bern|lausanne|zug)\b"]),
    ("AT", [r"\baustria\b|\bösterreich\b", r"\b(vienna|wien|graz|linz|salzburg)\b"]),
    ("BE", [r"\bbelgium\b|\bbelgië\b|\bbelgique\b", r"\b(brussels|bruxelles|antwerp|"
            r"antwerpen|ghent|gent|bruges|brugge|leuven)\b"]),
    ("LU", [r"\bluxembourg\b"]),
    ("PT", [r"\bportugal\b", r"\b(lisbon|lisboa|porto|oporto)\b"]),
    ("PL", [r"\bpoland\b|\bpolska\b", r"\b(warsaw|warszawa|krakow|kraków|wroclaw|wrocław|"
            r"poznan|poznań|gdansk|gdańsk)\b"]),
    ("CZ", [r"\bczech(\s*republic)?\b|\bčeská\b", r"\bprague\b|\bpraha\b"]),
    ("SG", [r"\bsingapore\b"]),
    ("JP", [r"\bjapan\b|\b日本\b", r"\b(tokyo|osaka|kyoto|yokohama)\b"]),
    ("CN", [r"\bchina\b|\b中国\b", r"\b(beijing|shanghai|shenzhen|guangzhou|hangzhou)\b"]),
    ("IN", [r"\bindia\b|\bbharat\b", r"\b(bengaluru|bangalore|mumbai|delhi|hyderabad|"
            r"chennai|pune|gurgaon|gurugram|noida|kolkata)\b"]),
    ("PH", [r"\bphilippines\b", r"\b(manila|quezon city|cebu)\b"]),
    ("MX", [r"\bmexico\b|\bméxico\b", r"\b(mexico city|ciudad de méxico|guadalajara|"
            r"monterrey)\b"]),
    ("BR", [r"\bbrazil\b|\bbrasil\b", r"\b(são paulo|sao paulo|rio de janeiro)\b"]),
    ("AE", [r"\bunited arab emirates\b|\bu\.?a\.?e\.?\b", r"\b(dubai|abu dhabi|sharjah)\b"]),
    ("ZA", [r"\bsouth africa\b", r"\b(cape town|johannesburg|durban)\b"]),
    ("IL", [r"\bisrael\b", r"\b(tel aviv|jerusalem|haifa)\b"]),
    ("RS", [r"\bserbia\b", r"\b(belgrade|beograd|novi sad)\b"]),
    ("RO", [r"\bromania\b", r"\b(bucharest|bucuresti|cluj)\b"]),
    ("GR", [r"\bgreece\b", r"\b(athens|athina|thessaloniki)\b"]),
    ("TR", [r"\bturkey\b|\btürkiye\b", r"\bistanbul\b|\bankara\b|\bizmir\b"]),
    ("KR", [r"\bsouth korea\b|\bkorea\b", r"\bseoul\b|\bbusan\b"]),
    ("TW", [r"\btaiwan\b", r"\btaipei\b"]),
    ("ID", [r"\bindonesia\b", r"\b(jakarta|surabaya|bandung)\b"]),
    ("SA", [r"\bsaudi arabia\b", r"\b(riyadh|jeddah|dammam)\b"]),
    ("VN", [r"\bvietnam\b", r"\b(ho chi minh|hanoi)\b"]),
    ("CR", [r"\bcosta rica\b"]),
    ("Remote", [r"\bremote\b", r"\bwork from home\b", r"\banywhere\b"]),
]

# کلماتی که در انتهای location یعنی چند مکان/نامشخص است
AMBIGUOUS_LOCATIONS = {"n/a", "na", "unknown", "various", "multiple", "", "-"}


def country_from_location(loc):
    """
    استخراج کشور از فیلد location با پارس چندالگویی.
    اولویت: نام کامل کشور > ایالت/ناحیه > شهر بزرگ.
    اگر چند کشور مطابقت خورد، اولویت با تطبیق «نام کامل کشور» است
    و در غیر این صورت اولین تطبیق. اگر هیچ چیز پیدا نشد: "؟"
    """
    loc_l = (loc or "").strip().lower()
    if loc_l in AMBIGUOUS_LOCATIONS:
        return "؟"

    # یک پاس: نام‌های کامل کشور اولویت دارند، بعد نواحی، بعد شهرها
    best_full, best_region, best_city = None, None, None
    for country, patterns in COUNTRY_PATTERNS:
        if country == "Remote":
            continue
        for i, pat in enumerate(patterns):
            if re.search(pat, loc_l, re.I):
                if i == 0 and best_full is None:
                    best_full = country
                elif i > 0 and best_region is None:
                    best_region = country
                break
    # Remote فقط اگر هیچ کشور دیگری مطابق نشد
    if best_full is None and best_region is None:
        for country, patterns in COUNTRY_PATTERNS:
            if country != "Remote":
                continue
            for pat in patterns:
                if re.search(pat, loc_l, re.I):
                    return "Remote"
        return "؟"

    return best_full or best_region


def strip_html(text):
    return re.sub(r"<[^>]+>", " ", text or "").strip()


def extract_jsonld_jobs(html, source_name):
    """استخراج آگهی از markup استاندارد schema.org/JobPosting."""
    jobs = []
    for m in re.finditer(
        r'<script[^>]*type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        html, re.S | re.I
    ):
        try:
            data = json.loads(m.group(1).strip())
        except Exception:
            continue
        items = data if isinstance(data, list) else data.get("@graph", [data])
        for item in items:
            if not isinstance(item, dict):
                continue
            if item.get("@type") not in ("JobPosting", ["JobPosting"]):
                continue
            org = item.get("hiringOrganization") or {}
            company = org.get("name", "نامشخص") if isinstance(org, dict) else str(org)
            loc_obj = item.get("jobLocation") or {}
            if isinstance(loc_obj, list):
                loc_obj = loc_obj[0] if loc_obj else {}
            addr = (loc_obj.get("address") or {}) if isinstance(loc_obj, dict) else {}
            loc = " ".join(filter(None, [
                addr.get("addressLocality") or "",
                addr.get("addressRegion") or "",
                addr.get("addressCountry") or "",
            ])).strip()
            jobs.append({
                "title": item.get("title") or item.get("name", "بدون عنوان"),
                "company": company,
                "location": loc,
                "country": country_from_location(str(addr.get("addressCountry", loc))),
                "source": source_name,
                "url": item.get("url") or item.get("@id", ""),
                "description": strip_html(str(item.get("description", "")))[:2000],
                "posted": item.get("datePosted", ""),
            })
    return jobs


def extract_fallback_jobs(html, source_name, base_url):
    """Fallback: لینک‌هایی که شبیه آگهی واقعی‌اند (نه nav)."""
    jobs = []
    seen = set()
    for m in re.finditer(r'<a[^>]+href="([^"]+)"[^>]*>([^<]{12,120})</a>', html, re.I):
        href, text = m.group(1), m.group(2).strip()
        if not re.search(r"job|vacan|position|careers?|posting|anzeige", href, re.I):
            continue
        if href.startswith(("#", "javascript:", "mailto:")):
            continue
        href = urljoin(base_url, href)
        key = href.split("?")[0]
        if key in seen:
            continue
        seen.add(key)
        jobs.append({
            "title": text,
            "company": "نامشخص",
            "location": "",
            "country": "؟",
            "source": source_name,
            "url": href,
            "description": "",
            "posted": "",
        })
    return jobs


def fetch_gov_source(name, country, url_tpl, try_jsonld, keyword):
    url = url_tpl.format(kw=quote_plus(keyword))
    html = fetch_html(url)
    if not html:
        return []
    jobs = extract_jsonld_jobs(html, name) if try_jsonld else []
    if not jobs:
        jobs = extract_fallback_jobs(html, name, url)
    for j in jobs:
        if j.get("country") in ("؟", ""):
            j["country"] = country
    return jobs


# ═══════════════════════════════════════════════════
# موتور صحت‌سنجی: تشخیص واقعی/فیک با پرچم‌های وزن‌دار
# ═══════════════════════════════════════════════════
RED_FLAGS = [
    (25, r"(application|processing|registration|visa)\s+(fee|payment|charge)",
     "درخواست پرداخت وجه"),
    (20, r"(western\s*union|money\s*gram|bitcoin|crypto|usdt)",
     "روش پرداخت غیرقابل پیگیری"),
    (15, r"@(gmail|yahoo|hotmail|outlook|aol|protonmail)\.",
     "ایمیل تماس رایگان (نه دامنه شرکتی)"),
    (15, r"(guaranteed|100%)\s*(visa|job|offer|work\s*permit)",
     "ادعای تضمینی بودن ویزا/شغل"),
    (12, r"(telegram|whatsapp)\.?(me|com)?[/\s]",
     "تماس فقط از طریق تلگرام/واتس‌اپ"),
    (10, r"no\s+(experience|qualification|interview)\s+(needed|required)",
     "بدون نیاز به تجربه/مصاحبه"),
    (10, r"(\$|\€|\£)?\s?(\d{2,3}[,.]?\d{3})\s*[-–]\s*(\$|\€|\£)?\s?\d{2,3}[,.]?\d{3}\s*(per\s*month|monthly)",
     "حقوق غیرمعمول بالا (ماهانه)"),
    (8, r"work\s*from\s*home.{0,40}(no\s*experience|immediate|daily\s*payout)",
     "الگوی دورکاری-درآمد سریع"),
]

GREEN_FLAGS = [
    (25, "gov|jobbank|jobsearch|arbeitsagentur|make-it-in-germany|eures|immigration",
     "منبع رسمی دولتی"),
    (20, r"(visa\s*sponsorship|work\s*visa|sponsor(s|hip)?\s*(available|provided)?|"
         r"relocation\s*(package|support|assistance))",
     "اشاره صریح به اسپانسری/جریمه‌ها به‌عنوان 'اسپانسری/جابجایی'"),
    (18, "greenhouse|lever", "پلتفرم رسمی ATS (Greenhouse/Lever)"),
    (12, r"(accredited|licensed)\s*(employer|sponsor)",
     "کارفرمای دارای مجوز/اسپانسری رسمی"),
    (10, r"(b\.?sc|m\.?sc|bachelor|master|degree|diploma|certification)",
     "ذکر مدرک تحصیلی موردنیاز"),
    (8, r"(registered|licensed|certified)\s+(midwife|nurse|engineer|nmc|ahpra|anzsco)",
     "ذکر ثبت حرفه‌ای"),
]


def verify_job(job):
    """
    امتیازدهی صحت‌سنجی.
    برخلاف نسخه قبلی (که فقط keyword match می‌کرد)، اینجا پرچم‌های
    واقعی کلاهبرداری و سیگنال‌های اعتبار بررسی می‌شوند و
    «توضیح» کامل برمی‌گردد تا در ستون Excel درج شود.
    """
    text = " ".join([
        job.get("title", ""), job.get("description", ""),
        job.get("url", ""), job.get("company", ""),
    ]).lower()

    red_hits, green_hits = [], []
    red_score = green_score = 0

    for weight, pattern, label in RED_FLAGS:
        if re.search(pattern, text, re.I):
            red_hits.append(f"🔴 {label} (+{weight} risk)")
            red_score += weight

    for weight, pattern, label in GREEN_FLAGS:
        try:
            if re.search(pattern, text, re.I) or pattern in text:
                green_hits.append(f"🟢 {label} (+{weight})")
                green_score += weight
        except re.error:
            if pattern in text:
                green_hits.append(f"🟢 {label} (+{weight})")
                green_score += weight

    net = green_score - red_score
    if net >= 25 and red_score <= 15:
        verdict = "✅ قابل اعتماد"
    elif net >= 10:
        verdict = "🟡 نیازمند بررسی"
    elif red_score >= 25:
        verdict = "❌ مشکوک به کلاهبرداری"
    else:
        verdict = "⚪ خنثی/کم‌اطلاعات"

    explanation = "; ".join(green_hits + red_hits) if (green_hits or red_hits) \
        else "هیچ سیگنال مثبت یا منفی‌ای یافت نشد"

    return {
        "verdict": verdict,
        "net_score": net,
        "red": red_score,
        "green": green_score,
        "explanation": explanation,
    }


# ═══════════════════════════════════════════════════
# Match score بر اساس پروفایل متقاضی
# ═══════════════════════════════════════════════════
def match_score(job, applicants):
    title = (job.get("title") or "").lower()
    desc = (job.get("description") or "").lower()
    best_applicant, best = None, 0
    for a in applicants:
        s = sum(2 for k in a.get("keywords", []) if k.lower() in title) \
            + sum(1 for k in a.get("keywords", []) if k.lower() in desc[:1000])
        if s > best:
            best, best_applicant = s, a
    return best, best_applicant


# ═══════════════════════════════════════════════════
# ساخت Excel با ستون صحت‌سنجی
# ═══════════════════════════════════════════════════
def build_excel(jobs, applicants):
    wb = Workbook()
    ws = wb.active
    ws.title = "فرصت‌های شکار شده"
    ws.sheet_view.rightToLeft = True

    thin = Border(left=Side("thin"), right=Side("thin"),
                  top=Side("thin"), bottom=Side("thin"))
    hfont = Font(size=9, bold=True, color="FFFFFF")
    hfill = PatternFill("solid", start_color="1B4F72")
    cfont = Font(size=9)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    headers = ["#", "عنوان شغل", "شرکت", "کشور", "منبع", "لینک", "تاریخ",
               "امتیاز تطبیق", "متقاضی پیشنهادی", "صحت‌سنجی",
               "توضیح صحت‌سنجی (چرا؟)"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font, c.fill, c.alignment, c.border = hfont, hfill, center, thin

    widths = [5, 40, 22, 10, 22, 35, 16, 10, 18, 20, 70]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i) if i <= 26 else "A"].width = w

    for idx, (job, v, ms, label) in enumerate(jobs, 1):
        row_vals = [idx, job["title"][:80], job["company"][:60],
                    job["country"], job["source"][:30], job["url"][:60],
                    DATE_STR, ms, label, v["verdict"], v["explanation"]]
        r = idx + 1
        for ci, val in enumerate(row_vals, 1):
            c = ws.cell(r, ci, val)
            c.font, c.border = cfont, thin
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if ci == 10:
                if "✅" in val:
                    c.fill = PatternFill("solid", start_color="C6EFCE")
                elif "❌" in val:
                    c.fill = PatternFill("solid", start_color="FFC7CE")
                elif "🟡" in val:
                    c.fill = PatternFill("solid", start_color="FFEB9C")

    ws.auto_filter.ref = f"A1:K{len(jobs) + 1}"
    return wb


# ═══════════════════════════════════════════════════
# main
# ═══════════════════════════════════════════════════
KEYWORDS = ["midwife", "IT manager", "systems administrator",
            "infrastructure engineer", "network engineer"]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true",
                    help="فقط ۳ کلیدواژه و یک منبع دولتی برای تست سریع")
    args = ap.parse_args()

    print("═" * 62)
    print("MigrationHunter v4 — Sponsor-First Crawler (parallel)")
    print(f"📅 {DATE_STR}")
    print("═" * 62)
    applicants = get_applicants()
    all_jobs = []

    keywords = KEYWORDS[:2] if args.quick else KEYWORDS
    gov_list = GOV_SEARCHES[:1] if args.quick else GOV_SEARCHES

    # همه‌ی کارها موازی: Greenhouse + Lever + Gov
    futures = {}
    done = 0
    total = len(GREENHOUSE_COMPANIES) + len(LEVER_COMPANIES) \
        + len(keywords) * len(gov_list)

    # --- مرحله ۱+۲: ATS APIs ---
    print("\n[1/2] 🏢 Greenhouse/Lever boards (public JSON APIs, parallel)")
    with ThreadPoolExecutor(max_workers=12) as ex:
        for co in GREENHOUSE_COMPANIES:
            futures[ex.submit(fetch_greenhouse, co)] = f"Greenhouse:{co}"
        for co in LEVER_COMPANIES:
            futures[ex.submit(fetch_lever, co)] = f"Lever:{co}"

        # --- مرحله ۳: منابع دولتی ---
        print("[2/2] 🏛️ Government boards (JSON-LD / server-rendered, parallel)")
        for kw in keywords:
            for (name, cc, tpl, jld) in gov_list:
                futures[ex.submit(fetch_gov_source, name, cc, tpl, jld, kw)] = \
                    f"{name}×{kw}"

        for fut in as_completed(futures):
            label = futures[fut]
            try:
                jobs = fut.result(timeout=30)
            except Exception:
                jobs = []
            all_jobs.extend(jobs)
            done += 1
            progress(done, total, f"{label}: {len(jobs)} jobs")

    # --- Dedup ---
    seen, unique = set(), []
    for j in all_jobs:
        k = (j["title"][:60].lower(), j["company"].lower()[:30],
             j["source"].split(" (")[0])
        if k not in seen:
            seen.add(k)
            unique.append(j)
    print(f"\n📥 مجموع خام: {len(all_jobs)} | یکتا: {len(unique)}")

    # --- Score + verify ---
    scored = []
    for j in unique:
        v = verify_job(j)
        ms, best = match_score(j, applicants)
        label = get_applicant_label(best["id"]) if best else "—"
        scored.append((j, v, ms, label))
    scored.sort(key=lambda t: (t[2], t[1]["net_score"]), reverse=True)

    # --- Excel ---
    os.makedirs(DASH, exist_ok=True)
    fn = f"Sponsor_Hunt_{FILE_DATE}.xlsx"
    wb = build_excel(scored, applicants)
    wb.save(os.path.join(DASH, fn))

    # --- JSON memory ---
    os.makedirs(os.path.join(BASE, "memory"), exist_ok=True)
    with open(os.path.join(BASE, "memory", "SPONSOR_RESULTS.json"), "w",
              encoding="utf-8") as f:
        json.dump({"found": len(scored), "at": DATE_STR,
                   "jobs": [j for j, *_ in scored]}, f,
                  ensure_ascii=False, indent=2)

    # --- Summary ---
    print(f"\n📊 خلاصه")
    trusted = sum(1 for _, v, *_ in scored if "✅" in v["verdict"])
    scam = sum(1 for _, v, *_ in scored if "❌" in v["verdict"])
    print(f"  ├─ کل: {len(scored)} | ✅ قابل اعتماد: {trusted} | "
          f"❌ مشکوک: {scam}")
    countries = {}
    for j, *_ in scored:
        countries[j["country"]] = countries.get(j["country"], 0) + 1
    for c, n in sorted(countries.items(), key=lambda x: -x[1])[:8]:
        print(f"  │  {c}: {n}")
    print(f"  └─ 📁 {fn}")
    print("═" * 62)


if __name__ == "__main__":
    main()
