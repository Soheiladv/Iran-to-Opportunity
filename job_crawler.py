#!/usr/bin/env python3
"""
MigrationHunter — جستجوی خودکار کاریابی
جستجو در سایت‌های مختلف + ذخیره در Excel + تشخیص واقعی/فیک آگهی
اجرای: python job_crawler.py
"""
import os, sys, json, re, io, time
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError
from urllib.parse import urljoin, urlparse, parse_qs
from html.parser import HTMLParser

try:
    from playwright.sync_api import sync_playwright  # رندر جاوااسکریپت — اختیاری
    HAS_PLAYWRIGHT = True
except ImportError:
    HAS_PLAYWRIGHT = False

# مرورگر headless — اول Chromium، اگر نبود Edge/Chrome ویندوز (نیازی به دانلود ندارد)
#
# نکته: دانلود Chromium از CDN پلی‌رایت در بعضی کشورها ۴۰۳ می‌خورد، ولی
# Edge روی هر ویندوزی هست. کانالی که جواب داد را کش می‌کنیم تا هر صفحه
# سه بار تلاش نکنیم.
_PW = None       # context playwright — یک‌بار start، در _close_browser متوقف می‌شود
_BROWSER = None  # نگه‌داشتن مرورگر بین منابع — یک‌بار launch می‌شود
_BROWSER_CHANNEL = None  # کانالی که واقعاً جواب داد
_BROWSER_PROBE = None    # نتیجهٔ تست در این thread (None = هنوز تست نشده)


def _launch_browser(playwright):
    global _BROWSER_CHANNEL, _HAS_ANY_BROWSER
    channels = ([_BROWSER_CHANNEL] if _BROWSER_CHANNEL else []) + ["chromium", "msedge", "chrome"]
    seen = set()
    for ch in channels:
        if ch in seen:
            continue
        seen.add(ch)
        try:
            browser = playwright.chromium.launch(
                headless=True, **({"channel": ch} if ch != "chromium" else {}))
            _BROWSER_CHANNEL = ch
            _HAS_ANY_BROWSER = True
            return browser
        except Exception:
            continue
    _HAS_ANY_BROWSER = False
    return None


def browser_available():
    """آیا مرورگر headless روی همین thread قابل استفاده است؟

    sync_playwright به یک thread قفل می‌شود: اگر اولین بار در یک thread کارگر
    باز شود، همان‌جا می‌ماند و در thread بعدی «Cannot switch to a different
    thread» می‌دهد. برای همین بررسی را سراسری نگه نمی‌داریم و هر بار در
    thread فعلی امتحان می‌کنیم.
    """
    if not HAS_PLAYWRIGHT:
        return False
    global _BROWSER_PROBE
    if _PW is not None:
        return True
    if _BROWSER_PROBE is not None:
        return _BROWSER_PROBE
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            if _launch_browser(pw) is None:
                _BROWSER_PROBE = False
            else:
                _BROWSER_PROBE = True
    except Exception:
        _BROWSER_PROBE = False
    return bool(_BROWSER_PROBE)


def browser_hint():
    """پیام صادقانه برای وقتی مرورگر headless در دسترس نیست."""
    if HAS_PLAYWRIGHT and not browser_available():
        return ("مرورگر headless در دسترس نیست — `python -m playwright install chromium` "
                "را بزن یا Edge/Chrome نصب باشد")
    return "مسدودسازی سایت (۴۰۳) یا محتوای JS که رندر نشده"


def _pw_fetch(url, timeout_ms=25000):
    """دریافت HTML رندرشده با مرورگر headless — برای منابع needs_js و مسدودشده.

    نکته: page در بلوک finally بسته می‌شود تا در timeout نشت نکند (نشت صفحه
    باعث قفل‌شدن sync_playwright و هنگ‌کردن بی‌نهایت کراولر می‌شود).
    اگر به هر دلیلی (مثل اجرا در thread کارگر) نشد، None برمی‌گرداند و
    فراخوان به HTTP ساده می‌افتد.
    """
    global _PW, _BROWSER
    if not HAS_PLAYWRIGHT:
        return None
    try:
        from playwright.sync_api import sync_playwright
        if _PW is None:
            _PW = sync_playwright().start()
        if _BROWSER is None or not _BROWSER.is_connected():
            _BROWSER = _launch_browser(_PW)
        if _BROWSER is None:
            return None
        page = _BROWSER.new_page(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            viewport={"width": 1366, "height": 900},
            locale="en-US",
        )
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
            try:
                page.wait_for_load_state("networkidle", timeout=5000)
            except Exception:
                pass
            html = page.content()
            return html if html and len(html) > 500 else None
        finally:
            try:
                page.close()
            except Exception:
                pass
    except Exception:
        return None


def _close_browser():
    global _PW, _BROWSER
    try:
        if _BROWSER is not None:
            _BROWSER.close()
    except Exception:
        pass
    try:
        if _PW is not None:
            _PW.stop()
    except Exception:
        pass
    _BROWSER = _PW = None

# Fix Windows console encoding
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

BASE = os.path.dirname(os.path.abspath(__file__))
MEM = os.path.join(BASE, "memory")
DASH = os.path.join(BASE, "dashboard")
OUT = os.path.join(BASE, "output")

NOW = datetime.now()
DATE_STR = NOW.strftime("%Y-%m-%d %H:%M")
FILE_DATE = NOW.strftime("%Y%m%d_%H%M")

SOURCES_PATH = os.path.join(BASE, "sources.json")
DISCOVERED_PATH = os.path.join(MEM, "discovered_sources.json")
YIELD_HISTORY_PATH = os.path.join(MEM, "SOURCE_YIELD_HISTORY.json")

# ══════════════════════════════════════════════════════════════════
# progress سراسری — روی «تک‌تک آدرس‌ها» جلو می‌رود، نه هر سایت یک‌بار
#
# چرا روی URL و نه روی سایت؟ چون هر سایت چند آدرس دارد و هر آدرس
# جداگانه یک صفحه‌ی سنگین باز می‌کند. اگر progress را فقط با «شمار
# سایت‌های تمام‌شده» حساب کنیم، نوار یک‌هو از ۳٪ به ۹٪ می‌پرد و کاربر
# فکر می‌کند هنگ کرده. با شمارش URLها نوار پیوسته و واقعی حرکت می‌کند.
#
# P کلیدهایی دارد که web_ui.py از روی لاگ پر می‌کند:
#   url_done / url_total  → درصد واقعی
#   sites_visited         → جدول زندهٔ هر آدرسی که باز شد
# ══════════════════════════════════════════════════════════════════
P = {
    # شمارش روی URL
    "total": 0, "done": 0, "pct": 0,
    # شمارس روی سایت (فقط برای برچسب «منبع ۳ از ۲۵»)
    "site_total": 0, "site_done": 0,
    "current_source": "", "current_url": "", "current_keyword": "",
    "jobs_found_total": 0,
    "sites_visited": [],  # [{name, url, keyword, jobs, ms, status}] — به‌ترتیب
}

# پروتکل لاگ — web_ui.py این قالب‌ها را با regex می‌خواند.
# اگر این قالب‌ها را عوض کنی، همان regex ها را در web_ui هم عوض کن.
# نکته: 🏁 برای «پایان کل منبع» جدا از ✅ («پایان یک آدرس») است — هر دو با
# ✅ بودند و پنل زنده هر منبع را دوبار می‌شمرد.
LG_SOURCE = "📡"   # شروع یک منبع
LG_FETCH  = "🔎"   # شروع یک آدرس
LG_DONE   = "✅"   # یک آدرس تمام شد
LG_FAIL   = "⚠️"   # یک آدرس جواب نداد
LG_SITE   = "🏁"   # پایان کامل یک منبع (تعداد تجمیعی)


def emit(kind, payload):
    """یک رویداد progress را چاپ می‌کند — تنها نقطهٔ خروج لاگ زنده."""
    print(f"{kind} {payload}", flush=True)


def progress_note(msg):
    print(f"    ⏳ {msg}", flush=True)


def show_progress():
    """نوار progress واقعی — مخرجش تعداد کل آدرس‌هاست، نه تعداد سایت‌ها."""
    w = 30
    filled = int(w * P["pct"] / 100) if P["pct"] > 0 else 0
    bar = "█" * filled + "░" * (w - filled)
    print(f"  [{bar}] آدرس {P['done']}/{P['total']} ({P['pct']}%)"
          f" · سایت {P['site_done']}/{P['site_total']}"
          f" · 🧲 {P['jobs_found_total']} آگهی", flush=True)


def update_progress(jobs=0):
    """بعد از هر آدرسِ تمام‌شده صدا زده می‌شود."""
    P["done"] += 1
    P["jobs_found_total"] += jobs
    if P["total"] > 0:
        P["pct"] = round(P["done"] * 100 / P["total"], 1)
    show_progress()


def plan_progress(ordered):
    """قبل از شروع، کل کار را می‌شمارد تا نوار از همان ابتدا مخرج درست دارد."""
    P["site_total"] = len(ordered)
    P["site_done"] = 0
    P["total"] = sum(max(1, len(s.get("search_urls") or [s["url"]])) for s in ordered)
    P["done"] = 0
    P["pct"] = 0
    P["jobs_found_total"] = 0
    P["sites_visited"] = []


def record_url(source_name, url, keyword, jobs, ms, status):
    """یک ردیف برای جدول زندهٔ آدرس‌ها نگه می‌دارد."""
    row = {
        "name": source_name, "url": url, "keyword": keyword,
        "jobs": jobs, "ms": ms, "status": status,
    }
    P["sites_visited"].append(row)
    return row

# ════════════════════════════════════════════════════
# بانک منابع — مرجع واحد: sources.json (هیچ لیست هاردکد)
# ════════════════════════════════════════════════════

def validate_source(s, idx=0):
    """اعتبارسنجی سادهٔ هر منبع — منابع خراب حذف می‌شوند."""
    if not isinstance(s, dict):
        return None
    name = str(s.get("name", "")).strip()
    url = str(s.get("url", "")).strip()
    if not name or not url.startswith("http"):
        return None
    s["name"] = name
    s["url"] = url
    s.setdefault("country", "؟")
    s.setdefault("type", "job_board")
    s.setdefault("enabled", True)
    s.setdefault("trust", 60)
    # مسیر: job (آگهی شغلی) یا education (برنامهٔ تحصیلی) — هیچ چیزی هاردکد نیست
    track = str(s.get("track", "job")).strip().lower()
    s["track"] = track if track in ("job", "education") else "job"
    urls = s.get("search_urls") or []
    s["search_urls"] = [u for u in urls if isinstance(u, str) and u.startswith("http")] or [url]
    return s


def load_sources(track=None, country=None, include_disabled=False):
    """منابع را از sources.json می‌خواند — مرجع واحد. هیچ لیست پیش‌فرض هاردکد.

    track='job' فقط آگهی‌های شغلی، track='education' فقط برنامه‌های تحصیلی،
    track=None یعنی هر دو با هم.
    """
    if not os.path.exists(SOURCES_PATH):
        print(f"  ❌ sources.json پیدا نشد در {SOURCES_PATH}")
        print(f"     این فایل در ریپو هست — `git checkout sources.json` یا از web_ui تب تنظیمات اضافه کن.")
        return []
    try:
        with open(SOURCES_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        sources = data.get("sources", []) if isinstance(data, dict) else data
        valid, invalid = [], 0
        for i, s in enumerate(sources):
            v = validate_source(s, i)
            if v:
                valid.append(v)
            else:
                invalid += 1
        if invalid:
            print(f"  ⚠️ {invalid} منبع نامعتبر در sources.json نادیده گرفته شد.")
        if track:
            valid = [s for s in valid if s["track"] == track]
        if country:
            cc = country.upper()
            valid = [s for s in valid if s.get("country") == cc]
        if not include_disabled:
            valid = [s for s in valid if s.get("enabled", True)]
        return valid
    except (json.JSONDecodeError, OSError) as e:
        print(f"  ❌ sources.json قابل خواندن نیست: {e}")
        return []


def source_bank_meta():
    """فراداده‌های sources.json (tracks و countries) — برای نمایش در web_ui."""
    try:
        with open(SOURCES_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}
    if not isinstance(data, dict):
        return {}
    return {"tracks": data.get("tracks", {}), "countries": data.get("countries", {})}


def save_sources(sources):
    with open(SOURCES_PATH, "w", encoding="utf-8") as f:
        json.dump({"sources": sources}, f, ensure_ascii=False, indent=2)


# ════════════════════════════════════════════════════
# بانک منابع کشف‌شده (آگهی‌دار) — discovered_sources.json
# هر سایتی که آگهی واقعی داد اینجا ثبت می‌شود و در اجرای
# بعدی «قبل از همه» و با اولویت بالا دوباره چک می‌شود.
# ════════════════════════════════════════════════════

def load_discovered():
    if not os.path.exists(DISCOVERED_PATH):
        return {}
    try:
        with open(DISCOVERED_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_discovered(bank):
    os.makedirs(MEM, exist_ok=True)
    with open(DISCOVERED_PATH, "w", encoding="utf-8") as f:
        json.dump(bank, f, ensure_ascii=False, indent=2)


def record_discovered_sources(all_sources, active_sources, per_source_jobs):
    """
    هر منبعی که آگهی داد → در discovered_sources.json ثبت/بروزرسانی می‌شود.
    خروجی: دیکشنری {name: {url, country, jobs_last, total_found, last_seen, hits}}
    """
    bank = load_discovered()
    now = DATE_STR
    for s in all_sources:
        name = s["name"]
        jobs = per_source_jobs.get(name, 0)
        entry = bank.get(name, {
            "url": s["url"], "country": s.get("country", "؟"),
            "total_found": 0, "hits": 0, "first_seen": now,
        })
        entry["url"] = s["url"]
        entry["country"] = s.get("country", "؟")
        entry["jobs_last"] = jobs
        if jobs > 0:
            entry["total_found"] = entry.get("total_found", 0) + jobs
            entry["hits"] = entry.get("hits", 0) + 1
        entry["last_seen"] = now
        bank[name] = entry
    save_discovered(bank)
    return bank


def get_priority_sources(active_sources):
    """
    ترتیب جستجو: اول منابعِ «آگهی‌دارِ» بانک (بر اساس total_found نزولی)،
    بعد منابعی که هنوز امتحان نشده‌اند.
    """
    bank = load_discovered()
    proven = [(bank[n]["total_found"], s) for s in active_sources if (n := s["name"]) in bank and bank[n].get("total_found", 0) > 0]
    proven.sort(key=lambda x: -x[0])
    proven_names = {s["name"] for _, s in proven}
    fresh = [s for s in active_sources if s["name"] not in proven_names]
    # ثابت: proven اول، fresh بعد (حتی اگر fresh امتیاز بالاتری داشته باشند)
    return [s for _, s in proven] + fresh


# ════════════════════════════════════════════════════
# تاریخچهٔ بازده منابع — SOURCE_YIELD_HISTORY.json
# هر اجرا: چند آگهی از هر منبع؟ → تب «📈 بازده منابع» روی همین رشد می‌کند
# ════════════════════════════════════════════════════

def record_yield_history(per_source_jobs):
    """یک رکورد تاریخی برای هر اجرا ذخیره می‌کند (حداکثر ۶۰ اجرای آخر)."""
    try:
        history = []
        if os.path.exists(YIELD_HISTORY_PATH):
            with open(YIELD_HISTORY_PATH, "r", encoding="utf-8") as f:
                history = json.load(f)
        if not isinstance(history, list):
            history = []
        history.append({
            "date": DATE_STR,
            "total": sum(per_source_jobs.values()),
            "sources": {k: v for k, v in per_source_jobs.items()},
        })
        with open(YIELD_HISTORY_PATH, "w", encoding="utf-8") as f:
            json.dump(history[-60:], f, ensure_ascii=False, indent=1)
    except Exception as e:
        print(f"  ⚠️ ذخیرهٔ تاریخچهٔ بازده ناموفق: {e}")


def load_yield_history():
    if not os.path.exists(YIELD_HISTORY_PATH):
        return []
    try:
        with open(YIELD_HISTORY_PATH, "r", encoding="utf-8") as f:
            history = json.load(f)
        return history if isinstance(history, list) else []
    except Exception:
        return []


def yield_stats(history=None):
    """آمار تجمعی هر منبع: مجموع/میانگین/بیشترین آگهی + چند اجرای کارا."""
    history = history if history is not None else load_yield_history()
    stats = {}
    for run in history:
        for name, cnt in (run.get("sources") or {}).items():
            st = stats.setdefault(name, {
                "runs": 0, "with_jobs": 0, "total": 0, "best": 0, "last": 0, "last_date": "",
            })
            cnt = int(cnt or 0)
            st["runs"] += 1
            st["total"] += cnt
            st["best"] = max(st["best"], cnt)
            if cnt > 0:
                st["with_jobs"] += 1
            st["last"] = cnt
            st["last_date"] = run.get("date", "")
    for st in stats.values():
        st["avg"] = round(st["total"] / st["runs"], 1) if st["runs"] else 0
        st["rate"] = round(st["with_jobs"] * 100 / st["runs"]) if st["runs"] else 0
    return stats


# ═══════════════ discovered-sites scanner ═══════════════
# قبل از هر جستجو، صفحهٔ اصلی هر منبع را می‌کَوید و هر دامنهٔ
# جدیدی که آگهی دارد به بانک اضافه می‌شود. خروجی به sources.json
# دست نمی‌زند — فقط بانک درونی discovered_sources.json.
# ════════════════════════════════════════════════════

AD_DOMAIN_HINTS = re.compile(
    r"(job|jobs|career|careers|vacan|recruit|hiring|emploi|stellen|vacature|stellenangebot|arbeit|baan|jobsite|position)",
    re.I
)

def scan_for_new_sources(sources, max_scan=10, timeout=8):
    """
    قبل از هر جستجو: صفحهٔ اصلی برخی منابع را می‌کَوید و دامنه‌های
    «آگهی‌دار» جدید را پیدا می‌کند. منابع جدید به discovered_sources.json
    (نه sources.json) اضافه می‌شوند تا لیست اصلی دست‌نخورده بماند.

    تیونینگ نسبت به نسخهٔ قبل:
    - شروع اسکن از منابع پربازدهٔ بانک (منابع کاری که لینک‌های خوبی دارند)
    - تأیید دومرحله‌ای: دامنهٔ نامزد باید صفحهٔ اصلی‌اش هم آگهی واقعی داشته باشد
      (JSON-LD یا لینک آگهی) — وگرنه فقط دامنهٔ تبلیغاتی است و ثبت نمی‌شود.
    - از هر منبعِ میزبان حداکثر ۳ دامنهٔ جدید تا بانک پرِ آشغال نشود.
    """
    bank = load_discovered()
    added = 0
    seen_domains = {urlparse(s["url"]).netloc for s in sources}
    # منابعی که تا الان بازده داشته‌اند اول اسکن می‌شوند — لینک‌های بهتری دارند
    ordered = sorted(
        sources,
        key=lambda s: -load_discovered().get(s["name"], {}).get("total_found", 0),
    )[:max_scan]
    for s in ordered:
        per_host = 0
        # HTTP ساده — نه مرورگر؛ اسکن نباید کند و ریسکی شود
        html = _fetch_plain(s["url"], timeout=timeout)
        if not html:
            continue
        for m in re.finditer(r'href="(https?://[^"]+)"', html):
            href = m.group(1)
            try:
                dom = urlparse(href).netloc.lower()
            except Exception:
                continue
            if not dom or dom in seen_domains:
                continue
            if not AD_DOMAIN_HINTS.search(dom):
                continue
            seen_domains.add(dom)
            # ── تأیید دومرحله‌ای: این دامنه واقعاً آگهی دارد؟ ──
            candidate_html = fetch_page(f"https://{dom}", timeout=timeout, allow_browser=False)
            if not candidate_html:
                continue
            probe = extract_jobs_from_html(candidate_html, {"name": dom, "url": f"https://{dom}", "country": s.get("country", "؟")})
            if not probe:
                continue  # بدون آگهی واقعی → تبلیغاتی است، ثبت نمی‌شود
            # ── ثبت ──
            name = f"[کشف‌شده] {dom}"
            if name in bank:
                continue
            bank[name] = {
                "url": f"https://{dom}", "country": s.get("country", "؟"),
                "discovered_from": s["name"], "discovered_at": DATE_STR,
                "total_found": len(probe), "hits": 1,
            }
            per_host += 1
            added += 1
            print(f"      🛰 دامنهٔ آگهی‌دار جدید: {dom} ({len(probe)} آگهی)", flush=True)
            if per_host >= 3 or added >= 12:
                break
        if added >= 12:
            break
    if added:
        save_discovered(bank)
    return added


# ════════════════════════════════════════════════════
# fetch page — با HEADERS صحیح
# ════════════════════════════════════════════════════
def fetch_page(url, timeout=15, allow_browser=True):
    """دریافت محتوای صفحه — اول HTTP ساده؛ اگر خالی/بلاک شد و Playwright بود، مرورگر headless."""
    html_text = _fetch_plain(url, timeout=timeout)
    if html_text and not (allow_browser and HAS_PLAYWRIGHT):
        return html_text
    # اگر HTTP ساده چیزی نداد (بلاک/JS) یا صفحه تقریباً خالی بود → مرورگر headless
    if not html_text or (allow_browser and HAS_PLAYWRIGHT and len(html_text) < 20000):
        rendered = _pw_fetch(url)
        if rendered:
            return rendered
    return html_text


def _fetch_plain(url, timeout=15):
    try:
        req = Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.5",
            }
        )
        with urlopen(req, timeout=timeout) as resp:
            if resp.status == 200:
                content = resp.read()
                for enc in ["utf-8", "cp1252", "latin-1"]:
                    try:
                        return content.decode(enc, errors="ignore")
                    except Exception:
                        continue
                return content.decode("utf-8", errors="ignore")
    except (HTTPError, URLError, OSError):
        pass
    return None


# ════════════════════════════════════════════════════
# JSON-LD (schema.org/JobPosting) — روش اصلی استخراج
# اکثر سایت‌های بزرگ (حتی JS-heavy) این schema را داخل HTML
# می‌گذارند چون Google for Jobs به آن نیاز دارد.
# ═══════════════════════════════════════════ JobPosting
def extract_jsonld_jobs(html, source_info):
    """استخراج آگهی‌ها از JSON-LD (schema.org/JobPosting) داخل HTML."""
    jobs = []
    for m in re.finditer(
        r'<script[^>]*type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        html, re.S | re.I
    ):
        try:
            raw = m.group(1).strip()
            data = json.loads(raw)
        except Exception:
            continue
        # JSON-LD می‌تواند یک لیست باشد یا @graph داشته باشد
        candidates = data if isinstance(data, list) else data.get("@graph", [data]) if isinstance(data, dict) else [data]
        for item in candidates:
            if not isinstance(item, dict):
                continue
            if item.get("@type") not in ("JobPosting", ["JobPosting"]):
                continue
            title = (item.get("title") or item.get("name") or "").strip()
            if not title:
                continue
            org = item.get("hiringOrganization") or {}
            company = ""
            if isinstance(org, dict):
                company = (org.get("name") or "").strip()
            elif isinstance(org, str):
                company = org.strip()
            url = item.get("url") or ""
            if not url and isinstance(item.get("sameAs"), str):
                url = item["sameAs"]
            location = item.get("jobLocation") or {}
            loc_str = ""
            if isinstance(location, dict):
                addr = location.get("address") or {}
                if isinstance(addr, dict):
                    loc_str = ", ".join(str(addr.get(k)) for k in ("addressLocality", "addressRegion") if addr.get(k))
            date_posted = item.get("datePosted") or ""
            jobs.append({
                "title": title[:150],
                "company": (company[:100] or "نامشخص"),
                "source": source_info.get("name", "نامشخص"),
                "country": source_info.get("country", "؟"),
                "url": (url or source_info.get("url", ""))[:300],
                "location": loc_str[:120],
                "date_posted": str(date_posted)[:40],
                "found_at": DATE_STR,
                "via": "json-ld",
            })
    return jobs


# ════════════════════════════════════════════════════
# پارس HTML دینامیک — لینک‌های آگهی + نام شرکت از متن اطراف
# ════════════════════════════════════════════════════
class DynamicJobParser(HTMLParser):
    """پارس HTML — لینک‌های آگهی + متن h2/h3 (شرکت/عنوان)."""
    JOB_PATH_RE = re.compile(r'(/job|/jobs/|/jobs\?|/vacan|/position|/careers?/|/stellen|/emploi|/offre|-jobs|jobdetail|/opening|/posting|search\?|/search/)', re.I)
    ABSOLUTE_URL_RE = re.compile(r'^https?://')
    NOISE_TEXT_RE = re.compile(
        r'^(apply now|read more|see more|view all|sign in|log in|register|home|about|contact|jobs? list|'
        r'next|previous|more|search|filter|save|share|print|email|all jobs|browse jobs|find a job|'
        r'create alert|français.*|language selection|skip to.*|main navigation menu|secondary menu|'
        r'account menu|job seekers|employers|training and careers|job search.*|my workspace|date modified|'
        r'همه|بیشتر|جستجو|ورود|ثبت نام|صفحه اصلی|تماس|درباره ما)$', re.I
    )
    # نویزهایی که فقط باید «شامل» متن باشند (نه تطبیق کامل)
    NOISE_TEXT_PARTIAL_RE = re.compile(
        r'(language selection|skip to|create (an )?alert|date modified|terms and conditions|privacy|'
        r'^\s*(français|english)\s*$|report a problem|feedback)', re.I
    )


    def __init__(self, base_url=""):
        super().__init__()
        self.base_url = base_url
        self.possible_links = []      # [{text, href}]
        self.nearby_company = {}      # id(anchor) → متن نزدیک h3/company
        self._anchor_depth = 0
        self._anchor_text = []
        self._anchor_href = ""
        self._current_heading = ""    # آخرین h1-h4 دیده‌شده
        self._capture_heading = False

    def handle_starttag(self, tag, attrs):
        d = dict(attrs)
        if tag in ("h1", "h2", "h3", "h4"):
            self._capture_heading = True
            self._current_heading = ""
        elif tag == "a" and d.get("href"):
            self._anchor_depth = 1
            self._anchor_text = []
            self._anchor_href = d["href"]
        elif self._anchor_depth > 0:
            self._anchor_depth += 1

    def handle_endtag(self, tag):
        if tag in ("h1", "h2", "h3", "h4"):
            self._capture_heading = False
            txt = " ".join(self._current_heading.split())
            if 4 <= len(txt) <= 120:
                self._current_heading = txt
            else:
                self._current_heading = ""
        elif tag == "a" and self._anchor_depth > 0:
            text = " ".join("".join(self._anchor_text).split())
            self._anchor_depth = 0
            href = self._anchor_href or ""
            # رد کردن لینک‌های غیرمحتوا
            if href.startswith(("#", "javascript:", "mailto:", "tel:")):
                return
            if 8 <= len(text) <= 200 and not self.NOISE_TEXT_RE.match(text) and not self.NOISE_TEXT_PARTIAL_RE.search(text):
                if href.startswith("/"):
                    href = urljoin(self.base_url, href)
                self.possible_links.append({"text": text, "href": href, "near_company": self._current_heading})

    def handle_data(self, data):
        if self._capture_heading:
            self._current_heading += data
        if self._anchor_depth > 0:
            self._anchor_text.append(data)


NOISE_COMPANY_RE = re.compile(
    r'^(apply|view|see|read|more|all|browse|find|search|save|share|home|about|contact|login|sign|'
    r'jobs|job|vacancy|vacancies|career|careers|new|latest|featured|remote|full.?time|part.?time|'
    r'session expired|error|page not found|not found|menu|navigation)$', re.I
)


def _same_text(a, b):
    """آیا دو متن تقریباً یکی هستند؟ (برای رد کردن عنوان به‌عنوان شرکت)"""
    a = re.sub(r'\W+', ' ', (a or '')).lower().strip()
    b = re.sub(r'\W+', ' ', (b or '')).lower().strip()
    if not a or not b:
        return False
    return a == b or a[:40] == b[:40] or a in b or b in a


# الگوی «شرکت: X» یا «Company: X» در متن لینک/عنوان
COMPANY_PREFIX_RE = re.compile(r'(?:شرکت|کمپانی|company|employer|organisation|organization)\s*[:：]\s*(.{2,60})', re.I)


def guess_company(link_info):
    """حدس نام شرکت از متن لینک و عنوان نزدیک — عنوانِ خود آگهی هرگز شرکت نیست."""
    title = link_info.get("text", "")
    for src in (title, link_info.get("near_company", "")):
        m = COMPANY_PREFIX_RE.search(src)
        if m:
            return m.group(1).strip(" ,-–|·")[:80]
    near = link_info.get("near_company", "")
    if near and 3 <= len(near) <= 80 and not NOISE_COMPANY_RE.match(near) and not _same_text(near, title):
        return near[:80]
    return ""


def _strip_session_ids(url):
    """حذف jsessionid و مسیرهای session از URL — key بهتر برای dedup"""
    return url.split(";jsessionid=")[0].split(";JSESSIONID=")[0].split(";sid=")[0].split(";CFID=")[0]


def extract_jobs_from_html(html, source_info):
    """استخراج آگهی‌ها: اول JSON-LD، بعد لینک‌های آگهی‌نما."""
    jobs = []

    # ۱) JSON-LD — دقیق‌ترین روش
    try:
        jobs.extend(extract_jsonld_jobs(html, source_info))
    except Exception:
        pass

    # ۲) لینک‌های آگهی‌نما از پارسر HTML
    parser = DynamicJobParser(base_url=source_info.get("url", ""))
    try:
        parser.feed(html)
    except Exception:
        pass

    seen_urls = {j["url"] for j in jobs if j.get("url")}
    for li in parser.possible_links:
        href = li["href"]
        # فیلتر: مسیر لینک باید شبیه آگهی باشد
        path = urlparse(href).path
        if not path or len(path) < 3:
            continue
        if not (DynamicJobParser.JOB_PATH_RE.search(href) or "job" in li["text"].lower()[:40]
                or "vacan" in li["text"].lower()[:40] or "career" in li["text"].lower()[:40]):
            continue
        if href in seen_urls:
            continue
        seen_urls.add(href)
        company = guess_company(li) or "نامشخص"
        jobs.append({
            "title": li["text"][:150],
            "company": company,
            "source": source_info.get("name", "نامشخص"),
            "country": source_info.get("country", "؟"),
            "url": href[:300],
            "location": "",
            "date_posted": "",
            "found_at": DATE_STR,
            "via": "html-link",
        })
        if len(jobs) > 150:  # سقف: هر صفحه حداکثر ۱۵۰ آگهی
            break
    return jobs


# ════════════════════════════════════════════════════
# تشخیص واقعی/فیک آگهی + امتیاز تطبیق
# ════════════════════════════════════════════════════
GENERIC_TITLE_WORDS = ["manager", "director", "engineer", "developer", "analyst", "coordinator", "specialist", "nurse", "midwife", "technician"]

def detect_job_realness(job_title, job_company, applicants_config):
    """تشخیص آیا آگهی واقعی است یا فیک + امتیاز تطبیق با متقاضیان"""
    score = 0
    max_possible = 0
    title_lower = (job_title or "").lower()
    company_lower = (job_company or "").lower() if isinstance(job_company, str) else ""

    for app in applicants_config:
        keywords = app.get("keywords", [])
        for kw in keywords:
            if kw.lower() in title_lower:
                score += 2
        max_possible += len(keywords) * 2
        for kw in keywords:
            if kw.lower() in company_lower:
                score += 1
        max_possible += len(keywords) * 1

    if score == 0 and title_lower:
        for gw in GENERIC_TITLE_WORDS:
            if gw in title_lower:
                score += 1
                break

    realness_ratio = score / max_possible if max_possible > 0 else 0
    is_real = realness_ratio > 0.2 or score >= 2
    return {
        "is_real": is_real,
        "realness_score": int(realness_ratio * 100),
        "match_score": score,
        "max_possible": max_possible
    }


# ════════════════════════════════════════════════════
# ساخت Excel
# ═════════════════════میکس══════════════════════════
def build_jobs_excel(jobs, applicants_config):
    """ساخت Excel با جزئیات کامل و پیگیری ایمیل"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "نتایج جستجو"
    ws.sheet_view.rightToLeft = True

    thin_border = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))
    header_font = Font(size=9, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid")
    cell_font = Font(size=9)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    headers = ["#", "عنوان شغل", "شرکت", "کشور", "منبع", "لینک آگهی",
               "محل", "تاریخ آگهی", "تاریخ یافت", "امتیاز تطبیق", "حقیقی/فیک", "متقاضی پیشنهادی"]
    for i, h in enumerate(headers):
        cell = ws.cell(row=1, column=i + 1, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    widths = [5, 45, 25, 8, 20, 45, 18, 14, 16, 10, 12, 24]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter_safe(i + 1)].width = w

    applicant_ids = [(a.get("id", "").lower(), a) for a in applicants_config]

    row = 2
    for idx, job in enumerate(jobs, 1):
        realness = detect_job_realness(job.get("title", ""), job.get("company", ""), applicants_config)
        applicant_suggestion = "—"
        for aid, a in applicant_ids:
            title_lower = (job.get("title", "") or "").lower()
            if any(k.lower() in title_lower for k in a.get("keywords", [])):
                applicant_suggestion = a.get("name_fa") or a.get("name", aid)
                break

        values = [
            idx,
            job.get("title", "")[:80],
            job.get("company", "")[:60],
            job.get("country", "") or "؟",
            job.get("source", "")[:30],
            job.get("url", ""),
            job.get("location", "")[:40],
            job.get("date_posted", "") or "—",
            job.get("found_at", ""),
            realness["match_score"],
            "✅ واقعی" if realness["is_real"] else "❌ فیک",
            applicant_suggestion,
        ]
        for ci, v in enumerate(values):
            cell = ws.cell(row=row, column=ci + 1, value=v)
            cell.font = cell_font
            cell.border = thin_border
            if ci == 6:
                cell.hyperlink = job.get("url", "")
            if ci == 10:
                if "✅" in str(v):
                    cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                else:
                    cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        row += 1

    ws.cell(row=row, column=1, value=len(jobs))
    ws.cell(row=row, column=2, value="جمع کل آگهی‌ها")
    ws.auto_filter.ref = f"A1:L{row}"
    return wb


def get_column_letter_safe(col_idx):
    """A, B, C, ... بدون نیاز به openpyxl.utils"""
    letters = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ═════════ Pattern/sourced search ═════════
def _guess_keyword(url):
    """حدس کلیدواژهٔ جستجو از URL — فقط برای نمایش زنده."""
    try:
        parsed = urlparse(url)
        qs = parse_qs(parsed.query)
        for key in ("q", "searchstring", "was", "keywords", "keyword", "k"):
            if key in qs and qs[key]:
                return qs[key][0].replace("+", " ")
        segment = parsed.path.strip("/").split("/")[-1]
        segment = segment.replace("-jobs", "").replace("-", " ").strip()
        return segment or "(کل صفحه)"
    except Exception:
        return "؟"


def fetch_one(source, url, timeout=None):
    """یک آدرس را می‌گیرد و HTML برمی‌گرداند — منابع JS-heavy مستقیم با مرورگر.

    (نکته: برای منابع needs_js اول مرورگر امتحان می‌شود چون HTTP ساده
    روی اکثر بوردهای بزرگ با 403 بلاک می‌شود و اتلاف وقت است. اگر مرورگر
    هم جواب نداد، به HTTP ساده هم فرصت داده می‌شود.)
    timeout اختیاری است — به ثانیه — و برای اسکریپت اعتبارسنجی لازم است.
    """
    secs = timeout if timeout else None
    # مرورگر سراسری نیست: هر thread وضعیت خودش را دارد، پس اینجا فقط
    # حاضر بودن پکیج را می‌پرسیم و اگر thread عوض شده باشد _pw_fetch خودش
    # None می‌دهد و به plain برمی‌گردیم.
    needs_js = bool(source.get("needs_js"))
    if needs_js and HAS_PLAYWRIGHT:
        html = _pw_fetch(url, timeout_ms=(int(secs * 1000) if secs else 25000))
        if html is None and secs:
            html = _fetch_plain(url, timeout=secs)
        elif html is None:
            html = _fetch_plain(url)
        return html
    if secs:
        return _fetch_plain(url, timeout=secs)
    return fetch_page(url)


def search_source(source, on_progress=None, on_url_done=None, extractor=None, kind="آگهی"):
    """جستجو در یک منبع، آدرس‌به‌آدرس — و گزارش زندهٔ هر آدرس.

    خروجی: (همهٔ آیتم‌ها، per_url) که per_url فهرستِ
    [{url, keyword, jobs, ms, status}] برای جدول زنده و آمار بازده است.
    on_progress(url, keyword, i, total) قبل از هر fetch،
    on_url_done(per_url_row) بعد از هر fetch — روی همین‌ها نوار زنده حرکت می‌کند.
    extractor پیش‌فرض استخراج آگهی؛ education_crawler استخراج برنامه می‌دهد.
    """
    extract = extractor or extract_jobs_from_html
    all_items = []
    per_url = []
    urls = source.get("search_urls") or [source["url"]]
    name = source["name"]
    total = len(urls)

    for i, url in enumerate(urls, 1):
        keyword = _guess_keyword(url)
        if on_progress:
            on_progress(url, keyword, i, total)
        emit(LG_FETCH, f"[{i}/{total}] کلیدواژه: «{keyword}» ← {url}")

        t0 = time.monotonic()
        html = fetch_one(source, url)
        ms = int((time.monotonic() - t0) * 1000)

        if html:
            try:
                items = extract(html, source)
            except Exception as e:
                print(f"       ⚠️ خطا در استخراج: {e}", flush=True)
                items = []
            all_items.extend(items)
            row = {"url": url, "keyword": keyword, "jobs": len(items),
                   "ms": ms, "status": "ok"}
            per_url.append(row)
            emit(LG_DONE, f"[{i}/{total}] {name} · «{keyword}» · {len(items)} {kind} · {ms}ms")
            time.sleep(0.4)  # فاصلهٔ کوتاه تا سایت ما را ریت‌لیمیت نکند
        else:
            hint = browser_hint() if source.get("needs_js") else "مسدودسازی یا نیاز به JS"
            row = {"url": url, "keyword": keyword, "jobs": 0, "ms": ms, "status": "blocked"}
            per_url.append(row)
            emit(LG_FAIL, f"[{i}/{total}] {name} · «{keyword}» · پاسخی نیامد ({hint}) · {ms}ms")

        if on_url_done:
            on_url_done(row)

    return all_items, per_url


# ════════════════════════════════════════════════════
# اجرا
# ════════════════════════════════════════════════════
def dedupe(items, source_key="source"):
    """حذف تکراری بر اساس URL (بدون session id و بدون query) یا عنوان+مصدر."""
    seen_keys = set()
    unique = []
    for j in items:
        url = _strip_session_ids((j.get("url") or "").split("?")[0])
        key = url if url.startswith("http") else (
            (j.get("title", "")[:50].lower(), j.get("company", "")[:30].lower(),
             j.get(source_key, "").lower()))
        if key in seen_keys:
            continue
        seen_keys.add(key)
        unique.append(j)
    return unique


def build_excel(unique_items, applicants_config, filename, sheet_title, headers, widths, row_fn, hyperlink_col):
    """ساخت فایل اکسل — مشترک بین مسیر کار و مسیر تحصیل."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.sheet_view.rightToLeft = True

    thin = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))
    header_font = Font(size=9, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid")
    cell_font = Font(size=9)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, h in enumerate(headers):
        c = ws.cell(row=1, column=i + 1, value=h)
        c.font, c.fill, c.alignment, c.border = header_font, header_fill, center, thin
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter_safe(i + 1)].width = w

    row = 2
    for n, item in enumerate(unique_items, 1):
        for ci, val in enumerate(row_fn(n, item, applicants_config)):
            c = ws.cell(row=row, column=ci + 1, value=val)
            c.font, c.border = cell_font, thin
            if ci == hyperlink_col and isinstance(val, str) and val.startswith("http"):
                c.hyperlink = val
        row += 1

    ws.cell(row=row, column=1, value=len(unique_items))
    ws.cell(row=row, column=2, value="جمع کل")
    ws.auto_filter.ref = f"A1:{get_column_letter_safe(len(headers))}{row}"
    return wb


def main(track="job", country=None, discover=True, make_excel=True):
    kind = "آگهی" if track == "job" else "برنامه"
    label = "کاریابی" if track == "job" else "تحصیل"
    print("═" * 60)
    print(f"MigrationHunter — جستجوی خودکار {label}" + (f" · کشور {country.upper()}" if country else ""))
    print(f"📅 {DATE_STR}")
    print("═" * 60)

    from config_loader import get_applicants
    applicants_config = get_applicants()
    print(f"  👥 {len(applicants_config)} متقاضی پیکربندی شده")

    all_sources = load_sources(track=track, country=country, include_disabled=True)
    if not all_sources:
        print(f"\n  ❌ هیچ منبع «{label}» فعالی برای جستجو نیست — خروج.")
        return 0
    active_sources = [s for s in all_sources if s.get("enabled", True)]
    skipped = len(all_sources) - len(active_sources)
    print(f"\n🔍 {len(active_sources)} منبع فعال"
          + (f" ({skipped} غیرفعال رد شد)" if skipped else "")
          + f" · {sum(len(s['search_urls']) for s in active_sources)} آدرس")

    # ── گام صفر: اسکن منابع آگهی‌دار جدید (فقط برای مسیر کار) ──
    if discover and track == "job":
        print("\n🛰  اسکن برای منابع آگهی‌دار جدید (discovered_sources)…")
        try:
            n_new = scan_for_new_sources(active_sources)
            print(f"  {'✅' if n_new else 'ℹ️'} {n_new} دامنهٔ آگهی‌دار جدید کشف شد"
                  f"{' — در memory/discovered_sources.json ثبت شد' if n_new else ''}")
        except Exception as e:
            print(f"  ⚠️ اسکن منابع جدید با خطا مواجه شد: {e}")

    # ── ترتیب: اول منابع آگهی‌دارِ اثبات‌شده (از اجرای قبل) ──
    ordered = get_priority_sources(active_sources)
    proven_count = sum(1 for s in ordered
                       if load_discovered().get(s["name"], {}).get("total_found", 0) > 0)
    if proven_count:
        print(f"  ⭐ {proven_count} منبع آگهی‌دارِ اثبات‌شده از اجرای قبل اول جستجو می‌شوند")

    plan_progress(ordered)
    show_progress()

    extractor = extract_jobs_from_html if track == "job" else None
    if extractor is None:
        from education_crawler import extract_programs_from_html
        extractor = extract_programs_from_html

    all_items = []
    per_source = {}

    for idx, source in enumerate(ordered, 1):
        name = source["name"]
        P["current_source"] = name
        emit(LG_SOURCE, f"[{idx}/{len(ordered)}] {name} ({source.get('country','؟')}) — در حال بررسی…")

        t0 = time.monotonic()
        # on_url_done نوار را تک‌تک آدرس جلو می‌برد و جدول زنده را پر می‌کند
        items, per_url = search_source(
            source,
            on_progress=lambda u, k, i, t: P.update({"current_url": u, "current_keyword": k}),
            on_url_done=lambda row: (
                record_url(name, row["url"], row["keyword"], row["jobs"], row["ms"], row["status"]),
                update_progress(jobs=row["jobs"]),
            ),
            extractor=extractor,
            kind=kind,
        )
        secs = round(time.monotonic() - t0, 1)
        all_items.extend(items)
        per_source[name] = len(items)

        status = "✅" if items else "⚠️"
        emit(LG_SITE, f"مجموعاً {len(items)} {kind} یافت شد از {name} · {secs}s")
        P["site_done"] = idx
        print(f"    {status} {name} → {len(items)} {kind} در {secs}s", flush=True)
        time.sleep(0.2)

    P["pct"] = 100.0
    P["done"] = P["total"]
    show_progress()

    unique_items = dedupe(all_items)
    unique_items.sort(key=lambda j: detect_job_realness(
        j.get("title", ""), j.get("company", ""), applicants_config)["match_score"], reverse=True)

    os.makedirs(DASH, exist_ok=True)
    os.makedirs(MEM, exist_ok=True)
    fn = (f"Job_Crawler_{FILE_DATE}.xlsx" if track == "job"
          else f"Education_Crawler_{FILE_DATE}.xlsx")
    fp = os.path.join(DASH, fn)

    if make_excel:
        if track == "job":
            wb = build_jobs_excel(unique_items, applicants_config)
        else:
            from education_crawler import build_programs_excel
            wb = build_programs_excel(unique_items, applicants_config)
        wb.save(fp)
        print(f"\n📊 اکسل ساخته شد: {fn} ({len(unique_items)} ردیف)")

    out_json = os.path.join(MEM, "CRAWLER_RESULTS.json" if track == "job"
                            else "EDUCATION_RESULTS.json")
    key = "jobs" if track == "job" else "programs"
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump({
            "date": DATE_STR, "track": track, "country": country or "همه",
            key: unique_items, "found": len(unique_items),
            "scanned_sources": len(ordered),
            "scanned_urls": P["total"],
            "per_source": per_source,
            "sites_visited": P["sites_visited"],
        }, f, ensure_ascii=False, indent=2)

    if track == "job":
        bank = record_discovered_sources(all_sources, active_sources, per_source)
        n_proven = sum(1 for v in bank.values() if v.get("total_found", 0) > 0)
        print(f"\n🏦 بانک منابع آگهی‌دار بروزرسانی شد — {n_proven} منبع با آگهی واقعی ثبت شده")
        record_yield_history(per_source)
    _close_browser()

    # ── خلاصهٔ نهایی ──
    print(f"\n  📊 خلاصه {label}")
    print(f"  ├─ سایت‌های چک‌شده: {len(ordered)} · آدرس‌ها: {P['total']}")
    print(f"  ├─ مجموع {kind}ها (قبل از dedup): {len(all_items)}")
    print(f"  ├─ یکتا: {len(unique_items)}")
    countries = {}
    for j in unique_items:
        c = j.get("country", "؟")
        countries[c] = countries.get(c, 0) + 1
    if countries:
        print(f"  ├─ آمار کشورها:")
        for c, cnt in sorted(countries.items(), key=lambda x: -x[1])[:12]:
            print(f"  │  {c}: {cnt}")
    blocked = sum(1 for r in P["sites_visited"] if r.get("status") == "blocked")
    if blocked:
        print(f"  └─ ⚠️ {blocked} آدرس پاسخی نداد (۴۰۳ یا نیاز به مرورگر — `playwright install` را بزن)")
    print(f"  └─ ذخیره شد: {fn if make_excel else out_json}")
    print("═" * 60)
    return len(unique_items)


def cli():
    """نقطهٔ ورود خط فرمان — انتخاب مسیر کار یا تحصیل و کشور."""
    import argparse
    ap = argparse.ArgumentParser(
        description="MigrationHunter — کراولر آگهی شغلی و برنامهٔ تحصیلی",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog="مثال‌ها:\n"
               "  python job_crawler.py                      # همهٔ آگهی‌های شغلی\n"
               "  python job_crawler.py --track education    # فقط برنامه‌های تحصیلی\n"
               "  python job_crawler.py --country FI         # فقط فنلاند\n"
               "  python job_crawler.py --country FI --track job --no-discover\n")
    ap.add_argument("--track", choices=["job", "education"], default="job",
                    help="مسیر: کاریابی یا تحصیل (پیش‌فرض job)")
    ap.add_argument("--country", help="فقط یک کشور، مثلاً FI یا DE")
    ap.add_argument("--no-discover", action="store_true",
                    help="اسکن خودکار دامنه‌های آگهی‌دار جدید را رد کن (سریع‌تر)")
    ap.add_argument("--no-excel", action="store_true", help="فایل اکسل نساز")
    args = ap.parse_args()
    return main(track=args.track, country=args.country,
                discover=not args.no_discover, make_excel=not args.no_excel)


if __name__ == "__main__":
    sys.exit(cli() or 0)
