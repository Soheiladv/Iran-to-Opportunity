#!/usr/bin/env python3
"""
MigrationHunter — جستجوی خودکار کاریابی
جستجو در سایت‌های مختلف + ذخیره در Excel

اجرا: python job_crawler.py
"""
import os, sys, json, re, io
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.error import URLError
from html.parser import HTMLParser

# Fix Windows console encoding for emoji support
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    os.environ['PYTHONIOENCODING'] = 'utf-8'

BASE = os.path.dirname(os.path.abspath(__file__))
MEM = os.path.join(BASE, "memory")
DASH = os.path.join(BASE, "dashboard")

NOW = datetime.now()
DATE_STR = NOW.strftime("%Y-%m-%d %H:%M")
FILE_DATE = NOW.strftime("%Y%m%d_%H%M")

# ═══════════════════════════════════════════════════
# سایت‌های جستجو
# ═══════════════════════════════════════════════════
SEARCH_SOURCES = [
    # ─── نیوزیلند ───
    {
        "name": "Seek NZ", "country": "NZ", "type": "job_board",
        "url": "https://www.seek.co.nz/midwife-jobs",
        "search_urls": [
            "https://www.seek.co.nz/midwife-jobs",
            "https://www.seek.co.nz/it-manager-jobs",
        ],
        "trust": 85,
    },
    {
        "name": "Trade Me Jobs NZ", "country": "NZ", "type": "job_board",
        "url": "https://www.trademe.co.nz/jobs",
        "search_urls": [
            "https://www.trademe.co.nz/jobs/healthcare",
        ],
        "trust": 80,
    },
    # ─── استرالیا ───
    {
        "name": "Seek AU", "country": "AU", "type": "job_board",
        "url": "https://www.seek.com.au/midwife-jobs",
        "search_urls": [
            "https://www.seek.com.au/midwife-jobs",
            "https://www.seek.com.au/it-infrastructure-jobs",
        ],
        "trust": 85,
    },
    # ─── کانادا ───
    {
        "name": "Job Bank Canada", "country": "CA", "type": "government",
        "url": "https://www.jobbank.gc.ca",
        "search_urls": [
            "https://www.jobbank.gc.ca/jobsearch/jobsearch?searchstring=midwife&locationstring=",
            "https://www.jobbank.gc.ca/jobsearch/jobsearch?searchstring=it+manager&locationstring=",
        ],
        "trust": 90,
    },
    {
        "name": "Indeed Canada", "country": "CA", "type": "job_board",
        "url": "https://ca.indeed.com",
        "search_urls": [
            "https://ca.indeed.com/jobs?q=midwife&l=",
            "https://ca.indeed.com/jobs?q=it+infrastructure+manager&l=",
        ],
        "trust": 75,
    },
    # ─── آلمان ───
    {
        "name": "StepStone DE", "country": "DE", "type": "job_board",
        "url": "https://www.stepstone.de",
        "search_urls": [
            "https://www.stepstone.de/jobs/it-manager",
        ],
        "trust": 80,
    },
    {
        "name": "Indeed DE", "country": "DE", "type": "job_board",
        "url": "https://de.indeed.com",
        "search_urls": [
            "https://de.indeed.com/jobs?q=it+manager&l=",
        ],
        "trust": 75,
    },
    # ─── ایرلند ───
    {
        "name": "IrishJobs", "country": "IE", "type": "job_board",
        "url": "https://www.irishjobs.ie",
        "search_urls": [
            "https://www.irishjobs.ie/jobs/midwife",
        ],
        "trust": 75,
    },
    # ─── هلند ───
    {
        "name": "Indeed NL", "country": "NL", "type": "job_board",
        "url": "https://nl.indeed.com",
        "search_urls": [
            "https://nl.indeed.com/jobs?q=it+manager&l=",
        ],
        "trust": 75,
    },
]

# ═══════════════════════════════════════════════════
# خواندن اطلاعات متقاضیان
# ═══════════════════════════════════════════════════
def load_applicants():
    applicants = {
        "NEDA": {"profession": "Midwife", "keywords": ["midwife", "maternity", "obstetric", "neonatal"]},
        "TOHID": {"profession": "IT Manager", "keywords": ["it manager", "infrastructure", "systems", "network", "devops"]},
    }
    return applicants

# ═══════════════════════════════════════════════════
# جستجوی ساده (بدون API)
# ═══════════════════════════════════════════════════
def fetch_page(url, timeout=10):
    """دریافت محتوای صفحه"""
    try:
        req = Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })
        with urlopen(req, timeout=timeout) as resp:
            return resp.read().decode("utf-8", errors="ignore")
    except Exception as e:
        return None

def extract_jobs_from_html(html, source):
    """استخراج آگهی‌ها از HTML"""
    jobs = []
    
    # الگوهای رایج برای استخراج عنوان شغل
    title_patterns = [
        r'<h\d[^>]*>([^<]*(?:midwife|manager|engineer|admin|developer|analyst)[^<]*)</h\d>',
        r'title["\s:=]+["\']([^"\']+(?:midwife|manager|engineer|admin)[^"\']*)["\']',
    ]
    
    # الگو برای استخراج لینک
    link_patterns = [
        r'href=["\']([^"\']*(?:job|career|position|vacancy)[^"\']*)["\']',
    ]
    
    titles_found = set()
    for pattern in title_patterns:
        matches = re.findall(pattern, html, re.IGNORECASE)
        for m in matches:
            title = m.strip()
            if len(title) > 10 and len(title) < 200:
                titles_found.add(title)
    
    for title in list(titles_found)[:20]:  # حداکثر 20
        jobs.append({
            "title": title,
            "source": source["name"],
            "country": source["country"],
            "url": source["url"],
            "found_at": DATE_STR,
        })
    
    return jobs

def match_applicant(job, applicants):
    """تطبیق آگهی با متقاضی"""
    title_lower = job.get("title", "").lower()
    
    for person, info in applicants.items():
        for kw in info["keywords"]:
            if kw in title_lower:
                return person
    return None

# ═══════════════════════════════════════════════════
# ذخیره نتایج
# ═══════════════════════════════════════════════════
def save_results(jobs, crawl_log):
    """ذخیره نتایج در فایل"""
    os.makedirs(MEM, exist_ok=True)
    
    # ذخیره در JSON
    fp = os.path.join(MEM, "CRAWLER_RESULTS.json")
    existing = []
    if os.path.exists(fp):
        with open(fp, "r", encoding="utf-8") as f:
            existing = json.load(f)
    
    # جلوگیری از تکرار
    existing_urls = {j.get("url", "") + j.get("title", "") for j in existing}
    new_jobs = [j for j in jobs if j.get("url", "") + j.get("title", "") not in existing_urls]
    
    all_jobs = existing + new_jobs
    with open(fp, "w", encoding="utf-8") as f:
        json.dump(all_jobs, f, ensure_ascii=False, indent=2)
    
    # ذخیره لاگ
    log_fp = os.path.join(MEM, "CRAWLER_LOG.json")
    with open(log_fp, "w", encoding="utf-8") as f:
        json.dump(crawl_log, f, ensure_ascii=False, indent=2)
    
    return len(new_jobs)

def build_crawler_excel(jobs):
    """ساخت Excel از نتایج جستجو"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    
    wb = Workbook()
    ws = wb.active
    ws.title = "نتایج جستجو"
    ws.sheet_view.rightToLeft = True
    
    thin = Border(left=Side("thin"), right=Side("thin"),
                  top=Side("thin"), bottom=Side("thin"))
    
    # عنوان
    cell = ws.cell(row=1, column=1, value=f"نتایج جستجوی خودکار — {DATE_STR}")
    cell.font = Font(name="B Mitra", size=14, bold=True, color="1B4F72")
    ws.merge_cells("A1:G1")
    
    # هدرها
    headers = ["#", "عنوان شغل", "کشور", "منبع", "متقاضی", "لینک", "تاریخ"]
    for i, h in enumerate(headers):
        cell = ws.cell(row=3, column=i+1, value=h)
        cell.font = Font(name="B Mitra", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin
    
    # داده‌ها
    applicants = load_applicants()
    row = 4
    for idx, job in enumerate(sorted(jobs, key=lambda x: x.get("found_at", ""), reverse=True), 1):
        applicant = match_applicant(job, applicants)
        app_label = "👩 ندا" if applicant == "NEDA" else "👨 توحید" if applicant == "TOHID" else "❓"
        
        vals = [idx, job.get("title", ""), job.get("country", ""),
                job.get("source", ""), app_label, job.get("url", ""), job.get("found_at", "")]
        
        for ci, v in enumerate(vals):
            cell = ws.cell(row=row, column=ci+1, value=v)
            cell.font = Font(name="B Mitra", size=9)
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            cell.border = thin
        row += 1
    
    # عرض ستون‌ها
    widths = [5, 50, 10, 18, 10, 40, 18]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i+1)].width = w
    
    # ذخیره
    os.makedirs(DASH, exist_ok=True)
    fn = f"Job_Crawler_{FILE_DATE}.xlsx"
    fp = os.path.join(DASH, fn)
    wb.save(fp)
    return fn

# ═══════════════════════════════════════════════════
# اجرا
# ═══════════════════════════════════════════════════
def main():
    print("═" * 50)
    print("MigrationHunter — جستجوی خودکار کاریابی")
    print(f"📅 {DATE_STR}")
    print("═" * 50)
    
    applicants = load_applicants()
    all_jobs = []
    crawl_log = []
    
    print(f"\n🔍 جستجو در {len(SEARCH_SOURCES)} منبع...")
    
    for source in SEARCH_SOURCES:
        print(f"\n  📡 {source['name']} ({source['country']})...")
        
        for url in source.get("search_urls", [source["url"]]):
            html = fetch_page(url)
            if html:
                jobs = extract_jobs_from_html(html, source)
                all_jobs.extend(jobs)
                crawl_log.append({
                    "source": source["name"],
                    "url": url,
                    "status": "success",
                    "jobs_found": len(jobs),
                    "time": DATE_STR,
                })
                print(f"    {OK} {len(jobs)} آگهی یافت شد")
            else:
                crawl_log.append({
                    "source": source["name"],
                    "url": url,
                    "status": "failed",
                    "jobs_found": 0,
                    "time": DATE_STR,
                })
                print(f"    ⚠️ خطا در دریافت صفحه")
    
    # ذخیره
    print(f"\n💾 ذخیره نتایج...")
    new_count = save_results(all_jobs, crawl_log)
    print(f"  {OK} {new_count} آگهی جدید ذخیره شد")
    
    # ساخت Excel
    print(f"\n📊 ساخت Excel...")
    fn = build_crawler_excel(all_jobs)
    print(f"  {OK} {fn}")
    
    # خلاصه
    print("\n" + "═" * 50)
    print("📊 خلاصه")
    print("═" * 50)
    
    countries = {}
    for j in all_jobs:
        c = j.get("country", "?")
        countries[c] = countries.get(c, 0) + 1
    
    for c, cnt in sorted(countries.items(), key=lambda x: -x[1]):
        print(f"  {c}: {cnt} آگهی")
    
    print(f"\n  مجموع: {len(all_jobs)} آگهی")
    print(f"  جدید: {new_count}")
    print("═" * 50)

if __name__ == "__main__":
    main()
