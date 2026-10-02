# -*- coding: utf-8 -*-
"""LinkedIn Live — ورود زنده، استخراج پروفایل، ریکروتریابی و ذخیرهٔ شغل.

هیچ وابستگی به Google Storage ندارد: درایور از chromedriver-py محلی (بستهٔ
PyPI که باینری داخل خودش دارد) خوانده می‌شود و فقط در صورت نبودِ آن، سراغ
PATH یا Selenium Manager می‌رود. چون گوگل در ایران مسدود است، گزینهٔ آخر
معمولاً کار نمی‌کند و همان chromedriver-py کافی است.

اعتبارنامه فقط از محیط خوانده می‌شود:
    LINKEDIN_EMAIL=... LINKEDIN_PASSWORD=... python linkedin_live.py login

خروجی‌ها:
    linkedin_profile.json        پروفایل استخراج‌شده
    memory/LINKEDIN_DB.json      پروفایل + ریکروترها (ادغام‌شده با قبل)
    memory/LINKEDIN_JOBS.json    آگهی‌های ذخیره‌شده
    memory/linkedin_live.log     لاگ برای نمایش در تب LinkedIn داشبورد
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import shutil
import sys
import time
from pathlib import Path
from urllib.parse import quote_plus

BASE = Path(__file__).resolve().parent
MEM = BASE / "memory"
PROFILE_OUT = BASE / "linkedin_profile.json"
DB_PATH = MEM / "LINKEDIN_DB.json"
JOBS_PATH = MEM / "LINKEDIN_JOBS.json"
LOG_PATH = MEM / "linkedin_live.log"

LOGIN_URL = "https://www.linkedin.com/login"
FEED_URL = "https://www.linkedin.com/feed/"


# --------------------------------------------------------------------------- #
# لاگ
# --------------------------------------------------------------------------- #
def log(msg: str) -> None:
    line = f"[{_dt.datetime.now():%H-%M:%S}] {msg}"
    print(line, flush=True)
    try:
        MEM.mkdir(exist_ok=True)
        with LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def read_log(n: int = 400) -> str:
    try:
        return "\n".join(LOG_PATH.read_text(encoding="utf-8").splitlines()[-n:])
    except Exception:
        return ""


# --------------------------------------------------------------------------- #
# اعتبارنامه (فقط محیط / .env محلی)
# --------------------------------------------------------------------------- #
def _load_dotenv() -> dict:
    p = BASE / ".env"
    out = {}
    try:
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            out[k.strip()] = v.strip().strip('"').strip("'")
    except Exception:
        pass
    return out


def get_credentials() -> tuple[str, str]:
    env = _load_dotenv()
    email = os.getenv("LINKEDIN_EMAIL") or env.get("LINKEDIN_EMAIL", "")
    password = os.getenv("LINKEDIN_PASSWORD") or env.get("LINKEDIN_PASSWORD", "")
    return email, password


# --------------------------------------------------------------------------- #
# درایور — بدون webdriver_manager
# --------------------------------------------------------------------------- #
def find_driver_path() -> str | None:
    """ chromedriver-py محلی ← PATH ← (آخرین راه) Selenium Manager. """
    try:
        from chromedriver_py import binary_path  # باینری داخل بستهٔ PyPI

        if binary_path and os.path.exists(binary_path):
            return binary_path
    except Exception:
        pass
    for name in ("chromedriver.exe", "chromedriver"):
        p = shutil.which(name)
        if p:
            return p
    return None


def get_driver(headless: bool = False, page_load: int = 45):
    """Chrome را با درایور محلی بالا می‌آورد — هیچ دانلودی انجام نمی‌شود.

    اگر Chrome همزمان جای دیگری در حال اجرا باشد گاهی با
    «session not created: Chrome instance exited» شکست می‌خورد؛
    یک بار دوباره تلاش می‌کنیم.
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service

    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--disable-notifications")
    opts.add_argument("--lang=en-US")
    # user-data-dir عمداً داده نمی‌شود: chromedriver خودش یک پروفایل موقت
    # یکتا می‌سازد، پس دو اجرا با هم برخورد نمی‌کنند.
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.page_load_strategy = "eager"

    path = find_driver_path()
    if path:
        svc = Service(executable_path=path)
        log(f"درایور محلی: {path}")
    else:
        svc = Service()
        log("⚠️ chromedriver-py پیدا نشد — سراغ Selenium Manager می‌روم (ممکن است به دلیل "
            "مسدودبودن گوگل شکست بخورد).")

    last = None
    for attempt in (1, 2):
        try:
            driver = webdriver.Chrome(service=svc, options=opts)
            driver.set_page_load_timeout(page_load)
            driver.implicitly_wait(3)
            return driver
        except Exception as e:
            last = e
            if attempt == 1:
                log("⚠️ مرورگر بالا نیامد، ۲ ثانیه دیگر دوباره تلاش می‌کنم…")
                time.sleep(2)
    raise last


# --------------------------------------------------------------------------- #
# ابزارهای سطح پایین
# --------------------------------------------------------------------------- #
def _find(driver, by, value, timeout=12):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))


def _clickable(driver, by, value, timeout=12):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))


def _safe(func, default=None):
    try:
        return func()
    except Exception:
        return default


def _txt(driver, by, value, default="—"):
    el = _safe(lambda: driver.find_element(by, value))
    return (el.text or default).strip() if el else default


# --------------------------------------------------------------------------- #
# ورود
# --------------------------------------------------------------------------- #
def login(driver, email: str, password: str, wait: int = 25) -> bool:
    from selenium.webdriver.common.by import By

    if not email or not password or email == "your_email@example.com":
        log("❌ LINKEDIN_EMAIL / LINKEDIN_PASSWORD تنظیم نشده — ورود انجام نشد.")
        return False

    log("🔐 باز کردن صفحهٔ لاگین لینکدین…")
    driver.get(LOGIN_URL)
    try:
        u = _find(driver, By.ID, "username", 20)
        p = driver.find_element(By.ID, "password")
        u.clear()
        u.send_keys(email)
        p.clear()
        p.send_keys(password)
        _clickable(driver, By.CSS_SELECTOR, "button[type='submit']").click()
        log("✅ دکمهٔ ورود کلیک شد.")
    except Exception as e:
        log(f"❌ فرم لاگین پیدا نشد: {e}")
        return False

    end = time.time() + wait
    while time.time() < end:
        time.sleep(1)
        url = driver.current_url or ""
        if "challenge" in url or "checkpoint" in url:
            log("⚠️ لینکدین چالش امنیتی/دومرحله‌ای نشان داد — لازم است دستی تأیید کنی.")
            return False
        if "/feed" in url or "linkedin.com/in/" in url:
            log("✅ ورود موفق.")
            return True
    log("❌ ورود در زمان مقرر تأیید نشد (ایمیل/رمز اشتباه یا تأیید دومرحله‌ای).")
    return False


# --------------------------------------------------------------------------- #
# استخراج پروفایل
# --------------------------------------------------------------------------- #
def extract_profile(driver) -> dict:
    from selenium.webdriver.common.by import By

    data = {
        "name": _txt(driver, By.CSS_SELECTOR, "h1.text-heading-xlarge", ""),
        "headline": _txt(driver, By.CSS_SELECTOR, "div.text-body-medium.break-words", ""),
        "location": _txt(driver, By.CSS_SELECTOR, "span.text-body-small.inline.t-black--light.break-words", ""),
        "url": driver.current_url.split("?")[0],
        "scraped_at": f"{_dt.datetime.now():%Y-%m-%d %H:%M}",
    }
    if not data["name"]:
        data["name"] = _txt(driver, By.CSS_SELECTOR, "main h1", "نامشخص")

    try:
        driver.get("https://www.linkedin.com/mynetwork/")
        time.sleep(2)
        data["connections"] = _txt(
            driver,
            By.CSS_SELECTOR,
            "a[href*='ring'] .t-bold, span[data-test-id='connections-count']",
            "—",
        )
    except Exception:
        data["connections"] = "—"
    return data


def save_profile(profile: dict, applicant: str = "") -> dict:
    """پروفایل را در linkedin_profile.json و memory/LINKEDIN_DB.json می‌نویسد."""
    PROFILE_OUT.write_text(
        json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    db = load_db()
    key = applicant or (profile.get("url", "").rstrip("/").rsplit("/", 1)[-1] or "me")
    db[key] = {**db.get(key, {}), **profile, "applicant": applicant or key}
    MEM.mkdir(exist_ok=True)
    DB_PATH.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"💾 پروفایل ذخیره شد: {PROFILE_OUT.name} و {DB_PATH.name} [{key}]")
    return db.get(key, {})


def load_db() -> dict:
    try:
        return json.loads(DB_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def load_jobs() -> list:
    try:
        d = json.loads(JOBS_PATH.read_text(encoding="utf-8"))
        return d if isinstance(d, list) else d.get("jobs", [])
    except Exception:
        return []


def _save_jobs(jobs: list) -> None:
    MEM.mkdir(exist_ok=True)
    JOBS_PATH.write_text(json.dumps(jobs, ensure_ascii=False, indent=2), encoding="utf-8")


# --------------------------------------------------------------------------- #
# ریکروتریابی (Recruiter discovery)
# --------------------------------------------------------------------------- #
def search_recruiters(driver, keywords: str, location: str = "Germany",
                      limit: int = 25) -> list:
    """اسکن نتایج جستجوی افراد لینکدین و برداشتن ریکروترها / تalent‌ها."""
    from selenium.webdriver.common.by import By

    kw = quote_plus(keywords)
    loc = quote_plus(location)
    url = (f"https://www.linkedin.com/search/results/people/?keywords={kw}"
           f"&origin=GLOBAL_SEARCH_HEADER")
    log(f"🔎 ریکروتریابی: «{keywords}» در {location}")
    driver.get(url)
    time.sleep(4)

    found, seen = [], set()
    cards = _safe(lambda: driver.find_elements(
        By.CSS_SELECTOR, "li.reusable-search__result-container, div.entity-result")) or []
    log(f"   {len(cards)} کارت نتیجه دیده شد.")

    for card in cards[: limit * 2]:
        if len(found) >= limit:
            break
        name = _safe(lambda c=card: (c.find_element(
            By.CSS_SELECTOR, "span.entity-result__title-text a span[aria-hidden='true']").text or "").strip())
        if not name:
            name = _safe(lambda c=card: (c.find_element(By.CSS_SELECTOR, "a[href*='/in/']").text or "").strip())
        if not name:
            continue
        headline = _safe(lambda c=card: (c.find_element(
            By.CSS_SELECTOR, ".entity-result__primary-subtitle, .t-normal.t-black").text or "").strip()) or ""
        locn = _safe(lambda c=card: (c.find_element(
            By.CSS_SELECTOR, ".entity-result__secondary-subtitle").text or "").strip()) or ""
        href = _safe(lambda c=card: c.find_element(
            By.CSS_SELECTOR, "a[href*='/in/']").get_attribute("href")) or ""
        href = href.split("?")[0]
        if href in seen:
            continue
        seen.add(href)
        role = (headline + " " + name).lower()
        is_rec = any(k in role for k in (
            "recruit", "talent", "hiring", "people partner", "hr ", "hr-", "headhunter",
            "staffing", "sourcing", "personalleiter", "personalreferent"))
        found.append({
            "name": name, "headline": headline, "location": locn, "url": href,
            "recruiter": bool(is_rec),
            "found_at": f"{_dt.datetime.now():%Y-%m-%d %H:%M}",
        })

    # ادغام در دیتابیس محلی
    db = load_db()
    bucket = db.setdefault("_recruiters", {})
    for r in found:
        bucket[r["url"]] = {**bucket.get(r["url"], {}), **r,
                            "keyword": keywords, "location": location}
    MEM.mkdir(exist_ok=True)
    DB_PATH.write_text(json.dumps(db, ensure_ascii=False, indent=2), encoding="utf-8")
    hits = sum(1 for r in found if r["recruiter"])
    log(f"✅ {len(found)} نفر ذخیره شد که {hits} نفرشان ریکروتر/تalent بودند.")
    return found


# --------------------------------------------------------------------------- #
# جستجوی شغل + ذخیرهٔ شغل
# --------------------------------------------------------------------------- #
def search_jobs(driver, keywords: str, location: str = "Germany",
                limit: int = 25) -> list:
    from selenium.webdriver.common.by import By

    kw, loc = quote_plus(keywords), quote_plus(location)
    url = (f"https://www.linkedin.com/jobs/search/?keywords={kw}"
           f"&location={loc}&f_WT=2")   # فقط فاصله‌کاری/ریموت‌پذیر
    log(f"💼 جستجوی شغل: «{keywords}» در {location}")
    driver.get(url)
    time.sleep(4)

    cards = _safe(lambda: driver.find_elements(
        By.CSS_SELECTOR, "div.job-card-container, li.jobs-search-results__list-item")) or []
    jobs, seen = [], set()
    for card in cards:
        if len(jobs) >= limit:
            break
        title = _safe(lambda c=card: (c.find_element(
            By.CSS_SELECTOR, "a.job-card-list__title, a[data-control-name='job_card']").text or "").strip())
        comp = _safe(lambda c=card: (c.find_element(
            By.CSS_SELECTOR, ".job-card-container__primary-description, "
                             "h4.base-search-card__subtitle, span.job-card-container__company-name").text or "").strip())
        href = _safe(lambda c=card: c.find_element(
            By.CSS_SELECTOR, "a[href*='/jobs/view/']").get_attribute("href")) or ""
        href = href.split("?")[0]
        if not title or href in seen:
            continue
        seen.add(href)
        jobs.append({
            "title": title, "company": (comp or "").strip(), "url": href,
            "keyword": keywords, "location": location, "source": "linkedin",
            "saved_at": f"{_dt.datetime.now():%Y-%m-%d %H:%M}",
        })
    log(f"✅ {len(jobs)} آگهی استخراج شد.")
    return jobs


def save_job(driver, job_url: str) -> bool:
    """روی صفحهٔ آگهی دکمهٔ Save را می‌زند (نیازمند لاگین)."""
    from selenium.webdriver.common.by import By

    driver.get(job_url)
    time.sleep(3)
    btn = None
    for sel in ("button[aria-label^='Save']",
                "button.jobs-save-button",
                "button[aria-label*='ذخیره']"):
        btn = _safe(lambda s=sel: driver.find_element(By.CSS_SELECTOR, s))
        if btn:
            break
    if not btn:
        log("❌ دکمهٔ Save پیدا نشد (شاید لاگین نیستی یا آگهی قدیمی است).")
        return False
    try:
        state = (btn.get_attribute("aria-label") or "")
        if "Unsave" in state or "حذف" in state:
            log("ℹ️ این آگهی قبلاً ذخیره شده بود.")
            return True
        driver.execute_script("arguments[0].click();", btn)
        time.sleep(1.5)
        log(f"⭐ آگهی ذخیره شد: {job_url}")
        _record_job(driver, job_url)
        return True
    except Exception as e:
        log(f"❌ خطا در ذخیره: {e}")
        return False


def _record_job(driver, url: str) -> None:
    from selenium.webdriver.common.by import By

    title = _txt(driver, By.CSS_SELECTOR, "h1.t-24, h1.jobs-unified-top-card__job-title", "")
    comp = _txt(driver, By.CSS_SELECTOR, ".jobs-unified-top-card__company-name a, "
                                         "span.jobs-unified-top-card__company-name", "")
    jobs = load_jobs()
    if not any(j.get("url") == url for j in jobs):
        jobs.append({
            "title": title or "(بدون عنوان)", "company": comp or "—", "url": url,
            "source": "linkedin", "saved": True,
            "saved_at": f"{_dt.datetime.now():%Y-%m-%d %H:%M}",
        })
        _save_jobs(jobs)


def save_job_list(jobs: list) -> list:
    """آگهی‌های استخراج‌شده را به فهرست ذخیره‌شده‌ها اضافه می‌کند (بدون تکرار)."""
    cur = load_jobs()
    have = {j.get("url") for j in cur}
    for j in jobs:
        if j.get("url") and j["url"] not in have:
            cur.append(j)
            have.add(j["url"])
    _save_jobs(cur)
    log(f"💾 {len(jobs)} آگهی به memory/LINKEDIN_JOBS.json اضافه شد.")
    return cur


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="LinkedIn Live — بدون webdriver_manager")
    ap.add_argument("action", choices=[
        "profile", "login", "recruiters", "jobs", "save", "scan", "log"],
        help="profile=لاگین+استخراج پروفایل · recruiters=ریکروتریابی · "
             "jobs=جستجوی شغل · save=ذخیرهٔ یک آدرس · scan=jobs+recruiters · log=نمایش لاگ")
    ap.add_argument("--keywords", default="recruiter HR talent acquisition")
    ap.add_argument("--location", default="Germany")
    ap.add_argument("--url", default="", help="آدرس آگهی برای save")
    ap.add_argument("--applicant", default="", help="شناسهٔ متقاضی برای ذخیره در LINKEDIN_DB")
    ap.add_argument("--limit", type=int, default=25)
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--no-login", action="store_true", help="بدون لاگین فقط جستجو کن")
    args = ap.parse_args(argv)

    if args.action == "log":
        print(read_log())
        return 0

    email, password = get_credentials()
    driver = None
    rc = 0
    try:
        driver = get_driver(headless=args.headless)
        logged = args.no_login
        if not args.no_login:
            logged = login(driver, email, password)

        if args.action in ("profile", "login"):
            if not logged and args.action == "profile":
                return 2
            prof = extract_profile(driver)
            save_profile(prof, args.applicant)
            for k in ("name", "headline", "location", "connections"):
                print(f"  • {k}: {prof.get(k, '—')}")

        if args.action in ("recruiters", "scan"):
            if not logged:
                log("⚠️ بدون لاگین فقط نتایج عمومی دیده می‌شود.")
            search_recruiters(driver, args.keywords, args.location, args.limit)

        if args.action in ("jobs", "scan"):
            js = search_jobs(driver, args.keywords, args.location, args.limit)
            save_job_list(js)

        if args.action == "save":
            if not args.url:
                print("--url لازم است")
                return 1
            rc = 0 if save_job(driver, args.url) else 3

    except Exception as e:
        log(f"❌ خطا: {type(e).__name__}: {e}")
        rc = 1
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
    return rc


if __name__ == "__main__":
    sys.exit(main())
