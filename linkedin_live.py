# -*- coding: utf-8 -*-
"""LinkedIn Live — ورود زنده، استخراج پروفایل، ریکروتریابی و ذخیرهٔ شغل.

هیچ وابستگی به Google Storage ندارد: درایور از chromedriver-py محلی (بستهٔ
PyPI که باینری داخل خودش دارد) خوانده می‌شود و فقط در صورت نبودِ آن، سراغ
PATH یا Selenium Manager می‌رود. چون گوگل در ایران مسدود است، گزینهٔ آخر
معمولاً کار نمی‌کند و همان chromedriver-py کافی است.

اعتبارنامه فقط از محیط خوانده می‌شود:
    LINKEDIN_EMAIL=... LINKEDIN_PASSWORD=... python linkedin_live.py login

اگر اکانتت با «ورود با گوگل» ساخته شده و رمزی نداری، به‌جای رمز
به مرورگر خودت وصل می‌شود (نشست لاگین‌شده‌ات همان‌جا استفاده می‌شود):
    ۱) کروم را کاملاً ببند
    ۲) & "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222
    ۳) وارد لینکدین شو  ۴) python linkedin_live.py profile --mode attach

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

    سه مرحله تلاش می‌شود، چون «session not created: Chrome instance exited»
    معمولاً موقتی است (مرورگر دیگری همزمان باز است یا پنجرهٔ تعاملی در دسترس نیست):
      ۱) همان حالت خواسته‌شده
      ۲) دوباره با همان حالت، ولی با درایور تازه (سرویس قبلی ممکن است خراب مانده باشد)
      ۳) حالت ناشناس (headless) — وقتی پنجرهٔ بصری در دسترس نیست
    لاگ verbose درایور در memory/chromedriver.log می‌ماند تا علت قابل بررسی باشد.
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service

    MEM.mkdir(exist_ok=True)
    driver_log = str(MEM / "chromedriver.log")

    def _options(hd: bool):
        opts = Options()
        if hd:
            opts.add_argument("--headless=new")
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--lang=en-US")
        # پرچم‌هایی که استارت را در ویندوز/سیستم‌های پر از Chromهای دیگر مقاوم می‌کنند
        opts.add_argument("--no-first-run")
        opts.add_argument("--no-default-browser-check")
        opts.add_argument("--disable-extensions")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--disable-dev-shm-usage")
        # user-data-dir عمداً داده نمی‌شود: chromedriver خودش یک پروفایل موقت
        # یکتا می‌سازد، پس دو اجرا با هم برخورد نمی‌کنند.
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
        opts.page_load_strategy = "eager"
        return opts

    path = find_driver_path()
    if path:
        log(f"درایور محلی: {path}")
    else:
        log("⚠️ chromedriver-py پیدا نشد — سراغ Selenium Manager می‌روم (ممکن است به دلیل "
            "مسدودبودن گوگل شکست بخورد).")

    attempts = [(headless, "همان حالت"), (headless, "درایور تازه"), (not headless, "ناشناس")]
    last = None
    for i, (hd, why) in enumerate(attempts, 1):
        if hd != headless:
            log(f"⚠️ مرورگر بالا نیامد → تلاش {i} در حالت ناشناس (headless)…")
        elif i > 1:
            log(f"⚠️ مرورگر بالا نیامد → تلاش {i} ({why})…")
        svc = Service(executable_path=path, log_output=driver_log) if path \
            else Service(log_output=driver_log)
        try:
            driver = webdriver.Chrome(service=svc, options=_options(hd))
            driver.set_page_load_timeout(page_load)
            driver.implicitly_wait(3)
            if hd != headless:
                log("ℹ️ مرورگر ناشناس بالا آمد (پنجرهٔ بصری در دسترس نبود).")
            return driver
        except Exception as e:
            last = e
            try:
                svc.stop()
            except Exception:
                pass
            time.sleep(1.5)
    log(f"❌ درایور: {type(last).__name__}: {str(last).splitlines()[0]}")
    log(f"   جزئیات در memory/chromedriver.log")
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
# اتصال به مرورگر خود کاربر (برای اکانت‌هایی که با «ورود با گوگل» ساخته شده‌اند
# و اصلاً رمز لینکدین ندارند). کاربر یک‌بار کروم خودش را با پورت دیباگ باز می‌کند:
#   & "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222
# بعد اسکریپت به همان نشستِ لاگین‌شده وصل می‌شود — بدون رمز، بدون فرم لاگین.
# --------------------------------------------------------------------------- #
DEBUG_PORT = 9222


def debug_port_open(port: int = DEBUG_PORT) -> bool:
    import socket

    s = socket.socket()
    s.settimeout(0.7)
    try:
        s.connect(("127.0.0.1", port))
        return True
    except Exception:
        return False
    finally:
        try:
            s.close()
        except Exception:
            pass


def try_attach(port: int = DEBUG_PORT):
    """به کرومِ در حال اجرای کاربر وصل می‌شود؛ اگر پورت دیباگ باز نباشد None."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service

    if not debug_port_open(port):
        return None
    MEM.mkdir(exist_ok=True)
    opts = Options()
    opts.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    path = find_driver_path()
    svc = Service(executable_path=path, log_output=str(MEM / "chromedriver.log")) \
        if path else Service(log_output=str(MEM / "chromedriver.log"))
    try:
        driver = webdriver.Chrome(service=svc, options=opts)
        driver.implicitly_wait(3)
        log(f"🔗 به مرورگر خودت وصل شدم (پورت {port}) — هیچ رمزی لازم نیست.")
        return driver
    except Exception as e:
        log(f"⚠️ اتصال به مرورگر خودت نشد: {type(e).__name__}")
        try:
            svc.stop()
        except Exception:
            pass
        return None


def already_logged_in(driver, timeout: int = 12) -> bool:
    """نشست لینکدین باز است؟ فقط ناوبری اصلی اپ ([data-testid="primary-nav"]) قبول است.

    عمداً سخت‌گیرانه: صفحات مهمان لینکدین هم لینک‌های عمومی دارند و نباید
    مثبت کاذب بدهند (کلاس‌های قدیمی global-nav دیگر در DOM نیست — لینکدین
    آن‌ها را obfuscate کرده). در پایان مهلت اگر ناوبری اصلی نبود، False.
    """
    from selenium.webdriver.common.by import By

    signed_in = '[data-testid="primary-nav"]'
    signed_out = ("a.nav__button-secondary, a.nav__button-tertiary, "
                  "form.login__form, form#login, .authwall, "
                  "input#username, input#password")
    try:
        driver.get("https://www.linkedin.com/feed/")
        end = time.time() + timeout
        while time.time() < end:
            time.sleep(1)
            url = (driver.current_url or "").lower()
            if ("login" in url or "authwall" in url
                    or "checkpoint" in url or "challenge" in url):
                return False
            try:
                if driver.find_elements(By.CSS_SELECTOR, signed_in):
                    return True
                if driver.find_elements(By.CSS_SELECTOR, signed_out):
                    return False
            except Exception:
                pass
        return False
    except Exception:
        return False


def open_own_tab(driver):
    """در حالت اتصال، یک تب تازه برای خودمان باز می‌کند تا تب‌های کاربر دست نخورد."""
    home = None
    try:
        home = driver.current_window_handle
        driver.switch_to.new_window("tab")
    except Exception:
        home = None
    return home


def close_own_tab(driver, home) -> None:
    """تب‌های خودمان را می‌بندد و به تب اصلی کاربر برمی‌گردد.

    مستقیم با CDP بسته می‌شود (بدون سوییچ تک‌تک) چون سوییچ روی پاپ‌آپ‌های
    گوگل گیر می‌کند. هر تب جداگانه try می‌شود.
    """
    if not home:
        return
    closed = 0
    try:
        res = driver.execute_cdp_cmd("Target.getTargets", {})
        infos = res.get("targetInfos", []) if isinstance(res, dict) else []
    except Exception:
        infos = []
    if infos:
        for info in infos:
            tid = info.get("targetId")
            if not tid or tid == home or info.get("type") != "page":
                continue
            try:
                driver.execute_cdp_cmd("Target.closeTarget", {"targetId": tid})
                closed += 1
            except Exception:
                continue
    else:
        # fallback: روش قدیمی سوییچ+بستن
        try:
            handles = list(driver.window_handles)
        except Exception:
            return
        for h in handles:
            if h == home:
                continue
            try:
                driver.switch_to.window(h)
                driver.close()
                closed += 1
            except Exception:
                continue
    try:
        driver.switch_to.window(home)
    except Exception:
        pass
    if closed:
        log(f"   {closed} تب بسته شد.")


# --------------------------------------------------------------------------- #
# ورود
# --------------------------------------------------------------------------- #
def login(driver, email: str, password: str, wait: int = 25) -> bool:
    from selenium.webdriver.common.by import By

    if already_logged_in(driver):
        log("✅ از قبل وارد لینکدین شده‌ای — بدون نیاز به رمز.")
        return True

    if not email or not password or email == "your_email@example.com":
        log("❌ رمز لینکدین ثبت نشده — ورود انجام نشد.")
        log("💡 اگر اکانتت با «ورود با گوگل» ساخته شده و رمزی نداری: کروم خودت را "
            "با پورت دیباگ باز کن (راهنما در تب LinkedIn داشبورد) و حالت «اتصال» را بزن؛ "
            "یا در همان مرورگر خودت وارد لینکدین شو و دوباره اجرا کن.")
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
    """پروفایل خود کاربر را از /in/me/ می‌خواند (به URL واقعی ریدایرکت می‌شود).

    کلاس‌های لینکدین obfuscate شده‌اند، پس به‌جای کلاس به ساختار تکیه می‌کنیم:
    نام = h2 کارت بالایی، ارتباط‌ها = لینک connections، سمت/موقعیت = تجزیهٔ
    متن کارت بالایی (نام | … | موقعیتِ دارای ویرگول | …).
    """
    from selenium.webdriver.common.by import By

    try:
        driver.get("https://www.linkedin.com/in/me/")
        time.sleep(5)
    except Exception:
        pass
    url = (driver.current_url or "").split("?")[0]
    data = {"name": "نامشخص", "headline": "—", "location": "—",
            "connections": "—", "url": url,
            "scraped_at": f"{_dt.datetime.now():%Y-%m-%d %H:%M}"}

    try:
        cands = driver.find_elements(By.CSS_SELECTOR, "main section")
    except Exception:
        return data
    # کارت بالایی پروفایل: تنها سکشنی که هم h2 نام و هم لینک Contact info دارد
    # (سکشن‌های About/Activity/Experience هیچ‌کدام لینک contact-info ندارند)
    top = None
    for s in cands:
        try:
            if (s.find_element(By.TAG_NAME, "h2").text or "").strip() \
                    and s.find_elements(By.CSS_SELECTOR, "a[href*='contact-info']"):
                top = s
                break
        except Exception:
            continue
    if top is None:
        return data

    try:
        data["name"] = (top.find_element(By.TAG_NAME, "h2").text or "").strip() or "نامشخص"
    except Exception:
        pass

    try:
        lines = [ln.strip() for ln in (top.text or "").splitlines()]
        lines = [ln for ln in lines if ln and ln != "·" and "Contact info" not in ln]
        rest = lines[1:] if len(lines) > 1 else []
        headline = rest[0][:300] if rest else "—"
        location = "—"
        for ln in rest[1:4]:
            if "," in ln and len(ln) <= 80:
                location = ln
                break
        if location == "—" and "," in headline and len(headline) <= 80:
            location, headline = headline, "—"
        data["headline"] = headline or "—"
        data["location"] = location
    except Exception:
        pass

    try:
        conn = top.find_elements(By.CSS_SELECTOR, "a[href*='connections']")
        if conn:
            data["connections"] = (conn[0].text or "").strip() or "—"
    except Exception:
        pass
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
# کمک‌تابع مشترک: متن کارتِ یک لینک بدون تکیه به کلاس (کلاس‌ها obfuscate شده‌اند)
# --------------------------------------------------------------------------- #
def _card_lines_of(a):
    """بالاترین جدی که متنش عیناً همان متن اولین جدِ معنادار است.

    wrapperهای میانی همان متن را دارند؛ وقتی متن عوض شد یعنی کارت‌های
    همسایه قاطی شدند و همان‌جا توقف می‌کنیم. لینک‌های کم‌متن (آواتار) []
    می‌دهند تا لینک غنی همان کارت پردازش شود.
    """
    from selenium.webdriver.common.by import By

    def _lines(t):
        return [ln.strip() for ln in (t or "").splitlines() if ln.strip()]

    best = []
    try:
        el = a
        for _ in range(10):
            try:
                el = el.find_element(By.XPATH, "..")
            except Exception:
                break
            try:
                lines = _lines(el.text)
            except Exception:
                continue
            if not lines:
                continue
            if not best:
                best = lines
                continue
            if lines == best:
                continue
            break
        if len(best) >= 2:
            return best
    except Exception:
        pass
    try:
        own = _lines(a.text)
    except Exception:
        own = []
    return own if len(own) >= 2 else []


def _name_matches_slug(name: str, href: str) -> bool:
    import re as _re

    slug = href.rstrip("/").rsplit("/", 1)[-1].lower()
    toks = [t for t in _re.split(r"[^a-z]+", name.lower()) if len(t) > 1]
    if not toks:
        return True  # نام غیرلاتین قابل تطبیق نیست؛ به ایزولاسیون کارت اعتماد کن
    return all(t in slug for t in toks)


def _card_lines_growing(a, cap: int = 1000):
    """متن کارت با قانون رشد: تا وقتی خط اول همان است و متن در حد کارت
    است بالاتر می‌رویم؛ با عوض شدن خط اول (ورود همسایه) یا انفجار طول، توقف.
    برای لینک‌هایی که خودشان عنوان کامل‌اند (مثل عنوان آگهی شغلی)."""
    from selenium.webdriver.common.by import By

    def _lines(t):
        return [ln.strip() for ln in (t or "").splitlines() if ln.strip()]

    best = []
    try:
        el = a
        for _ in range(12):
            try:
                el = el.find_element(By.XPATH, "..")
            except Exception:
                break
            try:
                raw = el.text or ""
            except Exception:
                continue
            if not raw.strip():
                continue
            if len(raw) > cap:
                break
            lines = _lines(raw)
            if not lines:
                continue
            if not best:
                best = lines
                continue
            if lines[0] == best[0] and len(lines) >= len(best):
                best = lines
                continue
            break
    except Exception:
        pass
    return best


# --------------------------------------------------------------------------- #
# ریکروتریابی (Recruiter discovery)
# --------------------------------------------------------------------------- #
def search_recruiters(driver, keywords: str, location: str = "Germany",
                      limit: int = 25) -> list:
    """اسکن نتایج جستجوی افراد لینکدین و برداشتن ریکروترها / تالنت‌ها.

    به کلاس‌ها تکیه نمی‌کند (لینکدین آن‌ها را obfuscate کرده): لینک‌های /in/
    در کل سند پیدا می‌شوند و کارت هر لینک = نزدیک‌ترین li بالادستی آن است،
    پس متن یک کارت هرگز به کارت همسایه نشت نمی‌کند.
    """
    import re

    from selenium.webdriver.common.by import By

    kw = quote_plus(keywords)
    url = (f"https://www.linkedin.com/search/results/people/?keywords={kw}"
           f"&origin=GLOBAL_SEARCH_HEADER")
    log(f"🔎 ریکروتریابی: «{keywords}» در {location}")
    driver.get(url)
    # انتظار هوشمند برای نتایج (تا ۱۵ ثانیه) به‌جای sleep ثابت —
    # گاهی لینکدین کند است یا اول دیوار امنیتی نشان می‌دهد.
    # عمداً در کل سند جستجو می‌شود نه فقط main، چون چیدمان نتایج ثابت نیست.
    def _has_results() -> bool:
        try:
            return bool(driver.find_elements(By.CSS_SELECTOR, "a[href*='/in/']"))
        except Exception:
            return False

    end = time.time() + 15
    while time.time() < end and not _has_results():
        time.sleep(1)
    if not _has_results():
        # یک بار رفرش؛ گاهی اولین بار خالی برمی‌گردد
        log("   ⏳ نتیجه‌ای نیامد — یک بار رفرش می‌کنم…")
        try:
            driver.refresh()
        except Exception:
            pass
        end = time.time() + 8
        while time.time() < end and not _has_results():
            time.sleep(1)
    time.sleep(2)

    skip = {"connect", "follow", "message", "more", "show more"}

    def clean_degree(s: str) -> str:
        return re.sub(r"\s*[•·]\s*(1st|2nd|3rd|you)\b.*$", "", s,
                      flags=re.IGNORECASE).strip()

    def is_noise(ln: str) -> bool:
        low = ln.strip().lower()
        return (not low or low in skip
                or re.fullmatch(r"[•·]?\s*(1st|2nd|3rd|you)", low) is not None)

    def _lines(t):
        return [ln.strip() for ln in (t or "").splitlines() if ln.strip()]

    # کارت هر لینک = بالاترین جدی که متنش عیناً همان متن اولین جدِ معنادار است
    # (wrapperهای میانی همان متن را دارند؛ وقتی متن عوض شد یعنی کارت‌های
    # همسایه قاطی شدند و همان‌جا توقف می‌کنیم — چون کارت‌ها li نیستند)
    def card_lines(a):
        best = _card_lines_of(a)
        return best

    def name_matches_slug(name: str, href: str) -> bool:
        return _name_matches_slug(name, href)

    try:
        links = driver.find_elements(By.CSS_SELECTOR, "a[href*='/in/']")
    except Exception:
        links = []
    log(f"   {len(links)} لینک پروفایل دیده شد.")

    found, seen = [], set()
    for a in links:
        if len(found) >= limit:
            break
        try:
            href = (a.get_attribute("href") or "").split("?")[0]
        except Exception:
            continue
        if not href or "/in/" not in href or href in seen:
            continue
        slug = href.rstrip("/").rsplit("/", 1)[-1]
        # لینک‌های سیستمی/سرویس (نه پروفایل واقعی) را رد کن
        if not slug or slug.lower() in ("in",):
            continue
        try:
            lines = card_lines(a)
        except Exception:
            continue
        if not lines:
            continue
        name = clean_degree(lines[0])
        if not name or len(name) > 80:
            continue
        # نگهبان نهایی: توکن‌های نام باید در slug آدرس باشند، وگرنه متن
        # کارت همسایه است (هر کارت دقیقاً یک لینک «غنی» با نام+سمت دارد
        # که همین‌جا درست پردازش می‌شود)
        if not name_matches_slug(name, href):
            continue
        seen.add(href)
        headline, location = "", ""
        for ln in lines[1:]:
            if is_noise(ln):
                continue
            headline = ln[:200]
            break
        for ln in lines[1:]:
            if ln == headline or is_noise(ln):
                continue
            if "," in ln and len(ln) <= 80:
                location = ln
                break
        role = (headline + " " + name).lower()
        is_rec = any(k in role for k in (
            "recruit", "talent", "hiring", "people partner", "hr ", "hr-",
            "headhunter", "staffing", "sourcing", "personalleiter",
            "personalreferent"))
        found.append({
            "name": name, "headline": headline, "location": location, "url": href,
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
    """جستجوی آگهی در لینکدین — مستقل از کلاس (obfuscate شده‌اند).

    لینک‌های /jobs/view/ در کل سند پیدا می‌شوند؛ عنوان/شرکت از متن کارتِ
    همان لینک (بالاترین جدِ هم‌متن) تجزیه می‌شود.
    """
    import re

    from selenium.webdriver.common.by import By

    kw, loc = quote_plus(keywords), quote_plus(location)
    url = (f"https://www.linkedin.com/jobs/search/?keywords={kw}"
           f"&location={loc}&f_WT=2")   # فقط فاصله‌کاری/ریموت‌پذیر
    log(f"💼 جستجوی شغل: «{keywords}» در {location}")
    driver.get(url)

    def _has_results() -> bool:
        try:
            return bool(driver.find_elements(By.CSS_SELECTOR, "a[href*='/jobs/view/']"))
        except Exception:
            return False

    end = time.time() + 15
    while time.time() < end and not _has_results():
        time.sleep(1)
    if not _has_results():
        log("   ⏳ نتیجه‌ای نیامد — یک بار رفرش می‌کنم…")
        try:
            driver.refresh()
        except Exception:
            pass
        end = time.time() + 8
        while time.time() < end and not _has_results():
            time.sleep(1)
    time.sleep(2)

    noise = {"promoted", "easy apply", "apply", "save", "saved", "dismiss",
             "actively hiring"}
    time_re = re.compile(r"(\d+\s*(m(in)?|h(ours?)?|d(ays?)?|w(eeks?)?|mo(nths?)?)"
                         r"\s*ago|reposted|just now)", re.I)
    meta_re = re.compile(r"(applicant|connection|alum|be among the first)", re.I)

    try:
        links = driver.find_elements(By.CSS_SELECTOR, "a[href*='/jobs/view/']")
    except Exception:
        links = []
    log(f"   {len(links)} لینک آگهی دیده شد.")

    jobs, seen = [], set()
    for a in links:
        if len(jobs) >= limit:
            break
        try:
            href = (a.get_attribute("href") or "").split("?")[0]
            atext = " ".join((a.text or "").split())
        except Exception:
            continue
        slug = href.rstrip("/").rsplit("/", 1)[-1]
        if not href or "/jobs/view/" not in href or not slug.isdigit() or href in seen:
            continue
        # عنوان فقط و فقط متن خود لینک است (نه متن کارت — آنجا دکمه‌ها قاطی‌اند)
        if not atext or len(atext) > 160 or atext.lower() in noise:
            continue
        title = atext
        try:
            lines = _card_lines_growing(a)
        except Exception:
            lines = []
        company = ""
        for ln in lines:
            low = ln.lower()
            if ln == title or not low or low in noise:
                continue
            if time_re.search(low):
                continue
            if meta_re.search(low) and len(low) < 70:
                continue
            company = ln[:120]
            break
        seen.add(href)
        jobs.append({
            "title": title, "company": company or "—", "url": href,
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
        if any(k in state for k in ("Unsave", "Saved", "ذخیره", "حذف")) \
                and "Save the job" not in state:
            log("ℹ️ این آگهی قبلاً ذخیره شده بود.")
            _record_job(driver, job_url)
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
    """عنوان/شرکت آگهی را می‌خواند — اول از <title> («عنوان | شرکت | LinkedIn»)
    چون کلاس‌های DOM مدام obfuscate می‌شوند؛ بعد لینک /company/."""
    from selenium.webdriver.common.by import By

    title, comp = "", ""
    try:
        parts = [p.strip() for p in (driver.title or "").split("|")]
        if parts and parts[0] and "LinkedIn" not in parts[0]:
            title = parts[0][:160]
        if len(parts) > 1 and "LinkedIn" not in parts[1]:
            comp = parts[1][:120]
    except Exception:
        pass
    if not comp:
        try:
            for a in driver.find_elements(By.CSS_SELECTOR, "a[href*='/company/']"):
                try:
                    t = (a.text or "").strip()
                    h = a.get_attribute("href") or ""
                except Exception:
                    continue
                if t and len(t) < 80 and "/life" not in h and "careers" not in h.lower():
                    comp = t
                    break
        except Exception:
            pass
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
    ap.add_argument("--mode", choices=["auto", "launch", "attach"], default="auto",
                    help="attach=وصل شدن به کروم خودت (پورت 9222، برای اکانت‌های گوگلی) · "
                         "launch=همیشه مرورگر تازه · auto=اول اتصال، اگر نبود مرورگر تازه")
    ap.add_argument("--port", type=int, default=DEBUG_PORT, help="پورت دیباگ کروم خودت")
    args = ap.parse_args(argv)

    if args.action == "log":
        print(read_log())
        return 0

    email, password = get_credentials()
    driver = None
    attached = False
    home_tab = None
    rc = 0
    try:
        if args.mode in ("auto", "attach") and not args.headless:
            driver = try_attach(args.port)
            attached = driver is not None
        if driver is None:
            if args.mode == "attach":
                log("❌ کرومی با پورت دیباگ پیدا نشد.")
                log("   ۱) کروم را کاملاً ببند  ۲) این را در PowerShell اجرا کن:")
                log('   & "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" '
                    f"--remote-debugging-port={args.port}")
                log("   ۳) وارد لینکدین شو (با گوگل یا رمز)  ۴) دوباره اجرا بزن.")
                return 4
            driver = get_driver(headless=args.headless)

        if attached:
            # تب‌های خود کاربر دست نمی‌خورد؛ در یک تب تازه کار می‌کنیم
            home_tab = open_own_tab(driver)
            log("ℹ️ در یک تب تازه کار می‌کنم؛ تب‌های خودت دست نمی‌خورد.")

        logged = args.no_login
        if not args.no_login:
            logged = login(driver, email, password)

        if args.action in ("profile", "login"):
            if not logged:
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
        if driver and attached:
            # مرورگر خود کاربر هرگز بسته نمی‌شود — فقط تب خودمان را جمع می‌کنیم
            close_own_tab(driver, home_tab)
            try:
                del driver
            except Exception:
                pass
            log("ℹ️ اتصال قطع شد؛ مرورگر خودت باز می‌ماند.")
        elif driver:
            try:
                driver.quit()
            except Exception:
                pass
    return rc


if __name__ == "__main__":
    sys.exit(main())
