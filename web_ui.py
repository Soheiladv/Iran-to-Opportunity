#!/usr/bin/env python3
"""
Migration Hunter — Web UI
داشبورد وب سبک، بدون هیچ وابستگی جدید (فقط کتابخانه‌ی استاندارد پایتون).
اگر openpyxl نصب باشد (که برای build_dashboard.py هم لازم است)، خروجی‌های
اکسل هم به‌صورت جدول در مرورگر نمایش داده می‌شوند؛ در غیر این صورت فقط
لینک دانلود نشان داده می‌شود.

اجرا:
    python web_ui.py            # پیش‌فرض روی پورت 8877
    python web_ui.py --port 9000

سپس در مرورگر باز کن: http://127.0.0.1:8877
"""
import argparse
import html
import json
import os
import subprocess
import sys
import threading
import time
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

BASE = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE, "output")
DASHBOARD_DIR = os.path.join(BASE, "dashboard")
MEM_DIR = os.path.join(BASE, "memory")
CONFIG_PATH = os.path.join(BASE, "config.json")
ENV_PATH = os.path.join(BASE, ".env")
GITIGNORE_PATH = os.path.join(BASE, ".gitignore")
SOURCES_PATH = os.path.join(BASE, "sources.json")

import unified_report

try:
    import job_crawler as _jc  # برای استفاده‌ی مجدد از load_sources/save_sources، بدون تکرار کد
    HAS_JOB_CRAWLER = True
except Exception:
    HAS_JOB_CRAWLER = False

try:
    import education_crawler as _ec
    HAS_EDU_CRAWLER = True
except Exception:
    HAS_EDU_CRAWLER = False

try:
    import visa_tracker as _vt
    HAS_VISA = True
except Exception:
    HAS_VISA = False

try:
    from openpyxl import load_workbook, Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ══════════════════════════════════════════════════════════════
# دو مسیر جدا: کاریابی و تحصیل
#
# هر مسیر پایپ‌لاین خودش را دارد و منابع خودش را از sources.json
# (track=job یا track=education) می‌خواند. انتخاب مسیر در صفحهٔ اصلی
# با ?track= انجام می‌شود و در سراسر داشبورد (تب‌ها، پایپ‌لاین، گزارش‌ها)
# حفظ می‌شود — یعنی همه‌جا یا کار می‌بینی یا تحصیل، قاطی نمی‌شود.
# ══════════════════════════════════════════════════════════════
TRACKS = {
    "job": {
        "label": "کاریابی", "emoji": "🔍", "crawler": "job_crawler.py",
        "results": "CRAWLER_RESULTS.json", "kind": "آگهی",
        "xlsx_prefix": "Job_Crawler", "accent": "var(--teal)",
    },
    "education": {
        "label": "تحصیل", "emoji": "🎓", "crawler": "education_crawler.py",
        "results": "EDUCATION_RESULTS.json", "kind": "برنامه",
        "xlsx_prefix": "Education_Crawler", "accent": "var(--amber)",
    },
}
DEFAULT_TRACK = "job"


def track_of(track=None):
    """نام مسیر را امن می‌کند — هر ورودی ناشناخته به مسیر پیش‌فرض می‌افتد."""
    t = (track or DEFAULT_TRACK).strip().lower()
    return t if t in TRACKS else DEFAULT_TRACK


def track_meta(track=None):
    return TRACKS[track_of(track)]


def track_results_path(track=None):
    return os.path.join(MEM_DIR, track_meta(track)["results"])


def with_track(path, track):
    """مسیر داخلی را با ?track= کامل می‌کند تا تب‌ها مسیر فعال را نگه دارند."""
    sep = "&" if "?" in path else "?"
    return f"{path}{sep}track={track_of(track)}"


# پایپ‌لاین اجرا — هر مسیر مراحل خودش را دارد
PIPELINE = {
    "job": [
        {"key": "email_analyze", "name": "تحلیل ایمیل شغلی", "script": "email_analyzer.py", "emoji": "📧"},
        {"key": "email_excel", "name": "ساخت Excel ایمیل", "script": "email_dashboard.py", "emoji": "📊"},
        {"key": "job_search", "name": "جستجوی خودکار کار", "script": "job_crawler.py", "emoji": "🔍"},
        {"key": "followup", "name": "یادآوری پیگیری", "script": "followup_reminder.py", "emoji": "⏰"},
        {"key": "dashboard", "name": "ساخت داشبورد اصلی", "script": "build_dashboard.py", "emoji": "📈"},
    ],
    "education": [
        {"key": "email_analyze", "name": "تحلیل ایمیل تحصیلی", "script": "email_analyzer.py", "emoji": "📧"},
        {"key": "edu_search", "name": "جستجوی برنامه‌های تحصیلی", "script": "education_crawler.py", "emoji": "🎓"},
        {"key": "email_excel", "name": "ساخت Excel ایمیل", "script": "email_dashboard.py", "emoji": "📊"},
        {"key": "followup", "name": "یادآوری پیگیری", "script": "followup_reminder.py", "emoji": "⏰"},
        {"key": "dashboard", "name": "ساخت داشبورد اصلی", "script": "build_dashboard.py", "emoji": "📈"},
    ],
}


def pipeline_steps(track=None):
    return PIPELINE[track_of(track)]


# ── وضعیت اجرای پایپ‌لاین در حافظه (thread-safe به‌قدر کافی برای یک کاربر محلی) ──
run_state = {
    "running": False, "log": [], "started_at": None, "finished_at": None,
    "step_index": 0, "step_total": 0, "current_step": "", "track": DEFAULT_TRACK,
    # 🔴 پنل زنده — اسکریپت الان کدام منبع و کدام آدرس را باز می‌کند
    "live": {
        "active": False, "track": DEFAULT_TRACK,
        "source": "", "source_idx": 0, "source_total": 0, "country": "",
        "url": "", "keyword": "", "url_idx": 0, "url_total": 0,
        # شمارش روی تک‌تک آدرس‌ها — نوار پیوسته حرکت می‌کند
        "url_done": 0, "url_total_all": 0, "pct": 0,
        "jobs_found": 0, "sources_done": [], "urls": [],
    },
}
run_lock = threading.Lock()

import re as _re
import urllib.parse

# ── پروتکل پارس لاگ ─────────────────────────────────────────────
# این regex ها دقیقاً با قالب print های job_crawler.py هم‌خوان‌اند.
# اگر آن قالب‌ها را عوض کردی، این‌ها را هم عوض کن.
_LIVE_PATTERNS = {
    # [███░░░] آدرس 12/191 (6.3%) · سایت 3/25 · 🧲 45 آگهی   ← دقیق‌ترین منبع
    "bar": _re.compile(
        r"\[[█▏░ ]*\]\s*آدرس\s*(\d+)\s*/\s*(\d+)\s*\(([\d.]+)%\)"
        r".*?سایت\s*(\d+)\s*/\s*(\d+).*?🧲\s*(\d+)\s*(\S+)"),
    # 📡 [3/25] Seek AU (AU) — در حال بررسی…
    "source": _re.compile(r"📡\s*\[(\d+)/(\d+)\]\s*(.+?)\s*\(([A-Z]{2}|[^\s()]+)\)\s*—+\s*در حال بررسی"),
    # 🔎 [2/5] کلیدواژه: «it manager» ← https://…
    "fetch": _re.compile(r"🔎\s*\[(\d+)/(\d+)\]\s*کلیدواژه:\s*«(.+?)»\s*←\s*(\S+)"),
    # ✅ [2/5] نام سایت · «it manager» · 12 آگهی · 4300ms
    "url_done": _re.compile(
        r"✅\s*\[(\d+)/(\d+)\]\s*(.+?)\s*·\s*«(.+?)»\s*·\s*(\d+)\s*(\S+)\s*·\s*(\d+)\s*ms"),
    # ⚠️ [3/5] نام سایت · «kw» · پاسخی نیامد (…)
    "url_fail": _re.compile(r"⚠️\s*\[(\d+)/(\d+)\]\s*(.+?)\s*·\s*«(.+?)»\s*·\s*پاسخی نیامد"),
    # 🏁 مجموعاً 12 آگهی یافت شد از نام سایت · 40.3s
    "site_done": _re.compile(r"🏁\s*مجموعاً\s*(\d+)\s*(\S+)\s*یافت شد از\s*(.+?)\s*·\s*([\d.]+)\s*s"),
}

MAX_LIVE_URLS = 400  # جدول زنده بی‌نهایت رشد نکند


def _live_reset():
    """وضعیت زنده را صفر می‌کند — موقع شروع هر اجرا."""
    run_state["live"] = {
        "active": False, "track": run_state.get("track", DEFAULT_TRACK),
        "source": "", "source_idx": 0, "source_total": 0, "country": "",
        "url": "", "keyword": "", "url_idx": 0, "url_total": 0,
        "url_done": 0, "url_total_all": 0, "pct": 0,
        "jobs_found": 0, "sources_done": [], "urls": [],
    }


def _update_live_state(line):
    """خطوط لاگِ در حال استریم را پارس می‌کند و پنل جستجوی زنده را به‌روز می‌کند."""
    live = run_state["live"]

    # نوار progress — دقیق‌تر از هر محاسبه‌ای، چون خود کراولر شمرده
    m = _LIVE_PATTERNS["bar"].search(line)
    if m:
        live["url_done"] = int(m.group(1))
        live["url_total_all"] = int(m.group(2))
        live["pct"] = float(m.group(3))
        live["site_done"] = int(m.group(4))
        live["source_total"] = int(m.group(5))
        live["jobs_found"] = int(m.group(6))
        live["active"] = True
        return

    m = _LIVE_PATTERNS["source"].search(line)
    if m:
        live["active"] = True
        live["source_idx"], live["source_total"] = int(m.group(1)), int(m.group(2))
        live["source"] = m.group(3).strip()
        live["country"] = m.group(4).strip()
        live["url"], live["keyword"] = "", ""
        live["url_idx"] = live["url_total"] = 0
        return

    m = _LIVE_PATTERNS["fetch"].search(line)
    if m:
        live["url_idx"], live["url_total"] = int(m.group(1)), int(m.group(2))
        live["keyword"], live["url"] = m.group(3), m.group(4)
        return

    # یک آدرس تمام شد → ردیف جدید در جدول زنده
    for key, status in (("url_done", "ok"), ("url_fail", "blocked")):
        m = _LIVE_PATTERNS[key].search(line)
        if not m:
            continue
        name, kw = m.group(3).strip(), m.group(4)
        jobs = int(m.group(5)) if status == "ok" else 0
        ms = int(m.group(7)) if status == "ok" else 0
        live["urls"].append({
            "name": name, "keyword": kw, "url": live["url"] or kw,
            "jobs": jobs, "ms": ms, "status": status,
            "idx": len(live["urls"]) + 1,
        })
        if len(live["urls"]) > MAX_LIVE_URLS:
            del live["urls"][: len(live["urls"]) - MAX_LIVE_URLS]
        return

    m = _LIVE_PATTERNS["site_done"].search(line)
    if m:
        found, name, secs = int(m.group(1)), m.group(3).strip(), float(m.group(4))
        live["sources_done"].append({"name": name, "jobs": found, "secs": secs})
        return


def app_bank_section():
    """خلاصه بانک درخواست‌ها را برای صفحه اصلی برمی‌گرداند."""
    bank_path = os.path.join(MEM_DIR, "APPLICATION_BANK.json")
    if not os.path.exists(bank_path):
        return ""
    try:
        with open(bank_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return ""
    apps = data.get("applications", [])
    if not apps:
        return ""
    status_counts = {}
    overdue = 0
    now = datetime.now()
    for a in apps:
        st = a.get("status", "?")
        status_counts[st] = status_counts.get(st, 0) + 1
        dl = a.get("reply_deadline")
        if dl and st in ("SENT", "FOLLOW_UP"):
            try:
                if datetime.fromisoformat(dl) < now:
                    overdue += 1
            except Exception:
                pass
    STATUS_EMOJI = {"SENT": "📤", "REPLIED": "📬", "FOLLOW_UP": "🔁",
                    "REJECTED": "❌", "INTERVIEW": "🎤", "OFFER": "🎉"}
    cards_html = ""
    for st, cnt in sorted(status_counts.items()):
        emoji = STATUS_EMOJI.get(st, "📋")
        cards_html += (f'<div class="card" style="text-align:center; padding:12px 8px;">'
                        f'<div style="font-size:1.4rem">{emoji}</div>'
                        f'<div style="font-size:1.6rem; font-weight:700">{cnt}</div>'
                        f'<div style="color:var(--muted); font-size:.85rem">{st}</div></div>')
    overdue_html = (f'<div style="margin-top:10px; padding:10px 14px; background:#fef3ed; '
                     f'border-radius:var(--radius); color:var(--err); font-size:.9rem">'
                     f'⚠️ <strong>{overdue} درخواست</strong> از مهلت پاسخ گذشته — نیاز به پیگیری فوری</div>')
    return f"""<section>
  <h2>📋 بانک درخواست‌ها</h2>
  <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(100px,1fr)); gap:10px">{cards_html}</div>
  {overdue_html if overdue else ''}
</section>"""


def load_applicants():
    if not os.path.exists(CONFIG_PATH):
        return []
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f).get("applicants", [])
    except Exception:
        return []


def save_applicants(applicants):
    data = {"version": "1.0", "applicants": applicants}
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    try:
        os.chmod(CONFIG_PATH, 0o600)
    except OSError:
        pass


def upsert_applicant(fields):
    """اضافه یا به‌روزرسانی یک متقاضی در config.json بر اساس id."""
    app_id = fields["id"].strip().lower()
    applicants = load_applicants()
    entry = {
        "id": app_id,
        "name": fields.get("name", "").strip() or app_id,
        "name_fa": fields.get("name_fa", "").strip() or fields.get("name", "").strip() or app_id,
        "emoji": fields.get("emoji", "").strip() or "👤",
        "profession": fields.get("profession", "").strip(),
        "keywords": [k.strip() for k in fields.get("keywords", "").split(",") if k.strip()],
        "linkedin": fields.get("linkedin", "").strip(),
        "email": fields.get("email", "").strip(),
        "english": fields.get("english", "").strip(),
        "german": fields.get("german", "").strip(),
    }
    for i, a in enumerate(applicants):
        if a.get("id") == app_id:
            applicants[i] = entry
            break
    else:
        applicants.append(entry)
    save_applicants(applicants)
    return app_id


def read_env():
    """خواندن .env به‌صورت دیکشنری key→value. کامنت‌ها و خطوط خالی رد می‌شوند."""
    data = {}
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if not s or s.startswith("#") or "=" not in s:
                    continue
                k, _, v = s.partition("=")
                data[k.strip()] = v.strip()
    return data


def write_env_updates(updates):
    """
    به‌روزرسانی چند کلید در .env بدون دست‌زدن به بقیه‌ی فایل.
    کلیدهایی که مقدار خالی برایشان داده نشده (رمز خالی یعنی «تغییر نده») نادیده گرفته می‌شوند.
    """
    updates = {k: v for k, v in updates.items() if v not in (None, "")}
    if not updates:
        return
    lines = []
    if os.path.exists(ENV_PATH):
        with open(ENV_PATH, "r", encoding="utf-8") as f:
            lines = f.readlines()

    seen = set()
    new_lines = []
    for line in lines:
        s = line.rstrip("\n")
        if "=" in s and not s.strip().startswith("#"):
            k = s.split("=", 1)[0].strip()
            if k in updates:
                new_lines.append(f"{k}={updates[k]}\n")
                seen.add(k)
                continue
        new_lines.append(line if line.endswith("\n") else line + "\n")

    if new_lines and not new_lines[-1].endswith("\n"):
        new_lines[-1] += "\n"
    for k, v in updates.items():
        if k not in seen:
            new_lines.append(f"{k}={v}\n")

    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.writelines(new_lines)
    try:
        os.chmod(ENV_PATH, 0o600)  # فقط خودت بتوانی بخوانی/بنویسی
    except OSError:
        pass


def gitignore_covers(name):
    """چک می‌کند آیا .gitignore الگویی دارد که این فایل را پوشش دهد (ساده و تقریبی)."""
    if not os.path.exists(GITIGNORE_PATH):
        return False
    with open(GITIGNORE_PATH, "r", encoding="utf-8") as f:
        patterns = [p.strip() for p in f if p.strip() and not p.strip().startswith("#")]
    return name in patterns


def read_sources():
    """
    منابع جستجو را برمی‌گرداند. اگر job_crawler.py قابل import باشد از
    load_sources همان‌جا استفاده می‌شود (که خودش sources.json را می‌سازد
    اگر نبود). اگر نه، مستقیم فایل sources.json خوانده می‌شود؛ اگر آن هم
    نبود، None برمی‌گردد (یعنی هنوز هیچ‌جا ساخته نشده).
    """
    if HAS_JOB_CRAWLER:
        try:
            return _jc.load_sources()
        except Exception:
            pass
    if os.path.exists(SOURCES_PATH):
        try:
            with open(SOURCES_PATH, "r", encoding="utf-8") as f:
                return json.load(f).get("sources", [])
        except Exception:
            return []
    return None


def write_sources(sources):
    if HAS_JOB_CRAWLER:
        try:
            _jc.save_sources(sources)
            return
        except Exception:
            pass
    with open(SOURCES_PATH, "w", encoding="utf-8") as f:
        json.dump({"sources": sources}, f, ensure_ascii=False, indent=2)


def list_files(folder, exts=None):
    if not os.path.isdir(folder):
        return []
    items = []
    for root, _dirs, files in os.walk(folder):
        for fn in files:
            if exts and not fn.lower().endswith(tuple(exts)):
                continue
            full = os.path.join(root, fn)
            rel = os.path.relpath(full, BASE)
            items.append({
                "rel": rel,
                "name": fn,
                "mtime": os.path.getmtime(full),
            })
    return sorted(items, key=lambda x: x["mtime"], reverse=True)


MAX_LOG_LINES = 800  # جلوگیری از رشد بی‌نهایت لاگ در اجراهای طولانی


def _append_log(line):
    run_state["log"].append(line)
    if len(run_state["log"]) > MAX_LOG_LINES:
        del run_state["log"][: len(run_state["log"]) - MAX_LOG_LINES]


def run_pipeline_background(track=None, countries=None):
    """پایپ‌لاین مسیر انتخاب‌شده را در پس‌زمینه اجرا می‌کند.

    track مسیر است (job یا education) و steps از همان مسیر خوانده می‌شود،
    پس «اجرای تحصیل» هیچ‌وقت اسکریپت کاریابی را صدا نمی‌زند.
    countries لیست اختیاری کد کشورهایی است که باید --country بگیرند.
    """
    t = track_of(track)
    steps = pipeline_steps(t)
    live_steps = ("job_search", "edu_search")  # مرحله‌هایی که پنل زنده دارند

    with run_lock:
        if run_state["running"]:
            return
        run_state["running"] = True
        run_state["track"] = t
        run_state["log"] = []
        run_state["started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        run_state["finished_at"] = None
        run_state["step_index"] = 0
        run_state["step_total"] = len(steps)
        run_state["current_step"] = ""
    _live_reset()
    run_state["live"]["track"] = t

    for i, step in enumerate(steps, 1):
        run_state["step_index"] = i
        run_state["current_step"] = f"{step['emoji']} {step['name']}"
        script_path = os.path.join(BASE, step["script"])
        if not os.path.exists(script_path):
            _append_log(f"⏭️  {step['emoji']} {step['name']} — فایل {step['script']} پیدا نشد، رد شد")
            continue
        _append_log(f"▶ [{i}/{len(steps)}] {step['emoji']} {step['name']} در حال اجرا…")
        run_state["live"]["active"] = step["key"] in live_steps

        # مسیر انتخابی به کراولر می‌رود و کشورها هم فقط برای همان مسیر
        cmd = [sys.executable, "-u", script_path]
        if step["key"] in live_steps:
            cmd += ["--track", t]
            if countries:
                for c in countries:
                    cmd += ["--country", c]

        try:
            # -u = خروجی بدون بافر، تا خط‌به‌خط همین‌جا زنده دیده شود
            proc = subprocess.Popen(
                cmd,
                cwd=BASE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1,
                encoding="utf-8", errors="replace",
            )
            start_ts = time.monotonic()
            for line in proc.stdout:
                line = line.rstrip("\r\n")
                if line.strip():
                    _append_log("   " + line)
                    _update_live_state(line.strip())
                if time.monotonic() - start_ts > 1800:  # سقف ۳۰ دقیقه برای هر مرحله
                    proc.kill()
                    _append_log(f"⏱️ {step['emoji']} {step['name']} — بیش از حد طول کشید، متوقف شد")
                    break
            proc.wait(timeout=5)
            ok = proc.returncode == 0
            _append_log(f"{'✅' if ok else '❌'} {step['emoji']} {step['name']} "
                        f"{'با موفقیت انجام شد' if ok else 'با خطا مواجه شد (کد خروج ' + str(proc.returncode) + ')'}")
        except Exception as e:
            _append_log(f"❌ {step['emoji']} {step['name']} — {e}")

    with run_lock:
        run_state["running"] = False
        run_state["current_step"] = ""
        run_state["live"]["active"] = False
        run_state["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


PAGE_STYLE = """
:root{
  --ink:#20302f; --paper:#f4f2ec; --panel:#ffffff; --line:#dcd8cc;
  --teal:#2f5e59; --teal-deep:#1c3a37; --amber:#b3722c; --amber-soft:#f1e2cd;
  --ok:#3d7a55; --err:#a8432f; --muted:#6f7a76;
  --radius:10px;
}
*{box-sizing:border-box}
body{
  margin:0; background:var(--paper); color:var(--ink);
  font-family:"Vazirmatn","IRANSans",Tahoma,"Segoe UI",sans-serif;
  direction:rtl; text-align:right; line-height:1.7;
}
a{color:var(--teal-deep)}
header.top{
  background:var(--teal-deep); color:#f4f2ec; padding:28px 32px;
  display:flex; justify-content:space-between; align-items:baseline; flex-wrap:wrap; gap:10px;
}
header.top h1{margin:0; font-size:1.4rem; font-weight:700}
header.top span.sub{color:#c9d8d4; font-size:.92rem}
main{max-width:920px; margin:0 auto; padding:28px 20px 60px}
section{margin-bottom:34px}
h2{font-size:1.05rem; color:var(--teal-deep); border-bottom:1px solid var(--line); padding-bottom:8px; margin-bottom:16px}
.grid{display:grid; grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:14px}
.card{
  background:var(--panel); border:1px solid var(--line); border-radius:var(--radius);
  padding:16px 18px;
}
.applicant .emoji{font-size:1.6rem}
.applicant .name{font-weight:700; margin-top:4px}
.applicant .meta{color:var(--muted); font-size:.85rem; margin-top:6px}
ol.steps{list-style:none; margin:0; padding:0; counter-reset:step}
ol.steps li{
  counter-increment:step; position:relative; padding:10px 0 10px 0; padding-right:44px;
  border-bottom:1px dashed var(--line);
}
ol.steps li:last-child{border-bottom:none}
ol.steps li::before{
  content:counter(step); position:absolute; right:0; top:8px;
  width:28px; height:28px; border-radius:50%; background:var(--amber-soft);
  color:var(--amber); font-weight:700; font-size:.85rem;
  display:flex; align-items:center; justify-content:center;
}
.btn{
  display:inline-block; background:var(--teal); color:#fff; border:none;
  padding:10px 22px; border-radius:var(--radius); font-size:.95rem; cursor:pointer;
  text-decoration:none;
}
.btn:hover{background:var(--teal-deep)}
.btn[disabled]{background:#a8b3b0; cursor:not-allowed}
#log{
  background:var(--teal-deep); color:#e6efec; font-family:monospace; font-size:.85rem;
  padding:14px 16px; border-radius:var(--radius); min-height:60px; white-space:pre-wrap;
  max-height:320px; overflow-y:auto;
}
table.files{width:100%; border-collapse:collapse; font-size:.92rem}
table.files td{padding:8px 6px; border-bottom:1px solid var(--line)}
table.files th{padding:8px 6px; border-bottom:2px solid var(--line); text-align:right; color:var(--teal-deep)}
table.files td.name a{text-decoration:none}
table.files td.date{color:var(--muted); font-size:.82rem; white-space:nowrap}
.empty{color:var(--muted); font-style:italic}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.25}}
.badge{display:inline-block; font-size:.78rem; padding:2px 8px; border-radius:20px; margin-right:6px}
.badge.ok{background:#e2efe6; color:var(--ok)}
.badge.err{background:#f4e2dd; color:var(--err)}
nav.tabs{display:flex; gap:6px; flex-wrap:wrap}
nav.tabs a{
  color:#c9d8d4; text-decoration:none; padding:6px 14px; border-radius:20px; font-size:.9rem;
}
nav.tabs a.active{background:rgba(255,255,255,.14); color:#fff}
.warn{
  background:#f4e2dd; border:1px solid #e0b3a6; color:var(--err); border-radius:var(--radius);
  padding:14px 16px; margin-bottom:20px;
}
form.inline{display:grid; grid-template-columns:1fr 1fr; gap:10px 16px; margin-top:10px}
form.inline label{display:flex; flex-direction:column; font-size:.85rem; color:var(--muted); gap:4px}
form.inline input, form.inline select{
  padding:8px 10px; border:1px solid var(--line); border-radius:6px; font-family:inherit; font-size:.95rem;
  background:#fff; color:var(--ink);
}
form.inline .full{grid-column:1 / -1}
.applicant-block{border:1px solid var(--line); border-radius:var(--radius); padding:16px 18px; margin-bottom:18px}
.applicant-block h3{margin:0 0 4px; font-size:1rem}
.hint{color:var(--muted); font-size:.82rem; margin-top:4px}

/* ── انتخابگر مسیر: کاریابی / تحصیل ── */
.track-switch{
  display:inline-flex; background:#fff; border:1px solid var(--line);
  border-radius:26px; padding:4px; gap:4px; margin-bottom:6px;
}
.track-switch a{
  text-decoration:none; color:var(--muted); padding:8px 20px; border-radius:22px;
  font-size:.92rem; font-weight:600; display:flex; align-items:center; gap:7px;
}
.track-switch a.active{background:var(--teal); color:#fff}
.track-switch a:hover:not(.active){background:var(--amber-soft); color:var(--ink)}

/* ── نوار پیشرفت ── */
.pbar{height:10px; background:#e6e2d6; border-radius:6px; overflow:hidden; margin:12px 0}
.pbar > span{display:block; height:100%; background:var(--teal); border-radius:6px; transition:width .35s ease}
.pbar.edu > span{background:var(--amber)}
.live-top{display:flex; justify-content:space-between; align-items:baseline; flex-wrap:wrap; gap:8px}
.live-nums{font-variant-numeric:tabular-nums; color:var(--muted); font-size:.88rem}
.live-now{
  background:var(--teal-deep); color:#e6efec; border-radius:var(--radius);
  padding:12px 16px; font-family:monospace; font-size:.83rem; direction:ltr; text-align:left;
  word-break:break-all;
}
.live-now .dim{color:#8fa8a3}
.live-now .kw{color:var(--amber-soft)}
.live-now .flag{color:#f0b8a8}

/* ── جدول زندهٔ آدرس‌ها ── */
table.live{width:100%; border-collapse:collapse; font-size:.84rem}
table.live th{
  padding:8px 6px; border-bottom:2px solid var(--line); text-align:right;
  color:var(--teal-deep); font-weight:700; white-space:nowrap;
}
table.live td{padding:6px; border-bottom:1px solid #eeebe2; vertical-align:top}
table.live td.idx{color:var(--muted); width:34px; font-variant-numeric:tabular-nums}
table.live td.num{font-variant-numeric:tabular-nums; text-align:left; white-space:nowrap; width:60px}
table.live td.ms{color:var(--muted); font-variant-numeric:tabular-nums; text-align:left; white-space:nowrap; width:70px}
table.live td.kw{white-space:nowrap; color:var(--ink); font-weight:600; width:120px}
table.live td.src{white-space:nowrap; color:var(--muted); width:150px}
table.live td.url{direction:ltr; text-align:left; word-break:break-all; color:var(--teal)}
table.live td.url a{color:var(--teal)}
.scroll-y{max-height:420px; overflow-y:auto}
.scroll-y table.live thead th{position:sticky; top:0; background:var(--paper); z-index:1}

/* ── چک‌لیست ویزا ── */
ul.check{list-style:none; margin:0; padding:0}
ul.check li{
  display:flex; align-items:flex-start; gap:10px; padding:9px 0;
  border-bottom:1px dashed var(--line);
}
ul.check li:last-child{border-bottom:none}
ul.check .box{
  width:21px; height:21px; border:2px solid var(--line); border-radius:5px; flex:0 0 21px;
  cursor:pointer; display:flex; align-items:center; justify-content:center;
  font-size:.8rem; color:transparent; margin-top:3px; text-decoration:none; background:#fff;
}
ul.check .box:hover{border-color:var(--teal)}
ul.check li.done .box{background:var(--ok); border-color:var(--ok); color:#fff}
ul.check li.done .t{text-decoration:line-through; color:var(--muted)}
ul.check .t{font-weight:600}
ul.check .h{color:var(--muted); font-size:.8rem}
ul.check .when{color:var(--ok); font-size:.78rem; white-space:nowrap}
ul.check form{display:inline}
.act-row{display:flex; align-items:flex-start; gap:10px; padding:10px 0; border-bottom:1px dashed var(--line)}
.act-row:last-child{border-bottom:none}
.act-row .cat{font-size:1.15rem; line-height:1.4}
.act-row .ttl{font-weight:600}
.act-row .dt{color:var(--muted); font-size:.79rem; margin-top:2px}
.act-row .dt .late{color:var(--err); font-weight:700}
.act-row.done .ttl{text-decoration:line-through; color:var(--muted)}
.kv{display:flex; gap:6px; flex-wrap:wrap; margin-top:3px}
.stat-row{display:flex; gap:20px; flex-wrap:wrap; margin:10px 0 4px}
.stat-row .st{background:#fff; border:1px solid var(--line); border-radius:var(--radius); padding:10px 16px; min-width:96px}
.stat-row .st b{display:block; font-size:1.35rem; color:var(--teal-deep); font-variant-numeric:tabular-nums}
.stat-row .st span{font-size:.78rem; color:var(--muted)}
"""


def nav_html(active, track=None):
    """نوار بالا: انتخابگر مسیر + تب‌ها.

    مسیر در ?track= جا می‌ماند تا هر تب با همان مسیر باز شود — یعنی اگر
    در حال دیدن «تحصیل» هستی و به فایل‌ها یا تنظیمات می‌روی، همان‌جا می‌مانی.
    """
    def cls(name):
        return "active" if name == active else ""
    t = track_of(track)
    return f"""<nav class="tabs">
  <a class="{cls('dashboard')}" href="{with_track('/', t)}">🔴 گزارش زنده</a>
  <a class="{cls('education')}" href="{with_track('/education', t)}">🎓 تحصیل</a>
  <a class="{cls('visa')}" href="{with_track('/visa', t)}">🛂 نورد ویزا</a>
  <a class="{cls('yield')}" href="{with_track('/yield', t)}">📈 بازده منابع</a>
  <a class="{cls('files')}" href="{with_track('/files', t)}">📁 فایل‌ها</a>
  <a class="{cls('reports')}" href="{with_track('/reports', t)}">📑 گزارش‌ها</a>
  <a class="{cls('settings')}" href="{with_track('/settings', t)}">⚙️ تنظیمات</a>
  <a class="{cls('about')}" href="{with_track('/about', t)}">ℹ️ توضیحات</a>
</nav>"""


# آیکون درون‌خطی — بدون فایل جدا، و مرورگر دیگر 404 برای /favicon.ico نمی‌زند
FAVICON = (
    '<link rel="icon" href="data:image/svg+xml,'
    '%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 viewBox=%270 0 32 32%27%3E'
    '%3Crect width=%2732%27 height=%2732%27 rx=%277%27 fill=%27%231c3a37%27/%3E'
    '%3Cpath d=%27M16 5l9 5-9 5-9-5 9-5zm-9 10l9 5 9-5v3l-9 5-9-5v-3z%27 fill=%27%23b3722c%27/%3E'
    '%3C/svg%3E">'
)


def track_switch_html(active_track, base_path="/"):
    """کلید انتخاب مسیر: کاریابی ⇄ تحصیل."""
    def link(t):
        cls = "active" if t == active_track else ""
        m = TRACKS[t]
        return (f'<a class="{cls}" href="{with_track(base_path, t)}">'
                f'{m["emoji"]} {m["label"]}</a>')

    return f"""<div class="track-switch">{link('job')}{link('education')}</div>
<div class="hint" style="margin-bottom:16px">
  مسیر انتخابی همه‌جا اعمال می‌شود — منابع، پایپ‌لاین، گزارش‌ها و خروجی‌ها.
  این دو مسیر از هم جدا هستند و قاطی نمی‌شوند.
</div>"""


def render_yield():
    """تب آمار: کدام منبع‌ها واقعاً آگهی می‌دهند؟ (بر اساس SOURCE_YIELD_HISTORY.json)"""
    stats = {}
    history_len = 0
    if HAS_JOB_CRAWLER:
        try:
            history_len = len(_jc.load_yield_history())
            stats = _jc.yield_stats()
        except Exception:
            stats = {}

    # بانک منابع اثبات‌شده (discovered) — total_found تجمعی
    bank = {}
    if HAS_JOB_CRAWLER:
        try:
            bank = _jc.load_discovered()
        except Exception:
            bank = {}

    sources = read_sources() or []
    idx_by_name = {s.get("name", ""): i for i, s in enumerate(sources)}

    if not stats:
        body = ('<p class="empty">هنوز تاریخچه‌ای نیست — یک‌بار «اجرای پایپ‌لاین» را بزن تا '
                'بازده هر منبع ثبت شود.</p>')
    else:
        rows = []
        ranked = sorted(stats.items(), key=lambda kv: -kv[1]["total"])
        for rank, (name, st) in enumerate(ranked, 1):
            bank_entry = bank.get(name, {})
            bank_total = bank_entry.get("total_found", 0)
            idx = idx_by_name.get(name, -1)
            toggle = ""
            if idx >= 0:
                enabled = sources[idx].get("enabled", True)
                toggle = f"""
                <form method="post" action="/settings/sources/toggle" style="display:inline">
                  <input type="hidden" name="index" value="{idx}">
                  <button class="btn" style="padding:4px 10px;font-size:.78rem" type="submit">
                    {"خاموش" if enabled else "روشن"}
                  </button>
                </form>"""
            quality = "🥇" if st["total"] >= 20 else ("✅" if st["total"] > 0 else "🚫")
            rows.append(
                f"<tr>"
                f"<td>{rank}</td>"
                f"<td>{quality} {html.escape(name)}</td>"
                f"<td>{st['runs']}</td>"
                f"<td>{st['with_jobs']} ({st['rate']}%)</td>"
                f"<td>{st['avg']}</td>"
                f"<td><strong>{st['total']}</strong></td>"
                f"<td>{bank_total}</td>"
                f"<td>{st['last']} <span class='hint'>({html.escape(st['last_date'][:16])})</span></td>"
                f"<td>{toggle}</td>"
                f"</tr>"
            )
        body = f"""
        <table class="files"><thead><tr>
          <th>#</th><th>منبع</th><th>اجرا</th><th>آگهی‌دار</th><th>میانگین</th>
          <th>مجموع</th><th>بانک</th><th>آخرین اجرا</th><th></th>
        </tr></thead><tbody>{''.join(rows)}</tbody></table>
        <p class="hint">«مجموع» = همهٔ آگهی‌های یافت‌شده از این منبع در کل تاریخچه ·
        «بانک» = total_found در بانک discovered · «🚫» = هنوز هیچ آگهی‌ای نداده
        (اگر چند اجرا صفر داد، خاموشش کن تا سرعت اجرا بالا برود) ·
        منابع پربازده به‌طور خودکار اولِ جستجوی اجرای بعدی قرار می‌گیرند.</p>"""

    proven = sum(1 for v in bank.values() if v.get("total_found", 0) > 0)
    discovered_cnt = sum(1 for k in bank if k.startswith("[کشف‌شده]"))
    summary = f"""
    <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:10px; margin-bottom:20px">
      <div class="card" style="text-align:center"><div style="font-size:1.5rem">🗂</div>
        <div style="font-size:1.4rem;font-weight:700">{len(sources)}</div>
        <div class="hint">منبع در sources.json</div></div>
      <div class="card" style="text-align:center"><div style="font-size:1.5rem">⭐</div>
        <div style="font-size:1.4rem;font-weight:700">{proven}</div>
        <div class="hint">منبع اثبات‌شدهٔ آگهی‌دار</div></div>
      <div class="card" style="text-align:center"><div style="font-size:1.5rem">🛰</div>
        <div style="font-size:1.4rem;font-weight:700">{discovered_cnt}</div>
        <div class="hint">دامنهٔ کشف‌شدهٔ خودکار</div></div>
      <div class="card" style="text-align:center"><div style="font-size:1.5rem">🕘</div>
        <div style="font-size:1.4rem;font-weight:700">{history_len}</div>
        <div class="hint">اجرای ثبت‌شده</div></div>
    </div>"""

    return f"""<!doctype html>
<html lang="fa"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Hunter — بازده منابع</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}</style>
</head>
<body>
<header class="top">
  <h1>📈 بازده منابع در طول زمان</h1>
  {nav_html('yield', 'job')}
</header>
<main>
  {summary}
  <section>
    <h2>کدام سایت‌ها واقعاً آگهی می‌دهند؟</h2>
    {body}
  </section>
</main>
</body></html>"""


def render_files():
    """تب جدید: همهٔ فایل‌های تولیدشده (output + dashboard) — از صفحهٔ اصلی حذف شد."""
    md_files = list_files(OUTPUT_DIR, exts=[".md", ".txt"])
    xlsx_files = list_files(DASHBOARD_DIR, exts=[".xlsx"])

    def file_rows(files, icon, empty_msg):
        if not files:
            return f'<tr><td colspan="2" class="empty">{empty_msg}</td></tr>'
        rows = []
        for f in files[:60]:
            dt = datetime.fromtimestamp(f["mtime"]).strftime("%Y-%m-%d %H:%M")
            rows.append(
                f'<tr><td class="name">{icon} <a href="/file?path={html.escape(f["rel"])}">{html.escape(f["name"])}</a></td>'
                f'<td class="date">{dt}</td></tr>'
            )
        return "".join(rows)

    return f"""<!doctype html>
<html lang="fa"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Hunter — فایل‌ها</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}</style>
</head>
<body>
<header class="top">
  <h1>📁 فایل‌های تولیدشده</h1>
  {nav_html('files', 'job')}
</header>
<main>
  <section>
    <h2>داشبوردهای اکسل (dashboard/)</h2>
    <table class="files"><tbody>{file_rows(xlsx_files, "📊", "هنوز اکسلی ساخته نشده — یک‌بار پایپ‌لاین را اجرا کن")}</tbody></table>
  </section>
  <section>
    <h2>گزارش‌های متنی (output/)</h2>
    <table class="files"><tbody>{file_rows(md_files, "📄", "هنوز گزارشی تولید نشده")}</tbody></table>
  </section>
  <footer style="text-align:center; padding:30px 0 10px; color:var(--muted); font-size:.8rem; border-top:1px solid var(--line); margin-top:30px">
    MigrationHunter v3.1 · آخرین به‌روزرسانی: {datetime.now().strftime("%Y-%m-%d %H:%M")}
  </footer>
</main>
</body></html>"""


def render_index(track=None, message=None):
    """صفحهٔ اصلی — تماماً گزارش لحظه‌ای.

    این صفحه هیچ فایل، داشبورد یا لیست بلندی ندارد: یک چیز نشان می‌دهد،
    «الان چه اتفاقی در حال رخ دادن است». فایل‌ها در تب جدا (/files) هستند.
    """
    t = track_of(track)
    mt = track_meta(t)
    steps = pipeline_steps(t)

    steps_html = "".join(
        f'<li>{s["emoji"]} {html.escape(s["name"])} '
        f'<span style="color:var(--muted);font-size:.85rem">({s["script"]})</span></li>'
        for s in steps
    )

    msg_html = f'<div class="warn">{html.escape(message)}</div>' if message else ""

    return f"""<!doctype html>
<html lang="fa"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Hunter — {mt['emoji']} {mt['label']}</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}</style>
</head>
<body>
<header class="top">
  <h1>🧭 Migration Hunter</h1>
  {nav_html('dashboard', t)}
</header>
<main>
  {msg_html}
  <section>
    {track_switch_html(t, '/')}
    <div class="live-top">
      <h2 style="border:none;margin:0">🔴 گزارش لحظه‌ای — {mt['emoji']} {mt['label']}</h2>
      <span class="live-nums" id="liveCounter"></span>
    </div>

    <div id="livePanel" style="display:none; margin:10px 0 18px">
      <div class="card" style="border-color:{mt['accent']}; background:#fdf8ef">
        <div style="display:flex; align-items:center; gap:8px; margin-bottom:8px">
          <span style="width:10px; height:10px; border-radius:50%; background:var(--err); display:inline-block; animation:pulse 1.2s infinite"></span>
          <strong style="color:var(--amber)">در حال اجرا…</strong>
          <span style="color:var(--muted); font-size:.85rem" id="liveStep"></span>
        </div>

        <div class="pbar {'edu' if t == 'education' else ''}">
          <span id="progressBar" style="width:0%"></span>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:.85rem; color:var(--muted)">
          <span id="progressLabel">آماده</span>
          <span id="progressPct">0%</span>
        </div>

        <div class="live-now" id="liveNow" style="margin-top:12px">
          <span class="dim">در انتظار شروع…</span>
        </div>
      </div>
    </div>

    <div id="log">آماده به اجرا. دکمهٔ پایین را بزن تا «{mt['label']}» شروع شود.</div>
  </section>

  <section id="liveTableSec" style="display:none">
    <h2>آدرس‌های بررسی‌شده <span class="hint" id="liveTableHint"></span></h2>
    <div class="scroll-y">
      <table class="live"><thead><tr>
        <th>#</th><th>منبع</th><th>کلیدواژه</th><th>آدرس</th>
        <th>{mt['kind']}</th><th>زمان</th><th>وضعیت</th>
      </tr></thead>
      <tbody id="liveRows"></tbody></table>
    </div>
  </section>

  <section>
    <h2>اجرای مسیر {mt['label']}</h2>
    <ol class="steps">{steps_html}</ol>
    <div class="stat-row">
      <div class="st"><b id="stJobs">0</b><span>{mt['kind']} یافت‌شده</span></div>
      <div class="st"><b id="stUrls">0</b><span>آدرس بررسی‌شده</span></div>
      <div class="st"><b id="stSrc">0</b><span>منبع تمام‌شده</span></div>
      <div class="st"><b id="stMs">0</b><span>میانگین زمان پاسخ</span></div>
    </div>
    <p style="margin-top:16px">
      <button class="btn" id="runBtn" onclick="startRun()">▶ اجرای {mt['label']}</button>
      <span class="hint" style="margin-right:10px">مسیر فعال: {mt['emoji']} {mt['label']}</span>
    </p>
  </section>

  <footer style="text-align:center; padding:30px 0 10px; color:var(--muted); font-size:.8rem; border-top:1px solid var(--line); margin-top:30px">
    MigrationHunter · {mt['emoji']} {mt['label']} · آخرین به‌روزرسانی: {datetime.now().strftime("%Y-%m-%d %H:%M")}
  </footer>
</main>

<script>
const TRACK = {json.dumps(t)};
let polling = null;
let seenUrls = 0;

function startRun(){{
  document.getElementById('runBtn').disabled = true;
  document.getElementById('log').textContent = 'شروع شد…';
  document.getElementById('livePanel').style.display = 'block';
  seenUrls = 0;
  fetch('/run?track=' + TRACK, {{method:'POST'}})
    .then(r => r.json()).then(() => {{
      polling = setInterval(pollStatus, 900);
      pollStatus();
    }});
}}

function fmtMs(ms){{ return ms >= 1000 ? (ms/1000).toFixed(1) + 's' : ms + 'ms'; }}

function pollStatus(){{
  fetch('/run-status').then(r => r.json()).then(s => {{
    const logBox = document.getElementById('log');
    logBox.textContent = s.log.join('\\n') || '…';
    logBox.scrollTop = logBox.scrollHeight;

    const lv = s.live || {{}};

    // نوار: اول از شمارش واقعی کراولر، وگرنه از مرحلهٔ پایپ‌لاین
    let pct, label;
    if (lv.active && lv.url_total_all) {{
      pct = Math.min(100, lv.pct || 0);
      label = `آدرس ${{lv.url_done || 0}} از ${{lv.url_total_all}} · منبع ${{lv.source_idx || 0}} از ${{lv.source_total || 0}}`;
    }} else if (!s.running) {{
      pct = 100; label = 'تمام شد';
    }} else {{
      pct = Math.round(((s.step_index || 0) / (s.step_total || 5)) * 100);
      label = `مرحله ${{s.step_index || 0}} از ${{s.step_total || 0}} — ${{s.current_step || '…'}}`;
    }}
    document.getElementById('progressBar').style.width = pct + '%';
    document.getElementById('progressPct').textContent = pct + '%';
    document.getElementById('progressLabel').textContent = label;
    document.getElementById('liveStep').textContent = s.current_step || '';
    document.getElementById('liveCounter').textContent =
      lv.source_total ? `منبع ${{lv.source_idx || 0}} از ${{lv.source_total}}` : '';

    // کادر «الان کجاییم»
    const now = document.getElementById('liveNow');
    if (lv.source) {{
      let h = `<div><span class="dim">منبع:</span> <strong>${{lv.source}}</strong>`;
      if (lv.country) h += ` <span class="dim">(${{lv.country}})</span>`;
      h += `</div>`;
      if (lv.keyword) h += `<div><span class="dim">کلیدواژه:</span> <span class="kw">«${{lv.keyword}}»</span>`;
      if (lv.url) h += `<div><span class="dim">آدرس ${{lv.url_idx || '?'}}/${{lv.url_total || '?'}}:</span> ${{lv.url}}</div>`;
      h += `<div class="dim">مجموع تا الان: ${{lv.jobs_found || 0}} یافته‌شده</div>`;
      now.innerHTML = h;
    }}

    // جدول ریز هر آدرس — فقط ردیف‌های تازه اضافه می‌شوند
    const urls = lv.urls || [];
    if (urls.length) document.getElementById('liveTableSec').style.display = 'block';
    const tbody = document.getElementById('liveRows');
    for (let i = seenUrls; i < urls.length; i++) {{
      const u = urls[i];
      const okc = u.status === 'ok';
      const tr = document.createElement('tr');
      tr.innerHTML =
        `<td class="idx">${{u.idx}}</td>` +
        `<td class="src">${{u.name || ''}}</td>` +
        `<td class="kw">${{u.keyword || ''}}</td>` +
        `<td class="url"><a href="${{u.url}}" target="_blank" rel="noopener">${{u.url}}</a></td>` +
        `<td class="num">${{okc ? u.jobs : '—'}}</td>` +
        `<td class="ms">${{u.ms ? fmtMs(u.ms) : '—'}}</td>` +
        `<td>${{okc
            ? '<span class="badge ok">پاسخ داد</span>'
            : '<span class="badge err">جواب نداد</span>'}}</td>`;
      tbody.appendChild(tr);
    }}
    if (urls.length > seenUrls) {{
      seenUrls = urls.length;
      tbody.parentElement.scrollTop = tbody.parentElement.scrollHeight;
      document.getElementById('liveTableHint').textContent =
        `${{seenUrls}} آدرس · ${{(lv.url_done || 0)}}/${{(lv.url_total_all || 0)}}`;
    }}

    // آمار پایین صفحه
    const doneTimes = urls.filter(u => u.ms).map(u => u.ms);
    const avg = doneTimes.length
      ? Math.round(doneTimes.reduce((a,b) => a+b, 0) / doneTimes.length) : 0;
    document.getElementById('stJobs').textContent = lv.jobs_found || 0;
    document.getElementById('stUrls').textContent = lv.url_done || 0;
    document.getElementById('stSrc').textContent = (lv.sources_done || []).length;
    document.getElementById('stMs').textContent = avg ? fmtMs(avg) : '0';

    if (!s.running) {{
      clearInterval(polling);
      document.getElementById('runBtn').disabled = false;
      document.getElementById('progressBar').style.width = '100%';
      document.getElementById('progressPct').textContent = '100%';
    }}
  }});
}}
pollStatus();
</script>
</body></html>"""

def _edu_load_results(track=None):
    """آخرین نتیجهٔ کراولر تحصیل را می‌خواند."""
    p = track_results_path(track or "education")
    if not os.path.exists(p):
        return {"programs": [], "stats": {}}
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"programs": [], "stats": {}}


def render_education(message=None):
    """تب تحصیل — سه چیز: برنامه‌های پیدا‌شده، بانک اپلیکیشن، و اجرای زنده.

    این صفحه هرچه لازم باشد را به مسیر education وصل می‌کند: برنامه‌های
    کشف‌شده از منابع تحصیلی، و اپلیکیشن‌هایی که خودت ثبت کرده‌ای و تا
    پذیرش پیگیری می‌کنی.
    """
    t = "education"
    mt = track_meta(t)
    applicants = load_applicants()
    names = [a.get("name") or a.get("name_fa") or "?" for a in applicants] or [""]

    msg_html = f'<div class="warn">{html.escape(message)}</div>' if message else ""

    # ── بانک اپلیکیشن‌ها ──
    if HAS_EDU_CRAWLER:
        try:
            apps = _ec.load_applications()
            edu_stats = _ec.app_stats(apps)
            # deadline_warnings دو لیست برمی‌گرداند: (نزدیک، گذشته)
            _soon, _passed = _ec.deadline_warnings(apps)
            warnings = list(_passed) + list(_soon)
        except Exception:
            apps, edu_stats, warnings = [], {}, []
    else:
        apps, edu_stats, warnings = [], {}, []

    stats_html = f"""
    <div class="stat-row">
      <div class="st"><b>{len(apps)}</b><span>اپلیکیشن ثبت‌شده</span></div>
      <div class="st"><b>{edu_stats.get('submitted', 0) + edu_stats.get('accepted', 0) + edu_stats.get('enrolled', 0)}</b><span>ارسال‌شده</span></div>
      <div class="st"><b>{edu_stats.get('accepted', 0) + edu_stats.get('enrolled', 0)}</b><span>پذیرش گرفته</span></div>
      <div class="st"><b>{edu_stats.get('open', 0)}</b><span>در جریان</span></div>
    </div>"""

    warn_html = ""
    if warnings:
        import datetime as _dt
        today = _dt.date.today()
        parts = []
        for w in warnings[:8]:
            dl = (w.get("deadline") or "")[:10]
            try:
                delta = (_dt.date.fromisoformat(dl) - today).days
                when = f"{abs(delta)} روز گذشته" if delta < 0 else f"{delta} روز مانده"
                tone = "var(--err)" if delta < 0 else "var(--amber)"
            except ValueError:
                when, tone = "—", "var(--muted)"
            parts.append(
                f"<li><b>{html.escape(w.get('title','?'))}</b> — مهلت {html.escape(dl or '—')}"
                f" <span style='color:{tone};font-weight:700'>({when})</span>"
                f" <span class='hint'>({html.escape(w.get('applicant',''))})</span></li>")
        warn_html = (f'<div class="warn"><strong>⏳ مهلت‌های اپلیکیشن‌ها</strong>'
                     f'<ul style="margin:8px 0 0">{"".join(parts)}</ul></div>')

    def app_row(a):
        st = a.get("status", "PLANNED")
        emoji = _ec.APP_STATUS_EMOJI.get(st, "•") if HAS_EDU_CRAWLER else "•"
        label = _ec.APP_STATUS_FA.get(st, st) if HAS_EDU_CRAWLER else st
        dl = a.get("deadline") or ""
        title = html.escape(a.get("title") or "?")
        url = a.get("url") or ""
        link = f'<a href="{html.escape(url)}" target="_blank" rel="noopener">{title}</a>' if url else title
        opts = ""
        if HAS_EDU_CRAWLER:
            for k, _e, fa in _ec.APP_STATUSES:
                sel = " selected" if k == st else ""
                opts += f'<option value="{k}"{sel}>{fa}</option>'
        dl_html = f'<span class="hint">{html.escape(dl)}</span>' if dl else '<span class="empty">—</span>'
        return f"""
        <div class="act-row">
          <div class="cat">{emoji}</div>
          <div style="flex:1; min-width:0">
            <div class="ttl">{link}</div>
            <div class="dt">
              {html.escape(a.get('provider',''))}
              {('· ' + html.escape(a.get('country',''))) if a.get('country') else ''}
              {('· ' + html.escape(a.get('degree',''))) if a.get('degree') else ''}
              · مهلت: {dl_html}
              {('· 👤 ' + html.escape(a.get('applicant',''))) if a.get('applicant') else ''}
            </div>
            <form method="post" action="/education/app-status" style="margin-top:6px; display:flex; gap:6px; flex-wrap:wrap; align-items:center">
              <input type="hidden" name="applicant" value="{html.escape(a.get('applicant',''))}">
              <input type="hidden" name="key" value="{html.escape(url or a.get('title',''))}">
              <select name="status" style="padding:4px 8px; border:1px solid var(--line); border-radius:6px; font-family:inherit; font-size:.85rem">
                {opts}
              </select>
              <input type="text" name="note" placeholder="یادداشت…" value=""
                     style="flex:1; min-width:120px; padding:4px 8px; border:1px solid var(--line); border-radius:6px; font-family:inherit; font-size:.85rem">
              <button class="btn" style="padding:5px 12px; font-size:.82rem" type="submit">ثبت</button>
            </form>
            <form method="post" action="/education/app-delete" style="display:inline">
              <input type="hidden" name="applicant" value="{html.escape(a.get('applicant',''))}">
              <input type="hidden" name="key" value="{html.escape(url or a.get('title',''))}">
              <button class="btn" style="padding:2px 8px; font-size:.75rem; background:#a8432f" type="submit">حذف</button>
            </form>
          </div>
        </div>"""

    apps_html = (''.join(app_row(a) for a in apps)
                 if apps else '<p class="empty">هنوز اپلیکیشنی ثبت نکرده‌ای — پایین‌تر فرم افزودن هست.</p>')

    # فرم افزودن اپلیکیشن دستی
    name_opts = "".join(f'<option value="{html.escape(n)}">{html.escape(n)}</option>' for n in names)
    add_form = f"""
    <section>
      <h2>افزودن اپلیکیشن تحصیلی</h2>
      <form method="post" action="/education/app-add" class="inline">
        <label>برنامه
          <input type="text" name="title" placeholder="مثلاً MSc Computer Science" required>
        </label>
        <label>دانشگاه / مؤسسه
          <input type="text" name="provider" placeholder="University of Helsinki">
        </label>
        <label>کشور
          <input type="text" name="country" placeholder="FI">
        </label>
        <label>مهلت اپلای
          <input type="date" name="deadline">
        </label>
        <label>مدرک
          <select name="degree">
            <option value="">—</option><option>Bachelor</option><option>Master</option>
            <option>PhD</option><option>Language</option><option>Certificate</option>
          </select>
        </label>
        <label>زبان
          <select name="language">
            <option value="">—</option><option>English</option><option>Finnish</option>
            <option>Swedish</option><option>Other</option>
          </select>
        </label>
        <label>برای چه کسی
          <select name="applicant">{name_opts}</select>
        </label>
        <label>لینک برنامه
          <input type="text" name="url" placeholder="https://…" dir="ltr">
        </label>
        <label class="full">یادداشت
          <input type="text" name="notes" placeholder="شرایط، زبان، هزینه، نکتهٔ مهم…">
        </label>
        <div class="full">
          <button class="btn" type="submit">➕ افزودن به بانک اپلیکیشن</button>
        </div>
      </form>
    </section>"""

    # ── برنامه‌های کشف‌شده ──
    results = _edu_load_results(t)
    programs = results.get("programs", []) or []
    prog_rows = ""
    for i, p in enumerate(programs[:80], 1):
        name = html.escape(p.get("name") or "?")
        url = p.get("url") or ""
        link = f'<a href="{html.escape(url)}" target="_blank" rel="noopener">{name}</a>' if url else name
        add_btn = ""
        if HAS_EDU_CRAWLER:
            add_btn = f"""
            <form method="post" action="/education/app-add" style="display:inline">
              <input type="hidden" name="title" value="{html.escape(p.get('name',''))}">
              <input type="hidden" name="provider" value="{html.escape(p.get('provider',''))}">
              <input type="hidden" name="country" value="{html.escape(p.get('country',''))}">
              <input type="hidden" name="url" value="{html.escape(url)}">
              <input type="hidden" name="degree" value="{html.escape(p.get('degree',''))}">
              <input type="hidden" name="language" value="{html.escape(p.get('language',''))}">
              <input type="hidden" name="deadline" value="{html.escape(p.get('deadline','') or '')}">
              <button class="btn" style="padding:3px 10px; font-size:.78rem" type="submit">➕ افزودن</button>
            </form>"""
        prog_rows += f"""
        <tr>
          <td>{i}</td>
          <td class="name">{link} {add_btn}</td>
          <td>{html.escape(p.get('provider','') or '—')}</td>
          <td>{html.escape(p.get('degree','') or '—')}</td>
          <td>{html.escape(p.get('language','') or '—')}</td>
          <td>{html.escape(p.get('deadline','') or '—')}</td>
        </tr>"""
    prog_html = (f'<table class="files"><thead><tr><th>#</th><th>برنامه</th>'
                 f'<th>دانشگاه</th><th>مدرک</th><th>زبان</th><th>مهلت</th>'
                 f'</tr></thead><tbody>{prog_rows}</tbody></table>'
                 if prog_rows else
                 '<p class="empty">هنوز برنامه‌ای کشف نشده — از صفحهٔ اصلی مسیر «تحصیل» را اجرا کن.</p>')

    return f"""<!doctype html>
<html lang="fa"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Hunter — 🎓 تحصیل</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}</style>
</head>
<body>
<header class="top">
  <h1>🎓 مسیر تحصیل</h1>
  {nav_html('education', t)}
</header>
<main>
  {msg_html}
  {warn_html}
  <section>
    <h2>بانک اپلیکیشن تحصیلی</h2>
    {stats_html}
    <p class="hint">هر اپلیکیشنی که ثبت کنی از «در نظر گرفته‌شده» تا «پذیرش» و «ثبت‌نام قطعی»
    دنبال می‌شود. هر تغییر وضعیت با تاریخ در تاریخچهٔ خودش می‌ماند.</p>
    {apps_html}
  </section>
  {add_form}
  <section>
    <h2>برنامه‌های کشف‌شده از منابع تحصیلی</h2>
    {prog_html}
    <p style="margin-top:14px">
      <a class="btn" href="{with_track('/', t)}">🔴 گزارش زندهٔ تحصیل</a>
      <a class="btn" href="{with_track('/files', t)}" style="background:var(--amber)">📁 فایل‌های تحصیل</a>
    </p>
  </section>
</main>
</body></html>"""

def render_visa(applicant=None, message=None):
    """تب نورد ویزای همراه — چک‌لیست مراحل + دفترچهٔ اکت‌ها.

    چک‌لیست می‌گوید «چه کارهایی مانده»، دفترچهٔ اکت‌ها ثبت می‌کند «چه
    کرده‌ای و کِی». اولی وضعیت است، دومی تاریخچه.
    """
    t = track_of(applicant and None)  # ویزا مستقل از مسیر کار/تحصیل است
    t = DEFAULT_TRACK
    if not HAS_VISA:
        return f"""<!doctype html><html lang="fa"><head><meta charset="utf-8">
<style>{PAGE_STYLE}</style></head><body><header class="top"><h1>🛂 نورد ویزا</h1>
{nav_html('visa', t)}</header><main>
<p class="empty">ماژول visa_tracker.py پیدا نشد.</p></main></body></html>"""

    people = list(_vt.load_journey().get("people", {}).keys())
    cur = applicant if applicant in people else (people[0] if people else None)

    msg_html = f'<div class="warn">{html.escape(message)}</div>' if message else ""

    # ── هشدار مهلت‌های نزدیک ──
    soon = _vt.deadline_soon(30)
    soon_html = ""
    if soon:
        items = "".join(
            f"<li>{html.escape(x['applicant'])} — <b>{html.escape(x['act'].get('title','?'))}</b>: "
            f"{('گذشته' if x['days'] < 0 else str(x['days']) + ' روز مانده')}</li>"
            for x in soon[:8]
        )
        soon_html = f'<div class="warn"><strong>⏳ کارهای باز با مهلت نزدیک</strong><ul style="margin:8px 0 0">{items}</ul></div>'

    if not cur:
        # ── انتخاب/ساخت نورد ──
        ppl = "".join(
            f'<li><a href="/visa?applicant={html.escape(p)}">{html.escape(p)}</a></li>'
            for p in people) or '<li class="empty">هنوز نوردی ساخته نشده</li>'
        return f"""<!doctype html><html lang="fa"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Hunter — نورد ویزا</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}</style></head><body>
<header class="top"><h1>🛂 نورد ویزای همراه</h1>{nav_html('visa', t)}</header>
<main>
  {msg_html}
  <section>
    <h2>نوردهای موجود</h2>
    <ul>{ppl}</ul>
  </section>
  <section>
    <h2>ساخت نورد جدید</h2>
    <form method="post" action="/visa/new" class="inline">
      <label>برای چه کسی
        <input type="text" name="applicant" placeholder="اسم متقاضی" required>
      </label>
      <label>هدف ویزا
        <input type="text" name="target" placeholder="مثلاً فنلاند — ویزای تحصیلی">
      </label>
      <div class="full">
        <button class="btn" type="submit">🛂 شروع نورد</button>
      </div>
    </form>
    <p class="hint">هر متقاضی نورد جداگانهٔ خودش را دارد — چک‌لیست و اکت‌ها قاطی نمی‌شوند.</p>
  </section>
</main></body></html>"""

    st = _vt.journey_stats(cur)
    journey = _vt.load_journey(cur)
    stages = journey.get("stages", {}) or {}
    acts = journey.get("acts", []) or []

    # ── چک‌لیست ──
    def stage_row(key, title, hint, cat):
        e = stages.get(key) or {}
        done = bool(e.get("done"))
        mark = "☑" if done else ""
        when = f'<span class="when">{html.escape(e.get("date",""))}</span>' if e.get("date") else ""
        note = e.get("note") or ""
        note_html = ""
        if note:
            note_html = (f'<form method="post" action="/visa/stage-note" style="display:inline">'
                         f'<input type="hidden" name="applicant" value="{html.escape(cur)}">'
                         f'<input type="hidden" name="key" value="{key}">'
                         f'<input type="text" name="note" value="{html.escape(note)}" '
                         f'placeholder="یادداشت…" style="width:130px;padding:2px 6px;font-size:.78rem;'
                         f'border:1px solid var(--line);border-radius:5px;font-family:inherit">'
                         f'<button class="btn" style="padding:1px 8px;font-size:.75rem" type="submit">💾</button>'
                         f'</form>')
        return f"""
        <li class="{'done' if done else ''}">
          <a class="box" href="/visa/toggle?applicant={html.escape(cur)}&key={key}" title="تیک بزن">{mark}</a>
          <div style="flex:1">
            <div><span class="t">{html.escape(title)}</span> {when} {note_html}</div>
            <div class="h">{html.escape(hint)}</div>
          </div>
        </li>"""

    checklist = ""
    for cat in _vt.CAT_ORDER:
        keys = [(k, t_, h_) for k, t_, h_, c_ in _vt.VISA_STAGES if c_ == cat]
        if not keys:
            continue
        c = st["per_cat"][cat]
        bar = f"{c['done']}/{c['total']}"
        checklist += f"""<h3 style="font-size:.95rem;margin:18px 0 6px;color:var(--teal-deep)">
          {_vt.CAT_FA[cat]} <span class="hint">({bar})</span></h3>
          <ul class="check">{''.join(stage_row(*k, cat) for k in keys)}</ul>"""

    # ── دفترچهٔ اکت‌ها ──
    def act_row(a):
        emoji = _vt.ACT_CAT_EMOJI.get(a.get("category", "other"), "📌")
        date = a.get("date", "")
        try:
            delta = (datetime.now().date() - datetime.strptime(date[:10], "%Y-%m-%d").date()).days
        except Exception:
            delta = None
        late = ""
        if delta is not None and delta > 0 and not a.get("done"):
            late = f' <span class="late">({delta} روز گذشته)</span>'
        detail = f'<div class="dt">{html.escape(a.get("detail",""))}</div>' if a.get("detail") else ""
        return f"""
        <div class="act-row {'done' if a.get('done') else ''}">
          <div class="cat">{emoji}</div>
          <div style="flex:1; min-width:0">
            <div class="ttl">{html.escape(a.get('title',''))}</div>
            <div class="dt">{html.escape(date)}{late}</div>
            {detail}
          </div>
          <div style="white-space:nowrap">
            <a class="box" style="display:inline-flex; width:21px;height:21px;border:2px solid var(--line);
               border-radius:5px; align-items:center; justify-content:center; text-decoration:none;
               color:{'#fff' if a.get('done') else 'transparent'};
               background:{'var(--ok)' if a.get('done') else '#fff'}"
               href="/visa/act-toggle?applicant={html.escape(cur)}&id={a.get('id')}"
               title="انجام شد">{'☑' if a.get('done') else ''}</a>
            <form method="post" action="/visa/act-delete" style="display:inline">
              <input type="hidden" name="applicant" value="{html.escape(cur)}">
              <input type="hidden" name="id" value="{a.get('id')}">
              <button class="btn" style="padding:1px 8px; font-size:.75rem; background:#a8432f" type="submit">✕</button>
            </form>
          </div>
        </div>"""

    acts_html = (''.join(act_row(a) for a in acts)
                 if acts else '<p class="empty">هنوز کاری ثبت نکرده‌ای — پایین‌تر فرم ثبت هست.</p>')

    act_cats = "".join(f'<option value="{k}">{e} {html.escape(fa)}</option>'
                       for k, e, fa in _vt.ACT_CATEGORIES)

    people_tabs = " · ".join(
        f'<a href="/visa?applicant={html.escape(p)}" style="{"" if p == cur else "opacity:.6"}">'
        f'{"● " if p == cur else ""}{html.escape(p)}</a>'
        for p in people)

    stale = _vt.overdue_stages(cur, 45)
    stale_html = ""
    if stale:
        items = "".join(f"<li>{html.escape(s['title'])} <span class='hint'>({s['stale_days']} روز بی‌تیک)</span></li>"
                        for s in stale[:6])
        stale_html = f'<div class="warn"><strong>🕰️ مراحلی که مدت‌هاست بی‌تیک مانده‌اند</strong><ul style="margin:8px 0 0">{items}</ul></div>'

    return f"""<!doctype html>
<html lang="fa"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Hunter — 🛂 نورد ویزا — {html.escape(cur)}</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}</style>
</head>
<body>
<header class="top">
  <h1>🛂 نورد ویزای همراه — {html.escape(cur)}</h1>
  {nav_html('visa', t)}
</header>
<main>
  {msg_html}
  {soon_html}
  {stale_html}

  <section>
    <div class="live-top">
      <h2 style="border:none;margin:0">وضعیت نورد</h2>
      <span class="hint">{people_tabs}</span>
    </div>
    <p class="hint">هدف: <b>{html.escape(st['target'] or '—')}</b></p>
    <div class="stat-row">
      <div class="st"><b>{st['pct']}%</b><span>پیشرفت چک‌لیست</span></div>
      <div class="st"><b>{st['stages_done']}/{st['stages_total']}</b><span>مرحلهٔ انجام‌شده</span></div>
      <div class="st"><b>{st['acts_total']}</b><span>کار ثبت‌شده</span></div>
      <div class="st"><b>{st['acts_open']}</b><span>کار باز</span></div>
    </div>
    <div class="pbar"><span style="width:{st['pct']}%"></span></div>
    <form method="post" action="/visa/target" style="margin-bottom:8px">
      <input type="hidden" name="applicant" value="{html.escape(cur)}">
      <input type="text" name="target" value="{html.escape(st['target'])}"
             placeholder="هدف این نورد — مثلاً فنلاند، ویزای تحصیلی"
             style="padding:6px 10px;border:1px solid var(--line);border-radius:6px;font-family:inherit;width:min(420px,80%)">
      <button class="btn" style="padding:6px 14px; font-size:.85rem" type="submit">ثبت هدف</button>
    </form>
  </section>

  <section>
    <h2>چک‌لیست مراحل</h2>
    <p class="hint">روی هر مربع کلیک کن تا تیک بخورد و تاریخش ثبت شود.</p>
    {checklist}
  </section>

  <section>
    <h2>دفترچهٔ اقدامات</h2>
    <p class="hint">هر کاری که انجام دادی را با تاریخ ثبت کن — بعداً همین تاریخچه ارزش دارد.</p>
    {acts_html}
    <form method="post" action="/visa/act-add" class="inline" style="margin-top:14px">
      <input type="hidden" name="applicant" value="{html.escape(cur)}">
      <label>دسته
        <select name="category">{act_cats}</select>
      </label>
      <label>تاریخ
        <input type="date" name="date" value="{datetime.now().strftime('%Y-%m-%d')}">
      </label>
      <label class="full">چه کاری کردی؟
        <input type="text" name="title" placeholder="مثلاً مدارک ترجمهٔ رسمی را فرستادم" required>
      </label>
      <label class="full">جزئیات
        <input type="text" name="detail" placeholder="شمارهٔ نامه، نتیجه، لینک، هرچه لازم است">
      </label>
      <div class="full"><button class="btn" type="submit">📌 ثبت اقدام</button></div>
    </form>
  </section>
</main>
</body></html>"""

def render_xlsx_as_html(full_path, rel_path):
    if not HAS_OPENPYXL:
        return f"""<!doctype html><html lang="fa"><head><meta charset="utf-8">
<style>{PAGE_STYLE}</style></head><body><main>
<p class="empty">openpyxl نصب نیست، فقط می‌توانی دانلود کنی: <a href="/download?path={html.escape(rel_path)}">دانلود {html.escape(os.path.basename(rel_path))}</a></p>
<p><a href="/">← بازگشت</a></p></main></body></html>"""

    wb = load_workbook(full_path, data_only=True, read_only=True)
    sections = []
    for ws in wb.worksheets:
        rows_html = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i > 300:
                rows_html.append("<tr><td colspan='99' class='empty'>… (بریده‌شده، فایل کامل را دانلود کن)</td></tr>")
                break
            cells = "".join(f"<td>{html.escape('' if c is None else str(c))}</td>" for c in row)
            tag = "th" if i == 0 else "td"
            if i == 0:
                cells = "".join(f"<th>{html.escape('' if c is None else str(c))}</th>" for c in row)
            rows_html.append(f"<tr>{cells}</tr>")
        sections.append(f"<h3>{html.escape(ws.title)}</h3><div style='overflow-x:auto'><table class='files'>{''.join(rows_html)}</table></div>")

    return f"""<!doctype html><html lang="fa"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>{PAGE_STYLE}</style></head><body>
<header class="top"><h1>📊 {html.escape(os.path.basename(rel_path))}</h1>
<a class="sub" href="/">← بازگشت به داشبورد</a></header>
<main>{''.join(sections)}
<p style="margin-top:20px"><a href="/download?path={html.escape(rel_path)}">⬇ دانلود فایل اصلی</a></p>
</main></body></html>"""


def render_text_file(full_path, rel_path):
    with open(full_path, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()
    return f"""<!doctype html><html lang="fa"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>{PAGE_STYLE}
pre.filebody{{background:var(--panel); border:1px solid var(--line); border-radius:var(--radius); padding:18px; white-space:pre-wrap; font-family:inherit}}
</style></head><body>
<header class="top"><h1>📄 {html.escape(os.path.basename(rel_path))}</h1>
<a class="sub" href="/">← بازگشت به داشبورد</a></header>
<main><pre class="filebody">{html.escape(content)}</pre></main></body></html>"""


def render_settings(message=None):
    applicants = load_applicants()
    env = read_env()

    warnings = []
    if not gitignore_covers(".env"):
        warnings.append("فایل <code>.env</code> در .gitignore نیست — اگه همین الان commit/push کنی، رمزها لو می‌رن.")
    if not gitignore_covers("config.json"):
        warnings.append("فایل <code>config.json</code> در .gitignore نیست — ایمیل و اطلاعات شخصی توش هست.")
    warn_html = ""
    if warnings:
        warn_html = '<div class="warn"><b>⚠️ قبل از هر commit این‌ها را برطرف کن:</b><ul>' + \
                    "".join(f"<li>{w}</li>" for w in warnings) + "</ul></div>"

    msg_html = f'<div class="card" style="border-color:var(--ok);margin-bottom:20px">{html.escape(message)}</div>' if message else ""

    sources = read_sources()
    if sources is None:
        sources_html = ('<p class="empty">هنوز sources.json ساخته نشده — یک‌بار از تب داشبورد «اجرای پایپ‌لاین» '
                         'را بزن (یا مستقیم <code>job_crawler.py</code> را اجرا کن) تا با لیست پیش‌فرض ساخته شود.</p>')
    else:
        by_country = {}
        for i, s in enumerate(sources):
            by_country.setdefault(s.get("country", "؟"), []).append((i, s))
        rows = []
        for country in sorted(by_country):
            rows.append(f'<tr><td colspan="5" style="background:var(--amber-soft);font-weight:700">{html.escape(country)}</td></tr>')
            for i, s in by_country[country]:
                enabled = s.get("enabled", True)
                badge = '<span class="badge ok">فعال</span>' if enabled else '<span class="badge err">غیرفعال</span>'
                rows.append(f"""
                <tr>
                  <td>{html.escape(s.get('name',''))}</td>
                  <td>{html.escape(s.get('type','job_board'))}</td>
                  <td>{s.get('trust','—')}</td>
                  <td>{badge}</td>
                  <td>
                    <form method="post" action="/settings/sources/toggle" style="display:inline">
                      <input type="hidden" name="index" value="{i}">
                      <button class="btn" style="padding:5px 12px;font-size:.82rem" type="submit">
                        {"غیرفعال کن" if enabled else "فعال کن"}
                      </button>
                    </form>
                  </td>
                </tr>""")
        sources_html = f"""
        <table class="files"><thead><tr>
          <th>نام منبع</th><th>نوع</th><th>اعتبار</th><th>وضعیت</th><th></th>
        </tr></thead><tbody>{''.join(rows)}</tbody></table>
        <p class="hint">این‌ها همان منابعی هستند که job_crawler.py موقع اجرای پایپ‌لاین جستجو می‌کند.
        غیرفعال‌کردن یک منبع یعنی در اجرای بعدی رد می‌شود، بدون نیاز به دست‌زدن به کد.</p>
        """

    def applicant_block(a):
        aid = a.get("id", "")
        aid_upper = aid.upper()
        has_pw = bool(env.get(f"EMAIL_PASSWORD_{aid_upper}"))
        existing_email = env.get(f"EMAIL_{aid_upper}", "")
        existing_provider = env.get(f"EMAIL_PROVIDER_{aid_upper}", "gmail")
        pw_status = '<span class="badge ok">رمز ثبت شده ✓</span>' if has_pw else '<span class="badge err">رمز ثبت نشده</span>'
        return f"""
        <div class="applicant-block">
          <h3>{html.escape(a.get('emoji','👤'))} {html.escape(a.get('name_fa') or a.get('name',aid))}</h3>
          <p class="hint">شناسه (id): <code>{html.escape(aid)}</code></p>

          <form class="inline" method="post" action="/settings/applicant">
            <input type="hidden" name="id" value="{html.escape(aid)}">
            <label>نام (انگلیسی)<input name="name" value="{html.escape(a.get('name',''))}"></label>
            <label>نام (فارسی)<input name="name_fa" value="{html.escape(a.get('name_fa',''))}"></label>
            <label>ایموجی<input name="emoji" value="{html.escape(a.get('emoji',''))}"></label>
            <label>شغل / تخصص<input name="profession" value="{html.escape(a.get('profession',''))}"></label>
            <label class="full">لینکدین<input name="linkedin" value="{html.escape(a.get('linkedin',''))}"></label>
            <label class="full">کلمات کلیدی (با کاما جدا کن)<input name="keywords" value="{html.escape(', '.join(a.get('keywords', [])))}"></label>
            <label>سطح انگلیسی<input name="english" value="{html.escape(a.get('english',''))}"></label>
            <label>سطح آلمانی<input name="german" value="{html.escape(a.get('german',''))}"></label>
            <div class="full"><button class="btn" type="submit">💾 ذخیره‌ی مشخصات</button></div>
          </form>

          <hr style="border:none;border-top:1px solid var(--line); margin:18px 0">
          <p><b>ایمیل برای دریافت خودکار پاسخ‌ها</b> {pw_status}</p>
          <form class="inline" method="post" action="/settings/credentials">
            <input type="hidden" name="id" value="{html.escape(aid)}">
            <label>آدرس ایمیل<input name="email" value="{html.escape(existing_email)}" placeholder="you@gmail.com"></label>
            <label>سرویس<select name="provider">
                <option value="gmail" {"selected" if existing_provider=="gmail" else ""}>Gmail</option>
                <option value="outlook" {"selected" if existing_provider=="outlook" else ""}>Outlook</option>
                <option value="other" {"selected" if existing_provider not in ("gmail","outlook") else ""}>سایر (IMAP دستی)</option>
            </select></label>
            <label class="full">App Password
              <input type="password" name="password" placeholder="{'برای عوض‌کردن رمز پر کن، وگرنه خالی بذار' if has_pw else 'مثلاً: abcd efgh ijkl mnop'}">
            </label>
            <p class="hint full">
              رمز عادی جیمیل کار نمی‌کند — باید از
              <a href="https://myaccount.google.com/apppasswords" target="_blank" rel="noopener">myaccount.google.com/apppasswords</a>
              یک App Password جدید بسازی. این رمز فقط در فایل محلی <code>.env</code> ذخیره می‌شود
              (که در .gitignore هست) و هیچ‌جا فرستاده نمی‌شود.
            </p>
            <div class="full"><button class="btn" type="submit">💾 ذخیره‌ی ایمیل</button></div>
          </form>
        </div>"""

    blocks = "".join(applicant_block(a) for a in applicants)
    if not blocks:
        blocks = '<p class="empty">هنوز متقاضی‌ای نداری — اول یکی با فرم پایین اضافه کن.</p>'

    return f"""<!doctype html>
<html lang="fa"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Hunter — تنظیمات</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}</style>
</head>
<body>
<header class="top">
  <h1>⚙️ تنظیمات</h1>
  {nav_html('settings', 'job')}
</header>
<main>
  {warn_html}
  {msg_html}

  <section>
    <h2>متقاضی‌ها و ایمیل‌هایشان</h2>
    {blocks}
  </section>

  <section>
    <h2>افزودن متقاضی جدید</h2>
    <form class="inline" method="post" action="/settings/applicant">
      <label>شناسه (id, فقط حروف انگلیسی)<input name="id" required placeholder="mehran"></label>
      <label>نام (انگلیسی)<input name="name" placeholder="Mehran"></label>
      <label>نام (فارسی)<input name="name_fa" placeholder="مهران"></label>
      <label>ایموجی<input name="emoji" placeholder="👨"></label>
      <label class="full">شغل / تخصص<input name="profession"></label>
      <label class="full">لینکدین<input name="linkedin" placeholder="linkedin.com/in/..."></label>
      <label class="full">کلمات کلیدی (با کاما)<input name="keywords" placeholder="developer, backend, python"></label>
      <label>سطح انگلیسی<input name="english" placeholder="B2"></label>
      <label>سطح آلمانی<input name="german" placeholder="A1"></label>
      <div class="full"><button class="btn" type="submit">➕ افزودن متقاضی</button></div>
    </form>
  </section>

  <section>
    <h2>🌐 منابع جستجوی آگهی</h2>
    {sources_html}
    <h3 style="font-size:.95rem;margin-top:22px">افزودن منبع سفارشی جدید</h3>
    <form class="inline" method="post" action="/settings/sources/add">
      <label>نام منبع<input name="name" required placeholder="مثلاً: Glassdoor DE"></label>
      <label>کشور (کد دو حرفی)<input name="country" required placeholder="DE"></label>
      <label class="full">آدرس جستجو (URL)<input name="url" required placeholder="https://..."></label>
      <label>امتیاز اعتبار (۰ تا ۱۰۰)<input name="trust" type="number" min="0" max="100" value="60"></label>
      <div class="full"><button class="btn" type="submit">➕ افزودن منبع</button></div>
    </form>
  </section>
</main>
</body></html>"""


def render_about():
    return f"""<!doctype html>
<html lang="fa"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Hunter — توضیحات</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}
.about p{{margin:0 0 12px}}
.about ul{{margin:0 0 14px; padding-right:22px}}
.about li{{margin-bottom:6px}}
.callout{{background:var(--amber-soft); border-radius:var(--radius); padding:14px 16px; margin:14px 0}}
</style>
</head>
<body>
<header class="top">
  <h1>ℹ️ این صفحه چیکار می‌کند</h1>
  {nav_html('about', 'job')}
</header>
<main class="about">

  <section>
    <h2>این داشبورد چیست</h2>
    <p>این یک ابزار محلی است که روی سیستم خودت اجرا می‌شود (نه سرور بیرونی، نه ابر).
    چهار تب دارد:</p>
    <ul>
      <li><b>داشبورد</b> — نشون می‌ده چند متقاضی داری، دکمه‌ی اجرای پایپ‌لاین (۵ مرحله‌ی <code>run.py</code>)،
      و لیست فایل‌های تولیدشده در <code>output/</code> و <code>dashboard/</code>.</li>
      <li><b>گزارش‌ها</b> — برای هر متقاضی، آگهی‌های مرتبط، خروجی اکسل مخصوص همون متقاضی،
      و تولید کاور لتر/ایمیل با AI (اگه کلید API تنظیم شده باشه).</li>
      <li><b>تنظیمات</b> — مشخصات متقاضی‌ها و ایمیل/رمزشون (ذخیره‌ی محلی در <code>config.json</code> و <code>.env</code>).</li>
      <li><b>همین صفحه</b> — توضیح صادقانه‌ی اینکه crawler دقیقاً چیکار می‌کنه و محدودیت‌هاش چیه.</li>
    </ul>
  </section>

  <section>
    <h2>«اصالت‌سنجی آگهی» دقیقاً چیه؟ (سؤال مهم — جواب صادقانه)</h2>
    <p>توی ستون «حقیقی/فیک» اکسل، این یک <b>سیستم تشخیص کلاه‌برداری واقعی نیست</b>.
    کاری که کد (<code>detect_job_realness</code> در <code>job_crawler.py</code>) انجام می‌ده این است:</p>
    <ul>
      <li>عنوان و شرکتِ آگهی رو با کلمات کلیدی‌ای که خودت برای هر متقاضی توی تنظیمات نوشتی مقایسه می‌کنه.</li>
      <li>اگه تطبیق کافی پیدا بشه (یا حتی فقط یک کلمه‌ی عمومی مثل «manager»/«engineer» توی عنوان باشه)
      برچسب «✅ واقعی» می‌خوره.</li>
      <li>هیچ چک واقعی روی <b>هویت شرکت</b>، ثبت رسمی، الگوهای شناخته‌شده‌ی کلاه‌برداری استخدامی،
      یا اعتبار آگهی انجام نمی‌شود.</li>
    </ul>
    <div class="callout">
      یعنی این ستون در واقع «چقدر به پروفایل تو نزدیکه» را می‌سنجه، نه «آیا این آگهی واقعیه یا کلاه‌برداریه».
      اسمش توی کد گمراه‌کننده‌ست. پیشنهاد می‌کنم قبل از هرگونه تصمیم (مخصوصاً پرداخت پول یا دادن مدارک)،
      خودت دستی شرکت رو در LinkedIn/سایت رسمی‌اش چک کنی.
    </div>
  </section>

  <section>
    <h2>ریکروتر/شرکت را از کجا تشخیص می‌دهد؟</h2>
    <p>فعلاً از <b>هیچ‌جا</b> — نکته‌ی مهم دیگه: توی تابع استخراج
    (<code>extract_jobs_from_html</code>)، فیلد «شرکت» همیشه خالی ساخته می‌شود
    (<code>company = ""</code> در کد) و در نتیجه در اکسل همیشه «نامشخص» نشون داده می‌شود.
    یعنی الان کد اصلاً نام شرکت/ریکروتر رو از صفحه استخراج نمی‌کنه — این یکی از چیزهایی‌ست
    که پیشنهاد می‌کنم اول از همه درستش کنیم (پایین صفحه، بخش «چی اضافه/درست کنیم»).</p>
  </section>

  <section>
    <h2>از کجاها و بر چه اساسی آگهی جمع می‌کند؟</h2>
    <p>لیست منابع در کد (<code>SEARCH_SOURCES</code>) از قبل ثابت تعریف شده:</p>
    <ul>
      <li>Seek NZ, Trade Me Jobs NZ (نیوزیلند)</li>
      <li>Seek AU (استرالیا)</li>
      <li>Job Bank Canada (رسمی دولت کانادا), Indeed Canada</li>
      <li>StepStone DE, Indeed DE (آلمان)</li>
      <li>IrishJobs (ایرلند)</li>
      <li>Indeed NL (هلند)</li>
    </ul>
    <p>روش کار: صفحه‌ی HTML خام هر URL رو با <code>urllib</code> می‌گیره (بدون اجرای جاوااسکریپت)،
    بعد یک پارسر عمومی همه‌ی تگ‌های <code>&lt;a&gt;</code> و تیترها (<code>h1..h6</code>) رو
    به‌عنوان «آگهی احتمالی» برمی‌داره. هیچ selector اختصاصی برای ساختار HTML هر سایت نداره.</p>
    <div class="callout">
      محدودیت مهم: خیلی از سایت‌های کاریابی امروزی (به‌خصوص Seek و Indeed) محتوایشان را با
      جاوااسکریپت می‌سازند. چون این کد جاوااسکریپت اجرا نمی‌کند، ممکن است از این سایت‌ها
      نتیجه‌ی خیلی کم یا حتی خالی بگیرد و شما اصلاً متوجه نشوید (لاگ فقط تعداد را نشان می‌دهد،
      نه اینکه واقعاً همه‌ی آگهی‌های آن سایت را دیده یا نه).
    </div>
  </section>

  <section>
    <h2>چقدر به‌روز است؟</h2>
    <p>فقط وقتی که خودت دکمه‌ی «اجرای پایپ‌لاین» را بزنی (یا <code>job_crawler.py</code> را دستی اجرا کنی).
    هیچ زمان‌بندی خودکار (cron/scheduler) در پروژه وجود ندارد. تاریخی که در اکسل کنار هر آگهی می‌بینی
    («تاریخ افزوده»)، تاریخ واقعی انتشار آگهی نیست — تاریخ همان لحظه‌ای است که تو crawler را اجرا کردی.</p>
  </section>

  <section>
    <h2>نظر من — چی اضافه/درست کنیم؟</h2>
    <p>به ترتیب اولویت پیشنهاد می‌کنم:</p>
    <ul>
      <li><b>استخراج واقعی نام شرکت</b> — الان همیشه خالیه؛ این پایه‌ای‌ترین باگه.</li>
      <li><b>تغییر اسم ستون</b> از «حقیقی/فیک» به چیزی مثل «میزان تطبیق با پروفایل» تا گمراه‌کننده نباشه.</li>
      <li><b>بررسی محتوای واقعی صفحات</b> با یک نمونه‌ی دستی از هر سایت، چون خیلی محتمله بعضی سایت‌ها
      (مخصوصاً Seek/Indeed که JS-heavy هستند) اصلاً چیزی برنگردونن.</li>
      <li><b>لینک مستقیم به صفحه‌ی جستجوی رسمی هر سایت</b> به‌جای تکیه بر پارس HTML، برای مواردی
      که پارس اتوماتیک شکست می‌خوره — حداقل کاربر بتونه دستی چک کنه.</li>
      <li><b>هشدار صریح در خود اکسل</b> (نه فقط اینجا) که «واقعی/فیک» یعنی تطبیق کلیدواژه، نه اصالت‌سنجی واقعی.</li>
    </ul>
    <p class="hint">اگه بخوای، می‌تونم همین حالا شروع کنم به درست‌کردن استخراج نام شرکت و
    عوض‌کردن اسم ستون — فقط بگو کدوم اول.</p>
  </section>

</main>
</body></html>"""


MEM_DIR = os.path.join(BASE, "memory")


def score_job_for_applicant(job, applicant):
    """امتیاز تطبیقِ ساده — همان منطق job_crawler.py: شمارش کلیدواژه در عنوان/شرکت."""
    title = (job.get("title") or "").lower()
    company = (job.get("company") or "").lower()
    score = 0
    for kw in applicant.get("keywords", []):
        kw = (kw or "").lower().strip()
        if not kw:
            continue
        if kw in title:
            score += 2
        if kw in company:
            score += 1
    return score


def build_applicant_excel(applicant, jobs):
    wb = Workbook()
    ws = wb.active
    ws.title = "گزارش آگهی‌ها"
    ws.sheet_view.rightToLeft = True
    from openpyxl.styles import Font, PatternFill, Alignment
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid")
    headers = ["#", "عنوان شغل", "شرکت", "کشور", "منبع", "لینک آگهی", "امتیاز تطبیق",
               "وضعیت اقدام", "دلیل تطبیق"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    st_fa = {k: v["fa"] for k, v in unified_report.APP_STATUS.items()}
    for i, j in enumerate(jobs, 1):
        ws.append([i, j.get("title", ""), j.get("company", "نامشخص"),
                   j.get("country", ""), j.get("source", ""), j.get("url", ""),
                   j.get("score", j.get("_score", 0)),
                   st_fa.get(j.get("app_status", ""), ""),
                   j.get("_why", "")])
    widths = {"A": 5, "B": 45, "C": 25, "D": 10, "E": 20, "F": 45, "G": 12,
              "H": 15, "I": 30}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.auto_filter.ref = f"A1:I{len(jobs)+1}"
    return wb


def call_ai(prompt, max_tokens=800):
    """
    تماس مستقیم با API هوش‌مصنوعی (فقط urllib، بدون وابستگی جدید).
    از AI_PROVIDER / AI_API_KEY در .env استفاده می‌کند (طبق .env.example).
    """
    import urllib.request as _ur
    import urllib.error as _ue

    env = read_env()
    provider = (env.get("AI_PROVIDER") or "openai").strip().lower()
    api_key = (env.get("AI_API_KEY") or "").strip()
    if not api_key:
        return None, "کلید AI_API_KEY در .env تنظیم نشده — از تب تنظیمات یا مستقیم در .env (طبق .env.example) یک کلید OpenAI یا Gemini اضافه کن."

    try:
        if provider == "gemini":
            url = (f"https://generativelanguage.googleapis.com/v1beta/models/"
                   f"gemini-1.5-flash:generateContent?key={api_key}")
            body = json.dumps({"contents": [{"parts": [{"text": prompt}]}]}).encode("utf-8")
            req = _ur.Request(url, data=body, headers={"Content-Type": "application/json"}, method="POST")
            with _ur.urlopen(req, timeout=45) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            return data["candidates"][0]["content"]["parts"][0]["text"], None
        else:
            url = "https://api.openai.com/v1/chat/completions"
            body = json.dumps({
                "model": "gpt-4o-mini",
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
            }).encode("utf-8")
            req = _ur.Request(url, data=body, headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}",
            }, method="POST")
            with _ur.urlopen(req, timeout=45) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            return data["choices"][0]["message"]["content"], None
    except _ue.HTTPError as e:
        return None, f"خطای API ({e.code}): {e.read().decode('utf-8', 'ignore')[:300]}"
    except Exception as e:
        return None, f"خطا در تماس با AI: {e}"


def build_cover_letter_prompt(applicant, job):
    kws = ", ".join(applicant.get("keywords", [])) or "—"
    return f"""You are a professional career-application assistant. Write for this candidate:
- Name: {applicant.get('name') or applicant.get('name_fa','')}
- Profession: {applicant.get('profession','')}
- Key skills/keywords: {kws}
- English level: {applicant.get('english','') or 'not specified'}
- German level: {applicant.get('german','') or 'not specified'}

Job they are applying to:
- Title: {job.get('title','')}
- Company: {job.get('company') or 'Unknown (not extracted from listing)'}
- Country: {job.get('country','')}
- Source: {job.get('source','')}

Write two clearly separated sections in English:
1. "## EMAIL" — a short, professional email (5-8 sentences) to accompany the application.
2. "## COVER LETTER" — a one-page formal cover letter.

Keep the tone professional and honest. Do NOT invent specific work-history facts, employers, or
achievements beyond the profession/skills given above — keep claims general and truthful."""


APPLICANT_COUNTRIES = ["FI", "SE", "NO", "DK", "DE", "NL", "CA", "AU", "GB", "IE", "ALL"]

APP_BANK_PATH = os.path.join(MEM_DIR, "APPLICATION_BANK.json")


def _load_app_bank():
    if not os.path.exists(APP_BANK_PATH):
        return {"applications": []}
    try:
        with open(APP_BANK_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"applications": []}


def _save_app_bank(bank):
    os.makedirs(MEM_DIR, exist_ok=True)
    with open(APP_BANK_PATH, "w", encoding="utf-8") as f:
        json.dump(bank, f, ensure_ascii=False, indent=2)


def _upsert_application(applicant, url, title, employer, country, status):
    """
    ثبت یا به‌روزرسانی یک اقدام در APPLICATION_BANK.

    اگر قبلاً برای همین (متقاضی، آدرس آگهی) چیزی ثبت شده باشد، فقط وضعیت
    عوض می‌شود و رکورد جدیدی ساخته نمی‌شود — وگرنه بانک پر از تکراری
    می‌شد. کلید تطبیق URL است چون شناسهٔ آگهی پایدار نیست.
    """
    bank = _load_app_bank()
    apps = bank.setdefault("applications", [])
    now = datetime.now()

    def same(rec):
        return unified_report._norm(rec.get("job_url") or rec.get("url")) == \
               unified_report._norm(url)

    for rec in apps:
        if unified_report._matches_applicant(rec, {"id": applicant,
                                                    "name": applicant,
                                                    "name_fa": applicant}) and same(rec):
            old = rec.get("status", "?")
            rec["status"] = status
            rec["job"] = title or rec.get("job", "")
            rec["employer"] = employer or rec.get("employer", "")
            rec["country"] = country or rec.get("country", "")
            rec["updated_at"] = now.strftime("%Y-%m-%d %H:%M")
            rec.setdefault("history", []).append(
                {"at": rec["updated_at"], "from": old, "to": status})
            if status not in ("PLANNED", "PREPARING") and not rec.get("sent_date"):
                rec["sent_date"] = rec["updated_at"][:10]
            _save_app_bank(bank)
            return True, f"وضعیت از {old} به {status} تغییر کرد"

    apps.append({
        "id": f"JOB-{now.strftime('%Y%m%d%H%M%S')}-{len(apps)+1}",
        "applicant": applicant,
        "employer": employer,
        "job": title,
        "country": country,
        "job_url": url,
        "url": url,
        "status": status,
        "sent_date": "",
        "created_at": now.strftime("%Y-%m-%d %H:%M"),
        "reply_deadline": "",
        "history": [{"at": now.strftime("%Y-%m-%d %H:%M"), "from": "", "to": status}],
    })
    _save_app_bank(bank)
    return True, ""


def _job_card(job, aid, country):
    """
    یک ردیف آگهی: عنوان، شرکت، کشور، لینک مستقیم و دکمه‌های اقدام.

    لینک آگهی واقعاً کلیک‌پذیر است (target=_blank) — کاربر لازم نیست بین
    فایل‌ها بگردد تا آدرسش را پیدا کند.
    """
    title = html.escape(job.get("title") or "بدون عنوان")
    company = html.escape(job.get("company") or "نامشخص")
    url = job.get("url") or ""
    ctry = job.get("country") or ""
    ctry_fa = unified_report.COUNTRY_FA.get(ctry, ctry)
    score = job.get("score", 0)
    why = job.get("_why") or ""
    applied = job.get("applied")

    # ── ستون وضعیت اقدام ──
    if applied:
        meta = unified_report.APP_STATUS.get(job.get("app_status") or "", {})
        badge = (f'<span class="badge" style="background:var(--ok)">'
                 f'{meta.get("emoji","📋")} {html.escape(meta.get("fa","اقدام‌شده"))}</span>')
        action_cell = (f'<form method="post" action="/reports/apply" style="display:inline">'
                       f'<input type="hidden" name="applicant" value="{html.escape(aid)}">'
                       f'<input type="hidden" name="url" value="{html.escape(url)}">'
                       f'<input type="hidden" name="title" value="{html.escape(job.get("title",""))}">'
                       f'<input type="hidden" name="employer" value="{html.escape(job.get("company",""))}">'
                       f'<select name="status" style="padding:4px 6px;font-size:.8rem" onchange="this.form.submit()">'
                       + "".join(f'<option value="{k}"{" selected" if k==job.get("app_status") else ""}>{v["emoji"]} {v["fa"]}</option>'
                                 for k, v in unified_report.APP_STATUS.items())
                       + f'</select></form>'
                       f'<a class="btn ghost" style="padding:4px 8px;font-size:.75rem" '
                       f'href="{html.escape(url)}" target="_blank" rel="noopener">↗ آگهی</a>')
    else:
        badge = ""
        action_cell = (
            f'<a class="btn" style="padding:5px 10px;font-size:.78rem" '
            f'href="{html.escape(url)}" target="_blank" rel="noopener">👁 دیدن آگهی</a> '
            f'<form method="post" action="/reports/apply" style="display:inline">'
            f'<input type="hidden" name="applicant" value="{html.escape(aid)}">'
            f'<input type="hidden" name="url" value="{html.escape(url)}">'
            f'<input type="hidden" name="title" value="{html.escape(job.get("title",""))}">'
            f'<input type="hidden" name="employer" value="{html.escape(job.get("company",""))}">'
            f'<input type="hidden" name="country" value="{html.escape(ctry)}">'
            f'<input type="hidden" name="status" value="PLANNED">'
            f'<button class="btn primary" style="padding:5px 10px;font-size:.78rem" '
            f'type="submit">🚀 اقدام کن</button></form>'
            f'<form method="post" action="/reports/letter" style="display:inline">'
            f'<input type="hidden" name="applicant" value="{html.escape(aid)}">'
            f'<input type="hidden" name="url" value="{html.escape(url)}">'
            f'<input type="hidden" name="country" value="{html.escape(ctry)}">'
            f'<button class="btn ghost" style="padding:5px 10px;font-size:.78rem" '
            f'type="submit">✍️ کاورلتر</button></form>')

    link = (f'<a href="{html.escape(url)}" target="_blank" rel="noopener" '
            f'style="color:var(--accent);text-decoration:none">{title}</a>'
            if url else title)

    return f"""
    <tr>
      <td style="font-weight:700;color:var(--accent)">{score}</td>
      <td>{link}<div style="font-size:.72rem;color:var(--muted)">{why}</div>{badge}</td>
      <td>{company}</td>
      <td>{html.escape(ctry_fa)}</td>
      <td style="font-size:.72rem;color:var(--muted)">{html.escape(job.get('source','')[:26])}</td>
      <td style="white-space:nowrap">{action_cell}</td>
    </tr>"""


def _next_step_block(steps):
    if not steps:
        return ""
    rows = []
    for s in steps:
        rows.append(f"""
        <li style="display:flex;gap:10px;align-items:flex-start;padding:9px 0;
                   border-bottom:1px solid var(--border)">
          <span style="font-size:1.15rem">{s['emoji']}</span>
          <span style="flex:1">{html.escape(s['text'])}</span>
          {f'<a class="btn ghost" style="padding:3px 10px;font-size:.75rem" href="{html.escape(s["href"])}" '
             f'target="_blank" rel="noopener">{html.escape(s["cta"])}</a>' if s.get('href') else ''}
        </li>""")
    return f"""
    <div style="background:linear-gradient(135deg,var(--surface2),var(--surface));
                border:1px solid var(--border);border-radius:var(--radius);padding:14px 18px">
      <h3 style="margin:0 0 6px">🎯 الان چه کار کن</h3>
      <ul style="list-style:none;margin:0;padding:0">{''.join(rows)}</ul>
    </div>"""


def render_reports(applicant_id=None, message=None, error=None, country="FI"):
    import unified_report

    applicants = load_applicants()

    msg_html = ""
    if message:
        msg_html = f'<div class="card" style="border-color:var(--ok);margin-bottom:20px">{html.escape(message)}</div>'
    if error:
        msg_html += f'<div class="warn">{html.escape(error)}</div>'

    # ── فیلتر کشور ──
    country = (country or "FI").upper()
    if country not in APPLICANT_COUNTRIES:
        country = "ALL"
    csel = "".join(
        f'<option value="{c}"{" selected" if c == country else ""}>'
        f'{("همهٔ کشورها" if c == "ALL" else unified_report.COUNTRY_FA.get(c, c))}</option>'
        for c in APPLICANT_COUNTRIES)

    applicant = next((a for a in applicants if a.get("id") == applicant_id), None) if applicant_id else None

    if not applicant:
        # ── نمای کلی: هر متقاضی با شمارش زنده ──
        if not applicants:
            body = '<p class="empty">هنوز متقاضی‌ای در تنظیمات ثبت نشده.</p>'
        else:
            cards = ""
            for a in applicants:
                d = unified_report.applicant_dossier(a, limit=1, country=country)
                s = d["stats"]
                cards += f"""
                <a class="card applicant" style="text-decoration:none;display:block"
                   href="/reports?applicant={html.escape(a['id'])}&country={country}">
                  <div class="emoji">{html.escape(a.get('emoji','👤'))}</div>
                  <div class="name">{html.escape(a.get('name_fa') or a.get('name',''))}</div>
                  <div class="meta">{html.escape(a.get('profession',''))}</div>
                  <div style="margin-top:10px;font-size:.8rem;color:var(--muted);line-height:1.8">
                    🎯 <strong style="color:var(--accent)">{s['jobs_matched']}</strong> آگهی تطبیقی<br>
                    🚀 {s['applied']} اقدام‌شده · {s['open_apps']} در جریان<br>
                    📨 {s['emails_found']} ایمیل شغلی
                  </div>
                </a>"""
            body = f'<div class="grid">{cards}</div>'
        content = f"""
        <section>
          <h2>هر متقاضی چه دارد؟</h2>
          <p class="hint">همهٔ آگهی‌ها، درخواست‌ها، ایمیل‌ها و یادآورهای هر نفر
             از همهٔ بانک‌ها جمع شده — دیگر لازم نیست بین فایل‌ها بگردی.</p>
          {body}
        </section>"""
    else:
        aid = applicant["id"]
        d = unified_report.applicant_dossier(applicant, country=country)
        s = d["stats"]

        # ── کارت‌های آمار ──
        stat_cards = f"""
        <div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(115px,1fr));gap:10px">
          <div class="card" style="text-align:center;padding:12px 8px">
            <div style="font-size:1.5rem">🎯</div>
            <div style="font-size:1.7rem;font-weight:700">{s['jobs_matched']}</div>
            <div style="color:var(--muted);font-size:.8rem">آگهی تطبیقی</div></div>
          <div class="card" style="text-align:center;padding:12px 8px">
            <div style="font-size:1.5rem">🚀</div>
            <div style="font-size:1.7rem;font-weight:700">{s['applied']}</div>
            <div style="color:var(--muted);font-size:.8rem">اقدام‌شده</div></div>
          <div class="card" style="text-align:center;padding:12px 8px">
            <div style="font-size:1.5rem">📤</div>
            <div style="font-size:1.7rem;font-weight:700">{s['open_apps']}</div>
            <div style="color:var(--muted);font-size:.8rem">در جریان</div></div>
          <div class="card" style="text-align:center;padding:12px 8px">
            <div style="font-size:1.5rem">📬</div>
            <div style="font-size:1.7rem;font-weight:700">{s['emails_found']}</div>
            <div style="color:var(--muted);font-size:.8rem">ایمیل شغلی</div></div>
          <div class="card" style="text-align:center;padding:12px 8px;
                  {'border-color:var(--err)' if s['overdue'] else ''}">
            <div style="font-size:1.5rem">🚨</div>
            <div style="font-size:1.7rem;font-weight:700;color:{'var(--err)' if s['overdue'] else 'inherit'}">{s['overdue']}</div>
            <div style="color:var(--muted);font-size:.8rem">سررسید گذشته</div></div>
        </div>"""

        # ── جدول آگهی‌ها ──
        if d["jobs"]:
            rows = "".join(_job_card(j, aid, country) for j in d["jobs"])
            job_section = f"""
            <table class="files"><thead><tr>
              <th style="width:44px">امتیاز</th><th>عنوان و شرکت</th>
              <th>شرکت</th><th>کشور</th><th>منبع</th><th>اقدام</th>
            </tr></thead><tbody>{rows}</tbody></table>
            <p class="hint">امتیاز = تطبیق کلیدواژه + معادل‌های فنلاندی/سوئدی.
               آگهی‌های اقدام‌شده بالاتر می‌آیند.</p>"""
        else:
            job_section = (
                '<p class="empty">هیچ آگهی تطبیقی در این کشور نیست. '
                'کشور را از فیلتر بالا عوض کن یا crawler را اجرا کن.</p>')

        # ── درخواست‌ها ──
        if d["applications"]:
            arows = ""
            for a in d["applications"]:
                meta = unified_report.APP_STATUS.get(a.get("status", ""), {})
                dl = a.get("reply_deadline") or ""
                overdue = ""
                if dl and a.get("status") in ("SENT", "FOLLOW_UP"):
                    try:
                        if datetime.fromisoformat(dl) < datetime.now():
                            overdue = '<span style="color:var(--err);font-size:.75rem">⏰ مهلت گذشته</span>'
                    except Exception:
                        pass
                arows += f"""
                <tr>
                  <td>{meta.get('emoji','📋')} {html.escape(meta.get('fa', a.get('status','')))}</td>
                  <td>{html.escape(a.get('job','')[:60])}</td>
                  <td>{html.escape(a.get('employer','')[:30])}</td>
                  <td style="font-size:.78rem">{html.escape(a.get('sent_date',''))}
                      {overdue}</td>
                </tr>"""
            app_section = f"""
            <table class="files"><thead><tr>
              <th>وضعیت</th><th>موقعیت</th><th>کارفرما</th><th>تاریخ</th>
            </tr></thead><tbody>{arows}</tbody></table>"""
        else:
            app_section = '<p class="empty">هنوز اقدامی ثبت نشده. روی «🚀 اقدام کن» هر آگهی بزن.</p>'

        # ── ایمیل‌ها ──
        if d["emails"]:
            erows = ""
            for e in d["emails"][:10]:
                fr = (e.get("from") or "").strip()
                mailto = f"mailto:{fr}" if fr else "#"
                erows += f"""
                <tr>
                  <td style="font-size:.78rem;white-space:nowrap">{html.escape((e.get('date') or '')[:16])}</td>
                  <td>{html.escape(e.get('category',''))}</td>
                  <td>{html.escape((e.get('subject') or '')[:66])}</td>
                  <td><a href="{html.escape(mailto)}" style="color:var(--accent)">{html.escape(fr[:38])}</a></td>
                </tr>"""
            email_section = f"""
            <table class="files"><thead><tr>
              <th>تاریخ</th><th>دسته</th><th>موضوع</th><th>فرستنده</th>
            </tr></thead><tbody>{erows}</tbody></table>"""
        else:
            email_section = '<p class="empty">ایمیل شغلی برای این شخص پیدا نشده.</p>'

        content = f"""
        <section>
          <p><a href="/reports">← بازگشت به لیست متقاضی‌ها</a></p>
          {stat_cards}
          <div style="margin-top:16px">{_next_step_block(d['next_steps'])}</div>
        </section>
        <section>
          <h2>🎯 آگهی‌های تطبیق‌یافته</h2>
          {job_section}
        </section>
        <section>
          <h2>📤 درخواست‌های ثبت‌شده</h2>
          {app_section}
        </section>
        <section>
          <h2>📬 ایمیل‌های شغلی</h2>
          {email_section}
        </section>"""

    return f"""<!doctype html>
<html lang="fa"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Hunter — گزارش‌ها</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}</style>
</head>
<body>
<header class="top">
  <h1>📑 گزارش‌ها</h1>
  {nav_html('reports', 'job')}
</header>
<main>
  <div class="card" style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:18px">
    <strong>🌍 کشور:</strong>
    <select onchange="location.href='/reports{('?applicant=' + html.escape(aid) + '&') if applicant else '?'}country=' + this.value">
      {csel}
    </select>
    {f'<a class="btn" href="/reports/excel?applicant={html.escape(aid)}&country={country}">📊 دانلود اکسل همین متقاضی</a>' if applicant else ''}
    <span style="flex:1"></span>
    <span class="hint" style="margin:0">همهٔ بانک‌ها یکجا: کراولر + اسپانسر + درخواست + ایمیل + یادآور</span>
  </div>
  {msg_html}
  {content}
</main>
</body></html>"""


def safe_resolve(rel_path):
    """جلوگیری از path traversal — فقط اجازه‌ی دسترسی به داخل output/ و dashboard/."""
    full = os.path.normpath(os.path.join(BASE, rel_path))
    allowed_roots = [os.path.normpath(OUTPUT_DIR), os.path.normpath(DASHBOARD_DIR)]
    if not any(full.startswith(root + os.sep) or full == root for root in allowed_roots):
        return None
    if not os.path.isfile(full):
        return None
    return full


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # ساکت — نویز کنسول را کم می‌کند

    def _send(self, body, status=200, content_type="text/html; charset=utf-8"):
        # آیکون درون‌خطی را به هر صفحهٔ HTML تزریق می‌کنیم تا مرورگر
        # برای /favicon.ico یک 404 بی‌دلیل نزند (در کنسول خطا نشان می‌داد)
        if isinstance(body, str) and content_type.startswith("text/html") and FAVICON not in body:
            body = body.replace("</head>", FAVICON + "\n</head>", 1)
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _redirect(self, location):
        self.send_response(303)
        self.send_header("Location", location)
        self.end_headers()

    def _read_form(self):
        length = int(self.headers.get("Content-Length", 0) or 0)
        raw = self.rfile.read(length).decode("utf-8") if length else ""
        parsed = parse_qs(raw, keep_blank_values=True)
        return {k: v[0] for k, v in parsed.items()}

    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        # مسیر فعال (job/education) از هر صفحه به صفحه منتقل می‌شود
        track = (qs.get("track") or [None])[0]

        if parsed.path == "/":
            self._send(render_index(track))
        elif parsed.path == "/education":
            self._send(render_education((qs.get("msg") or [None])[0]))
        elif parsed.path == "/yield":
            self._send(render_yield())
        elif parsed.path == "/files":
            self._send(render_files())
        elif parsed.path == "/visa":
            self._send(render_visa((qs.get("applicant") or [None])[0],
                                   (qs.get("msg") or [None])[0]))
        elif parsed.path == "/visa/toggle":
            ap = (qs.get("applicant") or [""])[0]
            key = (qs.get("key") or [""])[0]
            _vt.toggle_stage(ap, key)
            self._redirect(f"/visa?applicant={ap}")
        elif parsed.path == "/visa/act-toggle":
            ap = (qs.get("applicant") or [""])[0]
            aid = (qs.get("id") or [0])[0]
            _vt.toggle_act(ap, aid)
            self._redirect(f"/visa?applicant={ap}")
        elif parsed.path == "/settings":
            self._send(render_settings())
        elif parsed.path == "/about":
            self._send(render_about())
        elif parsed.path == "/reports":
            aid = (qs.get("applicant") or [None])[0]
            msg = (qs.get("msg") or [None])[0]
            err = (qs.get("err") or [None])[0]
            ctry = (qs.get("country") or ["FI"])[0]
            self._send(render_reports(applicant_id=aid, message=msg, error=err, country=ctry))
        elif parsed.path == "/reports/excel":
            aid = (qs.get("applicant") or [None])[0]
            ctry = (qs.get("country") or ["FI"])[0]
            applicants = load_applicants()
            applicant = next((a for a in applicants if a.get("id") == aid), None)
            if not applicant:
                self._send(render_reports(error="متقاضی پیدا نشد."), status=404)
                return
            if not HAS_OPENPYXL:
                self._send(render_reports(applicant_id=aid, error="openpyxl نصب نیست، اکسل ساخته نمی‌شود.", country=ctry))
                return
            matched = unified_report.applicant_dossier(applicant, country=ctry)["jobs"]
            if not matched:
                self._send(render_reports(applicant_id=aid, country=ctry,
                                          error="آگهی تطبیق‌یافته‌ای برای ساخت اکسل وجود ندارد."))
                return
            wb = build_applicant_excel(applicant, matched)
            os.makedirs(DASHBOARD_DIR, exist_ok=True)
            fn = f"Report_{aid}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            wb.save(os.path.join(DASHBOARD_DIR, fn))
            self._redirect(f"/file?path=dashboard/{fn}")
        elif parsed.path == "/run-status":
            self._send(json.dumps(run_state), content_type="application/json")
        elif parsed.path == "/file":
            rel = (qs.get("path") or [""])[0]
            full = safe_resolve(rel)
            if not full:
                self._send("<p>فایل پیدا نشد یا اجازه‌ی دسترسی نیست.</p>", status=404)
                return
            if full.lower().endswith(".xlsx"):
                self._send(render_xlsx_as_html(full, rel))
            else:
                self._send(render_text_file(full, rel))
        elif parsed.path == "/download":
            rel = (qs.get("path") or [""])[0]
            full = safe_resolve(rel)
            if not full:
                self._send("فایل پیدا نشد.", status=404)
                return
            with open(full, "rb") as f:
                self._send(f.read(), content_type="application/octet-stream")
        else:
            self._send("<p>404</p>", status=404)

    def do_POST(self):
        raw_path = urlparse(self.path).path
        if raw_path == "/run":
            with run_lock:
                already_running = run_state["running"]
            if not already_running:
                qs = parse_qs(urlparse(self.path).query)
                track = (qs.get("track") or [None])[0]
                threading.Thread(target=run_pipeline_background,
                                 args=(track,), daemon=True).start()
            self._send(json.dumps({"ok": True}), content_type="application/json")

        elif raw_path == "/education/app-add":
            if not HAS_EDU_CRAWLER:
                self._send(render_education("ماژول education_crawler.py پیدا نشد."), status=500)
                return
            f = self._read_form()
            _ec.add_application(
                applicant=f.get("applicant", "").strip(),
                title=f.get("title", "").strip(),
                provider=f.get("provider", "").strip(),
                country=f.get("country", "").strip(),
                url=f.get("url", "").strip(),
                deadline=f.get("deadline", "").strip(),
                degree=f.get("degree", "").strip(),
                language=f.get("language", "").strip(),
                notes=f.get("notes", "").strip(),
            )
            self._redirect("/education?msg=" + urllib.parse.quote("به بانک اپلیکیشن اضافه شد."))

        elif raw_path == "/education/app-status":
            if not HAS_EDU_CRAWLER:
                self._send(render_education("ماژول education_crawler.py پیدا نشد."), status=500)
                return
            f = self._read_form()
            _ec.set_app_status(f.get("applicant", "").strip(), f.get("key", ""),
                               f.get("status", "PLANNED"), f.get("note", "").strip())
            self._redirect("/education?msg=" + urllib.parse.quote("وضعیت اپلیکیشن ثبت شد."))

        elif raw_path == "/education/app-delete":
            if not HAS_EDU_CRAWLER:
                self._send(render_education("ماژول education_crawler.py پیدا نشد."), status=500)
                return
            f = self._read_form()
            _ec.remove_application(f.get("applicant", "").strip(), f.get("key", ""))
            self._redirect("/education?msg=" + urllib.parse.quote("اپلیکیشن حذف شد."))

        elif raw_path == "/visa/new":
            if not HAS_VISA:
                self._send("<p>ماژول ویزا نصب نیست.</p>", status=500)
                return
            f = self._read_form()
            ap = f.get("applicant", "").strip()
            if not ap:
                self._send(render_visa(message="خطا: نام متقاضی الزامی است."), status=400)
                return
            _vt.set_target(ap, f.get("target", "").strip())
            self._redirect("/visa?applicant=" + urllib.parse.quote(ap))

        elif raw_path == "/visa/target":
            f = self._read_form()
            _vt.set_target(f.get("applicant", "").strip(), f.get("target", "").strip())
            self._redirect("/visa?applicant=" + urllib.parse.quote(f.get("applicant", "")))

        elif raw_path == "/visa/stage-note":
            f = self._read_form()
            _vt.set_note_stage(f.get("applicant", "").strip(), f.get("key", ""),
                               f.get("note", "").strip())
            self._redirect("/visa?applicant=" + urllib.parse.quote(f.get("applicant", "")))

        elif raw_path == "/visa/act-add":
            f = self._read_form()
            _vt.add_act(f.get("applicant", "").strip(), f.get("category", "other"),
                        f.get("title", ""), f.get("detail", ""), f.get("date") or None)
            self._redirect("/visa?applicant=" + urllib.parse.quote(f.get("applicant", "")))

        elif raw_path == "/visa/act-delete":
            f = self._read_form()
            _vt.delete_act(f.get("applicant", "").strip(), f.get("id", "0"))
            self._redirect("/visa?applicant=" + urllib.parse.quote(f.get("applicant", "")))

        elif raw_path == "/reports/apply":
            # ثبت «اقدام» روی یک آگهی از صفحهٔ گزارش — همان چیزی که
            # کاربر قبلاً باید دستی در APPLICATION_BANK.json می‌نوشت.
            f = self._read_form()
            aid = f.get("applicant", "").strip()
            url = f.get("url", "").strip()
            title = f.get("title", "").strip()
            employer = f.get("employer", "").strip() or "نامشخص"
            status = (f.get("status") or "PLANNED").strip().upper()
            ctry = (f.get("country") or "FI").strip().upper()
            if not aid or not url:
                self._send(render_reports(error="خطا: متقاضی یا آدرس آگهی مشخص نیست."), status=400)
                return
            if status not in unified_report.APP_STATUS:
                status = "PLANNED"

            ok, note = _upsert_application(aid, url, title, employer, ctry, status)
            msg = (f"✅ «{title[:48]}» برای «{aid}» ثبت شد — وضعیت: "
                   f"{unified_report.APP_STATUS[status]['fa']}"
                   + (f" ({note})" if note else ""))
            dest = f"/reports?applicant={urllib.parse.quote(aid)}"
            if not ok:
                self._send(render_reports(applicant_id=aid, error=msg, country=ctry), status=400)
                return
            self._redirect(dest)

        elif raw_path == "/settings/applicant":
            form = self._read_form()
            if not form.get("id", "").strip():
                self._send(render_settings(message="خطا: شناسه (id) الزامی است."), status=400)
                return
            app_id = upsert_applicant(form)
            self._send(render_settings(message=f"مشخصات «{app_id}» ذخیره شد."))

        elif self.path == "/settings/credentials":
            form = self._read_form()
            app_id = form.get("id", "").strip().lower()
            if not app_id:
                self._send(render_settings(message="خطا: شناسه‌ی متقاضی مشخص نیست."), status=400)
                return
            aid_upper = app_id.upper()
            write_env_updates({
                f"EMAIL_{aid_upper}": form.get("email", "").strip(),
                f"EMAIL_PASSWORD_{aid_upper}": form.get("password", "").strip(),
                f"EMAIL_PROVIDER_{aid_upper}": form.get("provider", "").strip(),
            })
            self._send(render_settings(message=f"اطلاعات ایمیل «{app_id}» در .env ذخیره شد (رمز نمایش داده نمی‌شود)."))

        elif self.path == "/settings/sources/toggle":
            form = self._read_form()
            try:
                idx = int(form.get("index", "-1"))
            except ValueError:
                idx = -1
            sources = read_sources() or []
            if 0 <= idx < len(sources):
                sources[idx]["enabled"] = not sources[idx].get("enabled", True)
                write_sources(sources)
                state = "فعال" if sources[idx]["enabled"] else "غیرفعال"
                self._send(render_settings(message=f"منبع «{sources[idx].get('name','')}» {state} شد."))
            else:
                self._send(render_settings(message="منبع پیدا نشد (لیست را رفرش کن)."), status=400)

        elif self.path == "/settings/sources/add":
            form = self._read_form()
            name = form.get("name", "").strip()
            country = form.get("country", "").strip().upper()
            url = form.get("url", "").strip()
            try:
                trust = int(form.get("trust", "60"))
            except ValueError:
                trust = 60
            if not (name and country and url):
                self._send(render_settings(message="خطا: نام، کشور و URL الزامی است."), status=400)
                return
            sources = read_sources() or []
            sources.append({
                "name": name, "country": country, "type": "custom", "enabled": True,
                "url": url, "search_urls": [url], "trust": max(0, min(100, trust)),
            })
            write_sources(sources)
            self._send(render_settings(message=f"منبع «{name}» اضافه شد و در اجرای بعدی پایپ‌لاین جستجو می‌شود."))

        elif raw_path == "/reports/letter":
            form = self._read_form()
            aid = form.get("id", "").strip().lower() or form.get("applicant", "").strip().lower()
            ctry = (form.get("country") or "FI").strip().upper()
            # آدرس آگهی را می‌گیریم نه شمارهٔ ردیف — با عوض شدن فیلتر کشور
            # یا مرتب‌سازی، index به آگهی دیگری اشاره می‌کرد.
            job_url = form.get("url", "").strip()

            applicants = load_applicants()
            applicant = next((a for a in applicants if a.get("id") == aid), None)
            if not applicant:
                self._send(render_reports(error="متقاضی پیدا نشد.", country=ctry), status=404)
                return

            if job_url:
                target = unified_report._norm(job_url)
                job = next((j for j in unified_report.all_jobs()
                            if unified_report._norm(j.get("url")) == target), None)
            else:
                try:
                    job_index = int(form.get("job_index", "-1"))
                except ValueError:
                    job_index = -1
                matched = unified_report.applicant_dossier(applicant, country=ctry)["jobs"]
                job = matched[job_index] if 0 <= job_index < len(matched) else None

            if not job:
                self._send(render_reports(applicant_id=aid, country=ctry,
                                          error="این آگهی پیدا نشد (شاید لیست عوض شده — صفحه را رفرش کن)."))
                return

            prompt = build_cover_letter_prompt(applicant, job)
            text, err = call_ai(prompt)
            if err:
                self._send(render_reports(applicant_id=aid, error=err, country=ctry))
                return

            os.makedirs(OUTPUT_DIR, exist_ok=True)
            fn = f"COVER_{aid}_{datetime.now().strftime('%Y%m%d_%H%M')}.md"
            header = (f"# کاور لتر و ایمیل — {applicant.get('name_fa') or applicant.get('name','')}\n"
                      f"برای آگهی: {job.get('title','')} — {job.get('company') or 'نامشخص'} ({job.get('country','')})\n"
                      f"ساخته‌شده با AI در {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n---\n\n")
            with open(os.path.join(OUTPUT_DIR, fn), "w", encoding="utf-8") as f:
                f.write(header + text)
            self._redirect(f"/file?path=output/{fn}")

        else:
            self._send("<p>404</p>", status=404)


def main():
    parser = argparse.ArgumentParser(description="Migration Hunter — Web UI")
    parser.add_argument("--port", type=int, default=8877)
    args = parser.parse_args()

    server = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    print(f"✅ داشبورد آماده است → http://127.0.0.1:{args.port}")
    print("برای توقف: Ctrl+C")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nمتوقف شد.")


if __name__ == "__main__":
    main()
