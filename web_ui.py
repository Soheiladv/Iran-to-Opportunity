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
CONFIG_PATH = os.path.join(BASE, "config.json")
ENV_PATH = os.path.join(BASE, ".env")
GITIGNORE_PATH = os.path.join(BASE, ".gitignore")
SOURCES_PATH = os.path.join(BASE, "sources.json")

try:
    import job_crawler as _jc  # برای استفاده‌ی مجدد از load_sources/save_sources، بدون تکرار کد
    HAS_JOB_CRAWLER = True
except Exception:
    HAS_JOB_CRAWLER = False

try:
    from openpyxl import load_workbook, Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# پایپ‌لاین اجرا — همان ۵ مرحله‌ی run.py
PIPELINE_STEPS = [
    {"key": "email_analyze", "name": "تحلیل ایمیل شغلی", "script": "email_analyzer.py", "emoji": "📧"},
    {"key": "email_excel", "name": "ساخت Excel ایمیل", "script": "email_dashboard.py", "emoji": "📊"},
    {"key": "job_search", "name": "جستجوی خودکار کار", "script": "job_crawler.py", "emoji": "🔍"},
    {"key": "followup", "name": "یادآوری پیگیری", "script": "followup_reminder.py", "emoji": "⏰"},
    {"key": "dashboard", "name": "ساخت داشبورد اصلی", "script": "build_dashboard.py", "emoji": "📈"},
]

# ── وضعیت اجرای پایپ‌لاین در حافظه (thread-safe به‌قدر کافی برای یک کاربر محلی) ──
run_state = {
    "running": False, "log": [], "started_at": None, "finished_at": None,
    "step_index": 0, "step_total": len(PIPELINE_STEPS), "current_step": "",
    # 🔴 پنل زنده جستجو — اسکریپت‌ها الان کدام منبع/آدرس را می‌کَویند
    "live": {"active": False, "source": "", "source_idx": 0, "source_total": 0,
             "url": "", "keyword": "", "sources_done": [], "jobs_found": 0},
}
run_lock = threading.Lock()

import re as _re

_LIVE_PATTERNS = {
    # 📡 [3/18] Seek AU (AU) - در حال بررسی...
    "source": _re.compile(r"📡\s*\[(\d+)/(\d+)\]\s*(.+?)\s*\((\w+)\)\s*-\s*در حال بررسی"),
    # 🔎 [1/2] کلیدواژه: «midwife» ← https://...
    "url": _re.compile(r"🔎\s*\[\d+/\d+\]\s*کلیدواژه:\s*«(.+?)»\s*←\s*(\S+)"),
    # ⚠️ مجموعاً 0 آگهی یافت شد از Seek NZ  |  ✅ ... 12 آگهی ...
    "done": _re.compile(r"مجموعاً (\d+) آگهی یافت شد از (.+)"),
}


def _update_live_state(line):
    """خطوط لاگِ در حال استریم را پارس می‌کند و پنل جستجوی زنده را به‌روز می‌کند."""
    live = run_state["live"]
    m = _LIVE_PATTERNS["source"].search(line)
    if m:
        live["active"] = True
        live["source_idx"], live["source_total"] = int(m.group(1)), int(m.group(2))
        live["source"] = m.group(3).strip()
        live["url"], live["keyword"] = "", ""
        return
    m = _LIVE_PATTERNS["url"].search(line)
    if m:
        live["keyword"], live["url"] = m.group(1), m.group(2)
        return
    m = _LIVE_PATTERNS["done"].search(line)
    if m:
        found, name = int(m.group(1)), m.group(2).strip()
        live["jobs_found"] += found
        live["sources_done"].append({"name": name, "jobs": found})
        # url/keyword را پاک نکن — خط «📡 منبع جدید» خودش آنها را ریست می‌کند
        # و این‌طوری آخرین URL در پنل زنده تا منبع بعدی دیده می‌ماند.


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


def run_pipeline_background():
    with run_lock:
        if run_state["running"]:
            return
        run_state["running"] = True
        run_state["log"] = []
        run_state["started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        run_state["finished_at"] = None
        run_state["step_index"] = 0
        run_state["step_total"] = len(PIPELINE_STEPS)
        run_state["current_step"] = ""
        run_state["live"] = {"active": False, "source": "", "source_idx": 0, "source_total": 0,
                             "url": "", "keyword": "", "sources_done": [], "jobs_found": 0}

    for i, step in enumerate(PIPELINE_STEPS, 1):
        run_state["step_index"] = i
        run_state["current_step"] = f"{step['emoji']} {step['name']}"
        script_path = os.path.join(BASE, step["script"])
        if not os.path.exists(script_path):
            _append_log(f"⏭️  {step['emoji']} {step['name']} — فایل {step['script']} پیدا نشد، رد شد")
            continue
        _append_log(f"▶ [{i}/{len(PIPELINE_STEPS)}] {step['emoji']} {step['name']} در حال اجرا…")
        run_state["live"]["active"] = step["key"] in ("job_search",)
        try:
            # -u = خروجی بدون بافر، تا خط‌به‌خط همین‌جا زنده دیده شود
            proc = subprocess.Popen(
                [sys.executable, "-u", script_path],
                cwd=BASE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1,
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
nav.tabs{display:flex; gap:6px}
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
"""


def nav_html(active):
    def cls(name):
        return "active" if name == active else ""
    return f"""<nav class="tabs">
  <a class="{cls('dashboard')}" href="/">داشبورد</a>
  <a class="{cls('reports')}" href="/reports">📑 گزارش‌ها</a>
  <a class="{cls('settings')}" href="/settings">⚙️ تنظیمات</a>
  <a class="{cls('about')}" href="/about">ℹ️ توضیحات</a>
</nav>"""


def render_index():
    applicants = load_applicants()
    if applicants:
        cards = "".join(f"""
          <div class="card applicant">
            <div class="emoji">{html.escape(a.get('emoji','👤'))}</div>
            <div class="name">{html.escape(a.get('name_fa') or a.get('name','?'))}</div>
            <div class="meta">{html.escape(a.get('profession',''))}</div>
          </div>""" for a in applicants)
    else:
        cards = '<p class="empty">هنوز کسی در config.json تعریف نشده — از config.json.example شروع کن.</p>'

    steps_html = "".join(
        f'<li>{s["emoji"]} {html.escape(s["name"])} <span style="color:var(--muted);font-size:.85rem">({s["script"]})</span></li>'
        for s in PIPELINE_STEPS
    )

    md_files = list_files(OUTPUT_DIR, exts=[".md", ".txt"])
    xlsx_files = list_files(DASHBOARD_DIR, exts=[".xlsx"])

    def file_rows(files, icon):
        if not files:
            return '<tr><td colspan="2" class="empty">چیزی هنوز تولید نشده</td></tr>'
        rows = []
        for f in files[:25]:
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
<title>Migration Hunter — داشبورد</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css">
<style>{PAGE_STYLE}</style>
</head>
<body>
<header class="top">
  <h1>🧭 Migration Hunter</h1>
  {nav_html('dashboard')}
</header>
<main>
  <section>
    <h2>متقاضی‌ها</h2>
    <div class="grid">{cards}</div>
  </section>

  <section>
    <h2>اجرای پایپ‌لاین</h2>
    <p>پنج مرحله به‌ترتیب اجرا می‌شوند و نتیجه‌شان در پایین لاگ می‌شود:</p>
    <ol class="steps">{steps_html}</ol>
    <p style="margin-top:16px">
      <button class="btn" id="runBtn" onclick="startRun()">▶ اجرای پایپ‌لاین</button>
    </p>

    <div id="progressWrap" style="display:none; margin:14px 0 6px">
      <div style="display:flex; justify-content:space-between; font-size:.85rem; color:var(--muted); margin-bottom:4px">
        <span id="progressLabel">—</span>
        <span id="progressPct">0%</span>
      </div>
      <div style="background:var(--line); border-radius:20px; height:10px; overflow:hidden">
        <div id="progressBar" style="background:var(--amber); height:100%; width:0%; transition:width .3s"></div>
      </div>
    </div>

    <div id="livePanel" style="display:none; margin:14px 0 6px">
      <div class="card" style="border-color:var(--amber); background:#fdf8ef">
        <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px">
          <span style="width:10px; height:10px; border-radius:50%; background:var(--err); display:inline-block; animation:pulse 1.2s infinite"></span>
          <strong style="color:var(--amber)">🔴 در حال جستجوی زنده…</strong>
          <span id="liveCounter" style="color:var(--muted); font-size:.85rem"></span>
        </div>
        <div style="display:flex; gap:14px; flex-wrap:wrap; font-size:.9rem">
          <div><span style="color:var(--muted)">📡 منبع فعلی:</span> <strong id="liveSource">—</strong></div>
          <div><span style="color:var(--muted)">🔎 کلیدواژه:</span> <span id="liveKeyword">—</span></div>
        </div>
        <div id="liveUrl" style="direction:ltr; text-align:left; font-family:monospace; font-size:.8rem; color:var(--teal); margin-top:6px; word-break:break-all">—</div>
        <div style="margin-top:8px; font-size:.85rem">
          <span style="color:var(--muted)">✅ منابع تمام‌شده:</span>
          <span id="liveDone">هنوز هیچ منبعی تمام نشده</span>
        </div>
        <div style="margin-top:4px; font-size:.85rem; color:var(--ok)">
          <span style="color:var(--muted)">🧲 مجموع آگهی‌های یافت‌شده تا الان:</span> <strong id="liveJobs">0</strong>
        </div>
      </div>
    </div>

    <div id="log">آماده به اجرا. برای شروع دکمه‌ی بالا را بزن.</div>
  </section>

  <section>
    <h2>گزارش‌ها (output/)</h2>
    <table class="files"><tbody>{file_rows(md_files, "📄")}</tbody></table>
  </section>

  <section>
    <h2>داشبوردهای اکسل (dashboard/)</h2>
    <table class="files"><tbody>{file_rows(xlsx_files, "📊")}</tbody></table>
  </section>

  {app_bank_section()}

  <footer style="text-align:center; padding:30px 0 10px; color:var(--muted); font-size:.8rem; border-top:1px solid var(--line); margin-top:30px">
    MigrationHunter v1.0 · آخرین به‌روزرسانی: {datetime.now().strftime("%Y-%m-%d %H:%M")}
  </footer>
</main>

<script>
let polling = null;
function startRun(){{
  fetch('/run', {{method:'POST'}}).then(()=>{{
    document.getElementById('runBtn').disabled = true;
    document.getElementById('progressWrap').style.display = 'block';
    document.getElementById('log').textContent = 'شروع شد…';
    polling = setInterval(pollStatus, 900);
    pollStatus();
  }});
}}
function pollStatus(){{
  fetch('/run-status').then(r=>r.json()).then(s=>{{
    const logBox = document.getElementById('log');
    logBox.textContent = s.log.join('\\n') || '…';
    logBox.scrollTop = logBox.scrollHeight;  // همیشه آخرین خط لایو دیده شود

    const total = s.step_total || 1;
    const idx = s.step_index || 0;
    const pct = Math.round((idx / total) * 100);
    document.getElementById('progressBar').style.width = pct + '%';
    document.getElementById('progressPct').textContent = pct + '%';
    document.getElementById('progressLabel').textContent =
      s.running ? (`مرحله ${{idx}}/${{total}} — ${{s.current_step || '...'}}`) : 'تمام شد';

    // 🔴 پنل زندهٔ جستجو
    const lp = document.getElementById('livePanel');
    const lv = s.live || {{}};
    if (s.running && lv.active) {{
      lp.style.display = 'block';
      document.getElementById('liveSource').textContent = lv.source || '…';
      document.getElementById('liveKeyword').textContent = lv.keyword || '…';
      document.getElementById('liveUrl').textContent = lv.url || '—';
      document.getElementById('liveJobs').textContent = lv.jobs_found || 0;
      document.getElementById('liveCounter').textContent =
        lv.source_total ? `منبع ${{lv.source_idx || 0}} از ${{lv.source_total}}` : '';
      const done = lv.sources_done || [];
      document.getElementById('liveDone').textContent = done.length
        ? done.map(d => `${{d.name}} (${{d.jobs}})`).join(' · ')
        : 'هنوز هیچ منبعی تمام نشده';
    }} else {{
      lp.style.display = 'none';
    }}

    if(!s.running){{
      clearInterval(polling);
      document.getElementById('runBtn').disabled = false;
      document.getElementById('progressBar').style.width = '100%';
      document.getElementById('progressPct').textContent = '100%';
      logBox.textContent += '\\n\\n— پایان اجرا — صفحه را رفرش کن تا فایل‌های جدید را ببینی —';
      logBox.scrollTop = logBox.scrollHeight;
    }}
  }});
}}
</script>
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
  {nav_html('settings')}
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
  {nav_html('about')}
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
CRAWLER_RESULTS_PATH = os.path.join(MEM_DIR, "CRAWLER_RESULTS.json")


def load_crawler_jobs():
    """آگهی‌های آخرین اجرای job_crawler.py را می‌خواند (اگر وجود داشته باشد)."""
    if not os.path.exists(CRAWLER_RESULTS_PATH):
        return None
    try:
        with open(CRAWLER_RESULTS_PATH, "r", encoding="utf-8") as f:
            return json.load(f).get("jobs", [])
    except Exception:
        return []


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


def matched_jobs_for_applicant(applicant, jobs, limit=200):
    scored = [(score_job_for_applicant(j, applicant), j) for j in (jobs or [])]
    scored = [(s, j) for s, j in scored if s > 0]
    scored.sort(key=lambda x: x[0], reverse=True)
    out = []
    for s, j in scored[:limit]:
        jj = dict(j)
        jj["_score"] = s
        out.append(jj)
    return out


def build_applicant_excel(applicant, jobs):
    wb = Workbook()
    ws = wb.active
    ws.title = "گزارش آگهی‌ها"
    ws.sheet_view.rightToLeft = True
    from openpyxl.styles import Font, PatternFill, Alignment
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1B4F72", end_color="1B4F72", fill_type="solid")
    headers = ["#", "عنوان شغل", "شرکت", "کشور", "منبع", "لینک آگهی", "امتیاز تطبیق"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    for i, j in enumerate(jobs, 1):
        ws.append([i, j.get("title", ""), j.get("company", "نامشخص"),
                    j.get("country", ""), j.get("source", ""), j.get("url", ""), j.get("_score", 0)])
    widths = {"A": 5, "B": 45, "C": 25, "D": 10, "E": 20, "F": 45, "G": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.auto_filter.ref = f"A1:G{len(jobs)+1}"
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


def render_reports(applicant_id=None, message=None, error=None):
    applicants = load_applicants()
    jobs = load_crawler_jobs()  # None = crawler never run; [] = ran but nothing/no match stored

    msg_html = ""
    if message:
        msg_html = f'<div class="card" style="border-color:var(--ok);margin-bottom:20px">{html.escape(message)}</div>'
    if error:
        msg_html += f'<div class="warn">{html.escape(error)}</div>'

    applicant = next((a for a in applicants if a.get("id") == applicant_id), None) if applicant_id else None

    if not applicant:
        if not applicants:
            body = '<p class="empty">هنوز متقاضی‌ای در تنظیمات ثبت نشده.</p>'
        else:
            cards = "".join(f"""
              <a class="card applicant" style="text-decoration:none;display:block"
                 href="/reports?applicant={html.escape(a['id'])}">
                <div class="emoji">{html.escape(a.get('emoji','👤'))}</div>
                <div class="name">{html.escape(a.get('name_fa') or a.get('name',''))}</div>
                <div class="meta">{html.escape(a.get('profession',''))}</div>
              </a>""" for a in applicants)
            body = f'<div class="grid">{cards}</div>'
        content = f"""
        <section><h2>گزارش هر متقاضی را انتخاب کن</h2>{body}</section>"""
    else:
        aid = applicant["id"]
        if jobs is None:
            job_section = ('<p class="empty">هنوز crawler اجرا نشده. از تب داشبورد، «اجرای پایپ‌لاین» '
                            'را بزن تا آگهی جمع‌آوری شود، بعد دوباره اینجا را باز کن.</p>')
        else:
            matched = matched_jobs_for_applicant(applicant, jobs)
            if not matched:
                job_section = '<p class="empty">هیچ آگهی‌ای با کلیدواژه‌های این متقاضی تطبیق نداشت.</p>'
            else:
                rows = []
                for i, j in enumerate(matched[:50]):
                    rows.append(f"""
                    <tr>
                      <td>{i+1}</td>
                      <td>{html.escape(j.get('title','')[:70])}</td>
                      <td>{html.escape(j.get('company') or 'نامشخص')}</td>
                      <td>{html.escape(j.get('country',''))}</td>
                      <td>{html.escape(j.get('source',''))}</td>
                      <td>{j.get('_score',0)}</td>
                      <td>
                        <form method="post" action="/reports/letter" style="display:inline">
                          <input type="hidden" name="id" value="{html.escape(aid)}">
                          <input type="hidden" name="job_index" value="{i}">
                          <button class="btn" style="padding:5px 12px;font-size:.82rem" type="submit">✍️ کاور لتر + ایمیل</button>
                        </form>
                      </td>
                    </tr>""")
                job_section = f"""
                <table class="files"><thead><tr>
                  <th>#</th><th>عنوان</th><th>شرکت</th><th>کشور</th><th>منبع</th><th>امتیاز</th><th></th>
                </tr></thead><tbody>{''.join(rows)}</tbody></table>
                <p class="hint">امتیاز = تطبیق کلیدواژه (نه اصالت‌سنجی واقعی — توضیح کامل در تب «ℹ️ توضیحات»)</p>
                """

        content = f"""
        <section>
          <h2>{html.escape(applicant.get('emoji','👤'))} گزارش {html.escape(applicant.get('name_fa') or applicant.get('name',aid))}</h2>
          <p><a href="/reports">← بازگشت به لیست متقاضی‌ها</a></p>
          <p style="margin-top:14px">
            <a class="btn" href="/reports/excel?applicant={html.escape(aid)}">📊 دانلود گزارش اکسل این متقاضی</a>
          </p>
        </section>
        <section>
          <h2>آگهی‌های تطبیق‌یافته</h2>
          {job_section}
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
  {nav_html('reports')}
</header>
<main>
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

        if parsed.path == "/":
            self._send(render_index())
        elif parsed.path == "/settings":
            self._send(render_settings())
        elif parsed.path == "/about":
            self._send(render_about())
        elif parsed.path == "/reports":
            aid = (qs.get("applicant") or [None])[0]
            msg = (qs.get("msg") or [None])[0]
            err = (qs.get("err") or [None])[0]
            self._send(render_reports(applicant_id=aid, message=msg, error=err))
        elif parsed.path == "/reports/excel":
            aid = (qs.get("applicant") or [None])[0]
            applicants = load_applicants()
            applicant = next((a for a in applicants if a.get("id") == aid), None)
            if not applicant:
                self._send(render_reports(error="متقاضی پیدا نشد."), status=404)
                return
            if not HAS_OPENPYXL:
                self._send(render_reports(applicant_id=aid, error="openpyxl نصب نیست، اکسل ساخته نمی‌شود."))
                return
            jobs = load_crawler_jobs() or []
            matched = matched_jobs_for_applicant(applicant, jobs)
            if not matched:
                self._send(render_reports(applicant_id=aid, error="آگهی تطبیق‌یافته‌ای برای ساخت اکسل وجود ندارد."))
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
        if self.path == "/run":
            with run_lock:
                already_running = run_state["running"]
            if not already_running:
                threading.Thread(target=run_pipeline_background, daemon=True).start()
            self._send(json.dumps({"ok": True}), content_type="application/json")

        elif self.path == "/settings/applicant":
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

        elif self.path == "/reports/letter":
            form = self._read_form()
            aid = form.get("id", "").strip().lower()
            try:
                job_index = int(form.get("job_index", "-1"))
            except ValueError:
                job_index = -1

            applicants = load_applicants()
            applicant = next((a for a in applicants if a.get("id") == aid), None)
            if not applicant:
                self._send(render_reports(error="متقاضی پیدا نشد."), status=404)
                return

            jobs = load_crawler_jobs() or []
            matched = matched_jobs_for_applicant(applicant, jobs)
            if job_index < 0 or job_index >= len(matched):
                self._send(render_reports(applicant_id=aid, error="این آگهی پیدا نشد (شاید لیست تغییر کرده — صفحه را رفرش کن)."))
                return

            job = matched[job_index]
            prompt = build_cover_letter_prompt(applicant, job)
            text, err = call_ai(prompt)
            if err:
                self._send(render_reports(applicant_id=aid, error=err))
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
