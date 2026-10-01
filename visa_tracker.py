#!/usr/bin/env python3
"""
MigrationHunter — نورد ویزای همراه

دو چیز را نگه می‌دارد، هر دو در memory/VISA_JOURNEY.json:

۱) چک‌لیست مراحل ویزا — هر مرحله یک‌بار تیک می‌خورد و درصد پیشرفت می‌دهد.
   مراحل بر اساس واقعیت مسیر مهاجرت چیده شده‌اند و برای ویزای کاری/تحصیلی
   و همراه (خانواده) قابل استفاده‌اند.

۲) دفترچهٔ اکت‌ها — هر کاری که واقعاً انجام دادی با تاریخ ثبت می‌شود:
   «۱۴ مهر — مدارک را جمع کردم»، «۲۰ مهر — وقت سفارت گرفتم». این همان
   چیزی است که بعداً لازم می‌شود: تاریخچهٔ دقیقِ اینکه چه کردی و کِی.

این ماژول هیچ وابستگی بیرونی ندارد.
"""
import json
import os
from datetime import datetime, timedelta

BASE = os.path.dirname(os.path.abspath(__file__))
MEM = os.path.join(BASE, "memory")
VISA_PATH = os.path.join(MEM, "VISA_JOURNEY.json")

# ── چک‌لیست مراحل ────────────────────────────────────────────────
# (key, عنوان, راهنمای کوتاه، دسته)
# دسته‌ها: prep / apply / embassy / decision / arrival / family
VISA_STAGES = [
    ("docs_passport",    "پاسپورت",              "گذرنامه با حداقل ۶ ماه اعتبار + صفحهٔ تمدید", "prep"),
    ("docs_photos",      "عکس",                   "عکس پاسپورتی، پس‌زمینهٔ سفید، ۳۵×۴۵", "prep"),
    ("docs_diploma",     "مدارک تحصیلی",          "دیپلم/لیسانس + ریزنامه، ترجمهٔ رسمی", "prep"),
    ("docs_work_exp",    "سوابق شغلی",            "گواهی کار، فیش حقوقی، بیمه", "prep"),
    ("docs_language",    "مدرک زبان",             "IELTS/TOEFL یا معادل — مدرک نامعتبر یعنی رد", "prep"),
    ("docs_police",      "گواهی عدم سابقه",        "پلیس فینری — بعضی سفارت‌ها بیش از ۶ ماه نمی‌پذیرند", "prep"),
    ("docs_translation", "ترجمهٔ رسمی",           "همهٔ مدارک با مُهر دادگستری/سفارت", "prep"),
    ("docs_marriage",    "عقدنامه/گواهی ازدواج",  "برای ویزای همراه الزامی", "family"),
    ("form_filled",      "فرم درخواست",           "فرم رسمی پر و امضا شده", "apply"),
    ("fee_paid",         "پرداخت هزینهٔ درخواست",  "رسیدش را نگه دار — بدون رسید، درخواست بی‌اعتبار", "apply"),
    ("booked_appointment", "نوبت سفارت/کنسول",   "تاریخ و ساعت + آدرس کنسولگری", "embassy"),
    ("interview_done",   "مصاحبه",                "تاریخ مصاحبه و نتیجه‌اش", "embassy"),
    ("biometrics_done",  "اثر انگشت و عکس",       "حضوری در کنسولگری", "embassy"),
    ("decision",         "تصمیم نهایی",           "نتیجهٔ اعلان‌شده — با تاریخ", "decision"),
    ("permit_issued",    "ویزا/اجازهٔ اقامت",     "شمارهٔ ویزا و تاریخ انقضا", "decision"),
    ("travel_booked",    "بلیت هوا",               "تاریخ سفر", "arrival"),
    ("arrived",          "ورود به کشور مقصد",     "تاریخ ورود", "arrival"),
    ("registered",       "ثبت اقامت",             "ثبت آدرس در ادارهٔ مهاجرت", "arrival"),
    ("family_applied",   "اپلای همراه",           "برای همسر/فرزند اقدام شده", "family"),
    ("family_visa",      "ویزای همراه",           "همسر/فرزند ویزا گرفتند", "family"),
]

STAGE_KEYS = {k for k, _t, _h, _c in VISA_STAGES}
STAGE_TITLE = {k: t for k, t, _h, _c in VISA_STAGES}
STAGE_CAT = {k: c for k, _t, _h, c in VISA_STAGES}

CAT_FA = {
    "prep": "📋 آماده‌سازی مدارک",
    "apply": "✍️ ثبت درخواست",
    "embassy": "🏛️ سفارت و کنسولگری",
    "decision": "📬 تصمیم",
    "arrival": "✈️ سفر و اقامت",
    "family": "👨‍👩‍👧 همراه",
}
CAT_ORDER = ["prep", "apply", "embassy", "decision", "arrival", "family"]

# دسته‌بندی اکت‌ها — تا بتوانیم گزارش را مرتب کنیم
ACT_CATEGORIES = [
    ("doc",    "📄 مدارک",      "تهیه/ترجمه/ارسال مدرک"),
    ("booking", "📅 نوبت",       "گرفتن وقت از سفارت یا مرکز"),
    ("money",  "💳 پرداخت",      "پرداخت هزینه، شهریه، بلیت، بیمه"),
    ("exam",   "📝 آزمون",       "آزمون زبان، آزمون ورودی، مصاحبه"),
    ("submit", "📤 ارسال",       "ارسال درخواست یا فرم"),
    ("result", "📥 نتیجه",       "دریافت نتیجه، پذیرش، رد، یا ویزا"),
    ("study",  "🎓 تحصیل",       "ثبت‌نام دانشگاه، انتخاب واحد، بورسیه"),
    ("job",    "💼 کار",         "مصاحبهٔ کاری، پیشنهاد شغلی، قرارداد"),
    ("move",   "📦 جابه‌جایی",   "اسباب‌کشی، سفر، ثبت اقامت"),
    ("other",  "📌 سایر",        "هر چیز دیگر"),
]
ACT_CATS = {k for k, _e, _d in ACT_CATEGORIES}
ACT_CAT_EMOJI = {k: e for k, e, _d in ACT_CATEGORIES}
ACT_CAT_FA = {k: fa for k, _e, fa in ACT_CATEGORIES}


def _today():
    return datetime.now().strftime("%Y-%m-%d")


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def load_journey(applicant=None):
    """کل نورد را می‌خواند. applicant=None یعنی «همه با هم»."""
    if not os.path.exists(VISA_PATH):
        return {"people": {}}
    try:
        with open(VISA_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {"people": {}}
    if not isinstance(data, dict) or "people" not in data:
        return {"people": {}}
    if applicant:
        return data["people"].get(applicant, {"stages": {}, "acts": [], "target": ""})
    return data


def save_journey(data):
    os.makedirs(MEM, exist_ok=True)
    with open(VISA_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return data


def _person(root, applicant):
    return root.setdefault("people", {}).setdefault(
        applicant, {"stages": {}, "acts": [], "target": ""})


def toggle_stage(applicant, key, done=None, note=""):
    """تیک‌زدن/برداشتن تیک یک مرحله. done=None یعنی برعکسِ وضعیت فعلی."""
    if key not in STAGE_KEYS:
        return None
    root = load_journey()
    p = _person(root, applicant)
    cur = p["stages"].get(key, {})
    is_done = (not cur.get("done")) if done is None else bool(done)
    entry = {"done": is_done, "date": _today() if is_done else "",
             "note": note or cur.get("note", ""), "updated": _now()}
    p["stages"][key] = entry
    save_journey(root)
    return entry


def set_note_stage(applicant, key, note):
    root = load_journey()
    p = _person(root, applicant)
    p["stages"].setdefault(key, {"done": False})["note"] = note
    p["stages"][key]["updated"] = _now()
    save_journey(root)
    return p["stages"][key]


def set_target(applicant, target):
    """کشور/ویزایی که این نورد برای آن است — مثلاً «فنلاند — ویزای تحصیلی»."""
    root = load_journey()
    _person(root, applicant)["target"] = target
    save_journey(root)
    return target


def add_act(applicant, category, title, detail="", when=None, done=True):
    """ثبت یک اکت در دفترچهٔ اقدامات.

    category یکی از ACT_CATS است. when تاریخ YYYY-MM-DD (پیش‌فرض امروز).
    """
    category = category if category in ACT_CATS else "other"
    root = load_journey()
    p = _person(root, applicant)
    act = {
        "id": len(p["acts"]) + 1,
        "category": category,
        "title": (title or "").strip()[:160],
        "detail": (detail or "").strip()[:1000],
        "date": when or _today(),
        "done": bool(done),
        "logged_at": _now(),
    }
    p["acts"].insert(0, act)
    # شماره‌ها را دوباره بده تا ترتیب نمایش با ترتیب ذخیره یکی بماند
    for i, a in enumerate(p["acts"], 1):
        a["id"] = i
    save_journey(root)
    return act


def toggle_act(applicant, act_id):
    root = load_journey()
    p = _person(root, applicant)
    for a in p["acts"]:
        if a.get("id") == int(act_id):
            a["done"] = not a.get("done")
            a["logged_at"] = _now()
            save_journey(root)
            return a
    return None


def delete_act(applicant, act_id):
    root = load_journey()
    p = _person(root, applicant)
    before = len(p["acts"])
    p["acts"] = [a for a in p["acts"] if a.get("id") != int(act_id)]
    for i, a in enumerate(p["acts"], 1):
        a["id"] = i
    save_journey(root)
    return before - len(p["acts"])


def clear_acts(applicant):
    root = load_journey()
    _person(root, applicant)["acts"] = []
    save_journey(root)


# ── محاسبهٔ گزارش ────────────────────────────────────────────────
def journey_stats(applicant):
    """درصد پیشرفت چک‌لیست + شمارش اکت‌ها به تفکیک دسته."""
    p = load_journey(applicant)
    stages = p.get("stages", {}) or {}
    total = len(VISA_STAGES)
    done = sum(1 for k, _t, _h, _c in VISA_STAGES if (stages.get(k) or {}).get("done"))
    per_cat = {}
    for cat in CAT_ORDER:
        keys = [k for k, _t, _h, c in VISA_STAGES if c == cat]
        d = sum(1 for k in keys if (stages.get(k) or {}).get("done"))
        per_cat[cat] = {"done": d, "total": len(keys)}
    acts = p.get("acts", []) or []
    per_act_cat = {}
    for c in ACT_CATS:
        n = sum(1 for a in acts if a.get("category") == c)
        if n:
            per_act_cat[c] = n
    open_acts = [a for a in acts if not a.get("done")]
    return {
        "target": p.get("target", ""),
        "stages_done": done, "stages_total": total,
        "pct": round(done * 100 / total) if total else 0,
        "per_cat": per_cat,
        "acts_total": len(acts),
        "acts_open": len(open_acts),
        "acts_per_cat": per_act_cat,
        "last_act": acts[0] if acts else None,
    }


def overdue_stages(applicant, days=45):
    """مرحله‌هایی که خیلی وقت است تیک نخورده‌اند — احتمالاً جا افتاده‌اند."""
    p = load_journey(applicant)
    stages = p.get("stages", {}) or {}
    out = []
    for k, title, hint, cat in VISA_STAGES:
        st = stages.get(k) or {}
        if st.get("done"):
            continue
        if st.get("updated"):
            try:
                age = (datetime.now() - datetime.strptime(st["updated"], "%Y-%m-%d %H:%M")).days
            except ValueError:
                continue
            if age >= days:
                out.append({"key": k, "title": title, "hint": hint, "cat": cat, "stale_days": age})
    return out


def deadline_soon(within_days=30):
    """اکت‌های بازی که تاریخشان نزدیک است یا گذشته — برای هشدار بالای صفحه."""
    today = datetime.now().date()
    out = []
    for applicant in load_journey().get("people", {}):
        for a in load_journey(applicant).get("acts", []) or []:
            if a.get("done"):
                continue
            ds = (a.get("date") or "")[:10]
            try:
                d = datetime.strptime(ds, "%Y-%m-%d").date()
            except ValueError:
                continue
            if (d - today).days <= within_days:
                out.append({"applicant": applicant, "act": a,
                            "days": (d - today).days})
    out.sort(key=lambda x: x["days"])
    return out


def render_report(applicant):
    """گزارش متنی نورد — برای فایل خروجی و گزارش لحظه‌ای."""
    p = load_journey(applicant)
    st = journey_stats(applicant)
    lines = [
        f"🛂 نورد ویزای همراه — {applicant}",
        f"تاریخ گزارش: {_now()}",
        f"هدف: {st['target'] or '—'}",
        f"پیشرفت چک‌لیست: {st['stages_done']}/{st['stages_total']} ({st['pct']}%)",
        "",
    ]
    for cat in CAT_ORDER:
        keys = [k for k, _t, _h, c in VISA_STAGES if c == cat]
        if not keys:
            continue
        lines.append(f"── {CAT_FA[cat]} ──")
        for k in keys:
            entry = (p.get("stages", {}) or {}).get(k) or {}
            mark = "☑" if entry.get("done") else "☐"
            when = f"  ({entry['date']})" if entry.get("date") else ""
            note = f"  — {entry['note']}" if entry.get("note") else ""
            lines.append(f"  {mark} {STAGE_TITLE[k]}{when}{note}")
        lines.append("")
    lines.append(f"── 📌 دفترچهٔ اکت‌ها ({len(p.get('acts', []) or [])} مورد) ──")
    for a in p.get("acts", []) or []:
        mark = "☑" if a.get("done") else "☐"
        lines.append(f"  {mark} [{a.get('date')}] {ACT_CAT_EMOJI.get(a.get('category'),'')} "
                     f"{a.get('title','')}")
        if a.get("detail"):
            lines.append(f"      {a['detail']}")
    return "\n".join(lines)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print(__doc__)
        print("نمونه: python visa_tracker.py tohid")
        raise SystemExit(0)
    print(render_report(sys.argv[1]))
