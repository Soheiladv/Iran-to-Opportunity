#!/usr/bin/env python3
"""
MigrationHunter — اعتبارسنجی بانک منابع (sources.json)

هر URL را یک‌بار می‌گیرد و گزارش می‌دهد: چه کدی برگشت، چند بایت HTML بود،
و آیا اصلاً آگهی/برنامه‌ای داخلش پیدا شد یا نه.

هدف: بانک منابع را بدون حدس‌زدن آب نگه داریم — منبعی که ۴۰۴ می‌دهد یا
هیچ چیزی نمی‌دهد، باید خاموش شود (enabled=false) تا زمان اجرا هدر نرود.

نکتهٔ مهم: گرفتن صفحه از راه fetch_one انجام می‌شود، نه _fetch_plain تنها.
پس اگر منبعی needs_js داشته باشد (یا plain بلاک شود) با مرورگر headless
امتحان می‌شود و نتیجهٔ واقعی را می‌بینیم، نه یک «مردهٔ» ساختگی.

اجرا:
    python scripts/validate_sources.py                    # همهٔ منابع
    python scripts/validate_sources.py --track job        # فقط کاریابی
    python scripts/validate_sources.py --track education # فقط تحصیل
    python scripts/validate_sources.py --country FI      # فقط فنلاند
    python scripts/validate_sources.py --homepage        # فقط صفحهٔ اصلی (سریع)
    python scripts/validate_sources.py --no-browser      # بدون مرورگر (سریع ولی گمراه‌کننده برای JS)
    python scripts/validate_sources.py --timeout 8 --workers 6 --save
"""
import argparse
import json
import os
import sys
from concurrent.futures import ThreadPoolExecutor

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import job_crawler as jc  # noqa: E402

SOURCES_PATH = jc.SOURCES_PATH
MEM = jc.MEM
REPORT_PATH = os.path.join(MEM, "SOURCE_VALIDATION.json")

# وضعیت‌ها — چون نتیجهٔ "مرده" و "جواب نداد" فرق دارد، سه حالت داریم
OK = "OK"                 # صفحه آمد و HTML معنادار داشت
EMPTY = "EMPTY"           # صفحه آمد ولی هیچ آگهی/برنامه‌ای داخلش نبود
BLOCKED = "BLOCKED/DEAD"  # اصلاً صفحه‌ای نیامد (403 / تایم‌اوت / خطا)


def probe(entry, use_browser=True, timeout=12):
    """یک URL را می‌گیرد و نتیجهٔ کوتاه برمی‌گرداند.

    entry = (url, source_dict) — کل منبع را می‌خواهیم چون fetch_one و
    استخراج‌گر به فیلدهای source (مثل track) نیاز دارند.
    """
    url, source = entry
    name = source.get("name", "?")
    # fetch_one خودش plain → headless را به‌ترتیب امتحان می‌کند و نیاز به
    # timeout را هم رعایت می‌کند (به‌جای فراخوانی مستقیم _fetch_plain).
    source_for_fetch = dict(source)
    source_for_fetch["url"] = url
    if not use_browser or not jc.browser_available():
        source_for_fetch["needs_js"] = False
    html = jc.fetch_one(source_for_fetch, url, timeout=timeout)

    result = {"url": url, "source": name, "country": source.get("country", ""),
              "track": source.get("track", "job"), "bytes": 0, "found": 0,
              "status": BLOCKED}

    if not html:
        result["status"] = BLOCKED
        return result

    result["bytes"] = len(html)
    found = 0
    try:
        if source.get("track", "job") == "education":
            import education_crawler as ec
            found = len(ec.extract_programs_from_html(html, source))
        else:
            found = len(jc.extract_jobs_from_html(html, source))
    except Exception:
        found = 0
    result["found"] = found
    result["status"] = OK if found > 0 else EMPTY
    return result


def main():
    ap = argparse.ArgumentParser(description="اعتبارسنجی بانک منابع")
    ap.add_argument("--track", choices=["job", "education"], help="فقط یک مسیر")
    ap.add_argument("--country", help="فقط یک کشور (کد دو حرفی، مثلاً FI)")
    ap.add_argument("--homepage", action="store_true", help="فقط URL صفحهٔ اصلی هر منبع")
    ap.add_argument("--timeout", type=int, default=12, help="ثانیه برای هر آدرس")
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--no-browser", action="store_true",
                    help="فقط درخواست ساده — سریع‌تر، ولی منابع JS را مرده نشان می‌دهد")
    ap.add_argument("--save", action="store_true", help="ذخیرهٔ گزارش در memory/SOURCE_VALIDATION.json")
    args = ap.parse_args()

    sources = jc.load_sources(include_disabled=True)
    if args.track:
        sources = [s for s in sources if s.get("track", "job") == args.track]
    if args.country:
        cc = args.country.upper()
        sources = [s for s in sources if s.get("country") == cc]
    if not sources:
        print("هیچ منبعی با این فیلترها پیدا نشد.")
        return 1

    targets = []
    for s in sources:
        urls = [s["url"]] if args.homepage else (s.get("search_urls") or [s["url"]])
        for u in urls:
            targets.append((u, s))

    use_browser = not args.no_browser
    mode = "مروگر headless + plain" if (use_browser and jc.HAS_PLAYWRIGHT) else "فقط plain"
    print(f"🔍 بررسی {len(targets)} آدرس از {len(sources)} منبع "
          f"(حداکثر {args.workers} هم‌زمان، timeout {args.timeout}s، حالت: {mode})…\n")

    # ⚠️ نکتهٔ مهم: Playwright همزمان (sync) به یک thread قفل می‌شود — با
    # ThreadPool «Cannot switch to a different thread» می‌دهد. پس وقتی مرورگر
    # در کار است، همه را تک‌thread می‌زنیم (کندتر ولی درست). بدون مرورگر
    # می‌شود موازی کرد چون فقط urllib است.
    parallel = not (use_browser and jc.HAS_PLAYWRIGHT)
    if not parallel:
        args.workers = 1
        mode += " (تک‌thread — Playwright قابل موازی‌سازی نیست)"

    if parallel:
        with ThreadPoolExecutor(max_workers=args.workers) as pool:
            results = list(pool.map(
                lambda t: probe(t, use_browser=use_browser, timeout=args.timeout), targets))
    else:
        results = []
        for i, t in enumerate(targets, 1):
            r = probe(t, use_browser=use_browser, timeout=args.timeout)
            results.append(r)
            mark = {"OK": "✅", "EMPTY": "⚪"}.get(r["status"], "❌")
            print(f"  {mark} [{i}/{len(targets)}] {r['source'][:38]} · "
                  f"{r['found']} یافته · {r['bytes']//1024}KB")

    # ── خلاصه به تفکیک منبع ──
    by_source = {}
    for r in results:
        st = by_source.setdefault(r["source"], {"ok": 0, "empty": 0, "dead": 0,
                                                "found": 0, "bytes": 0})
        if r["status"] == OK:
            st["ok"] += 1
        elif r["status"] == EMPTY:
            st["empty"] += 1
        else:
            st["dead"] += 1
        st["found"] += r["found"]
        st["bytes"] += r["bytes"]

    print(f"{'منبع':<44}{'URL':>5}{'زنده':>6}{'خالی':>6}{'مرده':>6}{'یافته':>7}{'KB':>7}")
    print("─" * 81)
    weak, dead_all = [], []
    for s in sources:
        st = by_source.get(s["name"], {"ok": 0, "empty": 0, "dead": 0, "found": 0, "bytes": 0})
        flag = ""
        if st["ok"] == 0 and st["empty"] == 0:
            flag = "  ❌ همه بلاک/مرده"
            dead_all.append(s["name"])
        elif st["ok"] == 0:
            flag = "  ⚠️ زنده ولی بی‌برنامه"
            weak.append(s["name"])
        n_urls = len([s["url"]] if args.homepage else (s.get("search_urls") or [s["url"]]))
        print(f"{s['name'][:42]:<44}{n_urls:>5}{st['ok']:>6}{st['empty']:>6}"
              f"{st['dead']:>6}{st['found']:>7}{st['bytes']//1024:>7}{flag}")

    total_ok = sum(1 for r in results if r["status"] == OK)
    total_empty = sum(1 for r in results if r["status"] == EMPTY)
    total_found = sum(r["found"] for r in results)
    print("─" * 81)
    print(f"مجموع: {total_ok} آدرس با محتوا · {total_empty} خالی · "
          f"{len(results) - total_ok - total_empty} بلاک/مرده · {total_found} یافته")
    if dead_all:
        print(f"\n❌ منابعی که هیچ URL زنده‌ای ندارند ({len(dead_all)}) — enabled=false کن:")
        for n in dead_all:
            print(f"   - {n}")
    if weak:
        print(f"\n⚠️ منابع زنده ولی بدون نتیجه ({len(weak)}) — آدرس جستجو شاید عوض شده "
              f"یا نیاز به صفحه‌بندی/لاگین دارند:")
        for n in weak:
            print(f"   - {n}")
    if args.no_browser:
        print("\nℹ️ با --no-browser منابع needs_js حتماً «مرده» دیده می‌شوند — "
              "برای قضاوت واقعی بدون این فلگ اجرا کن.")

    if args.save:
        os.makedirs(MEM, exist_ok=True)
        with open(REPORT_PATH, "w", encoding="utf-8") as f:
            json.dump({"results": results, "by_source": by_source}, f, ensure_ascii=False, indent=1)
        print(f"\n💾 گزارش کامل ذخیره شد: {REPORT_PATH}")
    return 0


if __name__ == "__main__":
    sys.exit(main())