#!/usr/bin/env python3
"""
MigrationHunter — Email Analysis Excel
تبدیل تحلیل ایمیل به Excel با شیت‌های جداگانه + آمار درصدی
"""
import os, json
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, BarChart, Reference

BASE = os.path.dirname(os.path.abspath(__file__))
MEM = os.path.join(BASE, "memory")
DASH = os.path.join(BASE, "dashboard")
OUT = os.path.join(BASE, "output")

FONT_FA = "B Mitra"
FONT_EN = "Times New Roman"
NOW = datetime.now()
DATE_STR = NOW.strftime("%Y-%m-%d %H:%M")
FILE_DATE = NOW.strftime("%Y%m%d_%H%M")

# Colors
C_DARK = "1B4F72"; C_MED = "2E86C1"; C_LIGHT = "D6EAF8"
C_GREEN = "27AE60"; C_LGREEN = "D5F5E3"
C_YELLOW = "F39C12"; C_LYELLOW = "FEF9E7"
C_RED = "E74C3C"; C_LRED = "FADBD8"
C_PURPLE = "8E44AD"; C_LPURPLE = "E8DAEF"
C_GRAY = "95A5A6"; C_LGRAY = "F2F3F4"
C_WHITE = "FFFFFF"; C_DARK2 = "2C3E50"
C_ORANGE = "E67E22"; C_LORANGE = "FDEBD0"

thin = Border(left=Side("thin"), right=Side("thin"),
              top=Side("thin"), bottom=Side("thin"))

def rtl(ws):
    ws.sheet_view.rightToLeft = True

def fa(sz=10, bold=False, color="000000", italic=False):
    return Font(name=FONT_FA, size=sz, bold=bold, color=color, italic=italic)

def en(sz=10, bold=False, color="000000"):
    return Font(name=FONT_EN, size=sz, bold=bold, color=color)

def fill(c):
    return PatternFill(start_color=c, end_color=c, fill_type="solid")

def ac(h="right", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def wc(ws, r, c, val, font=None, bg=None, align=None, border=True):
    cell = ws.cell(row=r, column=c, value=val)
    if font: cell.font = font
    if bg: cell.fill = fill(bg)
    if align: cell.alignment = align
    else: cell.alignment = ac()
    if border: cell.border = thin
    return cell

def header_row(ws, row, cols, bg=C_DARK):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = fa(sz=9, bold=True, color=C_WHITE)
        cell.fill = fill(bg)
        cell.alignment = center()
        cell.border = thin

def auto_width(ws, col, w):
    ws.column_dimensions[get_column_letter(col)].width = w

def freeze(ws, cell="A2"):
    ws.freeze_panes = cell

# Category info
CATEGORIES = {
    "interview": {"label": "🗣️ مصاحبه", "color": C_GREEN, "bg": C_LGREEN, "emoji": "🗣️"},
    "offer": {"label": "🎉 پیشنهاد کار", "color": C_PURPLE, "bg": C_LPURPLE, "emoji": "🎉"},
    "rejection": {"label": "❌ رد شده", "color": C_RED, "bg": C_LRED, "emoji": "❌"},
    "follow_up": {"label": "⏰ پیگیری", "color": C_YELLOW, "bg": C_LYELLOW, "emoji": "⏰"},
    "inquiry": {"label": "💬 استعلام", "color": C_MED, "bg": C_LIGHT, "emoji": "💬"},
    "acknowledgment": {"label": "📩 تأیید دریافت", "color": C_GRAY, "bg": C_LGRAY, "emoji": "📩"},
    "unknown": {"label": "❓ نامشخص", "color": C_DARK2, "bg": C_LGRAY, "emoji": "❓"},
}

def load_email_data():
    filepath = os.path.join(MEM, "EMAIL_ANALYSIS.json")
    if not os.path.exists(filepath):
        print("❌ فایل EMAIL_ANALYSIS.json پیدا نشد")
        print("   ابتدا: python email_analyzer.py")
        return None
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)

def build_email_excel(data):
    wb = Workbook()
    wb.remove(wb.active)
    
    emails = data.get("emails", [])
    total = data.get("total_emails", 0)
    job_related = data.get("job_related", 0)
    per_account = data.get("per_account", [])
    
    # Categorize
    by_category = {}
    for e in emails:
        cat = e.get("category", "unknown")
        if cat not in by_category:
            by_category[cat] = []
        by_category[cat].append(e)
    
    # By applicant
    by_applicant = {"NEDA": [], "TOHID": [], "UNKNOWN": []}
    for e in emails:
        app = e.get("applicant", "UNKNOWN")
        if app not in by_applicant:
            by_applicant[app] = []
        by_applicant[app].append(e)
    
    # ═══════════════════════════════════════
    # Sheet 01: Dashboard with percentages
    # ═══════════════════════════════════════
    ws = wb.create_sheet("داشبورد ایمیل")
    rtl(ws)
    
    wc(ws, 1, 1, f"تحلیل ایمیل شغلی — {DATE_STR}", font=fa(sz=16, bold=True, color=C_DARK))
    ws.merge_cells("A1:H1")
    
    # KPI
    wc(ws, 3, 1, "کل ایمیل‌ها", font=fa(sz=9, bold=True, color=C_WHITE), bg=C_DARK, align=center())
    wc(ws, 3, 2, total, font=fa(sz=18, bold=True, color=C_DARK), align=center())
    wc(ws, 3, 3, "مرتبط با کار", font=fa(sz=9, bold=True, color=C_WHITE), bg=C_GREEN, align=center())
    pct_job = round(job_related / total * 100) if total else 0
    wc(ws, 3, 4, f"{job_related} ({pct_job}%)", font=fa(sz=18, bold=True, color=C_GREEN), align=center())
    
    # Category breakdown with percentages
    row = 6
    wc(ws, row, 1, "دسته‌بندی ایمیل‌ها", font=fa(sz=12, bold=True, color=C_DARK))
    ws.merge_cells(f"A{row}:H{row}")
    row += 1
    
    cat_headers = ["دسته", "تعداد", "درصد", "نمودار"]
    for i, h in enumerate(cat_headers):
        wc(ws, row, i + 1, h, font=fa(sz=9, bold=True, color=C_WHITE), bg=C_DARK, align=center())
    row += 1
    
    chart_start_row = row
    for cat_key in ["interview", "offer", "rejection", "follow_up", "inquiry", "acknowledgment", "unknown"]:
        info = CATEGORIES.get(cat_key, CATEGORIES["unknown"])
        count = len(by_category.get(cat_key, []))
        pct = round(count / job_related * 100) if job_related else 0
        
        # Bar visualization
        bar = "█" * max(1, round(pct / 5)) if pct > 0 else ""
        
        wc(ws, row, 1, info["label"], font=fa(sz=10, bold=True), bg=info["bg"], align=center())
        wc(ws, row, 2, count, font=fa(sz=12, bold=True), align=center())
        wc(ws, row, 3, f"{pct}%", font=fa(sz=12, bold=True, color=info["color"]), align=center())
        wc(ws, row, 4, bar, font=fa(sz=10, color=info["color"]), align=ac(h="left"))
        row += 1
    
    # Pie chart
    try:
        pie = PieChart()
        pie.title = "دسته‌بندی ایمیل‌ها"
        pie.style = 10
        data_ref = Reference(ws, min_col=2, min_row=chart_start_row - 1, 
                           max_row=row - 1)
        cats_ref = Reference(ws, min_col=1, min_row=chart_start_row, 
                           max_row=row - 1)
        pie.add_data(data_ref, titles_from_data=True)
        pie.set_categories(cats_ref)
        pie.width = 18
        pie.height = 12
        ws.add_chart(pie, f"F{chart_start_row}")
    except:
        pass
    
    # Applicant breakdown
    row += 1
    wc(ws, row, 1, "بر اساس متقاضی", font=fa(sz=12, bold=True, color=C_DARK))
    ws.merge_cells(f"A{row}:H{row}")
    row += 1
    
    for app_key, app_label in [("TOHID", "👨 توحید"), ("NEDA", "👩 ندا"), ("UNKNOWN", "❓ نامشخص")]:
        count = len(by_applicant.get(app_key, []))
        pct = round(count / job_related * 100) if job_related else 0
        wc(ws, row, 1, app_label, font=fa(sz=10, bold=True), align=center())
        wc(ws, row, 2, count, font=fa(sz=12, bold=True), align=center())
        wc(ws, row, 3, f"{pct}%", font=fa(sz=12, bold=True, color=C_MED), align=center())
        row += 1
    
    # Top senders
    row += 1
    wc(ws, row, 1, "فرستنده‌های پرتکرار", font=fa(sz=12, bold=True, color=C_DARK))
    ws.merge_cells(f"A{row}:H{row}")
    row += 1
    
    senders = {}
    for e in emails:
        sender = e.get("from", "").split("<")[0].strip().strip('"')
        if sender:
            senders[sender] = senders.get(sender, 0) + 1
    
    for sender, count in sorted(senders.items(), key=lambda x: -x[1])[:10]:
        pct = round(count / job_related * 100) if job_related else 0
        wc(ws, row, 1, sender[:50], font=fa(sz=9))
        wc(ws, row, 2, count, font=fa(sz=10, bold=True), align=center())
        wc(ws, row, 3, f"{pct}%", font=fa(sz=10, color=C_MED), align=center())
        row += 1
    
    # Per-account breakdown
    if per_account:
        row += 1
        wc(ws, row, 1, "بر اساس حساب ایمیل", font=fa(sz=12, bold=True, color=C_DARK))
        ws.merge_cells(f"A{row}:H{row}")
        row += 1
        wc(ws, row, 1, "حساب", font=fa(sz=9, bold=True, color=C_WHITE), bg=C_DARK, align=center())
        wc(ws, row, 2, "شخص", font=fa(sz=9, bold=True, color=C_WHITE), bg=C_DARK, align=center())
        wc(ws, row, 3, "LinkedIn", font=fa(sz=9, bold=True, color=C_WHITE), bg=C_DARK, align=center())
        wc(ws, row, 4, "ایمیل شغلی", font=fa(sz=9, bold=True, color=C_WHITE), bg=C_DARK, align=center())
        wc(ws, row, 5, "درصد", font=fa(sz=9, bold=True, color=C_WHITE), bg=C_DARK, align=center())
        row += 1
        for pa in per_account:
            pct = round(pa.get("job_related", 0) / job_related * 100) if job_related else 0
            wc(ws, row, 1, pa.get("email", ""), font=en(sz=9))
            person = pa.get("person", "?")
            app_label = "👩 ندا" if person == "NEDA" else "👨 توحید" if person == "TOHID" else person
            wc(ws, row, 2, app_label, font=fa(sz=9), align=center())
            wc(ws, row, 3, pa.get("linkedin", ""), font=en(sz=8, color="0563C1"))
            wc(ws, row, 4, pa.get("job_related", 0), font=fa(sz=10, bold=True), align=center())
            wc(ws, row, 5, f"{pct}%", font=fa(sz=10, bold=True, color=C_MED), align=center())
            row += 1
    
    # Widths
    for i, w in enumerate([30, 10, 10, 40, 10, 10, 10, 10]):
        auto_width(ws, i + 1, w)
    freeze(ws, "A2")
    
    # ═══════════════════════════════════════
    # Sheet 02: All job emails
    # ═══════════════════════════════════════
    ws2 = wb.create_sheet("تمام ایمیل‌ها")
    rtl(ws2)
    
    wc(ws2, 1, 1, f"تمام ایمیل‌های شغلی — {job_related} ایمیل", font=fa(sz=14, bold=True, color=C_DARK))
    ws2.merge_cells("A1:G1")
    
    headers = ["#", "تاریخ", "از", "موضوع", "دسته", "متقاضی", "کارفرما", "حساب"]
    header_row(ws2, 3, len(headers))
    for i, h in enumerate(headers):
        ws2.cell(row=3, column=i+1).value = h
    
    row = 4
    for idx, e in enumerate(sorted(emails, key=lambda x: x.get("date", ""), reverse=True), 1):
        cat = e.get("category", "unknown")
        info = CATEGORIES.get(cat, CATEGORIES["unknown"])
        applicant = e.get("applicant", "?")
        app_label = "👩 ندا" if applicant == "NEDA" else "👨 توحید" if applicant == "TOHID" else "❓"
        
        vals = [idx, e.get("date",""), e.get("from","")[:50], e.get("subject","")[:70],
                info["label"], app_label, e.get("employer",""), e.get("account_id","")]
        
        for ci, v in enumerate(vals):
            bg = info["bg"] if ci == 4 else None
            wc(ws2, row, ci + 1, v, font=fa(sz=9), bg=bg)
        row += 1
    
    widths = [5, 18, 40, 55, 16, 10, 20, 12]
    for i, w in enumerate(widths): auto_width(ws2, i + 1, w)
    freeze(ws2, "A4")
    ws2.auto_filter.ref = f"A3:H{row-1}"
    
    # ═══════════════════════════════════════
    # Sheets by category
    # ═══════════════════════════════════════
    for cat_key in ["interview", "offer", "rejection", "follow_up", "inquiry"]:
        cat_emails = by_category.get(cat_key, [])
        if not cat_emails:
            continue
        
        info = CATEGORIES[cat_key]
        ws_cat = wb.create_sheet(info["label"].split(" ")[-1] if " " in info["label"] else cat_key)
        rtl(ws_cat)
        
        wc(ws_cat, 1, 1, f"{info['label']} — {len(cat_emails)} ایمیل", 
           font=fa(sz=14, bold=True, color=info["color"]))
        ws_cat.merge_cells("A1:G1")
        
        cat_headers = ["#", "تاریخ", "از", "موضوع", "متقاضی", "کارفرما", "اقدام"]
        header_row(ws_cat, 3, len(cat_headers), bg=info["color"])
        for i, h in enumerate(cat_headers):
            ws_cat.cell(row=3, column=i+1).value = h
        
        row = 4
        for idx, e in enumerate(sorted(cat_emails, key=lambda x: x.get("date", ""), reverse=True), 1):
            applicant = e.get("applicant", "?")
            app_label = "👩 ندا" if applicant == "NEDA" else "👨 توحید" if applicant == "TOHID" else "❓"
            
            action = ""
            if cat_key == "interview": action = "پاسخ + حضور"
            elif cat_key == "offer": action = "بررسی + تصمیم"
            elif cat_key == "rejection": action = "بایگانی"
            elif cat_key == "follow_up": action = "پیگیری"
            elif cat_key == "inquiry": action = "بررسی"
            
            vals = [idx, e.get("date",""), e.get("from","")[:50], e.get("subject","")[:70],
                    app_label, e.get("employer",""), action]
            
            for ci, v in enumerate(vals):
                wc(ws_cat, row, ci + 1, v, font=fa(sz=9))
            row += 1
        
        widths = [5, 18, 40, 55, 10, 20, 16]
        for i, w in enumerate(widths): auto_width(ws_cat, i + 1, w)
        freeze(ws_cat, "A4")
    
    # ═══════════════════════════════════════
    # Sheet: By Applicant
    # ═══════════════════════════════════════
    for app_key, app_label, app_color in [("TOHID", "توحید — IT", C_MED), ("NEDA", "ندا — مامایی", C_PURPLE)]:
        app_emails = by_applicant.get(app_key, [])
        if not app_emails:
            continue
        
        ws_app = wb.create_sheet(app_label)
        rtl(ws_app)
        
        wc(ws_app, 1, 1, f"{app_label} — {len(app_emails)} ایمیل", 
           font=fa(sz=14, bold=True, color=app_color))
        ws_app.merge_cells("A1:G1")
        
        # Mini dashboard for this applicant
        app_cats = {}
        for e in app_emails:
            cat = e.get("category", "unknown")
            app_cats[cat] = app_cats.get(cat, 0) + 1
        
        row = 3
        wc(ws_app, row, 1, "دسته", font=fa(sz=9, bold=True, color=C_WHITE), bg=app_color, align=center())
        wc(ws_app, row, 2, "تعداد", font=fa(sz=9, bold=True, color=C_WHITE), bg=app_color, align=center())
        wc(ws_app, row, 3, "درصد", font=fa(sz=9, bold=True, color=C_WHITE), bg=app_color, align=center())
        row += 1
        
        for cat_key in ["interview", "offer", "rejection", "follow_up", "inquiry"]:
            count = app_cats.get(cat_key, 0)
            pct = round(count / len(app_emails) * 100) if app_emails else 0
            info = CATEGORIES.get(cat_key, CATEGORIES["unknown"])
            wc(ws_app, row, 1, info["label"], font=fa(sz=9), bg=info["bg"], align=center())
            wc(ws_app, row, 2, count, font=fa(sz=10, bold=True), align=center())
            wc(ws_app, row, 3, f"{pct}%", font=fa(sz=10, bold=True, color=info["color"]), align=center())
            row += 1
        
        row += 1
        app_headers = ["#", "تاریخ", "از", "موضوع", "دسته", "کارفرما", "اقدام"]
        header_row(ws_app, row, len(app_headers), bg=app_color)
        for i, h in enumerate(app_headers):
            ws_app.cell(row=row, column=i+1).value = h
        row += 1
        
        for idx, e in enumerate(sorted(app_emails, key=lambda x: x.get("date", ""), reverse=True), 1):
            cat = e.get("category", "unknown")
            info = CATEGORIES.get(cat, CATEGORIES["unknown"])
            vals = [idx, e.get("date",""), e.get("from","")[:50], e.get("subject","")[:70],
                    info["label"], e.get("employer",""), ""]
            for ci, v in enumerate(vals):
                bg = info["bg"] if ci == 4 else None
                wc(ws_app, row, ci + 1, v, font=fa(sz=9), bg=bg)
            row += 1
        
        widths = [5, 18, 40, 55, 16, 20, 16]
        for i, w in enumerate(widths): auto_width(ws_app, i + 1, w)
        freeze(ws_app, "A4")
    
    return wb

def main():
    print("=" * 60)
    print("MigrationHunter — Email Analysis Excel")
    print("=" * 60)
    
    data = load_email_data()
    if not data:
        return
    
    print(f"\n📂 داده: {data.get('job_related', 0)} ایمیل شغلی از {data.get('total_emails', 0)} کل")
    
    print("\n📊 ساخت Excel...")
    wb = build_email_excel(data)
    
    os.makedirs(DASH, exist_ok=True)
    filename = f"Email_Analysis_{FILE_DATE}.xlsx"
    filepath = os.path.join(DASH, filename)
    wb.save(filepath)
    
    print(f"  ✅ ذخیره شد: {filename}")
    print(f"  📋 شیت‌ها: {len(wb.sheetnames)}")
    for s in wb.sheetnames:
        print(f"    - {s}")
    
    print(f"\n{'=' * 60}")
    print(f"📊 خلاصه")
    print(f"{'=' * 60}")
    print(f"  📧 ایمیل شغلی: {data.get('job_related', 0)}")
    print(f"  📋 شیت‌ها: {len(wb.sheetnames)}")
    print(f"  📁 فایل: {filename}")

if __name__ == "__main__":
    main()
