# MigrationHunter — Excel Forensic Audit

**تاریخ اجرا:** 2026-08-27 17:00
**وضعیت:** READ-ONLY AUDIT — NO MODIFICATIONS

---

## 1. Executive Summary

| مورد | مقدار |
|------|-------|
| کل فایل‌ها | 140 |
| Python Scripts | 25 |
| فایل‌های Excel | 10 (1 فعال + 1 تکراری + 8 آرشیو) |
| Sheetهای فعال | 3 (از 10+ مورد نیاز) |
| ردیف‌های داده فعال | 22 |
| طراحی | NEEDS REDESIGN |
| کیفیت داده | LOW |
| Evidence | PARTIAL |
| خودکارسازی | LOW |

**نتیجه نهایی:**

```
CURRENT EXCEL:    REBUILD
DESIGN:           35/100
DATA QUALITY:     25/100
EVIDENCE:         40/100
ARCHITECTURE:     20/100
AUTOMATION:       15/100
OVERALL:          27/100
```

---

## 2. Repository Inventory

### فایل‌ها بر اساس نوع

| نوع | تعداد | مسیر |
|-----|-------|------|
| `.xlsx` | 10 | `dashboard/`, `dashboard/archive/`, `dashboard/dashboard/` |
| `.py` | 25 | ریشه پروژه |
| `.md` | 30+ | `memory/`, `output/`, `profiles/`, ریشه |
| `.json` | 3 | `memory/` |
| `.txt` | 9 | `output/emails/` |
| `.env` | 1 | ریشه |
| `.gitignore` | 1 | ریشه |
| **مجموع** | **140** | |

---

## 3. Excel Inventory

### فایل‌های Excel

| # | فایل | مسیر | وضعیت | Sheets | Rows |
|---|------|------|-------|--------|------|
| 1 | `MigrationHunter_Full_20260827_1440.xlsx` | `dashboard/` | **ACTIVE** | 3 | 22 |
| 2 | `MigrationHunter_Full_20260827_1443.xlsx` | `dashboard/dashboard/` | **DUPLICATE (nested)** | 3 | 22 |
| 3 | `MigrationHunter_Complete_20260827_1402.xlsx` | `archive/` | ARCHIVED | 8 | 121 |
| 4 | `MigrationHunter_v2_20260827_1336.xlsx` | `archive/` | ARCHIVED | 8 | 82 |
| 5 | `MigrationHunter_v2_20260827_1413.xlsx` | `archive/` | ARCHIVED | 8 | 82 |
| 6 | `MigrationHunter_Full_20260827_1433.xlsx` | `archive/` | ARCHIVED | 3 | 56 |
| 7 | `Europe_Complete_20260827_1420.xlsx` | `archive/` | ARCHIVED | 5 | 111 |
| 8 | `Saskatchewan_20260827_1336.xlsx` | `archive/` | ARCHIVED | ? | ? |
| 9 | `Saskatchewan_20260827_1413.xlsx` | `archive/` | ARCHIVED | ? | ? |
| 10 | `Search_Methods_20260827_1348.xlsx` | `archive/` | ARCHIVED | 6 | ? |

### مشکل ساختاری

```
dashboard/
├── MigrationHunter_Full_20260827_1440.xlsx     ← فعال
├── ~$MigrationHunter_Full_20260827_1440.xlsx   ← Lock file (Excel باز بوده)
├── dashboard/                                   ← پوشه تکراری اشتباهی!
│   └── MigrationHunter_Full_20260827_1443.xlsx  ← کپی تکراری
└── archive/                                     ← 8 فایل آرشیو
```

**⚠️ مشکل:** پوشه `dashboard/dashboard/` اشتباه ایجاد شده و فایل تکراری در آن قرار دارد.

---

## 4. Active Excel — Workbook Forensics

### فایل فعال: `MigrationHunter_Full_20260827_1440.xlsx`

#### Sheet 1: داشبورد اصلی

| ویژگی | مقدار |
|-------|-------|
| Rows | 22 |
| Columns | 22 |
| RTL | ✅ True |
| Font | B Mitra 16pt |
| Freeze | ❌ None |
| Filter | ❌ None |
| Hyperlinks | 0 |
| Charts | 0 |
| Conditional Formatting | 0 |
| Data Validation | 0 |
| Merged Cells | Yes (title row) |

#### Sheet 2: مقایسه کشورها

| ویژگی | مقدار |
|-------|-------|
| Rows | 14 |
| Columns | 8 |
| RTL | ✅ True |
| Font | B Mitra 14pt |

#### Sheet 3: ایمیل‌های معتبر

| ویژگی | مقدار |
|-------|-------|
| Rows | 20 |
| Columns | 4 |
| RTL | ✅ True |
| Font | B Mitra 14pt |

### خلاصه Active Excel

```
Total Sheets:    3
Total Rows:      56
Total Cols:      22 (max)
RTL:             ✅ All sheets
Font:            ✅ B Mitra
Freeze:          ❌ None
Filter:          ❌ None
Charts:          ❌ None
Hyperlinks:      ❌ None
Conditional:     ❌ None
```

---

## 5. Comparison: Active vs Archived

### Archive #3 — بهترین نسخه قبلی

`MigrationHunter_Complete_20260827_1402.xlsx` — 8 sheets, 121 rows

| Sheet | Rows | محتوا |
|-------|------|-------|
| داشبورد \| Dashboard | 9 | KPI خلاصه |
| ایمیل‌ها \| Emails | 8 | ایمیل‌های آماده |
| متن ایمیل \| Email Text | 16 | متن کامل ایمیل |
| نقاط ضعف ندا | 13 | LinkedIn weaknesses |
| نقاط ضعف توحید | 13 | LinkedIn weaknesses |
| لینک‌های جستجو | 42 | 39 لینک فعال |
| مقایسه کشورها | 7 | 4 کشور |
| اقدامات \| Actions | 13 | 10 اقدام |

### Archive #4 — بهترین ساختار

`MigrationHunter_v2_20260827_1336.xlsx` — 8 sheets, 82 rows

| Sheet | Rows | محتوا |
|-------|------|-------|
| Dashboard \| داشبورد | 13 | KPI cards |
| Top 10 \| ۱۰ برتر | 11 | Top opportunities |
| Emails & Links | 11 | ایمیل + لینک |
| Neda \| ندا | 10 | فرصت‌های ندا |
| Tohid \| توحید | 4 | فرصت‌های توحید |
| Countries \| کشورها | 7 | مقایسه |
| Weaknesses \| نقاط ضعف | 13 | LinkedIn |
| Actions \| اقدامات | 13 | اقدامات |

### Archive #7 — اروپا

`Europe_Complete_20260827_1420.xlsx` — 5 sheets, 111 rows

| Sheet | Rows | محتوا |
|-------|------|-------|
| مقایسه کشورها | 14 | 11 کشور |
| لینک‌های جستجو | 72 | 69 لینک |
| اتریش | 13 | راهنمای RWR |
| اسکاندیناوی | 7 | 4 کشور |
| هلند+ایرلند | 5 | مقایسه |

### ⚠️ نتیجه مقایسه

```
Active Excel:   3 sheets,  56 rows, ۳ شیت
Best Archive:   8 sheets, 121 rows, ۸ شیت
Europe Archive: 5 sheets, 111 rows, ۵ شیت

Active < Best Archive  ← نسخه فعال از آرشیو ضعیف‌تر است!
```

**نسخه فعال اطلاعات کمتری نسبت به نسخه‌های آرشیو شده دارد.**

---

## 6. Python Scripts — 25 اسکریپت

### دسته‌بندی

| دسته | تعداد | اسکریپت‌ها |
|------|-------|-----------|
| 📝 Writes Excel + Hardcoded | 8 | `run_full_dashboard.py`, `create_comprehensive_dashboard.py`, `create_dashboard_v2.py`, `create_dashboard_v3.py`, `create_dashboard_run3.py`, `run.py`, `run_search.py`, `run_dashboard_v2.py` |
| 📝 Writes Excel + Reads Files | 4 | `run_all.py`, `run_complete_cycle.py`, `email_tracker.py`, `add_application_status.py` |
| 📝 Writes Excel only | 7 | `create_dashboard.py`, `create_dashboard_final.py`, `create_final_excel.py`, `create_final_v2_complete.py`, `create_saskatchewan_excel.py`, `create_search_methods_excel.py`, `run_full_cycle.py`, `run_europe_complete.py` |
| 🔒 Hardcoded only | 1 | `linkedin_crawler.py`, `run_auto.py` |
| 📂 Reads Files only | 1 | `reminder.py` |
| 🔧 Utility | 1 | `run_cycle3.py` |

### ⚠️ مشکل اصلی

**8 اسکریپت داده Hardcoded دارند** — یعنی تمام فرصت‌ها، ایمیل‌ها و اطلاعات مستقیماً در کد نوشته شده‌اند و از فایل‌های memory/ خوانده نمی‌شوند.

### Current Generator

```
CURRENT EXCEL GENERATOR = run_full_dashboard.py
OUTPUT = dashboard/MigrationHunter_Full_YYYYMMDD_HHMM.xlsx
ARCHIVE = dashboard/archive/
```

### Read-Only Scripts (هیچ‌کدام memory/ را به‌روزرسانی نمی‌کنند)

هیچ اسکریپتی فایل‌های memory/*.md را خواندنی و نوشتني ندارد — فقط `run_all.py` و `run_complete_cycle.py` memory/ را write می‌کنند ولی **overwrite** می‌کنند، نه **merge**.

---

## 7. Data Lineage

### ستون‌های Excel از کجا می‌آیند؟

| ستون | Source | Reliable? |
|------|--------|-----------|
| متقاضی | Hardcoded در Python | ⚠️ Manual |
| کشور | Hardcoded در Python | ⚠️ Manual |
| کارفرما | Hardcoded در Python | ⚠️ Manual |
| عنوان شغل | Hardcoded در Python | ⚠️ Manual |
| لینک آگهی | Hardcoded در Python | ⚠️ Manual |
| ایمیل | Hardcoded — verified dict | ⚠️ Semi-verified |
| امتیاز | Hardcoded | ❌ Fake |
| حمایت مالی | Hardcoded | ⚠️ Semi-verified |
| ویزا | Hardcoded | ⚠️ Semi-verified |
| زبان | Hardcoded | ⚠️ Semi-verified |

**هیچ ستونی از memory/*.md خوانده نمی‌شود.**

---

## 8. Data Model Audit

### Entityهای مورد نیاز vs موجود

| Entity | در Excel | در Memory | در Both |
|--------|----------|-----------|---------|
| APPLICANT | ✅ | ✅ profiles/ | ✅ |
| EMPLOYER | ✅ hardcoded | ✅ EMPLOYER_BANK.md | ⚠️ disconnected |
| JOB | ✅ hardcoded | ✅ JOB_BANK.md | ⚠️ disconnected |
| EVIDENCE | ❌ | ✅ EVIDENCE_REGISTRY.md | ❌ |
| CONTACT | ✅ emails | ❌ | ⚠️ partial |
| VISA | ❌ | ✅ VISA_BANK.md | ❌ |
| APPLICATION | ❌ | ✅ APPLICATION_BANK.md | ❌ |
| FOLLOW_UP | ❌ | ❌ | ❌ |
| INTERVIEW | ❌ | ❌ | ❌ |
| OFFER | ❌ | ❌ | ❌ |
| SEARCH_HISTORY | ❌ | ✅ SEARCH_HISTORY.md | ❌ |

**نتیجه:** Excel و Memory Bankها از هم جدا هستند و اطلاعات را share نمی‌کنند.

---

## 9. Duplicate Audit

| نوع Duplicate | تعداد | جزئیات |
|---------------|-------|--------|
| Excel duplicate file | 1 | `dashboard/dashboard/` nested copy |
| Lock file | 1 | `~$MigrationHunter_Full_20260827_1440.xlsx` |
| Python scripts overlap | 8+ | `create_dashboard_v2.py`, `v3.py`, `run3.py` همه مشابه |
| Memory overwrite | 3 | `run_all.py` هر بار memory/ را overwrite می‌کند |
| Jobs hardcoded in 8+ scripts | 8 | هر اسکریپت لیست جداگانه دارد |

---

## 10. Conflict Audit

### تناقض‌های شناسایی شده

| # | File A | File B | Conflict | Resolution |
|---|--------|--------|----------|------------|
| 1 | Excel = APPROVED | Audit = WAITING FOR HUMAN APPROVAL | Status conflict | **Audit wins — APPROVED not granted** |
| 2 | Excel = 18 jobs | EVIDENCE_REGISTRY = 3 verified | Count mismatch | **Registry wins — only 3 truly verified** |
| 3 | run_all.py writes SOURCE_BANK | SOURCE_BANK.md has different data | Overwrite conflict | **Manual data wins** |
| 4 | Email hardcoded = 17 | APPLICATION_BANK = 0 sent | False impression | **0 sent is correct** |
| 5 | Excel "امتیاز مسیر" = calculated | Formula = NONE (hardcoded) | Score fake | **Scores are meaningless** |

---

## 11. Evidence Integrity

### برای هر Opportunity

| Opportunity | Job | Employer | Sponsorship | Visa | Email | Overall |
|-------------|-----|----------|-------------|------|-------|---------|
| #1 Health NZ → Neda | ✅ Verified | ✅ Verified | ✅ Explicit | ✅ Green List | ✅ Verified | **STRONG** |
| #2 RGH Global → Neda | ✅ Verified | ⚠️ Agency | ✅ Explicit | ✅ AEWV | ⚠️ General | **REVIEW** |
| #3 Germany Blue Card → Tohid | ❌ No Job | ❌ No Employer | N/A | ✅ Blue Card | ❌ None | **WEAK** |
| #4 Alberta HA → Neda | ⚠️ Partial | ✅ Verified | ⚠️ Possible | ✅ PNP | ✅ Verified | **REVIEW** |
| #5 Saskatchewan → Tohid | ⚠️ Partial | ✅ Verified | ⚠️ Unknown | ✅ PNP | ✅ Verified | **REVIEW** |

---

## 12. Score Integrity

### آیا امتیازها واقعی هستند؟

```
Evidence Score:    HARDCODED — بدون فرمول
Candidate Fit:     HARDCODED — بدون فرمول
Path Fit Score:    HARDCODED — بدون فرمول
Final Score:       HARDCODED — بدون فرمول
Success Rate:      HARDCODED — بدون مبنا
```

**هیچ فرمول واقعی وجود ندارد. تمام امتیازها Hard-coded و فاقد اعتبار هستند.**

---

## 13. Application Lifecycle

### Excel فعلی این مراحل را پشتیبانی می‌کند؟

| مرحله | وضعیت |
|-------|-------|
| DISCOVERED | ❌ |
| VALIDATING | ❌ |
| VERIFIED | ❌ |
| SHORTLISTED | ❌ |
| HUMAN_REVIEW | ❌ |
| APPROVED | ❌ |
| PREPARED | ❌ |
| SENT | ❌ |
| FOLLOW_UP | ❌ |
| RESPONSE | ❌ |
| INTERVIEW | ❌ |
| OFFER | ❌ |
| REJECTED | ❌ |
| CLOSED | ❌ |

**هیچ مرحله‌ای از Application Lifecycle در Excel وجود ندارد.**

---

## 14. Design Score

| معیار | امتیاز | توضیح |
|-------|--------|-------|
| Visual Hierarchy | 3/10 | فقط title row رنگی |
| RTL | 8/10 | ✅ فعال |
| Persian Typography | 7/10 | B Mitra ولی inconsistent |
| English Typography | 4/10 | Times New Roman همیشه اعمال نشده |
| Colors | 5/10 | رنگ هست ولی منطقی نیست |
| Borders | 6/10 | thin borders |
| Spacing | 3/10 | خیلی فشرده |
| Column Width | 2/10 | خیلی عریض/بیش از حد |
| Readability | 3/10 | سلول‌های طولانی |
| Navigation | 1/10 | بدون freeze/filter |
| Filtering | 0/10 | ❌ |
| Sorting | 0/10 | ❌ |
| Dashboard | 2/10 | فقط title |
| KPI | 1/10 | اعداد hardcoded |
| Charts | 0/10 | ❌ |
| Status Visualization | 2/10 | emoji فقط |
| User Experience | 2/10 | بسیار ضعیف |

**DESIGN SCORE: 35/100**

---

## 15. Root Causes

### چرا Excel ضعیف است؟

1. **داده Hardcoded** — هر بار اسکریپت جدید با داده جدید ساخته شده
2. **عدم اتصال به Memory** — Excel و MEMORY Bankها جداگانه کار می‌کنند
3. **نسخه‌سازی مکرر** — 25 اسکریپت، هرکدام نسخه‌ای جدا
4. **نبود Data Model** — جدول بزرگ به جای entity separation
5. **نبود Application Pipeline** — هیچ tracking برای فرآیند application
6. **نبود Evidence View** — اطلاعات evidence در Excel نیست
7. **نبود Chart/Dashboard** — فقط جدول متنی
8. **نبود Freeze/Filter** — غیرقابل استفاده برای مرور

---

## 16. Recommended Architecture

### Source of Truth

```
RECOMMENDED: C — Structured Data = Source of Truth
             Excel = Dashboard / Reporting only

Reason:
- Markdown files قابل grep, diff, version control هستند
- Python می‌تواند Markdown را بخواند و Excel بسازد
- Excel نباید Source of Truth باشد (قابل version control نیست)
```

### Proposed Structure

```
memory/                  ← Source of Truth
├── SOURCE_BANK.md
├── EMPLOYER_BANK.md
├── JOB_BANK.md
├── EVIDENCE_REGISTRY.md
├── APPLICATION_BANK.md
├── VISA_BANK.md
└── REGISTRATION_BANK.md

                    ↓ (Python reads memory)

dashboard/              ← Generated Dashboard
└── MigrationHunter_YYYYMMDD.xlsx
    ├── 01_داشبورد          (KPI cards, charts)
    ├── 02_فرصت‌ها           (All opportunities)
    ├── 03_ندا               (Neda-specific)
    ├── 04_توحید             (Tohid-specific)
    ├── 05_کارفرمایان        (Employers)
    ├── 06_ایمیل‌ها          (Emails + links)
    ├── 07_درخواست‌ها        (Application pipeline)
    ├── 08_پیگیری            (Follow-up tracker)
    ├── 09_ویزا              (Visa info)
    ├── 10_ثبت‌نام           (Registration)
    ├── 11_ارزیابی Evidence  (Evidence matrix)
    ├── 12_تاریخچه جستجو     (Search history)
    └── 99_تنظیمات           (Config + lookups)
```

---

## 17. Recommended Excel UX

### Sheet 01: داشبورد

```
┌─────────────────────────────────────────────┐
│  KPI CARDS                                  │
│  ┌──────┐ ┌──────┐ ┌──────┐ ┌──────┐       │
│  │  18  │ │   3  │ │   0  │ │   0  │       │
│  │فرصت  │ │verified│ │sent │ │resp │       │
│  └──────┘ └──────┘ └──────┘ └──────┘       │
│                                             │
│  TOP OPPORTUNITIES (filtered table)         │
│  Application Funnel (bar chart)             │
│  Country Distribution (pie chart)           │
│  Follow-up Due (highlighted rows)           │
└─────────────────────────────────────────────┘
```

### Sheet 02: فرصت‌ها

Columns:
```
ID | متقاضی | کشور | کارفرما | شغل | لینک | ایمیل | 
Evidence | Fit | حمایت | ویزا | زبان | ثبت‌نام | 
امتیاز نهایی | تصمیت | وضعیت | تاریخ | یادداشت
```

Features:
- Freeze top row
- Auto-filter on all columns
- Conditional formatting (green/yellow/red)
- Hyperlinks on URLs
- Data validation on Status/Decision columns

### Sheet 07: درخواست‌ها

Columns:
```
Applicant | Employer | Job | Sent Date | Method | 
Email | Response | Interview | Offer | Next Follow-up | Status
```

### Sheet 08: پیگیری

Columns:
```
Applicant | Employer | Sent | Days Since | 
Follow-up Due | Response | Next Action | Status
```

---

## 18. Final Verdict

```
CURRENT EXCEL:      REBUILD
ACTIVE FILE:        MigrationHunter_Full_20260827_1440.xlsx
SHEETS:             3 (need 12+)
ROWS:               56 (need 200+)
DATA SOURCE:        Hardcoded (need memory/*.md)
DESIGN:             35/100
DATA QUALITY:       25/100
EVIDENCE:           40/100
ARCHITECTURE:       20/100
AUTOMATION:         15/100
OVERALL:            27/100

CURRENT GENERATOR:  run_full_dashboard.py (25 scripts total)
SOURCE OF TRUTH:    SHOULD BE memory/*.md
MAIN PROBLEM:       Disconnected data + hardcoded values + no application pipeline
ROOT CAUSE:         No single data model; each script creates its own data
RECOMMENDED:        REBUILD with single Python script reading from memory/
```

---

## 19. Recommended Actions

| # | اقدام | اولویت | فایل مرتبط |
|---|-------|--------|-----------|
| 1 | حذف `dashboard/dashboard/` تکراری | P0 | dashboard/ |
| 2 | حذف `~$` lock file | P0 | dashboard/ |
| 3 | یکپارچه‌سازی 25 اسکریپت به 1 اسکریپت اصلی | P0 | *.py |
| 4 | اتصال Excel به memory/*.md | P0 | run_*.py |
| 5 | ساخت Excel جدید با 12 sheets | P0 | new script |
| 6 | Application Pipeline در Excel | P1 | new script |
| 7 | Evidence Matrix در Excel | P1 | new script |
| 8 | Follow-up Tracker در Excel | P1 | new script |
| 9 | Charts/Dashboard | P2 | new script |
| 10 | Conditional Formatting | P2 | new script |
| 11 | Freeze/Filter | P2 | new script |
| 12 | حذف اسکریپت‌های قدیمی | P3 | *.py |

---

> [!IMPORTANT]
> ## FINAL AUDIT STATUS
>
> **Audit:** COMPLETE
>
> **Excel Status:** REBUILD REQUIRED
>
> **Active Excel:** MigrationHunter_Full_20260827_1440.xlsx
>
> **Sheets:** 3 (need 12+)
>
> **Data Source:** HARDCODED (should be memory/*.md)
>
> **Design Score:** 35/100
>
> **Overall Score:** 27/100
>
> **Python Scripts:** 25 (need 1-2)
>
> **Conflicts Found:** 5
>
> **Duplicates Found:** 3
>
> **Emails Sent:** 0
>
> **Applications:** 0
>
> **Human Approval:** NOT GRANTED
>
> **Next Action:** REBUILD Excel with single script reading from memory/
>
> **Last Verified:** 2026-08-27 17:00
