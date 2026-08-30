# MigrationHunter — راهنمای اجرا

## ۱. پیش‌نیازها
- Python 3.10+
- Gmail + App Password (۱۶ کاراکتر) برای هر ایمیل
- `pip install -r requirements.txt` (openpyxl, requests, etc.)

## ۲. راه‌اندازی اولیه (یک بار)
```bash
python setup.py
```
پرسش‌ها:
- تعداد متقاضیان
- برای هر نفر: نام، حرفه، جنسیت، سطح زبان
- تعداد ایمیل‌ها و لینکدین‌ها
- App Password برای هر ایمیل

خروجی:
- `config.json` — تنظیمات اصلی
- `.env` — رمزها (Git-ignore شده)

## ۳. تست اتصال ایمیل
```bash
python email_analyzer.py --dry-run
```
خروجی: `✅ email@gmail.com: اتصال موفق` یا `⚠️ رمز تنظیم نشده`

## ۴. اجرای کامل پایپ‌لاین
```bash
python run.py
```

### مراحل اجرا:
| مرحله | اسکریپت | توضیح | خروجی |
|-------|---------|-------|-------|
| ۱ | `email_analyzer.py` | بررسی ۳۰ روز اخیر ایمیل‌ها | `memory/EMAIL_ANALYSIS.json` |
| ۲ | `email_dashboard.py` | اکسل تحلیل ایمیل | `dashboard/Email_Analysis_*.xlsx` |
| ۳ | `job_crawler.py` | جستجو در ۹ سایت شغلی | `dashboard/Job_Crawler_*.xlsx` |
| ۴ | `followup_reminder.py` | یادآوری پیگیری‌ها | `output/FOLLOWUP_REMINDER.md` |
| ۵ | `build_dashboard.py` | داشبورد اصلی ۱۳ شیتی | `dashboard/MigrationHunter_Dashboard_*.xlsx` |

## ۵. فایل‌های خروجی مهم
```
dashboard/
├── MigrationHunter_Dashboard_YYYYMMDD_HHMM.xlsx   ← داشبورد اصلی (۱۳ شیت)
├── Email_Analysis_YYYYMMDD_HHMM.xlsx              ← تحلیل ایمیل‌ها
└── Job_Crawler_YYYYMMDD_HHMM.xlsx                 ← نتایج جستجوی کار

output/
└── FOLLOWUP_REMINDER.md                           ← یادآوری‌ها

memory/
├── EMAIL_ANALYSIS.json                            ← داده‌های خام ایمیل
├── JOB_BANK.md                                    ← بانک مشاغل
├── EMPLOYER_BANK.md                               ← بانک کارفرمایان
├── EVIDENCE_REGISTRY.md                           ← ارزیابی‌ها
└── ... (سایر بانک‌های حافظه)
```

## ۶. داشبورد اصلی (شیت‌ها)
1. **داشبورد** — KPI کارت‌ها، توزیع کشور/متقاضی، تاپ ۵
2. **فرصت‌ها** — جدول کامل با فیلتر
3. **نیلوفر — English Teacher** — فرصت‌های متقاضی ۱
4. **کارفرمایان** — بانک کارفرمایان با ایمیل تأییدشده
5. **ایمیل‌ها** — لیست ایمیل‌های آماده ارسال
6. **درخواست‌ها** — پایپ‌لاین اپلیکیشن
7. **پیگیری** — ترکر پیگیری ۷ روزه
8. **ویزا** — اطلاعات ویزای هر کشور
9. **ثبت‌نام** — مسیر ثبت‌نام حرفه‌ای
10. **Evidence** — ماتریس ارزیابی
11. **تاریخچه** — لاک جستجوها
12. **تحلیل ایمیل** — آمار دسته‌بندی ایمیل‌ها

## ۷. عیب‌یابی رایج

| خطا | راه‌حل |
|------|--------|
| `UnicodeDecodeError` | `.env` با UTF-8 ذخیره شده باشد |
| `⚠️ رمز تنظیم نشده` | App Password ۱۶ رقمی را در `.env` وارد کنید |
| `❌ AuthenticationFailed` | App Password اشتباه / ۲FA غیرفعال است |
| `⚠️ خطا در دریافت صفحه` | سایت بلاک کرده — فقط Job Bank Canada مطمئن کار می‌کند |
| `NameError: 'neda_opps'` | `build_dashboard.py` به‌روز شده — `git pull` یا فایل جدید را کپی کنید |

## ۸. تغییر متقاضی/ایمیل
```bash
# مجدد اجرا کن (فایل‌ها overwrite می‌شوند)
python setup.py
```

## ۹. نکات مهم
- `.env` را **هرگز** در GitHub push نکن
- حافظه‌ها در `memory/*.md` به‌صورت Markdown جدول ذخیره می‌شوند
- داشبورد RTL است، فونت: B Mitra + Times New Roman
- برای اضافه کردن متقاضی دوم، `setup.py` مجدد اجرا کن