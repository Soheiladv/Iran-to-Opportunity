# راهنمای اجرای MigrationHunter

## ۱. پیش‌نیازها

- Python 3.10+
- Gmail + رمز عبور اپلیکیشن (۱۶ کاراکتر) برای هر ایمیل
- `pip install -r requirements.txt` (openpyxl, requests, etc.)

## ۲. راه‌اندازی اولیه (یک بار)

```bash
python setup.py
```

سوالات:
- تعداد متقاضیان
- برای هر نفر: نام (انگلیسی + فارسی)، حرفه، جنسیت، سطح زبان
- تعداد ایمیل‌ها و لینکدین‌ها
- رمز عبور اپلیکیشن برای هر ایمیل

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
├── FOLLOWUP_REMINDER.md                           ← یادآوری‌ها
├── TOP_JOBS.md                                    ← فرصت‌های برتر
├── EMPLOYERS_TO_CONTACT.md                        ← کارفرمایان برای تماس
├── RECRUITMENT_AGENCIES.md                        ← آژانس‌های کاریابی
├── GOVERNMENT_SOURCES.md                          ← منابع دولتی
├── APPLICATIONS_TO_PREPARE.md                     ← درخواست‌ها برای آماده‌سازی
├── LANGUAGE_REGISTRATION.md                       ← وضعیت زبان و ثبت‌نام
├── EMAILS_TO_SEND.md                              ← ایمیل‌های آماده ارسال
└── DAILY_ACTIONS.md                               ← اقدامات روزانه

memory/
├── EMAIL_ANALYSIS.json                            ← داده‌های خام ایمیل
├── JOB_BANK.md                                    ← بانک مشاغل
├── EMPLOYER_BANK.md                               ← بانک کارفرمایان
├── SOURCE_BANK.md                                 ← بانک منابع (یادگیری خودکار)
├── RECRUITER_BANK.md                              ← بانک کاریابی‌ها
├── APPLICATION_BANK.md                            ← بانک درخواست‌ها
├── VISA_BANK.md                                   ← بانک ویزا
├── REGISTRATION_BANK.md                           ← بانک ثبت‌نام حرفه‌ای
└── SEARCH_HISTORY.md                              ← تاریخچه جستجو
```

## ۶. داشبورد اصلی (شیت‌ها)

| شیت | عنوان | محتوا |
|-------|-------|---------|
| ۰۱ | **داشبورد** | KPI کارت‌ها، توزیع کشور/متقاضی، تاپ ۵ |
| ۰۲ | **فرصت‌ها** | جدول کامل با فیلتر |
| ۰۳+ | **متقاضی — حرفه** | برگه فرصت‌های هر متقاضی (پویا) |
| ۰۵ | **کارفرمایان** | بانک کارفرمایان با ایمیل تأییدشده |
| ۰۶ | **ایمیل‌ها** | لیست ایمیل‌های آماده ارسال |
| ۰۷ | **درخواست‌ها** | پایپ‌لاین اپلیکیشن |
| ۰۸ | **پیگیری** | ترکر پیگیری ۷ روزه |
| ۰۹ | **ویزا** | اطلاعات ویزای هر کشور |
| ۱۰ | **ثبت‌نام** | مسیر ثبت‌نام حرفه‌ای |
| ۱۱ | **مدرک** | ماتریس ارزیابی مدارک |
| ۱۲ | **تاریخچه** | لاک جستجوها |
| ۱۳ | **تحلیل ایمیل** | آمار دسته‌بندی ایمیل‌ها |

## ۷. عیب‌یابی رایج

| خطا | راه‌حل |
|------|--------|
| `UnicodeDecodeError` | `.env` با UTF-8 ذخیره شده باشد |
| `⚠️ رمز تنظیم نشده` | رمز عبور اپلیکیشن ۱۶ رقمی را در `.env` وارد کنید |
| `❌ AuthenticationFailed` | رمز عبور اپلیکیشن اشتباه / ۲FA غیرفعال است |
| `⚠️ خطا در دریافت صفحه` | سایت بلاک کرده — فقط Job Bank Canada مطمئن کار می‌کند |
| `NameError` | فایل به‌روز نشده — `git pull` یا فایل جدید را کپی کنید |

## ۸. تغییر متقاضی/ایمیل

```bash
# مجدد اجرا کن (فایل‌ها overwrite می‌شوند)
python setup.py
```

## ۹. فرمان‌های موجود

```bash
# پایپ‌لاین کامل
python run.py

# فقط تست ایمیل
python email_analyzer.py --dry-run

# فقط تحلیل ایمیل
python email_dashboard.py

# فقط جستجوی شغل
python job_crawler.py

# فقط یادآوری پیگیری
python followup_reminder.py

# فقط داشبورد
python build_dashboard.py

# رابط وب
python web_ui.py

# اجرای مجدد راه‌اندازی
python setup.py
```

## ۱۰. نکات مهم

- `.env` را **هرگز** در GitHub push نکنید
- حافظه‌ها در `memory/*.md` به‌صورت Markdown جدول ذخیره می‌شوند
- داشبورد RTL است، فونت: B Mitra + Times New Roman
- برای اضافه کردن متقاضی دوم، `setup.py` مجدد اجرا کنید
- ایمیل‌ها فقط با تأیید صریح کاربر ارسال می‌شوند
- هرگز اطلاعات جعلی تولید نکنید — اگر مشخص نیست: UNKNOWN
