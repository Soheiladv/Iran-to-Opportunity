# راهنمای فضای کاری Migration Hunter

## نسخه: 6.0
## تاریخ: ۱۸ آگوست ۲۰۲۶

---

## هدف

این فضای کاری برای شکار فرصت شغلی بین‌المللی طراحی شده است.

**هدف نهایی:**
Job Offer → Work Visa → Legal Exit → Family Relocation

---

## ساختار پوشه

```
MigrationHunter/
├── 00_README_FA.md          ← راهنما (این فایل)
├── 01_MASTER_PROMPT.md      ← دستورالعمل اصلی
├── 02_RUN_GUIDE_FA.md       ← راهنمای اجرا
│
├── profiles/
│   ├── TOHID_PROFILE.md     ← پروفایل توحید
│   └── NEDA_PROFILE.md      ← پروفایل ندا
│
├── memory/
│   ├── SOURCE_BANK.md       ← بانک منابع
│   ├── EMPLOYER_BANK.md     ← بانک کارفرمایان
│   ├── JOB_BANK.md          ← بانک مشاغل
│   ├── RECRUITER_BANK.md    ← بانک کاریابی‌ها
│   ├── APPLICATION_BANK.md  ← بانک درخواست‌ها
│   ├── VISA_BANK.md         ← بانک ویزا
│   ├── REGISTRATION_BANK.md ← بانک ثبت‌نام حرفه‌ای
│   └── SEARCH_HISTORY.md    ← تاریخچه جستجو
│
├── input/
│   └── LATEST_SEARCH.md     ← آخرین ورودی جستجو
│
├── output/
│   ├── TOHID_TOP_JOBS.md    ← فرصت‌های برتر توحید
│   ├── NEDA_TOP_JOBS.md     ← فرصت‌های برتر ندا
│   ├── EMPLOYERS_TO_CONTACT.md
│   ├── RECRUITMENT_AGENCIES.md
│   ├── GOVERNMENT_SOURCES.md
│   ├── APPLICATIONS_TO_PREPARE.md
│   ├── LANGUAGE_REGISTRATION.md
│   ├── SOURCE_BANK_UPDATE.md
│   └── DAILY_ACTIONS.md     ← اقدامات روزانه
│
└── archive/
    └── [فایل‌های قبلی]
```

---

## چگونه اجرا کنیم

برای اجرای جستجو، کافی است بنویسید:

- `SEARCH`
- `RUN`
- `کاریابی را اجرا کن`

سیستم به‌صورت خودکار:
1. پروفایل‌ها را می‌خواند
2. بانک‌های اطلاعاتی را بارگذاری می‌کند
3. جستجو را انجام می‌دهد
4. گزارش فارسی تولید می‌کند

---

## نکات مهم

- هرگز اطلاعات جعلی تولید نکنید
- اگر چیزی مشخص نیست: UNKNOWN
- اولویت با کارفرمای واقعی است
- گزارش‌ها به فارسی باشند
