# راهنمای اسکریپت‌ها — Migration Hunter

## 🚀 نحوه اجرا

### اسکریپت پایه
```bash
python run.py
```

### اسکریپت خودکار (با جستجوی وب)
```bash
python run_auto.py
```

### با پارامتر خاص
```bash
# جستجوی یک کشور
python run_auto.py --country nz
python run_auto.py --country de
python run_auto.py --country au
python run_auto.py --country ca

# جستجوی یک متقاضی
python run_auto.py --applicant neda
python run_auto.py --applicant tohid

# اجرای کامل
python run_auto.py --full
```

---

## 📁 ساختار فایل‌ها

```
MigrationHunter/
├── run.py                    ← اسکریپت پایه
├── run_auto.py               ← اسکریپت خودکار
│
├── profiles/
│   ├── TOHID_PROFILE.md      ← پروفایل توحید
│   └── NEDA_PROFILE.md       ← پروفایل ندا
│
├── memory/
│   ├── SOURCE_BANK.md        ← بانک منابع
│   ├── EMPLOYER_BANK.md      ← بانک کارفرمایان
│   ├── JOB_BANK.md           ← بانک مشاغل
│   ├── APPLICATION_BANK.md   ← بانک درخواست‌ها
│   ├── VISA_BANK.md          ← بانک ویزا
│   ├── REGISTRATION_BANK.md  ← بانک ثبت‌نام
│   └── SEARCH_HISTORY.md     ← تاریخچه جستجو
│
├── output/
│   ├── DAILY_ACTIONS.md      ← اقدامات روزانه
│   ├── NEDA_TOP_JOBS.md      ← فرصت‌های برتر ندا
│   └── TOHID_TOP_JOBS.md     ← فرصت‌های برتر توحید
│
└── dashboard/
    └── MigrationHunter_*.xlsx ← فایل Excel
```

---

## 🔄 چرخه اجرا

هر بار که اسکریپت اجرا می‌شود:

1. **📂 بارگذاری حافظه** — خواندن بانک‌های اطلاعاتی قبلی
2. **👤 بارگذاری پروفایل** — خواندن اطلاعات توحید و ندا
3. **🌐 جمع‌آوری وب** — خواندن صفحات سایت‌ها
4. **🔍 جمع‌آوری منابع** — بررسی منابع شناخته شده
5. **📊 تحلیل** — امتیازدهی به مشاغل
6. **💾 ذخیره حافظه** — بروزرسانی بانک‌ها
7. **📝 گزارش** — تولید گزارش‌های فارسی
8. **📊 Excel** — تولید فایل اکسل

---

## 📊 خروجی‌ها

### گزارش‌ها (Markdown)
- `output/DAILY_ACTIONS.md` — ۵ اقدام برتر امروز
- `output/NEDA_TOP_JOBS.md` — فرصت‌های برتر ندا
- `output/TOHID_TOP_JOBS.md` — فرصت‌های برتر توحید

### Excel Dashboard
- `dashboard/MigrationHunter_YYYYMMDD_HHMMSS.xlsx`
- شامل: Dashboard + All Jobs + Charts

---

## ⚙️ تنظیمات

### منابع شغلی
در فایل `run_auto.py`، کلاس `JobSources` را ویرایش کنید:

```python
class JobSources:
    SOURCES = {
        "nz": [
            {
                "name": "Health New Zealand",
                "url": "https://www.healthnz.govt.nz/careers/international",
                "type": "government",
                "trust_score": 95
            },
            # منابع بیشتر...
        ],
    }
```

### امتیازدهی
در فایل `run.py`، کلاس `JobAnalyzer` را ویرایش کنید.

---

## 🛠 نیازمندی‌ها

```bash
pip install openpyxl requests
```

---

## 📝 نکات مهم

1. **حافظه پاک نمی‌شود** — هر اجرا روی اجراهای قبلی سوار می‌شود
2. **تکراری حذف می‌شود** — مشاغل تکراری شناسایی و حذف می‌شوند
3. **امتیازدهی خودکار** — هر فرصت بر اساس Path Fit Score امتیازدهی می‌شود
4. **گزارش فارسی** — تمام گزارش‌ها به زبان فارسی هستند
5. **Excel خودکار** — فایل Excel با هر اجرا بروزرسانی می‌شود

---

## 🔮 قابلیت‌های آینده

- [ ] اتصال به AI API برای تحلیل هوشمندتر
- [ ] ارسال ایمیل خودکار (با تأیید کاربر)
- [ ] پیگیری وضعیت درخواست‌ها
- [ ] گزارش‌های هفتگی/ماهانه
- [ ] داشبورد وب
