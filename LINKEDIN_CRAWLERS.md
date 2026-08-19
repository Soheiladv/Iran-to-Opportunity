# LinkedIn Crawlers — راهنمای ابزارها

## 🎯 بهترین ابزارها برای کاریابی

### ۱. JobFunnel ⭐⭐⭐⭐⭐
**GitHub:** https://github.com/PaulMcInnis/JobFunnel

| ویژگی | مقدار |
|-------|-------|
| **Stars** | 2.2k |
| **زبان** | Python |
| **عملکرد** | جمع‌آوری مشاغل از سایت‌های مختلف |
| **خروجی** | CSV/Excel |
| **مزیت** | حذف خودکار تکراری |

```bash
pip install jobfunnel
```

### ۲. scrape-linkedin-selenium ⭐⭐⭐⭐
**GitHub:** https://github.com/austinoboyle/scrape-linkedin-selenium

| ویژگی | مقدار |
|-------|-------|
| **Stars** | 537 |
| **زبان** | Python |
| **عملکرد** | کراول پروفایل LinkedIn |
| **خروجی** | JSON/CSV |
| **مزیت** | استخراج اطلاعات کامل پروفایل |

```bash
pip install scrape-linkedin-selenium
```

### ۳. Scout ⭐⭐⭐⭐
**GitHub:** https://github.com/kiryano/Scout

| ویژگی | مقدار |
|-------|-------|
| **Stars** | 492 |
| **زبان** | Python |
| **عملکرد** | کراول LinkedIn + Instagram + TikTok |
| **خروجی** | CSV |
| **مزیت** | استخراج ایمیل از بیو |

### ۴. AI_Resume_Builder ⭐⭐⭐⭐
**GitHub:** https://github.com/feder-cr/resume_render_from_job_description

| ویژگی | مقدار |
|-------|-------|
| **Stars** | 415 |
| **زبان** | Python |
| **عملکرد** | تحلیل شغل + ساخت CV با AI |
| **خروجی** | PDF |
| **مزیت** | شخصی‌سازی CV بر اساس آگهی |

### ۵. LinkedIn Jobs Scraper ⭐⭐⭐
**GitHub:** https://github.com/kirkhunter/linkedin-jobs-scraper

| ویژگی | مقدار |
|-------|-------|
| **Stars** | 190 |
| **زبان** | Python |
| **عملکرد** | کراول آگهی‌های شغلی LinkedIn |
| **خروجی** | CSV/JSON |
| **مزیت** | ساده و سبک |

---

## 🔧 نصب و استفاده

### JobFunnel (پیشنهادی)

```bash
# نصب
pip install jobfunnel

# استفاده
jobfunnel --country nz --keywords "midwife"
jobfunnel --country de --keywords "IT manager"
```

### scrape-linkedin-selenium

```bash
# نصب
pip install scrape-linkedin-selenium

# استفاد
scrape_linkedin profile https://www.linkedin.com/in/username
```

---

## ⚠️ نکات مهم

### ۱. محدودیت‌های LinkedIn
- LinkedIn مکرراً کراولرها را بلاک می‌کند
- استفاده از VPN ضروری است
- سرعت کراول را کم کنید
- از代理 (proxy) استفاده کنید

### ۲. قوانین استفاده
- اطلاعات شخصی را بدون اجازه منتشر نکنید
- فقط برای استفاده شخصی کراول کنید
- از داده‌ها برای کاریابی استفاده کنید

### ۳. جایگزین کراول
- **LinkedIn API رسمی:** محدود ولی امن
- **LinkedIn Jobs API:** برای جستجوی شغل
- **وب‌سایت‌های کاریابی:** Indeed, Seek, Trade Me

---

## 🎯 پیشنهاد برای پروژه Iran-to-Opportunity

### مرحله ۱: کراول پروفایل خودتان
```bash
# پروفایل خودتان را کراول کنید
python -c "
from scrape_linkedin import Scraper
s = Scraper()
profile = s.scrape_profile('https://www.linkedin.com/in/YOUR_USERNAME')
print(profile)
"
```

### مرحله ۲: کراول آگهی‌های شغلی
```bash
# مشاغل مرتبط را جستجو کنید
python -c "
import requests
from bs4 import BeautifulSoup

# LinkedIn Jobs search
url = 'https://www.linkedin.com/jobs/search/?keywords=midwife&location=New+Zealand'
# Note: LinkedIn blocks direct scraping, use API or Selenium
"
```

### مرحله ۳: ترکیب با AI
```bash
# اطلاعات کراول شده را با AI تحلیل کنید
python run_auto.py --ai --linkedin-data profile.json
```

---

## 📚 منابع بیشتر

- [JobFunnel Documentation](https://jobfunnel.readthedocs.io/)
- [LinkedIn API Documentation](https://docs.microsoft.com/en-us/linkedin/)
- [Selenium Documentation](https://www.selenium.dev/documentation/)

---

**آخرین بروزرسانی:** 2026-08-19
