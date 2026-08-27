# راهنمای کامل اتصال ایمیل — MigrationHunter

## هدف

اتصال ایمیل شغلی به ابزار تحلیل ایمیل پروژه MigrationHunter
تا بتوانیم ایمیل‌های کاریابی را خودکار خوانده، دسته‌بندی و تطبیق دهیم.

---

## فهرست

1. [چرا App Password لازم است؟](#۱-چرا-app-password-لازم-است)
2. [مراحل Gmail — مرحله به مرحله](#۲-مراحل-gmail)
3. [مراحل Outlook / Office 365](#۳-مراحل-outlook--office-365)
4. [مراحل Yahoo](#۴-مراحل-yahoo)
5. [تنظیم فایل .env](#۵-تنظیم-fایل-env)
6. [تست اتصال](#۶-تست-اتصال)
7. [اجرای تحلیل ایمیل](#۷-اجرای-تحلیل-ایمیل)
8. [عیب‌یابی](#۸-عیب‌یابی)
9. [امنیت](#۹-امنیت)

---

## ۱. چرا App Password لازم است؟

Gmail و سرویس‌های مشابه **اجازه ورود مستقیم با پسورد اصلی را نمی‌دهند.**

به جای آن باید یک **App Password** بسازید:

- پسوردی 16 رقمی مخصوص برنامه
- فقط برای یک برنامه خاص
- قابل لغو در هر زمان
- پسورد اصلی Gmail شما را فاش نمی‌کند

---

## ۲. مراحل Gmail

### مرحله ۱: ورود به حساب Google

1. مرورگر را باز کن
2. برو به: **https://myaccount.google.com**
3. با ایمیل `t.arjmand1980@gmail.com` وارد شو

### مرحله ۲: فعال کردن 2-Step Verification

> ⚠️ بدون فعال کردن 2FA، App Password قابل ساخت نیست.

1. در صفحه اصلی، روی **Security** کلیک کن (در سمت چپ یا بالا)
2. بخش **"How you sign in to Google"** را پیدا کن
3. روی **"2-Step Verification"** کلیک کن
4. روی **"Get started"** بزن
5. شماره تلفن خودت را وارد کن (می‌تواند ایران باشد)
6. **SMS** یا **Phone call** را انتخاب کن
7. کد تأیید را وارد کن
8. **Turn on** بزن

✅ حالا 2-Step Verification فعال شده

### مرحله ۳: ساخت App Password

1. هنوز در صفحه **Security** هستی
2. به بخش **"How you sign in to Google"** برگرد
3. **"App passwords"** را پیدا کن و کلیک کن

> اگر "App passwords" را نمی‌بینی:
> - مستقیماً برو به: **https://myaccount.google.com/apppasswords**
> - یا در Google Search تایپ کن: "google app passwords"

4. در صفحه App passwords:
   - **Select app:** انتخاب کن **"Mail"**
   - **Select device:** انتخاب کن **"Windows Computer"** یا **"Other (Custom name)"**
   - اگر Custom name زدی، بنویس: **MigrationHunter**
5. روی **"Create"** کلیک کن

### مرحله ۴: کپی کردن کد

6. یک کد 16 رقمی نمایش داده می‌شود:

```
abcd efgh ijkl mnop
```

> ⚠️ این کد فقط **یک بار** نمایش داده می‌شود!
> حتماً آن را **کپی** یا **یادداشت** کن.

7. روی **"Done"** بزن

### مرحله ۵: استفاده از کد

حالا این کد را در فایل `.env` پروژه وارد کن (مرحله ۵ راهنما).

---

## ۳. مراحل Outlook / Office 365

1. وارد **https://account.microsoft.com** شو
2. **Security** → **Advanced security options**
3. **App passwords** → **Create a new app password**
4. کد 16 رقمی را کپی کن

یا مستقیماً:
**https://account.microsoft.com/security/app-passwords**

---

## ۴. مراحل Yahoo

1. وارد **https://login.yahoo.com** شو
2. **Account Security** → **Manage app passwords**
3. **Generate app password**
4. نام: **MigrationHunter**
5. کد را کپی کن

---

## ۵. تنظیم فایل `.env`

فایل `.env` در پوشه اصلی پروژه:

```
MigrationHunter/.env
```

محتوا را این‌طور تنظیم کن:

```env
# ═══════════════════════════════════════
# MigrationHunter — Email Config
# ═══════════════════════════════════════

EMAIL_ADDRESS=t.arjmand1980@gmail.com
EMAIL_PASSWORD=abcd efgh ijkl mnop
EMAIL_PROVIDER=gmail
```

### توضیح فیلدها

| فیلد | مقدار | توضیح |
|------|-------|-------|
| `EMAIL_ADDRESS` | `t.arjmand1980@gmail.com` | آدرس ایمیل شما |
| `EMAIL_PASSWORD` | `abcd efgh ijkl mnop` | **کد 16 رقمی App Password** — نه پسورد Gmail! |
| `EMAIL_PROVIDER` | `gmail` | سرویس ایمیل: `gmail`, `outlook`, `yahoo`, `icloud` |

### ⚠️ نکات مهم

- **پسورد اصلی Gmail را ننویسید** — فقط App Password
- **فاصله بین حروف اشکالی ندارد** — `abcd efgh ijkl mnop` درست است
- **بزرگی و کوچکی حروف مهم نیست** — `ABCD` همان `abcd` است
- **این فایل را در GitHub پوش نکنید** — `.gitignore` آن را نادیده می‌گیرد

---

## ۶. تست اتصال

قبل از اجرای کامل، فقط اتصال را تست کن:

```bash
cd MigrationHunter
python email_analyzer.py --dry-run
```

اگر موفق باشد:

```
✅ اتصال موفق! (dry-run mode)

📁 پوشه‌ها:
  - INBOX
  - [Gmail]/Sent Mail
  - [Gmail]/Drafts
  - ...
```

اگر خطا داد، بخش [عیب‌یابی](#۸-عیب‌یابی) را بخوان.

---

## ۷. اجرای تحلیل ایمیل

### تحلیل ۳۰ روز اخیر

```bash
python email_analyzer.py
```

### تحلیل ۶۰ روز اخیر

```bash
python email_analyzer.py --days 60
```

### تحلیل با محدودیت تعداد

```bash
python email_analyzer.py --limit 500
```

### خروجی‌ها

| فایل | مسیر | توضیح |
|------|------|-------|
| گزارش فارسی | `output/EMAIL_ANALYSIS_REPORT.md` | خلاصه تحلیل + اقدامات |
| داده JSON | `memory/EMAIL_ANALYSIS.json` | داده ساختاریافته |

---

## ۸. عیب‌یابی

### خطا: `AUTHENTICATIONFAILED`

```
❌ خطا در اتصال: [AUTHENTICATIONFAILED] Invalid credentials
```

**دلیل:** App Password اشتباه یا وجود ندارد.

**راه‌حل:**
1. مطمئن شو 2-Step Verification فعال است
2. یک App Password جدید بساز
3. کد را دقیقاً کپی کن (بدون فاصله اضافی)
4. در `.env` جایگزین کن

### خطا: `Application-specific password required`

```
❌ خطا: Application-specific password required
```

**دلیل:** داری از پسورد اصلی Gmail استفاده می‌کنی، نه App Password.

**راه‌حل:** حتماً App Password بساز (مرحله ۳ بالا).

### خطا: `Connection refused`

```
❌ خطا: Connection refused
```

**دلیل:** IMAP غیرفعال است یا فایروال بلاک می‌کند.

**راه‌حل:**
1. Gmail: IMAP خودکار فعال است — مشکل فایروال است
2. فایروال/آنتی‌ویروس را بررسی کن
3. VPN را خاموش کن و دوباره تست کن

### خطا: ` imaplib.IMAP4.error: [ALERT]`

**دلیل:** Gmail اجازه IMAP نمی‌دهد.

**راه‌حل:**
1. برو به: **https://mail.google.com/mail/u/0/#settings/fwdandpop**
2. تب **"Forwarding and POP/IMAP"**
3. **"Enable IMAP"** را فعال کن
4. **Save Changes**

### App Password را پیدا نمی‌کنم

اگر گزینه "App passwords" در صفحه Security نیست:
1. مستقیماً برو به: **https://myaccount.google.com/apppasswords**
2. یا در Google جستجو کن: **"google create app password"**

---

## ۹. امنیت

### ✅ انجام بده

- از **App Password** استفاده کن (نه پسورد اصلی)
- App Password را فقط در `.env` ذخیره کن
- `.env` را **هرگز** در Git پوش نکن
- اگر نیاز نداری، App Password را **حذف** کن

### ❌ انجام نده

- پسورد اصلی Gmail را هیچ‌جا ننویس
- App Password را به کسی نده
- App Password را در کد Python ننویس
- `.env` را در GitHub آپلود نکن

### 🔒 حذف App Password

اگر دیگر نیاز نداری:
1. **myaccount.google.com** → Security → App passwords
2. کنار App Password مربوطه، **垃圾桶** (Remove) را بزن
3. تأیید کن

---

## خلاصه سریع

```
۱. Gmail → Security → 2-Step Verification → ON
۲. Gmail → Security → App passwords → Create
۳. کد 16 رقمی را کپی کن
۴. در .env بنویس:
   EMAIL_PASSWORD=کد_۱۶_رقمی
۵. تست کن:
   python email_analyzer.py --dry-run
۶. اجرا کن:
   python email_analyzer.py
```

---

> **تاریخ ایجاد:** 2026-08-27
> **پروژه:** MigrationHunter — Iran-to-Opportunity
