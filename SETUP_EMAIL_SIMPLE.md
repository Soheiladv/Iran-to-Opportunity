# تنظیم ساده ایمیل — فقط یک فایل!

## فقط این یک فایل را ویرایش کن:

### `MigrationHunter/.env`

```env
# ═══════════════════════════════════════════════════
# ایمیل توحید
# ═══════════════════════════════════════════════════
EMAIL_TOHID_1=t.arjmand1980@gmail.com
EMAIL_PASSWORD_TOHID_1=اینجا_کد_۱۶_رقمی_توحید
EMAIL_PROVIDER_TOHID_1=gmail

# ═══════════════════════════════════════════════════
# ایمیل ندا
# ═══════════════════════════════════════════════════
EMAIL_NEDA_1=n.arjmand.85@gmail.com
EMAIL_PASSWORD_NEDA_1=اینجا_کد_۱۶_رقمی_ندا
EMAIL_PROVIDER_NEDA_1=gmail
```

## کد ۱۶ رقمی را از کجا بگیرم؟

1. برو به: **https://myaccount.google.com/apppasswords**
2. رمز عبور برنامه بساز
3. کد 16 رقمی را کپی کن
4. جای `اینجا_کد_۱۶_رقمی` بگذار

## تست کن:

```bash
python email_analyzer.py --dry-run
```

اگر دیدی:
```
✅ t.arjmand1980@gmail.com: اتصال موفق
✅ n.arjmand.85@gmail.com: اتصال موفق
```

یعنی درست شد! حالا اجرا کن:
```bash
python run.py
```
