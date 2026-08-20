# 🔒 سرورهای VPN رایگان — لیست بروزرسانی‌شده

**تاریخ بروزرسانی:** ۱۸ آگوست ۲۰۲۶ (۲۷ مرداد ۱۴۰۵)

---

## 🇯🇵 سرورهای ژاپن (پیشنهادی — سریع‌ترین)

| ردیف | IP Address | Uptime | Ping |
|------|-----------|--------|------|
| 1 | 219.100.37.176 | 85 days | 20 ms |
| 2 | 219.100.37.163 | 85 days | 22 ms |
| 3 | 219.100.37.31 | 85 days | 9 ms ⭐ |
| 4 | 219.100.37.94 | 85 days | 13 ms ⭐ |
| 5 | 219.100.37.117 | 85 days | 14 ms |
| 6 | 219.100.37.98 | 85 days | 17 ms |
| 7 | 219.100.37.209 | 85 days | 9 ms ⭐ |
| 8 | 219.100.37.199 | 85 days | 9 ms ⭐ |
| 9 | 219.100.37.100 | 85 days | 9 ms ⭐ |
| 10 | 219.100.37.172 | 85 days | 14 ms |
| 11 | 219.100.37.110 | 85 days | 14 ms |
| 12 | 219.100.37.81 | 85 days | 20 ms |
| 13 | 59.133.158.94 | 7 days | 9 ms ⭐ |
| 14 | 219.100.37.125 | 85 days | 20 ms |
| 15 | 219.100.37.13 | 85 days | 17 ms |
| 16 | 219.100.37.119 | 85 days | 25 ms |
| 17 | 115.179.206.241 | 3 days | 15 ms |
| 18 | 110.163.135.59 | 2 days | 8 ms ⭐ |
| 19 | 128.27.28.120 | 1 days | 7 ms ⭐ |

---

## 🔐 تنظیمات اتصال (L2TP/IPsec)

**Pre-shared key (PSK):** `vpn`
**Username:** `vpn`
**Password:** `vpn`

### نحوه اتصال در Windows:

1. **Settings** → **Network & Internet** → **VPN**
2. **Add VPN connection**
3. تنظیمات:
   - **VPN provider:** Windows (built-in)
   - **Connection name:** VPN Gate Japan
   - **Server name or address:** `219.100.37.31` (یا هر IP دیگر)
   - **VPN type:** L2TP/IPsec with pre-shared key
   - **Pre-shared key:** `vpn`
   - **Type of sign-in info:** Username and password
   - **Username:** `vpn`
   - **Password:** `vpn`
4. **Save** → **Connect**

---

## ⚠️ نکات مهم

1. **سرورهای ژاپن** معمولاً سریع‌ترین و پایدارترین هستند
2. **Ping کمتر** = سرعت بهتر
3. **Uptime بیشتر** = پایداری بیشتر
4. سرورها **رایگان** هستند ولی **محدودیت پهنای باند** ممکن است داشته باشند
5. اگر L2TP کار نکرد، **OpenVPN** یا **SSTP** را امتحان کنید

---

## 🔄 نحوه بروزرسانی سرورها

برای دریافت لیست تازه سرورها:
1. به `https://www.vpngate.net/en/` بروید
2. لیست سرورها را دانلود کنید
3. یا از **SoftEther VPN Client** با قابلیت **VPN Gate** استفاده کنید

---

## 🛠 عیب‌یابی

### اگر L2TP کار نکرد:
1. **Registry Fix:**
   ```
   HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\PolicyAgent
   ```
   DWORD: `AssumeUDPEncapsulationContextOnSendRule` = `2`

2. ** restarting پروxy:**
   - Internet Options → Connections → LAN Settings → Uncheck all

3. **Firewall:**
   - Windows Firewall را غیرفعال کنید
