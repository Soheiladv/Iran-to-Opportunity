@echo off
chcp 65001 >nul
title MigrationHunter — شروع سریع

echo ═══════════════════════════════════════════════════
echo   MigrationHunter — شروع سریع
echo   Iran-to-Opportunity
echo ═══════════════════════════════════════════════════
echo.

:menu
echo  چه کاری می‌خواهید انجام دهید؟
echo.
echo  [1] اجرای کامل (همه چیز)
echo  [2] فقط تحلیل ایمیل
echo  [3] فقط داشبورد
echo  [4] فقط جستجوی کار
echo  [5] فقط یادآوری پیگیری
echo  [6] راه‌اندازی سیستم جدید
echo  [7] تست اتصال ایمیل
echo  [8] خروج
echo.
set /p choice=انتخاب کنید (1-8): 

if "%choice%"=="1" goto full
if "%choice%"=="2" goto email
if "%choice%"=="3" goto dashboard
if "%choice%"=="4" goto crawl
if "%choice%"=="5" goto followup
if "%choice%"=="6" goto setup
if "%choice%"=="7" goto test
if "%choice%"=="8" goto end

echo انتخاب نامعتبر!
goto menu

:full
echo.
echo ▶ اجرای کامل...
echo.
python run.py
echo.
pause
goto menu

:email
echo.
echo ▶ تحلیل ایمیل...
echo.
python email_analyzer.py
echo.
pause
goto menu

:dashboard
echo.
echo ▶ ساخت داشبورد...
echo.
python build_dashboard.py
echo.
pause
goto menu

:crawl
echo.
echo ▶ جستجوی کار...
echo.
python job_crawler.py
echo.
pause
goto menu

:followup
echo.
echo ▶ یادآوری پیگیری...
echo.
python followup_reminder.py
echo.
pause
goto menu

:setup
echo.
echo ▶ راه‌اندازی سیستم...
echo.
python setup.py
echo.
pause
goto menu

:test
echo.
echo ▶ تست اتصال ایمیل...
echo.
python email_analyzer.py --dry-run
echo.
pause
goto menu

:end
echo.
echo خداحافظ!
timeout /t 2 >nul
