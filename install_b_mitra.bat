@echo off
echo ========================================
echo Installing B Mitra Font
echo ========================================
echo.

REM Check if font already exists
if exist "%WINDIR%\Fonts\BMitra.ttf" (
    echo B Mitra font is already installed.
    goto :end
)

REM Try to copy font from common locations
if exist "%~dp0fonts\BMitra.ttf" (
    copy "%~dp0fonts\BMitra.ttf" "%WINDIR%\Fonts\" /Y
    copy "%~dp0fonts\BMitra.ttf" "C:\Windows\Fonts\" /Y
    echo Font installed successfully from local folder.
) else (
    echo Font file not found in local folder.
    echo Please download B Mitra font and place it in the fonts folder.
    echo.
    echo Download from:
    echo https://www.fontyab.com/2095/b-mitra.html
    echo.
    echo Then run this script again.
)

:end
echo.
echo Done.
pause
