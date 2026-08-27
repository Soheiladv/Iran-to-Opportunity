# ============================================================
# 🚀 نصب اکستنشن‌های رایگان VSCode — AI Agents + API Tools
# ============================================================
# اجرای اسکریپت:
#   powershell -ExecutionPolicy Bypass -File install_vscode_extensions.ps1
# ============================================================

$ErrorActionPreference = "SilentlyContinue"

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  نصب اکستنشن‌های رایگان VSCode" -ForegroundColor Cyan
Write-Host "  AI Agents + API Tools" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --- بررسی وجود code در PATH ---
if (-not (Get-Command code -ErrorAction SilentlyContinue)) {
    Write-Host "❌ دستور 'code' پیدا نشد!" -ForegroundColor Red
    Write-Host "   ابتدا VSCode را نصب کنید و 'code' را به PATH اضافه کنید." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   راهنمای اضافه کردن به PATH:" -ForegroundColor Yellow
    Write-Host "   Ctrl+Shift+P → Shell Command: Install 'code' command in PATH" -ForegroundColor Gray
    exit 1
}

Write-Host "✅ VSCode CLI پیدا شد" -ForegroundColor Green
Write-Host ""

# ============================================================
# 🤖 بخش ۱: AI Agents رایگان
# ============================================================
Write-Host "--- 🤖 AI Agents رایگان ---" -ForegroundColor Magenta

$aiAgents = @(
    @{ Name="GitHub Copilot";      Id="GitHub.copilot";                 Desc="دستیار AI مایکروسافت (رایگان برای اکانت مایکروسافت)" }
    @{ Name="GitHub Copilot Chat"; Id="GitHub.copilot-chat";            Desc="چت AI در VSCode" }
    @{ Name="Continue";            Id="Continue.continue";              Desc="Agent متن‌باز — پشتیبانی از Ollama/Gemini/OpenAI" }
    @{ Name="Cline";               Id="saoudrizwan.claude-dev";         Desc="Agent خودکار — اجرای فایل، ترمینال، وب" }
    @{ Name="Amazon Q";            Id="Amazonwebservices.amazon-q-vscode"; Desc="AI رایگان آمازون — کد و تولید متن" }
    @{ Name="Codeium / Windsurf";  Id="Codeium.codeium";                Desc="Copilot رایگان با اکانت شخصی" }
    @{ Name="Cody (Sourcegraph)";  Id="sourcegraph.cody-ai";            Desc="AI متن‌باز از Sourcegraph" }
    @{ Name="CodeGPT";             Id="DanielSanMedium.dscodegpt";    Desc="پشتیبانی از Ox Alpha و مدل‌های رایگان" }
    @{ Name="Kilo Code";           Id="kilocode.kilo-code";           Desc="Agent متن‌باز — پشتیبانی از MiMo 2.5 از طریق OpenRouter" }
)

foreach ($ext in $aiAgents) {
    Write-Host "  📦 $($ext.Name): $($ext.Desc)" -ForegroundColor White
    code --install-extension $ext.Id --force 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "    ✅ نصب شد" -ForegroundColor Green
    } else {
        Write-Host "    ⚠️ قبلاً نصب است یا خطا" -ForegroundColor Yellow
    }
}

Write-Host ""

# ============================================================
# 🔌 بخش ۲: ابزارهای API رایگان
# ============================================================
Write-Host "--- 🔌 ابزارهای API ---" -ForegroundColor Magenta

$apiTools = @(
    @{ Name="REST Client";      Id="humao.rest-client";       Desc="تست API بدون خروج از VSCode" }
    @{ Name="Thunder Client";   Id="rangav.vscode-thunder-client"; Desc="GUI تست API مشابه Postman" }
    @{ Name="API Debug";        Id="rangav.vscode-api-debug";  Desc="ابزار دیباگ API" }
    @{ Name="Postman";          Id="Postman.postman-for-vscode"; Desc="نسخه رسمی Postman" }
)

foreach ($ext in $apiTools) {
    Write-Host "  📦 $($ext.Name): $($ext.Desc)" -ForegroundColor White
    code --install-extension $ext.Id --force 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "    ✅ نصب شد" -ForegroundColor Green
    } else {
        Write-Host "    ⚠️ قبلاً نصب است یا خطا" -ForegroundColor Yellow
    }
}

Write-Host ""

# ============================================================
# 🐍 بخش ۳: اکستنشن‌های مفید برای این پروژه
# ============================================================
Write-Host "--- 🐍 اکستنشن‌های مفید برای MigrationHunter ---" -ForegroundColor Magenta

$projectTools = @(
    @{ Name="Python";               Id="ms-python.python";                Desc="پشتیبانی کامل Python" }
    @{ Name="Python Debugger";      Id="ms-python.debugpy";               Desc="دیباگر Python" }
    @{ Name="dotenv support";       Id="mikestead.dotenv";                Desc="خواندن فایل .env" }
    @{ Name="GitLens";              Id="eamodio.gitlens";                 Desc="تاریخچه Git و BLAME" }
    @{ Name="PowerShell";           Id="ms-vscode.powershell";            Desc="ابزار PowerShell" }
    @{ Name="Markdown Preview";     Id="shd101wyy.markdown-preview-enhanced"; Desc="پیش‌نمایش Markdown فارسی" }
    @{ Name="TODO Highlight";       Id="wayou.vscode-todo-highlight";     Desc="نمایش TODO در کد" }
    @{ Name="Error Lens";           Id="usernamehw.errorlens";            Desc="نمایش خطا در خط کد" }
)

foreach ($ext in $projectTools) {
    Write-Host "  📦 $($ext.Name): $($ext.Desc)" -ForegroundColor White
    code --install-extension $ext.Id --force 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "    ✅ نصب شد" -ForegroundColor Green
    } else {
        Write-Host "    ⚠️ قبلاً نصب است یا خطا" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  ✅ نصب اکستنشن‌ها تمام شد!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📝 نکات:" -ForegroundColor Yellow
Write-Host "   • GitHub Copilot: با اکانت مایکروسافت رایگان است" -ForegroundColor Gray
Write-Host "   • Continue: از Ollama برای مدل‌های محلی رایگان استفاده کنید" -ForegroundColor Gray
Write-Host "   • Codeium: با اکانت شخصی رایگان است" -ForegroundColor Gray
Write-Host "   • CodeGPT: Ox Alpha رایگان از طریق OpenRouter" -ForegroundColor Gray
Write-Host "   • Kilo Code: MiMo 2.5 رایگان از طریق OpenRouter" -ForegroundColor Gray
Write-Host "   • REST Client: فایل .http بسازید و API تست کنید" -ForegroundColor Gray
Write-Host ""
