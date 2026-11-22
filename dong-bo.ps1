<#
.SYNOPSIS
    Hue Sound Express - Dong bo 1 cham sau khi chinh sua file CSV bang Excel.
    Tu dong: Kiem tra khoa file -> Build 1107 trang -> git add -A -> git commit -> git push.
.EXAMPLE
    .\dong-bo.ps1
    .\dong-bo.ps1 -Message "cap nhat gia dia than thang 11"
#>
param(
    [string]$Message = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Set-Location -Path $PSScriptRoot

Write-Host "=== HUE SOUND EXPRESS - DONG BO DU LIEU LEN WEBSITE ===" -ForegroundColor Cyan

# 1. Kiem tra xem Excel co dang mo khoa file CSV khong
$csvFiles = @(
    ".\data\vinyl_lp_enriched.csv",
    ".\data\templates\cd_sacd_template.csv",
    ".\data\templates\turntables_template.csv",
    ".\data\templates\accessories_care_template.csv"
)
foreach ($f in $csvFiles) {
    if (Test-Path $f) {
        try {
            $fs = [System.IO.File]::Open((Resolve-Path $f).Path, 'Open', 'ReadWrite', 'None')
            $fs.Close()
        } catch {
            Write-Host "[LOI] File '$f' dang mo trong Excel! Hay luu (Ctrl+S), dong han Excel roi chay lai .\dong-bo.ps1" -ForegroundColor Red
            exit 1
        }
    }
}

# 2. Tu dong kiem tra & lam sach loi HTML rac / link anh nham trong CSV
Write-Host "`n[1/5] Dang kiem tra & tu dong lam sach du lieu CSV..." -ForegroundColor Yellow
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\kiem-tra-va-sua-csv.ps1 -AutoFix

# 3. Chay build toan bo du lieu va trang tinh /p/
Write-Host "`n[2/5] Dang chay build_site_data.ps1 -GenerateProductPages..." -ForegroundColor Yellow
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\build_site_data.ps1 -GenerateProductPages
if ($LASTEXITCODE -ne 0) {
    Write-Host "[LOI] Qua trinh build gap su co!" -ForegroundColor Red
    exit 1
}

# 4. Dong goi tat ca file thay doi (git add -A)
Write-Host "`n[3/5] Dang dong goi tat ca thay doi (git add -A)..." -ForegroundColor Yellow
git add -A

# Kiem tra xem co gi thay doi de commit khong
$staged = git diff --cached --name-only
if (-not $staged) {
    Write-Host "`nKhong co thay doi moi nao so voi ban da luu. Dang kiem tra dong bo GitHub..." -ForegroundColor Green
    git push
    exit 0
}

# 5. Tao commit
$ts = Get-Date -Format "yyyy-MM-dd HH:mm"
$commitMsg = if ([string]::IsNullOrWhiteSpace($Message)) { "chore: dong bo du lieu CSV ($ts)" } else { $Message }
Write-Host "`n[4/5] Dang luu phien ban moi (git commit)..." -ForegroundColor Yellow
git commit -m $commitMsg

# 6. Day len GitHub Pages
Write-Host "`n[5/5] Dang day len GitHub Pages (git push)..." -ForegroundColor Yellow
git push
if ($LASTEXITCODE -ne 0) {
    Write-Host "Dang keo ban moi nhat tu GitHub roi day lai (git pull --rebase)..." -ForegroundColor DarkYellow
    git pull --rebase origin main
    git push
}

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n[THANH CONG] Da dong bo len GitHub Pages! Cho 1-2 phut roi bam Ctrl+F5 tai https://huesoundexpress.github.io/" -ForegroundColor Green
} else {
    Write-Host "`n[LOI] Khong the day len GitHub. Kiem tra ket noi mang hoac dang nhap Git." -ForegroundColor Red
    exit 1
}
