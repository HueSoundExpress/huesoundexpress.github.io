<#
.SYNOPSIS
    Hue Sound Express - Tu dong kiem tra & sua loi du lieu CSV (HTML rac, entity, link anh nham, slug loi).
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\scripts\kiem-tra-va-sua-csv.ps1 -AutoFix
#>
param(
    [switch]$AutoFix
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Clean-PlainText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = [string]$Text
    # Cat bo toan bo phan HTML bi dinh tu the dong </p> / <p class=... / <a href=... / <div...
    $t = [regex]::Replace($t, '(?is)</?(?:p|div|span|a|i|ul|li|table|tr|td|br)\b.*$', '')
    # Loai bo cac the HTML con sot
    $t = [regex]::Replace($t, '<[^>]+>', ' ')
    # Giai ma HTML entity (&amp; &quot; &#039; &nbsp; ...)
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    # Loai bo dau ngoac kep thua / xuong dong / khoang trang kep
    $t = [regex]::Replace($t, '[\r\n\t]+', ' ')
    $t = [regex]::Replace($t, '\s{2,}', ' ').Trim()
    return $t
}

function Save-CsvUtf8Bom($rows, $path) {
    $csvText = ($rows | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Resolve-Path $path).Path, $csvText + "`r`n", $utf8Bom)
}

$csvTargets = @(
    ".\data\vinyl_lp_enriched.csv",
    ".\data\templates\cd_sacd_template.csv",
    ".\data\templates\turntables_template.csv",
    ".\data\templates\accessories_care_template.csv"
)

$htmlPattern = '(?i)(<[^>]+>|class\s*=|href\s*=|</p|<p\b|<div|<span|<a\b|&amp;|&quot;|&#\d+;|&nbsp;)'
$totalFixed = 0
$totalWarnings = 0

foreach ($csvPath in $csvTargets) {
    if (-not (Test-Path $csvPath)) { continue }
    $rows = Import-Csv $csvPath
    $fileDirty = $false
    $seenSku = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($r in $rows) {
        # 1. Kiem tra trung SKU
        if ($r.PSObject.Properties['sku'] -and $r.sku) {
            if (-not $seenSku.Add($r.sku)) {
                Write-Host "  [CANH BAO] Trung SKU '$($r.sku)' trong $csvPath" -ForegroundColor Yellow
                $totalWarnings++
            }
        }

        # 1b. Tu dong lam sach dau ngoac kep thua / khoang trang / dau phay trong cac cot so (id, list_price, contact_only, year)
        foreach ($numCol in @('id','list_price','contact_only','year')) {
            if ($r.PSObject.Properties[$numCol] -and $r.$numCol) {
                $rawNum = [string]$r.$numCol
                $cleanNum = ($rawNum -replace '[^0-9]', '')
                if ($rawNum -ne $cleanNum) {
                    Write-Host "  [SO LOI DINH DANG] $csvPath | SKU=$($r.sku) | cot '$numCol': '$rawNum' => '$cleanNum'" -ForegroundColor Magenta
                    if ($AutoFix) {
                        $r.$numCol = $cleanNum
                        $fileDirty = $true
                        $totalFixed++
                    } else {
                        $totalWarnings++
                    }
                }
            }
        }

        # 2. Kiem tra & tu dong lam sach HTML rac trong title/artist/brand/label/description/tracklist
        foreach ($col in @('title','artist','brand','label','description','tracklist')) {
            if ($r.PSObject.Properties[$col] -and $r.$col) {
                $orig = [string]$r.$col
                if ($orig -match $htmlPattern) {
                    $clean = Clean-PlainText $orig
                    if ($col -eq 'description' -and $r.PSObject.Properties['title']) {
                        $cleanTitle = Clean-PlainText ([string]$r.title)
                        $cleanArtist = if ($r.PSObject.Properties['artist']) { Clean-PlainText ([string]$r.artist) } else { "" }
                        if ($clean.Length -lt 40 -and $cleanTitle) {
                            $clean = "Đĩa CD $cleanTitle – $cleanArtist. Ấn phẩm chính hãng, hỗ trợ đặt hàng & kiểm tra tồn kho trực tiếp qua Zalo / WhatsApp Hue Sound Express."
                        }
                    }
                    Write-Host "  [HTML RAC] $csvPath | ID=$($r.id) ($($r.sku)) | cot '$col' => '$clean'" -ForegroundColor Magenta
                    if ($AutoFix) {
                        $r.$col = $clean
                        $fileDirty = $true
                        $totalFixed++
                    } else {
                        $totalWarnings++
                    }
                }
            }
        }

        # 3. Canh bao link trang san pham bi dan nham vao cover_url
        if ($r.PSObject.Properties['cover_url'] -and $r.cover_url) {
            $u = ([string]$r.cover_url).Trim()
            if ($u -match '/products/' -and $u -notmatch '\.(jpg|jpeg|png|webp|gif)(\?|$)') {
                Write-Host "  [CANH BAO ANH] $csvPath | ID=$($r.id) ($($r.sku)) | cover_url la link trang web, khong phai link anh truc tiep: $u" -ForegroundColor Red
                $totalWarnings++
            }
        }
    }

    if ($AutoFix -and $fileDirty) {
        Save-CsvUtf8Bom $rows $csvPath
        Write-Host "  [DA SUA] Luu lai file sach: $csvPath" -ForegroundColor Green
    }
}

# 4. Don dep cac thu muc /p/ co slug bi dinh HTML rac tu ban build cu
if (Test-Path '.\p') {
    $badDirs = @(Get-ChildItem '.\p' -Directory | Where-Object { $_.Name -match 'class-price|het-hang|a-href|p-p-|span-|div-' })
    if ($badDirs.Count -gt 0) {
        foreach ($d in $badDirs) {
            if ($AutoFix) {
                Remove-Item -LiteralPath $d.FullName -Recurse -Force
                Write-Host "  [DA XOA SLUG LOI] p/$($d.Name)" -ForegroundColor Green
                $totalFixed++
            } else {
                Write-Host "  [SLUG LOI] p/$($d.Name)" -ForegroundColor Magenta
                $totalWarnings++
            }
        }
    }
}

Write-Host "`n=== KET QUA KIEM TRA DU LIEU CSV ===" -ForegroundColor Cyan
Write-Host "  Da tu dong sua: $totalFixed muc | Canh bao can xem: $totalWarnings muc" -ForegroundColor $(if ($totalWarnings -gt 0) { 'Yellow' } else { 'Green' })
