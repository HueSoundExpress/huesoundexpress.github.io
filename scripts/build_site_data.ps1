<#
.SYNOPSIS
    Hue Sound Express - Single Source of Truth Builder (0d tren GitHub Pages)
    Doc du lieu tu 4 bang (CSV noi bo hoac Google Sheets CSV URL) ->
    1. Hop nhat thanh data/catalog_all.json (4 nhom san pham)
    2. Cap nhat llms.txt chuan B2A (da ap dung quy tac thuong hieu & van hanh moi)
    3. Tu dong sinh trang chi tiet rieng cho tung san pham (/p/<slug>/index.html) & sitemap.xml chuan SEO Google

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\scripts\build_site_data.ps1
#>

param(
    [string]$VinylJsonFallback = "data\catalog_enriched.json",
    [string]$LegacyVinylJson = "catalog.json",
    [string]$VinylCsv = "data\vinyl_lp_enriched.csv",
    [string]$CdCsv = "data\templates\cd_sacd_template.csv",
    [string]$TurntableCsv = "data\templates\turntables_template.csv",
    [string]$AccessoryCsv = "data\templates\accessories_care_template.csv",
    [string]$SiteBaseUrl = "https://huesoundexpress.github.io",
    [switch]$GenerateProductPages
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Remove-Diacritics {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $normalized = $Text.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $normalized.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString().Replace([char]0x0111, 'd').Replace([char]0x0110, 'D')
}

function Get-Slug {
    param([string]$Text)
    $clean = (Remove-Diacritics $Text).ToLowerInvariant()
    $clean = [regex]::Replace($clean, '[^a-z0-9]+', '-')
    return $clean.Trim('-')
}

function Escape-Html {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Load-CsvOrUrl {
    param([string]$Source)
    if ([string]::IsNullOrWhiteSpace($Source)) { return @() }
    if ($Source -match '^https?://') {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $raw = (Invoke-WebRequest -Uri $Source -UseBasicParsing -TimeoutSec 20).Content
        return ($raw | ConvertFrom-Csv)
    }
    if (Test-Path $Source) {
        return @(Import-Csv -Path $Source -Encoding UTF8)
    }
    return @()
}

function Normalize-LegacyStatus {
    param([string]$Cond, [string]$OldStatus, [int]$Price)
    if ($Cond -match '(?i)(sign|mint|limited|picture)' -or $Price -le 0) { return "Liên hệ" }
    if ($OldStatus -eq "Còn hàng" -or $OldStatus -eq "New Seal") { return "New Seal" }
    if ($OldStatus -eq "Sắp có hàng") { return "Sắp có hàng" }
    return "Hết hàng"
}

$allItems = New-Object System.Collections.ArrayList

# 1. NAP DIA THAN (VINYL)
if (Test-Path $VinylCsv) {
    $rows = Load-CsvOrUrl $VinylCsv
    foreach ($r in $rows) {
        $tracks = if ($r.tracklist) { @($r.tracklist -split '\s*\|\s*' | Where-Object { $_ }) } else { @() }
        $missing = if ($r.missing_fields) { @($r.missing_fields -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
        $moods = if ($r.mood) { @($r.mood -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
        [void]$allItems.Add([PSCustomObject]@{
            id             = [int]$r.id
            sku            = [string]$r.sku
            slug           = if ($r.slug) { [string]$r.slug } else { Get-Slug "$($r.artist)-$($r.title)-$($r.id)" }
            category       = "vinyl"
            g              = [string]$r.genre
            mood           = $moods
            a              = [string]$r.artist
            t              = [string]$r.title
            l              = [string]$r.label
            f              = [string]$r.format
            c              = [string]$r.edition
            y              = [string]$r.year
            p              = if ($r.list_price) { [int]$r.list_price } else { 0 }
            s              = [string]$r.status
            d              = [string]$r.description
            tracklist      = $tracks
            img            = [string]$r.cover_url
            aud            = [string]$r.audio_url
            aud_track      = [string]$r.audio_track
            missing_fields = $missing
        })
    }
} elseif (Test-Path $VinylJsonFallback) {
    $rows = Get-Content -Raw -Encoding UTF8 $VinylJsonFallback | ConvertFrom-Json
    foreach ($r in $rows) { [void]$allItems.Add($r) }
} elseif (Test-Path $LegacyVinylJson) {
    $rows = Get-Content -Raw -Encoding UTF8 $LegacyVinylJson | ConvertFrom-Json
    foreach ($r in $rows) {
        $price = [int]$r.p
        $st = Normalize-LegacyStatus -Cond ([string]$r.c) -OldStatus ([string]$r.s) -Price $price
        [void]$allItems.Add([PSCustomObject]@{
            id             = [int]$r.id
            sku            = [string]$r.sku
            slug           = Get-Slug "$($r.a)-$($r.t)-$($r.id)"
            category       = "vinyl"
            g              = [string]$r.g
            mood           = @()
            a              = [string]$r.a
            t              = [string]$r.t
            l              = [string]$r.l
            f              = [string]$r.f
            c              = [string]$r.c
            y              = [string]$r.y
            p              = $price
            s              = $st
            d              = [string]$r.d
            tracklist      = @()
            img            = ""
            aud            = ""
            aud_track      = ""
            missing_fields = @("MISSING_COVER","MISSING_AUDIO","MISSING_TRACKLIST")
        })
    }
}

# 2. NAP DIA CD / SACD
foreach ($r in (Load-CsvOrUrl $CdCsv)) {
    if (-not $r.sku) { continue }
    $tracks = if ($r.tracklist) { @($r.tracklist -split '\s*\|\s*' | Where-Object { $_ }) } else { @() }
    $moods = if ($r.mood) { @($r.mood -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
    [void]$allItems.Add([PSCustomObject]@{
        id             = [int]$r.id
        sku            = [string]$r.sku
        slug           = if ($r.slug) { [string]$r.slug } else { Get-Slug "cd-$($r.artist)-$($r.title)-$($r.id)" }
        category       = "cd"
        g              = [string]$r.genre
        mood           = $moods
        a              = [string]$r.artist
        t              = [string]$r.title
        l              = [string]$r.label
        f              = [string]$r.format
        c              = [string]$r.edition
        y              = [string]$r.year
        p              = if ($r.list_price) { [int]$r.list_price } else { 0 }
        s              = [string]$r.status
        d              = [string]$r.description
        tracklist      = $tracks
        img            = [string]$r.cover_url
        aud            = [string]$r.audio_url
        aud_track      = [string]$r.audio_track
        missing_fields = @()
    })
}

# 3. NAP MAY HAT / MAM THAN
foreach ($r in (Load-CsvOrUrl $TurntableCsv)) {
    if (-not $r.sku) { continue }
    $specs = if ($r.specs_summary) { @($r.specs_summary -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
    $recSkus = if ($r.recommended_skus) { @($r.recommended_skus -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
    [void]$allItems.Add([PSCustomObject]@{
        id               = [int]$r.id
        sku              = [string]$r.sku
        slug             = if ($r.slug) { [string]$r.slug } else { Get-Slug "$($r.brand)-$($r.title)-$($r.id)" }
        category         = "turntable"
        g                = [string]$r.level
        a                = [string]$r.brand
        t                = [string]$r.title
        l                = [string]$r.drive_type
        f                = [string]$r.cartridge_type
        c                = [string]$r.cartridge_model
        has_phono        = [string]$r.has_phono
        has_bluetooth    = [string]$r.has_bluetooth
        specs            = $specs
        recommended_skus = $recSkus
        p                = if ($r.list_price) { [int]$r.list_price } else { 0 }
        s                = [string]$r.status
        d                = [string]$r.description
        img              = [string]$r.cover_url
        video_demo_url   = [string]$r.video_demo_url
    })
}

# 4. NAP PHU KIEN & HUE SOUND VINYL CARE
foreach ($r in (Load-CsvOrUrl $AccessoryCsv)) {
    if (-not $r.sku) { continue }
    [void]$allItems.Add([PSCustomObject]@{
        id            = [int]$r.id
        sku           = [string]$r.sku
        slug          = if ($r.slug) { [string]$r.slug } else { Get-Slug "$($r.brand)-$($r.title)-$($r.id)" }
        category      = if ($r.category) { [string]$r.category } else { "accessory" }
        g             = [string]$r.sub_type
        a             = [string]$r.brand
        t             = [string]$r.title
        l             = [string]$r.level
        f             = [string]$r.compatibility
        c             = [string]$r.sub_type
        p             = if ($r.list_price) { [int]$r.list_price } else { 0 }
        s             = [string]$r.status
        d             = [string]$r.description
        img           = [string]$r.cover_url
    })
}

# Chuan hoa Hang dia, Tinh trang theo Goldmine Grading Standard (M (New seal), NM, VG+, VG, G+, G, F, P) & Trang thai (Con hang / Het hang)
foreach ($it in $allItems) {
    if ($it.category -eq "vinyl" -and $it.id -ge 457 -and $it.id -le 536) {
        $chkTxt = "$($it.a) $($it.t)"
        if ($chkTxt -match '(?i)(totoro|mononoke|howl|kiki|porco rosso|thung lũng gió|lâu đài trên không|vùng đất linh hồn|ghibli)') {
            $it.g = "Classic & Soundtrack"
            $it.a = "Joe Hisaishi (Studio Ghibli)"
            $it.sku = $it.sku -replace '^VN-', 'NN-'
        } elseif ($chkTxt -notmatch '[ăâđêôơưĂÂĐÊÔƠƯáàảãạấầẩẫậắằẳẵặéèẻẽẹếềểễệíìỉĩịóòỏõọốồổỗộớờởỡợúùủũụứừửữựýỳỷỹỵ]') {
            $it.g = "Audiophile & Khác"
            $it.sku = $it.sku -replace '^VN-', 'NN-'
        }
    }

    # Chuan hoa ten hang dia tranh trung lap
    $lbl = ([string]$it.l).Trim()
    if ($lbl -match '^(?i)thời đại$') { $lbl = "Hãng Đĩa Thời Đại" }
    elseif ($lbl -match '^(?i)stockfish$') { $lbl = "Stockfisch Records" }
    elseif ($lbl -match '^(?i)bk production$') { $lbl = "BK Production" }
    elseif ($lbl -match '^(?i)viết tân audio$') { $lbl = "Viết Tân Studio" }
    $it.l = $lbl

    $rawCond = ([string]$it.c).Trim()
    $rawSt = ([string]$it.s).Trim()
    $price = [int]$it.p

    # Goldmine Grading Standard: M (New seal), NM, VG+, VG, G+, G, F, P
    $normCond = "M (New seal)"
    if ($rawCond -match '^(?i)vg\+$') { $normCond = "VG+" }
    elseif ($rawCond -match '^(?i)vg$') { $normCond = "VG" }
    elseif ($rawCond -match '^(?i)g\+$') { $normCond = "G+" }
    elseif ($rawCond -match '^(?i)g$') { $normCond = "G" }
    elseif ($rawCond -match '^(?i)(f|fair)$') { $normCond = "F" }
    elseif ($rawCond -match '^(?i)(p|poor)$') { $normCond = "P" }
    elseif ($rawCond -match '(?i)\b(nm|mint)\b') { $normCond = "NM" }
    else { $normCond = "M (New seal)" }

    # Trang thai kho: Con hang / Het hang
    $normSt = "Còn hàng"
    if ($price -le 0 -or $rawSt -match '(?i)(hết hàng|sắp có|đặt trước)') {
        $normSt = "Hết hàng"
    }

    if ($it.PSObject.Properties['cond']) { $it.cond = $normCond }
    else { $it | Add-Member -NotePropertyName 'cond' -NotePropertyValue $normCond }
    $it.s = $normSt
}

# Xuat data/catalog_all.json
if (-not (Test-Path "data")) { New-Item -ItemType Directory -Path "data" -Force | Out-Null }
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$jsonAll = $allItems | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path (Get-Location) "data\catalog_all.json"), $jsonAll, $utf8NoBom)

# Sinh llms.txt chuan B2A (da cap nhat quy tac thuong hieu & van hanh moi)
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("# HUE SOUND EXPRESS - OFFICIAL AI AGENT & B2A KNOWLEDGE BASE")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 1. Brand Identity & Contact")
[void]$sb.AppendLine("- **Brand Name:** Hue Sound Express")
[void]$sb.AppendLine("- **Care & Technical Service:** Hue Sound Vinyl Care")
[void]$sb.AppendLine("- **Slogan:** Where Sound Finds Its Color")
[void]$sb.AppendLine("- **Positioning:** Chuyên phân phối trực tuyến đĩa than (Vinyl LP), đĩa CD/SACD, máy hát đĩa than (Turntable), phụ kiện audio và dịch vụ chăm sóc đĩa Hue Sound Vinyl Care tại Hà Nội & giao hàng toàn quốc.")
[void]$sb.AppendLine("- **Ordering & Consultation (Zalo / Hotline):** 09323.13918 (https://zalo.me/0932313918)")
[void]$sb.AppendLine("- **Facebook Fanpage:** https://www.facebook.com/HueSoundExpress")
[void]$sb.AppendLine("- **Instagram:** https://www.instagram.com/huesoundexpress/")
[void]$sb.AppendLine("- **Website:** $SiteBaseUrl/")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 2. Operational & Quoting Policies")
[void]$sb.AppendLine("- **Online-First Experience:** Hue Sound Express phục vụ giao hàng hỏa tốc 2H nội thành Hà Nội và ship toàn quốc với quy chuẩn đóng gói 3 lớp chống rách mép (Seam Splits Prevention). Trường hợp khách hàng có nhu cầu gặp trực tiếp để bàn giao đĩa/máy, vui lòng liên hệ Zalo 09323.13918 để được nhân sự hỗ trợ sắp xếp lịch hẹn.")
[void]$sb.AppendLine("- **Pricing, Condition & Status:** Luôn hiển thị Giá niêm yết (nếu giá bằng 0 hiển thị Liên hệ). Tình trạng đĩa phân loại theo chuẩn quốc tế Goldmine Grading Standard: M (New seal), NM, VG+, VG, G+, G, F, P. Trạng thái kho gồm: Còn hàng, Hết hàng.")
[void]$sb.AppendLine("- **Corporate & VIP Gift Quoting:** Hỗ trợ xuất bảng báo giá nhanh phục vụ khách hàng cá nhân/doanh nghiệp làm thủ tục tạm ứng, hỗ trợ gói quà nghệ thuật và viết thiệp tay.")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 3. Hue Sound Vinyl Care - Maintenance Standards")
[void]$sb.AppendLine("1. Cầm đĩa đúng kỹ thuật: chỉ cầm mép ngoài và nhãn trung tâm.")
[void]$sb.AppendLine("2. Lưu trữ thẳng đứng như xếp sách, tránh nhiệt độ cao và ánh nắng trực tiếp.")
[void]$sb.AppendLine("3. Quét bụi và khử tĩnh điện bằng chổi sợi Carbon trước mỗi lần nghe.")
[void]$sb.AppendLine("4. Nâng cấp bao lót chống tĩnh điện (Anti-Static Inner Sleeve) thay cho túi giấy thô.")
[void]$sb.AppendLine("5. Vệ sinh đầu kim định kỳ (Audio-Technica AT607a) và cân chỉnh lực tì kim (Tracking Force) đúng chuẩn.")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 4. Unified Product Catalog ($($allItems.Count) Items)")
[void]$sb.AppendLine("| Category | SKU | Artist / Brand | Product / Album | Format / Spec | Condition | List Price (VND) | Status | URL |")
[void]$sb.AppendLine("| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |")
foreach ($it in $allItems) {
    $priceTxt = if ($it.p -gt 0) { $it.p.ToString("N0") } else { "Liên hệ" }
    $urlTxt = "$SiteBaseUrl/p/$($it.slug)/"
    [void]$sb.AppendLine("| $($it.category) | $($it.sku) | $($it.a) | $($it.t) | $($it.f) | $($it.cond) | $priceTxt | $($it.s) | $urlTxt |")
}
[System.IO.File]::WriteAllText((Join-Path (Get-Location) "llms.txt"), $sb.ToString(), $utf8NoBom)

# Tuy chon: Sinh trang chi tiet tinh /p/<slug>/index.html & sitemap.xml
if ($GenerateProductPages) {
    $pRoot = Join-Path (Get-Location) "p"
    if (-not (Test-Path $pRoot)) { New-Item -ItemType Directory -Path $pRoot -Force | Out-Null }

    $sm = New-Object System.Text.StringBuilder
    [void]$sm.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    [void]$sm.AppendLine('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">')
    [void]$sm.AppendLine("  <url><loc>$SiteBaseUrl/</loc><priority>1.0</priority></url>")

    $accLookup = $allItems | Where-Object { $_.category -eq "accessory" -or $_.category -eq "care" }

    foreach ($it in $allItems) {
        $dir = Join-Path $pRoot $it.slug
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        $titleEsc = Escape-Html "$($it.a) – $($it.t)"
        $descEsc = Escape-Html $it.d
        $priceFmt = if ($it.p -gt 0) { "$($it.p.ToString('N0')) đ" } else { "Liên hệ báo giá" }
        $pageUrl = "$SiteBaseUrl/p/$($it.slug)/"
        $ogImg = if ($it.img -match '^https?://') { $it.img } else { "$SiteBaseUrl/og-image.jpg" }
        $catLabel = switch ($it.category) {
            "vinyl" { "Đĩa Than (Vinyl LP)" }
            "cd" { "Đĩa CD / SACD" }
            "turntable" { "Máy Hát / Mâm Than" }
            "care" { "Hue Sound Vinyl Care" }
            default { "Phụ Kiện Audio" }
        }

        $trackHtml = ""
        if ($it.tracklist -and $it.tracklist.Count -gt 0) {
            $li = ($it.tracklist | ForEach-Object { "<li>$(Escape-Html $_)</li>" }) -join ""
            $trackHtml = "<section class='box'><h2>Danh sách bài hát (Tracklist)</h2><ol class='tracks'>$li</ol></section>"
        }

        $specHtml = ""
        if ($it.specs -and $it.specs.Count -gt 0) {
            $li = ($it.specs | ForEach-Object { "<li>$(Escape-Html $_)</li>" }) -join ""
            $specHtml = "<section class='box'><h2>Thông số nổi bật</h2><ul class='tracks'>$li</ul></section>"
        }

        $audioHtml = ""
        if ($it.aud -match '^https?://') {
            $audioHtml = "<div class='audio-box'><div class='small'>Nghe thử 30s: <strong>$(Escape-Html $it.aud_track)</strong></div><audio controls preload='none' src='$(Escape-Html $it.aud)' style='width:100%;margin-top:8px;'></audio></div>"
        }

        # Goi y mua kem thong minh
        $upsellItems = @()
        if ($it.category -eq "turntable") {
            $upsellItems = $accLookup | Where-Object { $_.sku -in @("ACC-AQ-CARBON","ACC-AT607A","ACC-SCALE-01","CARE-SETUP-01") } | Select-Object -First 3
        } elseif ($it.category -eq "vinyl") {
            $upsellItems = $accLookup | Where-Object { $_.sku -in @("ACC-SLEEVE-50","ACC-AQ-CARBON","CARE-ULTRASONIC-01") } | Select-Object -First 3
        }
        $upsellHtml = ""
        if ($upsellItems.Count -gt 0) {
            $cards = ($upsellItems | ForEach-Object {
                "<a class='up-card' href='../$($_.slug)/'><strong>$(Escape-Html $_.t)</strong><span>$($_.p.ToString('N0')) đ</span></a>"
            }) -join ""
            $upsellHtml = "<section class='box'><h2>Gợi ý đi kèm từ Hue Sound Vinyl Care</h2><div class='up-grid'>$cards</div></section>"
        }

        $zaloText = [Uri]::EscapeDataString("Chào Hue Sound Express, mình quan tâm sản phẩm: $($it.a) - $($it.t) (Mã: $($it.sku) - Giá niêm yết: $priceFmt)")
        $ctaLabel = if ($it.s -eq "Hết hàng" -or $it.s -eq "Sắp có hàng") { "🔔 Đăng ký báo khi hàng về" } else { "💬 Đặt mua / Báo giá qua Zalo" }
        $ctaHref = if ($it.s -eq "Hết hàng" -or $it.s -eq "Sắp có hàng") { "../../?waitlist=$($it.sku)" } else { "https://zalo.me/0932313918?text=$zaloText" }

        $schemaAvail = if ($it.s -eq "Còn hàng") { "https://schema.org/InStock" } else { "https://schema.org/OutOfStock" }

        $html = @"
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>$titleEsc | Hue Sound Express</title>
  <meta name="description" content="$descEsc">
  <meta name="theme-color" content="#E8D9B6">
  <link rel="canonical" href="$pageUrl">
  <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Ccircle cx='32' cy='32' r='30' fill='%23EADBB8'/%3E%3Ccircle cx='32' cy='32' r='22' fill='none' stroke='%23B5541F' stroke-width='3'/%3E%3Ccircle cx='32' cy='32' r='14' fill='none' stroke='%23C99A2E' stroke-width='2'/%3E%3Ccircle cx='32' cy='32' r='5' fill='%23B5541F'/%3E%3C/svg%3E">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,500;0,600;1,400&family=Montserrat:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <meta property="og:title" content="$titleEsc – $priceFmt | Hue Sound Express">
  <meta property="og:description" content="$descEsc">
  <meta property="og:image" content="$ogImg">
  <meta property="og:url" content="$pageUrl">
  <meta property="og:type" content="product">
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@type": "Product",
    "name": "$($it.a -replace '"','\"') - $($it.t -replace '"','\"')",
    "sku": "$($it.sku -replace '"','\"')",
    "image": "$ogImg",
    "description": "$($it.d -replace '"','\"')",
    "brand": { "@type": "Brand", "name": "$($it.l -replace '"','\"')" },
    "offers": {
      "@type": "Offer",
      "priceCurrency": "VND",
      "price": "$($it.p)",
      "availability": "$schemaAvail",
      "url": "$pageUrl"
    }
  }
  </script>
  <style>
    :root {
      --paper:#E7D5AE; --paper-light:#F4E9CC; --copper:#B5541F; --copper-dark:#7E3210;
      --copper-grad:linear-gradient(135deg, #DC8644 0%, #B85922 46%, #8C3A12 100%);
      --gold:#CFA042; --ink:#4A2312; --muted:#6F5139;
      --card:rgba(250,242,222,.82); --border:rgba(142,66,24,.27);
      --sans:"Montserrat",system-ui,-apple-system,sans-serif; --serif:"Lora",Georgia,serif;
    }
    * { box-sizing:border-box; margin:0; padding:0; }
    body {
      font-family:var(--sans); color:var(--ink); line-height:1.6;
      background-color:#C9B184;
      background-image:radial-gradient(ellipse at 50% 18%, #F4E9CC 0%, #E7D5AE 44%, #C5AC7E 82%, #AA9062 100%);
      background-attachment:fixed; padding:24px 16px 64px; position:relative;
    }
    body::before {
      content:''; position:fixed; inset:0; z-index:0; pointer-events:none; opacity:.56; mix-blend-mode:multiply;
      background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='220' height='220'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.88' numOctaves='3' stitchTiles='stitch'/%3E%3CfeColorMatrix values='0 0 0 0 .42 0 0 0 0 .30 0 0 0 0 .16 0 0 0 .25 0'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E");
    }
    body::after {
      content:''; position:fixed; inset:0; z-index:1; pointer-events:none;
      background:radial-gradient(ellipse at 50% 42%, transparent 56%, rgba(84,52,22,.24) 100%);
    }
    .wrap { max-width:900px; margin:0 auto; position:relative; z-index:2; }
    .topbar { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:20px; }
    .back { display:inline-flex; align-items:center; gap:6px; color:var(--copper-dark); text-decoration:none; font-size:13px; font-weight:600; background:var(--card); border:1px solid var(--border); padding:8px 14px; border-radius:99px; }
    .back:hover { background:#FFFBF0; border-color:var(--copper); }
    .brand-mini { font-weight:800; font-size:14px; color:var(--copper); letter-spacing:.04em; }
    .hero { display:grid; grid-template-columns:290px 1fr; gap:26px; background:var(--card); border:1px solid var(--border); border-radius:18px; padding:24px; backdrop-filter:blur(8px); box-shadow:0 12px 30px -12px rgba(110,60,20,.22); }
    @media (max-width:680px){ .hero{ grid-template-columns:1fr; } }
    .cover { aspect-ratio:1/1; border-radius:14px; overflow:hidden; background:radial-gradient(circle at center, #F3E9D0 0%, #DFC99B 100%); border:1px solid var(--border); display:flex; align-items:center; justify-content:center; text-align:center; padding:16px; position:relative; }
    .cover img { width:100%; height:100%; object-fit:cover; }
    .cover-art { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:8px; color:var(--ink); }
    .badge { display:inline-block; font-size:11px; font-weight:600; padding:4px 11px; border-radius:99px; background:rgba(181,84,31,.11); color:var(--copper-dark); border:1px solid rgba(181,84,31,.22); margin-bottom:10px; }
    h1 { font-family:var(--serif); font-size:26px; line-height:1.3; margin-bottom:4px; color:var(--ink); }
    .sub { color:var(--copper-dark); font-size:16px; font-weight:600; margin-bottom:10px; }
    .meta { font-size:12.5px; color:var(--muted); margin-bottom:14px; }
    .desc { font-size:14px; color:var(--ink); opacity:.9; }
    .price-row { display:flex; align-items:center; flex-wrap:wrap; gap:12px; margin:16px 0; }
    .price-label { font-size:11px; text-transform:uppercase; letter-spacing:.06em; color:var(--muted); display:block; }
    .price { font-size:25px; font-weight:800; color:var(--copper); }
    .status { font-size:12px; font-weight:600; padding:4px 11px; border-radius:99px; background:rgba(201,154,46,.18); border:1px solid rgba(201,154,46,.5); color:var(--ink); }
    .cta-row { display:flex; flex-wrap:wrap; gap:10px; margin-top:16px; }
    .btn { text-decoration:none; padding:11px 18px; border-radius:10px; font-weight:600; font-size:13.5px; display:inline-flex; align-items:center; gap:6px; transition:all .2s; }
    .btn-copper { background:var(--copper-grad); color:#FFF7E8; box-shadow:inset 0 1px 0 rgba(244,219,141,.4), 0 6px 16px -6px rgba(126,50,16,.5); }
    .btn-copper:hover { transform:translateY(-1px); background:linear-gradient(135deg, #E59654 0%, #C66429 48%, #9A4216 100%); }
    .btn-outline { border:1px solid var(--copper); color:var(--copper-dark); background:rgba(255,251,240,.6); }
    .btn-outline:hover { background:var(--copper); color:#FFF6E6; }
    .audio-box { margin-top:16px; padding:12px 14px; background:rgba(243,233,208,.8); border-radius:10px; border:1px solid var(--border); }
    .small { font-size:12px; color:var(--muted); }
    .box { margin-top:18px; background:var(--card); border:1px solid var(--border); border-radius:16px; padding:20px; backdrop-filter:blur(6px); }
    .box h2 { font-family:var(--serif); font-size:17px; color:var(--copper-dark); margin-bottom:12px; }
    .tracks { padding-left:20px; color:var(--ink); font-size:13.5px; columns:2; gap:26px; }
    .tracks li { margin-bottom:5px; }
    @media (max-width:600px){ .tracks{ columns:1; } }
    .up-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:12px; }
    .up-card { text-decoration:none; background:rgba(255,251,240,.75); border:1px solid var(--border); padding:14px; border-radius:12px; color:var(--ink); display:flex; flex-direction:column; gap:6px; transition:all .2s; }
    .up-card:hover { border-color:var(--copper); transform:translateY(-2px); box-shadow:0 8px 18px -10px rgba(143,62,20,.35); }
    .up-card span { color:var(--copper); font-weight:700; font-size:13.5px; }
    footer { text-align:center; margin-top:34px; font-size:12px; color:var(--muted); }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="topbar">
      <a class="back" href="../../">← Quay lại Phòng Nghe Số Hue Sound Express</a>
      <span class="brand-mini">HUE SOUND EXPRESS • Where sound finds its color</span>
    </div>
    <div class="hero">
      <div class="cover">
        $(if ($it.img -match '^https?://') { "<img src='$(Escape-Html $it.img)' alt='$titleEsc'>" } else { "<div class='cover-art'><svg width='96' height='96' viewBox='0 0 100 100'><circle cx='50' cy='50' r='44' fill='none' stroke='#B5541F' stroke-width='2' opacity='.5'/><circle cx='50' cy='50' r='34' fill='none' stroke='#C99A2E' stroke-width='1.5' opacity='.6'/><circle cx='50' cy='50' r='24' fill='none' stroke='#B5541F' stroke-width='1.5' opacity='.5'/><circle cx='50' cy='50' r='9' fill='#B5541F'/></svg><div style='font-weight:700;font-size:13px;'>$(Escape-Html $it.a)</div><div style='font-size:12px;color:#7A5C46;'>$(Escape-Html $it.t)</div></div>" })
      </div>
      <div>
        <span class="badge">$catLabel • $(Escape-Html $it.g)</span>
        <h1>$(Escape-Html $it.t)</h1>
        <div class="sub">$(Escape-Html $it.a)</div>
        <div class="meta">Mã SKU: <strong>$(Escape-Html $it.sku)</strong> • Định dạng: $(Escape-Html $it.f) $(if ($it.l) { "• $(Escape-Html $it.l)" }) $(if ($it.y) { "• Năm $($it.y)" })</div>
        <p class="desc">$descEsc</p>
        <div class="price-row">
          <div><span class="price-label">Giá niêm yết</span><div class="price">$priceFmt</div></div>
          <span class="status">$(Escape-Html $it.s)</span>
        </div>
        <div class="cta-row">
          <a class="btn btn-copper" href="../../?add=$([Uri]::EscapeDataString($it.sku))">+ Thêm vào Giỏ / Báo giá</a>
          <a class="btn btn-outline" href="$ctaHref" target="_blank">$ctaLabel</a>
        </div>
        $audioHtml
      </div>
    </div>
    $trackHtml
    $specHtml
    $upsellHtml
    <footer>© 2026 Hue Sound Express • Đóng gói 3 lớp chống rách mép • Hotline/Zalo: 09323.13918</footer>
  </div>
</body>
</html>
"@
        [System.IO.File]::WriteAllText((Join-Path $dir "index.html"), $html, $utf8NoBom)
        [void]$sm.AppendLine("  <url><loc>$pageUrl</loc><priority>0.8</priority></url>")
    }

    [void]$sm.AppendLine('</urlset>')
    [System.IO.File]::WriteAllText((Join-Path (Get-Location) "sitemap.xml"), $sm.ToString(), $utf8NoBom)
    Write-Host "Da sinh $($allItems.Count) trang chi tiet SEO trong /p/ va sitemap.xml!" -ForegroundColor Green
}

Write-Host "Build hoan tat! Tong so san pham hop nhat: $($allItems.Count)" -ForegroundColor Green
Write-Host " - data\catalog_all.json"
Write-Host " - llms.txt"
