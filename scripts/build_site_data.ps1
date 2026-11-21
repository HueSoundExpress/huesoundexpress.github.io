<#
.SYNOPSIS
    Hue Sound Express - Single Source of Truth Builder (0d tren GitHub Pages)
    Doc du lieu tu 4 bang CSV ->
    1. Hop nhat thanh data/catalog_all.json (chuan hoa ASCII SKU, Slug co tien to + legacy_slug, gia giam 50-60k)
    2. Xuat feed/products.json (chuan OpenAI / Google Shopping Feed 0d cho AI Agent)
    3. Cap nhat llms.txt chuan B2A (dan truc tiep toi feed/products.json & /cam-nang/)
    4. Sinh trang chi tiet SEO /p/<slug>/index.html (BreadcrumbList + Product Offer shippingDetails/returnPolicy/itemCondition + FAQPage rieng theo nhom + Redirect stub cho URL cu)
    5. Sinh 6 bai viet tinh /cam-nang/<slug>/index.html tu dong cap nhat san pham lien quan & sitemap.xml
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
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

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

function Escape-JsonStr {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    return ($Text -replace '\\', '\\' -replace '"', '\"' -replace "`r", ' ' -replace "`n", ' ')
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

function Normalize-AsciiSku {
    param([string]$RawSku, [string]$FallbackPrefix, [int]$Id, [System.Collections.Generic.HashSet[string]]$Used)
    $s = (Remove-Diacritics $RawSku).ToUpperInvariant()
    $s = [regex]::Replace($s, '[^A-Z0-9]+', '-').Trim('-')
    if (-not $s -or $s.Length -lt 3) { $s = "$FallbackPrefix-$Id" }
    $cand = $s
    $n = 2
    while ($Used.Contains($cand)) {
        $cand = "$s-$Id"
        if ($Used.Contains($cand)) { $cand = "$s-$Id-$n"; $n++ }
    }
    [void]$Used.Add($cand)
    return $cand
}

function Get-CanonicalSlug {
    param([string]$Category, [string]$Artist, [string]$Title, [int]$Id)
    $prefix = switch ($Category) {
        "vinyl"     { "dia-than-vinyl" }
        "cd"        { "dia-cd" }
        "turntable" { "may-hat-dia-than" }
        "care"      { "dich-vu-vinyl-care" }
        default     { "phu-kien-analog" }
    }
    $body = Get-Slug "$Artist-$Title"
    if ($body.Length -gt 64) { $body = $body.Substring(0, 64).Trim('-') }
    return "$prefix-$body-$Id"
}

function Get-GenreDisplayVi {
    param([string]$Genre)
    switch ($Genre) {
        "Nhạc Việt"          { return "Nhạc Việt trẻ" }
        "Nhạc Việt trẻ"      { return "Nhạc Việt trẻ" }
        "Nhạc Việt / Bolero" { return "Trữ tình" }
        "Tình khúc 1954–1975" { return "Tình khúc 1954–1975" }
        "Dân ca & Cách mạng" { return "Dân ca & Cách mạng" }
        default              { return $Genre }
    }
}

$allItems = New-Object System.Collections.ArrayList
$usedSkus = New-Object 'System.Collections.Generic.HashSet[string]'

# 1. NAP DIA THAN (VINYL)
if (Test-Path $VinylCsv) {
    $rows = Load-CsvOrUrl $VinylCsv
    foreach ($r in $rows) {
        $idNum = [int]$r.id
        $tracks = if ($r.tracklist) { @($r.tracklist -split '\s*\|\s*' | Where-Object { $_ }) } else { @() }
        $missing = if ($r.missing_fields) { @($r.missing_fields -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
        $moods = if ($r.mood) { @($r.mood -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
        $legacySlug = if ($r.PSObject.Properties['legacy_slug'] -and $r.legacy_slug) { [string]$r.legacy_slug } elseif ($r.slug) { [string]$r.slug } else { "" }
        [void]$allItems.Add([PSCustomObject]@{
            id             = $idNum
            sku            = Normalize-AsciiSku -RawSku ([string]$r.sku) -FallbackPrefix "VN" -Id $idNum -Used $usedSkus
            slug           = Get-CanonicalSlug -Category "vinyl" -Artist ([string]$r.artist) -Title ([string]$r.title) -Id $idNum
            legacy_slug    = $legacySlug
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
            contact_only   = $false
            d              = [string]$r.description
            tracklist      = $tracks
            img            = [string]$r.cover_url
            aud            = [string]$r.audio_url
            aud_track      = [string]$r.audio_track
            missing_fields = $missing
        })
    }
}

# 2. NAP DIA CD / SACD
foreach ($r in (Load-CsvOrUrl $CdCsv)) {
    if (-not $r.sku) { continue }
    $idNum = [int]$r.id
    $tracks = if ($r.tracklist) { @($r.tracklist -split '\s*\|\s*' | Where-Object { $_ }) } else { @() }
    $moods = if ($r.mood) { @($r.mood -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
    $legacySlug = if ($r.PSObject.Properties['legacy_slug'] -and $r.legacy_slug) { [string]$r.legacy_slug } elseif ($r.slug) { [string]$r.slug } else { "" }
    $isContact = ($r.PSObject.Properties['contact_only'] -and ($r.contact_only -eq "1" -or $r.contact_only -eq "true"))
    [void]$allItems.Add([PSCustomObject]@{
        id             = $idNum
        sku            = Normalize-AsciiSku -RawSku ([string]$r.sku) -FallbackPrefix "CD-VN" -Id $idNum -Used $usedSkus
        slug           = Get-CanonicalSlug -Category "cd" -Artist ([string]$r.artist) -Title ([string]$r.title) -Id $idNum
        legacy_slug    = $legacySlug
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
        contact_only   = $isContact
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
    $idNum = [int]$r.id
    $specs = if ($r.specs_summary) { @($r.specs_summary -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
    $recSkus = if ($r.recommended_skus) { @($r.recommended_skus -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
    $legacySlug = if ($r.PSObject.Properties['legacy_slug'] -and $r.legacy_slug) { [string]$r.legacy_slug } elseif ($r.slug) { [string]$r.slug } else { "" }
    [void]$allItems.Add([PSCustomObject]@{
        id               = $idNum
        sku              = Normalize-AsciiSku -RawSku ([string]$r.sku) -FallbackPrefix "TT" -Id $idNum -Used $usedSkus
        slug             = Get-CanonicalSlug -Category "turntable" -Artist ([string]$r.brand) -Title ([string]$r.title) -Id $idNum
        legacy_slug      = $legacySlug
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
        contact_only     = $false
        d                = [string]$r.description
        img              = [string]$r.cover_url
        video_demo_url   = [string]$r.video_demo_url
    })
}

# 4. NAP PHU KIEN & HUE SOUND VINYL CARE
foreach ($r in (Load-CsvOrUrl $AccessoryCsv)) {
    if (-not $r.sku) { continue }
    $idNum = [int]$r.id
    $cat = if ($r.category) { [string]$r.category } else { "accessory" }
    $legacySlug = if ($r.PSObject.Properties['legacy_slug'] -and $r.legacy_slug) { [string]$r.legacy_slug } elseif ($r.slug) { [string]$r.slug } else { "" }
    [void]$allItems.Add([PSCustomObject]@{
        id            = $idNum
        sku           = Normalize-AsciiSku -RawSku ([string]$r.sku) -FallbackPrefix "ACC" -Id $idNum -Used $usedSkus
        slug          = Get-CanonicalSlug -Category $cat -Artist ([string]$r.brand) -Title ([string]$r.title) -Id $idNum
        legacy_slug   = $legacySlug
        category      = $cat
        g             = [string]$r.sub_type
        a             = [string]$r.brand
        t             = [string]$r.title
        l             = [string]$r.level
        f             = [string]$r.compatibility
        c             = [string]$r.sub_type
        p             = if ($r.list_price) { [int]$r.list_price } else { 0 }
        s             = [string]$r.status
        contact_only  = $false
        d             = [string]$r.description
        img           = [string]$r.cover_url
    })
}

# Chuan hoa Hang dia, Nghe si, Tinh trang Goldmine, Trang thai & Hieu ung giam gia
foreach ($it in $allItems) {
    if ($it.category -eq "vinyl" -and $it.id -ge 457 -and $it.id -le 536) {
        $chkTxt = "$($it.a) $($it.t)"
        if ($chkTxt -match '(?i)(totoro|mononoke|howl|kiki|porco rosso|thung lũng gió|lâu đài trên không|vùng đất linh hồn|ghibli)') {
            $it.g = "Classic & Soundtrack"
            $it.a = "Joe Hisaishi (Studio Ghibli)"
            $it.sku = $it.sku -replace '^VN-', 'NN-'
        } elseif ($chkTxt -notmatch '[ăâđêôơưĂÂĐÊÔƠƯáàảãạấầẩẫậắằẳẵặéèẻẽẹếềểễệíìỉĩịóòỏõọốồổỗộớờởỡợúùủũụứừửữựýỳỷỹỵ]') {
            if ($it.g -notin @('Classic & Soundtrack','Pop & Disco','Rock & Metal','Jazz & Vocal','Instrument')) {
                $it.g = "Audiophile & Khác"
            }
            $it.sku = $it.sku -replace '^VN-', 'NN-'
        }
    }

    if (([string]$it.a).Trim() -match '^(?i)nhiều ca sĩ$') { $it.a = "Nhiều nghệ sĩ" }
    if ($it.g -eq "Nhạc Việt") { $it.g = "Nhạc Việt trẻ" }

    $lbl = ([string]$it.l).Trim()
    if ($lbl -match '^(?i)thời đại$') { $lbl = "Hãng Đĩa Thời Đại" }
    elseif ($lbl -match '^(?i)stockfish$') { $lbl = "Stockfisch Records" }
    elseif ($lbl -match '^(?i)bk production$') { $lbl = "BK Production" }
    elseif ($lbl -match '^(?i)viết tân audio$') { $lbl = "Viết Tân Studio" }
    $it.l = $lbl

    $rawCond = ([string]$it.c).Trim()
    $rawSt = ([string]$it.s).Trim()
    $price = [int]$it.p

    $normCond = "M (New seal)"
    if ($rawCond -match '^(?i)vg\+$') { $normCond = "VG+" }
    elseif ($rawCond -match '^(?i)vg$') { $normCond = "VG" }
    elseif ($rawCond -match '^(?i)g\+$') { $normCond = "G+" }
    elseif ($rawCond -match '^(?i)g$') { $normCond = "G" }
    elseif ($rawCond -match '^(?i)(f|fair)$') { $normCond = "F" }
    elseif ($rawCond -match '^(?i)(p|poor)$') { $normCond = "P" }
    elseif ($rawCond -match '(?i)\b(nm|mint)\b') { $normCond = "NM" }
    else { $normCond = "M (New seal)" }

    $normSt = "Còn hàng"
    if ($price -le 0 -or $rawSt -match '(?i)(hết hàng|sắp có|đặt trước)') {
        $normSt = "Hết hàng"
    }

    if ($it.PSObject.Properties['cond']) { $it.cond = $normCond }
    else { $it | Add-Member -NotePropertyName 'cond' -NotePropertyValue $normCond }
    $it.s = $normSt

    # Hieu ung marketing giam gia 50.000 - 60.000d so voi Gia goc (20k cho CD 100-199k, 10k cho DV < 100k)
    $origP = [int]$it.p
    $saleP = $origP
    if ($origP -ge 200000) {
        $disc = if (($origP % 100000) -eq 50000) { 60000 } else { 50000 }
        $saleP = $origP - $disc
    } elseif ($origP -ge 100000) {
        $saleP = $origP - 20000
    } elseif ($origP -gt 10000) {
        $saleP = $origP - 10000
    }
    if ($it.PSObject.Properties['p_orig']) { $it.p_orig = $origP }
    else { $it | Add-Member -NotePropertyName 'p_orig' -NotePropertyValue $origP }
    $it.p = $saleP
}

# Sap xep uu tien san pham moi them len tren cung (giu nguyen thu tu danh sach cu)
$orderedItems = New-Object System.Collections.ArrayList
$catOrderRules = @(
    @{ cats = @('vinyl');            baseMax = 536 },
    @{ cats = @('cd');               baseMax = 1600 },
    @{ cats = @('turntable');        baseMax = 2004 },
    @{ cats = @('accessory','care'); baseMax = 3007 }
)
foreach ($rule in $catOrderRules) {
    $group = @($allItems | Where-Object { $rule.cats -contains $_.category })
    $newItems = @($group | Where-Object { $_.id -gt $rule.baseMax } | Sort-Object -Property id -Descending)
    $oldItems = @($group | Where-Object { $_.id -le $rule.baseMax })
    foreach ($x in $newItems) { [void]$orderedItems.Add($x) }
    foreach ($x in $oldItems) { [void]$orderedItems.Add($x) }
}
$allItems = $orderedItems

# 1. Xuat data/catalog_all.json
if (-not (Test-Path "data")) { New-Item -ItemType Directory -Path "data" -Force | Out-Null }
$jsonAll = $allItems | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path (Get-Location) "data\catalog_all.json"), $jsonAll, $utf8NoBom)

# 2. Xuat feed/products.json (Chuan OpenAI / Google Shopping Feed 0d cho AI Agent)
if (-not (Test-Path "feed")) { New-Item -ItemType Directory -Path "feed" -Force | Out-Null }
$feedProducts = New-Object System.Collections.ArrayList
$priceValidUntil = (Get-Date).AddDays(90).ToString("yyyy-MM-dd")

foreach ($it in $allItems) {
    $seoPrefix = switch ($it.category) {
        "vinyl"     { "Đĩa than (Vinyl) " }
        "cd"        { "Đĩa CD / SACD " }
        "turntable" { "Máy hát đĩa than (Turntable) " }
        default     { "Analog Audio – " }
    }
    $pageUrl = "$SiteBaseUrl/p/$($it.slug)/"
    $condVi = if ($it.cond -eq "M (New seal)") { "Mới 100% (M)" } else { $it.cond }
    $priceFmt = if ($it.p -gt 0) { "$($it.p.ToString('N0')) đ" } else { "Liên hệ" }
    $chatRaw = "Chào Hue Sound Express, mình quan tâm: $seoPrefix$($it.a) - $($it.t) (SKU: $($it.sku) | $condVi | Giá: $priceFmt). Link: $pageUrl"
    $waUrl = "https://wa.me/84932313918?text=$([Uri]::EscapeDataString($chatRaw))"
    $gKw = @("Analog Audio", (Get-GenreDisplayVi $it.g))
    if ($it.g -eq "Nhạc Việt / Bolero") { $gKw += @("Bolero", "Nhạc Vàng", "Nhạc Trữ Tình") }
    elseif ($it.g -eq "Tình khúc 1954–1975") { $gKw += @("Tình khúc 1954-1975", "Tình ca", "Nhạc Tiền Chiến") }
    elseif ($it.g -eq "Instrument") { $gKw += @("Instrument", "Hòa tấu nhạc cụ") }
    if ($it.category -eq "vinyl") { $gKw += @("Đĩa than", "Vinyl LP") }

    [void]$feedProducts.Add([PSCustomObject]@{
        id                    = $it.id
        sku                   = $it.sku
        title                 = $it.t
        seo_title             = "$seoPrefix$($it.t) – $($it.a)"
        artist_or_brand       = $it.a
        category              = $it.category
        genre                 = (Get-GenreDisplayVi $it.g)
        genre_keywords        = $gKw
        label                 = $it.l
        format                = $it.f
        year                  = $it.y
        condition_goldmine    = $it.cond
        condition_vi          = $condVi
        condition_schema      = if ($it.cond -eq "M (New seal)") { "https://schema.org/NewCondition" } else { "https://schema.org/UsedCondition" }
        availability          = if ($it.s -eq "Còn hàng") { "in_stock" } else { "out_of_stock" }
        order_mode            = if ($it.contact_only) { "contact" } else { "cart" }
        list_price_vnd        = $it.p_orig
        sale_price_vnd        = $it.p
        discount_vnd          = [Math]::Max(0, $it.p_orig - $it.p)
        accessory_bundle_only = ($it.category -eq "accessory" -or $it.category -eq "care")
        url                   = $pageUrl
        image_url             = if ($it.img -match '^https?://') { $it.img } else { "$SiteBaseUrl/og-image.jpg" }
        audio_preview_url     = if ($it.aud -match '^https?://') { $it.aud } else { "" }
        zalo_url              = "https://zalo.me/0932313918"
        whatsapp_order_url    = $waUrl
    })
}

$guideMeta = @(
    @{ slug = "bao-quan-dia-than-thoi-tiet-viet-nam";      title = "Cách bảo quản đĩa than với thời tiết nóng ẩm ở Việt Nam" },
    @{ slug = "dia-than-moi-co-can-rua-truoc-khi-nghe";    title = "Đĩa than mới mua (Mới 100% / New Seal) có cần rửa trước khi nghe không?" },
    @{ slug = "dong-goi-dia-than-chong-rach-mep-di-tinh";  title = "Đóng gói đĩa than gửi đi tỉnh có sợ cong vênh hay rách mép (Seam Splits) không?" },
    @{ slug = "cach-chon-mam-than-dau-tien";               title = "Cách chọn mâm đĩa than (Turntable) đầu tiên cho người mới bắt đầu" },
    @{ slug = "phan-biet-kim-mm-va-mc";                    title = "Phân biệt kim đĩa than MM và MC – Nên chọn loại nào cho dàn Analog Audio?" },
    @{ slug = "giai-thich-thang-goldmine-grading";         title = "Giải thích thang đánh giá tình trạng đĩa Goldmine Grading Standard (M, NM, VG+, VG...)" }
)

$feedPayload = [PSCustomObject]@{
    schema_version      = "1.0"
    store               = "Hue Sound Express"
    care_service        = "Hue Sound Vinyl Care"
    slogan              = "Where sound finds its color"
    updated_at          = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss+07:00")
    currency            = "VND"
    vat_note            = "Giá báo chưa bao gồm thuế VAT 10%. Bảng báo giá không phải hóa đơn VAT."
    accessory_promo     = "Chỉ giảm giá phụ kiện khi mua cùng đĩa than / mâm than."
    contact             = [PSCustomObject]@{
        hotline_zalo = "09323.13918"
        zalo_url     = "https://zalo.me/0932313918"
        whatsapp_url = "https://wa.me/84932313918"
        website      = "$SiteBaseUrl/"
    }
    total_products      = $feedProducts.Count
    guides              = @($guideMeta | ForEach-Object { [PSCustomObject]@{ title = $_.title; url = "$SiteBaseUrl/cam-nang/$($_.slug)/" } })
    products            = $feedProducts
}
[System.IO.File]::WriteAllText((Join-Path (Get-Location) "feed\products.json"), ($feedPayload | ConvertTo-Json -Depth 6), $utf8NoBom)

# 3. Sinh llms.txt chuan B2A (dan truc tiep toi feed/products.json & /cam-nang/)
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("# HUE SOUND EXPRESS - OFFICIAL AI AGENT & B2A KNOWLEDGE BASE")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 1. Brand Identity, Shopping Feed & Contact")
[void]$sb.AppendLine("- **Brand Name:** Hue Sound Express")
[void]$sb.AppendLine("- **Care & Technical Service:** Hue Sound Vinyl Care")
[void]$sb.AppendLine("- **Slogan:** Where Sound Finds Its Color")
[void]$sb.AppendLine("- **Machine-Readable Shopping Feed (1-Call JSON):** $SiteBaseUrl/feed/products.json (chứa đầy đủ tồn kho, giá gốc/giá giảm, thể loại Nhạc vàng / Bolero / Tình khúc 1954–1975 / Nhạc Việt trẻ / Instrument / Analog Audio và link đặt Zalo/WhatsApp)")
[void]$sb.AppendLine("- **Positioning:** Chuyên phân phối trực tuyến **Analog Audio** chính hãng: Đĩa than (Vinyl LP), Đĩa CD/SACD, Máy hát đĩa than (Turntable), Phụ kiện & Dịch vụ Hue Sound Vinyl Care tại Hà Nội & giao hàng toàn quốc.")
[void]$sb.AppendLine("- **Vietnamese Music Taxonomy:** Tình khúc 1954–1975 (Phạm Duy, Trịnh Công Sơn, Ngô Thụy Miên, Vũ Thành An, Từ Công Phụng, Lê Uyên Phương, Nguyễn Ánh 9), Trữ tình (Nhạc Bolero / Nhạc Vàng), Nhạc Việt trẻ, Dân ca & Cách mạng, Instrument (Hòa tấu nhạc cụ).")
[void]$sb.AppendLine("- **Ordering & Consultation (Zalo / Hotline / WhatsApp):** 09323.13918 (https://zalo.me/0932313918 | https://wa.me/84932313918)")
[void]$sb.AppendLine("- **Facebook Fanpage:** https://www.facebook.com/HueSoundExpress")
[void]$sb.AppendLine("- **Instagram:** https://www.instagram.com/huesoundexpress/")
[void]$sb.AppendLine("- **Website:** $SiteBaseUrl/")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 2. Operational, Pricing & Quoting Policies")
[void]$sb.AppendLine("- **Online-First Experience:** Giao hỏa tốc 2H nội thành Hà Nội & chuyển phát nhanh toàn quốc với quy chuẩn đóng gói 3 lớp chống rách mép (Seam Splits Prevention). Nếu khách có nhu cầu hẹn bàn giao trực tiếp, chuyển qua nhân sự tại Zalo 09323.13918.")
[void]$sb.AppendLine("- **Pricing & Promo Rules:** Hiển thị giá niêm yết gốc & giá ưu đãi (chưa bao gồm VAT 10%; giá bằng 0 hiển thị Liên hệ). Lưu ý: **Chỉ giảm giá phụ kiện khi mua cùng đĩa than / mâm than**. Các tựa CD đặt trước/kiểm kho hiển thị giá tham khảo kèm nút Liên hệ.")
[void]$sb.AppendLine("- **Condition (Goldmine Grading Standard):** Mới 100% (M) / New seal (M), NM, VG+, VG, G+, G, F, P. **Status:** Còn hàng, Hết hàng.")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 3. Analog Audio Guides & FAQs (/cam-nang/)")
foreach ($g in $guideMeta) {
    [void]$sb.AppendLine("- **$($g.title):** $SiteBaseUrl/cam-nang/$($g.slug)/")
}
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 4. Unified Analog Audio Catalog ($($allItems.Count) Items)")
[void]$sb.AppendLine("| Category | SKU | Artist / Brand | Product / Album | Genre | Condition | Sale Price (VND) | Status | URL |")
[void]$sb.AppendLine("| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |")
foreach ($it in $allItems) {
    $priceTxt = if ($it.p -gt 0) { $it.p.ToString("N0") } else { "Liên hệ" }
    $urlTxt = "$SiteBaseUrl/p/$($it.slug)/"
    $llmsPrefix = switch ($it.category) {
        "vinyl"     { "Đĩa than (Vinyl) " }
        "cd"        { "Đĩa CD " }
        "turntable" { "Máy hát đĩa than " }
        default     { "" }
    }
    [void]$sb.AppendLine("| $($it.category) | $($it.sku) | $($it.a) | $llmsPrefix$($it.t) | $(Get-GenreDisplayVi $it.g) | $($it.cond) | $priceTxt | $($it.s) | $urlTxt |")
}
[System.IO.File]::WriteAllText((Join-Path (Get-Location) "llms.txt"), $sb.ToString(), $utf8NoBom)

# Ham tra ve FAQ rieng theo tung nhom san pham (hien thi cheo tren /p/ va FAQPage JSON-LD)
function Get-CategoryFaqs {
    param([string]$Category)
    switch ($Category) {
        "vinyl" {
            return @(
                @{ q = "Đĩa than mới mua (Mới 100% / M) có cần rửa trước khi nghe không?"; a = "Nên rửa nhẹ hoặc quét sạch tĩnh điện bằng chổi sợi Carbon trước khi hạ kim lần đầu để loại bỏ bụi giấy từ bao lót nguyên bản và lớp dầu khuôn ép, giúp âm thanh trong trẻo và bền đầu kim hơn."; guide = "../../cam-nang/dia-than-moi-co-can-rua-truoc-khi-nghe/" },
                @{ q = "Đóng gói đĩa than gửi đi tỉnh có sợ cong vênh hay rách mép (Seam Splits) không?"; a = "Hue Sound Express áp dụng chuẩn đóng gói 3 lớp chống rách mép (bọc xốp nổ đa lớp, chèn bìa carton cứng chống gập góc và hộp mailer 12-inch chuyên dụng), an toàn tuyệt đối khi vận chuyển toàn quốc."; guide = "../../cam-nang/dong-goi-dia-than-chong-rach-mep-di-tinh/" },
                @{ q = "Cách bảo quản đĩa than với thời tiết nóng ẩm ở Việt Nam?"; a = "Luôn xếp đĩa thẳng đứng như xếp sách ở nơi khô thoáng (độ ẩm lý tưởng 45–55%), thay bao lót chống tĩnh điện (Anti-Static Inner Sleeve) và bọc áo ngoài nilon để tránh mốc bìa giấy vào mùa nồm."; guide = "../../cam-nang/bao-quan-dia-than-thoi-tiet-viet-nam/" }
            )
        }
        "cd" {
            return @(
                @{ q = "Đĩa CD nhạc Việt tại Hue Sound Express có phải bản chính hãng không?"; a = "Toàn bộ đĩa CD/SACD đều là ấn phẩm phát hành chính hãng. Với các tựa CD hiển thị nút Liên hệ, bạn chỉ cần nhắn Zalo 09323.13918 hoặc WhatsApp để shop kiểm tra tồn kho và giữ đĩa ngay."; guide = "../../cam-nang/giai-thich-thang-goldmine-grading/" },
                @{ q = "Giao hàng đĩa CD đi tỉnh có sợ vỡ vỏ hộp nhựa (Jewel Case) không?"; a = "Mỗi đĩa CD đều được quấn nhiều lớp xốp chống sốc và đóng hộp carton cứng 5 lớp; nếu không may nứt vỡ do vận chuyển, Hue Sound Express hỗ trợ đổi trả trong 7 ngày."; guide = "../../cam-nang/dong-goi-dia-than-chong-rach-mep-di-tinh/" }
            )
        }
        "turntable" {
            return @(
                @{ q = "Người mới bắt đầu nên chọn mâm đĩa than (Turntable) loại nào?"; a = "Người mới nên ưu tiên mâm than có tích hợp sẵn mạch Phono Preamp (Built-in Phono) và kim MM phổ biến để dễ dàng cắm trực tiếp vào loa active hoặc ampli sẵn có tại nhà."; guide = "../../cam-nang/cach-chon-mam-than-dau-tien/" },
                @{ q = "Kim đĩa than MM và MC khác nhau như thế nào?"; a = "Kim MM (Moving Magnet) dễ chơi, dễ thay mũi kim và hợp mọi phono cơ bản; kim MC (Moving Coil) cho độ chi tiết và không gian âm thanh cao cấp hơn, phù hợp hệ thống Audiophile có Phono MC chuyên dụng."; guide = "../../cam-nang/phan-biet-kim-mm-va-mc/" },
                @{ q = "Hue Sound Express có hỗ trợ cân chỉnh mâm than và lắp kim không?"; a = "Đội ngũ Hue Sound Vinyl Care hỗ trợ cân lực tì kim (Tracking Force), chống trượt (Anti-Skating) và tư vấn phối ghép trực tiếp qua Zalo 09323.13918."; guide = "../../cam-nang/cach-chon-mam-than-dau-tien/" }
            )
        }
        default {
            return @(
                @{ q = "Chính sách giảm giá phụ kiện tại Hue Sound Express như thế nào?"; a = "Lưu ý: Chương trình giảm giá phụ kiện & dịch vụ chỉ áp dụng khi quý khách mua kèm đĩa than (Vinyl LP) hoặc mâm đĩa than (Turntable) trong cùng đơn hàng."; guide = "../../cam-nang/bao-quan-dia-than-thoi-tiet-viet-nam/" },
                @{ q = "Tại sao nên dùng chổi sợi Carbon và túi lót chống tĩnh điện cho đĩa than?"; a = "Chổi sợi Carbon giúp khử tĩnh điện và quét sạch bụi mịn trước khi hạ kim, còn túi lót chống tĩnh điện thay cho túi giấy thô giúp đĩa không bị xước dăm và giảm tiếng nổ lách tách."; guide = "../../cam-nang/dia-than-moi-co-can-rua-truoc-khi-nghe/" }
            )
        }
    }
}

# 4. Sinh trang chi tiet tinh /p/<slug>/index.html + Redirect stub cho legacy_slug + 6 trang /cam-nang/ + sitemap.xml
if ($GenerateProductPages) {
    $pRoot = Join-Path (Get-Location) "p"
    if (-not (Test-Path $pRoot)) { New-Item -ItemType Directory -Path $pRoot -Force | Out-Null }

    $sm = New-Object System.Text.StringBuilder
    [void]$sm.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    [void]$sm.AppendLine('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">')
    [void]$sm.AppendLine("  <url><loc>$SiteBaseUrl/</loc><priority>1.0</priority></url>")

    $accLookup = @($allItems | Where-Object { $_.category -eq "accessory" -or $_.category -eq "care" })
    $vinylLookup = @($allItems | Where-Object { $_.category -eq "vinyl" -and $_.s -eq "Còn hàng" })
    $ttLookup = @($allItems | Where-Object { $_.category -eq "turntable" })
    $redirectCount = 0

    foreach ($it in $allItems) {
        $dir = Join-Path $pRoot $it.slug
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        $seoPrefix = switch ($it.category) {
            "vinyl"     { "Đĩa than (Vinyl) " }
            "cd"        { "Đĩa CD / SACD " }
            "turntable" { "Máy hát đĩa than (Turntable) " }
            default     { "Analog Audio – " }
        }
        $catShort = switch ($it.category) {
            "vinyl"     { "Đĩa Than (Vinyl LP)" }
            "cd"        { "Đĩa CD / SACD" }
            "turntable" { "Máy Hát / Mâm Than" }
            "care"      { "Phụ kiện & Dịch vụ" }
            default     { "Phụ kiện & Dịch vụ" }
        }
        $catLabel = "$catShort • Analog Audio"
        $genreDisplay = Get-GenreDisplayVi $it.g
        $extraSeoDesc = if ($it.g -eq "Nhạc Việt / Bolero") { " Thể loại Nhạc Bolero (Nhạc Vàng / Nhạc Trữ Tình) chuẩn Analog Audio." } elseif ($it.g -eq "Tình khúc 1954–1975") { " Dòng nhạc Tình khúc 1954–1975 kinh điển chuẩn Analog Audio." } elseif ($it.category -eq "vinyl") { " Đĩa than Vinyl chuẩn Analog Audio tại Hue Sound Express." } else { " Sản phẩm Analog Audio chính hãng tại Hue Sound Express." }

        $titleEsc = Escape-Html "$($it.a) – $($it.t)"
        $seoTitleEsc = Escape-Html "$seoPrefix$($it.t) – $($it.a)"
        $descEsc = Escape-Html "$($it.d)$extraSeoDesc"
        $priceFmt = if ($it.p -gt 0) { "$($it.p.ToString('N0')) đ" } else { "Liên hệ báo giá" }
        $origPriceFmt = if ($it.p_orig -gt $it.p) { "$($it.p_orig.ToString('N0')) đ" } else { "" }
        $saveK = if ($it.p_orig -gt $it.p) { [int](($it.p_orig - $it.p) / 1000) } else { 0 }
        $condVi = if ($it.cond -eq "M (New seal)") { "Mới 100% (M)" } else { $it.cond }
        $isAccBundle = ($it.category -eq "accessory" -or $it.category -eq "care")
        $accNoteHtml = if ($isAccBundle) { "<div class='acc-note'>* Chỉ giảm giá phụ kiện khi mua cùng đĩa than / mâm than</div>" } else { "" }
        $priceHtml = if ($it.p -gt 0 -and $origPriceFmt) {
            "<div class='price-stack'><span class='price-label'>Giá</span><span class='price-old'>$origPriceFmt</span><span class='price-save'>-$($saveK)k</span></div><div class='price'>$priceFmt</div>$accNoteHtml"
        } else {
            "<div class='price-stack'><span class='price-label'>Giá</span></div><div class='price'>$priceFmt</div>$accNoteHtml"
        }
        $pageUrl = "$SiteBaseUrl/p/$($it.slug)/"
        $ogImg = if ($it.img -match '^https?://') { $it.img } else { "$SiteBaseUrl/og-image.jpg" }

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
        } elseif ($it.category -eq "vinyl" -or $it.category -eq "cd") {
            $upsellItems = $accLookup | Where-Object { $_.sku -in @("ACC-SLEEVE-50","ACC-AQ-CARBON","CARE-ULTRASONIC-01") } | Select-Object -First 3
        } else {
            $upsellItems = $vinylLookup | Select-Object -First 3
        }
        $upsellHtml = ""
        if ($upsellItems.Count -gt 0) {
            $cards = ($upsellItems | ForEach-Object {
                "<a class='up-card' href='../$($_.slug)/'><strong>$(Escape-Html $_.t)</strong><span>$($_.p.ToString('N0')) đ</span></a>"
            }) -join ""
            $upHead = if ($isAccBundle) { "Đĩa than mới về đề xuất phối ghép" } else { "Gợi ý đi kèm từ Hue Sound Vinyl Care" }
            $upsellHtml = "<section class='box'><h2>$upHead</h2><div class='up-grid'>$cards</div></section>"
        }

        # FAQ rieng theo nhom san pham (HTML + JSON-LD FAQPage)
        $catFaqs = Get-CategoryFaqs -Category $it.category
        $faqItemsHtml = ($catFaqs | ForEach-Object {
            "<details class='faq-item'><summary>$(Escape-Html $_.q)</summary><p>$(Escape-Html $_.a) <a class='faq-more' href='$($_.guide)'>Xem bài viết chi tiết →</a></p></details>"
        }) -join ""
        $faqBoxHtml = "<section class='box'><h2>Câu hỏi thường gặp (FAQ)</h2>$faqItemsHtml</section>"
        $faqJsonArr = ($catFaqs | ForEach-Object {
            "{ `"@type`": `"Question`", `"name`": `"$(Escape-JsonStr $_.q)`", `"acceptedAnswer`": { `"@type`": `"Answer`", `"text`": `"$(Escape-JsonStr $_.a)`" } }"
        }) -join ",`n      "

        $chatRaw = "Chào Hue Sound Express, mình quan tâm sản phẩm: $seoPrefix$($it.a) - $($it.t) (Mã: $($it.sku) | Tình trạng đĩa: $condVi | Giá: $priceFmt). Link: $pageUrl"
        $chatEnc = [Uri]::EscapeDataString($chatRaw)
        $waHref = "https://wa.me/84932313918?text=$chatEnc"
        $zaloHref = "https://zalo.me/0932313918"
        $chatHtmlVal = Escape-Html $chatRaw

        $schemaAvail = if ($it.s -eq "Còn hàng") { "https://schema.org/InStock" } else { "https://schema.org/OutOfStock" }
        $schemaCond = if ($it.cond -eq "M (New seal)") { "https://schema.org/NewCondition" } else { "https://schema.org/UsedCondition" }

        $primaryCtaHtml = if ($it.contact_only -or $it.s -eq "Hết hàng") {
            "<button class='btn btn-copper' type='button' onclick='openZaloPopup()'>💬 Liên hệ đặt hàng</button>"
        } else {
            "<a class='btn btn-copper' href='../../?add=$([Uri]::EscapeDataString($it.sku))'>+ Thêm vào Giỏ / Báo giá</a>"
        }

        $html = @"
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>$seoTitleEsc | Hue Sound Express – Analog Audio</title>
  <meta name="description" content="$descEsc">
  <meta name="theme-color" content="#E8D9B6">
  <link rel="canonical" href="$pageUrl">
  <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Ccircle cx='32' cy='32' r='30' fill='%23EADBB8'/%3E%3Ccircle cx='32' cy='32' r='22' fill='none' stroke='%23B5541F' stroke-width='3'/%3E%3Ccircle cx='32' cy='32' r='14' fill='none' stroke='%23C99A2E' stroke-width='2'/%3E%3Ccircle cx='32' cy='32' r='5' fill='%23B5541F'/%3E%3C/svg%3E">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,500;0,600;1,400&family=Montserrat:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <meta property="og:title" content="$seoTitleEsc – $priceFmt | Hue Sound Express">
  <meta property="og:description" content="$descEsc">
  <meta property="og:image" content="$ogImg">
  <meta property="og:url" content="$pageUrl">
  <meta property="og:type" content="product">
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@graph": [
      {
        "@type": "BreadcrumbList",
        "itemListElement": [
          { "@type": "ListItem", "position": 1, "name": "Hue Sound Express", "item": "$SiteBaseUrl/" },
          { "@type": "ListItem", "position": 2, "name": "$(Escape-JsonStr $catShort)", "item": "$SiteBaseUrl/?cat=$($it.category)" },
          { "@type": "ListItem", "position": 3, "name": "$(Escape-JsonStr "$($it.a) - $($it.t)")", "item": "$pageUrl" }
        ]
      },
      {
        "@type": "Product",
        "name": "$(Escape-JsonStr "$seoPrefix$($it.a) - $($it.t)")",
        "sku": "$(Escape-JsonStr $it.sku)",
        "category": "$(Escape-JsonStr "$catLabel / $genreDisplay")",
        "image": "$ogImg",
        "description": "$(Escape-JsonStr "$($it.d)$extraSeoDesc")",
        "brand": { "@type": "Brand", "name": "$(Escape-JsonStr $(if ($it.l) { $it.l } else { 'Hue Sound Express' }))" },
        "offers": {
          "@type": "Offer",
          "priceCurrency": "VND",
          "price": "$($it.p)",
          "priceValidUntil": "$priceValidUntil",
          "itemCondition": "$schemaCond",
          "availability": "$schemaAvail",
          "url": "$pageUrl",
          "shippingDetails": {
            "@type": "OfferShippingDetails",
            "shippingDestination": { "@type": "DefinedRegion", "addressCountry": "VN" },
            "deliveryTime": {
              "@type": "ShippingDeliveryTime",
              "handlingTime": { "@type": "QuantitativeValue", "minValue": 0, "maxValue": 1, "unitCode": "DAY" },
              "transitTime": { "@type": "QuantitativeValue", "minValue": 1, "maxValue": 4, "unitCode": "DAY" }
            }
          },
          "hasMerchantReturnPolicy": {
            "@type": "MerchantReturnPolicy",
            "applicableCountry": "VN",
            "returnPolicyCategory": "https://schema.org/MerchantReturnFiniteReturnWindow",
            "merchantReturnDays": 7,
            "returnMethod": "https://schema.org/ReturnByMail",
            "returnFees": "https://schema.org/FreeReturn"
          }
        }
      },
      {
        "@type": "FAQPage",
        "mainEntity": [
      $faqJsonArr
        ]
      }
    ]
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
    body::after { content:''; position:fixed; inset:0; z-index:1; pointer-events:none; background:radial-gradient(ellipse at 50% 42%, transparent 56%, rgba(84,52,22,.24) 100%); }
    .wrap { max-width:900px; margin:0 auto; position:relative; z-index:2; }
    .topbar { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:16px; }
    .back { display:inline-flex; align-items:center; gap:6px; color:var(--copper-dark); text-decoration:none; font-size:13px; font-weight:600; background:var(--card); border:1px solid var(--border); padding:8px 14px; border-radius:99px; }
    .back:hover { background:#FFFBF0; border-color:var(--copper); }
    .brand-mini { font-weight:800; font-size:13px; color:var(--copper); letter-spacing:.04em; }
    .crumb { font-size:12px; color:var(--muted); margin-bottom:14px; }
    .crumb a { color:var(--copper-dark); text-decoration:none; font-weight:600; }
    .hero { display:grid; grid-template-columns:290px 1fr; gap:26px; background:var(--card); border:1px solid var(--border); border-radius:18px; padding:24px; backdrop-filter:blur(8px); box-shadow:0 12px 30px -12px rgba(110,60,20,.22); }
    @media (max-width:680px){ .hero{ grid-template-columns:1fr; } }
    .cover { aspect-ratio:1/1; border-radius:14px; overflow:hidden; background:radial-gradient(circle at center, #F3E9D0 0%, #DFC99B 100%); border:1px solid var(--border); display:flex; align-items:center; justify-content:center; text-align:center; padding:16px; position:relative; }
    .cover img { width:100%; height:100%; object-fit:cover; }
    .cover-art { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:8px; color:var(--ink); }
    .badge { display:inline-block; font-size:11px; font-weight:600; padding:4px 11px; border-radius:99px; background:rgba(181,84,31,.11); color:var(--copper-dark); border:1px solid rgba(181,84,31,.22); margin-bottom:10px; }
    h1 { font-family:var(--serif); font-size:25px; line-height:1.3; margin-bottom:4px; color:var(--ink); }
    .sub { color:var(--copper-dark); font-size:16px; font-weight:600; margin-bottom:10px; }
    .meta { font-size:12.5px; color:var(--muted); margin-bottom:12px; }
    .desc { font-size:14px; color:var(--ink); opacity:.92; }
    .price-row { display:flex; align-items:center; flex-wrap:wrap; gap:12px; margin:16px 0; }
    .price-label { font-size:11px; text-transform:uppercase; letter-spacing:.06em; color:var(--muted); font-weight:600; }
    .price-stack { display:flex; align-items:center; gap:7px; }
    .price-old { font-size:13.5px; color:var(--muted); text-decoration:line-through; opacity:.82; font-weight:500; }
    .price-save { font-size:10.5px; font-weight:700; color:#FFF7E8; background:#B5541F; padding:1px 6px; border-radius:99px; }
    .price { font-size:25px; font-weight:800; color:var(--copper); line-height:1.15; }
    .acc-note { font-size:11.5px; color:#8C3A12; font-weight:600; margin-top:4px; }
    .status { font-size:12px; font-weight:600; padding:4px 11px; border-radius:99px; background:rgba(201,154,46,.18); border:1px solid rgba(201,154,46,.5); color:var(--ink); }
    .cta-row { display:flex; flex-wrap:wrap; gap:9px; margin-top:14px; }
    .btn { text-decoration:none; padding:10px 15px; border-radius:10px; font-weight:600; font-size:13px; display:inline-flex; align-items:center; gap:6px; transition:all .2s; border:none; cursor:pointer; font-family:inherit; }
    .btn-copper { background:var(--copper-grad); color:#FFF7E8; box-shadow:inset 0 1px 0 rgba(244,219,141,.4), 0 6px 16px -6px rgba(126,50,16,.5); }
    .btn-copper:hover { transform:translateY(-1px); }
    .btn-wa { background:#25D366; color:#fff; }
    .btn-wa:hover { background:#1EBE5D; }
    .btn-zalo { background:#0068FF; color:#fff; }
    .btn-zalo:hover { background:#0052CC; }
    .btn-share { background:rgba(255,251,240,.85); color:var(--ink); border:1px solid rgba(181,84,31,.35); }
    .btn-share:hover { border-color:var(--copper); color:var(--copper-dark); }
    .audio-box { margin-top:16px; padding:12px 14px; background:rgba(243,233,208,.8); border-radius:10px; border:1px solid var(--border); }
    .small { font-size:12px; color:var(--muted); }
    .box { margin-top:18px; background:var(--card); border:1px solid var(--border); border-radius:16px; padding:20px; backdrop-filter:blur(6px); }
    .box h2 { font-family:var(--serif); font-size:17px; color:var(--copper-dark); margin-bottom:12px; }
    .tracks { padding-left:20px; color:var(--ink); font-size:13.5px; columns:2; gap:26px; }
    .tracks li { margin-bottom:5px; }
    @media (max-width:600px){ .tracks{ columns:1; } }
    .up-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:12px; }
    .up-card { text-decoration:none; background:rgba(255,251,240,.75); border:1px solid var(--border); padding:14px; border-radius:12px; color:var(--ink); display:flex; flex-direction:column; gap:6px; transition:all .2s; }
    .up-card:hover { border-color:var(--copper); transform:translateY(-2px); }
    .up-card span { color:var(--copper); font-weight:700; font-size:13.5px; }
    .faq-item { border-bottom:1px dashed rgba(142,66,24,.22); padding:9px 0; }
    .faq-item:last-child { border-bottom:none; }
    .faq-item summary { font-weight:700; font-size:13.5px; cursor:pointer; color:var(--ink); }
    .faq-item p { font-size:13px; color:var(--muted); margin-top:6px; }
    .faq-more { color:var(--copper); font-weight:600; text-decoration:none; }
    footer { text-align:center; margin-top:34px; font-size:12px; color:var(--muted); }
    .chat-backdrop { position:fixed; inset:0; background:rgba(58,28,10,.5); backdrop-filter:blur(3px); z-index:90; display:none; align-items:center; justify-content:center; padding:16px; }
    .chat-backdrop.open { display:flex; }
    .chat-popup { background:#FAF2DE; border:1px solid rgba(142,66,24,.45); border-radius:16px; max-width:460px; width:100%; padding:18px; box-shadow:0 20px 50px rgba(58,28,10,.35); }
    .chat-popup h3 { font-family:var(--serif); color:var(--copper-dark); font-size:17px; margin-bottom:6px; }
    .chat-popup textarea { width:100%; padding:10px; border-radius:10px; border:1px solid var(--border); background:#FFFDF7; color:var(--ink); font-family:inherit; font-size:13px; margin:8px 0 10px; }
    .chat-pop-note { font-size:12px; color:#29523A; background:rgba(61,107,79,.12); padding:7px 10px; border-radius:8px; margin-bottom:10px; }
    .chat-pop-actions { display:flex; flex-wrap:wrap; gap:8px; }
    .share-toast { position:fixed; bottom:24px; left:50%; transform:translateX(-50%); background:#4A2312; color:#FFF7E8; padding:9px 16px; border-radius:99px; font-size:12.5px; font-weight:600; display:none; z-index:100; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="topbar">
      <a class="back" href="../../">← Quay lại Phòng Nghe Số Hue Sound Express</a>
      <span class="brand-mini">HUE SOUND EXPRESS • Analog Audio</span>
    </div>
    <nav class="crumb" aria-label="Breadcrumb">
      <a href="../../">Trang chủ</a> › <a href="../../?cat=$($it.category)">$catShort</a> › <span>$(Escape-Html $it.t)</span>
    </nav>
    <div class="hero">
      <div class="cover">
        $(if ($it.img -match '^https?://') { "<img src='$(Escape-Html $it.img)' alt='$seoTitleEsc'>" } else { "<div class='cover-art'><svg width='96' height='96' viewBox='0 0 100 100'><circle cx='50' cy='50' r='44' fill='none' stroke='#B5541F' stroke-width='2' opacity='.5'/><circle cx='50' cy='50' r='34' fill='none' stroke='#C99A2E' stroke-width='1.5' opacity='.6'/><circle cx='50' cy='50' r='24' fill='none' stroke='#B5541F' stroke-width='1.5' opacity='.5'/><circle cx='50' cy='50' r='9' fill='#B5541F'/></svg><div style='font-weight:700;font-size:13px;'>$(Escape-Html $it.a)</div><div style='font-size:12px;color:#7A5C46;'>$(Escape-Html $it.t)</div></div>" })
      </div>
      <div>
        <span class="badge">$catLabel • $(Escape-Html $genreDisplay)</span>
        <h1>$(Escape-Html $it.t)</h1>
        <div class="sub">$(Escape-Html $it.a)</div>
        <div class="meta">Mã SKU: <strong>$(Escape-Html $it.sku)</strong> • Định dạng: $(Escape-Html $it.f) $(if ($it.l) { "• $(Escape-Html $it.l)" }) $(if ($it.y) { "• Năm $($it.y)" })</div>
        <p class="desc">$descEsc</p>
        <div class="price-row">
          <div>$priceHtml</div>
          <span class="status">$(Escape-Html $condVi)</span>
          <span class="status">$(Escape-Html $it.s)</span>
        </div>
        <div class="cta-row">
          $primaryCtaHtml
          <button class="btn btn-zalo" type="button" onclick="openZaloPopup()">💬 Chat Zalo</button>
          <a class="btn btn-wa" href="$waHref" target="_blank" rel="noopener noreferrer">💬 WhatsApp</a>
          <button class="btn btn-share" type="button" onclick="shareProduct()">📤 Chia sẻ link</button>
        </div>
        $audioHtml
      </div>
    </div>
    $trackHtml
    $specHtml
    $faqBoxHtml
    $upsellHtml
    <footer>© 2026 Hue Sound Express • Analog Audio • Đóng gói 3 lớp chống rách mép • Hotline/Zalo &amp; WhatsApp: 09323.13918</footer>
  </div>

  <div class="share-toast" id="shareToast">✓ Đã sao chép link sản phẩm để chia sẻ!</div>

  <!-- Pop-up Chat Zalo kèm sẵn thông tin sản phẩm -->
  <div class="chat-backdrop" id="zaloPopup" onclick="if(event.target===this)this.classList.remove('open')">
    <div class="chat-popup">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <h3>💬 Gửi thông tin sản phẩm qua Zalo</h3>
        <button type="button" onclick="document.getElementById('zaloPopup').classList.remove('open')" style="background:none;border:none;font-size:18px;cursor:pointer;color:var(--ink);">✕</button>
      </div>
      <div class="chat-pop-note">✓ Đã tự động sao chép sẵn nội dung bên dưới — khi mở Zalo bạn chỉ cần Dán (Paste) &amp; Gửi!</div>
      <textarea id="zaloMsgBox" rows="4">$chatHtmlVal</textarea>
      <div class="chat-pop-actions">
        <a class="btn btn-zalo" href="$zaloHref" target="_blank" rel="noopener noreferrer" onclick="copyChatMsg()">Mở Zalo (09323.13918)</a>
        <a class="btn btn-wa" id="waPopBtn" href="$waHref" target="_blank" rel="noopener noreferrer">Mở WhatsApp (Tự điền sẵn)</a>
      </div>
    </div>
  </div>
  <script>
    function copyChatMsg() {
      var el = document.getElementById('zaloMsgBox');
      if (navigator.clipboard) navigator.clipboard.writeText(el.value).catch(function(){});
      else { el.select(); try { document.execCommand('copy'); } catch(e){} }
      document.getElementById('waPopBtn').href = 'https://wa.me/84932313918?text=' + encodeURIComponent(el.value);
    }
    function openZaloPopup() {
      document.getElementById('zaloPopup').classList.add('open');
      copyChatMsg();
    }
    function shareProduct() {
      var url = '$pageUrl';
      var title = document.title;
      if (navigator.share) {
        navigator.share({ title: title, url: url }).catch(function(){});
      } else {
        if (navigator.clipboard) navigator.clipboard.writeText(url).catch(function(){});
        var t = document.getElementById('shareToast');
        t.style.display = 'block';
        setTimeout(function(){ t.style.display = 'none'; }, 2400);
      }
    }
  </script>
</body>
</html>
"@
        [System.IO.File]::WriteAllText((Join-Path $dir "index.html"), $html, $utf8NoBom)
        [void]$sm.AppendLine("  <url><loc>$pageUrl</loc><priority>0.8</priority></url>")

        # Tao trang chuyen huong SEO (giu nguyen URL cu da index de khong mat thu hang Google!)
        if ($it.legacy_slug -and ($it.legacy_slug -ne $it.slug)) {
            $oldDir = Join-Path $pRoot $it.legacy_slug
            if (-not (Test-Path $oldDir)) { New-Item -ItemType Directory -Path $oldDir -Force | Out-Null }
            $redirHtml = @"
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <title>Đang chuyển hướng – $seoTitleEsc</title>
  <meta name="robots" content="noindex, follow">
  <link rel="canonical" href="$pageUrl">
  <meta http-equiv="refresh" content="0; url=../$($it.slug)/">
  <script>window.location.replace('../$($it.slug)/');</script>
</head>
<body><p>Đang chuyển hướng tới <a href="../$($it.slug)/">$seoTitleEsc</a>...</p></body>
</html>
"@
            [System.IO.File]::WriteAllText((Join-Path $oldDir "index.html"), $redirHtml, $utf8NoBom)
            $redirectCount++
        }
    }

    # 5. SINH 6 BAI VIET TINH CAM NANG ANALOG AUDIO (/cam-nang/<slug>/index.html) + TU DONG CAP NHAT SAN PHAM LIEN QUAN
    $guideRoot = Join-Path (Get-Location) "cam-nang"
    if (-not (Test-Path $guideRoot)) { New-Item -ItemType Directory -Path $guideRoot -Force | Out-Null }

    $guideDefs = @(
        @{
            slug = "bao-quan-dia-than-thoi-tiet-viet-nam"
            title = "Cách bảo quản đĩa than với thời tiết nóng ẩm ở Việt Nam"
            desc = "Hướng dẫn chi tiết từ Hue Sound Vinyl Care về cách chống mốc bìa, chống cong vênh và giữ rãnh đĩa than bền đẹp hàng chục năm trong khí hậu nhiệt đới gió mùa Việt Nam."
            body = @"
<p>Khí hậu Việt Nam đặc trưng bởi <strong>mùa hè nắng nóng trên 38°C</strong> và <strong>mùa nồm ẩm miền Bắc có lúc vượt 85–90%</strong>. Đây là hai kẻ thù lớn nhất khiến đĩa than (Vinyl LP) bị cong vênh phôi PVC và mốc loang bìa giấy.</p>
<h2>1. Kiểm soát nhiệt độ &amp; độ ẩm phòng nghe</h2>
<ul>
  <li><strong>Độ ẩm lý tưởng:</strong> duy trì trong khoảng <strong>45% – 55%</strong> bằng máy hút ẩm hoặc điều hòa chế độ Dry vào mùa nồm.</li>
  <li><strong>Nhiệt độ an toàn:</strong> giữ dưới 30°C, tuyệt đối tránh đặt kệ đĩa cạnh cửa sổ nắng chiếu trực tiếp, gần ampli đèn tỏa nhiệt mạnh hoặc nóc loa.</li>
</ul>
<h2>2. Luôn xếp đĩa thẳng đứng &amp; nâng cấp bao lót</h2>
<ul>
  <li>Xếp đĩa thẳng đứng như xếp sách, không xếp chồng ngang quá 3 đĩa để tránh lực nén gây cong vênh rãnh đĩa.</li>
  <li>Thay thế túi giấy nguyên bản (dễ hút ẩm, bám mốc và gây xước dăm) bằng <strong>túi lót chống tĩnh điện (Anti-Static Inner Sleeve)</strong> và bọc <strong>áo nilon ngoài (Outer Sleeve)</strong> hở cạnh để bìa đĩa "thở" nhẹ, tránh hấp hơi nước.</li>
</ul>
<h2>3. Vệ sinh định kỳ bằng chổi Carbon &amp; rửa siêu âm</h2>
<p>Trước mỗi lần hạ kim, lướt nhẹ chổi sợi Carbon 2–3 vòng để triệt tiêu bụi mịn và tĩnh điện. Với đĩa nghe lâu ngày hoặc có dấu hiệu mốc nhẹ, hãy sử dụng <strong>dịch vụ rửa đĩa siêu âm chuẩn Audiophile tại Hue Sound Vinyl Care</strong>.</p>
"@
            related = @($accLookup | Where-Object { $_.sku -in @("ACC-SLEEVE-50","ACC-AQ-CARBON","CARE-ULTRASONIC-01") }) + @($vinylLookup | Select-Object -First 3)
        },
        @{
            slug = "dia-than-moi-co-can-rua-truoc-khi-nghe"
            title = "Đĩa than mới mua (Mới 100% / New Seal) có cần rửa trước khi nghe không?"
            desc = "Giải đáp thắc mắc phổ biến nhất của người chơi Analog Audio: Vì sao đĩa than nguyên seal vừa bóc vẫn có tiếng nổ lách tách và cách xử lý chuẩn Hue Sound Vinyl Care."
            body = @"
<p>Rất nhiều người mới chơi bất ngờ khi bóc một chiếc đĩa than <strong>Mới 100% (M / New Seal)</strong> nhưng đặt lên mâm vẫn nghe vài tiếng nổ lách tách nhẹ. Nguyên nhân hoàn toàn không phải do đĩa lỗi mà đến từ quy trình ép đĩa tại nhà máy:</p>
<h2>1. Hai nguyên nhân đĩa mới vẫn bám bụi vi mô</h2>
<ul>
  <li><strong>Lớp chất tách khuôn (Mold Release Compound):</strong> một lượng nhỏ dầu/sáp công nghiệp giúp tách đĩa PVC khỏi khuôn ép vẫn có thể sót lại dưới đáy rãnh sâu.</li>
  <li><strong>Bụi giấy &amp; tĩnh điện:</strong> trong quá trình đóng gói bằng túi giấy thô và vận chuyển đường dài, ma sát tạo ra tĩnh điện hút sạch vụn giấy vào mặt đĩa.</li>
</ul>
<h2>2. Quy trình 2 bước chuẩn trước lần nghe đầu tiên</h2>
<ol>
  <li><strong>Cơ bản tại nhà:</strong> Quét sạch mặt đĩa bằng chổi sợi Carbon chống tĩnh điện (AudioQuest) và thay ngay bằng túi lót chống tĩnh điện mới.</li>
  <li><strong>Chuẩn Audiophile:</strong> Rửa đĩa bằng máy rửa siêu âm (Ultrasonic Cleaning) giúp đẩy sạch 100% lớp dầu khuôn ép, làm nền âm tĩnh lặng (black background) và kéo dài tuổi thọ đầu kim lên gấp đôi.</li>
</ol>
"@
            related = @($accLookup | Where-Object { $_.sku -in @("CARE-ULTRASONIC-01","ACC-AQ-CARBON","ACC-SLEEVE-50","ACC-AT607A") }) + @($vinylLookup | Select-Object -First 2)
        },
        @{
            slug = "dong-goi-dia-than-chong-rach-mep-di-tinh"
            title = "Đóng gói đĩa than gửi đi tỉnh có sợ cong vênh hay rách mép (Seam Splits) không?"
            desc = "Quy chuẩn đóng gói 3 lớp Seam Splits Prevention độc quyền tại Hue Sound Express giúp bảo vệ góc bìa và phôi đĩa than an toàn tuyệt đối khi vận chuyển toàn quốc."
            body = @"
<p>Nỗi lo lớn nhất khi mua đĩa than online ở xa là tình trạng <strong>rách mép bìa (Seam Splits)</strong> do phôi đĩa bên trong trượt mạnh khi vận chuyển, hoặc <strong>gãy góc bìa</strong> do hộp đóng gói mỏng.</p>
<h2>Quy chuẩn đóng gói 3 lớp tại Hue Sound Express</h2>
<ol>
  <li><strong>Lớp 1 – Cố định &amp; chống sốc:</strong> Bọc đĩa trong nhiều lớp xốp nổ dày, chèn đệm cố định để phôi đĩa không bị xê dịch va đập vào mép bìa.</li>
  <li><strong>Lớp 2 – Tấm bìa carton cứng bảo vệ 4 góc:</strong> Ép kẹp hai mặt bằng bìa carton cứng kích thước lớn hơn bìa đĩa giúp phân tán lực khi va chạm.</li>
  <li><strong>Lớp 3 – Hộp Mailer 12-inch chuyên dụng:</strong> Đóng trong hộp carton 5 lớp chuyên biệt cho đĩa than, dán nhãn hàng dễ vỡ và hỗ trợ đồng kiểm / đổi trả miễn phí trong 7 ngày nếu phát sinh sự cố vận chuyển.</li>
</ol>
<p>Nội thành Hà Nội được <strong>giao hỏa tốc 2H</strong>; các tỉnh thành toàn quốc nhận hàng nhanh từ 1–4 ngày làm việc.</p>
"@
            related = @($vinylLookup | Select-Object -First 4) + @($accLookup | Where-Object { $_.sku -in @("ACC-SLEEVE-50","ACC-AQ-CARBON") })
        },
        @{
            slug = "cach-chon-mam-than-dau-tien"
            title = "Cách chọn mâm đĩa than (Turntable) đầu tiên cho người mới bắt đầu"
            desc = "Cẩm nang chọn mua máy hát đĩa than đầu tiên: phân biệt truyền động Belt-Drive vs Direct-Drive, Phono tích hợp, Bluetooth và ngân sách hợp lý."
            body = @"
<p>Bước vào thế giới <strong>Analog Audio</strong> không hề phức tạp hay đắt đỏ nếu bạn hiểu rõ 4 tiêu chí chọn mâm than đầu tiên:</p>
<h2>1. Tránh xa các máy hát vali giá rẻ không chỉnh được lực tì kim</h2>
<p>Các loại máy hát vali không có đối trọng (Counterweight) thường đè lực kim quá nặng (5–6 gram), rất nhanh làm mòn rãnh đĩa quý và cho âm thanh mỏng.</p>
<h2>2. Ưu tiên mâm có Phono Preamp tích hợp (Built-in Phono)</h2>
<p>Tín hiệu từ kim đĩa than rất nhỏ. Chọn mâm tích hợp sẵn Phono (có công tắc gạt Line/Phono) giúp bạn cắm thẳng vào loa kiểm âm / loa Bluetooth có cổng AUX/RCA là nghe được ngay, sau này vẫn dễ dàng nâng cấp Phono rời.</p>
<h2>3. Nên chọn Belt-Drive hay Direct-Drive?</h2>
<ul>
  <li><strong>Belt-Drive (dây curoa):</strong> cách ly rung động motor cực tốt, nền âm êm ấm — lựa chọn số 1 để nghe nhạc gia đình và Audiophile.</li>
  <li><strong>Direct-Drive (truyền động trực tiếp):</strong> tốc độ ổn định tức thì, lực kéo khỏe, bền bỉ.</li>
</ul>
<h2>4. Đừng quên cân chỉnh lực kim &amp; phụ kiện đi kèm</h2>
<p>Chỉ cần thêm một chiếc <strong>chổi sợi Carbon</strong> và <strong>cân lực tì kim điện tử 0.01g</strong> (được giảm giá khi mua kèm mâm than tại Hue Sound Express), dàn mâm than của bạn đã sẵn sàng mang lại chất âm mộc mạc chuẩn mực.</p>
"@
            related = @($ttLookup) + @($accLookup | Where-Object { $_.sku -in @("ACC-SCALE-01","CARE-SETUP-01") })
        },
        @{
            slug = "phan-biet-kim-mm-va-mc"
            title = "Phân biệt kim đĩa than MM và MC – Nên chọn loại nào cho dàn Analog Audio?"
            desc = "So sánh chi tiết cấu tạo, chất âm, chi phí thay mũi kim và yêu cầu phối ghép Phono giữa kim Moving Magnet (MM) và Moving Coil (MC)."
            body = @"
<p>Đầu kim (Cartridge) là "trái tim" chuyển đổi rung động cơ học trong rãnh đĩa thành tín hiệu âm thanh. Hiện nay có hai chuẩn phổ biến nhất là <strong>MM</strong> và <strong>MC</strong>:</p>
<h2>1. Kim MM (Moving Magnet – Nam châm động)</h2>
<ul>
  <li><strong>Ưu điểm:</strong> Điện áp ra cao (3–5mV), tương thích với hầu hết mọi Phono tích hợp sẵn; <strong>thay riêng mũi kim (Stylus) rất dễ dàng và tiết kiệm</strong>; chất âm dày dặn, ấm áp, dễ nghe.</li>
  <li><strong>Phù hợp với:</strong> Người mới chơi đến trung cấp, nghe đa dạng từ Nhạc Việt trẻ, Tình khúc 1954–1975 đến Bolero, Jazz.</li>
</ul>
<h2>2. Kim MC (Moving Coil – Cuộn dây động)</h2>
<ul>
  <li><strong>Ưu điểm:</strong> Khối lượng chuyển động siêu nhẹ giúp bắt chi tiết vi mô, dải cao tơi mịn và âm hình sân khấu 3D vượt trội.</li>
  <li><strong>Lưu ý:</strong> Điện áp ra thấp (0.2–0.5mV), đòi hỏi Phono Preamp có chế độ MC hoặc Step-up Transformer (SUT); khi mòn kim thường phải thay nguyên củ kim hoặc gửi retip.</li>
</ul>
<h2>3. Bảo vệ đầu kim đúng cách</h2>
<p>Dù dùng MM hay MC, hãy vệ sinh đầu kim định kỳ bằng <strong>dung dịch Audio-Technica AT607a</strong> và đo lực tì kim chuẩn hãng bằng cân điện tử 0.01g.</p>
"@
            related = @($ttLookup | Select-Object -First 3) + @($accLookup | Where-Object { $_.sku -in @("ACC-AT607A","ACC-SCALE-01","CARE-SETUP-01") })
        },
        @{
            slug = "giai-thich-thang-goldmine-grading"
            title = "Giải thích thang đánh giá tình trạng đĩa Goldmine Grading Standard (M, NM, VG+, VG...)"
            desc = "Cách đọc hiểu chuẩn phân loại tình trạng đĩa than quốc tế Goldmine Grading Standard được áp dụng minh bạch trên toàn bộ sản phẩm tại Hue Sound Express."
            body = @"
<p>Để đảm bảo minh bạch tuyệt đối khi mua bán đĩa than trực tuyến, <strong>Hue Sound Express</strong> áp dụng chuẩn đánh giá quốc tế <strong>Goldmine Grading Standard</strong> cho từng sản phẩm:</p>
<ul>
  <li><strong>Mới 100% (M / New Seal):</strong> Đĩa nguyên seal nhà sản xuất, chưa từng bóc mở hay hạ kim.</li>
  <li><strong>NM (Near Mint):</strong> Gần như mới hoàn hảo, bề mặt đĩa bóng sạch, không tiếng ồn nền đáng kể, bìa đẹp nguyên vẹn.</li>
  <li><strong>VG+ (Very Good Plus):</strong> Đĩa đã qua sử dụng cẩn thận, có thể có vài vết xước tơ rất nhẹ dưới ánh đèn mạnh nhưng không ảnh hưởng đến chất lượng âm thanh.</li>
  <li><strong>VG (Very Good):</strong> Có tiếng lách tách nhẹ ở những đoạn nhạc tĩnh, vẫn nghe trọn vẹn mượt mà không vấp đĩa.</li>
  <li><strong>G+ / G / F / P:</strong> Các cấp đĩa sưu tầm sâu, có dấu ấn thời gian rõ rệt trên mặt đĩa hoặc bìa.</li>
</ul>
<p>Bạn có thể lọc nhanh trực tiếp theo từng mức tình trạng đĩa ngay tại thanh <strong>Tình trạng đĩa</strong> trên trang chủ Hue Sound Express.</p>
"@
            related = @($vinylLookup | Select-Object -First 6)
        }
    )

    foreach ($gd in $guideDefs) {
        $gDir = Join-Path $guideRoot $gd.slug
        if (-not (Test-Path $gDir)) { New-Item -ItemType Directory -Path $gDir -Force | Out-Null }
        $gUrl = "$SiteBaseUrl/cam-nang/$($gd.slug)/"
        $gTitleEsc = Escape-Html $gd.title
        $gDescEsc = Escape-Html $gd.desc

        $relCards = ($gd.related | Select-Object -First 6 | ForEach-Object {
            $pStr = if ($_.p -gt 0) { "$($_.p.ToString('N0')) đ" } else { "Liên hệ" }
            "<a class='up-card' href='../../p/$($_.slug)/'><strong>$(Escape-Html $_.t)</strong><small style='color:#6F5139;'>$(Escape-Html $_.a)</small><span>$pStr</span></a>"
        }) -join ""

        $otherLinks = ($guideMeta | Where-Object { $_.slug -ne $gd.slug } | ForEach-Object {
            "<li><a href='../$($_.slug)/'>$(Escape-Html $_.title)</a></li>"
        }) -join ""

        $gHtml = @"
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>$gTitleEsc | Cẩm Nang Analog Audio – Hue Sound Express</title>
  <meta name="description" content="$gDescEsc">
  <meta name="theme-color" content="#E8D9B6">
  <link rel="canonical" href="$gUrl">
  <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Ccircle cx='32' cy='32' r='30' fill='%23EADBB8'/%3E%3Ccircle cx='32' cy='32' r='22' fill='none' stroke='%23B5541F' stroke-width='3'/%3E%3Ccircle cx='32' cy='32' r='14' fill='none' stroke='%23C99A2E' stroke-width='2'/%3E%3Ccircle cx='32' cy='32' r='5' fill='%23B5541F'/%3E%3C/svg%3E">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,500;0,600;1,400&family=Montserrat:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@graph": [
      {
        "@type": "BreadcrumbList",
        "itemListElement": [
          { "@type": "ListItem", "position": 1, "name": "Hue Sound Express", "item": "$SiteBaseUrl/" },
          { "@type": "ListItem", "position": 2, "name": "Cẩm nang FAQ & Giao vận", "item": "$SiteBaseUrl/?cat=guide" },
          { "@type": "ListItem", "position": 3, "name": "$(Escape-JsonStr $gd.title)", "item": "$gUrl" }
        ]
      },
      {
        "@type": "Article",
        "headline": "$(Escape-JsonStr $gd.title)",
        "description": "$(Escape-JsonStr $gd.desc)",
        "author": { "@type": "Organization", "name": "Hue Sound Express – Hue Sound Vinyl Care" },
        "publisher": { "@type": "Organization", "name": "Hue Sound Express" },
        "mainEntityOfPage": "$gUrl"
      }
    ]
  }
  </script>
  <style>
    :root { --copper:#B5541F; --copper-dark:#7E3210; --ink:#4A2312; --muted:#6F5139; --card:rgba(250,242,222,.86); --border:rgba(142,66,24,.27); --sans:"Montserrat",system-ui,sans-serif; --serif:"Lora",Georgia,serif; }
    * { box-sizing:border-box; margin:0; padding:0; }
    body { font-family:var(--sans); color:var(--ink); line-height:1.7; background:radial-gradient(ellipse at 50% 18%, #F4E9CC 0%, #E7D5AE 44%, #C5AC7E 85%, #AA9062 100%); background-attachment:fixed; padding:24px 16px 64px; }
    .wrap { max-width:860px; margin:0 auto; }
    .topbar { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:16px; }
    .back { color:var(--copper-dark); text-decoration:none; font-size:13px; font-weight:600; background:var(--card); border:1px solid var(--border); padding:8px 14px; border-radius:99px; }
    .article-box, .box { background:var(--card); border:1px solid var(--border); border-radius:18px; padding:26px; margin-bottom:18px; }
    h1 { font-family:var(--serif); font-size:26px; line-height:1.3; color:var(--ink); margin-bottom:10px; }
    .meta { font-size:12.5px; color:var(--muted); margin-bottom:16px; padding-bottom:12px; border-bottom:1px dashed var(--border); }
    h2 { font-family:var(--serif); font-size:18px; color:var(--copper-dark); margin:18px 0 8px; }
    p { font-size:14.5px; margin-bottom:10px; }
    ul, ol { padding-left:22px; margin-bottom:12px; font-size:14px; }
    li { margin-bottom:6px; }
    .up-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:12px; margin-top:10px; }
    .up-card { text-decoration:none; background:rgba(255,251,240,.8); border:1px solid var(--border); padding:14px; border-radius:12px; color:var(--ink); display:flex; flex-direction:column; gap:4px; transition:all .2s; }
    .up-card:hover { border-color:var(--copper); transform:translateY(-2px); }
    .up-card span { color:var(--copper); font-weight:800; font-size:14px; margin-top:4px; }
    .other-list a { color:var(--copper-dark); text-decoration:none; font-weight:600; }
    .other-list a:hover { text-decoration:underline; }
    .cta-bar { display:flex; flex-wrap:wrap; gap:10px; margin-top:16px; }
    .btn { text-decoration:none; padding:10px 16px; border-radius:10px; font-weight:600; font-size:13px; color:#FFF7E8; background:linear-gradient(135deg,#DC8644,#B85922,#8C3A12); border:none; cursor:pointer; }
    .btn-wa { background:#25D366; }
    footer { text-align:center; font-size:12px; color:var(--muted); margin-top:24px; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="topbar">
      <a class="back" href="../../">← Quay lại Hue Sound Express</a>
      <span style="font-weight:800;font-size:13px;color:var(--copper);">CẨM NANG ANALOG AUDIO • HUE SOUND VINYL CARE</span>
    </div>
    <article class="article-box">
      <h1>$gTitleEsc</h1>
      <div class="meta">Biên soạn bởi <strong>Hue Sound Vinyl Care</strong> • Cập nhật tự động cùng danh mục sản phẩm mới nhất</div>
      $($gd.body)
      <div class="cta-bar">
        <a class="btn" href="https://zalo.me/0932313918" target="_blank" rel="noopener noreferrer">💬 Tư vấn kỹ thuật qua Zalo (09323.13918)</a>
        <a class="btn btn-wa" href="https://wa.me/84932313918" target="_blank" rel="noopener noreferrer">💬 WhatsApp</a>
      </div>
    </article>
    <section class="box">
      <h2>🎵 Sản phẩm liên quan tại Hue Sound Express</h2>
      <div class="up-grid">$relCards</div>
    </section>
    <section class="box">
      <h2>📖 Các bài Cẩm nang &amp; FAQ Analog Audio khác</h2>
      <ul class="other-list">$otherLinks</ul>
    </section>
    <footer>© 2026 Hue Sound Express – Where sound finds its color • Analog Audio • Hotline/Zalo: 09323.13918</footer>
  </div>
</body>
</html>
"@
        [System.IO.File]::WriteAllText((Join-Path $gDir "index.html"), $gHtml, $utf8NoBom)
        [void]$sm.AppendLine("  <url><loc>$gUrl</loc><priority>0.85</priority></url>")
    }

    [void]$sm.AppendLine('</urlset>')
    [System.IO.File]::WriteAllText((Join-Path (Get-Location) "sitemap.xml"), $sm.ToString(), $utf8NoBom)
    Write-Host "Da sinh $($allItems.Count) trang SEO /p/, $redirectCount trang redirect URL cu, 6 bai /cam-nang/ va sitemap.xml!" -ForegroundColor Green
}

Write-Host "Build hoan tat! Tong so san pham hop nhat: $($allItems.Count)" -ForegroundColor Green
Write-Host " - data\catalog_all.json"
Write-Host " - feed\products.json"
Write-Host " - llms.txt"
