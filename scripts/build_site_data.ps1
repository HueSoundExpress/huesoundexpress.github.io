<#
.SYNOPSIS
    Hue Sound Express - Single Source of Truth Builder (0d tren GitHub Pages)
    Doc du lieu chuan snake_case tu 4 bang CSV ->
    1. Hop nhat thanh data/catalog_all.json (ma noi bo tieng Anh snake_case, ASCII SKU, Slug co tien to, Keywords phong phu)
    2. Xuat feed/products.json (chuan OpenAI / Google Shopping Feed 0d + UTM sach khong dau)
    3. Cap nhat llms.txt chuan B2A (dan toi feed/products.json, 10 /bo-suu-tap/ & 6 /cam-nang/)
    4. Sinh 1107+ trang chi tiet SEO /p/<slug>/index.html + redirect stub URL cu
    5. Sinh 6 bai /cam-nang/<slug>/index.html + 10 trang bo suu tap /bo-suu-tap/<slug>/index.html & sitemap.xml
#>

param(
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

function Clean-PlainText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = [string]$Text
    $t = [regex]::Replace($t, '(?is)</?(?:p|div|span|a|i|ul|li|table|tr|td|br)\b.*$', '')
    $t = [regex]::Replace($t, '<[^>]+>', ' ')
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $t = [regex]::Replace($t, '[\r\n\t]+', ' ')
    return ([regex]::Replace($t, '\s{2,}', ' ')).Trim()
}

function Get-Slug {
    param([string]$Text)
    $clean = (Remove-Diacritics (Clean-PlainText $Text)).ToLowerInvariant()
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

function Get-VinylFallbackSvg {
    param([string]$Title, [string]$Artist, [string]$Fmt, [string]$Uid = "vf1", [bool]$Hidden = $false)
    $tRaw = ([string]$Title).Trim()
    $aRaw = ([string]$Artist).Trim()
    $tShort = if ($tRaw.Length -gt 26) { $tRaw.Substring(0, 25) + [char]0x2026 } else { $tRaw }
    $aShort = if ($aRaw.Length -gt 24) { $aRaw.Substring(0, 23) + [char]0x2026 } else { $aRaw }
    $fShort = if ($Fmt) { $Fmt } else { "LP" }
    $tTxt = Escape-Html $tShort
    $aTxt = Escape-Html $aShort
    $fTxt = Escape-Html $fShort
    $disp = if ($Hidden) { "display:none;" } else { "display:flex;" }
    return "<div class='fb-vinyl' style='$($disp)align-items:center;justify-content:center;width:100%;height:100%;padding:6px;'><svg viewBox='0 0 120 120' fill='none' style='width:100%;height:100%;max-width:190px;max-height:190px;'><defs><path id='$($Uid)t' d='M 16,60 A 44,44 0 1,1 104,60'/><path id='$($Uid)b' d='M 18,63 A 42,42 0 0,0 102,63'/></defs><circle cx='60' cy='60' r='55' stroke='#B5541F' stroke-width='1.6' opacity='.42'/><circle cx='60' cy='60' r='46' stroke='#C99A2E' stroke-width='1.1' opacity='.45' stroke-dasharray='2 2'/><circle cx='60' cy='60' r='34' stroke='#B5541F' stroke-width='1.4' opacity='.5'/><circle cx='60' cy='60' r='23' stroke='#C99A2E' stroke-width='1.3' opacity='.62'/><circle cx='60' cy='60' r='13' fill='rgba(181,84,31,.14)' stroke='#B5541F' stroke-width='1.2'/><circle cx='60' cy='60' r='4' fill='#B5541F'/><text font-family='Montserrat,sans-serif' font-size='7.1' font-weight='800' fill='#6E2E10' letter-spacing='.35'><textPath href='#$($Uid)t' startOffset='50%' text-anchor='middle'>$tTxt</textPath></text><text font-family='Lora,Georgia,serif' font-size='6.7' font-weight='600' fill='#8B4513' letter-spacing='.25'><textPath href='#$($Uid)b' startOffset='50%' text-anchor='middle'>$aTxt</textPath></text><text x='60' y='74' text-anchor='middle' font-family='Montserrat,sans-serif' font-size='6.2' font-weight='800' fill='#B5541F'>$fTxt</text></svg></div>"
}

function Load-CsvOrUrl {
    param([string]$Source)
    if ([string]::IsNullOrWhiteSpace($Source)) { return @() }
    if ($Source -match '^https?://') {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $raw = (Invoke-WebRequest -Uri $Source -UseBasicParsing -TimeoutSec 20).Content
        return ($raw | ConvertFrom-Csv)
    }
    if (Test-Path $Source) { return @(Import-Csv -Path $Source -Encoding UTF8) }
    return @()
}

function Normalize-AsciiSku {
    param([string]$RawSku, [string]$FallbackPrefix, [int]$Id, [System.Collections.Generic.HashSet[string]]$Used)
    $s = (Remove-Diacritics $RawSku).ToUpperInvariant()
    $s = [regex]::Replace($s, '[^A-Z0-9]+', '-').Trim('-')
    if (-not $s -or $s.Length -lt 3) { $s = "$FallbackPrefix-$Id" }
    $cand = $s; $n = 2
    while ($Used.Contains($cand)) {
        $cand = "$s-$Id"
        if ($Used.Contains($cand)) { $cand = "$s-$Id-$n"; $n++ }
    }
    [void]$Used.Add($cand)
    return $cand
}

function Normalize-GenreCode {
    param([string]$Val, [string]$Category)
    $v = ([string]$Val).Trim()
    if ($Category -eq "turntable") {
        if ($v -match '(?i)beginner|người mới') { return "beginner" }
        if ($v -match '(?i)mid_range|trung cấp') { return "mid_range" }
        return "audiophile"
    }
    if ($Category -eq "accessory" -or $Category -eq "care") {
        if ($v -match '(?i)record_cleaning|vệ sinh') { return "record_cleaning" }
        if ($v -match '(?i)sleeves|bao lót') { return "sleeves_protection" }
        if ($v -match '(?i)cartridge|kim') { return "cartridge_stylus" }
        if ($v -match '(?i)calibration|cân chỉnh') { return "calibration" }
        return "vinyl_care_service"
    }
    switch ($v) {
        "love_songs_1954_1975"      { return "love_songs_1954_1975" }
        "bolero_golden_oldies"      { return "bolero_golden_oldies" }
        "vietnamese_pop_modern"     { return "vietnamese_pop_modern" }
        "vietnamese_folk_patriotic" { return "vietnamese_folk_patriotic" }
        "instrumental"              { return "instrumental" }
        "jazz_vocal"                { return "jazz_vocal" }
        "audiophile_reference"      { return "audiophile_reference" }
        "rock_metal"                { return "rock_metal" }
        "pop_disco"                 { return "pop_disco" }
        "classic_soundtrack"        { return "classic_soundtrack" }
        "Tình khúc 1954–1975"       { return "love_songs_1954_1975" }
        "Nhạc Việt / Bolero"        { return "bolero_golden_oldies" }
        "Nhạc Việt trẻ"             { return "vietnamese_pop_modern" }
        "Nhạc Việt"                 { return "vietnamese_pop_modern" }
        "Dân ca & Cách mạng"        { return "vietnamese_folk_patriotic" }
        "Instrument"                { return "instrumental" }
        "Jazz & Vocal"              { return "jazz_vocal" }
        "Audiophile & Khác"         { return "audiophile_reference" }
        "Rock & Metal"              { return "rock_metal" }
        "Pop & Disco"               { return "pop_disco" }
        "Classic & Soundtrack"      { return "classic_soundtrack" }
        default                     { return "audiophile_reference" }
    }
}

function Normalize-CondCode {
    param([string]$Val)
    $v = ([string]$Val).Trim()
    if ($v -in @("mint_sealed","near_mint","very_good_plus","very_good","good_plus","good","fair","poor")) { return $v }
    if ($v -match '^(?i)vg\+$') { return "very_good_plus" }
    if ($v -match '^(?i)vg$') { return "very_good" }
    if ($v -match '^(?i)g\+$') { return "good_plus" }
    if ($v -match '^(?i)g$') { return "good" }
    if ($v -match '^(?i)(f|fair)$') { return "fair" }
    if ($v -match '^(?i)(p|poor)$') { return "poor" }
    if ($v -match '(?i)\b(nm|mint)\b' -and $v -notmatch '(?i)new\s*seal') { return "near_mint" }
    return "mint_sealed"
}

function Normalize-StatusCode {
    param([string]$Val, [int]$Price)
    $v = ([string]$Val).Trim()
    if ($Price -le 0 -or $v -match '(?i)(out_of_stock|hết hàng|sắp có|đặt trước)') { return "out_of_stock" }
    return "in_stock"
}

function Normalize-MoodArray {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }
    $out = New-Object System.Collections.ArrayList
    foreach ($p in ($Raw -split '\s*;\s*' | Where-Object { $_ })) {
        $code = switch ($p.Trim()) {
            "late_night_chill"        { "late_night_chill" }
            "nostalgic_classic"       { "nostalgic_classic" }
            "audiophile_reference"    { "audiophile_reference" }
            "weekend_upbeat"          { "weekend_upbeat" }
            "refined_gifting"         { "refined_gifting" }
            "morning_coffee_acoustic" { "morning_coffee_acoustic" }
            "Đêm muộn & Thư giãn"     { "late_night_chill" }
            "Hoài niệm & Nhạc xưa"    { "nostalgic_classic" }
            "Audiophile & Phê âm"     { "audiophile_reference" }
            "Năng lượng & Cuối tuần"  { "weekend_upbeat" }
            "Quà tặng tinh tế"        { "refined_gifting" }
            "Cà phê sáng & Acoustic"  { "morning_coffee_acoustic" }
            default                   { "" }
        }
        if ($code -and -not $out.Contains($code)) { [void]$out.Add($code) }
    }
    return @($out)
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
    param([string]$Code)
    switch ($Code) {
        "love_songs_1954_1975"      { return "Tình khúc 1954–1975" }
        "bolero_golden_oldies"      { return "Trữ tình" }
        "vietnamese_pop_modern"     { return "Nhạc Việt trẻ" }
        "vietnamese_folk_patriotic" { return "Dân ca & Cách mạng" }
        "instrumental"              { return "Instrument" }
        "jazz_vocal"                { return "Jazz & Vocal" }
        "audiophile_reference"      { return "Audiophile" }
        "rock_metal"                { return "Rock & Metal" }
        "pop_disco"                 { return "Pop & Disco" }
        "classic_soundtrack"        { return "Classic & OST" }
        "beginner"                  { return "Người mới bắt đầu" }
        "mid_range"                 { return "Trung cấp" }
        "audiophile"                { return "Audiophile" }
        "record_cleaning"           { return "Vệ sinh đĩa" }
        "sleeves_protection"        { return "Bao lót & Bìa bảo vệ" }
        "cartridge_stylus"          { return "Kim & Máng" }
        "calibration"               { return "Cân chỉnh" }
        "vinyl_care_service"        { return "Hue Sound Vinyl Care" }
        default                     { return $Code }
    }
}

function Get-CondDisplayVi {
    param([string]$Code)
    switch ($Code) {
        "mint_sealed"    { return "Mới 100% (M)" }
        "near_mint"      { return "NM" }
        "very_good_plus" { return "VG+" }
        "very_good"      { return "VG" }
        "good_plus"      { return "G+" }
        "good"           { return "G" }
        "fair"           { return "F" }
        "poor"           { return "P" }
        default          { return "Mới 100% (M)" }
    }
}

function Get-StatusDisplayVi {
    param([string]$Code)
    if ($Code -eq "in_stock") { return "Còn hàng" }
    return "Hết hàng"
}

function Get-ItemOriginCode {
    param($Item)
    $sku = ([string]$Item.sku).ToUpperInvariant()
    if ($sku.StartsWith("NN-")) { return "INTL" }
    if ($sku.StartsWith("VN-") -or $sku.StartsWith("CD-VN-") -or $sku.StartsWith("CD-GDA-") -or $sku.StartsWith("CD-SSR-") -or $sku.StartsWith("CD-SSS-")) { return "VN" }
    $vnGenres = @("love_songs_1954_1975","bolero_golden_oldies","vietnamese_pop_modern","vietnamese_folk_patriotic")
    if ($vnGenres -contains $Item.g) { return "VN" }
    return "INTL"
}

function Format-LinkedDescription {
    param(
        [string]$DescText,
        $Item,
        [string]$CatParam,
        [string]$OriginCode,
        [string]$GenreDisplay,
        [string]$CatUrl,
        [string]$OriginUrl,
        [string]$GenreUrl,
        [string]$ArtistUrl,
        [string]$LabelUrl
    )
    if ([string]::IsNullOrWhiteSpace($DescText)) { return "" }
    $cands = New-Object System.Collections.ArrayList

    if ($Item.a -and $Item.a.Length -ge 2) {
        [void]$cands.Add([PSCustomObject]@{ txt = $Item.a; url = $ArtistUrl; tip = "Xem tất cả sản phẩm của $($Item.a)" })
    }
    if ($LabelUrl -and $Item.l -and $Item.l.Length -ge 2 -and $Item.l -notmatch '^(?i)(other|việt nam)$') {
        [void]$cands.Add([PSCustomObject]@{ txt = $Item.l; url = $LabelUrl; tip = "Xem tất cả ấn phẩm của hãng $($Item.l)" })
    }
    if ($OriginUrl) {
        if ($OriginCode -eq "VN") {
            [void]$cands.Add([PSCustomObject]@{ txt = "nhạc Việt Nam"; url = $OriginUrl; tip = "Xem toàn bộ Nhạc Việt" })
            [void]$cands.Add([PSCustomObject]@{ txt = "Nhạc Việt"; url = $OriginUrl; tip = "Xem toàn bộ Nhạc Việt" })
        } else {
            [void]$cands.Add([PSCustomObject]@{ txt = "nhạc Quốc tế"; url = $OriginUrl; tip = "Xem toàn bộ Nhạc Quốc Tế" })
            [void]$cands.Add([PSCustomObject]@{ txt = "Nhạc Quốc Tế"; url = $OriginUrl; tip = "Xem toàn bộ Nhạc Quốc Tế" })
        }
    }
    if ($GenreDisplay -and $GenreDisplay.Length -ge 3) {
        [void]$cands.Add([PSCustomObject]@{ txt = $GenreDisplay; url = $GenreUrl; tip = "Xem thể loại $GenreDisplay" })
    }
    $genreAliases = switch ($Item.g) {
        "classic_soundtrack"        { @("Classic & Nhạc phim OST", "Nhạc phim OST", "Cổ điển") }
        "bolero_golden_oldies"      { @("Nhạc Vàng & Bolero", "Nhạc Bolero", "Nhạc Trữ Tình", "Trữ tình") }
        "love_songs_1954_1975"      { @("Tình khúc 1954–1975", "Tình khúc 1954-1975", "Tiền chiến & Trữ tình", "Tiền chiến") }
        "vietnamese_pop_modern"     { @("Nhạc Việt trẻ") }
        "vietnamese_folk_patriotic" { @("Dân ca & Cách mạng") }
        "instrumental"              { @("Hòa tấu Violin Instrument", "Hòa tấu", "Instrument") }
        "jazz_vocal"                { @("Jazz & Vocal") }
        "audiophile_reference"      { @("Audiophile") }
        default                     { @() }
    }
    foreach ($ga in $genreAliases) {
        [void]$cands.Add([PSCustomObject]@{ txt = $ga; url = $GenreUrl; tip = "Xem thể loại $GenreDisplay" })
    }
    if ($Item.PSObject.Properties['genres'] -and $Item.genres -and $Item.genres.Count -gt 1) {
        foreach ($eg in ($Item.genres | Select-Object -Skip 1)) {
            $egDisp = Get-GenreDisplayVi ([string]$eg)
            $egUrl = if ($OriginUrl) { "../../?cat=$CatParam&origin=$OriginCode&genre=$([Uri]::EscapeDataString([string]$eg))" } else { "../../?cat=$CatParam&genre=$([Uri]::EscapeDataString([string]$eg))" }
            [void]$cands.Add([PSCustomObject]@{ txt = $egDisp; url = $egUrl; tip = "Xem thể loại $egDisp" })
            if ($eg -eq "instrumental") {
                [void]$cands.Add([PSCustomObject]@{ txt = "Hòa tấu Cello Instrument"; url = $egUrl; tip = "Xem thể loại $egDisp" })
            }
        }
    }
    $catPhrases = switch ($Item.category) {
        "vinyl"     { @("Đĩa than") }
        "cd"        { @("Đĩa CD") }
        "turntable" { @("Máy hát đĩa than", "Mâm than") }
        default     { @("Phụ kiện đĩa than", "Hue Sound Vinyl Care") }
    }
    foreach ($cp in $catPhrases) {
        [void]$cands.Add([PSCustomObject]@{ txt = $cp; url = $CatUrl; tip = "Xem danh mục" })
    }
    $namedArtists = @("London Symphony Orchestra","Trịnh Công Sơn","Phạm Duy","Ngô Thụy Miên","Vũ Thành An","Lam Phương","Nguyễn Ánh 9","Từ Công Phụng","Lê Uyên Phương","Trúc Phương","Trần Thiện Thanh","Hoàng Thi Thơ","Phạm Đình Chương","Quốc Bảo","Bảo Chấn","Đức Huy","Đức Trí","Phú Quang","Trần Tiến","Dương Thụ","Thanh Tùng","Văn Cao","Dzoãn Mẫn","Dương Thiệu Tước","Nguyễn Văn Tý","Đỗ Bảo","Thành Nghiệp","Thân Sao Mai","Thanh Lam","Hồng Nhung","Mỹ Linh","Bằng Kiều","Nguyên Thảo","Khánh Linh","Nguyễn Trần Minh Đức","Trần Thu Hà","Khánh Hà","Tuấn Ngọc","Thanh Tuyền","Hoàng Oanh","Chế Linh","Phương Dung","Trường Vũ","Như Quỳnh","Lân Nhã","Uyên Linh","Lê Hiếu","Quang Dũng","Trọng Bắc","Hoàng Yến Chibi","Thùy Chi","Quốc Thiên","Vũ Thắng Lợi","Khánh Ly","Lệ Thu","Thái Hiền","Elvis Phương","Duy Khánh","Hương Lan","Tuấn Vũ","Sơn Tuyền","Nhật Trường","Trần Thái Hòa","Thu Phương","Ngọc Hạ","Cẩm Vân","Văn Mai Hương","AMEE","Hoàng Dũng","Quỳnh Anh")
    foreach ($na in $namedArtists) {
        if ($DescText -match [regex]::Escape($na)) {
            $qUrl = "../../?cat=$CatParam&q=$([Uri]::EscapeDataString($na))"
            [void]$cands.Add([PSCustomObject]@{ txt = $na; url = $qUrl; tip = "Tìm các ấn phẩm liên quan $na" })
        }
    }

    $sortedCands = $cands | Sort-Object { $_.txt.Length } -Descending
    $work = $DescText
    $tokens = New-Object System.Collections.ArrayList
    $usedUrls = New-Object 'System.Collections.Generic.HashSet[string]'
    $idx = 0

    foreach ($c in $sortedCands) {
        if ($usedUrls.Contains($c.url)) { continue }
        $rx = "(?i)(?<![\p{L}\p{N}])" + [regex]::Escape($c.txt) + "(?![\p{L}\p{N}])"
        $m = [regex]::Match($work, $rx)
        if ($m.Success) {
            $tok = "___XLNK$($idx)___"
            $matchedEsc = Escape-Html $m.Value
            $tipEsc = Escape-Html $c.tip
            [void]$tokens.Add("<a class='desc-link' href='$($c.url)' title='$tipEsc'>$matchedEsc</a>")
            $work = $work.Substring(0, $m.Index) + $tok + $work.Substring($m.Index + $m.Length)
            [void]$usedUrls.Add($c.url)
            $idx++
        }
    }

    $outHtml = Escape-Html $work
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $outHtml = $outHtml.Replace("___XLNK$($i)___", $tokens[$i])
    }
    return $outHtml
}

# Trich xuat tu khoa giau ngu nghia: Ca si, Nhac si, Nhac cu, Bai hat, Chu de trong album
function Get-ProductKeywords {
    param($Item)
    $set = New-Object 'System.Collections.Generic.List[string]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $addKw = {
        param([string]$k)
        $clean = ([string]$k).Trim()
        if ($clean -and $clean.Length -ge 2 -and -not $seen.Contains($clean)) {
            [void]$seen.Add($clean)
            [void]$set.Add($clean)
        }
    }

    & $addKw "Analog Audio"
    if ($Item.category -eq "vinyl") { & $addKw "Đĩa than"; & $addKw "Vinyl LP" }
    elseif ($Item.category -eq "cd") { & $addKw "Đĩa CD"; & $addKw "CD Nhạc Việt" }
    elseif ($Item.category -eq "turntable") { & $addKw "Máy hát đĩa than"; & $addKw "Mâm than"; & $addKw "Turntable" }
    else { & $addKw "Phụ kiện đĩa than"; & $addKw "Hue Sound Vinyl Care" }

    # 1. Ca si / Nghe si & Ten album & Hang dia
    & $addKw $Item.a
    & $addKw $Item.t
    if ($Item.l -and $Item.l -ne "Việt Nam") { & $addKw $Item.l }

    $corpus = "$($Item.a) $($Item.t) $($Item.d) $($Item.f) $(($Item.tracklist) -join ' ')"

    # 2. Nhac si xuat hien trong album / tracklist
    $composers = @(
        @{ rx = 'trịnh công sơn|\btcs\b|diễm xưa|vết lăn trầm|bống là ai|xin mặt trời ngủ yên'; lbl = 'Trịnh Công Sơn' },
        @{ rx = 'phạm duy|áo lụa hà đông|trả lại em yêu|requiem'; lbl = 'Phạm Duy' },
        @{ rx = 'ngô thụy miên|mắt biếc|niệm khúc cuối|riêng một góc trời'; lbl = 'Ngô Thụy Miên' },
        @{ rx = 'vũ thành an|bài không tên'; lbl = 'Vũ Thành An' },
        @{ rx = 'lam phương|thành phố buồn|biển tình|phút cuối'; lbl = 'Lam Phương' },
        @{ rx = 'nguyễn ánh 9|buồn ơi chào mi|cô đơn|lặng lẽ tiếng dương cầm'; lbl = 'Nguyễn Ánh 9' },
        @{ rx = 'từ công phụng|bây giờ tháng mấy|mắt lệ cho người'; lbl = 'Từ Công Phụng' },
        @{ rx = 'lê uyên phương|dạ khúc cho tình nhân|vũng lầy của chúng ta'; lbl = 'Lê Uyên Phương' },
        @{ rx = 'trúc phương|nửa đêm ngoài phố|tàu đêm năm cũ'; lbl = 'Trúc Phương' },
        @{ rx = 'trần thiện thanh|hoa trinh nữ|lâu đài tình ái'; lbl = 'Trần Thiện Thanh' },
        @{ rx = 'hoàng thi thơ|đường xưa lối cũ|tạ từ trong đêm'; lbl = 'Hoàng Thi Thơ' },
        @{ rx = 'phạm đình chương|ly rượu mừng|mộng dưới hoa'; lbl = 'Phạm Đình Chương' },
        @{ rx = 'quốc bảo|bình yên|em về tinh khôi'; lbl = 'Quốc Bảo' },
        @{ rx = 'bảo chấn|nơi ấy bình yên|giấc mơ tuyệt vời'; lbl = 'Bảo Chấn' },
        @{ rx = 'đức huy|đừng xa em đêm nay|như đã dấu yêu'; lbl = 'Đức Huy' },
        @{ rx = 'phú quang|em ơi hà nội phố|biển nỗi nhớ và em|điều giản dị|mơ về nơi xa lắm'; lbl = 'Phú Quang' },
        @{ rx = 'trần tiến|sắc màu|mặt trời bé con|quê nhà|về đi em'; lbl = 'Trần Tiến' },
        @{ rx = 'dương thụ|cho em một ngày|tháng tư về|nghe mưa|lắng nghe mùa xuân về'; lbl = 'Dương Thụ' },
        @{ rx = 'thanh tùng|giọt nắng bên thềm|lối cũ ta về|hoa cúc vàng|một mình'; lbl = 'Thanh Tùng' },
        @{ rx = 'anh quân|huy tuấn|tóc ngắn|hương ngọc lan|trưa vắng'; lbl = 'Anh Quân & Huy Tuấn' },
        @{ rx = 'hoàng trọng|ngàn thu áo tím'; lbl = 'Hoàng Trọng' },
        @{ rx = 'văn cao|đoàn chuẩn|từ linh|gửi gió cho mây ngàn bay'; lbl = 'Nhạc Tiền Chiến' },
        @{ rx = 'joe hisaishi|ghibli|totoro|mononoke|howl'; lbl = 'Joe Hisaishi' },
        @{ rx = 'beethoven|mozart|chopin|vivaldi|bach|tchaikovsky|strauss'; lbl = 'Nhạc Cổ Điển' }
    )
    foreach ($c in $composers) { if ($corpus -match "(?i)$($c.rx)") { & $addKw $c.lbl } }

    # 3. Nhac cu xuat hien
    $instruments = @(
        @{ rx = 'piano|dương cầm'; lbl = 'Piano' },
        @{ rx = 'saxophone|kèn sax'; lbl = 'Saxophone' },
        @{ rx = 'guitar|ghi-ta'; lbl = 'Guitar' },
        @{ rx = 'violin|vĩ cầm|viollage|qinie'; lbl = 'Violin' },
        @{ rx = 'cello|2cellos'; lbl = 'Cello' },
        @{ rx = 'đàn bầu|đàn tranh|nhạc cụ dân tộc|trần văn khê|ngô hồng quang'; lbl = 'Nhạc cụ dân tộc' },
        @{ rx = 'trumpet|miles davis|chet baker'; lbl = 'Trumpet' },
        @{ rx = 'acoustic|unplugged'; lbl = 'Acoustic' }
    )
    foreach ($ins in $instruments) { if ($corpus -match "(?i)$($ins.rx)") { & $addKw $ins.lbl } }

    # 4. Chu de trong album (Tet, Ngay le, Noel, Ha Noi, Sai Gon, Mua thu, Bolero, Tinh khuc...)
    $themes = @(
        @{ rx = 'xuân|tết|ly rượu mừng|mùa xuân|chúc xuân|em ơi mùa xuân'; lbl = 'Nhạc Tết & Mùa Xuân' },
        @{ rx = 'christmas|noel|giáng sinh|mùa sao sáng|neujahr|new year|holiday|winter'; lbl = 'Nhạc Ngày Lễ & Giáng Sinh' },
        @{ rx = 'hà nội|hanoi|phố cổ'; lbl = 'Nhạc Hà Nội' },
        @{ rx = 'sài gòn|saigon'; lbl = 'Nhạc Sài Gòn' },
        @{ rx = 'mùa thu|thu vàng|heo may'; lbl = 'Nhạc Mùa Thu' },
        @{ rx = 'quê hương|dân ca|xứ huế|xứ nghệ|hà tĩnh'; lbl = 'Nhạc Quê Hương' },
        @{ rx = 'ost|soundtrack|nhạc phim|score'; lbl = 'Nhạc Phim OST' }
    )
    foreach ($th in $themes) { if ($corpus -match "(?i)$($th.rx)") { & $addKw $th.lbl } }

    # The loai ngu nghia & SEO / AI Agent keywords
    & $addKw (Get-GenreDisplayVi $Item.g)
    if ($Item.category -eq "vinyl") {
        & $addKw "Đĩa than Vinyl LP"
        & $addKw "Đĩa than chính hãng"
        & $addKw "Cửa hàng đĩa than Hue Sound Express"
        & $addKw "Nơi bán đĩa nhạc uy tín"
    } elseif ($Item.category -eq "cd") {
        & $addKw "Đĩa CD chính hãng"
        & $addKw "Cửa hàng đĩa than Hue Sound Express"
        & $addKw "Nơi bán đĩa nhạc uy tín"
    }
    if ($Item.g -eq "bolero_golden_oldies") { & $addKw "Bolero"; & $addKw "Nhạc Vàng"; & $addKw "Nhạc Trữ Tình" }
    elseif ($Item.g -eq "love_songs_1954_1975") { & $addKw "Tình khúc 1954–1975"; & $addKw "Nhạc Xưa"; & $addKw "Tình Ca" }
    elseif ($Item.g -eq "instrumental") { & $addKw "Instrument"; & $addKw "Hòa tấu nhạc cụ" }

    # 5. Cac bai hat noi bat trong tracklist (toi da 8 bai)
    if ($Item.tracklist -and $Item.tracklist.Count -gt 0) {
        foreach ($tr in ($Item.tracklist | Select-Object -First 8)) { & $addKw $tr }
    }

    return @($set)
}

# Tu dong chuan hoa ten file trong images/ & audio/ sang kebab-case neu ton tai script
if (Test-Path "scripts\normalize_media_names.ps1") {
    & ".\scripts\normalize_media_names.ps1" -TargetFolders @("images","audio") | Out-Null
}

$allItems = New-Object System.Collections.ArrayList
$usedSkus = New-Object 'System.Collections.Generic.HashSet[string]'

# 1. NAP DIA THAN (VINYL)
foreach ($r in (Load-CsvOrUrl $VinylCsv)) {
    if (-not $r.id) { continue }
    $idNum = [int]$r.id
    $tracks = if ($r.tracklist) { @($r.tracklist -split '\s*\|\s*' | Where-Object { $_ }) } else { @() }
    $missing = if ($r.missing_fields) { @($r.missing_fields -split '\s*;\s*' | Where-Object { $_ }) } else { @() }
    $legacySlug = if ($r.PSObject.Properties['legacy_slug'] -and $r.legacy_slug) { [string]$r.legacy_slug } elseif ($r.slug) { [string]$r.slug } else { "" }
    $isContact = ($r.PSObject.Properties['contact_only'] -and ($r.contact_only -eq "1" -or $r.contact_only -eq "true"))
    $vinylWeight = if ($r.PSObject.Properties['weight'] -and -not [string]::IsNullOrWhiteSpace($r.weight)) { ([string]$r.weight).Trim() } else { "180g" }
    $rawGenres = @(([string]$r.genre) -split '\s*;\s*' | Where-Object { $_ } | ForEach-Object { Normalize-GenreCode -Val $_ -Category "vinyl" } | Select-Object -Unique)
    if ($rawGenres.Count -eq 0) { $rawGenres = @("audiophile_reference") }
    [void]$allItems.Add([PSCustomObject]@{
        id             = $idNum
        sku            = Normalize-AsciiSku -RawSku ([string]$r.sku) -FallbackPrefix "VN" -Id $idNum -Used $usedSkus
        slug           = Get-CanonicalSlug -Category "vinyl" -Artist ([string]$r.artist) -Title ([string]$r.title) -Id $idNum
        legacy_slug    = $legacySlug
        category       = "vinyl"
        g              = $rawGenres[0]
        genres         = $rawGenres
        mood           = Normalize-MoodArray ([string]$r.mood)
        a              = [string]$r.artist
        t              = [string]$r.title
        l              = [string]$r.label
        f              = [string]$r.format
        w              = $vinylWeight
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
        missing_fields = $missing
    })
}

# 2. NAP DIA CD / SACD
foreach ($r in (Load-CsvOrUrl $CdCsv)) {
    if (-not $r.sku) { continue }
    $idNum = [int]$r.id
    $tracks = if ($r.tracklist) { @($r.tracklist -split '\s*\|\s*' | Where-Object { $_ }) } else { @() }
    $legacySlug = if ($r.PSObject.Properties['legacy_slug'] -and $r.legacy_slug) { [string]$r.legacy_slug } elseif ($r.slug) { [string]$r.slug } else { "" }
    $isContact = ($r.PSObject.Properties['contact_only'] -and ($r.contact_only -eq "1" -or $r.contact_only -eq "true"))
    $rawCdGenres = @(([string]$r.genre) -split '\s*;\s*' | Where-Object { $_ } | ForEach-Object { Normalize-GenreCode -Val $_ -Category "cd" } | Select-Object -Unique)
    if ($rawCdGenres.Count -eq 0) { $rawCdGenres = @("vietnamese_pop_modern") }
    [void]$allItems.Add([PSCustomObject]@{
        id             = $idNum
        sku            = Normalize-AsciiSku -RawSku ([string]$r.sku) -FallbackPrefix "CD-VN" -Id $idNum -Used $usedSkus
        slug           = Get-CanonicalSlug -Category "cd" -Artist ([string]$r.artist) -Title ([string]$r.title) -Id $idNum
        legacy_slug    = $legacySlug
        category       = "cd"
        g              = $rawCdGenres[0]
        genres         = $rawCdGenres
        mood           = Normalize-MoodArray ([string]$r.mood)
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
    $isContact = ($r.PSObject.Properties['contact_only'] -and ($r.contact_only -eq "1" -or $r.contact_only -eq "true"))
    [void]$allItems.Add([PSCustomObject]@{
        id               = $idNum
        sku              = Normalize-AsciiSku -RawSku ([string]$r.sku) -FallbackPrefix "TT" -Id $idNum -Used $usedSkus
        slug             = Get-CanonicalSlug -Category "turntable" -Artist ([string]$r.brand) -Title ([string]$r.title) -Id $idNum
        legacy_slug      = $legacySlug
        category         = "turntable"
        g                = Normalize-GenreCode -Val ([string]$r.level) -Category "turntable"
        mood             = @("refined_gifting")
        a                = [string]$r.brand
        t                = [string]$r.title
        l                = [string]$r.drive_type
        f                = [string]$r.cartridge_type
        c                = "mint_sealed"
        has_phono        = if ([string]$r.has_phono -match '(?i)yes|có') { "yes" } else { "no" }
        has_bluetooth    = if ([string]$r.has_bluetooth -match '(?i)yes|có') { "yes" } else { "no" }
        specs            = $specs
        recommended_skus = $recSkus
        p                = if ($r.list_price) { [int]$r.list_price } else { 0 }
        s                = [string]$r.status
        contact_only     = $isContact
        d                = [string]$r.description
        img              = [string]$r.cover_url
        video_demo_url   = [string]$r.video_demo_url
    })
}

# 4. NAP PHU KIEN & HUE SOUND VINYL CARE
foreach ($r in (Load-CsvOrUrl $AccessoryCsv)) {
    if (-not $r.sku) { continue }
    $idNum = [int]$r.id
    $cat = if ($r.category) { ([string]$r.category).Trim().ToLowerInvariant() } else { "accessory" }
    $legacySlug = if ($r.PSObject.Properties['legacy_slug'] -and $r.legacy_slug) { [string]$r.legacy_slug } elseif ($r.slug) { [string]$r.slug } else { "" }
    $isContact = ($r.PSObject.Properties['contact_only'] -and ($r.contact_only -eq "1" -or $r.contact_only -eq "true" -or $r.contact_only -eq "2"))
    $isForceContact = ($r.PSObject.Properties['contact_only'] -and $r.contact_only -eq "2")
    [void]$allItems.Add([PSCustomObject]@{
        id            = $idNum
        sku           = Normalize-AsciiSku -RawSku ([string]$r.sku) -FallbackPrefix "ACC" -Id $idNum -Used $usedSkus
        slug          = Get-CanonicalSlug -Category $cat -Artist ([string]$r.brand) -Title ([string]$r.title) -Id $idNum
        legacy_slug   = $legacySlug
        category      = $cat
        g             = Normalize-GenreCode -Val ([string]$r.sub_type) -Category $cat
        mood          = @("refined_gifting")
        a             = [string]$r.brand
        t             = [string]$r.title
        l             = [string]$r.level
        f             = [string]$r.compatibility
        c             = "mint_sealed"
        p             = if ($r.list_price) { [int]$r.list_price } else { 0 }
        s             = [string]$r.status
        contact_only  = $isContact
        force_contact = $isForceContact
        d             = [string]$r.description
        img           = [string]$r.cover_url
    })
}

# Chuan hoa toan bo ma noi bo snake_case, lam sach HTML rac, gia giam 50-60k & trich xuat keywords
foreach ($it in $allItems) {
    $it.a = Clean-PlainText ([string]$it.a)
    $it.t = Clean-PlainText ([string]$it.t)
    $it.l = Clean-PlainText ([string]$it.l)
    $it.d = Clean-PlainText ([string]$it.d)
    if (([string]$it.a).Trim() -match '^(?i)nhiều ca sĩ$') { $it.a = "Nhiều nghệ sĩ" }
    $it.slug = Get-CanonicalSlug -Category $it.category -Artist $it.a -Title $it.t -Id $it.id

    $lbl = ([string]$it.l).Trim()
    if ($lbl -match '^(?i)(hãng đĩa\s+)?thời đại$') { $lbl = "Thời Đại" }
    elseif ($lbl -match '^(?i)stockfish$') { $lbl = "Stockfisch Records" }
    elseif ($lbl -match '^(?i)bk production$') { $lbl = "BK Production" }
    elseif ($lbl -match '^(?i)viết tân audio$') { $lbl = "Viết Tân Studio" }
    elseif ($lbl -match '^(?i)music on vinyl$') { $lbl = "Music On Vinyl" }
    $it.l = $lbl

    $origP = [int]$it.p
    $normCond = Normalize-CondCode ([string]$it.c)
    $normSt   = Normalize-StatusCode ([string]$it.s) $origP

    if ($it.PSObject.Properties['cond']) { $it.cond = $normCond }
    else { $it | Add-Member -NotePropertyName 'cond' -NotePropertyValue $normCond }
    $it.s = $normSt

    if ($it.category -eq "cd" -and $origP -gt 0 -and $origP -lt 150000) {
        $origP = 150000
    }
    $saleP = $origP
    if ($it.category -eq "cd" -and $origP -ge 200000 -and $origP -le 300000) {
        $saleP = $origP - 20000
    } elseif ($origP -ge 200000) {
        $disc = if (($origP % 100000) -eq 50000) { 60000 } else { 50000 }
        $saleP = $origP - $disc
    } elseif ($origP -ge 100000) {
        $saleP = $origP - 20000
    } elseif ($origP -gt 10000) {
        $saleP = $origP - 10000
    }
    $regularSaleP = $saleP
    if ($it.sku -eq "VN-QVTH-DUONGVE-560" -and $origP -gt 100000) {
        $novDeadline = [DateTime]::Parse("2026-11-30T23:59:59+07:00")
        $inNovPromo = ((Get-Date) -le $novDeadline)
        $promoP = $origP - 100000
        if ($inNovPromo) {
            $saleP = $promoP
            $it | Add-Member -NotePropertyName 'promo_until' -NotePropertyValue '2026-11-30' -Force
            $it | Add-Member -NotePropertyName 'p_regular' -NotePropertyValue $regularSaleP -Force
            $it | Add-Member -NotePropertyName 'promo_note' -NotePropertyValue 'Ưu đãi tháng 11: giảm 100k (từ 01/12 giảm 50k)' -Force
        }
    }
    if ($it.PSObject.Properties['p_orig']) { $it.p_orig = $origP }
    else { $it | Add-Member -NotePropertyName 'p_orig' -NotePropertyValue $origP }
    $it.p = $saleP

    $kwList = Get-ProductKeywords -Item $it
    if ($it.PSObject.Properties['kw']) { $it.kw = $kwList }
    else { $it | Add-Member -NotePropertyName 'kw' -NotePropertyValue $kwList }
}

# Sap xep uu tien san pham moi them len tren cung (giu nguyen thu tu cu)
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

# Dinh nghia 11 trang bo suu tap tinh (/bo-suu-tap/<slug>/)
$collectionMeta = @(
    @{ slug = "dia-than-nhac-tet";             utm = "dia_than_nhac_tet";             title = "Đĩa Than Nhạc Tết & Mùa Xuân – Analog Audio"; desc = "Tuyển tập đĩa than (Vinyl LP) & CD chủ đề Mùa Xuân, Nhạc Tết, Ly Rượu Mừng và những giai điệu rộn ràng đón năm mới tại Hue Sound Express." },
    @{ slug = "dia-than-ngay-le";              utm = "dia_than_ngay_le";              title = "Đĩa Than Ngày Lễ, Giáng Sinh (Noel) & Năm Mới"; desc = "Bộ sưu tập đĩa than Giáng Sinh (Christmas / Noel), Hòa nhạc Năm Mới (Neujahrskonzert) và các ấn phẩm quà tặng dịp lễ hội tinh tế." },
    @{ slug = "dia-than-nhac-trinh-cong-son";  utm = "dia_than_nhac_trinh_cong_son";  title = "Đĩa Than & CD Nhạc Trịnh Công Sơn"; desc = "Tuyển tập đĩa than (Vinyl LP) & CD nhạc Trịnh Công Sơn bất hủ qua các giọng ca Khánh Ly, Hồng Nhung, Lệ Thu, Cẩm Vân, Quang Dũng, Trần Mạnh Tuấn... với Diễm Xưa, Cát Bụi, Hạ Trắng, Biển Nhớ, Mưa Hồng, Tuổi Đá Buồn." },
    @{ slug = "dia-than-nhac-vang-bolero";     utm = "dia_than_nhac_vang_bolero";     title = "Đĩa Than Nhạc Vàng, Bolero & Trữ Tình"; desc = "Kho đĩa than & CD Nhạc Vàng, Bolero, Nhạc Trữ Tình tuyển chọn (Chế Linh, Thanh Tuyền, Duy Khánh, Tuấn Vũ, Hoàng Oanh, Như Quỳnh, Lệ Quyên...)." },
    @{ slug = "dia-than-tinh-khuc-1954-1975";  utm = "dia_than_tinh_khuc_1954_1975";  title = "Đĩa Than Tình Khúc 1954–1975 & Nhạc Tiền Chiến"; desc = "Tuyển tập Tình khúc 1954–1975 & Nhạc Tiền Chiến bất hủ của Phạm Duy, Trịnh Công Sơn, Ngô Thụy Miên, Vũ Thành An, Nguyễn Ánh 9, Từ Công Phụng..." },
    @{ slug = "dia-than-nhac-viet-tre";        utm = "dia_than_nhac_viet_tre";        title = "Đĩa Than Nhạc Việt Trẻ & Đương Đại"; desc = "Các ấn phẩm đĩa than Nhạc Việt trẻ hiện đại từ Mỹ Tâm, Hà Anh Tuấn, Quỳnh Anh, SOOBIN, AMEE, Văn Mai Hương, Gia Định Audio, Thời Đại..." },
    @{ slug = "dia-than-instrument-hoa-tau";   utm = "dia_than_instrument_hoa_tau";   title = "Đĩa Than Instrument & Hòa Tấu Nhạc Cụ"; desc = "Đĩa than & CD hòa tấu nhạc cụ (Piano, Saxophone, Violin, Guitar, Nhạc cụ dân tộc): Viollage – Qinie, Piano Nguyễn Ánh 9, Đặng Thái Sơn, Trần Mạnh Tuấn, GS Trần Văn Khê." },
    @{ slug = "dia-than-jazz-audiophile";      utm = "dia_than_jazz_audiophile";      title = "Đĩa Than Jazz, Vocal & Audiophile Tham Chiếu"; desc = "Kho đĩa than Jazz, Vocal và Audiophile thu âm tham chiếu từ các hãng đĩa danh tiếng: Stockfisch Records, Blue Note, Verve, Venus Records..." },
    @{ slug = "may-hat-dia-than-cho-nguoi-moi";utm = "may_hat_dia_than_cho_nguoi_moi";title = "Máy Hát Đĩa Than (Turntable) Cho Người Mới Bắt Đầu"; desc = "Gợi ý các dòng máy hát đĩa than (Turntable) tích hợp Phono Preamp, dễ chơi, an toàn cho đĩa cùng phụ kiện cân chỉnh & vệ sinh từ Hue Sound Vinyl Care." },
    @{ slug = "dia-than-nhac-ha-noi-sai-gon";  utm = "dia_than_nhac_ha_noi_sai_gon";  title = "Đĩa Than Ký Ức Hà Nội & Sài Gòn Xưa"; desc = "Hành trình âm thanh qua những bản thu gắn liền với ký ức Hà Nội và Sài Gòn xưa: Saigon Super Sound 1965–75, Saigon Soul Revival, Hà Nội Nhớ, Hà Nội Riêng Tôi, HANOI MASTERS." },
    @{ slug = "qua-tang-dia-than-cao-cap";     utm = "qua_tang_dia_than_cao_cap";     title = "Quà Tặng Đĩa Than & Analog Audio Cao Cấp"; desc = "Gợi ý quà tặng đĩa than, mâm than & phụ kiện Audiophile sang trọng cho sếp, đối tác, người thân — kèm gói quà nghệ thuật, viết thiệp tay & xuất báo giá PDF nhanh." }
)

$guideMeta = @(
    @{ slug = "bao-quan-dia-than-thoi-tiet-viet-nam";      utm = "guide_bao_quan_dia_than";   title = "Cách bảo quản đĩa than với thời tiết nóng ẩm ở Việt Nam" },
    @{ slug = "dia-than-moi-co-can-rua-truoc-khi-nghe";    utm = "guide_rua_dia_than_moi";    title = "Đĩa than mới mua (Mới 100% / New Seal) có cần rửa trước khi nghe không?" },
    @{ slug = "dong-goi-dia-than-chong-rach-mep-di-tinh";  utm = "guide_dong_goi_3_lop";      title = "Đóng gói đĩa than gửi đi tỉnh có sợ cong vênh hay rách mép (Seam Splits) không?" },
    @{ slug = "cach-chon-mam-than-dau-tien";               utm = "guide_chon_mam_than";       title = "Cách chọn mâm đĩa than (Turntable) đầu tiên cho người mới bắt đầu" },
    @{ slug = "phan-biet-kim-mm-va-mc";                    utm = "guide_kim_mm_mc";           title = "Phân biệt kim đĩa than MM và MC – Nên chọn loại nào cho dàn Analog Audio?" },
    @{ slug = "giai-thich-thang-goldmine-grading";         utm = "guide_goldmine_grading";    title = "Giải thích thang đánh giá tình trạng đĩa Goldmine Grading Standard (M, NM, VG+, VG...)" }
)

# 1. Xuat data/catalog_all.json
if (-not (Test-Path "data")) { New-Item -ItemType Directory -Path "data" -Force | Out-Null }
[System.IO.File]::WriteAllText((Join-Path (Get-Location) "data\catalog_all.json"), ($allItems | ConvertTo-Json -Depth 6), $utf8NoBom)

# 2. Xuat feed/products.json (OpenAI / Google Shopping Feed 0d + UTM khong dau)
if (-not (Test-Path "feed")) { New-Item -ItemType Directory -Path "feed" -Force | Out-Null }
$feedProducts = New-Object System.Collections.ArrayList
$priceValidUntil = (Get-Date).AddDays(90).ToString("yyyy-MM-dd")
$accNoteVi = "* Giảm giá phụ kiện khi mua cùng đĩa than / mâm than — inbox để được tư vấn"

foreach ($it in $allItems) {
    $seoPrefix = switch ($it.category) {
        "vinyl"     { "Đĩa than (Vinyl) " }
        "cd"        { "Đĩa CD / SACD " }
        "turntable" { "Máy hát đĩa than (Turntable) " }
        default     { "Analog Audio – " }
    }
    $pageUrl = "$SiteBaseUrl/p/$($it.slug)/"
    $utmPageUrl = "$pageUrl?utm_source=ai_agent&utm_medium=shopping_feed&utm_campaign=b2a_catalog"
    $condVi = Get-CondDisplayVi $it.cond
    $priceFmt = if ($it.p -gt 0) { "$($it.p.ToString('N0')) đ" } else { "Liên hệ" }
    $chatRaw = "Chào Hue Sound Express, mình quan tâm: $seoPrefix$($it.a) - $($it.t) (SKU: $($it.sku) | $condVi | Giá: $priceFmt). Link: $pageUrl"
    $waUrl = "https://wa.me/84932313918?text=$([Uri]::EscapeDataString($chatRaw))"

    [void]$feedProducts.Add([PSCustomObject]@{
        id                    = $it.id
        sku                   = $it.sku
        title                 = $it.t
        seo_title             = "$seoPrefix$($it.t) – $($it.a)"
        artist_or_brand       = $it.a
        category              = $it.category
        genre_code            = $it.g
        genre_vi              = (Get-GenreDisplayVi $it.g)
        mood_codes            = $it.mood
        keywords              = $it.kw
        label                 = $it.l
        format                = $it.f
        year                  = $it.y
        condition_code        = $it.cond
        condition_vi          = $condVi
        condition_schema      = if ($it.cond -eq "mint_sealed") { "https://schema.org/NewCondition" } else { "https://schema.org/UsedCondition" }
        availability          = $it.s
        order_mode            = if ($it.PSObject.Properties['force_contact'] -and $it.force_contact) { "contact" } elseif ($it.s -eq "out_of_stock") { "waitlist" } elseif ($it.contact_only) { "contact" } else { "cart" }
        list_price_vnd        = $it.p_orig
        sale_price_vnd        = $it.p
        discount_vnd          = [Math]::Max(0, $it.p_orig - $it.p)
        accessory_bundle_only = ($it.category -eq "accessory" -or $it.category -eq "care")
        url                   = $utmPageUrl
        canonical_url         = $pageUrl
        image_url             = if ($it.img -match '^https?://') { $it.img } else { "$SiteBaseUrl/og-image.jpg" }
        audio_preview_url     = if ($it.aud -match '^https?://') { $it.aud } else { "" }
        zalo_url              = "https://zalo.me/0932313918"
        whatsapp_order_url    = $waUrl
    })
}

$feedPayload = [PSCustomObject]@{
    schema_version      = "1.1"
    store               = "Hue Sound Express"
    care_service        = "Hue Sound Vinyl Care"
    slogan              = "Where sound finds its color"
    updated_at          = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss+07:00")
    currency            = "VND"
    vat_note            = "Giá báo chưa bao gồm thuế VAT 10%. Bảng báo giá không phải hóa đơn VAT."
    accessory_promo     = $accNoteVi
    contact             = [PSCustomObject]@{
        hotline_zalo = "09323.13918"
        zalo_url     = "https://zalo.me/0932313918"
        whatsapp_url = "https://wa.me/84932313918"
        website      = "$SiteBaseUrl/"
    }
    total_products      = $feedProducts.Count
    collections         = @($collectionMeta | ForEach-Object { [PSCustomObject]@{ slug = $_.slug; title = $_.title; url = "$SiteBaseUrl/bo-suu-tap/$($_.slug)/" } })
    guides              = @($guideMeta | ForEach-Object { [PSCustomObject]@{ slug = $_.slug; title = $_.title; url = "$SiteBaseUrl/cam-nang/$($_.slug)/" } })
    products            = $feedProducts
}
[System.IO.File]::WriteAllText((Join-Path (Get-Location) "feed\products.json"), ($feedPayload | ConvertTo-Json -Depth 6), $utf8NoBom)

# 3. Sinh llms.txt chuan B2A
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("# HUE SOUND EXPRESS - OFFICIAL AI AGENT & B2A KNOWLEDGE BASE")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 1. Brand Identity, Shopping Feed & Contact")
[void]$sb.AppendLine("- **Brand Name:** Hue Sound Express")
[void]$sb.AppendLine("- **Care & Technical Service:** Hue Sound Vinyl Care")
[void]$sb.AppendLine("- **Slogan:** Where Sound Finds Its Color")
[void]$sb.AppendLine("- **Machine-Readable Shopping Feed (1-Call JSON):** $SiteBaseUrl/feed/products.json (chứa đầy đủ tồn kho `in_stock`/`out_of_stock`, giá gốc/giá giảm, keywords ca sĩ/nhạc sĩ/nhạc cụ/bài hát/chủ đề và link đặt Zalo/WhatsApp)")
[void]$sb.AppendLine("- **Positioning:** Chuyên phân phối trực tuyến **Analog Audio** chính hãng: Đĩa than (Vinyl LP), Đĩa CD/SACD, Máy hát đĩa than (Turntable), Phụ kiện & Dịch vụ Hue Sound Vinyl Care tại Hà Nội & giao hàng toàn quốc.")
[void]$sb.AppendLine("- **Internal Codes (snake_case):** status (`in_stock`, `out_of_stock`), condition (`mint_sealed`, `near_mint`, `very_good_plus`, `very_good`, `good_plus`, `good`, `fair`, `poor`), genre (`love_songs_1954_1975`, `bolero_golden_oldies`, `vietnamese_pop_modern`, `vietnamese_folk_patriotic`, `instrumental`, `jazz_vocal`, `audiophile_reference`, `rock_metal`, `pop_disco`, `classic_soundtrack`).")
[void]$sb.AppendLine("- **Ordering & Consultation (Zalo / Hotline / WhatsApp):** 09323.13918 (https://zalo.me/0932313918 | https://wa.me/84932313918)")
[void]$sb.AppendLine("- **Website:** $SiteBaseUrl/")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 2. Operational, Pricing & Quoting Policies")
[void]$sb.AppendLine("- **Online-First Experience:** Giao hỏa tốc 2H nội thành Hà Nội & chuyển phát nhanh toàn quốc với quy chuẩn đóng gói 3 lớp chống rách mép (Seam Splits Prevention). Nếu khách có nhu cầu hẹn bàn giao trực tiếp, chuyển qua nhân sự tại Zalo 09323.13918.")
[void]$sb.AppendLine("- **Pricing & Promo Rules:** Giá báo chưa bao gồm VAT 10% (giá 0 hiển thị Liên hệ). $accNoteVi.")
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 3. Themed Collection Landing Pages (/bo-suu-tap/)")
foreach ($col in $collectionMeta) {
    [void]$sb.AppendLine("- **$($col.title):** $SiteBaseUrl/bo-suu-tap/$($col.slug)/")
}
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 4. Analog Audio Guides & FAQs (/cam-nang/)")
foreach ($g in $guideMeta) {
    [void]$sb.AppendLine("- **$($g.title):** $SiteBaseUrl/cam-nang/$($g.slug)/")
}
[void]$sb.AppendLine("")
[void]$sb.AppendLine("## 5. Unified Analog Audio Catalog ($($allItems.Count) Items)")
[void]$sb.AppendLine("| Category | SKU | Artist / Brand | Product / Album | Genre Code | Condition | Sale Price (VND) | Status | URL |")
[void]$sb.AppendLine("| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |")
foreach ($it in $allItems) {
    $priceTxt = if ($it.p -gt 0) { $it.p.ToString("N0") } else { "Liên hệ" }
    $urlTxt = "$SiteBaseUrl/p/$($it.slug)/"
    $llmsPrefix = switch ($it.category) { "vinyl" { "Đĩa than (Vinyl) " }; "cd" { "Đĩa CD " }; "turntable" { "Máy hát đĩa than " }; default { "" } }
    [void]$sb.AppendLine("| $($it.category) | $($it.sku) | $($it.a) | $llmsPrefix$($it.t) | $($it.g) | $($it.cond) | $priceTxt | $($it.s) | $urlTxt |")
}
[System.IO.File]::WriteAllText((Join-Path (Get-Location) "llms.txt"), $sb.ToString(), $utf8NoBom)

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
                @{ q = "Chính sách giảm giá phụ kiện tại Hue Sound Express như thế nào?"; a = "Giảm giá phụ kiện khi mua cùng đĩa than / mâm than — inbox Zalo 09323.13918 để được tư vấn phối ghép và ưu đãi combo tốt nhất."; guide = "../../cam-nang/bao-quan-dia-than-thoi-tiet-viet-nam/" },
                @{ q = "Tại sao nên dùng chổi sợi Carbon và túi lót chống tĩnh điện cho đĩa than?"; a = "Chổi sợi Carbon giúp khử tĩnh điện và quét sạch bụi mịn trước khi hạ kim, còn túi lót chống tĩnh điện thay cho túi giấy thô giúp đĩa không bị xước dăm và giảm tiếng nổ lách tách."; guide = "../../cam-nang/dia-than-moi-co-can-rua-truoc-khi-nghe/" }
            )
        }
    }
}

# 4. Sinh trang chi tiet /p/ + Redirect URL cu + 6 trang /cam-nang/ + 10 trang /bo-suu-tap/ + sitemap.xml
if ($GenerateProductPages) {
    $pRoot = Join-Path (Get-Location) "p"
    if (-not (Test-Path $pRoot)) { New-Item -ItemType Directory -Path $pRoot -Force | Out-Null }

    $sm = New-Object System.Text.StringBuilder
    [void]$sm.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
    [void]$sm.AppendLine('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">')
    [void]$sm.AppendLine("  <url><loc>$SiteBaseUrl/</loc><priority>1.0</priority></url>")

    $accLookup = @($allItems | Where-Object { $_.category -eq "accessory" -or $_.category -eq "care" })
    $vinylLookup = @($allItems | Where-Object { $_.category -eq "vinyl" -and $_.s -eq "in_stock" })
    $ttLookup = @($allItems | Where-Object { $_.category -eq "turntable" })
    $redirectCount = 0

    foreach ($it in $allItems) {
        $dir = Join-Path $pRoot $it.slug
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        $seoPrefix = switch ($it.category) { "vinyl" { "Đĩa than (Vinyl) " }; "cd" { "Đĩa CD / SACD " }; "turntable" { "Máy hát đĩa than (Turntable) " }; default { "Analog Audio – " } }
        $catShort = switch ($it.category) { "vinyl" { "Đĩa Than (Vinyl LP)" }; "cd" { "Đĩa CD / SACD" }; "turntable" { "Máy Hát / Mâm Than" }; default { "Phụ kiện và dịch vụ" } }
        $catLabel = "$catShort • Analog Audio"
        $genreDisplay = Get-GenreDisplayVi $it.g
        $catParam = if ($it.category -eq "care") { "accessory" } else { $it.category }
        $hasOrigin = ($it.category -eq "vinyl" -or $it.category -eq "cd")
        $originCode = if ($hasOrigin) { Get-ItemOriginCode $it } else { "" }
        $originLabelVi = if ($originCode -eq "VN") { "Nhạc Việt" } elseif ($originCode -eq "INTL") { "Nhạc Quốc Tế" } else { "" }
        $catUrl    = "../../?cat=$catParam"
        $originUrl = if ($hasOrigin) { "../../?cat=$catParam&origin=$originCode" } else { "" }
        $genreUrl  = if ($hasOrigin) { "../../?cat=$catParam&origin=$originCode&genre=$([Uri]::EscapeDataString($it.g))" } else { "../../?cat=$catParam&genre=$([Uri]::EscapeDataString($it.g))" }
        $artistUrl = "../../?cat=$catParam&artist=$([Uri]::EscapeDataString($it.a))"
        $hasLabelLink = ($it.l -and $it.l -notmatch '^(?i)(other|việt nam)$')
        $labelUrl  = if ($hasLabelLink) { "../../?cat=$catParam&label=$([Uri]::EscapeDataString($it.l))" } else { "" }
        $condUrl   = "../../?cat=$catParam&cond=$([Uri]::EscapeDataString($it.cond))"
        $statusUrl = "../../?cat=$catParam&status=$([Uri]::EscapeDataString($it.s))"

        $extraSeoDesc = if ($it.d -match 'Hue Sound Express') { "" } elseif ($it.g -eq "bolero_golden_oldies") { " Thể loại Nhạc Bolero (Nhạc Vàng / Nhạc Trữ Tình) chuẩn Analog Audio." } elseif ($it.g -eq "love_songs_1954_1975") { " Dòng nhạc Tình khúc 1954–1975 kinh điển chuẩn Analog Audio." } elseif ($it.category -eq "vinyl") { " Đĩa than Vinyl chuẩn Analog Audio tại Hue Sound Express." } else { " Sản phẩm Analog Audio chính hãng tại Hue Sound Express." }

        $seoTitleEsc = Escape-Html "$seoPrefix$($it.t) – $($it.a)"
        $fullDescRaw = "$($it.d)$extraSeoDesc"
        $descEsc = Escape-Html $fullDescRaw
        $descLinkedHtml = Format-LinkedDescription -DescText $fullDescRaw -Item $it -CatParam $catParam -OriginCode $originCode -GenreDisplay $genreDisplay -CatUrl $catUrl -OriginUrl $originUrl -GenreUrl $genreUrl -ArtistUrl $artistUrl -LabelUrl $labelUrl
        $kwJoined = ($it.kw -join ", ")
        $kwEsc = Escape-Html $kwJoined
        $priceFmt = if ($it.p -gt 0) { "$($it.p.ToString('N0')) đ" } else { "Liên hệ báo giá" }
        $origPriceFmt = if ($it.p_orig -gt $it.p) { "$($it.p_orig.ToString('N0')) đ" } else { "" }
        $saveK = if ($it.p_orig -gt $it.p) { [int](($it.p_orig - $it.p) / 1000) } else { 0 }
        $condVi = Get-CondDisplayVi $it.cond
        $statusVi = Get-StatusDisplayVi $it.s
        $isAccBundle = ($it.category -eq "accessory" -or $it.category -eq "care")
        $accNoteHtml = if ($isAccBundle) { "<div class='acc-note'>$(Escape-Html $accNoteVi)</div>" } else { "" }
        $promoNoteHtml = if ($it.PSObject.Properties['promo_note'] -and $it.promo_note) { "<div class='acc-note' id='promoNoteEl'>🎁 $(Escape-Html $it.promo_note)</div>" } else { "" }
        $priceHtml = if ($it.p -gt 0 -and $origPriceFmt) {
            "<div class='price-stack'><span class='price-label'>Giá</span><span class='price-old'>$origPriceFmt</span><span class='price-save' id='priceSaveEl'>-$($saveK)k</span></div><div class='price' id='priceMainEl'>$priceFmt</div>$promoNoteHtml$accNoteHtml"
        } else {
            "<div class='price-stack'><span class='price-label'>Giá</span></div><div class='price' id='priceMainEl'>$priceFmt</div>$promoNoteHtml$accNoteHtml"
        }
        $promoExpireJs = ""
        if ($it.PSObject.Properties['promo_until'] -and $it.promo_until -and $it.PSObject.Properties['p_regular']) {
            $regSaveK = [int](($it.p_orig - $it.p_regular) / 1000)
            $regFmt = "$($it.p_regular.ToString('N0')) đ"
            $promoExpireJs = "<script>(function(){if(Date.now()>new Date('$($it.promo_until)T23:59:59+07:00').getTime()){var pm=document.getElementById('priceMainEl');var ps=document.getElementById('priceSaveEl');var pn=document.getElementById('promoNoteEl');if(pm)pm.textContent='$regFmt';if(ps)ps.textContent='-$($regSaveK)k';if(pn)pn.style.display='none';}})();</script>"
        }
        $genreBadgesHtml = ""
        $gList = if ($it.PSObject.Properties['genres'] -and $it.genres -and $it.genres.Count -gt 0) { $it.genres } else { @($it.g) }
        $genreBadgesHtml = ($gList | ForEach-Object {
            $gCode = [string]$_
            $gDisp = Get-GenreDisplayVi $gCode
            $gHref = if ($hasOrigin) { "../../?cat=$catParam&origin=$originCode&genre=$([Uri]::EscapeDataString($gCode))" } else { "../../?cat=$catParam&genre=$([Uri]::EscapeDataString($gCode))" }
            "<a class='badge' href='$gHref' title='Xem thể loại $(Escape-Html $gDisp)'>$(Escape-Html $gDisp)</a>"
        }) -join "`n          "
        $pageUrl = "$SiteBaseUrl/p/$($it.slug)/"
        $shareUtmUrl = "$pageUrl?utm_source=share&utm_medium=referral&utm_campaign=product_share"
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

        $upsellItems = @()
        if ($it.category -eq "turntable") { $upsellItems = $accLookup | Where-Object { $_.sku -in @("ACC-AQ-CARBON","ACC-AT607A","ACC-SCALE-01","CARE-SETUP-01") } | Select-Object -First 3 }
        elseif ($it.category -eq "vinyl" -or $it.category -eq "cd") { $upsellItems = $accLookup | Where-Object { $_.sku -in @("ACC-SLEEVE-50","ACC-AQ-CARBON","CARE-ULTRASONIC-01") } | Select-Object -First 3 }
        else { $upsellItems = $vinylLookup | Select-Object -First 3 }
        $upsellHtml = ""
        if ($upsellItems.Count -gt 0) {
            $cards = ($upsellItems | ForEach-Object {
                "<a class='up-card' href='../$($_.slug)/'><strong>$(Escape-Html $_.t)</strong><span>$($_.p.ToString('N0')) đ</span></a>"
            }) -join ""
            $upHead = if ($isAccBundle) { "Đĩa than mới về đề xuất phối ghép" } else { "Gợi ý đi kèm từ Hue Sound Vinyl Care" }
            $upsellHtml = "<section class='box'><h2>$upHead</h2><div class='up-grid'>$cards</div></section>"
        }

        $catFaqs = Get-CategoryFaqs -Category $it.category
        $faqItemsHtml = ($catFaqs | ForEach-Object {
            "<details class='faq-item'><summary>$(Escape-Html $_.q)</summary><p>$(Escape-Html $_.a) <a class='faq-more' href='$($_.guide)'>Xem bài viết chi tiết →</a></p></details>"
        }) -join ""
        $faqBoxHtml = "<section class='box'><h2>Câu hỏi thường gặp (FAQ)</h2>$faqItemsHtml</section>"
        $faqJsonArr = ($catFaqs | ForEach-Object {
            "{ `"@type`": `"Question`", `"name`": `"$(Escape-JsonStr $_.q)`", `"acceptedAnswer`": { `"@type`": `"Answer`", `"text`": `"$(Escape-JsonStr $_.a)`" } }"
        }) -join ",`n      "

        $chatRaw = "Chào Hue Sound Express, mình quan tâm sản phẩm: $seoPrefix$($it.a) - $($it.t) (Mã: $($it.sku) | Tình trạng đĩa: $condVi | Giá: $priceFmt). Link: $pageUrl"
        $waHref = "https://wa.me/84932313918?text=$([Uri]::EscapeDataString($chatRaw))"
        $zaloHref = "https://zalo.me/0932313918"
        $chatHtmlVal = Escape-Html $chatRaw

        $schemaAvail = if ($it.s -eq "in_stock") { "https://schema.org/InStock" } else { "https://schema.org/OutOfStock" }
        $schemaCond = if ($it.cond -eq "mint_sealed") { "https://schema.org/NewCondition" } else { "https://schema.org/UsedCondition" }
        $weightJsonProp = ""
        if ($it.PSObject.Properties['w'] -and $it.w -match '(\d+)') {
            $weightJsonProp = "`n        `"weight`": { `"@type`": `"QuantitativeValue`", `"value`": $($matches[1]), `"unitCode`": `"GRM`" },"
        }
        $primaryCtaHtml = if ($it.PSObject.Properties['force_contact'] -and $it.force_contact) {
            "<button class='btn btn-copper' type='button' onclick='openZaloPopup()'>💬 Liên hệ đặt hàng</button>"
        } elseif ($it.s -eq "out_of_stock") {
            "<a class='btn btn-copper' href='../../?waitlist=$([Uri]::EscapeDataString($it.sku))'>🔔 Báo khi về</a>"
        } elseif ($it.contact_only) {
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
  <meta name="keywords" content="$kwEsc">
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
        "keywords": "$(Escape-JsonStr $kwJoined)",
        "image": "$ogImg",
        "description": "$(Escape-JsonStr "$($it.d)$extraSeoDesc")",
        "brand": { "@type": "Brand", "name": "$(Escape-JsonStr $(if ($it.l) { $it.l } else { 'Hue Sound Express' }))" },$weightJsonProp
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
    :root { --copper:#B5541F; --copper-dark:#7E3210; --copper-grad:linear-gradient(135deg, #DC8644 0%, #B85922 46%, #8C3A12 100%); --ink:#4A2312; --muted:#6F5139; --card:rgba(250,242,222,.82); --border:rgba(142,66,24,.27); --sans:"Montserrat",system-ui,-apple-system,sans-serif; --serif:"Lora",Georgia,serif; }
    * { box-sizing:border-box; margin:0; padding:0; }
    body { font-family:var(--sans); color:var(--ink); line-height:1.6; background-color:#C9B184; background-image:radial-gradient(ellipse at 50% 18%, #F4E9CC 0%, #E7D5AE 44%, #C5AC7E 82%, #AA9062 100%); background-attachment:fixed; padding:24px 16px 64px; position:relative; }
    .wrap { max-width:900px; margin:0 auto; position:relative; z-index:2; }
    .topbar { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:16px; }
    .back { display:inline-flex; align-items:center; gap:6px; color:var(--copper-dark); text-decoration:none; font-size:13px; font-weight:600; background:var(--card); border:1px solid var(--border); padding:8px 14px; border-radius:99px; }
    .brand-mini { font-weight:800; font-size:13px; color:var(--copper); letter-spacing:.04em; }
    .crumb { font-size:12px; color:var(--muted); margin-bottom:14px; }
    .crumb a { color:var(--copper-dark); text-decoration:none; font-weight:600; }
    .hero { display:grid; grid-template-columns:290px 1fr; gap:26px; background:var(--card); border:1px solid var(--border); border-radius:18px; padding:24px; backdrop-filter:blur(8px); box-shadow:0 12px 30px -12px rgba(110,60,20,.22); }
    @media (max-width:680px){ .hero{ grid-template-columns:1fr; } }
    .cover { aspect-ratio:1/1; border-radius:14px; overflow:hidden; background:radial-gradient(circle at center, #F3E9D0 0%, #DFC99B 100%); border:1px solid var(--border); display:flex; align-items:center; justify-content:center; text-align:center; padding:16px; }
    .cover img { width:100%; height:100%; object-fit:cover; }
    .badge-row { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:10px; }
    .badge { display:inline-block; font-size:11px; font-weight:600; padding:4px 11px; border-radius:99px; background:rgba(181,84,31,.11); color:var(--copper-dark); border:1px solid rgba(181,84,31,.22); text-decoration:none; transition:all .18s; }
    a.badge:hover { background:rgba(181,84,31,.22); border-color:var(--copper); color:var(--ink); }
    h1 { font-family:var(--serif); font-size:25px; line-height:1.3; margin-bottom:4px; color:var(--ink); }
    .sub { color:var(--copper-dark); font-size:16px; font-weight:600; margin-bottom:10px; }
    .xlink { color:var(--copper-dark); text-decoration:none; border-bottom:1px dashed rgba(181,84,31,.45); transition:all .18s; }
    .xlink:hover { color:var(--copper); border-bottom-style:solid; }
    .meta { font-size:12.5px; color:var(--muted); margin-bottom:12px; }
    .desc { font-size:14px; color:var(--ink); opacity:.92; }
    .desc-link { color:var(--copper-dark); font-weight:600; text-decoration:none; border-bottom:1px dotted rgba(181,84,31,.55); transition:all .18s; }
    .desc-link:hover { color:var(--copper); border-bottom-style:solid; }
    .price-row { display:flex; align-items:center; flex-wrap:wrap; gap:12px; margin:16px 0; }
    .price-label { font-size:11px; text-transform:uppercase; letter-spacing:.06em; color:var(--muted); font-weight:600; }
    .price-stack { display:flex; align-items:center; gap:7px; }
    .price-old { font-size:13.5px; color:var(--muted); text-decoration:line-through; opacity:.82; font-weight:500; }
    .price-save { font-size:10.5px; font-weight:700; color:#FFF7E8; background:#B5541F; padding:1px 6px; border-radius:99px; }
    .price { font-size:25px; font-weight:800; color:var(--copper); line-height:1.15; }
    .acc-note { font-size:11.5px; color:#8C3A12; font-weight:600; margin-top:4px; }
    .status { font-size:12px; font-weight:600; padding:4px 11px; border-radius:99px; background:rgba(201,154,46,.18); border:1px solid rgba(201,154,46,.5); color:var(--ink); text-decoration:none; transition:all .18s; }
    a.status:hover { background:rgba(201,154,46,.32); border-color:var(--copper); }
    .cta-row { display:flex; flex-wrap:wrap; gap:9px; margin-top:14px; }
    .btn { text-decoration:none; padding:10px 15px; border-radius:10px; font-weight:600; font-size:13px; display:inline-flex; align-items:center; gap:6px; transition:all .2s; border:none; cursor:pointer; font-family:inherit; }
    .btn-copper { background:var(--copper-grad); color:#FFF7E8; }
    .btn-wa { background:#25D366; color:#fff; }
    .btn-zalo { background:#0068FF; color:#fff; }
    .btn-share { background:rgba(255,251,240,.85); color:var(--ink); border:1px solid rgba(181,84,31,.35); }
    .audio-box { margin-top:16px; padding:12px 14px; background:rgba(243,233,208,.8); border-radius:10px; border:1px solid var(--border); }
    .small { font-size:12px; color:var(--muted); }
    .box { margin-top:18px; background:var(--card); border:1px solid var(--border); border-radius:16px; padding:20px; }
    .box h2 { font-family:var(--serif); font-size:17px; color:var(--copper-dark); margin-bottom:12px; }
    .tracks { padding-left:20px; color:var(--ink); font-size:13.5px; columns:2; gap:26px; }
    .tracks li { margin-bottom:5px; }
    @media (max-width:600px){ .tracks{ columns:1; } }
    .up-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:12px; }
    .up-card { text-decoration:none; background:rgba(255,251,240,.75); border:1px solid var(--border); padding:14px; border-radius:12px; color:var(--ink); display:flex; flex-direction:column; gap:6px; }
    .up-card span { color:var(--copper); font-weight:700; font-size:13.5px; }
    .faq-item { border-bottom:1px dashed rgba(142,66,24,.22); padding:9px 0; }
    .faq-item:last-child { border-bottom:none; }
    .faq-item summary { font-weight:700; font-size:13.5px; cursor:pointer; }
    .faq-item p { font-size:13px; color:var(--muted); margin-top:6px; }
    .faq-more { color:var(--copper); font-weight:600; text-decoration:none; }
    .kw-cloud { display:flex; flex-wrap:wrap; gap:6px; margin-top:10px; }
    .kw-chip { font-size:11px; padding:3px 9px; border-radius:99px; background:rgba(181,84,31,.09); color:var(--muted); border:1px solid rgba(181,84,31,.18); text-decoration:none; transition:all .18s; }
    a.kw-chip:hover { background:rgba(181,84,31,.2); color:var(--copper-dark); border-color:var(--copper); }
    footer { text-align:center; margin-top:34px; font-size:12px; color:var(--muted); }
    .chat-backdrop { position:fixed; inset:0; background:rgba(58,28,10,.5); backdrop-filter:blur(3px); z-index:90; display:none; align-items:center; justify-content:center; padding:16px; }
    .chat-backdrop.open { display:flex; }
    .chat-popup { background:#FAF2DE; border:1px solid rgba(142,66,24,.45); border-radius:16px; max-width:460px; width:100%; padding:18px; }
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
      <a href="../../">Trang chủ</a> › <a href="$catUrl">$catShort</a> $(if ($hasOrigin) { "› <a href='$originUrl'>$originLabelVi</a>" }) › <a href="$genreUrl">$(Escape-Html $genreDisplay)</a> › <span>$(Escape-Html $it.t)</span>
    </nav>
    <div class="hero">
      <div class="cover">
        $(if ($it.img -match '^https?://') { "<img src='$(Escape-Html $it.img)' alt='$seoTitleEsc' onerror=`"this.style.display='none';this.nextElementSibling.style.display='flex';`">$(Get-VinylFallbackSvg -Title $it.t -Artist $it.a -Fmt $it.f -Uid "p$($it.id)" -Hidden $true)" } else { Get-VinylFallbackSvg -Title $it.t -Artist $it.a -Fmt $it.f -Uid "p$($it.id)" -Hidden $false })
      </div>
      <div>
        <div class="badge-row">
          <a class="badge" href="$catUrl" title="Xem danh mục $catShort">$catShort</a>
          $(if ($hasOrigin) { "<a class='badge' href='$originUrl' title='Xem phân loại $originLabelVi'>$originLabelVi</a>" })
          $genreBadgesHtml
        </div>
        <h1>$(Escape-Html $it.t)</h1>
        <div class="sub"><a class="xlink" href="$artistUrl" title="Xem tất cả sản phẩm của $(Escape-Html $it.a)">$(Escape-Html $it.a)</a></div>
        <div class="meta">Mã SKU: <strong>$(Escape-Html $it.sku)</strong> • Định dạng: $(Escape-Html $it.f) $(if ($it.PSObject.Properties['w'] -and $it.w) { "• Trọng lượng: <strong>$(Escape-Html $it.w)</strong>" }) $(if ($hasLabelLink) { "• Hãng: <a class='xlink' href='$labelUrl' title='Xem tất cả ấn phẩm của hãng $(Escape-Html $it.l)'><strong>$(Escape-Html $it.l)</strong></a>" } elseif ($it.l) { "• $(Escape-Html $it.l)" }) $(if ($it.y) { "• Năm $($it.y)" })</div>
        <p class="desc">$descLinkedHtml</p>
        <div class="price-row">
          <div>$priceHtml</div>
          <a class="status" href="$condUrl" title="Lọc theo tình trạng $(Escape-Html $condVi)">$(Escape-Html $condVi)</a>
          <a class="status" href="$statusUrl" title="Lọc theo trạng thái $(Escape-Html $statusVi)">$(Escape-Html $statusVi)</a>
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
    <section class="box">
      <h2>Từ khóa &amp; Chủ đề liên quan</h2>
      <div class="kw-cloud">$(($it.kw | Select-Object -First 14 | ForEach-Object { "<a class='kw-chip' href='../../?cat=$catParam&q=$([Uri]::EscapeDataString($_))'>$(Escape-Html $_)</a>" }) -join "")</div>
    </section>
    <footer>© 2026 Hue Sound Express • Analog Audio • Đóng gói 3 lớp chống rách mép • Hotline/Zalo &amp; WhatsApp: 09323.13918</footer>
  </div>

  <div class="share-toast" id="shareToast">✓ Đã sao chép link sản phẩm để chia sẻ!</div>

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
    function openZaloPopup() { document.getElementById('zaloPopup').classList.add('open'); copyChatMsg(); }
    function shareProduct() {
      var url = '$shareUtmUrl';
      if (navigator.share) { navigator.share({ title: document.title, url: url }).catch(function(){}); }
      else {
        if (navigator.clipboard) navigator.clipboard.writeText(url).catch(function(){});
        var t = document.getElementById('shareToast'); t.style.display = 'block';
        setTimeout(function(){ t.style.display = 'none'; }, 2400);
      }
    }
  </script>
  $promoExpireJs
</body>
</html>
"@
        [System.IO.File]::WriteAllText((Join-Path $dir "index.html"), $html, $utf8NoBom)
        [void]$sm.AppendLine("  <url><loc>$pageUrl</loc><priority>0.8</priority></url>")

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

    # 5. SINH 10 TRANG BO SUU TAP TINH (/bo-suu-tap/<slug>/index.html) + TU DONG PHAN PHOI & BAO CAO
    $colRoot = Join-Path (Get-Location) "bo-suu-tap"
    if (-not (Test-Path $colRoot)) { New-Item -ItemType Directory -Path $colRoot -Force | Out-Null }

    # Quy tac loai tru chung (bo Xuan Hao, To Nga khoi tat ca bo suu tap) & rieng cho Tet / Ngay Le
    $globalColExcludeRx = '(?i)(xuân hảo|xuan hao|tố nga|to nga)'
    $festColExcludeRx   = '(?i)(căn nhà ngoại ô|can nha ngoai o|ngày còn em bên tôi|ngay con em ben toi|tuyển tập nhạc xưa|tuyen tap nhac xua|\boplus\b)'
    $colEligibleItems   = @($allItems | Where-Object { "$($_.a) $($_.t)" -notmatch $globalColExcludeRx })
    $vinylColFallback   = @($vinylLookup | Where-Object { "$($_.a) $($_.t)" -notmatch $globalColExcludeRx })

    $colSelectors = @{
        "dia-than-nhac-tet"              = { param($list) @($list | Where-Object { ($_.category -eq 'vinyl' -or $_.category -eq 'cd') -and ("$($_.a) $($_.t)" -notmatch $festColExcludeRx) -and ("$($_.t) $(($_.tracklist) -join ' ')" -match '(?i)(mùa xuân|tết|ly rượu mừng|chúc xuân|đón xuân|em ơi mùa xuân|xuân về|xuân muộn|ngày xuân)') }) }
        "dia-than-ngay-le"               = { param($list) @($list | Where-Object { ($_.category -eq 'vinyl' -or $_.category -eq 'cd') -and ("$($_.a) $($_.t)" -notmatch $festColExcludeRx) -and ("$($_.t) $(($_.tracklist) -join ' ')" -match '(?i)(christmas|noel|giáng sinh|mùa sao sáng|neujahr|new year|holiday|winter|silent night)') }) }
        "dia-than-nhac-trinh-cong-son"   = { param($list) @($list | Where-Object { ($_.category -eq 'vinyl' -or $_.category -eq 'cd') -and ("$($_.a) $($_.t) $($_.d) $(($_.tracklist) -join ' ')" -match '(?i)(trịnh công sơn|nhạc trịnh|diễm xưa|cát bụi|hạ trắng|biển nhớ|mưa hồng|tuổi đá buồn|một cõi đi về|ru em từng ngón xuân nồng|vết lăn trầm|bống là ai|xin mặt trời ngủ yên|phôi pha|nắng thủy tinh|đóa hoa vô thường|còn tuổi nào cho em|hoa vàng mấy độ|tiến thoái lưỡng nan|ru tình|chiều một mình qua phố|ngủ đi con|tình nhớ|ướt mi|gọi tên bốn mùa|quỳnh hương|sóng về đâu|lặng lẽ nơi này|ở trọ|em còn nhớ hay em đã quên|để gió cuốn đi|nhớ mùa thu hà nội|tôi ơi đừng tuyệt vọng|như cánh vạc bay|cuối cùng cho một tình yêu)') }) }
        "dia-than-nhac-vang-bolero"      = { param($list) @($list | Where-Object { $_.g -eq 'bolero_golden_oldies' -or ($_.genres -contains 'bolero_golden_oldies') }) }
        "dia-than-tinh-khuc-1954-1975"   = { param($list) @($list | Where-Object { $_.g -eq 'love_songs_1954_1975' -or ($_.genres -contains 'love_songs_1954_1975') }) }
        "dia-than-nhac-viet-tre"         = { param($list) @($list | Where-Object { $_.category -eq 'vinyl' -and ($_.g -eq 'vietnamese_pop_modern' -or ($_.genres -contains 'vietnamese_pop_modern')) }) }
        "dia-than-instrument-hoa-tau"    = { param($list) @($list | Where-Object { $_.g -eq 'instrumental' -or ($_.genres -contains 'instrumental') }) }
        "dia-than-jazz-audiophile"       = { param($list) @($list | Where-Object { $_.category -eq 'vinyl' -and ($_.g -eq 'jazz_vocal' -or $_.g -eq 'audiophile_reference') -and $_.s -eq 'in_stock' } | Select-Object -First 36) }
        "may-hat-dia-than-cho-nguoi-moi" = { param($list) @($list | Where-Object { $_.category -eq 'turntable' -or $_.category -eq 'accessory' -or $_.category -eq 'care' }) }
        "dia-than-nhac-ha-noi-sai-gon"   = { param($list) @($list | Where-Object { ($_.category -eq 'vinyl' -or $_.category -eq 'cd') -and ("$($_.a) $($_.t) $(($_.tracklist) -join ' ')" -match '(?i)(hà nội|hanoi|sài gòn|saigon)') }) }
        "qua-tang-dia-than-cao-cap"      = { param($list) @($list | Where-Object { $_.s -eq 'in_stock' -and (($_.mood -contains 'refined_gifting') -or $_.category -eq 'turntable') } | Select-Object -First 30) }
    }

    # Theo doi phan phoi san pham vao bo suu tap de bao cao khi co san pham moi
    $snapPath = Join-Path (Get-Location) "data\collection_snapshot.json"
    $prevSkus = New-Object 'System.Collections.Generic.HashSet[string]'
    $prevColCounts = @{}
    if (Test-Path $snapPath) {
        try {
            $prevSnap = Get-Content $snapPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($s in $prevSnap.skus) { [void]$prevSkus.Add([string]$s) }
            if ($prevSnap.collection_counts) {
                $prevSnap.collection_counts.PSObject.Properties | ForEach-Object { $prevColCounts[$_.Name] = [int]$_.Value }
            }
        } catch {}
    }
    $skuToCols = @{}
    $newColCounts = @{}

    foreach ($col in $collectionMeta) {
        $cDir = Join-Path $colRoot $col.slug
        if (-not (Test-Path $cDir)) { New-Item -ItemType Directory -Path $cDir -Force | Out-Null }
        $cUrl = "$SiteBaseUrl/bo-suu-tap/$($col.slug)/"
        $cTitleEsc = Escape-Html $col.title
        $cDescEsc  = Escape-Html $col.desc

        $matched = @(& $colSelectors[$col.slug] $colEligibleItems)
        if ($matched.Count -lt 6) {
            # Bo sung them cac dia than con hang tieu bieu (da loai tru Xuan Hao, To Nga & muc loai tru le hoi)
            $extra = @($vinylColFallback | Where-Object { ($matched.sku -notcontains $_.sku) -and ("$($_.a) $($_.t)" -notmatch $festColExcludeRx) } | Select-Object -First (8 - $matched.Count))
            $matched = @($matched) + $extra
        }
        $shown = @($matched | Select-Object -First 36)
        $newColCounts[$col.slug] = $shown.Count
        foreach ($mItem in $shown) {
            if (-not $skuToCols.ContainsKey($mItem.sku)) { $skuToCols[$mItem.sku] = New-Object System.Collections.ArrayList }
            [void]$skuToCols[$mItem.sku].Add($col.slug)
        }

        $itemCardsHtml = ($shown | ForEach-Object {
            $fbUid = "c$($col.utm.Substring(0,[Math]::Min(6,$col.utm.Length)))$($_.id)"
            $imgTag = if ($_.img -match '^https?://') { "<img src='$(Escape-Html $_.img)' alt='$(Escape-Html "$($_.a) - $($_.t)")' loading='lazy' onerror=`"this.style.display='none';this.nextElementSibling.style.display='flex';`">$(Get-VinylFallbackSvg -Title $_.t -Artist $_.a -Fmt $_.f -Uid $fbUid -Hidden $true)" } else { Get-VinylFallbackSvg -Title $_.t -Artist $_.a -Fmt $_.f -Uid $fbUid -Hidden $false }
            $pTxt = if ($_.p -gt 0) { "$($_.p.ToString('N0')) đ" } else { "Liên hệ" }
            $oldTxt = if ($_.p_orig -gt $_.p -and $_.p -gt 0) { "<span class='old'>$($_.p_orig.ToString('N0')) đ</span>" } else { "" }
            $prodLink = "../../p/$($_.slug)/"
            $wTag = if ($_.PSObject.Properties['w'] -and $_.w) { " • $(Escape-Html $_.w)" } else { "" }
            "<a class='prod-card' href='$prodLink'><div class='thumb'>$imgTag</div><div class='p-body'><div class='p-title'>$(Escape-Html $_.t)</div><div class='p-art'>$(Escape-Html $_.a)</div><div class='p-meta'>$(Escape-Html (Get-CondDisplayVi $_.cond)) • $(Escape-Html (Get-StatusDisplayVi $_.s))$wTag</div><div class='p-price'>$oldTxt <strong>$pTxt</strong></div></div></a>"
        }) -join ""

        $listJson = ($shown | Select-Object -First 15 | ForEach-Object -Begin { $pos = 1 } -Process {
            $u = "$SiteBaseUrl/p/$($_.slug)/"
            "{ `"@type`": `"ListItem`", `"position`": $pos, `"url`": `"$u`", `"name`": `"$(Escape-JsonStr "$($_.a) - $($_.t)")`" }"
            $pos++
        }) -join ",`n      "

        $otherColsHtml = ($collectionMeta | Where-Object { $_.slug -ne $col.slug } | ForEach-Object {
            "<a class='col-chip' href='../$($_.slug)/'>$(Escape-Html $_.title)</a>"
        }) -join ""

        $cHtml = @"
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>$cTitleEsc | Hue Sound Express</title>
  <meta name="description" content="$cDescEsc">
  <meta name="theme-color" content="#E8D9B6">
  <link rel="canonical" href="$cUrl">
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
          { "@type": "ListItem", "position": 2, "name": "Bộ sưu tập Analog Audio", "item": "$SiteBaseUrl/" },
          { "@type": "ListItem", "position": 3, "name": "$(Escape-JsonStr $col.title)", "item": "$cUrl" }
        ]
      },
      {
        "@type": "CollectionPage",
        "name": "$(Escape-JsonStr $col.title)",
        "description": "$(Escape-JsonStr $col.desc)",
        "url": "$cUrl",
        "mainEntity": {
          "@type": "ItemList",
          "numberOfItems": $($shown.Count),
          "itemListElement": [
      $listJson
          ]
        }
      }
    ]
  }
  </script>
  <style>
    :root { --copper:#B5541F; --copper-dark:#7E3210; --ink:#4A2312; --muted:#6F5139; --card:rgba(250,242,222,.85); --border:rgba(142,66,24,.27); --sans:"Montserrat",system-ui,sans-serif; --serif:"Lora",Georgia,serif; }
    * { box-sizing:border-box; margin:0; padding:0; }
    body { font-family:var(--sans); color:var(--ink); line-height:1.6; background:radial-gradient(ellipse at 50% 18%, #F4E9CC 0%, #E7D5AE 44%, #C5AC7E 85%, #AA9062 100%); background-attachment:fixed; padding:24px 16px 64px; }
    .wrap { max-width:1040px; margin:0 auto; }
    .topbar { display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px; margin-bottom:16px; }
    .back { color:var(--copper-dark); text-decoration:none; font-size:13px; font-weight:600; background:var(--card); border:1px solid var(--border); padding:8px 14px; border-radius:99px; }
    .hero-col { background:var(--card); border:1px solid var(--border); border-radius:18px; padding:24px; margin-bottom:18px; }
    h1 { font-family:var(--serif); font-size:26px; color:var(--ink); margin-bottom:8px; }
    .hero-col p { font-size:14px; color:var(--muted); margin-bottom:14px; }
    .cta-bar { display:flex; flex-wrap:wrap; gap:9px; }
    .btn { text-decoration:none; padding:9px 15px; border-radius:10px; font-weight:600; font-size:13px; color:#FFF7E8; background:linear-gradient(135deg,#DC8644,#B85922,#8C3A12); }
    .btn-wa { background:#25D366; }
    .grid { display:grid; grid-template-columns:repeat(auto-fill,minmax(230px,1fr)); gap:14px; margin-bottom:22px; }
    .prod-card { text-decoration:none; background:var(--card); border:1px solid var(--border); border-radius:14px; overflow:hidden; color:var(--ink); display:flex; flex-direction:column; transition:all .2s; }
    .prod-card:hover { transform:translateY(-3px); border-color:var(--copper); }
    .thumb { aspect-ratio:1/1; background:#E7D5AE; display:flex; align-items:center; justify-content:center; overflow:hidden; }
    .thumb img { width:100%; height:100%; object-fit:cover; }
    .thumb .fb { font-weight:800; color:var(--copper); font-size:13px; }
    .p-body { padding:12px; display:flex; flex-direction:column; gap:4px; }
    .p-title { font-family:var(--serif); font-weight:600; font-size:14.5px; line-height:1.3; }
    .p-art { font-size:12.5px; font-weight:600; color:var(--copper-dark); }
    .p-meta { font-size:11px; color:var(--muted); }
    .p-price { margin-top:4px; font-size:14px; color:var(--copper); }
    .p-price .old { font-size:11.5px; color:var(--muted); text-decoration:line-through; margin-right:5px; }
    .box { background:var(--card); border:1px solid var(--border); border-radius:16px; padding:18px; margin-bottom:16px; }
    .box h2 { font-family:var(--serif); font-size:17px; color:var(--copper-dark); margin-bottom:10px; }
    .col-chips { display:flex; flex-wrap:wrap; gap:8px; }
    .col-chip { text-decoration:none; font-size:12px; font-weight:600; color:var(--ink); background:rgba(255,251,240,.85); border:1px solid var(--border); padding:6px 12px; border-radius:99px; }
    .col-chip:hover { border-color:var(--copper); color:var(--copper-dark); }
    footer { text-align:center; font-size:12px; color:var(--muted); margin-top:20px; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="topbar">
      <a class="back" href="../../">← Quay lại Phòng Nghe Số Hue Sound Express</a>
      <span style="font-weight:800;font-size:13px;color:var(--copper);">BỘ SƯU TẬP ANALOG AUDIO</span>
    </div>
    <header class="hero-col">
      <h1>$cTitleEsc</h1>
      <p>$cDescEsc Hiển thị <strong>$($shown.Count)</strong> sản phẩm tuyển chọn (tự động cập nhật khi có đĩa mới).</p>
      <div class="cta-bar">
        <a class="btn" href="https://zalo.me/0932313918" target="_blank" rel="noopener noreferrer">💬 Tư vấn chọn đĩa qua Zalo (09323.13918)</a>
        <a class="btn btn-wa" href="https://wa.me/84932313918" target="_blank" rel="noopener noreferrer">💬 WhatsApp</a>
      </div>
    </header>
    <section class="grid">$itemCardsHtml</section>
    <section class="box">
      <h2>✨ Khám phá các Bộ Sưu Tập Analog Audio khác</h2>
      <div class="col-chips">$otherColsHtml</div>
    </section>
    <footer>© 2026 Hue Sound Express – Where sound finds its color • Analog Audio • Hotline/Zalo: 09323.13918</footer>
  </div>
</body>
</html>
"@
        [System.IO.File]::WriteAllText((Join-Path $cDir "index.html"), $cHtml, $utf8NoBom)
        [void]$sm.AppendLine("  <url><loc>$cUrl</loc><priority>0.9</priority></url>")
    }

    # Bao cao tu dong phan phoi san pham moi vao Nhom & Bo Suu Tap
    $newItemsAdded = @($allItems | Where-Object { $prevSkus.Count -gt 0 -and (-not $prevSkus.Contains($_.sku)) })
    $reportSb = New-Object System.Text.StringBuilder
    [void]$reportSb.AppendLine("=== BAO CAO PHAN PHOI TU DONG VAO NHOM & BO SUU TAP ($(Get-Date -Format 'yyyy-MM-dd HH:mm')) ===")
    if ($newItemsAdded.Count -gt 0) {
        Write-Host "`n[PHAN PHOI TU DONG] Phat hien $($newItemsAdded.Count) san pham moi duoc phan loai & dua vao bo suu tap:" -ForegroundColor Cyan
        [void]$reportSb.AppendLine("San pham moi phat hien: $($newItemsAdded.Count)")
        foreach ($ni in $newItemsAdded) {
            $colsJoined = if ($skuToCols.ContainsKey($ni.sku)) { ($skuToCols[$ni.sku] -join ', ') } else { "(chua thuoc bo suu tap chu de)" }
            $line = "  + [$($ni.category.ToUpper()) / $($ni.g)] $($ni.sku) | $($ni.a) - $($ni.t) => Bo suu tap: $colsJoined"
            Write-Host $line -ForegroundColor Green
            [void]$reportSb.AppendLine($line)
        }
    } elseif ($prevSkus.Count -eq 0) {
        [void]$reportSb.AppendLine("Khoi tao snapshot dau tien cho $($allItems.Count) san pham va $($collectionMeta.Count) bo suu tap.")
    } else {
        [void]$reportSb.AppendLine("Khong co SKU moi trong lan build nay (tong: $($allItems.Count) san pham).")
    }
    [void]$reportSb.AppendLine("Thong ke 10 bo suu tap:")
    foreach ($col in $collectionMeta) {
        $oldC = if ($prevColCounts.ContainsKey($col.slug)) { $prevColCounts[$col.slug] } else { $newColCounts[$col.slug] }
        $newC = $newColCounts[$col.slug]
        $diff = $newC - $oldC
        $diffStr = if ($diff -gt 0) { " (+$diff)" } elseif ($diff -lt 0) { " ($diff)" } else { "" }
        [void]$reportSb.AppendLine("  - $($col.slug): $newC sp$diffStr")
    }
    [System.IO.File]::WriteAllText((Join-Path (Get-Location) "data\last_distribution_report.txt"), $reportSb.ToString(), $utf8NoBom)

    $snapObj = [PSCustomObject]@{
        updated_at        = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss+07:00")
        total_items       = $allItems.Count
        skus              = @($allItems | ForEach-Object { $_.sku })
        collection_counts = $newColCounts
    }
    [System.IO.File]::WriteAllText($snapPath, ($snapObj | ConvertTo-Json -Depth 3), $utf8NoBom)

    # 6. SINH 6 BAI VIET TINH CAM NANG ANALOG AUDIO (/cam-nang/<slug>/index.html)
    $guideRoot = Join-Path (Get-Location) "cam-nang"
    if (-not (Test-Path $guideRoot)) { New-Item -ItemType Directory -Path $guideRoot -Force | Out-Null }

    $guideDefs = @(
        @{
            slug = "bao-quan-dia-than-thoi-tiet-viet-nam"; utm = "guide_bao_quan_dia_than"
            title = "Cách bảo quản đĩa than với thời tiết nóng ẩm ở Việt Nam"
            desc = "Hướng dẫn chi tiết từ Hue Sound Vinyl Care về cách chống mốc bìa, chống cong vênh và giữ rãnh đĩa than bền đẹp hàng chục năm trong khí hậu nhiệt đới gió mùa Việt Nam."
            body = "<p>Khí hậu Việt Nam đặc trưng bởi <strong>mùa hè nắng nóng trên 38°C</strong> và <strong>mùa nồm ẩm miền Bắc có lúc vượt 85–90%</strong>. Đây là hai kẻ thù lớn nhất khiến đĩa than (Vinyl LP) bị cong vênh phôi PVC và mốc loang bìa giấy.</p><h2>1. Kiểm soát nhiệt độ &amp; độ ẩm phòng nghe</h2><ul><li><strong>Độ ẩm lý tưởng:</strong> duy trì trong khoảng <strong>45% – 55%</strong> bằng máy hút ẩm hoặc điều hòa chế độ Dry vào mùa nồm.</li><li><strong>Nhiệt độ an toàn:</strong> giữ dưới 30°C, tuyệt đối tránh đặt kệ đĩa cạnh cửa sổ nắng chiếu trực tiếp hoặc gần ampli đèn tỏa nhiệt mạnh.</li></ul><h2>2. Luôn xếp đĩa thẳng đứng &amp; nâng cấp bao lót</h2><ul><li>Xếp đĩa thẳng đứng như xếp sách, không xếp chồng ngang quá 3 đĩa để tránh lực nén gây cong vênh rãnh đĩa.</li><li>Thay thế túi giấy nguyên bản bằng <strong>túi lót chống tĩnh điện (Anti-Static Inner Sleeve)</strong> và bọc áo nilon ngoài để tránh mốc bìa giấy.</li></ul><h2>3. Vệ sinh định kỳ bằng chổi Carbon &amp; rửa siêu âm</h2><p>Trước mỗi lần hạ kim, lướt nhẹ chổi sợi Carbon 2–3 vòng để triệt tiêu bụi mịn và tĩnh điện. Với đĩa lâu ngày, sử dụng <strong>dịch vụ rửa đĩa siêu âm tại Hue Sound Vinyl Care</strong>.</p>"
            related = @($accLookup | Where-Object { $_.sku -in @("ACC-SLEEVE-50","ACC-AQ-CARBON","CARE-ULTRASONIC-01") }) + @($vinylLookup | Select-Object -First 3)
        },
        @{
            slug = "dia-than-moi-co-can-rua-truoc-khi-nghe"; utm = "guide_rua_dia_than_moi"
            title = "Đĩa than mới mua (Mới 100% / New Seal) có cần rửa trước khi nghe không?"
            desc = "Giải đáp thắc mắc phổ biến nhất của người chơi Analog Audio: Vì sao đĩa than nguyên seal vừa bóc vẫn có tiếng nổ lách tách và cách xử lý chuẩn Hue Sound Vinyl Care."
            body = "<p>Rất nhiều người mới chơi bất ngờ khi bóc một chiếc đĩa than <strong>Mới 100% (M / New Seal)</strong> nhưng đặt lên mâm vẫn nghe vài tiếng nổ lách tách nhẹ. Nguyên nhân đến từ quy trình ép đĩa tại nhà máy:</p><h2>1. Hai nguyên nhân đĩa mới vẫn bám bụi vi mô</h2><ul><li><strong>Lớp chất tách khuôn (Mold Release Compound):</strong> một lượng nhỏ dầu/sáp công nghiệp giúp tách đĩa PVC khỏi khuôn ép vẫn có thể sót lại dưới đáy rãnh sâu.</li><li><strong>Bụi giấy &amp; tĩnh điện:</strong> trong quá trình đóng gói bằng túi giấy thô và vận chuyển đường dài, ma sát tạo ra tĩnh điện hút vụn giấy vào mặt đĩa.</li></ul><h2>2. Quy trình 2 bước chuẩn trước lần nghe đầu tiên</h2><ol><li><strong>Cơ bản tại nhà:</strong> Quét sạch mặt đĩa bằng chổi sợi Carbon chống tĩnh điện (AudioQuest) và thay ngay bằng túi lót chống tĩnh điện mới.</li><li><strong>Chuẩn Audiophile:</strong> Rửa đĩa bằng máy rửa siêu âm (Ultrasonic Cleaning) giúp đẩy sạch 100% lớp dầu khuôn ép, làm nền âm tĩnh lặng và kéo dài tuổi thọ đầu kim.</li></ol>"
            related = @($accLookup | Where-Object { $_.sku -in @("CARE-ULTRASONIC-01","ACC-AQ-CARBON","ACC-SLEEVE-50","ACC-AT607A") }) + @($vinylLookup | Select-Object -First 2)
        },
        @{
            slug = "dong-goi-dia-than-chong-rach-mep-di-tinh"; utm = "guide_dong_goi_3_lop"
            title = "Đóng gói đĩa than gửi đi tỉnh có sợ cong vênh hay rách mép (Seam Splits) không?"
            desc = "Quy chuẩn đóng gói 3 lớp Seam Splits Prevention độc quyền tại Hue Sound Express giúp bảo vệ góc bìa và phôi đĩa than an toàn tuyệt đối khi vận chuyển toàn quốc."
            body = "<p>Nỗi lo lớn nhất khi mua đĩa than online ở xa là tình trạng <strong>rách mép bìa (Seam Splits)</strong> do phôi đĩa bên trong trượt mạnh khi vận chuyển, hoặc <strong>gãy góc bìa</strong> do hộp đóng gói mỏng.</p><h2>Quy chuẩn đóng gói 3 lớp tại Hue Sound Express</h2><ol><li><strong>Lớp 1 – Cố định &amp; chống sốc:</strong> Bọc đĩa trong nhiều lớp xốp nổ dày, chèn đệm cố định để phôi đĩa không bị xê dịch va đập vào mép bìa.</li><li><strong>Lớp 2 – Tấm bìa carton cứng bảo vệ 4 góc:</strong> Ép kẹp hai mặt bằng bìa carton cứng kích thước lớn hơn bìa đĩa giúp phân tán lực khi va chạm.</li><li><strong>Lớp 3 – Hộp Mailer 12-inch chuyên dụng:</strong> Đóng trong hộp carton 5 lớp chuyên biệt cho đĩa than, dán nhãn hàng dễ vỡ và hỗ trợ đổi trả miễn phí trong 7 ngày nếu phát sinh sự cố vận chuyển.</li></ol>"
            related = @($vinylLookup | Select-Object -First 4) + @($accLookup | Where-Object { $_.sku -in @("ACC-SLEEVE-50","ACC-AQ-CARBON") })
        },
        @{
            slug = "cach-chon-mam-than-dau-tien"; utm = "guide_chon_mam_than"
            title = "Cách chọn mâm đĩa than (Turntable) đầu tiên cho người mới bắt đầu"
            desc = "Cẩm nang chọn mua máy hát đĩa than đầu tiên: phân biệt truyền động Belt-Drive vs Direct-Drive, Phono tích hợp, Bluetooth và ngân sách hợp lý."
            body = "<p>Bước vào thế giới <strong>Analog Audio</strong> rất đơn giản nếu bạn nắm vững 4 tiêu chí chọn mâm than đầu tiên:</p><h2>1. Tránh xa các máy hát vali giá rẻ không chỉnh được lực tì kim</h2><p>Các máy hát vali không có đối trọng (Counterweight) thường đè lực kim quá nặng (5–6 gram), rất nhanh làm mòn rãnh đĩa quý.</p><h2>2. Ưu tiên mâm có Phono Preamp tích hợp (Built-in Phono)</h2><p>Chọn mâm tích hợp sẵn Phono giúp bạn cắm thẳng vào loa active / loa Bluetooth có cổng AUX/RCA là nghe được ngay, sau này vẫn dễ nâng cấp Phono rời.</p><h2>3. Belt-Drive hay Direct-Drive?</h2><ul><li><strong>Belt-Drive (dây curoa):</strong> cách ly rung động motor cực tốt, nền âm êm ấm — lựa chọn hàng đầu cho nghe nhạc gia đình và Audiophile.</li><li><strong>Direct-Drive (truyền động trực tiếp):</strong> tốc độ ổn định tức thì, bền bỉ.</li></ul><p><em>Lưu ý: Giảm giá phụ kiện khi mua cùng đĩa than / mâm than — inbox Zalo 09323.13918 để được tư vấn combo tối ưu.</em></p>"
            related = @($ttLookup) + @($accLookup | Where-Object { $_.sku -in @("ACC-SCALE-01","CARE-SETUP-01") })
        },
        @{
            slug = "phan-biet-kim-mm-va-mc"; utm = "guide_kim_mm_mc"
            title = "Phân biệt kim đĩa than MM và MC – Nên chọn loại nào cho dàn Analog Audio?"
            desc = "So sánh chi tiết cấu tạo, chất âm, chi phí thay mũi kim và yêu cầu phối ghép Phono giữa kim Moving Magnet (MM) và Moving Coil (MC)."
            body = "<p>Đầu kim (Cartridge) là trái tim chuyển đổi rung động trong rãnh đĩa thành tín hiệu âm thanh. Có hai chuẩn phổ biến nhất là <strong>MM</strong> và <strong>MC</strong>:</p><h2>1. Kim MM (Moving Magnet – Nam châm động)</h2><ul><li><strong>Ưu điểm:</strong> Điện áp ra cao (3–5mV), tương thích mọi Phono tích hợp; <strong>thay riêng mũi kim (Stylus) rất dễ và tiết kiệm</strong>; chất âm dày ấm.</li><li><strong>Phù hợp:</strong> Người mới chơi đến trung cấp, nghe đa dạng từ Nhạc Việt trẻ, Tình khúc 1954–1975 đến Bolero, Jazz.</li></ul><h2>2. Kim MC (Moving Coil – Cuộn dây động)</h2><ul><li><strong>Ưu điểm:</strong> Khối lượng chuyển động siêu nhẹ giúp bắt chi tiết vi mô, dải cao tơi mịn và âm hình sân khấu 3D vượt trội.</li><li><strong>Lưu ý:</strong> Cần Phono Preamp hỗ trợ MC hoặc Step-up Transformer (SUT).</li></ul>"
            related = @($ttLookup | Select-Object -First 3) + @($accLookup | Where-Object { $_.sku -in @("ACC-AT607A","ACC-SCALE-01","CARE-SETUP-01") })
        },
        @{
            slug = "giai-thich-thang-goldmine-grading"; utm = "guide_goldmine_grading"
            title = "Giải thích thang đánh giá tình trạng đĩa Goldmine Grading Standard (M, NM, VG+, VG...)"
            desc = "Cách đọc hiểu chuẩn phân loại tình trạng đĩa than quốc tế Goldmine Grading Standard được áp dụng minh bạch trên toàn bộ sản phẩm tại Hue Sound Express."
            body = "<p>Để đảm bảo minh bạch tuyệt đối khi mua đĩa than trực tuyến, <strong>Hue Sound Express</strong> áp dụng chuẩn quốc tế <strong>Goldmine Grading Standard</strong>:</p><ul><li><strong>Mới 100% (M / New Seal):</strong> Nguyên seal nhà sản xuất, chưa từng bóc mở hay hạ kim.</li><li><strong>NM (Near Mint):</strong> Gần như mới hoàn hảo, bề mặt đĩa bóng sạch, bìa nguyên vẹn.</li><li><strong>VG+ (Very Good Plus):</strong> Đĩa đã qua sử dụng cẩn thận, phát sạch không ảnh hưởng thưởng thức.</li><li><strong>VG (Very Good):</strong> Có tiếng lách tách nhẹ ở đoạn tĩnh, vẫn nghe mượt mà không vấp đĩa.</li><li><strong>G+ / G / F / P:</strong> Các cấp đĩa sưu tầm sâu, có dấu ấn thời gian rõ rệt.</li></ul>"
            related = @($vinylLookup | Select-Object -First 6)
        }
    )

    foreach ($gd in $guideDefs) {
        $gDir = Join-Path $guideRoot $gd.slug
        if (-not (Test-Path $gDir)) { New-Item -ItemType Directory -Path $gDir -Force | Out-Null }
        $gUrl = "$SiteBaseUrl/cam-nang/$($gd.slug)/"
        $gTitleEsc = Escape-Html $gd.title
        $gDescEsc  = Escape-Html $gd.desc

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
          { "@type": "ListItem", "position": 2, "name": "Cẩm nang FAQ và giao vận", "item": "$SiteBaseUrl/?cat=guide" },
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
    .up-card { text-decoration:none; background:rgba(255,251,240,.8); border:1px solid var(--border); padding:14px; border-radius:12px; color:var(--ink); display:flex; flex-direction:column; gap:4px; }
    .up-card span { color:var(--copper); font-weight:800; font-size:14px; margin-top:4px; }
    .other-list a { color:var(--copper-dark); text-decoration:none; font-weight:600; }
    .cta-bar { display:flex; flex-wrap:wrap; gap:10px; margin-top:16px; }
    .btn { text-decoration:none; padding:10px 16px; border-radius:10px; font-weight:600; font-size:13px; color:#FFF7E8; background:linear-gradient(135deg,#DC8644,#B85922,#8C3A12); }
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
    Write-Host "Da sinh $($allItems.Count) trang SEO /p/, $redirectCount redirect URL cu, 10 /bo-suu-tap/, 6 /cam-nang/ va sitemap.xml!" -ForegroundColor Green
}

Write-Host "Build hoan tat! Tong so san pham hop nhat: $($allItems.Count)" -ForegroundColor Green
Write-Host " - data\catalog_all.json"
Write-Host " - feed\products.json"
Write-Host " - llms.txt"
