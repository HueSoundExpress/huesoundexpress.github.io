<#
.SYNOPSIS
    Hue Sound Express - Dong bo va sua anh bia / bo sung dia than tu cac nguon tham khao Viet Nam:
    1. Gia Dinh Audio (https://giadinhaudio.vn/dia-than)
    2. Hang Dia Thoi Dai (https://shop.hangdiathoidai.com/collections/dia-than-vinyl & an-ban-thoi-dai)
    3. Playlist YouTube VIOLLAGE (cho ID 40 Viollage - Qinie)
    - Sua anh/am thanh bi khop nham (single 1 bai / khac nghe si)
    - Uu tien anh bia dia than that cho cac an ban Nhac Viet
    - Bo sung cac dia than chua co trong kho voi Gia=0 (Lien he), Tinh trang=Lien he, Trang thai=Lien he.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\scripts\sync_vn_sources.ps1
#>

param(
    [string]$VinylCsv = "data\vinyl_lp_enriched.csv",
    [string]$OutputJson = "data\catalog_enriched.json",
    [string]$MissingReportCsv = "data\missing_media_report.csv",
    [int]$MaxHtdPages = 10,
    [int]$MaxGdaPages = 12,
    [int]$MaxNewItems = 120
)

$ErrorActionPreference = "Continue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Headers = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) HueSoundExpressBot/1.0" }

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
    $c = (Remove-Diacritics $Text).ToLowerInvariant()
    $c = [regex]::Replace($c, '[^a-z0-9]+', '-')
    return $c.Trim('-')
}

function Normalize-Key {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = [System.Net.WebUtility]::HtmlDecode($Text)
    $t = [regex]::Replace($t, '\[[^\]]*\]', ' ')
    $t = [regex]::Replace($t, '\([^\)]*\)', ' ')
    $t = [regex]::Replace($t, '(?i)(đĩa than|dia than|vinyl|2xlp|2lp|3lp|lp|180g|180gr|45rpm|33rpm|nhiều ca sĩ|nhieu ca si|bản chữ ký|ban chu ky|gia định audio|hãng đĩa thời đại)', ' ')
    $t = (Remove-Diacritics $t).ToLowerInvariant()
    $t = [regex]::Replace($t, '[^a-z0-9]+', ' ').Trim()
    return $t
}

function Get-Tokens {
    param([string]$Text)
    $k = Normalize-Key $Text
    if (-not $k) { return @() }
    return @($k -split '\s+' | Where-Object { $_.Length -ge 2 } | Select-Object -Unique)
}

function Get-TokenScoreFast {
    param([string[]]$tA, [string[]]$tB)
    if (-not $tA -or -not $tB -or $tA.Count -eq 0 -or $tB.Count -eq 0) { return 0.0 }
    $hits = 0
    foreach ($tk in $tA) { if ($tB -contains $tk) { $hits++ } }
    if ($hits -eq 0) { return 0.0 }
    $minLen = [Math]::Min($tA.Count, $tB.Count)
    $cov = [double]$hits / [double]$minLen
    $jac = [double]$hits / [double]($tA.Count + $tB.Count - $hits)
    return ($cov * 0.55) + ($jac * 0.45)
}

function Get-OverlapScore {
    param([string]$A, [string]$B)
    return Get-TokenScoreFast -tA (Get-Tokens $A) -tB (Get-Tokens $B)
}

function Write-Utf8BomCsv {
    param([array]$Rows, [string]$Path)
    if (-not $Rows -or $Rows.Count -eq 0) { return }
    $csvLines = $Rows | ConvertTo-Csv -NoTypeInformation
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines((Join-Path (Get-Location) $Path), $csvLines, $utf8Bom)
}

function Is-VietnameseText {
    param([string]$Text, [string]$Vendor)
    if ($Vendor -match '(?i)(thời đại|gia định|viết tân|phương nam|thúy nga|bến thành|sài gòn|tmt|bk production|năm sao|tuấn trinh)') { return $true }
    # Kiem tra cac ky tu dau tieng Viet dac trung
    if ($Text -match '[ăâđêôơưĂÂĐÊÔƠƯáàảãạấầẩẫậắằẳẵặéèẻẽẹếềểễệíìỉĩịóòỏõọốồổỗộớờởỡợúùủũụứừửữựýỳỷỹỵ]') { return $true }
    return $false
}

function Parse-ArtistAndTitle {
    param([string]$RawTitle, [string]$DefaultArtist = "Nhiều ca sĩ")
    $clean = [System.Net.WebUtility]::HtmlDecode($RawTitle).Trim()
    $clean = [regex]::Replace($clean, '(?i)^LP\s+', '')
    $clean = [regex]::Replace($clean, '(?i)\s*[-–]\s*Đĩa Than\s*$', '')
    $clean = [regex]::Replace($clean, '(?i)\s*\((?:[^)]*(?:vinyl|lp|180g|gram|translucent|black|colour|color|edition|limited|disc)[^)]*)\)', '')
    $clean = [regex]::Replace($clean, '\s+', ' ').Trim()

    $artist = $DefaultArtist
    $title = $clean
    if ($clean -match '^(.+?)\s+[-–]\s+(.+)$') {
        $left = $Matches[1].Trim()
        $right = $Matches[2].Trim()
        if ($left.Length -ge 2 -and $left.Length -le 48 -and $right.Length -ge 2) {
            $artist = $left
            $title = $right
        }
    }
    return [PSCustomObject]@{ artist = $artist; title = $title }
}

function Search-iTunesStrictAudio {
    param([string]$Artist, [string]$Title, [bool]$IsVN)
    $country = if ($IsVN) { "vn" } else { "us" }
    $qArtist = if ($Artist -match '(?i)nhiều ca sĩ|various') { "" } else { $Artist }
    $q = "$qArtist $Title".Trim()
    if (-not $q) { return $null }
    try {
        $url = "https://itunes.apple.com/search?term=$([Uri]::EscapeDataString($q))&entity=album&country=$country&limit=6"
        $res = Invoke-RestMethod -Uri $url -Headers $Headers -TimeoutSec 12 -ErrorAction Stop
        if (-not $res.results -or $res.results.Count -eq 0) { return $null }
        $best = $null; $bestScore = 0.0
        foreach ($r in $res.results) {
            if ([int]$r.trackCount -lt 3) { continue } # Loai bo single 1-2 bai de tranh khop nham
            $tScore = Get-OverlapScore -A $Title -B ([string]$r.collectionName)
            $aScore = if ($qArtist) { Get-OverlapScore -A $qArtist -B ([string]$r.artistName) } else { 0.65 }
            if ($qArtist -and $aScore -lt 0.5) { continue }
            if ($tScore -lt 0.5) { continue }
            $sc = ($tScore * 0.6) + ($aScore * 0.4)
            if ($sc -gt $bestScore) { $bestScore = $sc; $best = $r }
        }
        if (-not $best -or $bestScore -lt 0.58) { return $null }
        Start-Sleep -Milliseconds 120
        $lk = Invoke-RestMethod -Uri "https://itunes.apple.com/lookup?id=$($best.collectionId)&entity=song&country=$country" -Headers $Headers -TimeoutSec 12 -ErrorAction Stop
        $tracks = @(); $aud = ""; $audTrack = ""
        foreach ($tr in $lk.results) {
            if ($tr.wrapperType -eq "track" -and $tr.trackName) {
                $tracks += [string]$tr.trackName
                if (-not $aud -and $tr.previewUrl) { $aud = [string]$tr.previewUrl; $audTrack = [string]$tr.trackName }
            }
        }
        $yr = if ($best.releaseDate -and ([string]$best.releaseDate).Length -ge 4) { ([string]$best.releaseDate).Substring(0,4) } else { "" }
        return [PSCustomObject]@{ audio_url = $aud; audio_track = $audTrack; tracklist = $tracks; year = $yr }
    } catch {
        return $null
    }
}

# === BUOC 1: THU THAP DU LIEU TU CAC NGUON THAM KHAO ===
Write-Host "=== Buoc 1: Thu thap danh muc dia than tu Gia Dinh Audio & Hang Dia Thoi Dai ===" -ForegroundColor Yellow
$refCatalog = New-Object System.Collections.ArrayList

# 1A. Gia Dinh Audio (giadinhaudio.vn/dia-than)
for ($p = 1; $p -le $MaxGdaPages; $p++) {
    $url = if ($p -eq 1) { "https://giadinhaudio.vn/dia-than" } else { "https://giadinhaudio.vn/dia-than?page=$p" }
    try {
        $html = (Invoke-WebRequest -Uri $url -Headers $Headers -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop).Content
        $matches = [regex]::Matches($html, '<img[^>]+(?:data-src|src)="([^"]*bizweb\.dktcdn\.net[^"]*products[^"]*)"[^>]*alt="([^"]+)"')
        if ($matches.Count -eq 0) { break }
        $addedPage = 0
        foreach ($m in $matches) {
            $img = $m.Groups[1].Value.Trim()
            if ($img.StartsWith("//")) { $img = "https:$img" }
            $rawTitle = [System.Net.WebUtility]::HtmlDecode($m.Groups[2].Value.Trim())
            if ([string]::IsNullOrWhiteSpace($rawTitle)) { continue }
            $dup = $refCatalog | Where-Object { $_.raw_title -eq $rawTitle -and $_.source -eq "giadinhaudio.vn" }
            if (-not $dup) {
                $fmt = if ($rawTitle -match '(?i)2\s*x?\s*lp') { "2xLP" } else { "LP" }
                [void]$refCatalog.Add([PSCustomObject]@{
                    raw_title = $rawTitle
                    vendor    = "Gia Định Audio"
                    format    = $fmt
                    cover_url = $img
                    source    = "giadinhaudio.vn"
                    is_vn     = (Is-VietnameseText -Text $rawTitle -Vendor "Gia Định Audio")
                })
                $addedPage++
            }
        }
        Write-Host "  [Gia Dinh Audio] Trang $p : +$addedPage dia" -ForegroundColor Cyan
        if ($addedPage -eq 0) { break }
        Start-Sleep -Milliseconds 250
    } catch {
        Write-Host "  [Gia Dinh Audio] Dung o trang $p" -ForegroundColor DarkGray
        break
    }
}

# 1B. Hang Dia Thoi Dai (shop.hangdiathoidai.com)
$htdCollections = @("an-ban-thoi-dai", "an-ban-co-chu-ky", "dia-than-vinyl")
foreach ($col in $htdCollections) {
    $maxP = if ($col -eq "dia-than-vinyl") { $MaxHtdPages } else { 3 }
    for ($p = 1; $p -le $maxP; $p++) {
        $url = "https://shop.hangdiathoidai.com/collections/$col/products.json?limit=50&page=$p"
        try {
            $res = Invoke-RestMethod -Uri $url -Headers $Headers -TimeoutSec 15 -ErrorAction Stop
            if (-not $res.products -or $res.products.Count -eq 0) { break }
            $addedPage = 0
            foreach ($pr in $res.products) {
                $ptype = [string]$pr.product_type
                $rtitle = [string]$pr.title
                # Chi lay Dia Than (Vinyl / LP), bo CD/Cassette/Phu kien ao thun
                if ($ptype -notmatch '(?i)vinyl' -and $rtitle -notmatch '(?i)(đĩa than|vinyl|\blp\b)') { continue }
                if ($rtitle -match '(?i)(cassette|băng cối|\bcd\b|merch|túi|áo|sách)' -and $rtitle -notmatch '(?i)(vinyl|đĩa than)') { continue }
                $img = if ($pr.images -and $pr.images.Count -gt 0) { [string]$pr.images[0].src } else { "" }
                if (-not $img) { continue }
                $dup = $refCatalog | Where-Object { $_.raw_title -eq $rtitle -and $_.source -eq "hangdiathoidai.com" }
                if (-not $dup) {
                    $vendor = if ($pr.vendor -and $pr.vendor -ne "Khác") { [string]$pr.vendor } else { "Hãng Đĩa Thời Đại" }
                    $fmt = if ($rtitle -match '(?i)3\s*x?\s*lp') { "3xLP" } elseif ($rtitle -match '(?i)2\s*x?\s*lp') { "2xLP" } else { "LP" }
                    [void]$refCatalog.Add([PSCustomObject]@{
                        raw_title = $rtitle
                        vendor    = $vendor
                        format    = $fmt
                        cover_url = $img
                        source    = "hangdiathoidai.com"
                        is_vn     = (Is-VietnameseText -Text $rtitle -Vendor $vendor)
                    })
                    $addedPage++
                }
            }
            Write-Host "  [Hang Dia Thoi Dai / $col] Trang $p : +$addedPage dia" -ForegroundColor Cyan
            if ($res.products.Count -lt 50) { break }
            Start-Sleep -Milliseconds 250
        } catch {
            break
        }
    }
}

Write-Host "Tong so dia than thu thap duoc tu cac nguon tham khao: $($refCatalog.Count)" -ForegroundColor Green

# Pre-cache tokens cho toan bo refCatalog de so khop sieu nhanh
foreach ($ref in $refCatalog) {
    $ref | Add-Member -NotePropertyName 'tokens' -NotePropertyValue (Get-Tokens $ref.raw_title) -Force
}

# === BUOC 2: LAM SACH LOI KHOP NHAM & SUA ANH BIA CHO KHO HIEN TAI ===
Write-Host "=== Buoc 2: Sua anh/am thanh sai & dong bo anh bia that cho 456 dia hien tai ===" -ForegroundColor Yellow
$currentRows = @(Import-Csv -Path $VinylCsv -Encoding UTF8)
$cleanedWrongCount = 0
$upgradedCoverCount = 0
$existingTokensList = New-Object System.Collections.ArrayList

foreach ($row in $currentRows) {
    $isVN = ($row.genre -match 'Nhạc Việt|Bolero')
    $trackArr = if ($row.tracklist) { @($row.tracklist -split '\s*\|\s*' | Where-Object { $_ }) } else { @() }

    # 2A. Phat hien va xoa bo cac ket qua khop nham ro rang (single 1 bai / bai nhac thieu nhi / nam 2026 gia)
    $isBadMatch = $false
    if ($row.audio_track -match '(?i)(cá vàng bơi|anh muốn nghe giọng em|tôi nhớ em vào một chiều thu)') { $isBadMatch = $true }
    if ($isVN -and $trackArr.Count -eq 1 -and $row.media_source -match 'itunes|deezer') { $isBadMatch = $true }
    if ($isVN -and $row.year -eq "2026" -and $row.artist -match '(?i)quỳnh anh|hoài phan') { $isBadMatch = $true }

    if ($isBadMatch) {
        $row.cover_url = ""
        $row.audio_url = ""
        $row.audio_track = ""
        $row.tracklist = ""
        $row.year = ""
        $row.media_source = "cleared-mismatch"
        $cleanedWrongCount++
    }

    # 2B. Dac tri ID 40: Viollage - Qinie tu Playlist YouTube chinh chu
    if ($row.artist -match '(?i)viollage' -and -not $row.cover_url) {
        $row.cover_url = "https://i9.ytimg.com/s_p/OLAK5uy_miu5FEalaoYT18g3QiEKvSL1tfKcaM87g/mqdefault.jpg?sqp=CKzfgdgGir7X7AMICK7zz84GEAE=&rs=AOn4CLAy9hOGvrxZ1HPLv1bmefSvtq46cA&v=1775499694"
        $row.media_source = "youtube-playlist"
        $upgradedCoverCount++
    }

    # 2C. Tim anh bia dia than THAT tu Gia Dinh Audio & Hang Dia Thoi Dai (pre-cached tokens)
    $tTitle = Get-Tokens $row.title
    $tFull = Get-Tokens "$($row.artist) $($row.title)"
    $tArtist = Get-Tokens $row.artist
    $isVA = ($row.artist -match '(?i)nhiều ca sĩ|various')
    [void]$existingTokensList.Add([PSCustomObject]@{ tTitle = $tTitle; tFull = $tFull })

    $bestRef = $null
    $bestScore = 0.0
    foreach ($ref in $refCatalog) {
        $sTitle = Get-TokenScoreFast -tA $tTitle -tB $ref.tokens
        $sFull = Get-TokenScoreFast -tA $tFull -tB $ref.tokens
        $sc = [Math]::Max($sTitle, $sFull)
        if (-not $isVA -and $sc -lt 0.78) {
            $aSc = Get-TokenScoreFast -tA $tArtist -tB $ref.tokens
            if ($aSc -lt 0.5) { continue }
        }
        if ($sc -gt $bestScore) { $bestScore = $sc; $bestRef = $ref }
    }

    if ($bestRef -and $bestScore -ge 0.72) {
        if (-not $row.cover_url -or $isVN) {
            $row.cover_url = $bestRef.cover_url
            $row.media_source = if ($row.audio_url) { "$($bestRef.source)+audio" } else { $bestRef.source }
            $upgradedCoverCount++
        }
    }

    # Cap nhat lai missing_fields
    $miss = @()
    if (-not $row.cover_url) { $miss += "MISSING_COVER" }
    if (-not $row.audio_url) { $miss += "MISSING_AUDIO" }
    if (-not $row.tracklist) { $miss += "MISSING_TRACKLIST" }
    if (-not $row.year) { $miss += "MISSING_YEAR" }
    if ([int]$row.list_price -le 0) { $miss += "MISSING_PRICE" }
    $row.missing_fields = ($miss -join "; ")
}

Write-Host "  - Da go bo media khop nham: $cleanedWrongCount dia" -ForegroundColor Yellow
Write-Host "  - Da nang cap / bo sung anh bia dia than that: $upgradedCoverCount dia" -ForegroundColor Green

# === BUOC 3: BO SUNG CAC DIA THAN CHUA CO (TRANG THAI & TINH TRANG = LIEN HE, GIA = 0) ===
Write-Host "=== Buoc 3: Bo sung cac dia than chua co trong kho (Uu tien Nhac Viet + dia chon loc) ===" -ForegroundColor Yellow
$outList = New-Object System.Collections.ArrayList
foreach ($r in $currentRows) { [void]$outList.Add($r) }

$nextId = ($currentRows | ForEach-Object { [int]$_.id } | Measure-Object -Maximum).Maximum + 1
$addedNew = 0

$sortedRefs = $refCatalog | Sort-Object -Property @{Expression={$_.is_vn}; Descending=$true}, @{Expression={$_.source}; Ascending=$true}

foreach ($ref in $sortedRefs) {
    if ($addedNew -ge $MaxNewItems) { break }

    $exists = $false
    foreach ($ex in $existingTokensList) {
        $s1 = Get-TokenScoreFast -tA $ex.tTitle -tB $ref.tokens
        $s2 = Get-TokenScoreFast -tA $ex.tFull -tB $ref.tokens
        if ([Math]::Max($s1, $s2) -ge 0.70) { $exists = $true; break }
    }
    if ($exists) { continue }

    $parsed = Parse-ArtistAndTitle -RawTitle $ref.raw_title
    if ($parsed.title.Length -lt 2) { continue }
    [void]$existingTokensList.Add([PSCustomObject]@{ tTitle = (Get-Tokens $parsed.title); tFull = $ref.tokens })

    $prefix = if ($ref.is_vn) { "VN" } else { "NN" }
    $srcTag = if ($ref.source -eq "giadinhaudio.vn") { "GDA" } else { "HTD" }
    $sku = "$prefix-$srcTag-$nextId"
    $slug = Get-Slug "$($parsed.artist)-$($parsed.title)-$nextId"
    $genre = if ($ref.is_vn) { "Nhạc Việt" } else { "Audiophile & Khác" }
    $desc = "$($parsed.title) – $($parsed.artist) ($genre). Ấn bản đĩa than $($ref.format) ($($ref.vendor)). Vui lòng liên hệ Hue Sound Express để kiểm tra tồn kho và đặt trước."

    # Thu keo nhanh audio 30s chuan tu iTunes (chi lay khi khop chat)
    $audUrl = ""; $audTrack = ""; $trackStr = ""; $yr = ""
    $it = Search-iTunesStrictAudio -Artist $parsed.artist -Title $parsed.title -IsVN $ref.is_vn
    if ($it) {
        $audUrl = $it.audio_url
        $audTrack = $it.audio_track
        $trackStr = ($it.tracklist -join " | ")
        $yr = $it.year
    }

    $miss = @("MISSING_PRICE")
    if (-not $audUrl) { $miss += "MISSING_AUDIO" }
    if (-not $trackStr) { $miss += "MISSING_TRACKLIST" }
    if (-not $yr) { $miss += "MISSING_YEAR" }

    [void]$outList.Add([PSCustomObject]@{
        id             = $nextId
        sku            = $sku
        slug           = $slug
        category       = "vinyl"
        genre          = $genre
        mood           = ""
        artist         = $parsed.artist
        title          = $parsed.title
        label          = $ref.vendor
        format         = $ref.format
        edition        = "Liên hệ"
        year           = $yr
        list_price     = 0
        status         = "Liên hệ"
        description    = $desc
        tracklist      = $trackStr
        cover_url      = $ref.cover_url
        audio_url      = $audUrl
        audio_track    = $audTrack
        media_source   = if ($audUrl) { "$($ref.source)+itunes" } else { $ref.source }
        missing_fields = ($miss -join "; ")
    })
    $nextId++
    $addedNew++
}

# Luu vinyl_lp_enriched.csv & missing_media_report.csv
Write-Utf8BomCsv -Rows $outList -Path $VinylCsv

$missingRows = foreach ($r in $outList) {
    if ($r.missing_fields) {
        $parts = @($r.missing_fields -split '\s*;\s*' | Where-Object { $_ })
        $qEnc = [Uri]::EscapeDataString("$($r.artist) $($r.title) vinyl")
        [PSCustomObject]@{
            id                 = $r.id
            sku                = $r.sku
            artist             = $r.artist
            title              = $r.title
            genre              = $r.genre
            status             = $r.status
            list_price         = $r.list_price
            missing_count      = $parts.Count
            missing_fields     = $r.missing_fields
            current_cover_url  = $r.cover_url
            current_audio_url  = $r.audio_url
            search_discogs_url = "https://www.discogs.com/search/?q=$qEnc&type=release"
            search_google_img  = "https://www.google.com/search?tbm=isch&q=$qEnc"
            search_youtube_url = "https://www.youtube.com/results?search_query=$qEnc"
        }
    }
}
Write-Utf8BomCsv -Rows ($missingRows | Sort-Object -Property missing_count -Descending) -Path $MissingReportCsv

Write-Host ""
Write-Host "================ KET QUA DONG BO NGUON THAM KHAO ================" -ForegroundColor Green
Write-Host "Tong so dia than sau dong bo   : $($outList.Count)"
Write-Host "Da sua/go media khop nham      : $cleanedWrongCount"
Write-Host "Da nang cap anh bia dia that   : $upgradedCoverCount"
Write-Host "Da bo sung dia moi (Lien he)   : $addedNew"
Write-Host "=================================================================" -ForegroundColor Green
