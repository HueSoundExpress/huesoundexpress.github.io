<#
.SYNOPSIS
    Hue Sound Express - Cong cu tu dong keo Anh bia, Tracklist, Nam phat hanh va Audio Preview 30s.
    Ho tro da nguon mien phi (0d): iTunes Search API (VN/US) -> Deezer Public API -> MusicBrainz + Cover Art Archive -> Discogs (tuy chon).

.EXAMPLE
    # Chay thu nghiem 15 dia dau tien de kiem tra ket qua:
    powershell -ExecutionPolicy Bypass -File .\scripts\enrich_catalog.ps1 -Limit 15

    # Chay toan bo 456 dia (tu dong bo qua nhung dia da lay xong):
    powershell -ExecutionPolicy Bypass -File .\scripts\enrich_catalog.ps1 -Resume

    # Chay kem Discogs Personal Access Token (tuy chon, giup tang ty le tim thay dia hiem/Audiophile):
    powershell -ExecutionPolicy Bypass -File .\scripts\enrich_catalog.ps1 -Resume -DiscogsToken "YOUR_TOKEN"
#>

param(
    [string]$InputJson = "catalog.json",
    [string]$OutputJson = "data\catalog_enriched.json",
    [string]$OutputCsv = "data\vinyl_lp_enriched.csv",
    [string]$MissingReportCsv = "data\missing_media_report.csv",
    [int]$Limit = 0,
    [int]$StartAt = 0,
    [int]$DelayMs = 350,
    [switch]$Resume,
    [switch]$ForceRefresh,
    [string]$DiscogsToken = ""
)

$ErrorActionPreference = "Continue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Dam bao thu muc data ton tai
$dataDir = Split-Path -Parent $OutputJson
if ($dataDir -and -not (Test-Path $dataDir)) {
    New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
}

function Remove-Diacritics {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $normalized = $Text.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $normalized.ToCharArray()) {
        $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
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

function Clean-SearchText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = $Text
    # Bo cac the phien ban noi bo: [R], [black], [purple], (ban chu ky), (1981)...
    $t = [regex]::Replace($t, '\[[^\]]*\]', ' ')
    $t = [regex]::Replace($t, '\((?i)b[ảa]n ch[ữư] k[ýy].*?\)', ' ')
    $t = [regex]::Replace($t, '\((?i)limited.*?\)', ' ')
    $t = [regex]::Replace($t, '\((?i)reissue.*?\)', ' ')
    $t = [regex]::Replace($t, '\((?i)remaster.*?\)', ' ')
    $t = [regex]::Replace($t, '\(\d{4}\)', ' ')
    $t = [regex]::Replace($t, '\s+', ' ').Trim()
    return $t
}

function Normalize-ArtistForSearch {
    param([string]$Artist)
    $a = Clean-SearchText $Artist
    if ($a -match '^(?i)(nhi[ềe]u ca s[ĩi]|various|nh[ạa]c c[ụu] d[âa]n t[ộo]c|ti[ếe]ng h[áa]t vi[ệe]t nam)$') {
        return ""
    }
    # Dao "Shepp, Archie" -> "Archie Shepp"
    if ($a -match '^([^,]+),\s*([^,]+)$') {
        $a = "$($Matches[2].Trim()) $($Matches[1].Trim())"
    }
    # Bo danh hieu NSND / GS
    $a = [regex]::Replace($a, '^(?i)(NSND|NSUT|GS\.?)\s+', '').Trim()
    return $a
}

function Get-TokenOverlapScore {
    param([string]$Source, [string]$Candidate)
    if ([string]::IsNullOrWhiteSpace($Source) -or [string]::IsNullOrWhiteSpace($Candidate)) { return 0.0 }
    $sTokens = ((Remove-Diacritics $Source).ToLowerInvariant() -replace '[^a-z0-9\s]', ' ') -split '\s+' | Where-Object { $_.Length -ge 2 } | Select-Object -Unique
    $cTokens = ((Remove-Diacritics $Candidate).ToLowerInvariant() -replace '[^a-z0-9\s]', ' ') -split '\s+' | Where-Object { $_.Length -ge 2 } | Select-Object -Unique
    if (-not $sTokens -or $sTokens.Count -eq 0) { return 0.0 }
    $hits = 0
    foreach ($tk in $sTokens) {
        if ($cTokens -contains $tk) { $hits++ }
    }
    return [double]$hits / [double]$sTokens.Count
}

function Invoke-SafeJsonGet {
    param([string]$Url, [hashtable]$Headers = @{})
    try {
        if (-not $Headers.ContainsKey("User-Agent")) {
            $Headers["User-Agent"] = "HueSoundExpressCatalogBot/1.0 (https://huesoundexpress.github.io)"
        }
        $resp = Invoke-RestMethod -Uri $Url -Method Get -Headers $Headers -TimeoutSec 15 -ErrorAction Stop
        return $resp
    } catch {
        return $null
    }
}

function Search-iTunes {
    param([string]$Artist, [string]$Album, [string]$Country = "vn")
    $queryParts = @()
    if ($Artist) { $queryParts += $Artist }
    if ($Album) { $queryParts += $Album }
    $query = ($queryParts -join " ").Trim()
    if (-not $query) { return $null }

    $enc = [Uri]::EscapeDataString($query)
    $url = "https://itunes.apple.com/search?term=$enc&entity=album&country=$Country&limit=8"
    $res = Invoke-SafeJsonGet -Url $url
    if (-not $res -or -not $res.results -or $res.results.Count -eq 0) { return $null }

    $best = $null
    $bestScore = 0.0
    foreach ($item in $res.results) {
        $candArtist = [string]$item.artistName
        $candAlbum = [string]$item.collectionName
        $titleScore = Get-TokenOverlapScore -Source $Album -Candidate $candAlbum
        $artistScore = if ($Artist) { Get-TokenOverlapScore -Source $Artist -Candidate $candArtist } else { 0.7 }

        # Bat buoc khop nghe si khi co ten nghe si cu the (tranh nhan nham album trung ten cua nghe si khac)
        if ($Artist -and $artistScore -lt 0.5) { continue }
        if ($titleScore -lt 0.4) { continue }

        $score = ($titleScore * 0.55) + ($artistScore * 0.45)
        if ($score -gt $bestScore) {
            $bestScore = $score
            $best = $item
        }
    }

    if (-not $best -or $bestScore -lt 0.55) { return $null }

    $coverUrl = ""
    if ($best.artworkUrl100) {
        $coverUrl = ([string]$best.artworkUrl100) -replace '100x100bb', '600x600bb'
    }
    $year = ""
    if ($best.releaseDate -and ([string]$best.releaseDate).Length -ge 4) {
        $year = ([string]$best.releaseDate).Substring(0, 4)
    }

    $tracks = @()
    $previewUrl = ""
    $previewTrack = ""

    # Lookup chi tiet tracklist & audio 30s
    if ($best.collectionId) {
        Start-Sleep -Milliseconds 150
        $lookupUrl = "https://itunes.apple.com/lookup?id=$($best.collectionId)&entity=song&country=$Country"
        $lookup = Invoke-SafeJsonGet -Url $lookupUrl
        if ($lookup -and $lookup.results) {
            foreach ($r in $lookup.results) {
                if ($r.wrapperType -eq "track" -and $r.trackName) {
                    $tracks += [string]$r.trackName
                    if (-not $previewUrl -and $r.previewUrl) {
                        $previewUrl = [string]$r.previewUrl
                        $previewTrack = [string]$r.trackName
                    }
                }
            }
        }
    }

    return [PSCustomObject]@{
        cover_url   = $coverUrl
        year        = $year
        tracklist   = $tracks
        audio_url   = $previewUrl
        audio_track = $previewTrack
        source      = "itunes-$Country"
        score       = [Math]::Round($bestScore, 2)
    }
}

function Search-Deezer {
    param([string]$Artist, [string]$Album)
    $q = if ($Artist) { "$Artist $Album" } else { $Album }
    $enc = [Uri]::EscapeDataString($q.Trim())
    $url = "https://api.deezer.com/search/album?q=$enc&limit=8"
    $res = Invoke-SafeJsonGet -Url $url
    if (-not $res -or -not $res.data -or $res.data.Count -eq 0) { return $null }

    $best = $null
    $bestScore = 0.0
    foreach ($item in $res.data) {
        $candArtist = [string]$item.artist.name
        $candAlbum = [string]$item.title
        $tScore = Get-TokenOverlapScore -Source $Album -Candidate $candAlbum
        $aScore = if ($Artist) { Get-TokenOverlapScore -Source $Artist -Candidate $candArtist } else { 0.7 }

        # Bat buoc khop nghe si khi co ten nghe si cu the
        if ($Artist -and $aScore -lt 0.5) { continue }
        if ($tScore -lt 0.4) { continue }

        $score = ($tScore * 0.55) + ($aScore * 0.45)
        if ($score -gt $bestScore) {
            $bestScore = $score
            $best = $item
        }
    }

    if (-not $best -or $bestScore -lt 0.55) { return $null }

    $coverUrl = if ($best.cover_big) { [string]$best.cover_big } elseif ($best.cover_medium) { [string]$best.cover_medium } else { "" }
    $year = ""
    $tracks = @()
    $previewUrl = ""
    $previewTrack = ""

    if ($best.id) {
        Start-Sleep -Milliseconds 150
        $detail = Invoke-SafeJsonGet -Url "https://api.deezer.com/album/$($best.id)"
        if ($detail) {
            if ($detail.release_date -and ([string]$detail.release_date).Length -ge 4) {
                $year = ([string]$detail.release_date).Substring(0, 4)
            }
            if ($detail.tracks -and $detail.tracks.data) {
                foreach ($tr in $detail.tracks.data) {
                    if ($tr.title) { $tracks += [string]$tr.title }
                    if (-not $previewUrl -and $tr.preview) {
                        $previewUrl = [string]$tr.preview
                        $previewTrack = [string]$tr.title
                    }
                }
            }
        }
    }

    return [PSCustomObject]@{
        cover_url   = $coverUrl
        year        = $year
        tracklist   = $tracks
        audio_url   = $previewUrl
        audio_track = $previewTrack
        source      = "deezer"
        score       = [Math]::Round($bestScore, 2)
    }
}

function Search-MusicBrainz {
    param([string]$Artist, [string]$Album)
    $qParts = @("release:`"$Album`"")
    if ($Artist) { $qParts += "artist:`"$Artist`"" }
    $enc = [Uri]::EscapeDataString(($qParts -join " AND "))
    $url = "https://musicbrainz.org/ws/2/release/?query=$enc&fmt=json&limit=5"
    # MusicBrainz yeu cau gioi han 1 request/giay
    Start-Sleep -Milliseconds 1050
    $res = Invoke-SafeJsonGet -Url $url
    if (-not $res -or -not $res.releases -or $res.releases.Count -eq 0) { return $null }

    $best = $res.releases[0]
    $tScore = Get-TokenOverlapScore -Source $Album -Candidate ([string]$best.title)
    if ($tScore -lt 0.55) { return $null }
    if ($Artist -and $best.'artist-credit') {
        $mbArtist = ($best.'artist-credit' | ForEach-Object { $_.name }) -join " "
        $aScore = Get-TokenOverlapScore -Source $Artist -Candidate $mbArtist
        if ($aScore -lt 0.5) { return $null }
    }

    $year = ""
    if ($best.date -and ([string]$best.date).Length -ge 4) {
        $year = ([string]$best.date).Substring(0, 4)
    }

    $coverUrl = ""
    if ($best.id) {
        $caaUrl = "https://coverartarchive.org/release/$($best.id)"
        $caa = Invoke-SafeJsonGet -Url $caaUrl
        if ($caa -and $caa.images -and $caa.images.Count -gt 0) {
            $front = $caa.images | Where-Object { $_.front -eq $true } | Select-Object -First 1
            if (-not $front) { $front = $caa.images[0] }
            if ($front.thumbnails -and $front.thumbnails.'500') {
                $coverUrl = [string]$front.thumbnails.'500'
            } elseif ($front.image) {
                $coverUrl = [string]$front.image
            }
        }
    }

    return [PSCustomObject]@{
        cover_url   = $coverUrl
        year        = $year
        tracklist   = @()
        audio_url   = ""
        audio_track = ""
        source      = "musicbrainz"
        score       = [Math]::Round($tScore, 2)
    }
}

function Search-Discogs {
    param([string]$Artist, [string]$Album, [string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }
    $q = if ($Artist) { "$Artist $Album" } else { $Album }
    $enc = [Uri]::EscapeDataString($q.Trim())
    $url = "https://api.discogs.com/database/search?q=$enc&type=release&format=vinyl&per_page=5"
    $headers = @{
        "Authorization" = "Discogs token=$Token"
        "User-Agent"    = "HueSoundExpressCatalogBot/1.0"
    }
    Start-Sleep -Milliseconds 1000
    $res = Invoke-SafeJsonGet -Url $url -Headers $headers
    if (-not $res -or -not $res.results -or $res.results.Count -eq 0) { return $null }

    $best = $res.results[0]
    $coverUrl = if ($best.cover_image -and ($best.cover_image -notmatch 'spacer\.gif')) { [string]$best.cover_image } else { "" }
    $year = if ($best.year) { [string]$best.year } else { "" }
    $tracks = @()

    if ($best.resource_url) {
        Start-Sleep -Milliseconds 1000
        $rel = Invoke-SafeJsonGet -Url $best.resource_url -Headers $headers
        if ($rel -and $rel.tracklist) {
            foreach ($tr in $rel.tracklist) {
                if ($tr.title) {
                    $pos = if ($tr.position) { "$($tr.position). " } else { "" }
                    $tracks += "$pos$($tr.title)"
                }
            }
        }
    }

    return [PSCustomObject]@{
        cover_url   = $coverUrl
        year        = $year
        tracklist   = $tracks
        audio_url   = ""
        audio_track = ""
        source      = "discogs"
        score       = 0.8
    }
}

function Normalize-HueStatus {
    param($Item)
    $cond = [string]$Item.c
    $oldStatus = [string]$Item.s
    $price = [int]$Item.p

    # Hang hiem / Ban chu ky / Mint / Chua co gia -> Lien he
    if ($cond -match '(?i)(sign|mint|limited|picture)' -or $price -le 0) {
        return "Liên hệ"
    }
    if ($oldStatus -eq "Còn hàng") {
        return "New Seal"
    }
    return "Hết hàng"
}

function Write-Utf8BomCsv {
    param([array]$Rows, [string]$Path)
    if (-not $Rows -or $Rows.Count -eq 0) { return }
    $csvLines = $Rows | ConvertTo-Csv -NoTypeInformation
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines((Join-Path (Get-Location) $Path), $csvLines, $utf8Bom)
}

function Save-Outputs {
    param([array]$Records)
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $jsonOut = $Records | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText((Join-Path (Get-Location) $OutputJson), $jsonOut, $utf8NoBom)

    # Xuat CSV day du de mo bang Excel / Google Sheets
    $csvRows = foreach ($r in $Records) {
        [PSCustomObject]@{
            id             = $r.id
            sku            = $r.sku
            slug           = $r.slug
            category       = $r.category
            genre          = $r.g
            mood           = $r.mood
            artist         = $r.a
            title          = $r.t
            label          = $r.l
            format         = $r.f
            edition        = $r.c
            year           = $r.y
            list_price     = $r.p
            status         = $r.s
            description    = $r.d
            tracklist      = ($r.tracklist -join " | ")
            cover_url      = $r.img
            audio_url      = $r.aud
            audio_track    = $r.aud_track
            media_source   = $r.media_source
            missing_fields = ($r.missing_fields -join "; ")
        }
    }
    Write-Utf8BomCsv -Rows $csvRows -Path $OutputCsv

    # Xuat bao cao nhung dia thieu thong tin de bo sung manual
    $missingRows = foreach ($r in $Records) {
        if ($r.missing_fields -and $r.missing_fields.Count -gt 0) {
            $qEnc = [Uri]::EscapeDataString("$($r.a) $($r.t) vinyl")
            [PSCustomObject]@{
                id                 = $r.id
                sku                = $r.sku
                artist             = $r.a
                title              = $r.t
                genre              = $r.g
                status             = $r.s
                list_price         = $r.p
                missing_count      = $r.missing_fields.Count
                missing_fields     = ($r.missing_fields -join "; ")
                current_cover_url  = $r.img
                current_audio_url  = $r.aud
                search_discogs_url = "https://www.discogs.com/search/?q=$qEnc&type=release"
                search_google_img  = "https://www.google.com/search?tbm=isch&q=$qEnc"
                search_youtube_url = "https://www.youtube.com/results?search_query=$qEnc"
            }
        }
    }
    $missingSorted = $missingRows | Sort-Object -Property missing_count -Descending
    Write-Utf8BomCsv -Rows $missingSorted -Path $MissingReportCsv
}

# === BAT DAU THUC THI ===
Write-Host "=== Hue Sound Express - Auto Media & Metadata Enrichment ===" -ForegroundColor Yellow

$rawData = Get-Content -Raw -Encoding UTF8 $InputJson | ConvertFrom-Json
$existingMap = @{}
if ($Resume -and (Test-Path $OutputJson)) {
    $prev = Get-Content -Raw -Encoding UTF8 $OutputJson | ConvertFrom-Json
    foreach ($p in $prev) { $existingMap["$($p.id)"] = $p }
    Write-Host "Da nap $($existingMap.Count) ban ghi cu tu $OutputJson (Che do Resume)." -ForegroundColor Cyan
}

$enriched = New-Object System.Collections.ArrayList
$total = $rawData.Count
$processedNow = 0
$stats = @{ cover = 0; audio = 0; tracklist = 0; year = 0; complete = 0 }

for ($i = 0; $i -lt $total; $i++) {
    $item = $rawData[$i]
    $idKey = "$($item.id)"

    # Khoi tao doi tuong chuan hoa
    if ($existingMap.ContainsKey($idKey) -and -not $ForceRefresh) {
        $rec = $existingMap[$idKey]
    } else {
        $slug = Get-Slug "$($item.a)-$($item.t)-$($item.id)"
        $normStatus = Normalize-HueStatus $item
        $rec = [PSCustomObject]@{
            id             = $item.id
            sku            = $item.sku
            slug           = $slug
            category       = "vinyl"
            g              = $item.g
            mood           = ""
            a              = $item.a
            t              = $item.t
            l              = $item.l
            f              = $item.f
            c              = $item.c
            y              = [string]$item.y
            p              = [int]$item.p
            s              = $normStatus
            d              = $item.d
            tracklist      = @()
            img            = ""
            aud            = ""
            aud_track      = ""
            media_source   = ""
            missing_fields = @()
        }
    }

    $shouldFetch = ($i -ge $StartAt) -and (($Limit -eq 0) -or ($processedNow -lt $Limit))
    $alreadyHasMedia = ($rec.img -match '^https?://') -and ($rec.aud -match '^https?://') -and ($rec.tracklist.Count -gt 0) -and ($rec.y)

    if ($shouldFetch -and (-not $alreadyHasMedia -or $ForceRefresh)) {
        $processedNow++
        $cleanArtist = Normalize-ArtistForSearch $item.a
        $cleanAlbum = Clean-SearchText $item.t
        $isVN = ($item.g -match 'Nhạc Việt|Bolero')
        $primaryCountry = if ($isVN) { "vn" } else { "us" }
        $secondaryCountry = if ($isVN) { "us" } else { "vn" }

        Write-Host ("[{0}/{1}] Dang tim: {2} - {3} ..." -f ($i + 1), $total, $item.a, $item.t) -NoNewline

        $sourcesUsed = New-Object System.Collections.ArrayList

        # 1. Thu iTunes (Quoc gia chinh)
        $r1 = Search-iTunes -Artist $cleanArtist -Album $cleanAlbum -Country $primaryCountry
        # 1b. Thu iTunes khong dau neu chua thay
        if (-not $r1 -and $isVN) {
            $r1 = Search-iTunes -Artist (Remove-Diacritics $cleanArtist) -Album (Remove-Diacritics $cleanAlbum) -Country $primaryCountry
        }
        # 1c. Thu iTunes quoc gia phu
        if (-not $r1) {
            $r1 = Search-iTunes -Artist $cleanArtist -Album $cleanAlbum -Country $secondaryCountry
        }

        if ($r1) {
            if ($r1.cover_url -and -not $rec.img) { $rec.img = $r1.cover_url }
            if ($r1.audio_url -and -not $rec.aud) { $rec.aud = $r1.audio_url; $rec.aud_track = $r1.audio_track }
            if ($r1.tracklist.Count -gt 0 -and $rec.tracklist.Count -eq 0) { $rec.tracklist = $r1.tracklist }
            if ($r1.year -and -not $rec.y) { $rec.y = $r1.year }
            [void]$sourcesUsed.Add($r1.source)
        }

        # 2. Neu van thieu Cover hoac Audio hoac Tracklist -> Fallback sang Deezer Public API
        if (-not $rec.img -or -not $rec.aud -or $rec.tracklist.Count -eq 0) {
            Start-Sleep -Milliseconds $DelayMs
            $r2 = Search-Deezer -Artist $cleanArtist -Album $cleanAlbum
            if ($r2) {
                if ($r2.cover_url -and -not $rec.img) { $rec.img = $r2.cover_url }
                if ($r2.audio_url -and -not $rec.aud) { $rec.aud = $r2.audio_url; $rec.aud_track = $r2.audio_track }
                if ($r2.tracklist.Count -gt 0 -and $rec.tracklist.Count -eq 0) { $rec.tracklist = $r2.tracklist }
                if ($r2.year -and -not $rec.y) { $rec.y = $r2.year }
                [void]$sourcesUsed.Add($r2.source)
            }
        }

        # 3. Neu van thieu Cover hoac Nam -> Fallback sang MusicBrainz + Cover Art Archive
        if (-not $rec.img -or -not $rec.y) {
            $r3 = Search-MusicBrainz -Artist $cleanArtist -Album $cleanAlbum
            if ($r3) {
                if ($r3.cover_url -and -not $rec.img) { $rec.img = $r3.cover_url }
                if ($r3.year -and -not $rec.y) { $rec.y = $r3.year }
                [void]$sourcesUsed.Add($r3.source)
            }
        }

        # 4. Neu co Discogs Token va van thieu Cover hoac Tracklist
        if ($DiscogsToken -and (-not $rec.img -or $rec.tracklist.Count -eq 0)) {
            $r4 = Search-Discogs -Artist $cleanArtist -Album $cleanAlbum -Token $DiscogsToken
            if ($r4) {
                if ($r4.cover_url -and -not $rec.img) { $rec.img = $r4.cover_url }
                if ($r4.tracklist.Count -gt 0 -and $rec.tracklist.Count -eq 0) { $rec.tracklist = $r4.tracklist }
                if ($r4.year -and -not $rec.y) { $rec.y = $r4.year }
                [void]$sourcesUsed.Add($r4.source)
            }
        }

        if ($sourcesUsed.Count -gt 0) {
            $rec.media_source = ($sourcesUsed | Select-Object -Unique) -join "+"
        } else {
            $rec.media_source = "manual-needed"
        }

        $tag = if ($rec.img -and $rec.aud) { "OK (Cover+Audio)" } elseif ($rec.img) { "Cover only" } else { "THIEU" }
        Write-Host " -> $tag [$($rec.media_source)]" -ForegroundColor $(if ($tag -match 'OK') { 'Green' } elseif ($tag -match 'Cover') { 'Yellow' } else { 'DarkGray' })

        Start-Sleep -Milliseconds $DelayMs
    }

    # Danh dau cac truong con thieu de bo sung manual
    $missing = @()
    if (-not $rec.img) { $missing += "MISSING_COVER" }
    if (-not $rec.aud) { $missing += "MISSING_AUDIO" }
    if (-not $rec.tracklist -or $rec.tracklist.Count -eq 0) { $missing += "MISSING_TRACKLIST" }
    if (-not $rec.y) { $missing += "MISSING_YEAR" }
    if ($rec.p -le 0) { $missing += "MISSING_PRICE" }
    $rec.missing_fields = $missing

    if ($rec.img) { $stats.cover++ }
    if ($rec.aud) { $stats.audio++ }
    if ($rec.tracklist.Count -gt 0) { $stats.tracklist++ }
    if ($rec.y) { $stats.year++ }
    if ($missing.Count -eq 0) { $stats.complete++ }

    [void]$enriched.Add($rec)

    # Luu checkpoint sau moi 15 dia duoc fetch
    if ($processedNow -gt 0 -and ($processedNow % 15 -eq 0)) {
        Save-Outputs -Records $enriched
        Write-Host "  [Checkpoint] Da luu tien do ($($enriched.Count)/$total)..." -ForegroundColor Cyan
    }
}

# Luu ket qua cuoi cung
Save-Outputs -Records $enriched

Write-Host ""
Write-Host "================ TONG KET ENRICHMENT ================" -ForegroundColor Green
Write-Host "Tong so dia trong kho     : $total"
Write-Host "So dia vua quet luot nay  : $processedNow"
Write-Host "Co Anh bia (Cover)        : $($stats.cover) / $total"
Write-Host "Co Audio nghe thu 30s     : $($stats.audio) / $total"
Write-Host "Co Tracklist              : $($stats.tracklist) / $total"
Write-Host "Co Nam phat hanh (Year)   : $($stats.year) / $total"
Write-Host "Du 100% thong tin         : $($stats.complete) / $total"
Write-Host "File JSON da luu          : $OutputJson"
Write-Host "File CSV Excel/Sheets     : $OutputCsv"
Write-Host "Bao cao dia thieu (Manual): $MissingReportCsv"
Write-Host "=====================================================" -ForegroundColor Green
