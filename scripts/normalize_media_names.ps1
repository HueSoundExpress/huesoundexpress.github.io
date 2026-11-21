<#
.SYNOPSIS
    Tu dong chuan hoa ten file & thu muc trong images/, audio/, docs/ sang chu thuong khong dau noi bang gach ngang (kebab-case)
    de tranh ma hoa %20, %C3%A1 khi chia se Zalo/Facebook va toi uu CDN GitHub Pages.
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\scripts\normalize_media_names.ps1
#>

param(
    [string[]]$TargetFolders = @("images", "audio", "docs")
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function To-KebabAscii {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "file" }
    $norm = $Text.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $norm.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    $ascii = $sb.ToString().Replace([char]0x0111, 'd').Replace([char]0x0110, 'D').ToLowerInvariant()
    $ascii = [regex]::Replace($ascii, '[^a-z0-9]+', '-').Trim('-')
    if (-not $ascii) { return "file" }
    return $ascii
}

$renamedCount = 0
foreach ($folder in $TargetFolders) {
    if (-not (Test-Path $folder)) { continue }
    $files = Get-ChildItem -Path $folder -Recurse -File
    foreach ($f in $files) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $ext  = $f.Extension.ToLowerInvariant()
        $newName = (To-KebabAscii $base) + $ext
        if ($f.Name -cne $newName) {
            $dest = Join-Path $f.DirectoryName $newName
            $tmp  = $f.FullName + ".tmp_kebab"
            Move-Item -Path $f.FullName -Destination $tmp -Force
            Move-Item -Path $tmp -Destination $dest -Force
            Write-Host "Renamed: $($f.Name) -> $newName" -ForegroundColor Cyan
            $renamedCount++
        }
    }
}
Write-Host "Hoan tat chuan hoa ten file media/tai lieu ($renamedCount file da doi ten)." -ForegroundColor Green
