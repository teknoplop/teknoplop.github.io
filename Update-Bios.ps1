<#
.SYNOPSIS
Detects this PC's motherboard/vendor and current BIOS version, opens the correct
official support page for the latest BIOS, then offers to extract the downloaded
file onto a USB drive.

.DESCRIPTION
Covers ASUS, MSI, Gigabyte, and ASRock motherboards. None of them expose a
public API for "the latest BIOS by model" - their download pages are
JS-rendered search/listing pages - so this script automates the identification
step (board vendor, model, current BIOS version/date) and opens the correct
vendor search page pre-filled with your model, then automates the USB
extraction step once you've downloaded the file yourself.

# EXECUTION
# .\Update-Bios.ps1
# .\Update-Bios.ps1 -File "C:\Users\me\Downloads\SomeBios.zip"
#>

[CmdletBinding()]
param(
    # Path to an already-downloaded BIOS file. If omitted, the script watches
    # your Downloads folder for a new file after you confirm the download page opened.
    [string]$File,

    # Folder to watch for the downloaded BIOS file.
    [string]$DownloadsFolder = (Join-Path $env:USERPROFILE 'Downloads'),

    # Skip the "already checked recently" idempotency guard below, e.g. when
    # explicitly running this script rather than via an unattended setup pass.
    [switch]$Force,

    # How fresh the installed BIOS needs to be (by release date) to skip the check.
    [int]$SkipIfNewerThanMonths = 3
)

function Get-CleanBoardRevision($rawVersion) {
    if (-not $rawVersion) { return $null }
    $v = $rawVersion.Trim()
    # Many boards ship SMBIOS placeholder text here instead of a real revision.
    $placeholders = @('default string', 'to be filled by o.e.m.', 'system version', 'unknown', 'n/a', '')
    if ($placeholders -contains $v.ToLower()) { return $null }
    return $v
}

function Get-SystemFirmwareInfo {
    $baseBoard = Get-CimInstance Win32_BaseBoard
    $bios      = Get-CimInstance Win32_BIOS
    $cs        = Get-CimInstance Win32_ComputerSystem
    [PSCustomObject]@{
        Manufacturer   = ($cs.Manufacturer).Trim()
        Model          = ($cs.Model).Trim()
        BoardVendor    = ($baseBoard.Manufacturer).Trim()
        BoardProduct   = ($baseBoard.Product).Trim()
        BoardRevision  = Get-CleanBoardRevision $baseBoard.Version
        BiosVersion    = ($bios.SMBIOSBIOSVersion).Trim()
        BiosDate       = $bios.ReleaseDate
    }
}

function Get-VendorSupportUrl($info) {
    # Vendor support pages bake the board revision into the URL (e.g. Gigabyte's
    # /Motherboard/<model>-rev-10-11-12/support), which we usually can't determine
    # from firmware alone. Both the vendor sites and search engines actively block
    # scripted/non-browser requests (403s / CAPTCHA challenges), so this can only be
    # resolved by opening a real browser - a site-restricted search reliably lands
    # on the right page there regardless of revision, where a guessed URL 404s.
    $vendorText = "$($info.Manufacturer) $($info.BoardVendor)".ToLower()
    $model = if ($info.BoardProduct -and $info.BoardProduct -ne $info.Model) { $info.BoardProduct } else { $info.Model }

    $domain = switch -Regex ($vendorText) {
        'asus' { 'asus.com'; break }
        '(msi|micro-star)' { 'msi.com'; break }
        'gigabyte' { 'gigabyte.com'; break }
        'asrock' { 'asrock.com'; break }
        default { $null }
    }

    if (-not $domain) {
        Write-Host "Board vendor '$($info.BoardVendor)' is not one of ASUS/MSI/Gigabyte/ASRock - falling back to a general web search." -ForegroundColor Yellow
        $query = "$model bios update download"
    } else {
        # Quoting the model and searching for "BIOS" (rather than "support") consistently
        # surfaces the exact product's BIOS page as the top hit across all four vendors -
        # verified by hand for each. An unquoted "support" query can instead surface a
        # generic model-listing page (seen with ASRock).
        $query = "site:$domain `"$model`" BIOS"
    }
    return "https://duckduckgo.com/?q=$([uri]::EscapeDataString($query))"
}

function Get-RemovableDrives {
    @(Get-Volume | Where-Object { $_.DriveType -eq 'Removable' -and $_.DriveLetter })
}

Write-Host "=== Detecting hardware ===" -ForegroundColor Cyan
$info = Get-SystemFirmwareInfo
Write-Host "Manufacturer : $($info.Manufacturer)"
Write-Host "Model        : $($info.Model)"
Write-Host "Board vendor : $($info.BoardVendor)"
Write-Host "Board model  : $($info.BoardProduct)"
if ($info.BoardRevision) {
    Write-Host "Board rev.   : $($info.BoardRevision)"
} else {
    Write-Host "Board rev.   : not reported by firmware." -ForegroundColor Yellow
    Write-Host "               The only fully reliable source is the silkscreen/sticker on the board itself." -ForegroundColor Yellow
    Write-Host "               As a cross-check: vendors often split a model into per-revision tabs on their support" -ForegroundColor Yellow
    Write-Host "               page, each with its own BIOS version numbering track. Your current BIOS version below" -ForegroundColor Yellow
    Write-Host "               should fall within one tab's version list, not another - use that to confirm you're" -ForegroundColor Yellow
    Write-Host "               grabbing the update for the right revision before flashing." -ForegroundColor Yellow
}
Write-Host "Current BIOS : $($info.BiosVersion) ($($info.BiosDate))"

if (-not $File -and -not $Force -and $info.BiosDate) {
    $cutoff = (Get-Date).AddMonths(-$SkipIfNewerThanMonths)
    if ($info.BiosDate -gt $cutoff) {
        Write-Host "`nInstalled BIOS is less than $SkipIfNewerThanMonths month(s) old - skipping the update check." -ForegroundColor Green
        Write-Host "(pass -Force to check anyway, or -File <path> to go straight to the USB step)" -ForegroundColor DarkGray
        return
    }
}

if (-not $File) {
    $url = Get-VendorSupportUrl $info
    Write-Host "`n=== Opening vendor support page ===" -ForegroundColor Cyan
    Write-Host $url
    Start-Process $url

    Write-Host "`nDownload the latest BIOS update for your board from that page into:"
    Write-Host "  $DownloadsFolder" -ForegroundColor Yellow
    $before = @{}
    if (Test-Path $DownloadsFolder) {
        Get-ChildItem $DownloadsFolder -File | ForEach-Object { $before[$_.FullName] = $_.LastWriteTime }
    }
    Read-Host "`nPress Enter once the download has finished"

    $candidates = Get-ChildItem $DownloadsFolder -File |
        Where-Object { -not $before.ContainsKey($_.FullName) -or $_.LastWriteTime -gt $before[$_.FullName] } |
        Sort-Object LastWriteTime -Descending

    if (-not $candidates) {
        Write-Host "No new file detected in $DownloadsFolder. Re-run with -File <path> once you have it." -ForegroundColor Red
        return
    }
    $File = $candidates[0].FullName
    Write-Host "Detected downloaded file: $File" -ForegroundColor Green
}

if (-not (Test-Path $File)) {
    throw "File not found: $File"
}

Write-Host "`n=== USB transfer ===" -ForegroundColor Cyan
$drive = $null
while (-not $drive) {
    $drives = @(Get-RemovableDrives)
    if ($drives.Count -eq 0) {
        Write-Host "No removable USB drive detected." -ForegroundColor Red
        $retry = Read-Host "Insert one and press Enter to re-scan (or type 'c' to cancel)"
        if ($retry -match '^[Cc]') { return }
        continue
    }

    Write-Host "Removable drives found:"
    for ($i = 0; $i -lt $drives.Count; $i++) {
        $d = $drives[$i]
        Write-Host "  [$i] $($d.DriveLetter): $($d.FileSystemLabel) ($([math]::Round($d.SizeRemaining/1GB,1)) GB free)"
    }

    $choice = Read-Host "`nExtract '$([IO.Path]::GetFileName($File))' to which drive? (index, 'r' to re-scan, or blank to cancel)"
    if ([string]::IsNullOrWhiteSpace($choice)) {
        Write-Host "Cancelled."
        return
    }
    if ($choice -match '^[Rr]') { continue }

    if ($choice -notmatch '^\d+$' -or [int]$choice -ge $drives.Count) {
        Write-Host "Not a valid drive index." -ForegroundColor Red
        continue
    }
    $drive = $drives[[int]$choice]
}

$targetRoot = "$($drive.DriveLetter):\"

$confirm = Read-Host "This will copy/extract to the root of drive $($drive.DriveLetter): - proceed? (y/N)"
if ($confirm -notmatch '^[Yy]') {
    Write-Host "Cancelled."
    return
}

if ($File -match '\.zip$') {
    Expand-Archive -Path $File -DestinationPath $targetRoot -Force
    Write-Host "Extracted zip contents to the root of $($drive.DriveLetter): - some BIOS flash tools don't scan subfolders." -ForegroundColor Green
} else {
    Copy-Item -Path $File -Destination $targetRoot -Force
    Write-Host "Copied $([IO.Path]::GetFileName($File)) to the root of $($drive.DriveLetter):" -ForegroundColor Green
    Write-Host "Note: many BIOS updates ship as self-extracting/flash executables meant to be run as-is (e.g. from BIOS flashback) rather than unzipped." -ForegroundColor Yellow
}

Write-Host "`nDone. Safely eject drive $($drive.DriveLetter): before removing it." -ForegroundColor Cyan
