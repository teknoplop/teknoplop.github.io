<#
.SYNOPSIS
One-shot setup for a clean Windows PC. Reads and executes the command list from
the <code id="commands"> block on https://teknoplop.github.io, so that page is
the single source of truth - there is no separate hardcoded list here.

.DESCRIPTION
Re-launches itself elevated if not already running as Administrator, since global
scoop installs and clink's all-users autorun both require it.

# EXECUTION FROM GITHUB (Single Command Line)
# iex (irm "https://teknoplop.github.io/Setup.ps1")
#>

[CmdletBinding()]
param()

$IsElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if (-not $IsElevated) {
    Write-Host "Re-launching elevated..."
    $SetupUrl = "https://teknoplop.github.io/Setup.ps1"
    Start-Process powershell -Verb RunAs -ArgumentList @(
        '-NoExit', '-NoProfile', '-Command',
        "iex (irm '$SetupUrl')"
    )
    return
}

$PageUrl = "https://teknoplop.github.io/index.html"
Write-Host "Fetching command list from $PageUrl ..."
$Html = Invoke-RestMethod -Uri $PageUrl

if ($Html -notmatch '(?s)<code id="commands">(.*?)</code>') {
    throw "Could not find <code id=`"commands`"> block in $PageUrl"
}
$Block = $Matches[1]
$Block = $Block -replace '<br\s*/?>', "`n"
$Block = $Block -replace '<[^>]+>', ''
$Block = [System.Net.WebUtility]::HtmlDecode($Block)

$Commands = $Block -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notlike '#*' }

Write-Host "Running $($Commands.Count) commands...`n"
foreach ($Command in $Commands) {
    Write-Host "> $Command" -ForegroundColor Cyan
    Invoke-Expression $Command
}

Write-Host "`nSetup complete. 🚀"
