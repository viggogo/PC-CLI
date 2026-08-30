# Adds this project's bin\ folder to the CURRENT USER's PATH and creates .env from
# .env.example. No administrator rights.
#
# Only bin\ goes on PATH, never the project folder itself. Two reasons:
#   1. PowerShell searches .ps1 ahead of PATHEXT, so a study.ps1 sitting next to
#      study.cmd on PATH would win, and then die on the machine's execution policy.
#   2. The project folder also holds test.ps1 and install.ps1 -- generic names that
#      would become global commands in every terminal.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$binDir  = Join-Path $PSScriptRoot 'bin'
$envFile = Join-Path $PSScriptRoot '.env'
$example = Join-Path $PSScriptRoot '.env.example'

# Compare PATH entries case-insensitively, ignoring a trailing backslash.
function Get-NormalizedPathEntry {
    param([string]$Entry)
    return $Entry.TrimEnd('\').ToLowerInvariant()
}

# --- 1. PATH -----------------------------------------------------------------
$binKey = Get-NormalizedPathEntry $binDir

$current = [Environment]::GetEnvironmentVariable('Path', 'User')
if ($null -eq $current) { $current = '' }

# @() keeps this an array even when PATH holds a single entry.
$parts = @($current -split ';' | Where-Object { $_.Trim() -ne '' })

$normalized = @($parts | ForEach-Object { Get-NormalizedPathEntry $_ })
$alreadyPresent = $normalized -contains $binKey

if ($alreadyPresent) {
    Write-Host "Already on your PATH:`n  $binDir"
} else {
    [Environment]::SetEnvironmentVariable('Path', (@($parts + $binDir) -join ';'), 'User')
    Write-Host "Added to your user PATH:`n  $binDir"
}

# --- 2. .env -----------------------------------------------------------------
# Never overwrite an existing .env -- it holds your own path override.
if (Test-Path -LiteralPath $envFile) {
    Write-Host 'Found an existing .env -- left untouched.'
} else {
    Copy-Item -LiteralPath $example -Destination $envFile
    Write-Host 'Created .env from .env.example.'
}

Write-Host ''
# A new TAB is not enough: terminals inherit their environment from the host
# process, so a tab opened in an already-running VS Code or Windows Terminal
# still carries the old PATH and reports "study is not recognized".
Write-Host 'Fully quit and relaunch your terminal app, then run:  study --begin'
Write-Host ''
Write-Host 'Or refresh this session without restarting anything:'
Write-Host '  $env:Path = [Environment]::GetEnvironmentVariable(''Path'',''Machine'') + '';'' + [Environment]::GetEnvironmentVariable(''Path'',''User'')'
