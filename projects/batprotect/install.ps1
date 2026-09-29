# Adds this project's bin\ folder to the CURRENT USER's PATH. No administrator rights.
#
# Only bin\ goes on PATH, never the project folder itself. Two reasons:
#   1. PowerShell searches .ps1 ahead of PATHEXT, so a batprotect.ps1 sitting next to
#      batprotect.cmd on PATH would win, and then die on the machine's execution policy.
#   2. The project folder also holds test.ps1 and install.ps1 -- generic names that
#      would become global commands on every terminal.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$binDir = Join-Path $PSScriptRoot 'bin'

# Compare PATH entries case-insensitively, ignoring a trailing backslash.
function Get-NormalizedPathEntry {
    param([string]$Entry)
    return $Entry.TrimEnd('\').ToLowerInvariant()
}

$binKey = Get-NormalizedPathEntry $binDir

$current = [Environment]::GetEnvironmentVariable('Path', 'User')
if ($null -eq $current) { $current = '' }

# @() keeps this an array even when PATH holds a single entry.
$parts = @($current -split ';' | Where-Object { $_.Trim() -ne '' })

$normalized = @($parts | ForEach-Object { Get-NormalizedPathEntry $_ })
if ($normalized -contains $binKey) {
    Write-Host "Already on your PATH:`n  $binDir"
} else {
    [Environment]::SetEnvironmentVariable('Path', (@($parts + $binDir) -join ';'), 'User')
    Write-Host "Added to your user PATH:`n  $binDir"
}

Write-Host ''
Write-Host 'Open a NEW terminal, then run:  batprotect --status'
