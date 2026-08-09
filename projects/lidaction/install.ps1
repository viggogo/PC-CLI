# Adds this project's bin\ folder to the CURRENT USER's PATH. No administrator rights.
#
# Only bin\ goes on PATH, never the project folder itself. Two reasons:
#   1. PowerShell searches .ps1 ahead of PATHEXT, so a lidaction.ps1 sitting next to
#      lidaction.cmd on PATH would win, and then die on the machine's execution policy.
#   2. The project folder also holds test.ps1 and install.ps1 -- generic names that
#      would become global commands on every terminal.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$binDir = Join-Path $PSScriptRoot 'bin'
$oldDir = $PSScriptRoot

# Compare PATH entries case-insensitively, ignoring a trailing backslash.
function Get-NormalizedPathEntry {
    param([string]$Entry)
    return $Entry.TrimEnd('\').ToLowerInvariant()
}

$binKey = Get-NormalizedPathEntry $binDir
$oldKey = Get-NormalizedPathEntry $oldDir

$current = [Environment]::GetEnvironmentVariable('Path', 'User')
if ($null -eq $current) { $current = '' }

# @() keeps this an array even when PATH holds a single entry.
$parts = @($current -split ';' | Where-Object { $_.Trim() -ne '' })

# Migration: drop the pre-bin\ entry that pointed at the project folder. Leaving it
# would keep lidaction.ps1 shadowing the shim, so the fix would not take effect.
$kept = @($parts | Where-Object { (Get-NormalizedPathEntry $_) -ne $oldKey })
$removedOld = $kept.Count -lt $parts.Count

$normalized = @($kept | ForEach-Object { Get-NormalizedPathEntry $_ })
$alreadyPresent = $normalized -contains $binKey
if (-not $alreadyPresent) {
    $kept = @($kept + $binDir)
}

if ($removedOld -or (-not $alreadyPresent)) {
    [Environment]::SetEnvironmentVariable('Path', ($kept -join ';'), 'User')
}

if ($removedOld) {
    Write-Host "Removed the old entry from your user PATH:`n  $oldDir"
}
if ($alreadyPresent) {
    Write-Host "Already on your PATH:`n  $binDir"
} else {
    Write-Host "Added to your user PATH:`n  $binDir"
}

Write-Host ''
Write-Host 'Open a NEW terminal, then run:  lidaction --status'
