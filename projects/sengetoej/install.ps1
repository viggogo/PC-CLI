# Creates this tool's virtualenv, installs it, and puts `sengetoej` on your PATH.
# No administrator rights. Safe to re-run.
#
# Only bin\ goes on PATH, never the project folder itself. Two reasons:
#   1. PowerShell searches .ps1 ahead of PATHEXT, so a sengetoej.ps1 sitting
#      next to sengetoej.cmd on PATH would win, then die on the execution policy.
#   2. The project folder also holds install.ps1 -- a generic name that would
#      become a global command in every terminal.
#
# ASCII only on purpose: PowerShell 5.1 decodes a BOM-less script as ANSI, so a
# stray non-ASCII character here would come out mangled.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root    = $PSScriptRoot
$venv    = Join-Path $root '.venv'
$venvPy  = Join-Path $venv 'Scripts\python.exe'
$binDir  = Join-Path $root 'bin'
$envFile = Join-Path $root '.env'
$example = Join-Path $root '.env.example'

# --- 1. virtualenv -----------------------------------------------------------
if (Test-Path $venvPy) {
    Write-Host 'Using the existing .venv'
} else {
    Write-Host 'Creating .venv ...'
    python -m venv $venv
    if (-not (Test-Path $venvPy)) {
        throw "venv creation failed -- no interpreter at $venvPy"
    }
}

# --- 2. dependencies ---------------------------------------------------------
# Editable install, so edits to sengetoej\ take effect without reinstalling.
Write-Host 'Installing dependencies (this needs network) ...'
& $venvPy -m pip install --upgrade pip --quiet
if ($LASTEXITCODE -ne 0) { throw 'pip self-upgrade failed' }
& $venvPy -m pip install --editable "$root[dev]" --quiet
if ($LASTEXITCODE -ne 0) { throw 'pip install failed' }

$exe = Join-Path $venv 'Scripts\sengetoej.exe'
if (-not (Test-Path $exe)) { throw "install finished but $exe is missing" }

# --- 3. .env -----------------------------------------------------------------
# Optional here: both settings have working defaults in code.
if (Test-Path $envFile) {
    Write-Host 'Found an existing .env -- left untouched.'
} else {
    Copy-Item $example $envFile
    Write-Host 'Created .env from .env.example.'
}

# --- 4. PATH -----------------------------------------------------------------
# Compare PATH entries case-insensitively, ignoring a trailing backslash.
function Get-NormalizedPathEntry {
    param([string]$Entry)
    return $Entry.TrimEnd('\').ToLowerInvariant()
}

$binKey  = Get-NormalizedPathEntry $binDir
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

# --- done --------------------------------------------------------------------
Write-Host ''
Write-Host 'ONE THING LEFT: run the one-time migration, which adds the "cli"'
Write-Host 'column and moves the comments one column right. Close Excel first.'
Write-Host "  $venvPy -m sengetoej.migrate"
Write-Host ''
Write-Host 'Then open a NEW terminal and run:  sengetoej --last'
