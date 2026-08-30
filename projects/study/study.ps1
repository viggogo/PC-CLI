<#
.SYNOPSIS
    Open the Literature study repo in VS Code.
.DESCRIPTION
    `study --begin` launches VS Code on the Literature repo. The path defaults to
    the constant below and can be overridden with STUDY_REPO in this folder's .env.

    The .env lives HERE, not in the Literature repo: a config file inside the
    target folder would be circular, since the path is what finds that folder.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:DefaultRepo = 'C:\Users\viggo\Git Clone\Literature'
$script:EnvFile     = Join-Path $PSScriptRoot '.env'

# ---------------------------------------------------------------- pure helpers

function Get-StudyIntent {
    param([string[]]$Argv)

    if ($null -eq $Argv -or $Argv.Count -eq 0) { return @{ Kind = 'Help' } }

    $action = $null

    foreach ($raw in $Argv) {
        # Every option must be dash-prefixed; a bare "begin" is a usage error.
        if ($raw -notmatch '^--?') {
            return @{ Kind = 'Error'; Message = "Unknown option: $raw" }
        }
        $t = ($raw -replace '^--?', '').ToLowerInvariant()

        if ($t -eq 'help' -or $t -eq 'h' -or $t -eq '?') {
            return @{ Kind = 'Help' }
        }

        $next = $null
        if ($t -eq 'begin' -or $t -eq 'b') { $next = 'Begin' }
        elseif ($t -eq 'where' -or $t -eq 'w') { $next = 'Where' }
        else {
            return @{ Kind = 'Error'; Message = "Unknown option: $raw" }
        }

        if ($null -ne $action) {
            return @{ Kind    = 'Error'
                      Message = "Specify only one action (got --$($action.ToLowerInvariant()) and --$t)." }
        }
        $action = $next
    }

    return @{ Kind = $action }
}

function Read-DotEnv {
    param([string]$Path)

    $result = @{}
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $result }

    foreach ($line in (Get-Content -LiteralPath $Path)) {
        $trimmed = $line.Trim()
        if ($trimmed -eq '' -or $trimmed.StartsWith('#')) { continue }

        # First '=' only. Windows paths hold no '=', but splitting on every one
        # would quietly truncate any value that did.
        $i = $trimmed.IndexOf('=')
        if ($i -lt 1) { continue }   # -1 is no '=' at all, 0 is a missing key

        $key   = $trimmed.Substring(0, $i).Trim()
        $value = $trimmed.Substring($i + 1).Trim()

        # Strip one matched pair of surrounding quotes. Without this, a quoted
        # path fails as "not found" with the quotes buried in the message.
        if ($value.Length -ge 2 -and
            (($value.StartsWith('"') -and $value.EndsWith('"')) -or
             ($value.StartsWith("'") -and $value.EndsWith("'")))) {
            $value = $value.Substring(1, $value.Length - 2)
        }

        if ($key -ne '') { $result[$key] = $value }
    }

    return $result
}

function Test-HasRepoOverride {
    param([hashtable]$DotEnv)

    if ($null -eq $DotEnv -or -not $DotEnv.ContainsKey('STUDY_REPO')) { return $false }
    return (-not [string]::IsNullOrWhiteSpace($DotEnv['STUDY_REPO']))
}

function Resolve-RepoPath {
    param([hashtable]$DotEnv)

    if (Test-HasRepoOverride $DotEnv) { return $DotEnv['STUDY_REPO'] }
    return $script:DefaultRepo
}

function Get-PathSource {
    param([hashtable]$DotEnv)

    if (Test-HasRepoOverride $DotEnv) { return '.env' }
    return 'default in study.ps1'
}

# ------------------------------------------------------------------------ cli

# Returns the usage text rather than printing it. This matters: in PowerShell a
# function returns EVERYTHING written to the output stream, so a Write-Output here
# would become part of Invoke-Main's return value and corrupt the exit code.
function Get-UsageText {
    return @'
study - open the Literature repo in VS Code

USAGE
  study --begin       Open the Literature repo in VS Code
  study --where       Show which path the tool resolved, and whether it exists
  study --help        Show this help

The path defaults to a constant in study.ps1. Override it by setting
STUDY_REPO in this tool's own .env (see .env.example).
'@
}

function Invoke-Main {
    param([string[]]$Argv)

    $intent = Get-StudyIntent $Argv

    switch ($intent.Kind) {
        'Help' {
            [Console]::Out.WriteLine((Get-UsageText))
            return 0
        }
        'Error' {
            [Console]::Error.WriteLine("study: $($intent.Message)")
            [Console]::Error.WriteLine('')
            [Console]::Error.WriteLine((Get-UsageText))
            return 2
        }
        'Where' {
            $dotEnv = Read-DotEnv $script:EnvFile
            $path   = Resolve-RepoPath $dotEnv
            $source = Get-PathSource $dotEnv
            $exists = Test-Path -LiteralPath $path -PathType Container
            # [Console]::Out, not Write-Output: anything on the output stream
            # becomes part of this function's return value and breaks the exit code.
            [Console]::Out.WriteLine("Repo path  $path")
            [Console]::Out.WriteLine("Source     $source")
            [Console]::Out.WriteLine("Exists     $exists")
            if (-not $exists) { return 1 }
            return 0
        }
        'Begin' {
            $path = Resolve-RepoPath (Read-DotEnv $script:EnvFile)

            # Check the folder BEFORE launching. `code` on a missing path opens an
            # empty window that looks like success, so a typo'd .env would fail
            # invisibly without this guard.
            if (-not (Test-Path -LiteralPath $path -PathType Container)) {
                [Console]::Error.WriteLine("study: no folder at $path")
                [Console]::Error.WriteLine('study: set STUDY_REPO in this tool''s .env, or run study --where.')
                return 1
            }

            if ($null -eq (Get-Command 'code' -ErrorAction SilentlyContinue)) {
                [Console]::Error.WriteLine('study: the VS Code CLI (code) is not on your PATH.')
                [Console]::Error.WriteLine('study: in VS Code run: Shell Command: Install ''code'' command in PATH')
                return 1
            }

            # `code` hands the folder to any running instance and returns at once.
            & code $path
            if ($LASTEXITCODE -ne 0) {
                [Console]::Error.WriteLine("study: code exited $LASTEXITCODE")
                return 1
            }

            [Console]::Out.WriteLine("Opened in VS Code: $path")
            return 0
        }
    }
    return 1
}

# Dot-source guard: when tests dot-source this file ($MyInvocation.InvocationName
# is '.'), define the functions but do not run. Only run main when executed.
if ($MyInvocation.InvocationName -ne '.') {
    try {
        exit (Invoke-Main $args)
    } catch {
        [Console]::Error.WriteLine("study: $($_.Exception.Message)")
        exit 1
    }
}
