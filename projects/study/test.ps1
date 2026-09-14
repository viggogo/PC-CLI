# Tests for study. Entirely inert: every assertion here is pure logic -- argument
# parsing, .env parsing, path resolution, and the exit codes for help and usage
# errors. Nothing in this file opens a window or touches your PATH.
#
# `study --read` and `study --write` are deliberately NOT covered: the only thing
# they do is launch VS Code, and a test for either would leave an editor window
# open with nothing to restore. Verify those by hand -- see README.md.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'study.ps1')

$script:Pass = 0
$script:Fail = 0

function Assert-Equal {
    param($Expected, $Actual, [string]$Name)
    if ($Expected -eq $Actual) {
        $script:Pass++
        Write-Host "  PASS  $Name" -ForegroundColor Green
    } else {
        $script:Fail++
        Write-Host "  FAIL  $Name" -ForegroundColor Red
        Write-Host "        expected: [$Expected]" -ForegroundColor Red
        Write-Host "        actual:   [$Actual]" -ForegroundColor Red
    }
}

$readSpec  = Get-RepoSpec 'Read'
$writeSpec = Get-RepoSpec 'Write'

Write-Host "`nGet-StudyIntent - help" -ForegroundColor Cyan
Assert-Equal 'Help' (Get-StudyIntent @()).Kind         'no args'
Assert-Equal 'Help' (Get-StudyIntent @('--help')).Kind 'double dash'
Assert-Equal 'Help' (Get-StudyIntent @('-help')).Kind  'single dash'
Assert-Equal 'Help' (Get-StudyIntent @('-h')).Kind     'short'
Assert-Equal 'Help' (Get-StudyIntent @('-?')).Kind     'question mark'
Assert-Equal 'Help' (Get-StudyIntent @('--HELP')).Kind 'case insensitive'

Write-Host "`nGet-StudyIntent - read" -ForegroundColor Cyan
Assert-Equal 'Read' (Get-StudyIntent @('--read')).Kind 'double dash'
Assert-Equal 'Read' (Get-StudyIntent @('-read')).Kind  'single dash'
Assert-Equal 'Read' (Get-StudyIntent @('-r')).Kind     'short'
Assert-Equal 'Read' (Get-StudyIntent @('--READ')).Kind 'case insensitive'

Write-Host "`nGet-StudyIntent - write" -ForegroundColor Cyan
Assert-Equal 'Write' (Get-StudyIntent @('--write')).Kind 'double dash'
Assert-Equal 'Write' (Get-StudyIntent @('-write')).Kind  'single dash'
Assert-Equal 'Write' (Get-StudyIntent @('--WRITE')).Kind 'case insensitive'

# -w must keep meaning --where. Handing it to --write would silently turn a
# "print the paths" habit into "open an editor".
Write-Host "`nGet-StudyIntent - where" -ForegroundColor Cyan
Assert-Equal 'Where' (Get-StudyIntent @('--where')).Kind 'double dash'
Assert-Equal 'Where' (Get-StudyIntent @('-w')).Kind      '-w is where, not write'

Write-Host "`nGet-StudyIntent - errors" -ForegroundColor Cyan
Assert-Equal 'Error' (Get-StudyIntent @('--bogus')).Kind           'unknown flag'
Assert-Equal 'Error' (Get-StudyIntent @('--begin')).Kind           'the old --begin is gone'
Assert-Equal 'Error' (Get-StudyIntent @('read')).Kind              'bare word needs a dash'
Assert-Equal 'Error' (Get-StudyIntent @('--read','--where')).Kind  'two actions'
Assert-Equal 'Error' (Get-StudyIntent @('--read','--write')).Kind  'read and write are exclusive'
Assert-Equal 'Error' (Get-StudyIntent @('--read','--bogus')).Kind  'good flag then bad'

Write-Host "`nGet-RepoSpec" -ForegroundColor Cyan
Assert-Equal 'STUDY_READ_REPO'  $readSpec.EnvKey   'read spec env key'
Assert-Equal 'STUDY_WRITE_REPO' $writeSpec.EnvKey  'write spec env key'
Assert-Equal $script:DefaultReadRepo  $readSpec.Default  'read spec default'
Assert-Equal $script:DefaultWriteRepo $writeSpec.Default 'write spec default'
Assert-Equal $true ($readSpec.Default -ne $writeSpec.Default) 'the two repos differ'

# .env fixtures go to a temp folder, never into the project -- writing them here
# would sit next to the real .env and risk clobbering it.
$fixtureDir = Join-Path ([System.IO.Path]::GetTempPath()) "study-tests-$PID"
New-Item -ItemType Directory -Path $fixtureDir -Force | Out-Null

function New-Fixture {
    param([string]$Name, [string[]]$Lines)
    $path = Join-Path $fixtureDir $Name
    Set-Content -Path $path -Value $Lines -Encoding UTF8
    return $path
}

try {
    Write-Host "`nRead-DotEnv" -ForegroundColor Cyan

    $missing = Join-Path $fixtureDir 'does-not-exist.env'
    Assert-Equal 0 (Read-DotEnv $missing).Count 'missing file yields no keys'

    $plain = New-Fixture 'plain.env' @('STUDY_READ_REPO=C:\Repos\Literature')
    Assert-Equal 'C:\Repos\Literature' (Read-DotEnv $plain)['STUDY_READ_REPO'] 'reads a simple value'

    $both = New-Fixture 'both.env' @(
        'STUDY_READ_REPO=C:\Repos\Literature'
        'STUDY_WRITE_REPO=C:\Repos\latex'
    )
    $bothEnv = Read-DotEnv $both
    Assert-Equal 2 $bothEnv.Count 'both keys are read'
    Assert-Equal 'C:\Repos\latex' $bothEnv['STUDY_WRITE_REPO'] 'write key survives alongside read'

    $noisy = New-Fixture 'noisy.env' @(
        '# a comment'
        ''
        '   '
        '   # indented comment'
        'STUDY_READ_REPO=C:\Repos\Literature'
    )
    $noisyEnv = Read-DotEnv $noisy
    Assert-Equal 1 $noisyEnv.Count 'comments and blank lines are skipped'
    Assert-Equal 'C:\Repos\Literature' $noisyEnv['STUDY_READ_REPO'] 'value survives the noise'

    $spaced = New-Fixture 'spaced.env' @('  STUDY_READ_REPO  =   C:\Repos\Literature   ')
    Assert-Equal 'C:\Repos\Literature' (Read-DotEnv $spaced)['STUDY_READ_REPO'] 'trims key and value'

    $quoted = New-Fixture 'quoted.env' @('STUDY_READ_REPO="C:\Repos\Literature"')
    Assert-Equal 'C:\Repos\Literature' (Read-DotEnv $quoted)['STUDY_READ_REPO'] 'strips surrounding quotes'

    # Split on the FIRST '=' only, so a value containing one stays intact.
    $equals = New-Fixture 'equals.env' @('STUDY_READ_REPO=C:\Repos\a=b')
    Assert-Equal 'C:\Repos\a=b' (Read-DotEnv $equals)['STUDY_READ_REPO'] 'splits on the first = only'

    $malformed = New-Fixture 'malformed.env' @(
        'this line has no equals sign'
        '=value with no key'
        'STUDY_READ_REPO=C:\Repos\Literature'
    )
    $malformedEnv = Read-DotEnv $malformed
    Assert-Equal 1 $malformedEnv.Count 'malformed lines are ignored'
    Assert-Equal 'C:\Repos\Literature' $malformedEnv['STUDY_READ_REPO'] 'good line still read'

    $empty = New-Fixture 'empty-value.env' @('STUDY_READ_REPO=')
    Assert-Equal '' (Read-DotEnv $empty)['STUDY_READ_REPO'] 'empty value reads as empty string'

    Write-Host "`nResolve-RepoPath" -ForegroundColor Cyan
    Assert-Equal 'C:\Repos\Literature' `
        (Resolve-RepoPath @{ 'STUDY_READ_REPO' = 'C:\Repos\Literature' } $readSpec) 'read override wins'
    Assert-Equal 'C:\Repos\latex' `
        (Resolve-RepoPath @{ 'STUDY_WRITE_REPO' = 'C:\Repos\latex' } $writeSpec) 'write override wins'
    Assert-Equal $readSpec.Default  (Resolve-RepoPath @{} $readSpec)  'no key falls back to the read default'
    Assert-Equal $writeSpec.Default (Resolve-RepoPath @{} $writeSpec) 'no key falls back to the write default'
    Assert-Equal $readSpec.Default `
        (Resolve-RepoPath @{ 'STUDY_READ_REPO' = '' } $readSpec) 'empty value falls back to the default'
    Assert-Equal $readSpec.Default `
        (Resolve-RepoPath @{ 'STUDY_READ_REPO' = '   ' } $readSpec) 'whitespace value falls back to the default'

    # Each action reads only its own key, so one override never moves the other repo.
    Assert-Equal $writeSpec.Default `
        (Resolve-RepoPath @{ 'STUDY_READ_REPO' = 'C:\Repos\Literature' } $writeSpec) 'read override leaves write alone'
    Assert-Equal $readSpec.Default `
        (Resolve-RepoPath @{ 'STUDY_WRITE_REPO' = 'C:\Repos\latex' } $readSpec) 'write override leaves read alone'

    # The old single-key name must no longer steer anything.
    Assert-Equal $readSpec.Default `
        (Resolve-RepoPath @{ 'STUDY_REPO' = 'C:\Repos\Old' } $readSpec) 'legacy STUDY_REPO is ignored'

    # --where is only worth having if it names the source that actually won.
    Write-Host "`nGet-PathSource" -ForegroundColor Cyan
    Assert-Equal '.env (STUDY_READ_REPO)' `
        (Get-PathSource @{ 'STUDY_READ_REPO' = 'C:\Repos\Literature' } $readSpec) 'override reports .env and the key'
    Assert-Equal '.env (STUDY_WRITE_REPO)' `
        (Get-PathSource @{ 'STUDY_WRITE_REPO' = 'C:\Repos\latex' } $writeSpec) 'write override reports its own key'
    Assert-Equal 'default in study.ps1' (Get-PathSource @{} $readSpec) 'no key reports the default'
    Assert-Equal 'default in study.ps1' `
        (Get-PathSource @{ 'STUDY_READ_REPO' = '' } $readSpec) 'empty value reports the default'
}
finally {
    Remove-Item -Path $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`nInvoke-Main exit codes" -ForegroundColor Cyan
Assert-Equal 0 (Invoke-Main @('--help'))  'help exits 0'
Assert-Equal 0 (Invoke-Main @())          'no args exits 0'
Assert-Equal 2 (Invoke-Main @('--bogus')) 'unknown flag exits 2'
Assert-Equal 2 (Invoke-Main @('--begin')) 'the retired --begin exits 2'
Assert-Equal 2 (Invoke-Main @('--read','--where')) 'two actions exits 2'

Write-Host "`nGet-UsageText" -ForegroundColor Cyan
$usage = Get-UsageText
Assert-Equal $true ($usage -like '*--read*')  'usage mentions --read'
Assert-Equal $true ($usage -like '*--write*') 'usage mentions --write'
Assert-Equal $true ($usage -like '*--where*') 'usage mentions --where'
Assert-Equal $true ($usage -like '*--help*')  'usage mentions --help'
Assert-Equal $false ($usage -like '*--begin*') 'usage no longer mentions --begin'

Write-Host ''
if ($script:Fail -gt 0) {
    Write-Host "$($script:Pass) passed, $($script:Fail) FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "$($script:Pass) passed, 0 failed" -ForegroundColor Green
exit 0
