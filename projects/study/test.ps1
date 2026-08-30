# Tests for study. Entirely inert: every assertion here is pure logic -- argument
# parsing, .env parsing, path resolution, and the exit codes for help and usage
# errors. Nothing in this file opens a window or touches your PATH.
#
# `study --begin` is deliberately NOT covered: the only thing it does is launch
# VS Code, and a test for it would leave an editor window open with nothing to
# restore. Verify that one by hand -- see README.md.
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

Write-Host "`nGet-StudyIntent - help" -ForegroundColor Cyan
Assert-Equal 'Help' (Get-StudyIntent @()).Kind         'no args'
Assert-Equal 'Help' (Get-StudyIntent @('--help')).Kind 'double dash'
Assert-Equal 'Help' (Get-StudyIntent @('-help')).Kind  'single dash'
Assert-Equal 'Help' (Get-StudyIntent @('-h')).Kind     'short'
Assert-Equal 'Help' (Get-StudyIntent @('-?')).Kind     'question mark'
Assert-Equal 'Help' (Get-StudyIntent @('--HELP')).Kind 'case insensitive'

Write-Host "`nGet-StudyIntent - begin" -ForegroundColor Cyan
Assert-Equal 'Begin' (Get-StudyIntent @('--begin')).Kind 'double dash'
Assert-Equal 'Begin' (Get-StudyIntent @('-begin')).Kind  'single dash'
Assert-Equal 'Begin' (Get-StudyIntent @('-b')).Kind      'short'
Assert-Equal 'Begin' (Get-StudyIntent @('--BEGIN')).Kind 'case insensitive'

Write-Host "`nGet-StudyIntent - where" -ForegroundColor Cyan
Assert-Equal 'Where' (Get-StudyIntent @('--where')).Kind 'double dash'
Assert-Equal 'Where' (Get-StudyIntent @('-w')).Kind      'short'

Write-Host "`nGet-StudyIntent - errors" -ForegroundColor Cyan
Assert-Equal 'Error' (Get-StudyIntent @('--bogus')).Kind          'unknown flag'
Assert-Equal 'Error' (Get-StudyIntent @('begin')).Kind            'bare word needs a dash'
Assert-Equal 'Error' (Get-StudyIntent @('--begin','--where')).Kind 'two actions'
Assert-Equal 'Error' (Get-StudyIntent @('--begin','--bogus')).Kind 'good flag then bad'

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

    $plain = New-Fixture 'plain.env' @('STUDY_REPO=C:\Repos\Literature')
    Assert-Equal 'C:\Repos\Literature' (Read-DotEnv $plain)['STUDY_REPO'] 'reads a simple value'

    $noisy = New-Fixture 'noisy.env' @(
        '# a comment'
        ''
        '   '
        '   # indented comment'
        'STUDY_REPO=C:\Repos\Literature'
    )
    $noisyEnv = Read-DotEnv $noisy
    Assert-Equal 1 $noisyEnv.Count 'comments and blank lines are skipped'
    Assert-Equal 'C:\Repos\Literature' $noisyEnv['STUDY_REPO'] 'value survives the noise'

    $spaced = New-Fixture 'spaced.env' @('  STUDY_REPO  =   C:\Repos\Literature   ')
    Assert-Equal 'C:\Repos\Literature' (Read-DotEnv $spaced)['STUDY_REPO'] 'trims key and value'

    $quoted = New-Fixture 'quoted.env' @('STUDY_REPO="C:\Repos\Literature"')
    Assert-Equal 'C:\Repos\Literature' (Read-DotEnv $quoted)['STUDY_REPO'] 'strips surrounding quotes'

    # Split on the FIRST '=' only, so a value containing one stays intact.
    $equals = New-Fixture 'equals.env' @('STUDY_REPO=C:\Repos\a=b')
    Assert-Equal 'C:\Repos\a=b' (Read-DotEnv $equals)['STUDY_REPO'] 'splits on the first = only'

    $malformed = New-Fixture 'malformed.env' @(
        'this line has no equals sign'
        '=value with no key'
        'STUDY_REPO=C:\Repos\Literature'
    )
    $malformedEnv = Read-DotEnv $malformed
    Assert-Equal 1 $malformedEnv.Count 'malformed lines are ignored'
    Assert-Equal 'C:\Repos\Literature' $malformedEnv['STUDY_REPO'] 'good line still read'

    $empty = New-Fixture 'empty-value.env' @('STUDY_REPO=')
    Assert-Equal '' (Read-DotEnv $empty)['STUDY_REPO'] 'empty value reads as empty string'

    Write-Host "`nResolve-RepoPath" -ForegroundColor Cyan
    Assert-Equal 'C:\Repos\Literature' `
        (Resolve-RepoPath @{ 'STUDY_REPO' = 'C:\Repos\Literature' }) '.env override wins'
    Assert-Equal $script:DefaultRepo (Resolve-RepoPath @{}) 'no key falls back to the default'
    Assert-Equal $script:DefaultRepo `
        (Resolve-RepoPath @{ 'STUDY_REPO' = '' }) 'empty value falls back to the default'
    Assert-Equal $script:DefaultRepo `
        (Resolve-RepoPath @{ 'STUDY_REPO' = '   ' }) 'whitespace value falls back to the default'

    # --where is only worth having if it names the source that actually won.
    Write-Host "`nGet-PathSource" -ForegroundColor Cyan
    Assert-Equal '.env' (Get-PathSource @{ 'STUDY_REPO' = 'C:\Repos\Literature' }) 'override reports .env'
    Assert-Equal 'default in study.ps1' (Get-PathSource @{}) 'no key reports the default'
    Assert-Equal 'default in study.ps1' (Get-PathSource @{ 'STUDY_REPO' = '' }) 'empty value reports the default'
}
finally {
    Remove-Item -Path $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "`nInvoke-Main exit codes" -ForegroundColor Cyan
Assert-Equal 0 (Invoke-Main @('--help'))  'help exits 0'
Assert-Equal 0 (Invoke-Main @())          'no args exits 0'
Assert-Equal 2 (Invoke-Main @('--bogus')) 'unknown flag exits 2'
Assert-Equal 2 (Invoke-Main @('--begin','--where')) 'two actions exits 2'

Write-Host "`nGet-UsageText" -ForegroundColor Cyan
$usage = Get-UsageText
Assert-Equal $true ($usage -like '*--begin*') 'usage mentions --begin'
Assert-Equal $true ($usage -like '*--where*') 'usage mentions --where'
Assert-Equal $true ($usage -like '*--help*')  'usage mentions --help'

Write-Host ''
if ($script:Fail -gt 0) {
    Write-Host "$($script:Pass) passed, $($script:Fail) FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "$($script:Pass) passed, 0 failed" -ForegroundColor Green
exit 0
