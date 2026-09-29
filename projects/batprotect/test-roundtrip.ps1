# End-to-end test of the write path.
#
# CHANGES THE REAL SETTING and opens/closes Samsung Settings: cycles Battery
# protection off/on and through two charge limits, then restores your original
# on/off, limit, and whether the Samsung Settings window was open -- in a finally
# block. A hard kill mid-run can still leave it changed; run `batprotect --status`.
#
# Don't touch the Samsung Settings window while this runs.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'batprotect.ps1')

$script:Pass = 0
$script:Fail = 0

function Assert-Equal {
    param($Expected, $Actual, [string]$Name)
    # Same-type check: plain -eq coerces, so `$true -eq 'x'` would pass.
    $sameType = ($null -eq $Expected -and $null -eq $Actual) -or
                ($null -ne $Expected -and $null -ne $Actual -and $Expected.GetType() -eq $Actual.GetType())
    if ($sameType -and $Expected -eq $Actual) {
        $script:Pass++
        Write-Host "  PASS  $Name" -ForegroundColor Green
    } else {
        $script:Fail++
        Write-Host "  FAIL  $Name" -ForegroundColor Red
        Write-Host "        expected: [$Expected]" -ForegroundColor Red
        Write-Host "        actual:   [$Actual]" -ForegroundColor Red
    }
}

# Independent of batprotect: is a Samsung Settings window open right now?
function Test-WindowOpen {
    return [bool](Get-Process SamsungSettingsHost -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero })
}

function Close-Window {
    Get-Process SamsungSettingsHost -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } |
        ForEach-Object { [void]$_.CloseMainWindow() }
    $deadline = (Get-Date).AddSeconds(10)
    while ((Test-WindowOpen) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 200 }
}

function Step {
    param([string[]]$Argv, [bool]$WantOn, [int]$WantLevel, [bool]$WantWindow)
    $label = "batprotect $($Argv -join ' ')"
    Write-Host "`n$label" -ForegroundColor Cyan
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Assert-Equal 0 (Invoke-Main $Argv) "$label exits 0"
    Write-Host "        ($($sw.ElapsedMilliseconds) ms)"
    # The registry is the witness: Samsung's service writes it after the app acts.
    $s = Get-BatState
    Assert-Equal $WantOn     $s.On            "On is $WantOn"
    Assert-Equal $WantLevel  $s.Level         "Level is $WantLevel"
    Assert-Equal $WantWindow (Test-WindowOpen) "window open afterwards: $WantWindow"
}

$orig       = Get-BatState
$origWindow = Test-WindowOpen
Write-Host "Original: On=$($orig.On) Level=$($orig.Level) WindowOpen=$origWindow"
if ($orig.Level -notin $script:ValidLevels) {
    throw "Original level '$($orig.Level)' is not one of $($script:ValidLevels -join ', '); refusing to run because it could not be restored exactly."
}

# Test levels that differ from the original, so every step is a real change.
$levelA = @(60, 80) | Where-Object { $_ -ne $orig.Level } | Select-Object -First 1
$levelB = @(50, 90) | Where-Object { $_ -ne $orig.Level } | Select-Object -First 1

try {
    # Window closed: batprotect must open Samsung Settings and close it again.
    Close-Window
    Step @('--off')                       $false $orig.Level $false
    Step @('--level', "$levelA")          $false $levelA     $false   # moved while off, stays off

    # Window already open: batprotect must leave it open.
    Start-SamsungSettings | Out-Null
    Step @('--on')                        $true  $levelA     $true
    Step @('--on')                        $true  $levelA     $true    # no-op
    Step @('--on', '--level', "$levelB")  $true  $levelB     $true
} finally {
    Write-Host "`nrestoring..." -ForegroundColor Cyan
    [void](Invoke-Main @('--on', '--level', "$($orig.Level)"))
    if (-not $orig.On) { [void](Invoke-Main @('--off')) }
    if ($origWindow -and -not (Test-WindowOpen)) { Start-SamsungSettings | Out-Null }
    if (-not $origWindow) { Close-Window }
}

$s = Get-BatState
Assert-Equal $orig.On    $s.On    'On restored'
Assert-Equal $orig.Level $s.Level 'Level restored'
Assert-Equal $origWindow (Test-WindowOpen) 'window state restored'

Write-Host "`n$($script:Pass) passed, $($script:Fail) failed`n"
if ($script:Fail -gt 0) { exit 1 }
exit 0
