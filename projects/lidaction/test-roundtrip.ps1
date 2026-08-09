# End-to-end round trip for lidaction.
# WARNING: this CHANGES the real lid-close setting, then restores it.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'lidaction.ps1')

$pass = 0
$fail = 0

function Check {
    param($Expected, $Actual, [string]$Name)
    if ($Expected -eq $Actual) {
        $script:pass++
        Write-Host "  PASS  $Name" -ForegroundColor Green
    } else {
        $script:fail++
        Write-Host "  FAIL  $Name (expected [$Expected], got [$Actual])" -ForegroundColor Red
    }
}

$scheme = Get-ActiveScheme
$orig   = Get-LidActionIndices $scheme.Guid

if ($orig.Source -eq 'unknown') {
    Write-Host 'Cannot read the current setting; refusing to run so nothing is lost.' -ForegroundColor Red
    exit 1
}

Write-Host "Plan: $($scheme.Name)"
Write-Host "Original: AC=$(Get-ActionName $orig.Ac) DC=$(Get-ActionName $orig.Dc)"
Write-Host 'Cycling through all four values...' -ForegroundColor Cyan

try {
    foreach ($n in 0, 1, 2, 3) {
        Set-LidActionIndices $n
        $now = Get-LidActionIndices $scheme.Guid
        Check $n $now.Ac "AC set to $n ($(Get-ActionName $n))"
        Check $n $now.Dc "DC set to $n ($(Get-ActionName $n))"
    }
} finally {
    # Always restore, even on Ctrl-C or a thrown error, so an interrupted run
    # never leaves the machine's lid behavior changed.
    Write-Host 'Restoring original setting...' -ForegroundColor Cyan
    Invoke-Powercfg @('/setacvalueindex', 'SCHEME_CURRENT', $script:SubButtons, $script:LidActionGuid, "$($orig.Ac)")
    Invoke-Powercfg @('/setdcvalueindex', 'SCHEME_CURRENT', $script:SubButtons, $script:LidActionGuid, "$($orig.Dc)")
    Invoke-Powercfg @('/setactive', 'SCHEME_CURRENT')
}

$final = Get-LidActionIndices $scheme.Guid
Check $orig.Ac $final.Ac 'AC restored to original'
Check $orig.Dc $final.Dc 'DC restored to original'

Write-Host "`n$pass passed, $fail failed`n"
if ($fail -gt 0) { exit 1 }
exit 0
