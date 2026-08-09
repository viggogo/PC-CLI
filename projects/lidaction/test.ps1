# Tests for lidaction. Mostly pure logic, but NOT inert: the last section runs
# Invoke-Main @('--1') against the live system to cover the set path, then restores
# your original AC and DC values in a finally block.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'lidaction.ps1')

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

Write-Host "`nGet-ActionName" -ForegroundColor Cyan
Assert-Equal 'Do nothing' (Get-ActionName 0) 'index 0'
Assert-Equal 'Sleep'      (Get-ActionName 1) 'index 1'
Assert-Equal 'Hibernate'  (Get-ActionName 2) 'index 2'
Assert-Equal 'Shut down'  (Get-ActionName 3) 'index 3'
Assert-Equal 'unknown'    (Get-ActionName 4) 'out of range high'
# Parenthesize -1 so PowerShell does not try to bind it as a parameter name.
Assert-Equal 'unknown'    (Get-ActionName (-1)) 'out of range low'
Assert-Equal 'unknown'    (Get-ActionName $null) 'null'

Write-Host "`nGet-LidIntent - help" -ForegroundColor Cyan
Assert-Equal 'Help' (Get-LidIntent @()).Kind          'no args'
Assert-Equal 'Help' (Get-LidIntent @('--help')).Kind  '--help'
Assert-Equal 'Help' (Get-LidIntent @('-help')).Kind   '-help'
Assert-Equal 'Help' (Get-LidIntent @('-h')).Kind      '-h'
Assert-Equal 'Help' (Get-LidIntent @('-?')).Kind      '-?'
Assert-Equal 'Help' (Get-LidIntent @('--HELP')).Kind  'case insensitive'

Write-Host "`nGet-LidIntent - status" -ForegroundColor Cyan
Assert-Equal 'Status' (Get-LidIntent @('--status')).Kind 'double dash'
Assert-Equal 'Status' (Get-LidIntent @('-status')).Kind  'single dash'
Assert-Equal 'Status' (Get-LidIntent @('-s')).Kind       'short'
Assert-Equal 'Status' (Get-LidIntent @('--S')).Kind      'short uppercase'

Write-Host "`nGet-LidIntent - set" -ForegroundColor Cyan
Assert-Equal 'Set' (Get-LidIntent @('--0')).Kind  'kind for --0'
Assert-Equal 0     (Get-LidIntent @('--0')).Value 'value for --0'
Assert-Equal 1     (Get-LidIntent @('--1')).Value 'value for --1'
Assert-Equal 2     (Get-LidIntent @('--2')).Value 'value for --2'
Assert-Equal 3     (Get-LidIntent @('--3')).Value 'value for --3'
Assert-Equal 1     (Get-LidIntent @('-1')).Value  'single dash form'

Write-Host "`nGet-LidIntent - errors" -ForegroundColor Cyan
Assert-Equal 'Error' (Get-LidIntent @('--4')).Kind        'out of range action'
Assert-Equal 'Error' (Get-LidIntent @('--bogus')).Kind    'unknown flag'
Assert-Equal 'Error' (Get-LidIntent @('0')).Kind          'bare value needs a dash'
Assert-Equal 'Error' (Get-LidIntent @('--1','--2')).Kind  'two action flags'
Assert-Equal 'Error' (Get-LidIntent @('--1','--status')).Kind 'action plus status'

Write-Host "`nFormat-LidStatus" -ForegroundColor Cyan
$expected = @(
    'Lid close action  (plan: Balanced)'
    '  Plugged in   Sleep'
    '  On battery   Hibernate'
) -join [Environment]::NewLine
Assert-Equal $expected (Format-LidStatus 'Balanced' 1 2) 'formats both rows'

$expectedUnknown = @(
    'Lid close action  (plan: Balanced)'
    '  Plugged in   unknown'
    '  On battery   unknown'
) -join [Environment]::NewLine
Assert-Equal $expectedUnknown (Format-LidStatus 'Balanced' $null $null) 'formats unknown'

Write-Host "`nInvoke-Main exit codes" -ForegroundColor Cyan
Assert-Equal 0 (Invoke-Main @('--help')) 'help exits 0'
Assert-Equal 0 (Invoke-Main @())         'no args exits 0'
Assert-Equal 2 (Invoke-Main @('--bogus')) 'unknown flag exits 2'
Assert-Equal 2 (Invoke-Main @('--1','--2')) 'two actions exits 2'

Write-Host "`nGet-UsageText content" -ForegroundColor Cyan
$usage = Get-UsageText
Assert-Equal $true ($usage -like '*--status*')  'usage mentions --status'
Assert-Equal $true ($usage -like '*--0*')       'usage mentions --0'
Assert-Equal $true ($usage -like '*--3*')       'usage mentions --3'
Assert-Equal $true ($usage -like '*Hibernate*') 'usage mentions Hibernate'

Write-Host "`nGet-ActiveScheme (live)" -ForegroundColor Cyan
$scheme = Get-ActiveScheme
Assert-Equal $true ($scheme.Guid -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') 'guid looks like a guid'
Assert-Equal $true ($scheme.Name.Length -gt 0) 'plan has a name'

Write-Host "`nGet-LidActionIndices (live, read-only)" -ForegroundColor Cyan
$idx = Get-LidActionIndices $scheme.Guid
Assert-Equal $true ($idx.Source -in @('registry','powercfg','unknown')) 'source is a known value'
Assert-Equal $true ($null -eq $idx.Ac -or ($idx.Ac -ge 0 -and $idx.Ac -le 3)) 'AC in range or null'
Assert-Equal $true ($null -eq $idx.Dc -or ($idx.Dc -ge 0 -and $idx.Dc -le 3)) 'DC in range or null'

Write-Host "`n--status" -ForegroundColor Cyan
Assert-Equal 0 (Invoke-Main @('--status')) 'status exits 0'

Write-Host "`nTest-HibernateEnabled" -ForegroundColor Cyan
$hib = Test-HibernateEnabled
Assert-Equal $true ($hib -is [bool]) 'returns a boolean'

Write-Host "`nSet round trip (restores original)" -ForegroundColor Cyan
$scheme0 = Get-ActiveScheme
$orig    = Get-LidActionIndices $scheme0.Guid
try {
    Assert-Equal 0 (Invoke-Main @('--1')) 'set --1 exits 0'
    $after = Get-LidActionIndices $scheme0.Guid
    Assert-Equal 1 $after.Ac 'AC became Sleep'
    Assert-Equal 1 $after.Dc 'DC became Sleep'
} finally {
    # Restore AC and DC INDEPENDENTLY. Set-LidActionIndices writes the same value
    # to both, which would silently corrupt the restore when they differ.
    if ($null -ne $orig.Ac -and $null -ne $orig.Dc) {
        Invoke-Powercfg @('/setacvalueindex', 'SCHEME_CURRENT', $script:SubButtons, $script:LidActionGuid, "$($orig.Ac)")
        Invoke-Powercfg @('/setdcvalueindex', 'SCHEME_CURRENT', $script:SubButtons, $script:LidActionGuid, "$($orig.Dc)")
        Invoke-Powercfg @('/setactive', 'SCHEME_CURRENT')
    }
}
$restored = Get-LidActionIndices $scheme0.Guid
Assert-Equal $orig.Ac $restored.Ac 'AC restored'
Assert-Equal $orig.Dc $restored.Dc 'DC restored'

Write-Host "`n$($script:Pass) passed, $($script:Fail) failed`n"
if ($script:Fail -gt 0) { exit 1 }
exit 0
