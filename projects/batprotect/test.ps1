# Tests for batprotect. Pure logic plus READ-ONLY live reads of the registry.
# Nothing here changes the Battery protection setting or opens Samsung Settings --
# the write path lives in test-roundtrip.ps1.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'batprotect.ps1')

$script:Pass = 0
$script:Fail = 0

function Assert-Equal {
    param($Expected, $Actual, [string]$Name)
    # Require the same type too: plain -eq coerces the right side to the left side's
    # type, so `$true -eq 'any string'` is $true and would pass a broken test.
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

# Read a key from an intent without tripping StrictMode on a null intent.
function Get-Field {
    param($Intent, [string]$Key)
    if ($null -eq $Intent) { return '<null intent>' }
    return $Intent[$Key]
}

# Steps as one comparable string, e.g. 'on,level:80,off'. '' means "nothing to do".
function Get-StepString {
    param($Current, $Intent)
    return (@(Get-BatSteps $Current $Intent) -join ',')
}

Write-Host "`nGet-BatIntent - help" -ForegroundColor Cyan
Assert-Equal 'Help' (Get-Field (Get-BatIntent @()) 'Kind')                 'no args'
Assert-Equal 'Help' (Get-Field (Get-BatIntent @('--help')) 'Kind')         '--help'
Assert-Equal 'Help' (Get-Field (Get-BatIntent @('-h')) 'Kind')             '-h'
Assert-Equal 'Help' (Get-Field (Get-BatIntent @('-?')) 'Kind')             '-?'
Assert-Equal 'Help' (Get-Field (Get-BatIntent @('--HELP')) 'Kind')         'case insensitive'
Assert-Equal 'Help' (Get-Field (Get-BatIntent @('--on', '--help')) 'Kind') 'help wins over other flags'

Write-Host "`nGet-BatIntent - status" -ForegroundColor Cyan
Assert-Equal 'Status' (Get-Field (Get-BatIntent @('--status')) 'Kind') 'double dash'
Assert-Equal 'Status' (Get-Field (Get-BatIntent @('-status')) 'Kind')  'single dash'
Assert-Equal 'Status' (Get-Field (Get-BatIntent @('-s')) 'Kind')       'short'
Assert-Equal 'Status' (Get-Field (Get-BatIntent @('--STATUS')) 'Kind') 'case insensitive'

Write-Host "`nGet-BatIntent - on/off" -ForegroundColor Cyan
Assert-Equal 'Set'  (Get-Field (Get-BatIntent @('--on')) 'Kind')   '--on kind'
Assert-Equal $true  (Get-Field (Get-BatIntent @('--on')) 'On')     '--on sets On=true'
Assert-Equal $null  (Get-Field (Get-BatIntent @('--on')) 'Level')  '--on leaves Level unset'
Assert-Equal $false (Get-Field (Get-BatIntent @('--off')) 'On')    '--off sets On=false'
Assert-Equal $true  (Get-Field (Get-BatIntent @('-ON')) 'On')      'single dash, upper case'
Assert-Equal $true  (Get-Field (Get-BatIntent @('--on', '--on')) 'On') 'repeating --on is harmless'

Write-Host "`nGet-BatIntent - level" -ForegroundColor Cyan
Assert-Equal 'Set' (Get-Field (Get-BatIntent @('--level', '80')) 'Kind')  '--level kind'
Assert-Equal 80    (Get-Field (Get-BatIntent @('--level', '80')) 'Level') '--level 80'
Assert-Equal $null (Get-Field (Get-BatIntent @('--level', '80')) 'On')    '--level alone leaves On unset'
Assert-Equal 70    (Get-Field (Get-BatIntent @('--level', '70%')) 'Level') 'accepts a trailing %'
foreach ($n in 50, 60, 70, 80, 90) {
    Assert-Equal $n (Get-Field (Get-BatIntent @('--level', "$n")) 'Level') "accepts $n"
}
Assert-Equal $true (Get-Field (Get-BatIntent @('--on', '--level', '60')) 'On')    '--on --level: On'
Assert-Equal 60    (Get-Field (Get-BatIntent @('--on', '--level', '60')) 'Level') '--on --level: Level'
Assert-Equal 60    (Get-Field (Get-BatIntent @('--level', '60', '--on')) 'Level') 'order does not matter'

Write-Host "`nGet-BatIntent - errors" -ForegroundColor Cyan
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--bogus')) 'Kind')                 'unknown flag'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('80')) 'Kind')                      'bare value needs --level'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--on', '--off')) 'Kind')           '--on with --off'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--off', '--level', '70')) 'Kind')  '--off with --level'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--status', '--on')) 'Kind')        '--status with --on'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--status', '--level', '70')) 'Kind') '--status with --level'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--level')) 'Kind')                 '--level without a number'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--level', '--on')) 'Kind')         '--level followed by a flag'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--level', '85')) 'Kind')           'level not in the list'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--level', '100')) 'Kind')          'level 100'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--level', 'abc')) 'Kind')          'level not a number'
Assert-Equal 'Error' (Get-Field (Get-BatIntent @('--level', '70', '--level', '80')) 'Kind') '--level twice'
Assert-Equal $true   ((Get-Field (Get-BatIntent @('--level', '85')) 'Message') -like '*50, 60, 70, 80, 90*') 'bad level lists the valid ones'

Write-Host "`nGet-BatSteps - on/off" -ForegroundColor Cyan
$onAt70  = @{ On = $true;  Level = 70 }
$offAt70 = @{ On = $false; Level = 70 }
Assert-Equal 'off' (Get-StepString $onAt70  @{ On = $false; Level = $null }) 'on -> --off'
Assert-Equal 'on'  (Get-StepString $offAt70 @{ On = $true;  Level = $null }) 'off -> --on'
Assert-Equal ''    (Get-StepString $onAt70  @{ On = $true;  Level = $null }) '--on when already on does nothing'
Assert-Equal ''    (Get-StepString $offAt70 @{ On = $false; Level = $null }) '--off when already off does nothing'

Write-Host "`nGet-BatSteps - level" -ForegroundColor Cyan
Assert-Equal 'level:80'        (Get-StepString $onAt70  @{ On = $null; Level = 80 }) 'on: just move the slider'
Assert-Equal 'on,level:80,off' (Get-StepString $offAt70 @{ On = $null; Level = 80 }) 'off: slider is disabled, so on -> level -> off'
Assert-Equal ''                (Get-StepString $onAt70  @{ On = $null; Level = 70 }) 'same level does nothing'
Assert-Equal ''                (Get-StepString $offAt70 @{ On = $null; Level = 70 }) 'same level while off does nothing'
Assert-Equal 'on,level:60'     (Get-StepString $offAt70 @{ On = $true; Level = 60 }) '--on --level from off'
Assert-Equal 'on'              (Get-StepString $offAt70 @{ On = $true; Level = 70 }) '--on --level, level already right'
Assert-Equal 'level:60'        (Get-StepString $onAt70  @{ On = $true; Level = 60 }) '--on --level when already on'
Assert-Equal 'on,level:50,off' (Get-StepString @{ On = $false; Level = $null } @{ On = $null; Level = 50 }) 'unknown current level counts as different'

Write-Host "`nGet-LevelIndex" -ForegroundColor Cyan
# The slider is an index 0..4, not a percentage.
Assert-Equal 0 (Get-LevelIndex 50) '50 -> 0'
Assert-Equal 2 (Get-LevelIndex 70) '70 -> 2'
Assert-Equal 4 (Get-LevelIndex 90) '90 -> 4'

Write-Host "`nFormat-BatStatus" -ForegroundColor Cyan
$nl = [Environment]::NewLine
Assert-Equal ("Battery protection: On$nl  Charge limit   70%") (Format-BatStatus $true 70) 'on'
Assert-Equal ("Battery protection: Off$nl  Charge limit   80% (not active)") (Format-BatStatus $false 80) 'off'
Assert-Equal ("Battery protection: On$nl  Charge limit   unknown") (Format-BatStatus $true $null) 'no level stored'
Assert-Equal ("Battery protection: Off$nl  Charge limit   unknown") (Format-BatStatus $false 0) 'unknown level is never "(not active)"'

Write-Host "`nGet-UsageText content" -ForegroundColor Cyan
$usage = "$(Get-UsageText)"
foreach ($flag in '--status', '--on', '--off', '--level', '--help') {
    Assert-Equal $true ($usage -like "*$flag*") "usage mentions $flag"
}
Assert-Equal $true ($usage -like '*50, 60, 70, 80, 90*') 'usage lists the levels'

Write-Host "`nInvoke-Main exit codes (nothing touched)" -ForegroundColor Cyan
Assert-Equal 0 (Invoke-Main @('--help'))           'help exits 0'
Assert-Equal 0 (Invoke-Main @())                   'no args exits 0'
Assert-Equal 2 (Invoke-Main @('--bogus'))          'unknown flag exits 2'
Assert-Equal 2 (Invoke-Main @('--on', '--off'))    'conflict exits 2'
Assert-Equal 2 (Invoke-Main @('--level', '85'))    'bad level exits 2'

Write-Host "`nGet-BatState (live, read-only)" -ForegroundColor Cyan
$state = Get-BatState
Assert-Equal $true ($state -is [hashtable]) 'returns a hashtable'
if ($state -is [hashtable]) {
    Assert-Equal $true ($state.On -is [bool]) 'On is a bool'
    Assert-Equal $true ($null -eq $state.Level -or $state.Level -is [int]) 'Level is an int or null'
}

Write-Host "`n--status (live, read-only)" -ForegroundColor Cyan
Assert-Equal 0 (Invoke-Main @('--status')) 'status exits 0'

Write-Host "`n$($script:Pass) passed, $($script:Fail) failed`n"
if ($script:Fail -gt 0) { exit 1 }
exit 0
