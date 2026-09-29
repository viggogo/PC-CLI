<#
.SYNOPSIS
    Turn Samsung's "Battery protection" on or off, and set its charge limit.
.DESCRIPTION
    Drives the same setting as Samsung Settings -> Battery and performance ->
    Battery protection. No administrator rights.
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ValidLevels = @(50, 60, 70, 80, 90)

# ---------------------------------------------------------------- pure helpers

function Get-BatIntent {
    param([string[]]$Argv)

    if ($null -eq $Argv -or $Argv.Count -eq 0) { return @{ Kind = 'Help' } }

    # --help anywhere wins, so `batprotect --on --help` shows help instead of acting.
    foreach ($raw in $Argv) {
        if ($raw -match '^--?(help|h|\?)$') { return @{ Kind = 'Help' } }
    }

    $wantStatus = $false
    $on         = $null
    $level      = $null
    $validList  = $script:ValidLevels -join ', '

    for ($i = 0; $i -lt $Argv.Count; $i++) {
        $raw = $Argv[$i]
        # Every option must be dash-prefixed; a bare "80" is a usage error.
        if ($raw -notmatch '^--?') {
            return @{ Kind = 'Error'; Message = "Unknown option: $raw" }
        }
        $t = ($raw -replace '^--?', '').ToLowerInvariant()

        if ($t -eq 'status' -or $t -eq 's') {
            $wantStatus = $true
        }
        elseif ($t -eq 'on' -or $t -eq 'off') {
            $want = ($t -eq 'on')
            if ($null -ne $on -and $on -ne $want) {
                return @{ Kind = 'Error'; Message = 'Use --on or --off, not both.' }
            }
            $on = $want
        }
        elseif ($t -eq 'level') {
            if ($null -ne $level) {
                return @{ Kind = 'Error'; Message = 'Give --level only once.' }
            }
            if ($i + 1 -ge $Argv.Count -or $Argv[$i + 1] -match '^--?') {
                return @{ Kind = 'Error'; Message = "--level needs a number: $validList." }
            }
            $i++
            $value = $Argv[$i] -replace '%$', ''
            if ($value -notmatch '^\d+$' -or [int]$value -notin $script:ValidLevels) {
                return @{ Kind = 'Error'; Message = "Invalid level '$($Argv[$i])'. Choose one of: $validList." }
            }
            $level = [int]$value
        }
        else {
            return @{ Kind = 'Error'; Message = "Unknown option: $raw" }
        }
    }

    if ($wantStatus -and ($null -ne $on -or $null -ne $level)) {
        return @{ Kind = 'Error'; Message = 'Use --status on its own.' }
    }
    if ($on -eq $false -and $null -ne $level) {
        return @{ Kind = 'Error'; Message = 'Use --off or --level, not both.' }
    }
    if ($wantStatus) { return @{ Kind = 'Status' } }
    return @{ Kind = 'Set'; On = $on; Level = $level }
}

function Format-BatStatus {
    param([bool]$On, $Level)

    if ($null -ne $Level -and [int]$Level -in $script:ValidLevels) {
        $limit = "$Level%"
        if (-not $On) { $limit += ' (not active)' }
    } else {
        $limit = 'unknown'
    }
    return (@(
        "Battery protection: $(if ($On) { 'On' } else { 'Off' })"
        "  Charge limit   $limit"
    ) -join [Environment]::NewLine)
}

# Works out which UI actions turn $Current into what $Intent asks for.
# Returns e.g. @('on', 'level:80', 'off'); an empty list means nothing to do, so
# Samsung Settings is never opened for a no-op.
function Get-BatSteps {
    param($Current, $Intent)

    $finalOn = $Current.On
    if ($null -ne $Intent.On) { $finalOn = $Intent.On }

    $steps = @()
    if ($null -ne $Intent.Level -and $Intent.Level -ne $Current.Level) {
        # Samsung disables the slider while protection is off, so turn it on long
        # enough to move the slider, then back off if that's where it should end.
        if (-not $Current.On) { $steps += 'on' }
        $steps += "level:$($Intent.Level)"
        if (-not $finalOn) { $steps += 'off' }
    }
    elseif ($finalOn -ne $Current.On) {
        $steps += $(if ($finalOn) { 'on' } else { 'off' })
    }
    return $steps
}

# The slider's value is a position 0..4, not a percentage.
function Get-LevelIndex {
    param([int]$Level)
    return [array]::IndexOf($script:ValidLevels, $Level)
}

# ------------------------------------------------------------- samsung reads

# Samsung Settings does not store the setting itself: it asks Samsung's SYSTEM
# service (SamsungSystemSupportEngine), which writes it here and hands the limit
# on to Windows Smart charging. Everyone can read this key; only admins can write
# it -- and writing it directly would bypass the service, so the limit would never
# reach Windows. Read here, change through the app.
$script:RegKey = 'HKLM:\SOFTWARE\Samsung\SamsungSettings\ModuleProtectBattery'

function Get-BatState {
    $r = Get-ItemProperty -Path $script:RegKey -ErrorAction SilentlyContinue
    if ($null -eq $r -or $r.PSObject.Properties.Match('OnOff').Count -eq 0) {
        throw 'Battery protection setting not found. Is Samsung Settings installed?'
    }
    # Value only appears once the level has been changed at least once.
    $level = $null
    if ($r.PSObject.Properties.Match('Value').Count -gt 0) { $level = [int]$r.Value }
    return @{ On = ([int]$r.OnOff -eq 1); Level = $level }
}

# ------------------------------------------------------------ samsung writes
#
# Changes go through the Samsung Settings window via UI Automation. Samsung's
# service only accepts changes from Samsung-signed programs (it answers anything
# else with "No Verify Client: Digital Sign check Fail"), so the app has to be the
# one that flips the switch. Controls are found by AutomationId and structure,
# never by their text, so a change of Windows display language doesn't break this.

$script:TimeoutSec = 20

function Initialize-Uia {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
}

function Wait-Until {
    param([scriptblock]$Condition, [string]$What, [int]$Seconds = $script:TimeoutSec)
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ($true) {
        $result = & $Condition
        if ($result) { return $result }
        if ((Get-Date) -gt $deadline) { throw "Timed out after $Seconds s waiting for $What." }
        Start-Sleep -Milliseconds 100
    }
}

# The Samsung Settings window, or $null. Found through its process rather than its
# title: SamsungSettingsHost also runs as a windowless background service, so the
# one with a main window is the app.
function Get-SamsungWindow {
    $p = Get-Process SamsungSettingsHost -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } |
        Select-Object -First 1
    if ($null -eq $p) { return $null }
    return [Windows.Automation.AutomationElement]::FromHandle($p.MainWindowHandle)
}

function Start-SamsungSettings {
    Initialize-Uia
    # Match by prefix: the "1.5" is part of the package name and may change.
    $pkg = Get-AppxPackage -Name 'SAMSUNGELECTRONICSCO.LTD.SamsungSettings*' |
        Where-Object { $_.Name -notlike '*Runtime*' } | Select-Object -First 1
    if ($null -eq $pkg) {
        throw 'Samsung Settings is not installed. Get it from the Microsoft Store.'
    }
    $appId = (Get-AppxPackageManifest $pkg).Package.Applications.Application |
        Select-Object -First 1 -ExpandProperty Id
    Start-Process "shell:AppsFolder\$($pkg.PackageFamilyName)!$appId"
    return Wait-Until { Get-SamsungWindow } 'Samsung Settings to open'
}

function Find-ById {
    param($Root, [string]$Id)
    $c = New-Object Windows.Automation.PropertyCondition(
        [Windows.Automation.AutomationElement]::AutomationIdProperty, $Id)
    return $Root.FindFirst([Windows.Automation.TreeScope]::Descendants, $c)
}

# Navigates to "Battery and performance" and returns the switch and the slider.
function Get-BatControls {
    param($Window)

    $slider = Find-ById $Window '_LEVEL_SLIDER'
    if ($null -eq $slider) {
        $page = Wait-Until { Find-ById $Window 'BatteryPerformance' } 'the "Battery and performance" menu item'
        $page.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
        $slider = Wait-Until { Find-ById $Window '_LEVEL_SLIDER' } 'the Battery protection slider'
    }

    # The page has several switches with the same AutomationId (Battery protection,
    # USB charging, ...). Walk up from the slider to the nearest ancestor that holds
    # a switch: that is the Battery protection group, and its switch is ours.
    $walker = [Windows.Automation.TreeWalker]::ControlViewWalker
    $node   = $walker.GetParent($slider)
    $toggle = $null
    while ($null -ne $node -and $null -eq $toggle) {
        $toggle = Find-ById $node 'ToggleSwitch'
        $node   = $walker.GetParent($node)
    }
    if ($null -eq $toggle) { throw 'Could not find the Battery protection switch in Samsung Settings.' }

    return @{
        Toggle = $toggle.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern)
        Slider = $slider.GetCurrentPattern([Windows.Automation.RangeValuePattern]::Pattern)
    }
}

function Invoke-BatStep {
    param($Controls, [string]$Step)

    if ($Step -eq 'on' -or $Step -eq 'off') {
        $want = ($Step -eq 'on')
        $isOn = $Controls.Toggle.Current.ToggleState -eq [Windows.Automation.ToggleState]::On
        if ($isOn -ne $want) { $Controls.Toggle.Toggle() }
        [void](Wait-Until { (Get-BatState).On -eq $want } "Battery protection to turn $Step" 5)
    }
    elseif ($Step -match '^level:(\d+)$') {
        $level = [int]$Matches[1]
        $Controls.Slider.SetValue((Get-LevelIndex $level))
        [void](Wait-Until { (Get-BatState).Level -eq $level } "the charge limit to become $level%" 5)
    }
    else {
        throw "Unknown step: $Step"
    }
}

function Invoke-BatSteps {
    param([string[]]$Steps)

    Initialize-Uia
    $window = Get-SamsungWindow
    $opened = $null -eq $window
    if ($opened) { $window = Start-SamsungSettings }
    try {
        $controls = Get-BatControls $window
        foreach ($s in $Steps) { Invoke-BatStep $controls $s }
    } finally {
        # Leave things as we found them: close the app only if we opened it.
        if ($opened) {
            $appPid = $window.Current.ProcessId
            $window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close()
            # Close() only ASKS the window to close and returns at once. Wait until it
            # is really gone, or a batprotect run right after this could pick up the
            # dying window instead of opening a fresh one.
            [void](Wait-Until {
                $p = Get-Process -Id $appPid -ErrorAction SilentlyContinue
                $null -eq $p -or $p.MainWindowHandle -eq [IntPtr]::Zero
            } 'Samsung Settings to close' 10)
        }
    }
}

# ------------------------------------------------------------------------ cli

# Returns the usage text rather than printing it. In PowerShell a function returns
# EVERYTHING written to the output stream, so a Write-Output here would become part
# of Invoke-Main's return value and corrupt the exit code.
function Get-UsageText {
    return @"
batprotect - Samsung Battery protection from the terminal

USAGE
  batprotect --status          Show on/off and the charge limit
  batprotect --on              Turn Battery protection on
  batprotect --off             Turn Battery protection off
  batprotect --level N         Set the charge limit (on/off unchanged)
  batprotect --on --level N    Set the charge limit and turn on
  batprotect --help            Show this help

N is one of: $($script:ValidLevels -join ', ')

Same setting as Samsung Settings -> Battery and performance -> Battery protection.
Changes are made through the Samsung Settings app, which may flash open for a
moment. No administrator rights required.
"@
}

function Invoke-Main {
    param([string[]]$Argv)

    $intent = Get-BatIntent $Argv

    switch ($intent.Kind) {
        'Help' {
            [Console]::Out.WriteLine((Get-UsageText))
            return 0
        }
        'Error' {
            [Console]::Error.WriteLine("batprotect: $($intent.Message)")
            [Console]::Error.WriteLine('')
            [Console]::Error.WriteLine((Get-UsageText))
            return 2
        }
        'Status' {
            $state = Get-BatState
            # [Console]::Out, not Write-Output: anything on the output stream
            # becomes part of this function's return value and breaks the exit code.
            [Console]::Out.WriteLine((Format-BatStatus $state.On $state.Level))
            return 0
        }
        'Set' {
            $steps = @(Get-BatSteps (Get-BatState) $intent)
            if ($steps.Count -gt 0) { Invoke-BatSteps $steps }

            # Re-read and report verified reality, not intent.
            $state = Get-BatState
            $ok = ($null -eq $intent.On -or $state.On -eq $intent.On) -and
                  ($null -eq $intent.Level -or $state.Level -eq $intent.Level)
            [Console]::Out.WriteLine((Format-BatStatus $state.On $state.Level))
            if (-not $ok) {
                [Console]::Error.WriteLine('batprotect: the change did not take effect.')
                return 1
            }
            if ($intent.On -eq $false) {
                [Console]::Out.WriteLine('Note: Samsung may turn protection back on by itself if the charger stays connected for a long time.')
            }
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
        [Console]::Error.WriteLine("batprotect: $($_.Exception.Message)")
        exit 1
    }
}
