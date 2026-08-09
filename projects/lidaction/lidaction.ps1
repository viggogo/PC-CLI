<#
.SYNOPSIS
    Set what closing the laptop lid does (Windows).
.DESCRIPTION
    Reads and writes the "Lid close action" power setting for both AC (plugged in)
    and DC (on battery) on the active power plan. Requires no administrator rights.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# NOTE: "Lid close action" is a HIDDEN power setting. `powercfg /query` will NOT
# list it -- only `powercfg -qh` does. Don't go hunting when /query shows nothing.
$script:SubButtons    = '4f971e89-eebd-4455-a8de-9e59040e7347'  # SUB_BUTTONS
$script:LidActionGuid = '5ca83367-6e45-459f-a27b-476b1d01c936'  # LIDACTION
$script:ActionNames   = @('Do nothing', 'Sleep', 'Hibernate', 'Shut down')

# ---------------------------------------------------------------- pure helpers

function Get-ActionName {
    param($Index)
    if ($null -eq $Index) { return 'unknown' }
    $i = [int]$Index
    if ($i -lt 0 -or $i -ge $script:ActionNames.Count) { return 'unknown' }
    return $script:ActionNames[$i]
}

function Get-LidIntent {
    param([string[]]$Argv)

    if ($null -eq $Argv -or $Argv.Count -eq 0) { return @{ Kind = 'Help' } }

    $wantStatus = $false
    $action     = $null

    foreach ($raw in $Argv) {
        # Every option must be dash-prefixed; a bare "0" is a usage error.
        if ($raw -notmatch '^--?') {
            return @{ Kind = 'Error'; Message = "Unknown option: $raw" }
        }
        $t = ($raw -replace '^--?', '').ToLowerInvariant()

        if ($t -eq 'help' -or $t -eq 'h' -or $t -eq '?') {
            return @{ Kind = 'Help' }
        }
        elseif ($t -eq 'status' -or $t -eq 's') {
            $wantStatus = $true
        }
        elseif ($t -match '^[0-3]$') {
            if ($null -ne $action) {
                return @{ Kind    = 'Error'
                          Message = "Specify only one action (got --$action and --$t)." }
            }
            $action = [int]$t
        }
        else {
            return @{ Kind = 'Error'; Message = "Unknown option: $raw" }
        }
    }

    if ($null -ne $action -and $wantStatus) {
        return @{ Kind = 'Error'; Message = 'Use --status or an action flag, not both.' }
    }
    if ($null -ne $action) { return @{ Kind = 'Set'; Value = $action } }
    if ($wantStatus)       { return @{ Kind = 'Status' } }
    return @{ Kind = 'Help' }
}

function Format-LidStatus {
    param([string]$PlanName, $AcIndex, $DcIndex)
    return (@(
        "Lid close action  (plan: $PlanName)"
        "  Plugged in   $(Get-ActionName $AcIndex)"
        "  On battery   $(Get-ActionName $DcIndex)"
    ) -join [Environment]::NewLine)
}

# --------------------------------------------------------------- system reads

function Get-ActiveScheme {
    # A GUID is a GUID in every locale, so this regex is language-safe even
    # though the surrounding powercfg text is localized.
    $out = (& powercfg /getactivescheme 2>&1 | Out-String)
    if ($out -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\s*\(([^)]*)\)') {
        return @{ Guid = $Matches[1]; Name = $Matches[2].Trim() }
    }
    if ($out -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
        return @{ Guid = $Matches[1]; Name = 'unknown plan' }
    }
    throw "Could not determine the active power plan. powercfg said: $($out.Trim())"
}

function Get-LidActionFromPowercfg {
    # Fallback only. These labels ARE localized, so this is second choice.
    $out = (& powercfg -qh SCHEME_CURRENT $script:SubButtons $script:LidActionGuid 2>&1 | Out-String)
    $ac = $null
    $dc = $null
    if ($out -match '(?m)^\s*Current AC.*?:\s*0x([0-9a-fA-F]+)') { $ac = [Convert]::ToInt32($Matches[1], 16) }
    if ($out -match '(?m)^\s*Current DC.*?:\s*0x([0-9a-fA-F]+)') { $dc = [Convert]::ToInt32($Matches[1], 16) }
    if ($null -eq $ac -or $null -eq $dc) { return $null }
    return @{ Ac = $ac; Dc = $dc; Source = 'powercfg' }
}

function Get-LidActionIndices {
    param([string]$SchemeGuid)

    $path = "HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\" +
            "$SchemeGuid\$($script:SubButtons)\$($script:LidActionGuid)"

    $v = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
    if ($null -ne $v) {
        # StrictMode turns a missing property into a thrown PropertyNotFoundException,
        # so probe PSObject.Properties instead of comparing the property to $null.
        $hasAc = $v.PSObject.Properties.Match('ACSettingIndex').Count -gt 0
        $hasDc = $v.PSObject.Properties.Match('DCSettingIndex').Count -gt 0
        if ($hasAc -and $hasDc) {
            return @{ Ac = [int]$v.ACSettingIndex; Dc = [int]$v.DCSettingIndex; Source = 'registry' }
        }
    }

    $fallback = Get-LidActionFromPowercfg
    if ($null -ne $fallback) { return $fallback }

    # Never guess a value.
    return @{ Ac = $null; Dc = $null; Source = 'unknown' }
}

# -------------------------------------------------------------- system writes

function Test-HibernateEnabled {
    # Read the registry rather than parsing `powercfg /a`: that output is
    # localized AND prints the word "Hibernate" in both its available and
    # not-available sections, so text matching there is unreliable.
    $h = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' `
                          -Name 'HibernateEnabled' -ErrorAction SilentlyContinue
    if ($null -eq $h) { return $false }
    if ($h.PSObject.Properties.Match('HibernateEnabled').Count -eq 0) { return $false }
    return ([int]$h.HibernateEnabled -eq 1)
}

function Invoke-Powercfg {
    param([string[]]$PowercfgArgs)
    $out = (& powercfg @PowercfgArgs 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "powercfg $($PowercfgArgs -join ' ') failed (exit $LASTEXITCODE): $($out.Trim())"
    }
}

function Set-LidActionIndices {
    param([int]$Index)
    Invoke-Powercfg @('/setacvalueindex', 'SCHEME_CURRENT', $script:SubButtons, $script:LidActionGuid, "$Index")
    Invoke-Powercfg @('/setdcvalueindex', 'SCHEME_CURRENT', $script:SubButtons, $script:LidActionGuid, "$Index")
    # Re-apply the plan so the change takes effect immediately; without this
    # Windows can keep the old behavior until something else touches the plan.
    Invoke-Powercfg @('/setactive', 'SCHEME_CURRENT')
}

# ------------------------------------------------------------------------ cli

# Returns the usage text rather than printing it. This matters: in PowerShell a
# function returns EVERYTHING written to the output stream, so a Write-Output here
# would become part of Invoke-Main's return value and corrupt the exit code.
function Get-UsageText {
    return @'
lidaction - set what closing the laptop lid does (Windows)

USAGE
  lidaction --status        Show the current lid close action
  lidaction --0             Do nothing
  lidaction --1             Sleep
  lidaction --2             Hibernate
  lidaction --3             Shut down
  lidaction --help          Show this help

Changes apply to both power states (plugged in and on battery)
for the active power plan. No administrator rights required.
'@
}

function Invoke-Main {
    param([string[]]$Argv)

    $intent = Get-LidIntent $Argv

    switch ($intent.Kind) {
        'Help' {
            [Console]::Out.WriteLine((Get-UsageText))
            return 0
        }
        'Error' {
            [Console]::Error.WriteLine("lidaction: $($intent.Message)")
            [Console]::Error.WriteLine('')
            [Console]::Error.WriteLine((Get-UsageText))
            return 2
        }
        'Status' {
            $scheme = Get-ActiveScheme
            $idx    = Get-LidActionIndices $scheme.Guid
            # [Console]::Out, not Write-Output: anything on the output stream
            # becomes part of this function's return value and breaks the exit code.
            [Console]::Out.WriteLine((Format-LidStatus $scheme.Name $idx.Ac $idx.Dc))
            if ($idx.Source -eq 'unknown') {
                [Console]::Error.WriteLine('lidaction: could not read the current setting.')
                return 1
            }
            return 0
        }
        'Set' {
            $n = $intent.Value

            if ($n -eq 2 -and -not (Test-HibernateEnabled)) {
                [Console]::Error.WriteLine(
                    'lidaction: warning - hibernation is disabled on this system, so the lid may not hibernate. Enable it with: powercfg /h on')
            }

            Set-LidActionIndices $n

            # Re-read and report verified reality, not intent.
            $scheme = Get-ActiveScheme
            $idx    = Get-LidActionIndices $scheme.Guid
            if ($idx.Ac -ne $n -or $idx.Dc -ne $n) {
                [Console]::Error.WriteLine(
                    "lidaction: setting did not take effect (wanted $n, read AC=$($idx.Ac) DC=$($idx.Dc)).")
                return 1
            }
            [Console]::Out.WriteLine("Lid close action set to: $(Get-ActionName $n)")
            [Console]::Out.WriteLine((Format-LidStatus $scheme.Name $idx.Ac $idx.Dc))
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
        [Console]::Error.WriteLine("lidaction: $($_.Exception.Message)")
        exit 1
    }
}
