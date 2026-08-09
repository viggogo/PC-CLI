# lidaction Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A `lidaction` CLI that sets what closing the laptop lid does — for both plugged-in (AC) and on-battery (DC) — from any Windows terminal.

**Architecture:** A single PowerShell 5.1 script wrapping `powercfg`. Pure functions (argument parsing, name lookup, output formatting) are separated from thin system-touching wrappers, and the script is dot-source guarded so tests can load the functions without running `main`. State is *read* from the registry (locale-independent) and *written* via `powercfg`. A `.cmd` shim puts the bare command on PATH for PowerShell and cmd alike.

**Tech Stack:** Windows PowerShell 5.1, `powercfg.exe`, registry reads via `HKLM:`. No external dependencies, no Pester, no admin rights.

**Spec:** `docs/superpowers/specs/2026-08-07-lidaction-design.md`

## Global Constraints

- Target: Windows PowerShell **5.1** (not PowerShell 7). No `??`, no `?.`, no ternary, no `&&`/`||`.
- **No external dependencies.** No Pester, no modules to install.
- **No administrator rights** may be required at any point, including install.
- All files live under `projects/lidaction/`. Do not modify sibling tools.
- Every set operation writes **both** AC and DC. There is no independent control.
- Changes apply to the **active power plan only**.
- Exit codes: `0` success, `1` runtime failure, `2` usage error.
- Subgroup GUID (`SUB_BUTTONS`): `4f971e89-eebd-4455-a8de-9e59040e7347`
- Setting GUID (`LIDACTION`): `5ca83367-6e45-459f-a27b-476b1d01c936`
- Action indices: `0`=Do nothing, `1`=Sleep, `2`=Hibernate, `3`=Shut down
- Write user-facing errors to **stderr** via `[Console]::Error.WriteLine(...)`, not `Write-Error` (which emits a PowerShell error record with a stack trace).
- Scripts use `Set-StrictMode -Version Latest`. **Consequence, verified:** accessing a missing property on an object throws `PropertyNotFoundException` — never test for a missing registry value with `$obj.Foo -eq $null`; use `$obj.PSObject.Properties.Match('Foo').Count -gt 0`.

## File Structure

```
projects/lidaction/
├── README.md              purpose, stack, install, usage
├── lidaction.ps1          all functions + main, dot-source guarded
├── lidaction.cmd          PATH shim, forwards exit code
├── install.ps1            adds folder to user PATH (idempotent)
├── test.ps1               pure-logic tests — SAFE, changes nothing
└── test-roundtrip.ps1     end-to-end test — CHANGES real power settings, restores them
```

**Note — deviation from spec:** the spec listed a single `test.ps1`. It is split in two here because the two test kinds have very different risk profiles: `test.ps1` is safe to run any time, while the round-trip test actually cycles the machine's lid setting through all four values. Keeping them in one file means a casual `.\test.ps1` mutates system state. Everything else follows the spec exactly.

`lidaction.ps1` holds every function. At ~200 lines with a clean pure/system split it stays readable in one screen-scroll, and splitting into a library file would force the `.cmd` shim and tests to track two paths for no benefit.

---

### Task 1: Pure logic and test harness

Argument parsing, index→name lookup, and status formatting — all pure, no system calls. Plus the minimal assertion harness the later tasks reuse.

**Files:**
- Create: `projects/lidaction/lidaction.ps1`
- Create: `projects/lidaction/test.ps1`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `Get-ActionName([object]$Index) -> [string]` — `'Do nothing'|'Sleep'|'Hibernate'|'Shut down'|'unknown'`. `unknown` for `$null` or out-of-range.
  - `Get-LidIntent([string[]]$Argv) -> [hashtable]` with key `Kind` = `'Help'|'Status'|'Set'|'Error'`; when `Kind='Set'` also `Value` = `[int]` 0–3; when `Kind='Error'` also `Message` = `[string]`.
  - `Format-LidStatus([string]$PlanName, [object]$AcIndex, [object]$DcIndex) -> [string]`
  - Module-scope constants `$script:SubButtons`, `$script:LidActionGuid`, `$script:ActionNames`.

- [ ] **Step 1: Write the failing tests**

Create `projects/lidaction/test.ps1`:

```powershell
# Pure-logic tests for lidaction. SAFE: changes no system state.
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

Write-Host "`n$($script:Pass) passed, $($script:Fail) failed`n"
if ($script:Fail -gt 0) { exit 1 }
exit 0
```

- [ ] **Step 2: Run tests to verify they fail**

```powershell
cd "projects\lidaction"
.\test.ps1
```

Expected: FAIL — `lidaction.ps1` does not exist, so the dot-source on line 5 errors with `The term '...lidaction.ps1' is not recognized` / `Cannot find path`.

- [ ] **Step 3: Write the minimal implementation**

Create `projects/lidaction/lidaction.ps1`:

```powershell
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
                return @{ Kind  = 'Error'
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
```

- [ ] **Step 4: Run tests to verify they pass**

```powershell
.\test.ps1
```

Expected: `30 passed, 0 failed`, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add projects/lidaction/lidaction.ps1 projects/lidaction/test.ps1
git commit -m "feat(lidaction): pure argument parsing and formatting"
```

---

### Task 2: CLI entry point, help, and PATH shim

Makes the tool runnable as a command. After this task `lidaction --help` works from a real terminal with correct exit codes; `--status` and the action flags are recognized but report "not implemented" (wired up in Tasks 3 and 4).

**Files:**
- Modify: `projects/lidaction/lidaction.ps1` (append main dispatch + dot-source guard)
- Create: `projects/lidaction/lidaction.cmd`

**Interfaces:**
- Consumes: `Get-LidIntent` (Task 1).
- Produces:
  - `Show-Usage() -> [void]` — writes usage to stdout.
  - `Invoke-Main([string[]]$Argv) -> [int]` — returns the process exit code.
  - Dot-source guard: `lidaction.ps1` runs `Invoke-Main` only when executed, never when dot-sourced.

- [ ] **Step 1: Write the failing tests**

Append to `projects/lidaction/test.ps1`, immediately **before** the final `Write-Host "...passed..."` summary block:

```powershell
Write-Host "`nInvoke-Main exit codes" -ForegroundColor Cyan
Assert-Equal 0 (Invoke-Main @('--help')) 'help exits 0'
Assert-Equal 0 (Invoke-Main @())         'no args exits 0'
Assert-Equal 2 (Invoke-Main @('--bogus')) 'unknown flag exits 2'
Assert-Equal 2 (Invoke-Main @('--1','--2')) 'two actions exits 2'

Write-Host "`nShow-Usage content" -ForegroundColor Cyan
$usage = (Show-Usage | Out-String)
Assert-Equal $true ($usage -like '*--status*')  'usage mentions --status'
Assert-Equal $true ($usage -like '*--0*')       'usage mentions --0'
Assert-Equal $true ($usage -like '*--3*')       'usage mentions --3'
Assert-Equal $true ($usage -like '*Hibernate*') 'usage mentions Hibernate'
```

- [ ] **Step 2: Run tests to verify they fail**

```powershell
.\test.ps1
```

Expected: FAIL with `The term 'Invoke-Main' is not recognized as the name of a cmdlet`.

- [ ] **Step 3: Write the minimal implementation**

Append to `projects/lidaction/lidaction.ps1`:

```powershell
# ------------------------------------------------------------------------ cli

function Show-Usage {
    Write-Output @'
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
            Show-Usage
            return 0
        }
        'Error' {
            [Console]::Error.WriteLine("lidaction: $($intent.Message)")
            [Console]::Error.WriteLine('')
            [Console]::Error.WriteLine((Show-Usage | Out-String).TrimEnd())
            return 2
        }
        'Status' {
            [Console]::Error.WriteLine('lidaction: --status not implemented yet')
            return 1
        }
        'Set' {
            [Console]::Error.WriteLine('lidaction: setting not implemented yet')
            return 1
        }
    }
    return 1
}

# Dot-source guard: when tests dot-source this file ($MyInvocation.InvocationName
# is '.'), define the functions but do not run. Only run main when executed.
if ($MyInvocation.InvocationName -ne '.') {
    exit (Invoke-Main $args)
}
```

- [ ] **Step 4: Run tests to verify they pass**

```powershell
.\test.ps1
```

Expected: `38 passed, 0 failed`, exit 0.

- [ ] **Step 5: Create the PATH shim**

Create `projects/lidaction/lidaction.cmd`:

```bat
@echo off
REM Bare `lidaction` cannot resolve to a .ps1 because .PS1 is not in PATHEXT.
REM This shim makes the command work from PowerShell, cmd, and Windows Terminal.
REM -ExecutionPolicy Bypass applies to THIS invocation only, nothing system-wide.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lidaction.ps1" %*
exit /b %ERRORLEVEL%
```

- [ ] **Step 6: Verify the shim end-to-end**

```powershell
.\lidaction.cmd --help
"exit=$LASTEXITCODE"
.\lidaction.cmd --bogus
"exit=$LASTEXITCODE"
```

Expected: usage text then `exit=0`; then an error plus usage on stderr and `exit=2`. If the second shows `exit=0`, the `exit /b %ERRORLEVEL%` line is missing or `exit` is absent from the `.ps1` guard.

- [ ] **Step 7: Commit**

```bash
git add projects/lidaction/lidaction.ps1 projects/lidaction/lidaction.cmd projects/lidaction/test.ps1
git commit -m "feat(lidaction): cli entry point, help, and PATH shim"
```

---

### Task 3: Read current setting and implement `--status`

**Files:**
- Modify: `projects/lidaction/lidaction.ps1`

**Interfaces:**
- Consumes: `$script:SubButtons`, `$script:LidActionGuid`, `Format-LidStatus` (Task 1); `Invoke-Main` (Task 2).
- Produces:
  - `Get-ActiveScheme() -> [hashtable]` with `Guid` = `[string]`, `Name` = `[string]`. Throws on failure.
  - `Get-LidActionIndices([string]$SchemeGuid) -> [hashtable]` with `Ac`, `Dc` (`[int]` or `$null`) and `Source` = `'registry'|'powercfg'|'unknown'`.

- [ ] **Step 1: Write the failing tests**

Append to `test.ps1` before the summary block:

```powershell
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
```

- [ ] **Step 2: Run tests to verify they fail**

```powershell
.\test.ps1
```

Expected: FAIL with `The term 'Get-ActiveScheme' is not recognized`.

- [ ] **Step 3: Write the minimal implementation**

Insert into `lidaction.ps1` **after** `Format-LidStatus` and **before** the `# ---- cli` section:

```powershell
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
```

- [ ] **Step 4: Wire `--status` into `Invoke-Main`**

In `Invoke-Main`, replace the entire `'Status' { ... }` branch with:

```powershell
        'Status' {
            $scheme = Get-ActiveScheme
            $idx    = Get-LidActionIndices $scheme.Guid
            Write-Output (Format-LidStatus $scheme.Name $idx.Ac $idx.Dc)
            if ($idx.Source -eq 'unknown') {
                [Console]::Error.WriteLine('lidaction: could not read the current setting.')
                return 1
            }
            return 0
        }
```

- [ ] **Step 5: Run tests to verify they pass**

```powershell
.\test.ps1
```

Expected: `44 passed, 0 failed`, exit 0.

- [ ] **Step 6: Verify against the real machine**

```powershell
.\lidaction.cmd --status
powercfg -qh SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936
```

Expected: the tool's two rows match the `Current AC/DC Power Setting Index` values from `powercfg` (`0x00000000` = `Do nothing`). **On this machine both are currently `Do nothing`.**

- [ ] **Step 7: Commit**

```bash
git add projects/lidaction/lidaction.ps1 projects/lidaction/test.ps1
git commit -m "feat(lidaction): read current lid action and implement --status"
```

---

### Task 4: Write the setting and implement `--0`..`--3`

**Files:**
- Modify: `projects/lidaction/lidaction.ps1`

**Interfaces:**
- Consumes: `Get-ActiveScheme`, `Get-LidActionIndices` (Task 3); `Invoke-Main` (Task 2).
- Produces:
  - `Test-HibernateEnabled() -> [bool]`
  - `Set-LidActionIndices([int]$Index) -> [void]` — writes AC and DC, re-applies the plan. Throws on `powercfg` failure.

- [ ] **Step 1: Write the failing tests**

Append to `test.ps1` before the summary block:

```powershell
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
```

- [ ] **Step 2: Run tests to verify they fail**

```powershell
.\test.ps1
```

Expected: FAIL with `The term 'Test-HibernateEnabled' is not recognized`.

- [ ] **Step 3: Write the minimal implementation**

Insert into `lidaction.ps1` after `Get-LidActionIndices`:

```powershell
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
```

- [ ] **Step 4: Wire the `Set` branch into `Invoke-Main`**

Replace the entire `'Set' { ... }` branch in `Invoke-Main` with:

```powershell
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
            Write-Output "Lid close action set to: $(Get-ActionName $n)"
            Write-Output (Format-LidStatus $scheme.Name $idx.Ac $idx.Dc)
            return 0
        }
```

- [ ] **Step 5: Wrap `Invoke-Main` so thrown errors become exit 1**

Replace the dot-source guard at the bottom of `lidaction.ps1` with:

```powershell
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
```

- [ ] **Step 6: Run tests to verify they pass**

```powershell
.\test.ps1
```

Expected: `50 passed, 0 failed`, exit 0. The suite restores the original setting; confirm with `.\lidaction.cmd --status` that it reads `Do nothing` on both rows afterward.

- [ ] **Step 7: Commit**

```bash
git add projects/lidaction/lidaction.ps1 projects/lidaction/test.ps1
git commit -m "feat(lidaction): write lid action for AC and DC"
```

---

### Task 5: Destructive round-trip test

Exercises all four values against the real system and restores the original in a `finally` block.

**Files:**
- Create: `projects/lidaction/test-roundtrip.ps1`

**Interfaces:**
- Consumes: `Get-ActiveScheme`, `Get-LidActionIndices`, `Set-LidActionIndices`, `Get-ActionName` (Tasks 1, 3, 4).
- Produces: nothing consumed by later tasks.

- [ ] **Step 1: Write the test script**

Create `projects/lidaction/test-roundtrip.ps1`:

```powershell
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
```

Note: the restore path calls `Invoke-Powercfg` directly rather than `Set-LidActionIndices`, because `Set-LidActionIndices` writes the *same* value to AC and DC — and the original AC and DC values may differ.

- [ ] **Step 2: Run it**

```powershell
.\test-roundtrip.ps1
```

Expected: 10 PASS lines, `10 passed, 0 failed`, exit 0.

- [ ] **Step 3: Confirm the machine was left as found**

```powershell
.\lidaction.cmd --status
```

Expected: both rows read `Do nothing` — the values recorded during design.

- [ ] **Step 4: Commit**

```bash
git add projects/lidaction/test-roundtrip.ps1
git commit -m "test(lidaction): end-to-end round trip with guaranteed restore"
```

---

### Task 6: Installer and documentation

**Files:**
- Create: `projects/lidaction/install.ps1`
- Create: `projects/lidaction/README.md`
- Modify: `projects/README.md` (replace `_None yet._` in the Tools list)

**Interfaces:**
- Consumes: nothing.
- Produces: nothing consumed by later tasks.

- [ ] **Step 1: Write the installer**

Create `projects/lidaction/install.ps1`:

```powershell
# Adds this folder to the CURRENT USER's PATH. No administrator rights needed.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$dir = $PSScriptRoot

$current = [Environment]::GetEnvironmentVariable('Path', 'User')
if ($null -eq $current) { $current = '' }

$parts = @($current -split ';' | Where-Object { $_.Trim() -ne '' })

# Idempotent: compare case-insensitively and ignore a trailing backslash.
$normalized = $parts | ForEach-Object { $_.TrimEnd('\').ToLowerInvariant() }
if ($normalized -contains $dir.TrimEnd('\').ToLowerInvariant()) {
    Write-Host "Already on your PATH:`n  $dir"
} else {
    [Environment]::SetEnvironmentVariable('Path', (($parts + $dir) -join ';'), 'User')
    Write-Host "Added to your user PATH:`n  $dir"
}

Write-Host ''
Write-Host 'Open a NEW terminal, then run:  lidaction --status'
```

- [ ] **Step 2: Run it and verify idempotency**

```powershell
.\install.ps1
.\install.ps1
```

Expected: first run prints `Added to your user PATH:`; second prints `Already on your PATH:`. Then confirm no duplicate entry:

```powershell
([Environment]::GetEnvironmentVariable('Path','User') -split ';' | Where-Object { $_ -like '*lidaction*' }).Count
```

Expected: `1`

- [ ] **Step 3: Verify the command works from a new shell**

```powershell
powershell -NoProfile -Command "lidaction --status"
```

Expected: the status output. If `lidaction` is not found, the new process inherited the old environment — open a genuinely new terminal instead.

- [ ] **Step 4: Write the README**

Create `projects/lidaction/README.md`:

```markdown
# lidaction

Set what closing the laptop lid does — from the terminal, without clicking through
Control Panel → Power Options → "Choose what closing the lid does".

Every change applies to **both** power states: plugged in (AC) and on battery (DC).

## Stack

Windows PowerShell 5.1. No dependencies, no admin rights.

## Install

```powershell
.\install.ps1
```

Adds this folder to your user `PATH`. Open a new terminal afterwards.

## Usage

```
lidaction --status        Show the current lid close action
lidaction --0             Do nothing
lidaction --1             Sleep
lidaction --2             Hibernate
lidaction --3             Shut down
lidaction --help          Show help
```

Example:

```
> lidaction --status
Lid close action  (plan: Balanced)
  Plugged in   Do nothing
  On battery   Do nothing

> lidaction --1
Lid close action set to: Sleep
  Plugged in   Sleep
  On battery   Sleep
```

Exit codes: `0` success, `1` runtime failure, `2` usage error.

## How it works

Reads the current value from the registry (locale-independent) and writes with
`powercfg /setacvalueindex` + `/setdcvalueindex`, then `powercfg /setactive` to
apply it immediately. Changes affect the **active power plan** only, matching what
the Control Panel page does.

Note: "Lid close action" is a *hidden* power setting — `powercfg /query` won't show
it, only `powercfg -qh` does.

## Tests

```powershell
.\test.ps1             # pure logic; safe, changes nothing
.\test-roundtrip.ps1   # end-to-end; changes the real setting, then restores it
```
```

- [ ] **Step 5: Update the projects index**

In `projects/README.md`, replace the line `_None yet._` with:

```markdown
- [`lidaction`](lidaction/) — set what closing the laptop lid does (AC and DC).
```

- [ ] **Step 6: Full verification**

```powershell
.\test.ps1
.\lidaction.cmd --help
.\lidaction.cmd --status
```

Expected: all tests pass; help and status both render correctly.

- [ ] **Step 7: Commit**

```bash
git add projects/lidaction/install.ps1 projects/lidaction/README.md projects/README.md
git commit -m "feat(lidaction): installer and documentation"
```

---

## Verification Checklist

Run after Task 6. Every line must hold before calling this done.

- [ ] `.\test.ps1` → `50 passed, 0 failed`, exit 0
- [ ] `.\test-roundtrip.ps1` → `10 passed, 0 failed`, exit 0
- [ ] `lidaction --status` matches `powercfg -qh SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936`
- [ ] `lidaction --1` then physically close the lid → machine sleeps
- [ ] `lidaction --0` then physically close the lid → machine stays awake
- [ ] `lidaction --bogus` → usage on stderr, `$LASTEXITCODE` is 2
- [ ] `lidaction --1 --2` → error on stderr, `$LASTEXITCODE` is 2
- [ ] `lidaction` with no args → usage, exit 0
- [ ] Works in a **non-elevated** terminal (no UAC prompt at any point)
- [ ] Works from `cmd.exe` as well as PowerShell
- [ ] Original setting restored: both rows read `Do nothing`
