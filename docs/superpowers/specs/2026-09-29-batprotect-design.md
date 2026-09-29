# batprotect — design

**Date:** 2026-09-29
**Status:** Implemented. The user asked to skip the separate spec review and plan
steps, so this records the design as agreed and built.
**Project:** `projects/batprotect/`

## Purpose

A CLI that turns Samsung's **Battery protection** on or off and sets its charge limit
from the terminal. It is the same setting as Samsung Settings → Battery and
performance → Battery protection. The usual pattern: protection on normally, off
now and then for a full charge, then on again.

## Interface

```
batprotect --status          Show on/off and the charge limit
batprotect --on              Turn Battery protection on
batprotect --off             Turn Battery protection off
batprotect --level N         Set the charge limit (on/off unchanged)
batprotect --on --level N    Set the charge limit and turn on
batprotect --help            Show help (also with no arguments)
```

Decisions made with the user:

- Name: `batprotect`, which mirrors Samsung's own wording.
- `--level N` only changes the limit. If protection is off it stays off.
- Levels: 50, 60, 70, 80, 90 (the slider's stops). `80%` is also accepted.
- Output in English.
- A one-line note after `--off`: Samsung can turn protection back on by itself after
  a long time on the charger. This comes from the app's own strings.

Rules: `--x` and `-x` forms, case-insensitive. `--help` anywhere wins. Usage errors
(exit 2): unknown flag, bare value, `--on` with `--off`, `--off` with `--level`,
`--status` with anything, `--level` without a valid number, `--level` twice.
Repeating `--on` is harmless. A request that matches the current state does nothing
and never opens the app. Output is the state re-read after the change, and a mismatch
is exit 1.

## Platform facts (verified on the target machine)

| Fact | Value |
|---|---|
| Machine | Samsung Galaxy Book6 Pro, BIOS PRHK.2.5.59.585 |
| App | Samsung Settings 8.0.13 (`SAMSUNGELECTRONICSCO.LTD.SamsungSettings1.5`), app id `App` |
| State | `HKLM\SOFTWARE\Samsung\SamsungSettings\ModuleProtectBattery`: `OnOff` (0/1), `Value` (percent, absent until first changed), `Support` |
| Registry ACL | Users: read. Admins/SYSTEM: write |
| Who writes it | `SamsungSystemSupportEngine` (SYSTEM service), on request from the app over pipe `{9187E7E8-…}_SamsungSystemSupportEngine` |
| Enforcement | The service passes the limit to Windows Smart charging. `powercfg` does not change |
| Direct service calls | The app's `SettingsSDK.dll` (.NET 8) can be loaded (from bytes; `LoadFrom` gives "Access is denied"). Reads work. Writes are refused: `No Verify Client: Digital Sign check Fail` |
| UI Automation | Switch: `ToggleSwitch` (TogglePattern), inside the same group as the slider. Slider: `_LEVEL_SLIDER` (RangeValuePattern, index 0..4). Page: list item `BatteryPerformance` (SelectionItemPattern) |
| Slider while off | Disabled |
| Window | The `SamsungSettingsHost` process with a main window. A second, windowless `SamsungSettingsHost` is Samsung's background service |

## Approach

Three were considered:

1. **Samsung's SDK from PowerShell 7.** Rejected after a spike. Reads work, but the
   service rejects changes from non-Samsung-signed callers. Getting around that
   would mean defeating a deliberate security check, so it is not pursued.
2. **Speaking the pipe protocol directly.** Rejected for the same signature check,
   and it would also need reverse-engineering.
3. **UI Automation of the Samsung Settings window.** Chosen. The app makes the
   change, so the check is respected.

The user accepted the trade-offs of approach 3: changes take about 3 s when the app
has to open (the window may briefly appear), and a Samsung redesign can break
changes. It fails loudly, never silently.

## Stack

Windows PowerShell 5.1, same as `lidaction`. No PowerShell 7, no .NET SDK, no admin.
Files follow `lidaction`: `batprotect.ps1`, `bin\batprotect.cmd` (the only thing on
PATH), `install.ps1`, `test.ps1`, `test-roundtrip.ps1`, `README.md`.

## Components (`batprotect.ps1`)

| Function | Kind | Responsibility |
|---|---|---|
| `Get-BatIntent` | pure | argv → `Help` / `Status` / `Set {On, Level}` / `Error` |
| `Get-BatSteps` | pure | current state + intent → UI steps, e.g. `on,level:80,off` |
| `Get-LevelIndex` | pure | percent → slider index |
| `Format-BatStatus` | pure | state → output text |
| `Get-BatState` | read | registry → `{On, Level}` |
| `Start-SamsungSettings` | ui | launch `shell:AppsFolder\<PFN>!<AppId>`, wait for the window |
| `Get-BatControls` | ui | go to the page, find the slider by id and the switch next to it |
| `Invoke-BatStep(s)` | ui | apply the steps, waiting for the registry after each; close the app only if we opened it, and wait until it is really closed |

`--level` while off runs `on → level → off` because the slider is disabled when
protection is off.

## Testing

No Pester, same as `lidaction`.

- `test.ps1`: parsing, step planning, index mapping, formatting, help text, and
  exit codes, plus a read-only live `--status`. It never changes anything.
  `Assert-Equal` also compares types, because plain `-eq` coerces (`$true -eq 'x'`
  is true), and that let broken tests pass against the first stub.
- `test-roundtrip.ps1`: the real write path, with the app closed and with it already
  open. The registry is the witness. It checks the window is left as found, and
  restores the original state in `finally`.

A bug found this way: `WindowPattern.Close()` is asynchronous, so the tool returned
before the window was gone. It now waits for the window to close.
