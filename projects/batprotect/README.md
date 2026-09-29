# batprotect

Turn Samsung's **Battery protection** on or off and set its charge limit from the
terminal, without clicking through Samsung Settings → Battery and performance.

Battery protection stops charging at the chosen limit (50–90%) so the battery
doesn't sit at 100% while the charger is connected. The usual pattern: leave it on,
turn it off when you need a full charge, turn it back on afterwards.

Built for a Samsung Galaxy Book6 Pro with Samsung Settings 8.0.13.

## Stack

Windows PowerShell 5.1 and UI Automation, both built into Windows. No dependencies,
no admin rights. Needs the **Samsung Settings** app (Microsoft Store).

## Install

```powershell
.\install.ps1
```

Adds this folder's `bin` subfolder to your user `PATH`. Open a new terminal afterwards.

Only `bin\batprotect.cmd` goes on `PATH`, never the project folder itself. PowerShell
resolves `.ps1` from `PATH` ahead of `PATHEXT`, so a `batprotect.ps1` on `PATH` would
shadow the shim and then be blocked by the execution policy. It would also turn
`test.ps1` and `install.ps1` into global commands.

## Usage

```
batprotect --status          Show on/off and the charge limit
batprotect --on              Turn Battery protection on
batprotect --off             Turn Battery protection off
batprotect --level N         Set the charge limit (on/off unchanged)
batprotect --on --level N    Set the charge limit and turn on
batprotect --help            Show help
```

`N` is one of `50`, `60`, `70`, `80`, `90`, and `80%` also works. Both `--on` and
`-on` forms work, and flags are case-insensitive.

Example:

```
> batprotect --status
Battery protection: On
  Charge limit   70%

> batprotect --off
Battery protection: Off
  Charge limit   70% (not active)
Note: Samsung may turn protection back on by itself if the charger stays connected for a long time.

> batprotect --level 80
Battery protection: Off
  Charge limit   80% (not active)

> batprotect --on
Battery protection: On
  Charge limit   80%
```

Exit codes: `0` success, `1` runtime failure (Samsung Settings missing, a control
not found, the change didn't take effect), `2` usage error.

**While a change runs, don't click in the Samsung Settings window.** A change takes
about 3 s when the tool has to open Samsung Settings itself, and well under a second
when the app is already open. `--status` never opens anything.

## How it works

Samsung Settings doesn't store the setting itself. It asks Samsung's SYSTEM service
(`SamsungSystemSupportEngine`) over a named pipe. The service writes the result to

```
HKLM\SOFTWARE\Samsung\SamsungSettings\ModuleProtectBattery
    OnOff   1 = on, 0 = off
    Value   the charge limit in percent (only present once it has been changed)
```

and passes the limit on to Windows Smart charging.

- **Reading** (`--status`) reads that registry key. Every user can read it.
- **Changing** goes through the Samsung Settings window via UI Automation. The tool
  opens the app if it isn't running, goes to *Battery and performance*, sets the
  switch and the slider, and closes the app again if it opened it. After each step
  it waits until the registry shows the new value, and then prints the verified
  state rather than what it asked for.

Why not something simpler:

- **Writing the registry directly** needs admin and bypasses the service, so the
  limit would never reach Windows.
- **Calling Samsung's service directly** (through the app's own `SettingsSDK.dll`)
  works for reads. For changes, the service answers `No Verify Client: Digital Sign
  check Fail`: it only accepts changes from Samsung-signed programs. The app has to
  make the change, so the tool drives the app. The signature check is respected, not
  bypassed.

Controls are found by their internal AutomationId (`_LEVEL_SLIDER`,
`BatteryPerformance`, `ToggleSwitch`) and their position in the tree, never by their
text, so a change of Windows display language doesn't break the tool. There are
several `ToggleSwitch` controls on the page, and the tool takes the one that shares a
parent group with the slider.

The slider is **disabled while protection is off**, as it is in the app. So
`--level N` while off runs on → set level → off. Protection is on for a split second
and ends up off, as asked.

If a Samsung update redesigns the page, changes will fail with a clear error, such as
a timeout waiting for the slider, rather than doing the wrong thing. `--status` keeps
working as long as the registry key is there.

## Tests

```powershell
.\test.ps1             # parsing, step planning, formatting + a read-only --status
.\test-roundtrip.ps1   # end-to-end; changes the real setting, then restores it
```

`test.ps1` never changes the setting or opens Samsung Settings.

**`test-roundtrip.ps1` changes the real setting** and opens and closes Samsung
Settings. It runs `--off`, `--level` while off, `--on`, a no-op `--on`, and
`--on --level`, with the app both closed and already open. It checks the registry
and whether the window was left as found. It records your original on/off, limit,
and window state first, and restores them in a `finally` block. A hard kill
mid-run can still leave the setting changed, so run `batprotect --status` to check.
