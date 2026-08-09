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

Adds this folder's `bin` subfolder to your user `PATH`. Open a new terminal afterwards.

Only `bin\lidaction.cmd` is exposed, never the project folder itself — PowerShell
resolves `.ps1` from `PATH` ahead of `PATHEXT`, so a `lidaction.ps1` on `PATH` would
shadow the shim and then be blocked by the execution policy, and `test.ps1` /
`install.ps1` would become global commands.

## Usage

```
lidaction --status        Show the current lid close action
lidaction --0             Do nothing
lidaction --1             Sleep
lidaction --2             Hibernate
lidaction --3             Shut down
lidaction --help          Show help
```

Both `--0` and `-0` forms work, and flags are case-insensitive.

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

After writing, the tool re-reads the value and prints what it actually found, so
the output reflects reality rather than intent.

Note: "Lid close action" is a *hidden* power setting — `powercfg /query` won't show
it, only `powercfg -qh` does.

If you ask for Hibernate (`--2`) while hibernation is disabled on the system, the
tool warns on stderr but still writes the setting.

## Tests

```powershell
.\test.ps1             # parsing/formatting logic + one live set; restores it
.\test-roundtrip.ps1   # end-to-end; changes the real setting, then restores it
```

**Both scripts touch the real setting.** Most of `test.ps1` is pure parsing and
formatting logic, but its last section runs `lidaction --1` against the live system to
check the set path end to end. `test-roundtrip.ps1` cycles the lid setting through all
four values.

Each records your original AC and DC values first and restores them in a `finally`
block, so an interrupted run still puts them back. A hard kill between the write and
the `finally` can still leave the setting changed — run `lidaction --status` to check.
