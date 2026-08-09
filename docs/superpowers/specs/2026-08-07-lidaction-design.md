# lidaction — design

**Date:** 2026-08-07
**Status:** Approved
**Project:** `projects/lidaction/`

## Purpose

A CLI to set what closing the laptop lid does — the setting behind Control Panel's
"Choose what closing the lid does" — from any terminal, without clicking through
Power Options.

Every change writes **both** power states: plugged in (AC) and on battery (DC).
There is no way to set them independently; that is the point of the tool.

## Scope

In scope:

- Set the lid close action to one of four values.
- Report the current setting.
- Show usage.
- Install itself onto the user's PATH.

Explicitly out of scope (decided, not deferred):

- No toggle or "restore previous" mode. The user was asked and chose the flat
  set-exact-value interface.
- No control over the power button or sleep button, though they live in the same
  `SUB_BUTTONS` subgroup.
- No independent AC/DC control.
- No power-plan management.

## Platform facts

Verified on the target machine (Windows 11 Pro 26200, PowerShell 5.1) during design:

| Fact | Value |
|---|---|
| Chassis | Notebook (`Win32_SystemEnclosure.ChassisTypes = 10`) |
| Battery | Present |
| Power plans | One: Balanced `381b4222-f694-41f0-9685-ff5bb260df2e` |
| Current lid action | AC = 0, DC = 0 (Do nothing) |
| Hibernate available | Yes (`powercfg /a`) |
| Admin required to write | **No** — verified with a no-op `powercfg /setacvalueindex` |
| Pester available | 3.4.0 only (the version Windows bundles) |

GUIDs:

- Subgroup `SUB_BUTTONS` = `4f971e89-eebd-4455-a8de-9e59040e7347`
- Setting `LIDACTION` = `5ca83367-6e45-459f-a27b-476b1d01c936`

Action values:

| Index | Meaning |
|---|---|
| 0 | Do nothing |
| 1 | Sleep |
| 2 | Hibernate |
| 3 | Shut down |

Note: `LIDACTION` is a *hidden* power setting. Plain `powercfg /query` does not
list it; `powercfg -qh` does. This surprises people reading the code later, so it
is worth a comment in the source.

## Stack

PowerShell 5.1 script, no dependencies. Chosen over Python and Go because the tool
is fundamentally a `powercfg` wrapper — a runtime dependency or a compile step buys
nothing here.

## Files

```
projects/lidaction/
├── README.md          purpose, stack, install, usage
├── lidaction.ps1      the tool
├── lidaction.cmd      PATH shim
├── install.ps1        adds this folder to the user PATH
└── test.ps1           verification
```

Self-contained under `projects/`, per the repo's CLAUDE.md.

## Components

### `lidaction.cmd` — PATH shim

```bat
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0lidaction.ps1" %*
```

Exists because PowerShell will not resolve a bare `lidaction` to a `.ps1` on PATH —
`.PS1` is not in `PATHEXT` by default. The shim makes the bare command work from
PowerShell, cmd, and Windows Terminal alike.

`-ExecutionPolicy Bypass` is a deliberate choice so the tool works regardless of the
machine's execution policy. It applies to this single invocation only and changes
nothing system-wide. `-NoProfile` keeps startup fast and avoids profile side effects.

The shim must forward the child's exit code so the documented exit codes survive.

### `install.ps1`

Appends the project folder to the **user** `PATH` environment variable (`HKCU`, via
`[Environment]::SetEnvironmentVariable(..., 'User')`). No admin needed. Must be
idempotent: re-running does not duplicate the entry. Prints the path added and a
note that already-open terminals need restarting.

### `lidaction.ps1` — the tool

Internal structure. The pure functions carry the logic and are unit-testable with no
system calls; the `powercfg`-touching wrappers stay thin.

| Function | Kind | Responsibility |
|---|---|---|
| `Parse-Args` | pure | `$args` → intent (`Help`/`Status`/`Set n`/`Error`) |
| `Get-ActionName` | pure | index → display name |
| `Format-Status` | pure | plan name + AC/DC indices → output text |
| `Get-ActiveSchemeGuid` | system | parse GUID from `powercfg /getactivescheme` |
| `Get-LidAction` | system | read AC/DC indices |
| `Set-LidAction` | system | write AC/DC indices, re-apply plan |

#### Argument parsing

Hand-rolled over `$args`. PowerShell's `param()` block **cannot** declare a
parameter named `-0`, so the `--0`-style flags cannot use the built-in binder. This
is a hard language constraint, not a preference.

Accepted forms — both single and double dash:

| Input | Result |
|---|---|
| `--status`, `-status`, `--s`, `-s` | print current setting |
| `--0`..`--3`, `-0`..`-3` | set that action |
| `--help`, `-help`, `-h`, `-?`, *(no args)* | print usage |
| anything else | usage to stderr, exit 2 |
| two action flags | error, exit 2 |

Matching is case-insensitive.

#### Reading current state

Read the registry, not `powercfg` text output:

```
HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\<scheme-guid>\
    4f971e89-eebd-4455-a8de-9e59040e7347\5ca83367-6e45-459f-a27b-476b1d01c936
```

Values `ACSettingIndex` and `DCSettingIndex`. Confirmed present on the target
machine.

Rationale: `powercfg -qh` label strings ("Current AC Power Setting Index") are
localized, so a text parser silently breaks if the Windows display language ever
changes. The registry is locale-independent. Extracting the scheme GUID from
`powercfg /getactivescheme` by regex is locale-safe because a GUID is a GUID.

Fallback chain: registry → parse `powercfg -qh` → report the value as *unknown*.
Never guess a value.

#### Writing

```
powercfg /setacvalueindex SCHEME_CURRENT SUB_BUTTONS 5ca83367-... <n>
powercfg /setdcvalueindex SCHEME_CURRENT SUB_BUTTONS 5ca83367-... <n>
powercfg /setactive SCHEME_CURRENT
```

Applies to the **active plan only**, matching what the Control Panel page does. The
machine has a single plan, so this is equivalent to "all plans" today and stays
predictable if plans are added later.

`/setactive` re-applies the scheme so the change takes effect immediately; without
it Windows can retain the old behavior until something else touches the plan.

After writing, re-read from the registry and print the verified state, so output
reflects reality rather than intent.

## Output

Status:

```
> lidaction --status
Lid close action  (plan: Balanced)
  Plugged in   Do nothing
  On battery   Do nothing
```

Set:

```
> lidaction --1
Lid close action set to: Sleep
  Plugged in   Sleep
  On battery   Sleep
```

Help:

```
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
```

## Error handling

| Case | Behavior | Exit |
|---|---|---|
| Success | normal output | 0 |
| Unknown flag | usage → stderr | 2 |
| Two action flags | error → stderr | 2 |
| `powercfg` non-zero exit | report its stderr | 1 |
| Value doesn't read back as written | error → stderr | 1 |
| Registry and `powercfg` read both fail | print `unknown` | 1 |
| `--2` while hibernation disabled | warn on stderr, proceed | 0 |

Exit codes: `0` ok, `1` runtime failure, `2` usage error.

The hibernate warning is live-checked via `powercfg /a`. Hibernate is available on
the machine today, but `powercfg /h off` would make `--2` silently behave closer to
sleep, so the warning earns its place. It warns rather than blocks — the setting is
still written correctly and becomes right again if hibernation is re-enabled.

Non-laptop hardware is not blocked; the setting is harmless on a machine with no lid.

## Testing

`test.ps1`, with **no Pester dependency**. The machine ships Pester 3.4.0, which is
awkward at mocking external executables, and requiring Pester 5 adds an install step
for a ~150-line wrapper.

**Pure-logic tests** — no system calls:

- Argument-parsing table: every accepted form maps to the right intent.
- Unknown flag and double-action-flag both produce usage errors.
- Index → name mapping for 0–3.
- Status formatting.

**Round-trip test** — real end-to-end, no mocks:

1. Record the current AC and DC indices.
2. For each of 0, 1, 2, 3: set it, then verify both AC and DC via the registry.
3. Restore the originally recorded values.
4. Verify the restore.

The restore step runs in a `finally` block so an interrupted run does not leave the
lid behavior changed.

## README

Per the repo's CLAUDE.md, `projects/lidaction/README.md` covers purpose, stack,
install, and usage. `projects/README.md` gains a `lidaction` entry in its Tools
list, replacing the current `_None yet._`.

## Open questions

None.
