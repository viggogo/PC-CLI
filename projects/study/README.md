# study

Open the Literature study repo in VS Code from any terminal, without navigating
there by hand.

```
study --begin
```

## Stack

Windows PowerShell 5.1. No dependencies, no admin rights. Requires VS Code's `code`
command on your PATH.

## Install

```powershell
.\install.ps1
```

Adds this folder's `bin` subfolder to your user `PATH` and creates `.env` from
`.env.example`.

**Then fully quit and relaunch your terminal app.** A new *tab* is not enough: every
terminal inherits its environment from the host process, so a tab opened inside a VS
Code or Windows Terminal window that was already running still carries the old `PATH`
and will report `study : The term 'study' is not recognized`.

To pick up the change without restarting anything, refresh the current session:

```powershell
$env:Path = [Environment]::GetEnvironmentVariable('Path','Machine') + ';' +
            [Environment]::GetEnvironmentVariable('Path','User')
```

Running it a second time is safe: it reports the PATH entry as already present and
leaves an existing `.env` untouched.

Only `bin\study.cmd` is exposed, never the project folder itself — PowerShell
resolves `.ps1` from `PATH` ahead of `PATHEXT`, so a `study.ps1` on `PATH` would
shadow the shim and then be blocked by the execution policy, and `test.ps1` /
`install.ps1` would become global commands.

## Usage

```
study --begin       Open the Literature repo in VS Code
study --where       Show which path was resolved, and whether it exists
study --help        Show help
```

Both `--begin` and `-begin` forms work, `-b` and `-w` are short forms, and flags are
case-insensitive.

```
> study --begin
Opened in VS Code: C:\Users\viggo\Git Clone\Literature

> study --where
Repo path  C:\Users\viggo\Git Clone\Literature
Source     .env
Exists     True
```

Exit codes: `0` success, `1` runtime failure, `2` usage error. `--where` also exits
`1` when the folder it resolved does not exist.

## Configuration

The repo path comes from two places, in order:

1. `STUDY_REPO` in this folder's `.env`
2. otherwise the default constant at the top of `study.ps1`

The tool works with no `.env` at all. `study --where` prints which of the two won,
so you never have to guess.

The `.env` lives in **this** folder, not in the Literature repo. A config file inside
the target folder would be circular — the path is what finds that folder in the first
place. `.env` is gitignored; `.env.example` is the committed template.

## How it works

`--begin` resolves the path, then checks two things before doing anything:

- **The folder exists.** `code` on a missing path opens an empty window that looks
  like success, so a typo in `.env` would otherwise fail invisibly.
- **`code` is on your PATH.** Otherwise you get a clear message instead of a raw
  command-not-found.

Then it runs `code "<path>"`, which hands the folder to any running VS Code instance
and returns immediately. The success line prints the path actually opened, so the
output reflects reality rather than intent.

## Tests

```powershell
.\test.ps1
```

Fully inert — argument parsing, `.env` parsing, path resolution, and the exit codes
for help and usage errors. Nothing in the suite opens a window or touches your PATH.
`.env` fixtures are written to a temp folder and deleted in a `finally` block.

`--begin` is deliberately not covered: the only thing it does is launch VS Code, and
a test for it would leave an editor window open with nothing to restore. Verify it by
hand with the two commands under **Usage**.
