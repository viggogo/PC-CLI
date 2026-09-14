# study

Open a study repo in VS Code from any terminal, without navigating there by hand.
Two repos, split by what you do in them:

```
study --read     the Literature repo -- what you read
study --write    the latex repo      -- what you write
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
study --read        Open the Literature repo in VS Code
study --write       Open the latex repo in VS Code
study --where       Show both resolved paths, and whether they exist
study --help        Show help
```

Both `--read` and `-read` forms work and flags are case-insensitive. `-r` is short
for `--read` and `-w` for `--where`.

**`--write` has no short form on purpose.** `-w` already meant `--where` before the
latex repo existed, and moving it would silently turn a "print the paths" habit into
"open an editor" — a mistake you notice only after the window is up.

```
> study --read
Opened in VS Code: C:\Users\viggo\Git Clone\Literature

> study --write
Opened in VS Code: C:\Users\viggo\Git Clone\latex

> study --where
--read  (Literature)
  Repo path  C:\Users\viggo\Git Clone\Literature
  Source     .env (STUDY_READ_REPO)
  Exists     True

--write  (latex)
  Repo path  C:\Users\viggo\Git Clone\latex
  Source     .env (STUDY_WRITE_REPO)
  Exists     True
```

Exit codes: `0` success, `1` runtime failure, `2` usage error. `--where` also exits
`1` when **either** folder it resolved does not exist.

`--read` replaced the old `--begin`, which is now an unknown flag and exits `2`.

## Configuration

Each repo path comes from two places, in order:

1. `STUDY_READ_REPO` / `STUDY_WRITE_REPO` in this folder's `.env`
2. otherwise the matching default constant at the top of `study.ps1`

The two are resolved independently, so overriding one never moves the other. The tool
works with no `.env` at all. `study --where` prints which of the two sources won for
each repo, so you never have to guess.

The single-repo `STUDY_REPO` key is gone. It is not read as a fallback: with two
repos in play the name no longer says which one it means, and guessing would be worse
than ignoring it. If an old `.env` still has it, replace it with the two keys above —
`--where` will show `default in study.ps1` until you do.

The `.env` lives in **this** folder, not in either repo. A config file inside a target
folder would be circular — the path is what finds that folder in the first place.
`.env` is gitignored; `.env.example` is the committed template.

## How it works

`--read` and `--write` run the same routine against different paths. It resolves the
path, then checks two things before doing anything:

- **The folder exists.** `code` on a missing path opens an empty window that looks
  like success, so a typo in `.env` would otherwise fail invisibly.
- **`code` is on your PATH.** Otherwise you get a clear message instead of a raw
  command-not-found.

Then it runs `code "<path>"`, which hands the folder to any running VS Code instance
and returns immediately. The success line prints the path actually opened, so the
output reflects reality rather than intent.

Adding a third repo is one entry in `Get-RepoSpec` (label, `.env` key, default) plus
one flag in `Get-StudyIntent` — nothing else branches per repo.

## Tests

```powershell
.\test.ps1
```

Fully inert — argument parsing, `.env` parsing, path resolution, and the exit codes
for help and usage errors. Nothing in the suite opens a window or touches your PATH.
`.env` fixtures are written to a temp folder and deleted in a `finally` block.

`--read` and `--write` are deliberately not covered: the only thing they do is launch
VS Code, and a test for either would leave an editor window open with nothing to
restore. Verify them by hand with the commands under **Usage**.
