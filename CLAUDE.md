# PC-CLI

A collection of independent command-line tools for my PC. This is **not** a single
application — each tool stands on its own, and tools do not import from each other.

## Structure

```
PC-CLI/
├── CLAUDE.md
└── projects/
    ├── <tool-name>/
    │   ├── README.md
    │   └── ...
    └── <tool-name>/
        ├── README.md
        └── ...
```

Every tool lives in its own folder under `projects/` and is fully self-contained:
its own source, its own dependency/lockfile, its own `README.md`.

## Adding a new tool

1. Create `projects/<tool-name>/` — lowercase kebab-case (e.g. `disk-report`, `wifi-switch`).
2. Add a `README.md` in that folder covering: purpose, stack, install, usage.
3. Keep everything the tool needs inside that folder.

## Language policy

The language and runtime are chosen **per project**, not repo-wide. There is no
shared stack, build system, or test runner at the root.

Before working in a tool folder, read that folder's `README.md` to learn its stack
and commands. Do not assume Python, Node, or PowerShell from the repo root.

## Scope

When working on one tool, confine changes to that tool's folder. Don't refactor or
"tidy up" sibling tools as a side effect.

## Environment

Windows 11, PowerShell is the primary shell. Prefer cross-platform-safe paths where
it costs nothing, but Windows-only is acceptable — these are tools for this PC.
