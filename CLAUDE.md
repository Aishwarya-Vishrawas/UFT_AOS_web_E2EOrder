# CLAUDE.md

Guidance for Claude Code and other AI assistants working in this repository.

## What this repository is

`AOS_web_E2EOrder` is a **Micro Focus Unified Functional Testing (UFT) 14.51 GUI test** — not
an application, library, or conventional codebase. It is a demo/reference test that exercises an
end-to-end order flow against **Advantage Online Shopping (AOS)**, a demo web store.

The whole repository *is* one UFT test. There is no build, no package manager, no test runner,
no dependency file, and no CI. The "source" is a mix of VBScript fragments and UFT's own binary
project files.

- Application under test: `http://nimbusserver.aos.com:8000/#/` (a locally hosted AOS instance;
  the object repository also records `http://nimbusserver:8000/#`). Public AOS lives at
  `advantageonlineshopping.com`.
- Authoring/execution tool: UFT 14.51 on Windows, Web add-in.
- Upstream: `github.com/Aishwarya-Vishrawas/UFT_AOS_web_E2EOrder`.

## The most important constraint

**Most of this repository cannot be edited outside UFT.** Binary formats — OLE compound
documents, Berkeley DB, legacy Excel — hold the test's real configuration. Hand-editing them
corrupts the test, and a corrupted `.tsr`/`.bdb`/`.mtr` is not recoverable by inspection.

| Path | Format | Editable by an assistant? |
|---|---|---|
| `Action*/Script.mts` | Text (UTF-8 BOM, VBScript) | **Yes**, with care — see below |
| `README.md`, `CLAUDE.md` | Markdown | Yes |
| `AOS_web_E2EOrder.usr`, `default.cfg`, `default.usp` | INI text | Yes, but UFT rewrites them |
| `.gitignore` | Text | Yes |
| `AOS_web_OR.tsr` | Berkeley DB (shared object repository) | **No — UFT only** |
| `Action*/ObjectRepository.bdb` | Berkeley DB (local object repository) | **No — UFT only** |
| `Test.tsp`, `Parameters.mtr`, `Action*/Resource.mtr` | OLE compound document | **No — UFT only** |
| `Default.xls`, `Configurations/*.xls` | Legacy BIFF Excel (data tables) | **No — UFT only** |
| `Action*/SnapShots/*` | JPEG + zlib-compressed HTML/XML | **No — generated** |

When a change requires touching any "UFT only" file, do not attempt it. Describe the exact
sequence of UFT IDE steps the user should perform instead.

Read-only inspection of the binaries is fine and often useful:
`strings -e l Test.tsp` (UTF-16), `strings -e l Action1/Resource.mtr` (yields readable XML),
`python3 -c "import xlrd; ..."` for the `.xls` data tables.

## Structure

```
AOS_web_E2EOrder.usr      Test definition: action name -> folder mapping, add-ins, file manifest
Test.tsp                  Test settings (browser startup, run/iteration, recovery, add-in config)
default.usp               Run logic: which top-level actions run, and how many iterations
default.cfg               Runtime config (timeouts, logging, think time, web emulation)
Default.xls               Design-time data table (Global + per-action sheets)
Parameters.mtr            Test-level input/output parameters
AOS_web_OR.tsr            SHARED object repository — all AOS web objects live here
Configurations/
  AOS_web_E2EOrder_default.xls   Run-time data table override used by test settings
Action0/ ... Action3/
  Script.mts              The VBScript steps (the only human-readable "code")
  Resource.mtr            Per-action metadata: OR associations, action params, dependencies
  ObjectRepository.bdb    Per-action LOCAL object repository
  SnapShots/              Active Screen captures recorded during authoring
```

### Actions

Folder names and logical names differ — always map through `AOS_web_E2EOrder.usr` `[Actions]`:

| Folder | Logical name | Reusable | Role |
|---|---|---|---|
| `Action0` | `Action0` | No | Driver. Calls the other three in order. |
| `Action1` | `CreatAcct` | Yes | Register a new account, set SafePay credentials. |
| `Action2` | `E2EOrder` | Yes | Browse speakers, add to cart, checkout, pay via SafePay. |
| `Action3` | `DelAcct` | Yes | Delete the account created in `CreatAcct`. |

`default.usp` runs **only `Action0`, one iteration**. `Action0/Script.mts` is the whole call chain:

```vbscript
RunAction "CreatAcct", oneIteration
RunAction "E2EOrder", oneIteration
RunAction "DelAcct", oneIteration
```

`Action0` also carries action parameters used for browser/device selection:
`Browser` (default `CHROME`), `device_model`, `device_manufacturer`, `device_ostype`.

## Working with `Script.mts` files

These are the only files you should normally edit. Conventions:

- **UTF-8 BOM at the start of every file.** Preserve it; UFT may misread the first step without it.
- Steps are flat VBScript, one object action per line, using the standard UFT hierarchy
  `Browser("Advantage Shopping").Page("Advantage Shopping").<Class>("<name>").<Method>`.
- **Do not invent object names.** Every quoted name must already exist in `AOS_web_OR.tsr`.
  Adding a new object means recording/adding it in the UFT Object Repository Manager — a task
  for the user, not an assistant. Confirm an object exists with:
  `strings -e l AOS_web_OR.tsr | grep -i "<name>"`
- **The `@@ hightlight id_;_..._;_script infofile_;_ZIP::ssfN.xml_;_` suffixes are UFT-generated
  Active Screen bindings**, not comments you wrote. (The misspelling of "highlight" is UFT's.)
  They tie a step to `SnapShots/ssfN.*`. Leave existing ones intact; never author new ones by
  hand. A step without the suffix is perfectly valid — several already lack it.
- Comment out steps with a leading `'`, matching the disabled sign-out steps at the end of
  `Action2/Script.mts`.
- Checkpoints (`.Check CheckPoint("QTY: 2")`) are defined inside the repository/`Resource.mtr`,
  not in the script. Referencing a checkpoint that does not exist fails at run time. Current
  checkpoints: `QTY: 2`, `$438.00` (Action2) and `yes` (Action3).

### Data and secrets

Values come from the run-time data table via `DataTable("<column>", dtGlobalSheet)`.

- `Default.xls` Global sheet: `uName`, `encPassword`, `emailAdrs` (one row: `User003`).
- `Configurations/AOS_web_E2EOrder_default.xls` Global sheet adds `Browser`, `device_model`,
  `device_manufacturer`, `device_ostype` and holds two rows — `Dwight`/`FIREFOX64` and
  `Angela`/`CHROME`. Test settings parameterize the browser with
  `DataTable("Browser", dtGlobalSheet)`, which is how "multiple browsers" works here.
- Passwords use `.SetSecure` with a **UFT-encoded** string (e.g. `5c76d2ed9e6d...`), produced by
  UFT's Password Encoder. These are obfuscated, not securely encrypted — treat them as demo
  credentials for a demo app. Do not put real credentials in the data tables, and do not try to
  generate encoded values by hand; `.Set` with a plaintext password is also acceptable for a
  demo but changes the recorded step semantics.
- `default.cfg` `[RtsUserInfo] Password` is a leftover LoadRunner-style field, not the AUT
  password.

### Object identification

Objects live in the **shared** repository `AOS_web_OR.tsr`, which all three functional actions
associate. Some descriptions use regular expressions so they survive changing usernames — e.g.
`Link("User My account")` is described as `.* My account My orders Sign out `. Prefer extending
that pattern over adding user-specific objects. Per-action `ObjectRepository.bdb` files also
exist; the shared repository is the one that matters.

## Setup and running (Windows + UFT only)

There is nothing to run on Linux/macOS or in CI. To execute the test:

1. Open `AOS_web_E2EOrder.usr` in UFT 14.51 (or later) with the **Web add-in** loaded.
2. **Re-associate the shared object repository.** The associations in `Action*/Resource.mtr` are
   stored as an absolute path from the original author's machine:
   `C:\Users\demo\IdeaProjects\AOS_web_E2EOrder\AOS_web_OR.tsr`. On any other machine the
   objects will not resolve. Per the README: Resources > Associate Repositories… > remove the
   existing entry > add this checkout's `AOS_web_OR.tsr` > associate it with all three actions
   (`CreatAcct`, `E2EOrder`, `DelAcct`).
3. Point the AUT URL at a reachable AOS instance. `Test.tsp` records
   `http://nimbusserver.aos.com:8000/#/`; change it under File > Settings > Web (or use
   `advantageonlineshopping.com`).
4. Ensure the browser named in the `Browser` data-table column is installed.
5. Run. Results land in a `Res*` folder, which `.gitignore` excludes.

The test is not idempotent by design: `CreatAcct` registers an account and `DelAcct` removes it,
so a partial run can leave an orphaned account behind. Re-running after a mid-test failure may
fail at registration if the username still exists — bump `uName` in the data table or delete the
account manually.

## Git conventions

- Work on the branch you were assigned; the default branch is `master`.
- `.gitignore` excludes `*.lck` (UFT lock files) and `/Res*` (result folders). Keep these
  ignored — never commit lock files or run results.
- **Almost every commit is a binary diff.** `git diff` shows `Bin NNN -> NNN bytes` and nothing
  more, and merge conflicts in `.tsr`/`.bdb`/`.mtr`/`.xls` cannot be resolved by hand — one side
  must be taken whole. Avoid concurrent edits to the same test in UFT.
- Simply opening and saving the test in UFT rewrites timestamps inside `Resource.mtr` and
  `Test.tsp`, producing noisy no-op binary diffs. Check `git status` before committing and drop
  binary changes you did not intend.
- Historical commit messages are terse (`UFT commit`, `minor updates`). Prefer descriptive
  messages, and state explicitly which action or resource changed since the diff cannot show it.

## Things assistants get wrong here

- Treating this as a normal codebase and looking for `package.json`, tests, or a build. There
  are none.
- Trying to "fix" the `@@ hightlight ...` suffixes or the `hightlight` typo. They are UFT's.
- Editing `.xls`, `.tsr`, `.bdb`, `.mtr`, or `.tsp` with text tools or a Python writer.
- Adding a step for an object that is not in the shared repository.
- Renaming a folder to match its logical action name — `Action1` **is** `CreatAcct`, and the
  mapping in the `.usr` plus the dependency paths inside every `Resource.mtr` would break.
- Dropping the UTF-8 BOM when rewriting a `Script.mts`.
