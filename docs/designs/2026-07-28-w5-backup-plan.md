# W5 - Protect & Back Up

> **For agentic workers:** REQUIRED SUB-SKILL: use `superpowers:subagent-driven-development`
> to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship `PkgBackup`, four actions in the `backup` domain (already active in contract
section 7): `backup.timestamped_copy`, `backup.snapshot`, `backup.month_end_version`,
`backup.archive_output`. All four write to the file system; `archive_output` also moves one.
`backup.snapshot` is the load-bearing action of the wave - it is what the design doc names as
the unblocker several *Deferred* catalog items are waiting on - so this plan is explicit and
loud about exactly how much of that promise it delivers and how much it does not.

**Architecture:** One new package, one new module, no Core behaviour change beyond promoting
one existing helper. Three of the four actions (`timestamped_copy`, `snapshot`,
`month_end_version`) are the same underlying move - copy this workbook, as it stands right
now, to a computed path that does not yet exist - with a different destination-naming rule
each. They share two small private helpers (`PkgBackup_EnsureFolder`, `PkgBackup_Ext`) but do
their own precondition checks and write their own refusal text, because the earliest-cheapest-
failure rule and the per-action wording both matter more than collapsing three call sites into
one. The fourth (`archive_output`) is the wave's only genuinely destructive filesystem action -
it moves a file that already exists - and reuses `Core_FindFile` (existing) and a newly
promoted `Core_AlreadyOpen` (moved from `PkgImport`) rather than inventing either.

**Tech stack:** VBA (Excel/Windows, one concatenated `.bas`), PowerShell for the build, Node
for the syntax check. No test framework - see *Verification*.

**Runs after W4** (`2026-07-28-w4-summarize-enrich-plan.md`) and **before W6**. Sequencing:
`2026-07-28-track2-completion-plan-of-record.md` section 2. Written against the Core API from
`2026-07-28-core-helper-promotion-plan.md` - every `Core_*` helper this plan calls is assumed
already merged to `main` by the time this wave branches.

**Also depends on `2026-07-28-w3-format-export-plan.md` having landed**, because Task 5 calls
`Core_SafeFileName`, which that plan adds to Core (its Task 2, Step 2). W3 ships as phase 4 and
this wave as phase 6, so `Core_SafeFileName` is expected to already be in `main` by the time this
wave branches - but if it is not there when this plan is picked up, W3 has not landed; reconcile
before starting, do not write a local substitute.

---

## Why this wave matters more than its size suggests

The catalog (`docs/macro-catalog-plug-and-play.md` line 231) calls `backup.snapshot` "the
unblocker for several *Deferred* destructive items above, since it supplies the missing undo
story." That claim needs checking, not repeating, so this plan checked it.

**Grepping the catalog for `Deferred` finds five items, not "several" left vague:**

| Item | Line | Blocked on |
|---|---|---|
| Pull data from a closed workbook | 51 | An undefined failure mode when the source is missing, locked, or has moved. |
| Delete all hidden rows / columns | 85 | Destructive with no preview - the user cannot see what they are about to lose. |
| Hide all sheets except the active one | 211 | Destructive to workbook state with no preview and no one-click undo. |
| Delete all empty sheets | 217 | Destructive with no preview; "empty" is easy to get wrong. |
| Loop a folder of workbooks and process each | 252 | A defined per-file continue/abort contract and failure reporting. |

**Only three of the five actually name a missing-undo problem** (rows 2-4). The other two are
blocked on something a restore point cannot fix: a closed-workbook read failure mode, and a
chain-contract question. Snapshot is irrelevant to those two regardless of how good its restore
story is. This is stated here, once, so a later reader does not have to re-derive it from a
looser "several" and conclude this wave unblocked five items when it is at most three - and, per
the next section, actually zero.

**And of those three, this wave does not unblock any of them, and says so instead of implying
otherwise.** See "Restore path" below.

---

## Decisions taken in this plan

| # | Question | Decision |
|---|---|---|
| 1 | Where does a snapshot live - a hidden sheet in this workbook, or a separate file? | **A separate file: a full `SaveCopyAs` copy of the whole workbook.** A hidden sheet is recoverable in one click and travels with the workbook, but it dies with the workbook - and the entire reason a restore point exists is the scenario where the workbook itself is the thing that broke (a crash mid-write, a macro that corrupts a structure, a bad save). A sheet-shaped snapshot also only ever captures one sheet; the actions it is meant to guard (Import, Clear, and by extension anything W5 cannot see the shape of yet) can touch more than one. A separate file is heavier to produce and has no one-click recovery, but it is the only shape that survives the failure it exists to protect against. |
| 2 | Does this wave provide a restore path? | **No, and that is stated loudly rather than implied away.** None of the four frozen action ids for W5 is a restore action, and no restore id appears anywhere in the catalog. `backup.snapshot` supplies the save half of "an undo story"; the restore half does not exist in this codebase. **The three Deferred items that cite a missing-undo blocker (hide-all-except-active, delete-hidden-rows/cols, delete-empty-sheets) stay Deferred after this wave ships.** Unblocking any of them needs a `backup.restore_*`-shaped action - itself a destructive overwrite of live data, needing its own confirm-and-preview design - which is not in this wave's frozen scope and is not proposed here. |
| 3 | Chosen restore workaround for the user, given no restore action exists | The completion message for `backup.snapshot` says explicitly: "There is no automatic restore for this snapshot. To use it, close this workbook and open the snapshot file directly, or copy what you need from it by hand." A safety feature that implies a capability it does not have is worse than one that is honest about its limit. |
| 4 | Snapshot naming and retention | Snapshots are named `SNAPSHOT_<yyyy-mm-dd_HHnnss>[_<label>]<ext>` (timestamp first, so the newest sorts last by name too, though pruning below does not rely on that) in `<root>\Backup\Snapshots\`. `k_backup_snapshot_keep` (default 20) caps how many are kept; each successful snapshot prunes its own folder down to that count by file modified time, oldest first. The prune touches only files that start with `SNAPSHOT_` inside that one subfolder - never a user's own file, never a plain timestamped backup. |
| 5 | One shared `k_backup_folder` field, or a field per action? | **One shared root.** All four actions write under the same "where does this deployment keep its backups" concept; the specific subfolder name (`Backup`, `Backup\Snapshots`, `Archive\yyyy_mm`) is a package convention, not something worth a Setup-tab field each. Splitting it would add four fields to answer one question. |
| 6 | Create a missing backup/archive folder, or refuse? | **Create it.** `k_import_folder` (existing) refuses on a missing folder because that folder is meant to already hold the user's source files - an empty freshly-made one is never useful, and creating it would hide a real misconfiguration. A backup or archive folder is the opposite: it is this package's own workspace, and refusing would make every deployment's very first backup fail with "folder not found" for a folder this action is supposed to own. `Core_FolderExists` is checked first in both cases; only the response to "false" differs, and it differs for a stated reason in each package, not by accident. |
| 7 | Never overwrite | Every one of the four computes its destination path, then checks `Dir(destPath)` **before** `Core_SpeedOn` runs and before anything is written. A hit refuses with the path named and changes nothing. This is the earliest-cheapest-failure ordering `PkgImport` already uses for its destination-sheet check. |
| 8 | Prior month end, exactly | `refDate = Date + Core_CfgLong(ctx, "k_backup_offset")`, then `priorMonthEnd = DateSerial(Year(refDate), Month(refDate), 1) - 1`. This depends only on `refDate`'s year and month, never its day - the 1st and the 28th of the same month give the same answer, which a fixed day-count shortcut ("today minus 30") does not, because that lands in a different month depending on where in the current month "today" falls and gets short months wrong. January needs no special case: `DateSerial(y, 1, 1) - 1` rolls back to 31 December of `y-1` on its own. `k_backup_offset` (default 0) exists for the same reason `k_export_offset` does: a deployment that runs its close a few days after month end can backdate the reference date so "prior month" tracks the close period, not the calendar day the macro happens to run. This is the same day-offset convention `PkgImportExport_Export` uses, not a second one. |
| 9 | How does `archive_output` find "the generated file"? | Reuses the existing `file` field kind and `Core_FindFile` exactly as `k_master_prefix` already does: newest file in **this workbook's own folder** whose name starts with a configured prefix. Per contract section 3, a `file` kind is always resolved against the deployment's data location - `k_backup_folder` is deliberately **not** consulted for the search, only for where the archived copy lands. `Core_FindFile` only matches `*.xls*`, so this action only ever archives an Excel file - stated as a known limitation, not fixed here. |
| 10 | Snapshot/backup/month-end file format | Every copy is written with **this workbook's own extension** (`PkgBackup_Ext`, reads it off `ThisWorkbook.name`), never a hardcoded `.xlsm`. A macro-enabled workbook copied out as `.xlsx` silently loses every macro, which is the opposite of what a restore point is for. |
| 11 | Chain-quiet behaviour | No action in this package calls `MsgBox` directly. Every outcome goes through `Core_Refuse`, `Core_Done`, or `Core_Fail`, all three already quiet-aware from phase 1 - nothing new needed here to make these four safe as chain steps. `backup.snapshot` as step 1 of a chain that also runs something destructive is the wave's flagship usage: if it refuses (unsaved workbook, a same-second collision), the phase-1 chain-halt rule stops the chain before the destructive step runs, which is the entire safety story this action is bought for. |

---

## Escalations to Core

**One helper is promoted in this wave's diff: `PkgImport_AlreadyOpen` becomes `Core_AlreadyOpen`
(Task 2).** The phase-2 promotion plan deliberately left it in `PkgImport`, calling W5's need
for it "speculative," and named the correct move if that turned out wrong: "promote it then, in
W5's own diff, and say so in that PR." It turned out not speculative. `backup.archive_output`
moves a file with `FileSystemObject.MoveFile`, and moving a file Excel currently has open fails
partway and can leave neither a clean source nor a clean destination - exactly the failure mode
`PkgImport_AlreadyOpen`'s own header describes for the combine actions. The probe (is a workbook
of this name open, and is it *this* file or a same-name collision) is identical; only the policy
built on top of it differs; `PkgImport` opens-in-place, `PkgBackup` refuses outright, because
this wave has no legitimate reason to use an already-open file in place of the one on disk.

**`PkgImport_WorkbookPaths` is not promoted and not needed.** No W5 action loops a folder of
workbooks; `archive_output` targets exactly one file via `Core_FindFile`. The phase-2 decision
to leave it in `PkgImport` stands unchanged.

**Nothing else escalates.** `PkgBackup_EnsureFolder`, `PkgBackup_Ext`, and
`PkgBackup_PruneSnapshots` are used only by this package's own four actions; no other wave's
plan names a consumer for any of them, so by the same YAGNI reasoning phase 2 used, they stay
package-private and prefixed.

**One helper is removed, not added: this package declares no package-private "safe file name"
function.** An earlier revision of this plan declared its own copy of that job, but
`2026-07-28-w3-format-export-plan.md` (Task 2, Step 2) already adds `Core_SafeFileName` to Core
for the identical "make this a legal Windows file name" job - flagged there for this exact
wave at the time it was written. Keeping a second, package-private copy here would be the
cross-package duplicate CLAUDE.md rule 4 forbids, so every call in this plan uses
`Core_SafeFileName` instead. This wave now depends on W3 having landed first; see the
dependency note below and the gap noted in "Unknowns and risks."

---

## Global constraints

Copied verbatim from `2026-07-28-track2-completion-plan-of-record.md` section 4, so this plan
does not need to be read alongside that one.

- **`SCHEMA_VER` stays `"v2.0"`.** If a build ever reports `v3.0`, something is wrong.
- **Action ids are immutable**, two-segment `<domain>.<operation>`, lowercase ASCII and
  underscores, each segment starting with a letter. **Never camelCase. No alias table.**
  `build.ps1` parses the vocabulary table out of contract section 7 and fails on an unknown
  domain, so a new domain must be added to that table in the same commit or earlier - not needed
  here, since `backup` is already `active`.
- **Shared helpers live ONLY in Core. Cross-package calls are forbidden.**
- **Pure ASCII in every `.bas`**; `Option Explicit`; max 25 line continuations per statement;
  no module-level declaration after the first procedure (in the assembled bundle); no VBA
  reserved word as an identifier.
- **Every action entry point must be `Public Sub <Name>(ctx As Object)`** - one `Object`
  parameter, nothing else. `build.ps1` fails the build otherwise, because `Core_Run` dispatches
  through `Application.Run`.
- **Every package private carries its package prefix** (`PkgBackup_*`), enforced by
  `build.ps1`'s prefix guard since phase 2.
- **TDD is off.** VBA has no headless runner and Excel cannot be driven on CI. **Do not propose
  a test suite.** Verification is `pwsh build/build.ps1`, `node scripts/vba-syntax-check.mjs
  dist/MacrosBundle.bas`, and deliberate injections.
- **`dist/` is gitignored** - never `git add -f` the bundle. `.vscode/` stays untracked;
  do not commit it and do not remove it.
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.
- **No manual gate document.** Run-time-only claims append to
  `docs/designs/2026-07-28-deferred-validation-log.md` (Task 9).

---

## Verification, since TDD is off

1. `pwsh build/build.ps1` - exit 0, `Core + <n+1> package(s)` (one more than the W4 baseline),
   balanced `Sub`/`Function` counts, no id-gate or prefix-gate failures.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` - no diagnostics.
3. **Two deliberate injections** (Task 7's step, kept there since that is where the guards being
   tested already exist in full): an unprefixed private in `PkgBackup.bas`, to prove the phase-2
   prefix guard fires against a *new* module and not just the three it was written against; and
   a malformed action id, to prove the id gate's vocabulary/shape check fires against `backup`
   specifically rather than only the domains already exercised.
4. **A grep proving `CoreCandidates` was actually edited** (Task 8) - the one change none of the
   build gates can see, per the plan-of-record's own finding that a missing candidate "builds,
   bundles, passes every gate, and never runs."
5. Everything genuinely run-time-only (a real `SaveCopyAs` producing an openable file with its
   macros intact; folder auto-creation on a real filesystem; the prune deleting the right files
   and no others; `MoveFile` actually moving across drives) is not claimed here - it is logged in
   Task 9 for the single end-to-end validation, per plan-of-record section 3.5.

---

## File structure

| File | Change |
|---|---|
| `src/packages/PkgBackup.bas` | **New.** Manifest, four actions, four private helpers. |
| `src/core/Core.bas` | `Core_AlreadyOpen` added (promoted from `PkgImport`). |
| `src/packages/PkgImport.bas` | `PkgImport_AlreadyOpen` deleted; its one call site renamed. |
| `src/core/Core.bas` (`CoreCandidates`) | `c.Add "Backup"`. |
| `spec/contract.md` | v2.8.0. New section 8 subsection `Package backup`; one section 10 bullet. |
| `docs/designs/2026-07-28-deferred-validation-log.md` | + W5's run-time-only claims. |
| `CHANGELOG.md`, `VERSION`, `docs/code-map.md` | v0.16.0. |

---

### Task 1: Contract v2.8.0

**Owner: Opus, written directly.** Contract work.

**Files:** Modify `spec/contract.md` - header, section 8, section 10.

**Interfaces:** Produces the normative text Tasks 3-7 implement. Consumes nothing.

- [ ] **Step 1: Bump the version header**

Line 3. Expected prior value per the plan-of-record's table is `2.7.0` (W4's bump); if it reads
something else, stop and reconcile before continuing rather than assuming the table is current.

```
**Contract version: 2.8.0** (2026-07-28)
```

- [ ] **Step 2: Confirm the domain needs no section 7 edit**

```bash
grep -n '| `backup` | active' spec/contract.md
```

Expected: one line, already present. This wave adds nothing to section 7 - the domain was
declared active ahead of the wave that would use it, exactly as `spec/contract.md`'s own note
that domain and module are decoupled predicts.

- [ ] **Step 3: Add the section 8 subsection**

Insert as a new subsection at the end of section 8's existing package subsections (after
whichever package W4 most recently added):

````markdown
### Package `backup` — "Protect & Back Up" (v1.0.0, contract v2.8.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_backup_folder` | text | simple | *(blank)* | Root folder for every action in this package. **Blank means the deployment's data location** (section 9), which is the common case. Backups land in a `Backup` subfolder here, snapshots in `Backup\Snapshots`, and archived files in `Archive\yyyy_mm`. `usedBy` all four actions. |
| `k_backup_base` | text | simple | `Backup` | Base name used inside a timestamped or month-end file name. `usedBy` `backup.timestamped_copy`, `backup.month_end_version`. |
| `k_backup_offset` | int | advanced | `0` | Days to shift today's date before computing the prior month end. Same convention as `k_export_offset` (package `importexport`): a negative value looks back further than today. `usedBy` `backup.month_end_version`. |
| `k_backup_snapshot_label` | text | simple | *(blank)* | Short tag appended to a snapshot's file name, e.g. `before Import`. Blank means no tag. `usedBy` `backup.snapshot`. |
| `k_backup_snapshot_keep` | int | advanced | `20` | How many of the most recent snapshots to keep. On each new snapshot, older ones past this count are deleted automatically; nothing else in the folder is touched. `usedBy` `backup.snapshot`. |
| `k_backup_archive_prefix` | file | simple | *(blank)* | Name-prefix of the generated file this action relocates, resolved the same way every `file` field is (section 3): the newest file in the deployment's data location whose name starts with this text. `usedBy` `backup.archive_output`. |

**Actions**

| id | caption | semantics |
|---|---|---|
| `backup.timestamped_copy` | Backup Now (Timestamped Copy) | Save a full copy of this workbook, exactly as it stands at the moment the action runs (including any edits not yet saved to the main file), into `<root>\Backup\` with a datetime-stamped name. This workbook's own name, path, and unsaved-changes state MUST be unaffected. MUST refuse rather than overwrite if the computed destination name already exists. |
| `backup.snapshot` | Snapshot Before a Risky Step | As `backup.timestamped_copy`, into `<root>\Backup\Snapshots\`, named with a full timestamp and an optional configured label. After a successful save, prune that folder to the configured keep-count by file modified time, oldest first, touching only files this action's own naming pattern produced. **This action provides no restore.** A conforming implementation MUST NOT claim or imply an automatic restore capability for the file it writes. |
| `backup.month_end_version` | Save Month-End Version | As `backup.timestamped_copy`, named for the last day of the calendar month before the reference date's month (`k_backup_offset` days from today), independent of which day of that month the action runs on. |
| `backup.archive_output` | Archive a Generated Output | Locate the newest file matching `k_backup_archive_prefix` (the `file` kind rule). MUST refuse if that file is currently open in this Excel instance rather than attempting to move it. Move it into `<root>\Archive\yyyy_mm\` (the current date's year and month), creating that folder if it does not exist. MUST refuse rather than overwrite if a file of that name already exists at the destination. |
````

- [ ] **Step 4: Add the section 10 entry**

Insert as the first versioned bullet in section 10, above whatever entry W4 added at v2.7.0:

````markdown
- **v2.8.0 (additive):** a new package `backup` with six fields and four actions (§8). The
  `backup` domain itself is not new - section 7 already listed it `active` ahead of any package
  using it - so this bullet adds only the package, not the domain. **Additive by the two-axes
  rule:** no existing `key`, action `id`, `kind`, or encoding changes meaning; every prior stored
  configuration keeps working exactly as it did. `SCHEMA_VER` does **not** move.
````

- [ ] **Step 5: Verify**

```bash
grep -n 'Contract version: 2.8.0' spec/contract.md
grep -c '### Package `backup`' spec/contract.md
pwsh build/build.ps1
```

Expected: version line present; exactly `1`; build exits 0 (a contract edit that breaks the
domain table or the action-id table fails the build immediately, before `PkgBackup.bas` even
exists to be checked against it).

- [ ] **Step 6: Commit**

```bash
git add spec/contract.md
git commit -m "spec: contract v2.8.0 - the backup package"
```

---

### Task 2: Core gains `Core_AlreadyOpen`

**Owner: Opus, written directly.** Core edit, and the reasoning for promoting it now (rather
than leaving it, per phase 2's default) needs to be right, not mechanical.

**Files:** Modify `src/core/Core.bas` (grep `docs/code-map.md` for `Core_FolderExists` and
insert immediately after it - do not trust a fixed line number here, the map will have moved
by the time this wave lands). Modify `src/packages/PkgImport.bas`.

**Interfaces:**
- Consumes: `Core_FileName` (existing).
- Produces: `Core_AlreadyOpen(ByVal fullPath As String) As Workbook`. Task 7 depends on it.

- [ ] **Step 1: Move the function, rename it, make it Public**

The body is a verbatim move from `PkgImport_AlreadyOpen` (`PkgImport.bas:652`), except the
`Err.Raise` source string and the header comment's first line, which now explains a second
consumer:

```vb
' The workbook at this path IF it is already open in this Excel instance, else
' Nothing. PkgImport uses what comes back IN PLACE and does not close it; PkgBackup
' (archive_output) uses it only to REFUSE - it has no legitimate reason to move a
' file out from under an open handle. The probe is shared; the policy on top of it
' is each caller's own.
'
' This exists because Workbooks.Open is not a question - it is an answer. Handed a
' path that is already open it returns the open workbook, ReadOnly:=True and all,
' complete with whatever the user has typed and not yet saved. Closing that with
' SaveChanges:=False discards their work, and under Core_SpeedOn DisplayAlerts is
' False, so the prompt Excel would normally raise ("reopening will discard your
' changes") never reaches them.
'
' Probed by FILE NAME, because that is how the Workbooks collection is keyed - there
' is no lookup by path. The name alone is not proof of identity, so the path is
' confirmed afterwards. A same-name / different-path collision is raised rather than
' guessed at.
Public Function Core_AlreadyOpen(ByVal fullPath As String) As Workbook
    Dim wb As Workbook, nm As String, openPath As String
    nm = Core_FileName(fullPath)
    If Len(nm) = 0 Then Exit Function

    On Error Resume Next
    Set wb = Workbooks(nm)
    On Error GoTo 0
    If wb Is Nothing Then Exit Function          ' not open: the caller opens it itself

    On Error Resume Next
    openPath = wb.FullName
    On Error GoTo 0
    If StrComp(openPath, fullPath, vbTextCompare) = 0 Then
        Set Core_AlreadyOpen = wb
        Exit Function
    End If
    Err.Raise vbObjectError + 1025, "Core_AlreadyOpen", _
        "A different workbook named '" & nm & "' is already open (" & openPath & "). " & _
        "Excel cannot have two workbooks of the same name open at once. Close that one, " & _
        "then run this again."
End Function
```

- [ ] **Step 2: Delete the original and rename its one call site**

In `PkgImport.bas`, delete the `Private Function PkgImport_AlreadyOpen` block (`:652-673`).
At its one call site (`PkgImport_CombineEngine`, `:211`):

```vb
' was:  Set wbSrc = PkgImport_AlreadyOpen(CStr(one))
Set wbSrc = Core_AlreadyOpen(CStr(one))
```

- [ ] **Step 3: Prove the old name is gone**

```bash
grep -nE '(^|[^_[:alnum:]])PkgImport_AlreadyOpen\b' src/packages/*.bas
```

Expected: no output.

- [ ] **Step 4: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics. `Function` count unchanged overall (one moved, not added);
`PkgImport.bas`'s own count falls by one, `Core.bas`'s rises by one.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas src/packages/PkgImport.bas
git commit -m "refactor(core): promote PkgImport_AlreadyOpen to Core_AlreadyOpen for backup.archive_output"
```

---

### Task 3: `PkgBackup.bas` - module skeleton, manifest, shared helpers

**Owner: delegate to codex (sonnet).** New file, but every line below is exact.

**Files:** New `src/packages/PkgBackup.bas`.

**Interfaces:**
- Consumes: `Core_DeclarePackage`, `Core_AddSection`, `Core_AddField`, `Core_AddAction`,
  `Core_ResolveFolder`, `Core_FolderExists`, `Core_FileName` (all existing).
- Produces: `PkgBackup_Manifest`, `PkgBackup_EnsureFolder`, `PkgBackup_Ext`,
  `PkgBackup_PruneSnapshots`. Tasks 4-7 depend on all four.

- [ ] **Step 1: Header and manifest**

```vb
Attribute VB_Name = "PkgBackup"
Option Explicit

'================================================================
' PkgBackup - restore points and archival. Wave 5 of the catalog
' build-out (docs/macro-catalog-plug-and-play.md, section 9).
'
' Four actions, three of which are the same underlying move - copy
' THIS workbook, exactly as it stands right now, to a path that
' does not yet exist - with a different destination-naming rule
' each:
'   backup.timestamped_copy    a manual "back this up right now"
'   backup.snapshot             an automatic pre-risky-step restore point
'   backup.month_end_version    named for the last closed calendar month
' The fourth is the only genuinely destructive filesystem action in
' this package - it MOVES a file that already exists:
'   backup.archive_output       relocate a generated output into Archive\yyyy_mm
'
' backup.snapshot writes a restore point but this package has no
' restore action. See the plan this shipped from
' (docs/designs/2026-07-28-w5-backup-plan.md) for why that is
' deliberate and which catalog items it does NOT unblock.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgBackup_Manifest(reg As Object)
    Core_DeclarePackage reg, "backup", "Protect & Back Up", "1.0.0"

    Core_AddSection reg, "BACKUP & SNAPSHOT"
    Core_AddField reg, "k_backup_folder", "Folder to write backup and snapshot files into", "text", "", _
        "simple", _
        "Leave blank to use the folder this workbook is saved in - that is the usual answer. " & _
        "Otherwise a full path, e.g. C:\Close\Backups. Backups go in a 'Backup' subfolder here, " & _
        "snapshots in 'Backup\Snapshots', and archived files in 'Archive\yyyy_mm'.", _
        "backup.timestamped_copy,backup.snapshot,backup.month_end_version,backup.archive_output"
    Core_AddField reg, "k_backup_base", "Base name for a backup file", "text", "Backup", _
        "simple", _
        "Used to build the saved file's name, e.g. 'Backup 2026-07-28_143059.xlsm'.", _
        "backup.timestamped_copy,backup.month_end_version"
    Core_AddField reg, "k_backup_offset", "Date shift for month-end (days)", "int", "0", _
        "advanced", _
        "Days to shift today's date before finding the calendar month it falls in, then " & _
        "stepping back one month for the prior month end. 0 = use today as-is. Same " & _
        "convention as Export Dashboard's date shift.", _
        "backup.month_end_version"
    Core_AddField reg, "k_backup_snapshot_label", "Snapshot label", "text", "", _
        "simple", _
        "A short tag added to a snapshot's file name, e.g. 'before Import'. Leave blank for no tag.", _
        "backup.snapshot"
    Core_AddField reg, "k_backup_snapshot_keep", "Snapshots to keep", "int", "20", _
        "advanced", _
        "How many of the most recent snapshots to keep. Older ones are deleted " & _
        "automatically each time a new snapshot is taken. Nothing else in that folder is touched.", _
        "backup.snapshot"

    Core_AddSection reg, "ARCHIVE"
    Core_AddField reg, "k_backup_archive_prefix", "Generated output to archive (name starts with...)", "file", "", _
        "simple", _
        "Name-prefix of the file this action relocates. Matched in the folder this workbook " & _
        "is saved in, same rule as any other 'file' field: the newest file whose name starts " & _
        "with this text.", _
        "backup.archive_output"

    Core_AddAction reg, "backup.timestamped_copy", "Backup Now (Timestamped Copy)", "PkgBackup_TimestampedCopy", _
        "Saves a full copy of this workbook, exactly as it stands right now, into a 'Backup' " & _
        "subfolder with a datetime stamp in the file name. This workbook itself is untouched.", _
        "Protect & Back Up"
    Core_AddAction reg, "backup.snapshot", "Snapshot Before a Risky Step", "PkgBackup_Snapshot", _
        "Saves a full copy of this workbook as a restore point, meant to run right before a " & _
        "step that changes data. There is no automatic restore - open the snapshot file " & _
        "directly to recover from it.", _
        "Protect & Back Up"
    Core_AddAction reg, "backup.month_end_version", "Save Month-End Version", "PkgBackup_MonthEndVersion", _
        "Saves a full copy of this workbook named for the most recently closed calendar " & _
        "month, for finance packs that keep one file per close.", _
        "Protect & Back Up"
    Core_AddAction reg, "backup.archive_output", "Archive a Generated Output", "PkgBackup_ArchiveOutput", _
        "Moves the newest file matching the configured prefix into an 'Archive\yyyy_mm' " & _
        "subfolder. Refuses if the file is open in Excel or a file of that name is already " & _
        "archived there.", _
        "Protect & Back Up"
End Sub
```

- [ ] **Step 2: Shared private helpers**

Placed after the manifest, ahead of any action - every action in this file calls at least one
of these:

```vb
'================================================================
' Package-private helpers
'================================================================

' Creates folderPath (and any missing parent) if it does not already exist. Neither
' MkDir nor FileSystemObject.CreateFolder can create more than one new level at a
' time - CreateFolder("A\B") fails if A does not already exist - so a fresh
' deployment's very first backup run needs each level built up one at a time, or an
' action that is supposed to OWN this folder fails on the folder it owns.
Private Sub PkgBackup_EnsureFolder(ByVal folderPath As String)
    Dim fso As Object, sep As String, p As String, parts() As String, i As Long, built As String
    sep = Application.PathSeparator
    p = folderPath
    If Right$(p, 1) = sep Then p = Left$(p, Len(p) - 1)
    If Core_FolderExists(p) Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    parts = Split(p, sep)
    built = parts(LBound(parts))                  ' e.g. "C:" - the drive root, always exists
    For i = LBound(parts) + 1 To UBound(parts)
        built = built & sep & parts(i)
        If Not fso.FolderExists(built) Then fso.CreateFolder built
    Next i
End Sub

' This workbook's own file extension, INCLUDING the dot, e.g. ".xlsm". A backup or
' snapshot must be written in the SAME format the live workbook is in - a macro-
' enabled book copied out as .xlsx would silently drop every macro, which defeats
' the point of a restore point.
Private Function PkgBackup_Ext() As String
    Dim nm As String, dot As Long
    nm = ThisWorkbook.name
    dot = InStrRev(nm, ".")
    If dot > 0 Then PkgBackup_Ext = Mid$(nm, dot) Else PkgBackup_Ext = ".xlsm"
End Function

' NOTE: a package-private "user-typed label reduced to something safe inside a
' Windows file name" helper used to live here. It is gone - Task 5 now calls
' Core_SafeFileName (added to Core by 2026-07-28-w3-format-export-plan.md, Task 2,
' Step 2) instead, so this package does not carry a second copy of the identical
' job. See "Escalations to Core" above and "Unknowns and risks" below for the one
' behavior gap that removal opens up (no length cap on the token).

' Deletes the oldest files matching "SNAPSHOT_*" in folderPath, keeping the newest
' keepCount by file MODIFIED time (not by parsing the name - the label segment
' makes lexical order unreliable). Only files this action's own naming pattern
' would have produced are ever touched; nothing else in the folder is examined.
'
' A single locked or in-use old snapshot must not fail the run that just wrote a
' GOOD new one, so Kill runs under Resume Next - a prune that skips one file this
' time is a smaller problem than a snapshot action that reports failure after it
' already did the part that matters.
Private Sub PkgBackup_PruneSnapshots(ByVal folderPath As String, ByVal keepCount As Long)
    Dim f As String, names() As String, times() As Date, n As Long, i As Long, j As Long
    Dim tmpS As String, tmpD As Date
    If keepCount < 1 Then keepCount = 1

    ReDim names(1 To 500)
    ReDim times(1 To 500)
    n = 0
    f = Dir(folderPath & "SNAPSHOT_*")
    Do While Len(f) > 0
        n = n + 1
        If n > UBound(names) Then Exit Do          ' hard cap: never scan without bound
        names(n) = folderPath & f
        times(n) = FileDateTime(folderPath & f)
        f = Dir
    Loop
    If n <= keepCount Then Exit Sub

    ' Insertion sort, oldest first. n is small (bounded by keepCount plus however
    ' many snapshots ran since the last prune), so O(n^2) costs nothing here and
    ' needs no extra library.
    For i = 2 To n
        tmpS = names(i): tmpD = times(i)
        j = i - 1
        Do While j >= 1
            If times(j) <= tmpD Then Exit Do
            names(j + 1) = names(j): times(j + 1) = times(j)
            j = j - 1
        Loop
        names(j + 1) = tmpS: times(j + 1) = tmpD
    Next i

    On Error Resume Next
    For i = 1 To n - keepCount
        Kill names(i)
    Next i
    On Error GoTo 0
End Sub
```

- [ ] **Step 3: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: **FAIL, deliberately, at this point.** `PkgBackup_Manifest` declares four actions
whose entry points (`PkgBackup_TimestampedCopy` etc.) do not exist yet - the id gate's
"no source declares a dispatchable handler" check should name all four. If the build instead
reports something about a duplicate name or a prefix violation, stop and fix that before Task 4,
since those two guards must already be clean before the entry-point gate becomes the only
failure.

- [ ] **Step 4: Commit**

```bash
git add src/packages/PkgBackup.bas
git commit -m "feat(backup): PkgBackup module skeleton, manifest, and shared helpers"
```

---

### Task 4: `backup.timestamped_copy`

**Owner: delegate to codex (sonnet).** One action, exact code below.

**Files:** Modify `src/packages/PkgBackup.bas`.

**Interfaces:** Consumes `Core_ResolveFolder`, `Core_Cfg`, `Core_Refuse`, `Core_SpeedOn`,
`Core_SpeedOff`, `Core_Done`, `Core_Fail`, and Task 3's `PkgBackup_EnsureFolder` / `PkgBackup_Ext`.
Produces `PkgBackup_TimestampedCopy`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Save a timestamped backup copy right now
'================================================================
Public Sub PkgBackup_TimestampedCopy(ctx As Object)
    Dim root As String, destFolder As String, destPath As String, baseName As String
    Const TITLE As String = "Backup Now (Timestamped Copy)"
    On Error GoTo Fail

    root = Core_ResolveFolder(Core_Cfg(ctx, "k_backup_folder"))
    If Len(root) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to back up " & _
               "into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write backup and " & _
               "snapshot files into' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    baseName = Trim(Core_Cfg(ctx, "k_backup_base"))
    If Len(baseName) = 0 Then baseName = "Backup"
    destFolder = root & "Backup" & Application.PathSeparator
    destPath = destFolder & baseName & " " & Format(Now, "yyyy-mm-dd_HHnnss") & PkgBackup_Ext()

    ' Checked BEFORE Core_SpeedOn: the cheapest failure should be the earliest one,
    ' and a refusal after the workbook copy has already started reads as a bug.
    If Len(Dir(destPath)) > 0 Then
        Core_Refuse ctx, "A backup file already exists at:" & vbCrLf & vbCrLf & "   " & destPath & _
               vbCrLf & vbCrLf & "This can happen if this action runs twice in the same second. " & _
               "Wait a moment and run it again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgBackup_EnsureFolder destFolder
    ThisWorkbook.SaveCopyAs destPath
    Core_SpeedOff

    Core_Done ctx, "Saved a timestamped backup copy to:" & vbCrLf & vbCrLf & "   " & destPath & _
           vbCrLf & vbCrLf & "This workbook itself was not touched - its name, path, and any " & _
           "unsaved edits are exactly as they were before this ran.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "backup.timestamped_copy", Err
End Sub
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: the entry-point gate no longer names `PkgBackup_TimestampedCopy`; it should still name
the other three (Tasks 5-7 not yet done).

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgBackup.bas
git commit -m "feat(backup): backup.timestamped_copy"
```

---

### Task 5: `backup.snapshot` - the flagship action

**Owner: delegate to codex (sonnet).** The code below is complete and exact, but this is the
action the whole wave is judged by (see Decisions 1-4 and "Why this wave matters more than its
size suggests") - review the diff against this task more carefully than Tasks 4, 6, and 7.

**Files:** Modify `src/packages/PkgBackup.bas`.

**Interfaces:** Consumes the same Core helpers as Task 4, plus `Core_CfgLong`,
`Core_SafeFileName` (from W3, `2026-07-28-w3-format-export-plan.md` Task 2), and Task 3's
`PkgBackup_PruneSnapshots`. Produces `PkgBackup_Snapshot`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Save a restore-point snapshot before a risky step
'================================================================
' NO RESTORE. This writes a full copy of the workbook and prunes old copies; it does
' NOT read one back. See docs/designs/2026-07-28-w5-backup-plan.md for why that is
' deliberate for this wave and which catalog items it does NOT unblock as a result -
' do not extend this action's completion message to imply a restore capability it
' does not have.
Public Sub PkgBackup_Snapshot(ctx As Object)
    Dim root As String, destFolder As String, destPath As String
    Dim label As String, keepCount As Long
    Const TITLE As String = "Snapshot Before a Risky Step"
    On Error GoTo Fail

    root = Core_ResolveFolder(Core_Cfg(ctx, "k_backup_folder"))
    If Len(root) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to save a " & _
               "snapshot into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write " & _
               "backup and snapshot files into' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' 40 is this caller's cap, not a property of the helper. The rest of the name
    ' ("SNAPSHOT_" + a 17-character timestamp + the extension) is fixed, so bounding
    ' the one free-typed segment is what keeps the whole path inside MAX_PATH.
    label = Core_SafeFileName(Core_Cfg(ctx, "k_backup_snapshot_label"), 40)
    destFolder = root & "Backup" & Application.PathSeparator & "Snapshots" & Application.PathSeparator
    destPath = destFolder & "SNAPSHOT_" & Format(Now, "yyyy-mm-dd_HHnnss") & _
               IIf(Len(label) = 0, "", "_" & label) & PkgBackup_Ext()

    If Len(Dir(destPath)) > 0 Then
        Core_Refuse ctx, "A snapshot file already exists at:" & vbCrLf & vbCrLf & "   " & destPath & _
               vbCrLf & vbCrLf & "This can happen if this action runs twice in the same second. " & _
               "Wait a moment and run it again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgBackup_EnsureFolder destFolder
    ThisWorkbook.SaveCopyAs destPath
    keepCount = Core_CfgLong(ctx, "k_backup_snapshot_keep")
    PkgBackup_PruneSnapshots destFolder, keepCount
    Core_SpeedOff

    Core_Done ctx, "Snapshot saved:" & vbCrLf & vbCrLf & "   " & destPath & vbCrLf & vbCrLf & _
           "This is a full copy of the workbook exactly as it stood the moment this ran, " & _
           "including any edits not yet saved to the main file." & vbCrLf & vbCrLf & _
           "There is no automatic restore for this snapshot. To use it, close this workbook " & _
           "and open the snapshot file directly, or copy what you need from it by hand.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "backup.snapshot", Err
End Sub
```

**Why the refusal above matters more here than in Task 4's near-identical block:** this action's
entire reason to exist is running as step 1 of a chain that also runs something destructive.
Phase 1's chain-halt rule means a refusal here - workbook never saved, a same-second collision -
stops the chain before the destructive step runs. If this action's refusal path were weaker
(silently swallowed, or set `failed` without `refused`), the chain would ask "continue?" and a
user under close-deadline pressure would say yes, and the destructive step would run with no
restore point behind it - the exact failure this action is bought to prevent.

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgBackup.bas
git commit -m "feat(backup): backup.snapshot - restore point, no restore path (see plan doc)"
```

---

### Task 6: `backup.month_end_version`

**Owner: delegate to codex (sonnet).** One action, exact code below.

**Files:** Modify `src/packages/PkgBackup.bas`.

**Interfaces:** Consumes the same Core helpers as Task 4, plus `Core_CfgLong`. Produces
`PkgBackup_MonthEndVersion`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Save a version named for the prior closed month
'================================================================
Public Sub PkgBackup_MonthEndVersion(ctx As Object)
    Dim root As String, destFolder As String, destPath As String, baseName As String
    Dim refDate As Date, priorMonthEnd As Date
    Const TITLE As String = "Save Month-End Version"
    On Error GoTo Fail

    root = Core_ResolveFolder(Core_Cfg(ctx, "k_backup_folder"))
    If Len(root) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to save a " & _
               "month-end version into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to " & _
               "write backup and snapshot files into' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' "Prior month end" is computed from the CALENDAR MONTH of the reference date,
    ' not by counting back a fixed number of days - the 1st and the 28th of the
    ' same month must give the SAME answer (the month that closed before this one),
    ' and a fixed-day-count shortcut ("today minus 30") does not: it lands in a
    ' different month depending on where in the current month "today" falls, and it
    ' gets shorter months wrong. DateSerial(y, m, 1) - 1 depends only on refDate's
    ' YEAR and MONTH, never its day, and needs no January special case:
    ' DateSerial(y, 1, 1) - 1 rolls back to 31 December of y-1 on its own.
    refDate = Date + Core_CfgLong(ctx, "k_backup_offset")
    priorMonthEnd = DateSerial(Year(refDate), Month(refDate), 1) - 1

    baseName = Trim(Core_Cfg(ctx, "k_backup_base"))
    If Len(baseName) = 0 Then baseName = "Backup"
    destFolder = root & "Backup" & Application.PathSeparator
    destPath = destFolder & Format(priorMonthEnd, "yyyy-mm-dd") & " " & baseName & _
               " Month End" & PkgBackup_Ext()

    If Len(Dir(destPath)) > 0 Then
        Core_Refuse ctx, "A month-end version already exists at:" & vbCrLf & vbCrLf & "   " & destPath & _
               vbCrLf & vbCrLf & "This action never overwrites one. Rename or remove the " & _
               "existing file, then run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgBackup_EnsureFolder destFolder
    ThisWorkbook.SaveCopyAs destPath
    Core_SpeedOff

    Core_Done ctx, "Saved the month-end version for " & Format(priorMonthEnd, "mmmm yyyy") & _
           " to:" & vbCrLf & vbCrLf & "   " & destPath, TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "backup.month_end_version", Err
End Sub
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgBackup.bas
git commit -m "feat(backup): backup.month_end_version"
```

---

### Task 7: `backup.archive_output` - the wave's one file-move

**Owner: delegate to codex (sonnet).** One action, exact code below.

**Files:** Modify `src/packages/PkgBackup.bas`.

**Interfaces:** Consumes `Core_FindFile` (existing), `Core_AlreadyOpen` (Task 2), plus the same
Core helpers as Task 4. Produces `PkgBackup_ArchiveOutput`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Relocate a generated output into a dated archive folder
'================================================================
Public Sub PkgBackup_ArchiveOutput(ctx As Object)
    Dim prefix As String, srcPath As String
    Dim root As String, destFolder As String, destPath As String
    Dim openWb As Workbook, fso As Object
    Const TITLE As String = "Archive a Generated Output"
    On Error GoTo Fail

    If Len(ThisWorkbook.Path) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so there is no folder to find a " & _
               "generated file in." & vbCrLf & vbCrLf & "Save this workbook first, then run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    prefix = Trim(Core_Cfg(ctx, "k_backup_archive_prefix"))
    If Len(prefix) = 0 Then
        Core_Refuse ctx, "No file prefix is set for 'Generated output to archive'." & vbCrLf & vbCrLf & _
               "Fill it in on the Setup tab, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Core_FindFile is the SAME "file" kind lookup k_master_prefix already uses:
    ' newest match in THIS workbook's own folder, per contract section 3. It is not
    ' scoped by k_backup_folder - that field only says where the ARCHIVED COPY
    ' lands, not where the source is searched for.
    srcPath = Core_FindFile(prefix)
    If Len(srcPath) = 0 Then
        Core_Refuse ctx, "No file starting with '" & prefix & "' was found in:" & vbCrLf & vbCrLf & _
               "   " & ThisWorkbook.Path & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' A file cannot be safely moved while Excel holds it open: Name and
    ' FileSystemObject.MoveFile both fail against a locked file, and failing
    ' partway through a move can leave neither a clean source nor a clean
    ' destination. Refused here, before anything on disk changes - see
    ' Core_AlreadyOpen's own header for why the probe is by file name.
    Set openWb = Core_AlreadyOpen(srcPath)
    If Not openWb Is Nothing Then
        Core_Refuse ctx, "'" & Core_FileName(srcPath) & "' is currently open in Excel, and this " & _
               "action cannot move a file that is open." & vbCrLf & vbCrLf & "Close it, then run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' ThisWorkbook.Path is already proven non-blank above, so Core_ResolveFolder
    ' cannot return "" here even with k_backup_folder left blank.
    root = Core_ResolveFolder(Core_Cfg(ctx, "k_backup_folder"))
    destFolder = root & "Archive" & Application.PathSeparator & Format(Date, "yyyy_mm") & _
                 Application.PathSeparator
    destPath = destFolder & Core_FileName(srcPath)

    If Len(Dir(destPath)) > 0 Then
        Core_Refuse ctx, "A file named '" & Core_FileName(srcPath) & "' already exists in:" & _
               vbCrLf & vbCrLf & "   " & destFolder & vbCrLf & vbCrLf & "This action never " & _
               "overwrites one. Rename or remove the existing file, then run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgBackup_EnsureFolder destFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile srcPath, destPath                ' not the Name statement: MoveFile works
                                                    ' across drives, Name does not
    Core_SpeedOff

    Core_Done ctx, "Archived:" & vbCrLf & vbCrLf & "   " & srcPath & vbCrLf & "to" & vbCrLf & _
           "   " & destPath, TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "backup.archive_output", Err
End Sub
```

- [ ] **Step 2: Prove every entry point now resolves**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics, no entry-point-gate failures naming any `PkgBackup_*`.

- [ ] **Step 3: Two deliberate injections**

First, prove the phase-2 prefix guard fires against this module specifically:

```bash
printf '\nPrivate Function TotallyUnprefixed() As Boolean\nEnd Function\n' >> src/packages/PkgBackup.bas
pwsh build/build.ps1; echo "exit=$?"
git checkout -- src/packages/PkgBackup.bas
pwsh build/build.ps1; echo "exit=$?"
```

Expected: `exit=1` the first time, naming `PkgBackup.bas` and `TotallyUnprefixed`; `exit=0` after
the revert.

Second, prove the id gate's vocabulary/shape check fires against `backup` specifically, not only
the domains it has been exercised against so far:

```bash
sed -i 's/"backup.snapshot"/"Backup.Snapshot"/' src/packages/PkgBackup.bas
pwsh build/build.ps1; echo "exit=$?"
git checkout -- src/packages/PkgBackup.bas
pwsh build/build.ps1; echo "exit=$?"
```

Expected: `exit=1` the first time, naming the malformed id (camelCase / uppercase segment);
`exit=0` after the revert. **Record both injections in the PR body.**

- [ ] **Step 4: Commit**

```bash
git add src/packages/PkgBackup.bas
git commit -m "feat(backup): backup.archive_output"
```

---

### Task 8: Register `Backup` in `CoreCandidates`

**Owner: Opus.** Core edit; this is the step the plan-of-record names as the single easiest way
to ship a dead wave, so it gets its own task and its own verification rather than riding along
inside Task 2 or 7.

**Files:** Modify `src/core/Core.bas` (`CoreCandidates`, grep `docs/code-map.md` for the anchor).

**Interfaces:** Produces nothing new; makes `PkgBackup_Manifest` actually reachable.

- [ ] **Step 1: Add the candidate**

By the time this wave branches, `CoreCandidates` already lists whatever W2-W4 added. Add one
line:

```vb
    c.Add "Backup"                  ' probes PkgBackup_Manifest
```

- [ ] **Step 2: Verify it is really there, and that the build still passes**

```bash
grep -n 'c.Add "Backup"' src/core/Core.bas
pwsh build/build.ps1
```

Expected: one match; exit 0, package count in the build's summary line one higher than before
this task.

- [ ] **Step 3: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): register Backup in CoreCandidates"
```

---

### Task 9: Version, changelog, code map, deferred-validation log, and the PR

**Owner: Opus.**

**Files:** `VERSION`, `CHANGELOG.md`, `docs/code-map.md`,
`docs/designs/2026-07-28-deferred-validation-log.md`.

- [ ] **Step 1: `VERSION` -> `0.16.0`, and a `## [0.16.0]` changelog entry**

House style: prose about what changed for the user, and honest about the limit. State plainly
that `backup.snapshot` writes a restore point with no restore action, and name which Deferred
catalog items that does and does not affect (none, per this plan's own count above).

- [ ] **Step 2: Append this wave's run-time-only claims**

If `docs/designs/2026-07-28-deferred-validation-log.md` does not exist yet, create it with a
one-line header (`# Deferred validation log`) before appending - it is meant to exist by now
(phase 1 creates it), but this task must not silently depend on that.

```markdown
## W5 - Protect & Back Up

- **`CoreCandidates` actually reaching `backup.*`:** confirmed statically (Task 8's grep) that
  the line was added and the build still passes, but whether the four actions actually appear
  on a rendered Setup/button page, and actually dispatch through `Core_Run`, is unverifiable by
  `build.ps1` - a missing entry here "builds, bundles, passes every gate, and never runs"
  (plan-of-record section 3.1).
- **`SaveCopyAs` produces a file that reopens with its macros intact.** The extension is derived
  from `ThisWorkbook.name` specifically so format and extension always match, but whether a real
  Excel session actually writes and reopens a working `.xlsm` copy is a run-time fact.
- **`PkgBackup_EnsureFolder` actually creates a multi-level path on a real filesystem**,
  including a network share or a mapped drive, where `Scripting.FileSystemObject.CreateFolder`'s
  behaviour has more edge cases than a local disk.
- **`PkgBackup_PruneSnapshots` deletes exactly the files it should and no others** - run against
  a real folder with more than `k_backup_snapshot_keep` snapshots in it, including one snapshot
  that happens to be open or locked at prune time (the `Resume Next` path around `Kill`).
- **`fso.MoveFile` in `backup.archive_output` actually moves a file across drives** (the reason
  `MoveFile` was chosen over the `Name` statement) - untested against a real second drive or
  mapped network location.
- **The prior-month-end date arithmetic (`backup.month_end_version`) on 1 January**, and on the
  1st of any month generally - the formula is argued correct in this plan's Decision 8, but has
  not run inside Excel.
```

- [ ] **Step 3: Refresh the code map**

```bash
node scripts/code-map.mjs
```

- [ ] **Step 4: Final gates**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas
grep -nE '^\s*Private\s+(Sub|Function)\s+' src/packages/*.bas | grep -vE 'Private\s+(Sub|Function)\s+Pkg'
grep -n 'c.Add "Backup"' src/core/Core.bas
git status --short
```

Expected: exit 0; no diagnostics; `1`; no output from the prefix grep; one match; a tree clean
but for the untracked `.vscode/`.

- [ ] **Step 5: Review the diff, then one PR against `main`**

Branch from `main` after W4 has merged - not stacked. The PR body must state, in one line, which
Deferred catalog items this wave does and does not unblock (none - see the opening section of
this plan), so a reviewer does not have to re-derive that from the diff.

---

## Unknowns and risks

- **The catalog's own framing overstates what this wave delivers.** Its text for
  `backup.snapshot` calls it "the unblocker for several *Deferred* destructive items" without
  naming which ones or noting that a restore action does not exist anywhere in the catalog. This
  plan corrects the record in its opening section and in Decision 2; a future wave that adds a
  restore action should update the catalog's own wording at the same time, since right now the
  catalog itself is the thing most likely to mislead a reader who has not seen this plan.
- **`backup.archive_output`'s reliance on `Core_FindFile` inherits that function's existing
  blind spot for an unsaved workbook** (`folder = ThisWorkbook.Path`, empty when never saved,
  degrading to a root-of-drive search rather than a clean error). This plan's own precondition
  check (Task 7, Step 1) catches the unsaved case before `Core_FindFile` is ever called, so the
  blind spot is not reachable through this action - but `Core_FindFile` itself is unchanged and
  the same latent issue exists for its other caller (`PkgImportExport_SuggestTypes`), unfixed
  here as out of scope.
- **Snapshot pruning assumes `FileDateTime` is trustworthy for ordering.** On most filesystems it
  is, but a snapshot folder synced through a cloud-backed drive (OneDrive) can rewrite modified
  timestamps on sync in ways that are not fully within this plan's control. Flagged, not solved -
  the folder is expected to be a local or mapped-drive backup location, not itself a synced one,
  but nothing in this wave enforces that.
- **`k_backup_folder` is one field for four different subfolder conventions
  (`Backup`, `Backup\Snapshots`, `Archive\yyyy_mm`).** Decision 5 argues this is the right amount
  of Setup-tab surface; if a deployment genuinely needs its backups and its archive to live in
  unrelated locations, this wave does not support that, and splitting the field is the fix,
  not a workaround inside the existing one.
- **No restore action exists, and none is proposed here.** Repeated deliberately, once more, so
  it cannot be missed by a reader who only reads this section: this wave saves; it does not
  restore.
- **`Core_SafeFileName` (W3) does not cap length; the package-private helper this plan removed
  in its favor did (`Left$(s, 40)`).** Checked directly against W3's plan
  (`2026-07-28-w3-format-export-plan.md`, Task 2, Step 2) before reusing it: the character
  substitution is a full superset (it also replaces Core's own `[` `]`, which the removed
  helper left alone, and the two agree on `Trim` and on treating every other bad character
  identically), so on that axis W3's helper fully covers this wave's need. The one thing it
  does not do - and says so in its own doc comment ("Does NOT make the name unique or check
  length") - is truncate. `backup.snapshot`'s `k_backup_snapshot_label` is free-typed on
  Setup, so an unusually long label now produces an unusually long snapshot file name with
  nothing in this wave's own code stopping it, where before the 40-character cap kept it
  bounded. **This gap was found here and closed in W3 rather than carried.** `Core_SafeFileName`
  now takes `Optional ByVal maxLen As Long = 0`, `0` meaning no cap, so W3's own two call sites
  keep their unbounded behaviour and `backup.snapshot` passes `40` — the same bound the removed
  helper had. See `2026-07-28-w3-format-export-plan.md` Task 2 Step 2 for the current signature.
  **W5 therefore depends on W3's version of that helper, not merely on its existence**; if W3
  ships without `maxLen`, this call does not compile, which is the failure mode to want.
