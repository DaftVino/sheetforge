# W6 - Workbook housekeeping

> **For agentic workers:** REQUIRED SUB-SKILL: use `superpowers:subagent-driven-development`
> to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship `src/packages/PkgWorkbook.bas`, the domain `workbook`'s package, covering the
four catalog items at `docs/macro-catalog-plug-and-play.md` lines 196-205: unhide every
worksheet, protect all sheets, unprotect all sheets, build a Table of Contents, and reset row/
column visibility on the active sheet. Five actions from four items - line 199 is one item
that deliberately produces two ids.

**Architecture:** One new package, thin, built entirely on the Core API phases 1-2 and W2-W5
add (`Core_Refuse`, `Core_IsEngineSheet`, `Core_AskRange`, `Core_DestUsable`, `Core_NewSheet`).
`Core_DestUsable` already refuses a destination name the toolkit's own rebuild will claim later
(`_Config`, `Macros 3`, `Macros Home`, ...), not only a name that already exists - phase 2
(`2026-07-28-core-helper-promotion-plan.md`, Task 2 Step 3) added that branch ahead of this
wave, merging long before W6, because every wave that creates a sheet from a *user-typed* name
benefits from it and this wave is only the first to need it, for the Table of Contents. This
wave makes no Core change of its own beyond registering the new package with `CoreCandidates()`;
it consumes the fix rather than adding it.

**Tech stack:** VBA (Excel/Windows, one concatenated `.bas`), PowerShell for the build, Node
for the syntax check. No test framework - see *Verification*.

**Runs after** phases 1-2 and W2-W5 (`2026-07-28-track2-completion-plan-of-record.md` §2). It
is **phase 7 of 8** and the **last wave** - phase 8's single end-to-end validation is next, and
this plan's wrap-up task hands it the accumulated `docs/designs/2026-07-28-deferred-validation-
log.md` rather than writing its own gate document.

---

## Why this wave collides with the engine harder than any other

Every other wave's actions read a picked range or a source file and write to a **new** sheet.
This wave's actions read and write the **engine's own territory**: every sheet's visibility,
every sheet's protection state, and (for the TOC) a sheet listing every other sheet by name.
`Core_IsEngineSheet` exists precisely so a package answers "is this sheet mine to touch?" from
Core's own constants rather than a hardcoded copy - and this wave calls it more than any other,
because getting it wrong here is worse than anywhere else:

- **`workbook.unhide_sheets` unhiding `_Config` or `_Config_staged`** would put a finance user
  in front of the engine's own configuration store - two sheets that are hidden **by design**
  and invite editing damage the moment they are visible. This is the single most important
  correctness point in the wave, and it is handled by the same `Core_IsEngineSheet` exclusion
  every other action in this file uses, not a special case.
- **`workbook.protect_sheets` protecting a generated `Macros <n>` page** with Excel's own
  default (`DrawingObjects:=True`) would block the object-model deletion the next rebuild needs
  to remove its own buttons - contract §6 already forbids the engine unprotecting a surface to
  fix that itself, so the trap, once sprung, has no engine-side escape. This wave avoids
  springing it (Decision 3 below), rather than relying on the rule that forbids un-springing it.
- **`workbook.build_toc` writing to a reserved name** (`Macros Home`, `Macros 1`, `_Config`) is
  destroyed unconditionally, alerts suppressed, on the very next Finalize - a defect that would
  not surface until the rebuild *after* the one that created the illusion of success. This is
  the gap `Core_DestUsable` closes, per phase 2 (`2026-07-28-core-helper-promotion-plan.md`,
  Task 2 Step 3) - this wave inherits the fix rather than adding it.

---

## Decisions taken in this plan

| # | Question | Decision |
|---|---|---|
| 1 | `unhide_sheets` and `_Config`/`_Config_staged` | **Excluded, via `Core_IsEngineSheet`.** Those two sheets are hidden by design as the engine's config store; unhiding them is not a convenience, it is exposing internal state. No override, no confirm-to-unhide-anyway - the exclusion is unconditional, the same way `import.combine_worksheets` unconditionally excludes engine sheets from what it combines. |
| 2 | Which sheets does `protect_sheets` protect? | **Every sheet where `Core_IsEngineSheet` is False.** Setup, the config sheets, and every generated `Macros <n>` page are never protected by this action, so a later rebuild can still manage its own controls on them. |
| 3 | What protection parameters does `protect_sheets` pass? | **`DrawingObjects:=False`, not Excel's own default (`True`).** A sheet protected with `DrawingObjects:=True` blocks the object-model deletion a future rebuild needs for its own anchored buttons (contract §6's "best-effort" clause exists for exactly this failure, and the same section forbids the engine unprotecting a surface to escape it). The only safe moment to avoid the trap is the one time a **person**, not the engine, is choosing to protect - so this action chooses parameters that keep the engine's own future access, even though it is locking everything else down. |
| 4 | Where does the password come from? | **Config (`k_workbook_password`, text, advanced, default blank), not a per-run prompt.** Stated plainly rather than implied: this is not a secret. It sits in a cell on the hidden `_Config` sheet, in plain text, exactly like every other Setup answer - anyone who can reach that sheet (or the config store on disk) can read it. It only stops accidental edits from the ribbon, the same as Excel's own Protect Sheet with no server-side enforcement behind it. A per-run prompt would not fix this (the same password gets typed by the same people either way) and it would break the quiet-chain convention (an interactive prompt cannot be suppressed - contract §6). Storing it in config is consistent with everything else on Setup and is honestly no worse. |
| 5 | Wrong password on `unprotect_sheets` | **Caught per sheet, not aborted halfway.** Each `Unprotect` call is wrapped in its own error trap; a sheet whose password does not match is added to a "skipped" list and the loop continues. `Core_Done`'s message reports both counts by name, so a partial run is never mistaken for a complete one. |
| 6 | `Unprotect`'s `Password:=` argument when the config password is blank | **Always passed explicitly, never omitted**, even when it is `""`. Calling `Unprotect` with **no** argument shows Excel's own interactive password box when the sheet carries one - which would surface mid-chain, breaking the quiet convention. An explicit argument (blank or not) never prompts: a mismatch raises error 1004 instead, which the per-sheet trap catches. **Marked as an assumption about Excel's own behaviour** (`CLAUDE.md` §7) - see Unknowns. |
| 7 | Does `build_toc` create-only, or refresh? | **Refresh in place, using a written marker to prove ownership.** A create-only TOC would refuse on every run after the first - the exact shape of the `import.combine_worksheets` bug the whole refusal pass (phase 1) is named for, except here re-running after adding a tab is the entire point of the feature, not an edge case. The sheet's `A1` holds a fixed sentence (`PkgWorkbook_TocMarker`); on a rerun, if a sheet by the configured name exists and `A1` matches that sentence, its contents are cleared and rebuilt. If a sheet exists with that name and `A1` does **not** match, the action refuses exactly the way `Core_DestUsable` refuses any other name collision - a user's own sheet that happens to share the configured name is never touched. |
| 8 | Does `Core_DestUsable` need a reserved-name check, and where does it live? | **Yes, and phase 2 already added it - this wave consumes it rather than adding it.** `Core_DestUsable` used to refuse only an *existing* sheet of that name; it had no opinion on a name the engine would claim **later** (`Macros 3` before any rebuild has produced one, `_Config` if it were somehow briefly absent). `Core_IsEngineSheet` already answers "is this name in the reserved family?" as a pure string test with no dependency on whether a sheet currently exists by that name, so phase 2 added the check as one call inside `Core_DestUsable` (`2026-07-28-core-helper-promotion-plan.md`, Task 2 Step 3), and every existing and future caller - including this wave's TOC - inherits it the moment that PR merges. Hardcoding the reserved family inside `PkgWorkbook` instead would still be the exact defect phases 1 and 2 exist to remove: a second copy of state Core already owns. Phase 2's version also checks `ThisWorkbook.Sheets` rather than `Core_SheetExists`, so a name held by a **chart sheet** is refused too - a second refusal this wave did not know about when it was first drafted. |
| 9 | Is the `build_toc` name-collision refusal classed additive or breaking? | **Additive**, on the same reasoning contract v2.4.0 used for the chain-halt rule (§10): a config value like `k_import_dest = "Macros 3"` was never actually *usable* correctly before this fix - it produced a sheet the very next Finalize destroyed unconditionally. Refusing it now prevents a data-loss trap rather than withdrawing behaviour anyone could rely on. |
| 10 | `unhide_rows_cols` scope: whole sheet, active sheet, or a picked range? | **The whole active sheet, not a `Core_AskRange` pick.** The catalog names this "Reset visibility **across the sheet**" (singular), matching the classic snippet it replaces (select all cells, Unhide Rows, Unhide Columns) - a range-scoped version would leave hidden rows/columns just outside the pick untouched, which is not "reset," it is "reset some of it." |
| 11 | `unhide_rows_cols` and an active AutoFilter | **Warned, not silently overridden, not refused.** `Range.Hidden` is the same property for a manually hidden row and a filter-hidden row - there is no way to tell them apart or to unhide only one kind. The action detects `ws.AutoFilterMode` and, only then, shows a Yes/No confirm naming the effect (every filtered-out row reappears) before proceeding; without an active filter it just does the reset and reports the count, like `unhide_sheets`. |
| 12 | `unhide_rows_cols` and an engine sheet | **Refused if the active sheet is engine-owned.** Setup and the generated `Macros <n>` pages rely on rows/columns hidden on purpose as part of their own layout (the button-page grid, admin rows); unhiding them exposes the toolkit's internal layout, not something the user can use. |
| 13 | Does `build_toc` list hidden sheets? | **No.** A hyperlink's `SubAddress` cannot navigate to a hidden sheet, so a link to one is a dead click that looks like a bug. Only sheets where `Visible = xlSheetVisible` are listed, and the completion message says so. |
| 14 | Safety story per action (one line each, per `CLAUDE.md`'s bar) | `unhide_sheets`: non-destructive and reversible, no confirm, engine sheets excluded unconditionally. `protect_sheets`: preview-and-confirm (count + which sheets are excluded and why) before locking editing broadly. `unprotect_sheets`: non-destructive, no confirm, per-sheet failure reporting instead of an all-or-nothing abort. `build_toc`: additive-only on a fresh name, refuse-on-collision otherwise; the only content it ever destroys is its own prior output, proven by the marker, never a user's. `unhide_rows_cols`: non-destructive and reversible, confirm gated specifically on the one case (an active filter) where the result looks like something broke. |

---

## Global constraints

Copied from `2026-07-28-track2-completion-plan-of-record.md` §4, verbatim.

- **`SCHEMA_VER` stays `"v2.0"`.**
- **Action ids are immutable**, two-segment `<domain>.<operation>`, lowercase ASCII and
  underscores. This wave's five ids are frozen at `docs/macro-catalog-plug-and-play.md` lines
  196-205 and used verbatim: `workbook.unhide_sheets`, `workbook.protect_sheets`,
  `workbook.unprotect_sheets`, `workbook.build_toc`, `workbook.unhide_rows_cols`.
- **Shared helpers live ONLY in Core. Cross-package calls are forbidden.**
- **Pure ASCII in every `.bas`**; `Option Explicit`; max 25 line continuations per statement;
  **no module-level declaration after the first procedure** (a package's `Const` sits above its
  own first `Sub`); no VBA reserved word as an identifier.
- **Every action entry point is `Public Sub PkgWorkbook_<Verb>(ctx As Object)`** - one `Object`
  parameter, nothing else. `build.ps1` fails the build otherwise.
- **Every package private is prefixed `PkgWorkbook_`** (phase 2's guard enforces this at build
  time).
- **TDD is off.** VBA has no headless runner. **Do not propose a test suite.** Verification is
  `pwsh build/build.ps1`, `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`, and
  deliberate injections where one is meaningful.
- **No per-wave manual gate document.** Run-time-only claims are appended to
  `docs/designs/2026-07-28-deferred-validation-log.md` for phase 8, not written as a checklist.
- **`dist/` is gitignored** - never `git add -f` the bundle. **`.vscode/` stays untracked.**
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.

---

## Verification, since TDD is off

1. `pwsh build/build.ps1` - exit 0, `Core + N package(s)` including `PkgWorkbook.bas`, balanced
   `Sub`/`Function` counts, no duplicate-procedure or unprefixed-private failure.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` - no diagnostics.
3. **Grep proofs** for each mechanical claim (id uniqueness, field-key uniqueness, the
   `CoreCandidates` line, no leftover old text) - see each task's own step.
4. **What the build genuinely cannot see**: whether `Core_IsEngineSheet` actually excludes
   `_Config` at run time, whether `DrawingObjects:=False` actually preserves a rebuild's control
   placement on a protected sheet, whether `Unprotect Password:=""` really avoids Excel's
   interactive prompt, whether the TOC marker check really distinguishes ownership correctly,
   and whether the AutoFilter warning fires when it should. **These are logged to
   `docs/designs/2026-07-28-deferred-validation-log.md` in Task 8**, not gated here - per
   plan-of-record §3.5, phase 8 is where the owner runs the workbook once, for real.

---

## File structure

| File | Change |
|---|---|
| `src/core/Core.bas` | `CoreCandidates()` gains `c.Add "Workbook"`. |
| `src/packages/PkgWorkbook.bas` | **New.** Header comment, `PkgWorkbook_Manifest`, five `Public Sub` actions, three `Private` helpers (`PkgWorkbook_FailedNote`, `PkgWorkbook_TocMarker`, `PkgWorkbook_CountHidden`). |
| `spec/contract.md` | v2.9.0. §7 gets one new "Reserved" normative bullet (Decision 8/9); §8 gains a `Package workbook` subsection; §10 gains the version bullet. |
| `docs/designs/2026-07-28-deferred-validation-log.md` | Gains this wave's run-time-only claims. |
| `VERSION`, `CHANGELOG.md`, `docs/code-map.md` | v0.17.0. |

---

### Task 1: New module - header, manifest, registration proof

**Owner: Opus.** The manifest encodes the frozen ids and the field `usedBy` wiring; getting a
typo here silently breaks Setup's disclosure rule (§3's "omitting the field entirely remains
non-conforming" note exists because this exact mistake is easy to make and easy to miss).

**Files:** New `src/packages/PkgWorkbook.bas`. Modify `src/core/Core.bas` - `CoreCandidates()`
gains one line registering the new package (Step 2 below).

**Interfaces:**
- Consumes: `Core_DeclarePackage`, `Core_AddSection`, `Core_AddField`, `Core_AddAction` (all
  existing).
- Produces: `PkgWorkbook_Manifest(reg As Object)`. Tasks 2-6 add the five action subs it
  declares as entries. Also registers `"Workbook"` in `CoreCandidates()` - the two-place
  registration the whole wave depends on running at all.

- [ ] **Step 1: Write the file header and manifest**

```vb
Attribute VB_Name = "PkgWorkbook"
Option Explicit

'================================================================
' PkgWorkbook - housekeeping on the workbook itself: sheet
' visibility, protection, and navigation. Wave 6 of the catalog
' build-out (docs/designs/2026-07-28-w6-workbook-plan.md).
'
' Five actions from FOUR catalog items - "Protect / unprotect all
' sheets at once" is one catalog item that deliberately produces
' two ids, because a single toggle button hides which way it will
' go (docs/macro-catalog-plug-and-play.md line 199):
'   workbook.unhide_sheets       every worksheet visible again
'   workbook.protect_sheets      lock every sheet this deployment owns
'   workbook.unprotect_sheets    unlock every sheet this deployment owns
'   workbook.build_toc           an index sheet linking to every visible tab
'   workbook.unhide_rows_cols    reset row/column visibility, active sheet
'
' Every action here walks all sheets or changes sheet state, and the
' engine owns several sheets outright. Every action excludes them
' via Core_IsEngineSheet rather than a local copy of their names -
' see the plan's Decisions table for why each exclusion exists.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgWorkbook_Manifest(reg As Object)
    Core_DeclarePackage reg, "workbook", "Workbook Housekeeping", "1.0.0"

    Core_AddSection reg, "TABLE OF CONTENTS"
    Core_AddField reg, "k_workbook_toc_name", "Name for the Table of Contents sheet", "text", "Contents", _
        "simple", _
        "A sheet by this name is created the first run and refreshed in place on every later " & _
        "run - it is never treated as a sheet you already own, so a rerun after adding a tab " & _
        "is the normal way to use this. If the name is already taken by a sheet this action " & _
        "did not build, or by a name the toolkit itself reserves (Setup, the config sheets, or " & _
        "a generated Macros page), the action stops and asks for a different name. Kind is " & _
        "text, not sheet, because a sheet field is validated to already EXIST and this one " & _
        "must not.", _
        "workbook.build_toc"

    Core_AddSection reg, "PROTECTION"
    Core_AddField reg, "k_workbook_password", "Password for Protect / Unprotect All Sheets", "text", "", _
        "advanced", _
        "Leave blank to protect with no password. This is stored in plain text on the hidden " & _
        "config sheet, the same as every other Setup answer - it is not a secret, and storing " & _
        "it here does not make a sheet's contents confidential. It only stops accidental edits " & _
        "from the ribbon, the same protection Excel's own Protect Sheet provides on its own.", _
        "workbook.protect_sheets,workbook.unprotect_sheets"

    Core_AddAction reg, "workbook.unhide_sheets", "Unhide All Worksheets", "PkgWorkbook_UnhideSheets", _
        "Makes every hidden or very-hidden worksheet visible again. Setup and the toolkit's " & _
        "own internal sheets - including the hidden config store - are never touched.", _
        "Workbook Housekeeping"
    Core_AddAction reg, "workbook.protect_sheets", "Protect All Sheets", "PkgWorkbook_ProtectSheets", _
        "Locks every sheet you own against editing, with the password above (or no password " & _
        "if it is blank). Setup and the toolkit's own internal sheets are left unprotected, so " & _
        "a later rebuild can still manage its own buttons on them.", _
        "Workbook Housekeeping"
    Core_AddAction reg, "workbook.unprotect_sheets", "Unprotect All Sheets", "PkgWorkbook_UnprotectSheets", _
        "Unlocks every sheet you own, using the password above. A sheet locked with a " & _
        "different password is reported by name rather than stopping the rest.", _
        "Workbook Housekeeping"
    Core_AddAction reg, "workbook.build_toc", "Build Table of Contents", "PkgWorkbook_BuildToc", _
        "Creates, or refreshes, one sheet listing every visible tab, each linked to jump " & _
        "straight to it.", _
        "Workbook Housekeeping"
    Core_AddAction reg, "workbook.unhide_rows_cols", "Unhide All Rows & Columns", "PkgWorkbook_UnhideRowsCols", _
        "Resets row and column visibility on the active sheet. Warns first if an AutoFilter " & _
        "on that sheet would also lose its filtered-out rows.", _
        "Workbook Housekeeping"
End Sub
```

- [ ] **Step 2: Register the package in `CoreCandidates()`**

In `CoreCandidates()`, add one line before `Set CoreCandidates = c`:

```vb
    c.Add "Workbook"                 ' probes PkgWorkbook_Manifest
```

This is the two-place registration the whole wave depends on running at all - `build.ps1`
bundles `PkgWorkbook.bas` by glob whether or not this line exists; only this line makes Core
ever call it.

- [ ] **Step 3: Verify field-key, action-id, and registration (static)**

```bash
grep -rn '"k_workbook_toc_name"\|"k_workbook_password"' src/packages/*.bas
grep -rn '"workbook\.' src/packages/*.bas
grep -n 'c.Add "Workbook"' src/core/Core.bas
```

Expected: every field-key/action-id hit is inside `PkgWorkbook.bas` itself - no other package
independently declared either key or any `workbook.*` id. One hit for the `CoreCandidates` line.

- [ ] **Step 4: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: **build failure**, and this is expected at this point - the manifest declares five
entry points (`PkgWorkbook_UnhideSheets` etc.) that Tasks 2-6 have not written yet, so
`build.ps1`'s entry-point resolver has nothing to find. Confirm the failure names the missing
subs rather than something else; that is the proof the manifest itself is wired correctly.

- [ ] **Step 5: Commit**

```bash
git add src/packages/PkgWorkbook.bas src/core/Core.bas
git commit -m "feat(workbook): PkgWorkbook manifest - two fields, five action entries; register with CoreCandidates"
```

---

### Task 2: `workbook.unhide_sheets`

**Owner: delegate to codex (sonnet).** Mechanical: one loop, one exclusion, one report.

**Files:** Modify `src/packages/PkgWorkbook.bas`.

**Interfaces:** Consumes `Core_IsEngineSheet`, `Core_SpeedOn`/`Core_SpeedOff`, `Core_Done`,
`Core_Fail`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Unhide every worksheet, except the toolkit's own
'================================================================
Public Sub PkgWorkbook_UnhideSheets(ctx As Object)
    Dim ws As Worksheet, n As Long
    Const TITLE As String = "Unhide All Worksheets"
    On Error GoTo Fail

    Core_SpeedOn
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) Then
            If ws.Visible <> xlSheetVisible Then
                ws.Visible = xlSheetVisible
                n = n + 1
            End If
        End If
    Next ws
    Core_SpeedOff

    ' The toolkit's own sheets - Setup, the hidden config store and its staging
    ' sheet, the package-enable sheet, every generated Macros page - are hidden
    ' or shown for the toolkit's own reasons and are excluded unconditionally,
    ' the same exclusion every other action in this file uses. Unhiding _Config
    ' specifically would put a finance user in front of the engine's own working
    ' state and invite them to edit it - the single most important correctness
    ' point in this action.
    If n = 0 Then
        Core_Done ctx, "Nothing was hidden. Every sheet you own is already visible." & vbCrLf & vbCrLf & _
               "The toolkit's own internal sheets are never touched by this action, hidden or " & _
               "not.", TITLE
    Else
        Core_Done ctx, "Unhid " & n & " sheet(s)." & vbCrLf & vbCrLf & _
               "The toolkit's own internal sheets - Setup, the config store, and any generated " & _
               "Macros page - were left exactly as they were.", TITLE
    End If
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.unhide_sheets", Err
End Sub
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: still fails (four entries remain missing), but the failure list drops
`PkgWorkbook_UnhideSheets`.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgWorkbook.bas
git commit -m "feat(workbook): workbook.unhide_sheets"
```

---

### Task 3: `workbook.protect_sheets`, and the shared failure-note helper

**Owner: delegate to codex (sonnet).** Mechanical, following the count-then-confirm shape at
`src/packages/PkgClean.bas:292-310` (`PkgClean_ConfirmBlanks`).

**Files:** Modify `src/packages/PkgWorkbook.bas`.

**Interfaces:**
- Consumes `Core_IsEngineSheet`, `Core_Cfg`, `Core_SpeedOn`/`Core_SpeedOff`, `Core_Refuse`,
  `Core_Done`, `Core_Fail`.
- Produces `PkgWorkbook_FailedNote` (private), reused unchanged by Task 4.

- [ ] **Step 1: Add the action and the shared helper**

```vb
'================================================================
' ACTION - Protect every sheet this deployment owns
'================================================================
Public Sub PkgWorkbook_ProtectSheets(ctx As Object)
    Dim targets As Collection, ws As Worksheet, pwd As String
    Dim doneCount As Long, failed As Collection, msg As String
    Const TITLE As String = "Protect All Sheets"
    On Error GoTo Fail

    Set targets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) Then targets.Add ws
    Next ws
    If targets.Count = 0 Then
        Core_Refuse ctx, "There are no sheets to protect - every sheet in this workbook " & _
               "belongs to the toolkit itself." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    pwd = Core_Cfg(ctx, "k_workbook_password")
    msg = "This will protect " & targets.Count & " sheet(s) against editing" & _
          IIf(Len(pwd) = 0, ", with no password", ", with the password set on Setup") & "." & _
          vbCrLf & vbCrLf & _
          "The toolkit's own sheets - Setup, the config store, and any generated Macros page - " & _
          "are never protected by this action, so a later rebuild can still update them: " & _
          "protecting a generated page with Excel's own default would block the object-model " & _
          "deletion a rebuild needs to remove its own buttons." & vbCrLf & vbCrLf & "Continue?"
    If MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Set failed = New Collection
    Core_SpeedOn
    For Each ws In targets
        On Error Resume Next
        Err.Clear
        ' DrawingObjects:=False, deliberately NOT Excel's own default (True). A
        ' sheet protected with DrawingObjects:=True blocks a future rebuild from
        ' removing its own anchored controls on it - and contract section 6
        ' already forbids the engine unprotecting a user's sheet to fix that
        ' itself. This is the one moment a PERSON, not the engine, is choosing to
        ' protect, so it is the one place that choice can still be made safely.
        ws.Protect Password:=pwd, DrawingObjects:=False, Contents:=True, Scenarios:=True
        If Err.Number <> 0 Then
            failed.Add ws.name
            Err.Clear
        Else
            doneCount = doneCount + 1
        End If
        On Error GoTo 0
    Next ws
    Core_SpeedOff

    Core_Done ctx, "Protected " & doneCount & " of " & targets.Count & " sheet(s)." & _
           PkgWorkbook_FailedNote(failed, "already protected, or otherwise could not be locked"), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.protect_sheets", Err
End Sub

'================================================================
' Package-private helpers. Prefixed PkgWorkbook_ because the bundle
' is one flat namespace, even though these are Private.
'================================================================

' Names which target sheets did not get the requested change, so a partial
' result is STATED rather than silently rounded up to "done". Returns "" when
' every target succeeded, so a caller can test Len() rather than compare text.
' Shared by protect_sheets and unprotect_sheets - both hit the same shape of
' partial failure (one sheet's password or state does not cooperate) and both
' need the same "here is exactly what did not happen" honesty.
Private Function PkgWorkbook_FailedNote(failed As Collection, ByVal reason As String) As String
    Dim i As Long, names As String
    If failed.Count = 0 Then Exit Function
    For i = 1 To failed.Count
        If i > 1 Then names = names & ", "
        names = names & failed(i)
    Next i
    PkgWorkbook_FailedNote = vbCrLf & vbCrLf & failed.Count & " sheet(s) were skipped (" & _
        reason & "): " & names
End Function
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: still fails (three entries remain missing); the failure list drops
`PkgWorkbook_ProtectSheets`.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgWorkbook.bas
git commit -m "feat(workbook): workbook.protect_sheets"
```

---

### Task 4: `workbook.unprotect_sheets`

**Owner: delegate to codex (sonnet).** Mechanical; reuses Task 3's `PkgWorkbook_FailedNote`.

**Files:** Modify `src/packages/PkgWorkbook.bas`.

**Interfaces:** Consumes `Core_IsEngineSheet`, `Core_Cfg`, `Core_SpeedOn`/`Core_SpeedOff`,
`Core_Done`, `Core_Fail`, and Task 3's `PkgWorkbook_FailedNote`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Unprotect every sheet this deployment owns
'================================================================
Public Sub PkgWorkbook_UnprotectSheets(ctx As Object)
    Dim ws As Worksheet, pwd As String
    Dim doneCount As Long, failed As Collection
    Const TITLE As String = "Unprotect All Sheets"
    On Error GoTo Fail

    pwd = Core_Cfg(ctx, "k_workbook_password")
    Set failed = New Collection
    Core_SpeedOn
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) Then
            If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
                On Error Resume Next
                Err.Clear
                ' Password:=pwd is passed EXPLICITLY, even when pwd is "". Calling
                ' Unprotect with NO argument shows Excel's OWN interactive
                ' password box when the sheet carries one - which would surface
                ' mid-chain, breaking the quiet convention (contract section 6).
                ' An explicit argument, blank or not, never prompts: a mismatch
                ' raises error 1004 instead, which this catches per sheet.
                ' Assumption about Excel's own behaviour - see the plan's
                ' Unknowns.
                ws.Unprotect Password:=pwd
                If Err.Number <> 0 Then
                    failed.Add ws.name
                    Err.Clear
                Else
                    doneCount = doneCount + 1
                End If
                On Error GoTo 0
            End If
        End If
    Next ws
    Core_SpeedOff

    If doneCount = 0 And failed.Count = 0 Then
        Core_Done ctx, "Nothing was protected, so there was nothing to unprotect.", TITLE
    Else
        Core_Done ctx, "Unprotected " & doneCount & " sheet(s)." & _
               PkgWorkbook_FailedNote(failed, "the password did not match"), TITLE
    End If
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.unprotect_sheets", Err
End Sub
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: still fails (two entries remain missing); the failure list drops
`PkgWorkbook_UnprotectSheets`.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgWorkbook.bas
git commit -m "feat(workbook): workbook.unprotect_sheets"
```

---

### Task 5: `workbook.build_toc`

**Owner: delegate to codex (sonnet).** Mechanical, but Step 1 implements the marker-based
refresh from Decision 7 exactly as specified - do not simplify it back to create-only.

**Files:** Modify `src/packages/PkgWorkbook.bas`.

**Interfaces:** Consumes `Core_Cfg`, `Core_SheetExists`, `Core_DestUsable`, `Core_NewSheet`,
`Core_IsEngineSheet`, `Core_SpeedOn`/`Core_SpeedOff`, `Core_Refuse`, `Core_Done`, `Core_Fail`.

- [ ] **Step 1: Add the action and its marker helper**

```vb
'================================================================
' ACTION - Build (or refresh) a Table of Contents sheet
'================================================================
Public Sub PkgWorkbook_BuildToc(ctx As Object)
    Dim destName As String, ws As Worksheet, dest As Worksheet
    Dim rowOut As Long, n As Long
    Const TITLE As String = "Build Table of Contents"
    On Error GoTo Fail

    destName = Trim(Core_Cfg(ctx, "k_workbook_toc_name"))
    If Len(destName) = 0 Then
        Core_Refuse ctx, "No name is set for the Table of Contents sheet." & vbCrLf & vbCrLf & _
               "Fill 'Name for the Table of Contents sheet' in on the Setup tab, Finalize, and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Core_SheetExists(destName) Then
        Set dest = ThisWorkbook.Worksheets(destName)
        ' Only a sheet THIS action built before is refreshed in place. Anything
        ' else sharing the configured name - a user's own sheet, or a leftover
        ' from before the name was changed - is refused exactly the way
        ' Core_DestUsable refuses any other collision, so a rerun can never
        ' overwrite a sheet it did not build. A create-only version would
        ' refuse on every run after the first - the exact shape of bug the
        ' refusal pass (phase 1) is named for, except here rerunning after
        ' adding a tab is the entire point of the feature.
        If dest.Range("A1").Value <> PkgWorkbook_TocMarker() Then
            Core_Refuse ctx, "A sheet named '" & destName & "' already exists, and it was not " & _
                   "built by this action." & vbCrLf & vbCrLf & _
                   "Rename it, or change 'Name for the Table of Contents sheet' on the Setup " & _
                   "tab, and run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
            Exit Sub
        End If
        Core_SpeedOn
        dest.Cells.Clear
    Else
        If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub
        Core_SpeedOn
        Set dest = Core_NewSheet(destName)
    End If

    dest.Range("A1").Value = PkgWorkbook_TocMarker()
    dest.Range("A1").Font.Bold = True
    dest.Range("A1").Font.Italic = True
    rowOut = 2
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) And StrComp(ws.name, destName, vbTextCompare) <> 0 Then
            ' A hidden or very-hidden sheet is skipped, not merely un-linked: a
            ' hyperlink's SubAddress cannot navigate to a hidden sheet, so a
            ' link to one is a dead click that reads as a bug in this action
            ' rather than as the sheet being hidden on purpose.
            If ws.Visible = xlSheetVisible Then
                dest.Hyperlinks.Add Anchor:=dest.Cells(rowOut, 1), Address:="", _
                    SubAddress:="'" & ws.name & "'!A1", TextToDisplay:=ws.name
                rowOut = rowOut + 1
                n = n + 1
            End If
        End If
    Next ws
    dest.Columns("A").AutoFit
    Core_SpeedOff

    If n = 0 Then
        Core_Done ctx, "Built '" & destName & "', but there were no other sheets to list - " & _
               "every other sheet is either hidden or belongs to the toolkit itself.", TITLE
    Else
        Core_Done ctx, "Built '" & destName & "' with " & n & " link(s)." & vbCrLf & vbCrLf & _
               "A hidden sheet is not listed - linking to one that cannot be navigated to would " & _
               "defeat the point of hiding it. Run this again after adding, removing, or " & _
               "renaming a tab; the list is rebuilt from scratch each time, not appended to.", TITLE
    End If
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.build_toc", Err
End Sub

' The fixed value written to A1 of a Table of Contents sheet this action built,
' so a rerun can tell "my own prior output, safe to clear and rebuild" apart
' from "someone else's sheet that happens to share the configured name" - the
' distinction Core_DestUsable draws for a brand-new name, extended here to the
' refresh case Core_DestUsable does not have to handle, because it never
' overwrites anything. Deliberately human-readable rather than a hidden
' technical token: the user sees exactly why this cell holds what it holds.
Private Function PkgWorkbook_TocMarker() As String
    PkgWorkbook_TocMarker = "SheetForge Table of Contents - rebuilt automatically each run. Do not edit this sheet by hand."
End Function
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: still fails (one entry remains missing); the failure list drops
`PkgWorkbook_BuildToc`.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgWorkbook.bas
git commit -m "feat(workbook): workbook.build_toc"
```

---

### Task 6: `workbook.unhide_rows_cols`

**Owner: delegate to codex (sonnet).** Mechanical, following the confirm-on-surprise shape at
`src/packages/PkgClean.bas:292-310`, gated on one condition instead of always shown.

**Files:** Modify `src/packages/PkgWorkbook.bas`.

**Interfaces:** Consumes `Core_IsEngineSheet`, `Core_SpeedOn`/`Core_SpeedOff`, `Core_Refuse`,
`Core_Done`, `Core_Fail`.

- [ ] **Step 1: Add the action and its counting helper**

```vb
'================================================================
' ACTION - Reset row/column visibility on the active sheet
'================================================================
Public Sub PkgWorkbook_UnhideRowsCols(ctx As Object)
    Dim ws As Worksheet
    Dim rowsHidden As Long, colsHidden As Long
    Const TITLE As String = "Unhide All Rows & Columns"
    On Error GoTo Fail

    Set ws = ActiveSheet
    If Core_IsEngineSheet(ws.name) Then
        Core_Refuse ctx, "'" & ws.name & "' is one of the toolkit's own sheets, not a work " & _
               "sheet." & vbCrLf & vbCrLf & _
               "Some of the toolkit's own rows and columns are hidden on purpose, as part of " & _
               "its own layout - unhiding them here would expose the toolkit's internal " & _
               "working state, not something you can use." & vbCrLf & vbCrLf & _
               "Switch to the sheet you want to reset, then run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Counted across the sheet's USED range only, for the same reason
    ' Core_NarrowToUsed narrows a pick there: a per-row/per-column COM read
    ' across a whole 1,048,576-row sheet freezes the screen for no benefit.
    ' The unhide itself still clears the WHOLE sheet below, so a hidden column
    ' sitting outside the used range is still unhidden - it is just not
    ' counted in the number reported here.
    rowsHidden = PkgWorkbook_CountHidden(ws.UsedRange.Rows)
    colsHidden = PkgWorkbook_CountHidden(ws.UsedRange.Columns)
    If rowsHidden = 0 And colsHidden = 0 Then
        Core_Done ctx, "Nothing was hidden on '" & ws.name & "'.", TITLE
        Exit Sub
    End If

    ' An AutoFilter hides rows through the SAME Range.Hidden property as a
    ' manual Hide - there is no way to tell the two apart, and so no way to
    ' unhide only one kind. Warned explicitly, and only here, rather than
    ' silently changing what the user sees as "the filtered view": without an
    ' active filter, unhiding is not a surprise and this action just does it,
    ' the same as workbook.unhide_sheets.
    If ws.AutoFilterMode Then
        If MsgBox("'" & ws.name & "' has an active AutoFilter." & vbCrLf & vbCrLf & _
               "Unhiding rows will make every filtered-out row visible again - it will look " & _
               "like the filter changed, because the filter's own hidden rows are reset along " & _
               "with any manually hidden ones. There is no way to unhide one kind and not the " & _
               "other." & vbCrLf & vbCrLf & "Re-apply the filter afterward if you want it back." & _
               vbCrLf & vbCrLf & "Continue?", _
               vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub
    End If

    Core_SpeedOn
    ws.Rows.EntireRow.Hidden = False
    ws.Columns.EntireColumn.Hidden = False
    Core_SpeedOff

    Core_Done ctx, "Unhid " & rowsHidden & " row(s) and " & colsHidden & " column(s) on '" & _
           ws.name & "'.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.unhide_rows_cols", Err
End Sub

' Counts hidden rows or columns inside a Rows/Columns collection, one COM read
' per entry. Bounded to the caller's used range - see the comment at the call
' site for why the whole sheet is never walked this way.
Private Function PkgWorkbook_CountHidden(rng As Range) As Long
    Dim one As Range, n As Long
    For Each one In rng
        If one.Hidden Then n = n + 1
    Next one
    PkgWorkbook_CountHidden = n
End Function
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: **exit 0.** All five manifest entries now resolve; `Core + N package(s)` includes
`PkgWorkbook.bas`; no diagnostics.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgWorkbook.bas
git commit -m "feat(workbook): workbook.unhide_rows_cols"
```

---

### Task 7: Contract v2.9.0

**Owner: Opus.** Contract work; the reserved-name normative bullet and the additive/breaking
call (Decision 9) need care, not a template fill-in.

**Files:** Modify `spec/contract.md` - header, §7, §8, §10.

**Interfaces:** Produces the normative text this wave's code (Tasks 1-6) already implements.
Consumes nothing.

- [ ] **Step 1: Bump the version header**

Line 3: `**Contract version: 2.8.0**` -> `**Contract version: 2.9.0**` (the exact prior number
depends on what W5 landed; confirm with `grep -n 'Contract version:' spec/contract.md` before
editing).

- [ ] **Step 2: Confirm `workbook` is active in §7 (no edit expected)**

```bash
grep -n '| `workbook` |' spec/contract.md
```

Expected: one hit, `active`. This was verified during planning (already in the table at line
598 as of this writing) - this step exists to catch drift if an earlier wave's PR touched §7's
table shape. If the domain is missing or not `active`, stop and add it as its own step before
continuing; do not proceed on an assumption.

- [ ] **Step 3: Add the reserved-name normative bullet to §7**

Insert alongside the existing "Reserved (sheet names, v2.2.0)" bullet (grep `Reserved` to find
the block):

````markdown
- **Reserved (v2.9.0, normative):** an implementation creating a NEW destination sheet from a
  `text`-kind config field MUST refuse a value naming a sheet in this section's reserved family,
  or a sheet the platform's own generated-page naming rule claims (the Excel reference
  implementation: `Core_IsEngineSheet`), and MUST NOT check only for an EXISTING sheet of that
  name. Discovered by `workbook.build_toc` (§8): a value like `"Macros 3"` names no sheet today,
  so an exists-only check lets it through, and the sheet it creates is destroyed unconditionally
  by the very next rebuild (the `Macros`/`Macros <n>` rule above) - a data-loss trap that would
  not surface until the rebuild AFTER the one that hid it. This closes the same gap for every
  existing `text`-kind destination field (`k_import_dest` and any added by W2-W5), not only this
  wave's field, because the underlying check is one shared engine function in the reference
  implementation.
````

- [ ] **Step 4: Add the `Package workbook` subsection to §8**

Insert after whichever package subsection W5 added (find the end of §8's package list; this
subsection follows the same shape as `Package import` at line ~710):

````markdown
### Package `workbook` - "Workbook Housekeeping" (v1.0.0, contract v2.9.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_workbook_toc_name` | text | simple | `Contents` | Name of the Table of Contents sheet. A conforming implementation MUST create this sheet on first run and refresh it in place (clear and rewrite) on later runs, using a marker it writes and checks to prove the sheet is its own prior output - it MUST NOT overwrite a sheet by that name it did not build, and MUST NOT refuse solely because a sheet by that name already exists (the create-only shape refuses on every run after the first). Kind is `text`, not `sheet`, for the same reason as `k_import_dest` (§8): a `sheet` field is validated to already exist, and this one must not. |
| `k_workbook_password` | text | advanced | *(blank)* | Password used by `workbook.protect_sheets` and `workbook.unprotect_sheets`. Blank means protect with no password. **Not a secret** - stored as plain text in the config store, like every other field; it deters accidental edits, the same as Excel's own Protect Sheet, and provides no confidentiality. |

**Actions**

| id | caption | semantics |
|---|---|---|
| `workbook.unhide_sheets` | Unhide All Worksheets | Set every hidden or very-hidden worksheet to visible, EXCLUDING every sheet the implementation's engine owns (the Setup surface, the config store and its staging sheet, the package-enable sheet, every generated button page). The engine-owned config store MUST NOT be unhidden by this action under any configuration. |
| `workbook.protect_sheets` | Protect All Sheets | Protect every sheet NOT owned by the engine, with `k_workbook_password` (or no password if blank). An implementation MUST choose protection parameters that do not block its own future ability to remove and replace its own placed controls on a protected surface (Excel reference: `DrawingObjects:=False`) - protecting with the platform's own default MUST NOT be assumed safe. Engine-owned sheets MUST NOT be protected by this action, so a later rebuild retains its own access to them. |
| `workbook.unprotect_sheets` | Unprotect All Sheets | Unprotect every protected sheet NOT owned by the engine, using `k_workbook_password`. A sheet whose protection does not accept that password MUST be reported by name as not unprotected; the action MUST NOT stop partway and MUST attempt every remaining sheet regardless of one sheet's failure. |
| `workbook.build_toc` | Build Table of Contents | Create or refresh (§8 field semantics above) a sheet named `k_workbook_toc_name` listing every VISIBLE sheet not owned by the engine and not the Table of Contents sheet itself, each as a same-workbook hyperlink to that sheet. A hidden sheet MUST NOT be listed (a link to it cannot be navigated to). The destination name MUST be refused per the reserved-name rule (§7) and per an existing-and-not-mine collision, exactly as any other engine-created destination sheet. |
| `workbook.unhide_rows_cols` | Unhide All Rows & Columns | On the ACTIVE sheet only, set every row and every column to visible. MUST refuse if the active sheet is engine-owned (an engine-owned surface may hide rows/columns as part of its own generated layout). Because a filter-hidden row and a manually hidden row are not distinguishable, an implementation MUST warn, and obtain confirmation, before unhiding on a sheet with an active AutoFilter; outside that case no confirmation is required. |
````

- [ ] **Step 5: Add the §10 version bullet**

Insert as the newest versioned bullet in §10:

````markdown
- **v2.9.0 (additive):** the `workbook` domain's package, fields, and five actions (§8); one
  new reserved-name normative rule for `text`-kind destination fields (§7). The reserved-name
  rule reclassifies one previously-succeeding configuration (a destination name inside the
  reserved family, e.g. `"Macros 3"`, on a field that had no prior existing-name collision) as a
  refusal. Classed additive on the same reasoning v2.4.0 used for the chain-halt rule: the
  configuration being refused was never actually usable - it produced a sheet the very next
  rebuild destroyed unconditionally - so refusing it prevents a data-loss trap rather than
  withdrawing behaviour a deployment could rely on. `SCHEMA_VER` does not move.
````

- [ ] **Step 6: Verify**

```bash
grep -n 'Contract version: 2.9.0' spec/contract.md
grep -c 'workbook' spec/contract.md
pwsh build/build.ps1
```

Expected: the version line present once; a non-zero count; `build.ps1` still exits 0 (it
re-parses §7's vocabulary table on every run, so a malformed edit here fails the build
immediately, same as every prior contract task in this track).

- [ ] **Step 7: Commit**

```bash
git add spec/contract.md
git commit -m "spec: contract v2.9.0 - workbook package, and the reserved-name rule for destination fields"
```

---

### Task 8: Version, changelog, code map, deferred-validation log, and the PR

**Owner: Opus.**

**Files:** `VERSION`, `CHANGELOG.md`, `docs/code-map.md`,
`docs/designs/2026-07-28-deferred-validation-log.md`.

- [ ] **Step 1: `VERSION` -> `0.17.0`, and a `## [0.17.0]` changelog entry**

House style: state what changed for the user, and the failure each safety choice prevents -
match the tone of the `[0.10.0]` entry already in `CHANGELOG.md`. Cover: the five new actions;
that `_Config`/`_Config_staged` are never unhidden; that protection deliberately leaves the
engine's own sheets alone so a rebuild keeps working; that the password is plainly not a secret;
that the Table of Contents refreshes in place rather than refusing on every rerun.

- [ ] **Step 2: Refresh the code map**

```bash
node scripts/code-map.mjs
```

Run before committing; regenerating after a commit only matters if the stamp itself is being
checked.

- [ ] **Step 3: Append this wave's run-time-only claims to the deferred validation log**

Add a `## W6 - Workbook housekeeping` section to
`docs/designs/2026-07-28-deferred-validation-log.md` (create the file if phase 1's Task 5' has
not landed it yet - do not wait) with these bullets, each stating the claim and why the build
cannot confirm it:

- `Core_DestUsable` refuses a value naming a reserved/engine sheet (e.g. `"Macros 3"`,
  `"_Config"`) BEFORE any sheet exists there. Unconfirmed at runtime - the build only proves the
  branch compiles and is reachable, not that `Core_IsEngineSheet`'s string match behaves as
  believed against a real value typed into Setup.
- `workbook.unhide_sheets` never unhides `_Config` or `_Config_staged` against a real workbook
  where those sheets are actually hidden.
- `workbook.protect_sheets`'s `DrawingObjects:=False` sheets remain manageable by a real
  Finalize/rebuild afterward - i.e. that this choice actually prevents the stranded-control
  failure it is designed to prevent, not merely that it protects the sheet.
- `workbook.unprotect_sheets`'s explicit `Password:=` argument avoids Excel's own interactive
  password prompt on a protected sheet, rather than showing it. Marked as an assumption in the
  plan (`CLAUDE.md` §7); this is the item that assumption is about.
- `workbook.build_toc`'s marker check correctly distinguishes "my own prior TOC, safe to clear"
  from "someone else's sheet with the same configured name" against a real second run.
- `workbook.unhide_rows_cols`'s AutoFilter warning fires when, and only when, `AutoFilterMode`
  is genuinely True on a real filtered sheet.
- `Select Case LCase(<Const>)` inside `Core_IsEngineSheet` (phase 1) compiles and behaves as
  believed - carried forward from phase 1's own Unknowns, not new to this wave, but every
  action in this file depends on it being true.

- [ ] **Step 4: Note that phase 8 is next**

At the end of the same log section, add one line: "W6 is the last wave (plan-of-record §2).
Phase 8 - the single end-to-end validation over the finished workbook - is next, and this log
is its input. This wave does not write a validation document of its own."

- [ ] **Step 5: Final gates**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas
grep -nE '^\s*Private\s+(Sub|Function)\s+' src/packages/*.bas | grep -vE 'Private\s+(Sub|Function)\s+Pkg'
git status --short
```

Expected: exit 0; no diagnostics; `1`; no output from the prefix grep (phase 2's guard would
otherwise fail the build itself, so this is a belt-and-suspenders check); a tree clean but for
the untracked `.vscode/`.

- [ ] **Step 6: Review the diff, then one PR against `main`**

Not stacked on any other wave's branch - branch from `main` after W5 has merged. State in the
PR body, in one line each: which sheets `unhide_sheets` will never touch and why; that
`protect_sheets` deliberately excludes the engine's own sheets and why; that the password is
plainly not a secret; that `build_toc` refreshes rather than refusing on every rerun; and that
this is the last wave, with phase 8 next.

---

## Unknowns and risks

- **This wave's real risk is a false sense of safety from static verification.** Every one of
  the five actions makes a claim about Excel's own runtime behaviour (`Core_IsEngineSheet`'s
  string match, `DrawingObjects:=False`'s effect on a later rebuild, `Unprotect`'s prompt
  behaviour, a hyperlink's `SubAddress` against a hidden sheet, `AutoFilterMode`'s reliability).
  `build.ps1` and the syntax check can only prove the code compiles and is reachable - none of
  these claims are confirmed until phase 8's real run. Task 8 exists to make sure none of them
  is silently dropped between now and then.
- **`Core_DestUsable`'s reserved-name check changes behaviour for every existing caller**, not
  just this wave's TOC action. If W2-W5 shipped a destination-name field whose default or a real
  deployment's configured value happens to collide with the reserved family, that caller starts
  refusing where it previously (silently, incorrectly) succeeded. This is the intended fix
  (Decision 9), but it means phase 2's `Core_DestUsable` diff
  (`2026-07-28-core-helper-promotion-plan.md`, Task 2 Step 3) should be reviewed with W2-W5's
  actual shipped field defaults in hand, not assumed safe from this plan alone.
- **The TOC marker is a plain sentence in cell A1, not a hidden or protected cell.** A user who
  edits or deletes that cell by hand (encouraged not to, by the sentence itself) turns the next
  run into a refusal ("not built by this action") rather than a silent misfire - a deliberate
  fail-closed choice, but it means a curious user poking at the TOC sheet can make the action
  stop working until they either restore the marker or rename the sheet. Not a defect; stated so
  it is not read as one when it happens.
- **`PkgWorkbook_CountHidden` bounds its count to `UsedRange`.** The unhide operation itself
  covers the whole sheet (`Rows.EntireRow.Hidden = False`), so nothing is left hidden - but a
  column hidden far outside the used range is unhidden without being counted in the completion
  message, which will then slightly undercount what actually changed. Accepted as the same
  trade-off `Core_NarrowToUsed` already makes, not treated as a bug.
- **Not confirmed by this plan: whether `ws.Protect` on an already-protected sheet reliably
  raises an error** (rather than silently no-op'ing or silently changing its parameters) for the
  per-sheet failure tracking in `workbook.protect_sheets` to behave as designed. Believed true
  from ordinary VBA behaviour; logged in Task 8 alongside the `Unprotect`-prompt assumption
  rather than asserted as fact.
