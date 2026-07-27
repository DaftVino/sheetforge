# v0.4.0 track 1 — the four owner-gate items — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development`
> (recommended) or `superpowers:executing-plans` to implement this plan task-by-task.
> Steps use checkbox (`- [ ]`) syntax for tracking.

**Spec:** `docs/designs/2026-07-25-track1-owner-feedback-design.md` — normative for this
plan. Where this plan and the design disagree, the design wins **except** at the two
points explicitly flagged `DEVIATION` below, both of which carry the owner's decision
from this session.

**Goal:** Land the four items the owner's manual Excel gate (Task 10) produced, so
Task 11 (teardown + ship) unblocks.

**Architecture:** All engine work is in the single `src/core/Core.bas` module plus three
lines and one manifest field in `src/packages/PkgClean.bas`. One new non-build file,
`src/attic/HomeIndex.bas.txt`, preserves the withdrawn index feature. `spec/contract.md`
§3, §6, §7, §8, §10 are amended in the tasks that cause the change, never in a separate
documentation pass — the ledger records four review rounds lost to strings that drifted
from the behaviour they described.

**Tech Stack:** VBA (Excel/Windows, late-bound `Scripting.Dictionary`), PowerShell 7
(`build/build.ps1`), Node (`scripts/vba-syntax-check.mjs`, `scripts/code-map.mjs`).

---

## Global Constraints

Every task's requirements implicitly include this section.

- **`Option Explicit` everywhere.** Every new local is declared.
- **Module-level `Const`/`Dim` lives ONLY in the declarations section** — above
  `Core_DeclarePackage` at `Core.bas:144`. A `Const` between procedures is a compile
  error, and `build.ps1` fails the build on a declaration after the first procedure.
- **Flat namespace (CLAUDE.md rule 4).** Anything `Public` must be `Core_*` or
  `Pkg<Name>_*`. Everything else is `Private`. Shared helpers live ONLY in Core.
- **No VBA reserved word as a `Dim`/parameter identifier.** `build.ps1:160-166` fails the
  build on the list; every identifier introduced by this plan was checked against it.
  **Do not add words to that list** — it is conservative and empirical by design (do-not-revert item 8).
- **ASCII only** in every `.bas`. `build.ps1` fails on a non-ASCII byte. No smart quotes,
  no em dashes, no `→`. Use `-` and `->`.
- **25 line-continuation limit per statement.** One `.Add` per line.
- **The speed pattern.** Any new action-level work wraps in `Core_SpeedOn` /
  `Core_SpeedOff` with `RestoreSpeed` guaranteed in the error handler. No task here adds a
  new top-level entry point, so no task adds a new speed wrapper.
- **Iterate `Shapes` backwards** when deleting — the established idiom at
  `EnsurePanelSheet` (`Core.bas:2613`).
- **`src/` is the source of truth.** A `.bas` edit does nothing in Excel until the bundle
  is rebuilt and re-imported. There is no round-trip back to `src/`.
- **No contract version bump.** v1.2.0 is unreleased; every change here is additive or
  engine presentation under §10's two-axis rule.
- **TDD is formally OFF for this repo** (CLAUDE.md, "Off"). There is no headless Excel, no
  VBA compiler on the agent side, and no test suite. **Do not propose adding one.**
  Verification per task is the two structural gates below plus the owner's manual import.

**The two structural gates, run after every task:**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
```

They are **structural only**. They cannot resolve a call target, see a wrong argument
count, or catch a type error. This session's own history is the proof: a
`Dim ... any As Boolean` compile error passed four review rounds and both gates, and only
Excel found it. Never describe a task as verified on the strength of these alone.

---

## File Structure

| File | Status | Responsibility |
|---|---|---|
| `src/core/Core.bas` | modify | All engine work. ~3300 lines — **grep, never read whole.** |
| `src/packages/PkgClean.bas` | modify | Item 1's package-side changes (8K, safe to read whole). |
| `src/attic/HomeIndex.bas.txt` | **create** | The withdrawn index feature, preserved verbatim. `.txt` so `build.ps1`'s `Pkg*.bas` glob and its ASCII/reserved-word gates cannot reach it. |
| `spec/contract.md` | modify | §3, §6, §7, §8, §10. |
| `docs/designs/2026-07-24-finish-excel-track1-plan.md` | modify | Task 10 section only — the checklist being re-cut. |
| `docs/code-map.md` | regenerate | Once, at the end. |

**`src/attic/` does not exist yet** — Task 5 creates it. `build.ps1` globs only
`src/packages/Pkg*.bas` (verified at `build.ps1:95` and `:121`), so nothing under
`src/attic/` can reach a bundle or a gate.

---

## Task order and why

**1 → 2 → 3 → 4 → 5 → 6 → 7 → 8 → 9**, implementing design items **1 → 2 → 4 → 3**.

Item 4 (remove `Macros Home`, Task 5) runs before item 3 (`Sheet!Cell` placement,
Tasks 6-9) because item 3 must write a placed-button count into the completion message
and must reason about "zero generated pages" — and item 4 is what creates both. Doing
item 3 first means writing that code twice and reviewing it twice. **Priority order and
build order are different things here.**

Item 3's own split (Tasks 6-9) is by what a reviewer can accept or reject independently:
the `Placement` relabel, the grammar, the rendering, the validation.

---

## Interfaces introduced (whole-plan summary)

A task's implementer sees only their own task. This is the single index of names and
signatures crossing task boundaries.

| Symbol | Kind | Task | Signature |
|---|---|---|---|
| `Core_RangeRef` | public function | 1 | `(rng As Range) As String` |
| `NeedsSheetQuotes` | private function | 1 | `(ByVal nm As String) As Boolean` |
| `Core_CfgBool` | public function | 2 | `(ctx As Object, ByVal key As String) As Boolean` |
| `ParseBool` | private function | 2 | `(ByVal raw As String, ByRef flag As Boolean) As Boolean` |
| `BoolLabel` | private function | 2 | `(ByVal raw As String) As String` |
| `BoolDropdown` | private sub | 2 | `(rng As Range)` |
| `mRegCache` / `mRegCacheOk` | module vars | 3 | `Collection` / `Boolean` |
| `InvalidateRegistry` | private sub | 3 | `()` |
| `LandOnFirstPage` | private sub | 5 | `(layout As Object)` |
| `DeleteHomePanel` | private sub | 5 | `()` |
| `PLACE_HEADER` | private const | 6 | `String = "Placement"` |
| `CodeKind` | private function | 7 | `(ByVal code As String) As String` |
| `ParseAnchor` | private function | 7 | `(ByVal code As String, ByRef sheetName As String, ByRef cellRef As String) As Boolean` |
| `IsA1Ref` | private function | 7 | `(ByVal s As String) As Boolean` |
| `PLACED_PREFIX` | private const | 8 | `String = "SFPlaced_"` |
| `ANCHOR_BTN_W` | private const | 8 | `Double = 110` |
| `SafeShapeName` | private function | 8 | `(ByVal s As String) As String` |
| `SplitAnchored` | private sub | 8 | `(buttons As Collection, codeMap As Object, ByRef paged As Collection, ByRef anchored As Collection)` |
| `SweepPlacedButtons` | private sub | 8 | `()` |
| `AnchorSheet` | private function | 8 | `(ByVal code As String, ByRef cellRef As String) As Worksheet` |
| `PlaceAnchoredButtons` | private function | 8 | `(anchored As Collection, codeMap As Object) As Long` |
| `AddBtn` | private sub | 8 | gains `Optional ByVal place As XlPlacement = xlMoveAndSize` |
| `IsEngineOwnedSheet` | private function | 9 | `(ByVal nm As String) As Boolean` |
| `AnchorAddrKey` | private function | 9 | `(ByVal sheetName As String, ByVal cellRef As String) As String` |

---

## Two DEVIATIONS from the design, both owner-decided this session

**DEVIATION 1 — `ANCHOR_BTN_W` is 110pt, not 150pt.** The design proposed 150pt and
flagged it as needing the owner's eyes. The owner chose **110pt** (compact), so an
anchored button stays closer to a single default-width column and is less intrusive
next to real data. `Placement = xlMove` is confirmed unchanged from the design.

**DEVIATION 2 — the `bool` store spelling.** The design says §3 should record the
encoding as "`yes` / `no`, canonical lowercase" while ALSO saying the renderer mirrors
`AvailDropdown`, which offers and stores the title-case display spelling. Those two
cannot both be literally true. This plan resolves it exactly the way `availability`
already resolves it (§6: "abstract tokens, not a mandated on-screen spelling"):

- the **canonical tokens** are `yes` / `no`;
- a conforming **reader MUST be case-insensitive**;
- the Excel reference implementation **presents and stores `Yes` / `No`**.

§3's row is written to say that. If the owner prefers a literally-lowercase store, it is a
one-line change in `BoolDropdown`'s `Formula1` and one in `BoolLabel`.

---

## Do NOT revert

The eight items under "Do NOT revert" in
`docs/designs/2026-07-25-track1-owner-feedback-handoff.md` are earned across nine tasks
and ~20 review rounds. Four bear on this plan and are restated at the point they apply:

1. **The `hide` filter lives in `GatherButtons` and nowhere else.** Task 8's
   `SplitAnchored` runs strictly downstream of it. Do not add a second filter, and do not
   assemble either half any other way. (Task 8, Step 4.)
2. **Nothing on the card grid may be merged** — AutoFit ignores merged cells. Untouched
   by this plan; do not "tidy" it.
3. **`build.ps1`'s reserved-word list stays conservative and empirical.** Add a word only
   when Excel rejects one.
4. **No `Core_SheetExists` guard on a nav-link target.** One was added once, and because
   the target is built last it made every link dead text on a first build. Task 5's
   `LandOnFirstPage` does use `Core_SheetExists`, but at **run** time on an
   already-rendered sheet, not at **render** time on a not-yet-built one — a different
   thing, and Step 4 of that task says so in the code comment.

---

## Task 1: Sheet-qualified range references

Design item 1a. Three lines in `PkgClean` plus one shared helper in Core.

**Scope discipline — read this before touching anything.** The input brief said to
sheet-qualify "every message that names a range, across both packages and Core." That was
**wrong and is not to be reverted to.** Verified by grep: `.Address(` appears only at
`PkgClean.bas:45`, `:81`, `:112`. Every `PkgImportExport` message already names its
configured tab or file explicitly (`mDest`, `mSheet`, `tlist`, `outPath`), and Core has
none. **Do not go hunting in `PkgImportExport`** — there is nothing there, and "fixing" a
message that already names its tab would make it worse.

**Files:**
- Modify: `src/core/Core.bas` — insert after `Core_SheetExists`'s `End Function`
  (currently `Core.bas:558`+)
- Modify: `src/packages/PkgClean.bas:45`, `:81`, `:112`

**Interfaces:**
- Consumes: nothing.
- Produces: `Public Function Core_RangeRef(rng As Range) As String`

- [ ] **Step 1: Confirm the blast radius before editing**

Run:

```
grep -rn '\.Address(' src/
```

Expected: exactly three hits, all in `src/packages/PkgClean.bas` (lines 45, 81, 112).
If `src/packages/PkgImportExport.bas` appears, **stop and report** — the design's
diagnosis is wrong and this task needs re-scoping.

- [ ] **Step 2: Add `Core_RangeRef` and its helper to Core**

Insert immediately after `Core_SheetExists`'s `End Function` in `src/core/Core.bas`:

```vba
' "TabName!M2:M12", or "'TabName Data'!M2:M12" when the sheet name needs quoting.
' Range.Address omits the sheet name BY DESIGN, so every message that names a
' range must qualify it here or a cross-tab run is indistinguishable from a
' same-tab one - which is exactly what the owner hit at the v0.4.0 gate: three
' PkgClean actions reported "M2:M12" whether they had acted on TabName or on the
' button's own tab.
' Lives in Core, not PkgClean: CLAUDE.md rule 4 puts shared helpers only in Core,
' and a copy in a package would be a standards deviation.
Public Function Core_RangeRef(rng As Range) As String
    Dim nm As String
    If rng Is Nothing Then Exit Function
    nm = rng.Parent.Name
    If NeedsSheetQuotes(nm) Then nm = "'" & Replace(nm, "'", "''") & "'"
    Core_RangeRef = nm & "!" & rng.Address(False, False)
End Function

' Excel's own rule: a sheet name is written bare only when it is entirely
' A-Z / a-z / 0-9 / "_" and does not begin with a digit. Anything else is
' quoted, and an embedded apostrophe is doubled inside the quotes.
Private Function NeedsSheetQuotes(ByVal nm As String) As Boolean
    Dim i As Long, ch As String
    If Len(nm) = 0 Then NeedsSheetQuotes = True: Exit Function
    ch = Left$(nm, 1)
    If ch >= "0" And ch <= "9" Then NeedsSheetQuotes = True: Exit Function
    For i = 1 To Len(nm)
        ch = Mid$(nm, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") _
                Or (ch >= "0" And ch <= "9") Or ch = "_") Then
            NeedsSheetQuotes = True
            Exit Function
        End If
    Next i
End Function
```

- [ ] **Step 3: Qualify the three PkgClean messages**

`src/packages/PkgClean.bas:45` — replace:

```vba
    Core_Done ctx, "Remove Duplicates complete on " & rng.Address(False, False) & _
           " using key column(s) " & keys & ".", "Remove Duplicates"
```

with:

```vba
    Core_Done ctx, "Remove Duplicates complete on " & Core_RangeRef(rng) & _
           " using key column(s) " & keys & ".", "Remove Duplicates"
```

`src/packages/PkgClean.bas:81` — replace:

```vba
    Core_Done ctx, "Trimmed & cleaned text on " & rng.Address(False, False) & ".", _
           "Trim & Clean Text"
```

with:

```vba
    Core_Done ctx, "Trimmed & cleaned text on " & Core_RangeRef(rng) & ".", _
           "Trim & Clean Text"
```

`src/packages/PkgClean.bas:112` — replace:

```vba
    Core_Done ctx, "Filled blanks down on " & rng.Address(False, False) & ".", _
           "Fill Blanks Down"
```

with:

```vba
    Core_Done ctx, "Filled blanks down on " & Core_RangeRef(rng) & ".", _
           "Fill Blanks Down"
```

- [ ] **Step 4: Run both structural gates**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
```

Expected: all four succeed. A duplicate-proc-name failure would mean `Core_RangeRef` or
`NeedsSheetQuotes` collides with an existing name — rename rather than deleting the other.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas src/packages/PkgClean.bas
git commit -m "fix(clean): sheet-qualify every range named in a completion message

Range.Address omits the sheet name by design, so all three PkgClean actions
reported 'M2:M12' whether they ran on TabName or on the button's own tab. New
Core_RangeRef qualifies (and quotes) the parent sheet.

Nothing was ever acting on the wrong sheet - Application.InputBox(Type:=8)
returns a Range carrying its own parent, and TrimClean/FillDown address every
cell through rng.Cells(r, c). The message just could not say which sheet."
```

---

## Task 2: The `bool` kind's first use, and `clean_hasheader`

Design items 1b + 1c. This is the contract-reserved **first use** of the `bool` kind, so
fixing its encoding is contract-mandated work (§3, §10), not scope creep. The owner
approved the `bool` path over the cheaper `text` one with that cost stated.

**Files:**
- Modify: `src/core/Core.bas` — declarations region untouched; five insertions (see steps)
- Modify: `src/packages/PkgClean.bas` — manifest, `PkgClean_RemoveDupes`, `AskRange`
- Modify: `spec/contract.md` — §3, §8, §10

**Interfaces:**
- Consumes: nothing from Task 1.
- Produces: `Public Function Core_CfgBool(ctx As Object, ByVal key As String) As Boolean`;
  private `ParseBool`, `BoolLabel`, `BoolDropdown`.

- [ ] **Step 1: Add the reader and parser to Core**

Insert immediately after `Core_CfgTable`'s `End Function` (currently `Core.bas:314`):

```vba
' Reads a "bool" field (contract section 3; clean_hasheader is its first use).
'
' A BLANK cell means "never set" and resolves to the field's DECLARED manifest
' default - not a hardcoded False and not an optional caller-supplied fallback.
' That fallback is already free: Core_Cfg -> ScalarValue returns
' FieldDefaultReg(ctx("reg"), key) when the stored value is empty. The declared
' default is on the manifest; two places to state it would be two places to
' disagree.
'
' An unrecognised NON-blank value fails closed to False. It cannot normally get
' this far - Core_FinalizeSetup blocks activation on one, naming the field - but
' a reader that guessed would be exactly the "user typed 'ye' and got a silent
' no" failure the availability parser exists to prevent.
Public Function Core_CfgBool(ctx As Object, ByVal key As String) As Boolean
    Dim flag As Boolean
    If ParseBool(Core_Cfg(ctx, key), flag) Then Core_CfgBool = flag
End Function
```

Then insert this **next to `ParseAvail`** (search for `Private Function ParseAvail`; put
these three procedures directly after it, so the two value parsers sit together):

```vba
' "yes"/"y"/"true"/"1" -> True; "no"/"n"/"false"/"0" -> False. Returns False for
' anything else, so every caller fails CLOSED.
'
' Deliberately NOT IsTrue: that one accepts only TRUE/1 and fails OPEN to False
' for everything else, which is right for the optional _Packages enable column
' and wrong here.
'
' The spellings beyond "yes"/"no" are a REFERENCE-IMPLEMENTATION convenience
' (contract section 3), mirroring the same tolerance clause section 6 already
' grants availability. A port MAY accept only "yes" / "no".
Private Function ParseBool(ByVal raw As String, ByRef flag As Boolean) As Boolean
    Select Case LCase(Trim(raw))
        Case "yes", "y", "true", "1"
            flag = True
            ParseBool = True
        Case "no", "n", "false", "0"
            flag = False
            ParseBool = True
    End Select
End Function

' Display spelling for a bool cell, mirroring AvailLabel. An UNRECOGNISED value
' is handed back untouched rather than normalised - Finalize has to be able to
' see and block on what the user actually typed.
Private Function BoolLabel(ByVal raw As String) As String
    Dim flag As Boolean
    If ParseBool(raw, flag) Then
        BoolLabel = IIf(flag, "Yes", "No")
    Else
        BoolLabel = raw
    End If
End Function

' A validated dropdown removes free-typing from a two-value cell. Wrapped in
' On Error Resume Next for the same reason AvailDropdown is: validation can be
' refused (a protected sheet, some locales) and a refused dropdown must not abort
' the build - ParseBool still guards the value, and Finalize still blocks on
' anything unrecognised.
Private Sub BoolDropdown(rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                       Operator:=xlBetween, Formula1:="Yes,No"
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub
```

- [ ] **Step 2: Render a `bool` field as a Yes/No dropdown**

`Core.bas:1384-1392`, `RenderScalar`. Replace:

```vba
Private Function RenderScalar(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    ws.Cells(r, KEY_COL).Value = f("key")
    LabelCell ws.Cells(r, 2), CStr(f("label"))
    ws.Cells(r, VAL_COL).Value = ScalarValue(reg, CStr(f("key")))
    Yellow ws.Cells(r, VAL_COL)
    r = r + 1
```

with:

```vba
Private Function RenderScalar(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    ws.Cells(r, KEY_COL).Value = f("key")
    LabelCell ws.Cells(r, 2), CStr(f("label"))
    If LCase(CStr(f("kind"))) = "bool" Then
        ' Write the DISPLAY spelling, exactly as RenderPlacementBlock writes
        ' AvailLabel rather than the raw token - otherwise a seeded lowercase
        ' "yes" would sit in a cell whose dropdown offers "Yes", which reads as
        ' a value the validation does not recognise.
        ws.Cells(r, VAL_COL).Value = BoolLabel(ScalarValue(reg, CStr(f("key"))))
        Yellow ws.Cells(r, VAL_COL)
        BoolDropdown ws.Cells(r, VAL_COL)
    Else
        ws.Cells(r, VAL_COL).Value = ScalarValue(reg, CStr(f("key")))
        Yellow ws.Cells(r, VAL_COL)
    End If
    r = r + 1
```

`RenderField` (`Core.bas:1283`) needs **no change** — its `Case Else` already routes
`bool` to `RenderScalar`.

- [ ] **Step 3: Make Finalize BLOCK on an unrecognised bool**

This must add to `structural`, not to `warn`. `warn` is the soft "item(s) to note"
counter; only `structural > 0` blocks activation (`Core.bas:2363`).

In `Core_FinalizeSetup`, extend the `Dim` block at `Core.bas:2315-2316` to:

```vba
    Dim ws As Worksheet, cfg As Worksheet, reg As Object, f As Object
    Dim rpt As String, warn As Long, v As String, msg As String
    Dim boolBad As String, boolBadN As Long, bTmp As Boolean
```

In the same procedure's `Select Case f("kind")` (`Core.bas:2330-2345`), add a third case
after `Case "file"`'s `End If`:

```vba
                Case "bool"
                    ' STRUCTURAL, not a soft note: a user who typed "ye" meaning
                    ' "yes" must not be silently handed a "no". Contract section 3.
                    If v <> "" Then
                        If Not ParseBool(v, bTmp) Then
                            boolBad = boolBad & "  - '" & f("label") & "': '" & v & "'" & vbCrLf
                            boolBadN = boolBadN + 1
                        End If
                    End If
```

Immediately after `composeRep = ComposeReport(ws, structural)` (`Core.bas:2361`) and
**before** the `msg = msg & ...` line, insert:

```vba
    If boolBadN > 0 Then
        composeRep = composeRep & vbCrLf & "UNRECOGNISED Yes/No value (use Yes or No):" & _
                     vbCrLf & boolBad
        structural = structural + boolBadN
    End If
```

- [ ] **Step 4: Declare `clean_hasheader` and use it**

`src/packages/PkgClean.bas:23-24`, immediately after the `clean_keycols` field, add:

```vba
    Core_AddField reg, "clean_hasheader", "Does the range you pick include a header row?", "bool", "yes", _
        help:="Yes: the first row of your selection is treated as a header and is never removed. No: every row in the selection is considered, including the first one."
```

`src/packages/PkgClean.bas:26` — the `clean.dupes` help text still says "(include the
header row)", which is now a lie half the time. Replace:

```vba
    Core_AddAction reg, "clean.dupes", "Remove Duplicates", "PkgClean_RemoveDupes", "Pick a range (include the header row); removes duplicate rows using the key column(s) from Setup."
```

with:

```vba
    Core_AddAction reg, "clean.dupes", "Remove Duplicates", "PkgClean_RemoveDupes", "Pick a range; removes duplicate rows using the key column(s) from Setup. Whether the first row is treated as a header is the 'header row' setting on Setup."
```

`src/packages/PkgClean.bas:42` — replace:

```vba
    rng.RemoveDuplicates Columns:=ParseCols(keys), Header:=xlYes
```

with:

```vba
    rng.RemoveDuplicates Columns:=ParseCols(keys), _
                         Header:=IIf(Core_CfgBool(ctx, "clean_hasheader"), xlYes, xlNo)
```

`src/packages/PkgClean.bas:129` — `AskRange`'s prompt is shared by all three actions and
its "(include the header row)" is now wrong for two of them and conditional for the third.
Replace:

```vba
        Prompt:="Select the range for '" & title & "' (include the header row):", _
```

with:

```vba
        Prompt:="Select the range for '" & title & "':", _
```

- [ ] **Step 5: Amend contract §3**

`spec/contract.md:104` — replace the `bool` row:

```
| `bool` | yes/no | **reserved** | — | **Reserved: defined but not yet used by any package.** Encoding (`yes`/`no` vs `TRUE`/`FALSE`) is fixed on first use and this table updated then. |
```

with:

```
| `bool` | yes/no | `yes` / `no` (canonical tokens; a reader MUST be case-insensitive) | validated Yes/No dropdown; a **blank** cell means "never set" and reads as the field's declared default | Encoding fixed at v1.2.0 by `clean_hasheader`, the kind's first use. Like `availability` (§6), the tokens are abstract, not a mandated on-screen spelling: the Excel reference implementation presents and stores `Yes` / `No`. A port MAY accept only `yes` / `no`; the reference additionally tolerates `y` / `true` / `1` and `n` / `false` / `0`. An unrecognised non-blank value MUST block Finalize, naming the field — it MUST NOT fall back to either state. |
```

- [ ] **Step 6: Amend contract §8, package `clean`**

`spec/contract.md:494`, in the `clean` **Fields** table, add a row after `clean_keycols`:

```
| `clean_hasheader` | bool | simple | `yes` | Whether the picked range's first row is a header. `yes` (default) excludes it from the duplicate comparison and never removes it; `no` considers every row in the selection. |
```

`spec/contract.md:503`, replace the `clean.dupes` semantics cell:

```
| `clean.dupes` | Remove Duplicates | Prompt for a range (incl. header row); remove duplicate rows using `clean_keycols`. |
```

with:

```
| `clean.dupes` | Remove Duplicates | Prompt for a range; remove duplicate rows using `clean_keycols`. `clean_hasheader` decides whether the range's first row is treated as a header (excluded from comparison and never removed) or as data. |
```

- [ ] **Step 7: Amend contract §10**

`spec/contract.md:557-558` — replace:

```
- **Reserved for first use:** the `bool` kind (§3) — its encoding is fixed, and
  this document updated, when the first package uses it.
```

with:

```
- **Reserved for first use — DISCHARGED at v1.2.0.** The `bool` kind (§3) had its
  encoding reserved until first use. `clean_hasheader` (package `clean`, §8) is that
  first use, so §3's row now states the fixed encoding and this bullet is a record,
  not an outstanding obligation.
```

In the **v1.2.0 (additive)** bullet (`spec/contract.md:563`), after "the optional field
`usedBy` attribute," insert:

```
the `clean_hasheader` field and the now-fixed `bool` encoding (§3),
```

Then, in the same bullet's closing sentence about what is additive, no change is needed:
`clean_hasheader` is a **new** field with a default that reproduces today's behaviour
exactly, so no v1.0.0/v1.1.0 config is reinterpreted and §10's "changing an action's
semantics is breaking" does not fire. **Do not bump the contract version** — v1.2.0 is
unreleased.

- [ ] **Step 8: Run both structural gates**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
```

- [ ] **Step 9: Commit**

```bash
git add src/core/Core.bas src/packages/PkgClean.bas spec/contract.md
git commit -m "feat(clean): add clean_hasheader, fixing the bool kind's encoding

PkgClean_RemoveDupes hardcoded Header:=xlYes, so a user selecting M2:M12
intending eleven values considered got only M3:M12 deduped - very likely the
owner's 'works if there is something in M1' observation.

This is the contract-reserved FIRST USE of the bool kind, so contract 3's
encoding is fixed here too (yes/no, case-insensitive; Excel presents Yes/No),
with the reader, parser, validated dropdown and Finalize block that go with it.

Additive: the default 'yes' reproduces today's behaviour exactly, so no
v1.0.0/v1.1.0 config is reinterpreted. Contract 3, 8 and 10 amended; no
version bump (v1.2.0 is unreleased)."
```

---

## Task 3: Memoize `BuildRegistry`, collapse `AvailDropdown`

Design item 2. Both changes are **behaviour-neutral**. No contract change.

The owner measured the first Setup build at several minutes with the 82-action fixture.
`BuildRegistry` is uncached and is called from `RenderOffBand` (twice), `HasAnyTier`'s
per-package loops, `ActionCaptionFor` and `ActionTokenFor` **once per chain step**,
`StepsInline` once per step, and `ActionIdSet` **twice** inside a single `ComposeReport`.

**Files:**
- Modify: `src/core/Core.bas` — declarations section, `BuildRegistry`, three entry points,
  `AvailDropdown`, `RenderPlacementBlock`

**Interfaces:**
- Consumes: nothing.
- Produces: `mRegCache`, `mRegCacheOk`, `Private Sub InvalidateRegistry()`.
  `BuildRegistry`'s signature is unchanged.

- [ ] **Step 1: THE STOP GATE — re-verify that no consumer mutates a `reg`**

This is a hard gate. A cache handed out by reference is exactly the shape where being
wrong is silent.

Run:

```
grep -rn 'reg\.Add\|reg\.Remove\|reg("[a-zA-Z]*") *=\|reg("fields")\.Add\|reg("actions")\.Add' src/
```

Expected: hits **only** at `src/core/Core.bas:146-224` (inside `Core_DeclarePackage`,
`Core_AddSection`, `Core_AddField`, `Core_AddList`, `Core_AddTable`, `Core_AddDefault`,
`Core_AddAction`) plus two comment lines at `:26-27`. Every one of those runs *during*
`TryManifest`, inside `BuildRegistry` itself, against a dictionary it has just created.

Then confirm the read-only side:

```
grep -rn 'For Each reg In BuildRegistry()' src/ | wc -l
```

Expected: 13.

**If any mutation appears outside `Core.bas:146-224`, STOP. Do not implement this task.
Report the exact line.** Memoizing is unsafe and the design's premise is broken.

- [ ] **Step 2: Add the cache to the declarations section**

Insert immediately after `Private mLastBuildOk As Boolean` (`Core.bas:139`) and **above**
`Core_DeclarePackage` — module-level `Dim` in the declarations section is mandatory
(CLAUDE.md rule 5; `build.ps1` fails on a declaration placed after the first procedure):

```vba
' ---- memoized package registry (v0.4.0 performance) ----
' BuildRegistry re-runs CoreCandidates, one Application.Run "Pkg<Name>_Manifest"
' per candidate, and a _Packages worksheet scan per candidate. It is called from
' RenderOffBand (twice), HasAnyTier's per-package loops, ActionCaptionFor and
' ActionTokenFor ONCE PER CHAIN STEP, StepsInline once per step, and ActionIdSet
' TWICE inside a single ComposeReport - on top of the main render loops. At the
' 82-action fixture that is thousands of manifest rebuilds, and the owner
' measured the first Setup build in minutes.
'
' Handing the SAME Collection to every caller is safe because no consumer mutates
' a reg. Every mutation of a reg dictionary in this module is inside
' Core_DeclarePackage / Core_AddSection / Core_AddField / Core_AddList /
' Core_AddTable / Core_AddDefault / Core_AddAction, all of which run DURING
' TryManifest, inside BuildRegistry itself, against a dictionary it has just
' created. All thirteen "For Each reg In BuildRegistry()" sites only read.
'
' IF YOU ADD A CONSUMER THAT MUTATES A reg, this cache becomes a silent
' cross-contamination bug: the mutation persists into every later caller in the
' same build. Deep-copy in that consumer, or call InvalidateRegistry after it.
' Do not "just try it" - nothing in this repo's tooling can see the failure.
Private mRegCache   As Collection
Private mRegCacheOk As Boolean
```

- [ ] **Step 3: Rewrite `BuildRegistry` and add `InvalidateRegistry`**

`Core.bas:710-723`. Replace:

```vba
' The installed + enabled packages, each as a fully-built reg.
Private Function BuildRegistry() As Collection
    Dim names As Collection, nm As Variant, reg As Object, out As New Collection
    Set names = CoreCandidates()
    For Each nm In names
        Set reg = CreateObject("Scripting.Dictionary")
        If TryManifest(CStr(nm), reg) Then
            If reg.Exists("id") Then
                If PackageEnabled(CStr(reg("id"))) Then out.Add reg
            End If
        End If
    Next nm
    Set BuildRegistry = out
End Function
```

with:

```vba
' The installed + enabled packages, each as a fully-built reg. MEMOIZED - see
' mRegCache's declaration for why sharing one Collection is safe, and for the
' one thing that would make it unsafe.
Private Function BuildRegistry() As Collection
    Dim names As Collection, nm As Variant, reg As Object, out As Collection
    If mRegCacheOk Then
        Set BuildRegistry = mRegCache
        Exit Function
    End If
    ' Explicit New, not "Dim out As New Collection": an auto-instantiating local
    ' re-creates itself on the next reference after being set to Nothing, which
    ' is a surprise nobody needs sitting next to a cache.
    Set out = New Collection
    Set names = CoreCandidates()
    For Each nm In names
        Set reg = CreateObject("Scripting.Dictionary")
        If TryManifest(CStr(nm), reg) Then
            If reg.Exists("id") Then
                If PackageEnabled(CStr(reg("id"))) Then out.Add reg
            End If
        End If
    Next nm
    Set mRegCache = out
    mRegCacheOk = True
    Set BuildRegistry = out
End Function

' Drop the memoized registry. Called at the top of every entry point that BEGINS
' a build - Core_BuildSetup, Core_FinalizeSetup, Core_BuildControlPanel - and
' deliberately NOT from Core_Run.
'
' The registry depends on the installed modules (fixed for the session) and on
' the optional _Packages sheet's enable/disable values, which a user can only
' have edited BETWEEN builds. Those three entry points are every place a build
' begins, so an edited _Packages is always picked up. Invalidating in Core_Run
' would rebuild the registry once per chain STEP, which is most of what this
' cache exists to stop.
Private Sub InvalidateRegistry()
    Set mRegCache = Nothing
    mRegCacheOk = False
End Sub
```

- [ ] **Step 4: Invalidate at the three build entry points, and only those three**

`Core_BuildSetup` (`Core.bas:1128`) — insert as the first statement after
`On Error GoTo Fail`, above `If Not GateConfigMigration() Then Exit Sub`:

```vba
    InvalidateRegistry              ' a build begins here - see InvalidateRegistry
```

`Core_FinalizeSetup` (`Core.bas:2314`) — insert as the first statement after
`On Error GoTo Fail` (`Core.bas:2317`):

```vba
    InvalidateRegistry              ' a build begins here - see InvalidateRegistry
```

`Core_BuildControlPanel` (`Core.bas:2433`) — insert immediately after
`mLastBuildOk = False` and before `On Error GoTo Fail`:

```vba
    InvalidateRegistry              ' a build begins here - see InvalidateRegistry
```

**Do not add a fourth call site.** In particular, do not add one to `Core_Run` — that is
the whole point.

- [ ] **Step 5: Collapse `AvailDropdown` to one call per block**

First confirm it has exactly one caller:

```
grep -rn 'AvailDropdown' src/
```

Expected: the definition at `Core.bas:1780` and one call at `Core.bas:1832`.
If a second caller exists, **stop and report** — it would need the same treatment.

`Core.bas:1776-1788`. Replace:

```vba
' A validated dropdown removes free-typing from a three-value column. Wrapped in
' On Error Resume Next because validation can be refused (a protected sheet, some
' locales) and a refused dropdown must not abort the build - ParseAvail still
' guards the value, and Finalize still blocks on anything unrecognised.
Private Sub AvailDropdown(cell As Range)
    On Error Resume Next
    cell.Validation.Delete
    cell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, Formula1:="Button,Chains only,Off"
    cell.Validation.IgnoreBlank = True
    cell.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub
```

with:

```vba
' A validated dropdown removes free-typing from a three-value column. Wrapped in
' On Error Resume Next because validation can be refused (a protected sheet, some
' locales) and a refused dropdown must not abort the build - ParseAvail still
' guards the value, and Finalize still blocks on anything unrecognised.
'
' v0.4.0: takes the whole block of action rows at once, not one cell at a time.
' Per-cell it made TWO COM validation calls per row inside RenderPlacementBlock's
' loop - about 164 at the 82-action fixture scale. Nothing depends on the
' per-cell form: the value is guarded by ParseAvail and blocked at Finalize
' whether or not the dropdown was accepted.
Private Sub AvailDropdown(rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                       Operator:=xlBetween, Formula1:="Button,Chains only,Off"
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub
```

- [ ] **Step 6: Call it once, bounded to the rows actually written**

`RenderPlacementBlock` (`Core.bas:1795-1837`). Extend the `Dim` line:

```vba
    Dim prior As Object, priorAvail As Object, reg As Object, a As Object, id As String
    Dim firstDataRow As Long, lastDataRow As Long
```

After `r = r + 1` (`Core.bas:1818`, the line following the schema-marker write), insert:

```vba
    firstDataRow = r
```

Delete the in-loop call at `Core.bas:1832`:

```vba
            AvailDropdown ws.Cells(r, VAL_COL + 1)
```

After the `Next reg` that closes the loop (`Core.bas:1835`) and before
`RenderPlacementBlock = r`, insert:

```vba
    lastDataRow = r - 1
    ' ONE validation call over the exact block of action rows. Bounded to the
    ' rows actually written - never the whole column - so it cannot leak
    ' validation onto the key/marker row above it or onto anything below the
    ' END_KEY sentinel.
    If lastDataRow >= firstDataRow Then
        AvailDropdown ws.Range(ws.Cells(firstDataRow, VAL_COL + 1), _
                               ws.Cells(lastDataRow, VAL_COL + 1))
    End If
```

- [ ] **Step 7: Run both structural gates**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
```

- [ ] **Step 8: Commit**

```bash
git add src/core/Core.bas
git commit -m "perf(core): memoize BuildRegistry; one validation call per block

BuildRegistry re-ran CoreCandidates, an Application.Run per candidate and a
_Packages scan per candidate on EVERY call - and it is called once per chain
step from four places and twice inside one ComposeReport. The owner measured
the first Setup build in minutes at the 82-action fixture.

Cached behind a module-level Collection, invalidated at Core_BuildSetup,
Core_FinalizeSetup and Core_BuildControlPanel - the three places a build
begins, which is every place _Packages could have changed - and deliberately
NOT in Core_Run, so a 20-step chain builds the registry once.

Re-verified before caching: no consumer mutates a returned reg. Every mutation
is inside the manifest helpers (Core.bas:146-224), which run during TryManifest
inside BuildRegistry itself; all thirteen consumer sites only read.

AvailDropdown made two COM validation calls per action row inside
RenderPlacementBlock's loop (~164 at fixture scale); now one call over the
exact block of data rows. Behaviour-neutral: ParseAvail still guards the value
and Finalize still blocks regardless of whether the dropdown was accepted."
```

---

## Task 4: Preserve the index feature to the attic

Design item 4a. **Move, do not delete.** The owner reversed their own earlier call after
seeing the feature built; that call stands and is not re-argued. "Save the work" means
preserve the code.

Splitting the preservation (this task) from the unwiring (Task 5) means the attic file can
be reviewed against the original as a pure copy, with no unwiring noise in the diff.

**Files:**
- Create: `src/attic/HomeIndex.bas.txt`

**Interfaces:**
- Consumes: nothing.
- Produces: nothing consumed by later tasks. `src/core/Core.bas` is **not** modified in
  this task.

- [ ] **Step 1: Create the directory**

```bash
mkdir -p src/attic
```

- [ ] **Step 2: Copy the five procedures verbatim into the attic file**

Create `src/attic/HomeIndex.bas.txt` containing the header below, followed by
**`src/core/Core.bas` lines 3327-3537 copied byte-for-byte** (the comment block that
opens `' The generated index - ALWAYS built...` through `AddText`'s `End Function`),
and then **`src/core/Core.bas` lines 2673-2679** (`PageButtonCount`, whose only two
callers are inside `BuildHomePage` — verified by grep, so it is orphaned by the removal).

Do not retype the bodies. Copy them.

The header, placed at the top of the file:

```
'================================================================
' ATTIC - Macros Home, the generated index surface.  NOT BUILT.
'================================================================
' WHAT THIS WAS
'   The "Macros Home" tab: a generated index built LAST in
'   Core_BuildControlPanel's pipeline. Three parts - a page summary
'   (one row per generated page, with its button count and a link),
'   the full button catalog (every button, its page, and its
'   unabbreviated description), and a legend explaining the
'   placement code and the Availability states. It was also the only
'   thing in the whole pipeline that activated a sheet and selected
'   a cell, so it doubled as the build's landing place.
'
' IT NEVER REACHED A RELEASE.
'   Built during v0.4.0 track 1 (tasks 6-9) on branch
'   feat/finish-excel-track1. main is still v0.3.0. No shipped
'   deployment has ever had this tab, so nothing in the field
'   depends on it.
'
' WHY IT WAS WITHDRAWN - the owner, at the v0.4.0 manual gate:
'   "macros home tab is basically useless and shouldn't exist. its
'   needless clutter and since most folks will only ever load a
'   handful of macros or 1 or 2 chains into a workbook, there's
'   really no use for it other than as a demo. save the work for
'   later reference, but remove from the bundle."
'   This reversed their own earlier decision, made before they could
'   see it built. That is what a manual gate is for. Do not re-argue
'   it; if it comes back, it comes back because the owner asks.
'
' THE FILE EXTENSION IS LOAD-BEARING.
'   ".bas.txt", not ".bas": build/build.ps1 discovers packages with a
'   src/packages/Pkg*.bas glob, and runs its ASCII and reserved-word
'   gates over that same glob plus src/core/Core.bas. Nothing under
'   src/attic/ can reach a bundle or a gate, and nobody can mistake
'   this for a live module.
'
' WHAT A RESTORE WOULD HAVE TO RE-WIRE (all removed in v0.4.0):
'   1. The pipeline call. Core_BuildControlPanel ended with
'        BuildHomePage layout, manual
'      as its LAST step, after BuildAdminPanel. It is now
'        LandOnFirstPage layout
'      which activates the first generated page instead. Restoring
'      means putting BuildHomePage back LAST (it activates and
'      selects, so anything after it changes where the user lands)
'      and deciding what happens to LandOnFirstPage.
'   2. The nav link. RenderNavRow wrote an unconditional
'        NavLink ws.Cells(1, 2), "<< Home", HOME_SHEET
'      and put "< Previous" in cell D and "Next >" in cell F. Those
'      two now sit in B and D. Restoring means shifting them back.
'      Do NOT re-add a Core_SheetExists guard on the Home link - one
'      was tried and, because BuildHomePage runs last, it made every
'      link dead text on a first build. See the deleted comment.
'   3. IsPanelName's carve-out. "Macros Home" starts with PANEL_BASE
'      but has a non-numeric tail, so IsPanelName returns False and
'      DeletePanels spares it - which is what let a rebuild REUSE
'      rather than recreate the sheet. That rule is still in place
'      and unchanged. But DeleteHomePanel now deletes the tab
'      unconditionally on every build (it retires any index a
'      pre-removal build left behind, exactly as DeleteLegacyPanel
'      retires the pre-v0.4.0 single "Macros" tab). A restore must
'      remove DeleteHomePanel or the index is deleted every build.
'   4. Contract section 6. The "Generated index" clause used to say
'      the engine ALWAYS generates an index surface. It now says a
'      platform MAY provide one. The previous/next navigation
'      requirement was kept unconditional either way. The sheet name
'      "Macros Home" is still RESERVED in section 7 - deliberately,
'      so a restore can have the name back.
'   5. PageButtonCount. Included below because BuildHomePage was its
'      only caller; it was removed from Core.bas with the rest.
'
' The bodies below are byte-for-byte as they stood at the last commit
' before removal. They reference Core-private helpers that still
' exist (EnsurePanelSheet, NavLink, SortedKeys, HOME_SHEET,
' PANEL_BASE, ADMIN_SHEET, CARD_COLS), so a restore is a paste-back
' plus the five wiring points above.
'================================================================

```

- [ ] **Step 3: Verify the copy is faithful**

```bash
sed -n '3327,3537p' src/core/Core.bas > /tmp/orig1.txt
sed -n '2673,2679p' src/core/Core.bas > /tmp/orig2.txt
grep -c 'Private Sub BuildHomePage\|Private Function HomeCatalog\|Private Function HomeLegend\|Private Function HomeLine\|Private Function AddText\|Private Function PageButtonCount' src/attic/HomeIndex.bas.txt
```

Expected from the `grep -c`: **6**. Then diff the copied region against the originals to
confirm nothing was retyped or reflowed. Every one of the six procedures must be present
with its full body and its original comments.

- [ ] **Step 4: Confirm the attic is invisible to the build**

```
pwsh build/build.ps1
grep -c 'BuildHomePage' dist/MacrosBundle.bas
```

Expected: the build succeeds and the grep returns the same count as before this task
(the procedure is still in `Core.bas` — Task 5 removes it). Confirm `src/attic` does not
appear anywhere in the build's output listing.

- [ ] **Step 5: Commit**

```bash
git add src/attic/HomeIndex.bas.txt
git commit -m "chore(attic): preserve the Macros Home index before withdrawing it

The owner reversed their own earlier call after seeing the feature built:
'save the work for later reference, but remove from the bundle.' This is the
save. Task 5 does the removal.

BuildHomePage, HomeCatalog, HomeLegend, HomeLine, AddText and PageButtonCount
copied verbatim, headed with what it was, that it never reached a release, the
owner's reason, and the five wiring points a restore would have to redo.

.bas.txt so build.ps1's Pkg*.bas glob and its ASCII and reserved-word gates
cannot pick it up and nobody mistakes it for a live module."
```

---

## Task 5: Remove `Macros Home` from the bundle

Design items 4a (removal half), 4b, 4c. Four non-obvious consequences, all handled here.

**Files:**
- Modify: `src/core/Core.bas`
- Modify: `spec/contract.md` — §6, §7 (no change, verified), §10

**Interfaces:**
- Consumes: `src/attic/HomeIndex.bas.txt` from Task 4 (referenced by comments only).
- Produces: `Private Sub LandOnFirstPage(layout As Object)`,
  `Private Sub DeleteHomePanel()`. `HOME_SHEET` **survives** — `DeleteHomePanel` needs the
  name and §7 keeps it reserved.

- [ ] **Step 1: Delete the six procedures from Core**

Delete `src/core/Core.bas` lines **3327-3537** inclusive — from the comment block opening
`' The generated index - ALWAYS built, even for a single page...` through `AddText`'s
`End Function`. That is `BuildHomePage`, `HomeCatalog`, `HomeLegend`, `HomeLine` and
`AddText`.

Also delete `PageButtonCount` (`Core.bas:2673-2679`) and its preceding blank line — its
only two callers were inside `BuildHomePage`.

Confirm nothing else referenced them:

```
grep -n 'BuildHomePage\|HomeCatalog\|HomeLegend\|HomeLine\|AddText\|PageButtonCount' src/core/Core.bas
```

Expected after deletion: only **comment** mentions remain (at `Core.bas:96`, `:138`,
`:2298`, `:2302`, `:2457`, `:2627`, `:2790`, `:2829`, `:3143`, `:3559` by the pre-edit
numbering). Steps 2-8 below clear the ones that have become false. A comment mention that
is a historical note (e.g. `:2790`, `:3143` describing why a card's chain recipe is
bounded) must be **re-aimed, not left claiming Home still shows the full recipe** — see
Step 8.

- [ ] **Step 2: Add `LandOnFirstPage`**

`BuildHomePage` was the only thing in the pipeline that activated a sheet and selected a
cell. Insert this where `BuildHomePage` was (after the procedure that preceded it):

```vba
' Activate the lowest-numbered generated page and put the cursor on A1, so a
' build ends somewhere deliberate. This replaces the ONLY activation the pipeline
' ever had: BuildHomePage, withdrawn in v0.4.0 (src/attic/HomeIndex.bas.txt).
'
' Activate BEFORE selecting. Range.Select on a non-active sheet raises 1004,
' which is how RenderPageLegacy's bare Range("A1").Select survived for so long -
' it only ever ran on the fresh-sheet path, and Worksheets.Add activates. See the
' Task 8 ledger; do not let that habit migrate forward.
'
' Does nothing when there are NO generated pages. That is reachable now: every
' button can be hidden or switched Off (and, from v0.4.0, placed on a user's own
' sheet).
'
' Does nothing either when the target sheet is not xlSheetVisible. Hiding a
' generated tab is a supported habit here, and Activate raises 1004 on a hidden
' sheet. Deliberately does NOT unhide it the way BuildHomePage unhid the index:
' the index was the one surface a user had to be able to get back to; a numbered
' page is not, and silently unhiding a tab the user hid is worse than landing
' wherever they already were.
'
' The Core_SheetExists check here is NOT the banned nav-link guard (see
' RenderNavRow). That one fired at RENDER time against a sheet built later in the
' pipeline, and made every link dead text on a first build. This one fires at the
' END of the pipeline against a sheet RenderAllPages has already created - and it
' is here only so a stale layout entry cannot raise on the build's last statement
' and flip mLastBuildOk to False after everything actually succeeded.
Private Sub LandOnFirstPage(layout As Object)
    Dim tabKeys As Variant, ws As Worksheet, nm As String
    If layout.Count = 0 Then Exit Sub
    tabKeys = SortedKeys(layout)
    If UBound(tabKeys) < LBound(tabKeys) Then Exit Sub
    nm = PANEL_BASE & CLng(tabKeys(LBound(tabKeys)))
    If Not Core_SheetExists(nm) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(nm)
    If ws.Visible <> xlSheetVisible Then Exit Sub
    ws.Activate                     ' Range.Select needs its sheet active
    ws.Range("A1").Select
End Sub
```

- [ ] **Step 3: Add `DeleteHomePanel`**

Insert directly after `DeleteLegacyPanel`'s `End Sub` (`Core.bas:2486`), the procedure it
mirrors:

```vba
' Retire the withdrawn "Macros Home" index, mirroring DeleteLegacyPanel's
' retirement of the pre-v0.4.0 single "Macros" tab. Called UNCONDITIONALLY by
' Core_BuildControlPanel.
'
' Unconditional matters. DeletePanels deliberately spares "Macros Home"
' (IsPanelName requires a numeric tail), so an index left behind by a
' PRE-removal build would otherwise be never deleted and never rebuilt: it would
' survive every rebuild forever, showing a page list, a button catalog and a
' legend that describe a workbook that has moved on, with links to pages that may
' no longer exist. A stale index is worse than no index.
Private Sub DeleteHomePanel()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(HOME_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub
```

- [ ] **Step 4: Rewire the pipeline and settle the completion message**

`Core_BuildControlPanel` (`Core.bas:2447-2463`). Replace:

```vba
    DeletePanels                            ' remove any prior "Macros N" pages (spares
                                             ' "Macros Home" by design - see IsPanelName)
    DeleteLegacyPanel                       ' v0.4.0: the single "Macros" tab is retired in BOTH modes
```

with:

```vba
    DeletePanels                            ' remove any prior "Macros N" pages (spares
                                             ' "Macros Home" by design - see IsPanelName)
    DeleteLegacyPanel                       ' v0.4.0: the single "Macros" tab is retired in BOTH modes
    DeleteHomePanel                         ' v0.4.0: so is the withdrawn "Macros Home" index
```

Replace:

```vba
    BuildHomePage layout, manual            ' LAST: activates Macros Home and selects A1
```

with:

```vba
    LandOnFirstPage layout                  ' LAST: the build's deliberate landing place
```

Replace the success message:

```vba
    If Not silent Then _
        MsgBox "Built " & layout.Count & " macro page(s) with " & buttons.Count & _
               " button(s), plus the '" & HOME_SHEET & "' index.", _
               vbInformation, "Setup Complete"
```

with:

```vba
    If Not silent Then
        ' "Built 0 macro page(s)" is true but reads as a failure, and zero pages
        ' is a real, reachable state now that every button can be hidden or
        ' switched Off. Say what actually happened instead.
        If layout.Count = 0 Then
            MsgBox "No macro pages were built - every button is either hidden or switched Off.", _
                   vbInformation, "Setup Complete"
        Else
            MsgBox "Built " & layout.Count & " macro page(s) with " & buttons.Count & _
                   " button(s).", vbInformation, "Setup Complete"
        End If
    End If
```

**Note for Task 8:** this sentence is now settled. Task 8 edits it exactly once, to add
the placed-button count. That single-edit property is the whole reason item 4 was ordered
before item 3.

Also update the procedure's own header comment (`Core.bas:2423-2432`) — its last sentence
currently reads "Home is always built, even for a single page, so the workbook has one
shape at every scale." Replace that sentence with:

```vba
' The generated index this pipeline used to end with was withdrawn in v0.4.0
' (see src/attic/HomeIndex.bas.txt); DeleteHomePanel retires any copy a
' pre-removal build left behind, and LandOnFirstPage takes over its role as the
' build's landing place.
```

- [ ] **Step 5: Drop the dead `<< Home` nav link**

`RenderNavRow` (`Core.bas:2873-2878`). Replace the body:

```vba
Private Sub RenderNavRow(ws As Worksheet, ByVal prevSheet As String, ByVal nextSheet As String)
    NavLink ws.Cells(1, 2), "<< Home", HOME_SHEET
    If Len(prevSheet) > 0 Then NavLink ws.Cells(1, 4), "< Previous", prevSheet
    If Len(nextSheet) > 0 Then NavLink ws.Cells(1, 6), "Next >", nextSheet
    ws.Rows(1).RowHeight = 18
End Sub
```

with:

```vba
Private Sub RenderNavRow(ws As Worksheet, ByVal prevSheet As String, ByVal nextSheet As String)
    If Len(prevSheet) > 0 Then NavLink ws.Cells(1, 2), "< Previous", prevSheet
    If Len(nextSheet) > 0 Then NavLink ws.Cells(1, 4), "Next >", nextSheet
    ws.Rows(1).RowHeight = 18
End Sub
```

`NavLink` itself is unchanged.

**Replace the entire 13-line comment block above it (`Core.bas:2860-2872`).** It is
exclusively about why the `<< Home` link must not carry a `Core_SheetExists` guard, and
with the link gone it instructs the next reader to preserve something that no longer
exists. Delete it and put this in its place:

```vba
' Previous / Next only. The "<< Home" link that used to lead this row was dropped
' in v0.4.0 with the index it pointed at (src/attic/HomeIndex.bas.txt); Previous
' and Next shifted left into cells B and D.
'
' Worth keeping from the deleted comment, because it will come back the moment
' anyone re-adds a link to a sheet built later in the pipeline: do NOT guard a
' NavLink target with Core_SheetExists. One was added once, it fired at RENDER
' time against a sheet built last, and on a virgin workbook's first build it drew
' plain bold text instead of a link on every single page - then self-healed on
' the second build, so it was invisible to anyone testing a rebuild. Hyperlinks.Add
' does not validate SubAddress, and the target always exists by the time the macro
' returns and the user can click anything. Write the link unconditionally.
```

- [ ] **Step 6: Rewrite `IsPanelName`'s LOAD-BEARING comment**

`Core.bas:3552-3559`. The carve-out it defends has become vestigial, and a stale
load-bearing comment is worse than none. Replace the comment block (the function body is
**unchanged**) with:

```vba
' "Macros Home" starts with PANEL_BASE ("Macros ") but has a non-numeric tail, so
' it reads False here and DeletePanels never touches it.
'
' That WAS load-bearing: it is what let a rebuild reuse rather than recreate the
' index sheet, so EnsurePanelSheet found it already there and cleared and
' rewrote it in place. The index was withdrawn in v0.4.0
' (src/attic/HomeIndex.bas.txt) and DeleteHomePanel now removes any copy a
' pre-removal build left behind, so nothing depends on the carve-out any more.
'
' The numeric-tail rule is kept UNCHANGED all the same, for two reasons that
' outlive the index: the name stays RESERVED (contract section 7) against a
' future restore, and "delete exactly the pages this engine generates, and
' nothing else a user happened to name 'Macros Something'" is the correct rule
' on its own merits.
```

- [ ] **Step 7: Clear the remaining false claims about the index**

`Core.bas:2407-2409`, `Core_FinalizeSetup`'s success message. Replace:

```vba
        MsgBox "Configuration activated." & vbCrLf & vbCrLf & _
               "The Macros pages and 'Macros Home' index have been rebuilt to match your Setup answers.", _
               vbInformation, "Finalize Setup Tab"
```

with:

```vba
        MsgBox "Configuration activated." & vbCrLf & vbCrLf & _
               "The Macros pages have been rebuilt to match your Setup answers.", _
               vbInformation, "Finalize Setup Tab"
```

`Core.bas:2596`, `BuildAdminPanel`'s description cell. Replace:

```vba
        .Merge: .Value = "Validates your Setup answers, activates them, and rebuilds the Macros pages and 'Macros Home' index."
```

with:

```vba
        .Merge: .Value = "Validates your Setup answers, activates them, and rebuilds the Macros pages."
```

`Core.bas:138`, `mLastBuildOk`'s comment says "True once the whole pipeline (through
BuildHomePage) has actually finished." Change `BuildHomePage` to `LandOnFirstPage`.

`Core.bas:2302`, `AnyPlacementCode`'s comment ends "...and to set BuildHomePage's
isManual." Delete that clause; the sentence becomes "...both to choose BuildAutoLayout vs
BuildManualLayout."

`Core.bas:2298`, in the same block: "...and HomeLegend would tell them their page codes
control a layout they never authored" — replace `HomeLegend` with `the Setup tab's own
help text`, which is now the only place that claim is made.

- [ ] **Step 8: Re-aim the two chain-card comments that point at Home**

`Core.bas:2790` and `Core.bas:3143` both explain that a chain's card is deliberately
bounded *because* the full recipe is available on `Macros Home`. With Home gone, the
bounding is still correct but the stated reason is false — this is precisely the
string-drift class the ledger records four review rounds on.

At both sites, replace the "the full chain recipe stays in HomeCatalog" / "HomeCatalog
(Macros Home) is untouched and always shows the full recipe" clauses with:

```vba
' the full recipe stays authored on the Setup tab's Chains block
```

Adjust the surrounding sentence to read naturally at each site. Do not change any code.
Also check `Core.bas:96` and `:2627` and `:2829` and apply the same treatment — they
mention `HomeCatalog` as an existing surface.

- [ ] **Step 9: Amend contract §6 — the index becomes OPTIONAL**

`spec/contract.md:399-409`. Replace the **Generated index (v1.2.0)** paragraph:

```
**Generated index (v1.2.0):** the engine **always** generates an index surface,
even for a single page, so the workbook has one shape at every scale. It lists one
entry per built page with its title, its button count and a link; a **button
catalog** (caption, description, and the page each button is on) covering every
action whose `availability` placed it as a button — an action that is `off` or
`chains-only` has no button and no page, so it does not appear in this catalog;
it remains listed on the authoring surface (§6, per-action availability), which
is where it stays discoverable; a link to the admin surface; and a legend
explaining the placement code and availability. Every built page carries a
persistent link back to the index and links to its previous/next page. Reference
implementation: a `Macros Home` sheet; that name is **reserved** (§7).
```

with:

```
**Index surface (v1.2.0, amended v1.2.0 pre-freeze):** a platform **MAY** generate
an index surface. It is **not required**: the Excel reference implementation built one
and then withdrew it before release, because at the scale a real deployment runs —
a handful of actions and one or two chains — an index is clutter rather than
navigation. **When a platform does provide one**, it lists one entry per built page
with its title, its button count and a link; a **button catalog** (caption,
description, and the page each button is on) covering every action whose
`availability` placed it as a button — an action that is `off` or `chains-only` has
no button and no page, so it does not appear in this catalog; it remains listed on
the authoring surface (§6, per-action availability), which is where it stays
discoverable; a link to the admin surface; and a legend explaining the placement
code and availability.

**Page-to-page navigation is required regardless, and is unconditional:** every
built page MUST carry links to its previous and next page. Only the index itself,
and the per-page link back to it, are optional.

The reference implementation's withdrawn index was a `Macros Home` sheet. That name
stays **reserved** (§7) against a future restore.
```

- [ ] **Step 10: Verify contract §7 needs no change**

`spec/contract.md:426-428` already reserves `Macros Home`. It is correct as written and
**must not be relaxed** — reserving an unused name costs nothing and a future restore
wants it. Confirm by reading, change nothing, and note in the commit that it was checked.

- [ ] **Step 11: Amend contract §10's v1.2.0 bullet**

`spec/contract.md:563-582`. In the v1.2.0 bullet, replace `the always-generated index
surface with its navigation and button catalog,` with:

```
the optional index surface with its button catalog (built, then **withdrawn before
release** at the owner's call at the v0.4.0 manual gate — "needless clutter... most
folks will only ever load a handful of macros or 1 or 2 chains into a workbook";
the code is preserved unbuilt at `src/attic/HomeIndex.bas.txt` and the `Macros Home`
name stays reserved in §7), the unconditional previous/next page navigation,
```

This is engine presentation under §10's two-axis rule — no stored `key`, `id`, `kind`,
encoding or action semantics changes — so it stays additive. **No version bump**;
v1.2.0 is unreleased.

- [ ] **Step 12: Run both structural gates and confirm the removal**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
grep -c 'BuildHomePage\|HomeCatalog\|HomeLegend\|Sub HomeLine\|Function AddText\|PageButtonCount' dist/MacrosBundle.bas
```

Expected: all four gates pass and the final `grep -c` returns **0**. A non-zero count
means a procedure was missed; a build failure with an unbalanced `Sub`/`Function` means
the 3327-3537 cut landed on the wrong boundary.

- [ ] **Step 13: Commit**

```bash
git add src/core/Core.bas spec/contract.md
git commit -m "feat(core)!: withdraw the Macros Home index from the bundle

The owner's call at the v0.4.0 manual gate, reversing their own earlier one
after seeing it built. The code is preserved unbuilt at
src/attic/HomeIndex.bas.txt (previous commit).

Four non-obvious consequences, all handled:
- BuildHomePage was the ONLY thing in the pipeline that activated a sheet and
  selected a cell. New LandOnFirstPage takes over: activates the lowest-numbered
  generated page and selects A1, activating BEFORE selecting (Range.Select on a
  non-active sheet raises 1004), and doing nothing when there are no generated
  pages or the target is hidden.
- RenderNavRow's unconditional '<< Home' link is dropped; Previous/Next shift
  left into cells B and D.
- DeletePanels deliberately spares 'Macros Home', so an index left behind by a
  pre-removal build would be never deleted and never rebuilt. New
  DeleteHomePanel retires it unconditionally, mirroring DeleteLegacyPanel.
- Contract 6 mandated an index; left alone the Apps Script port would be built
  to a spec the reference implementation violates. The index is now optional,
  previous/next navigation stays required unconditionally, and 'Macros Home'
  stays reserved in 7 against a restore.

IsPanelName's rule is unchanged but its LOAD-BEARING comment is rewritten - it
was instructing the next reader to preserve a carve-out that has become
vestigial. No contract version bump; v1.2.0 is unreleased and this is engine
presentation under 10's two-axis rule."
```

---

## Task 6: Relabel the placement column to `Placement`

Design item 3a's string half, landed on its own so the ripple can be reviewed without
grammar changes mixed in. **No behaviour changes in this task.**

**Two facts, verified by grep, not assumed:**

1. **The relabel cannot disarm the migration tamper test.** `AVAIL_HEADER`
   (`Core.bas:60`) is compared at `Core.bas:1954`. `"Page code"` is only ever *written* —
   `Core.bas:1656` and `Core.bas:1810` — and never read back.
2. **...but it is written from two independent literals**, which is the exact coupling the
   ledger closed for `AVAIL_HEADER` in Task 4. Since this task edits both anyway, the new
   caption is introduced as a `Const` and both sites use it.

**A half-relabelled surface is worse than no relabel.** The ledger records four separate
review rounds spent on user-facing strings that had drifted from the behaviour they
described. Every site in Step 3 is updated in this task.

**Files:**
- Modify: `src/core/Core.bas`
- Modify: `spec/contract.md` — §6 prose

**Interfaces:**
- Consumes: Task 5's removal of `HomeLegend` (which held six of the "Page code" strings —
  this is why item 4 was ordered first).
- Produces: `Private Const PLACE_HEADER As String = "Placement"`.

- [ ] **Step 1: Add the Const to the declarations section**

Insert immediately after `AVAIL_HEADER` (`Core.bas:60`), inside the declarations section:

```vba
' Column header caption for the action catalog's and the Chains block's placement
' column. Introduced as a Const in v0.4.0 because the caption was written from
' two independent literals - the exact coupling closed for AVAIL_HEADER above.
'
' Unlike AVAIL_HEADER this one is NOT load-bearing, and that was checked rather
' than assumed: no check anywhere reads this caption back. "Page code" appeared
' only as a written value, at two sites. Kept as a Const anyway, because two
' literals for one visible string is how they drift.
'
' Relabelled from "Page code" because a placement value is no longer only a page
' code: it may also be a "Sheet!Cell" anchor onto a user's own tab.
Private Const PLACE_HEADER   As String = "Placement"
```

- [ ] **Step 2: Point both write sites at it**

`Core.bas:1656` (Chains block) and `Core.bas:1810` (action catalog) — both currently:

```vba
    ws.Cells(r, VAL_COL + 2).Value = "Page code"
```

Replace both with:

```vba
    ws.Cells(r, VAL_COL + 2).Value = PLACE_HEADER
```

- [ ] **Step 3: Update every user-facing string that says "page code"**

Each of these is a live string a user reads. Reword to "placement" / "placement code",
keeping each sentence's meaning intact.

| Site | Current text | Change |
|---|---|---|
| `Core.bas:1107` | `"the page code(s) named above there, then click 'Finalize Setup Tab'. "` | `"the placement code(s) named above there, then click 'Finalize Setup Tab'. "` |
| `Core.bas:1117-1118` | `"the page code(s) named above (those actions return to automatic placement); every other saved page code is kept."` | `"the placement code(s) named above (those actions return to automatic placement); every other saved placement is kept."` |
| `Core.bas:1123` | `"every saved page-code and Availability choice in the Action catalog..."` | `"every saved placement and Availability choice in the Action catalog..."` |
| `Core.bas:1510` | `"onto pages (Page code). Leave every code blank for the automatic, "` | `"onto pages (Placement). Leave every code blank for the automatic, "` |
| `Core.bas:1652` | `"Button name + Steps (paste tokens from the Actions list, one per line) + Page code."` | `"Button name + Steps (paste tokens from the Actions list, one per line) + Placement."` |
| `Core.bas:1804-1806` | `"Page code: leave every code blank for the automatic layout; ..."` | `"Placement: leave every code blank for the automatic layout; ..."` |
| `Core.bas:2031` | `"These saved page codes could not be read from the older configuration format:"` | `"These saved placements could not be read from the older configuration format:"` |
| `Core.bas:1791` | comment: `"a yellow Page-code cell in E"` | `"a yellow Placement cell in E"` |

Then run a sweep to catch anything this table missed:

```
grep -ni 'page code\|page-code' src/core/Core.bas
```

Every remaining hit must be either (a) a historical comment describing what a **past**
version did, which is fine and should stay, or (b) a live user-facing string, which must
be updated. There should be **no** live user-facing hits left. `ParseCode`'s own name and
its comments stay — it still parses page codes specifically.

- [ ] **Step 4: Update contract §6 prose**

`spec/contract.md:275-290`, the **Grouping / layout (v1.1.0)** paragraph. Replace the
phrase `each button (action or chain) may carry a **placement code**` — already correct —
but change the surrounding sentence that names the authoring column. Search §6 for the
strings `page code` and `Page code` and reword each to `placement` / `placement code`,
matching the terms now used in the engine. Where §6 refers to the *column* on the
authoring surface, name it `Placement`.

Add one sentence to that paragraph, immediately after the placement-code grammar:

```
The reference implementation heads this column **Placement** on its authoring surface
(renamed from `Page code` at v1.2.0, when the value stopped being only a page code).
The column caption is presentation, not data: no conforming reader may key off it.
```

- [ ] **Step 5: Run both structural gates**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
grep -c '"Page code"' dist/MacrosBundle.bas
```

Expected: gates pass, final `grep -c` returns **0**.

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas spec/contract.md
git commit -m "refactor(core): relabel the placement column 'Placement', via one Const

A placement value is about to stop being only a page code (Sheet!Cell anchors
land next), so 'Page code' becomes wrong on the Setup tab. The caption was
written from TWO independent literals - the exact coupling closed for
AVAIL_HEADER in Task 4 - so it becomes a PLACE_HEADER Const and both sites use
it.

Verified by grep, not assumed: no check anywhere reads this caption back
('Page code' appeared only as a written value at Core.bas:1656 and :1810), so
unlike AVAIL_HEADER the relabel cannot silently disarm the migration tamper
test.

Every user-facing string that said 'page code' is updated in the same change -
the two block headers, both blocks' help text, RenderComposeSection's intro,
GateConfigMigration's three recovery messages, ConfigSchemaStatus's detail, and
contract 6 prose. A half-relabelled surface is worse than no relabel."
```

---

## Task 7: The `Sheet!Cell` anchor grammar

Design items 3a (grammar), 3b (parsing), 3c (mode axis). Anchors are **parsed, classified
and validated for form**, and they do **not** flip the workbook into manual mode. Nothing
draws them yet — that is Task 8.

**Known intermediate state, by design, not a defect.** After this task, an anchor code on
a workbook that *also* has a real page code somewhere will route through
`BuildManualLayout`, fail `ParseCode`, and degrade to page 1's Unassigned column. Task 8's
`SplitAnchored` removes anchored buttons from the page pipeline entirely. A reviewer
should expect this and not "fix" it here — the design is explicit that
`BuildManualLayout` gets **no** defensive anchor branch, because the ledger records that a
per-filer duplicate of a funnel-level filter is exactly what produced the `hide` regression
in Task 7 of the prior plan.

**Files:**
- Modify: `src/core/Core.bas`

**Interfaces:**
- Consumes: `PLACE_HEADER` from Task 6.
- Produces: `CodeKind`, `ParseAnchor`, `IsA1Ref`.

- [ ] **Step 1: Add `CodeKind`, `ParseAnchor` and `IsA1Ref`**

Insert directly after `ParseCode`'s `End Function` (currently around `Core.bas:3680`), so
the placement parsers sit together:

```vba
' "" | "hide" | "page" | "anchor" - the SINGLE place a placement value's form is
' decided. Every reader goes through this rather than testing for "!" itself.
'
' The "!" discriminator is safe for two reasons, and NEITHER is "Excel forbids !
' in a sheet name" - it does not. The forbidden set is : \ / ? * [ ], so a sheet
' really can be called "Sales!Q1".
'   1. A valid PAGE CODE can never contain "!": ParseCode's grammar is digits
'      then letters with a trailing-junk guard, so "1A!" is already rejected.
'      That is what makes the two forms unconfusable - even for a sheet named "1A".
'   2. ParseAnchor splits on the LAST "!" (InStrRev), so "Sales!Q1!B4" and
'      "'Sales!Q1'!B4" both resolve to sheet "Sales!Q1", cell "B4".
' Do NOT "simplify" that InStrRev to InStr - it would silently break every anchor
' onto a sheet whose name legitimately contains "!".
Private Function CodeKind(ByVal code As String) As String
    Dim s As String
    s = Trim(code)
    If Len(s) = 0 Then CodeKind = "": Exit Function
    If LCase(s) = "hide" Then CodeKind = "hide": Exit Function
    If InStr(1, s, "!") > 0 Then CodeKind = "anchor": Exit Function
    CodeKind = "page"
End Function

' "TabName!B4" / "'TabName Data'!$B$4" / "TabName Data!B4" -> sheetName "TabName Data",
' cellRef "B4". Returns False on anything malformed.
'
' The sheet part may be quoted; an embedded apostrophe is doubled inside the
' quotes, matching Excel's own spelling. The cell part is a SINGLE A1 reference.
' A colon (a multi-cell range) is rejected: this anchors one button, not a block.
'
' Deliberately a SIBLING of ParseCode, not a branch inside it. ParseCode's four
' call sites keep their page-code meaning and its hard-won hardening - the digit
' cap, the 3-letter cap, the VSPLIT_MULT guard - stays exactly as the ledger left
' it.
Private Function ParseAnchor(ByVal code As String, ByRef sheetName As String, _
                             ByRef cellRef As String) As Boolean
    Dim s As String, p As Long, sh As String, cl As String
    sheetName = ""
    cellRef = ""
    s = Trim(code)
    p = InStrRev(s, "!")
    If p < 2 Then Exit Function              ' no "!", or nothing before it
    If p = Len(s) Then Exit Function         ' nothing after it
    sh = Left$(s, p - 1)
    cl = Mid$(s, p + 1)
    If Left$(sh, 1) = "'" Then
        If Len(sh) < 3 Then Exit Function
        If Right$(sh, 1) <> "'" Then Exit Function
        sh = Mid$(sh, 2, Len(sh) - 2)
        sh = Replace(sh, "''", "'")
    End If
    sh = Trim(sh)
    If Len(sh) = 0 Then Exit Function
    If Not IsA1Ref(cl) Then Exit Function
    sheetName = sh
    cellRef = UCase(Replace(cl, "$", ""))
    ParseAnchor = True
End Function

' A single A1-style cell reference: optional "$", 1-3 column letters, optional
' "$", 1-7 row digits, and nothing else. Rejects a colon (a range), a
' whole-column reference, and any trailing character.
' Row bounds beyond the digit cap are NOT checked here - a 7-digit row that
' exceeds the sheet's real last row simply fails to resolve when the caller asks
' for the Range, which is the one place that knows the platform's limit.
Private Function IsA1Ref(ByVal s As String) As Boolean
    Dim t As String, i As Long, ch As String, letters As String, digits As String
    t = UCase(Trim(s))
    If Len(t) = 0 Then Exit Function
    i = 1
    If Mid$(t, i, 1) = "$" Then i = i + 1
    Do While i <= Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "A" And ch <= "Z" Then letters = letters & ch Else Exit Do
        i = i + 1
    Loop
    If i <= Len(t) Then
        If Mid$(t, i, 1) = "$" Then i = i + 1
    End If
    Do While i <= Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch Else Exit Do
        i = i + 1
    Loop
    If i <= Len(t) Then Exit Function                     ' trailing junk, or a ":"
    If Len(letters) = 0 Or Len(letters) > 3 Then Exit Function
    If Len(digits) = 0 Or Len(digits) > 7 Then Exit Function
    If CLng(digits) < 1 Then Exit Function
    IsA1Ref = True
End Function
```

- [ ] **Step 2: Anchors are their own axis — they must NOT flip manual mode**

`AnyPlacementCode` (`Core.bas:2293-2309`). Replace the body:

```vba
Private Function AnyPlacementCode(map As Object) As Boolean
    Dim k As Variant, v As String
    For Each k In map.Keys
        v = Trim(CStr(map(k)))
        If Len(v) > 0 And LCase(v) <> "hide" Then AnyPlacementCode = True: Exit Function
    Next k
End Function
```

with:

```vba
Private Function AnyPlacementCode(map As Object) As Boolean
    Dim k As Variant
    For Each k In map.Keys
        If CodeKind(CStr(map(k))) = "page" Then AnyPlacementCode = True: Exit Function
    Next k
End Function
```

Append to that function's existing comment block (`Core.bas:2293-2302`), which already
explains the `hide` carve-out:

```vba
' v0.4.0: a "Sheet!Cell" ANCHOR gets the same carve-out, and for the same reason.
' An anchor does not place a button on a generated page at all - it puts one on a
' sheet the engine otherwise never touches - so it is not a page-layout decision.
' Without this, typing "TabName!B4" on ONE action would flip the workbook into manual
' mode and dump the other 81 buttons into page 1's Unassigned column. This is the
' third carve-out of the same shape; the contract clause is written to match, in
' both directions, so a port reading section 6 literally cannot reintroduce the bug -
' which the ledger records almost happening with "hide".
```

- [ ] **Step 3: Give `ConfigSchemaStatus` an anchor branch**

`Core.bas:1995-2038`. This one is **not cosmetic**: the `invalid` path *clears* the
offending rows, so an anchor classified `invalid` would be silently deleted on a marker
recovery.

Extend the `Dim` line at `Core.bas:1997` to add two locals:

```vba
    Dim id As String, code As String, tn As Long, cx As Long, cl As String
    Dim aSheet As String, aCell As String
```

Replace the validation body inside `Case "absent"` (`Core.bas:2021-2027`):

```vba
                    code = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
                    If Len(code) > 0 And LCase(code) <> "hide" Then
                        If Not ParseCode(code, tn, cx, cl) Then
                            detail = detail & "  - " & id & ": '" & code & "'" & vbCrLf
                            badIds(LCase(id)) = True
                        End If
                    End If
```

with:

```vba
                    code = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
                    ' Route through CodeKind so an anchor is never classified
                    ' "invalid". That path CLEARS the offending rows, so getting
                    ' this wrong would silently delete every anchor on a marker
                    ' recovery. A legacy store predates anchors, so this branch
                    ' is defensive rather than reachable today - which is exactly
                    ' why it has to be right, because nothing will exercise it.
                    Select Case CodeKind(code)
                        Case "page"
                            If Not ParseCode(code, tn, cx, cl) Then
                                detail = detail & "  - " & id & ": '" & code & "'" & vbCrLf
                                badIds(LCase(id)) = True
                            End If
                        Case "anchor"
                            If Not ParseAnchor(code, aSheet, aCell) Then
                                detail = detail & "  - " & id & ": '" & code & "'" & vbCrLf
                                badIds(LCase(id)) = True
                            End If
                    End Select
```

- [ ] **Step 4: Give `ComposeReport`'s two code loops an anchor branch**

`ComposeReport` (`Core.bas:2101`). Extend the `Dim` block to add:

```vba
    Dim aSheet As String, aCell As String
```

**Chain codes** (`Core.bas:2128-2132`). Replace:

```vba
        code = CStr(ch("code"))
        If Len(code) > 0 And LCase(code) <> "hide" Then
            If Not ParseCode(code, tabNum, colIx, colL) Then _
                s = s & "  {bad code '" & code & "'}": structural = structural + 1
        End If
```

with:

```vba
        code = CStr(ch("code"))
        Select Case CodeKind(code)
            Case "page"
                If Not ParseCode(code, tabNum, colIx, colL) Then _
                    s = s & "  {bad code '" & code & "'}": structural = structural + 1
            Case "anchor"
                If Not ParseAnchor(code, aSheet, aCell) Then _
                    s = s & "  {bad code '" & code & "'}": structural = structural + 1
        End Select
```

A chain can be anchored to a user's sheet exactly like an action — both blocks have always
used the same grammar, and `PlaceAnchoredButtons` will receive both from the same
collection.

**Placement codes** (`Core.bas:2138-2151`). Replace:

```vba
    For Each k In map.Keys
        code = CStr(map(k))
        If LCase(code) = "hide" Then
            hidden = hidden & "  - " & HiddenTokenFor(CStr(k), chains) & vbCrLf
        ElseIf Len(code) = 0 Then
            ' Fix round 2, fix 2: was the bare lowercased key - same
            ' inconsistency fix round 1's fix 5 removed from the HIDDEN list
            ' just above, now closed here too so both lists in one report
            ' agree on "caption [id]".
            unassigned = unassigned & "  - " & HiddenTokenFor(CStr(k), chains) & vbCrLf
        ElseIf Not ParseCode(code, tabNum, colIx, colL) Then
            s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf: structural = structural + 1
        End If
    Next k
```

with:

```vba
    For Each k In map.Keys
        code = CStr(map(k))
        Select Case CodeKind(code)
            Case "hide"
                hidden = hidden & "  - " & HiddenTokenFor(CStr(k), chains) & vbCrLf
            Case ""
                ' Fix round 2, fix 2: was the bare lowercased key - same
                ' inconsistency fix round 1's fix 5 removed from the HIDDEN list
                ' just above, now closed here too so both lists in one report
                ' agree on "caption [id]".
                unassigned = unassigned & "  - " & HiddenTokenFor(CStr(k), chains) & vbCrLf
            Case "page"
                If Not ParseCode(code, tabNum, colIx, colL) Then _
                    s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf: structural = structural + 1
            Case "anchor"
                ' Form only here. The sheet-exists, reserved-sheet and
                ' duplicate-target checks are added in the next task, and they
                ' all run AFTER this one has confirmed the anchor parses.
                If Not ParseAnchor(code, aSheet, aCell) Then _
                    s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf: structural = structural + 1
        End Select
    Next k
```

- [ ] **Step 5: Do NOT touch `BuildManualLayout`**

`ParseCode`'s fourth call site (`Core.bas:2923`) gets **no** anchor branch. Anchors are
partitioned out one layer above in the next task, and the ledger records that a per-filer
duplicate of a funnel-level filter is exactly what produced the `hide` regression. Read
`BuildManualLayout`'s existing comment (`Core.bas:2909-2913`) — it says the same thing
about `hide`, and says the dead branch was removed rather than left, "so nobody reads it
as evidence 'hide' is still handled per-filer and re-adds a per-filer check that can never
fire."

Confirm you have not touched it:

```
git diff src/core/Core.bas | grep -A5 -B5 'BuildManualLayout'
```

Expected: no output.

- [ ] **Step 6: Update the placement block's help text with the new form**

`Core.bas:1801-1806`, `RenderPlacementBlock`'s `BlockIntro`. After the existing
`"Placement: ... number = page, letters = column (e.g. 1A). "` sentence and before the
`"Type  hide  to leave any one button off..."` sentence, insert:

```vba
                   "Or put a button on one of your OWN tabs with  Sheet!Cell  - " & _
                   "for example  TabName!B4  places it on cell B4 of the TabName tab, and " & _
                   "does not change where any other button goes. " & _
```

Do the same for the Chains block's `BlockIntro` (`Core.bas:1651-1652`), whose one-line
summary becomes:

```vba
                   "Button name + Steps (paste tokens from the Actions list, one per line) + Placement (1A, hide, or TabName!B4).")
```

- [ ] **Step 7: Run both structural gates**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
```

- [ ] **Step 8: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): parse Sheet!Cell anchors as a second placement form

Grammar, classification and form-validation only - nothing draws an anchored
button yet.

New CodeKind is the single place a placement value's form is decided
('' | hide | page | anchor), so no site re-derives the rule. New ParseAnchor is
a SIBLING of ParseCode, not a branch inside it: ParseCode's four call sites keep
their page-code meaning and its hardening (digit cap, 3-letter cap, VSPLIT_MULT
guard) is untouched. The '!' discriminator is safe because a valid PAGE CODE can
never contain '!' (ParseCode's grammar rejects it), and because ParseAnchor
splits on the LAST '!' - so an anchor can never be confused with a page code,
not even for a sheet named '1A', and a sheet whose name legitimately contains
'!' still resolves.

Anchors are their own axis: AnyPlacementCode now ignores them exactly as it
ignores 'hide', so typing TabName!B4 on one action does NOT flip the workbook into
manual mode and dump the other 81 buttons into Unassigned. Third carve-out of
the same shape.

Three of ParseCode's four call sites gain an anchor branch. ConfigSchemaStatus's
is the one that matters: its 'invalid' path CLEARS the offending rows, so a
missing branch would silently delete every anchor on a marker recovery.
BuildManualLayout deliberately gets NO branch - anchors are partitioned out one
layer above, and a per-filer duplicate of a funnel-level filter is what produced
the 'hide' regression."
```

---

## Task 8: Place anchored buttons, and sweep them on rebuild

Design items 3d, 3e, 3f. A **genuinely separate code path** from the card grid.

**The card-grid renderer owns its sheet outright.** `EnsurePanelSheet` (`Core.bas:2602`)
does `Cells.Clear`, `Cells.UnMerge`, `Hyperlinks.Delete`, `Cells.RowHeight`, six column
widths, tab colour and interior fill. **A user's own sheet must never be treated that
way**, which is why this is a different procedure and not a parameter on that one.

**Files:**
- Modify: `src/core/Core.bas`

**Interfaces:**
- Consumes: `CodeKind`, `ParseAnchor` (Task 7); `LandOnFirstPage`, `DeleteHomePanel`,
  the settled completion message (Task 5).
- Produces: `PLACED_PREFIX`, `ANCHOR_BTN_W`, `SafeShapeName`, `SplitAnchored`,
  `SweepPlacedButtons`, `AnchorSheet`, `PlaceAnchoredButtons`; `AddBtn` gains an optional
  trailing parameter.

- [ ] **Step 1: Add the two constants to the declarations section**

Insert after `BTN_ROW_H` (`Core.bas:88`), inside the declarations section:

```vba
' ---- Sheet!Cell anchored buttons (v0.4.0) ----
' Shape-name prefix for every button this engine places on a user's OWN sheet.
' RESERVED (contract section 7): nothing else in the workbook may use it.
' SweepPlacedButtons finds and removes buttons by this prefix alone - it is the
' only handle the engine has on a sheet it must otherwise never touch.
Private Const PLACED_PREFIX As String = "SFPlaced_"
' Width of an anchored button, in points. Deliberately NOT the target cell's own
' width: a user's column may be three characters wide. Height reuses BTN_H.
' 110pt is the owner's call at the v0.4.0 gate (the design proposed 150pt), sized
' to stay close to one default-width column so a button is not intrusive sitting
' next to real data.
Private Const ANCHOR_BTN_W  As Double = 110
```

- [ ] **Step 2: Give `AddBtn` an optional placement parameter**

`Core.bas:3604-3619`. Replace:

```vba
Private Sub AddBtn(ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal action As String, _
                   ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double)
```

with:

```vba
Private Sub AddBtn(ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal action As String, _
                   ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, _
                   Optional ByVal place As XlPlacement = xlMoveAndSize)
```

and replace the final line `btn.Placement = xlMoveAndSize` with `btn.Placement = place`.

Extend that line's existing comment with:

```vba
    ' v0.4.0: xlMoveAndSize stays the DEFAULT, so every existing call site is
    ' unchanged. An anchored button on a user's own sheet passes xlMove instead:
    ' on a generated page the engine controls the column width and wants the
    ' button to track it, but on a user's sheet, narrowing a column must not
    ' squash their button.
```

- [ ] **Step 3: Add `SafeShapeName`**

Insert near `AddBtn`:

```vba
' Map any character outside A-Za-z0-9_ to "_", so the action id "clean.dupes"
' composes into the legal shape name "SFPlaced_clean_dupes".
' A theoretical collision between two ids differing only by that substitution is
' harmless: placement is sweep-then-add and never looks a button up by name.
Private Function SafeShapeName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") _
           Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            out = out & ch
        Else
            out = out & "_"
        End If
    Next i
    SafeShapeName = out
End Function
```

- [ ] **Step 4: Add `SplitAnchored` — strictly downstream of the single funnel**

Insert directly after `GatherButtons`'s `End Function` (`Core.bas:2528`), so the funnel and
its partition read together:

```vba
' Partition the SINGLE funnel's output into the buttons the page pipeline draws
' and the buttons anchored onto a user's own sheet.
'
' Runs strictly DOWNSTREAM of GatherButtons, so "hide" has already been applied to
' both halves. DO NOT re-filter "hide" here, and DO NOT assemble either
' collection any other way - GatherButtons' own comment records what happened the
' last time a filer was fed something assembled separately: "hide" silently
' stopped working, on the live path, and it took a BLOCKING fix round to find.
'
' Only "paged" reaches either layout filer, so an anchored button never affects
' page counts, column caps, category packing or continuation pages.
Private Sub SplitAnchored(buttons As Collection, codeMap As Object, _
                          ByRef paged As Collection, ByRef anchored As Collection)
    Dim e As Variant, id As String, code As String
    Set paged = New Collection
    Set anchored = New Collection
    For Each e In buttons
        id = LCase(CStr(e(0)))
        code = ""
        If codeMap.Exists(id) Then code = Trim(CStr(codeMap(id)))
        If CodeKind(code) = "anchor" Then
            anchored.Add e
        Else
            paged.Add e
        End If
    Next e
End Sub
```

- [ ] **Step 5: Add `SweepPlacedButtons`**

Insert near `DeletePanels`:

```vba
' Delete every button this engine has EVER placed on a user's sheet, across the
' whole workbook, before re-placing from the CURRENT codes.
'
' Self-healing by construction: a moved cell, a cleared code, a renamed sheet,
' and an orphan left by a build that died halfway are all fixed by the same pass.
' It runs on EVERY build, including builds with zero anchors - that is exactly
' how clearing a code removes its button.
'
' DeletePanels cannot help here: it only deletes generated "Macros N" sheets, and
' an anchored button lives on a sheet the engine must otherwise never touch.
'
' Shapes are iterated BACKWARDS - the established idiom in EnsurePanelSheet -
' because deleting one shifts every later index down.
Private Sub SweepPlacedButtons()
    Dim ws As Worksheet, si As Long
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For si = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(si).Name, Len(PLACED_PREFIX)) = PLACED_PREFIX Then _
                ws.Shapes(si).Delete
        Next si
    Next ws
    On Error GoTo 0
End Sub
```

- [ ] **Step 6: Add `AnchorSheet` and `PlaceAnchoredButtons`**

Insert after `SweepPlacedButtons`:

```vba
' Resolve an anchor's sheet part to a Worksheet (Nothing when it does not exist)
' and hand back its cell part. The single place an anchor becomes objects, so the
' Finalize checks and the placer cannot disagree about what "TabName!B4" resolves to.
Private Function AnchorSheet(ByVal code As String, ByRef cellRef As String) As Worksheet
    Dim sh As String
    cellRef = ""
    If Not ParseAnchor(code, sh, cellRef) Then Exit Function
    If Not Core_SheetExists(sh) Then Exit Function
    Set AnchorSheet = ThisWorkbook.Worksheets(sh)
End Function

' Anchor one button per placement onto a sheet the engine otherwise leaves
' COMPLETELY alone. Writes to no cell: no Clear, no UnMerge, no column width, no
' row height, no value, no format, no hyperlink. The ONLY mutation to a user's
' sheet is adding - and, via SweepPlacedButtons, removing - a shape whose name
' carries this engine's reserved prefix.
'
' This is deliberately a separate procedure, NOT a parameter on RenderPage. The
' card-grid renderer owns its sheet outright: EnsurePanelSheet does Cells.Clear,
' Cells.UnMerge, Hyperlinks.Delete, Cells.RowHeight, six column widths, tab
' colour and interior fill. A user's own sheet must never be treated that way.
'
' Returns how many buttons were actually placed, for the completion message.
'
' A failure skips that one button rather than raising. Finalize has already
' blocked on a missing sheet, an engine-owned sheet, a duplicate target and a
' malformed anchor, so anything still failing here is a workbook that changed
' underneath its own config - and a build must not die on its last step because
' one anchor went stale.
Private Function PlaceAnchoredButtons(anchored As Collection, codeMap As Object) As Long
    Dim e As Variant, id As String, code As String, aCell As String
    Dim ws As Worksheet, target As Range, n As Long
    For Each e In anchored
        id = LCase(CStr(e(0)))
        code = ""
        If codeMap.Exists(id) Then code = Trim(CStr(codeMap(id)))
        Set ws = AnchorSheet(code, aCell)
        If Not ws Is Nothing Then
            Set target = Nothing
            On Error Resume Next
            Set target = ws.Range(aCell)
            On Error GoTo 0
            If Not target Is Nothing Then
                AddBtn ws, PLACED_PREFIX & SafeShapeName(CStr(e(0))), _
                       CStr(e(1)), CStr(e(2)), _
                       target.Left, target.Top, ANCHOR_BTN_W, BTN_H, xlMove
                n = n + 1
            End If
        End If
    Next e
    PlaceAnchoredButtons = n
End Function
```

- [ ] **Step 7: Wire the pipeline**

`Core_BuildControlPanel` (`Core.bas:2433`). Extend the `Dim` line:

```vba
    Dim buttons As Collection, codeMap As Object, layout As Object, manual As Boolean
    Dim paged As Collection, anchored As Collection, nPlaced As Long
```

After `Set buttons = GatherButtons(codeMap)` insert:

```vba
    SplitAnchored buttons, codeMap, paged, anchored   ' downstream of the funnel - "hide"
                                             ' has already been applied to both halves
```

After `DeleteHomePanel` insert:

```vba
    SweepPlacedButtons                      ' after the deletes: fewer sheets left to walk
```

Change both filer calls to take `paged`, not `buttons`:

```vba
    If manual Then
        BuildManualLayout paged, codeMap, layout
    Else
        BuildAutoLayout paged, layout
    End If
```

After `BuildAdminPanel` and **before** `LandOnFirstPage layout`, insert:

```vba
    nPlaced = PlaceAnchoredButtons(anchored, codeMap)
```

- [ ] **Step 8: The one edit to the settled completion message**

`Core_BuildControlPanel`'s success message (settled in Task 5). Replace:

```vba
    If Not silent Then
        ' "Built 0 macro page(s)" is true but reads as a failure, and zero pages
        ' is a real, reachable state now that every button can be hidden or
        ' switched Off. Say what actually happened instead.
        If layout.Count = 0 Then
            MsgBox "No macro pages were built - every button is either hidden or switched Off.", _
                   vbInformation, "Setup Complete"
        Else
            MsgBox "Built " & layout.Count & " macro page(s) with " & buttons.Count & _
                   " button(s).", vbInformation, "Setup Complete"
        End If
    End If
```

with:

```vba
    If Not silent Then
        Dim placedNote As String
        If nPlaced > 0 Then _
            placedNote = vbCrLf & vbCrLf & nPlaced & " button(s) placed on your own sheets."
        ' "Built 0 macro page(s)" is true but reads as a failure, and zero pages
        ' is a real, reachable state: every button can be hidden, switched Off,
        ' or anchored onto one of the user's own tabs.
        If layout.Count = 0 Then
            MsgBox "No macro pages were built - every button is hidden, switched Off, or " & _
                   "placed on one of your own sheets." & placedNote, _
                   vbInformation, "Setup Complete"
        Else
            ' paged.Count, NOT buttons.Count: an anchored button is not on a
            ' generated page and must not be counted against one. This is the
            ' same class of miscount the ledger records for BuildHomePage's
            ' totalButtons - one honest definition, not the nearest handy number.
            MsgBox "Built " & layout.Count & " macro page(s) with " & paged.Count & _
                   " button(s)." & placedNote, vbInformation, "Setup Complete"
        End If
    End If
```

- [ ] **Step 9: Run both structural gates**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
```

Then confirm `buttons` is no longer fed to either filer:

```
grep -n 'BuildManualLayout buttons\|BuildAutoLayout buttons' src/core/Core.bas
```

Expected: no output.

- [ ] **Step 10: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): place anchored buttons on a user's own sheet

The owner's original intent, unbuilt until now: 'I specifically wanted the
option to add a macro to a specific page and cell, not just organize the
buttons on new pages.'

A genuinely separate code path, not a parameter on RenderPage. The card-grid
renderer OWNS its sheet - EnsurePanelSheet does Cells.Clear, Cells.UnMerge,
Hyperlinks.Delete, Cells.RowHeight, six column widths, tab colour and interior
fill - and a user's own sheet must never be treated that way.
PlaceAnchoredButtons writes to no cell at all: the only mutation is adding a
shape carrying the reserved SFPlaced_ prefix.

SplitAnchored partitions strictly DOWNSTREAM of GatherButtons, so 'hide' has
already been applied to both halves and the single funnel stays single. Only
the paged half reaches either filer, so an anchored button never affects page
counts, column caps, category packing or continuation pages - and the
completion message counts paged.Count, not buttons.Count.

SweepPlacedButtons deletes every previously-placed button across the whole
workbook on EVERY build, including builds with zero anchors - which is exactly
how clearing a code removes its button, and what makes a moved cell, a renamed
sheet and a half-finished build all self-heal.

Placement = xlMove, not the card grid's xlMoveAndSize: on a generated page the
engine controls the column width, but on a user's sheet narrowing a column must
not squash their button. Width is a fixed 110pt (owner's call at the gate), not
the target cell's - a user's column may be three characters wide."
```

---

## Task 9: What blocks Finalize, and the contract for anchors

Design items 3g, 3h. All four conditions are **structural** — they block activation —
consistent with how a bad page code is already treated.

**Files:**
- Modify: `src/core/Core.bas`
- Modify: `spec/contract.md` — §6, §7, §10

**Interfaces:**
- Consumes: `CodeKind`, `ParseAnchor` (Task 7); `PLACED_PREFIX` (Task 8).
- Produces: `IsEngineOwnedSheet`, `AnchorAddrKey`.

- [ ] **Step 1: Add `IsEngineOwnedSheet`**

Insert directly after `IsPanelName`'s `End Function`:

```vba
' Every sheet name this engine owns and rebuilds, in ONE place so a future
' generated surface is added once rather than four times.
'
' An anchor pointing at one of these is blocked at Finalize: the engine deletes
' or clears these tabs on every build, so a button placed there would vanish -
' silently - on the very next rebuild, which is the worst possible way for a
' placement to fail.
Private Function IsEngineOwnedSheet(ByVal nm As String) As Boolean
    Select Case LCase(Trim(nm))
        Case LCase(SETUP_SHEET), LCase(ADMIN_SHEET), LCase(MACRO_SHEET), _
             LCase(HOME_SHEET), LCase(CONFIG_SHEET), LCase(PACKAGES_SHEET)
            IsEngineOwnedSheet = True
        Case Else
            IsEngineOwnedSheet = IsPanelName(nm)      ' every generated "Macros <n>"
    End Select
End Function
```

`HOME_SHEET` is in the list even though the index is withdrawn: the name stays reserved
(§7), and a restore must not find user buttons squatting on it.

- [ ] **Step 2: Add `AnchorAddrKey`**

Insert after `IsEngineOwnedSheet`:

```vba
' The resolved ABSOLUTE address of an anchor's cell ("$B$4"), or "" when the cell
' part does not resolve on that sheet. Normalising through Range.Address is what
' makes "B4" and "$B$4" collide correctly in the duplicate-target check - two
' spellings of one cell are one target.
' Callers must have confirmed the sheet exists first; a missing sheet yields ""
' here, which the caller reports as a bad code rather than as a duplicate.
Private Function AnchorAddrKey(ByVal sheetName As String, ByVal cellRef As String) As String
    Dim ws As Worksheet, rng As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then Set rng = ws.Range(cellRef)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    AnchorAddrKey = rng.Address(True, True)
End Function
```

- [ ] **Step 3: Add the four blocking checks and the soft placed list**

`ComposeReport` (`Core.bas:2101`). Extend the `Dim` block with:

```vba
    Dim aKey As String, placedList As String, seenTargets As Object
```

and initialise before the placement-code loop:

```vba
    Set seenTargets = CreateObject("Scripting.Dictionary")
```

Replace Task 7's `Case "anchor"` branch inside the placement-code loop with:

```vba
            Case "anchor"
                If Not ParseAnchor(code, aSheet, aCell) Then
                    s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf
                    structural = structural + 1
                ElseIf Not Core_SheetExists(aSheet) Then
                    s = s & "NO SUCH SHEET for " & k & ": '" & code & "' - this workbook " & _
                        "has no tab named " & Chr(34) & aSheet & Chr(34) & "." & vbCrLf
                    structural = structural + 1
                ElseIf IsEngineOwnedSheet(aSheet) Then
                    s = s & "RESERVED SHEET for " & k & ": '" & code & "' - the engine " & _
                        "rebuilds that tab on every build; pick one of your own." & vbCrLf
                    structural = structural + 1
                Else
                    ' Normalise through the resolved absolute address so "B4" and
                    ' "$B$4" collide correctly. Runs only after the sheet-exists
                    ' check has passed, so an empty key here means a cell part
                    ' that parses but does not resolve (a row past the sheet's
                    ' last) - a bad code, not a duplicate.
                    aKey = LCase(aSheet) & "!" & AnchorAddrKey(aSheet, aCell)
                    If Right$(aKey, 1) = "!" Then
                        s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf
                        structural = structural + 1
                    ElseIf seenTargets.Exists(aKey) Then
                        ' Two buttons drawn at identical coordinates means one
                        ' physically covers the other and is unclickable. Blocking
                        ' is consistent with how a bad page code is already
                        ' treated: every placement error here is structural.
                        s = s & "DUPLICATE TARGET: " & CStr(seenTargets(aKey)) & " and " & _
                            k & " both point at " & code & "." & vbCrLf
                        structural = structural + 1
                    Else
                        seenTargets(aKey) = CStr(k)
                        placedList = placedList & "  - " & HiddenTokenFor(CStr(k), chains) & _
                                     "  ->  " & code & vbCrLf
                    End If
                End If
```

Then, next to the existing Hidden list at the end of the function (`Core.bas:2278`), add
the soft informational list in the same `caption [id]` format both existing lists use:

```vba
    If Len(placedList) > 0 Then _
        s = s & "Placed on your own sheets (not on a generated page):" & vbCrLf & placedList
```

**Note on `map` iteration order.** `ReadCodeMap` inserts actions first (registry order),
then chains, and `Scripting.Dictionary` preserves insertion order — so
`DUPLICATE TARGET: <A> and <B>` names them deterministically, and two runs of the same
config produce the same message.

- [ ] **Step 4: Amend contract §6 with the anchor form**

`spec/contract.md`, in the **Grouping / layout** paragraph, after the placement-code
grammar, insert:

```
**Anchored placement (v1.2.0).** A placement value may instead take the form
`<sheet>!<cell>` (e.g. `TabName!B4`), which anchors that one button to a single cell on a
surface the user owns, rather than filing it onto a generated page. The
discriminator is `!`. The two forms can never be confused — not because the platform
forbids `!` in a surface name (Excel, for one, does not) but because a valid **page
code** cannot contain `!` under the grammar above, so even a surface legitimately
named `1A` is unambiguous. An implementation MUST split on the **last** `!`, so a
surface whose own name contains one still resolves. The sheet part may
be quoted (`'TabName Data'!B4`) or bare even when it contains spaces (`TabName Data!B4`); an
embedded apostrophe is doubled inside the quotes. The cell part is a single A1
reference with optional `$` (`TabName!$B$4`); a range (a colon) is rejected.

- **Normative: an anchor is NOT a placement code for the mode decision.** It does not
  count toward "any code is present" in either direction — the same carve-out `hide`
  already has. A store whose only non-blank values are anchors is a store with **no**
  placement codes, and auto-layout applies to every other button unchanged. Stated in
  both directions because reading this clause literally in the other direction is how
  an implementation dumps every uncoded button into Unassigned the moment one button
  is anchored.
- **Normative: the engine MUST NOT modify the target surface** beyond adding and
  removing its own control. No clearing, no resizing, no writing a value, no
  formatting, no row or column sizing. The target is a surface the user owns; the
  engine is a guest on it.
- **Normative: an implementation MUST remove its previously-placed controls on
  rebuild**, across the whole deployment, before placing from the current values — so
  a moved cell, a cleared value, a renamed sheet, or a build that died halfway cannot
  leave a live orphan behind.
- **Blocking conditions.** Finalize MUST refuse to activate, naming the button, when:
  the named surface does not exist; the named surface is one the engine itself
  generates and rebuilds; two buttons resolve to the same target cell (two controls at
  identical coordinates means one covers the other and is unreachable); or the value is
  malformed.
- Both the action catalog and the chains block use this same grammar, so a **chain**
  may be anchored exactly like an action.
```

- [ ] **Step 5: Amend contract §7 to reserve the shape-name prefix**

`spec/contract.md`, after the existing v1.2.0 reserved-names bullet, add:

```
- **Reserved (v1.2.0):** in the Excel reference implementation, the shape-name prefix
  `SFPlaced_`. Every control the engine anchors onto a user's own surface carries it,
  and the engine removes controls by that prefix alone on every rebuild — so anything
  else in the workbook using the prefix is deleted without warning. A platform
  reserves its own equivalent namespace for the same purpose.
```

- [ ] **Step 6: Amend contract §10's v1.2.0 bullet**

Add to the v1.2.0 bullet's list of additions:

```
the anchored `<sheet>!<cell>` placement form with its engine-must-not-modify and
must-sweep-on-rebuild rules,
```

and to the same bullet's justification, append:

```
The anchored placement form adds no stored `key`, `id` or `kind`; it extends an
existing value's grammar with a form no valid v1.1.0 config could contain (a value
containing `!` was previously invalid and blocked Finalize), so no existing config is
reinterpreted. **Additive.**
```

**No version bump.** v1.2.0 is unreleased.

- [ ] **Step 7: Run both structural gates**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
```

- [ ] **Step 8: Commit**

```bash
git add src/core/Core.bas spec/contract.md
git commit -m "feat(core): block Finalize on a bad, reserved or duplicated anchor

Four conditions, all STRUCTURAL - consistent with how a bad page code is
already treated:
- NO SUCH SHEET: the workbook has no tab by that name.
- RESERVED SHEET: an engine-owned tab (Setup, Setup Menu, Macros, any
  Macros <n>, Macros Home, _Config, _Packages), collected into one new
  IsEngineOwnedSheet so the list lives in exactly one place. A button placed
  there would vanish silently on the next rebuild.
- DUPLICATE TARGET: two buttons at identical coordinates means one covers the
  other and is unclickable. Normalised through the resolved Range.Address, so
  B4 and \$B\$4 collide correctly.
- BAD CODE: reuses the existing malformed-code message.

ComposeReport also gains a soft, informational list of buttons placed on the
user's own tabs, alongside the existing Unassigned and Hidden lists and in the
same 'caption [id]' format both already use.

Contract 6 gains the anchor grammar with the mode carve-out stated in BOTH
directions (so a port reading it literally cannot reintroduce the bug the ledger
records almost happening with 'hide'), the MUST-NOT-modify-the-target-surface
rule, the MUST-sweep-on-rebuild rule and the four blocking conditions. 7
reserves the SFPlaced_ shape-name prefix. 10's v1.2.0 bullet records the form as
additive. No version bump; v1.2.0 is unreleased."
```

---

## Task 10: Re-cut the owner's Task 10 checklist

The checklist in `docs/designs/2026-07-24-finish-excel-track1-plan.md` (Task 10 section,
lines **2168-2462**) currently asks the owner to verify a feature that no longer exists.
The ledger records that this exact drift already happened once: two checks in that section
"would have FAILED a correct build."

**Files:**
- Modify: `docs/designs/2026-07-24-finish-excel-track1-plan.md` — Task 10 section only
- Regenerate: `docs/code-map.md`

**Interfaces:** none.

- [ ] **Step 1: Delete every item that verifies the withdrawn index**

From the Task 10 section, delete:

- Step 2's `` - [ ] `Macros Home` exists and is the first tab. ``
- Step 2's `` - [ ] `<< Home` is in the same cell on every page and works `` — **but keep
  the `< Previous` / `Next >` half of that item**, re-worded to check they now sit in
  cells **B and D**.
- Step 2's `` - [ ] Home's button catalog lists all 82 buttons... ``
- Step 3's `Home's counts drop by one` clause (the rest of that item stays).
- Step 5's entire **"Force the failure path once, deliberately"** item — its trigger is
  protecting the `Macros Home` tab, which no longer exists. **Do not simply drop it**:
  re-aim it at a tab that still exists. Replace the trigger with: *right-click the
  `Setup Menu` tab ▸ Protect Sheet ▸ OK, then Finalize* — `BuildAdminPanel` reuses that
  sheet (it is never deleted by `DeletePanels`) and `EnsurePanelSheet`'s `Cells.Clear`
  raises 1004 on a protected sheet, so the same two-dialog sequence is exercised. Keep
  the "exactly two dialogs, in this order" requirement and the recovery step verbatim.
- Step 6's `` - [ ] Home lists **one** page holding two category columns `` — re-word to
  check the same fact on the pages themselves.
- Step 7b's `` - [ ] **Macros Home's own links** `` item entirely.
- Step 7b's `` - [ ] **A chain's card vs. its Home catalog row.** `` — **keep the card
  half** (the bounded `+N more` and the "no `+0 more`" check) and delete the Home-catalog
  comparison, re-aiming the "see the full recipe" pointer at the Setup tab's Chains block.
- Step 7b's shrink-on-rebuild item: delete `no dead links on 'Macros Home' pointing at a
  page that no longer exists` and `exactly one 'Macros Home' tab (not a 'Macros Home (2)')`;
  **keep** the "no tall rows left behind from the fixture's longer descriptions" check.
- Step 7b's all-off item: replace `Confirm 'Macros Home' still builds and reads 'No actions
  are available as buttons yet.'` with a check on the new completion message — *confirm the
  dialog reads "No macro pages were built - every button is hidden, switched Off, or placed
  on one of your own sheets." and that no error appears.*

- [ ] **Step 2: Re-aim the AutoFit check — do NOT drop it**

The AutoFit risk was Home's un-floored catalog. With that gone, `Macros 1!F3` becomes the
**primary** AutoFit check.

In Step 2's **THE AUTOFIT CHECK** item, delete site **2** (`HomeCatalog`, "no floor at
all", the `Suggest Import Types` / 194-character check) and delete the entire
**"The risk is asymmetric"** paragraph — its whole argument is about `Macros Home` being
reused rather than recreated, and it concludes "if this bug fires anywhere, it fires only
on `Macros Home`". With Home gone that paragraph is not just stale, it actively tells the
owner the remaining check cannot fail.

Replace it with:

```
      **This is now the primary AutoFit check, and the risk is no longer
      asymmetric.** Every `Macros N` page is freshly created by `Worksheets.Add`,
      which activates it, so its AutoFit always runs on the active sheet - on a
      first build and on every rebuild alike. There is no "only from the second
      build onward" caveat any more, and no reused sheet to blame. If `Macros 1!F3`
      wraps correctly, tick this.
```

Keep sites **1** (`FitDescRows`, floored by `DESC_MIN_H`) and **3** (`RenderPage`'s
category-header row, no floor) exactly as written, and promote site 3's `Macros 1!F3` check
— `Clean & Standardize (cont. 2)`, **29** characters, 11pt bold, in the 30-character card
column — to the item's headline.

- [ ] **Step 3: Add the new items**

Add to **Step 2 (fixture pass — packing and geometry)**:

```
- [ ] **Timed build.** Note the wall-clock time from clicking `Create Setup Tab` to
      the Setup tab appearing, and again for `Finalize`. The owner measured
      **several minutes** for the first build before the registry was memoized.
      Anything still in minutes means the cache is not being hit - check that
      `InvalidateRegistry` was not also wired into `Core_Run`.
- [ ] The action catalog's third column header reads **`Placement`**, not
      `Page code`, in BOTH the Action catalog block and the Chains block.
- [ ] The `Availability` cells still have a working dropdown offering exactly
      `Button` / `Chains only` / `Off`. It is now applied to the whole block in one
      call instead of cell by cell, so this is confirming a refactor did not drop it.
- [ ] No `Macros Home` tab exists. If you are rebuilding a workbook that had one
      from an earlier build, confirm it has been **deleted**, not left behind.
- [ ] The build lands you on **`Macros 1`**, cell A1.
- [ ] The completion dialog does **not** mention an index.
- [ ] **The nav row on a SINGLE-page build.** `RenderNavRow` still sets row 1 to
      18pt unconditionally, but with only one page there is no `< Previous` and no
      `Next >` — and `<< Home` no longer occupies B1 — so row 1 renders as a blank
      18pt band. Behaviourally correct, but your original complaint was clutter and
      a reserved-but-empty band is the mirror image of it. **Judge it:** if it reads
      as wasted space, suppressing the row height when both links are absent is a
      two-line change.
```

Add a new **Step 5b: anchored placements** after Step 5:

```
- [ ] **Step 5b: Fixture pass — Sheet!Cell anchored placements**

- [ ] Create a scratch tab named `TabName`. Put `TabName!B4` in one action's Placement
      cell, leave every other code blank, and Finalize. Confirm: a button appears on
      `TabName` at cell B4, **and the other ~81 buttons are still packed automatically**
      onto numbered pages - NOT dumped into page 1's `Unassigned` column. An anchor
      must not flip the workbook into manual mode.
- [ ] The button on `TabName` is about **110pt** wide. Judge it against real data: it is
      the one number in this release that only your eyes can settle. If it reads
      wrong, `ANCHOR_BTN_W` is a one-line change.
- [ ] **`TabName` is otherwise untouched.** No cleared cells, no changed column widths or
      row heights, no formatting, no new values, no tab colour change. The button is
      the ONLY thing that appeared.
- [ ] Narrow column B on `TabName`. The button does **not** squash (it is `xlMove`, not
      `xlMoveAndSize`). Insert a row above row 4: the button moves down with its cell.
- [ ] Click it. It runs the action.
- [ ] **Clear that Placement cell and Finalize again: the button is GONE from `TabName`**,
      and the action is back on a generated page. This is the sweep working; if the
      button is still there, `SweepPlacedButtons` is not running.
- [ ] Change the code to `TabName!D10` and Finalize: exactly **one** button on `TabName`, at
      D10. Not two.
- [ ] Try each accepted spelling and confirm all are accepted: `TabName!B4`, `TabName!$B$4`.
      Rename the tab to `TabName Data` and confirm both `'TabName Data'!B4` and
      `TabName Data!B4` work.
- [ ] Anchor a **chain** the same way. It behaves identically - a chain is one button
      with one placement.
- [ ] **The four blocking conditions.** Each must refuse to activate, name the button,
      and change nothing:
      - `MICC!B4` (no such tab) -> `NO SUCH SHEET`
      - `Macros 1!B4` -> `RESERVED SHEET`
      - two different actions both set to `TabName!B4` -> `DUPLICATE TARGET` naming both
      - `TabName!B4:C9` (a range, not a cell) -> `BAD CODE`
      - also confirm `TabName!$B$4` and `TabName!B4` on two different actions **do** collide
        (they are the same cell spelled two ways)
- [ ] **`Macros 1!B4` on a workbook that has NOT been built yet** (a fresh import,
      before any `Macros N` tab exists). Finalize must say **`RESERVED SHEET`**, not
      `NO SUCH SHEET`. The wrong message here would tell you to create a tab called
      `Macros 1`, which `DeletePanels` destroys on the very next build.
- [ ] **A PROTECTED target sheet.** Right-click `TabName` ▸ Protect Sheet ▸ OK (the
      default protects `DrawingObjects`), then Finalize. Expected: **the build
      COMPLETES** — it must not die with "rebuilding the Macros pages failed" — and
      the completion dialog reports `1 anchored button(s) could NOT be placed`.
      **This is the single most important item in this section**: it is the one
      assumption that could not be verified without Excel (`ws.Buttons.Add` raising
      1004 on a protected sheet), and before the fix it took down the whole build.
      **Known residual, not a bug to report:** a button placed on that sheet by an
      *earlier* build survives, because the sweep cannot delete it either. Removing
      it would mean the engine unprotecting your sheet, which it must never do.
      Unprotect and re-Finalize to clear it.
- [ ] **A RENAMED target sheet.** Anchor a button to `TabName!B4`, Finalize, then rename
      `TabName` to something else and Finalize again. Expected: the completion dialog
      says `1 anchored button(s) could NOT be placed - the sheet or cell no longer
      exists...`. Confirm that action now has **no** button anywhere; that is
      correct and is what the message exists to tell you.
- [ ] With an anchor active, the Finalize report shows a soft
      `Placed on your own sheets` list in the same `caption [id]` format the
      `Hidden` and `Unassigned` lists use, and it does **not** block.
- [ ] Set every button to `hide` except one anchored button, then Finalize. Confirm
      the completion dialog reads `No macro pages were built - every button is
      hidden, switched Off, or placed on one of your own sheets.` followed by
      `1 button(s) placed on your own sheets.` - not an error, and not
      `Built 0 macro page(s)`.
```

Add to **Step 6 (real pass — the shipping bundle)**:

```
- [ ] **The duplicates header field.** On Setup, `Does the range you pick include a
      header row?` is present with a working **Yes/No dropdown**, defaulting to
      **Yes**.
- [ ] With it **Yes**, run `Remove Duplicates` over a range whose first row IS a
      header: the header row survives and is not compared.
- [ ] With it **No**, run it over `M2:M12` with no header: **all eleven values are
      considered**, not ten. This is the owner's original symptom.
- [ ] **Paste** (do not type) `ye` into that cell - typing is refused by the
      validation, so pasting is the only way in. Finalize **BLOCKS** with
      `UNRECOGNISED Yes/No value` and does not activate.
- [ ] Clear the cell entirely: it reads as the default (**Yes**) and activates.
- [ ] **Sheet-qualified messages.** Put a button on one tab and run
      `Remove Duplicates` against a range on a DIFFERENT tab. The completion dialog
      must now read `TabName!M2:M12`, not `M2:M12`.
      **If it names the WRONG tab, that is a genuinely new finding and the diagnosis
      in the design document is wrong - stop and report it.** The prediction is that
      it names the right one, and that nothing was ever acting on the wrong sheet.
- [ ] Do the same for `Trim & Clean Text` and `Fill Blanks Down`.
- [ ] Rename a tab to something needing quotes (e.g. `TabName Data` or `2024`) and
      confirm the message quotes it: `'TabName Data'!M2:M12`.
```

- [ ] **Step 4: Regenerate the code map**

```
node scripts/code-map.mjs
```

Then confirm it is current:

```
git log -1 --format=%h -- '*.bas' '*.cls' '*.frm' '*.mjs' '*.js' '*.ps1'
```

The header of `docs/code-map.md` must name that commit.

- [ ] **Step 5: Final full verification**

```
pwsh build/build.ps1
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
```

All four must pass. **Then stop.** The next step belongs to the owner: importing
`dist/MacrosBundle.fixture.bas` and running the re-cut checklist. Nothing in this plan may
be described as verified until that passes — this session's own history is the proof, since
a `Dim ... any As Boolean` compile error passed four review rounds and both automated gates.

- [ ] **Step 6: Commit**

```bash
git add docs/designs/2026-07-24-finish-excel-track1-plan.md docs/code-map.md
git commit -m "docs: re-cut the Task 10 checklist for the four gate items

Deleted every item verifying the withdrawn index - the page list, the button
catalog, the legend, the '<< Home' link, Home's own links and the 'no duplicate
Home tab' case.

Nothing was silently dropped. Two items were RE-AIMED rather than deleted:
- The AutoFit check. Home's un-floored catalog is gone, so the card-page
  category header at 'Macros 1!F3' ('Clean & Standardize (cont. 2)', 29 chars,
  11pt bold, in a 30-char column, no height floor) becomes the primary check.
  The 'risk is asymmetric' paragraph is deleted with it: its whole argument was
  that Macros Home is reused rather than recreated, and it concluded the bug
  could only fire there. Every Macros N page is freshly created and therefore
  active when its AutoFit runs, so the caveat is gone.
- The deliberate failure path. Its trigger was protecting the Macros Home tab;
  re-aimed at Setup Menu, which EnsurePanelSheet also reuses, so the same
  two-dialog sequence is still exercised.

Added: sheet-qualified messages, the duplicates header field with its four
states, a timed build, the Placement relabel, and a new Step 5b covering
anchored placements - the untouched-target-sheet check, the sweep-on-rebuild
check, all four blocking conditions, and the 110pt width judgement, which is the
one number in this release only the owner's eyes can settle.

Code map regenerated."
```

---

## Self-review against the spec

**Spec coverage.** Every section of `docs/designs/2026-07-25-track1-owner-feedback-design.md`
maps to a task:

| Design section | Task |
|---|---|
| 1a sheet-qualified references | 1 |
| 1b `clean_hasheader` | 2 |
| 1c first use of the `bool` kind | 2 |
| 1d contract §3/§8/§10 | 2 |
| 2a memoize `BuildRegistry` | 3 |
| 2b collapse `AvailDropdown` | 3 |
| 2c verification (timed build) | 10 |
| 4a move, do not delete | 4 (preserve) + 5 (remove) |
| 4b four non-obvious consequences | 5 |
| 4c completion message | 5 (settled) + 8 (one edit) |
| 4d Task 10 re-cut | 10 |
| 3a grammar + relabel ripple | 6 (relabel) + 7 (grammar) |
| 3b parsing, four call sites | 7 |
| 3c anchors are their own axis | 7 |
| 3d partitioning | 8 |
| 3e rendering | 8 |
| 3f rebuild cleanup | 8 |
| 3g validation | 9 |
| 3h contract §6/§7/§10 | 9 |
| Verification | every task's gate steps + 10 |

**Explicitly out of scope**, unchanged from the design and not to be picked up
opportunistically: the five procedures that walk the registry to find one action by id
(one `FindAction(id)` collapses four); manual mode's missing category concept at catalog
scale; `HasAnyTier` / `HasAdvanced` being named backwards; `"hide"` as five independent
literal comparisons where one `Const` would do; and the ~31 deferred minors in the ledger,
a third of which were already stale when triaged.

**Type consistency.** Checked against the Interfaces table: `CodeKind` returns `String` in
Task 7 and is compared to `"anchor"` / `"page"` in Tasks 7, 8 and 9. `ParseAnchor`'s
`ByRef` outs are `sheetName`/`cellRef`, and every call site declares `aSheet`/`aCell` as
`String`. `PlaceAnchoredButtons` is a `Function` returning `Long`, consumed as `nPlaced`.
`AddBtn`'s new parameter is `XlPlacement`, and `xlMove` is the only non-default value
passed. `SplitAnchored`'s `paged`/`anchored` are `ByRef ... As Collection` and are
`Set ... = New Collection` inside the sub, so the caller declares them without `New`.

**Identifier safety.** Every new identifier was checked against `build.ps1:160-166`'s
reserved-word list: `place`, `out`, `flag`, `raw`, `target`, `paged`, `anchored`,
`nPlaced`, `placedList`, `placedNote`, `seenTargets`, `aKey`, `aSheet`, `aCell`, `bTmp`,
`boolBad`, `boolBadN`, `firstDataRow`, `lastDataRow`, `letters`, `digits`, `sh`, `cl`,
`nm`, `si`, `p`, `t` — none collide. **Do not add words to that list.**

---

## Handoff log

### 2026-07-26 — all ten tasks built and reviewed; owner's Excel gate IN PROGRESS

**Branch.** `feat/finish-excel-track1`, **71 commits, NOT pushed, no PR.** Nothing
merged. `main` is still v0.3.0 / contract v1.1.0. This plan contributed **31** of
those commits (`f9436e6..HEAD`). Contract on the branch is v1.2.0, unreleased,
deliberately never bumped.

**What shipped.** All four owner-gate items, Tasks 1-10, each with a dedicated
spec+quality review; five needed one fix round, all resolved. Then a whole-branch
review (verdict: *blocked on the verification instrument, not the code*), one
consolidated fix wave of nine findings, and one scoped re-review — all addressed.

Test count before and after: **n/a, and not an omission.** No VBA runner exists and
`test-driven-development` is formally off (CLAUDE.md, "Off"). Verification is
`pwsh build/build.ps1` on both bundles plus `node scripts/vba-syntax-check.mjs` on
both — structural only — plus the owner's manual import.

**TWO REAL CODE DEFECTS found by the manual gate. Do not revert either fix.**

1. **`AddBtn` was not atomic** (`816fd47`). It has no error handler and runs under
   the caller's `On Error Resume Next`; in VBA an error inside a handler-less
   procedure resumes at the statement *after* the call, abandoning the rest of the
   callee. `Buttons.Add` had already created the shape, so every failed placement
   left a button carrying Excel's **default** name (`Button 11`), no caption and no
   `OnAction` — and because `SweepPlacedButtons` matches on the `SFPlaced_` name
   prefix, it could never remove it. **Permanent orphan, one per failed placement.**
   Now all-or-nothing: delete the partial shape, re-raise the original error.
2. **`eNum` IS the reserved word `Enum`** (`8f3a719`). VBA identifiers are
   case-insensitive. I introduced it in the fix for (1) and it broke
   `Debug ▸ Compile` on import. **Second instance of this class after `any`**, found
   the same way — by a human importing the bundle, after both automated gates passed
   it clean. `'enum'` added to `build.ps1`'s reserved list and **demonstrated**
   failing on a reintroduced `eNum` / passing clean, matching how the `any` gate was
   proven. That list's rule stands: add a word when Excel rejects one, never on
   suspicion.

**What the plan got wrong — do not revert these corrections.**

1. **`'TabName Data'!B4` could never work in Excel.** A leading apostrophe is Excel's
   "treat as text" prefix: consumed by the cell, stripped from `.Value`, still shown
   in the formula bar. The engine received `TabName Data'!B4` and reported `NO SUCH
   SHEET` for a tab in plain view. That spelling was listed **first** in contract §6
   and in the checklist. Owner chose accept-and-repair; `ParseAnchor` now treats an
   orphaned closing quote as the eaten pair, and §6 carries a normative clause for
   platforms whose cells do not eat the character. **NOT YET VERIFIED IN EXCEL.**
2. **"Excel forbids `!` in a sheet name" is FALSE.** The forbidden set is
   `: \ / ? * [ ]`. The claim came from the design doc, propagated into the plan and
   the landed comment, and was headed for `spec/contract.md` as normative text for
   the Apps Script port. The conclusion still holds for two other reasons: a valid
   page code can never contain `!` (`ParseCode`'s grammar rejects it), and
   `ParseAnchor` splits on the **last** `!`. **Do not "simplify" `InStrRev` to
   `InStr`** — it would silently break anchors onto a sheet whose name contains `!`.
   The design doc is annotated in place; its body is deliberately unedited.
3. **`Core_Run` and `Core_BuildControlPanel` are NOT in Alt+F8** (`055d24b`).
   Excel's Macro dialog lists only `Public Sub`s taking **no arguments** — an
   *optional* argument is enough to hide one. Two checklist steps and `build.ps1`'s
   shipped bundle header all told users to run them that way. Use the VBE Immediate
   window (Alt+F11, Ctrl+G).
4. **The Finalize check order was wrong** (Task 9 fix round). `Macros <n>` tabs are
   created *after* Finalize, so `Macros 1!B4` on a new deployment reported `NO SUCH
   SHEET` and the obvious fix — create that tab — is something `DeletePanels`
   destroys next build. Order is now malformed → engine-owned → sheet-missing →
   duplicate.

**What was discovered — empirical, with evidence.**

- The macros-disabled banner the owner hit was **Mark-of-the-Web** (file in
  `D:\Downloads\`), not a settings change. Unblock via file Properties.
- Chain anchors failed during the gate with "couldn't be placed" while actions never
  did. **On current code the chain places correctly**, and a step-by-step diagnostic
  showed `Buttons.Add` / `.Name` / `.Caption` / `.OnAction` / `.Font` / `.Placement`
  all OK (shape name 26 chars — the name-length hypothesis was **wrong**). Both
  defects above were live at the time and that workbook had accumulated orphans.
  **Cause not determined; deliberately left unattributed. Do not invent one.** If it
  recurs on clean code it is a new finding.
- The checklist itself was the main source of friction: **six items** could not be
  executed or judged as written. Root cause was naming internal identifiers where
  the Setup tab shows labels. Now rewritten and consolidated (37 remaining items →
  24) and filed as GitHub **issue #3** with a proposed blind-dry-run gate.

**What is still open, and what needs a human.**

- **The gate itself.** Steps 1-5 pass; 5b partly. **24 items remain** across
  5b/6/7/7b. Only the owner can run them.
- **Never verified in Excel:** the apostrophe repair, the sweep-on-clear, the four
  blocking conditions, and — highest value — the **protected-sheet** item, which is
  the one assumption nothing agent-side can execute (`ws.Buttons.Add` raising 1004)
  and which took down the whole build before `816fd47`.
- **Needs the owner's eyes:** `ANCHOR_BTN_W` at 110pt was judged acceptable once;
  worth a second look against real data.
- **Blocking Task 11** (teardown + ship): the remaining 24 items.

**Next phase.** Finish the manual gate, then Task 11. Read manifest:

| File | Size | Why |
|---|---|---|
| `docs/designs/2026-07-24-finish-excel-track1-plan.md` | 140K | **Task 10 section only** (~line 2168+) — the live checklist; Task 11 follows it |
| `.superpowers/sdd/2026-07-25-track1-owner-feedback-plan/progress.md` | 32K | the ledger: every review, fix round, parked finding and gate result |
| `docs/designs/2026-07-25-track1-owner-feedback-plan.md` | 129K | **this file's handoff log only** unless a task needs re-reading |
| `spec/contract.md` | 42K | normative; §3, §6, §7, §8, §10 all amended by this work |
| `docs/code-map.md` | 29K | grep for an anchor, then read the slice |
| `src/core/Core.bas` | 208K | **grep, never read whole** |

### 2026-07-25 — plan written, implementation not started

**Branch.** `feat/finish-excel-track1`, 40 commits, NOT pushed, no PR. `main` is still
v0.3.0 / contract v1.1.0.

**Owner decisions taken while writing this plan** (all three settle open items the design
flagged as needing a human):

1. **`ANCHOR_BTN_W` = 110pt**, not the design's 150pt. Compact, closer to a single
   default-width column, less intrusive next to real data. Still the one number only the
   owner's eyes can settle — Task 10's Step 5b asks for that judgement explicitly.
2. **`Placement = xlMove`** confirmed, unchanged from the design.
3. **A duplicate anchor target BLOCKS Finalize**, unchanged from the design. The owner
   pushed back on the first framing of this question — correctly, because it conflated the
   *button's drawing position* with *the data a macro operates on*. They are unrelated:
   contract §3 records that `clean` configures no target range at all and every action
   prompts for its range at run time. Two actions coded `TabName!B4` means two button graphics
   at identical coordinates, one covering the other. A chain is a single button with one
   placement, so chains never produce this.

**Verified independently this session, not taken on faith:**

- **The memoization stop condition holds.** `grep 'reg\.Add\|reg\.RemoveAll\|reg("fields")\.Add\|reg("actions")\.Add'`
  over `src/` returns hits only at `Core.bas:146-224` plus two comment lines. Task 3 Step 1
  re-runs that grep as a hard gate anyway.
- **`src/attic/` does not exist yet** — Task 4 creates it. `build.ps1` globs only
  `src/packages/Pkg*.bas` (`build.ps1:95`, `:121`), so nothing under `src/attic/` can reach
  a bundle or a gate.
- **`PageButtonCount` is orphaned by item 4.** Its only two callers are inside
  `BuildHomePage`. It moves to the attic with the rest, so the attic file is restorable on
  its own.
- **Every new identifier clears `build.ps1`'s reserved-word list.**

**One internal tension in the design resolved, flagged as DEVIATION 2.** The design says
§3 should record the `bool` encoding as "canonical lowercase" while also saying the
renderer mirrors `AvailDropdown`, which offers and stores title case. This plan resolves it
the way `availability` already resolves the same tension (§6: abstract tokens, not a
mandated on-screen spelling) — canonical `yes`/`no`, readers case-insensitive, Excel
presents and stores `Yes`/`No`. One-line reversal if the owner prefers otherwise.

**Next phase.** Execute Tasks 1-10 in order. Then the owner's manual import of
`dist/MacrosBundle.fixture.bas` against the re-cut Task 10 checklist. **That import is the
only real gate**, and it is what unblocks Task 11 (teardown + ship).
