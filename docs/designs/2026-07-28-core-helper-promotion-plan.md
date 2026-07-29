# Core pass 2 — the shared helpers move into Core

> **For agentic workers:** REQUIRED SUB-SKILL: use `superpowers:subagent-driven-development`
> to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give Core the twelve helpers W2–W6 would otherwise each re-invent, and leave every
package private carrying its package prefix — so the flat namespace stops being a landmine
and `CLAUDE.md` rule 4 is true rather than aspirational.

**Architecture:** A pure refactor. Twelve helpers move from `PkgClean` and `PkgImport` into
Core as `Core_*`; their old call sites change name only. Three helpers whose prompt text was
baked to one caller gain a parameter for it. Every remaining unprefixed package private is
renamed to carry its prefix. **No action changes behaviour, no action is added, and no
contract value moves** — the contract calls these names "an Excel/VBA detail, not part of the
portable contract" (§7, last bullet), so there is nothing here for `spec/contract.md` to say.

**Tech stack:** VBA (Excel/Windows, one concatenated `.bas`), PowerShell for the build, Node
for the syntax check. No test framework — see *Verification*.

**Runs after phase 1** (`2026-07-28-core-refuse-pass-plan.md`) and **before W2**. Sequencing
and rationale: `2026-07-28-track2-completion-plan-of-record.md` §3.1.

---

## Why this runs before W2, not after

Phase 1 exists because W1 duplicated two helpers for want of a Core API. The same thing
happens again at W2 in six more places, and phase 1 touches none of them.

**The flat namespace makes this sharper than ordinary duplication.** The bundle is one
module. `AskRange` sitting unprefixed and private in `PkgClean` does not merely fail to be
reusable — *it reserves the name*. `build.ps1:482` fails any build in which a second package
declares its own `AskRange`, private or not. So W2 has exactly two legal moves: promote it, or
invent a differently-named twin. Copying is what the standards forbid; inventing a twin is
what produced this pass.

| Helper | Today | Consumers after W2 |
|---|---|---|
| `AskRange` | `PkgClean.bas:603`, private, unprefixed | `PkgClean` + all 5 `validate.*` |
| `AskHasHeader` | `PkgClean.bas:626`, **title hardcoded `"Remove Duplicates"`** | `PkgClean` + `validate.reconcile`, `validate.scan` |
| `AskKeyCols` | `PkgClean.bas:645`, prompt + title baked to one caller | `PkgClean` + `validate.reconcile` |
| `ParseCols` | `PkgClean.bas:676` | `PkgClean` + `validate.*` |
| `CleanText` | `PkgClean.bas:692` | `PkgClean` + `validate.scan` |
| `IsBlankVal` | `PkgClean.bas:696` | `PkgClean` + `validate.missing_fields` |
| `PkgClean_NarrowToUsed` | `PkgClean.bas:582` | every action that takes a range |
| `PkgClean_IsAlpha` | `PkgClean.bas:496` | `PkgClean` + any column-letter parser |
| `IsAllDigits` | added by phase 1, private in Core | already Core's; only needs publishing |
| `PkgImport_DestUsable` | `PkgImport.bas:451` | + W2 report, W4 pivots, W5 snapshot, W6 TOC |
| the add-sheet-and-name recipe | **duplicated in-file** at `PkgImport.bas:192` and `:415` | same four |
| `PkgImport_SheetNameFor` | `PkgImport.bas:678`; its forbidden-char list is a **second copy** of the one in `PkgImport_DestUsable:469` | + W3 `export.per_sheet` |
| `PkgImport_ResolveFolder` | `PkgImport.bas:571` | + W3 export folder, W5 backup folder |
| `PkgImport_FolderExists` | `PkgImport.bas:585` | same |

**`AskHasHeader`'s hardcoded title is a live defect, not just an obstacle.** Any caller other
than Remove Duplicates gets a dialog captioned "Remove Duplicates" while asking about
something else. It has one caller today, so it has never shown; W2 gives it three.

---

## What is deliberately NOT promoted

Stated so W2–W6 do not read the omission as an oversight and quietly copy them.

- **`PkgImport_WorkbookPaths` (`:607`)** and **`PkgImport_AlreadyOpen` (`:652`)** stay in
  `PkgImport`. They have exactly one consumer today and W5's need for them is speculative —
  `backup.timestamped_copy` uses `SaveCopyAs` and opens nothing. Promoting on a guess is the
  YAGNI violation this pass would otherwise be evidence against. **If W5 turns out to need
  either, promote it then, in W5's own diff, and say so in that PR.**
- **`PkgClean_HasMerge` (`:275`)**, **`PkgClean_Contains` (`:409`)**,
  **`PkgClean_ScanBlanks`**, **`PkgClean_ConfirmBlanks`**, the two `*Note` builders, and
  `PkgClean_ResolveFormatCol` stay put. Each is one package's own job.
- **`PkgImportExport`'s five privates** are renamed, not promoted (Task 5). They are that
  package's own inference logic.

---

## Decisions taken in this plan

| # | Question | Decision |
|---|---|---|
| 1 | Promote everything generic, or only what has two consumers? | **Two consumers, or already duplicated.** Every promotion below is justified by an existing duplicate or a named W2–W6 consumer. The two speculative ones are listed above as not promoted. |
| 2 | `Core_AskHasHeader`'s title — new required parameter, or optional with a default? | **Optional, defaulting to `"Header row"`.** A required parameter would be a better forcing function, but it makes the `PkgClean` call site change shape rather than name, which is the one thing a pure refactor should avoid. The default is never what a caller wants, so it reads as a placeholder if anyone forgets. |
| 3 | Does `Core_DestUsable` keep showing its own dialog? | **Yes — it calls `Core_Refuse`,** which phase 1 made the single record-and-present call. Nothing here re-opens that. |
| 4 | Contract bump? | **None.** §7's last bullet puts VBA binding names outside the portable contract. Bumping would claim a portable change that did not happen. |

---

## Global constraints

- **`SCHEMA_VER` stays `"v2.0"`.**
- **`Option Explicit`**; every variable declared. **Pure ASCII** in every `.bas`. Max 25 line
  continuations per statement. **No module-level declaration after the first procedure**
  (assembled-bundle scope). No VBA reserved word as an identifier (`eNum` → `enum`).
- **Shared helpers live ONLY in Core.** Cross-package calls are forbidden.
- **Never read `src/core/Core.bas` whole** (232KB). Grep `docs/code-map.md` for an anchor and
  read that slice.
- **`dist/` is gitignored.** Never `git add -f` the bundle. `.vscode/` stays untracked.
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.
- **This plan changes no observable behaviour except `Core_AskHasHeader`'s dialog title.**
  If a diff changes anything a user could see, it is wrong — say so rather than absorbing it.

---

## Verification, since TDD is off

`CLAUDE.md` turns off `test-driven-development`: VBA has no headless runner and Excel cannot
be driven on CI. **Do not propose a test suite.** Every task verifies with:

1. `pwsh build/build.ps1` — exit 0, `Core + 3 package(s)`, balanced counts.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` — no diagnostics.
3. **A grep proving the old name is gone**, which for a pure rename is the actual proof.
4. **A deliberate injection** (Task 6) for the one thing meant to fail.

**A pure refactor is the one kind of change the build CAN largely verify**, because the
duplicate-procedure guard and the entry-point resolver both fire on a half-finished move.
That is why this phase carries no run-time-only claims into the phase-8 validation log.

---

## File structure

| File | Change |
|---|---|
| `src/core/Core.bas` | **+12 `Core_*` helpers**, in two new sections; `IsAllDigits` (phase 1) becomes `Core_IsDigits`. |
| `src/packages/PkgClean.bas` | 8 helpers deleted, call sites renamed; 3 remaining privates gain the `PkgClean_` prefix. |
| `src/packages/PkgImport.bas` | 5 helpers deleted or shrunk, call sites renamed; the in-file duplicated sheet-creation recipe collapses to one `Core_NewSheet` call. |
| `src/packages/PkgImportExport.bas` | 5 privates gain the `PkgImportExport_` prefix. |
| `build/build.ps1` | **+ a guard** that fails an unprefixed package private (Task 6). |
| `CHANGELOG.md`, `VERSION`, `docs/code-map.md` | v0.12.0. |

---

### Task 1: Core gains the prompt helpers

**Owner: Opus, written directly.** This is Core, and `Core_AskHasHeader` carries the one
behaviour fix in the pass.

**Files:** Modify `src/core/Core.bas` — add a new section immediately after
`Core_RangeRef` ends (line ~804, before `' ---- last row / last column ----`).

**Interfaces:**
- Consumes: `Core_RangeRef` (existing, public).
- Produces: `Core_AskRange`, `Core_AskHasHeader`, `Core_AskCols`, `Core_ParseCols`. Tasks 3
  and 4 depend on all four; every W2–W6 plan is written against them.

- [ ] **Step 1: Add the section header and the four procedures**

Insert after `Core_RangeRef`'s `End Function`:

```vb
'---- user prompts ----
' These live in Core because a package that asks the user for a range is not doing
' anything package-specific, and because the bundle is ONE module: an unprefixed
' Private AskRange in PkgClean reserves that name against every other package
' (build.ps1's duplicate-procedure guard), so the second package to want one cannot
' copy it and cannot re-declare it. Promoting is the only move that is not a
' standards deviation.

' Prompt the user to select a range. Returns Nothing if they cancel
' (InputBox Type:=8 returns Boolean False on cancel, which fails the Set).
Public Function Core_AskRange(ByVal title As String) As Range
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select the range for '" & title & "':", _
        title:=title, Type:=8)
    On Error GoTo 0
    If Not rng Is Nothing Then
        If rng.Areas.Count > 1 Then
            MsgBox "Please select a single, continuous block (not several separate areas).", _
                   vbExclamation, title
            Set rng = Nothing
        End If
    End If
    Set Core_AskRange = rng
End Function

' Ask whether the picked range's first row is a header. Returns False on Cancel,
' leaving hasHeader untouched and the caller responsible for stopping.
'
' vbYesNoCancel, not vbYesNo: a two-button dialog has no way out, and the title-bar
' X on a vbYesNo MsgBox returns vbNo - which for a caller that deletes rows would
' silently choose "compare every row" and delete a header row. Cancel has to be
' reachable and distinct.
'
' title is a parameter because this used to hardcode "Remove Duplicates" while it
' had exactly one caller. Every other caller would have shown a dialog captioned
' with an action it was not running.
Public Function Core_AskHasHeader(rng As Range, ByRef hasHeader As Boolean, _
                                  Optional ByVal title As String = "Header row") As Boolean
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Does " & Core_RangeRef(rng) & " include a header row?" & vbCrLf & vbCrLf & _
                 "Yes  -  the first row is a title row: it is kept, and never " & _
                 "treated as data." & vbCrLf & _
                 "No  -  every row is data, including the first one.", _
                 vbQuestion + vbYesNoCancel + vbDefaultButton1, title)
    If ans = vbCancel Then Exit Function
    hasHeader = (ans = vbYes)
    Core_AskHasHeader = True
End Function

' Ask for one or more column numbers, counted from the LEFT of rng. Returns False
' on Cancel or on a number outside the range's width; keys comes back as the raw
' "1,2,4" string, which Core_ParseCols turns into numbers.
'
' Application.InputBox, not VBA's InputBox: the latter returns "" for BOTH Cancel
' and an empty OK, and those must not mean the same thing. Application.InputBox
' returns Boolean False on Cancel, which VarType can tell apart from any string.
'
' question and title are parameters for the same reason Core_AskHasHeader's title
' is: this was written for Remove Duplicates and said so in both.
Public Function Core_AskCols(rng As Range, ByRef keys As String, _
                             ByVal question As String, ByVal title As String) As Boolean
    Dim v As Variant, cols As Variant, i As Long, nCols As Long
    nCols = rng.Columns.Count
    v = Application.InputBox( _
        Prompt:=question & vbCrLf & vbCrLf & _
                "Numbered from the LEFT of " & Core_RangeRef(rng) & ", which is " & _
                nCols & " column(s) wide. One number, or several like 1,2,4.", _
        title:=title, Default:="1", Type:=2)
    If VarType(v) = vbBoolean Then Exit Function     ' Cancel
    keys = Trim(CStr(v))
    If Len(keys) = 0 Then keys = "1"                 ' empty OK means the default

    ' Validate against the picked range's width. A column number past the right edge
    ' raises 1004 deep inside whatever the caller does with it, which surfaces through
    ' Core_Fail as an engine error - it reads as a broken macro rather than as the typo
    ' it is. Core_ParseCols already drops non-numeric entries and falls back to 1.
    cols = Core_ParseCols(keys)
    For i = LBound(cols) To UBound(cols)
        If cols(i) < 1 Or cols(i) > nCols Then
            MsgBox "Column " & cols(i) & " is outside " & Core_RangeRef(rng) & _
                   ", which is " & nCols & " column(s) wide." & vbCrLf & vbCrLf & _
                   "Nothing has been changed. Run " & title & " again.", _
                   vbExclamation, title
            Exit Function
        End If
    Next i
    Core_AskCols = True
End Function

' "1,2,4" -> Variant array of Longs, 0-based. Non-numeric entries are ignored;
' an empty result defaults to column 1.
Public Function Core_ParseCols(ByVal s As String) As Variant
    Dim parts() As String, i As Long, c As Collection, out() As Variant, k As Long
    Set c = New Collection
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        If IsNumeric(Trim(parts(i))) Then c.Add CLng(Trim(parts(i)))
    Next i
    If c.Count = 0 Then c.Add CLng(1)
    ReDim out(0 To c.Count - 1)
    For k = 1 To c.Count
        out(k - 1) = c(k)
    Next k
    Core_ParseCols = out
End Function
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: **FAIL**, and this is the point. The duplicate-procedure guard is silent here
(`Core_AskRange` and `AskRange` are different names), so the build should PASS with `Core + 3
package(s)` and `Function` up by four. If it instead reports a duplicate, a name collision
was missed — stop and re-check rather than renaming around it.

- [ ] **Step 3: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): Core_AskRange, Core_AskHasHeader, Core_AskCols, Core_ParseCols"
```

---

### Task 2: Core gains the value, range, sheet and folder helpers

**Owner: Opus, written directly.** Core.

**Files:** Modify `src/core/Core.bas` — the value/range helpers go beside `Core_LastCol`
(~line 850); the sheet and folder helpers go after `Core_FileName` (~line 908).

**Interfaces:**
- Consumes: `Core_Refuse` (phase 1), `Core_SheetExists` (`:782`), `Core_FileName` (`:902`).
- Produces: `Core_CleanText`, `Core_IsBlank`, `Core_IsAlpha`, `Core_IsDigits`,
  `Core_NarrowToUsed`, `Core_DestUsable`, `Core_NewSheet`, `Core_SafeSheetName`,
  `Core_ResolveFolder`, `Core_FolderExists`.

- [ ] **Step 1: Publish phase 1's `IsAllDigits`**

Phase 1 Task 3 added `Private Function IsAllDigits(ByVal s As String) As Boolean`. Rename it
`Core_IsDigits` and make it `Public`, and update the one caller inside
`Core_IsEngineSheet`:

```vb
' Every character is 0-9. Empty is False.
Public Function Core_IsDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    Core_IsDigits = True
End Function
```

and inside `Core_IsEngineSheet`, `IsAllDigits(tail)` becomes `Core_IsDigits(tail)`.

- [ ] **Step 2: Add the value and range helpers**

```vb
' Every character is A-Z (case-insensitively). Empty is False. Used to tell a
' column LETTER spec apart from a header NAME spec.
Public Function Core_IsAlpha(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = UCase$(Mid$(s, i, 1))
        If ch < "A" Or ch > "Z" Then Exit Function
    Next i
    Core_IsAlpha = True
End Function

' Excel TRIM (collapse runs of spaces) + CLEAN (strip non-printing characters).
Public Function Core_CleanText(ByVal s As String) As String
    Core_CleanText = Application.WorksheetFunction.Clean(Application.WorksheetFunction.Trim(s))
End Function

' Empty, or the empty string. NOT the same as IsEmpty: a cell holding "" is blank
' to a user and is not Empty to VBA.
Public Function Core_IsBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        Core_IsBlank = True
    ElseIf VarType(v) = vbString Then
        Core_IsBlank = (v = "")
    End If
End Function

' Narrow a picked range to the part of the sheet that actually holds anything.
' Returns Nothing when the pick lies entirely outside the used area - the caller
' reports that and stops, having changed nothing.
'
' Selecting whole columns is how people pick "this table" - A:D is one click on the
' column headers - and it hands an action roughly four million cells. Anything doing
' COM work per cell or per row then freezes the screen for minutes. Nothing outside
' UsedRange can hold a value, a formula or a merge, so the intersection examines the
' same data and finishes in seconds.
'
' Callers must use the NARROWED range everywhere downstream, including in the range
' they name in their messages, so what the user is told matches what was looked at.
Public Function Core_NarrowToUsed(rng As Range) As Range
    If rng Is Nothing Then Exit Function
    Set Core_NarrowToUsed = Application.Intersect(rng, rng.Parent.UsedRange)
End Function
```

- [ ] **Step 3: Add the sheet helpers**

```vb
' Characters Excel forbids in a sheet name. Declared once, here, because it was
' written out twice inside PkgImport alone - once to reject a name and once to
' repair one - and the two copies could drift apart silently.
Public Function Core_BadSheetChars() As String
    Core_BadSheetChars = ":\/?*[]"
End Function

' A destination name this deployment may create. It REFUSES rather than overwrites:
' the library snippet this replaces opened with Worksheets("Master").Delete under
' On Error Resume Next, which silently destroys whatever the user had there.
' Nothing in this toolkit deletes a sheet it did not create in the same run.
'
' Refuses via Core_Refuse, so a chain STOPS here rather than running its remaining
' steps against a sheet this one did not produce (contract section 6).
Public Function Core_DestUsable(ctx As Object, ByVal destName As String, _
                                ByVal title As String) As Boolean
    Dim bad As String, i As Long
    If Len(destName) = 0 Then
        Core_Refuse ctx, "No name is set for the sheet this creates." & vbCrLf & vbCrLf & _
               "Fill the destination name in on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", title
        Exit Function
    End If
    If Len(destName) > 31 Then
        Core_Refuse ctx, "'" & destName & "' is " & Len(destName) & " characters. Excel " & _
               "sheet names stop at 31." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", title
        Exit Function
    End If
    bad = Core_BadSheetChars()
    For i = 1 To Len(bad)
        If InStr(destName, Mid$(bad, i, 1)) > 0 Then
            Core_Refuse ctx, "'" & destName & "' contains " & Mid$(bad, i, 1) & ", which " & _
                   "Excel does not allow in a sheet name." & vbCrLf & vbCrLf & _
                   "None of these are allowed:  : \ / ? * [ ]" & vbCrLf & vbCrLf & _
                   "Nothing has been changed.", title
            Exit Function
        End If
    Next i
    If Core_SheetExists(destName) Then
        Core_Refuse ctx, "A sheet named '" & destName & "' already exists in this workbook, " & _
               "and this action will not overwrite it." & vbCrLf & vbCrLf & _
               "Rename or delete that sheet, then run " & title & " again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", title
        Exit Function
    End If
    Core_DestUsable = True
End Function

' Add a new sheet at the END of the workbook and name it. Call Core_DestUsable
' FIRST - this does no validating, deliberately, so that the refusal and the
' creation are not tangled into one call whose failure mode is ambiguous.
'
' Written once here because the same two lines were repeated verbatim at two
' places inside PkgImport, and four more waves each need a destination sheet.
Public Function Core_NewSheet(ByVal destName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.name = destName
    Set Core_NewSheet = ws
End Function

' Any string reduced to something Excel will accept as a sheet name: forbidden
' characters replaced with "-", trimmed, cut to 31. It does NOT make the name
' unique - a clash is Core_DestUsable's job, and is refused rather than resolved.
Public Function Core_SafeSheetName(ByVal raw As String) As String
    Dim nm As String, bad As String, i As Long
    nm = raw
    bad = Core_BadSheetChars()
    For i = 1 To Len(bad)
        nm = Replace(nm, Mid$(bad, i, 1), "-")
    Next i
    nm = Trim(nm)
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    Core_SafeSheetName = nm
End Function
```

- [ ] **Step 4: Add the folder helpers**

```vb
' Blank means the folder this workbook is saved in - contract section 9's "data
' location", and the answer in the overwhelming majority of cases. Returns a path
' ending in a separator, or "" when the workbook has never been saved and no folder
' was configured (the caller reports that case by name).
Public Function Core_ResolveFolder(ByVal raw As String) As String
    Dim p As String
    p = Trim(raw)
    If Len(p) = 0 Then p = ThisWorkbook.Path
    If Len(p) = 0 Then Exit Function                 ' workbook never saved
    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    Core_ResolveFolder = p
End Function

' Folder existence, via FileSystemObject rather than Dir(path, vbDirectory).
' Dir's behaviour with a trailing separator is not something to bet a refusal dialog
' on, and it also returns a non-empty result for a FILE, so a path pointing at a
' workbook would read as a valid folder. Late-bound: no reference to add, and no
' VBIDE trust dependency (CLAUDE.md rule 3).
Public Function Core_FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object, p As String
    p = folderPath
    If Len(p) > 3 Then
        If Right$(p, 1) = Application.PathSeparator Then p = Left$(p, Len(p) - 1)
    End If
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then Core_FolderExists = fso.FolderExists(p)
    On Error GoTo 0
End Function
```

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 3 package(s)`, no diagnostics. `Function` is now up by 14 from the
phase-1 baseline and `Sub` unchanged; nothing has been deleted yet, so both packages still
hold their own copies and **the build passing here is the proof the new names collide with
nothing.**

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): value, range, sheet and folder helpers the waves would each re-invent"
```

---

### Task 3: `PkgClean` moves onto the Core API

**Owner: delegate to codex (`sonnet`).** Mechanical: eight deletions and their call sites.

**Files:** Modify `src/packages/PkgClean.bas`.

**Interfaces:** Consumes every helper from Tasks 1 and 2.

- [ ] **Step 1: Rename the call sites**

Across the whole file, with no other change:

| Old call | New call |
|---|---|
| `AskRange(` | `Core_AskRange(` |
| `ParseCols(` | `Core_ParseCols(` |
| `CleanText(` | `Core_CleanText(` |
| `IsBlankVal(` | `Core_IsBlank(` |
| `PkgClean_NarrowToUsed(` | `Core_NarrowToUsed(` |
| `PkgClean_IsAlpha(` | `Core_IsAlpha(` |

`AskHasHeader` and `AskKeyCols` gain arguments rather than only a name, so they are done by
hand:

```vb
' was:  If Not AskHasHeader(rng, hasHeader) Then Exit Sub
If Not Core_AskHasHeader(rng, hasHeader, "Remove Duplicates") Then Exit Sub
```

```vb
' was:  If Not AskKeyCols(rng, keys) Then Exit Sub
If Not Core_AskCols(rng, keys, _
        "Which column(s) decide whether two rows are duplicates?", _
        "Remove Duplicates") Then Exit Sub
```

**Passing `"Remove Duplicates"` explicitly is the point** — the string that used to be
hardcoded inside the helper is now supplied by the one action it was ever right for.

- [ ] **Step 2: Delete the eight now-dead privates**

`AskRange` (603), `AskHasHeader` (626), `AskKeyCols` (645), `ParseCols` (676), `CleanText`
(692), `IsBlankVal` (696), `PkgClean_NarrowToUsed` (582), `PkgClean_IsAlpha` (496).

**Their explanatory comments moved into Core with them in Tasks 1 and 2 — do not leave an
orphaned comment block behind, and do not copy a comment into both places.**

- [ ] **Step 3: Prove nothing still references the old names**

```bash
grep -nE '(^|[^_[:alnum:]])(AskRange|AskHasHeader|AskKeyCols|ParseCols|CleanText|IsBlankVal|PkgClean_NarrowToUsed|PkgClean_IsAlpha)\b' src/packages/*.bas
```

Expected: **no output.** The `[^_[:alnum:]]` guard is what stops `Core_ParseCols` matching
`ParseCols` and reading as a leftover.

- [ ] **Step 4: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics. `Function` falls by 8 from Task 2's build.

- [ ] **Step 5: Commit**

```bash
git add src/packages/PkgClean.bas
git commit -m "refactor(clean): use the Core helpers instead of package-private copies"
```

---

### Task 4: `PkgImport` moves onto the Core API

**Owner: delegate to codex (`sonnet`).** Mechanical, but Step 2 collapses a real duplication.

**Files:** Modify `src/packages/PkgImport.bas`.

**Interfaces:** Consumes `Core_DestUsable`, `Core_NewSheet`, `Core_SafeSheetName`,
`Core_ResolveFolder`, `Core_FolderExists`, `Core_BadSheetChars`.

- [ ] **Step 1: Rename the call sites**

| Old call | New call |
|---|---|
| `PkgImport_DestUsable(ctx, destName, TITLE)` | `Core_DestUsable(ctx, destName, TITLE)` |
| `PkgImport_ResolveFolder(` | `Core_ResolveFolder(` |
| `PkgImport_FolderExists(` | `Core_FolderExists(` |

- [ ] **Step 2: Collapse both copies of the sheet-creation recipe**

At `PkgImport.bas:192-194` (inside `PkgImport_CombineEngine`) and again at `:415-417` (inside
`PkgImport_ImportDelimited`), replace:

```vb
    Set dest = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    dest.name = destName
```

with:

```vb
    Set dest = Core_NewSheet(destName)
```

- [ ] **Step 3: Shrink `PkgImport_SheetNameFor` onto `Core_SafeSheetName`**

It keeps its name and stays private — dropping a file extension is import's own business —
but its forbidden-character list, the second copy in this file, goes:

```vb
' A sheet name from a file name: drop the extension, then hand the rest to Core to
' sanitise and cut to 31. It does NOT make the name unique - a clash is reported by
' Core_DestUsable and refused, exactly like a combine's.
Private Function PkgImport_SheetNameFor(ByVal fullPath As String) As String
    Dim nm As String
    nm = Core_FileName(fullPath)
    If InStrRev(nm, ".") > 1 Then nm = Left$(nm, InStrRev(nm, ".") - 1)
    PkgImport_SheetNameFor = Core_SafeSheetName(nm)
End Function
```

- [ ] **Step 4: Delete the three now-dead privates**

`PkgImport_DestUsable` (451), `PkgImport_ResolveFolder` (571), `PkgImport_FolderExists` (585).

**`PkgImport_WorkbookPaths` and `PkgImport_AlreadyOpen` stay.** They are not promoted, by the
decision recorded above — do not delete them and do not move them.

- [ ] **Step 5: Prove nothing still references the deleted names**

```bash
grep -n "PkgImport_DestUsable\|PkgImport_ResolveFolder\|PkgImport_FolderExists" src/packages/*.bas
grep -n 'Worksheets.Add' src/packages/*.bas
```

Expected: nothing from the first. The second must return **no lines in `PkgImport.bas`** —
both recipes are gone.

- [ ] **Step 6: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 3 package(s)`, no diagnostics. `Function` falls by 3 more.

- [ ] **Step 7: Commit**

```bash
git add src/packages/PkgImport.bas
git commit -m "refactor(import): use the Core sheet and folder helpers, and stop duplicating the sheet-creation recipe"
```

---

### Task 5: Every package private carries its package prefix

**Owner: delegate to codex (`sonnet`).** Pure rename, no logic.

**Files:** Modify `src/packages/PkgImportExport.bas`.

**Interfaces:** Produces nothing. This is the last of the unprefixed names, so Task 6's guard
can be switched on.

- [ ] **Step 1: Rename all five privates and their call sites**

| Old | New |
|---|---|
| `InferTypeMap` (`:503`) | `PkgImportExport_InferTypeMap` |
| `InferColumn` (`:530`) | `PkgImportExport_InferColumn` |
| `LooksLikeId` (`:579`) | `PkgImportExport_LooksLikeId` |
| `ClearTabs` (`:589`) | `PkgImportExport_ClearTabs` |
| `CollectionToArray` (`:609`) | `PkgImportExport_CollectionToArray` |

**Not promoted to Core**, deliberately: the four inference helpers are that package's own
type-guessing, and `CollectionToArray` has one caller. `Core_ParseCols` already covers the
one shape another package would want.

- [ ] **Step 2: Prove no unprefixed private remains anywhere**

```bash
grep -nE '^\s*Private\s+(Sub|Function)\s+' src/packages/*.bas | grep -vE 'Private\s+(Sub|Function)\s+Pkg'
```

Expected: **no output.** Every private in every package file now begins `Pkg`.

- [ ] **Step 3: Build, syntax-check, commit**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add src/packages/PkgImportExport.bas
git commit -m "refactor(importexport): prefix every package private, so the flat namespace stays claimable"
```

---

### Task 6: A build guard, so this cannot regress

**Owner: Opus writes the guard** (it parses source and must not false-positive);
**delegate the injection test.**

**Files:** Modify `build/build.ps1` — a new guard beside the duplicate-procedure guard
(~line 482).

**Interfaces:** Consumes nothing. Produces a build failure on an unprefixed package private.

**This is what replaces a gate item.** "Did anyone add an unprefixed private?" is a question
the build can answer on every run, forever. A checklist answers it once, if someone remembers.

- [ ] **Step 1: Add the guard**

```powershell
# ---- guard: a package private must carry its package prefix -------------------
# The bundle is ONE module, so an unprefixed Private in a package reserves that
# name against every other package - and the duplicate-procedure guard below then
# fails a build for a collision whose cause is invisible in the failing file. This
# says so at the point the name is introduced instead. Core is exempt: it IS the
# shared namespace, and its privates are the ones packages must not shadow.
$prefixProblems = New-Object System.Collections.ArrayList
foreach ($f in $pkgFiles) {
    $leaf = Split-Path $f -Leaf
    if ($leaf -notmatch '^(Pkg[A-Za-z0-9]+)\.bas$') { continue }
    $expected = $matches[1] + '_'
    $lines = Get-Content -LiteralPath $f
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $code = Get-VbaCodePortion $lines[$i]
        if ($code -match '(?i)^\s*Private\s+(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)') {
            $nm = $matches[1]
            if (-not $nm.StartsWith($expected, [System.StringComparison]::Ordinal)) {
                [void]$prefixProblems.Add("${leaf}:$($i + 1): private procedure '$nm' does not start with '$expected'. The bundle is one flat namespace, so an unprefixed private reserves that name against every other package. Rename it '$expected$nm', or promote it to Core if another package needs it.")
            }
        }
    }
}
if ($prefixProblems.Count -gt 0) {
    Write-Host "BUILD FAILED - unprefixed package private(s):" -ForegroundColor Red
    $prefixProblems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}
```

**`StartsWith(..., Ordinal)` and not `-like`** — the do-not-revert list requires the gate's
comparers to be ordinal. **`Get-VbaCodePortion` is used** so a commented-out `Private Sub` in
a package cannot fail a build.

- [ ] **Step 2: Verify it passes on the current tree**

```bash
pwsh build/build.ps1
```

Expected: exit 0. Task 5 removed the last unprefixed private, so a failure here means Task 5
is incomplete — fix that rather than loosening the guard.

- [ ] **Step 3: Deliberate injection — prove the guard actually fires**

A guard that has never failed is not known to work.

```bash
printf '\nPrivate Function TotallyUnprefixed() As Boolean\nEnd Function\n' >> src/packages/PkgClean.bas
pwsh build/build.ps1; echo "exit=$?"
git checkout -- src/packages/PkgClean.bas
pwsh build/build.ps1; echo "exit=$?"
```

Expected: **exit 1** the first time, naming `PkgClean.bas` and `TotallyUnprefixed`; **exit 0**
after the revert. **Record both in the PR body** — this is the phase's only "meant to fail"
evidence.

- [ ] **Step 4: Commit**

```bash
git add build/build.ps1
git commit -m "build: fail an unprefixed package private, so the flat namespace stays claimable"
```

---

### Task 7: Version, changelog, map, and the PR

**Owner: Opus.**

**Files:** `VERSION`, `CHANGELOG.md`, `docs/code-map.md`, `CLAUDE.md`.

- [ ] **Step 1: State the prefix rule where it will be read**

`CLAUDE.md` rule 4 says "Shared helpers live ONLY in Core" but does not say a package private
must be prefixed. Add one clause, since the build now enforces it:

> **Naming is mandatory — flat namespace.** Entry points `Pkg<Name>_<Verb>`; manifests
> `Pkg<Name>_Manifest`; Core's API `Core_*`; **every package private prefixed
> `Pkg<Name>_` too, enforced by `build.ps1`** — the bundle is one module, so an
> unprefixed private reserves that name against every other package. Everything else
> `Private`. Shared helpers live ONLY in Core.

- [ ] **Step 2: `VERSION` → `0.12.0`, and a `## [0.12.0]` changelog entry**

House style: prose about what changed for the user. **The honest lead here is that almost
nothing did** — one dialog stops being captioned "Remove Duplicates" when it is not running
Remove Duplicates, and the rest is groundwork. Say that rather than dressing a refactor up.

- [ ] **Step 3: Refresh the code map**

```bash
node scripts/code-map.mjs
```

Note: run before committing and the map stamps the *previous* commit, so a fresh map can look
one commit stale. Regenerate, commit, and only regenerate again if the stamp matters.

- [ ] **Step 4: Final gates**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas
grep -nE '^\s*Private\s+(Sub|Function)\s+' src/packages/*.bas | grep -vE 'Private\s+(Sub|Function)\s+Pkg'
git status --short
```

Expected: exit 0; no diagnostics; `1`; no output from the prefix grep; a tree clean but for
the untracked `.vscode/`.

- [ ] **Step 5: Review the diff, then one PR against `main`**

Not stacked on the phase-1 branch. Branch from `main` after phase 1 has merged.

**The PR body must state, in one line, that no action's behaviour changed except
`Core_AskHasHeader`'s dialog title** — a reviewer reading a 12-helper diff needs to know what
they are checking *for*, and "nothing" is the claim being made.

---

## Unknowns and risks

- **This phase's risk is silent behaviour drift, not breakage.** A helper that moved and lost
  a line still builds. The mitigation is that every promotion in Tasks 1 and 2 is a verbatim
  move except the three noted title/prompt parameters — **diff the moved bodies against the
  originals** rather than reading them fresh.
- **`Core_AskCols`'s message now says "Run <title> again"** where the original said "Run
  Remove Duplicates again". For the one existing caller these are the same string. For a new
  caller the old text would have been wrong, which is the fix.
- **The prefix guard's regex does not follow line continuations.** A `Private Function` split
  across a `_` is not matched and would slip through. No such declaration exists in this
  codebase and none is plausible, but it is a known limit rather than a covered case.
- **Not in scope, still parked:** `build.ps1`'s `Get-VbaCodePortion` string-literal blind spot
  (it assumes a string opens and closes on one physical line), the guard-scope inconsistency
  between `$asciiCandidates` and `$sources`, and `Core_SheetExists` not seeing chart sheets.
