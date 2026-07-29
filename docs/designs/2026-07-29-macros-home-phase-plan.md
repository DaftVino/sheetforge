# Macros Home phase — implementation plan

> **For agentic workers:** implement this plan task by task. Steps use checkbox (`- [ ]`)
> syntax for tracking. Every task ends with a build, a syntax check, and a commit. **Do not
> batch tasks.** Each one is reviewed before the next begins.

**Goal:** Take the toolkit from 43 actions to 51, restore the withdrawn `Macros Home` index,
finish the Setup surface with the six corrections the owner raised on 2026-07-29, and ship the
result as a public versioned release with a guide a non-technical user can follow.

**Architecture:** Eight new actions land in four existing packages, so no package registration
changes and no new action domain. `Macros Home` is restored from `src/attic/HomeIndex.bas.txt`
and extended to list every registered action with its availability state. The Setup sheet gains
Excel notes on four headers, a formula-driven Finalize reminder, comma-separated chain ids, and
a `Complete Macro Setup` button that strips the configuration tabs once a workbook is finished.

**Design:** `docs/designs/2026-07-29-macros-home-phase-design.md`. Read it before Task 1.
Anything this plan does not answer, that document does.

**Tech stack:** VBA (Excel/Windows), built by `build/build.ps1` into a single-module
`dist/MacrosBundle.bas`; `scripts/vba-syntax-check.mjs` (Node) as a second gate;
`spec/contract.md` as the language-neutral contract the later Apps Script port is written from.

---

## Global constraints

Every task's requirements implicitly include this section. Values are verbatim from
`CLAUDE.md`, `spec/contract.md` and the plan of record.

- **`SCHEMA_VER` stays `"v2.0"`.** If a build ever reports `v3.0`, something is wrong. Stop.
- **Action ids are immutable**, two-segment `<domain>.<operation>`, lowercase ASCII and
  underscores, each segment starting with a letter. **Never camelCase. No alias table.**
- **`build.ps1` parses the domain vocabulary table out of contract §7** and fails on an unknown
  domain. This phase adds no domain — all eight ids use `clean`, `format`, `export`,
  `workbook`, already `active`. **If you find yourself adding a domain, stop: you have the
  wrong id.**
- **Shared helpers live ONLY in Core. Cross-package calls are forbidden.** A package may call
  `Core_*` and its own `Pkg<Name>_*`, nothing else.
- **Every package private must be prefixed `Pkg<Name>_`.** `build.ps1` fails the build
  otherwise. The bundle is one module, so an unprefixed private reserves that name against
  every other package.
- **Every action entry point must be `Public Sub <Name>(ctx As Object)`** — one `Object`
  parameter, nothing else. `build.ps1` fails the build otherwise, because `Core_Run` dispatches
  through `Application.Run`.
- **Pure ASCII in every `.bas`.** `Option Explicit` everywhere. Maximum 25 line continuations
  per statement — one `Collection.Add` per line, never a giant `Array(...)`.
- **No module-level declaration after the first procedure** *in the assembled bundle*, so a
  package's `Const`/`Dim` must sit above its own first procedure.
- **No VBA reserved word as an identifier.** `eNum` lower-cases to `enum` and the build rejects
  it — this happened once already on track 2.
- **VBA `And` / `Or` do not short-circuit.** Both sides always evaluate. Nest the tests.
- **`On Error GoTo 0` inside a procedure that has a `Fail:` handler disables that handler.**
  Both of these caused review findings across four wave plans.
- **Wrap every action in the speed pattern** (ScreenUpdating / Calculation / EnableEvents off →
  on) with a guaranteed restore in the error handler. **Quote every path** — folders have
  spaces and capitals.
- **TDD is off.** VBA has no headless runner and Excel cannot be driven on CI. **Do not propose
  or write a test suite.** Verification is `pwsh build/build.ps1`,
  `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`, and deliberate injections for
  anything meant to fail.
- **`dist/` is gitignored — never `git add -f` the bundle.** `.vscode/` stays untracked,
  deliberately. Do not commit it and do not delete it.
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit message or PR body.
  End the message at its last substantive line.
- **A new package needs registering in TWO places** — the file in `src/packages/`, and
  `c.Add "<Name>"` in `CoreCandidates()`. **This phase adds no package**, so if you are editing
  `CoreCandidates()` you have gone wrong.

### The verification cycle every task ends with

```bash
pwsh build/build.ps1
# expect: exit 0, "Core + 10 package(s)", Sub n/n and Function n/n balanced

node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
# expect: "OK    dist/MacrosBundle.bas - no diagnostics"
```

A task is not done until both pass. Neither proves the macro works — Excel has not run. Say
"reviewed", never "verified", about run-time behaviour.

---

## File map

**Modified in this phase:**

| File | Responsibility | Tasks |
|---|---|---|
| `spec/contract.md` | The language-neutral contract. §2, §6, §7, §8, §9, §10. | 1, 3, 4, 7 |
| `src/packages/PkgClean.bas` | `clean.*` actions. +2 actions. | 2a |
| `src/packages/PkgFormat.bas` | `format.*` actions. +2 actions. | 2b |
| `src/packages/PkgExport.bas` | `export.*` actions. +2 actions. | 2c |
| `src/packages/PkgWorkbook.bas` | `workbook.*` actions. +2 actions. | 2d |
| `src/packages/PkgImportExport.bas` | Four caption strings lose their numbering. | 4 |
| `src/core/Core.bas` | Chains parsing, `Macros Home`, the reminder, `Complete Macro Setup`, header notes, two helpers made `Private`. | 3, 5, 6, 7, 8 |
| `docs/designs/2026-07-28-end-to-end-validation.md` | The owner-facing checklist. Appended, never rewritten. | 4, 9 |
| `README.md` and `docs/*` | Brought current. | 10 |
| `CHANGELOG.md`, `VERSION` | 0.18.0. | 10 |

**Created in this phase:**

| File | Responsibility | Task |
|---|---|---|
| `docs/setup-guide.md` | The non-technical walkthrough. Source for the PDF. | 11 |
| `docs/setup-guide.pdf` | Generated from the above. | 11 |
| `scripts/reading-level.mjs` | Measures Flesch–Kincaid grade so the 6th-grade rule is enforced, not asserted. | 11 |

**Deleted in this phase:**

| File | Why | Task |
|---|---|---|
| `src/attic/HomeIndex.bas.txt` | Its contents move into `Core.bas`. The attic exists to preserve unbuilt code; once built, a second copy is a divergence risk. | 5 |

---

## Task 1: Contract v2.10.0 — freeze the ids before any code references them

**Files:**
- Modify: `spec/contract.md` — §2 (chain step delimiter), §6 (index surface), §7 (the
  `Macros Home` sweep clause), §8 (eight new ids), §10 (change record).

**Interfaces:**
- Consumes: nothing. This is the first task.
- Produces: the eight frozen action ids every later task uses verbatim —
  `clean.split_column`, `clean.change_case`, `format.highlight_special`, `format.chart_style`,
  `export.split_by_value`, `export.email`, `workbook.sort_sheets`, `workbook.toc_links`.

**Why this is first.** `build.ps1` extracts literal action-id declarations from source and
fails on an unknown domain or a malformed id. Writing the actions first and the contract
second means the build gate is judging code against a contract that does not yet describe it.
Freezing the ids here also means Tasks 2a–2d can be worked in any order without inventing
names.

- [ ] **Step 1: Confirm no new domain is needed**

```bash
grep -n '^| `clean`\|^| `format`\|^| `export`\|^| `workbook`' spec/contract.md
```

Expected: four rows, each with status `active`. If any is missing or not `active`, **stop** —
the design assumed otherwise and the plan needs revising before anything else happens.

- [ ] **Step 2: Add the eight ids to §8, the frozen registry**

Each row follows the shape already used by `workbook.build_toc` (`spec/contract.md:1019`):
`` | `<id>` | <Display Caption> | <normative semantics> | ``. Write the semantics from the
design's §2.2 and §2.3 for `export.email` and `workbook.toc_links`; for the other six, state
what the action does, what it refuses, and what it must not touch.

The two that carry settled behaviour and must say so normatively:

- `export.email` — MUST write the attachment file before contacting the mail client; MUST open
  a draft for review and MUST NOT send; when the mail client cannot be reached it MUST refuse
  with a message naming the file it already wrote, and MUST NOT delete that file.
- `workbook.toc_links` — MUST write into exactly one cell per sheet, whose address comes from a
  `text`-kind config field; MUST skip any sheet whose target cell is non-empty and MUST report
  every skipped sheet by name; MUST NOT insert or shift rows; MUST exclude engine-owned sheets
  and the table-of-contents sheet itself.

- [ ] **Step 3: Rewrite §6's index-surface clause**

The current text (`spec/contract.md:556-574`) says a platform MAY generate an index, that the
Excel reference implementation withdrew one, and that the button catalog covers only actions
that became buttons. Two changes:

1. The reference implementation generates an index again.
2. **The catalog covers every registered action with its availability state**, not only those
   placed as buttons. State the reason normatively, because it is the whole argument: every
   action arrives `off` (§6, v2.1.0), so a button-only catalog is empty on a fresh import — at
   exactly the moment a reader most needs to know what the toolkit does.

Keep unchanged: page-to-page navigation stays required and unconditional; only the index itself
and the per-page link back to it are optional.

- [ ] **Step 4: Rewrite §7's `Macros Home` sweep clause**

`spec/contract.md:671-679` says the engine *"deletes `Macros Home` unconditionally with alerts
suppressed on every rebuild"*. That describes the withdrawal and is now false. The name stays
**reserved**; it is now **generated** on every rebuild rather than swept. `Macros` and
`Macros <n>` keep their existing unconditional-delete rule — do not touch that sentence.

- [ ] **Step 5: Amend §2's chain-step delimiter note**

`spec/contract.md:218-220` reads *"Stored as an id-list joined by a delimiter that cannot occur
in an id (reference: newline)"*. Change `(reference: newline)` to `(reference: comma)`. The
normative requirement — an ordered list of action ids, not captions, joined by a delimiter that
cannot occur in an id — does not change, and a comma cannot occur in an id.

- [ ] **Step 6: Write the §10 change record for v2.10.0**

Additive, not breaking. It must say why the withdrawal is being reversed rather than merely
that it is: §6's reasoning was *"at the scale a real deployment runs — a handful of actions and
one or two chains — an index is clutter rather than navigation."* That was written against
seven actions. At 51 it no longer holds. A contract that reverses itself without saying why
teaches the next reader that its reasoning is decorative.

Also record: the four retitled `importexport` captions (Task 4) are display strings, not ids,
so retitling is additive; `SCHEMA_VER` does not move.

- [ ] **Step 7: Bump the version header**

`spec/contract.md` currently declares v2.9.0. Set it to **v2.10.0**.

- [ ] **Step 8: Verify the build gate still parses the vocabulary table**

```bash
pwsh build/build.ps1
```

Expected: exit 0, `Core + 10 package(s)`. Nothing in source changed, so this is a pure
regression check on the table parser — §7's table is machine-read, and Step 1 through Step 7
edited the document around it.

- [ ] **Step 9: Deliberate injection — prove the id gate is live**

Temporarily add `Core_AddAction reg, "notadomain.thing", "X", "PkgWorkbook_UnhideSheets", "x", "Workbook Housekeeping"`
to `src/packages/PkgWorkbook.bas`, run `pwsh build/build.ps1`, and confirm it **fails** naming
the unknown domain. Then revert the line. Record the observed failure text in the PR body.

Doing this now, on a contract that has just been edited, is the point: it proves the parser
still finds the table after the edits.

- [ ] **Step 10: Commit**

```bash
git add spec/contract.md
git commit -m "spec: contract v2.10.0 - restore the index, freeze eight ids, comma-delimit chain steps"
```

---

## Task 2a: PkgClean - Split One Column Into Many and Change Case

**Files:**
- Modify: `src/packages/PkgClean.bas`

**Interfaces:**
- Consumes: the frozen ids from Task 1.
- Produces: `PkgClean_SplitColumn(ctx As Object)` and `PkgClean_ChangeCase(ctx As Object)`, plus `PkgClean_Delimiter`, `PkgClean_DestinationColumn`, `PkgClean_FieldCount`, and `PkgClean_AskCase` privates.

**New config fields:** `k_clean_split_delimiter`, label `Split-column delimiter`, kind `text`, default `COMMA`, tier `simple`, usedBy `clean.split_column`; `k_clean_split_destination`, label `Split-column destination`, kind `text`, default `SOURCE`, tier `simple`, usedBy `clean.split_column`. Change Case has no config field because the casing choice describes the range being acted on and is requested at run time.

- [ ] **Step 1: Declare the config fields in the manifest**

Place these declarations after the existing `k_clean_formats` table and before every `Core_AddAction` call:

```vba
    Core_AddSection reg, "Split Column"
    Core_AddField reg, "k_clean_split_delimiter", "Split-column delimiter", "text", _
        "COMMA", "simple", _
        "Enter one delimiter character, or use COMMA, TAB, SEMICOLON, PIPE, or SPACE.", _
        "clean.split_column"
    Core_AddField reg, "k_clean_split_destination", "Split-column destination", "text", _
        "SOURCE", "simple", _
        "Use SOURCE to replace the picked column, or enter an absolute worksheet column letter such as D. A separate destination leaves the source column unchanged.", _
        "clean.split_column"
```

- [ ] **Step 2: Declare the actions in the manifest**

Add these after all field declarations, alongside the existing action declarations:

```vba
    Core_AddAction reg, "clean.split_column", "Split One Column Into Many", _
        "PkgClean_SplitColumn", "For exports that land everything in column A"
    Core_AddAction reg, "clean.change_case", "Change Case", _
        "PkgClean_ChangeCase", "Normalize text casing across a range"
```

- [ ] **Step 3: Write Split One Column Into Many**

The action accepts exactly one selected column. The delimiter and first output column come from Setup. `SOURCE` replaces the selected column; an absolute column letter writes elsewhere and leaves the source unchanged. Every output field is forced to text so identifiers such as `00123` are not coerced. Before changing anything, the action refuses formulas, invalid configuration, a delimiter that matches no selected value, an output that would run past the last worksheet column, or any occupied destination cell that would be overwritten.

```vba
Public Sub PkgClean_SplitColumn(ctx As Object)
    Dim src As Range, ws As Worksheet, destCell As Range, target As Range
    Dim c As Range, delimiter As String, destSpec As String, firstOccupied As String
    Dim sourceRef As String, destNote As String
    Dim destCol As Long, maxFields As Long, fields As Long, i As Long
    Dim formulaState As Variant, fieldInfo() As Variant
    Const TITLE As String = "Split One Column Into Many"
    On Error GoTo Fail

    Set src = Core_AskRange(TITLE)
    If src Is Nothing Then Exit Sub
    If src.Columns.Count <> 1 Then
        Core_Refuse ctx, "This action accepts exactly one selected column." & vbCrLf & vbCrLf & _
               "No Setup change is needed; run it again and pick one column only." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set src = Core_NarrowToUsed(src)
    If src Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to split." & vbCrLf & vbCrLf & _
               "No Setup change is needed; run it again and pick a populated column." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If
    Set ws = src.Parent
    sourceRef = Core_RangeRef(src)

    delimiter = PkgClean_Delimiter(Core_Cfg(ctx, "k_clean_split_delimiter"))
    If Len(delimiter) = 0 Then
        Core_Refuse ctx, "The split-column delimiter is not valid." & vbCrLf & vbCrLf & _
               "Set 'Split-column delimiter' on the Setup tab to one character, or to " & _
               "COMMA, TAB, SEMICOLON, PIPE, or SPACE, Finalize, and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    destSpec = Trim$(Core_Cfg(ctx, "k_clean_split_destination"))
    destCol = PkgClean_DestinationColumn(ws, destSpec, src.Column)
    If destCol = 0 Then
        Core_Refuse ctx, "The split-column destination is not valid." & vbCrLf & vbCrLf & _
               "Set 'Split-column destination' on the Setup tab to SOURCE or an absolute " & _
               "worksheet column letter, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    formulaState = src.HasFormula
    If IsNull(formulaState) Then
        Core_Refuse ctx, "The selected column contains one or more formulas. Splitting it " & _
               "would replace those formulas with text values." & vbCrLf & vbCrLf & _
               "No Setup change is needed; pick a values-only export column and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    Else
        If CBool(formulaState) Then
            Core_Refuse ctx, "The selected column contains one or more formulas. Splitting " & _
                   "it would replace those formulas with text values." & vbCrLf & vbCrLf & _
                   "No Setup change is needed; pick a values-only export column and run this " & _
                   "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
            Exit Sub
        End If
    End If

    For Each c In src.Cells
        fields = PkgClean_FieldCount(c.Value2, delimiter)
        If fields > maxFields Then maxFields = fields
    Next c
    If maxFields <= 1 Then
        Core_Refuse ctx, "The configured delimiter was not found in the selected column." & _
               vbCrLf & vbCrLf & "Check 'Split-column delimiter' on the Setup tab, " & _
               "Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If destCol + maxFields - 1 > ws.Columns.Count Then
        Core_Refuse ctx, "The split needs " & maxFields & " destination columns, but they " & _
               "would run past the last column on the worksheet." & vbCrLf & vbCrLf & _
               "Set 'Split-column destination' to an earlier column on the Setup tab, " & _
               "Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set destCell = ws.Cells(src.Row, destCol)
    Set target = ws.Range(destCell, _
        ws.Cells(src.Row + src.Rows.Count - 1, destCol + maxFields - 1))

    For Each c In target.Cells
        If destCol = src.Column Then
            If c.Column <> src.Column Then
                If Application.CountA(c) > 0 Then
                    firstOccupied = c.Address(False, False)
                    Exit For
                End If
            End If
        Else
            If Application.CountA(c) > 0 Then
                firstOccupied = c.Address(False, False)
                Exit For
            End If
        End If
    Next c

    If Len(firstOccupied) > 0 Then
        Core_Refuse ctx, "Splitting would overwrite existing content or a formula at " & _
               firstOccupied & "." & vbCrLf & vbCrLf & _
               "Clear the intended output area, or change 'Split-column destination' on " & _
               "the Setup tab, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ReDim fieldInfo(0 To maxFields - 1)
    For i = 0 To maxFields - 1
        fieldInfo(i) = Array(i + 1, xlTextFormat)
    Next i

    Core_SpeedOn
    src.TextToColumns Destination:=destCell, DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=True, OtherChar:=delimiter, FieldInfo:=fieldInfo
    Core_SpeedOff

    If destCol = src.Column Then
        destNote = "The selected source column was replaced by the split fields."
    Else
        destNote = "The source column was left unchanged."
    End If
    Core_Done ctx, "Split " & src.Rows.Count & " row(s) from " & sourceRef & _
           " into as many as " & maxFields & " text column(s), starting at " & _
           destCell.Address(False, False) & "." & vbCrLf & vbCrLf & destNote, TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "clean.split_column", Err
End Sub

Private Function PkgClean_Delimiter(ByVal raw As String) As String
    Dim token As String
    token = UCase$(Trim$(raw))

    Select Case token
        Case "COMMA"
            PkgClean_Delimiter = ","
        Case "TAB"
            PkgClean_Delimiter = vbTab
        Case "SEMICOLON"
            PkgClean_Delimiter = ";"
        Case "PIPE"
            PkgClean_Delimiter = "|"
        Case "SPACE"
            PkgClean_Delimiter = " "
        Case Else
            If Len(raw) = 1 Then PkgClean_Delimiter = raw
    End Select
End Function

Private Function PkgClean_DestinationColumn(ws As Worksheet, ByVal spec As String, _
                                            ByVal sourceColumn As Long) As Long
    Dim token As String, i As Long, colNum As Long
    token = UCase$(Trim$(spec))

    If token = "SOURCE" Then
        PkgClean_DestinationColumn = sourceColumn
        Exit Function
    End If
    If Len(token) = 0 Then Exit Function
    If Len(token) > 3 Then Exit Function
    If Not Core_IsAlpha(token) Then Exit Function

    For i = 1 To Len(token)
        colNum = colNum * 26 + Asc(Mid$(token, i, 1)) - Asc("A") + 1
    Next i
    If colNum <= ws.Columns.Count Then PkgClean_DestinationColumn = colNum
End Function

Private Function PkgClean_FieldCount(ByVal rawValue As Variant, _
                                     ByVal delimiter As String) As Long
    Dim s As String, ch As String, i As Long, n As Long
    Dim inQuotes As Boolean
    n = 1
    PkgClean_FieldCount = n

    If IsError(rawValue) Then Exit Function
    If IsEmpty(rawValue) Then Exit Function
    s = CStr(rawValue)

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If inQuotes Then
                If i < Len(s) Then
                    If Mid$(s, i + 1, 1) = """" Then
                        i = i + 1
                    Else
                        inQuotes = False
                    End If
                Else
                    inQuotes = False
                End If
            Else
                inQuotes = True
            End If
        Else
            If Not inQuotes Then
                If ch = delimiter Then n = n + 1
            End If
        End If
    Next i

    PkgClean_FieldCount = n
End Function
```

- [ ] **Step 4: Write Change Case**

The user picks the mode at run time by entering `UPPER`, `lower`, or `Proper`. Cancel exits silently. Only constant text cells inside the selected used range are eligible; formulas, numbers, errors, and blanks are left unchanged. The completion report states both how many text constants were examined and how many actually changed, and explicitly names formulas as skipped.

```vba
Public Sub PkgClean_ChangeCase(ctx As Object)
    Dim rng As Range, arr As Variant, converted As String, caseMode As String
    Dim r As Long, c As Long, textCount As Long, changed As Long
    Dim formulaCount As Long
    Const TITLE As String = "Change Case"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is no text in it to change." & vbCrLf & vbCrLf & _
               "No Setup change is needed; run it again and pick a populated range." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not PkgClean_AskCase(caseMode, TITLE) Then Exit Sub

    Core_SpeedOn
    arr = rng.Value2
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            For c = 1 To UBound(arr, 2)
                If rng.Cells(r, c).HasFormula Then
                    formulaCount = formulaCount + 1
                ElseIf VarType(arr(r, c)) = vbString Then
                    textCount = textCount + 1
                    Select Case caseMode
                        Case "UPPER"
                            converted = UCase$(CStr(arr(r, c)))
                        Case "lower"
                            converted = LCase$(CStr(arr(r, c)))
                        Case "Proper"
                            converted = StrConv(CStr(arr(r, c)), vbProperCase)
                    End Select
                    If converted <> CStr(arr(r, c)) Then
                        rng.Cells(r, c).Value = converted
                        changed = changed + 1
                    End If
                End If
            Next c
        Next r
    Else
        If rng.HasFormula Then
            formulaCount = 1
        ElseIf VarType(arr) = vbString Then
            textCount = 1
            Select Case caseMode
                Case "UPPER"
                    converted = UCase$(CStr(arr))
                Case "lower"
                    converted = LCase$(CStr(arr))
                Case "Proper"
                    converted = StrConv(CStr(arr), vbProperCase)
            End Select
            If converted <> CStr(arr) Then
                rng.Value = converted
                changed = 1
            End If
        End If
    End If
    Core_SpeedOff

    Core_Done ctx, "Applied " & caseMode & " case to " & textCount & _
           " constant text cell(s) in " & Core_RangeRef(rng) & "; " & changed & _
           " cell(s) changed." & vbCrLf & vbCrLf & _
           "Text already in that case was not rewritten. Numbers, errors, blanks, and " & _
           formulaCount & " formula cell(s) were left exactly as they were.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "clean.change_case", Err
End Sub

Private Function PkgClean_AskCase(ByRef caseMode As String, _
                                  ByVal title As String) As Boolean
    Dim answer As Variant, token As String

    Do
        answer = Application.InputBox( _
            Prompt:="Enter UPPER, lower, or Proper." & vbCrLf & vbCrLf & _
                    "Only constant text cells will be changed.", _
            title:=title, Default:="Proper", Type:=2)

        If VarType(answer) = vbBoolean Then
            If answer = False Then Exit Function
        End If

        token = LCase$(Trim$(CStr(answer)))
        Select Case token
            Case "upper"
                caseMode = "UPPER"
                PkgClean_AskCase = True
                Exit Function
            Case "lower"
                caseMode = "lower"
                PkgClean_AskCase = True
                Exit Function
            Case "proper"
                caseMode = "Proper"
                PkgClean_AskCase = True
                Exit Function
            Case Else
                MsgBox "Enter UPPER, lower, or Proper, or click Cancel to stop.", _
                       vbInformation, title
        End Select
    Loop
End Function
```

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 10 package(s)`, no diagnostics.

- [ ] **Step 6: Confirm both ids registered**

```bash
grep -c 'Core_AddAction reg, "clean.split_column"\|Core_AddAction reg, "clean.change_case"' src/packages/PkgClean.bas
```

Expected: `2`.

- [ ] **Step 7: Append the run-time claims to the validation checklist**

- With delimiter `COMMA`, destination `SOURCE`, and source text `00123,North`, the result is two text cells containing `00123` and `North`; losing the leading zeros, leaving the row unsplit, or writing outside the calculated output footprint is failure. This proves the configured delimiter, source replacement, and text-preserving field information are active.
- With destination `SOURCE` and an occupied cell immediately to the right of a value that needs two fields, the action refuses and both cells retain their original formulas or values; any overwritten or cleared cell is failure. This proves overwrite detection runs before Text to Columns mutates the sheet.
- With destination `D`, the split fields begin in column D while the selected source column remains byte-for-byte unchanged; changing the source or ignoring column D is failure. This proves the absolute destination setting and separate-destination contract.
- Running Change Case with `UPPER`, `lower`, and `Proper` produces the requested casing for constant text while formulas and numeric cells retain their formulas, values, and types; identical results for all three choices or any changed formula or number is failure. This proves the run-time choice is honored and the action does not flatten formulas or coerce non-text cells.

- [ ] **Step 8: Commit**

```bash
git add src/packages/PkgClean.bas docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(clean): split columns and normalize text case"
```

### Review notes on Task 2a — fix these while implementing

Drafted by a codex agent, reviewed here. The structure and the refusal set are sound; these are
the corrections.

1. **The overwrite scan is O(cells) and will crawl on a real sheet.** `For Each c In
   target.Cells` with `Application.CountA(c)` per cell is one COM call per cell — on 10,000 rows
   splitting into five fields that is 50,000 round trips, and this workbook is for finance data.
   Test the block in one call instead. When `destCol = src.Column`, test the block to the right
   of the source column as a single range; otherwise test the whole target block:

   ```vba
   If Application.CountA(scanRange) > 0 Then
   ```

   Only walk cell by cell to *name* the first occupied address, and only after the fast test has
   already said there is one to find.

2. **The duplicated formula refusal.** The `IsNull(formulaState)` branch and the
   `CBool(formulaState)` branch build the same message twice. Collapse to one test — a mixed
   range and an all-formula range get the same answer — and keep one copy of the string.

3. **`PkgClean_AskCase` asks the user to type `UPPER`, `lower` or `Proper` into a text box.**
   That is the weakest part of the draft: this toolkit is for people who do not want to think
   about macros, and free-text entry of a mode string is a typo away from a re-prompt loop.
   Prefer a Setup field — the choice is a property of how the deployment wants its data, not of
   the moment. If it stays a prompt, keep the loop's Cancel path exactly as drafted; it is
   correct.

4. **Confirm the `Core_AddAction` category argument matches the file.** The draft omits it, which
   defaults to the package display name. `PkgWorkbook` passes it explicitly. Read
   `src/packages/PkgClean.bas` and match whatever the other `clean.*` actions already do rather
   than introducing a second convention inside one file.

5. **Verify `TextToColumns` is not affected by the speed pattern.** `Core_SpeedOn` sets
   `Calculation = xlCalculationManual`. Add this to the checklist rather than assuming — the
   action's whole value is that the split fields are text, and a calculation-mode interaction
   that silently re-coerces them would be invisible in the build.

---

## Task 2b: Format - Highlight Comments, Formulas and Blanks and Apply One Style To All Charts

**Files:**
- Modify: `src/packages/PkgFormat.bas`

**Interfaces:**
- Consumes: the frozen ids from Task 1.
- Produces: `PkgFormat_HighlightSpecial(ctx As Object)` and `PkgFormat_ChartStyle(ctx As Object)`, plus `PkgFormat_` privates for identifying and removing action-owned conditional-format rules.
- `format.highlight_special` is one pass over the active report sheet's used range, not a choice prompt. Comments are orange, formulas are blue, and blanks are yellow; a commented formula is orange because comment highlighting has highest priority. The fixed pastel palette makes the audit view consistent without adding configuration.
- The audit view uses tagged conditional-format rules rather than direct fills. Running the action again removes only its own tagged rules, revealing the user's original fills and other conditional formatting unchanged.
- `format.chart_style` means Excel's built-in `ChartStyle` number, from 1 through 48. It changes only `Chart.ChartStyle`; it does not change chart data, chart type, titles, series formulas, size, or position. It processes embedded charts and chart sheets, but never charts on engine-owned sheets.

**New config fields:** `k_format_chart_style`, label `Built-in chart style number`, kind `number`, default `10`, tier `simple`, usedBy `format.chart_style`.

- [ ] **Step 1: Declare the config fields in the manifest**

Place this section after the existing field declarations and before every `Core_AddAction` call:

```vba
    Core_AddSection reg, "CHARTS"
    Core_AddField reg, "k_format_chart_style", "Built-in chart style number", "number", "10", _
        "simple", "Excel built-in ChartStyle number from 1 through 48. The style uses the " & _
        "workbook theme, so its exact colors and fonts follow that theme.", "format.chart_style"
```

- [ ] **Step 2: Declare the actions in the manifest**

Place both calls with the other action declarations, after all field declarations:

```vba
    Core_AddAction reg, "format.highlight_special", _
        "Highlight Comments, Formulas and Blanks", "PkgFormat_HighlightSpecial", _
        "Popular audit-view helper. On the active report sheet, marks comments orange, " & _
        "formulas blue, and blanks yellow with action-owned conditional formatting. Run it " & _
        "again to remove only those audit highlights and restore the previous appearance."
    Core_AddAction reg, "format.chart_style", "Apply One Style To All Charts", _
        "PkgFormat_ChartStyle", _
        "Enforce one visual standard across chart objects. Applies the configured built-in " & _
        "Excel ChartStyle to every embedded chart and chart sheet except charts on " & _
        "engine-owned sheets."
```

- [ ] **Step 3: Write Highlight Comments, Formulas and Blanks**

```vba
Public Sub PkgFormat_HighlightSpecial(ctx As Object)
    Dim ws As Worksheet, rng As Range
    Dim commentCells As Range, formulaCells As Range, blankCells As Range
    Dim fc As FormatCondition
    Dim commentCount As Double, formulaCount As Double, blankCount As Double
    Dim ownedCount As Long, removedCount As Long
    Dim anchor As String
    Const TITLE As String = "Highlight Comments, Formulas and Blanks"
    On Error GoTo Fail

    If TypeName(ActiveSheet) <> "Worksheet" Then
        Core_Refuse ctx, "The active tab is a chart sheet, so it has no cells to audit." & _
               vbCrLf & vbCrLf & "Switch to a report worksheet and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set ws = ActiveSheet
    If Core_IsEngineSheet(ws.name) Then
        Core_Refuse ctx, "'" & ws.name & "' belongs to the macro toolkit itself, not to a " & _
               "report." & vbCrLf & vbCrLf & "Switch to the report tab and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ownedCount = PkgFormat_SpecialRuleCount(ws)
    If ownedCount > 0 Then
        If MsgBox("This sheet already has audit highlights added by this action." & _
                  vbCrLf & vbCrLf & "Remove those " & ownedCount & " conditional-format " & _
                  "rule(s)? The sheet's direct fills and all other conditional formatting " & _
                  "will remain in place.", _
                  vbQuestion + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

        Core_SpeedOn
        removedCount = PkgFormat_RemoveSpecialRules(ws)
        Core_SpeedOff

        Core_Done ctx, "Removed " & removedCount & " audit-highlight rule(s) from '" & _
               ws.name & "'." & vbCrLf & vbCrLf & "The underlying fills and all other " & _
               "conditional formatting were left unchanged.", TITLE
        Exit Sub
    End If

    Set rng = ws.UsedRange

    On Error Resume Next
    Set commentCells = rng.SpecialCells(xlCellTypeComments)
    Set formulaCells = rng.SpecialCells(xlCellTypeFormulas)
    Set blankCells = rng.SpecialCells(xlCellTypeBlanks)
    Err.Clear
    On Error GoTo Fail

    If Not commentCells Is Nothing Then commentCount = commentCells.CountLarge
    If Not formulaCells Is Nothing Then formulaCount = formulaCells.CountLarge
    If Not blankCells Is Nothing Then blankCount = blankCells.CountLarge

    If Application.CountA(rng) = 0 Then
        If commentCells Is Nothing Then
            Core_Done ctx, "'" & ws.name & "' has no used cells or comments to audit." & _
                   vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
            Exit Sub
        End If
    End If

    If commentCount + formulaCount + blankCount = 0 Then
        Core_Done ctx, "No comments, formulas, or blanks were found in the used area of '" & _
               ws.name & "'." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If MsgBox("Add a reversible audit view to the used area of '" & ws.name & "'?" & _
              vbCrLf & vbCrLf & "Comments: orange" & vbCrLf & "Formulas: blue" & vbCrLf & _
              "Blanks: yellow" & vbCrLf & vbCrLf & _
              "This adds tagged conditional-format rules. It does not replace direct fills " & _
              "or other conditional formatting. Run this action again to remove only these " & _
              "audit highlights.", _
              vbQuestion + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    anchor = rng.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)

    Core_SpeedOn

    If Not blankCells Is Nothing Then
        Set fc = rng.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=AND(ISBLANK(" & anchor & ")," & _
                     "N(""SheetForge format.highlight_special blank"")=0)")
        fc.Interior.Color = RGB(255, 235, 156)
        fc.StopIfTrue = True
        fc.SetFirstPriority
    End If

    If Not formulaCells Is Nothing Then
        Set fc = rng.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=AND(ISFORMULA(" & anchor & ")," & _
                     "N(""SheetForge format.highlight_special formula"")=0)")
        fc.Interior.Color = RGB(189, 215, 238)
        fc.StopIfTrue = True
        fc.SetFirstPriority
    End If

    If Not commentCells Is Nothing Then
        Set fc = commentCells.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=N(""SheetForge format.highlight_special comment"")=0")
        fc.Interior.Color = RGB(252, 213, 180)
        fc.StopIfTrue = True
        fc.SetFirstPriority
    End If

    Core_SpeedOff

    Core_Done ctx, "Added reversible audit highlights to '" & ws.name & "': " & _
           Format$(commentCount, "#,##0") & " comment cell(s), " & _
           Format$(formulaCount, "#,##0") & " formula cell(s), and " & _
           Format$(blankCount, "#,##0") & " blank cell(s)." & vbCrLf & vbCrLf & _
           "No value or direct fill changed. Run this action again to remove the audit view.", _
           TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.highlight_special", Err
End Sub

Private Function PkgFormat_SpecialRuleCount(ws As Worksheet) As Long
    Dim i As Long
    Dim fc As FormatCondition
    Const TOKEN As String = "SheetForge format.highlight_special"

    For i = 1 To ws.Cells.FormatConditions.Count
        Set fc = ws.Cells.FormatConditions(i)
        If fc.Type = xlExpression Then
            If InStr(1, CStr(fc.Formula1), TOKEN, vbTextCompare) > 0 Then
                PkgFormat_SpecialRuleCount = PkgFormat_SpecialRuleCount + 1
            End If
        End If
    Next i
End Function

Private Function PkgFormat_RemoveSpecialRules(ws As Worksheet) As Long
    Dim i As Long
    Dim fc As FormatCondition
    Const TOKEN As String = "SheetForge format.highlight_special"

    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        Set fc = ws.Cells.FormatConditions(i)
        If fc.Type = xlExpression Then
            If InStr(1, CStr(fc.Formula1), TOKEN, vbTextCompare) > 0 Then
                fc.Delete
                PkgFormat_RemoveSpecialRules = PkgFormat_RemoveSpecialRules + 1
            End If
        End If
    Next i
End Function
```

- [ ] **Step 4: Write Apply One Style To All Charts**

```vba
Public Sub PkgFormat_ChartStyle(ctx As Object)
    Dim wb As Workbook, ws As Worksheet
    Dim co As ChartObject, chartSheet As Chart
    Dim styleNo As Long, eligibleCount As Long, engineSkipped As Long
    Dim styledCount As Long, failedCount As Long
    Dim errNum As Long, failureText As String
    Const TITLE As String = "Apply One Style To All Charts"
    On Error GoTo Fail

    styleNo = Core_CfgLong(ctx, "k_format_chart_style")
    If styleNo < 1 Or styleNo > 48 Then
        Core_Refuse ctx, "'Built-in chart style number' must be a whole number from 1 " & _
               "through 48." & vbCrLf & vbCrLf & "Fix that field on the Setup tab, " & _
               "Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        If Core_IsEngineSheet(ws.name) Then
            engineSkipped = engineSkipped + ws.ChartObjects.Count
        Else
            eligibleCount = eligibleCount + ws.ChartObjects.Count
        End If
    Next ws

    For Each chartSheet In wb.Charts
        If Core_IsEngineSheet(chartSheet.name) Then
            engineSkipped = engineSkipped + 1
        Else
            eligibleCount = eligibleCount + 1
        End If
    Next chartSheet

    If eligibleCount = 0 Then
        If engineSkipped = 0 Then
            Core_Done ctx, "This workbook has no charts." & vbCrLf & vbCrLf & _
                   "Nothing has been changed.", TITLE
        Else
            Core_Done ctx, "This workbook has no charts on report sheets." & vbCrLf & _
                   vbCrLf & engineSkipped & " chart(s) on engine-owned sheets were left " & _
                   "untouched. Nothing has been changed.", TITLE
        End If
        Exit Sub
    End If

    If MsgBox("Apply Excel built-in chart style " & styleNo & " to " & eligibleCount & _
              " chart(s) in this workbook?" & vbCrLf & vbCrLf & _
              "This replaces each eligible chart's current visual preset. It does not " & _
              "change chart data, chart type, titles, series formulas, size, or position." & _
              vbCrLf & vbCrLf & engineSkipped & " chart(s) on engine-owned sheets will " & _
              "remain untouched.", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn

    For Each ws In wb.Worksheets
        If Not Core_IsEngineSheet(ws.name) Then
            For Each co In ws.ChartObjects
                errNum = 0
                failureText = ""
                Err.Clear
                On Error Resume Next
                co.Chart.ChartStyle = styleNo
                errNum = Err.Number
                failureText = Err.Description
                Err.Clear
                On Error GoTo Fail

                If errNum = 0 Then
                    styledCount = styledCount + 1
                Else
                    failedCount = failedCount + 1
                    Core_Log ctx, "format.chart_style skipped embedded chart '" & co.name & _
                           "' on '" & ws.name & "': " & failureText
                End If
            Next co
        End If
    Next ws

    For Each chartSheet In wb.Charts
        If Not Core_IsEngineSheet(chartSheet.name) Then
            errNum = 0
            failureText = ""
            Err.Clear
            On Error Resume Next
            chartSheet.ChartStyle = styleNo
            errNum = Err.Number
            failureText = Err.Description
            Err.Clear
            On Error GoTo Fail

            If errNum = 0 Then
                styledCount = styledCount + 1
            Else
                failedCount = failedCount + 1
                Core_Log ctx, "format.chart_style skipped chart sheet '" & _
                       chartSheet.name & "': " & failureText
            End If
        End If
    Next chartSheet

    Core_SpeedOff

    If failedCount = 0 Then
        Core_Done ctx, "Applied built-in chart style " & styleNo & " to " & styledCount & _
               " chart(s)." & vbCrLf & vbCrLf & engineSkipped & " chart(s) on " & _
               "engine-owned sheets were left untouched. No chart data, type, title, " & _
               "series formula, size, or position changed.", TITLE
    Else
        Core_Done ctx, "Applied built-in chart style " & styleNo & " to " & styledCount & _
               " chart(s)." & vbCrLf & vbCrLf & "Excel refused " & failedCount & _
               " eligible chart(s); each skipped chart and the reason are named in the " & _
               "Log. " & engineSkipped & " chart(s) on engine-owned sheets were left " & _
               "untouched.", TITLE
    End If
    Exit Sub
Fail:
    Core_Fail ctx, "format.chart_style", Err
End Sub
```

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 10 package(s)`, no diagnostics.

- [ ] **Step 6: Confirm both ids registered**

```bash
grep -c 'Core_AddAction reg, "format.highlight_special"\|Core_AddAction reg, "format.chart_style"' src/packages/PkgFormat.bas
```

Expected: `2`.

- [ ] **Step 7: Append the run-time claims to the validation checklist**

- [ ] A commented formula appears orange, an uncommented formula appears blue, and a blank inside the used area appears yellow; blue winning over orange or any category remaining unmarked is failure. This proves the three classifiers run together and that comment priority is deterministic.
- [ ] Running Highlight Comments, Formulas and Blanks a second time restores a deliberately chosen direct fill and leaves a pre-existing unrelated conditional-format rule intact; a residual audit color, lost fill, or deleted unrelated rule is failure. This proves the action has a durable undo path and does not destroy user formatting.
- [ ] Every embedded chart and chart sheet on report tabs receives the configured built-in style while a deliberately different chart on an engine-owned sheet retains its original style; an eligible chart retaining the old style or an engine chart changing is failure. This proves both chart coverage and the engine-sheet exclusion.
- [ ] In a workbook with no eligible report charts, the action reports that nothing changed and leaves any engine-owned chart untouched; a runtime error, a success count above zero, or a changed engine chart is failure. This proves the empty-workbook path is an intentional no-op rather than an unhandled collection edge case.

- [ ] **Step 8: Commit**

```bash
git add src/packages/PkgFormat.bas docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(format): add reversible audit highlights and consistent chart styles"
```

### Review notes on Task 2b — fix these while implementing

The reversible-highlight design is better than what the catalog asked for: a direct fill would
have destroyed the user's own formatting with no way back, and tagged conditional-format rules
avoid that. Keep it. These are the corrections.

1. **The comment rule is a snapshot pretending to be a rule.** Excel has no `ISCOMMENT`, so the
   draft applies an always-true rule to whichever cells had comments *at the moment it ran*.
   Delete a comment afterwards and the cell stays orange; add one and it never colours. That is
   defensible — there is no alternative — but it must be **said in the completion message and in
   the action's help text**, not left for the user to discover. The formula and blank rules are
   live; the comment rule is not, and the difference is exactly the kind of thing that reads as a
   bug.

2. **`ISFORMULA` requires Excel 2013 or later.** Mark it as an assumption and put it on the
   checklist. Everything else in this codebase works further back.

3. **`ChartStyle` 1-48 is the pre-2007 numbering.** Modern Excel accepts a much wider range and
   the same number can mean different things across versions. **Do not ship the 1-48 validation
   until it has been confirmed in the target Excel** — a refusal that rejects a number Excel
   would have accepted is worse than no validation. Checklist item: set the field to the
   documented maximum, run it, and confirm the charts change rather than the action refusing.

4. **Both actions raise a `MsgBox` confirmation, which will block a chain.** `Core_Done`
   suppresses its dialog when the action runs as a quiet chain step; a raw `MsgBox` does not, so
   a chain containing either action stops and waits for a human who may not be there. See the
   cross-cutting note after Task 2d — this needs the owner's decision before either action is
   written.

5. **`ws.Cells.FormatConditions` as the enumeration source.** Confirm it really enumerates every
   conditional format on the sheet rather than only those whose applies-to covers the entire
   sheet. If it does not, `PkgFormat_RemoveSpecialRules` silently fails to remove some rules and
   the promised undo is not an undo. This is the single most important thing to verify in this
   task, because the reversibility claim is what makes the action safe to run at all.

---

## Task 2c: Export - Save Each Group As Its Own File and Email A Sheet As An Attachment

**Files:**
- Modify: `src/packages/PkgExport.bas`

**Interfaces:**
- Consumes: the frozen ids from Task 1.
- Produces: `PkgExport_SplitByValue(ctx As Object)` and `PkgExport_Email(ctx As Object)`, plus `PkgExport_CloseNoSave`.

**New config fields:**
- `k_export_email_to`: label `Email recipient`, kind `text`, default empty, tier `simple`, usedBy `export.email`
- `k_export_email_subject`: label `Email subject`, kind `text`, default `Sheet export`, tier `simple`, usedBy `export.email`
- `export.split_by_value` adds no field. The user picks a range and chooses exactly one split column at run time.

- [ ] **Step 1: Declare the config fields in the manifest**

Update the existing GENERAL declarations so both actions consume the shared folder and date settings, then add the EMAIL section after the existing PER-TAB EXPORT section:

```vba
    Core_AddField reg, "k_export_folder", "Folder to write exported files to", "text", "", _
        "simple", _
        "Leave blank to use the folder this workbook is saved in - that is the usual " & _
        "answer. Otherwise a full path, e.g. C:\Close\January.", _
        "export.pdf,export.per_sheet,export.split_by_value,export.email"
    Core_AddField reg, "k_export_date_offset", "Date shift for exported file names (days)", _
        "number", "-1", "advanced", _
        "How many days to shift the date put in front of an exported file's name. -1 = " & _
        "yesterday, 0 = today, 1 = tomorrow.", _
        "export.pdf,export.per_sheet,export.split_by_value,export.email"

    Core_AddSection reg, "EMAIL"
    Core_AddField reg, "k_export_email_to", "Email recipient", "text", "", _
        "simple", _
        "The address pre-filled in the Outlook draft. The action displays the draft and " & _
        "never sends it.", _
        "export.email"
    Core_AddField reg, "k_export_email_subject", "Email subject", "text", "Sheet export", _
        "simple", _
        "The subject pre-filled in the Outlook draft.", _
        "export.email"
```

- [ ] **Step 2: Declare the actions in the manifest**

Place these after every field declaration and after the existing action declarations:

```vba
    Core_AddAction reg, "export.split_by_value", "Save Each Group As Its Own File", _
        "PkgExport_SplitByValue", _
        "Split by a column value into separate files"
    Core_AddAction reg, "export.email", "Email A Sheet As An Attachment", _
        "PkgExport_Email", _
        "Write one sheet to an .xlsx file, attach it to an Outlook draft with the recipient " & _
        "and subject from Setup, and display the draft without sending it."
```

- [ ] **Step 3: Write Save Each Group As Its Own File**

The user picks one rectangular range, says whether its first row is a header, and chooses exactly one column within that range. Each distinct typed value in that column is a group; blank cells form a separate `(blank)` group. Error values are refused instead of guessed.

Each output is a values-only `.xlsx` containing the optional header and only that group's rows. Its name is `yyyy-mm-dd <group>.xlsx`, with the complete base name passed through `Core_SafeFileName`. Preflight all names before writing. If two distinct groups produce the same case-insensitive sanitized or truncated name, refuse the whole batch and name both groups. Also refuse the whole batch if any destination already exists. A failure after writing begins keeps and reports the files already saved.

```vba
Public Sub PkgExport_SplitByValue(ctx As Object)
    Dim rng As Range, sourceRow As Range, targetRow As Range
    Dim outWb As Workbook, outWs As Worksheet
    Dim groups As Object, groupLabels As Object, seenPaths As Object
    Dim groupOrder As Collection, groupRows As Collection, paths As Collection
    Dim splitValue As Variant, hasHeader As Boolean, keys As String
    Dim splitCol As Long, firstDataRow As Long, rowIndex As Long, colIndex As Long
    Dim i As Long, j As Long, outRow As Long, n As Long
    Dim folderPath As String, dateStr As String, groupKey As String
    Dim groupLabel As String, fileName As String, outPath As String
    Dim duplicateMsg As String, existing As String, listMsg As String
    Const TITLE As String = "Save Each Group As Its Own File"
    On Error GoTo Fail

    folderPath = Core_ResolveFolder(Core_Cfg(ctx, "k_export_folder"))
    If Len(folderPath) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to " & _
               "export into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write " & _
               "exported files to' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_FolderExists(folderPath) Then
        Core_Refuse ctx, "This folder was not found: " & folderPath & vbCrLf & vbCrLf & _
               "Check 'Folder to write exported files to' on the Setup tab, Finalize, and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to export." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Core_IsEngineSheet(rng.Parent.Name) Then
        Core_Refuse ctx, "'" & rng.Parent.Name & "' belongs to the macro toolkit itself, " & _
               "not to source data that can be split into exported files." & vbCrLf & _
               vbCrLf & "Pick a range on a non-toolkit sheet and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub
    firstDataRow = 1
    If hasHeader Then firstDataRow = 2
    If rng.Rows.Count < firstDataRow Then
        Core_Refuse ctx, "The picked range contains a header but no data rows to split." & _
               vbCrLf & vbCrLf & "Pick a range that includes at least one data row and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(rng, keys, _
        "Choose exactly one column whose values should define the output groups.", _
        TITLE) Then Exit Sub
    splitCol = Core_ResolveColumn(rng, keys)
    If splitCol = 0 Then
        Core_Refuse ctx, "'" & keys & "' does not identify exactly one column inside the " & _
               "picked range." & vbCrLf & vbCrLf & "Run this again and choose one valid " & _
               "column name, letter, or position." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbBinaryCompare
    Set groupLabels = CreateObject("Scripting.Dictionary")
    groupLabels.CompareMode = vbBinaryCompare
    Set groupOrder = New Collection

    For rowIndex = firstDataRow To rng.Rows.Count
        splitValue = rng.Cells(rowIndex, splitCol).Value2
        If IsError(splitValue) Then
            Core_Refuse ctx, "The split column contains an error value at " & _
                   rng.Cells(rowIndex, splitCol).Address(False, False) & "." & _
                   vbCrLf & vbCrLf & "Fix that cell or choose another split column, then " & _
                   "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
            Exit Sub
        End If

        If Core_IsBlank(splitValue) Then
            groupKey = "Blank:"
            groupLabel = "(blank)"
        Else
            groupKey = TypeName(splitValue) & ":" & CStr(splitValue)
            groupLabel = CStr(splitValue)
        End If

        If Not groups.Exists(groupKey) Then
            Set groupRows = New Collection
            groups.Add groupKey, groupRows
            groupLabels.Add groupKey, groupLabel
            groupOrder.Add groupKey
        End If
        Set groupRows = groups.Item(groupKey)
        groupRows.Add rowIndex
    Next rowIndex

    dateStr = Format(Date + Core_CfgLong(ctx, "k_export_date_offset"), "yyyy-mm-dd")
    Set paths = New Collection
    Set seenPaths = CreateObject("Scripting.Dictionary")
    seenPaths.CompareMode = vbTextCompare

    For i = 1 To groupOrder.Count
        groupKey = CStr(groupOrder(i))
        groupLabel = CStr(groupLabels.Item(groupKey))
        fileName = Core_SafeFileName(dateStr & " " & groupLabel, 120) & ".xlsx"
        outPath = folderPath & fileName
        paths.Add outPath

        If seenPaths.Exists(fileName) Then
            duplicateMsg = duplicateMsg & vbCrLf & "   '" & groupLabel & "' and '" & _
                CStr(seenPaths.Item(fileName)) & "' both produce: " & fileName
        Else
            seenPaths.Add fileName, groupLabel
        End If

        If Len(Dir(outPath)) > 0 Then
            existing = existing & vbCrLf & "   " & outPath
        End If
    Next i

    If Len(duplicateMsg) > 0 Then
        Core_Refuse ctx, "These distinct groups produce the same safe file name, so this " & _
               "batch would overwrite its own output:" & duplicateMsg & vbCrLf & vbCrLf & _
               "Change or shorten one of each pair of group values, then run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Len(existing) > 0 Then
        Core_Refuse ctx, "These file(s) already exist and this action will not overwrite " & _
               "them:" & existing & vbCrLf & vbCrLf & "Rename or delete them, or change " & _
               "the folder or date offset on Setup, then run this again." & vbCrLf & _
               vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    For i = 1 To paths.Count
        listMsg = listMsg & vbCrLf & "   " & Core_FileName(CStr(paths(i)))
    Next i
    If MsgBox("This writes " & paths.Count & " values-only .xlsx file(s) to:" & vbCrLf & _
              "   " & folderPath & vbCrLf & vbCrLf & listMsg & vbCrLf & vbCrLf & _
              "The source range will not be changed. Continue?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    For i = 1 To groupOrder.Count
        groupKey = CStr(groupOrder(i))
        Set groupRows = groups.Item(groupKey)
        Set outWb = Workbooks.Add(xlWBATWorksheet)
        Set outWs = outWb.Worksheets(1)
        outWs.Name = "Data"
        outRow = 1

        If hasHeader Then
            Set sourceRow = rng.Rows(1)
            Set targetRow = outWs.Cells(outRow, 1).Resize(1, rng.Columns.Count)
            sourceRow.Copy Destination:=targetRow
            targetRow.Value = sourceRow.Value
            outRow = outRow + 1
        End If

        For j = 1 To groupRows.Count
            Set sourceRow = rng.Rows(CLng(groupRows(j)))
            Set targetRow = outWs.Cells(outRow, 1).Resize(1, rng.Columns.Count)
            sourceRow.Copy Destination:=targetRow
            targetRow.Value = sourceRow.Value
            outRow = outRow + 1
        Next j

        For colIndex = 1 To rng.Columns.Count
            outWs.Columns(colIndex).ColumnWidth = rng.Columns(colIndex).ColumnWidth
        Next colIndex

        Application.CutCopyMode = False
        outWb.SaveAs Filename:=CStr(paths(i)), FileFormat:=51
        n = n + 1
        outWb.Close SaveChanges:=False
        Set outWb = Nothing
    Next i
    Core_SpeedOff

    Core_Done ctx, "Wrote " & n & " group file(s) to:" & vbCrLf & _
           "   " & folderPath & vbCrLf & vbCrLf & _
           "The source range was not changed.", TITLE
    Exit Sub
Fail:
    If n > 0 Then
        Err.Description = Err.Description & vbCrLf & vbCrLf & n & " group file(s) were " & _
            "already written to " & folderPath & " before this failed. They were not " & _
            "removed; check the folder before re-running."
    End If
    If Not outWb Is Nothing Then PkgExport_CloseNoSave outWb
    Core_Fail ctx, "export.split_by_value", Err
End Sub

Private Sub PkgExport_CloseNoSave(ByVal wb As Workbook)
    On Error Resume Next
    wb.Close SaveChanges:=False
End Sub
```

- [ ] **Step 4: Write Email A Sheet As An Attachment**

The recipient and subject are mandatory Setup values. The user chooses one non-toolkit worksheet. The action writes a dated, values-only `.xlsx` attachment using the package's existing `PkgExport_OneSheetWorkbook` helper before attempting to reach Outlook.

Only the `CreateObject("Outlook.Application")` statement uses `On Error Resume Next`, immediately followed by `On Error GoTo Fail`. If Outlook is unavailable or blocked, call `Core_Refuse`, name the full attachment path, and leave that file on disk. Create and display a mail item with the configured recipient, subject, and attachment. Never call `.Send`.

```vba
Public Sub PkgExport_Email(ctx As Object)
    Dim recipient As String, subjectText As String
    Dim folderPath As String, outPath As String, dateStr As String
    Dim sheetName As Variant, pickedSheet As Object, ws As Worksheet
    Dim outlookApp As Object, mailItem As Object
    Dim attachmentWritten As Boolean
    Const TITLE As String = "Email A Sheet As An Attachment"
    On Error GoTo Fail

    recipient = Trim(Core_Cfg(ctx, "k_export_email_to"))
    If Len(recipient) = 0 Then
        Core_Refuse ctx, "No email recipient is configured." & vbCrLf & vbCrLf & _
               "Fill in 'Email recipient' on the Setup tab, Finalize, and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    subjectText = Trim(Core_Cfg(ctx, "k_export_email_subject"))
    If Len(subjectText) = 0 Then
        Core_Refuse ctx, "No email subject is configured." & vbCrLf & vbCrLf & _
               "Fill in 'Email subject' on the Setup tab, Finalize, and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    folderPath = Core_ResolveFolder(Core_Cfg(ctx, "k_export_folder"))
    If Len(folderPath) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to " & _
               "export into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write " & _
               "exported files to' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_FolderExists(folderPath) Then
        Core_Refuse ctx, "This folder was not found: " & folderPath & vbCrLf & vbCrLf & _
               "Check 'Folder to write exported files to' on the Setup tab, Finalize, and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    sheetName = Application.InputBox( _
        Prompt:="Sheet to export and attach:", _
        title:=TITLE, Default:=ActiveSheet.Name, Type:=2)
    If VarType(sheetName) = vbBoolean Then Exit Sub

    If Not Core_SheetExists(CStr(sheetName)) Then
        Core_Refuse ctx, "There is no sheet named '" & CStr(sheetName) & "'." & _
               vbCrLf & vbCrLf & "Choose an existing worksheet and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set pickedSheet = ThisWorkbook.Sheets(CStr(sheetName))
    If Not TypeOf pickedSheet Is Worksheet Then
        Core_Refuse ctx, "'" & CStr(sheetName) & "' is not a worksheet and cannot be " & _
               "exported as an .xlsx attachment." & vbCrLf & vbCrLf & _
               "Choose a worksheet and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    Set ws = pickedSheet

    If Core_IsEngineSheet(ws.Name) Then
        Core_Refuse ctx, "'" & ws.Name & "' belongs to the macro toolkit itself, not to " & _
               "a report that can be emailed." & vbCrLf & vbCrLf & _
               "Choose a non-toolkit worksheet and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    dateStr = Format(Date + Core_CfgLong(ctx, "k_export_date_offset"), "yyyy-mm-dd")
    outPath = folderPath & Core_SafeFileName(dateStr & " " & ws.Name, 120) & ".xlsx"
    If Len(Dir(outPath)) > 0 Then
        Core_Refuse ctx, "A file already exists at:" & vbCrLf & vbCrLf & _
               "   " & outPath & vbCrLf & vbCrLf & _
               "This action will not overwrite it. Rename or delete that file, or change " & _
               "the folder or date offset on Setup, then run this again." & vbCrLf & _
               vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If MsgBox("This first writes the attachment file to:" & vbCrLf & vbCrLf & _
              "   " & outPath & vbCrLf & vbCrLf & _
              "It then opens an Outlook draft addressed to " & recipient & "." & _
              vbCrLf & vbCrLf & "The draft will be displayed and will not be sent. Continue?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    PkgExport_OneSheetWorkbook ws.Name, outPath
    attachmentWritten = True
    Core_SpeedOff

    Set outlookApp = Nothing
    On Error Resume Next
    Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo Fail

    If outlookApp Is Nothing Then
        Err.Clear
        Core_Refuse ctx, "The attachment file was written to:" & vbCrLf & vbCrLf & _
               "   " & outPath & vbCrLf & vbCrLf & _
               "Outlook could not be reached. It may not be installed, or access may be " & _
               "blocked by policy. Attach this file by hand instead. The file has not been " & _
               "deleted." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set mailItem = outlookApp.CreateItem(0)
    With mailItem
        .To = recipient
        .Subject = subjectText
        .Attachments.Add outPath
        .Display
    End With

    Core_Done ctx, "Wrote the attachment to:" & vbCrLf & vbCrLf & _
           outPath & vbCrLf & vbCrLf & _
           "An Outlook draft has been displayed. It has not been sent.", TITLE
    Exit Sub
Fail:
    If attachmentWritten Then
        Err.Description = Err.Description & vbCrLf & vbCrLf & _
            "The attachment file was already written and was not deleted. Its full path is:" & _
            vbCrLf & outPath
    End If
    Core_Fail ctx, "export.email", Err
End Sub
```

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 10 package(s)`, no diagnostics.

- [ ] **Step 6: Confirm both ids registered**

```bash
grep -c 'Core_AddAction reg, "export.split_by_value"\|Core_AddAction reg, "export.email"' src/packages/PkgExport.bas
```

Expected: `2`.

- [ ] **Step 7: Append the run-time claims to the validation checklist**

- [ ] Splitting a range with three distinct group values writes exactly three `.xlsx` files, and each file contains the optional header plus only its group's values-only rows; missing files, live formulas, or rows from another group are failures. This proves the grouping, row selection, and non-destructive export behavior.
- [ ] Groups `A/B` and `A:B` produce a refusal naming both values and write zero files; any output file is a failure. This proves every name passes through `Core_SafeFileName` and same-name collisions are detected before mutation.
- [ ] With Outlook available, the attachment exists on disk before a displayed draft appears, and the draft has the Setup recipient, Setup subject, and attachment but is not sent; a missing file, incomplete draft, or sent message is a failure. This proves the required disk-first ordering and the prohibition on `.Send`.
- [ ] With Outlook unavailable or blocked, the attachment remains on disk and the refusal names its full path for manual attachment; a deleted file, missing path, or generic unhandled failure is a failure. This proves the mandated recoverable fallback.

- [ ] **Step 8: Commit**

```bash
git add src/packages/PkgExport.bas docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(export): split groups and draft sheet emails"
```

### Review notes on Task 2c — fix these while implementing

`export.email` follows the owner's specification exactly: the file is written before Outlook is
touched, the draft is displayed and never sent, and an unreachable Outlook refuses while naming
the file and leaving it on disk. The `On Error Resume Next` is correctly scoped to the
`CreateObject` alone and correctly restored with `On Error GoTo Fail`. These are the
corrections.

1. **The Outlook-unreachable refusal ends with `"Nothing has been changed."` and that is
   false.** A file was written — the same message says so three lines earlier. This is the one
   correction that must not be skipped: that sentence is load-bearing across the whole toolkit,
   and a message that says both "the attachment was written to X" and "nothing has been changed"
   teaches the user that the sentence means nothing. Replace it with something true, for
   example: `"The file above was kept so you can attach it yourself. Nothing was emailed."`

   Sweep the rest of the draft for the same error. Any refusal that fires **after** work has
   been committed to disk needs a different closing sentence from the ones that fire before.

2. **Verify `k_export_date_offset` and `PkgExport_OneSheetWorkbook` actually exist** before
   relying on them. The draft assumes both. If either is absent the build fails at the first
   compile, which is cheap — but the plan should not carry a call to something imaginary.

3. **`export.split_by_value` copies row by row through the clipboard.** `sourceRow.Copy
   Destination:=targetRow` inside a per-row loop is a clipboard round trip per row. For a
   10,000-row source that is slow enough to look hung. Copy each group's rows in one operation,
   or build a values array and assign it in one write, then apply formatting once.

4. **`Application.CutCopyMode = False` only after the inner loop.** Move it so the clipboard is
   not left populated between groups; on a failure mid-run the user is otherwise left with a
   marching-ants selection and a clipboard full of someone else's data.

5. **A name collision against an open workbook is not covered.** The existing-file check uses
   `Dir(outPath)`, which is correct for disk. But `Workbooks.Add` + `SaveAs` also fails when a
   workbook of the same *name* is already open from a different folder. Call `Core_AlreadyOpen`
   on each intended output path in the pre-flight loop and refuse naming the open workbook —
   the helper exists precisely for this, and the pre-flight loop is already walking every path.

6. **`Err.Description = Err.Description & ...` in the `Fail:` handler.** Legal, and the intent —
   telling the user which files were already written before the failure — is right and worth
   keeping. Confirm `Core_Fail` reads `.Description` from the object it is passed rather than
   re-reading `Err` after something has reset it.

---

## Task 2d: Workbook - Sort Worksheets Alphabetically and Add A Back To Contents Link To Every Sheet

**Files:**

- Modify: `src/packages/PkgWorkbook.bas`

**Interfaces:**

- Consumes: the frozen ids from Task 1.
- Produces: `PkgWorkbook_SortSheets(ctx As Object)` and `PkgWorkbook_TocLinks(ctx As Object)`, plus `PkgWorkbook_` privates.

**New config fields:** `k_workbook_toc_link_cell`, label `Cell for Back to Contents link`, kind `text`, default `A1`, tier `simple`, usedBy `workbook.toc_links`.

**Action decisions:** `workbook.sort_sheets` has no prompt or config field. It sorts ascending only, case-insensitively. Hidden user sheets participate and retain their visibility. Engine-owned sheets are never moved or crossed; each uninterrupted group of user sheets between engine sheets is sorted independently. `workbook.toc_links` processes visible and hidden user sheets, excludes engine sheets and the contents sheet, writes to exactly one configured cell without inserting or shifting rows, and never overwrites an occupied cell.

- [ ] **Step 1: Declare the config fields in the manifest**

Replace the existing `TABLE OF CONTENTS` declaration block with:

```vba
    Core_AddSection reg, "TABLE OF CONTENTS"
    Core_AddField reg, "k_workbook_toc_name", "Name for the Table of Contents sheet", "text", "Contents", _
        "simple", _
        "A sheet by this name is created the first run. On a later run, if a sheet by this " & _
        "name already exists - including this action's own sheet from a prior run - the " & _
        "action stops and asks for a different name; it never overwrites an existing sheet, " & _
        "so to rebuild the Table of Contents, delete the old one by hand first, then run this " & _
        "again. The same name identifies the existing contents sheet used by Back to Contents " & _
        "links.", _
        "workbook.build_toc,workbook.toc_links"
    Core_AddField reg, "k_workbook_toc_link_cell", "Cell for Back to Contents link", "text", "A1", _
        "simple", _
        "The single cell on each user worksheet where the Back to Contents link is placed. " & _
        "Use one A1-style cell address, such as A1 or D4. A non-empty target cell is never " & _
        "overwritten; that sheet and the cell's existing contents are reported instead.", _
        "workbook.toc_links"
```

- [ ] **Step 2: Declare the actions in the manifest**

Place both calls with the other action declarations, after all field declarations:

```vba
    Core_AddAction reg, "workbook.sort_sheets", "Sort Worksheets Alphabetically", "PkgWorkbook_SortSheets", _
        "Sorts each uninterrupted group of worksheets you own into case-insensitive A-Z order. " & _
        "Hidden sheets participate without being unhidden. The toolkit's own sheets are never " & _
        "moved, and user sheets do not cross them.", _
        "Workbook Housekeeping"

    Core_AddAction reg, "workbook.toc_links", "Add A Back To Contents Link To Every Sheet", "PkgWorkbook_TocLinks", _
        "Adds a Back to Contents hyperlink in the configured cell on every visible or hidden " & _
        "worksheet you own. Requires the configured Table of Contents sheet to exist first. " & _
        "An occupied target cell is never overwritten and is reported with its contents.", _
        "Workbook Housekeeping"
```

- [ ] **Step 3: Write Sort Worksheets Alphabetically**

```vba
Public Sub PkgWorkbook_SortSheets(ctx As Object)
    Dim ws As Worksheet
    Dim i As Long, firstIndex As Long
    Dim userCount As Long, movedCount As Long
    Const TITLE As String = "Sort Worksheets Alphabetically"
    On Error GoTo Fail

    Core_SpeedOn
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) Then userCount = userCount + 1
    Next ws

    i = 1
    Do While i <= ThisWorkbook.Worksheets.Count
        If Core_IsEngineSheet(ThisWorkbook.Worksheets(i).name) Then
            If firstIndex > 0 Then
                movedCount = movedCount + PkgWorkbook_SortSheetBlock(firstIndex, i - 1)
                firstIndex = 0
            End If
        Else
            If firstIndex = 0 Then firstIndex = i
        End If
        i = i + 1
    Loop
    If firstIndex > 0 Then
        movedCount = movedCount + _
            PkgWorkbook_SortSheetBlock(firstIndex, ThisWorkbook.Worksheets.Count)
    End If
    Core_SpeedOff

    If userCount < 2 Then
        Core_Done ctx, "There are fewer than two user worksheets, so there was nothing to sort." & _
               vbCrLf & vbCrLf & "The toolkit's own sheets were left in their existing positions.", TITLE
    ElseIf movedCount = 0 Then
        Core_Done ctx, "The user worksheets are already in case-insensitive A-Z order within " & _
               "each uninterrupted user tab group." & vbCrLf & vbCrLf & _
               "Hidden sheets participated without being unhidden. The toolkit's own sheets " & _
               "were left in their existing positions, and no user sheet crossed one.", TITLE
    Else
        Core_Done ctx, "Sorted " & userCount & " user worksheet(s) in case-insensitive A-Z " & _
               "order using " & movedCount & " tab move(s)." & vbCrLf & vbCrLf & _
               "Hidden sheets participated without being unhidden. The toolkit's own sheets " & _
               "were left in their existing positions, and no user sheet crossed one.", TITLE
    End If
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.sort_sheets", Err
End Sub

Private Function PkgWorkbook_SortSheetBlock(ByVal firstIndex As Long, _
        ByVal lastIndex As Long) As Long
    Dim names() As String, sortName As String
    Dim blockCount As Long, i As Long, j As Long
    Dim targetIndex As Long, movedCount As Long
    Dim ws As Worksheet

    blockCount = lastIndex - firstIndex + 1
    If blockCount < 2 Then Exit Function

    ReDim names(1 To blockCount)
    For i = 1 To blockCount
        names(i) = ThisWorkbook.Worksheets(firstIndex + i - 1).name
    Next i

    For i = 2 To blockCount
        sortName = names(i)
        j = i - 1
        Do While j >= 1
            If StrComp(names(j), sortName, vbTextCompare) <= 0 Then Exit Do
            names(j + 1) = names(j)
            j = j - 1
        Loop
        names(j + 1) = sortName
    Next i

    For i = 1 To blockCount
        targetIndex = firstIndex + i - 1
        Set ws = ThisWorkbook.Worksheets(names(i))
        If ws.Index <> targetIndex Then
            ws.Move Before:=ThisWorkbook.Worksheets(targetIndex)
            movedCount = movedCount + 1
        End If
    Next i

    PkgWorkbook_SortSheetBlock = movedCount
End Function
```

- [ ] **Step 4: Write Add A Back To Contents Link To Every Sheet**

```vba
Public Sub PkgWorkbook_TocLinks(ctx As Object)
    Dim destName As String, linkAddress As String, normalized As String
    Dim contents As Worksheet, ws As Worksheet
    Dim probe As Range, target As Range
    Dim skipped As Collection, failed As Collection
    Dim eligibleCount As Long, addedCount As Long
    Dim detail As String, validAddress As Boolean
    Const TITLE As String = "Add A Back To Contents Link To Every Sheet"
    On Error GoTo Fail

    destName = Trim(Core_Cfg(ctx, "k_workbook_toc_name"))
    If Len(destName) = 0 Then
        Core_Refuse ctx, "No name is set for the Table of Contents sheet." & vbCrLf & vbCrLf & _
               "Fill 'Name for the Table of Contents sheet' in on the Setup tab, Finalize, and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    linkAddress = Trim(Core_Cfg(ctx, "k_workbook_toc_link_cell"))
    If Len(linkAddress) = 0 Then
        Core_Refuse ctx, "No cell is set for the Back to Contents link." & vbCrLf & vbCrLf & _
               "Fill 'Cell for Back to Contents link' in on the Setup tab with one A1-style " & _
               "cell address, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_SheetExists(destName) Then
        Core_Refuse ctx, "The configured Table of Contents sheet, '" & destName & "', does not " & _
               "exist yet." & vbCrLf & vbCrLf & _
               "Run 'Build Table of Contents' first, or correct 'Name for the Table of " & _
               "Contents sheet' on the Setup tab, Finalize, and run this again." & vbCrLf & _
               vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Core_IsEngineSheet(destName) Then
        Core_Refuse ctx, "'" & destName & "' is one of the toolkit's own sheets and cannot be " & _
               "used as the Table of Contents." & vbCrLf & vbCrLf & _
               "Correct 'Name for the Table of Contents sheet' on the Setup tab, Finalize, " & _
               "and run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set contents = ThisWorkbook.Worksheets(destName)
    Set probe = Nothing
    On Error Resume Next
    Err.Clear
    Set probe = contents.Range(linkAddress)
    Err.Clear
    On Error GoTo Fail

    If Not probe Is Nothing Then
        If probe.Cells.CountLarge = 1 Then
            normalized = probe.Address(False, False, xlA1)
            If StrComp(Replace(linkAddress, "$", ""), normalized, vbTextCompare) = 0 Then
                validAddress = True
            End If
        End If
    End If
    If Not validAddress Then
        Core_Refuse ctx, "'" & linkAddress & "' is not one A1-style cell address." & _
               vbCrLf & vbCrLf & _
               "Correct 'Cell for Back to Contents link' on the Setup tab with a value such " & _
               "as A1 or D4, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set skipped = New Collection
    Set failed = New Collection
    Core_SpeedOn
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) Then
            If StrComp(ws.name, contents.name, vbTextCompare) <> 0 Then
                eligibleCount = eligibleCount + 1
                Set target = ws.Range(normalized)
                If Application.CountA(target) > 0 Then
                    detail = "'" & ws.name & "' at " & normalized & ": " & _
                        PkgWorkbook_CellFound(target)
                    skipped.Add detail
                    Core_Log ctx, "Skipped " & detail
                Else
                    On Error Resume Next
                    Err.Clear
                    ws.Hyperlinks.Add Anchor:=target, Address:="", _
                        SubAddress:="'" & Replace(contents.name, "'", "''") & "'!A1", _
                        TextToDisplay:="Back to Contents"
                    If Err.Number <> 0 Then
                        detail = "'" & ws.name & "' at " & normalized & ": " & Err.Description
                        failed.Add detail
                        Core_Log ctx, "Could not add link on " & detail
                        Err.Clear
                    Else
                        addedCount = addedCount + 1
                    End If
                    On Error GoTo Fail
                End If
            End If
        End If
    Next ws
    Core_SpeedOff

    Core_Done ctx, "Added Back to Contents links to " & addedCount & " of " & _
           eligibleCount & " eligible sheet(s)." & vbCrLf & vbCrLf & _
           skipped.Count & " sheet(s) were skipped because " & normalized & _
           " was already non-empty; no existing cell was overwritten." & vbCrLf & _
           failed.Count & " empty target cell(s) could not be linked." & _
           PkgWorkbook_DetailNote(skipped, "Skipped sheets and existing contents:") & _
           PkgWorkbook_DetailNote(failed, "Sheets where the link could not be added:"), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.toc_links", Err
End Sub

Private Function PkgWorkbook_CellFound(cell As Range) As String
    Dim s As String

    If cell.HasFormula Then
        s = CStr(cell.Formula)
        s = Replace(s, vbCr, " ")
        s = Replace(s, vbLf, " ")
        s = Replace(s, vbTab, " ")
        PkgWorkbook_CellFound = "formula " & s
    ElseIf IsError(cell.Value) Then
        PkgWorkbook_CellFound = "error " & CStr(cell.Text)
    Else
        s = CStr(cell.Value2)
        s = Replace(s, vbCr, " ")
        s = Replace(s, vbLf, " ")
        s = Replace(s, vbTab, " ")
        s = Replace(s, """", """""")
        PkgWorkbook_CellFound = "value """ & s & """"
    End If
End Function

Private Function PkgWorkbook_DetailNote(items As Collection, ByVal heading As String) As String
    Dim i As Long, note As String

    If items.Count = 0 Then Exit Function
    note = vbCrLf & vbCrLf & heading
    For i = 1 To items.Count
        note = note & vbCrLf & "- " & items(i)
    Next i
    PkgWorkbook_DetailNote = note
End Function
```

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 10 package(s)`, no diagnostics.

- [ ] **Step 6: Confirm both ids registered**

```bash
grep -c 'Core_AddAction reg, "workbook.sort_sheets"\|Core_AddAction reg, "workbook.toc_links"' src/packages/PkgWorkbook.bas
```

Expected: `2`.

- [ ] **Step 7: Append the run-time claims to the validation checklist**

- A deliberately scrambled group of visible and hidden user tabs finishes in case-insensitive A-Z order while the hidden tab remains hidden; failure is an out-of-order tab or changed visibility, and this proves hidden sheets participate without being exposed.
- Deliberately scrambled user groups on both sides of an engine sheet are each reordered while the engine sheet retains its recorded index and no user sheet crosses it; failure is movement of the engine sheet or a user tab across it, and this protects the engine-owned tab layout.
- With `k_workbook_toc_link_cell` set to non-default `D4`, every eligible empty `D4`, including one on a hidden sheet, contains exactly one working link to `Contents!A1`, while a prefilled value and a formula remain unchanged and are each reported with the sheet name and contents; failure is a missing link, changed visibility, overwritten cell, inserted row, or incomplete skip report.
- When `k_workbook_toc_name` names a sheet that does not exist, the action refuses and every candidate target cell remains unchanged; failure is any created hyperlink or created contents sheet, and this proves `workbook.toc_links` requires `workbook.build_toc` to have run first.

- [ ] **Step 8: Commit**

```bash
git add src/packages/PkgWorkbook.bas docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(workbook): sort sheets and add contents backlinks"
```

### Review notes on Task 2d — fix these while implementing

`workbook.toc_links` follows the owner's specification: one configured cell, occupied cells
skipped and reported with their contents, no rows inserted, engine sheets and the contents sheet
excluded. The address validation — resolve it, then check it round-trips to a single cell — is a
better guard than a regular expression would have been. These are the corrections.

1. **`Worksheet.Index` is not an index into `Worksheets`.** This is a real defect, not a style
   note. `Worksheet.Index` returns the position among **all** sheets, including chart sheets,
   while `PkgWorkbook_SortSheetBlock` compares it against a position computed by walking
   `ThisWorkbook.Worksheets`. In any workbook containing a chart sheet the two diverge and the
   sort moves tabs to wrong positions — silently, and on the user's own tabs.

   `format.chart_style` in Task 2b establishes that chart sheets are a case this phase expects
   to exist, so this is not hypothetical. Fix by working in one index space throughout: walk
   `ThisWorkbook.Sheets` and skip non-worksheets, or track the intended position without ever
   reading `.Index`. Add a checklist item that sorts a workbook containing a chart sheet.

2. **`PkgWorkbook_DetailNote` duplicates `PkgWorkbook_FailedNote`, which already exists** at
   `src/packages/PkgWorkbook.bas:196` and does the same job — naming the targets that did not
   get the change so a partial result is stated rather than rounded up to "done". Use the
   existing helper. A second helper with the same purpose is how a file grows two conventions.

3. **The completion message has no length bound.** One line per skipped sheet, in a `MsgBox`, on
   a 60-tab workbook where every `A1` is occupied, produces a dialog that is unreadable and may
   be truncated by Excel. Cap the list — the first ten by name, then "and N more" — and put the
   full list in the log via `Core_Log`, which the draft is already calling per sheet.

4. **The zero cases are reported as though they were findings.** `"0 sheet(s) were skipped"` and
   `"0 empty target cell(s) could not be linked"` should not appear at all. Build the message
   from the branches that actually happened.

5. **`workbook.sort_sheets` runs the whole sort before checking `userCount < 2`.** Harmless, but
   check first and return early — it makes the no-op case obviously a no-op rather than
   something that ran and reported nothing.

6. **Hidden sheets get a link they cannot be seen to have.** This is fine and is the opposite of
   the `build_toc` case, where a link *to* a hidden sheet is unnavigable. Say so in the
   completion message so it does not read as an inconsistency with the action next to it.

---

### Cross-cutting: confirmation dialogs inside a chain — SETTLED

Several of the new actions ask something before they do work: a confirmation before overwriting
cells, restyling every chart or writing files, and a prompt for the case mode and the sheet to
email. `Core_Done` and `Core_Refuse` suppress their dialogs when an action runs as a **quiet
chain step**; a raw `MsgBox` or `Application.InputBox` does not. So a chain containing one of
these stops and waits for a human.

**Owner's decision, 2026-07-29: accept that they block a chain, and say so in their help text.**
Do not add a quiet-mode query to the Core API, and do not remove the confirmations — the actions
that have them are the ones that change the most.

Applies to every new action that asks anything before acting:

| Action | What it asks |
|---|---|
| `clean.change_case` | which case to apply |
| `format.highlight_special` | confirm before adding, and before removing, the audit view |
| `format.chart_style` | confirm before restyling every chart |
| `export.split_by_value` | confirm before writing the batch of files |
| `export.email` | which sheet, and confirm before writing and drafting |

**Each of those five `Core_AddAction` help strings must end with this sentence, verbatim, so it
reads the same everywhere it appears:**

```text
In a chain, this one stops and waits for your answer before it does anything.
```

It lands in the help text, which is shown on the button card and — after Task 5 — in full on
`Macros Home`, so a user composing a chain sees it at the moment they are choosing steps.

`clean.split_column`, `workbook.sort_sheets` and `workbook.toc_links` ask nothing and do not get
the sentence. Do not add it to them; a sentence that appears on every action stops being read.

**This is not a new class of behaviour** — `Core_AskRange` already blocks a chain the same way in
existing actions. What this decision settles is that the phase does not try to fix that, and
states it plainly instead.

---

## Task 3: Chain steps become comma-separated action ids

**Files:**
- Modify: `src/core/Core.bas` — `ParseStepTokens` (`:2570-2589`), `StepsToTokens`
  (`:2653-2659`), `RenderActionTokens` (`:2661-2673`), `RenderChainsBlock` (`:2699-2728`).

**Interfaces:**
- Consumes: nothing from Task 2. This task is independent of the eight actions and may be done
  before or after them.
- Produces: `ParseStepTokens(ByVal cell As String) As Collection` — unchanged signature,
  widened acceptance. `StepsToTokens` now emits bare ids joined by `", "`.

**What the user sees.** The Steps cell today holds `Remove Duplicates [clean.dupes]` on one
line, `Trim Spaces [clean.trim]` on the next. Afterwards it holds
`clean.dupes, clean.trim` on one line. The header changes from
`Steps (one token per line)` to `Steps (action ids, separated by commas)`.

- [ ] **Step 1: Widen `ParseStepTokens` — bracket-aware, not a blind comma split**

The naive change is `cell = Replace(cell, ",", vbLf)` beside the existing
`cell = Replace(cell, ";", vbLf)`. **Do not do that.** It splits a legacy bracketed token
whose caption contains a comma into two broken halves. No caption contains one today —
verified, zero of 43 — but the parser must not silently depend on that staying true, because
the failure is a chain that loses a step at Finalize with an `[UNKNOWN]` beside a fragment.

Split on newline and semicolon first; then split a line on commas **only when it carries no
bracket**:

```vba
' "Label [id]" or a bare id per line, and bare ids may be comma-separated on one line.
' Only the bracketed id is authoritative. A line carrying a bracket is NEVER comma-split:
' a caption may contain a comma and splitting it would break a legacy token in half.
Private Function ParseStepTokens(ByVal cell As String) As Collection
    Dim out As Collection, lines As Variant, parts As Variant
    Dim i As Long, j As Long, ln As String, id As String
    Dim lo As Long, hi As Long
    Set out = New Collection
    If Len(Trim(cell)) = 0 Then Set ParseStepTokens = out: Exit Function
    cell = Replace(cell, ";", vbLf)
    lines = Split(cell, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim(lines(i))
        If Len(ln) > 0 Then
            lo = InStrRev(ln, "["): hi = InStrRev(ln, "]")
            If lo > 0 And hi > lo Then
                id = Trim(Mid(ln, lo + 1, hi - lo - 1))
                If Len(id) > 0 Then out.Add id
            Else
                parts = Split(ln, ",")
                For j = LBound(parts) To UBound(parts)
                    id = Trim(parts(j))
                    If Len(id) > 0 Then out.Add id
                Next j
            End If
        End If
    Next i
    Set ParseStepTokens = out
End Function
```

Read the live procedure first and keep whatever local declarations and guards it already has;
the block above shows the shape of the change, not a licence to discard existing behaviour.
**`lo > 0 And hi > lo` does not short-circuit** — that is fine here because both operands are
plain comparisons on already-assigned Longs, but do not add a term that could fault.

- [ ] **Step 2: Make `StepsToTokens` write bare ids**

It currently joins `ActionTokenFor(id)` — the `Caption [id]` form — with `vbLf`. Join the
**raw ids** with `", "` instead. Leave `ActionTokenFor` itself alone: `RenderActionTokens` and
`ActionCaptionFor` still use it, and Step 3 changes only where it is called from.

- [ ] **Step 3: Make the column-G copy list lead with the id**

`RenderActionTokens` writes `Caption [id]` lines into column 7 (G) under the header
`"Available actions - copy a token into Steps:"`. Change it to write the **bare id into G** and
the **caption into H**, and change the header to
`"Available actions - copy an id into Steps:"`. The id is now the thing to copy, so it must be
the thing in the column the user drags across.

Check what else column H holds on that row band before writing into it —
`Core.bas:2170-2188` puts `"Reference: Excel files in THIS folder"` in H4. Do not collide;
`RenderActionTokens` starts one row above the first chain row, far below row 4, but confirm on
the rendered sheet rather than assuming.

- [ ] **Step 4: Change the header caption**

In `RenderChainsBlock`, `ws.Cells(r, VAL_COL + 1).Value = "Steps (one token per line)"` becomes
`"Steps (action ids, separated by commas)"`. Also update the `BlockIntro` text two lines above,
which currently reads `"Button name + Steps (paste tokens from the Actions list, one per line) + Placement (1A, hide, or TabName!B4)."`

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 10 package(s)`, no diagnostics.

- [ ] **Step 6: Confirm the old format is still parsed**

```bash
grep -n 'InStrRev(ln, "\[")' dist/MacrosBundle.bas
```

Expected: one hit. This is the backward-compatibility guarantee — an existing `_Config`
carrying `Caption [id]` steps must survive the upgrade rather than failing Finalize with
`[UNKNOWN]` on every step. Losing this line silently breaks every deployed workbook that has a
chain, and nothing else in the build would catch it.

- [ ] **Step 7: Append the run-time claims to the validation checklist**

Add to `docs/designs/2026-07-28-end-to-end-validation.md`, in section 4:

- an item that authors a chain by typing `clean.dupes, clean.trim` into Steps, clicks
  **Finalize Setup Tab**, and expects a working chain button — failure is `(no steps!)` or
  `[UNKNOWN]` in the Finalize report;
- an item that confirms a chain authored in the **old** `Caption [id]` form still finalizes
  after the upgrade.

Both state the observable result first, then what failure looks like, then why — the checklist's
house style.

- [ ] **Step 8: Commit**

```bash
git add src/core/Core.bas docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(core): chain steps are comma-separated action ids"
```

---

## Task 4: Drop the numbering from the ImportExport captions

**Files:**
- Modify: `src/packages/PkgImportExport.bas:80-83` — four caption strings.
- Modify: `spec/contract.md` §8 — the four matching registry rows.
- Modify: `docs/designs/2026-07-28-end-to-end-validation.md:67,74` — two live checklist items.

**Interfaces:**
- Consumes: nothing.
- Produces: the four final captions — `Clear Input Tabs`, `Import Source Data`,
  `Update Ref_Master`, `Export Dashboard (.xlsx)`. Task 11's guide names these; Task 9's
  checklist items must match them exactly.

**Why it is more than four strings.** These captions are the button faces, they are frozen in
contract §8's registry table, and two of them appear by name in checklist items that are still
unrun. Changing one and not the others leaves a checklist telling the owner to click a button
that does not exist — which is precisely the defect the last blind dry-run found five instances
of.

- [ ] **Step 1: Edit the four captions**

In `src/packages/PkgImportExport.bas`, lines 80-83. The ids, entry points, help strings and
categories do not change — only the caption argument.

```vba
    Core_AddAction reg, "importexport.clear", "Clear Input Tabs", "PkgImportExport_Clear", "Deletes data (row 2 down) on the input tabs. Keeps headers + formatting."
    Core_AddAction reg, "importexport.import", "Import Source Data", "PkgImportExport_Import", "Clears the input tabs, then imports the matching source files from this folder (auto-handles (1)/(2) name suffixes)."
    Core_AddAction reg, "importexport.updateMaster", "Update Ref_Master", "PkgImportExport_UpdateMaster", "Refreshes the Ref_Master tab from the master-list file in this folder."
    Core_AddAction reg, "importexport.export", "Export Dashboard (.xlsx)", "PkgImportExport_Export", "Saves a values-only copy of the main tabs (formatting + filters + conditional formatting kept)."
```

**Do not touch the ids.** `importexport.updateMaster` is camelCase and stays camelCase — it is
one of the eight permanent shipped exceptions named in contract §7.

- [ ] **Step 2: Sweep contract §8**

Find the four rows in the frozen registry whose caption column carries the numbering and remove
it. Search:

```bash
grep -n '`importexport\.' spec/contract.md
```

- [ ] **Step 3: Sweep the checklist**

`docs/designs/2026-07-28-end-to-end-validation.md` line 67 says *"Run **2.  Import Source
Data**"* and line 74 says *"Repeat it for **3.  Update Ref_Master**"*. Drop the numbers.

- [ ] **Step 4: Prove no numbered caption survives anywhere that matters**

```bash
grep -rn '[1-4]\.  \(Clear Input\|Import Source\|Update Ref_Master\|Export Dashboard\)' src/ spec/ docs/designs/2026-07-28-end-to-end-validation.md
```

Expected: **no matches**. Hits in `src/legacy/` are dead code and hits in older design documents
are historical record — neither is swept. Confine the check to the paths above.

- [ ] **Step 5: Build, syntax-check, and confirm the captions in the bundle**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c '"Clear Input Tabs"' dist/MacrosBundle.bas
```

Expected: build clean, and exactly `1`.

- [ ] **Step 6: Commit**

```bash
git add src/packages/PkgImportExport.bas spec/contract.md docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(importexport): drop the step numbering from the four button captions"
```

---

## Task 5: Restore Macros Home

**Files:**
- Modify: `src/core/Core.bas` — `Core_BuildControlPanel` (`:3815`), `DeleteHomePanel`
  (`:3943`), `RenderNavRow` (`:4344`), `LandOnFirstPage` (`:4815`); add the six procedures from
  the attic.
- Delete: `src/attic/HomeIndex.bas.txt`.
- Modify: `docs/designs/2026-07-28-end-to-end-validation.md`.

**Interfaces:**
- Consumes: the eight ids from Task 1 and the eight actions from Task 2 — Home is rendered at
  full 51-action scale, which is the point of sequencing it after them.
- Produces: `HOME_SHEET` (`"Macros Home"`) as a generated tab. Task 7's `Complete Macro Setup`
  deletes it and strips its nav link; Task 11's guide names it.

**Read first:** `src/attic/HomeIndex.bas.txt` in full — all 295 lines. It is small, it is the
source material for this task, and its header block enumerates the five wiring points below
with more detail than this plan repeats. This is the one file in this plan you read whole.

- [ ] **Step 1: Paste the six procedures into `Core.bas`**

`BuildHomePage`, `HomeCatalog`, `HomeLegend`, `HomeLine`, `AddText`, `PageButtonCount`. Put
them beside the other panel-rendering procedures, after `BuildAdminPanel`. They reference
Core privates that still exist — `EnsurePanelSheet`, `NavLink`, `SortedKeys`, `HOME_SHEET`,
`PANEL_BASE`, `ADMIN_SHEET`, `CARD_COLS` — so this is a paste, not a rewrite.

They are all `Private`. Check that no name collides with something added since the withdrawal:

```bash
grep -nE '^(Private|Public) (Sub|Function) (BuildHomePage|HomeCatalog|HomeLegend|HomeLine|AddText|PageButtonCount)\b' src/core/Core.bas
```

Expected before the paste: **no matches**. `AddText` is the one to watch — it is a generic name
and the likeliest collision. If it collides, rename the pasted one `HomeText`; the bundle is a
single module and a duplicate procedure name fails the build.

- [ ] **Step 2: Make the catalog list all 51 actions, not just the buttons**

This is the substantive change, and the reason this is not a pure restore. The attic's
`HomeCatalog` walks the **layout** — which contains only buttons — so on a fresh import, where
every action is `off`, it renders nothing.

Walk the **registry** instead, exactly the way `GatherButtons` does. **Read
`Core.bas:3969-3991` before writing this step** and reuse its registry-iteration and
`ReadAvailMap` / `AvailFor` calls rather than inventing an accessor. For each action emit:
caption, its full `help` text, its availability (`Button` / `Chains only` / `Off`), and a link
to its page when the layout has one.

Do not abbreviate the description. `CardHelp` truncates for the button cards because a card is
small; the catalog is where the untruncated text belongs, and that is most of what makes this
page worth having.

Chains keep their existing treatment — they are registry entries under `CHAIN_CAT`.

- [ ] **Step 3: Build Home last, and retire `LandOnFirstPage`**

In `Core_BuildControlPanel`, the pipeline currently ends `PlaceAnchoredButtons` →
`LandOnFirstPage layout`. Replace that last call with `BuildHomePage layout, manual`.

`BuildHomePage` activates and selects, so it must stay last — anything after it changes where
the user lands. Then check `LandOnFirstPage` has no other caller and delete it:

```bash
grep -n 'LandOnFirstPage' src/core/Core.bas
```

Expected after the edit: only the declaration, which you then remove. A dead private that
nothing calls is a name reserved against the whole bundle for nothing.

**This fixes the no-pages case as a side effect.** When every action is `off` the layout is
empty and there is currently no page to land on at all. Home is always built, so there is
always somewhere to land.

- [ ] **Step 4: Remove `DeleteHomePanel` entirely**

Delete the procedure at `Core.bas:3943` and its call at `Core.bas:3839`. It exists only to stop
a stale index surviving a rebuild; now that the index is rebuilt every time, it deletes the
sheet that was just made. `EnsurePanelSheet` reuses and clears, which covers the stale case.

Leave `DeletePanels` alone. Its `IsPanelName` numeric-tail carve-out is what spares
`Macros Home`, and it is already correct for a restore — `Core.bas:4980-4993` records that the
carve-out was deliberately kept against exactly this day.

- [ ] **Step 5: Restore the nav row**

`RenderNavRow` currently writes `< Previous` to cell B and `Next >` to cell D. Restore the
three-cell row: `<< Home` in B, `< Previous` in D, `Next >` in F.

```vba
    NavLink ws.Cells(1, 2), "<< Home", HOME_SHEET
    If Len(prevSheet) > 0 Then NavLink ws.Cells(1, 4), "< Previous", prevSheet
    If Len(nextSheet) > 0 Then NavLink ws.Cells(1, 6), "Next >", nextSheet
```

**Do not guard the Home link with `Core_SheetExists`.** `Core.bas:4348-4355` records that this
was tried: because Home builds last, the guard rendered plain bold text instead of a link on
every page of a first build, and self-healed only on the second. The link is written
unconditionally. Read that comment before touching this procedure and leave it in place.

- [ ] **Step 6: Build, syntax-check, and confirm the wiring**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c 'DeleteHomePanel\|LandOnFirstPage' dist/MacrosBundle.bas   # expect 0
grep -c 'BuildHomePage' dist/MacrosBundle.bas                      # expect 2 (declaration + call)
grep -c '"<< Home"' dist/MacrosBundle.bas                          # expect 1
```

- [ ] **Step 7: Delete the attic copy**

```bash
git rm src/attic/HomeIndex.bas.txt
```

The attic exists to preserve code that is not built. Once it is built, a second copy is a
divergence risk: the next person to fix a bug in `BuildHomePage` has no reason to know the
attic exists, and the stale twin then reads as authoritative. The history keeps it.

- [ ] **Step 8: Append the run-time claims to the validation checklist**

Section 4 of `docs/designs/2026-07-28-end-to-end-validation.md`. At minimum:

- **`Macros Home` exists after Finalize and lists 51 actions.** Failure: fewer than 51, or the
  tab is missing.
- **An action set to `Off` still appears, marked `Off`, with no page link.** *This is the item
  that proves the §6 catalog rewrite actually shipped* — on a fresh import it is the difference
  between a full page and an empty one.
- **Every `Macros <n>` page carries a working `<< Home` link in B1, on the first build.** The
  "first build" wording is the whole point; the defect this replaces self-healed on the second.
- **The workbook lands on `Macros Home` after Finalize**, not on `Macros 1`.

- [ ] **Step 9: Commit**

```bash
git add src/core/Core.bas src/attic docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(core): restore Macros Home as a full catalog of every action and its state"
```

---

## Task 6: The Finalize reminder

**Files:**
- Modify: `src/core/Core.bas` — `RenderPlacementBlock` (`:2968`) and the surrounding
  `Core_BuildSetup` region.
- Modify: `docs/designs/2026-07-28-end-to-end-validation.md`.

**Interfaces:**
- Consumes: `AVAIL_HEADER`, `AVAIL_OFF`, `LAYOUT_KEY` — existing Core constants.
- Produces: nothing later tasks depend on. Task 11's guide describes the behaviour.

**What it is.** A line above the Action Catalog block that appears whenever any action's
availability is not `Off`, telling the user to click **Finalize Setup Tab** and then
**Create Setup Tab** to finish setting up the buttons they just switched on — plus a hyperlink
to `Setup Menu` beside it.

**The trade, already settled — do not re-open it.** The condition stays true forever once
anything is configured, so the line never turns off again. The owner chose this over comparing
Setup against the activated `_Config`, on 2026-07-29. It is recorded in the design's §4.4.
Implement what was chosen.

- [ ] **Step 1: Write the reminder row**

Directly above the Action Catalog block header, inside `RenderPlacementBlock` before it writes
`LAYOUT_KEY`. Two cells:

- the message cell, holding a formula that resolves to the reminder text when any availability
  cell is not `Off`, and to an empty string otherwise;
- a hyperlink cell pointing at `ADMIN_SHEET`, written with `NavLink` — the helper the nav rows
  already use, so the link styling matches the rest of the workbook.

The formula counts availability cells that are non-blank and not `Off` over the catalog's own
row range, which the renderer knows because it is about to write it. Compute the first and last
data rows, then write an absolute range into the formula string. A worked shape:

```vba
    ' rFirst/rLast are the catalog's first and last DATA rows, known here because the
    ' block is about to be written. Absolute refs: the sheet is copied wholesale to
    ' _Config at Finalize and a relative ref would drift.
    f = "=IF(COUNTIFS(" & colLetter & rFirst & ":" & colLetter & rLast & ",""<>"" ," _
      & colLetter & rFirst & ":" & colLetter & rLast & ",""<>Off"")>0,""" _
      & REMINDER_TEXT & ""","""")"
    ws.Cells(rNote, VAL_COL).Formula = f
```

Declare `REMINDER_TEXT` as a module-level `Private Const` **above Core's first procedure** —
the bundle is one module and a declaration after the first procedure fails the build. Keep it
under 25 line continuations.

- [ ] **Step 2: Give the row no machine key**

Column A of the reminder row stays **empty**. Blocks are located by key through `FindKeyRow`,
never by row offset, so a keyless row inserted above the catalog cannot shift any block's
parse. A key would make it look like a block.

- [ ] **Step 3: Do not reword the `Availability` header**

`MarkerVerdict`'s tamper test reads the `AVAIL_HEADER` string back off the sheet
(`Core.bas:63-71`). The reminder sits above that header and must not touch it. `PLACE_HEADER`
is explicitly not read back and is safe — but nothing in this task needs to change either.

- [ ] **Step 4: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

- [ ] **Step 5: Append the run-time claims to the validation checklist**

- **On a fresh import with everything `Off`, the reminder line is blank.** Failure: the text
  shows when nothing is switched on, which means the formula's `<>Off` test is matching blanks.
- **Setting one action to `Button` makes the line appear immediately, before Finalize.** This
  is the whole value of using a formula rather than static text — it reacts as the user types.
- **The link beside it opens `Setup Menu`.**
- **The Setup tab still finalizes** — the added row has not shifted any block's parse. Failure
  here looks like a Finalize that reports the wrong number of actions or cannot find a block.

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(core): remind the user to finalize and rebuild Setup once an action is on"
```

---

## Task 7: `Complete Macro Setup`, and the Alt+F8 tidy

**Files:**
- Modify: `src/core/Core.bas` — `BuildAdminPanel` (`:4052-4090`); new `Core_CompleteSetup`;
  `Core_SpeedOn` (`:719`) and `Core_SpeedOff` (`:734`) become `Private`.
- Modify: `spec/contract.md` §9 — only if it names either helper.
- Modify: `docs/designs/2026-07-28-end-to-end-validation.md`.

**Interfaces:**
- Consumes: `HOME_SHEET`, `SETUP_SHEET`, `ADMIN_SHEET`, `PANEL_BASE`, `IsPanelName`,
  `CfgSheet`, `Core_SheetExists` — all existing.
- Produces: `Public Sub Core_CompleteSetup()` — no arguments, so Alt+F8 lists it. Task 11's
  guide ends on this button.

**What it is.** The user has the macros they want, set up the way they want. This strips the
configuration surface so the workbook is just their buttons.

- [ ] **Step 1: Add the button to `BuildAdminPanel`**

The existing two buttons sit at `topPos` and `topPos + 45`. Put this one a further **75 points**
below — about five rows — not another 45. The gap is deliberate: this is not a button to click
by accident, and visual separation is the cheapest guard there is.

```vba
    topPos = topPos + 75
    AddBtn ws, "btnCompleteSetup", "Complete Macro Setup", "Core_CompleteSetup", leftPos, topPos, 220, 32
```

Give it a description in column C in the same shape as the other two, saying what it removes and
that `Alt+F8` then `Core_BuildSetup` brings Setup back.

- [ ] **Step 2: Write the two refusals, before anything else happens**

```vba
Public Sub Core_CompleteSetup()
    Dim ws As Worksheet, target As Worksheet, ans As VbMsgBoxResult
    Dim nPages As Long
    Const TITLE As String = "Complete Macro Setup"
    On Error GoTo Fail

    If Not Core_SheetExists(CONFIG_SHEET) Then
        MsgBox "Setup has never been finalized, so there is nothing to complete." & vbCrLf & vbCrLf & _
               "Click 'Finalize Setup Tab' first." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    For Each ws In ThisWorkbook.Worksheets
        If IsPanelName(ws.name) Then
            nPages = nPages + 1
            If target Is Nothing Then Set target = ws
        End If
    Next ws

    If nPages = 0 Then
        MsgBox "There are no macro pages in this workbook yet, so removing the setup tabs " & _
               "would leave you with nothing at all." & vbCrLf & vbCrLf & _
               "Switch on the macros you want, click 'Finalize Setup Tab', then run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If
```

`Core_CompleteSetup` takes no `ctx`, so it cannot use `Core_Refuse` — that sink needs a ctx to
set `failed`/`refused` on. A plain `MsgBox` is correct here, the same as the other engine entry
points. Keep the `"Nothing has been changed."` sentence: it is the house phrasing and the user
has seen it everywhere else.

- [ ] **Step 3: Write the confirmation**

It must name every tab that will go, state plainly that settings survive, give the way back, and
warn about unfinalized edits. The owner chose warn-and-proceed over automatic detection, so the
warning is the only thing standing between a pending edit and its loss — it has to be legible,
not a formality.

```vba
    ans = MsgBox("This removes the setup tabs from this workbook:" & vbCrLf & _
                 "    Setup" & vbCrLf & "    Setup Menu" & vbCrLf & "    Macros Home" & vbCrLf & vbCrLf & _
                 "Your macro buttons and all your settings are kept." & vbCrLf & vbCrLf & _
                 "IF YOU HAVE CHANGED ANYTHING ON THE SETUP TAB SINCE YOU LAST CLICKED " & _
                 "'Finalize Setup Tab', CLICK NO NOW, FINALIZE FIRST, THEN COME BACK. " & _
                 "Unfinalized changes will be lost." & vbCrLf & vbCrLf & _
                 "To get the Setup tab back later: press Alt+F8, choose Core_BuildSetup, " & _
                 "click Run." & vbCrLf & vbCrLf & "Remove the setup tabs now?", _
                 vbYesNo + vbExclamation + vbDefaultButton2, TITLE)
    If ans <> vbYes Then Exit Sub
```

`vbDefaultButton2` makes **No** the default, so a reflexive Enter does nothing.

- [ ] **Step 4: Activate the landing page before deleting anything**

```vba
    Core_SpeedOn
    target.Activate
```

Activate first. If `Setup Menu` is the active sheet when it is deleted, Excel picks the
replacement itself and the user lands somewhere arbitrary. Choosing first makes the outcome
defined.

- [ ] **Step 5: Strip the `<< Home` link from every page**

```vba
    For Each ws In ThisWorkbook.Worksheets
        If IsPanelName(ws.name) Then
            ws.Cells(1, 2).Hyperlinks.Delete
            ws.Cells(1, 2).Clear
        End If
    Next ws
```

Both lines. `Clear` alone can leave the `Hyperlink` object behind, and a hyperlink pointing at a
deleted sheet is a click that raises an error rather than doing nothing.

- [ ] **Step 6: Delete the three tabs, `Setup Menu` last**

```vba
    If Core_SheetExists(HOME_SHEET) Then ThisWorkbook.Worksheets(HOME_SHEET).Delete
    If Core_SheetExists(SETUP_SHEET) Then ThisWorkbook.Worksheets(SETUP_SHEET).Delete
    If Core_SheetExists(ADMIN_SHEET) Then ThisWorkbook.Worksheets(ADMIN_SHEET).Delete
    Core_SpeedOff
    MsgBox "Setup is complete. Your macro buttons are ready to use." & vbCrLf & vbCrLf & _
           "Your settings are saved and hidden. To change them later: press Alt+F8, " & _
           "choose Core_BuildSetup, click Run.", vbInformation, TITLE
    Exit Sub
Fail:
    RestoreSpeed
    MsgBox "Could not finish: " & Err.Description, vbCritical, TITLE
End Sub
```

**`Setup Menu` is deleted last on purpose.** It hosts the button that is currently running, so
it is the delete most likely to misbehave; ordering it last means a failure leaves the most
recoverable state — the user still has Setup Menu and can try again.

`Core_SpeedOn` has already set `DisplayAlerts = False`, so the deletes do not prompt. Do not
set it separately and do not reset it by hand — `Core_SpeedOff` and `RestoreSpeed` own that
state, and touching it directly breaks the depth counting.

**Do not delete `_Config`.** It is read on every action run (`Core.bas:1742`), chains resolve
from it (`:1884`), and every setting the user typed lives there (`:549-563`). Deleting it does
not tidy the workbook, it stops all 51 buttons working. **Do not touch `_Packages` either** —
an absent `_Packages` means every package is enabled (`PackageEnabled`, `Core.bas:1629`), so
deleting one that disables a package silently turns it back on.

- [ ] **Step 7: Make `Core_SpeedOn` and `Core_SpeedOff` `Private`**

Alt+F8 lists only public `Sub`s that take no arguments. Today that is four names, two of which
are internal helpers a user must never run — running one leaves Excel with calculation and
screen updating off and nothing on screen saying so.

Change both declarations to `Private Sub`. The bundle is a single module, so every existing
caller still resolves. Confirm nothing outside Core calls them by name:

```bash
grep -n 'Core_SpeedOn\|Core_SpeedOff' src/packages/*.bas
```

Every hit is legal — the packages are concatenated into the same module — but read them to
confirm none is in a comment claiming a public API. Then check the contract:

```bash
grep -n 'Core_SpeedOn\|Core_SpeedOff' spec/contract.md
```

If §9's platform-abstraction appendix names either as part of the public surface, amend it in
this same commit.

- [ ] **Step 8: Build, syntax-check, and confirm the Alt+F8 list**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -oE '^Public Sub [A-Za-z_0-9]+\(\)' dist/MacrosBundle.bas
```

Expected exactly three, and no others:

```text
Public Sub Core_BuildSetup()
Public Sub Core_FinalizeSetup()
Public Sub Core_CompleteSetup()
```

This grep is the whole verification for Step 7 and it is worth running deliberately — nothing
else in the build would notice a helper leaking back into the macro dialog.

- [ ] **Step 9: Append the run-time claims to the validation checklist**

These are the highest-consequence items this phase adds. Put them in **section 2**, "Can it
destroy your work?", not section 4.

- **On a workbook that has never been finalized, the button refuses and deletes nothing.**
- **On a finalized workbook with no macro pages, it refuses and deletes nothing.**
- **Answering No deletes nothing**, and Enter alone answers No.
- **After Yes: `Setup`, `Setup Menu` and `Macros Home` are gone; every `Macros <n>` page and
  every button on it still works.** Click one and confirm it runs. *This is the item that
  proves `_Config` survived* — if it were deleted the buttons would still be there and would
  silently do nothing, which is the worst possible failure shape.
- **B1 of every remaining page is empty**, with no dead link.
- **`Alt+F8` then `Core_BuildSetup` brings the Setup tab back**, populated with the settings
  that were there before. This discharges the recovery route the confirmation promises.
- **`Alt+F8` lists exactly three names.**

**Mark as an assumption:** that a macro can delete the sheet hosting the button that invoked it.
Nothing in this codebase exercises it. If Excel refuses, the fallback is `Application.OnTime` to
defer the final delete past the end of the handler.

- [ ] **Step 10: Commit**

```bash
git add src/core/Core.bas spec/contract.md docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(core): Complete Macro Setup removes the configuration tabs when a workbook is finished"
```

---

## Task 8: Explanatory notes on four Setup headers — the last code task

**Files:**
- Modify: `src/core/Core.bas` — a new note helper; `RenderPlacementBlock` (`:2968`),
  `RenderChainsBlock` (`:2703`), `RenderTable` (`:2512`).
- Modify: `src/packages/PkgImportExport.bas:51` — the `k_import` declaration gains notes.
- Modify: `spec/contract.md` §1 — the `Core_AddTable` signature.
- Modify: `docs/designs/2026-07-28-end-to-end-validation.md`.

**Interfaces:**
- Consumes: everything. **This task runs after every other code task**, per the owner, so the
  note text describes the finished sheet instead of being written against a moving target and
  revised.
- Produces: `Core_AddTable` gains an eighth parameter,
  `Optional ByVal colNotes As String = ""` — pipe-delimited, positionally matching
  `colHeaders`.

**Why a package has to be able to do this.** Three of the four headers are Core literals, so
Core annotates them directly. `Column types` is not: it is package-supplied table metadata at
`PkgImportExport.bas:51`. Core cannot annotate it without knowing which packages exist, which is
the layering rule. So the declaration API carries the note.

**No comment API exists anywhere in this repo** — zero hits for `AddComment`, `.Comment`,
`Validation.InputMessage`. This helper is new.

- [ ] **Step 1: Write the Core note helper**

```vba
' Attach a hover note to a header cell. Clears any existing note first: Setup is
' deleted and rebuilt on every Core_BuildSetup, but Setup Menu is REUSED, and
' AddComment raises on a cell that already carries one.
Private Sub HeaderNote(cell As Range, ByVal txt As String)
    On Error Resume Next
    cell.ClearComments
    On Error GoTo 0
    cell.AddComment txt
    cell.Comment.Shape.TextFrame.AutoSize = True
End Sub
```

`On Error GoTo 0` is correct **here specifically** — `HeaderNote` has no `Fail:` handler of its
own, so there is no handler to disable. Do not copy this line into a procedure that has one.

- [ ] **Step 2: Annotate Availability and Placement in the Action Catalog**

In `RenderPlacementBlock`, after the three header cells are written and bolded. **Do not change
the `AVAIL_HEADER` string itself** — `MarkerVerdict`'s tamper test reads it back
(`Core.bas:63-71`) and rewording it breaks Finalize.

The Availability note explains that every macro starts as `Off`, that `Button` puts it on a
Macros page, and that `Chains only` makes it usable inside a chain but gives it no button of its
own. The Placement note gives all four accepted forms with an example of each: blank for
automatic, `1A` for page 1 column A, `hide` to leave it off the pages, and `TabName!B4` to put the
button on your own sheet at that cell.

- [ ] **Step 3: Annotate Steps and Placement in the Chains block**

In `RenderChainsBlock`. The Steps note must describe the format Task 3 introduced — action ids
separated by commas, copied from the list to the right — and say that the old
`Caption [id]` form still works. This is the field that changed most, so it is the note that
earns its place most clearly.

- [ ] **Step 4: Extend `Core_AddTable` and `RenderTable`**

Add `Optional ByVal colNotes As String = ""` as the final parameter. Store it on the field
dictionary beside `headers`. In `RenderTable`, after each header cell is written, split
`colNotes` on `|` and call `HeaderNote` for any position with a non-empty entry.

Optional and last, so every existing `Core_AddTable` call compiles unchanged. Same treatment for
`Core_AddList` **only if** a list field turns out to need a note — do not add an unused
parameter. Nothing in this phase needs one.

- [ ] **Step 5: Declare the `Column types` note in `PkgImportExport`**

`PkgImportExport.bas:51` declares `k_import` with headers
`"Input tab (in this workbook)|Source file (name starts with...)|Column types"`. Add a
`colNotes` argument whose third entry explains the codes: `T` keeps text exactly as it is,
including IDs and leading zeros; `N` is a number; `D` is a date; `DT` is date and time; `B` is
yes/no; `G` lets Excel decide. One code per column, left to right, separated by commas.

**Say what happens to a code that is not on that list**, because it is invisible today:
`Core_Coerce`'s `Case Else` treats anything unrecognised as `G`, so a typo silently becomes
auto-detect and an ID column can lose its leading zeros with no warning anywhere.

- [ ] **Step 6: Record the API change in contract §1**

The manifest object gains an optional per-column note on table fields. Additive — no existing
declaration changes shape, and a platform without hover notes may render them any way it likes
or not at all. Say that explicitly so the Apps Script port is not obliged to invent tooltips.

- [ ] **Step 7: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c 'AddComment' dist/MacrosBundle.bas
```

Expected: build clean; exactly `1` — the helper, called from four places, not four copies.

- [ ] **Step 8: Append the run-time claims to the validation checklist**

- **Hovering each of the four headers shows a readable note**, not clipped, not empty. Name all
  four: Availability, Placement in the Action Catalog, Steps and Placement in Chains, and
  Column types.
- **Running `Core_BuildSetup` twice in a row leaves one note per header, not two, and does not
  error.** *This is the item that proves `ClearComments` runs* — `AddComment` raises on a cell
  that already has a comment, and the second build is the only thing that exposes it.

**Mark as an assumption:** that a note anchors correctly on a header cell whose row is later
`AutoFit`, given the row directly above every header is merged B:E.

- [ ] **Step 9: Commit**

```bash
git add src/core/Core.bas src/packages/PkgImportExport.bas spec/contract.md docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "feat(core): explain Availability, Placement, Steps and Column types on the Setup headers"
```

---

## Task 9: The validation checklist, and the blind dry-run

**Files:**
- Modify: `docs/designs/2026-07-28-end-to-end-validation.md`.

**Interfaces:**
- Consumes: every earlier task, each of which appended its own items as it went.
- Produces: the finished owner-facing checklist. Task 11's guide must not contradict it.

**This is an assembly task, not a writing task.** Tasks 2a–8 each appended their run-time claims
as they were built. This one reconciles the result, fixes the counts, and puts it in front of a
reader who has never seen the code.

- [ ] **Step 1: Update every count the phase changed**

The document says **43 actions** in items 1.3 and 1.4. It is now **51**. Item 1.4's diagnostic
list ("5 missing is `PkgValidate`, 8 is `PkgFormat`+`PkgExport`") needs its arithmetic redone
against the new totals.

```bash
grep -n '43' docs/designs/2026-07-28-end-to-end-validation.md
```

Read every hit. Some are action counts and must change; some are line references and must not.

- [ ] **Step 2: Re-order by consequence**

The document is ordered *can it destroy your work* → *can it hand you a wrong number* → *does it
behave* → *appearance*. New items must land in the right section, not at the end of the file.
`Complete Macro Setup` and `workbook.toc_links` belong in section 2; `export.email` belongs
there too, because a draft addressed to the wrong person is the same class of harm.

- [ ] **Step 3: Verify every button caption the document names actually exists**

```bash
grep -ohE '\*\*[0-9.]+ — [^*]+\*\*' docs/designs/2026-07-28-end-to-end-validation.md
```

For each caption an item tells the owner to click, confirm it is in the bundle:

```bash
grep -c '"<the exact caption>"' dist/MacrosBundle.bas
```

The previous blind dry-run found **five** buttons the checklist named that do not exist. This is
the cheap mechanical half of that check; Step 5 is the half that needs a reader.

- [ ] **Step 4: Delete any item that would pass on a broken build**

Ask it out loud of every new item: *could this pass no matter what the build did?* An item whose
expected result equals the default state proves nothing and costs the owner time to run. Delete
it rather than keeping it for completeness.

- [ ] **Step 5: Blind dry-run — this is the gate, not a formality**

Dispatch a subagent with:

- `docs/designs/2026-07-28-end-to-end-validation.md`
- `dist/MacrosBundle.bas`
- an explicit prohibition on reading anything under `src/`

Ask it to flag every item it cannot execute (undefined term, missing precondition, unstated
setup), cannot judge (no observable expected result), or that would pass regardless of what the
build did.

**The blindness is the mechanism** — it reproduces the only reader who matters. An agent that
has read `src/` fills the gaps from memory and reports a checklist that works, which is exactly
the failure this catches. One agent costs less than one keyboard round trip with the owner.

Fix everything it finds, then run it again if the fixes were substantial.

- [ ] **Step 6: Commit**

```bash
git add docs/designs/2026-07-28-end-to-end-validation.md
git commit -m "docs: the validation checklist at 51 actions, dry-run blind"
```

---

## Task 10: Documentation sweep

**Files:**
- Modify: `README.md`, `docs/architecture.md`, `docs/macro-catalog-plug-and-play.md`,
  `docs/quick-ref-workflow.md`, `docs/code-map.md`, `CHANGELOG.md`, `VERSION`.

**Interfaces:**
- Consumes: the finished tree. Nothing after this changes behaviour.
- Produces: `VERSION` at `0.18.0`, which Task 12 releases.

**Owner's instruction: give this to a codex agent at medium effort.** It is a sweep across many
files with a clear definition of done, which is what that tool is for. **Review its output
rather than trusting it** — the agent has not watched this phase happen and will not know which
statements were true yesterday and are false today.

- [ ] **Step 1: Find everything the phase invalidated**

```bash
grep -rn '43 action\|43 catalog\|ten packages\|v0\.17\|v2\.9\.0' README.md docs/ --include=*.md
```

Plus the eight `**Later**` markers in `docs/macro-catalog-plug-and-play.md` at lines 48, 79,
139, 161, 183, 186, 208 and 214 — each becomes a built id with its wave recorded, in the same
notation the other built entries use.

- [ ] **Step 2: Brief the codex agent**

Give it the list from Step 1, the eight new ids, the new totals (51 actions, ten packages,
`VERSION` 0.18.0, contract v2.10.0), and these three prohibitions:

- **Do not touch `spec/contract.md`** — it is the frozen contract and Task 1 already amended it.
- **Do not edit `docs/code-map.md` by hand** — it is generated; Step 4 regenerates it.
- **Preserve the owner's `SheetForge` capitalisation** in `README.md`.

- [ ] **Step 3: Review what it produced, then fix what it missed**

Read the diff in full. Two specific things to check, because they are the ones an agent without
the phase's history gets wrong: that no doc still describes `Macros Home` as withdrawn, and that
no doc still describes chain steps as one-per-line.

While you are in `README.md`: the owner's centring uses `<div style="text-align: center;">`.
GitHub strips `style` attributes from markdown, so it does not centre. `<div align="center">`
does.

- [ ] **Step 4: Regenerate the code map**

```bash
node scripts/code-map.mjs
```

Never by hand. It stamps the commit it indexed, and a hand edit makes that stamp a lie — which
is worse than a stale map, because the staleness check then passes.

- [ ] **Step 5: `CHANGELOG.md` and `VERSION`**

`VERSION` to `0.18.0`. A `## [0.18.0]` entry above `## [0.17.0]`, in the existing house style —
which explains *why*, not only *what*. The entries that need a reason rather than a line:
restoring `Macros Home` after it was withdrawn; the chain-step format change and why comma is
now safe; and `Complete Macro Setup` keeping `_Config`.

- [ ] **Step 6: Commit**

```bash
git add README.md docs/ CHANGELOG.md VERSION
git commit -m "docs: bring the documentation up to 51 actions and v0.18.0"
```

---

## Task 11: The non-technical setup guide

**Files:**
- Create: `scripts/reading-level.mjs`, `docs/setup-guide.md`, `docs/setup-guide.pdf`.

**Interfaces:**
- Consumes: every final caption, sheet name and procedure name. It runs after Task 10 because it
  names things that must already be settled.
- Produces: the artifact Task 12 links from the release.

**HARD RULE: no more than a 6th-grade reading level.** Owner's instruction. Measured, not
asserted — Step 1 exists because a document written by someone fluent in the subject reliably
lands three grades above where its author thinks it did.

- [ ] **Step 1: Write the measuring script**

```javascript
// scripts/reading-level.mjs - Flesch-Kincaid grade for a markdown file.
// Strips code fences, inline code, link targets and HTML before measuring:
// a keystroke like Alt+F11 is not prose and must not be scored as a word.
import { readFileSync } from 'node:fs'

const syllables = (w) => {
  w = w.toLowerCase().replace(/[^a-z]/g, '')
  if (w.length <= 3) return 1
  w = w.replace(/(?:[^laeiouy]es|ed|[^laeiouy]e)$/, '').replace(/^y/, '')
  return (w.match(/[aeiouy]{1,2}/g) || ['x']).length
}

const file = process.argv[2]
if (!file) { console.error('usage: node scripts/reading-level.mjs <file.md>'); process.exit(2) }

const prose = readFileSync(file, 'utf8')
  .replace(/```[\s\S]*?```/g, ' ')
  .replace(/`[^`]*`/g, ' ')
  .replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
  .replace(/<[^>]+>/g, ' ')
  .replace(/^[#>|*-]+\s*/gm, ' ')

const sentences = (prose.match(/[.!?]+(?=\s|$)/g) || []).length || 1
const words = prose.split(/\s+/).filter((w) => /[a-z]/i.test(w))
const syl = words.reduce((n, w) => n + syllables(w), 0)
const grade = 0.39 * (words.length / sentences) + 11.8 * (syl / words.length) - 15.59

console.log(`${file}: grade ${grade.toFixed(2)} (${words.length} words, ${sentences} sentences)`)
process.exit(grade > 6.0 ? 1 : 0)
```

It exits non-zero above 6.0, so it is a gate rather than a report.

- [ ] **Step 2: Write the guide**

`docs/setup-guide.md`. Structure:

1. **What this is** — two sentences.
2. **The one thing to know first**, before any step:

   > To change your settings, click **Finalize Setup Tab**. Then click **Create Setup Tab**.
   > That makes the new questions show up. Answer them. Then click **Finalize Setup Tab**
   > again to save them.

3. **The steps.** One action each, with the exact keystroke and the exact on-screen words, and
   what they will see after each one so they can tell it worked:

   - Download the file. Click the link. Click the file whose name ends in `.bas`. Click
     **Download**.
   - Open your Excel file.
   - Press **Alt + F11**. A new window opens.
   - Click **File** in the top left. Click **Import File...**
   - Find the file you downloaded. Click it. Click **Open**.
   - Press **Alt + Q**. You are back in Excel.
   - Press **Alt + F8**. Click **Core_BuildSetup**. Click **Run**.
   - A new tab named **Setup** appears. Fill in the yellow boxes.
   - In the **Availability** column, pick **Button** for each macro you want.
   - Go to the **Setup Menu** tab. Click **Finalize Setup Tab**.
   - Click **Create Setup Tab**. New yellow boxes appear for the macros you picked.
   - Answer them. Click **Finalize Setup Tab** again.
   - Your buttons are on the **Macros 1** tab. **Macros Home** lists every macro and what it
     does.
   - When you are done for good: **Setup Menu** ▸ **Complete Macro Setup**.

4. **Three short asks** — you can ask for a new macro on the Issues page; check back now and
   then for a newer version; if this helped, please star the repo.

**The Alt + F8 snag.** After Task 7 the list shows exactly three names. Say which one to click
and that the other two are for later. Do not write "ignore the others" without saying what they
are — a first-time user reads an unexplained list as a sign they are in the wrong place.

- [ ] **Step 3: Measure, and rewrite until it passes**

```bash
node scripts/reading-level.mjs docs/setup-guide.md
```

Expected: `grade` at or under **6.00**, exit 0. Rewriting to hit it means shorter sentences and
plainer words — not deleting steps. A guide that passes by being incomplete has failed.

- [ ] **Step 4: Blind dry-run the guide**

Same discipline as Task 9. A subagent gets `docs/setup-guide.md` and `dist/MacrosBundle.bas`,
is **forbidden from reading `src/`**, and follows the guide step by step, flagging anything it
cannot do, any name that does not exist, and any step that assumes knowledge the guide has not
given. A setup guide is the artifact where blindness is the entire test.

- [ ] **Step 5: Verify every name in the guide is real**

```bash
for s in "Core_BuildSetup" "Core_FinalizeSetup" "Core_CompleteSetup" \
         "Create Setup Tab" "Finalize Setup Tab" "Complete Macro Setup"; do
  printf '%s: %s\n' "$s" "$(grep -c "\"$s\"" dist/MacrosBundle.bas)"
done
```

Every count must be at least 1. Sheet names `Setup`, `Setup Menu`, `Macros 1` and `Macros Home`
come from `SETUP_SHEET`, `ADMIN_SHEET`, `PANEL_BASE` and `HOME_SHEET` — check those constants
rather than guessing at the rendered names.

- [ ] **Step 6: Generate the PDF**

Use the `/make-pdf` skill against `docs/setup-guide.md`. Check the result opens, that no code
block or keystroke is split across a page break mid-instruction, and that it is legible printed
in black and white — a finance user will print it.

- [ ] **Step 7: Commit**

```bash
git add scripts/reading-level.mjs docs/setup-guide.md docs/setup-guide.pdf
git commit -m "docs: a setup guide for someone who has never opened the VBA editor"
```

---

## Task 12: Sanitize, publish, release

**Files:** repository settings and a GitHub release. No source changes.

**Interfaces:**
- Consumes: everything. **Do not start until Tasks 1–11 are merged.**
- Produces: `MacrosBundle-v0.18.0.bas` as a public release asset.

**Making a repository public cannot be undone.** Clones and forks survive a revert. Steps 1–3
are the reason this is its own task rather than a line at the end of Task 10.

- [ ] **Step 1: Scan the whole history, not the tip**

```bash
gitleaks detect --source . --log-opts="--all" --verbose
```

The pre-commit hook already scans each commit as it lands, so this is the belt to that
suspenders — it covers commits made before the hook existed.

Then read for the things a secret-scanner does not catch: client and franchise identifiers,
internal absolute paths, and anything in `docs/designs/` written on the assumption that only
this team would read it.

```bash
grep -rniE '[A-Za-z]:\\\\|\\\\\\\\[a-z0-9._-]+\\\\|@[a-z0-9.-]+\.(com|net|org)' --include=*.md --include=*.bas --include=*.ps1 --include=*.mjs .
```

Known and fine: three `e.g. C:\Close\January` example paths inside help strings in
`src/packages/`. Anything else needs a decision before Step 3.

- [ ] **Step 2: Confirm the public-facing basics**

- A `LICENSE` file. Check whether one exists; if not, the owner picks — this is not an agent's
  call.
- `README.md` reads to a stranger, not to someone who already knows the project.
- The Issues page is somewhere a non-technical user can post, since the guide sends them there.

- [ ] **Step 3: The owner makes the repository public**

**Owner's action, not the agent's.** Do not run `gh repo edit --visibility public`. Report that
Steps 1 and 2 are clean and stop.

- [ ] **Step 4: Cut the release**

```bash
pwsh build/build.ps1
cp dist/MacrosBundle.bas "dist/MacrosBundle-v0.18.0.bas"
gh release create v0.18.0 "dist/MacrosBundle-v0.18.0.bas" \
  --title "v0.18.0 - Macros Home, 51 actions" --notes-file <(sed -n '/## \[0.18.0\]/,/## \[0.17.0\]/p' CHANGELOG.md | head -n -1)
```

The version in the filename is what lets a user tell two downloads apart — the owner's
requirement. `dist/` stays gitignored: a release asset is not a tracked file, so nothing
generated enters the source tree.

- [ ] **Step 5: Check the link the guide promises**

The guide points at `/releases/latest`, never a pinned version, so it does not go stale on the
next release. Open it as a logged-out user and confirm the asset downloads.

- [ ] **Step 6: Open the PR**

Everything above is one PR against `main`. Record in the body: the deliberate injections from
Tasks 1 and their observed failure text, the build and syntax-check output, and — plainly —
that **nothing in this phase has run in Excel**, with a link to the checklist that discharges
it.
