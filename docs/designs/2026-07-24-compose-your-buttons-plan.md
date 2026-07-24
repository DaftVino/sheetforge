# Compose Your Buttons — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add two linked Setup-page capabilities to SheetForge — user-defined **chains** (one button that runs several actions in sequence) and **grouping/layout** (arrange action+chain buttons onto themed pages) — as an additive contract v1.1.0.

**Architecture:** Everything is keyed on stable ids (`chain.<token>` for chains, `pkg.verb` for actions), never captions. `Core_Run` stays the sole public dispatcher and routes `chain.*` ids to a `Private` chain runner built on a shared `RunAction` that returns a structured result. Chains + per-button placement codes are Core-owned config authored on the Setup sheet (rendered by a dedicated Core hook, not a fake package) and auto-persisted by the existing Finalize sheet-copy. Grouping rebuilds MacroLayout's `1A`-code model inside `Core_BuildControlPanel`; `MacroLayout.bas` is deleted.

**Tech Stack:** Excel VBA (single-module bundle after build). PowerShell build (`build/build.ps1`). No headless test runner.

## Global Constraints

- **Design source of truth:** `docs/designs/2026-07-24-compose-your-buttons-design.md`. Contract semantics: `spec/contract.md` (freeze at **v1.1.0**).
- **Verification model (no TDD):** there is no headless VBA test. Every task's gate is: (1) `pwsh build/build.ps1` exits 0 (it FAILS on a duplicate proc name across modules or a Sub/Function block imbalance), and (2) the task's stated manual-Excel check. Manual checks are run against the imported bundle; they are consolidated in Task 10 and recorded in the PR.
- **Naming:** public entry points `Pkg<Name>_<Verb>`, manifests `Pkg<Name>_Manifest`, Core API `Core_*`, everything else `Private`. Shared helpers live ONLY in Core.
- **VBA hygiene:** `Option Explicit` everywhere; ≤25 line-continuations per statement; every action wraps its work in `Core_SpeedOn`/`Core_SpeedOff` with guaranteed restore; quote every path.
- **No VBIDE trust dependency** — never enumerate modules; discovery stays via `Application.Run` against `_Packages`. The Compose section is Core-owned and must NOT be smuggled into `BuildRegistry()`.
- **`chain.` is a reserved id namespace** — no action id may begin `chain.`.
- **Reference (existing Core, `src/core/Core.bas`):** constants `KEY_COL=1`(A, hidden keys), `VAL_COL=3`(C), `MAX_LIST=20`, `END_KEY="_core_end"`, sheet names `SETUP_SHEET="Setup"`, `CONFIG_SHEET="_Config"`, `MACRO_SHEET="Macros"`. Helpers to reuse: `FindKeyRow(ws,key)`, `NextKeyRow(ws,afterRow)`, `ReadScalar`, `TintRow`, `Yellow`, `SectionHeader`, `BlockIntro`, `CfgSheet()`, `SetupSheet()`, `Core_SpeedOn/Off`, `Core_Fail`, `Core_SheetExists`, `AddBtn`, `BuildRegistry()`, `FindActionOwner(actionId, entry)`, `NewCtx(actionId, reg)`.
- **Finalize persists by copying the whole Setup sheet → hidden `_Config`.** Anything rendered on Setup (including hidden columns) is auto-persisted; readers read `_Config` via `CfgSheet()`.

---

## Task 1: Contract v1.1.0 (spec/contract.md)

**Files:**
- Modify: `spec/contract.md` (header version; §3, §5, §6, §8, §10 additions)

**Interfaces:**
- Produces: the normative vocabulary every later task conforms to — the `chain` config entity, the reserved `chain.*` namespace, engine-owned config, placement-code grammar, and the quiet/dispatch convention.

- [ ] **Step 1: Bump the version and add a v1.1.0 change-control note**

In `spec/contract.md`, change the title/version line to **v1.1.0** and under §10 add:

```markdown
- **v1.1.0 (additive):** adds the `chain` config entity, the reserved `chain.*`
  id namespace, engine-owned config, the grouping/layout placement code, and the
  quiet/dispatch convention. No existing key, action id, kind, value encoding, or
  semantics changed — v1.0.0 deployments remain valid.
```

- [ ] **Step 2: Add the `chain` config entity to §5 (Config storage & read semantics)**

Add a subsection:

```markdown
### Composite actions — chains (v1.1.0)

A **chain** is deployment config (not manifest data) that composes existing
actions into one button:
- **id** — an immutable, opaque token `chain.<token>`, minted once and preserved
  across label edits. NEVER derived live from the label. `chain.*` is reserved:
  no action id may use it.
- **label** — the button caption; freely editable without changing identity.
- **steps** — an ordered list of **action ids** (not captions). Stored as an
  id-list joined by a delimiter that cannot occur in an id (reference: newline).
- **semantics** — run steps in list order; each step's completion dialog is
  suppressed; on a step failure the engine prompts the user Continue/Stop; one
  combined summary is shown at the end. Chains resolve + validate at activate time.
```

- [ ] **Step 3: Add engine-owned config to §5**

```markdown
**Engine-owned config (v1.1.0):** config MAY originate from the engine itself
(chains, placement) via an explicit engine schema hook, not only from package
manifests. It is authored on the same Setup grid with the same key/value model,
and MUST NOT be discovered through the package-probing mechanism (§ discovery).
```

- [ ] **Step 4: Add the placement code + dispatch/quiet convention to §6**

```markdown
**Grouping / layout (v1.1.0):** each button (action or chain) may carry a
**placement code** `<page-number><column-letters>` (e.g. `1A` = page 1, column A;
`12AB` = page 12, column "AB"). Semantics: if NO button anywhere has a code, a
single default button page is built (legacy behavior). If any code is present, a
layout is active: numbered pages are built; a blank-coded button goes to the
default page's **Unassigned** area; the explicit code **`hide`** omits a button.
Same-code buttons stack in authoring order. Core's own admin controls are never
user-placeable.

**Dispatch (v1.1.0):** the single public dispatch entry resolves either an action
id or a reserved `chain.*` id. A **quiet convention** lets composed steps suppress
their **completion** dialogs: a handler reports completion through the engine's
done-sink, which suppresses the dialog when the context is quiet. Suppression is
guaranteed only for completion dialogs routed through that sink; interactive
prompts (e.g. a range picker) and error dialogs are not suppressible.
```

- [ ] **Step 5: Note the reserved namespace in §8's naming rules**

Add to the naming/registry section: `- **Reserved:** the id namespace \`chain.*\` is reserved for engine chains; no package action id may use it.`

- [ ] **Step 6: Commit**

```bash
git add spec/contract.md
git commit -m "spec: extend contract to v1.1.0 (chains, placement, engine config)"
```

**Manual check:** none (doc-only). Re-read §5/§6 for internal consistency with the design doc.

---

## Task 2: Dispatch core — `RunAction`, `Core_Done`, quiet plumbing, `Core_Run` rewrite

**Files:**
- Modify: `src/core/Core.bas` (§4 dispatcher region, ~L601–640)

**Interfaces:**
- Produces:
  - `Private Function RunAction(actionId As String, quiet As Boolean) As Object` — runs ONE action under a fresh child ctx; never dialogs on its own; returns a result `Scripting.Dictionary` with keys `id` (String), `label` (String), `ok` (Boolean), `message` (String), `errNumber` (Long), `errText` (String), `found` (Boolean).
  - `Public Sub Core_Done(ctx As Object, ByVal msg As String, Optional ByVal title As String = "Done")` — completion sink; suppresses the dialog when `ctx("quiet")` is True.
  - `Private Function IsChainId(id As String) As Boolean` — True when `id` begins `chain.` (case-insensitive).
- Consumes: existing `FindActionOwner`, `NewCtx`, `RestoreSpeed` (the private restore behind `Core_SpeedOff`).

- [ ] **Step 1: Add `IsChainId` and `Core_Done` near the dispatcher**

```vba
Private Function IsChainId(ByVal id As String) As Boolean
    IsChainId = (LCase(Left$(id, 6)) = "chain.")
End Function

Private Function IsQuiet(ctx As Object) As Boolean
    On Error Resume Next
    IsQuiet = (ctx.Exists("quiet"))
    If IsQuiet Then IsQuiet = (ctx("quiet") = True)
End Function

' Completion sink. Handlers call this instead of MsgBox for their final summary.
' When the context is quiet (a chain step), the message is captured on the ctx
' and the dialog is suppressed; otherwise it shows exactly as a direct MsgBox.
Public Sub Core_Done(ctx As Object, ByVal msg As String, Optional ByVal title As String = "Done")
    On Error Resume Next
    ctx("doneMsg") = msg
    ctx("done") = True
    On Error GoTo 0
    If Not IsQuiet(ctx) Then MsgBox msg, vbInformation, title
End Sub
```

- [ ] **Step 2: Add `RunAction` (the shared inner dispatch)**

```vba
' Run one action under an isolated child context. Returns a result dictionary;
' never shows its own error/success dialog (callers decide). quiet=True suppresses
' the handler's Core_Done completion dialog and captures its message instead.
Private Function RunAction(ByVal actionId As String, ByVal quiet As Boolean) As Object
    Dim res As Object, reg As Object, ctx As Object, entry As String
    Set res = CreateObject("Scripting.Dictionary")
    res("id") = actionId
    res("label") = actionId
    res("ok") = False
    res("found") = False
    res("message") = ""
    res("errNumber") = 0
    res("errText") = ""

    Set reg = FindActionOwner(actionId, entry)
    If reg Is Nothing Then
        Set RunAction = res            ' found stays False
        Exit Function
    End If
    res("found") = True
    res("label") = ActionCaption(reg, actionId)

    Set ctx = NewCtx(actionId, reg)
    ctx("quiet") = quiet
    ctx("done") = False
    ctx("doneMsg") = ""

    On Error GoTo StepFail
    Application.Run entry, ctx
    On Error GoTo 0
    res("ok") = True
    res("message") = CStr(ctx("doneMsg"))
    Set RunAction = res
    Exit Function

StepFail:
    RestoreSpeed                       ' fail-safe: never leave Excel frozen
    res("ok") = False
    res("errNumber") = Err.Number
    res("errText") = Err.Description
    Set RunAction = res
End Function

' Caption for an action id from its owning manifest (for result labels/summaries).
Private Function ActionCaption(reg As Object, ByVal actionId As String) As String
    Dim a As Object
    ActionCaption = actionId
    For Each a In reg("actions")
        If LCase(CStr(a("id"))) = LCase(actionId) Then
            ActionCaption = CStr(a("caption")): Exit Function
        End If
    Next a
End Function
```

Note: `RestoreSpeed` is the existing private restore invoked by `Core_Run`'s old handler (see current `Core_Run` `Fail:` block, which calls `RestoreSpeed`). Confirm the name by grep; if the private is named differently, call that name.

- [ ] **Step 3: Rewrite `Core_Run` to delegate to `RunAction` (chain routing added in Task 8)**

Replace the body of `Public Sub Core_Run(ByVal actionId As String)` with:

```vba
Public Sub Core_Run(ByVal actionId As String)
    Dim res As Object
    If IsChainId(actionId) Then
        MsgBox "Chains are not available until a configuration is activated.", _
               vbExclamation, "Core"      ' replaced by Core_RunChain routing in Task 8
        Exit Sub
    End If
    Set res = RunAction(actionId, False)
    If Not res("found") Then
        MsgBox "Action not installed: " & actionId, vbExclamation, "Core"
    ElseIf Not res("ok") Then
        MsgBox "Action '" & actionId & "' failed:" & vbCrLf & vbCrLf & _
               "#" & res("errNumber") & "  " & res("errText"), vbExclamation, "Macro Error"
    End If
End Sub
```

- [ ] **Step 4: Build**

Run: `pwsh build/build.ps1`
Expected: exit 0; proc counts increase by the new procs; no duplicate-name/imbalance error.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): shared RunAction dispatch + Core_Done quiet sink"
```

**Manual check (after Task 3 import):** existing single buttons still run and still show their completion dialog (the `quiet=False` path). Deferred to the Task 3 import since packages aren't migrated yet.

---

## Task 3: Migrate package completion dialogs to `Core_Done`

**Files:**
- Modify: `src/packages/PkgImportExport.bas` (each action's final summary `MsgBox`)
- Modify: `src/packages/PkgClean.bas` (each action's final summary `MsgBox`)

**Interfaces:**
- Consumes: `Core_Done(ctx, msg, title)` from Task 2.
- Produces: both bundled packages are chain-conforming (their completion dialogs suppress under quiet).

- [ ] **Step 1: Find every completion dialog in the two packages**

Run: `grep -nE "MsgBox" src/packages/PkgImportExport.bas src/packages/PkgClean.bas`
For each hit, classify: a **completion/summary** dialog (shown at the end of a successful action, reporting what it did) → migrate. A **pre-flight warning** ("No import files are configured") or an **error** dialog → leave as-is (not suppressible per design §2.4).

- [ ] **Step 2: Replace each completion `MsgBox` with `Core_Done`**

For a completion dialog currently written like:

```vba
MsgBox logTxt, vbInformation, "Import Source Data"
```

replace with:

```vba
Core_Done ctx, logTxt, "Import Source Data"
```

Do this for every completion/summary dialog in both files (Import, Update Master, Export in PkgImportExport; and any end-of-run summary in PkgClean's dupes/trim/filldown). Keep the title string identical to today's so standalone UX is unchanged.

- [ ] **Step 3: Build**

Run: `pwsh build/build.ps1`
Expected: exit 0.

- [ ] **Step 4: Commit**

```bash
git add src/packages/PkgImportExport.bas src/packages/PkgClean.bas
git commit -m "refactor(pkg): route completion dialogs through Core_Done"
```

**Manual check (import the bundle):** run each button standalone — each still shows its completion dialog with the same title (proves `quiet=False` unchanged). Pre-flight warnings/errors unchanged.

---

## Task 4: Core-owned Setup section scaffold + hook into `Core_BuildSetup`

**Files:**
- Modify: `src/core/Core.bas` (`Core_BuildSetup` ~L695–716; new §"Core-owned Setup schema" region)

**Interfaces:**
- Produces: `Private Function RenderComposeSection(ws As Worksheet, ByVal r As Long) As Long` — renders the "Compose Your Buttons" banner and returns the next free row. Tasks 5 and 6 fill in the chains block and placement block inside it.
- Consumes: `SectionHeader`, `TintRow`.

- [ ] **Step 1: Add the section scaffold**

In a new region of `Core.bas`:

```vba
'================================================================
' 5b. CORE-OWNED SETUP SCHEMA  (chains + placement; NOT a package)
'   Rendered by Core_BuildSetup outside the package-discovery loop.
'================================================================
Private Const CHAINS_KEY As String = "k_core_chains"   ' block sentinel (col A)
Private Const LAYOUT_KEY As String = "k_core_layout"   ' block sentinel (col A)
Private Const RID_COL    As Long = 10                  ' col J: hidden per-row id

Private Function RenderComposeSection(ws As Worksheet, ByVal r As Long) As Long
    r = SectionHeader(ws, r, "COMPOSE YOUR BUTTONS")
    r = TintRow(ws, r, "Build multi-step buttons (Chains) and arrange buttons " & _
                "onto pages (Page code). Leave everything blank to keep one Macros tab.", _
                10, False, True, RGB(90, 90, 90))
    r = r + 1
    ' Task 5 inserts the Chains block here.
    ' Task 6 inserts the Placement block here.
    RenderComposeSection = r
End Function
```

Confirm `TintRow`'s exact signature by grep before use (`Private Function TintRow(ws, r, txt, size, bold, italic, color)`); match arg order.

- [ ] **Step 2: Call it from `Core_BuildSetup` before the END_KEY sentinel**

In `Core_BuildSetup`, after the `For Each reg In BuildRegistry() ... Next reg` loop and BEFORE the line that writes `ws.Cells(r, KEY_COL).Value = END_KEY`, insert:

```vba
    r = RenderComposeSection(ws, r)
```

so the Compose blocks render after the packages and the END_KEY sentinel bounds the LAST Compose block.

- [ ] **Step 3: Hide the row-id column**

In `Core_BuildSetup`, where column widths/hidden are set (near `ws.Columns("A").Hidden = True`), add:

```vba
    ws.Columns(RID_COL).Hidden = True
```

- [ ] **Step 4: Build**

Run: `pwsh build/build.ps1`
Expected: exit 0.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): Compose Your Buttons Setup section scaffold + hook"
```

**Manual check (import):** run **Create Setup Tab** → the Setup sheet shows a "COMPOSE YOUR BUTTONS" banner + intro line after the package sections; existing sections unchanged; column J hidden.

---

## Task 5: Chains block — render, id mint/preserve, reader, validators

**Files:**
- Modify: `src/core/Core.bas` (`RenderComposeSection`; new chain helpers + reader)

**Interfaces:**
- Produces:
  - `Private Function RenderChainsBlock(ws As Worksheet, ByVal r As Long) As Long`
  - `Private Function ReadChains(ws As Worksheet) As Collection` — each item a `Scripting.Dictionary` `{ id, label, steps (Collection of action-id strings), code }`, read from the block on `ws` (pass `CfgSheet()` at run time or `SetupSheet()` for live validation).
  - `Private Function MintChainId(ByVal label As String, used As Object) As String`
  - `Private Function ParseStepTokens(ByVal cell As String) As Collection` — extracts bracketed action ids from `Label [id]` tokens (newline/`;`-separated).
- Consumes: `FindKeyRow`, `NextKeyRow`, `Yellow`, `BlockIntro`, `MAX_LIST`, `RID_COL`, `CHAINS_KEY`.

- [ ] **Step 1: Render the chains block (in `RenderComposeSection`, replacing the Task-4 placeholder comment)**

```vba
Private Function RenderChainsBlock(ws As Worksheet, ByVal r As Long) As Long
    Dim existing As Collection, ch As Object, i As Long, tokens As String
    r = BlockIntro(ws, r, "Chains  -  one button that runs several actions in order", _
                   "Button name + Steps (paste tokens from the Actions list, one per line) + Page code.")
    ws.Cells(r, KEY_COL).Value = CHAINS_KEY
    ws.Cells(r, VAL_COL).Value = "Button name"
    ws.Cells(r, VAL_COL + 1).Value = "Steps (one token per line)"
    ws.Cells(r, VAL_COL + 2).Value = "Page code"
    ws.Range(ws.Cells(r, VAL_COL), ws.Cells(r, VAL_COL + 2)).Font.Bold = True
    r = r + 1

    Set existing = ReadChains(CfgSheet())          ' re-seed from activated config
    For i = 1 To MAX_LIST
        If i <= existing.Count Then
            Set ch = existing(i)
            ws.Cells(r, VAL_COL).Value = ch("label")
            ws.Cells(r, VAL_COL + 1).Value = StepsToTokens(ch("steps"))   ' Label [id] per line
            ws.Cells(r, VAL_COL + 2).Value = ch("code")
            ws.Cells(r, RID_COL).Value = ch("id")   ' preserve immutable id
        End If
        Yellow ws.Cells(r, VAL_COL)
        Yellow ws.Cells(r, VAL_COL + 1)
        Yellow ws.Cells(r, VAL_COL + 2)
        ws.Cells(r, VAL_COL + 1).WrapText = True
        r = r + 1
    Next i
    ' Copy-ready action tokens in G/H for authoring Steps.
    RenderActionTokens ws, r - MAX_LIST
    RenderChainsBlock = r
End Function

' "Label [id]" per line -> Collection of ids (only the bracketed id is authoritative)
Private Function ParseStepTokens(ByVal cell As String) As Collection
    Dim out As New Collection, lines() As String, ln As String, i As Long
    Dim lo As Long, hi As Long, id As String
    Set ParseStepTokens = out
    cell = Replace(cell, ";", vbLf)
    lines = Split(cell, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim(lines(i))
        If Len(ln) > 0 Then
            lo = InStrRev(ln, "["): hi = InStrRev(ln, "]")
            If lo > 0 And hi > lo Then
                id = Trim(Mid(ln, lo + 1, hi - lo - 1))
            Else
                id = ln            ' bare id with no brackets is tolerated
            End If
            If Len(id) > 0 Then out.Add id
        End If
    Next i
End Function

Private Function StepsToTokens(steps As Collection) As String
    Dim s As String, v As Variant
    For Each v In steps
        s = s & ActionTokenFor(CStr(v)) & vbLf
    Next v
    StepsToTokens = Left$(s, Len(s) - IIf(Len(s) > 0, 1, 0))
End Function
```

- [ ] **Step 2: Add `ActionTokenFor`, `RenderActionTokens`, `MintChainId`**

```vba
' "Import Source Data [importexport.import]" for an action id (falls back to the id).
Private Function ActionTokenFor(ByVal actionId As String) As String
    Dim reg As Object, a As Object
    ActionTokenFor = actionId & " [" & actionId & "]"
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If LCase(CStr(a("id"))) = LCase(actionId) Then
                ActionTokenFor = CStr(a("caption")) & " [" & CStr(a("id")) & "]"
                Exit Function
            End If
        Next a
    Next reg
End Function

Private Sub RenderActionTokens(ws As Worksheet, ByVal startRow As Long)
    Dim reg As Object, a As Object, gr As Long
    ws.Cells(startRow - 1, 7).Value = "Available actions - copy a token into Steps:"
    ws.Cells(startRow - 1, 7).Font.Bold = True
    gr = startRow
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            ws.Cells(gr, 7).Value = CStr(a("caption")) & " [" & CStr(a("id")) & "]"
            ws.Cells(gr, 7).Font.Color = RGB(120, 120, 120)
            gr = gr + 1
        Next a
    Next reg
End Sub

' chain.<slug>, unique within `used` (a dictionary of lowercased ids already taken).
Private Function MintChainId(ByVal label As String, used As Object) As String
    Dim base As String, s As String, ch As String, i As Long, cand As String, n As Long
    For i = 1 To Len(label)
        ch = LCase(Mid(label, i, 1))
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            base = base & ch
        ElseIf Len(base) > 0 And Right$(base, 1) <> "_" Then
            base = base & "_"
        End If
    Next i
    base = Trim(Replace(base, "__", "_"))
    If Len(base) = 0 Then base = "chain"
    cand = "chain." & base
    Do While used.Exists(LCase(cand))
        n = n + 1
        cand = "chain." & base & "_" & n
    Loop
    MintChainId = cand
End Function
```

- [ ] **Step 3: Add `ReadChains` (reads the block from a given sheet)**

```vba
' Read the chains block from ws (CfgSheet at run time; SetupSheet for validation).
' Each item: {id, label, steps(Collection of ids), code}. IDs are read from the
' hidden RID column when present; a row with a label but no id has none yet.
Private Function ReadChains(ws As Worksheet) As Collection
    Dim out As New Collection, kr As Long, nk As Long, i As Long
    Dim label As String, ch As Object
    Set ReadChains = out
    If ws Is Nothing Then Exit Function
    kr = FindKeyRow(ws, CHAINS_KEY)
    If kr = 0 Then Exit Function
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1
        label = Trim(CStr(ws.Cells(i, VAL_COL).Value))
        If Len(label) > 0 Then
            Set ch = CreateObject("Scripting.Dictionary")
            ch("label") = label
            ch("id") = Trim(CStr(ws.Cells(i, RID_COL).Value))
            ch("code") = Trim(CStr(ws.Cells(i, VAL_COL + 2).Value))
            Set ch("steps") = ParseStepTokens(CStr(ws.Cells(i, VAL_COL + 1).Value))
            out.Add ch
        End If
    Next i
End Function
```

- [ ] **Step 4: Build**

Run: `pwsh build/build.ps1`
Expected: exit 0.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): chains Setup block, id mint/preserve, token parse + reader"
```

**Manual check (import):** Create Setup Tab → the Chains block shows Button name / Steps / Page code with yellow cells and a grey "Available actions" token list in column G. (Empty until chains are authored + activated.)

---

## Task 6: Placement block — render (actions), preserve codes, reader

**Files:**
- Modify: `src/core/Core.bas` (`RenderComposeSection`; placement render + reader)

**Interfaces:**
- Produces:
  - `Private Function RenderPlacementBlock(ws As Worksheet, ByVal r As Long) As Long` — one row per registry action; hidden id in `RID_COL`, `caption [id]` display in C (grey, locked feel), yellow Page-code in D; codes preserved by id from `CfgSheet()`.
  - `Private Function ReadCodeMap(ws As Worksheet) As Object` — a `Scripting.Dictionary` mapping **lowercased id → code string** for BOTH actions (placement block) and chains (chains block, `code` column). Empty codes included as `""`.
- Consumes: `ReadChains` (Task 5), `BuildRegistry`, `FindKeyRow`, `NextKeyRow`, `Yellow`, `LAYOUT_KEY`, `RID_COL`.

- [ ] **Step 1: Render the placement (actions) block**

Append inside `RenderComposeSection` after the chains block:

```vba
Private Function RenderPlacementBlock(ws As Worksheet, ByVal r As Long) As Long
    Dim prior As Object, reg As Object, a As Object, id As String
    r = BlockIntro(ws, r, "Actions on pages  -  where each action button goes", _
                   "Page code: number = page, letters = column (e.g. 1A). Blank = page 1. " & _
                   "Type  hide  to leave a button off entirely.")
    ws.Cells(r, KEY_COL).Value = LAYOUT_KEY
    ws.Cells(r, VAL_COL).Value = "Action"
    ws.Cells(r, VAL_COL + 1).Value = "Page code"
    ws.Range(ws.Cells(r, VAL_COL), ws.Cells(r, VAL_COL + 1)).Font.Bold = True
    r = r + 1

    Set prior = ReadCodeMap(CfgSheet())            ' preserve prior codes by id
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            id = CStr(a("id"))
            ws.Cells(r, VAL_COL).Value = CStr(a("caption")) & " [" & id & "]"
            ws.Cells(r, VAL_COL).Font.Color = RGB(90, 90, 90)
            ws.Cells(r, RID_COL).Value = id
            If prior.Exists(LCase(id)) Then ws.Cells(r, VAL_COL + 1).Value = prior(LCase(id))
            Yellow ws.Cells(r, VAL_COL + 1)
            r = r + 1
        Next a
    Next reg
    RenderPlacementBlock = r
End Function
```

Wire both blocks in `RenderComposeSection`:

```vba
    r = RenderChainsBlock(ws, r)
    r = r + 1
    r = RenderPlacementBlock(ws, r)
```

- [ ] **Step 2: Add `ReadCodeMap` (actions from placement block + chains from chains block)**

```vba
' id(lower) -> code, merging the placement (actions) block and the chains block.
Private Function ReadCodeMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String
    Dim chains As Collection, ch As Object
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadCodeMap = map
    If ws Is Nothing Then Exit Function
    ' actions
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr > 0 Then
        nk = NextKeyRow(ws, kr)
        For i = kr + 1 To nk - 1
            id = Trim(CStr(ws.Cells(i, RID_COL).Value))
            If Len(id) > 0 Then map(LCase(id)) = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
        Next i
    End If
    ' chains
    Set chains = ReadChains(ws)
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 Then map(LCase(CStr(ch("id")))) = CStr(ch("code"))
    Next ch
End Function
```

- [ ] **Step 3: Build**

Run: `pwsh build/build.ps1`
Expected: exit 0.

- [ ] **Step 4: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): placement Setup block + id->code reader"
```

**Manual check (import):** Create Setup Tab → an "Actions on pages" block lists every action as `caption [id]` (grey) with a yellow Page-code cell. Enter `1A` for one action, Finalize, Create Setup Tab again → the code is preserved.

---

## Task 7: Finalize — validate chains + placement, review-and-confirm dialog

**Files:**
- Modify: `src/core/Core.bas` (`Core_FinalizeSetup` ~L910–974)

**Interfaces:**
- Produces:
  - `Private Function AssignChainIds(ws As Worksheet) As String` — mints ids into empty `RID_COL` cells of the chains block on the LIVE Setup sheet (so the copy persists them); returns "" or a problem report.
  - `Private Function ComposeReport(ws As Worksheet, ByRef problems As Long) As String` — plain-language review text (resolved chains + steps, unknown steps, reserved-namespace/collision issues, invalid codes, the pages that will be built, Unassigned + hidden buttons).
- Consumes: `ReadChains`, `ReadCodeMap`, `ParseCode` (added in Task 9 — if Task 9 not yet merged, gate the code-validity portion behind `Core_SheetExists`; simplest: implement `ParseCode` in Task 7 as it is pure and Task 9 reuses it). **Decision: add `ParseCode` in this task** (pure helper, no dependency), Task 9 consumes it.

- [ ] **Step 1: Port `ParseCode` + `IndexToCol` from MacroLayout into Core (Private)**

Copy verbatim from `src/packages/MacroLayout.bas` (lines ~358–404) into Core's small-utilities region, as `Private`:

```vba
' "12AB" -> tabNum 12, colIndex for "AB". Returns False on malformed input.
Private Function ParseCode(ByVal code As String, ByRef tabNum As Long, _
                           ByRef colIndex As Long, ByRef colLetter As String) As Boolean
    Dim s As String, i As Long, ch As String, digits As String, letters As String
    tabNum = 0: colIndex = 0: colLetter = ""
    s = UCase(Trim(code))
    If Len(s) = 0 Then Exit Function
    i = 1
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch Else Exit Do
        i = i + 1
    Loop
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch >= "A" And ch <= "Z" Then letters = letters & ch Else Exit Do
        i = i + 1
    Loop
    If i <= Len(s) Then Exit Function
    If Len(digits) = 0 Or Len(letters) = 0 Then Exit Function
    tabNum = CLng(digits)
    colLetter = letters
    For i = 1 To Len(letters)
        colIndex = colIndex * 26 + (Asc(Mid(letters, i, 1)) - Asc("A") + 1)
    Next i
    ParseCode = True
End Function
```

Also port `IndexToCol` (MacroLayout ~L392) and `SortedKeys` (~L406) as `Private` — Task 9 uses them. (Porting them now avoids a forward-reference; the build only fails on *duplicate* names, and MacroLayout is not in the build, so there is no collision.)

- [ ] **Step 2: Add `AssignChainIds` (mint ids into empty hidden cells on live Setup)**

```vba
' Mint chain.<slug> into any chains-block row that has a Button name but no id.
' Runs on the LIVE Setup sheet so Finalize's copy persists the ids. Returns a
' problem string (reserved-namespace or duplicate id) or "".
Private Function AssignChainIds(ws As Worksheet) As String
    Dim kr As Long, nk As Long, i As Long, used As Object, label As String, id As String, rep As String
    Set used = CreateObject("Scripting.Dictionary")
    kr = FindKeyRow(ws, CHAINS_KEY)
    If kr = 0 Then Exit Function
    nk = NextKeyRow(ws, kr)
    ' First pass: register existing ids, flag duplicates.
    For i = kr + 1 To nk - 1
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(id) > 0 Then
            If used.Exists(LCase(id)) Then rep = rep & "- duplicate chain id '" & id & "'." & vbCrLf
            used(LCase(id)) = True
        End If
    Next i
    ' Second pass: mint for named rows without an id.
    For i = kr + 1 To nk - 1
        label = Trim(CStr(ws.Cells(i, VAL_COL).Value))
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(label) > 0 And Len(id) = 0 Then
            id = MintChainId(label, used)
            ws.Cells(i, RID_COL).Value = id
            used(LCase(id)) = True
        End If
    Next i
    AssignChainIds = rep
End Function
```

- [ ] **Step 3: Add `ComposeReport` (validation + human-readable review)**

```vba
' Build the review text + count problems. Validates chain steps against the
' registry, action ids against the reserved namespace, and codes against ParseCode.
Private Function ComposeReport(ws As Worksheet, ByRef problems As Long) As String
    Dim s As String, chains As Collection, ch As Object, ids As Object, stepId As Variant
    Dim tabNum As Long, colIx As Long, colL As String, code As String
    Dim map As Object, k As Variant, unassigned As String, hidden As String
    Set ids = ActionIdSet()                        ' lowercased set of valid action ids
    Set chains = ReadChains(ws)

    s = "CHAINS:" & vbCrLf
    If chains.Count = 0 Then s = s & "  (none)" & vbCrLf
    For Each ch In chains
        s = s & "  - " & ch("label") & "  ->  "
        If ch("steps").Count = 0 Then
            s = s & "(no steps!)": problems = problems + 1
        Else
            For Each stepId In ch("steps")
                s = s & CStr(stepId)
                If Not ids.Exists(LCase(CStr(stepId))) Then
                    s = s & "[UNKNOWN]": problems = problems + 1
                End If
                s = s & "  "
            Next stepId
        End If
        code = CStr(ch("code"))
        If Len(code) > 0 And LCase(code) <> "hide" Then
            If Not ParseCode(code, tabNum, colIx, colL) Then _
                s = s & "  {bad code '" & code & "'}": problems = problems + 1
        End If
        s = s & vbCrLf
    Next ch

    ' Placement codes: validity + collect unassigned/hidden.
    Set map = ReadCodeMap(ws)
    For Each k In map.Keys
        code = CStr(map(k))
        If LCase(code) = "hide" Then
            hidden = hidden & "  - " & k & vbCrLf
        ElseIf Len(code) = 0 Then
            unassigned = unassigned & "  - " & k & vbCrLf
        ElseIf Not ParseCode(code, tabNum, colIx, colL) Then
            s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf: problems = problems + 1
        End If
    Next k

    If AnyCode(map) Then
        s = s & vbCrLf & "A layout is ACTIVE (numbered pages will be built)." & vbCrLf
        If Len(unassigned) > 0 Then s = s & "Unassigned -> page 1:" & vbCrLf & unassigned
        If Len(hidden) > 0 Then s = s & "Hidden (not placed):" & vbCrLf & hidden
    Else
        s = s & vbCrLf & "No page codes -> a single 'Macros' tab (current behavior)." & vbCrLf
    End If
    ComposeReport = s
End Function

Private Function ActionIdSet() As Object
    Dim d As Object, reg As Object, a As Object
    Set d = CreateObject("Scripting.Dictionary")
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            d(LCase(CStr(a("id")))) = True
        Next a
    Next reg
    Set ActionIdSet = d
End Function

Private Function AnyCode(map As Object) As Boolean
    Dim k As Variant
    For Each k In map.Keys
        If Len(Trim(CStr(map(k)))) > 0 Then AnyCode = True: Exit Function
    Next k
End Function
```

- [ ] **Step 4: Wire into `Core_FinalizeSetup`**

In `Core_FinalizeSetup`, AFTER the existing scalar validation builds `rpt`/`warn` and BEFORE the existing `If MsgBox(msg & ... "Activate this configuration now?")` confirm, insert:

```vba
    ' --- Compose Your Buttons: mint ids, validate, and add to the review ---
    Dim composeRep As String, composeProblems As Long, idRep As String
    idRep = AssignChainIds(ws)
    composeRep = ComposeReport(ws, composeProblems)
    msg = msg & vbCrLf & vbCrLf & String(40, "-") & vbCrLf & composeRep
    If Len(idRep) > 0 Then msg = msg & vbCrLf & idRep
    If composeProblems > 0 Then _
        msg = msg & vbCrLf & composeProblems & " problem(s) above - fix or continue."
```

The existing confirm dialog then shows the combined report; on Yes, the existing `ws.Copy → _Config` persists the chains block (with minted ids in column J), the placement codes, everything.

- [ ] **Step 5: Build**

Run: `pwsh build/build.ps1`
Expected: exit 0.

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): Finalize validates + reviews chains and placement"
```

**Manual check (import):** author a chain (name + two tokens + a code), give an action a bad code like `9Z9`; run Finalize → the review lists the chain, its resolved steps, the bad code flagged, unassigned/hidden buttons, and whether a layout is active; Cancel aborts, Yes activates. Rebuild Setup → the chain's minted id persists (edit its label, re-finalize → same id in column J).

---

## Task 8: Chain runner — `Core_RunChain` + `Core_Run` routing

**Files:**
- Modify: `src/core/Core.bas` (`Core_Run` chain branch; new `Core_RunChain`)

**Interfaces:**
- Produces: `Private Sub Core_RunChain(ByVal chainId As String)` — resolves the chain from `CfgSheet()`, runs each step via `RunAction(stepId, True)`, prompts Continue/Stop on failure, shows one summary.
- Consumes: `RunAction` (Task 2), `ReadChains` (Task 5), `Core_SpeedOn/Off`.

- [ ] **Step 1: Replace the Task-2 placeholder chain branch in `Core_Run`**

In `Core_Run`, replace the `If IsChainId(actionId) Then ... Exit Sub` placeholder with:

```vba
    If IsChainId(actionId) Then Core_RunChain actionId: Exit Sub
```

- [ ] **Step 2: Add `Core_RunChain`**

```vba
' Run a chain: each step quiet; prompt Continue/Stop on failure; one end summary.
Private Sub Core_RunChain(ByVal chainId As String)
    Dim chains As Collection, ch As Object, target As Object
    Dim stepId As Variant, res As Object, n As Long, ok As Long, summary As String
    Dim stopped As Boolean

    Set chains = ReadChains(CfgSheet())
    For Each ch In chains
        If LCase(CStr(ch("id"))) = LCase(chainId) Then Set target = ch: Exit For
    Next ch
    If target Is Nothing Then
        MsgBox "Chain not found: " & chainId, vbExclamation, "Core": Exit Sub
    End If

    Core_SpeedOn                        ' one depth-counted wrapper around the whole run
    On Error GoTo Done
    For Each stepId In target("steps")
        n = n + 1
        Set res = RunAction(CStr(stepId), True)
        If Not res("found") Then
            summary = summary & n & ". " & stepId & "  X not installed" & vbCrLf
            If Not AskContinue(n, CStr(stepId), "action not installed") Then stopped = True: Exit For
        ElseIf res("ok") Then
            ok = ok + 1
            summary = summary & n & ". " & res("label") & "  " & Chr(&H2713) & _
                      IIf(Len(res("message")) > 0, "  (" & FirstLine(CStr(res("message"))) & ")", "") & vbCrLf
        Else
            summary = summary & n & ". " & res("label") & "  X  #" & res("errNumber") & _
                      " " & res("errText") & vbCrLf
            ' RunAction already restored speed on the failing step; re-open for the rest.
            If AskContinue(n, CStr(res("label")), CStr(res("errText"))) Then
                Core_SpeedOn
            Else
                stopped = True: Exit For
            End If
        End If
    Next stepId
Done:
    Core_SpeedOff
    Dim head As String
    head = target("label") & IIf(stopped, "  - stopped at step " & n, "  - done") & _
           "  (" & ok & "/" & target("steps").Count & " ok)"
    MsgBox head & vbCrLf & vbCrLf & summary, _
           IIf(stopped, vbExclamation, vbInformation), "Chain"
End Sub

' "Step N <label> failed: <reason>. Continue with the rest?"  Yes=continue.
Private Function AskContinue(ByVal n As Long, ByVal label As String, ByVal reason As String) As Boolean
    AskContinue = (MsgBox("Step " & n & "  '" & label & "'  failed:" & vbCrLf & vbCrLf & _
        reason & vbCrLf & vbCrLf & "Continue with the remaining steps?", _
        vbExclamation + vbYesNo, "Chain - step failed") = vbYes)
End Function

Private Function FirstLine(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, vbCr): If p = 0 Then p = InStr(s, vbLf)
    FirstLine = IIf(p > 0, Left$(s, p - 1), s)
End Function
```

Note on speed nesting: `RunAction` calls `RestoreSpeed` only on a *failing* step (its `StepFail` path), which resets the depth-counted wrapper; the `Core_SpeedOn` after a Continue re-opens it. Successful steps call their own `Core_SpeedOn/Off` which nest under the chain's outer wrapper and do not prematurely restore. Confirm `Core_SpeedOn/Off` are depth-counted (they are: `mSpeedDepth`).

- [ ] **Step 3: Build**

Run: `pwsh build/build.ps1`
Expected: exit 0.

- [ ] **Step 4: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): Core_RunChain sequential runner with failure prompt"
```

**Manual check:** deferred to Task 9 (chains have no button until the panel builds one). Can be smoke-tested via the Immediate window: `Core_Run "chain.<your id>"` after activating a chain.

---

## Task 9: Grouping builder — rewrite `Core_BuildControlPanel`

**Files:**
- Modify: `src/core/Core.bas` (`Core_BuildControlPanel` ~L979–1065; `PanelEntries`)

**Interfaces:**
- Consumes: `ReadCodeMap`, `ReadChains`, `AnyCode`, `ParseCode`, `IndexToCol`, `SortedKeys`, `AddBtn`, `BuildRegistry`.
- Produces: the panel builder — single legacy "Macros" tab when no codes, else numbered `Macros N` pages with an Unassigned area (page 1), `hide` omission, and a fixed admin area.

- [ ] **Step 1: Add a button-gathering helper (actions + chains + admin)**

```vba
' Each item: Array(id, caption, onAction, help). Actions from the registry,
' chains from activated config, then Core's admin buttons (id "" = admin).
Private Function GatherButtons() As Collection
    Dim out As New Collection, reg As Object, a As Object, chains As Collection, ch As Object
    Dim onAct As String
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            onAct = "'Core_Run " & Chr(34) & CStr(a("id")) & Chr(34) & "'"
            out.Add Array(CStr(a("id")), CStr(a("caption")), onAct, CStr(a("help")))
        Next a
    Next reg
    Set chains = ReadChains(CfgSheet())
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 Then
            onAct = "'Core_Run " & Chr(34) & CStr(ch("id")) & Chr(34) & "'"
            out.Add Array(CStr(ch("id")), CStr(ch("label")), onAct, _
                          "Chain: " & StepsInline(ch("steps")))
        End If
    Next ch
    Set GatherButtons = out
End Function

Private Function StepsInline(steps As Collection) As String
    Dim s As String, v As Variant
    For Each v In steps
        s = s & IIf(Len(s) > 0, " > ", "") & CStr(v)
    Next v
    StepsInline = s
End Function
```

- [ ] **Step 2: Rewrite `Core_BuildControlPanel` to branch legacy vs grouped**

Replace the button-placement portion of `Core_BuildControlPanel`. Keep the existing sheet-clear/title/width setup for the legacy path; factor the two modes:

```vba
Public Sub Core_BuildControlPanel()
    Dim buttons As Collection, codeMap As Object
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set buttons = GatherButtons()
    Set codeMap = ReadCodeMap(CfgSheet())
    DeletePanels                        ' remove any prior "Macros N" pages (ported helper)
    If AnyCode(codeMap) Then
        BuildGroupedPanels buttons, codeMap
    Else
        BuildSinglePanel buttons        ' the legacy single "Macros" tab
    End If
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Core_BuildControlPanel error: " & Err.Description, vbExclamation, "Setup"
End Sub
```

- [ ] **Step 3: `BuildSinglePanel` — today's behavior, now also placing chain buttons + admin**

Move the current `Core_BuildControlPanel` body (sheet setup, title, the `PanelEntries` loop, admin buttons) into `Private Sub BuildSinglePanel(buttons As Collection)`, iterating `buttons` (id/caption/onAction/help) instead of `PanelEntries`, and appending the two admin buttons exactly as `PanelEntries` does today:

```vba
Private Sub BuildSinglePanel(buttons As Collection)
    Dim ws As Worksheet, e As Variant, i As Long, leftPos As Double, topPos As Double
    Dim btnW As Double, btnH As Double, gap As Double, helpRow As Long
    Set ws = EnsurePanelSheet(MACRO_SHEET, True)     ' clear+format (extract from current code)
    ' ...title rows B2/B3 exactly as today...
    btnW = 220: btnH = 32: gap = 13
    leftPos = ws.Range("B5").Left: topPos = ws.Range("B5").Top
    helpRow = 5: i = 0
    For Each e In buttons
        i = i + 1
        AddBtn ws, "btn" & i, CStr(e(1)), CStr(e(2)), leftPos, topPos, btnW, btnH
        With ws.Range(ws.Cells(helpRow, 3), ws.Cells(helpRow + 1, 3))
            .Merge: .Value = CStr(e(3)): .Font.Size = 10: .WrapText = True: .VerticalAlignment = xlCenter
        End With
        topPos = topPos + btnH + gap: helpRow = helpRow + 3
    Next e
    AddAdminButtons ws, topPos          ' the two Core_BuildSetup/Core_FinalizeSetup buttons
    ws.Activate: ws.Range("A1").Select
    MsgBox "Macros tab built with " & buttons.Count & " button(s).", vbInformation, "Setup Complete"
End Sub
```

Extract the sheet clear/format into `EnsurePanelSheet(name, isMain)` and the admin buttons into `AddAdminButtons ws, topPos` (using the two `entries.Add Array("Create Setup Tab", ...)` / `("Finalize Setup Tab", ...)` definitions currently in `PanelEntries`). Delete `PanelEntries` once `BuildSinglePanel`/`GatherButtons` replace it.

- [ ] **Step 4: `BuildGroupedPanels` — numbered pages by code, Unassigned + hide + admin**

```vba
' Place buttons onto "Macros N" pages by code. Blank code -> page 1 Unassigned
' column; code "hide" -> omitted. Same-code buttons stack in `buttons` order.
Private Sub BuildGroupedPanels(buttons As Collection, codeMap As Object)
    Dim pages As Object, e As Variant, id As String, code As String
    Dim tabNum As Long, colIx As Long, colL As String, unassigned As New Collection
    Set pages = CreateObject("Scripting.Dictionary")   ' tabNum -> (colIx -> Collection of button-array)

    For Each e In buttons
        id = LCase(CStr(e(0)))
        code = ""
        If codeMap.Exists(id) Then code = Trim(CStr(codeMap(id)))
        If LCase(code) = "hide" Then GoTo NextBtn
        If Len(code) = 0 Then
            unassigned.Add e
        ElseIf ParseCode(code, tabNum, colIx, colL) Then
            AddToPage pages, tabNum, colIx, e
        Else
            unassigned.Add e            ' unparseable code degrades to Unassigned (already flagged at Finalize)
        End If
NextBtn:
    Next e
    ' Unassigned + admin live on page 1.
    Dim uCol As Long
    For Each e In unassigned
        AddToPage pages, 1, 9999, e     ' col 9999 = the Unassigned column, rendered last
    Next e

    Dim tabKeys As Variant, ti As Long, panel As Worksheet, firstPanel As Worksheet
    tabKeys = SortedKeys(pages)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        Set panel = RenderPage(pages, CLng(tabKeys(ti)))
        If firstPanel Is Nothing Then Set firstPanel = panel
        If CLng(tabKeys(ti)) = 1 Then AddAdminButtonsGrouped panel
    Next ti
    If Not firstPanel Is Nothing Then firstPanel.Activate
    MsgBox "Built " & pages.Count & " page(s).", vbInformation, "Setup Complete"
End Sub
```

Add the small helpers `AddToPage pages, tabNum, colIx, e` (nested-dictionary insert, mirroring MacroLayout's `ML_ApplyLayout` read loop) and `RenderPage(pages, tabNum) As Worksheet` (mirror MacroLayout's per-tab build loop: `EnsurePanelSheet("Macros " & tabNum, False)`, iterate `SortedKeys(cols)`, place `AddBtn` per button with the button geometry constants `BTN_W/BTN_H/ROW_GAP/COL_GAP/HDR_H` ported from MacroLayout; when a column index is `9999`, title it "Unassigned" instead of `"Column " & IndexToCol(colIx)`). `AddAdminButtonsGrouped panel` places the two admin buttons in a labeled admin strip at the top-right of page 1.

- [ ] **Step 5: Build**

Run: `pwsh build/build.ps1`
Expected: exit 0. Watch for duplicate-name errors against the ported MacroLayout helpers — they exist only in Core now (MacroLayout deleted in Task 10; if it is still present but not in the build, no collision since `Pkg*.bas` glob excludes it — confirm the glob).

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): grouped page builder (codes, Unassigned, hide, admin)"
```

**Manual check (import):** (a) no codes → one "Macros" tab identical to before, now also showing any chain buttons; click a chain → it runs its steps quietly with one summary, and a mid-chain failure prompts Continue/Stop. (b) Give buttons codes `1A`,`1B`,`2A`, one `hide`, leave one blank → Finalize → Build → pages "Macros 1"/"Macros 2" with columns, the blank in an "Unassigned" area on page 1, the hidden one absent, admin buttons present on page 1.

---

## Task 10: Delete MacroLayout, docs, final verification

**Files:**
- Delete: `src/packages/MacroLayout.bas`
- Modify: `CHANGELOG.md`, `TODOS.md`, `docs/handoff.md`

**Interfaces:** none (cleanup + docs).

- [ ] **Step 1: Delete the prototype**

```bash
git rm src/packages/MacroLayout.bas
```

- [ ] **Step 2: Confirm the build excludes/ignores it and stays green**

Run: `pwsh build/build.ps1`
Expected: exit 0. (The build globs `src/packages/Pkg*.bas`; `MacroLayout.bas` was never matched. Its ported helpers now live only in Core.)

- [ ] **Step 3: Update `CHANGELOG.md`**

Under `## [Unreleased]` add an `### Added` entry describing Compose Your Buttons (chains + grouping), contract v1.1.0, and the MacroLayout removal.

- [ ] **Step 4: Update `TODOS.md` and `docs/handoff.md`**

Mark handoff **step 7** done; move "Compose your buttons" from Next-active to Done; note the deferred built-page polish (design §8) as a new roadmap item.

- [ ] **Step 5: Full manual integration run (record in PR)**

Import `dist/MacrosBundle.bas` into the sample workbook and run the full checklist:
1. Create Setup Tab → Compose section renders (chains block, actions-on-pages block, token reference).
2. Author a chain `Refresh All` = `Clear Input Tabs [importexport.clear]` + `Import Source Data [importexport.import]`, code `1A`; give a couple of actions codes; mark one `hide`; leave one blank.
3. Finalize → review lists resolved chain, pages, unassigned, hidden; activate.
4. Build control panel → grouped pages built; chain button present.
5. Click the chain → runs quietly, one summary; force a failure (rename a source file) → Continue/Stop prompt; verify Excel is responsive at the prompt (speed restored).
6. Clear all codes, re-finalize, rebuild → single "Macros" tab, chain still present.
Record pass/fail per step in the PR.

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "chore: remove MacroLayout prototype; docs for compose-your-buttons v1.1.0"
```

---

## Self-Review (completed against the design)

**Spec coverage:** chaining data model + id immutability (Tasks 1,5,7) · reserved `chain.*` (Tasks 1,2,7) · `Label [id]` token authoring + validation (Tasks 5,7) · quiet suppression convention + package migration (Tasks 2,3) · structured step results + fresh child ctx (Task 2) · depth-aware speed restore before the failure prompt (Tasks 2,8) · engine-owned config via a Core hook, not a fake package (Tasks 4,5,6) · blank→page-1 Unassigned + `hide` (Task 9) · Finalize review-and-confirm (Task 7) · grouping salvage/rewrite + MacroLayout deletion (Tasks 7,9,10) · additive contract v1.1.0 (Task 1). Deferred polish (design §8) is explicitly out (Task 10 roadmap note).

**Placeholder scan:** the two intentional forward-ports (`ParseCode`/`IndexToCol`/`SortedKeys` in Task 7, geometry consts in Task 9) show the code to copy and its source lines — no "TBD". `EnsurePanelSheet`/`AddAdminButtons`/`AddToPage`/`RenderPage` are defined by extraction from named existing code with exact behavior stated.

**Type consistency:** result dictionary keys (`id/label/ok/found/message/errNumber/errText`) are identical in Task 2 (producer) and Task 8 (consumer). `ReadCodeMap` returns lowercased-id→code in Task 6 and is read lowercased in Tasks 7,9. `ReadChains` item shape `{id,label,steps,code}` identical across Tasks 5,7,8,9. Block sentinels `CHAINS_KEY`/`LAYOUT_KEY` and `RID_COL` shared across Tasks 4–9.
