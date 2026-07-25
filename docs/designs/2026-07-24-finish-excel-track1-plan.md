# Finish Excel — Track 1 (built-page polish) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the generated button pages from "disorganized and cluttered at seven
buttons" into a navigable, auto-laid-out index — `Macros Home` plus category pages
with descriptions and navigation — and add per-action enable/disable so a user only
sees the actions and the settings they use.

**Architecture:** Core-only; no package logic changes. Three additive manifest
attributes (`category` on an action, `usedBy` on a field, plus an `Enabled` column
on the existing engine-owned placement block) feed a rewritten panel builder. The
placement block becomes the **Action Catalog** (`Action | Enabled | Page code`);
`GatherButtons` — the single funnel for both actions and chains — filters on
Enabled; the Setup renderer relegates settings belonging only to switched-off
actions into one collapsed outline band at the bottom rather than omitting them.
When no button carries a page code, buttons auto-distribute one page per display
category; a single manual code still overrides everything.

**Tech Stack:** VBA (Excel/Windows), one required `src/core/Core.bas` + thin
`src/packages/Pkg<Name>.bas`. Build: `pwsh build/build.ps1` → `dist/MacrosBundle.bas`.
Syntax gate: `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`. No headless
VBA runner exists — see **Verification model** below.

---

## Global Constraints

Copied from `CLAUDE.md`, `spec/contract.md` and the design's do-not-revert list.
Every task's requirements implicitly include this section.

- **Target version:** repo `VERSION` → `0.4.0`; contract → **v1.2.0** (additive).
- **`src/` is the source of truth.** A `.bas` edit does nothing in Excel until the
  bundle is rebuilt and re-imported. There is no round-trip back to `src/`.
- **Flat namespace, mandatory naming.** Everything `Public` is `Core_*` or
  `Pkg<Name>_*`; everything else `Private`. Shared helpers live ONLY in Core.
  `build.ps1` FAILS on a duplicate procedure name across modules — every new
  private helper name in this plan must be globally unique.
- **`Option Explicit` everywhere.** Max **25 line continuations** per statement.
- **Module-level `Const`/`Dim` stays in the declarations section** at the top of
  `Core.bas` (above `Core_DeclarePackage` at line 70). A declaration after the
  first procedure is a VBA compile error and `build.ps1` fails on it.
- **ASCII only.** `build.ps1` writes UTF-8 **no BOM** on the stated assumption that
  content is ASCII and therefore import-safe. Use `<< Home`, `< Previous`,
  `Next >` — **never** `←`/`→`/`▶`.
- **No VBIDE trust dependency, ever.** Packages are discovered by probing
  `Pkg<Name>_Manifest` via `Application.Run` against `CoreCandidates`.
- **Every action wraps in the speed pattern** with guaranteed restore in the error
  handler. Quote every path.
- **Contract is normative.** Where this plan and `spec/contract.md` disagree, the
  contract wins — which is why Task 1 lands the contract first.

### Do NOT revert (design §5 + the compose-your-buttons handoff log)

1. **Setup COLLAPSES settings for disabled actions; it does NOT omit them.**
   `Core_BuildSetup` deletes and re-seeds Setup from the **finalized** `_Config`, so
   omit-unless-enabled costs toggle → Finalize → rebuild → fill → Finalize *and*
   discards anything typed but not yet finalized. This deliberately overrides all
   three `/deliberate` panelists.
2. **Module-level `Const` stays in the declarations section.**
3. **`_Config` is unhidden before `.Delete`** (`Core_FinalizeSetup`, Core.bas:1488) —
   a very-hidden sheet resists delete, which orphaned `Setup (2)`/`(3)` copies.
4. **Admin controls stay on the `Setup Menu` tab**, never on an action page.
5. **Action ids stay two segments** `<domain>.<operation>`; the domain is NOT the
   module name; the eight shipped ids are grandfathered literally including the two
   camelCase ones; **no alias table**.
6. **No search widget, and no Recent Actions, ever** (Recent Actions needs runtime
   writes the next Finalize overwrites). Favorites stays deferred.

### Verification model (replaces TDD — `test-driven-development` is OFF for this repo)

VBA has no headless runner and Excel cannot be driven on CI, so there is no test to
write first. Every task ends with the same two-command gate plus a commit:

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Both are **structural checks, not a compiler.** The real gate is **Task 9**: the
owner imports the bundle, runs `Debug ▸ Compile VBAProject`, and clicks every
action plus the cancel and error paths. **Never claim this work is verified before
Task 9 passes.** Task 8 builds the `PkgFixture` harness that makes Task 9 exercise
a 75-button layout while only 7 real actions exist.

---

## File Structure

| File | Responsibility | Change |
|---|---|---|
| `spec/contract.md` | Language-neutral source of truth | Modify — v1.1.0 → v1.2.0 |
| `src/core/Core.bas` | Engine: manifest API, Setup renderer, panel builder | Modify — all of Tasks 2–7 |
| `src/packages/PkgFixture.bas` | Throwaway 75-action layout harness | Create (Task 8), delete (Task 10) |
| `build/build.ps1` | Bundle + structural gates | Modify — exclude fixtures from the default bundle |
| `docs/code-map.md` | Symbol index | Regenerate (Task 10) |
| `CHANGELOG.md`, `TODOS.md`, `VERSION` | Release docs | Modify (Task 10) |

`Core.bas` is 1927 lines and already sectioned 1–7. Every change below lands inside
an existing section — no restructuring, and no new file. Read it **in slices via
`docs/code-map.md`**, never whole.

**Line anchors drift.** Every `Core.bas:NNN` in this plan is measured against the
file as of commit `514f6f4` (1927 lines), *before* any task has landed. Tasks 2–7
each insert code, so from Task 3 onward the anchors are approximate. Locate the
target by **procedure name** — grep `docs/code-map.md`, or grep `Core.bas` directly
— and treat the number as a hint about which section you are in. Do not edit by
line number alone.

---

### Task 1: Contract v1.2.0

Lands first because the contract is normative: the implementation tasks below
conform to this text, not the reverse.

**Files:**
- Modify: `spec/contract.md:3` (version header), `:41-70` (§2 field model),
  `:184-224` (§6 action contract), `:228-238` (§7 naming), `:337-355` (§10)

**Interfaces:**
- Consumes: nothing.
- Produces: the normative definitions of `category`, `usedBy`, `enabled`, the
  auto-layout rule, and the reserved `Macros Home` name that Tasks 2–7 implement.

- [ ] **Step 1: Bump the version header**

Replace `spec/contract.md:3-4`:

```markdown
**Contract version: 1.2.0** (2026-07-24) · SemVer, independent of the
repo `VERSION` and of per-package versions.
```

- [ ] **Step 2: Add `usedBy` to the §2 field record**

In the JSON block at `spec/contract.md:46-57`, add one line after `"help"`:

```
  "help":     "string — optional plain-language note shown under the field",
  "usedBy":   "string — optional comma-separated action ids this field configures (v1.2.0)"
```

Then add this bullet to the list that follows (after the `default`/`defaults[]`
bullet at `:68-69`):

```markdown
- **`usedBy` (v1.2.0, optional).** A comma-separated list of action `id`s this
  field configures. A field that declares nothing is **package-wide** and always
  renders in the working area. A field whose every named action is disabled (§6)
  is **relegated, never omitted**: the engine still renders it, still stores it,
  and still falls back to its default — it is simply moved into one collapsed
  "settings for actions you haven't turned on" band at the end of the Setup
  surface. An unrecognised id in `usedBy` counts as in-use, so a typo cannot make
  a field unreachable.
  **Portability note:** collapsing is a presentation affordance (reference: Excel
  row outlining). A platform that cannot collapse MUST still render the band,
  visibly and last. Omitting the fields is non-conforming — it would force a
  second activate cycle and discard un-activated input.
```

- [ ] **Step 3: Add `category` and `enabled` to §6**

In the action table at `spec/contract.md:188-193`, add two rows after `help`:

```markdown
| `category` | Optional display category (v1.2.0). Defaults to the package `name`. Groups actions on the generated index; deliberately decoupled from the package, so two packages may feed one category and one package may span several. |
| `enabled` | Deployment config, not manifest data (v1.2.0): whether this action gets a **button**. Defaults to enabled. |
```

- [ ] **Step 4: Add the enable + auto-layout semantics to §6**

Insert after the "Grouping / layout (v1.1.0)" paragraph (`spec/contract.md:224`):

```markdown
**Per-action enable (v1.2.0):** each action carries a deployment-level `enabled`
flag, authored on the same engine-owned surface as the placement code (§5). It
defaults to **enabled**, so a deployment that predates v1.2.0 is unchanged.

`enabled` controls **whether a button is drawn — not whether the action can run.**
A disabled action:
- draws no button on any page, and is excluded from the index counts;
- **stays listed** on the authoring surface, so it never becomes undiscoverable;
- remains a legal chain step and remains dispatchable by id. Hiding an action's
  button so it is reachable *only* inside a chain is a supported configuration, so
  a chain referencing a disabled action is reported as information, not an error.

**Auto-layout (v1.2.0) — supersedes the "no codes" clause of the v1.1.0 grouping
rule above.** When **no** button anywhere carries a placement code, the engine
distributes buttons automatically: one page per display `category`, in
first-declaration order with the reserved `Chains` category last; each page fills a
column to a fixed maximum before wrapping to the next column, and a category that
exceeds the page's column capacity continues onto a numbered continuation page.
Pages are titled by category. The v1.1.0 behaviour — a single default button page —
is withdrawn: it does not survive contact with a large catalog, which is what this
version exists to support. **When any code is present the v1.1.0 rule is unchanged**
and codes remain a complete override.

**Generated index (v1.2.0):** the engine generates an index surface listing one
entry per built page with its title, its enabled-action count, and a link to it,
plus a link to the admin surface and a legend explaining the placement code and
`enabled`. Every built page carries a persistent link back to the index and links
to its previous/next page. Reference implementation: a `Macros Home` sheet; that
name is **reserved** (§7).
```

- [ ] **Step 5: Reserve the index name in §7**

Add after the `chain.*` reserved bullet (`spec/contract.md:234`):

```markdown
- **Reserved:** the display category `Chains` (engine-generated, holds chain
  buttons) and, in the Excel reference implementation, the sheet name
  `Macros Home`. A platform substitutes its own index surface name.
```

- [ ] **Step 6: Record the change in §10**

Append to `spec/contract.md` after the v1.1.0 bullet (`:351-354`):

```markdown
- **v1.2.0 (additive):** adds the optional action `category` attribute, the
  optional field `usedBy` attribute, the deployment-level per-action `enabled`
  flag, the auto-layout rule for uncoded buttons, the generated index surface with
  its back/previous/next navigation, and the reserved `Chains` category +
  `Macros Home` name. No `key`, action `id`, kind, value encoding, or **action**
  semantics changed, so v1.0.0/v1.1.0 configs remain valid and read as
  all-actions-enabled. The one *behaviour* replaced is engine layout for the
  uncoded case (single page → auto-distributed pages), which §10 does not classify
  as breaking: it renames nothing and re-encodes nothing. Owner-approved in
  `docs/designs/2026-07-24-finish-excel-design.md` §2.2.
```

- [ ] **Step 7: Commit**

```bash
git add spec/contract.md
git commit -m "spec: freeze contract v1.2.0 (category, usedBy, per-action enable, auto-layout)"
```

---

### Task 2: Manifest API — `category` on actions, `usedBy` on fields

Pure additive plumbing: three optional arguments and one shared predicate. Nothing
reads them yet, so this task changes no behaviour — which is exactly why it is its
own reviewable step.

**Files:**
- Modify: `src/core/Core.bas:80-93` (`Core_AddSection`, `Core_AddField`),
  `:95-115` (`Core_AddList`, `Core_AddTable`), `:125-134` (`Core_AddAction`),
  `:136-155` (`NewField`)
- Modify: `src/core/Core.bas:957-967` (`HasAdvanced` — see Task 4; **not** here)

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `Core_AddField(reg, key, label, kind, dflt, [tier], [help], [usedBy])`
  - `Core_AddList(reg, key, label, colHeader, [tier], [help], [usedBy])`
  - `Core_AddTable(reg, key, label, colHeaders, [tier], [help], [usedBy])`
  - `Core_AddAction(reg, actionId, caption, entryPoint, help, [category])`
  - field dictionaries gain a `"usedBy"` key (lowercased, trimmed, possibly `""`)
  - action dictionaries gain a `"category"` key (never empty — defaults to
    `reg("name")`)
  - `Private Function FieldInUse(f As Object, enab As Object) As Boolean`

- [ ] **Step 1: Add `usedBy` to `NewField` and its four callers**

Replace `src/core/Core.bas:136-155` (`NewField`) with:

```vba
Private Function NewField(ByVal key As String, ByVal label As String, _
                          ByVal kind As String, ByVal dflt As String, _
                          ByVal tier As String, ByVal help As String, _
                          ByVal usedBy As String) As Object
    Dim f As Object
    Set f = CreateObject("Scripting.Dictionary")
    f.Add "key", key
    f.Add "label", label
    f.Add "kind", kind
    f.Add "default", dflt
    f.Add "headers", ""
    f.Add "defaults", New Collection
    f.Add "help", help
    ' Comma-separated action ids this field configures; "" = package-wide (always
    ' shown). Stored lowercased so FieldInUse can compare without re-normalizing.
    f.Add "usedBy", LCase(Trim(usedBy))
    ' Normalize the tier so an unknown or mis-cased value degrades to "simple"
    ' (still shown) rather than matching neither render pass and vanishing.
    Dim t As String
    t = LCase(Trim(tier))
    If t <> "advanced" Then t = "simple"
    f.Add "tier", t
    Set NewField = f
End Function
```

- [ ] **Step 2: Thread `usedBy` through the four public field builders**

Replace `src/core/Core.bas:80-115` (`Core_AddSection` through `Core_AddTable`) with:

```vba
Public Sub Core_AddSection(reg As Object, ByVal title As String)
    reg("fields").Add NewField("", title, "section", "", "simple", "", "")
End Sub

' tier is "simple" (default) or "advanced". Advanced fields are grouped under an
' "ADVANCED - most users can leave these as-is" band so regular users can stop
' after the simple fields. help is an optional plain-language note shown in grey
' under the field (pass via help:= to skip tier). usedBy is an optional
' comma-separated list of action ids this field configures; a field that names
' only disabled actions is moved into the collapsed band at the bottom of Setup
' (never omitted - see Core_BuildSetup).
Public Sub Core_AddField(reg As Object, ByVal key As String, ByVal label As String, _
                         ByVal kind As String, ByVal dflt As String, _
                         Optional ByVal tier As String = "simple", _
                         Optional ByVal help As String = "", _
                         Optional ByVal usedBy As String = "")
    reg("fields").Add NewField(key, label, kind, dflt, tier, help, usedBy)
End Sub

Public Sub Core_AddList(reg As Object, ByVal key As String, _
                        ByVal label As String, ByVal colHeader As String, _
                        Optional ByVal tier As String = "simple", _
                        Optional ByVal help As String = "", _
                        Optional ByVal usedBy As String = "")
    Dim f As Object
    Set f = NewField(key, label, "list", "", tier, help, usedBy)
    f("headers") = colHeader
    reg("fields").Add f
End Sub

' colHeaders is a "|"-delimited string (NOT comma - a header may itself contain
' commas, e.g. "Type map (T,N,D,...)"). Up to 3 columns are read (C, D, E).
Public Sub Core_AddTable(reg As Object, ByVal key As String, _
                         ByVal label As String, ByVal colHeaders As String, _
                         Optional ByVal tier As String = "simple", _
                         Optional ByVal help As String = "", _
                         Optional ByVal usedBy As String = "")
    Dim f As Object
    Set f = NewField(key, label, "table", "", tier, help, usedBy)
    f("headers") = colHeaders
    reg("fields").Add f
End Sub
```

Both existing packages call these positionally with at most 7 arguments, so
appending an eighth `Optional` leaves all 10 call sites in `PkgClean.bas` and
`PkgImportExport.bas` working untouched. Do not edit either package in this task.

- [ ] **Step 3: Add `category` to `Core_AddAction`**

Replace `src/core/Core.bas:125-134` with:

```vba
' category is the DISPLAY grouping used by the generated index and the page
' titles. It defaults to the package's display name, so an action that declares
' nothing keeps today's grouping. It is deliberately decoupled from the module:
' two packages may feed one category, and one package may span several, so a
' future package split never moves a button in the user's mental model.
Public Sub Core_AddAction(reg As Object, ByVal actionId As String, ByVal caption As String, _
                          ByVal entryPoint As String, ByVal help As String, _
                          Optional ByVal category As String = "")
    Dim a As Object, cat As String
    Set a = CreateObject("Scripting.Dictionary")
    a.Add "id", actionId
    a.Add "caption", caption
    a.Add "entry", entryPoint
    a.Add "help", help
    cat = Trim(category)
    If Len(cat) = 0 Then cat = CStr(reg("name"))
    a.Add "category", cat
    reg("actions").Add a
End Sub
```

- [ ] **Step 4: Add the `FieldInUse` predicate**

Append at the end of section 1, immediately after `FindField` ends
(`src/core/Core.bas:167`) and **before** the `'====` band that opens section 2:

```vba
' Does this field belong to the user's working area, given the enabled map?
' True when the field is package-wide (no usedBy) or when at least ONE action it
' names is enabled. False means "relegate to the collapsed band" - never "omit":
' Core_BuildSetup deletes and re-seeds Setup from the FINALIZED _Config, so
' omitting a field would cost toggle -> Finalize -> rebuild -> fill -> Finalize
' and would discard anything typed but not yet finalized (design section 2.3).
' An unknown id counts as in-use, so a typo in usedBy cannot hide a field forever.
Private Function FieldInUse(f As Object, enab As Object) As Boolean
    Dim ids() As String, i As Long, one As String
    FieldInUse = True
    If f Is Nothing Then Exit Function
    If Not f.Exists("usedBy") Then Exit Function
    If Len(Trim(CStr(f("usedBy")))) = 0 Then Exit Function
    ids = Split(CStr(f("usedBy")), ",")
    For i = LBound(ids) To UBound(ids)
        one = LCase(Trim(ids(i)))
        If Len(one) > 0 Then
            If Not enab.Exists(one) Then Exit Function     ' unknown id -> in use
            If enab(one) Then Exit Function                ' one enabled is enough
        End If
    Next i
    FieldInUse = False
End Function
```

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: build PASSES (`Core + 2 package(s)`, no duplicate-procedure error, Sub and
Function counts balanced, no module-level-declaration error) and the syntax check
reports no errors. `FieldInUse` is currently uncalled — that is expected and is not
a warning either tool emits.

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): add optional action category and field usedBy to the manifest API"
```

---

### Task 3: The Action Catalog block — per-action enable

Turns the existing placement block into `Action | Enabled | Page code`, adds the
readers, filters `GatherButtons`, and carries the v0.3.0 migration shim. This is
the task most likely to silently destroy a user's saved layout if the shim is
skipped — read Step 3's comment before writing it.

**Files:**
- Modify: `src/core/Core.bas:36-57` (declarations — add `CHAIN_CAT`)
- Modify: `src/core/Core.bas:1242-1266` (`RenderPlacementBlock`)
- Modify: `src/core/Core.bas:1269-1289` (`ReadCodeMap`)
- Modify: `src/core/Core.bas:1549-1567` (`GatherButtons`)
- Modify: `src/core/Core.bas:1318-1396` (`ComposeReport` — soft reporting)

**Interfaces:**
- Consumes: `FieldInUse` (Task 2, not used here), `a("category")` (Task 2).
- Produces:
  - `Private Function ReadEnabledMap(ws As Worksheet) As Object` — id(lower) →
    Boolean; empty dictionary when the block or the Enabled column is absent
  - `Private Function EnabledFor(enab As Object, ByVal id As String) As Boolean` —
    defaults True
  - `Private Function HasEnabledColumn(ws As Worksheet, ByVal keyRow As Long) As Boolean`
  - `Private Function IsYes(ByVal v As String) As Boolean`
  - `GatherButtons` items become **5**-element arrays:
    `Array(id, caption, onAction, help, category)`
  - module constant `CHAIN_CAT As String = "Chains"`

- [ ] **Step 1: Add the `CHAIN_CAT` constant to the declarations section**

In `src/core/Core.bas`, immediately after the `PACKAGES_SHEET` line (`:40`), add:

```vba
Private Const CHAIN_CAT      As String = "Chains"      ' reserved display category for chain buttons
```

This must stay in the declarations section above the first procedure —
`build.ps1` fails on a `Const` after the first `Sub`.

- [ ] **Step 2: Rewrite `RenderPlacementBlock` as the Action Catalog**

Replace `src/core/Core.bas:1240-1266` (the comment block and the whole function)
with:

```vba
' One row per registry action: hidden id in RID_COL, "caption [id]" (grey) in C, a
' yellow Yes/No Enabled cell in D, a yellow Page-code cell in E. Prior answers are
' preserved from CfgSheet() by id. Disabled actions STAY listed here - that is the
' owner's hard requirement: turning a button off must never make the action
' undiscoverable.
Private Function RenderPlacementBlock(ws As Worksheet, ByVal r As Long) As Long
    Dim prior As Object, priorEnab As Object, reg As Object, a As Object, id As String
    r = BlockIntro(ws, r, "Action catalog  -  turn actions on or off, and place them on pages", _
                   "Enabled: Yes draws a button, No hides it (the action stays listed here, and a " & _
                   "Chain can still run it). Page code: leave every code blank for the automatic " & _
                   "layout; fill in even one and codes take over - number = page, letters = " & _
                   "column (e.g. 1A). Type  hide  to leave one button off a manual layout.")
    ws.Cells(r, KEY_COL).Value = LAYOUT_KEY
    ws.Cells(r, VAL_COL).Value = "Action"
    ws.Cells(r, VAL_COL + 1).Value = "Enabled"
    ws.Cells(r, VAL_COL + 2).Value = "Page code"
    ws.Range(ws.Cells(r, VAL_COL), ws.Cells(r, VAL_COL + 2)).Font.Bold = True
    r = r + 1

    Set prior = ReadCodeMap(CfgSheet())            ' preserve prior codes by id
    Set priorEnab = ReadEnabledMap(CfgSheet())     ' preserve prior on/off by id
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            id = CStr(a("id"))
            ws.Cells(r, VAL_COL).Value = CStr(a("caption")) & " [" & id & "]"
            ws.Cells(r, VAL_COL).Font.Color = RGB(90, 90, 90)
            ws.Cells(r, RID_COL).Value = id
            ws.Cells(r, VAL_COL + 1).Value = IIf(EnabledFor(priorEnab, id), "Yes", "No")
            If prior.Exists(LCase(id)) Then ws.Cells(r, VAL_COL + 2).Value = prior(LCase(id))
            Yellow ws.Cells(r, VAL_COL + 1)
            Yellow ws.Cells(r, VAL_COL + 2)
            r = r + 1
        Next a
    Next reg
    RenderPlacementBlock = r
End Function
```

The `Enabled` cell is written as literal `"Yes"`/`"No"` text, **not** as the
contract's `bool` kind — §3 reserves `bool` for its first *field* use, and this is
a column on an engine-owned block, not a field.

- [ ] **Step 3: Add the readers and the v0.3.0 migration shim**

Insert immediately after the new `RenderPlacementBlock` (before the `ReadCodeMap`
comment at what was `:1268`):

```vba
' The v0.3.0 catalog block was "Action | Page code" with the code in VAL_COL + 1.
' v0.4.0 is "Action | Enabled | Page code" with the code in VAL_COL + 2. A _Config
' written by a v0.3.0 Finalize still carries the OLD shape, so the header cell -
' which Finalize copies along with everything else - decides which column holds
' the code. Without this shim, upgrading reads blanks out of the old sheet and the
' next rebuild silently discards the user's entire manual layout.
Private Function HasEnabledColumn(ws As Worksheet, ByVal keyRow As Long) As Boolean
    HasEnabledColumn = (LCase(Trim(CStr(ws.Cells(keyRow, VAL_COL + 1).Value))) = "enabled")
End Function

' Tolerant read of an Enabled cell. Blank is handled by the caller (blank = Yes);
' here anything not recognisably negative counts as Yes, so a user typing "y",
' "TRUE" or "on" gets what they meant instead of a silently hidden button.
Private Function IsYes(ByVal v As String) As Boolean
    Dim s As String
    s = LCase(Trim(v))
    IsYes = Not (s = "no" Or s = "n" Or s = "false" Or s = "0" Or s = "off")
End Function

' id(lower) -> Boolean from the Enabled column of the action catalog block. An id
' with no row, a blank cell, or a pre-v0.4.0 block with no Enabled column at all
' is absent from the map, and EnabledFor treats absence as ENABLED - so upgrading
' a v0.3.0 workbook changes nothing the user sees.
Private Function ReadEnabledMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String, cell As String
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadEnabledMap = map
    If ws Is Nothing Then Exit Function
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr = 0 Then Exit Function
    If Not HasEnabledColumn(ws, kr) Then Exit Function      ' pre-v0.4.0: all enabled
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(id) > 0 Then
            cell = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
            If Len(cell) = 0 Then
                map(LCase(id)) = True
            Else
                map(LCase(id)) = IsYes(cell)
            End If
        End If
    Next i
End Function

' Enabled unless the map says otherwise. Absence means enabled, deliberately: an
' unconfigured or pre-v0.4.0 deployment must show every button.
Private Function EnabledFor(enab As Object, ByVal id As String) As Boolean
    EnabledFor = True
    If enab Is Nothing Then Exit Function
    If enab.Exists(LCase(id)) Then EnabledFor = enab(LCase(id))
End Function
```

- [ ] **Step 4: Shift `ReadCodeMap` onto the code column, via the shim**

Replace `src/core/Core.bas:1269-1289` (`ReadCodeMap`) with:

```vba
' id(lower) -> code, merging the action catalog block and the chains block.
Private Function ReadCodeMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String, codeCol As Long
    Dim chains As Collection, ch As Object
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadCodeMap = map
    If ws Is Nothing Then Exit Function
    ' actions
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr > 0 Then
        ' v0.4.0 inserted Enabled between Action and Page code; an older _Config
        ' still has the code one column left. See HasEnabledColumn.
        If HasEnabledColumn(ws, kr) Then
            codeCol = VAL_COL + 2
        Else
            codeCol = VAL_COL + 1
        End If
        nk = NextKeyRow(ws, kr)
        For i = kr + 1 To nk - 1
            id = Trim(CStr(ws.Cells(i, RID_COL).Value))
            If Len(id) > 0 Then map(LCase(id)) = Trim(CStr(ws.Cells(i, codeCol).Value))
        Next i
    End If
    ' chains
    Set chains = ReadChains(ws)
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 Then map(LCase(CStr(ch("id")))) = CStr(ch("code"))
    Next ch
End Function
```

The chains block is untouched: its Page code was already at `VAL_COL + 2`, which is
now the same column as the catalog's — a happy consequence, not a coincidence to
rely on.

- [ ] **Step 5: Filter `GatherButtons` and carry the category**

Replace `src/core/Core.bas:1547-1567` (the comment and `GatherButtons`) with:

```vba
' Each item: Array(id, caption, onAction, help, category). Actions from the
' registry, chains from activated config. This is the SINGLE funnel for both, so
' the disabled-action filter lives here and nowhere else. Admin buttons are never
' in this list - they live on the Setup Menu tab.
Private Function GatherButtons() As Collection
    Dim out As New Collection, reg As Object, a As Object, chains As Collection, ch As Object
    Dim onAct As String, enab As Object
    Set enab = ReadEnabledMap(CfgSheet())
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If EnabledFor(enab, CStr(a("id"))) Then
                onAct = "'Core_Run " & Chr(34) & CStr(a("id")) & Chr(34) & "'"
                out.Add Array(CStr(a("id")), CStr(a("caption")), onAct, _
                              CStr(a("help")), CStr(a("category")))
            End If
        Next a
    Next reg
    Set chains = ReadChains(CfgSheet())
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 Then
            onAct = "'Core_Run " & Chr(34) & CStr(ch("id")) & Chr(34) & "'"
            out.Add Array(CStr(ch("id")), CStr(ch("label")), onAct, _
                          "Chain: " & StepsInline(ch("steps")), CHAIN_CAT)
        End If
    Next ch
    Set GatherButtons = out
End Function
```

A chain is never filtered by the Enabled column — it has no such column, and a
chain is an explicit user authoring act.

- [ ] **Step 6: Report disabled actions in `ComposeReport` — softly**

In `src/core/Core.bas`, inside `ComposeReport`, add to the `Dim` block (after the
`Dim ch2 As Object, actIds As Object, seenIds As Object, cid As String` line at
`:1323`):

```vba
    Dim enab As Object, offList As String, offInChain As String, nOff As Long
```

Then insert this block immediately **before** the `If AnyCode(map) Then` line
(`:1388`):

```vba
    ' Disabled actions: informational, never structural. Enabled controls whether a
    ' BUTTON is drawn, not whether the action can run - hiding a button so the
    ' action is reachable only inside a chain is a supported configuration, so a
    ' chain step pointing at a disabled action is reported, not blocked.
    Set enab = ReadEnabledMap(ws)
    For Each regX In BuildRegistry()
        For Each aX In regX("actions")
            If Not EnabledFor(enab, CStr(aX("id"))) Then
                nOff = nOff + 1
                offList = offList & "  - " & CStr(aX("caption")) & " [" & CStr(aX("id")) & "]" & vbCrLf
            End If
        Next aX
    Next regX
    For Each ch In chains
        For Each stepId In ch("steps")
            If ids.Exists(LCase(CStr(stepId))) Then
                If Not EnabledFor(enab, CStr(stepId)) Then
                    offInChain = offInChain & "  - " & CStr(ch("label")) & _
                                 " uses " & CStr(stepId) & vbCrLf
                End If
            End If
        Next stepId
    Next ch
    If nOff > 0 Then
        s = s & vbCrLf & "TURNED OFF (" & nOff & " action(s), no button drawn; " & _
            "their settings collapse to the bottom of Setup):" & vbCrLf & offList
    End If
    If Len(offInChain) > 0 Then
        s = s & vbCrLf & "Chains that use a turned-off action (they will still run it):" & _
            vbCrLf & offInChain
    End If
```

`regX`, `aX`, `ch`, `stepId` and `ids` are already declared and already used
earlier in this function; reusing them after their earlier loops is fine.

- [ ] **Step 7: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: PASS on both. `Yellow`, `BlockIntro`, `FindKeyRow`, `NextKeyRow` and
`CfgSheet` are all pre-existing.

- [ ] **Step 8: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): per-action enable on the action catalog block, with v0.3.0 config migration"
```

---

### Task 4: Collapse settings for switched-off actions

The design's most-likely-to-be-wrongly-"fixed" decision. Every field still renders
and still saves; fields belonging only to disabled actions move into one outline
band at the bottom.

**Files:**
- Modify: `src/core/Core.bas:820-916` (`Core_BuildSetup`)
- Modify: `src/core/Core.bas:923-944` (`RenderTier`)
- Modify: `src/core/Core.bas:957-967` (`HasAdvanced`)

**Interfaces:**
- Consumes: `FieldInUse` (Task 2), `ReadEnabledMap` (Task 3), `CfgSheet`,
  `SectionHeader`, `TintRow`, `RenderField` (all pre-existing).
- Produces:
  - `RenderTier(ws, r, reg, tier, showSections, bannerText, enab, wantInUse)`
  - `HasAdvanced(reg, enab)` — now "has an **in-use** advanced field"
  - `Private Function HasAnyTier(reg As Object, enab As Object, ByVal wantInUse As Boolean) As Boolean`
  - `Private Function RenderOffBand(ws, r, enab, ByRef bandFirst As Long, ByRef bandLast As Long) As Long`
  - `Private Sub CollapseRows(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)`

- [ ] **Step 1: Add the in-use filter to `RenderTier`**

Replace `src/core/Core.bas:918-944` (comment + `RenderTier`) with:

```vba
' Render every field of the given tier whose in-use state matches wantInUse. When
' showSections is True, a section header is emitted just before its first matching
' field (so a section whose fields are all in the other tier is not drawn). When
' False (the advanced pass and the off-band pass), section headers are skipped and
' the fields render flat under the one band.
Private Function RenderTier(ws As Worksheet, ByVal r As Long, reg As Object, _
                            ByVal tier As String, ByVal showSections As Boolean, _
                            ByVal bannerText As String, enab As Object, _
                            ByVal wantInUse As Boolean) As Long
    Dim f As Object, pending As String, havePending As Boolean
    For Each f In reg("fields")
        If f("kind") = "section" Then
            ' Skip a section header that just repeats the package banner (a
            ' description identical to the header above it reads as a mistake).
            If UCase(Trim(CStr(f("label")))) <> bannerText Then
                pending = CStr(f("label"))
                havePending = True
            End If
        ElseIf CStr(f("tier")) = tier Then
            If FieldInUse(f, enab) = wantInUse Then
                If showSections And havePending Then
                    r = SectionHeader(ws, r, pending)
                    havePending = False
                End If
                r = RenderField(ws, r, reg, f)
            End If
        End If
    Next f
    RenderTier = r
End Function
```

- [ ] **Step 2: Make `HasAdvanced` in-use-aware and add `HasAnyTier`**

Replace `src/core/Core.bas:957-967` (`HasAdvanced`) with:

```vba
' Is there an advanced field that belongs in the working area? An advanced field
' whose actions are all turned off goes to the off band instead, so drawing the
' ADVANCED header for it would leave an empty band.
Private Function HasAdvanced(reg As Object, enab As Object) As Boolean
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If CStr(f("tier")) = "advanced" Then
                If FieldInUse(f, enab) Then
                    HasAdvanced = True
                    Exit Function
                End If
            End If
        End If
    Next f
End Function

' Does this package have ANY non-section field on the requested side of the
' in-use split? Used to skip both an empty ADVANCED band and an empty off band.
Private Function HasAnyTier(reg As Object, enab As Object, ByVal wantInUse As Boolean) As Boolean
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If FieldInUse(f, enab) = wantInUse Then
                HasAnyTier = True
                Exit Function
            End If
        End If
    Next f
End Function
```

- [ ] **Step 3: Add `RenderOffBand` and `CollapseRows`**

Insert both immediately after the new `HasAnyTier`, before `RenderScalar`
(`src/core/Core.bas:969`):

```vba
' Every field whose usedBy names only turned-off actions, from every package,
' collected under one band at the very bottom of the field area and collapsed with
' Excel's row outline.
'
' They are RENDERED, never omitted. DO NOT "improve" this to skip them:
' Core_BuildSetup deletes and re-seeds Setup from the FINALIZED _Config, so
' omit-unless-enabled costs toggle -> Finalize -> rebuild -> fill -> Finalize and
' discards anything typed but not yet finalized (design section 2.3, which
' deliberately overrides all three /deliberate panelists).
'
' Returns the next free row; reports the collapsible row span through bandFirst /
' bandLast so the caller can group AFTER its AutoFit pass. bandFirst = 0 means
' nothing was relegated and no band was drawn.
Private Function RenderOffBand(ws As Worksheet, ByVal r As Long, enab As Object, _
                               ByRef bandFirst As Long, ByRef bandLast As Long) As Long
    Dim reg As Object, any As Boolean
    bandFirst = 0
    bandLast = 0
    RenderOffBand = r
    For Each reg In BuildRegistry()
        If HasAnyTier(reg, enab, False) Then any = True
    Next reg
    If Not any Then Exit Function              ' nothing off: draw no band at all

    r = SectionHeader(ws, r, "SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON")
    r = TintRow(ws, r, "These belong to actions whose Enabled is No in the action catalog " & _
                "above. They are kept and still saved, so nothing you typed is lost - turn " & _
                "the action back on and they move up into the main list. Click the + in the " & _
                "left margin to open this section.", _
                9, False, True, RGB(90, 70, 120))
    bandFirst = r
    For Each reg In BuildRegistry()
        If HasAnyTier(reg, enab, False) Then
            r = TintRow(ws, r, CStr(reg("name")), 10, True, False, RGB(90, 70, 120))
            r = RenderTier(ws, r, reg, "simple", False, UCase(CStr(reg("name"))), enab, False)
            r = RenderTier(ws, r, reg, "advanced", False, UCase(CStr(reg("name"))), enab, False)
        End If
    Next reg
    bandLast = r - 1
    RenderOffBand = r
End Function

' Group + collapse the off band, with the band's own header row left visible above
' it (SummaryRow = xlAbove). Degrades to a plain VISIBLE band if outlining is
' refused - the fields are rendered and readable by Finalize either way, because
' the config readers key off column A and never care whether a row is hidden.
Private Sub CollapseRows(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    If firstRow < 1 Or lastRow < firstRow Then Exit Sub
    On Error Resume Next
    ws.Outline.SummaryRow = xlAbove
    ws.Rows(firstRow & ":" & lastRow).Group
    ws.Outline.ShowLevels RowLevels:=1
    On Error GoTo 0
End Sub
```

- [ ] **Step 4: Wire it into `Core_BuildSetup`**

In `src/core/Core.bas`, add to `Core_BuildSetup`'s declarations (replacing
`:821-822`):

```vba
    Dim ws As Worksheet, reg As Object, f As Object, r As Long, nPkgs As Long
    Dim firstFieldRow As Long, enab As Object, bandFirst As Long, bandLast As Long
```

Replace the package loop and what follows, `src/core/Core.bas:868-893`, with:

```vba
    ' The enabled map comes from the ACTIVATED config, so a rebuild reflects the
    ' last Finalize. A fresh workbook has no _Config, so the map is empty and
    ' EnabledFor reports everything enabled - the v0.4.0 default.
    Set enab = ReadEnabledMap(CfgSheet())

    firstFieldRow = r
    Dim banner As String
    For Each reg In BuildRegistry()
        nPkgs = nPkgs + 1
        banner = UCase(CStr(reg("name")))
        r = SectionHeader(ws, r, banner)
        r = RenderTier(ws, r, reg, "simple", True, banner, enab, True)   ' simple, in use
        If HasAdvanced(reg, enab) Then
            r = SectionHeader(ws, r, "ADVANCED  -  most users can leave these as-is")
            r = RenderTier(ws, r, reg, "advanced", False, banner, enab, True)
        End If
        r = r + 1
    Next reg

    r = RenderComposeSection(ws, r)

    ' The off band goes BELOW Compose on purpose: the action catalog is where the
    ' user turns things on and off, so it must never sit under a collapsed section.
    r = r + 1
    r = RenderOffBand(ws, r, enab, bandFirst, bandLast)

    ' Sentinel key in column A. The config readers bound a list/table block by
    ' scanning column A for the next non-blank cell; without a trailing key the
    ' LAST list/table on the sheet has no lower bound (End(xlUp) stops at its own
    ' key row, whose value rows are blank in column A) and its values are dropped.
    ' It stays last, after the off band, because the off band may contain the
    ' sheet's final list/table block.
    ws.Cells(r, KEY_COL).Value = END_KEY

    ' Long labels in column B are set to wrap (below); auto-size the field rows
    ' so a wrapped label is fully visible instead of being clipped by the value
    ' cell in column C.
    ws.Rows(firstFieldRow & ":" & r).AutoFit

    ' Group AFTER AutoFit: AutoFit sizes rows and would otherwise fight the
    ' collapse, and a collapsed row's height is irrelevant anyway.
    CollapseRows ws, bandFirst, bandLast
```

Everything from `ws.Columns("A").Font.Color = ...` (`:895`) onward is unchanged.

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: PASS. Confirm the two `RenderTier` call sites in the package loop and the
two in `RenderOffBand` all pass **8** arguments, and that the only `HasAdvanced`
call passes **2** — a missing argument here is a compile error Excel catches but
neither of these tools does.

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): collapse settings belonging only to turned-off actions"
```

---

### Task 5: Panel layout structure — one page model, descriptions, navigation

Replaces the `pages` dictionary with a `layout` dictionary that carries titles and
column headers, tightens the geometry, draws each button's description, and adds
the Home/Previous/Next row. Auto-layout itself is Task 6; this task keeps the
existing manual-code behaviour and only changes how a page is drawn.

**Files:**
- Modify: `src/core/Core.bas:42-48` (geometry constants + `HOME_SHEET`)
- Modify: `src/core/Core.bas:1723-1730` (`AddToPage`)
- Modify: `src/core/Core.bas:1732-1764` (`RenderPage`)
- Modify: `src/core/Core.bas:1636-1665` (`BuildGroupedPanels` → `BuildManualLayout`)
- Modify: `src/core/Core.bas:1517-1536` (`Core_BuildControlPanel`)
- Delete: `src/core/Core.bas:1577-1630` (`BuildSinglePanel`)

**Interfaces:**
- Consumes: 5-element button arrays from `GatherButtons` (Task 3), `AnyCode`,
  `ParseCode`, `IndexToCol`, `SortedKeys`, `EnsurePanelSheet`, `AddBtn`, `AddText`
  (all pre-existing).
- Produces the `layout` structure every later task builds on:
  `layout(tabNum)` → dictionary with
  `"title"` (String), `"cols"` (colIx → Collection of button arrays),
  `"hdrs"` (colIx → String).
  Plus:
  - `Private Sub AddToPage(layout, tabNum, colIx, btn)` — creates the page lazily
  - `Private Sub SetPageTitle(layout, tabNum, title)`
  - `Private Sub SetColHeader(layout, tabNum, colIx, hdr)`
  - `Private Sub BuildManualLayout(buttons, codeMap, layout)` (files buttons only)
  - `Private Sub LabelManualPages(layout)`
  - `Private Sub RenderAllPages(layout)`
  - `Private Function RenderPage(layout, tabNum, prevSheet, nextSheet) As Worksheet`
  - `Private Sub RenderNavRow(ws, prevSheet, nextSheet)`
  - `Private Sub NavLink(cell As Range, caption As String, sheetName As String)`
  - `Private Function PageButtonCount(page As Object) As Long`

- [ ] **Step 1: Tighten the geometry and add `HOME_SHEET`**

Replace `src/core/Core.bas:42-48` with:

```vba
' ---- grouped control-panel geometry (ported from MacroLayout; Task 9) ----
Private Const PANEL_BASE As String = "Macros "   ' grouped output tabs: "Macros 1", "Macros 2"...
Private Const HOME_SHEET As String = "Macros Home"  ' generated index; RESERVED name (contract 7)
Private Const BTN_W   As Double = 210
Private Const BTN_H   As Double = 26             ' tightened in v0.4.0 (was 30)
Private Const DESC_H  As Double = 24             ' description textbox under each button
Private Const ROW_GAP As Double = 4              ' tightened in v0.4.0 (was 8)
Private Const COL_GAP As Double = 16             ' tightened in v0.4.0 (was 24)
Private Const HDR_H   As Double = 20             ' tightened in v0.4.0 (was 22)
Private Const AUTO_COL_MAX   As Long = 8         ' buttons per column before wrapping
Private Const AUTO_PAGE_COLS As Long = 3         ' columns per page before a new page
```

Button pitch becomes `BTN_H + DESC_H + ROW_GAP` = 54pt, so a full 8-button column
is 432pt plus the header — one screen without scrolling.

- [ ] **Step 2: Replace `AddToPage` and add the two setters**

Replace `src/core/Core.bas:1722-1730` (comment + `AddToPage`) with:

```vba
' The layout structure every panel builder fills and RenderPage consumes:
'   layout(tabNum)("title")        page title            (String)
'   layout(tabNum)("cols")(colIx)  buttons in that column (Collection of Arrays)
'   layout(tabNum)("hdrs")(colIx)  that column's header  (String)
' A page is created lazily by the first button filed onto it, so callers must
' file a button BEFORE setting its title or column header.
Private Sub AddToPage(layout As Object, ByVal tabNum As Long, ByVal colIx As Long, ByVal btn As Variant)
    Dim page As Object, cols As Object, items As Collection
    If Not layout.Exists(tabNum) Then
        Set page = CreateObject("Scripting.Dictionary")
        page.Add "title", PANEL_BASE & "- Page " & tabNum
        page.Add "cols", CreateObject("Scripting.Dictionary")
        page.Add "hdrs", CreateObject("Scripting.Dictionary")
        layout.Add tabNum, page
    End If
    Set page = layout(tabNum)
    Set cols = page("cols")
    If Not cols.Exists(colIx) Then cols.Add colIx, New Collection
    Set items = cols(colIx)
    items.Add btn                 ' btn = Array(id, caption, onAction, help, category)
End Sub

' No-ops when the page does not exist yet - file a button first (see AddToPage).
Private Sub SetPageTitle(layout As Object, ByVal tabNum As Long, ByVal title As String)
    If layout.Exists(tabNum) Then layout(tabNum)("title") = title
End Sub

Private Sub SetColHeader(layout As Object, ByVal tabNum As Long, ByVal colIx As Long, ByVal hdr As String)
    If layout.Exists(tabNum) Then layout(tabNum)("hdrs")(colIx) = hdr
End Sub

Private Function PageButtonCount(page As Object) As Long
    Dim cols As Object, k As Variant
    Set cols = page("cols")
    For Each k In cols.Keys
        PageButtonCount = PageButtonCount + cols(k).Count
    Next k
End Function
```

- [ ] **Step 3: Rewrite `RenderPage`, and add the nav row**

Replace `src/core/Core.bas:1732-1764` (comment + `RenderPage`) with:

```vba
' Draw one "Macros N" page: nav row, page title, per-column headers, and the
' stacked buttons each followed by its description.
Private Function RenderPage(layout As Object, ByVal tabNum As Long, _
                            ByVal prevSheet As String, ByVal nextSheet As String) As Worksheet
    Dim panel As Worksheet, page As Object, cols As Object, hdrs As Object
    Dim colKeys As Variant, ci As Long, colIx As Long, items As Collection
    Dim ii As Long, it As Variant
    Dim leftBase As Double, topBase As Double, colLeft As Double, topPos As Double, hdr As String
    Set page = layout(tabNum)
    Set cols = page("cols")
    Set hdrs = page("hdrs")
    Set panel = EnsurePanelSheet(PANEL_BASE & tabNum, False)
    ' Buttons float in points over the grid, so these narrow widths only affect
    ' the nav cells in row 1 - they do not move Range("B5").Left, which depends on
    ' column A alone.
    panel.Columns("B:D").ColumnWidth = 15
    RenderNavRow panel, prevSheet, nextSheet

    AddText panel, panel.Range("B3").Left, panel.Range("B3").Top, 520, 24, _
            CStr(page("title")), 15, True, RGB(31, 78, 121)
    leftBase = panel.Range("B5").Left
    topBase = panel.Range("B5").Top
    colKeys = SortedKeys(cols)
    For ci = LBound(colKeys) To UBound(colKeys)
        colIx = CLng(colKeys(ci))
        colLeft = leftBase + ci * (BTN_W + COL_GAP)          ' ci is 0-based across sorted columns
        Set items = cols(colIx)
        hdr = ""
        If hdrs.Exists(colIx) Then hdr = CStr(hdrs(colIx))
        If Len(hdr) > 0 Then _
            AddText panel, colLeft, topBase, BTN_W, HDR_H, hdr, 11, True, RGB(90, 90, 90)
        topPos = topBase + HDR_H + ROW_GAP
        For ii = 1 To items.Count
            it = items(ii)
            AddBtn panel, "b" & tabNum & "_" & colIx & "_" & ii, CStr(it(1)), CStr(it(2)), _
                   colLeft, topPos, BTN_W, BTN_H
            ' The description goes in a textbox under its own button, not in a cell
            ' as the old single-tab layout did: columns here are positioned in
            ' points, so column 2 onward has no cell to line up with.
            AddText panel, colLeft + 4, topPos + BTN_H, BTN_W - 8, DESC_H, _
                    CStr(it(4)), 8, False, RGB(110, 110, 110)
            topPos = topPos + BTN_H + DESC_H + ROW_GAP
        Next ii
    Next ci
    Set RenderPage = panel
End Function

' Row 1: a persistent Home link in a fixed position on every page, plus Previous /
' Next. Cells rather than textboxes, because a cell hyperlink needs no shape
' bookkeeping and EnsurePanelSheet's Cells.Clear removes it cleanly on rebuild.
' ASCII arrows only - build.ps1 writes UTF-8 with no BOM on the assumption that
' the content is ASCII, so a real arrow glyph would be an import hazard.
Private Sub RenderNavRow(ws As Worksheet, ByVal prevSheet As String, ByVal nextSheet As String)
    NavLink ws.Cells(1, 2), "<< Home", HOME_SHEET
    If Len(prevSheet) > 0 Then NavLink ws.Cells(1, 3), "< Previous", prevSheet
    If Len(nextSheet) > 0 Then NavLink ws.Cells(1, 4), "Next >", nextSheet
    ws.Rows(1).RowHeight = 18
End Sub

' One in-workbook hyperlink cell. The SubAddress sheet name is always quoted,
' because "Macros Home" and "Setup Menu" both contain a space.
Private Sub NavLink(cell As Range, ByVal caption As String, ByVal sheetName As String)
    cell.Parent.Hyperlinks.Add Anchor:=cell, Address:="", _
                               SubAddress:="'" & sheetName & "'!A1", _
                               TextToDisplay:=caption
    cell.Font.Size = 10
    cell.Font.Bold = True
End Sub
```

Note the removal of `panel.Range("A1").Select` from `RenderPage`. `Range.Select`
raises 1004 unless its sheet is active; the old code only survived because
`DeletePanels` forced every panel to be freshly `Worksheets.Add`ed (which
activates). Selection is now done once, on Home, in Task 7.

- [ ] **Step 4: Convert `BuildGroupedPanels` into `BuildManualLayout` + `LabelManualPages`**

Replace `src/core/Core.bas:1632-1665` (comment + `BuildGroupedPanels`) with:

```vba
' Manual layout: file each button onto a "Macros N" page by its code. Blank code
' -> page 1 Unassigned column (9999); "hide" -> omitted; unparseable -> Unassigned
' (already blocked at Finalize). Files only - rendering is RenderAllPages.
Private Sub BuildManualLayout(buttons As Collection, codeMap As Object, layout As Object)
    Dim e As Variant, id As String, code As String
    Dim tabNum As Long, colIx As Long, colL As String, unassigned As New Collection
    For Each e In buttons
        id = LCase(CStr(e(0)))
        code = ""
        If codeMap.Exists(id) Then code = Trim(CStr(codeMap(id)))
        If LCase(code) = "hide" Then GoTo NextBtn
        If Len(code) = 0 Then
            unassigned.Add e
        ElseIf ParseCode(code, tabNum, colIx, colL) Then
            AddToPage layout, tabNum, colIx, e
        Else
            unassigned.Add e            ' unparseable degrades to Unassigned
        End If
NextBtn:
    Next e
    For Each e In unassigned
        AddToPage layout, 1, 9999, e     ' blank-coded buttons go to page 1's Unassigned column
    Next e
End Sub

' Title and head a hand-built layout. Column headers keep the code-based names the
' user chose. A page is titled by category only when EVERY button on it agrees -
' a hand-built page may deliberately mix categories, and guessing a title for a
' mixed page would be worse than the generic one.
Private Sub LabelManualPages(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim cols As Object, colKeys As Variant, ci As Long, colIx As Long
    Dim items As Collection, ii As Long, it As Variant
    Dim cat As String, one As String, mixed As Boolean
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        cat = ""
        mixed = False
        colKeys = SortedKeys(cols)
        For ci = LBound(colKeys) To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            If colIx = 9998 Then
                SetColHeader layout, tabNum, colIx, "Setup"
            ElseIf colIx = 9999 Then
                SetColHeader layout, tabNum, colIx, "Unassigned"
            Else
                SetColHeader layout, tabNum, colIx, "Column " & IndexToCol(colIx)
            End If
            Set items = cols(colIx)
            For ii = 1 To items.Count
                it = items(ii)
                one = CStr(it(4))
                If Len(cat) = 0 Then
                    cat = one
                ElseIf LCase(cat) <> LCase(one) Then
                    mixed = True
                End If
            Next ii
        Next ci
        If Not mixed And Len(cat) > 0 Then SetPageTitle layout, tabNum, cat
    Next ti
End Sub

' Render every page in page order, threading each one's Previous/Next sheet names
' so the nav row is drawn in the same pass.
Private Sub RenderAllPages(layout As Object)
    Dim tabKeys As Variant, ti As Long, prevNm As String, nextNm As String
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        prevNm = ""
        nextNm = ""
        If ti > LBound(tabKeys) Then prevNm = PANEL_BASE & CLng(tabKeys(ti - 1))
        If ti < UBound(tabKeys) Then nextNm = PANEL_BASE & CLng(tabKeys(ti + 1))
        RenderPage layout, CLng(tabKeys(ti)), prevNm, nextNm
    Next ti
End Sub
```

- [ ] **Step 5: Delete `BuildSinglePanel`**

Delete `src/core/Core.bas:1577-1630` in full — the comment `' Legacy layout: one
"Macros" tab, ...` through the `End Sub` after the "Macros tab built with" MsgBox.
The single-tab default is withdrawn by contract v1.2.0, and its only unique
contribution — writing each button's description — is now in `RenderPage`.

Keep `MACRO_SHEET` (`:38`) and `DeleteLegacyPanel` (`:1539-1545`): they are how an
upgrading workbook's stale `Macros` tab gets removed.

- [ ] **Step 6: Rewire `Core_BuildControlPanel`**

Replace `src/core/Core.bas:1513-1536` (comment + `Core_BuildControlPanel`) with:

```vba
' Auto-layout by display category when no button carries a page code; the manual
' code layout when any does (codes are a complete override). Prior panels and the
' pre-v0.4.0 single "Macros" tab are cleared before either mode draws.
' silent:=True suppresses the "built N ..." confirmation (used when Finalize
' rebuilds the panel itself and shows its own combined message).
Public Sub Core_BuildControlPanel(Optional ByVal silent As Boolean = False)
    Dim buttons As Collection, codeMap As Object, layout As Object
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set buttons = GatherButtons()
    Set codeMap = ReadCodeMap(CfgSheet())
    Set layout = CreateObject("Scripting.Dictionary")
    DeletePanels                         ' remove any prior "Macros N" pages
    DeleteLegacyPanel                    ' v0.4.0: the single "Macros" tab is gone in BOTH modes
    If AnyCode(codeMap) Then
        BuildManualLayout buttons, codeMap, layout
        LabelManualPages layout
    Else
        BuildAutoLayout buttons, layout
    End If
    RenderAllPages layout
    BuildHomePage layout, buttons.Count
    BuildAdminPanel                      ' Create/Finalize controls on their own hideable tab
    Application.ScreenUpdating = True
    If Not silent Then _
        MsgBox "Built " & layout.Count & " macro page(s) with " & buttons.Count & _
               " button(s), plus the '" & HOME_SHEET & "' index.", _
               vbInformation, "Setup Complete"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Core_BuildControlPanel error: " & Err.Description, vbExclamation, "Setup"
End Sub
```

`BuildAutoLayout` (Task 6) and `BuildHomePage` (Task 7) do not exist yet, so **the
build will still pass** (`build.ps1` and the syntax checker are structural and do
not resolve call targets) while Excel would refuse to compile. That is why Task 6
and Task 7 must both land before Task 9's Excel gate — do not run an Excel import
between here and Task 7.

- [ ] **Step 7: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: PASS. `BuildSinglePanel` is gone from the unique-procedure count; Sub and
Function totals must still balance.

- [ ] **Step 8: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): page layout model with titles, descriptions and Home/Prev/Next nav"
```

---

### Task 6: Auto-layout by display category

**Files:**
- Modify: `src/core/Core.bas` — insert after `BuildManualLayout` (Task 5)

**Interfaces:**
- Consumes: 5-element button arrays (Task 3), `AddToPage` / `SetPageTitle` /
  `SetColHeader` (Task 5), `CHAIN_CAT`, `AUTO_COL_MAX`, `AUTO_PAGE_COLS`.
- Produces:
  - `Private Sub BuildAutoLayout(buttons As Collection, layout As Object)` — the
    call `Core_BuildControlPanel` already makes
  - `Private Function CategoryOrder(cats As Object, seen As Collection) As Collection`
  - `Private Function PageTitleFor(ByVal cat As String, ByVal partNum As Long) As String`
  - `Private Function ColHdrFor(ByVal cat As String, ByVal colInPage As Long, ByVal partNum As Long) As String`

- [ ] **Step 1: Write the auto-layout builder**

Insert immediately after `BuildManualLayout`'s `End Sub` and before
`LabelManualPages`:

```vba
' Auto-layout: no button carries a page code, so distribute by display category -
' one page per category, in first-appearance order with the reserved Chains
' category last. Each column fills to AUTO_COL_MAX before wrapping; a category
' needing more than AUTO_COL_MAX * AUTO_PAGE_COLS buttons continues onto a
' numbered continuation page. Replaces the pre-v0.4.0 single-"Macros"-tab default
' (contract v1.2.0), which did not survive contact with a large catalog.
Private Sub BuildAutoLayout(buttons As Collection, layout As Object)
    Dim cats As Object, seen As New Collection, order As Collection
    Dim e As Variant, cat As String, ci As Long, items As Collection, ii As Long
    Dim tabNum As Long, colInPage As Long, inCol As Long, partNum As Long

    Set cats = CreateObject("Scripting.Dictionary")
    For Each e In buttons
        cat = Trim(CStr(e(4)))
        If Len(cat) = 0 Then cat = "Other"       ' defensive: Core_AddAction never emits blank
        If Not cats.Exists(cat) Then
            cats.Add cat, New Collection
            seen.Add cat                         ' first-appearance order, stable
        End If
        cats(cat).Add e
    Next e

    Set order = CategoryOrder(cats, seen)
    tabNum = 0
    For ci = 1 To order.Count
        cat = CStr(order(ci))
        Set items = cats(cat)
        tabNum = tabNum + 1
        partNum = 1
        colInPage = 1
        inCol = 0
        For ii = 1 To items.Count
            If inCol >= AUTO_COL_MAX Then
                inCol = 0
                colInPage = colInPage + 1
                If colInPage > AUTO_PAGE_COLS Then
                    colInPage = 1
                    tabNum = tabNum + 1
                    partNum = partNum + 1
                End If
            End If
            ' File first: AddToPage creates the page, and the two setters no-op on
            ' a page that does not exist yet.
            AddToPage layout, tabNum, colInPage, items(ii)
            SetPageTitle layout, tabNum, PageTitleFor(cat, partNum)
            If inCol = 0 Then SetColHeader layout, tabNum, colInPage, ColHdrFor(cat, colInPage, partNum)
            inCol = inCol + 1
        Next ii
    Next ci
End Sub

' First-appearance order, with the reserved Chains category moved to the end: a
' composed button reads as a summary of the actions above it, so it belongs last.
Private Function CategoryOrder(cats As Object, seen As Collection) As Collection
    Dim out As New Collection, i As Long, cat As String
    Set CategoryOrder = out
    For i = 1 To seen.Count
        cat = CStr(seen(i))
        If LCase(cat) <> LCase(CHAIN_CAT) Then out.Add cat
    Next i
    For i = 1 To seen.Count
        cat = CStr(seen(i))
        If LCase(cat) = LCase(CHAIN_CAT) Then out.Add cat
    Next i
End Function

' "Clean & Standardize" for a category's first page, "Clean & Standardize (2)" for
' each continuation page.
Private Function PageTitleFor(ByVal cat As String, ByVal partNum As Long) As String
    PageTitleFor = cat
    If partNum > 1 Then PageTitleFor = cat & " (" & partNum & ")"
End Function

' The first column of a category's first page is headed with the category. Later
' columns and continuation pages are overflow of the same category, so repeating
' the name there would read as a second, different group.
Private Function ColHdrFor(ByVal cat As String, ByVal colInPage As Long, ByVal partNum As Long) As String
    If colInPage = 1 And partNum = 1 Then
        ColHdrFor = cat
    Else
        ColHdrFor = "continued"
    End If
End Function
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: PASS. `BuildHomePage` is still undefined — Task 7.

- [ ] **Step 3: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): auto-layout uncoded buttons one page per display category"
```

---

### Task 7: The `Macros Home` index

**Files:**
- Modify: `src/core/Core.bas` — insert after `RenderAllPages` (Task 5)

**Interfaces:**
- Consumes: the `layout` structure and `PageButtonCount` / `NavLink` (Task 5),
  `HOME_SHEET`, `ADMIN_SHEET`, `EnsurePanelSheet`, `SortedKeys`.
- Produces:
  - `Private Sub BuildHomePage(layout As Object, ByVal totalButtons As Long)` — the
    call `Core_BuildControlPanel` already makes
  - `Private Function HomeLegend(ws As Worksheet, ByVal r As Long, ByVal totalButtons As Long) As Long`
  - `Private Function HomeLine(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long`

- [ ] **Step 1: Write the index builder**

Insert immediately after `RenderAllPages`'s `End Sub`:

```vba
' The generated index: one card row per built page (title, enabled-action count,
' link), a link to the admin tab, and the legend that explains the page code and
' Enabled to a user who never reads the Setup help text. Built last and activated,
' so any rebuild lands the user here.
Private Sub BuildHomePage(layout As Object, ByVal totalButtons As Long)
    Dim ws As Worksheet, r As Long, tabKeys As Variant, ti As Long, tabNum As Long
    Dim page As Object
    Set ws = EnsurePanelSheet(HOME_SHEET, True)
    ws.Columns("B").ColumnWidth = 42
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 22

    With ws.Range("B2")
        .Value = "Macro Toolkit"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    With ws.Range("B3")
        .Value = "Pick a category below. Every page has a  << Home  link back here."
        .Font.Italic = True
    End With

    r = 5
    ws.Cells(r, 2).Value = "CATEGORY"
    ws.Cells(r, 3).Value = "ACTIONS"
    ws.Cells(r, 4).Value = "GO TO"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
        .Font.Bold = True
        .Interior.Color = RGB(221, 217, 240)
    End With
    r = r + 1

    tabKeys = SortedKeys(layout)
    If UBound(tabKeys) < LBound(tabKeys) Then
        ws.Cells(r, 2).Value = "No actions are turned on yet."
        ws.Cells(r, 2).Font.Italic = True
        r = r + 1
    Else
        For ti = LBound(tabKeys) To UBound(tabKeys)
            tabNum = CLng(tabKeys(ti))
            Set page = layout(tabNum)
            ws.Cells(r, 2).Value = CStr(page("title"))
            ws.Cells(r, 3).Value = PageButtonCount(page)
            NavLink ws.Cells(r, 4), PANEL_BASE & tabNum, PANEL_BASE & tabNum
            r = r + 1
        Next ti
    End If

    r = r + 1
    NavLink ws.Cells(r, 2), "Setup Menu  -  first-time setup / re-configure", ADMIN_SHEET
    r = r + 2
    r = HomeLegend(ws, r, totalButtons)

    ws.Activate                     ' Range.Select needs its sheet active
    ws.Range("A1").Select
End Sub

' The legend. This is the only place the page-code convention and Enabled are
' explained where a daily user will actually see them.
Private Function HomeLegend(ws As Worksheet, ByVal r As Long, ByVal totalButtons As Long) As Long
    ws.Cells(r, 2).Value = "HOW THESE PAGES ARE BUILT"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    r = HomeLine(ws, r, totalButtons & " button(s), laid out automatically - one page per category.")
    r = HomeLine(ws, r, "To take charge of the layout, fill in 'Page code' in the action catalog on the Setup tab:")
    r = HomeLine(ws, r, "    1A = page 1, column A.    12AB = page 12, column AB.    hide = leave that button off.")
    r = HomeLine(ws, r, "As soon as ONE code is filled in, the codes control everything and uncoded buttons land on page 1.")
    r = HomeLine(ws, r, "'Enabled = No' hides a button. The action stays listed on Setup, and a Chain can still run it.")
    HomeLegend = r
End Function

Private Function HomeLine(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Size = 10
    ws.Cells(r, 2).Font.Color = RGB(90, 90, 90)
    HomeLine = r + 1
End Function
```

`EnsurePanelSheet(HOME_SHEET, True)` reuses and clears an existing sheet, so
rebuilding does not multiply Home tabs. `DeletePanels` leaves it alone by
construction: `IsPanelName` requires a numeric tail, and `"Home"` is not numeric —
confirm that still holds before finishing this step.

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: PASS, and this is the first build since Task 5 with **no undefined
call targets left** — every procedure `Core_BuildControlPanel` invokes now exists.

- [ ] **Step 3: Sanity-check the reference count for the deleted single-panel path**

```bash
grep -n "BuildSinglePanel\|MACRO_SHEET" src/core/Core.bas
```

Expected: `BuildSinglePanel` returns **nothing**; `MACRO_SHEET` appears exactly
twice — its `Const` in the declarations section and its use in
`DeleteLegacyPanel`. Any other hit is a leftover reference to the removed path.

- [ ] **Step 4: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): generate the Macros Home index with legend, counts and links"
```

---

### Task 8: `PkgFixture` — 75 no-op actions, excluded from the shipping bundle

Design §2.5's answer to designing a 60-button layout while only 7 actions exist.
Deliberately **more** than 60, with uneven category sizes, so that a full column
(8), a wrap (9), a multi-column page (12) and a page overflow (26 > 8 × 3) are all
exercised in one Excel session.

**Files:**
- Create: `src/packages/PkgFixture.bas`
- Modify: `build/build.ps1:64-67` (default package glob)
- Modify: `src/core/Core.bas:574-579` (`CoreCandidates` — **temporary**, removed in
  Task 10)

**Interfaces:**
- Consumes: `Core_DeclarePackage`, `Core_AddSection`, `Core_AddField`,
  `Core_AddList`, `Core_AddTable` (with the Task 2 `usedBy` argument),
  `Core_AddAction` (with the Task 2 `category` argument), `Core_Done`.
- Produces: `PkgFixture_Manifest(reg)`, `PkgFixture_Noop(ctx)`, and 75 actions
  `fixture.a1` … `fixture.a75` across 9 categories.

- [ ] **Step 1: Exclude fixtures from the default bundle**

Replace `build/build.ps1:64-67` with:

```powershell
} else {
    # Fixture/harness packages are never part of a shipped bundle. They stay
    # excluded from the default glob rather than relying on someone remembering,
    # and are still buildable on purpose via -Packages PkgFixture.
    $excludeByDefault = @('PkgFixture.bas')
    $pkgFiles = Get-ChildItem -LiteralPath $pkgDir -Filter 'Pkg*.bas' |
                Where-Object { $excludeByDefault -notcontains $_.Name } |
                Sort-Object Name | ForEach-Object { $_.FullName }
}
```

- [ ] **Step 2: Write the fixture module**

Create `src/packages/PkgFixture.bas`:

```vba
Attribute VB_Name = "PkgFixture"
Option Explicit

'================================================================
' PkgFixture - TEST HARNESS, NEVER SHIPPED.
'
' Declares 75 no-op actions across 9 display categories so the v0.4.0 auto-layout
' can be exercised at catalog scale while only 7 real actions exist (design
' section 2.5). Category sizes are deliberately uneven so one Excel session
' covers: a part-full column, an exactly-full column (8), a wrap (9), a
' multi-column page (12) and a page overflow (26 > AUTO_COL_MAX * AUTO_PAGE_COLS).
'
' Every action binds to the SAME entry point. Nothing in the contract requires a
' distinct sub per action - the ctx carries actionId - so 75 actions cost one sub.
'
' Excluded from the default bundle by build/build.ps1; build it deliberately with
'   pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
' DELETE this module, its CoreCandidates entry and its build.ps1 exclusion once
' the polish pass is verified.
'================================================================

Public Sub PkgFixture_Manifest(reg As Object)
    Dim cats(1 To 9) As String, counts(1 To 9) As Long
    Dim ci As Long, k As Long, n As Long

    Core_DeclarePackage reg, "fixture", "Fixture (test only)", "0.0.1"

    cats(1) = "Ingest & Import":      counts(1) = 3
    cats(2) = "Clean & Standardize":  counts(2) = 5
    cats(3) = "Validate":             counts(3) = 8
    cats(4) = "Reconcile":            counts(4) = 9
    cats(5) = "Report":               counts(5) = 12
    cats(6) = "Export":               counts(6) = 4
    cats(7) = "Period Close":         counts(7) = 6
    cats(8) = "Admin & Housekeeping": counts(8) = 2
    cats(9) = "Bulk Fixture Set":     counts(9) = 26

    n = 0
    For ci = 1 To 9
        For k = 1 To counts(ci)
            n = n + 1
            Core_AddAction reg, "fixture.a" & n, "Fixture " & n & " - " & cats(ci), _
                           "PkgFixture_Noop", _
                           "No-op fixture action " & n & ", category " & cats(ci) & ".", _
                           cats(ci)
        Next k
    Next ci

    ' Fields that exercise the collapsed off band. fx_* keys cannot collide with
    ' the shipped k_* / clean_* namespaces.
    Core_AddSection reg, "FIXTURE SETTINGS"
    Core_AddField reg, "fx_always", "Always shown (declares no usedBy)", "text", "always", _
                  "simple", "Package-wide: stays in place whatever is turned off."
    Core_AddField reg, "fx_a1", "Used by Fixture 1 only", "text", "one", _
                  "simple", "Collapses into the off band when Fixture 1 is set to No.", _
                  "fixture.a1"
    Core_AddField reg, "fx_a2a3", "Used by Fixture 2 or Fixture 3", "text", "two", _
                  "simple", "Stays up while EITHER action is on; collapses only when both are off.", _
                  "fixture.a2,fixture.a3"
    Core_AddField reg, "fx_typo", "Names an action id that does not exist", "text", "safe", _
                  "simple", "An unknown usedBy id must count as in use, never hide the field.", _
                  "fixture.no_such_action"
    Core_AddField reg, "fx_adv", "Advanced, used by Fixture 4", "int", "7", _
                  "advanced", "An advanced field must move into the off band too.", _
                  "fixture.a4"
    Core_AddList reg, "fx_list", "List used by Fixture 5", "Value", _
                 "simple", "A list block must survive being collapsed.", "fixture.a5"
    Core_AddTable reg, "fx_table", "Table used by Fixture 6", "First|Second|Third", _
                  "simple", "A table block must survive being collapsed.", "fixture.a6"
    Core_AddDefault reg, "fx_list", "seed one"
    Core_AddDefault reg, "fx_list", "seed two"
    Core_AddDefault reg, "fx_table", "a1|b1|c1"
End Sub

' The single entry point for all 75 fixture actions. Routing through Core_Done
' (not MsgBox) means chain-step suppression is exercised too.
Public Sub PkgFixture_Noop(ctx As Object)
    Core_Done ctx, "Fixture action ran: " & CStr(ctx("actionId")), "Fixture"
End Sub
```

- [ ] **Step 3: Register the fixture in Core's candidate list — temporarily**

Replace `src/core/Core.bas:574-579` with:

```vba
Private Function CoreCandidates() As Collection
    Dim c As New Collection
    c.Add "ImportExport"            ' probes PkgImportExport_Manifest
    c.Add "Clean"                   ' probes PkgClean_Manifest
    c.Add "Fixture"                 ' TEMPORARY (v0.4.0 verification) - remove with PkgFixture
    Set CoreCandidates = c
End Function
```

`TryManifest` swallows the error when the module is absent, so this line is inert
in the shipping bundle — but it is still temporary, and Task 10 removes it.

- [ ] **Step 4: Build both bundles and syntax-check both**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
```

Expected: the first build reports `Core + 2 package(s): PkgClean.bas,
PkgImportExport.bas` — **`PkgFixture.bas` must NOT appear**. The second reports
`Core + 3 package(s)` including it. Both pass the duplicate-procedure,
block-balance and declaration-order gates, and both pass the syntax check.
`dist/` is gitignored, so neither bundle is committed.

- [ ] **Step 5: Commit**

```bash
git add src/packages/PkgFixture.bas build/build.ps1 src/core/Core.bas
git commit -m "test: add PkgFixture (75 no-op actions, 9 categories) and keep it out of the default bundle"
```

---

### Task 9: Owner verification in Excel — THE gate

There is no headless VBA compiler and no headless Excel. Everything above is
structurally checked and functionally unverified. **Only the owner can run this
task, and nothing in this plan may be described as verified until it passes.**

**Files:** none — this task produces the PR verification record.

**Interfaces:**
- Consumes: `dist/MacrosBundle.fixture.bas` and `dist/MacrosBundle.bas` (Task 8).
- Produces: a pass/fail record per check, pasted into the PR.

- [ ] **Step 1: Fixture pass — compile**

In a scratch copy of the workbook: import `dist/MacrosBundle.fixture.bas`, then
`Debug ▸ Compile VBAProject`. Expected: compiles clean. A missing-argument or
undefined-name error here is precisely what the structural gates cannot see.

- [ ] **Step 2: Fixture pass — build at scale**

Run `Core_BuildSetup` (Alt+F8), then `Finalize Setup Tab`, and confirm each:

- [ ] `Macros Home` exists, is the first tab, and lists **11 rows** (one per page)
      across **10 categories**, in this order: `Import / Export` (5),
      `Clean & Standardize` (8), `Ingest & Import` (3), `Validate` (8),
      `Reconcile` (9), `Report` (12), `Export` (4), `Period Close` (6),
      `Admin & Housekeeping` (2), `Bulk Fixture Set` (26, over two rows).
- [ ] **The merge works.** `Clean & Standardize` shows **8**, not 3 and 5 on two
      separate rows: `PkgClean`'s 3 real actions and the fixture's 5 declare the
      same category from two different modules and land on one page. This is the
      whole point of decoupling `category` from the package — if they split, Task 2
      Step 3 is defaulting the category wrongly.
- [ ] Action counts on Home match the buttons actually drawn on each page.
- [ ] `Validate` (8 actions) fills exactly one column with no wrap.
- [ ] `Reconcile` (9) wraps to a second column headed `continued`.
- [ ] `Report` (12) uses two columns on one page.
- [ ] `Bulk Fixture Set` (26) overflows to a second page titled
      `Bulk Fixture Set (2)`.
- [ ] Every page shows its category as the title — no `Macros - Page N` anywhere.
- [ ] Every button has its description underneath, unclipped and not overlapping
      the next button.
- [ ] `<< Home` is in the same place on every page and works; `< Previous` /
      `Next >` work and are correctly absent on the first/last page.
- [ ] No `Macros` tab remains (the single-tab default is withdrawn).

- [ ] **Step 3: Fixture pass — enable/disable and the collapsed band**

- [ ] On Setup, the action catalog block reads `Action | Enabled | Page code`
      with every row `Yes`.
- [ ] Set `fixture.a1` to `No`, Finalize. Its button disappears; **its row stays**
      in the catalog; Home's count for `Ingest & Import` drops by one.
- [ ] Run `Create Setup Tab` again. `fx_a1` is now inside a collapsed
      **SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON** band with a working `+`
      outline control, and the value you typed into `fx_a1` **survived**.
- [ ] `fx_typo` (unknown `usedBy` id) is still in the main area, not the band.
- [ ] Set `fixture.a2` to `No` but leave `fixture.a3` on: `fx_a2a3` stays up. Turn
      `fixture.a3` off too: it moves into the band.
- [ ] Turn `fixture.a5` and `fixture.a6` off: the list and table blocks land in
      the band, and after re-Finalizing, their values still read back correctly
      (this is the `_core_end` sentinel still bounding the last block).
- [ ] Turn everything back on: the band disappears entirely rather than showing
      empty.

- [ ] **Step 4: Fixture pass — manual codes still override**

- [ ] Put `1A` on one action and `2B` on another, leave the rest blank, Finalize.
      Numbered pages appear, everything uncoded lands in page 1's `Unassigned`
      column, and the column headers read `Column A` / `Column B` / `Unassigned`.
- [ ] `hide` on one action omits its button while leaving its catalog row.
- [ ] Clear every code and Finalize: auto-layout returns.

- [ ] **Step 5: Fixture pass — chains, cancel and error paths**

- [ ] Build a chain of three fixture actions. It appears on a `Chains` page,
      ordered last. Running it shows one combined summary, not three dialogs.
- [ ] Disable an action that a chain uses. Finalize reports it under *"Chains that
      use a turned-off action"* and **does not block activation**; the chain still
      runs that step.
- [ ] Answer **No** to the "Activate this configuration now?" prompt: nothing is
      rebuilt, no `_Config` is written.
- [ ] Answer **No** to `Create Setup Tab`'s rebuild confirmation: the existing
      Setup tab is untouched.
- [ ] Enter a deliberately bad page code (`0A`, `A1`, `1`): Finalize blocks with a
      `BAD CODE` line and refuses to activate.
- [ ] Run Finalize twice in a row: no *"That name is already taken"*, and no
      `Setup (2)` / `Setup (3)` orphans (the v0.3.0 unhide-before-delete fix).

- [ ] **Step 6: Real pass — the shipping bundle**

In a **fresh** copy of the real workbook: import `dist/MacrosBundle.bas` (no
fixture), `Debug ▸ Compile VBAProject`, `Core_BuildSetup`, Finalize, then:

- [ ] Home lists exactly two categories with 5 and 3 actions.
- [ ] **Click all 8 real actions** and confirm each does its real work.
- [ ] Cancel out of each action that prompts (range pickers, file dialogs) and
      confirm no error dialog and no half-finished state.
- [ ] Force one error path (point a `sheet` field at a missing tab, or a `file`
      field at a missing prefix) and confirm Finalize reports it and that a
      failing action restores `ScreenUpdating`/`Calculation`/`EnableEvents`
      (the next action still runs normally).

- [ ] **Step 7: Upgrade pass — the migration shim**

This is the check that protects an existing user's layout. In a copy of a workbook
**already finalized under v0.3.0** (so its `_Config` has the old
`Action | Page code` block), with manual page codes set:

- [ ] Import `dist/MacrosBundle.bas` over it and run `Core_BuildSetup`.
- [ ] **Every prior page code is still there**, now in the `Page code` column to
      the right of the new `Enabled` column, and every `Enabled` reads `Yes`.

If codes come back blank, `HasEnabledColumn` (Task 3 Step 3) is wrong or is not
being consulted — stop and fix it, do not re-enter the codes by hand.

- [ ] **Step 8: Record the result**

Paste the completed checklist into the PR. If any check fails, stop and fix it
before Task 10 — a failure here means the corresponding task is not done.

---

### Task 10: Tear down the fixture, refresh docs, ship v0.4.0

**Files:**
- Delete: `src/packages/PkgFixture.bas`
- Modify: `src/core/Core.bas:574-580` (remove the `Fixture` candidate)
- Modify: `build/build.ps1` (remove the now-pointless exclusion)
- Modify: `VERSION`, `CHANGELOG.md`, `TODOS.md`
- Regenerate: `docs/code-map.md`

**Interfaces:**
- Consumes: a passing Task 9.
- Produces: the v0.4.0 release commit and PR.

- [ ] **Step 1: Remove the fixture and its two hooks**

```bash
git rm src/packages/PkgFixture.bas
```

In `src/core/Core.bas`, delete the `c.Add "Fixture"` line from `CoreCandidates`.
In `build/build.ps1`, restore the plain default glob (the exclusion has nothing
left to exclude, and a dangling filter invites a future fixture to be added
silently):

```powershell
} else {
    $pkgFiles = Get-ChildItem -LiteralPath $pkgDir -Filter 'Pkg*.bas' |
                Sort-Object Name | ForEach-Object { $_.FullName }
}
```

- [ ] **Step 2: Verify the fixture is gone from every surface**

```bash
grep -rn "Fixture" src/ build/ ; rm -f dist/MacrosBundle.fixture.bas
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: `grep` prints **nothing**, and the build reports `Core + 2 package(s)`.

- [ ] **Step 3: Bump `VERSION`**

```bash
echo -n "0.4.0" > VERSION
```

- [ ] **Step 4: Add the `CHANGELOG.md` entry**

Under `## [Unreleased]`, insert:

```markdown
## [0.4.0] - 2026-07-24

### Added

- **`Macros Home`** — a generated index tab (always first): one row per built page
  with its title, its enabled-action count and a link, a link to `Setup Menu`, and
  a legend explaining page codes and Enabled. The name is reserved by the contract.
- **Auto-layout.** With no page codes filled in, buttons distribute automatically —
  one page per display **category**, columns filled to 8 buttons before wrapping,
  a continuation page past 3 columns. Filling in even one code still hands total
  control back to the codes.
- **Category pages.** Pages are titled by category instead of `"Macros - Page 1"`,
  columns are headed by category instead of `"Column A"`, and every button now
  carries its manifest description underneath — grouped pages previously had no
  descriptions at all.
- **Navigation.** A persistent `<< Home` link in a fixed position on every page,
  plus `< Previous` / `Next >`.
- **Per-action enable.** The placement block is now the **Action catalog**
  (`Action | Enabled | Page code`). `Enabled = No` hides a button while leaving the
  action listed on Setup and still runnable inside a chain; Finalize reports which
  actions are off, and which chains use one, without blocking.
- **`usedBy` on a field** — settings belonging only to switched-off actions are
  collected into one collapsed *"Settings for actions you haven't turned on"* band
  at the bottom of Setup. They are relegated, never omitted: Setup re-seeds from
  the finalized `_Config`, so omitting them would cost a second Finalize and
  discard un-finalized typing.
- **`category` on an action** — optional, defaults to the package name, and
  deliberately decoupled from the module so a future package split cannot move a
  button in the user's mental model.
- Contract frozen at **v1.2.0** (additive).
- **Build:** fixture/harness packages are kept out of the default bundle.

### Changed

- Tighter panel spacing (`BTN_H` 30→26, `ROW_GAP` 8→4, `COL_GAP` 24→16,
  `HDR_H` 22→20), with room for the new per-button description.

### Removed

- **The single `Macros` tab default.** With no page codes, v0.3.0 built one tab and
  stacked every button on it; v0.4.0 auto-distributes by category instead, and the
  stale `Macros` tab is deleted on rebuild. Manual page codes are unaffected.

### Fixed

- `RenderPage` no longer calls `Range.Select` on a possibly-inactive sheet — it
  only ever worked because rebuilding deleted and re-added every panel, which
  happens to activate it. Selection now happens once, on `Macros Home`.
- A v0.3.0 `_Config` (whose placement block has no `Enabled` column) is detected
  from its header row, so upgrading reads page codes from the column they are
  actually in instead of silently discarding the user's manual layout.
```

- [ ] **Step 5: Refresh `TODOS.md`**

Replace the `Current:` paragraph (`TODOS.md:8-13`) so it names v0.4.0, contract
v1.2.0, and track 2 (the catalog waves) as next; move the track-1 line from
`## Next — active` into `## Done (recent)`.

- [ ] **Step 6: Regenerate the code map**

```bash
node scripts/code-map.mjs
```

The map is generated output — never hand-edit it. Committing it in the same PR as
the code keeps its freshness stamp honest.

- [ ] **Step 7: Final gate, then commit**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add -A
git commit -m "chore: release v0.4.0 (built-page polish - Macros Home, auto-layout, per-action enable; contract v1.2.0)"
```

- [ ] **Step 8: Review and ship**

Run `/review` over the `src/` diff, then `/ship`. The PR body must carry Task 9's
completed checklist — the build is not a substitute for it.

---

## Self-Review

**Spec coverage — design §2 against the tasks:**

| Design §2 requirement | Task |
|---|---|
| `Macros Home`: title, legend, card per category with name + count + link, `Setup Menu` link | 7 |
| `Macros N` titled by category, not `"Page 3"` | 5 (`LabelManualPages`), 6 (`PageTitleFor`) |
| Columns headed by name, not `"Column A"` | 6 (`ColHdrFor`); manual mode keeps code headers by design |
| Every button followed by its description | 5 (`RenderPage`) |
| Persistent fixed-position `← Home` + Previous/Next | 5 (`RenderNavRow`) — ASCII `<< Home` per the encoding constraint |
| Auto-layout by category, fill-to-max then wrap | 6 |
| Manual codes remain a complete override; `hide` still omits | 5 (`BuildManualLayout`, unchanged semantics) |
| Tighter spacing via the existing constants | 5 Step 1 |
| No search, no Recent Actions, Favorites deferred | not built — correct |
| Placement block → `Action \| Enabled \| Page code` | 3 |
| Enabled defaults yes → v0.4.0 is a no-op for a current user | 3 (`EnabledFor` absence = True) + 9 Step 7 |
| Read by a sibling of `ReadCodeMap` keyed off `RID_COL` | 3 (`ReadEnabledMap`) |
| `GatherButtons` skips disabled actions | 3 |
| Disabled actions stay listed | 3 + 9 Step 3 |
| Fields always render; off ones collapse into a band | 4 |
| `Core_AddField` gains `usedBy`; undeclared = package-wide | 2 |
| Contract additive bump to v1.2.0 | 1 |
| `PkgFixture`, ~60 no-op actions, 9 categories, never ships, deleted after | 8 + 10 |

**Two things the design did not specify, decided here and flagged for the owner:**

1. **The v0.3.0 `_Config` migration.** Inserting `Enabled` moves the page code from
   column D to E, and an existing `_Config` still has it in D. Undetected, the
   first rebuild after upgrading throws away the user's whole manual layout. Task 3
   Step 3 adds a header-row shim and Task 9 Step 7 tests it. §2.4's "no behaviour
   changes" is true of the contract but not of the sheet geometry.
2. **What `Enabled = No` means for a chain.** Chosen: it hides the *button*, not the
   action — so a disabled action is still a legal chain step, and Finalize reports
   the situation without blocking. This makes "reachable only inside a chain" a
   supported configuration and keeps `Core_Run` untouched. Written into the
   contract in Task 1 Step 4; if the owner wants disabling to also break chains,
   that is a different (and blocking) rule.

**Placeholder scan:** none. Every code step carries the actual VBA or PowerShell.
The only intentionally prose-only steps are Task 9's owner checks, which are
manual by nature, and Task 10 Step 5's `TODOS.md` edit, whose content depends on
what the release actually contains.

**Type/name consistency checked:** `layout` is the dictionary name throughout
Tasks 5–7; `enab` is the enabled map throughout Tasks 2–4; button arrays are
5-element everywhere after Task 3 (`e(4)` = category, read only in
`LabelManualPages`, `BuildAutoLayout`); `RenderTier` takes 8 arguments at all four
call sites; `HasAdvanced` takes 2 at its single call site; `FieldInUse` /
`HasAnyTier` / `EnabledFor` / `IsYes` / `HasEnabledColumn` / `CollapseRows` /
`RenderOffBand` / `SetPageTitle` / `SetColHeader` / `PageButtonCount` /
`RenderNavRow` / `NavLink` / `BuildAutoLayout` / `CategoryOrder` / `PageTitleFor` /
`ColHdrFor` / `BuildManualLayout` / `LabelManualPages` / `RenderAllPages` /
`BuildHomePage` / `HomeLegend` / `HomeLine` are each declared exactly once and none
collides with an existing procedure name (`build.ps1` would fail the build if one
did). `IsYes` is deliberately **not** named `IsTrue` — that name exists at
`Core.bas:1857` with the opposite default.
