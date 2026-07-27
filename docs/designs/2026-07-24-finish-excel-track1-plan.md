# Finish Excel — Track 1 (built-page polish) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Turn the generated button pages from "disorganized and cluttered at seven
buttons" into a navigable, auto-laid-out index — `Macros Home` plus capacity-packed
category pages with legible descriptions and navigation — and add per-action
availability so a user only sees the actions and the settings they use.

**Architecture:** Core-only; no package logic changes. Three additive manifest
attributes (`category` on an action, `usedBy` on a field, plus an `Availability`
column on the existing engine-owned placement block) feed a rewritten panel
builder. The placement block becomes the **Action Catalog**
(`Action | Availability | Page code`); `GatherButtons` — the single funnel for both
actions and chains — filters on availability. The Setup renderer relegates settings
belonging only to switched-off actions into one collapsed outline band rather than
omitting them. When no button carries a page code, buttons auto-distribute into
**capacity-packed** pages of up to three category-headed columns, drawn as an
aligned, unmerged **cell-card grid** so descriptions AutoFit instead of clipping.

**Tech Stack:** VBA (Excel/Windows), one required `src/core/Core.bas` + thin
`src/packages/Pkg<Name>.bas`. Build: `pwsh build/build.ps1` → `dist/MacrosBundle.bas`.
Syntax gate: `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`. No headless
VBA runner exists — see **Verification model** below.

---

## Provenance — read this before changing any decision below

The first draft of this plan was argued by a `/deliberate` panel (native /
veteran / aesthete; `gpt-5.6-terra` ×2 and `gpt-5.6-sol`, medium effort, 2 rounds,
88.3k tokens, halted on the first loop). It found six problems. **Four were
unanimous, and all four were real:**

1. **Tasks 5–7 shipped three commits that would not compile.** The build script and
   syntax checker are structural and do not resolve call targets, so an entry point
   calling not-yet-written procedures passed the gate. In a repo whose only compiler
   is a human, that is not acceptable. **Fixed:** the new renderer and auto-layout
   land as dead-but-compilable code, and a small separate cutover task rewires the
   entry point last. Every commit compiles.
2. **The original layout made today's deployment worse.** 8 actions with no codes
   was ONE tab; one-page-per-category would have shipped 2 category tabs + Home —
   more tabs and an extra click, from a release motivated by "cluttered". **Fixed**
   by capacity packing (below): 8 actions now produce one page with two category
   columns.
3. **One page per category is an unbounded tab multiplier** (10 categories → 11
   page tabs, in a tab bar that shows 8–10). Home exists to solve tab sprawl and
   that rule would have created it. **Fixed:** a page is a *capacity* unit.
4. **The migration shim keyed off a user-editable header cell** to decide whether
   to preserve or destroy a user's saved layout. The repo already keeps identity in
   the hidden `RID_COL`, never in display text. **Fixed:** a machine-owned schema
   marker in that hidden column, plus validation, plus fail-closed abort.

**Owner decisions on the three contested points** (these override the panel):

- **Capacity packing, but Home is ALWAYS generated** — even for a single page. One
  page plan, one renderer, no threshold to test. The panel wanted a no-index floor;
  the owner chose the uniform workbook shape.
- **Descriptions stay on the action page**, as the aesthete's unmerged cell-card
  grid, **and** Home carries a full action catalog. The native and veteran wanted
  buttons-only pages; the owner kept point-of-click legibility.
- **Three-state availability now** — `Button` / `Chains only` / `Off`, with an Excel
  data-validation dropdown. The native and veteran both argued for two states on
  the grounds that no deployment has any chains yet; the owner chose to define the
  semantics completely in the release that introduces them.

---

## Global Constraints

Copied from `CLAUDE.md`, `spec/contract.md` and the design's do-not-revert list.
Every task's requirements implicitly include this section.

- **Target version:** repo `VERSION` → `0.4.0`; contract → **v1.2.0** (additive).
- **`src/` is the source of truth.** A `.bas` edit does nothing in Excel until the
  bundle is rebuilt and re-imported. There is no round-trip back to `src/`.
- **EVERY COMMIT MUST COMPILE.** New procedures may be unreachable, but no
  committed call site may name a procedure that does not exist yet. Dead but
  compilable beats a bundle that only compiles two tasks later.
- **Flat namespace, mandatory naming.** Everything `Public` is `Core_*` or
  `Pkg<Name>_*`; everything else `Private`. Shared helpers live ONLY in Core.
  `build.ps1` FAILS on a duplicate procedure name across modules — every new
  private helper name in this plan must be globally unique.
- **`Option Explicit` everywhere.** Max **25 line continuations** per statement.
- **Module-level `Const`/`Dim` stays in the declarations section** at the top of
  `Core.bas` (above `Core_DeclarePackage`). A declaration after the first procedure
  is a VBA compile error and `build.ps1` fails on it.
- **ASCII source bytes only.** `build.ps1` writes UTF-8 **no BOM** on the stated
  assumption that content is ASCII. Use `<< Home`, `< Previous`, `Next >` — never a
  literal `←`/`→`/`▶`. (A runtime `ChrW$(&H2713)` is fine and already used at
  `Core.bas:1786`-ish — the constraint is on source bytes, not on rendered glyphs.)
- **Nothing on the card grid may be merged.** Excel's AutoFit *ignores merged
  cells* — that is why the Setup band rows already have to estimate their height
  from a character count. The whole point of the card grid is that AutoFit works,
  so a merge anywhere on it silently reintroduces clipping.
- **No VBIDE trust dependency, ever.** Packages are discovered by probing
  `Pkg<Name>_Manifest` via `Application.Run` against `CoreCandidates`.
- **Every action wraps in the speed pattern** with guaranteed restore in the error
  handler. Quote every path.
- **Contract is normative.** Where this plan and `spec/contract.md` disagree, the
  contract wins — which is why Task 1 lands the contract first.

### Do NOT revert (design §5 + the compose-your-buttons handoff log)

1. **Setup COLLAPSES settings for unavailable actions; it does NOT omit them.**
   `Core_BuildSetup` deletes and re-seeds Setup from the **finalized** `_Config`, so
   omit-unless-enabled costs toggle → Finalize → rebuild → fill → Finalize *and*
   discards anything typed but not yet finalized. This deliberately overrides all
   three panelists of an earlier `/deliberate` round.
2. **Module-level `Const` stays in the declarations section.**
3. **`_Config` is unhidden before `.Delete`** (`Core_FinalizeSetup`) — a very-hidden
   sheet resists delete, which orphaned `Setup (2)`/`(3)` copies.
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

Both are **structural checks, not a compiler** — they do not resolve call targets,
which is exactly the hole that let the first draft plan three uncompilable commits.
The real gate is **Task 10**: the owner imports the bundle, runs `Debug ▸ Compile
VBAProject`, and clicks every action plus the cancel and error paths. **Never claim
this work is verified before Task 10 passes.** Task 9 builds the `PkgFixture`
harness that makes Task 10 exercise a 75-button layout while only 8 real actions
exist.

---

## File Structure

| File | Responsibility | Change |
|---|---|---|
| `spec/contract.md` | Language-neutral source of truth | Modify — v1.1.0 → v1.2.0 |
| `src/core/Core.bas` | Engine: manifest API, Setup renderer, panel builder | Modify — Tasks 2–8 |
| `src/packages/PkgFixture.bas` | Throwaway 75-action layout harness | Create (Task 9), delete (Task 11) |
| `build/build.ps1` | Bundle + structural gates | Modify — exclude fixtures from the default bundle |
| `docs/code-map.md` | Symbol index | Regenerate (Task 11) |
| `CHANGELOG.md`, `TODOS.md`, `VERSION` | Release docs | Modify (Task 11) |

`Core.bas` is 1927 lines and already sectioned 1–7. Every change below lands inside
an existing section — no restructuring, and no new file. Read it **in slices via
`docs/code-map.md`**, never whole.

**Line anchors drift.** Every `Core.bas:NNN` in this plan is measured against the
file as of commit `514f6f4` (1927 lines), *before* any task has landed. Tasks 2–8
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
- Produces: the normative definitions of `category`, `usedBy`, three-state
  `availability`, capacity-packed auto-layout, the always-generated index, the
  config schema marker, and the reserved names that Tasks 2–8 implement.

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

Then add this bullet after the `default`/`defaults[]` bullet at `:68-69`:

```markdown
- **`usedBy` (v1.2.0, optional).** A comma-separated list of action `id`s this
  field configures. A field that declares nothing is **package-wide** and always
  renders in the working area. A field is relegated only when **every** action it
  names is `off` (§6) — an action that is `chains-only` still runs, so it still
  needs its settings. Relegated means **moved, never omitted**: the engine still
  renders the field, still stores it, and still falls back to its default; it is
  simply placed in one collapsed "settings for actions you haven't turned on" band
  at the end of the Setup surface. An unrecognised id in `usedBy` counts as in-use,
  so a typo cannot make a field unreachable.
  **Portability note:** collapsing is a presentation affordance (reference: Excel
  row outlining). A platform that cannot collapse MUST still render the band,
  visibly and last. Omitting the fields is non-conforming — it would force a
  second activate cycle and discard un-activated input.
```

- [ ] **Step 3: Add `category` and `availability` to the §6 action table**

In the action table at `spec/contract.md:188-193`, add two rows after `help`:

```markdown
| `category` | Optional display category (v1.2.0). Defaults to the package `name`. Groups actions on the generated index and heads the columns of a button page; deliberately decoupled from the package, so two packages may feed one category and one package may span several. |
| `availability` | Deployment config, not manifest data (v1.2.0): one of `button`, `chains-only`, `off`. Defaults to `button`. |
```

- [ ] **Step 4: Add the availability + layout semantics to §6**

Insert after the "Grouping / layout (v1.1.0)" paragraph (`spec/contract.md:224`):

```markdown
**Per-action availability (v1.2.0):** each action carries a deployment-level
`availability`, authored on the same engine-owned surface as the placement code
(§5). It has exactly three values and defaults to `button`, so a deployment that
predates v1.2.0 is unchanged.

| Value | Button drawn | Runnable directly | Runnable as a chain step |
|---|---|---|---|
| `button` (default) | yes | yes | yes |
| `chains-only` | no | **no** | yes |
| `off` | no | **no** | **no** |

- **`off` means off.** The engine MUST refuse to dispatch an `off` action, and
  Finalize MUST **block activation** when any chain references one, naming the
  chain. An implementation where a switched-off action still executes inside a
  chain is non-conforming: a user who disables a destructive action and then
  watches it run has been actively misled.
- **`chains-only`** exists so "reachable only as part of a composed button" is an
  explicit, authored state rather than a side effect of hiding something. A direct
  dispatch of a `chains-only` id MUST be refused.
- An action that is not `button` **stays listed** on the authoring surface, so it
  never becomes undiscoverable. Only `off` relegates that action's fields (§2).
- **An unrecognised value MUST block Finalize** and report the accepted values. It
  MUST NOT fall back to `button` — a user who typed `of` meaning `off` would
  otherwise be handed a live destructive button. A **blank** cell is not
  unrecognised: it means "never set" and reads as `button`.
- **Portability note:** the reference implementation presents the three values as a
  validated dropdown on the authoring grid. A platform without cell validation
  still parses the same three tokens and still blocks on anything else.

**Auto-layout (v1.2.0) — supersedes the "no codes" clause of the v1.1.0 grouping
rule above.** When **no** button carries a placement code, the engine distributes
buttons automatically. A page is a **capacity unit, not a category unit**:

- a page holds up to **three category-headed columns**; a column holds up to
  **eight** buttons, all of one category;
- small categories therefore **share** a page;
- a category needing more than one column takes consecutive columns, and a
  category that does not fit in a page's remaining columns starts a new page
  rather than straddling one — splitting a category across tabs destroys the
  grouping that putting it in one place buys;
- a category larger than a whole page continues onto further pages.

The v1.1.0 behaviour — a single default button page holding everything — is
withdrawn: it does not survive contact with a large catalog, which is what this
version exists to support. **When any code is present the v1.1.0 rule is unchanged**
and codes remain a complete override.

**Generated index (v1.2.0):** the engine **always** generates an index surface,
even for a single page, so the workbook has one shape at every scale. It lists one
entry per built page with its title, its button count and a link; a full **action
catalog** (caption, description, and the page each button is on); a link to the
admin surface; and a legend explaining the placement code and availability. Every
built page carries a persistent link back to the index and links to its
previous/next page. Reference implementation: a `Macros Home` sheet; that name is
**reserved** (§7).

**Description legibility (v1.2.0, normative for grid platforms):** an action's
`help` text is displayed with its button and MUST NOT be clipped. Reference
implementation: an unmerged cell-card grid — button anchored to a content cell,
the description in the cell directly beneath it, that row auto-fitted. Merged
cells are forbidden on this grid precisely because Excel's AutoFit ignores them.
```

- [ ] **Step 5: Add the config schema marker to §5**

Insert after the "Engine-owned config (v1.1.0)" paragraph (`spec/contract.md:180`):

```markdown
**Config schema marker (v1.2.0, normative):** the engine writes a machine-owned
schema-version token into the config store alongside the engine-owned blocks,
outside any user-editable cell (reference: the hidden per-row id column, on the
placement block's key row). Readers branch on that token, **never on a visible
label** — a display header is presentation, is user-editable, and must never
decide whether saved configuration is preserved or discarded.

Migration is **fail-closed**. When the marker is absent the store predates v1.2.0
and is read with the older column layout; every value recovered that way MUST be
validated against its grammar (for placement, a valid code or `hide`), and any
conflict — an unknown marker, or a recovered value that does not validate — MUST
abort the rebuild with a report and leave the existing surface untouched. Silently
rebuilding with blanks is non-conforming.
```

- [ ] **Step 6: Reserve the names in §7**

Add after the `chain.*` reserved bullet (`spec/contract.md:234`):

```markdown
- **Reserved:** the display category `Chains` (engine-generated, holds chain
  buttons) and, in the Excel reference implementation, the sheet name
  `Macros Home`. A platform substitutes its own index surface name.
```

- [ ] **Step 7: Record the change in §10**

Append after the v1.1.0 bullet (`:351-354`):

```markdown
- **v1.2.0 (additive):** adds the optional action `category` attribute, the
  optional field `usedBy` attribute, the three-state deployment-level
  `availability` flag, the capacity-packed auto-layout rule for uncoded buttons,
  the always-generated index surface with its navigation and action catalog, the
  no-clipping rule for descriptions, the config schema marker with fail-closed
  migration, and the reserved `Chains` category + `Macros Home` name. No `key`,
  action `id`, kind, value encoding, or **action** semantics changed, so
  v1.0.0/v1.1.0 configs remain valid and read as all-actions-`button`. The one
  *behaviour* replaced is engine layout for the uncoded case (single page →
  capacity-packed pages), which §10 does not classify as breaking: it renames
  nothing and re-encodes nothing. Owner-approved; see the Provenance section of
  `docs/designs/2026-07-24-finish-excel-track1-plan.md` for the three contested
  calls and who lost them.
```

- [ ] **Step 8: Commit**

```bash
git add spec/contract.md
git commit -m "spec: freeze contract v1.2.0 (category, usedBy, three-state availability, packed auto-layout)"
```

---

### Task 2: Manifest API — `category` on actions, `usedBy` on fields

Pure additive plumbing: four optional arguments and one shared predicate. Nothing
reads them yet, so this task changes no behaviour — which is why it is its own
reviewable step.

**Files:**
- Modify: `src/core/Core.bas:80-93` (`Core_AddSection`, `Core_AddField`),
  `:95-115` (`Core_AddList`, `Core_AddTable`), `:125-134` (`Core_AddAction`),
  `:136-155` (`NewField`)

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `Core_AddField(reg, key, label, kind, dflt, [tier], [help], [usedBy])`
  - `Core_AddList(reg, key, label, colHeader, [tier], [help], [usedBy])`
  - `Core_AddTable(reg, key, label, colHeaders, [tier], [help], [usedBy])`
  - `Core_AddAction(reg, actionId, caption, entryPoint, help, [category])`
  - field dictionaries gain `"usedBy"` (lowercased, trimmed, possibly `""`)
  - action dictionaries gain `"category"` (never empty — defaults to `reg("name")`)
  - `Private Function FieldInUse(f As Object, availMap As Object) As Boolean`

- [ ] **Step 1: Add `usedBy` to `NewField`**

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
' comma-separated list of action ids this field configures; a field whose every
' named action is Off moves into the collapsed band at the bottom of Setup (never
' omitted - see Core_BuildSetup). A "Chains only" action still runs, so it still
' keeps its fields in the working area.
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
' category is the DISPLAY grouping used by the generated index and by the column
' headers on a button page. It defaults to the package's display name, so an
' action that declares nothing keeps today's grouping. It is deliberately
' decoupled from the module: two packages may feed one category, and one package
' may span several, so a future package split never moves a button in the user's
' mental model.
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

`FieldInUse` — which reads these `usedBy` values — deliberately lives in **Task 3**,
not here: it calls `IsRunnable`, and a procedure calling one that does not exist yet
is precisely the uncompilable commit this revision exists to prevent.

- [ ] **Step 4: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: build PASSES (`Core + 2 package(s)`, no duplicate-procedure error, Sub
and Function counts balanced, no module-level-declaration error) and the syntax
check reports no errors.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): add optional action category and field usedBy to the manifest API"
```

---

### Task 3: The Action Catalog block — three-state availability

Turns the placement block into `Action | Availability | Page code`, adds the
readers and the dropdown, filters `GatherButtons`, gates the dispatcher, and makes
an `off` action inside a chain a **blocking** Finalize error.

**Files:**
- Modify: `src/core/Core.bas:36-57` (declarations — `CHAIN_CAT`, `AVAIL_*`)
- Modify: `src/core/Core.bas:722-734` (`Core_Run` — direct-dispatch gate)
- Modify: `src/core/Core.bas:645-700` (`RunAction` — `off` gate)
- Modify: `src/core/Core.bas:1242-1266` (`RenderPlacementBlock`)
- Modify: `src/core/Core.bas:1549-1567` (`GatherButtons`)
- Modify: `src/core/Core.bas:1318-1396` (`ComposeReport`)

**Interfaces:**
- Consumes: `FieldInUse` (Task 2), `a("category")` (Task 2), `FindKeyRow`,
  `NextKeyRow`, `CfgSheet`, `Yellow`, `BlockIntro` (pre-existing).
- Produces:
  - module constants `CHAIN_CAT`, `AVAIL_BUTTON`, `AVAIL_CHAIN`, `AVAIL_OFF`
  - `Private Function ParseAvail(ByVal v As String) As String` — `""` = invalid
  - `Private Function ReadAvailMap(ws As Worksheet) As Object` — id(lower) → token
  - `Private Function AvailFor(availMap As Object, ByVal id As String) As String`
  - `Private Function HasButton(availMap As Object, ByVal id As String) As Boolean`
  - `Private Function IsRunnable(availMap As Object, ByVal id As String) As Boolean`
  - `Private Function AvailLabel(ByVal token As String) As String`
  - `Private Sub AvailDropdown(cell As Range)`
  - `GatherButtons` items become **5**-element arrays:
    `Array(id, caption, onAction, help, category)`

- [ ] **Step 1: Add the constants to the declarations section**

In `src/core/Core.bas`, immediately after the `PACKAGES_SHEET` line (`:40`), add:

```vba
Private Const CHAIN_CAT      As String = "Chains"      ' reserved display category for chain buttons

' ---- per-action availability (contract v1.2.0) ----
' Stored in the action catalog block as the display label ("Button" / "Chains
' only" / "Off") and parsed to these tokens. Blank means never set and reads as
' Button, which is what makes v0.4.0 a no-op for a v0.3.0 deployment.
Private Const AVAIL_BUTTON   As String = "button"
Private Const AVAIL_CHAIN    As String = "chains only"
Private Const AVAIL_OFF      As String = "off"

' Config-format marker, written into the hidden RID column of the action catalog's
' key row - a machine-owned cell, not a visible header. A display label is
' presentation, is user-editable, and must never decide whether saved config is
' preserved or discarded (contract section 5). Task 4 builds the migration on it;
' it is declared and read here because ReadAvailMap needs it.
Private Const SCHEMA_VER     As String = "1.2"
```

These must stay in the declarations section above the first procedure —
`build.ps1` fails on a `Const` after the first `Sub`.

- [ ] **Step 2: Add the schema-marker reader and `FieldInUse`**

Insert immediately before `ReadCodeMap`'s comment (`src/core/Core.bas:1268`):

```vba
' The config-format marker for the action catalog block, or "" when absent
' (meaning the block predates v1.2.0 and its page code is one column to the left).
' Deliberately NOT read from the visible header: RID_COL is hidden, is already how
' this codebase carries identity, and is copied into _Config with everything else.
Private Function BlockSchema(ws As Worksheet, ByVal keyRow As Long) As String
    If ws Is Nothing Or keyRow < 1 Then Exit Function
    BlockSchema = Trim(CStr(ws.Cells(keyRow, RID_COL).Value))
End Function
```

And append at the end of section 1, immediately after `FindField` ends
(`src/core/Core.bas:167`) and **before** the `'====` band that opens section 2:

```vba
' Does this field belong to the user's working area, given the availability map?
' True when the field is package-wide (no usedBy) or when at least ONE action it
' names is not Off - a "Chains only" action still runs, so it still needs its
' settings. False means "relegate to the collapsed band" - never "omit":
' Core_BuildSetup deletes and re-seeds Setup from the FINALIZED _Config, so
' omitting a field would cost toggle -> Finalize -> rebuild -> fill -> Finalize
' and would discard anything typed but not yet finalized (design section 2.3).
' An unknown id counts as in-use, so a typo in usedBy cannot hide a field forever.
Private Function FieldInUse(f As Object, availMap As Object) As Boolean
    Dim ids() As String, i As Long, one As String
    FieldInUse = True
    If f Is Nothing Then Exit Function
    If Not f.Exists("usedBy") Then Exit Function
    If Len(Trim(CStr(f("usedBy")))) = 0 Then Exit Function
    ids = Split(CStr(f("usedBy")), ",")
    For i = LBound(ids) To UBound(ids)
        one = LCase(Trim(ids(i)))
        If Len(one) > 0 Then
            If Not availMap.Exists(one) Then Exit Function     ' unknown id -> in use
            If IsRunnable(availMap, one) Then Exit Function    ' one runnable is enough
        End If
    Next i
    FieldInUse = False
End Function
```

- [ ] **Step 3: Add the availability parser and readers**

Insert immediately before `RenderPlacementBlock`'s comment block
(`src/core/Core.bas:1240`):

```vba
' Parse an Availability cell to one of the three tokens.
'   ""      -> AVAIL_BUTTON  (never set; a pre-v0.4.0 config reads as today)
'   invalid -> ""            (the CALLER must block Finalize)
' It must never fall back to Button on an unrecognised string: a user who typed
' "of" meaning "off" would otherwise be handed a live destructive button.
' Yes/No are accepted as synonyms so a user upgrading from a two-value habit, or
' one who types the obvious thing, is understood rather than blocked.
Private Function ParseAvail(ByVal v As String) As String
    Dim s As String
    s = LCase(Trim(v))
    If Len(s) = 0 Then ParseAvail = AVAIL_BUTTON: Exit Function
    Select Case s
        Case AVAIL_BUTTON, "yes", "y", "on", "true", "1"
            ParseAvail = AVAIL_BUTTON
        Case AVAIL_CHAIN, "chain only", "chains", "chain", "chains-only"
            ParseAvail = AVAIL_CHAIN
        Case AVAIL_OFF, "no", "n", "false", "0"
            ParseAvail = AVAIL_OFF
        Case Else
            ParseAvail = ""                  ' unrecognised - caller must block
    End Select
End Function

' The display label written into the cell for a token.
Private Function AvailLabel(ByVal token As String) As String
    Select Case token
        Case AVAIL_CHAIN: AvailLabel = "Chains only"
        Case AVAIL_OFF:   AvailLabel = "Off"
        Case Else:        AvailLabel = "Button"
    End Select
End Function

' id(lower) -> availability token, from the action catalog block. An id with no
' row is simply absent, and AvailFor treats absence as Button - so a v0.3.0
' _Config (which has no Availability column at all) reads as everything on.
' An unrecognised cell is stored as "" so ComposeReport can block on it.
Private Function ReadAvailMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadAvailMap = map
    If ws Is Nothing Then Exit Function
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr = 0 Then Exit Function
    If BlockSchema(ws, kr) <> SCHEMA_VER Then Exit Function   ' pre-v1.2 block: all Button
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(id) > 0 Then map(LCase(id)) = ParseAvail(CStr(ws.Cells(i, VAL_COL + 1).Value))
    Next i
End Function

' Button unless the map says otherwise. Absence means Button, deliberately: an
' unconfigured or pre-v1.2 deployment must show every button.
Private Function AvailFor(availMap As Object, ByVal id As String) As String
    AvailFor = AVAIL_BUTTON
    If availMap Is Nothing Then Exit Function
    If availMap.Exists(LCase(id)) Then AvailFor = CStr(availMap(LCase(id)))
End Function

' Fail closed on both: an invalid token ("") draws no button and does not run.
Private Function HasButton(availMap As Object, ByVal id As String) As Boolean
    HasButton = (AvailFor(availMap, id) = AVAIL_BUTTON)
End Function

Private Function IsRunnable(availMap As Object, ByVal id As String) As Boolean
    Dim t As String
    t = AvailFor(availMap, id)
    IsRunnable = (t = AVAIL_BUTTON Or t = AVAIL_CHAIN)
End Function

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

- [ ] **Step 4: Rewrite `RenderPlacementBlock` as the Action Catalog**

Replace `src/core/Core.bas:1240-1266` (comment + function) with:

```vba
' One row per registry action: hidden id in RID_COL, "caption [id]" (grey) in C, a
' yellow Availability cell with a validated dropdown in D, a yellow Page-code cell
' in E. Prior answers are preserved from CfgSheet() by id. Unavailable actions STAY
' listed here - that is the owner's hard requirement: turning a button off must
' never make the action undiscoverable.
Private Function RenderPlacementBlock(ws As Worksheet, ByVal r As Long) As Long
    Dim prior As Object, priorAvail As Object, reg As Object, a As Object, id As String
    r = BlockIntro(ws, r, "Action catalog  -  what each action is available as, and where it goes", _
                   "Availability: 'Button' shows a button; 'Chains only' hides the button but a " & _
                   "Chain can still run it; 'Off' hides it and stops it running anywhere. " & _
                   "Page code: leave every code blank for the automatic layout; fill in even one " & _
                   "and codes take over - number = page, letters = column (e.g. 1A). " & _
                   "Type  hide  to leave one button off a manual layout.")
    ws.Cells(r, KEY_COL).Value = LAYOUT_KEY
    ws.Cells(r, VAL_COL).Value = "Action"
    ws.Cells(r, VAL_COL + 1).Value = "Availability"
    ws.Cells(r, VAL_COL + 2).Value = "Page code"
    ws.Range(ws.Cells(r, VAL_COL), ws.Cells(r, VAL_COL + 2)).Font.Bold = True
    ' The machine-owned schema marker, in the hidden id column of the key row -
    ' NOT in the visible header. See Task 4 / contract section 5.
    ws.Cells(r, RID_COL).Value = SCHEMA_VER
    r = r + 1

    Set prior = ReadCodeMap(CfgSheet())            ' preserve prior codes by id
    Set priorAvail = ReadAvailMap(CfgSheet())      ' preserve prior availability by id
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            id = CStr(a("id"))
            ws.Cells(r, VAL_COL).Value = CStr(a("caption")) & " [" & id & "]"
            ws.Cells(r, VAL_COL).Font.Color = RGB(90, 90, 90)
            ws.Cells(r, RID_COL).Value = id
            ws.Cells(r, VAL_COL + 1).Value = AvailLabel(AvailFor(priorAvail, id))
            If prior.Exists(LCase(id)) Then ws.Cells(r, VAL_COL + 2).Value = prior(LCase(id))
            Yellow ws.Cells(r, VAL_COL + 1)
            Yellow ws.Cells(r, VAL_COL + 2)
            AvailDropdown ws.Cells(r, VAL_COL + 1)
            r = r + 1
        Next a
    Next reg
    RenderPlacementBlock = r
End Function
```

Note `AvailFor(priorAvail, id)` returns `AVAIL_BUTTON` for an id that was invalid
in the prior config only because `ReadAvailMap` stored `""` and `AvailFor` returns
the raw stored value — which for `""` is `""`, and `AvailLabel("")` yields
`"Button"`. That is deliberate: the rebuild **repairs** an invalid cell to the safe
default rather than propagating it, and Finalize would have blocked before that
config was ever activated.

- [ ] **Step 5: Filter `GatherButtons` and carry the category**

Replace `src/core/Core.bas:1547-1567` (comment + `GatherButtons`) with:

```vba
' Each item: Array(id, caption, onAction, help, category). Actions from the
' registry, chains from activated config. This is the SINGLE funnel for both, so
' the availability filter lives here and nowhere else. Admin buttons are never in
' this list - they live on the Setup Menu tab.
Private Function GatherButtons() As Collection
    Dim out As New Collection, reg As Object, a As Object, chains As Collection, ch As Object
    Dim onAct As String, availMap As Object
    Set availMap = ReadAvailMap(CfgSheet())
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If HasButton(availMap, CStr(a("id"))) Then
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

A chain is never filtered by availability — it has no such column, and a chain is
an explicit user authoring act.

- [ ] **Step 6: Gate the dispatcher**

Replace `src/core/Core.bas:722-734` (`Core_Run`) with:

```vba
' The single public dispatch entry. Resolves an action id or a reserved chain.* id.
' A DIRECT invocation of anything not available as a button is refused here: with
' the button gone the only ways to arrive are a stale shape or Alt+F8, and neither
' should be able to run something the user switched off. Chain steps do not pass
' through here - Core_RunChain calls RunAction, which independently refuses Off.
Public Sub Core_Run(ByVal actionId As String)
    Dim res As Object, availMap As Object, t As String
    If IsChainId(actionId) Then Core_RunChain actionId: Exit Sub
    Set availMap = ReadAvailMap(CfgSheet())
    t = AvailFor(availMap, actionId)
    If t <> AVAIL_BUTTON Then
        If t = AVAIL_CHAIN Then
            MsgBox "'" & actionId & "' is set to 'Chains only' on the Setup tab, so it " & _
                   "runs only as a step inside a Chain." & vbCrLf & vbCrLf & _
                   "Change its Availability to 'Button' to run it on its own.", _
                   vbInformation, "Not available on its own"
        Else
            MsgBox "'" & actionId & "' is switched Off on the Setup tab." & vbCrLf & vbCrLf & _
                   "Change its Availability to run it.", vbExclamation, "Action is off"
        End If
        Exit Sub
    End If
    Set res = RunAction(actionId, False)
    If Not res("found") And res("errNumber") = 0 Then
        MsgBox "Action not installed: " & actionId, vbExclamation, "Core"
    ElseIf Not res("ok") And Not res("reported") Then
        MsgBox "Action '" & actionId & "' failed:" & vbCrLf & vbCrLf & _
               "#" & res("errNumber") & "  " & res("errText"), vbExclamation, "Macro Error"
    End If
End Sub
```

Then in `RunAction` (`src/core/Core.bas:645`), immediately after the
`res("reported") = False` line and **before** `On Error GoTo StepFail`, insert the
defence-in-depth gate that covers the chain path too:

```vba
    ' Defence in depth: Finalize blocks a chain that references an Off action, so
    ' this should be unreachable through a validated config. It exists because
    ' "switched off but it ran anyway" is the one failure mode of this feature
    ' that a finance user cannot recover from.
    If Not IsRunnable(ReadAvailMap(CfgSheet()), actionId) Then
        res("found") = True
        res("ok") = False
        res("errNumber") = 0
        res("errText") = "Action '" & actionId & "' is switched Off on the Setup tab."
        Set RunAction = res
        Exit Function
    End If
```

- [ ] **Step 7: Make an `off` action in a chain a BLOCKING error**

In `ComposeReport`, add to the `Dim` block (after the
`Dim ch2 As Object, actIds As Object, seenIds As Object, cid As String` line):

```vba
    Dim availMap As Object, offList As String, badAvail As String
    Dim chainOnly As String, nOff As Long, t As String
```

Then insert immediately **before** the `If AnyCode(map) Then` line:

```vba
    ' Availability. An unrecognised cell and an Off action used by a chain are both
    ' STRUCTURAL - they block activation. "Chains only" is reported for confidence,
    ' not as a problem: it is an authored state, not an accident.
    Set availMap = ReadAvailMap(ws)
    For Each regX In BuildRegistry()
        For Each aX In regX("actions")
            t = AvailFor(availMap, CStr(aX("id")))
            If Len(t) = 0 Then
                badAvail = badAvail & "  - " & CStr(aX("id")) & vbCrLf
                structural = structural + 1
            ElseIf t = AVAIL_OFF Then
                nOff = nOff + 1
                offList = offList & "  - " & CStr(aX("caption")) & " [" & CStr(aX("id")) & "]" & vbCrLf
            ElseIf t = AVAIL_CHAIN Then
                chainOnly = chainOnly & "  - " & CStr(aX("caption")) & " [" & CStr(aX("id")) & "]" & vbCrLf
            End If
        Next aX
    Next regX
    If Len(badAvail) > 0 Then
        s = s & vbCrLf & "UNRECOGNISED Availability (use Button, Chains only, or Off):" & _
            vbCrLf & badAvail
    End If
    For Each ch In chains
        For Each stepId In ch("steps")
            If ids.Exists(LCase(CStr(stepId))) Then
                If AvailFor(availMap, CStr(stepId)) = AVAIL_OFF Then
                    s = s & "OFF ACTION IN CHAIN: '" & CStr(ch("label")) & "' uses " & _
                        CStr(stepId) & ", which is switched Off." & vbCrLf
                    structural = structural + 1
                End If
            End If
        Next stepId
    Next ch
    If nOff > 0 Then
        s = s & vbCrLf & "OFF (" & nOff & " action(s); no button, cannot run, and their " & _
            "settings collapse to the bottom of Setup):" & vbCrLf & offList
    End If
    If Len(chainOnly) > 0 Then
        s = s & vbCrLf & "CHAINS ONLY (no button; runs only inside a Chain):" & vbCrLf & chainOnly
    End If
```

`regX`, `aX`, `ch`, `stepId` and `ids` are already declared and already used
earlier in this function; reusing them after their earlier loops is fine.

- [ ] **Step 8: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

- [ ] **Step 9: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): three-state per-action availability (Button / Chains only / Off)"
```

---

### Task 4: Fail-closed config migration

The data-safety task. Inserting `Availability` moved the page code from column D to
column E. A v0.3.0 `_Config` still has it in D, and getting this wrong silently
destroys a user's hand-built layout.

> **Tasks 3 and 4 are one functionally verifiable unit — do NOT import to Excel
> between them.** Task 3 moves the page code to column E but leaves `ReadCodeMap`
> reading column D, so on a Task-3-only build `ComposeReport` reads the literal
> `"Button"` as every action's page code, `ParseCode` rejects it, and **Finalize
> refuses to activate with one `BAD CODE` per action and no way for the user to fix
> it.** Both commits compile — the plan's compile-every-commit rule holds — but
> Task 3 alone is not *runnable*, and an earlier draft of this plan wrongly claimed
> a reviewer could reject Task 4 while approving Task 3. Step 1 below is what makes
> Task 3 work at all. (Found by the Task 3 code review, not by the build: no
> structural gate can see a column-index mismatch.)

**Files:**
- Modify: `src/core/Core.bas:36-57` (declarations — `SCHEMA_VER`)
- Modify: `src/core/Core.bas:1269-1289` (`ReadCodeMap`)
- Modify: `src/core/Core.bas:820-916` (`Core_BuildSetup` — the abort path)

**Interfaces:**
- Consumes: `FindKeyRow`, `NextKeyRow`, `CfgSheet`, `ParseCode`, `RID_COL`.
- Consumes additionally: `SCHEMA_VER` and `BlockSchema` (both from Task 3).
- Produces:
  - `Private Function ConfigSchemaStatus(ByRef detail As String) As String` —
    `"none"` | `"ok"` | `"legacy"` | `"unknown"` | `"invalid"`

`SCHEMA_VER` and `BlockSchema` already landed in Task 3 (`ReadAvailMap` needs
them). This task builds the migration that uses them.

- [ ] **Step 1: Branch `ReadCodeMap` on the marker**

Replace `src/core/Core.bas:1269-1289` (`ReadCodeMap`) with:

```vba
' id(lower) -> code, merging the action catalog block and the chains block.
' v0.4.0 inserted Availability between Action and Page code, so a pre-v1.2 block
' keeps its code one column to the left. Callers that could DESTROY the recovered
' values (Core_BuildSetup) must consult ConfigSchemaStatus first; this reader is
' permissive on purpose so the report can describe what it found.
Private Function ReadCodeMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String, codeCol As Long
    Dim chains As Collection, ch As Object
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadCodeMap = map
    If ws Is Nothing Then Exit Function
    ' actions
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr > 0 Then
        If BlockSchema(ws, kr) = SCHEMA_VER Then
            codeCol = VAL_COL + 2
        Else
            codeCol = VAL_COL + 1                ' legacy (v0.3.0) layout
        End If
        nk = NextKeyRow(ws, kr)
        For i = kr + 1 To nk - 1
            id = Trim(CStr(ws.Cells(i, RID_COL).Value))
            If Len(id) > 0 And LCase(id) <> LCase(SCHEMA_VER) Then _
                map(LCase(id)) = Trim(CStr(ws.Cells(i, codeCol).Value))
        Next i
    End If
    ' chains
    Set chains = ReadChains(ws)
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 Then map(LCase(CStr(ch("id")))) = CStr(ch("code"))
    Next ch
End Function
```

The chains block is untouched: its page code was already at `VAL_COL + 2`, which is
now the same column as the catalog's.

- [ ] **Step 2: Add the fail-closed status check**

Insert immediately after `BlockSchema`:

```vba
' Classify the ACTIVATED config before anything destructive touches the Setup
' sheet. Returns one of:
'   "none"    - no config yet (a fresh workbook); nothing to lose
'   "ok"      - written by this version
'   "legacy"  - pre-v1.2, every recovered page code validates; safe to migrate
'   "invalid" - pre-v1.2, but a recovered code does not parse; DO NOT rebuild
'   "unknown" - a marker this build does not recognise (a NEWER version wrote it)
' detail is filled with a human-readable explanation for the last two.
Private Function ConfigSchemaStatus(ByRef detail As String) As String
    Dim ws As Worksheet, kr As Long, ver As String, nk As Long, i As Long
    Dim id As String, code As String, tn As Long, cx As Long, cl As String
    detail = ""
    Set ws = CfgSheet()
    If ws Is Nothing Then ConfigSchemaStatus = "none": Exit Function
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr = 0 Then ConfigSchemaStatus = "none": Exit Function
    ver = BlockSchema(ws, kr)
    If ver = SCHEMA_VER Then ConfigSchemaStatus = "ok": Exit Function
    If Len(ver) > 0 Then
        detail = "The saved configuration was written by a newer version of this " & _
                 "toolkit (format '" & ver & "'; this build understands '" & SCHEMA_VER & "')."
        ConfigSchemaStatus = "unknown"
        Exit Function
    End If
    ' Legacy: validate every page code recovered from the OLD column before we let
    ' anything overwrite the sheet holding them.
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(id) > 0 Then
            code = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
            If Len(code) > 0 And LCase(code) <> "hide" Then
                If Not ParseCode(code, tn, cx, cl) Then
                    detail = detail & "  - " & id & ": '" & code & "'" & vbCrLf
                End If
            End If
        End If
    Next i
    If Len(detail) > 0 Then
        detail = "These saved page codes could not be read from the older " & _
                 "configuration format:" & vbCrLf & detail
        ConfigSchemaStatus = "invalid"
    Else
        ConfigSchemaStatus = "legacy"
    End If
End Function
```

- [ ] **Step 3: Abort `Core_BuildSetup` before it destroys anything**

In `Core_BuildSetup`, insert at the very top — **before** the
`If Core_SheetExists(SETUP_SHEET) Then` block that deletes the sheet:

```vba
    Dim schemaState As String, schemaDetail As String
    schemaState = ConfigSchemaStatus(schemaDetail)
    If schemaState = "unknown" Or schemaState = "invalid" Then
        MsgBox schemaDetail & vbCrLf & vbCrLf & _
               "Nothing has been changed. Rebuilding the Setup tab now would " & _
               "discard those saved answers, so it has been stopped." & vbCrLf & vbCrLf & _
               "Note them down (or restore a backup of this workbook) before " & _
               "rebuilding.", vbExclamation, "Create Setup Tab"
        Exit Sub
    End If
```

This is the fail-closed requirement: on any ambiguity the existing surface is left
exactly as it was, and the user is told what could not be read.

- [ ] **Step 4: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): machine-owned config schema marker with fail-closed v0.3.0 migration"
```

---

### Task 5: Collapse settings for switched-off actions

The design's most-likely-to-be-wrongly-"fixed" decision. Every field still renders
and still saves; fields belonging only to `Off` actions move into one outline band.

**Files:**
- Modify: `src/core/Core.bas:820-916` (`Core_BuildSetup`)
- Modify: `src/core/Core.bas:923-944` (`RenderTier`)
- Modify: `src/core/Core.bas:957-967` (`HasAdvanced`)

**Interfaces:**
- Consumes: `FieldInUse` (Task 2), `ReadAvailMap` (Task 3), `SectionHeader`,
  `TintRow`, `RenderField`.
- Produces:
  - `RenderTier(ws, r, reg, tier, showSections, bannerText, availMap, wantInUse)`
  - `HasAdvanced(reg, availMap)` — now "has an **in-use** advanced field"
  - `Private Function HasAnyTier(reg, availMap, ByVal wantInUse As Boolean) As Boolean`
  - `Private Function RenderOffBand(ws, r, availMap, ByRef bandFirst, ByRef bandLast) As Long`
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
                            ByVal bannerText As String, availMap As Object, _
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
            If FieldInUse(f, availMap) = wantInUse Then
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
' whose actions are all Off goes to the off band instead, so drawing the ADVANCED
' header for it would leave an empty band.
Private Function HasAdvanced(reg As Object, availMap As Object) As Boolean
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If CStr(f("tier")) = "advanced" Then
                If FieldInUse(f, availMap) Then
                    HasAdvanced = True
                    Exit Function
                End If
            End If
        End If
    Next f
End Function

' Does this package have ANY non-section field on the requested side of the
' in-use split? Used to skip both an empty ADVANCED band and an empty off band.
Private Function HasAnyTier(reg As Object, availMap As Object, ByVal wantInUse As Boolean) As Boolean
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If FieldInUse(f, availMap) = wantInUse Then
                HasAnyTier = True
                Exit Function
            End If
        End If
    Next f
End Function
```

- [ ] **Step 3: Add `RenderOffBand` and `CollapseRows`**

Insert immediately after `HasAnyTier`, before `RenderScalar`:

```vba
' Every field whose usedBy names only Off actions, from every package, collected
' under one band at the very bottom of the field area and collapsed with Excel's
' row outline.
'
' They are RENDERED, never omitted. DO NOT "improve" this to skip them:
' Core_BuildSetup deletes and re-seeds Setup from the FINALIZED _Config, so
' omit-unless-enabled costs toggle -> Finalize -> rebuild -> fill -> Finalize and
' discards anything typed but not yet finalized (design section 2.3, which
' deliberately overrides all three panelists of an earlier review round).
'
' Returns the next free row; reports the collapsible span through bandFirst /
' bandLast so the caller can group AFTER its AutoFit pass. bandFirst = 0 means
' nothing was relegated and no band was drawn.
Private Function RenderOffBand(ws As Worksheet, ByVal r As Long, availMap As Object, _
                               ByRef bandFirst As Long, ByRef bandLast As Long) As Long
    Dim reg As Object, any As Boolean
    bandFirst = 0
    bandLast = 0
    RenderOffBand = r
    For Each reg In BuildRegistry()
        If HasAnyTier(reg, availMap, False) Then any = True
    Next reg
    If Not any Then Exit Function              ' nothing off: draw no band at all

    r = SectionHeader(ws, r, "SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON")
    r = TintRow(ws, r, "These belong to actions set to 'Off' in the action catalog above. " & _
                "They are kept and still saved, so nothing you typed is lost - set the " & _
                "action back to 'Button' and they move up into the main list. Click the + " & _
                "in the left margin to open this section.", _
                9, False, True, RGB(90, 70, 120))
    bandFirst = r
    For Each reg In BuildRegistry()
        If HasAnyTier(reg, availMap, False) Then
            r = TintRow(ws, r, CStr(reg("name")), 10, True, False, RGB(90, 70, 120))
            r = RenderTier(ws, r, reg, "simple", False, UCase(CStr(reg("name"))), availMap, False)
            r = RenderTier(ws, r, reg, "advanced", False, UCase(CStr(reg("name"))), availMap, False)
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

Add to `Core_BuildSetup`'s declarations:

```vba
    Dim ws As Worksheet, reg As Object, f As Object, r As Long, nPkgs As Long
    Dim firstFieldRow As Long, availMap As Object, bandFirst As Long, bandLast As Long
```

Replace the package loop and what follows (`src/core/Core.bas:868-893`) with:

```vba
    ' The availability map comes from the ACTIVATED config, so a rebuild reflects
    ' the last Finalize. A fresh workbook has no _Config, so the map is empty and
    ' AvailFor reports Button for everything - the v0.4.0 default.
    Set availMap = ReadAvailMap(CfgSheet())

    firstFieldRow = r
    Dim banner As String
    For Each reg In BuildRegistry()
        nPkgs = nPkgs + 1
        banner = UCase(CStr(reg("name")))
        r = SectionHeader(ws, r, banner)
        r = RenderTier(ws, r, reg, "simple", True, banner, availMap, True)
        If HasAdvanced(reg, availMap) Then
            r = SectionHeader(ws, r, "ADVANCED  -  most users can leave these as-is")
            r = RenderTier(ws, r, reg, "advanced", False, banner, availMap, True)
        End If
        r = r + 1
    Next reg

    r = RenderComposeSection(ws, r)

    ' The off band goes BELOW Compose on purpose: the action catalog is where the
    ' user turns things on and off, so it must never sit under a collapsed section.
    r = r + 1
    r = RenderOffBand(ws, r, availMap, bandFirst, bandLast)

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

Everything from `ws.Columns("A").Font.Color = ...` onward is unchanged.

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Confirm the two `RenderTier` call sites in the package loop and the two in
`RenderOffBand` all pass **8** arguments, and the only `HasAdvanced` call passes
**2** — a missing argument here is a compile error Excel catches but neither of
these tools does.

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): collapse settings belonging only to Off actions"
```

---

### Task 6: The card-grid page renderer — dead but compilable

Everything needed to draw a page, added **without touching the entry point**. The
existing `BuildSinglePanel` / `BuildGroupedPanels` path still runs after this task,
so the commit compiles and the workbook behaves exactly as v0.3.0 did. Task 8 does
the cutover.

**Files:**
- Modify: `src/core/Core.bas:42-48` (geometry constants + `HOME_SHEET`)
- Modify: `src/core/Core.bas:1723-1730` (`AddToPage`)
- Add: after `AddToPage` — the layout setters, the card grid, `RenderPage`,
  `RenderNavRow`, `NavLink`, `RenderAllPages`, `BuildManualLayout`,
  `LabelManualPages`

**Interfaces:**
- Consumes: 5-element button arrays (Task 3), `SortedKeys`, `EnsurePanelSheet`,
  `AddBtn`, `ParseCode`, `IndexToCol`.
- Produces the `layout` structure everything downstream uses:
  `layout(tabNum)` → dictionary with `"title"` (String),
  `"cols"` (colIx → Collection of button arrays), `"hdrs"` (colIx → String).
  Plus: `AddToPage`, `SetPageTitle`, `SetColHeader`, `PageButtonCount`,
  `SetupCardGrid`, `CardCol`, `AddBtnInCell`, `FitDescRows`, `RenderPage`,
  `RenderNavRow`, `NavLink`, `RenderAllPages`, `BuildManualLayout`,
  `LabelManualPages`.

- [ ] **Step 1: Replace the geometry constants**

Replace `src/core/Core.bas:42-48` with:

```vba
' ---- button page geometry (v0.4.0: an unmerged CELL-CARD grid) ----
' Cards live in worksheet columns B, D and F, with C and E as narrow gutters.
' NOTHING on this grid may be merged: Excel's AutoFit ignores merged cells (which
' is why the Setup band rows have to estimate their height instead), and AutoFit
' on the description rows is the entire reason a description cannot clip.
Private Const PANEL_BASE As String = "Macros "   ' grouped output tabs: "Macros 1", "Macros 2"...
Private Const HOME_SHEET As String = "Macros Home"  ' generated index; RESERVED name (contract 7)
Private Const CARD_COLS   As Long = 3            ' card columns per page -> B, D, F
Private Const CARD_W      As Double = 30         ' content column width, in characters
Private Const GUTTER_W    As Double = 3          ' spacer column width, in characters
Private Const CARD_HDR_ROW As Long = 5           ' category header row; cards start at 6
Private Const BTN_H      As Double = 26
Private Const BTN_ROW_H  As Double = 30          ' button row height (BTN_H + padding)
Private Const DESC_MIN_H As Double = 24          ' floor after AutoFit
Private Const DESC_MAX_H As Double = 60          ' ceiling: ~5 lines at 9pt in a 30-char column,
                                                 ' which clears the longest shipped help string
Private Const AUTO_COL_MAX As Long = 8           ' buttons per card column before a new column
```

`BTN_W`, `ROW_GAP`, `COL_GAP` and `HDR_H` are no longer used by the card grid. Leave
them in place for now — `BuildGroupedPanels` and the old `RenderPage` still
reference them until Task 8 deletes those, and removing a constant they use would
break this commit's compile.

- [ ] **Step 2: Replace `AddToPage` and add the layout helpers**

Replace `src/core/Core.bas:1722-1730` (comment + `AddToPage`) with:

```vba
' The layout structure every panel builder fills and RenderPage consumes:
'   layout(tabNum)("title")        page title             (String)
'   layout(tabNum)("cols")(colIx)  buttons in that column (Collection of Arrays)
'   layout(tabNum)("hdrs")(colIx)  that column's header   (String)
' A page is created lazily by the first button filed onto it, so callers must file
' a button BEFORE setting its title or column header.
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

' B/D/F carry content, C/E are gutters. Column A stays narrow as a margin.
Private Sub SetupCardGrid(ws As Worksheet)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = CARD_W
    ws.Columns("C").ColumnWidth = GUTTER_W
    ws.Columns("D").ColumnWidth = CARD_W
    ws.Columns("E").ColumnWidth = GUTTER_W
    ws.Columns("F").ColumnWidth = CARD_W
End Sub

' Card content columns are B, D, F -> 2, 4, 6. slot is 1-based; 0 means off-grid.
Private Function CardCol(ByVal slot As Long) As Long
    If slot < 1 Or slot > CARD_COLS Then Exit Function
    CardCol = slot * 2
End Function

' Anchor a button to a cell so it lines up with the description directly beneath.
Private Sub AddBtnInCell(ws As Worksheet, ByVal nm As String, ByVal cap As String, _
                         ByVal action As String, anchor As Range)
    AddBtn ws, nm, cap, action, anchor.Left, anchor.Top + 1, anchor.Width, BTN_H
End Sub

' AutoFit each description row, then clamp. AutoFit works here ONLY because
' nothing on the card grid is merged. The clamp is two-sided: the floor keeps
' one-line cards from looking broken next to three-line neighbours, and the
' ceiling stops one runaway string from pushing the page off screen.
Private Sub FitDescRows(ws As Worksheet, ByVal firstDescRow As Long, ByVal lastRow As Long)
    Dim r As Long
    For r = firstDescRow To lastRow Step 2
        ws.Rows(r).AutoFit
        If ws.Rows(r).RowHeight < DESC_MIN_H Then ws.Rows(r).RowHeight = DESC_MIN_H
        If ws.Rows(r).RowHeight > DESC_MAX_H Then ws.Rows(r).RowHeight = DESC_MAX_H
    Next r
End Sub
```

- [ ] **Step 3: Add the card-grid `RenderPage` and the nav row**

`RenderPage` already exists (the point-positioned version). **Rename the existing
one to `RenderPageLegacy`** so this commit still compiles with the old path
intact, and update its single call site inside `BuildGroupedPanels`. Then add:

```vba
' Draw one "Macros N" page as an unmerged card grid: nav row, page title, one
' category-headed card column per layout column, each button followed by its
' description in the cell below it.
Private Function RenderPage(layout As Object, ByVal tabNum As Long, _
                            ByVal prevSheet As String, ByVal nextSheet As String) As Worksheet
    Dim panel As Worksheet, page As Object, cols As Object, hdrs As Object
    Dim colKeys As Variant, si As Long, slot As Long, colIx As Long
    Dim items As Collection, ii As Long, it As Variant
    Dim cc As Long, bRow As Long, dRow As Long, lastRow As Long, hdr As String
    Set page = layout(tabNum)
    Set cols = page("cols")
    Set hdrs = page("hdrs")
    Set panel = EnsurePanelSheet(PANEL_BASE & tabNum, False)
    SetupCardGrid panel
    RenderNavRow panel, prevSheet, nextSheet
    With panel.Cells(3, 2)
        .Value = CStr(page("title"))
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With

    colKeys = SortedKeys(cols)
    lastRow = CARD_HDR_ROW + 1
    slot = 0
    For si = LBound(colKeys) To UBound(colKeys)
        colIx = CLng(colKeys(si))
        slot = slot + 1
        cc = CardCol(slot)
        If cc = 0 Then GoTo NextSlot        ' more columns than card slots: never draw off-grid
        Set items = cols(colIx)
        hdr = ""
        If hdrs.Exists(colIx) Then hdr = CStr(hdrs(colIx))
        If Len(hdr) > 0 Then
            With panel.Cells(CARD_HDR_ROW, cc)
                .Value = hdr
                .Font.Bold = True
                .Font.Size = 11
                .Font.Color = RGB(54, 32, 100)
                .Interior.Color = RGB(221, 217, 240)
            End With
        End If
        For ii = 1 To items.Count
            it = items(ii)
            bRow = CARD_HDR_ROW + 1 + (ii - 1) * 2
            dRow = bRow + 1
            panel.Rows(bRow).RowHeight = BTN_ROW_H
            AddBtnInCell panel, "b" & tabNum & "_" & colIx & "_" & ii, _
                         CStr(it(1)), CStr(it(2)), panel.Cells(bRow, cc)
            With panel.Cells(dRow, cc)
                .Value = CStr(it(3))
                .Font.Size = 9
                .Font.Color = RGB(110, 110, 110)
                .WrapText = True
                .VerticalAlignment = xlTop
            End With
            If dRow > lastRow Then lastRow = dRow
        Next ii
NextSlot:
    Next si
    FitDescRows panel, CARD_HDR_ROW + 2, lastRow
    Set RenderPage = panel
End Function

' Row 1: a persistent Home link in a fixed position on every page, plus Previous /
' Next. Cells rather than shapes, because a cell hyperlink needs no shape
' bookkeeping and EnsurePanelSheet's Cells.Clear removes it cleanly on rebuild.
' ASCII arrows only - build.ps1 writes UTF-8 with no BOM on the assumption that
' the SOURCE is ASCII, so a literal arrow glyph in the .bas is an import hazard.
Private Sub RenderNavRow(ws As Worksheet, ByVal prevSheet As String, ByVal nextSheet As String)
    NavLink ws.Cells(1, 2), "<< Home", HOME_SHEET
    If Len(prevSheet) > 0 Then NavLink ws.Cells(1, 4), "< Previous", prevSheet
    If Len(nextSheet) > 0 Then NavLink ws.Cells(1, 6), "Next >", nextSheet
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

Nav links go in B, D and F — the card content columns — so they sit under the
grid's own alignment instead of needing their own widths.

- [ ] **Step 4: Add the manual-layout filer and labeller**

Add after `RenderAllPages`. These do **not** replace `BuildGroupedPanels` yet;
they are the new-model equivalents that Task 8 switches to.

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
' user chose. A page is titled by category only when EVERY button on it agrees - a
' hand-built page may deliberately mix categories, and guessing a title for a
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
```

- [ ] **Step 5: Build and syntax-check**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: PASS, and **the bundle still behaves exactly as v0.3.0** — the entry
point is untouched and all of the above is unreachable. Confirm the old
`RenderPage` → `RenderPageLegacy` rename left exactly one call site:

```bash
grep -n "RenderPageLegacy\|RenderPage " src/core/Core.bas
```

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): add the unmerged cell-card page renderer and nav (not yet wired)"
```

---

### Task 7: Capacity-packed auto-layout and `Macros Home` — still not wired

**Files:**
- Add: `src/core/Core.bas` — after `LabelManualPages`

**Interfaces:**
- Consumes: `AddToPage` / `SetPageTitle` / `SetColHeader` / `PageButtonCount` /
  `NavLink` (Task 6), `CHAIN_CAT`, `CARD_COLS`, `AUTO_COL_MAX`, `ADMIN_SHEET`.
- Produces: `BuildAutoLayout`, `CategoryOrder`, `LabelAutoPages`, `BuildHomePage`,
  `HomeCatalog`, `HomeLegend`, `HomeLine`.

- [ ] **Step 1: Write the capacity-packed auto-layout**

```vba
' Auto-layout: no button carries a page code, so distribute by display category.
' A page is a CAPACITY unit, not a category unit: up to CARD_COLS category-headed
' columns per page, up to AUTO_COL_MAX buttons per column. Small categories share
' a page. A category that does not fit in the remaining columns starts a new page
' rather than straddling one - splitting a category across tabs destroys the
' grouping that putting it in one place buys. A category larger than a whole page
' continues onto further pages. Replaces the pre-v0.4.0 single-"Macros"-tab
' default (contract v1.2.0), which did not survive contact with a large catalog.
Private Sub BuildAutoLayout(buttons As Collection, layout As Object)
    Dim cats As Object, seen As New Collection, order As Collection
    Dim e As Variant, cat As String, ci As Long, items As Collection
    Dim needCols As Long, tabNum As Long, slotsUsed As Long
    Dim taken As Long, k As Long, n As Long, firstColOfCat As Boolean

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
    tabNum = 1
    slotsUsed = 0
    For ci = 1 To order.Count
        cat = CStr(order(ci))
        Set items = cats(cat)
        needCols = -Int(-items.Count / AUTO_COL_MAX)        ' ceiling division
        ' Start a fresh page rather than straddle, unless this category is bigger
        ' than a whole page (in which case straddling is unavoidable).
        If needCols <= CARD_COLS And slotsUsed > 0 And slotsUsed + needCols > CARD_COLS Then
            tabNum = tabNum + 1
            slotsUsed = 0
        End If
        taken = 0
        firstColOfCat = True
        For k = 1 To needCols
            If slotsUsed >= CARD_COLS Then
                tabNum = tabNum + 1
                slotsUsed = 0
            End If
            slotsUsed = slotsUsed + 1
            For n = 1 To AUTO_COL_MAX
                taken = taken + 1
                If taken > items.Count Then Exit For
                AddToPage layout, tabNum, slotsUsed, items(taken)
            Next n
            If firstColOfCat Then
                SetColHeader layout, tabNum, slotsUsed, cat
                firstColOfCat = False
            Else
                SetColHeader layout, tabNum, slotsUsed, cat & " (cont.)"   ' SUPERSEDED: Task 7 numbers these — see the note below
            End If
        Next k
    Next ci
    LabelAutoPages layout
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

' A packed page may hold up to CARD_COLS categories, so title it by what is
' actually on it: one category -> its name; several -> the distinct names joined.
Private Sub LabelAutoPages(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim hdrs As Object, hk As Variant, seen As Object, title As String, one As String
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set hdrs = page("hdrs")
        Set seen = CreateObject("Scripting.Dictionary")
        title = ""
        For Each hk In hdrs.Keys
            one = Replace(CStr(hdrs(hk)), " (cont.)", "")   ' SUPERSEDED: matches on " (cont." now — numbered suffixes
            If Not seen.Exists(LCase(one)) Then
                seen(LCase(one)) = True
                title = title & IIf(Len(title) > 0, "   |   ", "") & one
            End If
        Next hk
        If Len(title) > 70 Then title = Left$(title, 67) & "..."
        If Len(title) > 0 Then SetPageTitle layout, tabNum, title
    Next ti
End Sub
```

- [ ] **Step 2: Write `Macros Home`**

```vba
' The generated index - ALWAYS built, even for a single page, so the workbook has
' one shape at every scale (owner's call). Three parts: a page summary, the full
' action catalog, and the legend. Built last and activated, so any rebuild lands
' the user here.
Private Sub BuildHomePage(layout As Object, ByVal totalButtons As Long)
    Dim ws As Worksheet, r As Long, tabKeys As Variant, ti As Long, tabNum As Long
    Dim page As Object
    Set ws = EnsurePanelSheet(HOME_SHEET, True)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 62
    ws.Columns("E").ColumnWidth = 16

    With ws.Range("B2")
        .Value = "Macro Toolkit"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    With ws.Range("B3")
        .Value = "Pick a page below, or find an action in the catalog. Every page has a  << Home  link back here."
        .Font.Italic = True
    End With

    r = 5
    ws.Cells(r, 2).Value = "PAGE"
    ws.Cells(r, 3).Value = "BUTTONS"
    ws.Cells(r, 4).Value = "CATEGORIES ON IT"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Font.Bold = True
        .Interior.Color = RGB(221, 217, 240)
    End With
    r = r + 1

    tabKeys = SortedKeys(layout)
    If UBound(tabKeys) < LBound(tabKeys) Then
        ws.Cells(r, 2).Value = "No actions are available as buttons yet."
        ws.Cells(r, 2).Font.Italic = True
        r = r + 1
    Else
        For ti = LBound(tabKeys) To UBound(tabKeys)
            tabNum = CLng(tabKeys(ti))
            Set page = layout(tabNum)
            NavLink ws.Cells(r, 2), PANEL_BASE & tabNum, PANEL_BASE & tabNum
            ws.Cells(r, 3).Value = PageButtonCount(page)
            ws.Cells(r, 4).Value = CStr(page("title"))
            ws.Cells(r, 4).Font.Color = RGB(90, 90, 90)
            r = r + 1
        Next ti
    End If

    r = r + 1
    NavLink ws.Cells(r, 2), "Setup Menu  -  first-time setup / re-configure", ADMIN_SHEET
    r = r + 2
    r = HomeCatalog(ws, r, layout)
    r = r + 1
    r = HomeLegend(ws, r, totalButtons)

    ws.Activate                     ' Range.Select needs its sheet active
    ws.Range("A1").Select
End Sub

' Every button, with its full description and the page it is on. This is the one
' screen that answers "what can this workbook do?" - and unlike a card, a cell
' here is as wide as it needs to be, so nothing is abbreviated.
Private Function HomeCatalog(ws As Worksheet, ByVal r As Long, layout As Object) As Long
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim cols As Object, colKeys As Variant, ci As Long, colIx As Long
    Dim items As Collection, ii As Long, it As Variant, firstRow As Long
    ws.Cells(r, 2).Value = "ALL ACTIONS"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    ws.Cells(r, 2).Value = "ACTION"
    ws.Cells(r, 3).Value = "PAGE"
    ws.Cells(r, 4).Value = "WHAT IT DOES"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = RGB(240, 240, 245)
    End With
    r = r + 1
    firstRow = r
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        colKeys = SortedKeys(cols)
        For ci = LBound(colKeys) To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            Set items = cols(colIx)
            For ii = 1 To items.Count
                it = items(ii)
                ws.Cells(r, 2).Value = CStr(it(1))
                NavLink ws.Cells(r, 3), PANEL_BASE & tabNum, PANEL_BASE & tabNum
                With ws.Cells(r, 4)
                    .Value = CStr(it(3))
                    .Font.Size = 10
                    .Font.Color = RGB(90, 90, 90)
                    .WrapText = True
                    .VerticalAlignment = xlTop
                End With
                r = r + 1
            Next ii
        Next ci
    Next ti
    ' Nothing here is merged, so AutoFit gives every description its full height.
    If r > firstRow Then ws.Rows(firstRow & ":" & (r - 1)).AutoFit
    HomeCatalog = r
End Function

' The legend. This is the only place the page-code convention and Availability are
' explained where a daily user will actually see them.
Private Function HomeLegend(ws As Worksheet, ByVal r As Long, ByVal totalButtons As Long) As Long
    ws.Cells(r, 2).Value = "HOW THESE PAGES ARE BUILT"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    r = HomeLine(ws, r, totalButtons & " button(s), laid out automatically - categories packed onto pages, up to 3 per page.")
    r = HomeLine(ws, r, "To take charge of the layout, fill in 'Page code' in the action catalog on the Setup tab:")
    r = HomeLine(ws, r, "    1A = page 1, column A.    12AB = page 12, column AB.    hide = leave that button off.")
    r = HomeLine(ws, r, "As soon as ONE code is filled in, the codes control everything and uncoded buttons land on page 1.")
    r = HomeLine(ws, r, "Availability on the Setup tab: 'Button' = shown here. 'Chains only' = no button, but a Chain can run it.")
    r = HomeLine(ws, r, "'Off' = no button and it will not run at all; Finalize refuses to activate a Chain that uses one.")
    HomeLegend = r
End Function

Private Function HomeLine(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Size = 10
    ws.Cells(r, 2).Font.Color = RGB(90, 90, 90)
    HomeLine = r + 1
End Function
```

- [ ] **Step 3: Build, syntax-check, commit**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add src/core/Core.bas
git commit -m "feat(core): capacity-packed auto-layout and the Macros Home index (not yet wired)"
```

Still no behaviour change: nothing above is called yet.

---

### Task 8: The cutover

Small, and the only task that changes what a user sees. Everything it calls
already exists and already compiles.

**Files:**
- Modify: `src/core/Core.bas:1517-1536` (`Core_BuildControlPanel`)
- Delete: `src/core/Core.bas:1577-1630` (`BuildSinglePanel`)
- Delete: `BuildGroupedPanels`, `RenderPageLegacy`, and the now-unused `BTN_W` /
  `ROW_GAP` / `COL_GAP` / `HDR_H` constants

- [ ] **Step 1: Rewire the entry point**

Replace `src/core/Core.bas:1513-1536` (comment + `Core_BuildControlPanel`) with:

```vba
' Auto-layout by display category when no button carries a page code; the manual
' code layout when any does (codes are a complete override). Prior panels and the
' pre-v0.4.0 single "Macros" tab are cleared before either mode draws. Home is
' always built, even for one page, so the workbook has one shape at every scale.
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

- [ ] **Step 2: Delete the old path**

Delete, in full: `BuildSinglePanel`, `BuildGroupedPanels`, `RenderPageLegacy`, and
the `BTN_W` / `ROW_GAP` / `COL_GAP` / `HDR_H` constants (now unreferenced).

Keep `MACRO_SHEET` and `DeleteLegacyPanel`: they are how an upgrading workbook's
stale `Macros` tab is removed.

- [ ] **Step 3: Verify nothing references the removed path**

```bash
grep -n "BuildSinglePanel\|BuildGroupedPanels\|RenderPageLegacy\|BTN_W\|ROW_GAP\|COL_GAP\|HDR_H" src/core/Core.bas
grep -n "MACRO_SHEET" src/core/Core.bas
```

Expected: the first prints **nothing**; the second prints exactly two hits (the
`Const` and `DeleteLegacyPanel`).

- [ ] **Step 4: Build, syntax-check, commit**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add src/core/Core.bas
git commit -m "feat(core)!: build Macros Home plus packed category pages; retire the single Macros tab"
```

---

### Task 9: `PkgFixture` — 75 no-op actions, excluded from the shipping bundle

Design §2.5's answer to designing a large layout while only 8 actions exist.
Deliberately uneven category sizes so a full column (8), a wrap (9), a
multi-column category (12), page packing of small categories, and a
larger-than-a-page category (26) are all exercised in one Excel session.

**Files:**
- Create: `src/packages/PkgFixture.bas`
- Modify: `build/build.ps1:64-67` (default package glob)
- Modify: `src/core/Core.bas:574-579` (`CoreCandidates` — **temporary**)

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
' Declares 75 no-op actions across 9 display categories so the v0.4.0 packed
' auto-layout can be exercised at catalog scale while only 8 real actions exist
' (design section 2.5). Category sizes are deliberately uneven so one Excel
' session covers: several small categories packed onto one page, an exactly-full
' column (8), a wrap to a second column (9), a multi-column category (12) and a
' category larger than a whole page (26 > CARD_COLS * AUTO_COL_MAX).
'
' Note "Clean & Standardize" below: it deliberately COLLIDES with PkgClean's
' display name, so two different modules feed one category. That is the whole
' point of decoupling category from package, and it must be verified, not assumed.
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
                           "No-op fixture action " & n & " in category " & cats(ci) & _
                           ". This sentence is deliberately long enough to test that a " & _
                           "card description wraps and auto-fits instead of clipping.", _
                           cats(ci)
        Next k
    Next ci

    ' Fields that exercise the collapsed off band. fx_* keys cannot collide with
    ' the shipped k_* / clean_* namespaces.
    Core_AddSection reg, "FIXTURE SETTINGS"
    Core_AddField reg, "fx_always", "Always shown (declares no usedBy)", "text", "always", _
                  "simple", "Package-wide: stays in place whatever is turned off."
    Core_AddField reg, "fx_a1", "Used by Fixture 1 only", "text", "one", _
                  "simple", "Collapses into the off band when Fixture 1 is set to Off.", _
                  "fixture.a1"
    Core_AddField reg, "fx_a2a3", "Used by Fixture 2 or Fixture 3", "text", "two", _
                  "simple", "Stays up while EITHER action is runnable; collapses only when both are Off.", _
                  "fixture.a2,fixture.a3"
    Core_AddField reg, "fx_chain", "Used by Fixture 7 (test 'Chains only')", "text", "chain", _
                  "simple", "Fixture 7 set to 'Chains only' still runs, so this must STAY in the main list.", _
                  "fixture.a7"
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

- [ ] **Step 3: Register the fixture temporarily**

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
in the shipping bundle — but it is still temporary, and Task 11 removes it.

- [ ] **Step 4: Build both bundles, syntax-check both, commit**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
node scripts/vba-syntax-check.mjs dist/MacrosBundle.fixture.bas
git add src/packages/PkgFixture.bas build/build.ps1 src/core/Core.bas
git commit -m "test: add PkgFixture (75 no-op actions, 9 categories) and keep it out of the default bundle"
```

Expected: the first build reports `Core + 2 package(s): PkgClean.bas,
PkgImportExport.bas` — **`PkgFixture.bas` must NOT appear**. The second reports
`Core + 3 package(s)`. `dist/` is gitignored, so neither bundle is committed.

---

### Task 10: Owner verification in Excel — THE gate

There is no headless VBA compiler and no headless Excel. Everything above is
structurally checked and functionally unverified. **Only the owner can run this
task, and nothing may be described as verified until it passes.**

#### Before you start: the three workbooks, and which file goes in each

You need **three separate workbooks**. Import **one** `.bas` per workbook, never two —
both bundles declare the same module name and share every `Core_*` procedure, so a
second import collides.

| # | Workbook | Import | Covers |
|---|---|---|---|
| **A** | A **scratch copy** of any workbook. Its own data is irrelevant — nothing in these steps touches it. | `dist/MacrosBundle.fixture.bas` | Steps 1–5b |
| **B** | A **fresh copy of the real workbook** you actually use. | `dist/MacrosBundle.bas` | Step 6 |
| **C** | A workbook **already finalized under v0.3.0** — recipe below. | `dist/MacrosBundle.bas`, imported *over* the v0.3.0 one | Steps 7 and 7b |

**Workbook A gets the *fixture* bundle:** the 8 real actions plus 74 fake ones. The
fake ones exist only to force page-packing cases a real 8-action deployment never
reaches, and they do nothing but show a dialog. **The 5 `Import / Export` buttons are
real, though — do not click those here.** Step 6 is where real actions get exercised,
on workbook B.

**Workbook A also needs a scratch tab named `TabName`** for Step 5b's anchored-placement
checks. Create it before you start; it can be empty.

**What "finalized under v0.3.0" means, and how to make workbook C.** Step 7 checks
that upgrading an *existing* deployment does not silently lose its saved layout.
v0.4.0 inserted an `Availability` column between `Action` and `Placement` in the
hidden `_Config` sheet; v0.3.0 wrote only `Action | Page code`. So workbook C has to
be one whose `_Config` was written **by v0.3.0** — you cannot fake it by editing
cells, because the migration keys off a hidden format marker that v0.3.0 never wrote.
Build one from `main`, which is still v0.3.0:

```
git worktree add ../sheetforge-v030 main
cd ../sheetforge-v030
pwsh build/build.ps1 -Out dist/MacrosBundle.v030.bas
```

Import that `MacrosBundle.v030.bas` into a third scratch copy, run `Create Setup Tab`,
type a few page codes (say `1A` on one action and `2B` on another), and click
`Finalize Setup Tab`. That workbook is now "finalized under v0.3.0". **Keep it** —
Steps 7 and 7b both need it, and 7b's escape-hatch check works *only* on it. Clean up
afterwards with `git worktree remove ../sheetforge-v030`.

- [x] **Step 1: Fixture pass — compile**

Import `dist/MacrosBundle.fixture.bas` into a scratch workbook, then
`Debug ▸ Compile VBAProject`. Expected: clean. A missing-argument or undefined-name
error here is exactly what the structural gates cannot see.

- [x] **Step 2: Fixture pass — packing and geometry**

Run `Core_BuildSetup`, then `Finalize Setup Tab`, and confirm:

- [x] **Two modules feed one category, and that category wraps to a second column.**
      The fiddliest case in the build. Full expected result immediately below.

**What you should see** — two adjacent columns, somewhere on `Macros 1`:

| Column header | Buttons in it |
|---|---|
| `Clean & Standardize` | `Remove Duplicates`, `Trim & Clean Text`, `Fill Blanks Down`, `Fixture 4`, `Fixture 5`, `Fixture 6` — **6** |
| `Clean & Standardize (cont. 2)` | `Fixture 7`, `Fixture 8` — **2** |

**Why it is shaped like that.** Two *different* modules deliberately declare the
*same* category with *different* capitalisation: `PkgClean` says `Clean &
Standardize` (3 actions), `PkgFixture` says `CLEAN & STANDARDIZE` (5). The engine
must treat those as **one** category of **8**. A plain VBA dictionary would not — it
would produce two separate categories — so this is a real check, not a formality.
8 then exceeds the 6-per-column cap, so the category wraps into a numbered
continuation column.

The first three buttons have no number because they are the real `PkgClean` actions;
the numbered ones are fixtures. Fixture numbering runs across *all* categories and
`Ingest & Import` consumes 1–3, which is why this category starts at 4.

**Any of these is a failure:** one column of 8 (column cap broken); two separately
titled categories, e.g. `Clean & Standardize` *and* `CLEAN & STANDARDIZE` (the merge
failed); or a bare `(cont.)` with no number (see the `Reconcile` item below).

- [x] **Ten categories, 82 buttons in total.** Count the column headers across all
      the `Macros N` tabs against this — the counts are per *category*, which may
      span more than one column:

| Category | Buttons |
|---|---|
| `Import / Export` | 5 |
| `Clean & Standardize` | 8 |
| `Ingest & Import` | 3 |
| `Validate` | 5 |
| `Reconcile` | 9 |
| `Report` | 14 |
| `Export` | 4 |
| `Period Close` | 6 |
| `Admin & Housekeeping` | 2 |
| `Bulk Fixture Set` | 26 |
| **Total** | **82** |

That 82 is 74 fixture actions plus the 8 real ones (5 from `PkgImportExport`, 3 from
`PkgClean`). Only `Import / Export` and `Clean & Standardize` contain anything real.

- [x] **Small categories share pages** — no page has one 3-button category alone
      while a neighbouring page has room.
- [x] `Validate` (5 buttons) sits in **one** column with **no** `(cont.)` in its
      header. 5 is under the 6-per-column cap, so it must not wrap.
- [x] `Reconcile` (9) wraps to a second column headed `Reconcile (cont. 2)` —
      6 buttons then 3. **The number matters:** a bare `(cont.)` with no number is
      old behaviour and is a **failure**, because without it one page could show two
      identically-titled columns.
- [x] `Bulk Fixture Set` (26) spans **more than one tab**. At 6 per column and 3
      columns per tab, one tab holds at most 18, so 26 cannot fit on one.
- [x] **No other category is split across two tabs.** A category only splits when it
      needs more than 18 buttons; `Bulk Fixture Set` is the only one that does. A
      category wrapping to a second *column* on the same tab is normal and expected
      (`Clean & Standardize` and `Reconcile` both do).
- [x] Page titles name the categories on that page, joined by `|` when packed —
      no bare `Macros - Page N` anywhere.
- [x] **Descriptions do not clip.** Every card shows its full sentence, wrapped, at
      9pt — **except `Fixture 1 - Ingest & Import`, whose description is
      deliberately empty** (it is the test that the minimum-height floor holds a
      card with no text). A single blank card description is correct. Every other
      description must show in full; if one is cut off, see the AutoFit check
      immediately below — that is the likely cause.
- [x] **THE AUTOFIT CHECK — look for this one deliberately; it is the single most
      likely thing to be wrong on first open. The primary check is
      `Macros 1!F3`: `Clean & Standardize (cont. 2)`, 29 characters, 11pt bold,
      in the 30-character-wide card column, with no height floor.**
      `FitDescRows` and `RenderPage`'s category-header row AutoFit are the first
      `AutoFit` calls this build's *packed-page* code path has ever executed
      (Tasks 6–7 were unreachable), and after Task 8 they run with
      `ScreenUpdating = False` **and** `Calculation = xlCalculationManual` from
      `Core_SpeedOn`. AutoFit on a `WrapText` cell under suppressed screen
      updating is a well-known spot for Excel to hand back a **one-line** height.
      Note what is and is not new here: AutoFit under suppressed `ScreenUpdating`
      is **not** new — `Core_BuildSetup` has done exactly this since v0.3.0 and
      it works. The genuinely new variable this release adds is
      `Calculation = xlCalculationManual` from `Core_SpeedOn`, now wrapping the
      panel-build pipeline too.
      There are **two** AutoFit sites, with different exposure:
        1. `FitDescRows` (card descriptions) — half-protected by the
           `DESC_MIN_H` floor (24pt).
        2. `RenderPage`'s category-header row (`panel.Rows(CARD_HDR_ROW).AutoFit`)
           — a `WrapText` cell with **no floor either**, and nobody flagged it
           before this fix wave. The fixture's most borderline case is the merged
           category's continuation header, **`Clean & Standardize (cont. 2)`**
           (29 characters, 11pt bold, in the 30-character-wide card column) — the
           closest any header comes to wrapping. Check: on `Macros 1`, cell
           **F3** (the merged category's second column — column D holds the
           first, un-suffixed `Clean & Standardize` header), confirm **both
           lines** of the wrapped header are visible, not clipped to one.
      **This is now the primary AutoFit check, and the risk is no longer
      asymmetric.** Every `Macros N` page is freshly created by `Worksheets.Add`,
      which activates it, so its AutoFit always runs on the active sheet - on a
      first build and on every rebuild alike. There is no "only from the second
      build onward" caveat any more, and no reused sheet to blame. If
      `Macros 1!F3` wraps correctly, tick this.
- [x] Cards align across all three columns — button tops and description tops line
      up on every row.
- [x] `< Previous` and `Next >` sit in cells **B1** and **D1** on every page
      (`<< Home` no longer exists to occupy B1), work, and are absent on the
      first/last page.
- [x] No `Macros` tab remains.
- [x] **Timed build.** Note the wall-clock time from clicking `Create Setup Tab` to
      the Setup tab appearing, and again for `Finalize`. The owner measured
      **several minutes** for the first build before the registry was memoized.
      Anything still in minutes means the cache is not being hit - check that
      `InvalidateRegistry` was not also wired into `Core_Run`.
- [x] The action catalog's third column header reads **`Placement`**, not
      `Page code`, in BOTH the Action catalog block and the Chains block.
- [x] The `Availability` cells still have a working dropdown offering exactly
      `Button` / `Chains only` / `Off`. It is now applied to the whole block in one
      call instead of cell by cell, so this is confirming a refactor did not drop it.
- [x] No `Macros Home` tab exists. If you are rebuilding a workbook that had one
      from an earlier build, confirm it has been **deleted**, not left behind.
- [x] The build lands you on **`Macros 1`**, cell A1.
- [x] The completion dialog does **not** mention an index.
- [x] **The nav row on a SINGLE-page build.** `RenderNavRow` still sets row 1 to
      18pt unconditionally, but with only one page there is no `< Previous` and no
      `Next >` — and `<< Home` no longer occupies B1 — so row 1 renders as a blank
      18pt band. Behaviourally correct, but your original complaint was clutter and
      a reserved-but-empty band is the mirror image of it. **Judge it:** if it reads
      as wasted space, suppressing the row height when both links are absent is a
      two-line change.

- [x] **Step 3: Fixture pass — the three availability states**

- [x] The catalog block reads `Action | Availability | Placement`, every row
      `Button`, and the Availability cell has a **working dropdown** offering
      exactly `Button`, `Chains only`, `Off`.
- [x] Set `fixture.a1` to `Off`, Finalize: its button disappears, **its row stays**
      in the catalog.
- [x] Run `Core_Run "fixture.a1"` — it **refuses** with the "switched Off" message
      and does not run. **Not via Alt+F8:** that dialog lists only `Public Sub`s
      taking no arguments (just `Core_BuildSetup`, `Core_FinalizeSetup`,
      `Core_SpeedOn`, `Core_SpeedOff`), and `Core_Run` takes an action id. Use the
      VBA editor's Immediate window instead — **Alt+F11**, then **Ctrl+G**, type the
      line, press Enter. The dialog appears exactly as it would from a button.
- [x] Set `fixture.a7` to `Chains only`, Finalize: no button. In the Immediate
      window, `Core_Run "fixture.a7"` refuses with the "Chains only" message.
- [x] Build a chain containing `fixture.a7`: it **runs** as a step, one combined
      summary, no per-step dialogs.
- [x] Build a chain containing `fixture.a1` (which is `Off`), then Finalize.
      **Judge this on the DIALOG, not on the Macros tabs** — see immediately below.

**Expected dialog.** A warning containing `OFF ACTION IN CHAIN: '<your chain name>'
uses fixture.a1, which is switched Off.` and ending `1 problem(s) above must be fixed
before activating.` Finalize then stops.

**Read this before judging the result — it is the confusing part.** A blocked
Finalize exits *before* rebuilding anything. So the chain's button from your **last
successful** Finalize is **still sitting on its Macros tab**, and **clicking it still
runs**, `fixture.a1` step included.

Both of those are correct. `Off` was typed on the Setup tab but never *activated*,
and the engine dispatches from the last activated configuration — not from what is
currently typed on screen. **A working chain button after a blocked Finalize is the
block working**, because nothing was rebuilt.

**It is a failure only if the dialog says `Configuration activated.`** If it does,
stop — an `Off` action would be reachable inside a chain, which contract §6 forbids
outright.

- [x] **Paste** (do not type) `maybe` into an Availability cell — `AvailDropdown`
      validates with `AlertStyle:=xlValidAlertStop`, which hard-refuses a typed
      invalid entry before it ever reaches the cell, so typing `maybe` cannot
      exercise this check at all; pasting bypasses cell validation and is the
      only way to get an unrecognised value into the cell. Finalize **BLOCKS**
      with `UNRECOGNISED Availability` and does not activate.
- [x] **Clear the Availability cell of an action that is currently `Off`** — not one
      already showing `Button`, or nothing will appear to happen and the check
      proves nothing. Blank means "never set", which reads as `Button`. Expected:
      Finalize activates with no complaint, **and that action's button comes back**
      on its Macros tab. Set `fixture.a1` back to `Off` afterwards if a later step
      still needs it off.

- [x] **Step 4: Fixture pass — the collapsed band**

- [x] **The band exists, is collapsed, and holds the right field.** Set `fixture.a1`
      to `Off`, click `Finalize Setup Tab`, then click `Create Setup Tab` and scroll
      to the bottom. Confirm **(a)** a band headed **SETTINGS FOR ACTIONS YOU HAVEN'T
      TURNED ON** exists, with the row labelled `Used by Fixture 1 only` inside it;
      and **(b)** the band is **collapsed** — there is a `+` in Excel's outline
      gutter (the narrow strip left of the row numbers) that expands it when clicked.

- [x] **A relegated setting keeps its value.** This is the point of the whole
      feature: switching an action off must never throw away what you typed. Run
      these **in order** — the middle step is the one that matters:

**Step 1.** Expand the band and find the row labelled `Used by Fixture 1 only`. **The
cell to edit is the yellow one on that row** (column C, the same column every other
setting's value sits in). It currently reads `one`. Replace it with `two`.

**Step 2. Click `Finalize Setup Tab`.** ← **Do not skip this.**

**Step 3.** Click `Create Setup Tab`, accept the rebuild prompt, expand the band
again. **Expected: that cell still reads `two`.**

**If you skip step 2, the cell correctly reverts to `one` — and that is not a
failure.** `Create Setup Tab` rebuilds from your last *finalized* settings, never
from what is on screen, and the rebuild prompt says exactly that: *"DISCARDED:
anything you have typed on the Setup tab since the last 'Finalize Setup Tab'."* So
skipping step 2 tests the discard rule instead, which is also correct behaviour —
just not this check. `one` is the shipped default, which is what an unset field falls
back to.
**What this step is testing, in one sentence.** A setting that belongs *only* to
actions you have switched `Off` is moved into a collapsed band at the bottom of the
Setup tab — **never deleted**. A setting still reachable by anything stays in the
main list.

**You will not see any `fx_*` name on screen.** The Setup tab shows each field's
**label**; the `fx_*` names below are internal keys, listed only so you can match the
two. Actions likewise appear as `Fixture 1 - Ingest & Import [fixture.a1]` in the
Action catalog — caption first, id in brackets.

| Setup tab shows this label | Internal key | Belongs to | Moves into the band when |
|---|---|---|---|
| `Always shown (declares no usedBy)` | `fx_always` | nothing | **never** |
| `Used by Fixture 1 only` | `fx_a1` | `fixture.a1` | Fixture 1 is `Off` |
| `Used by Fixture 2 or Fixture 3` | `fx_a2a3` | `fixture.a2` **or** `a3` | **both** are `Off` |
| `Used by Fixture 7 (test 'Chains only')` | `fx_chain` | `fixture.a7` | Fixture 7 is `Off` — **not** when it is `Chains only` |
| `Names an action id that does not exist` | `fx_typo` | a typo'd id | **never** |
| `Advanced, used by Fixture 4` | `fx_adv` | `fixture.a4` | Fixture 4 is `Off` |
| `List used by Fixture 5` | `fx_list` | `fixture.a5` | Fixture 5 is `Off` |
| `Table used by Fixture 6` | `fx_table` | `fixture.a6` | Fixture 6 is `Off` |

Each check below is: set the Availability, click `Finalize Setup Tab`, then click
`Create Setup Tab` and look at the bottom of the rebuilt Setup tab. **The band only
redraws on `Create Setup Tab`** — Finalize alone never redraws Setup.

- [x] **A `Chains only` action keeps its settings up.** Set `fixture.a7` to
      `Chains only`. `Used by Fixture 7 (test 'Chains only')` must stay in the
      **main list**, not the band — `Chains only` still runs inside a chain, so its
      setting is still needed. Only `Off` relegates anything.
- [x] **An unknown owner never hides a field.** `Names an action id that does not
      exist` points at an action id that isn't real. It must **always** stay in the
      main list, whatever you switch off. Failing closed here matters: a typo in a
      package's manifest must never make a user's setting vanish.
- [x] **Either-owner-runnable keeps a field up.** Set `fixture.a2` `Off` but leave
      `fixture.a3` as `Button`: `Used by Fixture 2 or Fixture 3` stays in the main
      list. Now set `fixture.a3` `Off` too: it moves into the band.
- [x] **An advanced field relegates too.** Set `fixture.a4` `Off`. `Advanced, used
      by Fixture 4` normally lives under the **ADVANCED** band; confirm it moves out
      of there and into the off band, rather than staying put or disappearing.
- [x] **List and table blocks survive relegation.** Set `fixture.a5` and `fixture.a6`
      `Off`. `List used by Fixture 5` and `Table used by Fixture 6` both move into
      the band **with their rows intact** — the list still shows `seed one` /
      `seed two`, the table still shows `a1|b1|c1` across its three columns. Then
      Finalize again and re-run `Create Setup Tab`: the values must still read back.
      A multi-row block is the easiest thing to truncate when it moves, which is why
      it is checked explicitly.
- [x] **Empty band means no band.** Set every action back to `Button` and rebuild
      Setup. The **SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON** heading must be
      **gone entirely** — not present-but-empty.

- [x] **Step 5: Fixture pass — manual codes, cancel and error paths**

**Where you type these.** All placement codes go in the **Placement** column of the
**Action catalog** block on the Setup tab (third column, next to `Availability`). A
code means "put this button on page *N*, column *X*" — `1A` is page 1 column A.

- [x] **Two codes, everything else blank.** Put `1A` on one action and `2B` on
      another, leave every other Placement cell empty, and Finalize.

**Expected:** `1A`'s button is on `Macros 1` under a column headed `Column A`;
`2B`'s is on `Macros 2` under `Column B`.

**And the other ~80 buttons spread across several tabs under `Unassigned`,
`Unassigned (cont. 2)`, `Unassigned (cont. 3)` and so on. That is correct — do not
report it.** Filling in even one code switches the whole workbook out of automatic
layout, so every uncoded button falls into an "Unassigned" pile. That pile obeys the
same limits as everything else — 6 buttons per column, 3 columns per tab — so 80 of
them need many columns and therefore many tabs. At the real 8-action scale you would
see one small Unassigned column; the fixture makes it look dramatic.

- [x] **`hide`.** Type `hide` in one action's Placement cell and Finalize: its button
      is gone from every tab, but **its row is still in the Action catalog** on the
      Setup tab. Hiding a button must never make the action undiscoverable.
- [x] **Clearing every code returns automatic layout.** Empty every Placement cell
      and Finalize: the tidy category-packed pages from Step 2 come back.
- [x] **A bad code blocks.** Type `0A` (page 0 does not exist), `A1` (letters before
      digits) or `1` (no column letter) and Finalize. Expected: a warning listing
      `BAD CODE for <id>: '<what you typed>'` and ending `problem(s) above must be
      fixed before activating`. Nothing is rebuilt.
- [x] **Declining activation changes nothing.** Finalize, and at
      `Activate this configuration now?` answer **No**. No tabs are rebuilt and your
      previously activated settings are untouched.
- [x] **Declining a Setup rebuild changes nothing.** Click `Create Setup Tab` on a
      workbook that already has one and answer **No** to the rebuild prompt. The
      Setup tab is left exactly as it was, including anything you had typed but not
      yet finalized.
- [x] **Accepting a rebuild discards unfinalized typing — and warns you first.** The
      other half of the same prompt. Type something into any yellow settings cell,
      do **not** Finalize, then click `Create Setup Tab` and answer **Yes**.
      Expected: the prompt states plainly that unfinalized typing will be
      **DISCARDED** and your last finalized settings **KEPT**, it defaults to **No**,
      and after answering Yes your untyped-over value is back to whatever you last
      finalized. Losing unsaved typing here is correct; losing it *silently* would
      not be.
- [x] **Finalize twice in a row.** Click `Finalize Setup Tab`, let it finish, click
      it again. Expected: it just works. Specifically **no** *"That name is already
      taken"* error, and **no** stray `Setup (2)` or `Setup (3)` tab left behind —
      that was a real bug in v0.3.0.
- [x] **Break the build on purpose, once.** This tests what happens when your
      settings save but the tab rebuild then fails. No amount of code reading proves
      it, so it has to be forced.

**Do this:** right-click the **`Setup Menu`** tab ▸ **Protect Sheet** ▸ **OK** (no
password needed) ▸ then click `Finalize Setup Tab`.

**Expected: exactly two dialogs, in this order.**

1. `Core_BuildControlPanel error: ... Your settings are safe. Click 'Finalize Setup
   Tab' again to retry the rebuild.`
2. `Configuration activated, but rebuilding the Macros pages failed`

**A failure is: only one dialog, or a dialog claiming plain success.** In v0.3.0 this
path reported the error and then cheerfully said the rebuild had succeeded — on a
workbook whose button tabs had already been deleted, so it had none at all. Two
dialogs means the success flag is being read properly.

**Then recover:** right-click `Setup Menu` ▸ **Unprotect Sheet**, re-run Finalize, and
confirm the workbook comes back intact.

*Why `Setup Menu` specifically:* it is the one generated tab the engine **reuses**
rather than deletes and recreates, so protecting it makes the rebuild fail on a real
sheet. Two triggers people reach for instead do **not** work and verify nothing:
*hiding* `Setup Menu` (writing to a hidden sheet is perfectly legal), and renaming
`_Config` (a missing config is handled gracefully — the build just succeeds using
defaults).

- [x] **Step 5b: Fixture pass — Sheet!Cell anchored placements**

- [x] **It places, and it does not disturb the layout.** `TabName!B4` on one action,
      every other code blank, Finalize: the button appears on `TabName` at B4, the other
      ~81 buttons stay packed on numbered pages (an anchor must not flip the workbook
      into manual mode), and the button is about **110pt** wide — the one number only
      your eyes can settle. Same again for a **chain**: identical behaviour, one
      button, one placement.

- [x] **`TabName` is a guest surface, and the button behaves like one.** Four
      observations on the button you already have there: **(a)** nothing on `TabName`
      changed but the button — no cleared cells, no altered column widths or row
      heights, no formatting, no tab colour; **(b)** narrow column B — the button
      does **not** squash (`xlMove`, not `xlMoveAndSize`); **(c)** insert a row above
      row 4 — the button moves down with its cell; **(d)** click it — it runs.

- [x] **Self-healing, and every accepted spelling, in one pass.** Change the
      Placement cell and Finalize after each, confirming there is **exactly one**
      button on the target each time — never two:
      `TabName!$B$4` (same cell, absolute) → `TabName!D10` (moves) → rename the tab to
      `TabName Data` and use `TabName Data!B4` (unquoted, with a space) → `'TabName Data'!B4`
      (quoted) → finally **clear the cell** and confirm the button is **gone** and
      the action is back on a generated page.
      The quoted spelling is the one to watch: **Excel eats a leading apostrophe**,
      treating it as a "this cell is text" prefix, so the cell really stores
      `TabName Data'!B4`. You can watch it happen — your apostrophe shows in the formula
      bar and not in the cell. The engine repairs this; if it fails you get
      `NO SUCH SHEET` naming a tab sitting in plain view.

- [x] **What blocks Finalize.** Each must refuse to activate, name the button, and
      change nothing:
      `MICC!B4` → `NO SUCH SHEET` · `Macros 1!B4` → `RESERVED SHEET` (whether or not
      that tab currently exists — engine-owned names are checked first, deliberately,
      so you are never told to create a tab the engine would destroy) ·
      `TabName!B4:C9` → `BAD CODE` · two actions both on `TabName!B4` → `DUPLICATE TARGET`
      naming both · and `TabName!$B$4` against `TabName!B4` on two actions **must also**
      collide, being one cell spelled two ways.
      With a *valid* anchor, the report instead shows a soft `Placed on your own
      sheets` list in the same `caption [id]` format as the `Hidden` and `Unassigned`
      lists, and does **not** block.

- [x] **A PROTECTED target sheet — the single most important item in this section.**
      Right-click `TabName` ▸ Protect Sheet ▸ OK (the default protects `DrawingObjects`),
      then Finalize. Expected: **the build COMPLETES** — it must not die with
      "rebuilding the Macros pages failed" — and the `Configuration activated.`
      dialog reports `1 anchored button(s) could NOT be placed.` followed by two
      bullets, the second telling you to unprotect and Finalize again.
      This is the one assumption that could not be verified without Excel
      (`ws.Buttons.Add` raising 1004 on a protected sheet), and before the fix it
      took down the entire build. **If the dialog says only `Configuration
      activated.` with no count, the note is not reaching Finalize — a real finding,
      not cosmetic.**
      **Known residual, not a bug:** a button an *earlier* build put there survives,
      because the sweep cannot delete it either. The engine must never unprotect your
      sheet. Unprotect and re-Finalize to clear it.

- [x] **A target that disappears, and the empty workbook.**
      *You are renaming the worksheet **tab**, not the Placement cell. The cell keeps
      saying `TabName!B4`; the point is to leave the activated config pointing at a tab
      that is no longer there.*
      With `TabName!B4` anchored and finalized, rename the `TabName` tab to `MIC2`. **Do not
      Finalize** — Finalize re-validates first and correctly refuses with
      `NO SUCH SHEET` (the item above already checked that), so it can never show
      this dialog. Instead run `Core_BuildControlPanel` from the Immediate window
      (**Alt+F11**, **Ctrl+G**): it rebuilds from the activated `_Config` **without**
      re-validating, which is exactly the state a user reaches by renaming a tab some
      time after finalizing. Expected: the `Setup Complete` dialog reports
      `1 anchored button(s) could NOT be placed.` and its **first** bullet applies —
      that action now has **no** button anywhere, which is what the message exists to
      tell you. Rename the tab back to `TabName` before continuing.
      Then type `hide` in the **Placement** column — the third column of the Action
      catalog block, to the right of Availability, the same column that holds
      `TabName!B4`. **Not** Availability: that is a dropdown of `Button` / `Chains only`
      / `Off`, it will not let you type, and pasting `hide` into it lands an
      unrecognised token that blocks Finalize with one `UNRECOGNISED Availability`
      error per cell. Code every action row `hide` except the anchored one, **and
      the chain rows in the Compose Your Buttons block too** — chain buttons come
      through the same funnel, and one left uncoded keeps a page alive and hides the
      zero-pages wording this item exists to check. Leaving only the one anchored
      button, Finalize: a **success** dialog ending
      `1 button(s) placed on your own sheets.`
      To see the zero-pages wording as well, run `Core_BuildControlPanel` from the
      Immediate window (**Alt+F11**, **Ctrl+G**, type it, Enter — it is not in the
      Alt+F8 list, which shows only zero-argument macros). That non-silent path
      reads `No macro pages were built - every button is hidden, switched Off, or
      anchored to one of your own sheets.` — not `Built 0 macro page(s)`. The note
      must appear exactly **once per dialog**, never doubled.

- [x] **Step 6: Real pass — the shipping bundle**

Fresh copy of the real workbook, import `dist/MacrosBundle.bas` (no fixture),
`Debug ▸ Compile VBAProject`, `Core_BuildSetup`, Finalize:

- [x] **Exactly ONE `Macros` tab, holding two columns** — `Import / Export`
      (5 buttons) and `Clean & Standardize` (3). All 8 fit on one tab. The point of
      this check is that a real deployment must not end up with more tabs than
      v0.3.0 had; the 7 tabs you saw with the fixture were the fixture's doing.
- [x] **All 8 buttons work, and cancelling is clean.** Click each one and confirm it
      does its real work on real data. On every action that opens a range picker or
      file dialog, run it a second time and press **Cancel**: it should simply stop —
      no error dialog, nothing half-done.
- [x] **A failed action must not poison the next one.** Force a failure — point a tab
      setting at a tab that does not exist, or a file setting at a prefix matching no
      file — and run it. Expected: Finalize warned you beforehand, and afterwards
      **Excel still behaves normally**: the screen redraws, formulas recalculate, and
      the next button works. A macro that dies with screen updating off leaves Excel
      looking frozen; that is what this catches.
- [x] **The duplicates header field, all four states.** On Setup, `Does the range you
      pick include a header row?` is present with a working **Yes/No dropdown**
      defaulting to **Yes**. Then:
      **Yes** + a range whose first row IS a header → the header survives, uncompared.
      **No** + `M2:M12` with no header → **all eleven values considered**, not ten —
      your original symptom.
      **Paste** (do not type) `ye` into the cell — typing is refused by the dropdown,
      so pasting is the only way in → Finalize **BLOCKS** with
      `UNRECOGNISED Yes/No value`.
      **Clear** the cell → reads as the default `Yes` and activates.
- [x] **Sheet-qualified messages, all three Clean actions.** With a button on one tab,
      run `Remove Duplicates` against a range on a **different** tab: the completion
      dialog must read `TabName!M2:M12`, not `M2:M12`. Repeat for `Trim & Clean Text` and
      `Fill Blanks Down`. Then rename that tab to something needing quotes (`TabName Data`
      or `2024`) and confirm it comes back quoted: `'TabName Data'!M2:M12`.
      **If any of them names the WRONG tab, stop and report it** — that would mean the
      original diagnosis was wrong and something really was acting on the wrong sheet.

- [x] **Step 7: Upgrade pass — fail-closed migration**

**This is the check that protects a real user's existing work on upgrade.** Use
**workbook C** — the one finalized under v0.3.0, per the recipe at the top of Task 10.
v0.3.0 stored `Action | Page code`; v0.4.0 stores `Action | Availability | Placement`,
so the saved codes have to survive being shifted one column to the right.

- [x] **Upgrade it.** In the VBA editor (**Alt+F11**), **remove the existing
      `MacrosBundle` module first** — right-click it ▸ Remove ▸ **No** when asked to
      export. Both bundles declare `Attribute VB_Name = "MacrosBundle"`, so importing
      without removing gives you a second module named `MacrosBundle1` with every
      procedure duplicated, which fails to compile. Then import
      `dist/MacrosBundle.bas` and run `Create Setup Tab`.
      **Removing the module does not touch `_Config`** — the saved v0.3.0 settings
      live on that hidden worksheet, which is the whole point of the test.
- [x] **Every code you set under v0.3.0 is still there**, now under the `Placement`
      header, and every `Availability` cell reads `Button` (v0.3.0 had no such
      column, so everything defaults to "shown as a button").

**If any code comes back blank, stop.** Do not retype them — a blank means the
migration dropped saved settings, which is the exact failure this step exists to
catch, and retyping hides it.

- [x] **Now corrupt one deliberately.** Open the hidden `_Config` sheet (see the note
      in Step 7b for how — it cannot be unhidden from Excel's normal menu) and replace
      one of your saved page codes with `zzz`, then set the **sheet** back to
      `2 - xlSheetVeryHidden` (re-hide the *sheet*, not a column).
      **Which cell:** `Create Setup Tab` rewrites only the *Setup* sheet; `_Config` is
      rewritten only by Finalize, which has not run since the upgrade. So `_Config` is
      still in v0.3.0 layout — column C `Action`, column D the **page code** (`1A`,
      `2B`), column E empty, and **no Availability column at all**. Overwrite the code
      in **column D**. If column D shows `Button` instead, Finalize has run and
      rewritten `_Config` in v0.4.0 format: workbook C is spent and must be rebuilt
      from the v0.3.0 bundle. Run `Create Setup Tab` again. Expected:
      it **refuses to run**, names the action whose code it could not read, and
      leaves your Setup tab exactly as it was. Refusing is the correct behaviour —
      rebuilding around an unreadable value risks discarding it.

- [x] **Step 7b: Previously-unverified behaviour (added in the final-review fix
      wave)** — shipped but never exercised by the checklist above.

**First: how to open the hidden `_Config` sheet.** Several checks below need you to
edit it directly. It is stored *very* hidden, so **Excel's Format ▸ Sheet ▸ Unhide
will not list it**. Do this instead:

1. **Alt+F11** to open the VBA editor.
2. Find `_Config` in the Project pane (top left).
3. Select it, then in the **Properties** pane (bottom left) set **`Visible`** to
   **`-1 - xlSheetVisible`**.
4. Switch back to Excel and edit it. Set `Visible` back to `2 - xlSheetVeryHidden`
   when you are done.

Inside `_Config`, the **format marker** is the cell in **column J** on the row whose
column A reads `k_core_layout`. That marker is how the engine knows which version
wrote the file. On a v0.4.0 workbook it reads **`v1.2`**.

**The next two items run on workbook C**, continuing straight on from Step 7 — no need
to keep workbook B alive. Both verdicts are decided by the marker cell's contents
alone, so a blank starting value works as well as `v1.2`. Two conditions:

- **Restore J to blank afterwards, NOT `v1.2`.** C's `_Config` must keep looking
  v0.3.0-written for the escape hatch below. A valid marker means that path never
  runs and the check passes for the wrong reason.
- **Answer No on the damaged-marker item.** Yes runs the guarded rebuild, discards
  C's v0.3.0 config, and spends the workbook — it would have to be rebuilt from
  `MacrosBundle.v030.bas`.

- [x] **A damaged marker asks before discarding anything.** Open `_Config`, put
      nonsense (e.g. `xxx`) into the marker cell, re-hide, and run `Create Setup
      Tab`. Expected: a confirmation prompt that spells out explicitly what will be
      **DISCARDED** and what will be **KEPT**, and that defaults to **No**. Answering
      No must leave everything untouched. It must never rebuild silently.
- [x] **A marker from a newer version refuses outright.** Put a higher version string
      into the marker cell (e.g. `v9.9`), re-hide, and run `Create Setup Tab`.
      Expected: it **refuses** and tells you to re-import that version's bundle. This
      is the guard against an old build stomping on a newer format it cannot read —
      it must not offer to rebuild.
- [x] **The no-Setup-tab escape hatch.** **This one must be done on workbook C, the
      v0.3.0 one** — on a v0.4.0 workbook the marker is present and valid, so this
      path never runs and the rebuild proceeds silently, which would look like a
      failure and is not.

  On workbook C: leave the corrupted `zzz` code from Step 7 in the **old** code
  column (column D, not E), **delete the Setup tab entirely**, then click `Create
  Setup Tab`. Expected: a confirmation prompt naming exactly what will be discarded —
  not a silent rebuild. This is the recovery route for a user who has no Setup tab
  left to correct the value on.
**Reset workbook A before the last two items.** Step 5b left every action row *and*
every chain row coded `hide` in the Placement column, plus one anchor on `TabName`. Clear
that whole column — actions and chains — and Finalize, or the two items below render
nothing and look like failures. You should be back to the fixture's packed pages.

- [x] **Chain cards: bounded, and always last.** Build a chain with **5 or more
      steps**, Finalize, and look at its button's description. Expected: a partial
      step list ending `+N more - see the Setup tab` — a long chain must not draw one
      enormous description row that stretches every card beside it — and you must
      **never** see `+0 more`, which would claim there is more when nothing was left
      off. In the same build, confirm the reserved `Chains` category renders **after
      every other category**, regardless of when you authored the chain.

- [x] **Shrinking back to a small deployment leaves no debris, and an empty one is
      legal.** On workbook A: delete the `MacrosBundle` module, import
      `dist/MacrosBundle.bas` (the real 8-action bundle) instead, and Finalize.
      Expected: the tabs shrink back to one, with **no abnormally tall rows** left
      over from the fixture's longer descriptions — rows that grew for big text must
      shrink again when it goes away.
      Then **delete every chain** (otherwise a chain pointing at an `Off` action makes
      Finalize refuse and you never reach the empty state), set every action to `Off`,
      and Finalize. Expected: **no error** — an empty workbook is a legitimate
      configuration, not a failure — and Finalize's own dialog reads
      `Configuration activated.` followed by `No macro pages were built - every button
      is hidden, switched Off, or anchored to one of your own sheets.`
      **It must NOT say `The Macros pages have been rebuilt to match your Setup
      answers.`** — that sentence is false when none were built, and it sends the user
      looking for tabs that were never meant to exist.
      **Found by this gate and fixed in v0.4.0.** Finalize calls
      `Core_BuildControlPanel silent:=True`, so the zero-pages sentence originally
      reached nobody on the normal path — the same defect class as the anchored-button
      note that `mLastPlacedNote` was added to carry out, missed in the same place.
      `mLastNoPages` now carries the verdict out, and `NO_PAGES_MSG` holds the
      sentence once so the two callers cannot disagree.

- [ ] **Step 8: Record the result**

Paste the completed checklist into the PR. Any failure means the corresponding task
is not done.

---

### Task 11: Tear down the fixture, refresh docs, ship v0.4.0

- [ ] **Step 1: Remove the fixture and its two hooks**

```bash
git rm src/packages/PkgFixture.bas
```

Delete `c.Add "Fixture"` from `CoreCandidates`. In `build/build.ps1`, restore the
plain default glob — a dangling filter invites a future fixture to be added
silently:

```powershell
} else {
    $pkgFiles = Get-ChildItem -LiteralPath $pkgDir -Filter 'Pkg*.bas' |
                Sort-Object Name | ForEach-Object { $_.FullName }
}
```

- [ ] **Step 2: Verify the fixture is gone everywhere**

```bash
grep -rn "Fixture" src/ build/ ; rm -f dist/MacrosBundle.fixture.bas
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: `grep` prints **nothing**; the build reports `Core + 2 package(s)`.

- [ ] **Step 3: Bump `VERSION`**

```bash
echo -n "0.4.0" > VERSION
```

- [ ] **Step 4: Add the `CHANGELOG.md` entry**

Under `## [Unreleased]`, insert:

```markdown
## [0.4.0] - 2026-07-24

### Added

- **`Macros Home`** — a generated index tab, always first and always built: one row
  per page with its button count and a link, a full **button catalog** (every
  button, its complete description, and the page it is on), a `Setup Menu` link,
  and a legend explaining page codes and availability.
- **Capacity-packed auto-layout.** With no page codes filled in, buttons
  distribute automatically: up to three category-headed columns per page, up to 6
  buttons per column, small categories sharing a page, and a category never split
  across tabs unless it needs more columns than a page holds (more than 18
  buttons). Filling in even one code still hands total control back to the codes.
- **Legible descriptions.** Button pages are now an **unmerged cell-card grid** —
  button anchored to a content cell, its description in the cell beneath, that row
  auto-fitted. Grouped pages previously had no descriptions at all, and merged
  cells (which Excel's AutoFit ignores) are forbidden on the grid so nothing clips.
- **Navigation.** A persistent `<< Home` link in a fixed position on every page,
  plus `< Previous` / `Next >`.
- **Three-state per-action availability.** The placement block is now the **Action
  catalog** (`Action | Availability | Page code`) with a validated dropdown:
  `Button` (shown and runnable), `Chains only` (no button, runs only as a chain
  step), `Off` (no button, runs nowhere). `Off` means off — the dispatcher refuses
  it and Finalize **blocks activation** for any chain that uses one. An
  unrecognised value blocks Finalize rather than defaulting to `Button`.
- **`usedBy` on a field** — settings belonging only to `Off` actions collect into
  one collapsed *"Settings for actions you haven't turned on"* band at the bottom
  of Setup. Relegated, never omitted: Setup re-seeds from the finalized `_Config`,
  so omitting them would cost a second Finalize and discard un-finalized typing. A
  `Chains only` action still runs, so it keeps its settings in the main list.
- **`category` on an action** — optional, defaults to the package name, and
  decoupled from the module so two packages can feed one category and a future
  package split cannot move a button in the user's mental model.
- **A machine-owned config schema marker** in the hidden id column, with
  **fail-closed** migration: a v0.3.0 config is detected and its page codes read
  from the column they are actually in, every recovered code is validated, and on
  any conflict `Create Setup Tab` aborts with a report instead of rebuilding.
- Contract frozen at **v1.2.0** (additive).
- **Build:** fixture/harness packages are kept out of the default bundle.

### Removed

- **The single `Macros` tab default.** With no page codes, v0.3.0 built one tab and
  stacked every button on it; v0.4.0 packs categories onto pages instead, and the
  stale `Macros` tab is deleted on rebuild. Manual page codes are unaffected.

### Fixed

- `RenderPage` no longer calls `Range.Select` on a possibly-inactive sheet — it
  only ever worked because rebuilding deleted and re-added every panel, which
  happens to activate it. Selection now happens once, on `Macros Home`.
```

- [ ] **Step 5: Refresh `TODOS.md`**

Rewrite the `Current:` paragraph (`TODOS.md:8-13`) for v0.4.0 / contract v1.2.0
with track 2 (the catalog waves) next; move the track-1 line from
`## Next — active` into `## Done (recent)`.

- [ ] **Step 6: Regenerate the code map**

```bash
node scripts/code-map.mjs
```

Generated output — never hand-edit. Committing it in the same PR keeps its
freshness stamp honest.

- [ ] **Step 7: Final gate, then commit**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add -A
git commit -m "chore: release v0.4.0 (Macros Home, packed card pages, three-state availability; contract v1.2.0)"
```

- [ ] **Step 8: Review and ship**

Run `/review` over the `src/` diff, then `/ship`. The PR body must carry Task 10's
completed checklist — the build is not a substitute for it.

---

## Self-Review

**Spec coverage** — design §2 and the six panel findings against the tasks:

| Requirement | Task |
|---|---|
| `Macros Home`: legend, per-page rows with counts + links, `Setup Menu` link | 7 |
| Home action catalog with full descriptions (owner's Q4 addition) | 7 |
| Pages titled by category, columns headed by category | 6, 7 |
| Every button followed by its description, unclipped | 6 (card grid + `FitDescRows`) |
| Persistent fixed-position Home link + Previous/Next | 6 (`RenderNavRow`) |
| Capacity-packed auto-layout, no category straddling (panel Q2/Q3) | 7 (`BuildAutoLayout`) |
| Home always built, even for one page (owner's Q2 call) | 8 |
| Manual codes remain a complete override; `hide` still omits | 6 (`BuildManualLayout`) |
| Tighter spacing | 6 Step 1 |
| No search, no Recent Actions, Favorites deferred | not built — correct |
| `Action \| Availability \| Page code` with dropdown | 3 |
| Three states; `Off` unrunnable; chain-with-`Off` blocks Finalize (owner's Q5) | 3 |
| Unrecognised value blocks rather than defaults | 3 (`ParseAvail` → `""`) |
| Defaults to available → v0.4.0 is a no-op for a current user | 3 + 10 Step 7 |
| Unavailable actions stay listed | 3 + 10 Step 3 |
| Fields always render; `Off`-only ones collapse | 5 |
| `Chains only` keeps its fields in the main list | 5 (`FieldInUse` via `IsRunnable`) + 10 Step 4 |
| Machine-owned schema marker, validated, fail-closed (panel Q6) | 4 |
| Every commit compiles (panel Q1) | 6, 7 dead-but-compilable; 8 cutover |
| Contract additive bump to v1.2.0 | 1 |
| `PkgFixture`, 75 no-op actions, 9 categories, never ships, deleted after | 9 + 11 |

**Ordering, designed out rather than annotated:** two forward references existed in
the first cut of this revision — `FieldInUse` calling `IsRunnable`, and
`ReadAvailMap` calling `BlockSchema`. Both are now declared in **Task 3**, the task
that first needs them, so the tasks run in plain numeric order and **every single
commit compiles on its own**. Nothing in this plan requires the executor to
reorder steps; if a task seems to need a procedure from a later task, that is a
bug in the plan, not an instruction to skip ahead.

**Placeholder scan:** none. Every code step carries actual VBA or PowerShell. The
prose-only steps are Task 10's owner checks (manual by nature) and Task 11 Step 5
(`TODOS.md` content depends on what ships).

**Type/name consistency checked:** `layout` throughout Tasks 6–8; `availMap`
throughout Tasks 2–5; button arrays are 5-element after Task 3 (`e(4)` = category,
read in `LabelManualPages`, `BuildAutoLayout`); `RenderTier` takes 8 arguments at
all four call sites; `HasAdvanced` takes 2 at its single call site. New names —
`ParseAvail`, `AvailLabel`, `ReadAvailMap`, `AvailFor`, `HasButton`, `IsRunnable`,
`AvailDropdown`, `BlockSchema`, `ConfigSchemaStatus`, `FieldInUse`, `HasAnyTier`,
`RenderOffBand`, `CollapseRows`, `SetPageTitle`, `SetColHeader`, `PageButtonCount`,
`SetupCardGrid`, `CardCol`, `AddBtnInCell`, `FitDescRows`, `RenderNavRow`,
`NavLink`, `RenderAllPages`, `BuildManualLayout`, `LabelManualPages`,
`BuildAutoLayout`, `CategoryOrder`, `LabelAutoPages`, `BuildHomePage`,
`HomeCatalog`, `HomeLegend`, `HomeLine` — are each declared once and none collides
with an existing procedure (`build.ps1` fails the build if one did). `IsRunnable`
is deliberately not `IsTrue`, which already exists with different semantics.
