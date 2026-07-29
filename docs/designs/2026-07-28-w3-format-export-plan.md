# W3 - Format & Present, plus Export

> **For agentic workers:** REQUIRED SUB-SKILL: use `superpowers:subagent-driven-development`
> to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship the eight catalog items in the `format` and `export` domains as two new
packages, `PkgFormat` and `PkgExport` - five presentation-only actions that change how a
report looks and never a stored value, and three actions that get finished work out of the
workbook (frozen values, PDFs, one file per tab), each either refusing on a real collision or
asking before doing something that cannot be undone with Ctrl+Z.

**Architecture:** Two thin packages on Core, exactly like `PkgClean` and `PkgImport`. The one
piece of shared engine this wave introduces - flattening formulas and dynamic-array spills to
static values - moves into Core as `Core_FreezeValues`, because it already has two consumers
the moment `export.freeze_values` exists: the promoted action itself, and
`importexport.export`, which has done this inline since before this wave. `PkgImportExport`
is edited to call the Core engine instead of carrying its own copy - the alternative
(a second, independently-written copy of the same 30-line flatten loop) is the cross-package
duplicate `CLAUDE.md` rule 4 forbids. A second, smaller helper - `Core_SafeFileName` - is
added alongside it for the same reason one wave early: both new export actions write files to
disk, and W5's backup package will need the identical "make this a legal Windows file name"
logic. Flagged and added now rather than duplicated in W5's diff.

**Tech stack:** VBA (Excel/Windows, one concatenated `.bas`), PowerShell for the build, Node
for the syntax check. No test framework - see *Verification*.

**Runs after phase 3 (W2 - Validate & Reconcile)** and assumes phases 1 and 2
(`2026-07-28-core-refuse-pass-plan.md`, `2026-07-28-core-helper-promotion-plan.md`) are
already on `main`: this plan is written directly against `Core_Refuse`, `Core_IsEngineSheet`,
`Core_AskRange`, `Core_AskHasHeader`, `Core_NarrowToUsed`, `Core_IsAlpha`, `Core_IsBlank`,
`Core_DestUsable`, `Core_ResolveFolder`, and `Core_FolderExists`, none of which exist in the
tree at the time this plan was written. Sequencing: `2026-07-28-track2-completion-plan-of-record.md`
Task table, row 4. **Starting state assumed by this plan:** repo `VERSION` `0.13.0`, contract
`v2.5.0`, `CoreCandidates()` already listing `ImportExport`, `Import`, `Clean`, and whatever
W2 named its package. If the tree does not match that when this plan is picked up, W2 has not
landed (or landed differently) - reconcile before starting, do not guess.

---

## Why this wave is shaped the way it is

Five of the eight actions touch no stored value at all (`format.*`, by the contract's own
domain definition, §7: "Presentation only ... Changes no value"). The other three write
outside the workbook - a value overwrite in place (`export.freeze_values`), or a new file on
disk (`export.pdf`, `export.per_sheet`). Both kinds of change share one property this repo
treats as load-bearing: neither is undoable by the normal means a spreadsheet user reaches
for. Running any VBA clears Excel's own Undo stack, so "changes no value" does not mean
"harmless" - it means the harm, if any, is to *formatting and layout*, not to *data*. Every
action below states, explicitly, which side of that line it is on and what its safety story
is; none is left to be discovered by running it.

---

## Decisions taken in this plan

| # | Question | Decision |
|---|---|---|
| 1 | Does `importexport.export` keep its own copy of the flatten logic, or call the new engine? | **It calls `Core_FreezeValues` (Task 2).** The alternative - a second, hand-written copy for `export.freeze_values` - is the cross-package duplicate the standards forbid outright, not a judgement call. `PkgImportExport.bas`'s ~30-line flatten loop (`:357-390`) collapses to a three-line `For Each ws ... Core_FreezeValues ws.UsedRange ... Next`, with the spill/legacy-array/formula behavior moved verbatim, not rewritten. |
| 2 | Does `format.report_title` inserting rows violate "format changes no value"? | **No stored VALUE changes; a row COORDINATE does, and that is presentation, not data.** Every cell keeps its own value; what moves is which row number holds it. Excel's own `Rows.Insert` auto-adjusts formulas, named ranges and print areas on that sheet - this is native host behavior the action relies on, not something it re-implements, and is logged as a run-time-only claim (Task 14). The action still asks before inserting (Decision 3), because a shift is a real, visible change even though no value is reinterpreted. |
| 3 | Which of the five `format.*` actions need a preview-and-confirm, and which can run silently? | **Split by what gets overwritten.** `format.style_headers` and `format.print_layout` overwrite only what the action itself is unambiguously for (the header row's own styling; this sheet's own print setup) and run without a confirm, matching `clean.standardize_formats`'s no-confirm precedent for pure formatting. `format.report_title` (shifts every row below the insert point), `format.table_style` (replaces existing fill/border formatting across the whole picked range) and `format.conditional_highlight` (clears existing conditional formatting first) each show what will happen and ask Yes/No, default No, before doing it - the same bar `PkgClean_RemoveBlanks` set for a destructive, non-obviously-reversible change. |
| 4 | Are `export.pdf` / `export.per_sheet` allowed to overwrite an existing file? | **No - refuse, don't confirm.** This extends `Core_DestUsable`'s established "refuse rather than overwrite" philosophy from sheets to files: nothing in this toolkit destroys something it did not create in the same run. `export.per_sheet` checks **every** target file before writing **any** of them and refuses the whole batch on the first collision, so a batch export is all-or-nothing, never half-overwritten. |
| 5 | Do the new export actions read `k_export_base` / `k_export_offset` / `k_export_tabs` from `PkgImportExport`'s manifest? | **No - new keys (`k_export_folder`, `k_export_date_offset`, `k_export_pdf_base`, `k_export_persheet_format`), never those three.** Two independent reasons, not one: (a) the brief forbids reusing or shadowing another package's keys outright, and (b) reading `Core.bas:471` shows why that rule is not just etiquette. `Core_Cfg` reads the config-sheet VALUE globally (`ReadScalar` is a flat key lookup, no package scoping) but falls back to the DECLARED DEFAULT only via `FieldDefaultReg(ctx("reg"), key)` - and `ctx("reg")` is the CALLING package's own manifest (`FindActionOwner` resolves it by the action id, `NewCtx` stores it once). A `PkgExport` action reading `k_export_base` would get the right value once a user has typed one into Setup, and silently get `""` - never `PkgImportExport`'s declared `"My Report"` default - the moment the cell is blank. That is a landmine that only misfires on an unconfigured deployment, which is exactly the case a first run needs to behave correctly. New keys under the new package's own manifest sidestep it entirely. |
| 6 | Does `export.per_sheet` skip only engine sheets, or more? | **Engine sheets (required) and non-visible sheets (this wave's addition).** A hidden settings/lookup tab is not itself a deliverable, and exporting it by name would put internal thresholds straight into a file meant to leave the building. Stated as a deliberate widening of the brief's minimum, not an oversight. |
| 7 | `format.conditional_highlight` needs a per-cell rule (column, condition, comparison value, color) - four pieces of data. Does its config table get a fourth column? | **No - `table` fields are capped at three columns.** `Core_CfgTable(ctx, key, a, b, c)` only ever reads three (contract §3: "up to 3 columns are read (C, D, E)"); a fourth was never wired anywhere in Core. Condition and its comparison value are folded into one free-text cell (`"greater_than 1000"`, `"contains VIP"`, `"blank"`), parsed on the first space. Documented with worked examples in the field's `help` text, matching the precedent `k_clean_formats` already set for a field whose shape needs explaining. |
| 8 | Does column resolution (`"C"` or a header name, matched against a picked range) get a third copy in `PkgFormat`, or move to Core? | **A third copy, `PkgFormat_ResolveCol`, flagged rather than silently written.** The 2026-07-28 helper-promotion plan explicitly declined to promote `PkgClean_ResolveFormatCol`, reasoning "each is one package's own job" - a call made before `PkgFormat` existed as a second consumer of the identical logic. Re-litigating that phase's already-committed scope is out of this wave's reach, so this plan does the only honest thing available: duplicate, and say so here in writing. **Escalate-to-Core note, for the next wave that touches this:** `PkgClean_ResolveFormatCol` and `PkgFormat_ResolveCol` are now byte-for-byte the same algorithm in two packages. If a third consumer appears (a plausible W4 `enrich.*` candidate), promote both to a single `Core_ResolveColumn` in that wave's own diff, per the standing rule - do not let a fourth copy get written first. |
| 9 | Does `format.report_title`'s title text live on Setup as a stored field? | **No - prompted at run time via a plain `Application.InputBox`, never stored.** The contract already establishes this shape for "range" (§3: "range is intentionally NOT a kind ... range-scoped actions prompt for their range at run time"). A report's title text is the same kind of per-run fact, not a deployment fact - the same title is not correct on every run the way a house number format is. No new field, no new key, no contract kind needed. |
| 10 | Do `format.style_headers`, `format.table_style` and `format.print_layout` get a config field for their color/style/margin choices? | **No - each is a fixed, deterministic house preset, zero config.** The catalog frames all three as "one-click" macros; a picker for colors or margins is a feature nobody asked for yet (YAGNI), and it would be new Setup surface for three actions that are otherwise this wave's simplest. Revisit only if a real deployment asks for a second look. |
| 11 | Owners | Matches the helper-promotion plan's split: **Core.bas and `spec/contract.md` changes are Opus**, written directly - they are exactly the trigger CLAUDE.md names (Core / contract). **Package construction (`PkgFormat.bas`, `PkgExport.bas`) and the `PkgImportExport.bas` refactor are delegated to codex (`sonnet`)** - every judgement call is already resolved in this plan and every step below contains the actual VBA; what remains is typing it in and verifying, the same shape the helper-promotion plan used for its Task 3-5. |

---

## Global constraints

Copied from the plan-of-record §4, binding on this wave same as every other:

- **`SCHEMA_VER` stays `"v2.0"`.**
- **Action ids are frozen**, two-segment `<domain>.<operation>`, lowercase ASCII and
  underscores, each segment starting with a letter. Never camelCase, no alias table.
- **Shared helpers live ONLY in Core. Cross-package calls are forbidden.**
- **Pure ASCII in every `.bas`**; `Option Explicit`; max 25 line continuations per statement;
  no module-level declaration after the first procedure (a package's `Const`/`Dim` must sit
  above its own first procedure); no VBA reserved word as an identifier.
- **`Public Sub Pkg<Name>_<Verb>(ctx As Object)`** for every action entry point - one `Object`
  parameter, nothing else. `build.ps1` fails the build otherwise.
- **Every package private is prefixed `Pkg<Name>_`.** Phase 2's build guard enforces this on
  every build; a violation here fails the build, not just a review.
- **TDD is off.** Do not propose a test suite. Verification is `pwsh build/build.ps1`,
  `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`, and deliberate injections.
- **`dist/` is gitignored** - never `git add -f` the bundle. `.vscode/` stays untracked.
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.
- **No per-wave manual gate document** (plan-of-record §3.5). Run-time-only claims are
  appended to `docs/designs/2026-07-28-deferred-validation-log.md` (Task 14), not written up
  as a checklist for the owner.

---

## Verification, since TDD is off

Every task verifies with:

1. `pwsh build/build.ps1` - exit 0, `Core + N package(s)` (N grows by one, then two, across
   this wave), balanced `Sub`/`Function` counts, no duplicate-procedure or unprefixed-private
   failures.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` - no diagnostics.
3. **A grep proving every new action id and field key exists exactly once** (Task 13).
4. **A deliberate injection** for the one thing genuinely meant to fail at build time (an
   unprefixed private in the new packages - Task 13 proves the phase-2 guard covers new
   packages too, not just the three it shipped with).
5. Everything that can only be observed by running the workbook is logged to
   `docs/designs/2026-07-28-deferred-validation-log.md` (Task 14), not gated here.

---

## File structure

| File | Change |
|---|---|
| `src/core/Core.bas` | **+`Core_FreezeValues`**, **+`Core_SafeFileName`**; `CoreCandidates()` gains `c.Add "Format"` and `c.Add "Export"`. |
| `src/packages/PkgImportExport.bas` | `PkgImportExport_Export`'s flatten loop (`:357-390`) collapses to a call to `Core_FreezeValues`; now-unused local `fRange`/`ar`/`cel`/`sp`/`sv` removed. No behavior change. |
| `src/packages/PkgFormat.bas` | **New.** Manifest + 5 actions + 2 privates (`PkgFormat_ResolveCol`, `PkgFormat_ApplyRule`, `PkgFormat_ColorFromName`). |
| `src/packages/PkgExport.bas` | **New.** Manifest + 3 actions + 1 private (`PkgExport_OneSheetWorkbook`). |
| `spec/contract.md` | v2.6.0. §8 gains `Package format` and `Package export` subsections; §9 gains two appendix bullets; §10 gains the version entry. |
| `docs/designs/2026-07-28-deferred-validation-log.md` | + this wave's run-time-only claims (append, do not overwrite). |
| `CHANGELOG.md`, `VERSION`, `docs/code-map.md` | v0.14.0. |

---

### Task 1: Contract v2.6.0

**Owner: Opus, written directly.**

**Files:** Modify `spec/contract.md` - header, §8, §9, §10.

**Interfaces:** Produces the normative text Tasks 3-12 implement. Consumes nothing.

- [ ] **Step 1: Bump the version header**

Line 3: `**Contract version: 2.5.0**` -> `**Contract version: 2.6.0**` (same date format as
the surrounding entries). If line 3 does not currently read `2.5.0`, stop - W2 has not landed
as this plan assumes; reconcile before continuing rather than bumping from whatever is there.

- [ ] **Step 2: Verify §7 already lists both domains `active`** (it does - do not add rows)

```bash
grep -n '| `format` | active' spec/contract.md
grep -n '| `export` | active' spec/contract.md
```

Expected: one match each. This is a verification step, not an edit - §7's table already
carries both domains (v2.3.0), and the marker-comment rules in §7 forbid touching the table's
shape for any reason other than adding a genuinely new domain.

- [ ] **Step 3: Add the `format` package subsection to §8**

Insert as a new subsection, in domain order, after the `clean` subsection and before whatever
W2 added for `validate`:

````markdown
### Package `format` - "Format & Present" (v1.0.0, contract v2.6.0)

Presentation only (§7): every action changes how a report looks and never a stored value.

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_format_highlight_rules` | table | simple | **none** | Rules for `format.conditional_highlight`. Three columns: **Column** (a column letter, or a header name matched against the first row of the picked range - same resolution rule as `k_clean_formats`) - **Condition** (`greater_than <value>`, `less_than <value>`, `equal_to <value>`, `contains <value>`, or `blank`, one operator plus its comparison value in one cell) - **Fill color** (one of `red`, `yellow`, `green`, `orange`, `blue`, `gray`). A row whose Column resolves to nothing inside the picked range MUST be reported, not silently skipped. Rules are evaluated top to bottom; a cell matching more than one rule takes the LAST match. `usedBy` `format.conditional_highlight`. |

`k_format_highlight_rules` header: `Column (letter or header)|Condition|Fill color`.

**This field ships with NO default rows, for the same reason `k_clean_formats` does not
(§8, package `clean`).** §5 resolves an empty stored table to the field's declared defaults,
so demo rows would make `format.conditional_highlight`'s required "nothing configured"
refusal unreachable, and would risk coloring a first-time user's own data by a rule they
never wrote.

**Actions**

| id | caption | semantics |
|---|---|---|
| `format.conditional_highlight` | Highlight by Rule | Prompt for a range, then whether its first row is a header. Read `k_format_highlight_rules`; if empty, refuse naming where to configure it. Otherwise, before changing anything, show the range and the rule count and ask to proceed. On yes: clear any existing native conditional formatting in the picked range, then for each rule, resolve its Column against the range and set a static fill color (§ above - not a live conditional-formatting rule) on every data-row cell in that column matching the rule's condition. A rule whose Column resolves to nothing MUST be reported by name, not silently dropped. Changes cell fill only; no value changes. |
| `format.style_headers` | Style Header Row | Prompt for a range. Bold and fill the range's first row, then autofit every column in the range. Overwrites any formatting already on that row. No confirmation - this is unambiguously what the action is for. |
| `format.report_title` | Add Report Title | Refuse if the active sheet is engine-owned. Prompt for title text (run-time input, not stored config - §3's "range is not a kind" precedent extended to free text); refuse if left blank. Show the row count that will be inserted and ask to proceed, noting that formulas and defined names on that sheet follow the shift automatically (native host behavior, §9). On yes: insert that many rows at row 1, merge and center the title text across the sheet's used width, and style it. No stored value changes - every cell keeps its own value; only its row position moves. |
| `format.table_style` | Apply House Table Style | Prompt for a range. Show what will happen (existing fill/border formatting in the range is replaced with a fixed banded-row + thin-border house style) and ask to proceed. On yes: apply it. No value changes. |
| `format.print_layout` | Prep for Print | Refuse if the active sheet is engine-owned. Set that sheet to landscape orientation, fit-to-one-page-wide, a fixed narrow-margin preset, and a repeating first print row. Fixed preset, no configuration. No confirmation - touches print setup only, no cell. |
````

- [ ] **Step 4: Add the `export` package subsection to §8**

Insert immediately after the `format` subsection:

````markdown
### Package `export` - "Export" (v1.0.0, contract v2.6.0)

A second package in the `export` domain alongside the grandfathered `importexport.export`
(§7). Declares its own fields under new keys; it does **not** read `k_export_base`,
`k_export_offset` or `k_export_tabs` from `importexport`'s manifest - those are a different
package's declared defaults, and the reference implementation's config-read fallback is
scoped per owning package (see this wave's plan, Decision 5), so borrowing another package's
key is not just a naming collision but a live default-resolution bug waiting for an
unconfigured deployment.

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_export_folder` | text | simple | *(blank)* | Folder exported files are written to. **Blank means the deployment's data location** (§9), same convention as `k_import_folder`. `usedBy` `export.pdf, export.per_sheet`. |
| `k_export_date_offset` | int | advanced | `-1` | Days to shift the date placed in front of an exported file's name. Same semantics as `importexport`'s `k_export_offset`, under its own key (§ above). `usedBy` `export.pdf, export.per_sheet`. |
| `k_export_pdf_base` | text | simple | `Report` | Base name for `export.pdf`'s single dated PDF -> `<date> <base>.pdf`. `usedBy` `export.pdf`. |
| `k_export_persheet_format` | text | simple | `pdf` | File type `export.per_sheet` writes: `pdf` or `xlsx`. Any other stored value MUST be refused at run time, naming the two accepted tokens. `usedBy` `export.per_sheet`. |

**Actions**

| id | caption | semantics |
|---|---|---|
| `export.freeze_values` | Freeze Formulas to Values | Prompt for a range. Count its formula cells (including dynamic-array spill anchors); refuse if none. Show the count and ask to proceed, stating this cannot be undone. On yes: convert every formula in the range - including spills and legacy CSE arrays - to a static value, using the same flattening behavior `importexport.export` already applies per sheet (§ Core_FreezeValues). Promoted from being embedded inside `importexport.export` to a standalone, range-scoped action; `importexport.export`'s own behavior is unchanged. |
| `export.pdf` | Export to PDF | Resolve the export folder; refuse if the workbook has never been saved and no folder is configured, or if the configured folder does not exist. Ask whether to export the whole workbook, one sheet, or a range (workbook / sheet / range); an unrecognised answer is refused, naming the three accepted tokens. Build the dated output path (`k_export_date_offset` + `k_export_pdf_base`); refuse rather than overwrite if a file already exists at that path. For "sheet", refuse if the named sheet does not exist or is engine-owned. For "range", the picked range temporarily becomes the target sheet's print area for the export and is restored to its prior value immediately after, whether the export succeeds or fails. Writes one PDF. |
| `export.per_sheet` | Export Each Tab Separately | Validate `k_export_persheet_format` is `pdf` or `xlsx`; refuse otherwise. Resolve and validate the export folder as `export.pdf` does. Build the list of every VISIBLE, non-engine-owned sheet (§ Decision 6 - hidden sheets are excluded too, on top of the required engine-sheet exclusion); refuse if none. Check every target file path for an existing collision BEFORE writing any file; refuse the WHOLE batch, naming every colliding path, if even one exists - this action never partially overwrites. Otherwise show the full file list and ask to proceed. On yes: write one file per sheet, named `<date> <sheet name>.<ext>`. The `xlsx` path flattens every formula in a full workbook copy (via `Core_FreezeValues`, same technique `importexport.export` already relies on) before isolating each sheet, so a sheet with a cross-sheet formula does not export with a broken reference. |
````

- [ ] **Step 5: Add two appendix bullets to §9**

Insert after the existing "Export snapshot" / "Dated filename" bullets:

````markdown
- **Range-scoped PDF export (`export.pdf`, range mode)** - a temporary print-area override on
  the target sheet, restored immediately after export. Neutral: render exactly the picked
  cells to the platform's page-description format, without permanently changing the sheet's
  own print setup.
- **Per-tab explode (`export.per_sheet`)** - one output file per visible, non-engine sheet,
  named `<date+offset> <sheet name>.<ext>`. Neutral: the platform's nearest equivalent to "one
  file per tab, skip the toolkit's own surfaces."
````

- [ ] **Step 6: Add the §10 entry**

Insert as the first versioned bullet in §10 (immediately above whatever W2's v2.5.0 entry
added):

````markdown
- **v2.6.0 (additive):** two new packages, `format` (one field, five actions) and `export`
  (four fields, three actions) - §8. `export.freeze_values` promotes logic that already
  shipped embedded inside `importexport.export`; that action's own id, fields and behavior
  are unchanged. **Additive by the two-axes rule:** no existing `key`, action `id`, `kind` or
  encoding is touched or reinterpreted; every new field and action is new surface on an
  already-`active` domain (§7, v2.3.0). `SCHEMA_VER` does **not** move - it stays `"v2.0"`.
  **One judgement call worth stating:** `format.report_title` inserts rows on the sheet it
  targets, which moves every value below the insert point to a different cell address. This
  is classed as presentation, not a value change, because no cell's stored VALUE is
  reinterpreted - only its row coordinate - and because the domain's own definition (§7)
  already anticipates "print layout" as changing no value; a row shift is the same kind of
  layout fact. A platform whose row-insert does not auto-adjust formulas and defined names
  the way Excel's does MUST perform the equivalent adjustment itself, or refuse the action -
  silently breaking a reference is not a conforming realization of this semantics.
````

- [ ] **Step 7: Verify**

```bash
grep -n 'Contract version: 2.6.0' spec/contract.md
grep -c '^### Package `format`' spec/contract.md
grep -c '^### Package `export`' spec/contract.md
```

Expected: one match on the version line, `1` for each package-header count. Full build gate
verification happens once source exists (Task 13).

---

### Task 2: Core gains `Core_FreezeValues` and `Core_SafeFileName`; `importexport.export` is refactored onto it

**Owner: Opus, written directly.** Touches Core and moves logic out of a shipped action's
body - exactly the "Core / migration" trigger.

**Files:** Modify `src/core/Core.bas` (new section beside the sheet/folder helpers, ~line
442 per the helper-promotion plan's Task 2 layout); modify `src/packages/PkgImportExport.bas`
(`PkgImportExport_Export`, `:341-390`).

**Interfaces:**
- Consumes: nothing new (`Range.SpecialCells`, `Range.HasSpill`/`HasArray` are native Excel).
- Produces: `Core_FreezeValues(rng As Range) As Long`, `Core_SafeFileName(raw As String) As String`.
  Tasks 10 and 12 (`PkgExport`) consume both; Task 2's own Step 3 consumes `Core_FreezeValues`.

- [ ] **Step 1: Add `Core_FreezeValues`**

```vb
'---- export engine: freeze formulas (incl. spills) to values ----
' Flatten dynamic-array SPILLS and legacy CSE ARRAYS in rng to static values,
' then freeze every ordinary formula left. Two passes, in this order: a
' spill's anchor cell is itself one of the "formula" cells the second pass
' would try to freeze while it is still spilling, so flattening the spill
' first turns it into an ordinary value block the second pass can walk past
' safely.
'
' Moved here UNCHANGED from importexport.export (2026-07-28 W3 plan) so
' export.freeze_values can call the exact engine that action already relies
' on, instead of a second, untested copy of the same 30 lines - the
' cross-package duplicate CLAUDE.md rule 4 forbids.
'
' PRECONDITION: Application.Calculation must be xlCalculationAutomatic when
' this runs - it reads live formula results via .Value. This was written and
' tested that way inside importexport.export; Core_SpeedOn sets Manual, so a
' caller inside that wrapper must restore Automatic first, exactly as
' importexport.export always has. This function does not manage Calculation
' itself, so two callers running back to back never fight over restoring it.
'
' Returns the number of formula cells it found (before flattening), so a
' caller can report "N formula(s) converted" without a second pass of its own.
Public Function Core_FreezeValues(rng As Range) As Long
    Dim fRange As Range, ar As Range, cel As Range, sp As Range, sv As Variant
    If rng Is Nothing Then Exit Function
    On Error Resume Next
    Set fRange = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If fRange Is Nothing Then Exit Function
    Core_FreezeValues = fRange.Cells.Count

    For Each cel In fRange.Cells
        If cel.HasSpill Then
            Set sp = cel.SpillingToRange
            sv = sp.Value
            cel.ClearContents
            sp.Value = sv
        ElseIf cel.HasArray Then
            Set sp = cel.CurrentArray
            sv = sp.Value
            sp.ClearContents
            sp.Value = sv
        End If
    Next cel

    Set fRange = Nothing
    On Error Resume Next
    Set fRange = rng.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not fRange Is Nothing Then
        For Each ar In fRange.Areas
            On Error Resume Next
            ar.Value = ar.Value
            On Error GoTo 0
        Next ar
    End If
End Function
```

- [ ] **Step 2: Add `Core_SafeFileName`**

```vb
' Any string reduced to something Windows will accept as a FILE NAME - not a
' sheet name. Core_SafeSheetName's forbidden-character list is narrower than
' the filesystem's: Windows additionally forbids  " < > |  on top of Excel's
' own  : \ / ? * [ ]  - so a sheet name that is perfectly legal in Excel (one
' containing a quote, say) can still be an illegal file name. That is why
' this does not simply call Core_SafeSheetName and stop.
'
' Added in the W3 plan (2026-07-28) alongside export.pdf and
' export.per_sheet, and flagged there for W5: backup.timestamped_copy will
' need the identical logic. Promote-on-second-use already happened here, by
' writing it once in Core rather than in export.pdf's own package.
'
' Does NOT make the name unique - a caller writing to disk still checks for an
' existing file itself (Core_DestUsable's sheet precedent: refuse rather than
' resolve a collision by guessing).
'
' maxLen caps the result; 0 means no cap. It is a parameter rather than a fixed
' number because the right cap depends on what the caller is building. A whole
' path is bounded by MAX_PATH, but a caller assembling
' "<folder><date> <label> <timestamp>.xlsx" needs to bound its own middle
' segment, and only it knows how much room the rest takes. W5's
' backup.snapshot passes 40 for exactly this reason.
Public Function Core_SafeFileName(ByVal raw As String, _
                                  Optional ByVal maxLen As Long = 0) As String
    Dim nm As String, bad As String, i As Long
    nm = Trim(raw)
    bad = Core_BadSheetChars() & Chr(34) & "<>|"
    For i = 1 To Len(bad)
        nm = Replace(nm, Mid$(bad, i, 1), "-")
    Next i
    If maxLen > 0 Then
        If Len(nm) > maxLen Then nm = Left$(nm, maxLen)
    End If
    Core_SafeFileName = Trim(nm)
End Function
```

- [ ] **Step 3: Refactor `PkgImportExport_Export` onto `Core_FreezeValues`**

In `src/packages/PkgImportExport.bas`, `PkgImportExport_Export` (`:308-448`), replace the two
`For Each ws In newWb.Worksheets ... SpecialCells ... Next ws` blocks (`:359-390`, the flatten
loop) with:

```vb
    ' Flatten dynamic-array SPILLS, legacy CSE ARRAYS, and ordinary formulas
    ' to static values - Core_FreezeValues, shared with export.freeze_values
    ' (2026-07-28 W3 plan). Behavior is unchanged; only the code moved.
    For Each ws In newWb.Worksheets
        Core_FreezeValues ws.UsedRange
    Next ws
```

Remove the now-unused local declarations `fRange`, `ar`, `cel`, `sp`, `sv` from
`PkgImportExport_Export`'s `Dim` block at the top of the procedure - they were used only by
the deleted loop. Leave every other line of the sub (the sheet copy, the delete-except-kept
loop, the link-breaking, the `SaveAs`) untouched.

- [ ] **Step 4: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 3 package(s)` (still three - `PkgFormat`/`PkgExport` do not exist
yet), no diagnostics. `Function` count up by 2 from Core, `PkgImportExport.bas`'s own line
count drops by roughly 25.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas src/packages/PkgImportExport.bas
git commit -m "feat(core): Core_FreezeValues and Core_SafeFileName; importexport.export now calls the shared engine"
```

---

### Task 3: `PkgFormat` manifest and field

**Owner: delegate to codex (`sonnet`).** Every field and action is fully specified in Task 1;
this types the manifest that declares them.

**Files:** New `src/packages/PkgFormat.bas`.

**Interfaces:** Consumes `Core_DeclarePackage`, `Core_AddSection`, `Core_AddTable`,
`Core_AddAction`. Produces the `format` package registration Tasks 4-8's actions attach to.

- [ ] **Step 1: Create the file**

```vb
Attribute VB_Name = "PkgFormat"
Option Explicit

'================================================================
' PkgFormat - presentation only. Every action here changes how a
' report LOOKS - fonts, fills, borders, row layout, print setup -
' and never a stored value (contract section 7: the format domain
' "changes no value"). Two actions still touch the sheet beyond a
' cell's own formatting: report_title inserts rows, and table_style
' / conditional_highlight overwrite existing formatting - each says
' so and asks before it runs.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST - Core reads this to learn the package's fields + actions.
'================================================================
Public Sub PkgFormat_Manifest(reg As Object)
    Core_DeclarePackage reg, "format", "Format & Present", "1.0.0"

    Core_AddSection reg, "CONDITIONAL HIGHLIGHT"
    Core_AddTable reg, "k_format_highlight_rules", "Highlight rules", _
        "Column (letter or header)|Condition|Fill color", "simple", _
        "One row per rule. Column: a letter like C, or the column's header text (a " & _
        "one-to-three-letter entry is always read as a letter). Condition: greater_than " & _
        "<value>, less_than <value>, equal_to <value>, contains <value>, or blank (no value " & _
        "needed) - e.g. 'greater_than 1000' or 'contains VIP'. Fill color: red, yellow, " & _
        "green, orange, blue, or gray. Rules are checked top to bottom; if more than one " & _
        "matches the same cell, the LAST match wins. Nothing happens until you fill this in.", _
        "format.conditional_highlight"

    Core_AddAction reg, "format.conditional_highlight", "Highlight by Rule", "PkgFormat_ConditionalHighlight", _
        "Pick a range; colors cells whose value matches a rule from the 'Highlight rules' " & _
        "table on Setup. Clears any existing conditional formatting in the range first, and " & _
        "asks before doing either."
    Core_AddAction reg, "format.style_headers", "Style Header Row", "PkgFormat_StyleHeaders", _
        "Pick a range; bolds and fills its first row as a header and autofits every column " & _
        "in the range. Overwrites any formatting already on that row."
    Core_AddAction reg, "format.report_title", "Add Report Title", "PkgFormat_ReportTitle", _
        "Type a title; inserts two rows at the top of the active sheet, merges and centers " & _
        "the title across the sheet's used width, and shifts everything else down. Asks first."
    Core_AddAction reg, "format.table_style", "Apply House Table Style", "PkgFormat_TableStyle", _
        "Pick a range; applies banded row fill and thin borders in the house style. " & _
        "Overwrites any fill or border formatting already in the range. Asks first."
    Core_AddAction reg, "format.print_layout", "Prep for Print", "PkgFormat_PrintLayout", _
        "Sets the active sheet to landscape, fit-to-width, narrow margins, and a repeating " & _
        "header row - the board-pack preset. Changes no cell, only how the sheet prints."
End Sub
```

- [ ] **Step 2: Build**

```bash
pwsh build/build.ps1
```

Expected: **FAIL** - the manifest declares five `entry` points (`PkgFormat_ConditionalHighlight`
etc.) that do not exist as `Public Sub`s yet. This is expected until Task 8 lands; do not
"fix" it by removing actions from the manifest.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgFormat.bas
git commit -m "feat(format): PkgFormat manifest and the highlight-rules field"
```

---

### Task 4: `format.style_headers`

**Owner: delegate to codex (`sonnet`).**

**Files:** Modify `src/packages/PkgFormat.bas` - append below the manifest.

**Interfaces:** Consumes `Core_AskRange`, `Core_NarrowToUsed`, `Core_Refuse`, `Core_RangeRef`,
`Core_SpeedOn`/`Off`, `Core_Done`, `Core_Fail`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Style the header row of a picked range, autofit columns
'================================================================
Public Sub PkgFormat_StyleHeaders(ctx As Object)
    Dim rng As Range, hdr As Range
    Const TITLE As String = "Style Header Row"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                   ' user cancelled
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to style." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    Set hdr = rng.Rows(1)

    Core_SpeedOn
    With hdr
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    rng.Columns.AutoFit
    Core_SpeedOff

    Core_Done ctx, "Styled the header row and autofit " & rng.Columns.Count & " column(s) in " & _
           Core_RangeRef(rng) & "." & vbCrLf & vbCrLf & _
           "Any formatting already on that row was overwritten.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.style_headers", Err
End Sub
```

- [ ] **Step 2: Build**

```bash
pwsh build/build.ps1
```

Expected: still fails (four entry points remain undeclared) - `PkgFormat_StyleHeaders` alone
is now resolvable; confirm via `grep -n "PkgFormat_StyleHeaders" dist/MacrosBundle.bas` that
it appears exactly twice (manifest reference + declaration).

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgFormat.bas
git commit -m "feat(format): format.style_headers"
```

---

### Task 5: `format.report_title`

**Owner: delegate to codex (`sonnet`).** The row-insertion safety story is already resolved
by Decisions 2 and 3 above - this task types the code, it does not re-decide anything.

**Files:** Modify `src/packages/PkgFormat.bas`.

**Interfaces:** Consumes `Core_IsEngineSheet`, `Core_Refuse`, `Core_SpeedOn`/`Off`,
`Core_Done`, `Core_Fail`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Insert a formatted title block above a report
'================================================================
Public Sub PkgFormat_ReportTitle(ctx As Object)
    Dim ws As Worksheet, titleText As Variant, lastCol As Long
    Const TITLE As String = "Add Report Title"
    Const ROWS_INSERTED As Long = 2
    On Error GoTo Fail

    Set ws = ActiveSheet
    If Core_IsEngineSheet(ws.name) Then
        Core_Refuse ctx, "'" & ws.name & "' belongs to the macro toolkit itself, not to a " & _
               "report." & vbCrLf & vbCrLf & "Switch to the report tab and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    titleText = Application.InputBox( _
        Prompt:="Title text for the top of '" & ws.name & "':", title:=TITLE, Type:=2)
    If VarType(titleText) = vbBoolean Then Exit Sub          ' user cancelled
    If Len(Trim(CStr(titleText))) = 0 Then
        Core_Refuse ctx, "No title text was entered." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    lastCol = ws.UsedRange.Columns.Count
    If lastCol < 1 Then lastCol = 1

    If MsgBox("This inserts " & ROWS_INSERTED & " row(s) at the top of '" & ws.name & _
              "' and shifts everything currently there down by " & ROWS_INSERTED & " row(s)." & _
              vbCrLf & vbCrLf & _
              "Formulas and defined names on this sheet follow the shift automatically - " & _
              "that is Excel's own row-insert behavior, not something this macro adds." & _
              vbCrLf & vbCrLf & "Insert the title block?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    ws.Rows("1:" & ROWS_INSERTED).Insert Shift:=xlShiftDown
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Merge
        .Value = CStr(titleText)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Rows(1).RowHeight = 24
    Core_SpeedOff

    Core_Done ctx, "Inserted a title block on '" & ws.name & "'." & vbCrLf & vbCrLf & _
           "Everything that was on this sheet moved down " & ROWS_INSERTED & " row(s); no " & _
           "value changed, only its row position.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.report_title", Err
End Sub
```

- [ ] **Step 2: Build**

```bash
pwsh build/build.ps1
```

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgFormat.bas
git commit -m "feat(format): format.report_title"
```

---

### Task 6: `format.table_style`

**Owner: delegate to codex (`sonnet`).**

**Files:** Modify `src/packages/PkgFormat.bas`.

**Interfaces:** Consumes `Core_AskRange`, `Core_NarrowToUsed`, `Core_Refuse`, `Core_RangeRef`,
`Core_SpeedOn`/`Off`, `Core_Done`, `Core_Fail`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Apply the house table style (banded rows + borders)
'================================================================
Public Sub PkgFormat_TableStyle(ctx As Object)
    Dim rng As Range, r As Long, bandColor As Long
    Const TITLE As String = "Apply House Table Style"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                   ' user cancelled
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to style." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If MsgBox("This replaces any fill and border formatting already inside " & _
              Core_RangeRef(rng) & " with the house table style (banded rows + thin " & _
              "borders). It does not change any value." & vbCrLf & vbCrLf & "Apply it?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    bandColor = RGB(237, 241, 245)
    rng.Interior.ColorIndex = xlColorIndexNone
    For r = 1 To rng.Rows.Count
        If r Mod 2 = 0 Then rng.Rows(r).Interior.Color = bandColor
    Next r
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With
    rng.Rows(1).Font.Bold = True
    Core_SpeedOff

    Core_Done ctx, "Applied the house table style to " & Core_RangeRef(rng) & "." & vbCrLf & _
           vbCrLf & "Any fill or border formatting already there was replaced. No value " & _
           "changed.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.table_style", Err
End Sub
```

- [ ] **Step 2: Build**

```bash
pwsh build/build.ps1
```

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgFormat.bas
git commit -m "feat(format): format.table_style"
```

---

### Task 7: `format.conditional_highlight`

**Owner: delegate to codex (`sonnet`).** Every rule, safety decision and the three-column
table shape are already settled (Decisions 3, 7, 8); this types the code.

**Files:** Modify `src/packages/PkgFormat.bas`.

**Interfaces:** Consumes `Core_CfgTable`, `Core_AskRange`, `Core_NarrowToUsed`,
`Core_AskHasHeader`, `Core_Refuse`, `Core_RangeRef`, `Core_IsAlpha`, `Core_IsBlank`,
`Core_SpeedOn`/`Off`, `Core_Done`, `Core_Fail`. Produces `PkgFormat_ResolveCol`,
`PkgFormat_ApplyRule`, `PkgFormat_ColorFromName` (private, this package only).

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Color cells that match a configured rule
'================================================================
Public Sub PkgFormat_ConditionalHighlight(ctx As Object)
    Dim cols() As String, conds() As String, colors() As String
    Dim rng As Range, hasHeader As Boolean, startRow As Long
    Dim i As Long, tgt As Range, matched As Long, missed As String
    Const TITLE As String = "Highlight by Rule"
    On Error GoTo Fail

    If Not Core_CfgTable(ctx, "k_format_highlight_rules", cols, conds, colors) Then
        Core_Refuse ctx, "No highlight rules are configured yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Highlight rules' table on the Setup tab, Finalize, and run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                   ' user cancelled
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to check." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub    ' user cancelled

    If MsgBox("This clears any existing conditional formatting inside " & Core_RangeRef(rng) & _
              " and colors cells against " & (UBound(cols) - LBound(cols) + 1) & " rule(s) " & _
              "from Setup." & vbCrLf & vbCrLf & "Apply it?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    startRow = IIf(hasHeader, 2, 1)

    Core_SpeedOn
    rng.FormatConditions.Delete
    For i = startRow To rng.Rows.Count
        rng.Rows(i).Interior.ColorIndex = xlColorIndexNone
    Next i

    For i = LBound(cols) To UBound(cols)
        Set tgt = PkgFormat_ResolveCol(rng, cols(i))
        If tgt Is Nothing Then
            missed = missed & vbCrLf & "   " & cols(i)
        Else
            matched = matched + PkgFormat_ApplyRule(tgt, startRow, conds(i), colors(i))
        End If
    Next i
    Core_SpeedOff

    Core_Done ctx, "Colored " & matched & " cell(s) inside " & Core_RangeRef(rng) & "." & _
           vbCrLf & "Existing conditional formatting in the range was cleared first." & _
           IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These rows matched no column in your " & _
               "selection and were skipped:" & missed, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.conditional_highlight", Err
End Sub
```

- [ ] **Step 2: Add the three privates**

```vb
' Parse ONE "Condition" cell ("greater_than 1000", "contains VIP", "blank")
' and color every matching cell in colRng from startRow down. colRng spans
' the SAME rows as the picked range (PkgFormat_ResolveCol guarantees this),
' so startRow indexes it directly. Returns how many cells it colored.
'
' An unrecognised operator or color colors nothing and returns 0 - the row
' is silently skipped rather than raising, so one typo does not abort the
' whole rule table; the help text spells out the exact five operator tokens
' and six color names for exactly this reason.
Private Function PkgFormat_ApplyRule(colRng As Range, ByVal startRow As Long, _
                                     ByVal condSpec As String, ByVal colorName As String) As Long
    Dim op As String, val As String, sp As Long, i As Long, c As Range, hit As Boolean
    Dim fillColor As Long
    condSpec = Trim(condSpec)
    sp = InStr(condSpec, " ")
    If sp > 0 Then
        op = LCase(Left$(condSpec, sp - 1))
        val = Trim(Mid$(condSpec, sp + 1))
    Else
        op = LCase(condSpec)
        val = ""
    End If
    fillColor = PkgFormat_ColorFromName(colorName)
    If fillColor = -1 Then Exit Function

    For i = startRow To colRng.Rows.Count
        Set c = colRng.Cells(i, 1)
        hit = False
        Select Case op
            Case "blank"
                hit = Core_IsBlank(c.Value)
            Case "equal_to"
                hit = (StrComp(Trim(CStr(c.Value)), val, vbTextCompare) = 0)
            Case "contains"
                hit = (InStr(1, CStr(c.Value), val, vbTextCompare) > 0)
            Case "greater_than"
                hit = IsNumeric(c.Value) And IsNumeric(val) And (CDbl(c.Value) > CDbl(val))
            Case "less_than"
                hit = IsNumeric(c.Value) And IsNumeric(val) And (CDbl(c.Value) < CDbl(val))
        End Select
        If hit Then
            c.Interior.Color = fillColor
            PkgFormat_ApplyRule = PkgFormat_ApplyRule + 1
        End If
    Next i
End Function

' Named colors only - a non-technical user typing a color name is more
' reliable than one typing a hex code or a VBA color constant. Returns -1
' for anything else; the caller reads that as "skip this rule".
Private Function PkgFormat_ColorFromName(ByVal nm As String) As Long
    Select Case LCase(Trim(nm))
        Case "red":    PkgFormat_ColorFromName = RGB(255, 199, 206)
        Case "yellow": PkgFormat_ColorFromName = RGB(255, 235, 156)
        Case "green":  PkgFormat_ColorFromName = RGB(198, 239, 206)
        Case "orange": PkgFormat_ColorFromName = RGB(252, 213, 180)
        Case "blue":   PkgFormat_ColorFromName = RGB(189, 215, 238)
        Case "gray", "grey": PkgFormat_ColorFromName = RGB(217, 217, 217)
        Case Else: PkgFormat_ColorFromName = -1
    End Select
End Function

' Resolve one "Column" cell against the picked range, returned as a range
' spanning the SAME ROWS as rng (one column wide) so a caller can index it
' by row number regardless of where the header row is.
'
' Same letter-or-header logic as PkgClean_ResolveFormatCol - a DELIBERATE
' duplicate, not an oversight. See this wave's plan, Decision 8: the
' helper-promotion pass declined to promote the original with "each is one
' package's own job", a call made before this was a second consumer. Not
' re-litigated here; flagged instead for whichever wave adds a third.
Private Function PkgFormat_ResolveCol(rng As Range, ByVal spec As String) As Range
    Dim s As String, absCol As Long, i As Long, hdr As String
    s = Trim(spec)
    If Len(s) = 0 Then Exit Function

    If Len(s) <= 3 And Core_IsAlpha(s) Then
        On Error Resume Next
        absCol = rng.Parent.Range(s & "1").Column
        On Error GoTo 0
        If absCol > 0 Then
            If absCol >= rng.Column And absCol <= rng.Column + rng.Columns.Count - 1 Then
                Set PkgFormat_ResolveCol = Application.Intersect(rng.Parent.Columns(absCol), rng)
            End If
            Exit Function
        End If
    End If

    For i = 1 To rng.Columns.Count
        hdr = Trim(CStr(rng.Cells(1, i).Value))
        If StrComp(hdr, s, vbTextCompare) = 0 Then
            Set PkgFormat_ResolveCol = rng.Columns(i)
            Exit Function
        End If
    Next i
End Function
```

- [ ] **Step 3: Build**

```bash
pwsh build/build.ps1
```

- [ ] **Step 4: Commit**

```bash
git add src/packages/PkgFormat.bas
git commit -m "feat(format): format.conditional_highlight"
```

---

### Task 8: `format.print_layout`, and `PkgFormat` is complete

**Owner: delegate to codex (`sonnet`).**

**Files:** Modify `src/packages/PkgFormat.bas`.

**Interfaces:** Consumes `Core_IsEngineSheet`, `Core_Refuse`, `Core_SpeedOn`/`Off`,
`Core_Done`, `Core_Fail`. This is the last action `PkgFormat` needs - the file is complete
once this lands.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Board-pack print preset: landscape, fit-to-width, narrow
' margins, repeating header row. Fixed preset, no configuration -
' the whole point is one consistent look every close.
'================================================================
Public Sub PkgFormat_PrintLayout(ctx As Object)
    Dim ws As Worksheet
    Const TITLE As String = "Prep for Print"
    On Error GoTo Fail

    Set ws = ActiveSheet
    If Core_IsEngineSheet(ws.name) Then
        Core_Refuse ctx, "'" & ws.name & "' belongs to the macro toolkit itself, not to a " & _
               "report." & vbCrLf & vbCrLf & "Switch to the report tab and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        On Error Resume Next
        .PrintTitleRows = "$1:$1"
        On Error GoTo 0
    End With
    Core_SpeedOff

    Core_Done ctx, "'" & ws.name & "' is set to landscape, fit-to-width, narrow margins, " & _
           "and a repeating header row (row 1)." & vbCrLf & vbCrLf & _
           "No cell was changed - only how this sheet prints.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.print_layout", Err
End Sub
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: **still fails** - `CoreCandidates()` does not yet probe `PkgFormat_Manifest`
(Task 13), so the build's entry-point resolver cannot find any of these five `entry` values
declared anywhere it checks... actually it CAN: the resolver checks source text, not the
candidate list (the candidate list is a runtime-only concern). Expected here: **exit 0**, no
diagnostics, `Core + 3 package(s)` still (the glob-discovered `PkgFormat.bas` is bundled, but
still un-probed until Task 13 - the build succeeding does not mean the package runs; that is
exactly the two-place-registration trap Task 13 exists to close).

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgFormat.bas
git commit -m "feat(format): format.print_layout - PkgFormat is now complete"
```

---

### Task 9: `PkgExport` manifest and fields

**Owner: delegate to codex (`sonnet`).**

**Files:** New `src/packages/PkgExport.bas`.

**Interfaces:** Consumes `Core_DeclarePackage`, `Core_AddSection`, `Core_AddField`,
`Core_AddAction`. Produces the `export` package registration Tasks 10-12's actions attach to.

- [ ] **Step 1: Create the file**

```vb
Attribute VB_Name = "PkgExport"
Option Explicit

'================================================================
' PkgExport - getting finished work OUT of the workbook: frozen
' values, PDFs, one file per tab. All three write outside the
' workbook (a value overwrite in place, or a new file on disk), so
' every action here either refuses on a real collision or asks
' before doing the irreversible part - matching the safety bar
' PkgClean set for in-workbook deletes.
'
' export.freeze_values shares its flattening engine, Core_FreezeValues,
' with importexport.export (PkgImportExport.bas) - see the 2026-07-28
' W3 plan for why the engine lives in Core rather than being copied.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgExport_Manifest(reg As Object)
    Core_DeclarePackage reg, "export", "Export", "1.0.0"

    Core_AddSection reg, "GENERAL"
    Core_AddField reg, "k_export_folder", "Folder to write exported files to", "text", "", _
        "simple", _
        "Leave blank to use the folder this workbook is saved in - that is the usual " & _
        "answer. Otherwise a full path, e.g. C:\Close\January.", _
        "export.pdf,export.per_sheet"
    Core_AddField reg, "k_export_date_offset", "Date shift for exported file names (days)", "int", "-1", _
        "advanced", _
        "How many days to shift the date put in front of an exported file's name. -1 = " & _
        "yesterday, 0 = today, 1 = tomorrow.", _
        "export.pdf,export.per_sheet"

    Core_AddSection reg, "PDF EXPORT"
    Core_AddField reg, "k_export_pdf_base", "Base name for the exported PDF", "text", "Report", _
        "simple", _
        "The saved PDF will be named like:  2026-01-15 Report.pdf", _
        "export.pdf"

    Core_AddSection reg, "PER-TAB EXPORT"
    Core_AddField reg, "k_export_persheet_format", "File type for per-tab export", "text", "pdf", _
        "simple", _
        "pdf or xlsx. Each visible, non-toolkit sheet is written as its own file named " & _
        "after the sheet.", _
        "export.per_sheet"

    Core_AddAction reg, "export.freeze_values", "Freeze Formulas to Values", "PkgExport_FreezeValues", _
        "Pick a range; converts every formula in it (including dynamic-array spills) into a " & _
        "static value. Shows how many formulas it found and asks before doing it - this " & _
        "cannot be undone with Ctrl+Z."
    Core_AddAction reg, "export.pdf", "Export to PDF", "PkgExport_Pdf", _
        "Save the whole workbook, one sheet, or a range you pick, as a dated PDF. Refuses " & _
        "rather than overwriting a same-named file that is already there."
    Core_AddAction reg, "export.per_sheet", "Export Each Tab Separately", "PkgExport_PerSheet", _
        "Save every visible, non-toolkit sheet as its own dated PDF or .xlsx file. Shows " & _
        "the full file list and asks before writing anything; refuses the whole batch " & _
        "rather than overwriting even one file that is already there."
End Sub
```

- [ ] **Step 2: Build**

```bash
pwsh build/build.ps1
```

Expected: **FAIL** - three undeclared entry points, same as Task 3.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgExport.bas
git commit -m "feat(export): PkgExport manifest and fields"
```

---

### Task 10: `export.freeze_values`

**Owner: delegate to codex (`sonnet`).** The Core-promotion decision is already made (Task
2); this is the standalone action that calls it.

**Files:** Modify `src/packages/PkgExport.bas`.

**Interfaces:** Consumes `Core_AskRange`, `Core_NarrowToUsed`, `Core_Refuse`, `Core_RangeRef`,
`Core_SpeedOn`/`Off`, `Core_FreezeValues`, `Core_Done`, `Core_Fail`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Freeze formulas (incl. spills) in a picked range to values
'================================================================
Public Sub PkgExport_FreezeValues(ctx As Object)
    Dim rng As Range, cnt As Long
    Const TITLE As String = "Freeze Formulas to Values"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                   ' user cancelled
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to freeze." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    cnt = 0
    On Error Resume Next
    cnt = rng.SpecialCells(xlCellTypeFormulas).Cells.Count
    On Error GoTo 0
    If cnt = 0 Then
        Core_Refuse ctx, "There are no formulas inside " & Core_RangeRef(rng) & " to freeze." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If MsgBox("This converts " & cnt & " formula(s) - including any dynamic-array spills - " & _
              "inside " & Core_RangeRef(rng) & " into static values." & vbCrLf & vbCrLf & _
              "This cannot be undone with Ctrl+Z." & vbCrLf & vbCrLf & "Freeze them?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    ' Core_FreezeValues reads live formula results, so it needs Automatic
    ' calculation exactly as importexport.export always has - Core_SpeedOn
    ' sets Manual, so it is put back before the freeze; Core_SpeedOff
    ' restores whatever was there before this action ran.
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    Core_FreezeValues rng
    Core_SpeedOff

    Core_Done ctx, "Converted " & cnt & " formula(s) inside " & Core_RangeRef(rng) & " to " & _
           "static values." & vbCrLf & vbCrLf & "This cannot be undone with Ctrl+Z.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "export.freeze_values", Err
End Sub
```

- [ ] **Step 2: Build**

```bash
pwsh build/build.ps1
```

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgExport.bas
git commit -m "feat(export): export.freeze_values, promoted from importexport.export"
```

---

### Task 11: `export.pdf`

**Owner: delegate to codex (`sonnet`).** The refuse-not-overwrite and print-area-restore
decisions are already made (Decision 4); this types the code.

**Files:** Modify `src/packages/PkgExport.bas`.

**Interfaces:** Consumes `Core_ResolveFolder`, `Core_FolderExists`, `Core_Refuse`,
`Core_SafeFileName`, `Core_Cfg`, `Core_CfgLong`, `Core_SheetExists`, `Core_IsEngineSheet`,
`Core_AskRange`, `Core_NarrowToUsed`, `Core_SpeedOn`/`Off`, `Core_Done`, `Core_Fail`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Export the workbook, one sheet, or a range, to a dated PDF
'================================================================
Public Sub PkgExport_Pdf(ctx As Object)
    Dim scope As Variant, folderPath As String, baseName As String, outPath As String
    Dim ws As Worksheet, rng As Range, savedArea As String, sheetName As Variant
    Dim printAreaChanged As Boolean
    Const TITLE As String = "Export to PDF"
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

    scope = Application.InputBox( _
        Prompt:="Export the whole workbook, one sheet, or a range you pick?" & vbCrLf & _
                vbCrLf & "Type:  workbook,  sheet,  or  range.", _
        title:=TITLE, Default:="sheet", Type:=2)
    If VarType(scope) = vbBoolean Then Exit Sub               ' user cancelled
    scope = LCase(Trim(CStr(scope)))
    If scope <> "workbook" And scope <> "sheet" And scope <> "range" Then
        Core_Refuse ctx, "'" & scope & "' is not one of the accepted answers: workbook, " & _
               "sheet, or range." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    baseName = Core_SafeFileName(Core_Cfg(ctx, "k_export_pdf_base"))
    outPath = folderPath & Format(Date + Core_CfgLong(ctx, "k_export_date_offset"), "yyyy-mm-dd") & _
              " " & baseName & ".pdf"
    If Len(Dir(outPath)) > 0 Then
        Core_Refuse ctx, "A file already exists at:" & vbCrLf & vbCrLf & "   " & outPath & _
               vbCrLf & vbCrLf & "This action will not overwrite it. Rename or delete that " & _
               "file, or change the base name / date offset on Setup, then run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Select Case scope
        Case "workbook"
            Core_SpeedOn
            ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPath, _
                Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=False
            Core_SpeedOff

        Case "sheet"
            sheetName = Application.InputBox( _
                Prompt:="Sheet to export:", title:=TITLE, Default:=ActiveSheet.name, Type:=2)
            If VarType(sheetName) = vbBoolean Then Exit Sub   ' user cancelled
            If Not Core_SheetExists(CStr(sheetName)) Then
                Core_Refuse ctx, "There is no sheet named '" & sheetName & "'." & vbCrLf & _
                       vbCrLf & "Nothing has been changed.", TITLE
                Exit Sub
            End If
            Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
            If Core_IsEngineSheet(ws.name) Then
                Core_Refuse ctx, "'" & ws.name & "' belongs to the macro toolkit itself, " & _
                       "not to a report." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
                Exit Sub
            End If
            Core_SpeedOn
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPath, _
                Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=False
            Core_SpeedOff

        Case "range"
            Set rng = Core_AskRange(TITLE)
            If rng Is Nothing Then Exit Sub                   ' user cancelled
            Set rng = Core_NarrowToUsed(rng)
            If rng Is Nothing Then
                Core_Refuse ctx, "The range you picked is entirely outside the used area " & _
                       "of the sheet, so there is nothing in it to export." & vbCrLf & _
                       vbCrLf & "Nothing has been changed.", TITLE
                Exit Sub
            End If
            Set ws = rng.Parent
            ' The picked range becomes the sheet's Print Area only for the
            ' duration of this export, then is put back exactly as it was -
            ' contract section 6 forbids an engine touching a surface the
            ' user owns beyond what it is explicitly asked to do, and a
            ' stray Print Area left behind would silently change what a
            ' later manual Ctrl+P prints.
            savedArea = ws.PageSetup.PrintArea
            printAreaChanged = True
            Core_SpeedOn
            ws.PageSetup.PrintArea = rng.Address
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPath, _
                Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=False
            ws.PageSetup.PrintArea = savedArea
            printAreaChanged = False
            Core_SpeedOff
    End Select

    Core_Done ctx, "Exported the " & scope & " to:" & vbCrLf & vbCrLf & outPath, TITLE
    Exit Sub
Fail:
    On Error Resume Next
    If printAreaChanged And Not ws Is Nothing Then ws.PageSetup.PrintArea = savedArea
    Core_Fail ctx, "export.pdf", Err
End Sub
```

- [ ] **Step 2: Build**

```bash
pwsh build/build.ps1
```

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgExport.bas
git commit -m "feat(export): export.pdf"
```

---

### Task 12: `export.per_sheet`

**Owner: delegate to codex (`sonnet`).** The all-or-nothing batch-refuse decision (Decision
4) and the engine-sheet-plus-hidden-sheet skip (Decision 6) are already made.

**Files:** Modify `src/packages/PkgExport.bas`.

**Interfaces:** Consumes everything Task 11 does, plus `Core_FreezeValues`, `Core_FileName`.
Produces `PkgExport_OneSheetWorkbook` (private).

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Explode the workbook into one dated file per visible tab
'================================================================
Public Sub PkgExport_PerSheet(ctx As Object)
    Dim fmt As String, ext As String, folderPath As String, dateStr As String
    Dim ws As Worksheet, targets As Collection, paths As Collection
    Dim i As Long, existing As String, listMsg As String, n As Long
    Const TITLE As String = "Export Each Tab Separately"
    On Error GoTo Fail

    fmt = LCase(Trim(Core_Cfg(ctx, "k_export_persheet_format")))
    If fmt <> "pdf" And fmt <> "xlsx" Then
        Core_Refuse ctx, "'" & fmt & "' is not one of the accepted file types for per-tab " & _
               "export: pdf or xlsx." & vbCrLf & vbCrLf & "Fix 'File type for per-tab " & _
               "export' on the Setup tab, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    ext = "." & fmt

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

    ' Every VISIBLE, non-toolkit sheet is a candidate. Hidden sheets are
    ' skipped too, on top of the required engine-sheet skip: a hidden
    ' settings/lookup tab is not itself a deliverable, and exporting it by
    ' name would put internal thresholds straight into a file meant to
    ' leave the building.
    dateStr = Format(Date + Core_CfgLong(ctx, "k_export_date_offset"), "yyyy-mm-dd")
    Set targets = New Collection
    Set paths = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Not Core_IsEngineSheet(ws.name) Then
                targets.Add ws
                paths.Add folderPath & dateStr & " " & Core_SafeFileName(ws.name) & ext
            End If
        End If
    Next ws

    If targets.Count = 0 Then
        Core_Refuse ctx, "There are no visible, non-toolkit sheets to export." & vbCrLf & _
               vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Check EVERY target file before writing ANY of them - a batch action
    ' either runs in full or refuses in full, never half-overwrites.
    For i = 1 To paths.Count
        If Len(Dir(CStr(paths(i)))) > 0 Then existing = existing & vbCrLf & "   " & paths(i)
    Next i
    If Len(existing) > 0 Then
        Core_Refuse ctx, "These file(s) already exist and this action will not overwrite " & _
               "them:" & existing & vbCrLf & vbCrLf & "Rename or delete them, or change the " & _
               "date offset on Setup, then run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    For i = 1 To targets.Count
        listMsg = listMsg & vbCrLf & "   " & Core_FileName(CStr(paths(i)))
    Next i
    If MsgBox("This writes " & targets.Count & " file(s) to:" & vbCrLf & "   " & folderPath & _
              vbCrLf & vbCrLf & listMsg & vbCrLf & vbCrLf & "Continue?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    If fmt = "pdf" Then
        For i = 1 To targets.Count
            targets(i).ExportAsFixedFormat Type:=xlTypePDF, Filename:=paths(i), _
                Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=False
            n = n + 1
        Next i
    Else
        For i = 1 To targets.Count
            PkgExport_OneSheetWorkbook targets(i).name, CStr(paths(i))
            n = n + 1
        Next i
    End If
    Core_SpeedOff

    Core_Done ctx, "Wrote " & n & " file(s) to:" & vbCrLf & "   " & folderPath, TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "export.per_sheet", Err
End Sub
```

- [ ] **Step 2: Add the private engine**

```vb
'================================================================
' Write ONE sheet out as its own .xlsx. Copies the WHOLE workbook first
' and freezes every formula in the copy (Core_FreezeValues) before
' deleting every sheet but the target - exactly the technique
' importexport.export already relies on. A sheet copied out ALONE, with
' its cross-sheet formulas still live, would show #REF! in every cell
' that pointed at a sheet no longer there; freezing first is what
' importexport.export already proved works, so this reuses it rather
' than risking a cleverer, untested shortcut.
'================================================================
Private Sub PkgExport_OneSheetWorkbook(ByVal sheetName As String, ByVal outPath As String)
    Dim newWb As Workbook, ws As Worksheet, i As Long
    Dim savedCalc As XlCalculation, errNum As Long, errDesc As String
    On Error GoTo CleanFail

    savedCalc = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

    ThisWorkbook.Worksheets.Copy
    Set newWb = Application.ActiveWorkbook
    newWb.Application.CalculateFull

    For Each ws In newWb.Worksheets
        Core_FreezeValues ws.UsedRange
    Next ws

    For i = newWb.Worksheets.Count To 1 Step -1
        If StrComp(newWb.Worksheets(i).name, sheetName, vbTextCompare) <> 0 Then
            newWb.Worksheets(i).Visible = xlSheetVisible
            newWb.Worksheets(i).Delete
        End If
    Next i

    newWb.SaveAs Filename:=outPath, FileFormat:=51        ' 51 = .xlsx, no macros
    newWb.Close SaveChanges:=False
    Application.Calculation = savedCalc
    Exit Sub

CleanFail:
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    Application.Calculation = savedCalc
    On Error GoTo 0
    Err.Raise errNum, "PkgExport_OneSheetWorkbook", errDesc
End Sub
```

- [ ] **Step 3: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics, `Core + 3 package(s)` (unchanged - `PkgFormat`/`PkgExport`
are bundled but not yet probed; Task 13 closes that gap).

- [ ] **Step 4: Commit**

```bash
git add src/packages/PkgExport.bas
git commit -m "feat(export): export.per_sheet"
```

---

### Task 13: Two-place registration - `CoreCandidates()` gains `Format` and `Export`

**Owner: delegate to codex (`sonnet`).** One two-line edit, but it is the single easiest way
to ship a dead wave, so it gets its own task and its own verification rather than being folded
silently into Task 2.

**Files:** Modify `src/core/Core.bas`, `CoreCandidates()` (~line 952, exact line shifts by
whatever W2 added).

**Interfaces:** Consumes nothing. Produces the runtime discovery that makes every action in
Tasks 3-12 actually dispatchable - **without this step the build passes, the bundle imports,
and every button silently does not exist.**

- [ ] **Step 1: Add two lines**

By the time this task runs, `CoreCandidates()` already lists `"ImportExport"`, `"Import"`,
`"Clean"`, and whatever W2 (phase 3) added for `validate`. Add exactly two lines, after the
existing entries, preserving their order:

```vb
    c.Add "Format"                  ' probes PkgFormat_Manifest
    c.Add "Export"                  ' probes PkgExport_Manifest
```

- [ ] **Step 2: Prove both packages are now discoverable**

```bash
grep -n 'c.Add "Format"' src/core/Core.bas
grep -n 'c.Add "Export"' src/core/Core.bas
```

Expected: one match each.

- [ ] **Step 3: Build, syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 5 package(s)` (three prior + `PkgFormat` + `PkgExport`), no
diagnostics.

- [ ] **Step 4: A grep proving every new id and key exists exactly once**

```bash
grep -noE '"format\.[a-z_]+"' src/packages/PkgFormat.bas | sort | uniq -c
grep -noE '"export\.[a-z_]+"' src/packages/PkgExport.bas | sort | uniq -c
grep -noE '"k_format_[a-z_]+"' src/packages/PkgFormat.bas | sort -u
grep -noE '"k_export_[a-z_]+"' src/packages/PkgExport.bas | sort -u
grep -n 'k_export_base\|k_export_offset\|k_export_tabs' src/packages/PkgExport.bas
```

Expected: five distinct `format.*` ids, three distinct `export.*` ids (each appearing twice -
manifest declaration + `Public Sub` - grep the `.bas` file, not the bundle, to see the
manifest string only), one `k_format_*` key, four distinct `k_export_*` keys, and **no
output** from the last line - confirming `PkgExport` never references `PkgImportExport`'s
three keys (Decision 5).

- [ ] **Step 5: Deliberate injection - the phase-2 prefix guard covers new packages too**

```bash
printf '\nPrivate Function TotallyUnprefixed() As Boolean\nEnd Function\n' >> src/packages/PkgFormat.bas
pwsh build/build.ps1; echo "exit=$?"
git checkout -- src/packages/PkgFormat.bas
pwsh build/build.ps1; echo "exit=$?"
```

Expected: **exit 1** the first time, naming `PkgFormat.bas` and `TotallyUnprefixed`; **exit 0**
after the revert. Record both in the PR body.

- [ ] **Step 6: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): register Format and Export in CoreCandidates - both packages now run"
```

---

### Task 14: Version, changelog, code map, deferred-validation log, and the PR

**Owner: Opus.**

**Files:** `VERSION`, `CHANGELOG.md`, `docs/code-map.md`,
`docs/designs/2026-07-28-deferred-validation-log.md`.

- [ ] **Step 1: `VERSION` -> `0.14.0`, and a `## [0.14.0]` changelog entry**

House style: state what changed for the user and what breaks for whom, not a feature list.
Lead with the promotion (freeze-values is now its own button, and `importexport.export`'s
behavior did not change), then the five presentation actions, then the safety story each one
carries (which ones ask first, which refuse rather than overwrite a file). Do not omit the
refuse-not-overwrite behavior for `export.pdf`/`export.per_sheet` - it is the one thing an
owner running this against a real folder will hit on a same-day rerun.

- [ ] **Step 2: Refresh the code map**

```bash
node scripts/code-map.mjs
```

- [ ] **Step 3: Append this wave's run-time-only claims to the deferred-validation log**

If `docs/designs/2026-07-28-deferred-validation-log.md` does not exist yet, phase 1 has not
landed - stop and reconcile before continuing (plan-of-record §3.5 assumes it exists by now).
Otherwise, **append** (do not overwrite) a `## W3 - Format & Present, plus Export` section:

```markdown
## W3 - Format & Present, plus Export

- **`ExportAsFixedFormat` produces a real, openable PDF** for all three `export.pdf` scopes
  (workbook / sheet / range) and for `export.per_sheet`'s pdf mode. The build can confirm the
  call compiles; it cannot confirm the file opens or looks right.
- **Excel's row-insert auto-adjusts formulas, defined names, and print areas** on the sheet
  `format.report_title` targets. Assumed, per the contract's normative claim about this
  wave (§10, v2.6.0) - marked per CLAUDE.md 7, never independently exercised in this codebase
  before this wave.
- **`format.conditional_highlight`'s six fill colors are legible on the workbook's own Excel
  theme**, not just against the light-mode swatch they were picked against. Carries the same
  open concern the plan-of-record already flagged for track-1 gate item "Reskin 4" (dark-theme
  legibility) - this wave adds a second, independent surface that concern applies to.
- **`ws.Copy` with no destination makes the copy `Application.ActiveWorkbook`** in a
  single-window session. `PkgExport_OneSheetWorkbook` (export.per_sheet's xlsx path) relies on
  this exactly as `importexport.export` already does; neither has been run in a multi-window
  Excel session.
- **`PageSetup` properties set by `format.print_layout` render as expected in Print Preview** -
  landscape, fit-to-one-page-wide, the margin values chosen, and the repeating header row.
  Page setup has no headless equivalent to check against.
```

- [ ] **Step 4: Final gates**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas
grep -nE '^\s*Private\s+(Sub|Function)\s+' src/packages/*.bas | grep -vE 'Private\s+(Sub|Function)\s+Pkg'
git status --short
```

Expected: exit 0; no diagnostics; `1`; no output from the prefix grep; a tree clean but for the
untracked `.vscode/`.

- [ ] **Step 5: Review the diff, then one PR against `main`**

Not stacked on any prior phase's branch. Branch from `main` after phase 3 (W2) has merged.

**The PR body must state, in one line, the freeze-values promotion decision** (Core engine,
`importexport.export` now calls it, behavior unchanged) - a reviewer reading two new
2000+-line-equivalent packages needs the one sentence that says what to check the shared logic
*against*.

---

## Unknowns and risks

- **`format.conditional_highlight` sets static cell colors, not native `FormatConditions`
  rules, by design** ("beyond native CF" per the catalog entry) - but this means a value that
  later changes to no longer match its rule stays colored until the action is re-run. This is
  a real behavioral difference from Excel's native CF (which re-evaluates live) and is stated
  in the action's help text ("Clears any existing conditional formatting... first") but is
  easy to misread as "this behaves like CF." If real usage wants live re-evaluation, that is a
  different action, not a fix to this one.
- **`format.report_title`'s row insertion has not been checked against Excel Tables
  (ListObjects).** Whether an `Insert` above a Table's header row extends the Table, sits
  above it cleanly, or raises is genuinely untested in this codebase and not covered by this
  plan's reasoning about named ranges and formulas - flagged as a real gap, not glossed as
  "the same case."
- **`export.per_sheet`'s xlsx path re-copies the ENTIRE workbook once per exported sheet.**
  Correct (Task 12's reasoning), but O(sheets) full-workbook copies is a real cost on a large
  workbook with many tabs - acceptable for this wave's scope, worth a performance look if a
  deployment routinely exports a two-digit number of sheets.
- **`PkgFormat_ResolveCol` is a known, flagged duplicate of `PkgClean_ResolveFormatCol`**
  (Decision 8) - not a defect, but a debt this plan deliberately does not pay down, because
  paying it down means reopening an already-committed phase-2 diff.
- **The escalate-to-Core note for `Core_SafeFileName` assumes W5 (backup) needs it.** If W5's
  actual design turns out not to need a filename sanitizer (e.g., if it names backups by
  workbook name plus timestamp with no user-supplied text in the path), the promotion was
  still correct per the helper-promotion plan's own rule ("two consumers, or already
  duplicated" - `export.pdf` and `export.per_sheet` are two consumers by themselves), just
  not for the reason stated to W5. Not a risk to this wave either way.
