# W4 - Summarize & Analyze, plus Match & Enrich: implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: use `superpowers:subagent-driven-development`
> to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship the six frozen W4 actions - four `summarize.*` in a new `PkgSummarize`, two
`enrich.*` in a new `PkgEnrich` - behind contract v2.7.0, so a chain can refresh pivots,
recalculate, build a summary table, freeze a pivot to values, and join in reference columns
without a volatile lookup formula anywhere.

**Architecture:** Two new packages, no new Core helper. `PkgSummarize` has four independent
actions - no two of them share real work, so there is no shared engine. `PkgEnrich` has two
actions that ARE the same job with different inputs - `enrich.join` asks for everything at
run time, `enrich.mapped_columns` reads four of its five inputs from Setup - so both delegate
to one private `PkgEnrich_JoinEngine`, the same split `PkgImport`'s two combines already use
for `PkgImport_CombineEngine`. `PkgEnrich` consumes `Core_ResolveColumn`, the letter-or-header
column resolution phase 2 already promotes to Core
(`2026-07-28-core-helper-promotion-plan.md`, Task 2 Step 2) - phase 2 also migrates
`PkgClean_ResolveFormatCol`'s call site onto it, so this wave does not touch Core at all.

**Tech stack:** VBA (Excel/Windows, one concatenated `.bas`), PowerShell for the build, Node
for the syntax check. No test framework - see *Verification*.

**Written against the Core API phase 2 adds** (`2026-07-28-core-helper-promotion-plan.md`):
`Core_AskRange`, `Core_AskHasHeader`, `Core_AskCols`, `Core_ParseCols`, `Core_NarrowToUsed`,
`Core_DestUsable`, `Core_NewSheet`, `Core_SafeSheetName`, `Core_CleanText`, `Core_IsAlpha`,
plus phase 1's `Core_Refuse` and `Core_IsEngineSheet`. None of these exist in `src/` yet - all
eight wave plans are written before phase 1's first line of code, per plan-of-record §6 - so
every reference below is to the *plan text*, not to shipped code.

---

## Why this wave is shaped this way

The catalog (`docs/macro-catalog-plug-and-play.md`, sections 4 and 6) gives six items, two of
which - `summarize.refresh_pivots` and `summarize.recalculate` - already exist, buried inside
`PkgImportExport_Import`'s tail (`src/packages/PkgImportExport.bas:178-188`):

```vb
    Application.Calculation = xlCalculationAutomatic
    For Each sh In ThisWorkbook.Worksheets
        For Each pt In sh.PivotTables
            On Error Resume Next
            pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
            pt.RefreshTable
            On Error GoTo Fail
        Next pt
    Next sh
    Application.CalculateFull
```

and two - `enrich.join` and `enrich.mapped_columns` - read as near-identical catalog entries
("VLOOKUP/XLOOKUP-style join in code" vs. "add mapped columns from a lookup table") that turn
out, on design, to be the same match-and-copy operation with different input sources. Both
shapes force a real decision instead of a default, and both are made explicitly below rather
than by omission.

---

## Decisions taken in this plan

| # | Question | Decision |
|---|---|---|
| 1 | `summarize.refresh_pivots` / `summarize.recalculate` exist today inside `PkgImportExport_Import`. Legacy keeps its copy, calls the new action, or the logic moves to Core? | **Legacy keeps its own two-line copy. The new actions get their own, independent ~15-line implementation. Nothing is promoted to Core in this wave.** A cross-package call is forbidden, which rules out "calls the new action" outright. Promoting to Core was the other real option, and the reason it loses: phase 2's helper-promotion pass already audited `PkgImport` and `PkgClean` for exactly this kind of duplication and produced an explicit "deliberately NOT promoted" list for anything with one consumer or a speculative second one. This pivot-refresh loop was not in that audit at all - it lives in a third package phase 2 did not touch - so promoting it now would be a second wave silently expanding a already-reviewed pass's scope instead of that pass's own diff. The duplication is real but small (about 15 lines, no branching, copied from working code), and touching `PkgImportExport.bas` - a shipped package this wave does not own - for a 15-line saving raises the risk to an already-live finance workflow for no user-visible gain. **If a third consumer appears in a later wave, promote then, in that wave's own diff and PR** - the same YAGNI rule phase 2 states for `PkgImport_WorkbookPaths` and `PkgImport_AlreadyOpen`. |
| 2 | Do `enrich.join` and `enrich.mapped_columns` share one engine? | **Yes - one `PkgEnrich_JoinEngine`, matching PkgImport's two-combines split exactly.** The two actions differ only in WHERE their five inputs (target range, target key column, reference range, reference key column, which reference columns to copy and what to call them) come from: `enrich.join` asks for all five at run time; `enrich.mapped_columns` reads the last four from Setup config and asks only for the target range. That is the same shape as `import.combine_worksheets` vs. `import.combine_workbooks` differing only in where the source list comes from (design section 3.3, `PkgImport.bas:160-179`). The match-and-copy mechanics - build a dictionary, scan the target once, write one block - are identical, so they are capped at one coherent engine rather than forked. |
| 3 | Duplicate key in the reference table - refuse, first wins, or last wins? | **First occurrence wins, and the count of duplicate keys encountered is always reported in the completion message, never silent.** Refusing outright fails the action on any real reference table, which is rarely perfectly clean, and defeats the feature entirely. Last-wins is no more "correct" than first and is harder for a reader to predict by scanning the sheet top to bottom. First-wins also matches what VLOOKUP and INDEX/MATCH already return, so a user replacing a formula with this action gets the same answer they were getting - the duplicate count is the one thing the formula never told them, and is the actual safety improvement over what it replaces. |
| 4 | A key present in the target but not the reference table? | **Left blank in the new columns, counted as unmatched, never refused.** A partial match is the ordinary outcome at month-end (a new store not yet in the master list), not an error condition. |
| 5 | Does either side of a join, or the pivot builder, ever ask "is row 1 a header"? | **No - row 1 is always the header, on every range this wave prompts for.** `Core_AskHasHeader` exists and is available, but every use here needs the header text itself (to name new columns, or to name pivot fields) rather than just needing to skip it, so there is nothing to ask - the header's presence is a hard requirement stated in the prompt text ("include its header row"), not a run-time branch. This removes a whole class of "what if they say no" bugs at the cost of one constraint stated up front. |
| 6 | `enrich.join`'s run-time key-column prompt returns a `Core_AskCols` result, which can name more than one column. What if the user names two? | **Refused, not silently narrowed to the first column.** A silent narrowing here is the same failure shape the duplicate-key question warns about, just moved one step earlier: a user who typed "1,2" meaning a composite key would get single-column matching with no indication anything was dropped, and their numbers would be wrong. Composite match keys are out of scope for v1 - refuse and name the count, rather than guess. |
| 7 | `PkgEnrich` needs `PkgClean_ResolveFormatCol`'s letter-or-header column resolution. Copy it, or promote it? | **Neither - phase 2 already promoted it.** `Core_ResolveColumn` merges in phase 2 (`2026-07-28-core-helper-promotion-plan.md`, Task 2 Step 2), well before W4, with the exact signature this plan needs: `Core_ResolveColumn(rng As Range, ByVal spec As String) As Long`. Phase 2 also migrates `PkgClean_ResolveFormatCol`'s call site onto it and deletes that private, so the duplicate pair this plan once weighed - promote-and-leave-`PkgClean`-alone vs. copy - never exists: there is no second implementation to copy from and no private left in `PkgClean` to escalate a future cleanup pass against. W4 simply consumes the helper Core already has. |
| 8 | `summarize.build_pivot`'s dimension/measure configuration - a config table, or run-time prompts? | **Fully run-time prompted, no config fields at all.** Which columns to group by and which to sum describes the range the caller is about to pick, not a deployment fact - the same normative rule contract §8 already states for `clean`'s withdrawn `clean_keycols`. Every value column is aggregated by SUM only; a per-column aggregation choice (average, count, ...) would need a second answer per measure, which is the config DSL this plan was explicitly told not to build. Anyone who needs a different aggregation can change the field's function by hand on the finished PivotTable - nothing here prevents that. |
| 9 | `summarize.build_pivot`'s destination sheet name? | **Derived from the picked range's own reference (`Core_RangeRef`), safe-sheeted, refused via `Core_DestUsable` if taken.** A fixed name like "Summary" run twice in one session collides with itself on the second run; deriving from the range makes two summaries of two different ranges land on two different sheets without asking the user to type a name. |
| 10 | `summarize.freeze_pivot` - which pivot range, and what does it preserve? | **`pt.TableRange2` (the whole visible pivot, including the report-filter area), copied as values + number formats + column widths.** `PasteSpecial xlPasteFormats` does not carry column width, so it is copied in a second explicit loop - otherwise the frozen copy would read correctly but not look like what was on screen, which defeats a "make it distributable" action. The source PivotTable is never refreshed and never written to - only read from via `.Copy`. |
| 11 | `summarize.refresh_pivots` - does one bad pivot abort the rest? | **No. Every pivot on every non-engine sheet is attempted regardless of earlier failures; each failure is counted and named (sheet, pivot name, error) in the completion report.** Aborting would let one stale external connection block every other, unrelated pivot in the workbook from refreshing. This deliberately does NOT inherit the legacy `PkgImportExport_Import` loop's silent `On Error Resume Next` - that loop is import's own internal step and stays as it is (decision 1), but a *standalone, user-facing* action reporting nothing about a failure is the exact defect `Core_Refuse` exists to cure, just for a loop instead of a single precondition. If NO pivots exist anywhere in the workbook, the action refuses instead of reporting a vacuous "refreshed 0 of 0". |
| 12 | Field key prefixes - collisions? | **None.** `grep -rn 'Core_AddField reg, "k_' src/packages` (run 2026-07-28) shows every existing key prefixed `k_export_`, `k_settings_`, `k_master_`, `k_import_`, or `k_clean_`. `k_summarize_*` is not used - `PkgSummarize` declares no config fields at all, the same starting point `PkgClean` had before v2.3.0. `k_enrich_lookup_sheet`, `k_enrich_lookup_key`, and `k_enrich_mapped` are new and collide with nothing. |

### Safety story per action

| Action | Safety story |
|---|---|
| `summarize.refresh_pivots` | Never overwrites or deletes anything; a bad pivot is counted and named, not hidden. Refuses if there is nothing to refresh. |
| `summarize.recalculate` | Never refuses - recalculating changes no stored value, only formula results, so there is no precondition that can fail. |
| `summarize.build_pivot` | Never overwrites a sheet - `Core_DestUsable` refuses and names the conflict, same rule as the two combines. |
| `summarize.freeze_pivot` | Never modifies the source PivotTable (read-only `.Copy`); never overwrites a sheet, same `Core_DestUsable` rule. A click outside any pivot is refused by name instead of raising an engine error. |
| `enrich.join` | Never overwrites existing columns - new columns are always appended to the right of the picked range. Refuses on a multi-column match key rather than silently matching on the wrong thing. Reports duplicate-key and unmatched-row counts rather than hiding them. |
| `enrich.mapped_columns` | Same append-only, never-overwrite rule as `enrich.join`. Refuses if the configured reference sheet is missing, if the configured key column resolves nowhere, or if nothing is configured yet - the same "nothing configured" refusal shape `clean.standardize_formats` already uses for `k_clean_formats`. |

---

## Global constraints

Copied from `2026-07-28-track2-completion-plan-of-record.md` §4, binding on this wave:

- **`SCHEMA_VER` stays `"v2.0"`.**
- **Action ids are immutable**, two-segment `<domain>.<operation>`, lowercase ASCII and
  underscores, each segment starting with a letter. No camelCase, no alias table. `summarize`
  and `enrich` are already `active` in contract §7's vocabulary table (verified by grep,
  2026-07-28) - no vocabulary edit is owed.
- **Shared helpers live ONLY in Core. Cross-package calls are forbidden.**
- **Pure ASCII in every `.bas`**; `Option Explicit`; max 25 line continuations per statement;
  no module-level declaration after the first procedure; no VBA reserved word as an
  identifier.
- **Every action entry point is `Public Sub Pkg<Name>_<Verb>(ctx As Object)`** - one `Object`
  parameter, nothing else. `build.ps1` fails the build otherwise.
- **TDD is off.** Do not propose a test suite. Verification is `pwsh build/build.ps1`, `node
  scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`, and deliberate injections.
- **`dist/` is gitignored** - never `git add -f` the bundle. `.vscode/` stays untracked.
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.
- **No manual gate document.** Run-time-only claims go to
  `docs/designs/2026-07-28-deferred-validation-log.md` (Task 9), per plan-of-record §3.5.
- **Every package private carries its package prefix** (`PkgEnrich_JoinEngine`, never
  `JoinEngine`) - CLAUDE.md rule 4, enforced by `build.ps1`'s prefix guard once phase 2 ships.

---

## Verification, since TDD is off

Every task verifies with:

1. `pwsh build/build.ps1` - exit 0, `Core + N package(s)`, balanced `Sub`/`Function` counts.
   **N rises by exactly 2 over whatever it was immediately before this wave's first commit**
   - this plan cannot state an absolute number, because W2 and W3 land first and each adds
   packages of their own.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` - no diagnostics.
3. **A grep proving each new action id and field key exists exactly once** in its manifest.
4. **A deliberate injection** (Task 8) proving `CoreCandidates()` registration actually
   matters - the single easiest way to ship a dead wave, per the shared brief.
5. Anything genuinely run-time-only (a live PivotTable, a live join over real data) is logged
   in `docs/designs/2026-07-28-deferred-validation-log.md` (Task 9), not gated here.

---

## File structure

| File | Change |
|---|---|
| `spec/contract.md` | v2.7.0. §8 gains a `summarize` package subsection (no fields, four actions) and an `enrich` package subsection (three fields, two actions); §10 gains the version entry. |
| `src/core/Core.bas` | `CoreCandidates()` gains `c.Add "Summarize"` and `c.Add "Enrich"`. |
| `src/packages/PkgSummarize.bas` | **New.** Manifest, four public entry points, no private helpers. |
| `src/packages/PkgEnrich.bas` | **New.** Manifest, two public entry points, one private `PkgEnrich_JoinEngine`. |
| `docs/designs/2026-07-28-deferred-validation-log.md` | Appended, not created (phase 1 creates it). |
| `CHANGELOG.md`, `VERSION`, `docs/code-map.md` | v0.15.0. |

---

### Task 1: Contract v2.7.0

**Owner: Opus, written directly.** Every later task implements what this states; a slip here
is a slip every reviewer reads against.

**Files:** Modify `spec/contract.md` - header (line 3), §8 (after the `clean` subsection,
around line 797), §10 (top of the version list, around line 828).

**Interfaces:** Consumes nothing. Produces the normative text Tasks 2-7 implement.

- [ ] **Step 1: Bump the version header**

Line 3: `**Contract version: 2.6.0**` -> `**Contract version: 2.7.0**` (2026-07-28). (W3 ships
2.6.0 first; if this wave lands before W3 merges, use whatever the current header says and
bump its minor by one - the number that matters is that this wave's own bump is +1 minor,
additive.)

- [ ] **Step 2: Add the `summarize` package subsection**

Insert immediately after the `clean` package subsection ends (after its Actions table, before
the `---` that precedes "## 9. Platform-abstraction appendix"):

````markdown
### Package `summarize` - "Summarize & Analyze" (v1.0.0)

**Fields**

None. Every action either takes no input (`summarize.recalculate`), operates on every sheet
(`summarize.refresh_pivots`), or prompts for its range and columns at run time
(`summarize.build_pivot`, `summarize.freeze_pivot`) - none of the four describes a fact about
the deployment, so none is a config field. Same precedent `clean` set before v2.3.0 gave it
its first field.

**Actions**

| id | caption | semantics |
|---|---|---|
| `summarize.refresh_pivots` | Refresh All PivotTables | Refresh every PivotTable on every sheet in the workbook, skipping engine-owned sheets (§7's `Core_IsEngineSheet` test). A PivotTable whose refresh raises MUST NOT stop the others; it MUST be counted and named (sheet, pivot, error) in the completion report rather than silently skipped or aborting the run. With no PivotTables anywhere in the workbook the action MUST refuse rather than report a vacuous success. |
| `summarize.recalculate` | Recalculate Workbook | Force a full recalculation of every formula in the workbook. This action MUST NOT refuse under any condition: it changes no stored value, so no precondition can be unmet. |
| `summarize.build_pivot` | Build a Summary Table | Prompt for a range whose first row MUST be treated as its header, then which column(s) to group by and which to sum. Write the result as a new PivotTable on a new sheet. Every value column is aggregated by SUM; a per-column aggregation choice is out of scope for this version. The destination sheet MUST NOT overwrite an existing sheet - same refusal rule as `import.combine_worksheets`. |
| `summarize.freeze_pivot` | Freeze a Pivot to Values | Prompt for a cell; resolve the PivotTable that cell belongs to, refusing by name if it belongs to none. Copy that pivot's current output - values, number formats, and column widths - to a new sheet. The action MUST NOT modify the source PivotTable (no refresh, no write to its cells). Same destination-sheet refusal rule as `summarize.build_pivot`. |
````

- [ ] **Step 3: Add the `enrich` package subsection**

Insert immediately after the new `summarize` subsection:

````markdown
### Package `enrich` - "Match & Enrich" (v1.0.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_enrich_lookup_sheet` | sheet | simple | `Ref_Master` | The reference sheet `enrich.mapped_columns` reads from. Unlike a destination-sheet field (e.g. `k_import_dest`), this one MUST already exist - a conforming implementation MUST refuse rather than create it. `usedBy` `enrich.mapped_columns`. |
| `k_enrich_lookup_key` | text | simple | `Store ID` | A column letter or header-text spec, resolved the same way as `k_clean_formats`' Column cell (§8, package `clean`): a one-to-three-letter value is read as a letter, anything else is matched against the target's own first row. Used against BOTH the reference sheet and whatever range the caller selects - the matching column has one name for both sides. `usedBy` `enrich.mapped_columns`. |
| `k_enrich_mapped` | table | simple | none | One row per column to add. Two columns: **Reference column** (letter or header text in the reference sheet) and **New column header** (the header given to the added column). Declares NO default rows, for the same reason `k_clean_formats` does not (§8, package `clean`): §5 resolves an empty stored table to the field's declared defaults, so demo rows would make `enrich.mapped_columns`' required "nothing configured" refusal unreachable. `usedBy` `enrich.mapped_columns`. |

`k_enrich_mapped` header: `Reference column (letter or header)|New column header`.

**Actions**

| id | caption | semantics |
|---|---|---|
| `enrich.join` | Join / Enrich Columns | Prompt for a target range and a reference range, each with its header row included and treated as such - there is no "does this have a header" branch. Prompt for the ONE matching column on each side (a multi-column answer MUST be refused, not narrowed to its first entry) and which reference column(s) to copy in. Match by building an in-memory dictionary keyed on the reference's matching column, then scanning the target once - an implementation MUST NOT re-scan the reference table per target row (dictionary lookup, not nested iteration). A reference key seen more than once MUST resolve to its FIRST occurrence and MUST be counted and reported as a distinct number, never silently overwritten and never a reason to refuse the whole action. A target row matching no reference key MUST be left blank in the new columns and counted as unmatched, not treated as an error. New columns are written immediately to the right of the target range, headed with the reference table's own header text. |
| `enrich.mapped_columns` | Add Mapped Columns | Prompt only for the target range; the reference sheet, its matching column, and which columns to add come from `k_enrich_lookup_sheet` / `k_enrich_lookup_key` / `k_enrich_mapped`. MUST share the match implementation with `enrich.join` - same first-wins-and-counted duplicate-key rule, same unmatched-row rule, same append-to-the-right placement - differing only in where its five inputs come from. MUST refuse if the configured reference sheet does not exist, if the configured key does not resolve on the reference sheet, or if `k_enrich_mapped` is empty or none of its rows resolve to a real reference column. |
````

- [ ] **Step 4: Add the §10 entry**

Insert as the first versioned bullet in §10:

````markdown
- **v2.7.0 (additive):** two new packages, `summarize` (four actions, no fields) and `enrich`
  (three fields, two actions). Both domains were already listed `active` in §7's vocabulary
  table before this version - this bullet backs that listing with a shipping package, and
  changes no vocabulary. `SCHEMA_VER` does **not** move - it stays `"v2.0"` and every existing
  configuration keeps its meaning. No `key`, action `id`, or field `kind` from any earlier
  version is renamed, removed, or reinterpreted.
````

- [ ] **Step 5: Verify**

```bash
grep -n 'Contract version: 2.7.0' spec/contract.md
grep -n "Package \`summarize\`\|Package \`enrich\`" spec/contract.md
pwsh build/build.ps1
```

Expected: both greps hit; the build still passes (Task 2/5's manifests are what makes the
domain-vocabulary parse meaningful, but the contract table itself already lists both domains
`active`, so this build should already pass before any package code lands).

- [ ] **Step 6: Commit**

```bash
git add spec/contract.md
git commit -m "spec: contract v2.7.0 - summarize and enrich packages"
```

---

### Task 2: `PkgEnrich` manifest and the shared join engine

**Owner: Opus, written directly.** The duplicate-key rule is the central design decision the
brief calls out; it lives here and must be right once.

**Files:** Create `src/packages/PkgEnrich.bas`.

**Interfaces:**
- Consumes: `Core_DeclarePackage`, `Core_AddSection`, `Core_AddField`, `Core_AddTable`,
  `Core_AddAction`, `Core_CleanText` (phase 2).
- Produces: `PkgEnrich_Manifest(reg As Object)`, and the private
  `PkgEnrich_JoinEngine(targetRng As Range, ByVal targetKeyCol As Long, lookupRng As Range,
  ByVal lookupKeyCol As Long, valueCols As Variant, newHeaders As Variant, ByRef matched As
  Long, ByRef unmatched As Long, ByRef dupKeys As Long, ByRef appendedCols As Long)`. Tasks 3
  and 4 both call it; nothing outside `PkgEnrich.bas` may.

- [ ] **Step 1: Write the file header and manifest**

```vb
Attribute VB_Name = "PkgEnrich"
Option Explicit

'================================================================
' PkgEnrich - matching rows against a reference table and copying
' columns in, without a volatile lookup formula. Wave 4 of the catalog
' build-out (docs/macro-catalog-plug-and-play.md, section 4).
'
' Two actions, both the same match-and-copy operation with different
' input sources, so both delegate to ONE private engine:
'   enrich.join            everything picked at run time
'   enrich.mapped_columns  the reference sheet, its key, and which
'                          columns to add all come from Setup; only the
'                          target range is still picked at run time
'
' That is the same split PkgImport's two combine actions use for
' PkgImport_CombineEngine (PkgImport.bas:160-179) - one engine, two
' different sources for the same five inputs.
'
' The engine is dictionary-based (Scripting.Dictionary, late-bound),
' never a nested search, because this runs over a consolidated dataset
' that can be tens of thousands of rows: building the reference
' dictionary once is O(reference rows), then matching every target row
' is O(target rows) with an O(1) lookup each - a nested loop over both
' would be O(target rows * reference rows), minutes instead of seconds
' on a real month-end file.
'
' A reference key seen more than once resolves to its FIRST occurrence
' and is counted, never refused and never silently overwritten - see
' the plan's decision table for why. A target row matching nothing is
' left blank and counted, never an error - a partial match is the
' ordinary outcome at month-end, not a defect.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgEnrich_Manifest(reg As Object)
    Core_DeclarePackage reg, "enrich", "Match & Enrich", "1.0.0"

    Core_AddSection reg, "MAPPED COLUMNS"
    Core_AddField reg, "k_enrich_lookup_sheet", "Reference sheet to pull from", "sheet", "Ref_Master", _
        "simple", _
        "The sheet enrich.mapped_columns reads from. It must already exist here - unlike a destination sheet, this one is never created.", _
        "enrich.mapped_columns"
    Core_AddField reg, "k_enrich_lookup_key", "Matching column (in both sheets)", "text", "Store ID", _
        "simple", _
        "A column letter (e.g. B) or the exact header text in the reference sheet's first row. The range you pick when you run this must have a column with the same letter or header text - that is the column matched against.", _
        "enrich.mapped_columns"
    Core_AddTable reg, "k_enrich_mapped", "Columns to add", _
        "Reference column (letter or header)|New column header", "simple", _
        "One row per column to copy in from the reference sheet. Left cell: a column letter or header text in the reference sheet. Right cell: the header to give the new column here. Nothing happens until you fill this in.", _
        "enrich.mapped_columns"

    Core_AddAction reg, "enrich.join", "Join / Enrich Columns", "PkgEnrich_Join", _
        "Pick a range to enrich and a reference table (each including its header row); matches rows by a key column and copies in the reference columns you choose. Everything is picked at run time - nothing here is configured on Setup.", _
        "Match & Enrich"
    Core_AddAction reg, "enrich.mapped_columns", "Add Mapped Columns", "PkgEnrich_MappedColumns", _
        "Pick a range; adds the columns configured above from the reference sheet, matched on the key column. The mapping is set up once above and runs the same way every time.", _
        "Match & Enrich"
End Sub
```

- [ ] **Step 2: Write the shared engine**

```vb
'================================================================
' Shared join engine. Both actions call this - see the file header for
' why one engine, and what "first wins" and "unmatched stays blank"
' mean for a finance dataset.
'
' Reads both ranges into arrays and writes the result as ONE block,
' never cell by cell - the reasoning PkgImport_AppendSheet's own header
' comment gives applies here just as hard (do-not-revert list, plan-of-
' record section 4.1): a per-cell COM write over tens of thousands of rows is
' minutes, not seconds, even with screen updating off.
'
' Both ranges are assumed to have a header row in row 1 - the caller's
' job, not this engine's; see Decision 5.
'================================================================
Private Sub PkgEnrich_JoinEngine(targetRng As Range, ByVal targetKeyCol As Long, _
                                 lookupRng As Range, ByVal lookupKeyCol As Long, _
                                 valueCols As Variant, newHeaders As Variant, _
                                 ByRef matched As Long, ByRef unmatched As Long, _
                                 ByRef dupKeys As Long, ByRef appendedCols As Long)
    Dim tArr As Variant, lArr As Variant, outArr As Variant
    Dim dict As Object, key As String
    Dim nRows As Long, mRows As Long, nCols As Long, c As Long, r As Long
    Dim hitRow As Variant, srcCol As Long
    Dim destWs As Worksheet, firstNewCol As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1                            ' vbTextCompare - same rule as PkgImport's headerMap

    tArr = targetRng.Value
    lArr = lookupRng.Value
    nRows = UBound(tArr, 1)                         ' row 1 is the header, both sides always have one
    mRows = UBound(lArr, 1)
    nCols = UBound(valueCols) - LBound(valueCols) + 1

    ' Build the reference dictionary once. Row 1 is the header and is skipped.
    ' DUPLICATE KEYS: the row already in the dictionary wins - the first one
    ' seen, top to bottom - and every later duplicate only increments the
    ' count. Nothing here refuses on a duplicate: a reference table with a
    ' repeated key is common, not exceptional, and the count is how the user
    ' finds out rather than getting a wrong number with no signal at all.
    For r = 2 To mRows
        key = Core_CleanText(CStr(lArr(r, lookupKeyCol)))
        If Len(key) > 0 Then
            If dict.Exists(key) Then
                dupKeys = dupKeys + 1
            Else
                dict.Add key, r
            End If
        End If
    Next r

    ' Build the output block in memory: row 1 = the new headers, rows 2..nRows
    ' = the matched reference values, or blank for a row with no match.
    ReDim outArr(1 To nRows, 1 To nCols)
    For c = 1 To nCols
        outArr(1, c) = CStr(newHeaders(LBound(newHeaders) + c - 1))
    Next c
    For r = 2 To nRows
        key = Core_CleanText(CStr(tArr(r, targetKeyCol)))
        If Len(key) > 0 And dict.Exists(key) Then
            hitRow = dict(key)
            matched = matched + 1
            For c = 1 To nCols
                srcCol = CLng(valueCols(LBound(valueCols) + c - 1))
                outArr(r, c) = lArr(hitRow, srcCol)
            Next c
        Else
            unmatched = unmatched + 1               ' outArr row stays blank for this row's new columns
        End If
    Next r

    ' One write, immediately to the right of the target range - this never
    ' overwrites a column the caller already had, and it is the reason both
    ' actions are "append", never "replace", in their safety story.
    Set destWs = targetRng.Parent
    firstNewCol = targetRng.Column + targetRng.Columns.Count
    destWs.Range(destWs.Cells(targetRng.Row, firstNewCol), _
                 destWs.Cells(targetRng.Row + nRows - 1, firstNewCol + nCols - 1)).Value = outArr
    appendedCols = nCols
End Sub
```

- [ ] **Step 3: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: `PkgEnrich_Manifest` and `PkgEnrich_JoinEngine` are the only two procedures so far;
the build should still pass structurally even with no actions calling the engine yet (Tasks 3
and 4 add those). If `build.ps1`'s entry-point resolver complains about an action with no
handler, that is expected until Task 3/4 land - do not silence it by adding a stub.

- [ ] **Step 4: Commit**

```bash
git add src/packages/PkgEnrich.bas
git commit -m "feat(enrich): PkgEnrich manifest and the shared join engine"
```

---

### Task 3: `enrich.join`

**Owner: Opus, written directly.** The brief names this "the valuable one" - it is where a
finance user replaces a volatile formula with this action, and the multi-column-key refusal
(Decision 6) has to be right.

**Files:** Modify `src/packages/PkgEnrich.bas`.

**Interfaces:**
- Consumes: `Core_AskRange`, `Core_NarrowToUsed`, `Core_AskCols`, `Core_ParseCols`,
  `Core_Refuse`, `Core_Done`, `Core_Fail`, `Core_SpeedOn`, `Core_SpeedOff`, `Core_RangeRef`
  (all phase 1/2), `PkgEnrich_JoinEngine` (Task 2).
- Produces: `Public Sub PkgEnrich_Join(ctx As Object)`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Ad hoc join: every input picked at run time
'================================================================
Public Sub PkgEnrich_Join(ctx As Object)
    Dim targetRng As Range, lookupRng As Range
    Dim tKeys As String, lKeys As String, vKeys As String
    Dim targetKeyArr As Variant, lookupKeyArr As Variant, valueColsArr As Variant
    Dim newHeaders() As String, i As Long
    Dim matched As Long, unmatched As Long, dupKeys As Long, appendedCols As Long
    Const TITLE As String = "Join / Enrich Columns"
    On Error GoTo Fail

    Set targetRng = Core_AskRange(TITLE & " - the data to enrich (include its header row)")
    If targetRng Is Nothing Then Exit Sub                     ' user cancelled
    Set targetRng = Core_NarrowToUsed(targetRng)
    If targetRng Is Nothing Or targetRng.Rows.Count < 2 Then
        Core_Refuse ctx, "Your selection needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(targetRng, tKeys, _
            "Which ONE column holds the value to match on? Row 1 is treated as the header.", _
            TITLE) Then Exit Sub
    targetKeyArr = Core_ParseCols(tKeys)
    If UBound(targetKeyArr) > LBound(targetKeyArr) Then
        Core_Refuse ctx, "Only one column is supported as the match key; you named " & _
               (UBound(targetKeyArr) - LBound(targetKeyArr) + 1) & "." & vbCrLf & vbCrLf & _
               "Enter a single column number and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set lookupRng = Core_AskRange(TITLE & " - the reference table (include its header row)")
    If lookupRng Is Nothing Then Exit Sub
    Set lookupRng = Core_NarrowToUsed(lookupRng)
    If lookupRng Is Nothing Or lookupRng.Rows.Count < 2 Then
        Core_Refuse ctx, "The reference table needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(lookupRng, lKeys, _
            "Which ONE column in the reference table is the matching key?", TITLE) Then Exit Sub
    lookupKeyArr = Core_ParseCols(lKeys)
    If UBound(lookupKeyArr) > LBound(lookupKeyArr) Then
        Core_Refuse ctx, "Only one column is supported as the match key; you named " & _
               (UBound(lookupKeyArr) - LBound(lookupKeyArr) + 1) & "." & vbCrLf & vbCrLf & _
               "Enter a single column number and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(lookupRng, vKeys, _
            "Which column(s) in the reference table do you want copied in? " & _
            "Numbered from the left of the reference table.", TITLE) Then Exit Sub
    valueColsArr = Core_ParseCols(vKeys)

    ' New columns are named after the reference table's OWN header text - this
    ' action asks nothing about naming, which is exactly the difference from
    ' enrich.mapped_columns, where the header text is configured once.
    ReDim newHeaders(LBound(valueColsArr) To UBound(valueColsArr))
    For i = LBound(valueColsArr) To UBound(valueColsArr)
        newHeaders(i) = CStr(lookupRng.Cells(1, CLng(valueColsArr(i))).Value)
    Next i

    Core_SpeedOn
    PkgEnrich_JoinEngine targetRng, CLng(targetKeyArr(LBound(targetKeyArr))), _
        lookupRng, CLng(lookupKeyArr(LBound(lookupKeyArr))), _
        valueColsArr, newHeaders, matched, unmatched, dupKeys, appendedCols
    Core_SpeedOff

    Core_Done ctx, "Added " & appendedCols & " column(s) to " & Core_RangeRef(targetRng) & "." & _
           vbCrLf & matched & " row(s) matched, " & unmatched & " row(s) did not." & _
           IIf(dupKeys > 0, vbCrLf & vbCrLf & dupKeys & " key(s) in the reference table were " & _
               "not unique; the FIRST row for each was used.", ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "enrich.join", Err
End Sub
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c '"enrich.join"' dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics. The grep prints `2` (one in `Core_AddAction`, one in
`Core_Fail`) - matching the W1 plan's corrected verification shape, not `1`.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgEnrich.bas
git commit -m "feat(enrich): enrich.join - ad hoc dictionary join, everything picked at run time"
```

---

### Task 4: `enrich.mapped_columns`

**Owner: Opus, written directly.** Shares the engine and the duplicate-key contract Task 3
established; kept with the same owner for consistency across the pair.

**Files:** Modify `src/packages/PkgEnrich.bas`.

**Interfaces:**
- Consumes: `Core_Cfg`, `Core_CfgTable`, `Core_SheetExists`, `Core_ResolveColumn` (Core, added
  by phase 2), `Core_AskRange`, `Core_NarrowToUsed`, `Core_Refuse`, `Core_Done`, `Core_Fail`,
  `Core_SpeedOn`, `Core_SpeedOff`, `Core_RangeRef`, `PkgEnrich_JoinEngine` (Task 2).
- Produces: `Public Sub PkgEnrich_MappedColumns(ctx As Object)`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Add the columns configured on Setup, matched by key
'================================================================
Public Sub PkgEnrich_MappedColumns(ctx As Object)
    Dim lookupSheetName As String, lookupKeySpec As String
    Dim lookupWs As Worksheet, lookupRng As Range, targetRng As Range
    Dim refCols() As String, newHdrs() As String, spare() As String
    Dim lookupKeyCol As Long, targetKeyCol As Long, colIdx As Long
    Dim valueCols() As Long, headers() As String, i As Long, n As Long
    Dim missing As String
    Dim matched As Long, unmatched As Long, dupKeys As Long, appendedCols As Long
    Const TITLE As String = "Add Mapped Columns"
    On Error GoTo Fail

    lookupSheetName = Trim(Core_Cfg(ctx, "k_enrich_lookup_sheet"))
    If Not Core_SheetExists(lookupSheetName) Then
        Core_Refuse ctx, "The reference sheet '" & lookupSheetName & "' does not exist in this " & _
               "workbook." & vbCrLf & vbCrLf & "Fix 'Reference sheet to pull from' on the Setup " & _
               "tab, Finalize, and run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If
    Set lookupWs = ThisWorkbook.Worksheets(lookupSheetName)
    Set lookupRng = Core_NarrowToUsed(lookupWs.UsedRange)
    If lookupRng Is Nothing Or lookupRng.Rows.Count < 2 Then
        Core_Refuse ctx, "'" & lookupSheetName & "' has no data rows to match against." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    lookupKeySpec = Trim(Core_Cfg(ctx, "k_enrich_lookup_key"))
    lookupKeyCol = Core_ResolveColumn(lookupRng, lookupKeySpec)
    If lookupKeyCol = 0 Then
        Core_Refuse ctx, "'" & lookupKeySpec & "' is not a column in '" & lookupSheetName & "'." & _
               vbCrLf & vbCrLf & "Fix 'Matching column' on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_CfgTable(ctx, "k_enrich_mapped", refCols, newHdrs, spare) Then
        Core_Refuse ctx, "No columns are set up to add yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Columns to add' table on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ReDim valueCols(0 To UBound(refCols) - LBound(refCols))
    ReDim headers(0 To UBound(refCols) - LBound(refCols))
    n = 0
    For i = LBound(refCols) To UBound(refCols)
        colIdx = Core_ResolveColumn(lookupRng, refCols(i))
        If colIdx = 0 Then
            missing = missing & vbCrLf & "   " & refCols(i)
        Else
            valueCols(n) = colIdx
            headers(n) = newHdrs(i)
            n = n + 1
        End If
    Next i
    If n = 0 Then
        Core_Refuse ctx, "None of the configured reference columns exist in '" & lookupSheetName & _
               "'." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If
    ReDim Preserve valueCols(0 To n - 1)
    ReDim Preserve headers(0 To n - 1)

    Set targetRng = Core_AskRange(TITLE & " - the data to enrich (include its header row)")
    If targetRng Is Nothing Then Exit Sub                     ' user cancelled
    Set targetRng = Core_NarrowToUsed(targetRng)
    If targetRng Is Nothing Or targetRng.Rows.Count < 2 Then
        Core_Refuse ctx, "Your selection needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    targetKeyCol = Core_ResolveColumn(targetRng, lookupKeySpec)
    If targetKeyCol = 0 Then
        Core_Refuse ctx, "'" & lookupKeySpec & "' is not a column in your selection." & vbCrLf & vbCrLf & _
               "The matching column is the same setting for both sides - pick a range here whose " & _
               "header includes it." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgEnrich_JoinEngine targetRng, targetKeyCol, lookupRng, lookupKeyCol, valueCols, headers, _
        matched, unmatched, dupKeys, appendedCols
    Core_SpeedOff

    Core_Done ctx, "Added " & appendedCols & " column(s) to " & Core_RangeRef(targetRng) & _
           " from '" & lookupSheetName & "'." & vbCrLf & matched & " row(s) matched, " & _
           unmatched & " row(s) did not." & _
           IIf(dupKeys > 0, vbCrLf & vbCrLf & dupKeys & " key(s) in '" & lookupSheetName & _
               "' were not unique; the FIRST row for each was used.", "") & _
           IIf(Len(missing) > 0, vbCrLf & vbCrLf & "These configured reference columns were not " & _
               "found and were skipped:" & missing, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "enrich.mapped_columns", Err
End Sub
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics, `Core + N package(s)` where `N` includes `Enrich` for the
first time (registration is Task 8 - until then `PkgEnrich`'s actions build and bundle but
Core never calls them; that gap is expected and closed by Task 8, not here).

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgEnrich.bas
git commit -m "feat(enrich): enrich.mapped_columns - config-driven join sharing PkgEnrich_JoinEngine"
```

---

### Task 5: `PkgSummarize` manifest, `refresh_pivots`, `recalculate`

**Owner: delegate to codex (`sonnet`).** Both actions are short, pattern-matched closely off
the legacy code already in `PkgImportExport.bas:178-188`, and carry no design decision this
task has to invent - Task 1 already settled the shape.

**Files:** Create `src/packages/PkgSummarize.bas`.

**Interfaces:**
- Consumes: `Core_DeclarePackage`, `Core_AddAction`, `Core_IsEngineSheet`, `Core_Refuse`,
  `Core_Done`, `Core_Fail`, `Core_SpeedOn`, `Core_SpeedOff` (all phase 1/2).
- Produces: `Public Sub PkgSummarize_RefreshPivots(ctx As Object)`,
  `Public Sub PkgSummarize_Recalculate(ctx As Object)`, `PkgSummarize_Manifest(reg As Object)`.

- [ ] **Step 1: Write the file header and manifest**

```vb
Attribute VB_Name = "PkgSummarize"
Option Explicit

'================================================================
' PkgSummarize - turning a cleaned dataset into pivots and summaries.
' Wave 4 of the catalog build-out (docs/macro-catalog-plug-and-play.md,
' section 6).
'
' Two of these four actions already exist, buried inside
' PkgImportExport_Import's tail end (PkgImportExport.bas ~178-188):
' refreshing every pivot and forcing a full recalculation after an
' import. This package gives them their own action ids so they can run
' standalone or as their own chain step - summarize.refresh_pivots and
' summarize.recalculate. The legacy import KEEPS its own two-line copy
' rather than calling these; see the wave plan's decision table for why
' this is deliberate, not an oversight.
'
' Declares NO config fields. Every action either takes no input at all
' (recalculate), scans every sheet with no per-run choice
' (refresh_pivots), or prompts for its range/columns at run time
' (build_pivot, freeze_pivot) - none of the four describes a fact about
' the deployment, so none belongs on Setup. Same precedent PkgClean set
' before v2.3.0 gave it its first field.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgSummarize_Manifest(reg As Object)
    Core_DeclarePackage reg, "summarize", "Summarize & Analyze", "1.0.0"

    Core_AddAction reg, "summarize.refresh_pivots", "Refresh All PivotTables", "PkgSummarize_RefreshPivots", _
        "Refresh every PivotTable on every sheet in this workbook. Your own Setup tabs are left out.", _
        "Summarize & Analyze"
    Core_AddAction reg, "summarize.recalculate", "Recalculate Workbook", "PkgSummarize_Recalculate", _
        "Force a full recalculation of every formula in the workbook, the same as pressing Ctrl+Alt+F9.", _
        "Summarize & Analyze"
    Core_AddAction reg, "summarize.build_pivot", "Build a Summary Table", "PkgSummarize_BuildPivot", _
        "Pick a range (with its header row), then which column(s) to group by and which to total. Writes a new PivotTable to a new sheet.", _
        "Summarize & Analyze"
    Core_AddAction reg, "summarize.freeze_pivot", "Freeze a Pivot to Values", "PkgSummarize_FreezePivot", _
        "Click any cell inside a PivotTable, then copy its current layout to a new sheet as static values and formats. The original PivotTable is never changed.", _
        "Summarize & Analyze"
End Sub
```

- [ ] **Step 2: Add `summarize.refresh_pivots`**

```vb
'================================================================
' ACTION - Refresh every PivotTable on every sheet
'
' A bad pivot (source range deleted, external connection unreachable)
' raises on .RefreshTable. One bad pivot does NOT stop the others - see
' the wave plan's Decision 11 for why aborting or staying silent are
' both wrong here.
'================================================================
Public Sub PkgSummarize_RefreshPivots(ctx As Object)
    Dim sh As Worksheet, pt As PivotTable
    Dim total As Long, refreshed As Long, failedNames As String
    Const TITLE As String = "Refresh All PivotTables"
    On Error GoTo Fail

    Core_SpeedOn
    For Each sh In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(sh.name) Then
            For Each pt In sh.PivotTables
                total = total + 1
                On Error Resume Next
                Err.Clear
                pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
                pt.RefreshTable
                If Err.Number = 0 Then
                    refreshed = refreshed + 1
                Else
                    failedNames = failedNames & vbCrLf & "   " & sh.name & " ! " & pt.name & _
                                  "  (#" & Err.Number & " " & Err.Description & ")"
                End If
                Err.Clear
                On Error GoTo Fail
            Next pt
        End If
    Next sh
    Core_SpeedOff

    If total = 0 Then
        Core_Refuse ctx, "There are no PivotTables anywhere in this workbook to refresh." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_Done ctx, "Refreshed " & refreshed & " of " & total & " PivotTable(s)." & _
           IIf(Len(failedNames) > 0, vbCrLf & vbCrLf & "These did not refresh:" & failedNames, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "summarize.refresh_pivots", Err
End Sub
```

- [ ] **Step 3: Add `summarize.recalculate`**

```vb
'================================================================
' ACTION - Force a full recalculation
'
' Never refuses: recalculating is always valid to run and changes no
' stored value, only formula results, so there is no precondition to
' fail. Exists as its own action so it can be its own chain step -
' summarize.refresh_pivots and the legacy import both want "and then
' recalculate" as their last step.
'================================================================
Public Sub PkgSummarize_Recalculate(ctx As Object)
    Const TITLE As String = "Recalculate Workbook"
    On Error GoTo Fail

    Core_SpeedOn
    Application.CalculateFull
    Core_SpeedOff

    Core_Done ctx, "Every formula in the workbook has been recalculated.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "summarize.recalculate", Err
End Sub
```

- [ ] **Step 4: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c '"summarize.refresh_pivots"\|"summarize.recalculate"' dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics; the grep prints `4` (each id appears once in
`Core_AddAction`, once in `Core_Fail`).

- [ ] **Step 5: Commit**

```bash
git add src/packages/PkgSummarize.bas
git commit -m "feat(summarize): PkgSummarize manifest, refresh_pivots and recalculate"
```

---

### Task 6: `summarize.build_pivot`

**Owner: Opus, written directly.** Real `PivotCaches.Create` / `CreatePivotTable` /
`PivotFields` API risk - the most Excel-specific code in this wave, per the brief.

**Files:** Modify `src/packages/PkgSummarize.bas`.

**Interfaces:**
- Consumes: `Core_AskRange`, `Core_NarrowToUsed`, `Core_AskCols`, `Core_ParseCols`,
  `Core_RangeRef`, `Core_SafeSheetName`, `Core_DestUsable`, `Core_NewSheet`, `Core_Refuse`,
  `Core_Done`, `Core_Fail`, `Core_SpeedOn`, `Core_SpeedOff`.
- Produces: `Public Sub PkgSummarize_BuildPivot(ctx As Object)`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Build a PivotTable from a picked range
'
' Row 1 of the picked range is ALWAYS its header - a PivotTable source
' needs field names and Excel itself requires them, so there is no
' "does this have a header" question to ask (Decision 5).
'
' Every chosen measure is SUMMED. A per-measure aggregation choice
' would need a second answer per column, which is the config DSL this
' plan was told not to build (Decision 8) - anyone who needs a
' different aggregation can change the field's function by hand on the
' finished PivotTable.
'================================================================
Public Sub PkgSummarize_BuildPivot(ctx As Object)
    Dim srcRng As Range, rowKeys As String, valKeys As String
    Dim rowColsArr As Variant, valColsArr As Variant, i As Long
    Dim destName As String, destWs As Worksheet
    Dim pc As PivotCache, pt As PivotTable
    Dim fld As PivotField, hdr As String
    Const TITLE As String = "Build a Summary Table"
    On Error GoTo Fail

    Set srcRng = Core_AskRange(TITLE & " - your data, including its header row")
    If srcRng Is Nothing Then Exit Sub                        ' user cancelled
    Set srcRng = Core_NarrowToUsed(srcRng)
    If srcRng Is Nothing Or srcRng.Rows.Count < 2 Then
        Core_Refuse ctx, "Your selection needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(srcRng, rowKeys, _
            "Which column(s) do you want to GROUP BY (e.g. Store, Month)? " & _
            "Numbered from the left of your selection.", TITLE) Then Exit Sub
    rowColsArr = Core_ParseCols(rowKeys)

    If Not Core_AskCols(srcRng, valKeys, _
            "Which column(s) do you want SUMMED (e.g. Amount)? " & _
            "Numbered from the left of your selection.", TITLE) Then Exit Sub
    valColsArr = Core_ParseCols(valKeys)

    ' Derived from the picked range's own reference, not a fixed "Summary" -
    ' Decision 9: a fixed name run twice in one session collides with itself.
    destName = Core_SafeSheetName(Core_RangeRef(srcRng) & " Summary")
    If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub

    Core_SpeedOn
    Set destWs = Core_NewSheet(destName)
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcRng)
    Set pt = pc.CreatePivotTable(TableDestination:=destWs.Range("A1"), TableName:="Pivot_" & destWs.Index)

    For i = LBound(rowColsArr) To UBound(rowColsArr)
        hdr = CStr(srcRng.Cells(1, CLng(rowColsArr(i))).Value)
        Set fld = pt.PivotFields(hdr)
        fld.Orientation = xlRowField
    Next i
    For i = LBound(valColsArr) To UBound(valColsArr)
        hdr = CStr(srcRng.Cells(1, CLng(valColsArr(i))).Value)
        pt.AddDataField pt.PivotFields(hdr), "Sum of " & hdr, xlSum
    Next i
    Core_SpeedOff

    Core_Done ctx, "Built '" & destName & "' from " & Core_RangeRef(srcRng) & "," & vbCrLf & _
           "grouped by " & (UBound(rowColsArr) - LBound(rowColsArr) + 1) & " column(s), summing " & _
           (UBound(valColsArr) - LBound(valColsArr) + 1) & " column(s).", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "summarize.build_pivot", Err
End Sub
```

**Known limitation, stated because it is a real gap, not covered here:** a source range with
two columns sharing the same header text will make `pt.PivotFields(hdr)` ambiguous - Excel
itself already renamed the second occurrence internally, and this code does not detect or
resolve that. No shipped source produces duplicate headers today (the two combines already
guarantee unique headers by construction), so this is listed under *Unknowns and risks*
rather than fixed here.

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c '"summarize.build_pivot"' dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics, grep prints `2`.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgSummarize.bas
git commit -m "feat(summarize): summarize.build_pivot - PivotTable from a picked range"
```

---

### Task 7: `summarize.freeze_pivot`

**Owner: Opus, written directly.** Judgment calls on what "preserves the visible layout"
means (`TableRange2`, explicit column widths) and the never-modify-source-pivot guarantee.

**Files:** Modify `src/packages/PkgSummarize.bas`.

**Interfaces:**
- Consumes: `Core_AskRange`, `Core_SafeSheetName`, `Core_DestUsable`, `Core_NewSheet`,
  `Core_Refuse`, `Core_Done`, `Core_Fail`, `Core_SpeedOn`, `Core_SpeedOff`.
- Produces: `Public Sub PkgSummarize_FreezePivot(ctx As Object)`.

- [ ] **Step 1: Add the action**

```vb
'================================================================
' ACTION - Copy a PivotTable's current output to a new sheet as values
'
' Reads the pivot from wherever the user clicks - Range.PivotTable
' raises when the cell is not part of one, refused here rather than
' surfacing as an engine error.
'
' Copies TableRange2 (the WHOLE visible pivot, report-filter area
' included) as values + number formats + column widths, so what lands
' on the new sheet looks like what was on screen, not just its numbers
' (Decision 10). It never calls .RefreshTable and never writes back to
' the source pivot's own cells - only .Copy reads from it. Field
' interactivity (expand/collapse, the field list, any GetPivotData
' formula elsewhere pointing at it) is NOT preserved - a static values
' sheet cannot have any of that, which is the entire point of freezing
' it, stated so it does not read as an oversight.
'================================================================
Public Sub PkgSummarize_FreezePivot(ctx As Object)
    Dim sel As Range, pt As PivotTable, destName As String, destWs As Worksheet
    Dim srcRng As Range, c As Long
    Const TITLE As String = "Freeze a Pivot to Values"
    On Error GoTo Fail

    Set sel = Core_AskRange(TITLE & " - click any cell inside the PivotTable to freeze")
    If sel Is Nothing Then Exit Sub                           ' user cancelled

    On Error Resume Next
    Set pt = sel.Cells(1, 1).PivotTable
    On Error GoTo Fail
    If pt Is Nothing Then
        Core_Refuse ctx, "The cell you picked is not part of a PivotTable." & vbCrLf & vbCrLf & _
               "Click a cell inside the PivotTable you want to freeze, then run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    destName = Core_SafeSheetName(pt.name & " (values)")
    If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub

    Core_SpeedOn
    Set srcRng = pt.TableRange2
    Set destWs = Core_NewSheet(destName)
    srcRng.Copy
    destWs.Range("A1").PasteSpecial Paste:=xlPasteValues
    destWs.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    ' Column widths are not part of xlPasteFormats - copied separately so the
    ' frozen copy reads the way the pivot looked on screen, not just the way
    ' its cells are formatted.
    For c = 1 To srcRng.Columns.Count
        destWs.Columns(c).ColumnWidth = srcRng.Columns(c).ColumnWidth
    Next c
    Core_SpeedOff

    Core_Done ctx, "Copied '" & pt.name & "' to '" & destName & "' as values." & vbCrLf & _
           "The original PivotTable was not changed.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "summarize.freeze_pivot", Err
End Sub
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c '"summarize.freeze_pivot"' dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics, grep prints `2`.

- [ ] **Step 3: Commit**

```bash
git add src/packages/PkgSummarize.bas
git commit -m "feat(summarize): summarize.freeze_pivot - copy a pivot to static values, never touching the source"
```

---

### Task 8: `CoreCandidates()` registration

**Owner: delegate to codex (`sonnet`).** Two-line mechanical add, but this is the single
easiest way to ship a dead wave (shared brief) - the deliberate injection in Step 3 is not
optional.

**Files:** Modify `src/core/Core.bas` - `CoreCandidates()` (`:952` as of this writing; confirm
via `docs/code-map.md` before editing, since phases 1/2 and W2/W3 add code above it first).

**Interfaces:** Consumes `PkgSummarize_Manifest`, `PkgEnrich_Manifest` (Tasks 2 and 5).
Produces nothing new - this is what makes Tasks 2-7's actions reachable at all.

- [ ] **Step 1: Add both packages to the candidate list**

```vb
Private Function CoreCandidates() As Collection
    Dim c As New Collection
    c.Add "ImportExport"            ' probes PkgImportExport_Manifest
    c.Add "Import"                  ' probes PkgImport_Manifest
    c.Add "Clean"                   ' probes PkgClean_Manifest
    c.Add "Summarize"               ' probes PkgSummarize_Manifest
    c.Add "Enrich"                  ' probes PkgEnrich_Manifest
    Set CoreCandidates = c
End Function
```

(If W2/W3 already landed, this list has more entries by the time this task runs - append
`"Summarize"` and `"Enrich"` to whatever is there, do not reset the list to the four shown
above.)

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -n 'c.Add "Summarize"\|c.Add "Enrich"' src/core/Core.bas
```

Expected: exit 0, no diagnostics, both greps hit exactly once.

- [ ] **Step 3: Deliberate injection - prove registration actually matters**

A registration line nobody has watched fail is not known to matter.

```bash
# Comment out the Enrich line and confirm its actions vanish from the built
# Setup surface / action catalog rather than merely failing to build (a
# missing CoreCandidates entry is NOT a build-time failure - the whole
# danger is that it builds, bundles, and silently never runs).
sed -i 's/c.Add "Enrich"/'"'"'c.Add "Enrich" - DISABLED FOR INJECTION TEST/' src/core/Core.bas
pwsh build/build.ps1
grep -c 'PkgEnrich_Manifest' dist/MacrosBundle.bas   # still present - the module still builds
git checkout -- src/core/Core.bas
pwsh build/build.ps1
```

Expected: the build still succeeds with the line commented out (this is exactly the danger -
`build.ps1` cannot catch a missing registration, only a run-time check or a human reading the
Setup tab can). Record in the PR body that this was tried and that the bundle's own build
output gives no signal either way - the run-time confirmation belongs in
`docs/designs/2026-07-28-deferred-validation-log.md` (Task 9).

- [ ] **Step 4: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): register PkgSummarize and PkgEnrich in CoreCandidates"
```

---

### Task 9: Version, changelog, code map, deferred-validation log, PR

**Owner: Opus.**

**Files:** `VERSION`, `CHANGELOG.md`, `docs/code-map.md`,
`docs/designs/2026-07-28-deferred-validation-log.md`.

- [ ] **Step 1: `VERSION` -> `0.15.0`**

Per plan-of-record §2's table - this assumes phases 1-2 and W2-W3 have already shipped their
own version bumps (0.11.0 through 0.14.0) by the time this task runs. If this wave lands out
of that order, state the actual predecessor version in the PR body rather than silently
picking a different number.

- [ ] **Step 2: `CHANGELOG.md` - a `## [0.15.0]` entry**

House style: state what changed for the user, and be honest about scope. Lead with the join
being dictionary-based (the value proposition), and name the duplicate-key and unmatched-row
reporting explicitly - that is the safety feature, not an implementation detail. State plainly
that `summarize.refresh_pivots` and `summarize.recalculate` already existed inside Import,
and are now also available standalone.

- [ ] **Step 3: Append to the deferred validation log**

`docs/designs/2026-07-28-deferred-validation-log.md` is created by phase 1 (per plan-of-record
§3.5) and accumulates run-time-only claims for the single end-to-end validation at the end of
the track. Append a new `## W4` section:

```markdown
## W4 - Summarize & Analyze, Match & Enrich

- `PivotCaches.Create` / `CreatePivotTable` / `PivotFields(...).Orientation` /
  `AddDataField` (summarize.build_pivot) has not run against a real workbook. Confirm a
  built PivotTable groups and sums exactly the columns named, and that a source range with
  a duplicate header name produces the documented error rather than a silent misgroup.
- `Range.PivotTable` correctly resolves the containing pivot from an arbitrary clicked cell,
  including a cell in the report-filter area (summarize.freeze_pivot) - not exercised.
- `pt.TableRange2` genuinely captures the whole visible pivot, and the explicit column-width
  copy loop actually reproduces the on-screen layout - not visually confirmed.
- The dictionary-based join's stated complexity claim (fast over tens of thousands of rows)
  has not been measured against a real consolidated dataset from W1's combine actions.
- `enrich.join` and `enrich.mapped_columns`' duplicate-key count and unmatched-row count
  have not been checked against a reference table with deliberately seeded duplicates.
- `summarize.refresh_pivots` skipping engine-owned sheets via `Core_IsEngineSheet`, and
  correctly naming a pivot that fails to refresh, has not been run against a workbook with
  a genuinely broken pivot connection.
- `CoreCandidates()` registration for `Summarize` and `Enrich` (Task 8): confirmed the build
  gives no signal either way when the line is missing; NOT confirmed that both packages'
  buttons actually appear on the generated Setup/Macros pages when registered.
```

- [ ] **Step 4: Refresh the code map**

```bash
node scripts/code-map.mjs
```

- [ ] **Step 5: Final gates**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas
grep -nE '^\s*Private\s+(Sub|Function)\s+' src/packages/PkgSummarize.bas src/packages/PkgEnrich.bas | grep -vE 'Private\s+(Sub|Function)\s+Pkg'
git status --short
```

Expected: exit 0; no diagnostics; `1`; no output from the prefix grep (`PkgEnrich_JoinEngine`
is the only private in either file, and it is prefixed); a tree clean but for the untracked
`.vscode/`.

- [ ] **Step 6: Review the diff, then one PR against `main`**

Not stacked on any other wave's branch - branch from `main` after phases 1-2 and W2-W3 have
merged, per plan-of-record §2.

**The PR body must state, in one line each:** that the legacy `PkgImportExport_Import` pivot
refresh/recalc loop was deliberately left untouched (Decision 1); that
`enrich.join`/`enrich.mapped_columns` share one engine and its duplicate-key rule is
first-wins-and-counted, never a refusal (Decisions 2-4); and the two Task 8 injection results.

---

## Unknowns and risks

- **`summarize.build_pivot`'s duplicate-header-name gap** (Task 6): a source range with two
  columns sharing a header text makes `pt.PivotFields(hdr)` ambiguous. No shipped source
  produces this today, so it is left as a known gap rather than solved here - fixing it would
  mean generating and tracking synthetic unique field names, which is a bigger change than
  this wave's charter.
- **`enrich.join`'s reference-range re-selection cost.** Each run rebuilds the dictionary from
  scratch even inside a chain that calls it twice in a row against the same reference table.
  Acceptable for a per-run action; would need a cached dictionary keyed by range address if a
  later wave wants a "join against the same reference table N times" chain step.
- **The `On Error Resume Next` / `On Error GoTo Fail` toggle inside
  `PkgSummarize_RefreshPivots`'s inner loop** is the same shape the legacy code already uses
  and is assumed safe on the same basis - VBA's error handler stack tolerates being
  re-armed per iteration - but has not been stress-tested against a pivot that raises
  mid-`RefreshTable` in a way that corrupts the loop's own state.
- **All run-time-only claims** are listed in `docs/designs/2026-07-28-deferred-validation-log.md`
  (Task 9) rather than gated here, per plan-of-record §3.5 - none of this wave's PivotTable
  or join behavior has run against a real Excel session.
