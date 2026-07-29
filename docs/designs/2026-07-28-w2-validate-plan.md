# W2 — Validate & Reconcile

> **For agentic workers:** REQUIRED SUB-SKILL: use `superpowers:subagent-driven-development`
> to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship `PkgValidate` — a new package implementing the catalog's five `validate.*`
actions (`docs/macro-catalog-plug-and-play.md` lines 92-105) — so a finance close can reconcile
two extracts on a key, sweep a range against configured rules, and flag duplicates, outliers
and missing mandatory fields, without any of the five ever repairing what they find.

**Architecture:** One new package, one engine per action, no cross-package calls. The flagship,
`validate.reconcile`, builds one `Scripting.Dictionary` key index per input range (not nested
loops) and classifies every key into exactly one of four outcomes — `In A only`, `In B only`,
`Duplicate key`, `Differs` — writing a flat exceptions sheet only when there is something to
report. The three colour-writing actions (`highlight_duplicates`, `highlight_outliers`,
`missing_fields`) share one safety story end to end: count first, confirm, never overwrite a
cell that already carries its own fill, and leave an undo instruction that is guaranteed
correct because of that same rule. `validate.scan` reuses that same never-overwrite-a-report
posture by writing to a sheet instead of colouring. Two Setup tables
(`k_validate_scan_rules`, `k_validate_outliers`) resolve a column by letter or header name,
the same shape `PkgClean_StandardizeFormats` uses for `k_clean_formats`, and both declare no
default rows for the same reason that one does, extended: a demo row here would mean a first
run could colour real cells nobody asked it to.

**Tech stack:** VBA (Excel/Windows, one concatenated `.bas`), PowerShell for the build, Node
for the syntax check. No test framework — see *Verification*.

**Runs after phase 1** (`2026-07-28-core-refuse-pass-plan.md`) **and phase 2**
(`2026-07-28-core-helper-promotion-plan.md`), both merged to `main` first. This plan is written
against the Core API phase 2 adds — `Core_AskRange`, `Core_AskHasHeader`, `Core_AskCols`,
`Core_ParseCols`, `Core_CleanText`, `Core_IsBlank`, `Core_IsAlpha`, `Core_IsDigits`,
`Core_NarrowToUsed`, `Core_DestUsable`, `Core_NewSheet`, `Core_SafeSheetName`,
`Core_ResolveFolder`, `Core_FolderExists` — plus phase 1's `Core_Refuse`. Sequencing and
rationale: `2026-07-28-track2-completion-plan-of-record.md` §3.1, §3.4, table in §2 (this is
phase 3 of 8; new actions: 5; repo `VERSION` 0.13.0; contract v2.5.0).

---

## Why this order / what this wave is for

This is the first wave to build on the Core API phase 2 promotes, and the first new domain
package since `PkgClean`. It exists to prove two things at once: that a fifth action-bearing
domain slots into the two-place registration (`src/packages/Pkg<Name>.bas` +
`CoreCandidates()`) exactly as cheaply as `clean` did, and that a domain whose entire contract
is "check and report, never repair" (contract §7) can hold that line even when two of its five
actions write something visible to the workbook — a cell colour.

**`validate.reconcile` is the highest-value single item in the whole catalog for a multi-entity
close** (catalog line 93) and the one genuine engineering problem in this wave: matching two
datasets on a key that might not be unique on either side is the case a naive
"loop A, loop B, compare" implementation gets wrong without ever raising an error — it just
pairs the wrong rows. Everything else in this wave is a smaller, more mechanical variation on
patterns `PkgClean` and `PkgImport` already established (a confirm-before-mutate dialog, a
Setup table resolved by column letter or header name, a destination sheet created only when
there's something to put on it).

---

## Decisions taken in this plan

| # | Question | Decision |
|---|---|---|
| 1 | Does `validate.reconcile` ask whether each range has a header? | **No — row 1 of each picked range is always its header, not asked.** Unlike `clean.dupes`, the report names *columns*, by header text, not just rows; a header-less dataset has nothing to align by. Stated as a fixed rule rather than a sixth prompt. |
| 2 | Which columns does `validate.reconcile` compare beyond the key? | **Every header common to both ranges, automatically — no config field.** The comparison set describes the *selection*, and contract §8's normative rule (established for `clean`) says a setting describing the selection must be a run-time input, not stored. A column present on only one side is not compared per-row; it is not silently ignored either — see decision 3's report shape. |
| 3 | What happens when a key is not unique on one or both sides? | **Reported once, as its own status (`Duplicate key`), with both sides' counts, and its rows are never compared column-by-column.** Pairing same-key rows across sides by position (first A-row with first B-row) is a guess the data does not support and is exactly the failure mode a nested-loop implementation produces silently. An implementation MUST NOT pair them. |
| 4 | Case sensitivity of a key match? | **Case-insensitive**, via `Scripting.Dictionary.CompareMode = vbTextCompare` set once before the first `.Add` (it cannot be changed after). Matches every other name comparison already in this codebase (`StrComp(ws.name, destName, vbTextCompare)` in `PkgImport`). |
| 5 | How are two matched values compared for "Differs"? | **Numeric-tolerant, text otherwise.** Two numeric values are compared as numbers (`Abs(a-b) < 1E-9`), so `"100"` and `"100.00"` agree; anything else is trimmed and compared case-insensitively. Without this, formatting noise would flood the report with false differences. |
| 6 | One shared destination field for `reconcile` and `scan`, or two? | **Two — `k_validate_recon_dest` and `k_validate_scan_dest`.** A chain that runs both (see "Chains" below) would otherwise have the second action refuse because the first one's sheet now exists, even though nothing is actually wrong. |
| 7 | Does `reconcile` / `scan` create its sheet unconditionally? | **No — only when there is at least one exception.** A perfect reconciliation, or a clean scan, leaves the workbook exactly as it was. This mirrors `clean.remove_blanks`'s "nothing found" branch, which also touches nothing. |
| 8 | Safety story for the three colour-writing actions | **Count-then-confirm (the `PkgClean_ConfirmBlanks` pattern) plus a never-overwrite rule: a cell that already carries its own fill colour is left alone and reported, never recoloured.** That second half is what turns "select the range, Home > Fill Color > No Fill" into a *guaranteed*-correct undo instead of a typically-correct one — every cell the action touches is provably blank-filled beforehand, because a cell that was not got skipped. Rejected: a new-sheet copy (defeats the point — the flag needs to sit where the user is already working) and a named Excel Style (reapplying "Normal" to remove it would also strip any of the user's *own* prior formatting on that cell, which is not always-correct the way this rule is). |
| 9 | How much of the picked range gets coloured? | **Only the cell(s) that caused the flag** — the key column(s) for a duplicate, the failing cell itself for an outlier or a missing field — never the whole row. Keeps the mutated footprint no bigger than the finding it reports. |
| 10 | One highlight colour, or one per action? | **One reserved colour per action** (documented RGB, precomputed as a `Long` because `RGB()` is not a VBA compile-time constant and so cannot be a `Const`), so a user can tell which rule flagged a given cell without opening a report sheet. |
| 11 | Default rows on `k_validate_scan_rules` / `k_validate_outliers`? | **None, deliberately — extends the `k_clean_formats` rule (contract §8, package `clean`) to two more fields, and the stakes are higher here.** `k_clean_formats` with a demo row risks an invisible number-format change; these two, with a demo row, risk *colouring real cells* on a deployment's very first run, before anyone has looked at Setup. Worked examples live in the field's help text instead. |
| 12 | Where do `validate.highlight_outliers`'s bounds come from — a Setup table, or a prompt? | **A Setup table (`k_validate_outliers`), same shape as `k_validate_scan_rules`.** A bound like "flag anything over $50,000" is a fact about the deployment — the same threshold applies every close — not about the selection, so contract §8's rule places it on Setup, not behind a new ad-hoc numeric prompt this plan would otherwise have to invent (no such helper exists in the Core API this plan is written against). |
| 13 | Does `PkgValidate` need its own column-letter-or-header resolver, or can it call `PkgClean`'s? | **Neither — it calls `Core_ResolveColumn`, promoted to Core by phase 2** (`2026-07-28-core-helper-promotion-plan.md`, Task 2 Step 2). Phase 2's own rule for a promotion candidate is "two consumers, or already duplicated" — `PkgClean_ResolveFormatCol` had exactly one consumer (`PkgClean` itself) until this wave's `validate.scan` and `validate.highlight_outliers` became its second, with W4's `enrich.*` anticipated as a third; phase 2 promoted it precisely because drafting this wave surfaced that second consumer. A second, package-private copy of that resolver here would be the forbidden duplicate phase 2's promotion exists to prevent, not the safe default. **One shape change to carry at every call site:** `Core_ResolveColumn` returns a `Long` column index counted from 1 at the picked range's left edge, not a `Range`, and 0 means not-found where the old helper's contract was `Is Nothing`. A caller wanting the column itself writes `rng.Columns(idx)`. |
| 14 | Does the composite-key builder (`PkgValidate_KeyOfArr`) get promoted to Core? | **No, stays package-private — but flagged.** It is used twice inside this package (`reconcile`, `highlight_duplicates`), which justifies factoring it out *within* the package, not promoting it. W4's `enrich.join` (not yet planned) will plausibly want the identical shape. Per this plan's brief: **escalate to Core in W4's own plan/diff if that turns out to be true, not here on a guess** — the same YAGNI stance phase 2 took on `PkgImport_WorkbookPaths`/`PkgImport_AlreadyOpen`. |
| 15 | How do these five actions behave as steps inside an automated chain? | **Every one of them pauses the chain for at least one interactive prompt — a range picker at minimum, up to four dialogs for `reconcile`.** Contract §6 is explicit that "interactive prompts (e.g. a range picker) and error dialogs are not suppressible" — only completion dialogs are. This is not a regression: `clean.dupes` already has the identical property today. The alternative — storing the range and key columns on Setup so a chain runs unattended — was tried for `clean` (`clean_keycols`, `clean_hasheader`) and explicitly withdrawn (contract §8) for describing the selection, not the deployment. This plan does not reopen that. |
| 16 | Read cell values via `.Cells(r,c).Value` in a loop, or bulk-read? | **Bulk-read via `Range.Value2` into a single in-memory 2-D array per range (or per resolved column), through a shared `PkgValidate_Grid` helper that normalises the one-cell scalar edge case.** A per-cell `.Cells().Value` read is a COM round trip; at "a real month-end dataset" scale (the brief's own words) that is the difference between a key index that builds in under a second and one that visibly grinds. The *write* side (colouring, writing an exceptions sheet) still touches real `Range` objects, because formatting cannot be set from an array — only the *read* side is bulk. |

---

## Global constraints

Copied from `2026-07-28-track2-completion-plan-of-record.md` §4, verbatim, plus §4.1's
do-not-revert list where it applies to this wave.

- **`SCHEMA_VER` stays `"v2.0"`.**
- **Action ids are immutable**, two-segment `<domain>.<operation>`, lowercase ASCII and
  underscores, each segment starting with a letter. Never camelCase, no alias table. This
  wave's five ids are frozen by the catalog (see the table in the task prompt) and used
  verbatim.
- **Shared helpers live ONLY in Core. Cross-package calls are forbidden.**
- **Pure ASCII in every `.bas`**; `Option Explicit`; max 25 line continuations per statement;
  no module-level declaration after the first procedure (every `Const`/`Dim` in
  `PkgValidate.bas` sits above `PkgValidate_Manifest`); no VBA reserved word as an identifier.
- **Every action entry point is `Public Sub Pkg<Name>_<Verb>(ctx As Object)`** — one `Object`
  parameter, nothing else. `build.ps1` fails the build otherwise.
- **Every package private is prefixed `PkgValidate_`.** The phase-2 build guard
  (`build/build.ps1`, added in that phase) fails a build on an unprefixed one; this plan does
  not rely on remembering the rule by hand.
- **TDD is off.** VBA has no headless runner. Do not propose a test suite. Verification is
  `pwsh build/build.ps1`, `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`, and
  deliberate injections.
- **No per-wave manual gate document.** Run-time-only claims are appended to
  `docs/designs/2026-07-28-deferred-validation-log.md` (Task 9), not written up as a checklist
  for the owner.
- **`dist/` is gitignored** — never `git add -f` the bundle. **`.vscode/` stays untracked.**
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.
- **`k_clean_formats` declares no default rows** (do-not-revert). This plan's two Setup tables
  follow the identical rule — see Decision 11.

---

## Verification, since TDD is off

`CLAUDE.md` turns off `test-driven-development`. Every task verifies with:

1. `pwsh build/build.ps1` — exit 0, `Core + 4 package(s)` (ImportExport, Import, Clean,
   Validate), balanced `Sub`/`Function` counts.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` — no diagnostics.
3. **A grep proving no field-key collision** (Task 2 Step 3) — every existing `k_*` key plus
   this wave's four, checked for uniqueness.
4. **A grep proving every `PkgValidate` private carries its prefix** (Task 2 Step 4), which the
   phase-2 build guard also enforces on every future build.
5. **A deliberate injection** (Task 8, Step 3) — temporarily mis-forming one action id as
   camelCase and confirming `build.ps1`'s existing id-format gate (`build/build.ps1:452`)
   catches it by name, then reverting. This is the one thing in this wave meant to fail, and
   it proves a brand-new package doesn't get a pass from a gate that has only ever seen the
   four ids that shipped before it.
6. **A deliberate injection** proving `Core_DestUsable` refuses `validate.reconcile` a second
   run against its own output (Task 3 Step 6) — the same shape as the reference `import.*`
   actions, confirmed statically per plan-of-record §3.5 item 4 rather than by the owner.

---

## File structure

| File | Change |
|---|---|
| `spec/contract.md` | v2.5.0. §8 gains "Package `validate`" (4 fields, 5 actions); §10 gains the version entry. §7 is **unchanged** — `validate` has been `active` since v2.3.0; this wave is the first to populate it (verify by grep, Task 1 Step 1). |
| `src/packages/PkgValidate.bas` | **New.** Manifest, 4 config fields, 5 actions, and their private engines — all privates prefixed `PkgValidate_`. |
| `src/core/Core.bas` | `CoreCandidates()` gains `c.Add "Validate"`. |
| `docs/designs/2026-07-28-deferred-validation-log.md` | Appended: this wave's run-time-only claims (Task 9). |
| `CHANGELOG.md`, `VERSION`, `docs/code-map.md` | v0.13.0. |

---

### Task 1: Contract v2.5.0

**Owner: Opus, written directly.** Establishes the normative text every later task in this
plan implements, and states the "colouring is reporting, not repair" argument the whole wave
stands on.

**Files:** Modify `spec/contract.md` — header (line 3), §8 (new subsection after package
`clean`, before "## 9. Platform-abstraction appendix"), §10.

**Interfaces:** Produces the field/action semantics Tasks 2-7 implement. Consumes nothing.

- [ ] **Step 1: Confirm `validate` is already `active` in §7**

```bash
grep -n '`validate`' spec/contract.md
```

Expected: one line, `| \`validate\` | active | Checking data against a rule and reporting what
fails; never repairs silently. |` (around line 593). **If this is missing, adding it is step
one and this task is not a pure §8/§10 addition** — but it should not be missing; this grep is
the proof, not a formality.

- [ ] **Step 2: Bump the version header**

Line 3: whatever phase 1/phase 2 left it at (`2.4.0`) becomes
`**Contract version: 2.5.0** (2026-07-28)`.

- [ ] **Step 3: Add "Package `validate`" to §8**

Insert immediately after the "Package `clean`" subsection ends (before "## 9. Platform-
abstraction appendix"):

````markdown
### Package `validate` — "Validate & Reconcile" (v1.0.0, contract v2.5.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_validate_recon_dest` | text | simple | `Reconcile Exceptions` | Name of the **new** sheet `validate.reconcile` writes its exceptions report to. Created only when the run finds at least one exception; a perfect match creates no sheet. `usedBy` `validate.reconcile`. |
| `k_validate_scan_dest` | text | simple | `Scan Exceptions` | Name of the **new** sheet `validate.scan` writes its exceptions report to. Created only when a configured rule finds at least one failing cell. `usedBy` `validate.scan`. |
| `k_validate_scan_rules` | table | simple | **none** | Data-quality rules for `validate.scan`. Three columns: **Column** (a column letter, or a header name matched against the first row of the range the caller picks — a one-to-three-letter entry is always read as a letter), **Rule** (`blank`, `number`, `date`, or `range`), **Bound** (read only for `range`, as `min,max`; either side may be blank for an open-ended bound). A row whose Column resolves to nothing inside the picked range MUST be reported by name, not silently skipped. `usedBy` `validate.scan`. |
| `k_validate_outliers` | table | simple | **none** | Numeric bounds for `validate.highlight_outliers`. Three columns: **Column** (letter or header name, same resolution rule as above), **Min**, **Max** (either may be blank for an open-ended bound). A row whose Column resolves to nothing inside the picked range MUST be reported by name. `usedBy` `validate.highlight_outliers`. |

`k_validate_scan_rules` header: `Column (letter or header name)|Rule (blank, number, date, range)|Bound (range rule only: min,max)`.
`k_validate_outliers` header: `Column (letter or header name)|Min (blank = no floor)|Max (blank = no ceiling)`.

**Both table fields declare NO default rows, deliberately — extending the `k_clean_formats`
rule (§8, package `clean`) to two more fields, for a sharper reason.** `k_clean_formats` with a
demo row risks an invisible number-format change; these two, with a demo row, risk *colouring
real cells* — `validate.scan` writes an exceptions sheet named from a rule nobody configured,
and `validate.highlight_outliers` colours cells against a bound nobody set, on a deployment's
very first run, before anyone has looked at Setup. A conforming implementation MUST NOT ship
default rows for either field. Worked examples live in each field's help text instead.

**`validate.reconcile` needs no config field beyond its destination name.** Its two ranges, and
the key column(s) within each, describe the SELECTION the caller is about to make — not the
deployment — so per §5's rule (established for `clean`, §8) they are run-time inputs, asked at
the moment the action runs, never stored on Setup.

**Actions**

| id | caption | semantics |
|---|---|---|
| `validate.reconcile` | Reconcile Two Datasets | Prompt for two ranges (Dataset A, Dataset B); row 1 of each is always its header, not asked. Prompt for the key column(s) of each, numbered from the LEFT of that range. Build a key index per side (composite key -> the row(s) carrying it); a key present on only one side is reported `In A only` / `In B only`. A key whose count is greater than one on EITHER side MUST be reported as `Duplicate key`, with both sides' counts, and its rows MUST NOT be compared column-by-column — an implementation MUST NOT pair same-key rows across sides by position or any other guess. A key matched exactly once on each side is compared across every header common to both ranges (excluding each side's own key column(s)); a numeric comparison MUST tolerate formatting differences (e.g. `100` vs `100.00`); a differing column is reported as `Differs`, one row per differing column. Written to a new sheet named `k_validate_recon_dest`, only when at least one exception exists. Neither input range is changed. |
| `validate.scan` | Scan for Data Issues | Prompt for a range and whether its first row is a header. Apply every `k_validate_scan_rules` row to its resolved column; a cell failing its rule is written to a new sheet named `k_validate_scan_dest`, only when at least one exception exists. A rule row whose Column resolves to nothing in the picked range MUST be reported by name. The picked range is not changed. |
| `validate.highlight_duplicates` | Highlight Duplicates | Prompt for a range, whether its first row is a header, and the key column(s), numbered from the LEFT of that range. Colour every key-column cell belonging to a row whose key is shared with at least one other row in the range. A cell that already carries a fill colour of its own MUST be left alone and counted separately, never overwritten. Before colouring anything, show a count of exactly what will be marked and require an explicit confirmation; declining or cancelling MUST change nothing. No value, formula, or existing formatting is changed. |
| `validate.highlight_outliers` | Highlight Outliers | Prompt for a range and whether its first row is a header. For every `k_validate_outliers` row, colour every numeric, non-blank cell in its resolved column that falls outside `Min`/`Max`. Same already-filled-cell carve-out, same confirm-before-colouring gate, same "nothing else changes" guarantee as `validate.highlight_duplicates`. A configured column that resolves to nothing in the picked range MUST be reported by name. |
| `validate.missing_fields` | Flag Missing Required Fields | Prompt for a range, whether its first row is a header, and which column(s), numbered from the LEFT of that range, are mandatory. Colour every blank cell in those columns. Same already-filled-cell carve-out, same confirm-before-colouring gate, same "nothing else changes" guarantee as the other two highlight actions. |

**Colouring is reporting, not repair.** §7 classes `validate` as "checking data against a rule
and reporting what fails; never repairs silently." The three actions above that write a fill
colour change no value, no formula, and no shape of the data — they mark, in place, exactly
what a written exceptions list marks on a separate sheet, using the sheet the data already
lives on as the report surface instead of a new one. A conforming implementation MUST make
this reversible without disturbing any other formatting: a cell that already carried a fill
colour of its own MUST NOT be recoloured, and a cell this touches is guaranteed to have carried
no fill immediately beforehand — so "clear the fill on the affected cells" always, and only,
undoes what the action did. **No action in this package may delete, move, or overwrite a value
or formula.**
````

- [ ] **Step 4: Add the §10 entry**

Insert as the first versioned bullet:

````markdown
- **v2.5.0 (additive):** a new package `validate` with four fields and five actions (§8):
  `validate.reconcile`, `validate.scan`, `validate.highlight_duplicates`,
  `validate.highlight_outliers`, `validate.missing_fields`. The `validate` domain has been
  listed `active` in §7's vocabulary table since v2.3.0, before any package implemented it;
  this version is the first to populate it. **Additive by the two-axes rule:** every existing
  stored value keeps its meaning, no existing `key`, `id`, `kind` or encoding changes.
  `SCHEMA_VER` does **not** move — it stays `"v2.0"`.
````

- [ ] **Step 5: Verify**

```bash
grep -n 'Contract version: 2.5.0' spec/contract.md
grep -c 'k_validate_' spec/contract.md
pwsh build/build.ps1
```

Expected: exit 0. A contract edit that breaks §7's table format fails the build immediately
(`build/build.ps1:227` onward parses it on every build) — this is what makes Step 1's grep
meaningful rather than a courtesy.

- [ ] **Step 6: Commit**

```bash
git add spec/contract.md
git commit -m "spec: contract v2.5.0 - package validate (reconcile, scan, and three highlight actions)"
```

---

### Task 2: `PkgValidate.bas` — skeleton, manifest, and the shared engine helpers

**Owner: Opus, written directly.** Defines the field/action schema surface (must match Task 1
exactly) and every helper the five action tasks depend on.

**Files:** Create `src/packages/PkgValidate.bas`.

**Interfaces:**
- Consumes: `Core_DeclarePackage`, `Core_AddSection`, `Core_AddField`, `Core_AddTable`,
  `Core_AddAction`, `Core_IsAlpha`, `Core_IsBlank` (all Core, existing/phase 2).
- Produces: `PkgValidate_Manifest`; the private helpers `PkgValidate_Grid`,
  `PkgValidate_ConfirmColor`, `PkgValidate_MarkCell`,
  `PkgValidate_ClassifyMark`, `PkgValidate_ColorDuplicate`, `PkgValidate_ColorOutlier`,
  `PkgValidate_ColorMissing`. Tasks 3-7 depend on all of them.

- [ ] **Step 1: Module header and manifest**

```vb
Attribute VB_Name = "PkgValidate"
Option Explicit

'================================================================
' PkgValidate - Wave 2 of the catalog build-out (docs/macro-catalog-
' plug-and-play.md section 3; docs/designs/2026-07-28-w2-validate-plan.md).
'
' Five actions. None of them repairs anything - contract section 7's
' domain table says so plainly: "Checking data against a rule and
' reporting what fails; never repairs silently." Two of the five write
' a cell colour; that is reporting IN PLACE rather than to a separate
' sheet, not a repair - see PkgValidate_MarkCell for the rule that
' keeps a colour call honest (never overwrite a cell's own fill, so
' clearing the fill afterward is always a correct undo).
'
'   validate.reconcile             match two ranges on a key, report
'                                   missing / extra / differing rows
'   validate.scan                  sweep a range against configured
'                                   rules, report exceptions to a new
'                                   sheet
'   validate.highlight_duplicates  colour-flag rows sharing a key
'   validate.highlight_outliers    colour-flag values outside a bound
'   validate.missing_fields        colour-flag rows missing a
'                                   mandatory column
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgValidate_Manifest(reg As Object)
    Core_DeclarePackage reg, "validate", "Validate & Reconcile", "1.0.0"

    Core_AddSection reg, "RECONCILE"
    Core_AddField reg, "k_validate_recon_dest", "Name for the reconciliation report sheet", "text", "Reconcile Exceptions", _
        "simple", _
        "A NEW sheet by this name is created only when there is something to report. If a sheet by this name already exists, the action stops and asks you to rename or remove it - it never overwrites one.", _
        "validate.reconcile"

    Core_AddSection reg, "SCAN"
    Core_AddField reg, "k_validate_scan_dest", "Name for the scan exceptions sheet", "text", "Scan Exceptions", _
        "simple", _
        "A NEW sheet by this name is created only when a rule below finds something. If a sheet by this name already exists, the action stops and asks you to rename or remove it.", _
        "validate.scan"
    Core_AddTable reg, "k_validate_scan_rules", "Data rules", _
        "Column (letter or header name)|Rule (blank, number, date, range)|Bound (range rule only: min,max)", "simple", _
        "One row per rule. Left cell: a column letter like C, or the column's header text - a one-to-three-letter entry is always read as a letter, so a column headed exactly A or ID must be addressed by its letter. Middle cell: blank (flag empty cells), number (flag anything not a number), date (flag anything not a date), or range (flag a number outside the Bound). Bound is only read for range, as two numbers separated by a comma - leave one side blank for an open-ended bound, e.g. ',100000' flags anything over 100000. Nothing happens until you fill this in.", _
        "validate.scan"

    Core_AddSection reg, "OUTLIERS"
    Core_AddTable reg, "k_validate_outliers", "Outlier bounds", _
        "Column (letter or header name)|Min (blank = no floor)|Max (blank = no ceiling)", "simple", _
        "One row per column to check. Left cell: a column letter or header name, same rule as above. A value below Min or above Max is flagged; leave either side blank to leave that side unchecked. Nothing happens until you fill this in.", _
        "validate.highlight_outliers"

    Core_AddAction reg, "validate.reconcile", "Reconcile Two Datasets", "PkgValidate_Reconcile", _
        "Pick two ranges and the key column(s) in each; reports rows found only in one side, rows sharing a key that is not unique on either side, and rows whose other columns differ - to a new sheet. Nothing in either range is changed.", _
        "Validate / Reconcile"
    Core_AddAction reg, "validate.scan", "Scan for Data Issues", "PkgValidate_Scan", _
        "Pick a range; checks it against the Data rules table on Setup and writes every row that fails a rule to a new sheet. Nothing in the picked range is changed.", _
        "Validate / Reconcile"
    Core_AddAction reg, "validate.highlight_duplicates", "Highlight Duplicates", "PkgValidate_HighlightDuplicates", _
        "Pick a range and the column(s) that form a key; colour-flags the key cell(s) of every row sharing a key with another row. Nothing is deleted or moved.", _
        "Validate / Reconcile"
    Core_AddAction reg, "validate.highlight_outliers", "Highlight Outliers", "PkgValidate_HighlightOutliers", _
        "Pick a range; colour-flags any cell in a configured column that falls outside its bound from the Outlier bounds table on Setup. Nothing is deleted or moved.", _
        "Validate / Reconcile"
    Core_AddAction reg, "validate.missing_fields", "Flag Missing Required Fields", "PkgValidate_MissingFields", _
        "Pick a range and the column(s) that are mandatory; colour-flags every blank cell in those columns. Nothing is deleted or moved.", _
        "Validate / Reconcile"
End Sub
```

- [ ] **Step 2: The shared read helper**

```vb
'================================================================
' Shared helpers. Prefixed PkgValidate_ because the bundle is one flat
' namespace (CLAUDE.md rule 4), even though these are Private.
'================================================================

' Range.Value2, but normalised to always be a 2-D array. Value2 on a
' range that holds exactly ONE cell returns a scalar, not a 1x1 array -
' every array-indexed read in this module assumes 2-D, so that edge
' case is handled once, here, instead of guarded at every call site.
'
' Every engine in this module reads through THIS, not through
' rng.Cells(r, c).Value in a loop. A per-cell .Value read is a COM
' round trip; at "a real month-end dataset" scale that is the
' difference between a key index that builds in under a second and one
' that visibly grinds. Only the WRITE side (colouring cells, writing an
' exceptions sheet) still touches real Range objects, because
' formatting cannot be set from an array.
Private Function PkgValidate_Grid(rng As Range) As Variant
    If rng.Cells.Count = 1 Then
        Dim g(1 To 1, 1 To 1) As Variant
        g(1, 1) = rng.Value2
        PkgValidate_Grid = g
    Else
        PkgValidate_Grid = rng.Value2
    End If
End Function

' Column resolution (letter-or-header, against a picked range) is not
' defined in this package. It is Core_ResolveColumn, promoted to Core
' by phase 2 (2026-07-28-core-helper-promotion-plan.md, Task 2 Step 2)
' precisely because this wave makes it a second, independent consumer
' of the same shape PkgClean_StandardizeFormats already had - see
' "Decisions taken", W2 plan, decision 13. It returns a Long column
' index counted from 1 at rng's left edge (0 when the spec does not
' resolve inside rng), not a Range; every call site in this module
' converts with rng.Columns(idx) where it needs the Range back.
```

- [ ] **Step 3: The shared colour helpers**

```vb
' RGB(...) is not a compile-time constant in VBA, so these are computed
' Functions rather than Const - each value is precomputed and commented
' with the RGB() call it equals, so it can be checked by eye. Chosen to
' be visually distinct from each other and to echo Excel's own built-in
' conditional-formatting palette, so a user recognises them as "a rule
' flagged this" rather than "someone coloured this by hand".
Private Function PkgValidate_ColorDuplicate() As Long
    PkgValidate_ColorDuplicate = 13551615         ' RGB(255, 199, 206) - Excel's own light red
End Function

Private Function PkgValidate_ColorOutlier() As Long
    PkgValidate_ColorOutlier = 10284031           ' RGB(255, 235, 156) - Excel's own light amber
End Function

Private Function PkgValidate_ColorMissing() As Long
    PkgValidate_ColorMissing = 16761036           ' RGB(204, 192, 255) - light violet, distinct from the other two
End Function

' Colours ONE cell, but only if it carries no fill of its own already -
' ColorIndex = xlColorIndexNone is Excel's "nothing set" state. A cell
' that already has a fill is left untouched and counted in skipped,
' never overwritten.
'
' This is what makes "select the range, Home > Fill Color > No Fill" a
' SAFE undo instruction rather than a guess: every cell this marks is
' guaranteed to have had No Fill immediately before, because a cell
' that did not was never touched. Restoring No Fill can therefore never
' destroy formatting the user set themselves.
Private Sub PkgValidate_MarkCell(c As Range, ByVal clr As Long, _
                                 ByRef marked As Long, ByRef skipped As Long)
    If c.Interior.ColorIndex = xlColorIndexNone Then
        c.Interior.Color = clr
        marked = marked + 1
    Else
        skipped = skipped + 1
    End If
End Sub

' How many of a candidate cell Collection WOULD be marked vs skipped,
' without touching anything - used to give the confirm dialog an exact
' count before the user commits, using the same test PkgValidate_MarkCell
' applies when it actually runs.
Private Sub PkgValidate_ClassifyMark(cells As Collection, ByRef willMark As Long, ByRef willSkip As Long)
    Dim c As Range
    willMark = 0: willSkip = 0
    For Each c In cells
        If c.Interior.ColorIndex = xlColorIndexNone Then
            willMark = willMark + 1
        Else
            willSkip = willSkip + 1
        End If
    Next c
End Sub

' Shared confirm-before-colouring dialog for the three highlight-style
' actions. vbYesNo, not vbYesNoCancel: Windows disables the title-bar X
' on a message box with no Cancel button, so there is no accidental
' dismissal that could be mistaken for consent - the same reasoning
' PkgClean_ConfirmBlanks uses for a delete. A colour is not a delete,
' but it is just as unreachable by Ctrl+Z once applied through VBA, so
' the same bar applies. Defaults to No.
Private Function PkgValidate_ConfirmColor(ByVal what As String, ByVal willMark As Long, _
                                          ByVal willSkip As Long, ByVal rngRef As String, _
                                          ByVal title As String) As Boolean
    Dim msg As String
    msg = "Inside " & rngRef & " this will colour-flag " & willMark & " cell(s) " & what & "." & vbCrLf & vbCrLf
    If willSkip > 0 Then
        msg = msg & willSkip & " more cell(s) matched but already carry a fill colour of their " & _
              "own, and are left alone." & vbCrLf & vbCrLf
    End If
    msg = msg & "No value, formula or existing formatting is changed - only cells with no fill " & _
          "colour today receive one." & vbCrLf & vbCrLf & _
          "This cannot be undone with Ctrl+Z. To remove it afterwards: select the cells and use " & _
          "Home > Fill Color > No Fill - every cell this colours is guaranteed to have had no " & _
          "fill before, so that always reverses it cleanly." & vbCrLf & vbCrLf & "Continue?"
    PkgValidate_ConfirmColor = (MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, title) = vbYes)
End Function
```

- [ ] **Step 4: Build, syntax-check, and prove the key-uniqueness and prefix rules**

```bash
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -nE 'Core_Add(Field|List|Table) reg, "k_' src/packages/*.bas | sed -E 's/.*"(k_[a-z_]+)".*/\1/' | sort | uniq -d
grep -nE '^\s*Private\s+(Sub|Function)\s+' src/packages/PkgValidate.bas | grep -vE 'Private\s+(Sub|Function)\s+PkgValidate_'
```

Expected: exit 0 on both build steps (the module has no actions yet, so `CoreCandidates`
still does not probe it — Task 8 wires that up, and it is fine for `PkgValidate.bas` to sit
unregistered for now, deliberately, so this task's build proves only that it compiles). The
`uniq -d` line prints **nothing** — that is the proof no `k_*` key collides with an existing
one. The prefix grep prints **nothing** — every private declared so far starts `PkgValidate_`.

- [ ] **Step 5: Commit**

```bash
git add src/packages/PkgValidate.bas
git commit -m "feat(validate): PkgValidate skeleton, manifest, and the shared read/colour helpers"
```

---

### Task 3: `validate.reconcile`

**Owner: Opus, written directly.** The flagship, and the one item in this wave with real
engineering risk — the non-unique-key case is exactly what a naive nested-loop implementation
gets silently wrong.

**Files:** Modify `src/packages/PkgValidate.bas` — append after Task 2's helpers.

**Interfaces:**
- Consumes: `Core_Cfg`, `Core_DestUsable`, `Core_AskRange`, `Core_NarrowToUsed`, `Core_AskCols`,
  `Core_ParseCols`, `Core_RangeRef`, `Core_NewSheet`, `Core_SpeedOn`/`Off`, `Core_Refuse`,
  `Core_Done`, `Core_Fail`, `PkgValidate_Grid` (Task 2).
- Produces: `PkgValidate_Reconcile` (the action), and its private engine.

- [ ] **Step 1: The action entry point**

```vb
'================================================================
' ACTION - Reconcile two datasets on a key
'================================================================
Public Sub PkgValidate_Reconcile(ctx As Object)
    Dim destName As String
    Dim rngA As Range, rngB As Range
    Dim keysA As String, keysB As String
    Dim colsA As Variant, colsB As Variant
    Const TITLE As String = "Reconcile Two Datasets"
    On Error GoTo Fail

    ' Cheapest failure first, before any prompt - same reasoning as
    ' PkgImport_CombineWorksheets: a refusal after the user has already
    ' answered four questions reads as a bug, not a precondition.
    destName = Trim(Core_Cfg(ctx, "k_validate_recon_dest"))
    If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub

    Set rngA = Core_AskRange("Dataset A (the primary side)")
    If rngA Is Nothing Then Exit Sub
    Set rngA = Core_NarrowToUsed(rngA)
    If rngA Is Nothing Then
        Core_Refuse ctx, "The range you picked for Dataset A is entirely outside the used " & _
               "area of its sheet, so there is nothing in it to reconcile." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If rngA.Rows.Count < 2 Then
        ' Row 1 is ALWAYS the header here, not asked (Decision 1, W2 plan) - a
        ' one-row pick has a header and no data underneath it.
        Core_Refuse ctx, "Dataset A (" & Core_RangeRef(rngA) & ") is only a header row - there " & _
               "is no data underneath it to reconcile." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rngB = Core_AskRange("Dataset B (the comparison side)")
    If rngB Is Nothing Then Exit Sub
    Set rngB = Core_NarrowToUsed(rngB)
    If rngB Is Nothing Then
        Core_Refuse ctx, "The range you picked for Dataset B is entirely outside the used " & _
               "area of its sheet, so there is nothing in it to reconcile." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If rngB.Rows.Count < 2 Then
        Core_Refuse ctx, "Dataset B (" & Core_RangeRef(rngB) & ") is only a header row - there " & _
               "is no data underneath it to reconcile." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(rngA, keysA, _
            "Which column(s) of Dataset A hold the key to match rows on?", TITLE) Then Exit Sub
    If Not Core_AskCols(rngB, keysB, _
            "Which column(s) of Dataset B hold the SAME key?", TITLE) Then Exit Sub
    colsA = Core_ParseCols(keysA)
    colsB = Core_ParseCols(keysB)

    Dim outRows As Collection
    Core_SpeedOn
    PkgValidate_ReconcileEngine rngA, rngB, colsA, colsB, outRows
    Core_SpeedOff

    If outRows.Count = 0 Then
        Core_Done ctx, "Every row matched: no rows found on only one side, no duplicate keys, " & _
               "and no differing columns." & vbCrLf & vbCrLf & "No report sheet was created - " & _
               "there was nothing to report.", TITLE
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Core_NewSheet(destName)
    PkgValidate_WriteReconcileReport ws, outRows

    Core_Done ctx, "Wrote " & outRows.Count & " exception row(s) to '" & destName & "'." & _
           vbCrLf & vbCrLf & "Status column meanings:" & vbCrLf & _
           "   In A only      - key exists in Dataset A, not in Dataset B" & vbCrLf & _
           "   In B only      - key exists in Dataset B, not in Dataset A" & vbCrLf & _
           "   Duplicate key  - the key is not unique on one or both sides, so those rows " & _
           "were NOT compared row-to-row" & vbCrLf & _
           "   Differs        - the key matched exactly once on each side, but a column " & _
           "disagrees", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.reconcile", Err
End Sub
```

- [ ] **Step 2: The engine — dictionary index, not nested loops**

```vb
' The reconcile engine. Builds ONE Scripting.Dictionary index per side,
' keyed by the composite key (PkgValidate_KeyOfArr), NOT nested loops -
' a "for each row in A, for each row in B" pass is O(rows(A) x
' rows(B)), the shape that looks fine on a forty-row demo and then
' visibly freezes on a real month-end extract with thousands of rows on
' each side. Building both indexes once and then walking them is
' O(rows(A) + rows(B)).
'
' CompareMode = vbTextCompare is set on each Dictionary BEFORE the
' first .Add - Scripting.Dictionary raises if it is changed afterward -
' so key matching is case-insensitive the same way every other name
' comparison in this codebase is (e.g. PkgImport's sheet-name check
' uses StrComp ... vbTextCompare). Without it, "ABC123" and "abc123"
' would reconcile as two different accounts.
'
' A key that is NOT unique on a side is the case a naive implementation
' gets silently wrong: pairing the first A-row with the first B-row by
' POSITION looks plausible and is not what the data says. Nothing here
' pairs same-key rows by position. Any key whose count is greater than
' one on EITHER side is reported once, as "Duplicate key", with both
' counts, and none of its rows enter the row-by-row Differs comparison -
' there is no principled way to decide which A-row corresponds to which
' B-row from the key alone, so the honest answer is "not compared", not
' a guess. (A key absent from one side entirely still reports each of
' its occurrences as In A/B only, even if it is ALSO duplicated on its
' own side - that duplication does not matter to "does the other side
' have this at all".)
Private Sub PkgValidate_ReconcileEngine(rngA As Range, rngB As Range, colsA As Variant, colsB As Variant, _
                                        ByRef outRows As Collection)
    Dim valsA As Variant, valsB As Variant
    valsA = PkgValidate_Grid(rngA)
    valsB = PkgValidate_Grid(rngB)

    Dim dictA As Object, dictB As Object
    Dim r As Long, k As Variant, cA As Collection, cB As Collection
    Set outRows = New Collection
    Set dictA = CreateObject("Scripting.Dictionary")
    Set dictB = CreateObject("Scripting.Dictionary")
    dictA.CompareMode = vbTextCompare
    dictB.CompareMode = vbTextCompare

    For r = 2 To UBound(valsA, 1)
        k = PkgValidate_KeyOfArr(valsA, r, colsA)
        If Not dictA.Exists(k) Then dictA.Add k, New Collection
        dictA(k).Add r
    Next r
    For r = 2 To UBound(valsB, 1)
        k = PkgValidate_KeyOfArr(valsB, r, colsB)
        If Not dictB.Exists(k) Then dictB.Add k, New Collection
        dictB(k).Add r
    Next r

    Dim commonHdrs As Collection
    Set commonHdrs = PkgValidate_CommonHeadersArr(valsA, valsB, colsA, colsB)

    For Each k In dictA.Keys
        Set cA = dictA(k)
        If Not dictB.Exists(k) Then
            PkgValidate_AddInAOnly outRows, valsA, cA, k
        Else
            Set cB = dictB(k)
            If cA.Count > 1 Or cB.Count > 1 Then
                outRows.Add Array("Duplicate key", k, "", cA.Count & " row(s) in A", cB.Count & " row(s) in B")
            Else
                PkgValidate_CompareRow outRows, valsA, CLng(cA(1)), valsB, CLng(cB(1)), commonHdrs, k
            End If
        End If
    Next k
    For Each k In dictB.Keys
        If Not dictA.Exists(k) Then
            Set cB = dictB(k)
            PkgValidate_AddInBOnly outRows, valsB, cB, k
        End If
    Next k
End Sub

' Builds a composite key from one or more columns of ONE row, numbered
' from the left of the range exactly like Core_AskCols/Core_ParseCols
' number them. Chr(31) (the ASCII "unit separator") joins the parts
' because it is not a character a person can type into a cell by hand,
' so a key made of "A" + "BC" can never collide with one made of "AB" +
' "C" the way joining on a comma or pipe could.
Private Function PkgValidate_KeyOfArr(vals As Variant, ByVal r As Long, keyCols As Variant) As String
    Dim i As Long, parts As String
    For i = LBound(keyCols) To UBound(keyCols)
        parts = parts & Chr$(31) & Trim$(CStr(vals(r, keyCols(i))))
    Next i
    PkgValidate_KeyOfArr = Mid$(parts, 2)     ' drop the leading separator
End Function

' Headers shared by both ranges, EXCLUDING each side's own key
' column(s) - those are already equal by construction (they are what
' matched the rows), so comparing them again would just repeat the key
' back at the user as a false "difference" whenever the two sides
' format it differently (e.g. "007" vs "7").
Private Function PkgValidate_CommonHeadersArr(valsA As Variant, valsB As Variant, colsA As Variant, colsB As Variant) As Collection
    Dim mapA As Object, mapB As Object, c As Long, hdr As String
    Dim out As New Collection, k As Variant
    Set mapA = CreateObject("Scripting.Dictionary"): mapA.CompareMode = vbTextCompare
    Set mapB = CreateObject("Scripting.Dictionary"): mapB.CompareMode = vbTextCompare
    For c = 1 To UBound(valsA, 2)
        If Not PkgValidate_InArray(c, colsA) Then
            hdr = Trim$(CStr(valsA(1, c)))
            If Len(hdr) > 0 Then If Not mapA.Exists(hdr) Then mapA.Add hdr, c
        End If
    Next c
    For c = 1 To UBound(valsB, 2)
        If Not PkgValidate_InArray(c, colsB) Then
            hdr = Trim$(CStr(valsB(1, c)))
            If Len(hdr) > 0 Then If Not mapB.Exists(hdr) Then mapB.Add hdr, c
        End If
    Next c
    For Each k In mapA.Keys
        If mapB.Exists(k) Then out.Add Array(CStr(k), CLng(mapA(k)), CLng(mapB(k)))
    Next k
    Set PkgValidate_CommonHeadersArr = out
End Function

Private Function PkgValidate_InArray(ByVal n As Long, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = n Then PkgValidate_InArray = True: Exit Function
    Next i
End Function

' Compares one matched pair of rows across every header common to both
' sides, using PkgValidate_ValuesEqual so a formatting difference like
' "100" vs "100.00" does not read as a disagreement. One output row per
' column that actually differs - a pair with no differing column adds
' nothing, which is how a full match stays silent rather than
' cluttering the report with an all-clear line for every one of
' possibly thousands of matched rows.
Private Sub PkgValidate_CompareRow(outRows As Collection, valsA As Variant, ByVal rA As Long, _
                                   valsB As Variant, ByVal rB As Long, commonHdrs As Collection, ByVal k As Variant)
    Dim pair As Variant, vA As Variant, vB As Variant
    For Each pair In commonHdrs
        vA = valsA(rA, pair(1))
        vB = valsB(rB, pair(2))
        If Not PkgValidate_ValuesEqual(vA, vB) Then
            outRows.Add Array("Differs", k, CStr(pair(0)), Trim$(CStr(vA)), Trim$(CStr(vB)))
        End If
    Next pair
End Sub

' Numeric-tolerant, case-insensitive equality for one cell pair. Two
' numbers are compared as numbers (so "100" and "100.00" agree, and so
' does ordinary floating-point noise past the ninth decimal); anything
' else is compared as trimmed text, case-insensitively, matching every
' other text comparison in this codebase.
Private Function PkgValidate_ValuesEqual(ByVal vA As Variant, ByVal vB As Variant) As Boolean
    If IsNumeric(vA) And IsNumeric(vB) Then
        PkgValidate_ValuesEqual = (Abs(CDbl(vA) - CDbl(vB)) < 0.000000001)
    Else
        PkgValidate_ValuesEqual = (StrComp(Trim$(CStr(vA)), Trim$(CStr(vB)), vbTextCompare) = 0)
    End If
End Function

Private Sub PkgValidate_AddInAOnly(outRows As Collection, valsA As Variant, cA As Collection, ByVal k As Variant)
    Dim i As Long, r As Long
    For i = 1 To cA.Count
        r = cA(i)
        outRows.Add Array("In A only", k, "", PkgValidate_RowTextArr(valsA, r), "")
    Next i
End Sub

Private Sub PkgValidate_AddInBOnly(outRows As Collection, valsB As Variant, cB As Collection, ByVal k As Variant)
    Dim i As Long, r As Long
    For i = 1 To cB.Count
        r = cB(i)
        outRows.Add Array("In B only", k, "", "", PkgValidate_RowTextArr(valsB, r))
    Next i
End Sub

' All of one row's values, comma-joined, so an "In A only" / "In B
' only" line shows the whole record rather than just the key that
' matched - the report stays a fixed 5 columns regardless of how wide
' either source range was, by folding the extra columns into one cell
' rather than widening the sheet per source.
Private Function PkgValidate_RowTextArr(vals As Variant, ByVal r As Long) As String
    Dim c As Long, out As String
    For c = 1 To UBound(vals, 2)
        If c > 1 Then out = out & ", "
        out = out & Trim$(CStr(vals(r, c)))
    Next c
    PkgValidate_RowTextArr = out
End Function

' Writes the exceptions Collection built above to a fresh sheet, one
' row per exception. This side stays plain Range.Cells writes, not a
' bulk array write like the read side - the exceptions count is bounded
' by how much is actually wrong, not by the size of either source
' range, so it is never the hot path Decision 16 is about.
Private Sub PkgValidate_WriteReconcileReport(ws As Worksheet, outRows As Collection)
    Dim i As Long, item As Variant
    ws.Range("A1:E1").Value = Array("Status", "Key", "Column", "Value in A", "Value in B")
    ws.Range("A1:E1").Font.Bold = True
    i = 2
    For Each item In outRows
        ws.Cells(i, 1).Value = item(0)
        ws.Cells(i, 2).Value = item(1)
        ws.Cells(i, 3).Value = item(2)
        ws.Cells(i, 4).Value = item(3)
        ws.Cells(i, 5).Value = item(4)
        i = i + 1
    Next item
    ws.Columns("A:E").AutoFit
End Sub
```

- [ ] **Step 3: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics.

- [ ] **Step 4: Commit**

```bash
git add src/packages/PkgValidate.bas
git commit -m "feat(validate): validate.reconcile - dictionary-indexed reconciliation with a stated non-unique-key rule"
```

---

### Task 4: `validate.scan`

**Owner: delegate to codex (`sonnet`).** Follows `Core_ResolveColumn` and
`PkgClean_StandardizeFormats`'s "missed row" pattern closely; the VBA below is complete.

**Files:** Modify `src/packages/PkgValidate.bas` — append after Task 3.

**Interfaces:** Consumes `Core_Cfg`, `Core_DestUsable`, `Core_CfgTable`, `Core_AskRange`,
`Core_NarrowToUsed`, `Core_AskHasHeader`, `Core_RangeRef`, `Core_NewSheet`, `Core_IsBlank`,
`Core_SpeedOn`/`Off`, `Core_Refuse`, `Core_Done`, `Core_Fail`, `Core_ResolveColumn`,
`PkgValidate_Grid` (Task 2). Produces `PkgValidate_Scan` and its engine.

- [ ] **Step 1: The action and its engine**

```vb
'================================================================
' ACTION - Scan a range against configured rules
'================================================================
Public Sub PkgValidate_Scan(ctx As Object)
    Dim destName As String, rng As Range, hasHeader As Boolean
    Dim cols() As String, rules() As String, bounds() As String
    Dim outRows As Collection, missed As String
    Const TITLE As String = "Scan for Data Issues"
    On Error GoTo Fail

    destName = Trim(Core_Cfg(ctx, "k_validate_scan_dest"))
    If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub

    If Not Core_CfgTable(ctx, "k_validate_scan_rules", cols, rules, bounds) Then
        Core_Refuse ctx, "No data rules are set up yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Data rules' table on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to scan." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub

    Core_SpeedOn
    PkgValidate_ScanEngine rng, hasHeader, cols, rules, bounds, outRows, missed
    Core_SpeedOff

    If outRows.Count = 0 Then
        Core_Done ctx, "No exceptions found inside " & Core_RangeRef(rng) & "." & vbCrLf & vbCrLf & _
               "No report sheet was created - there was nothing to report." & _
               IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These configured rule(s) matched no " & _
                   "column in your selection and were skipped:" & missed, ""), TITLE
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Core_NewSheet(destName)
    PkgValidate_WriteScanReport ws, outRows

    Core_Done ctx, "Wrote " & outRows.Count & " exception(s) to '" & destName & "'." & _
           IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These configured rule(s) matched no column " & _
               "in your selection and were skipped:" & missed, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.scan", Err
End Sub

' One pass per configured rule row, not one pass per cell per rule -
' the column is resolved ONCE (Core_ResolveColumn) and its values
' read ONCE in bulk (PkgValidate_Grid), so N rules over an M-column
' range costs N bulk reads of one column each, not N x (rows x M)
' individual cell reads.
Private Sub PkgValidate_ScanEngine(rng As Range, ByVal hasHeader As Boolean, cols() As String, _
                                   rules() As String, bounds() As String, ByRef outRows As Collection, _
                                   ByRef missed As String)
    Dim i As Long, r As Long, firstRow As Long, idx As Long, tgt As Range, vals As Variant, v As Variant
    Dim ruleName As String, lo As Double, hi As Double, hasLo As Boolean, hasHi As Boolean
    Dim reason As String, hdr As String
    Set outRows = New Collection
    firstRow = IIf(hasHeader, 2, 1)

    For i = LBound(cols) To UBound(cols)
        idx = Core_ResolveColumn(rng, cols(i))
        If idx = 0 Then
            missed = missed & vbCrLf & "   " & cols(i)
        Else
            Set tgt = rng.Columns(idx)
            vals = PkgValidate_Grid(tgt)
            hdr = Trim$(CStr(vals(1, 1)))
            ruleName = LCase$(Trim$(rules(i)))
            PkgValidate_ParseBound bounds(i), lo, hasLo, hi, hasHi
            For r = firstRow To UBound(vals, 1)
                v = vals(r, 1)
                reason = PkgValidate_RuleFails(ruleName, v, lo, hasLo, hi, hasHi)
                If Len(reason) > 0 Then
                    outRows.Add Array(Core_RangeRef(tgt.Cells(r, 1)), hdr, ruleName, Trim$(CStr(v)), reason)
                End If
            Next r
        End If
    Next i
End Sub

' "min,max" -> two optional Doubles. Either side blank means unbounded
' on that side. A non-numeric side is treated the same as blank
' (unbounded) rather than raising - a typo'd bound should not crash the
' scan, it should just not constrain that side.
Private Sub PkgValidate_ParseBound(ByVal spec As String, ByRef lo As Double, ByRef hasLo As Boolean, _
                                   ByRef hi As Double, ByRef hasHi As Boolean)
    Dim parts() As String
    hasLo = False: hasHi = False: lo = 0: hi = 0
    parts = Split(spec, ",")
    If UBound(parts) >= 0 Then
        If IsNumeric(Trim$(parts(0))) Then lo = CDbl(Trim$(parts(0))): hasLo = True
    End If
    If UBound(parts) >= 1 Then
        If IsNumeric(Trim$(parts(1))) Then hi = CDbl(Trim$(parts(1))): hasHi = True
    End If
End Sub

' "" means the cell passes. Any non-empty string is the reason it did
' not, and becomes the report's Reason column verbatim. A blank cell
' passes "number" and "date" - a blank is not "not a number", it is
' just blank, and flagging it here would duplicate validate.missing_fields'
' job; only the "blank" rule itself flags blanks.
Private Function PkgValidate_RuleFails(ByVal ruleName As String, ByVal v As Variant, _
                                       ByVal lo As Double, ByVal hasLo As Boolean, _
                                       ByVal hi As Double, ByVal hasHi As Boolean) As String
    Select Case ruleName
        Case "blank"
            If Core_IsBlank(v) Then PkgValidate_RuleFails = "Blank"
        Case "number"
            If Not Core_IsBlank(v) Then
                If Not IsNumeric(v) Then PkgValidate_RuleFails = "Not a number"
            End If
        Case "date"
            If Not Core_IsBlank(v) Then
                If Not IsDate(v) Then PkgValidate_RuleFails = "Not a date"
            End If
        Case "range"
            If Not Core_IsBlank(v) Then
                If IsNumeric(v) Then
                    If hasLo Then
                        If CDbl(v) < lo Then PkgValidate_RuleFails = CStr(v) & " is below the minimum " & lo
                    End If
                    If Len(PkgValidate_RuleFails) = 0 And hasHi Then
                        If CDbl(v) > hi Then PkgValidate_RuleFails = CStr(v) & " is above the maximum " & hi
                    End If
                End If
            End If
    End Select
End Function

Private Sub PkgValidate_WriteScanReport(ws As Worksheet, outRows As Collection)
    Dim i As Long, item As Variant
    ws.Range("A1:E1").Value = Array("Cell", "Column", "Rule", "Value", "Reason")
    ws.Range("A1:E1").Font.Bold = True
    i = 2
    For Each item In outRows
        ws.Cells(i, 1).Value = item(0)
        ws.Cells(i, 2).Value = item(1)
        ws.Cells(i, 3).Value = item(2)
        ws.Cells(i, 4).Value = item(3)
        ws.Cells(i, 5).Value = item(4)
        i = i + 1
    Next item
    ws.Columns("A:E").AutoFit
End Sub
```

- [ ] **Step 2: Build, syntax-check, commit**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add src/packages/PkgValidate.bas
git commit -m "feat(validate): validate.scan - configured blank/number/date/range rules, reported to a new sheet"
```

---

### Task 5: `validate.highlight_duplicates`

**Owner: delegate to codex (`sonnet`).** Reuses Task 2's colour helpers and Task 3's
`PkgValidate_KeyOfArr`/`PkgValidate_Grid`; the VBA below is complete.

**Files:** Modify `src/packages/PkgValidate.bas` — append after Task 4.

**Interfaces:** Consumes `Core_AskRange`, `Core_NarrowToUsed`, `Core_AskHasHeader`,
`Core_AskCols`, `Core_ParseCols`, `Core_RangeRef`, `Core_SpeedOn`/`Off`, `Core_Refuse`,
`Core_Done`, `Core_Fail`, `PkgValidate_Grid`, `PkgValidate_KeyOfArr` (Task 3),
`PkgValidate_ClassifyMark`, `PkgValidate_ConfirmColor`, `PkgValidate_MarkCell`,
`PkgValidate_ColorDuplicate` (Task 2). Produces `PkgValidate_HighlightDuplicates`.

- [ ] **Step 1: The action and its engine**

```vb
'================================================================
' ACTION - Colour-flag rows sharing a duplicate key
'================================================================
Public Sub PkgValidate_HighlightDuplicates(ctx As Object)
    Dim rng As Range, hasHeader As Boolean, keys As String, cols As Variant
    Const TITLE As String = "Highlight Duplicates"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to check." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub
    If Not Core_AskCols(rng, keys, _
            "Which column(s) form the key - two rows sharing the same value(s) here are " & _
            "duplicates?", TITLE) Then Exit Sub
    cols = Core_ParseCols(keys)

    Dim dupCells As Collection, dupRows As Long
    Core_SpeedOn
    Set dupCells = PkgValidate_FindDuplicateCells(rng, hasHeader, cols, dupRows)
    Core_SpeedOff

    If dupCells.Count = 0 Then
        Core_Done ctx, "No duplicate keys found inside " & Core_RangeRef(rng) & "." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Dim willMark As Long, willSkip As Long
    PkgValidate_ClassifyMark dupCells, willMark, willSkip
    If Not PkgValidate_ConfirmColor("marking the key cell(s) of " & dupRows & " row(s) that " & _
            "share a key with another row", willMark, willSkip, Core_RangeRef(rng), TITLE) Then Exit Sub

    Dim marked As Long, skipped As Long, c As Range
    Core_SpeedOn
    marked = 0: skipped = 0
    For Each c In dupCells
        PkgValidate_MarkCell c, PkgValidate_ColorDuplicate(), marked, skipped
    Next c
    Core_SpeedOff

    Core_Done ctx, "Colour-flagged " & marked & " key cell(s) across " & dupRows & " row(s) " & _
           "sharing a duplicate key." & IIf(skipped > 0, vbCrLf & vbCrLf & skipped & " cell(s) " & _
           "already had a fill colour of their own and were left alone.", "") & vbCrLf & vbCrLf & _
           "To remove the flags: select " & Core_RangeRef(rng) & " and use Home > Fill Color > " & _
           "No Fill.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.highlight_duplicates", Err
End Sub

' Every KEY-COLUMN cell (not every cell in the row) belonging to a row
' whose composite key is shared with at least one other row in rng.
' Colouring only the key cells, not the whole row, keeps the footprint
' of a single confirmed action as small as the finding it reports - the
' row is a duplicate BECAUSE of what is in these cells, and a wider
' flag would touch formatting on columns that had nothing to do with
' it.
Private Function PkgValidate_FindDuplicateCells(rng As Range, ByVal hasHeader As Boolean, cols As Variant, _
                                                ByRef dupRowCount As Long) As Collection
    Dim vals As Variant, dict As Object, r As Long, firstRow As Long, k As Variant, i As Long, j As Long
    Dim out As New Collection
    vals = PkgValidate_Grid(rng)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    firstRow = IIf(hasHeader, 2, 1)
    For r = firstRow To UBound(vals, 1)
        k = PkgValidate_KeyOfArr(vals, r, cols)
        If Not dict.Exists(k) Then dict.Add k, New Collection
        dict(k).Add r
    Next r
    dupRowCount = 0
    For Each k In dict.Keys
        If dict(k).Count > 1 Then
            dupRowCount = dupRowCount + dict(k).Count
            For i = 1 To dict(k).Count
                r = dict(k)(i)
                For j = LBound(cols) To UBound(cols)
                    out.Add rng.Cells(r, cols(j))
                Next j
            Next i
        End If
    Next k
    Set PkgValidate_FindDuplicateCells = out
End Function
```

- [ ] **Step 2: Build, syntax-check, commit**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add src/packages/PkgValidate.bas
git commit -m "feat(validate): validate.highlight_duplicates - key-cell colour flags, never overwriting an existing fill"
```

---

### Task 6: `validate.highlight_outliers`

**Owner: delegate to codex (`sonnet`).** Reuses `Core_ResolveColumn` and Task 2's colour
helpers; the VBA below is complete.

**Files:** Modify `src/packages/PkgValidate.bas` — append after Task 5.

**Interfaces:** Consumes `Core_CfgTable`, `Core_AskRange`, `Core_NarrowToUsed`,
`Core_AskHasHeader`, `Core_RangeRef`, `Core_IsBlank`, `Core_SpeedOn`/`Off`, `Core_Refuse`,
`Core_Done`, `Core_Fail`, `Core_ResolveColumn`, `PkgValidate_Grid`,
`PkgValidate_ClassifyMark`, `PkgValidate_ConfirmColor`, `PkgValidate_MarkCell`,
`PkgValidate_ColorOutlier` (Task 2). Produces `PkgValidate_HighlightOutliers`.

- [ ] **Step 1: The action and its engine**

```vb
'================================================================
' ACTION - Colour-flag values outside a configured bound
'================================================================
Public Sub PkgValidate_HighlightOutliers(ctx As Object)
    Dim rng As Range, hasHeader As Boolean
    Dim cols() As String, mins() As String, maxs() As String, missed As String
    Const TITLE As String = "Highlight Outliers"
    On Error GoTo Fail

    If Not Core_CfgTable(ctx, "k_validate_outliers", cols, mins, maxs) Then
        Core_Refuse ctx, "No outlier bounds are set up yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Outlier bounds' table on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to check." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub

    Dim outCells As Collection
    Core_SpeedOn
    Set outCells = PkgValidate_FindOutlierCells(rng, hasHeader, cols, mins, maxs, missed)
    Core_SpeedOff

    If outCells.Count = 0 Then
        Core_Done ctx, "No values outside their configured bound(s) were found inside " & _
               Core_RangeRef(rng) & "." & vbCrLf & vbCrLf & "Nothing has been changed." & _
               IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These configured column(s) matched " & _
                   "nothing in your selection and were skipped:" & missed, ""), TITLE
        Exit Sub
    End If

    Dim willMark As Long, willSkip As Long
    PkgValidate_ClassifyMark outCells, willMark, willSkip
    If Not PkgValidate_ConfirmColor("that fall outside their configured bound", willMark, _
            willSkip, Core_RangeRef(rng), TITLE) Then Exit Sub

    Dim marked As Long, skipped As Long, c As Range
    Core_SpeedOn
    marked = 0: skipped = 0
    For Each c In outCells
        PkgValidate_MarkCell c, PkgValidate_ColorOutlier(), marked, skipped
    Next c
    Core_SpeedOff

    Core_Done ctx, "Colour-flagged " & marked & " value(s) outside their configured bound." & _
           IIf(skipped > 0, vbCrLf & vbCrLf & skipped & " cell(s) already had a fill colour " & _
           "of their own and were left alone.", "") & vbCrLf & vbCrLf & "To remove the flags: " & _
           "select " & Core_RangeRef(rng) & " and use Home > Fill Color > No Fill." & _
           IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These configured column(s) matched nothing " & _
               "in your selection and were skipped:" & missed, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.highlight_outliers", Err
End Sub

' An outlier check only applies to actual numbers - a blank or
' non-numeric cell is SKIPPED here, never flagged. Flagging blanks is
' validate.missing_fields' job; conflating the two would mean a single
' cell could be flagged twice, for two different reasons, by two
' different actions, with no way to tell which one fired.
Private Function PkgValidate_FindOutlierCells(rng As Range, ByVal hasHeader As Boolean, cols() As String, _
                                              mins() As String, maxs() As String, ByRef missed As String) As Collection
    Dim out As New Collection
    Dim i As Long, r As Long, firstRow As Long, idx As Long, tgt As Range, vals As Variant, v As Variant
    Dim lo As Double, hi As Double, hasLo As Boolean, hasHi As Boolean
    firstRow = IIf(hasHeader, 2, 1)
    For i = LBound(cols) To UBound(cols)
        idx = Core_ResolveColumn(rng, cols(i))
        If idx = 0 Then
            missed = missed & vbCrLf & "   " & cols(i)
        Else
            Set tgt = rng.Columns(idx)
            vals = PkgValidate_Grid(tgt)
            hasLo = IsNumeric(Trim$(mins(i)))
            If hasLo Then lo = CDbl(Trim$(mins(i)))
            hasHi = IsNumeric(Trim$(maxs(i)))
            If hasHi Then hi = CDbl(Trim$(maxs(i)))
            For r = firstRow To UBound(vals, 1)
                v = vals(r, 1)
                If IsNumeric(v) And Not Core_IsBlank(v) Then
                    If (hasLo And CDbl(v) < lo) Or (hasHi And CDbl(v) > hi) Then out.Add tgt.Cells(r, 1)
                End If
            Next r
        End If
    Next i
    Set PkgValidate_FindOutlierCells = out
End Function
```

- [ ] **Step 2: Build, syntax-check, commit**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add src/packages/PkgValidate.bas
git commit -m "feat(validate): validate.highlight_outliers - configured Min/Max bounds, colour flags only"
```

---

### Task 7: `validate.missing_fields`

**Owner: delegate to codex (`sonnet`).** The simplest of the five; reuses `Core_AskCols`
exactly as `clean.dupes` and Task 5 do, and Task 2's colour helpers.

**Files:** Modify `src/packages/PkgValidate.bas` — append after Task 6.

**Interfaces:** Consumes `Core_AskRange`, `Core_NarrowToUsed`, `Core_AskHasHeader`,
`Core_AskCols`, `Core_ParseCols`, `Core_RangeRef`, `Core_IsBlank`, `Core_SpeedOn`/`Off`,
`Core_Refuse`, `Core_Done`, `Core_Fail`, `PkgValidate_Grid`, `PkgValidate_ClassifyMark`,
`PkgValidate_ConfirmColor`, `PkgValidate_MarkCell`, `PkgValidate_ColorMissing` (Task 2).
Produces `PkgValidate_MissingFields`.

- [ ] **Step 1: The action and its engine**

```vb
'================================================================
' ACTION - Colour-flag blank cells in mandatory columns
'================================================================
Public Sub PkgValidate_MissingFields(ctx As Object)
    Dim rng As Range, hasHeader As Boolean, keys As String, cols As Variant
    Const TITLE As String = "Flag Missing Required Fields"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to check." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub
    If Not Core_AskCols(rng, keys, _
            "Which column(s) are mandatory - a blank cell in any of them gets flagged?", TITLE) Then Exit Sub
    cols = Core_ParseCols(keys)

    Dim blankCells As Collection
    Core_SpeedOn
    Set blankCells = PkgValidate_FindBlankCells(rng, hasHeader, cols)
    Core_SpeedOff

    If blankCells.Count = 0 Then
        Core_Done ctx, "No blank cells found in the mandatory column(s) inside " & _
               Core_RangeRef(rng) & "." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Dim willMark As Long, willSkip As Long
    PkgValidate_ClassifyMark blankCells, willMark, willSkip
    If Not PkgValidate_ConfirmColor("that are blank in a mandatory column", willMark, willSkip, _
            Core_RangeRef(rng), TITLE) Then Exit Sub

    Dim marked As Long, skipped As Long, c As Range
    Core_SpeedOn
    marked = 0: skipped = 0
    For Each c In blankCells
        PkgValidate_MarkCell c, PkgValidate_ColorMissing(), marked, skipped
    Next c
    Core_SpeedOff

    Core_Done ctx, "Colour-flagged " & marked & " blank mandatory cell(s)." & _
           IIf(skipped > 0, vbCrLf & vbCrLf & skipped & " cell(s) already had a fill colour " & _
           "of their own and were left alone.", "") & vbCrLf & vbCrLf & "To remove the flags: " & _
           "select " & Core_RangeRef(rng) & " and use Home > Fill Color > No Fill.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.missing_fields", Err
End Sub

Private Function PkgValidate_FindBlankCells(rng As Range, ByVal hasHeader As Boolean, cols As Variant) As Collection
    Dim out As New Collection
    Dim vals As Variant, r As Long, firstRow As Long, i As Long
    vals = PkgValidate_Grid(rng)
    firstRow = IIf(hasHeader, 2, 1)
    For r = firstRow To UBound(vals, 1)
        For i = LBound(cols) To UBound(cols)
            If Core_IsBlank(vals(r, cols(i))) Then out.Add rng.Cells(r, cols(i))
        Next i
    Next r
    Set PkgValidate_FindBlankCells = out
End Function
```

- [ ] **Step 2: Build, syntax-check, commit**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add src/packages/PkgValidate.bas
git commit -m "feat(validate): validate.missing_fields - blank-cell colour flags in mandatory columns"
```

---

### Task 8: Register `PkgValidate` with Core — the two-place step

**Owner: Opus, written directly.** Small, but per plan-of-record §3.4: "the single easiest way
to ship a silently dead wave" is skipping this. `PkgValidate.bas` has built and bundled since
Task 2 and has done nothing, because `Core_Run` never probes an unregistered package name.

**Files:** Modify `src/core/Core.bas` — `CoreCandidates()`.

**Interfaces:** Consumes nothing new. Produces: every `validate.*` action becomes reachable
through `Core_Run` and appears on a generated Macros page.

- [ ] **Step 1: Add the candidate**

```bash
grep -n 'Function CoreCandidates' src/core/Core.bas
```

Find the `c.Add "Clean"` line inside it and add immediately after:

```vb
    c.Add "Validate"                ' probes PkgValidate_Manifest
```

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -n 'c.Add "Validate"' src/core/Core.bas
```

Expected: exit 0, `Core + 5 package(s)`.

- [ ] **Step 3: Deliberate injection — prove the frozen-id gate still catches a mistake in a
      brand-new package**

`build/build.ps1:452` fails a malformed action id (not exactly two lowercase segments). This
wave is the first new package since that gate was last exercised against a *new* id (W1's four
ids passed clean); prove it still fires rather than assuming it does.

```bash
sed -i 's/"validate.reconcile", "Reconcile Two Datasets"/"validateReconcile", "Reconcile Two Datasets"/' src/packages/PkgValidate.bas
pwsh build/build.ps1; echo "exit=$?"
git checkout -- src/packages/PkgValidate.bas
pwsh build/build.ps1; echo "exit=$?"
```

Expected: **exit 1** the first time, naming `validateReconcile` and "must be exactly two
segments... Never camelCase"; **exit 0** after the revert. **Record both outputs in the PR
body.**

- [ ] **Step 4: Deliberate injection — prove `Core_DestUsable` refuses a second reconcile run**

Confirms statically (per plan-of-record §3.5 item 4, replacing an owner gate item) that
`validate.reconcile` will not silently overwrite its own prior output. This cannot be run
without Excel, so it is read, not executed: open `PkgValidate_Reconcile` and confirm the FIRST
executable line after the `On Error GoTo Fail` reads `destName = Trim(Core_Cfg(ctx,
"k_validate_recon_dest"))` immediately followed by `If Not Core_DestUsable(ctx, destName,
TITLE) Then Exit Sub`, **before** either `Core_AskRange` call. Record this read in the PR body
as the static equivalent of the withdrawn gate item; append the run-time confirmation to
`docs/designs/2026-07-28-deferred-validation-log.md` in Task 9.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): register PkgValidate as a candidate package - W2's actions are now reachable"
```

---

### Task 9: Version, changelog, code map, deferred-validation log, and the PR

**Owner: Opus.**

**Files:** `VERSION`, `CHANGELOG.md`, `docs/code-map.md`,
`docs/designs/2026-07-28-deferred-validation-log.md`.

- [ ] **Step 1: `VERSION` -> `0.13.0`, and a `## [0.13.0]` changelog entry**

House style: prose about what changed for the user. Lead with `validate.reconcile` — it is
catalog line 93's "highest-value single item... for a multi-entity close" — then the other
four, and state plainly that none of the five repairs anything.

- [ ] **Step 2: Append this wave's run-time-only claims to the deferred validation log**

```markdown
## W2 - Validate & Reconcile (v0.13.0)

- **`Core_AskRange`'s dialog title doubles as the Windows title-bar caption.** `reconcile`
  passes titles as long as "Dataset B (the comparison side)" - whether that truncates visibly
  in the title bar is unconfirmed until run.
- **The reconcile engine's dictionary-based key index has not been run against a dataset large
  enough to observe the intended performance difference over a nested-loop implementation.**
  The O(n) vs O(n x m) argument (Task 3 Step 2) is a static one.
- **The three highlight colours' legibility on a dark Excel theme is unconfirmed** - same class
  of finding as the carried-forward Reskin 4 item (plan-of-record §3.6): the palette was chosen
  by RGB value against a white-sheet assumption, and nothing has checked it against a dark
  theme.
- **`validate.reconcile`'s refusal on a second run against its own output sheet** was confirmed
  by static read (Task 8 Step 4), not by an actual second run in Excel.
- **`Range.Value2`'s one-cell-returns-a-scalar edge case** (`PkgValidate_Grid`) has not been
  exercised against an actual one-cell pick in Excel; the guard is reasoned from documented
  `Value2` behaviour, not observed.
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
grep -n 'c.Add "Validate"' src/core/Core.bas
git status --short
```

Expected: exit 0; no diagnostics; `1`; no output from the prefix grep; one match; a tree clean
but for the untracked `.vscode/`.

- [ ] **Step 5: Review the diff, then one PR against `main`**

Branch from `main` after phase 2 has merged; not stacked. The PR body must state, in one line,
that no `validate.*` action deletes, moves, or overwrites a value or formula — a reviewer
reading a brand-new package's diff needs the one-sentence version of the domain's constraint
before checking whether the diff honours it. Include both deliberate-injection outputs from
Task 8.

---

## Unknowns and risks

- **This wave is `Core_ResolveColumn`'s third consumer** (`PkgClean`, then phase 2's own
  promotion pass, now `PkgValidate`) — see Decision 13. Its signature returns a `Long` column
  index counted from 1 at the picked range's left edge, 0 when the spec does not resolve
  inside the range, not a `Range`; every call site here that wants the column itself converts
  with `rng.Columns(idx)`.
- **`PkgValidate_KeyOfArr` is a plausible Core promotion candidate for W4's `enrich.join`**,
  which is not yet planned. Flagged (Decision 14), not promoted speculatively.
- **The three highlight actions can each be run again over cells they already coloured.** A
  second run of `validate.highlight_duplicates` over the same range sees its own prior fill and
  skips those cells (reported as "already had a fill colour"), which is safe but means the
  count of newly-marked cells on a re-run reads as smaller than the true number of duplicate
  rows. This is a byproduct of the never-overwrite rule (Decision 8), not a defect in it, and is
  named here so it does not read as one later.
- **`validate.reconcile`'s "Differs" comparison ignores columns present on only one side.** They
  are not silently dropped from the *dataset* — `PkgValidate_RowTextArr` still includes them in
  an `In A only` / `In B only` row's full record — but a column unique to one side is never
  compared for a MATCHED key pair, because there is nothing on the other side to compare it to.
  This is inherent to comparing two datasets that are not guaranteed to share a schema, not an
  oversight.
- **Numeric tolerance in `PkgValidate_ValuesEqual` is a fixed epsilon (`1E-9`), not configurable.**
  Reasonable for typical finance figures; a dataset carrying values at that precision boundary
  deliberately is not a case this wave designs for.
- **Performance claims (Decision 16, Task 3's engine comment) are argued from VBA's documented
  COM-call cost, not measured** — see the deferred-validation-log entries in Task 9 Step 2.
