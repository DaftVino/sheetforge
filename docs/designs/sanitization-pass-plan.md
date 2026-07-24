# Sanitization Pass Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace one client's dashboard jargon with fully generic, domain-neutral placeholders across all 7 tracked files, so the repo is publishable to `github.com/DaftVino/sheetforge`.

**Architecture:** A mechanical 1:1 rename per the substitution map in `docs/designs/sanitization-pass.md`. No logic, control flow, type-map letters, or row/column counts change. "Tests" are token-sweep `git grep`s: red = jargon present, green = zero hits; the two built `.bas` files additionally pass `pwsh build/build.ps1`.

**Tech Stack:** VBA (`.bas`), PowerShell build script, `git grep` (via the Bash tool / Git Bash).

## Global Constraints

- **1:1 rename only** — never alter logic, control flow, type-map strings (`T,T,…,N`), or any row/tab/column **count** (5 export tabs, 4 import rows, 5 master cols stay 5/4/5).
- **Substitution map is authoritative** — every replacement comes from the table in `docs/designs/sanitization-pass.md`; do not invent new names.
- **Kept-as-is tokens** (do NOT change): `README`, `Ref_Master`, `Current Master`, the `IN_` prefix, all type-map letters.
- **Build must stay green:** `pwsh build/build.ps1` after any change to `src/core/` or `src/packages/` (legacy is not built).
- **Excluded from sweeps:** `dist/`, `docs/designs/setup-simplicity.md`, `docs/designs/sanitization-pass.md`, `docs/designs/sanitization-pass-plan.md` (they legitimately name the old tokens).
- **The remote push is out of scope** — this plan ends at "scrubbed + verified"; `git remote add` / `git push` is an outward action handed back to the user.

### The token-sweep command (reused by every task)

```bash
# $SWEEP = the jargon regex, case-insensitive, word-boundary on TabName
SWEEP='food|sampleDash|target inventory|inventory value|floater|TabName\b|Calc_Hours|net sales|master store|category var|var detail|inv value|clock in and out|store key|store number|store #|Sample Recap|IN_Day|IN_PTD|IN_InvValue|IN_Clock'
```

---

### Task 1: Scrub the two BUILT files (`PkgImportExport.bas` + `Core.bas`)

**Files:**
- Modify: `src/packages/PkgImportExport.bas` (manifest defaults, comments, help text, one user message)
- Modify: `src/core/Core.bas:7` (one provenance comment)

**Interfaces:**
- Consumes: nothing.
- Produces: the scrubbed bundle inputs; the build concatenates these into `dist/MacrosBundle.bas`.

- [ ] **Step 1: Run the sweep to confirm jargon is present (red)**

```bash
cd x:/Projects/sheetforge
SWEEP='food|sampleDash|target inventory|inventory value|floater|TabName\b|Calc_Hours|net sales|master store|category var|var detail|inv value|clock in and out|store key|store number|store #|Sample Recap|IN_Day|IN_PTD|IN_InvValue|IN_Clock'
git grep -i -n -E "$SWEEP" -- src/packages/PkgImportExport.bas src/core/Core.bas
```
Expected: multiple hits (L5, L13, L28, L39–43, L48, L49–52, L56, L63, L64–68, L194, L265–266, L477 in the package; L7 in Core).

- [ ] **Step 2: Apply the package renames**

In `src/packages/PkgImportExport.bas`, make these exact replacements:

- L5 `' PkgImportExport - the original Sample Dashboard's four` → `' PkgImportExport - the original Sample Dashboard's four`
- L13 `' Behavior is identical to the legacy SetupImportExportMacros module, with ONE` → `' Behavior is identical to the legacy SetupImportExportMacros module, with ONE`
- L28 default `"My Report"` → `"My Report"`
- L39 `"Sample Recap (Daily)"` → `"Summary"`
- L40 `"Sample Category"` → `"Details"`
- L41 `"Sample Detail"` → `"Trends"`
- L42 `"Inv Value Outliers"` → `"Outliers"`
- L43 `"Floater Finder"` → `"Breakdown"`
- L48 help, substring `T=Text (keeps IDs, store numbers, leading zeros)` → `T=Text (keeps IDs, codes, leading zeros)`
- L49 `"IN_Day|Target Inventory Data (day)|T,T,T,T,T,T,T,T,N,N,N,N,N,N,N"` → `"IN_Data1|Source File A|T,T,T,T,T,T,T,T,N,N,N,N,N,N,N"`
- L50 `"IN_PTD|Target Inventory Data (period to date)|T,T,T,T,T,T,T,T,N,N,N,N,N,N,N"` → `"IN_Data2|Source File B|T,T,T,T,T,T,T,T,N,N,N,N,N,N,N"`
- L51 `"IN_InvValue|Inventory Value Data|T,T,T,T,T,T,T,N,N,N,N"` → `"IN_Data3|Source File C|T,T,T,T,T,T,T,N,N,N,N"`
- L52 `"IN_Clock|Clock in and out Combine|T,T,T,T,T,T,T,T,T,N,N,N,N,N"` → `"IN_Data4|Source File D|T,T,T,T,T,T,T,T,T,N,N,N,N,N"`
- L56 default `"Master Store List"` → `"Master List Import"`
- L63 help `The exact column headers to pull from the master file, in order. The FIRST one is the store key and is always kept as text.` → `The exact column headers to pull from the master file, in order. The FIRST one is the key column and is always kept as text. Replace these with the exact headers from your own master file.`
- L64 `"Store #"` → `"Store ID"`
- L65 `"Location Name"` → `"Location Name or Address"`
- L66 `"Entity"` → `"Entity ID"`
- L67 `"DO Name"` → `"Mgr (DM) Group"`
- L68 `"GM Name"` → `"Mgr (Loc) Name"`
- L194 comment `' ACTION 3 - Update Ref_Master from the Master Store List` → `' ACTION 3 - Update Ref_Master from the Master List Import`
- L265–266 comment `' Force the Store column to Text so store numbers stay text and keep` / `' matching the dashboard's text-based store keys.` → `' Force the first (key) column to Text so codes/IDs keep leading zeros` / `' and keep matching the workbook's text-based keys.`
- L477 message substring `a column of ID / code numbers (e.g. store #) must` → `a column of ID / code numbers (e.g. an ID or code) must`

- [ ] **Step 3: Apply the Core rename**

In `src/core/Core.bas`:
- L7 `' Extracted from SetupImportExportMacros (src/legacy/SetupImportExportMacros.bas):` → `' Extracted from SetupImportExportMacros (src/legacy/SetupImportExportMacros.bas):`

- [ ] **Step 4: Run the sweep again (green)**

```bash
git grep -i -n -E "$SWEEP" -- src/packages/PkgImportExport.bas src/core/Core.bas
```
Expected: **no output** (exit 1).

- [ ] **Step 5: Build must stay green**

```bash
pwsh build/build.ps1
```
Expected: success — bundle written, no "duplicate procedure" / "Sub/Function imbalance" error.

- [ ] **Step 6: Commit**

```bash
git add src/packages/PkgImportExport.bas src/core/Core.bas
git commit -m "chore(sanitize): generic demo defaults in PkgImportExport + Core comment"
```

---

### Task 2: Scrub the legacy reference module (not built)

**Files:**
- Modify: `src/legacy/SetupImportExportMacros.bas` (module name, header, import-default block, fallback getters, comment jargon, Control-Panel title)

**Interfaces:**
- Consumes: the same substitution map as Task 1 (identical name choices).
- Produces: scrubbed provenance; not part of the build.

- [ ] **Step 1: Run the sweep to confirm jargon is present (red)**

```bash
cd x:/Projects/sheetforge
SWEEP='food|sampleDash|target inventory|inventory value|floater|TabName\b|Calc_Hours|net sales|master store|category var|var detail|inv value|clock in and out|store key|store number|store #|Sample Recap|IN_Day|IN_PTD|IN_InvValue|IN_Clock'
git grep -i -n -E "$SWEEP" -- src/legacy/SetupImportExportMacros.bas
```
Expected: hits at L1, L5, L73–90, L188, L210, L243, L307–308, L409, L749, L889, L897, L909–910, L924–925.

- [ ] **Step 2: Module identity + header**

- L1 `Attribute VB_Name = "SetupImportExportMacros"` → `Attribute VB_Name = "SetupImportExportMacros"`
- L5 `' Sample Dashboard - Automation Macros` → `' Sample Dashboard - Automation Macros`

- [ ] **Step 3: Import-default block (L73–90) — keys + prefixes, type maps unchanged**

- L73 `tabNames(1) = "IN_Day"` → `tabNames(1) = "IN_Data1"`
- L74 `prefixes(1) = "Target Inventory Data (day)"` → `prefixes(1) = "Source File A"`
- L77 `tabNames(2) = "IN_PTD"` → `tabNames(2) = "IN_Data2"`
- L78 `prefixes(2) = "Target Inventory Data (period to date)"` → `prefixes(2) = "Source File B"`
- L81 `tabNames(3) = "IN_InvValue"` → `tabNames(3) = "IN_Data3"`
- L82 `prefixes(3) = "Inventory Value Data"` → `prefixes(3) = "Source File C"`
- L88 `tabNames(4) = "IN_Clock"` → `tabNames(4) = "IN_Data4"`
- L89 `prefixes(4) = "Clock in and out Combine"` → `prefixes(4) = "Source File D"`
- L85–87 comment block → reword, dropping `Calc_Hours`/`store key`/`TabName`, keeping the technical point:
  ```vba
      ' IN_Data4 is kept ALL TEXT for its identifier/date columns: a downstream
      ' step parses the time text with DATEVALUE/TIMEVALUE, and the key column
      ' must stay text to match the master tab. Only the numeric columns are N.
  ```

- [ ] **Step 4: Comment jargon elsewhere**

- L188 comment `' (Net Sales, category %s, TabName).` → `' (e.g. totals, percentages, key metrics).`
- L210 comment `' (TabName, Sample Recap, Sample Category, etc.) update from the new data.` → `' (the summary tabs, etc.) update from the new data.`
- L243 comment `' BUTTON 3 - Update Ref_Master from the Master Store List` → `' BUTTON 3 - Update Ref_Master from the Master List Import`
- L307–308 comment `' Force the first (Store) column to Text so the store numbers stay text` / `' and keep matching the dashboard's text-based store keys.` → `' Force the first (key) column to Text so codes/IDs keep leading zeros` / `' and keep matching the workbook's text-based keys.`
- L409 comment substring `dynamic-array spill (e.g. Floater Finder FILTER)` → `dynamic-array spill (e.g. a FILTER formula)`

- [ ] **Step 5: Control-Panel title + fallback getters**

- L749 `.Value = "Sample Dashboard  -  Control Panel"` → `.Value = "Sample Dashboard  -  Control Panel"`
- L889 `ExportBaseName = SheetSingle(CfgSheet(), K_EXPBASE, "My Report")` → replace `"My Report"` with `"My Report"`
- L897 `MasterPrefix = SheetSingle(CfgSheet(), K_MPREFIX, "Master Store List")` → replace `"Master Store List"` with `"Master List Import"`
- L909–910 fallback array:
  ```vba
          GetExportTabs = Array("Summary", "Details", "Trends", _
                                "Outliers", "Breakdown")
  ```
- L924–925 fallback master-col adds:
  ```vba
          col.Add "Store ID": col.Add "Location Name or Address": col.Add "Entity ID"
          col.Add "Mgr (DM) Group": col.Add "Mgr (Loc) Name"
  ```

- [ ] **Step 6: Run the sweep again (green)**

```bash
git grep -i -n -E "$SWEEP" -- src/legacy/SetupImportExportMacros.bas
```
Expected: **no output** (exit 1).

- [ ] **Step 7: Sanity — Sub/Function balance unchanged**

```bash
grep -c -E "^\s*(Private |Public )?(Sub|Function) " src/legacy/SetupImportExportMacros.bas
grep -c -E "^\s*End (Sub|Function)" src/legacy/SetupImportExportMacros.bas
```
Expected: the two counts are equal (renames touched no block openers/closers).

- [ ] **Step 8: Commit**

```bash
git add src/legacy/SetupImportExportMacros.bas
git commit -m "chore(sanitize): generic names in legacy reference module + rename to SetupImportExportMacros"
```

---

### Task 3: Scrub the docs

**Files:**
- Modify: `docs/handoff.md` (L16, L64, L75)
- Modify: `docs/engine-architecture-draft.md` (L13, L58, L220, L221)
- Modify: `docs/macro-catalog-plug-and-play.md` (L9)
- Modify: `CHANGELOG.md` (L33)

**Interfaces:**
- Consumes: the module rename `SetupImportExportMacros` → `SetupImportExportMacros` and proper-name `Sample Dashboard` → `Sample Dashboard`.
- Produces: nothing downstream.

- [ ] **Step 1: Run the sweep to confirm jargon is present (red)**

```bash
cd x:/Projects/sheetforge
SWEEP='food|sampleDash|target inventory|inventory value|floater|TabName\b|Calc_Hours|net sales|master store|category var|var detail|inv value|clock in and out|store key|store number|store #|Sample Recap|IN_Day|IN_PTD|IN_InvValue|IN_Clock'
git grep -i -n -E "$SWEEP" -- docs/handoff.md docs/engine-architecture-draft.md docs/macro-catalog-plug-and-play.md CHANGELOG.md
```
Expected: hits at handoff L16/L64/L75, draft L13/L58/L220/L221, macro-catalog L9, CHANGELOG L33.

- [ ] **Step 2: Replace `SetupImportExportMacros` → `SetupImportExportMacros` everywhere it appears in these four docs**

- `docs/handoff.md` L16 `# Core module(s) — extracted from SetupImportExportMacros` → `… extracted from SetupImportExportMacros`
- `docs/handoff.md` L64 `(module `SetupImportExportMacros`)` → `(module `SetupImportExportMacros`)`
- `docs/handoff.md` L75 `**Extract Core** from `SetupImportExportMacros`:` → `… from `SetupImportExportMacros`:`
- `docs/engine-architecture-draft.md` L13 `Your current `SetupImportExportMacros.bas`` → `Your current `SetupImportExportMacros.bas``
- `docs/engine-architecture-draft.md` L58 `Your `SetupImportExportMacros.bas` already contains` → `Your `SetupImportExportMacros.bas` already contains`
- `docs/engine-architecture-draft.md` L220 `**Extract Core** from `SetupImportExportMacros.bas`:` → `… from `SetupImportExportMacros.bas`:`
- `docs/macro-catalog-plug-and-play.md` L9 `already implemented in `SetupImportExportMacros.bas`.` → `… in `SetupImportExportMacros.bas`.`

- [ ] **Step 3: Replace the proper-name / lowercase narrative mentions**

- `docs/handoff.md` L64 `The working dashboard automation for the Sample Dashboard.` → `The working dashboard automation for the Sample Dashboard workbook.`
- `docs/engine-architecture-draft.md` L58 `lift those out* of the sample-dashboard specifics` → `lift those out* of the dashboard-specific parts`
- `docs/engine-architecture-draft.md` L221 `Nothing about the sample dashboard changes for the user` → `Nothing about the sample dashboard changes for the user`
- `CHANGELOG.md` L33 `the Sample Dashboard's four buttons` → `the Sample Dashboard's four buttons`

- [ ] **Step 4: Run the sweep again (green)**

```bash
git grep -i -n -E "$SWEEP" -- docs/handoff.md docs/engine-architecture-draft.md docs/macro-catalog-plug-and-play.md CHANGELOG.md
```
Expected: **no output** (exit 1).

- [ ] **Step 5: Commit**

```bash
git add docs/handoff.md docs/engine-architecture-draft.md docs/macro-catalog-plug-and-play.md CHANGELOG.md
git commit -m "docs(sanitize): generic module + dashboard names across docs and changelog"
```

---

### Task 4: Final repo-wide verification (readiness gate — no new deliverable)

**Files:** none modified unless a straggler is found.

**Interfaces:**
- Consumes: Tasks 1–3 complete.
- Produces: a clean tree ready for `git remote add` (performed later, by the user).

- [ ] **Step 1: Repo-wide sweep with the documented exclusions**

```bash
cd x:/Projects/sheetforge
SWEEP='food|sampleDash|target inventory|inventory value|floater|TabName\b|Calc_Hours|net sales|master store|category var|var detail|inv value|clock in and out|store key|store number|store #|Sample Recap|IN_Day|IN_PTD|IN_InvValue|IN_Clock'
git grep -i -n -E "$SWEEP" -- . ':(exclude)dist/' ':(exclude)docs/designs/setup-simplicity.md' ':(exclude)docs/designs/sanitization-pass.md' ':(exclude)docs/designs/sanitization-pass-plan.md'
```
Expected: **no output** (exit 1). If any straggler appears, fix it inline, re-run, then commit with `chore(sanitize): straggler`.

- [ ] **Step 2: Build green on the final tree**

```bash
pwsh build/build.ps1
```
Expected: success.

- [ ] **Step 3: STOP — hand back to the user**

Report: branch `chore/sanitize-defaults` is scrubbed and verified; the repo is ready for `git remote add origin …` + push. Do **not** add the remote or push — that outward action is the user's call.

---

## Self-Review

**Spec coverage:** every row of the substitution map and every file in the spec's file-by-file list maps to a task —
built `.bas` + Core comment → Task 1; legacy module (incl. TabName/Calc_Hours/Net Sales comment rewrites) → Task 2;
handoff/draft/macro-catalog/CHANGELOG → Task 3; verification (`build` + zero-token sweep) → Tasks 1/2/3 locally and Task 4 globally. No gaps.

**Placeholder scan:** no TBD/TODO/"handle edge cases"; every step shows the exact old→new text or the exact command + expected output.

**Type/name consistency:** the same targets are used everywhere — `SetupImportExportMacros`→`SetupImportExportMacros`, tabs `Summary/Details/Trends/Outliers/Breakdown`, keys `IN_Data1..4`, prefixes `Source File A..D`, master cols `Store ID / Location Name or Address / Entity ID / Mgr (DM) Group / Mgr (Loc) Name`, `Master List Import`, `My Report`. Type-map letters and counts are copied verbatim and never altered.
