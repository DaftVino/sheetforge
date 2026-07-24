# Sanitization pass — strip client jargon before the first remote push

Status: approved design (2026-07-24) · Blocker for `git remote add origin` →
`github.com/DaftVino/sheetforge`. Nothing has been pushed yet.

## Goal

Make the tracked repo publishable to a (currently private, later possibly
public) remote by replacing one client's dashboard jargon with **fully generic,
domain-neutral placeholders**, so the toolkit reads as reusable by anyone. This
is a mechanical 1:1 rename — **no logic, control flow, type-map letters, or
column counts change.**

## Decisions (settled in brainstorming)

1. **Defaults strategy — neutral demo, tracked only.** Manifest defaults are just
   seeds; at runtime each workbook's Setup/`_Config` sheet overrides them, so the
   real team keeps their real values in their own workbook and nothing real needs
   to live in source. No private overlay file (YAGNI).
2. **Legacy module — sanitize in place.** `src/legacy/SetupImportExportMacros.bas`
   never ships in the bundle (the build concatenates only `src/core` +
   `src/packages/Pkg*.bas`) but is cited by path in `Core.bas` and
   `docs/handoff.md`, so it stays tracked as scrubbed provenance rather than being
   deleted or moved.
3. **Naming style — fully generic, not domain-plausible.** A reusable toolkit
   ships placeholders anyone can map to their own data. Where a name is arbitrary
   (source files, input tabs) the placeholder reads as "replace me"; where a
   concept is universal (Summary, Details) a plain word is used.
4. **Scope — all 7 tracked files, not just the two `.bas` the TODO named.** The
   same jargon leaks into `Core.bas`, `CHANGELOG.md`, and three docs; the whole
   repo goes to the remote, so a code-only scrub would still publish the client
   name. Rewriting the historical `CHANGELOG` [0.1.0] entry text is acceptable
   because nothing has been published.

## Substitution map (final)

| Real (sensitive) | → Generic placeholder |
|---|---|
| `‹client dashboard name›` (proper name, narrative + Control-Panel title) | **Sample Dashboard** |
| `‹client export-base name›` (export base name default) | **My Report** |
| `‹legacy module id›` / `‹legacy module id›.bas` (legacy module identifier) | **`SetupImportExportMacros`** / **`SetupImportExportMacros.bas`** (module name now matches its filename) |
| Export tabs: `‹retired export-tab name›`, `‹retired export-tab name›`, `‹retired export-tab name›`, `‹retired export-tab name›`, `‹retired export-tab name›` | **Summary, Details, Trends, Outliers, Breakdown** |
| Import tab keys: `‹retired input-tab key›`, `‹retired input-tab key›`, `‹retired input-tab key›`, `‹retired input-tab key›` | **IN_Data1, IN_Data2, IN_Data3, IN_Data4** |
| Source-file prefixes: `‹retired source-file prefix›`, `‹retired source-file prefix›`, `‹retired source-file prefix›`, `‹retired source-file prefix›` | **Source File A, Source File B, Source File C, Source File D** |
| `‹retired master-list prefix›` (master-list file prefix) | **Master List Import** |
| Master columns: `‹retired master column›`, `‹retired master column›`, `‹retired master column›`, `‹retired master column›`, `‹retired master column›` | **Store ID, Location Name or Address, Entity ID, Mgr (DM) Group, Mgr (Loc) Name** |
| Comment jargon (legacy only): `‹retired comment term›`, `‹retired comment term›`, `‹retired comment term›`, `‹retired comment term›`, `‹retired comment term›` | reworded to domain-neutral equivalents that keep the technical point |

**Kept as-is (already generic):** `README` (settings sheet), `Ref_Master`
(destination tab), `Current Master` (inner sheet name), the `IN_` prefix
convention, every type-map string (`T,T,…,N`), all row/tab/column **counts** (5
export tabs, 4 import rows, 5 master cols).

The five master columns are a **list field on the Setup sheet** — users add,
remove, and rename rows to match their own master file. The help text is updated
to say so explicitly, so the generic seed is obviously replaceable.

## File-by-file change list

Line numbers are pre-edit anchors; verify at edit time.

### `src/packages/PkgImportExport.bas`
- L5 comment "the original ‹client dashboard name›'s four buttons" → "Sample Dashboard".
- L13 comment "the legacy ‹legacy module id› module" → "SetupImportExportMacros".
- L28 `k_export_base` default `‹client export-base name›` → `My Report`.
- L39–43 export-tab defaults → Summary / Details / Trends / Outliers / Breakdown.
- L48 help: "keeps IDs, ‹retired comment term›, leading zeros" → "keeps IDs, codes, leading zeros".
- L49–52 import defaults → `IN_Data1|Source File A|…` … `IN_Data4|Source File D|…` (type maps unchanged).
- L56 `k_master_prefix` default `‹retired master-list prefix›` → `Master List Import`.
- L63 help → "The FIRST one is the key column and is always kept as text. Replace these with the exact headers from your own master file."
- L64–68 master-col defaults → Store ID / Location Name or Address / Entity ID / Mgr (DM) Group / Mgr (Loc) Name.
- L194 comment "Update Ref_Master from the ‹retired master-list prefix›" → "Master List Import".
- L265–266 comment "Force the ‹retired master column› column … dashboard's text-based ‹retired comment term›" → generic ("first (key) column … the workbook's text-based keys").
- L477 user message "(e.g. ‹retired comment term›)" → "(e.g. an ID or code)".

### `src/legacy/SetupImportExportMacros.bas`
- L1 `Attribute VB_Name = "‹legacy module id›"` → `"SetupImportExportMacros"`.
- L5 header "‹client dashboard name› - Automation Macros" → "Sample Dashboard - Automation Macros".
- L73–90 import-default block: tab keys → `IN_Data1..4`; prefixes → `Source File A..D`; type maps unchanged.
- L85–87 comment reworded: drop ‹retired comment term›, ‹retired comment term›, ‹retired comment term›; keep the "this input stays all-text so its id/date columns survive; only numeric columns are N" point.
- L188 comment "(‹retired comment term›, category %s, ‹retired comment term›)" → "(e.g. totals, percentages, key metrics)".
- L210 comment "(‹retired comment term›, ‹retired export-tab name›, ‹retired export-tab name›, etc.)" → "(the summary tabs, etc.)".
- L243 comment "Update Ref_Master from the ‹retired master-list prefix›" → "Master List Import".
- L307–308 comment "Force the first (‹retired master column›) column … text-based ‹retired comment term›" → generic.
- L409 comment "‹retired export-tab name› FILTER" → "a FILTER spill".
- L749 Control-Panel title string `‹client dashboard name›  -  Control Panel` → `Sample Dashboard  -  Control Panel`.
- L889 fallback `‹client export-base name›` → `My Report`.
- L897 fallback `‹retired master-list prefix›` → `Master List Import`.
- L909–910 fallback export-tab array → Summary / Details / Trends / Outliers / Breakdown.
- L924–925 fallback master-col adds → Store ID / Location Name or Address / Entity ID / Mgr (DM) Group / Mgr (Loc) Name.

### `src/core/Core.bas`
- L7 provenance comment "Extracted from ‹legacy module id›" → "SetupImportExportMacros".

### `docs/handoff.md`
- L16, L64, L75 module references `‹legacy module id›` → `SetupImportExportMacros`.
- L64 "for the ‹client dashboard name›" → "the Sample Dashboard workbook".

### `docs/engine-architecture-draft.md`
- L13, L58, L220 `‹legacy module id›.bas` → `SetupImportExportMacros.bas`.
- L58 "the ‹client dashboard name›-specifics" → "the dashboard-specific parts".
- L221 "Nothing about the ‹client dashboard name› changes" → "the sample dashboard".

### `docs/macro-catalog-plug-and-play.md`
- L9 `‹legacy module id›.bas` → `SetupImportExportMacros.bas`.

### `CHANGELOG.md`
- L33 "the ‹client dashboard name›'s four buttons" → "the Sample Dashboard's four buttons".

## Verification

1. `pwsh build/build.ps1` — confirms no duplicate-proc / Sub-Function imbalance
   regression (renames touch string literals, comments, one `Attribute VB_Name`,
   and the Core comment only; legacy is not built).
2. Token sweep must return **zero** hits across tracked files (excluding
   `dist/`, `docs/designs/setup-simplicity.md`, and this spec):
   `git grep -i -n -E "<case-insensitive alternation over the retired client tokens>" -- <tracked paths>`
3. No Excel re-verify needed — behavior is unchanged (demo seed defaults only;
   `_Config` overrides at runtime; legacy fallbacks are illustrative).
4. Only after all three pass: `git remote add origin …` and push.

## Out of scope

- No version bump / new `CHANGELOG` [Unreleased] entry here — that belongs to
  `/ship`.
- No private-config overlay, no file moves/deletes, no logic changes.
- `docs/designs/setup-simplicity.md` and this spec are excluded from the token
  sweep (they discuss the sanitization by describing the old names).
