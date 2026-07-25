# Macro Catalog — Plug & Play Library

A master list of candidate macros for the config-driven macro project, grouped by **what the user is trying to get done** (not by how the code works). The User-Input Config driver is the backbone that selects and parameterizes everything below, so it is not listed as an item.

## How to read this

- **⭐ (1–5)** — how important I think it is to *include this in the project*. 5 = core, must-have; 1 = nice-to-have / gimmick.
- **Grade (A–F)** — how *well-used / downloaded / requested* the macro is in the real world (forums, code libraries, Stack Overflow). A = everyone wants it; F = niche.
- **✅ DONE** — already implemented. `(embedded)` means it lives inside one of the button routines or as a Core helper rather than as its own selectable action.

The two ratings are independent on purpose: a few things are hugely popular but low-value here (row spotlight), and a few are lower-profile but critical to *your* finance work (reconciliation, type coercion).

### Build verdicts (added 2026-07-24)

Every item now carries a verdict, making this file the **authoritative build queue**. Rationale and the wave protocol live in `docs/designs/2026-07-24-finish-excel-design.md`.

- **W1**–**W6** — assigned to that wave, with the frozen action id.
- **Later** — in scope, no wave assigned yet.
- **Deferred** — blocked on a safety story (destructive with no preview, or an undefined failure mode).
- **Declined** — will not be built; reason given inline.

Waves W1–W6 assign 34 items (35 new actions). With the 13 already shipped or embedded that brings the toolkit to ~48 listed actions. 8 items are *Later*, 5 *Deferred*, 4 *Declined*. "Build all 60" is explicitly not the goal — an action earns its place only with a stable user story, bounded inputs/outputs, a safe cancel and error-restore path, and a repeatable hand-verification script.

Action ids follow the frozen `<domain>.<operation>` scheme (contract v1.2.0). The domain is a user-facing vocabulary, **not** the module that implements it.

---

## 1. Bring Data In — import & combine sources

### Import all files from a folder into their tabs (type-aware) — ⭐⭐⭐⭐⭐ · Grade C ✅ DONE
Matches source files by name prefix (handles ` (1)`/` (2)` suffixes), clears the target, pastes, and coerces column types. Your `ImportSourceData`. — `importexport.import`

### Coerce imported text to real types (T / N / D / DT / B / G) — ⭐⭐⭐⭐⭐ · Grade C ✅ DONE (embedded)
Forces text-based exports back to numbers/dates/booleans so downstream formulas keep matching. Your `CoerceVal`, now Core's `Core_Coerce`. Rare as a shared snippet, but it's the reason your imports don't silently break. **Stays embedded** — the type map is a config field of the import action, not a standalone button.

### Find & open the newest file matching a prefix — ⭐⭐⭐⭐ · Grade B ✅ DONE (embedded)
Safe folder scan (collects names first, then stats — avoids the `Dir()` iterator corruption bug). Your `FindSourceFile`, now Core's `Core_FindFile`. **Stays a Core helper** — infrastructure, not a user action.

### Combine all worksheets into one master sheet — ⭐⭐⭐⭐⭐ · Grade A — **W1** `import.combine_worksheets`
Stacks every tab into a single table, aligning headers. One of the most-requested macros anywhere. Shares `PkgImport_CombineEngine` with the next item.

### Combine all workbooks in a folder into one — ⭐⭐⭐⭐⭐ · Grade A — **W1** `import.combine_workbooks`
Loops a folder, appends each file's data. Universal for consolidations and rollups. Shares `PkgImport_CombineEngine` with the previous item — same header-aligning append, different source enumeration.

### Import CSV / TXT with delimiter parsing — ⭐⭐⭐⭐ · Grade A — **W1** `import.delimited_file`
Bring in a delimited file with explicit qualifiers and type handling.

### Split one column into many (Text to Columns) — ⭐⭐⭐ · Grade A — **Later**
For exports that land everything in column A.

### Pull data from a closed workbook — ⭐⭐⭐ · Grade B — **Deferred**
Read a range without opening/showing the source. *Blocked on:* an undefined failure mode when the source is missing, locked, or has moved.

---

## 2. Clean & Standardize — fix messy data

### Clear data rows, keep headers + formatting — ⭐⭐⭐⭐⭐ · Grade B ✅ DONE
Wipes row 2 down on the listed tabs, preserves headers/formats. Your `ClearInputTabs`. — `importexport.clear`

### Remove blank rows / columns — ⭐⭐⭐⭐⭐ · Grade A — **W1** `clean.remove_blanks`
Delete empty rows/cols in the used range.

### Remove duplicates — ⭐⭐⭐⭐⭐ · Grade A ✅ DONE
De-dupe on user-chosen key columns. — `clean.dupes`

### Trim spaces & strip non-printing characters — ⭐⭐⭐⭐⭐ · Grade A ✅ DONE
`TRIM` + `CLEAN` across a selection; kills the invisible junk that breaks lookups. — `clean.trim`

### Fill blanks with the value above — ⭐⭐⭐⭐ · Grade A ✅ DONE
Repairs "group label only on first row" exports. — `clean.filldown`

### Unmerge cells and fill down the value — ⭐⭐⭐⭐ · Grade B — **W1** `clean.unmerge_fill`
Turns merged report layouts back into analyzable data.

### Standardize number / date formats — ⭐⭐⭐⭐ · Grade B — **W1** `clean.standardize_formats`
Apply consistent formats to key columns after transfer.

### Change case (UPPER / lower / Proper) — ⭐⭐⭐ · Grade A — **Later**
Normalize text casing across a range.

### Convert text-numbers to real numbers — ⭐⭐⭐ · Grade B — **W1** `clean.text_to_numbers`
Strip `$`, `,`, `%`, leading apostrophes; re-type as numeric. Calls the existing `Core_Coerce` rather than reimplementing coercion.

### Delete all hidden rows / columns — ⭐⭐ · Grade C — **Deferred**
Remove filtered-out or hidden data before sharing. *Blocked on:* destructive with no preview — the user cannot see what they are about to lose.

---

## 3. Validate & Reconcile — catch problems before they ship

### Reconcile two datasets on a key, flag mismatches — ⭐⭐⭐⭐⭐ · Grade C — **W2** `validate.reconcile`
Match A vs B on an ID, report missing/extra/differing rows. High value for finance close; under-shared as a generic snippet. The highest-value single item in this catalog for a multi-entity close.

### Data-validation scan (blanks / type / out-of-range) → report — ⭐⭐⭐⭐ · Grade C — **W2** `validate.scan`
Sweep a range against rules and write an exceptions list.

### Highlight duplicates — ⭐⭐⭐⭐ · Grade A — **W2** `validate.highlight_duplicates`
Color-flag repeated keys without deleting.

### Highlight outliers / out-of-threshold values — ⭐⭐⭐ · Grade C — **W2** `validate.highlight_outliers`
Flag values past a parameterized bound (your dashboard already leans on this idea via CF).

### Flag missing required fields — ⭐⭐⭐ · Grade C — **W2** `validate.missing_fields`
Mark rows missing a mandatory column.

---

## 4. Match & Enrich — look things up

### Refresh a reference/master table from an external file (column-mapped) — ⭐⭐⭐⭐⭐ · Grade C ✅ DONE
Pulls named columns from a master file in header order, keeps the key as text. Your `UpdateRefMaster`. — `importexport.updateMaster`

### VLOOKUP / XLOOKUP-style join in code — ⭐⭐⭐⭐ · Grade B — **W4** `enrich.join`
Fast in-memory join over large sets (dictionary-based), avoids volatile lookup formulas.

### Add mapped columns from a lookup table — ⭐⭐⭐ · Grade B — **W4** `enrich.mapped_columns`
Enrich rows with attributes pulled from a reference sheet.

---

## 5. Format & Make Presentable

### Conditional highlighting by rule / parameter — ⭐⭐⭐⭐ · Grade A — **W3** `format.conditional_highlight`
Rules-driven coloring beyond native CF, controlled from config.

### Style header row + autofit columns — ⭐⭐⭐⭐ · Grade A — **W3** `format.style_headers`
Bold/fill/font the header, size every column. The classic "make an export readable" macro.

### Add & style a report title (merge + center) — ⭐⭐⭐ · Grade B — **W3** `format.report_title`
Drop a formatted title block on top of a report.

### Apply a house table style (banding / borders) — ⭐⭐⭐ · Grade B — **W3** `format.table_style`
One-click brand/format standard for any table.

### Print-layout prep (landscape / fit-to-width / margins) — ⭐⭐⭐ · Grade B — **W3** `format.print_layout`
Board-pack and period-end binder formatting.

### Highlight cells with comments / formulas / blanks — ⭐⭐ · Grade A — **Later**
Popular audit-view helper; light value in a driven library.

### Spotlight the active row / column — ⭐⭐ · Grade A — **Declined**
Very popular download, but mostly cosmetic. *Declined:* requires a `SelectionChange` event handler in a sheet module, which a single imported standard-module bundle cannot install. Faking support is worse than declining.

---

## 6. Summarize & Analyze

### Refresh all pivot tables — ⭐⭐⭐⭐ · Grade A ✅ DONE (embedded) — **W4** `summarize.refresh_pivots`
Loops every sheet and refreshes each pivot after import. Inside your `ImportSourceData`. **Promoted to a standalone action in W4** so it can be run on its own and used as a chain step.

### Force full recalculation — ⭐⭐⭐ · Grade B ✅ DONE (embedded) — **W4** `summarize.recalculate`
`CalculateFull` so downstream tabs update. Inside Import/Export. **Promoted to a standalone action in W4.**

### Build a summary / pivot from a raw range — ⭐⭐⭐⭐ · Grade B — **W4** `summarize.build_pivot`
Generate a summary table from chosen dimensions + measures.

### Copy a pivot to a static values sheet — ⭐⭐⭐ · Grade B — **W4** `summarize.freeze_pivot`
Freeze pivot output as values + formatting for distribution.

### Apply a consistent style to all charts — ⭐⭐ · Grade B — **Later**
Enforce one visual standard across chart objects.

---

## 7. Get It Out — export & share

### Export a values-only snapshot (.xlsx) — ⭐⭐⭐⭐⭐ · Grade B ✅ DONE
Copies main tabs to a new file, flattens spills + legacy arrays + formulas to values, keeps formatting/CF, breaks external links, dates the filename. Your `ExportDashboard` — the most sophisticated routine in the file. — `importexport.export`

### Convert formulas to values — ⭐⭐⭐⭐ · Grade A ✅ DONE (embedded) — **W3** `export.freeze_values`
Freeze formula outputs (incl. dynamic-array spills). Inside `ExportDashboard`. **Promoted to a standalone action in W3.**

### Dated / timestamped output filename — ⭐⭐⭐⭐ · Grade B ✅ DONE (embedded)
`yyyy-mm-dd` + base name with a configurable day offset. Inside `ExportDashboard`. **Stays embedded** — it is a config option of the export actions, not a button of its own.

### Export sheet / range / workbook to PDF — ⭐⭐⭐⭐⭐ · Grade A — **W3** `export.pdf`
One-click PDF of a range, sheet, or whole book.

### Save each worksheet as its own file / PDF — ⭐⭐⭐⭐ · Grade A — **W3** `export.per_sheet`
Explode a workbook into per-tab deliverables.

### Save each filtered group (or row) as its own file — ⭐⭐⭐ · Grade B — **Later**
Split by a column value into separate files.

### Email a sheet / range as an attachment (Outlook) — ⭐⭐⭐⭐ · Grade A — **Later**
Send the current output straight from Excel. In scope, but needs a defined behaviour when Outlook is absent or blocked by policy.

### Send bulk personalized emails from a list — ⭐⭐⭐ · Grade A — **Declined**
Mail-merge-style send driven by a table. *Declined:* not responsible without a preview, a recipient count, an explicit confirmation, and a draft-only mode. Revisit only as a deliberate, separately designed feature.

---

## 8. Move Around & Housekeeping — workbook management

### Unhide all worksheets — ⭐⭐⭐ · Grade A — **W6** `workbook.unhide_sheets`
The single most-downloaded VBA snippet in existence.

### Protect / unprotect all sheets at once — ⭐⭐⭐ · Grade A — **W6** `workbook.protect_sheets` + `workbook.unprotect_sheets`
One password across the whole book. Two actions, since a single toggle button hides which way it will go.

### Create a Table of Contents with hyperlinks — ⭐⭐⭐ · Grade A — **W6** `workbook.build_toc`
Auto-built index linking to every tab.

### Unhide all rows & columns — ⭐⭐ · Grade B — **W6** `workbook.unhide_rows_cols`
Reset visibility across the sheet.

### Sort worksheets alphabetically — ⭐⭐ · Grade B — **Later**
Reorder tabs.

### Hide all sheets except the active one — ⭐⭐ · Grade B — **Deferred**
Focus/presentation mode. *Blocked on:* destructive to workbook state with no preview and no one-click undo.

### Add a "Back to TOC" link on every sheet — ⭐⭐ · Grade B — **Later**
Companion to the TOC macro.

### Delete all empty sheets — ⭐⭐ · Grade C — **Deferred**
Cleanup pass. *Blocked on:* destructive with no preview; "empty" is easy to get wrong.

### Jump to next / previous sheet — ⭐ · Grade C — **Declined**
Minor navigation convenience. *Declined:* no value as a generated button; Ctrl+PgUp/PgDn already does it.

---

## 9. Protect & Back Up — safety & versioning

### Timestamped backup copy — ⭐⭐⭐⭐ · Grade B — **W5** `backup.timestamped_copy`
`SaveCopyAs` into a Backup folder with a datetime stamp.

### Snapshot a sheet / range before a risky operation — ⭐⭐⭐⭐ · Grade C — **W5** `backup.snapshot`
Auto-save a restore point before Import/Clear runs. A strong safety add for this library — and the unblocker for several *Deferred* destructive items above, since it supplies the missing undo story.

### Save a month-end versioned file — ⭐⭐⭐ · Grade B — **W5** `backup.month_end_version`
Prior-month-end naming for finance packs.

### Move / archive a generated file to a dated folder — ⭐⭐⭐ · Grade B — **W5** `backup.archive_output`
Relocate outputs into `Archive\yyyy_mm\`.

---

## 10. Run Faster / Run Automatically — engine & events

### Speed wrapper (ScreenUpdating / Calculation / Events off → on) — ⭐⭐⭐⭐⭐ · Grade B ✅ DONE (Core helper)
Every button wraps work in the manual-calc / no-events / no-repaint pattern with guaranteed restore in the error handler. Now Core's `Core_SpeedOn` / `Core_SpeedOff` / `RestoreSpeed`.

### Robust last-row / last-column finder — ⭐⭐⭐⭐ · Grade A ✅ DONE (Core helper)
`End(xlUp)` / `End(xlToLeft)` + `SpecialCells` fallback. Now Core's `Core_LastRow` / `Core_LastRowUp` / `Core_LastCol`.

### Structured error handler + status log — ⭐⭐⭐⭐ · Grade C ✅ DONE (Core helper)
`On Error GoTo Fail`, state restore, and a run summary. Now Core's `Core_Fail` / `Core_Done` / `Core_Log`, with the quiet-inside-a-chain convention.

### Loop a folder of workbooks and process each — ⭐⭐⭐⭐ · Grade B — **Deferred**
Batch-run any of the above across many files. *Blocked on:* a defined per-file continue/abort contract and failure reporting. Revisit once chains have proven their Continue/Stop model in real use.

### Timestamp a row on edit (Worksheet_Change) — ⭐⭐⭐ · Grade A — **Declined**
Auto-stamp when a row is entered/changed. *Declined:* requires an event handler in a sheet or `ThisWorkbook` module, which a single imported standard-module bundle cannot install. Same architectural reason as the active-row spotlight.

---

## Notes for the build

Look at how many items are the *same verb with a different parameter*: **export** (PDF vs xlsx vs per-sheet vs per-row), **combine** (sheets vs workbooks), **clean** (blanks / dupes / spaces / case), **highlight** (dupes / outliers / missing).

**How this is reconciled with "every action individually listed":** the collapsing happens *underneath*, never in the list. Each listed action keeps its own permanent public `Sub`, manifest entry, id and button; variants that share real work delegate to one private, package-prefixed family engine (`PkgImport_CombineEngine` serving both combines, for example). An engine is capped at one coherent workflow and split when it starts branching on unrelated concerns. The user sees sixty distinct actions; the tested surface underneath stays small.

The ✅ Core helpers (speed wrapper, last-row finder, error handler, type coercion) are the shared plumbing every engine calls, and they live only in Core — the flat single-module namespace makes cross-package sharing impossible by design.
