# Macro Catalog — Plug & Play Library

A master list of candidate macros for the config-driven macro project, grouped by **what the user is trying to get done** (not by how the code works). The User-Input Config driver is the backbone that selects and parameterizes everything below, so it is not listed as an item.

## How to read this

- **⭐ (1–5)** — how important I think it is to *include this in the project*. 5 = core, must-have; 1 = nice-to-have / gimmick.
- **Grade (A–F)** — how *well-used / downloaded / requested* the macro is in the real world (forums, code libraries, Stack Overflow). A = everyone wants it; F = niche.
- **✅ DONE** — already implemented in `FoodDashMacros.bas`. `(embedded)` means it lives inside one of the four button routines rather than as its own selectable macro — worth exposing as a standalone in the library.

The two ratings are independent on purpose: a few things are hugely popular but low-value here (row spotlight), and a few are lower-profile but critical to *your* finance work (reconciliation, type coercion).

---

## 1. Bring Data In — import & combine sources

### Import all files from a folder into their tabs (type-aware) — ⭐⭐⭐⭐⭐ · Grade C ✅ DONE
Matches source files by name prefix (handles ` (1)`/` (2)` suffixes), clears the target, pastes, and coerces column types. Your `ImportSourceData`.

### Coerce imported text to real types (T / N / D / DT / B / G) — ⭐⭐⭐⭐⭐ · Grade C ✅ DONE (embedded)
Forces text-based exports back to numbers/dates/booleans so downstream formulas keep matching. Your `CoerceVal`. Rare as a shared snippet, but it's the reason your imports don't silently break.

### Find & open the newest file matching a prefix — ⭐⭐⭐⭐ · Grade B ✅ DONE (embedded)
Safe folder scan (collects names first, then stats — avoids the `Dir()` iterator corruption bug). Your `FindSourceFile`.

### Combine all worksheets into one master sheet — ⭐⭐⭐⭐⭐ · Grade A
Stacks every tab into a single table, aligning headers. One of the most-requested macros anywhere.

### Combine all workbooks in a folder into one — ⭐⭐⭐⭐⭐ · Grade A
Loops a folder, appends each file's data. Universal for consolidations and rollups.

### Import CSV / TXT with delimiter parsing — ⭐⭐⭐⭐ · Grade A
Bring in a delimited file with explicit qualifiers and type handling.

### Split one column into many (Text to Columns) — ⭐⭐⭐ · Grade A
For exports that land everything in column A.

### Pull data from a closed workbook — ⭐⭐⭐ · Grade B
Read a range without opening/showing the source.

---

## 2. Clean & Standardize — fix messy data

### Clear data rows, keep headers + formatting — ⭐⭐⭐⭐⭐ · Grade B ✅ DONE
Wipes row 2 down on the listed tabs, preserves headers/formats. Your `ClearInputTabs`.

### Remove blank rows / columns — ⭐⭐⭐⭐⭐ · Grade A
Delete empty rows/cols in the used range.

### Remove duplicates — ⭐⭐⭐⭐⭐ · Grade A
De-dupe on user-chosen key columns.

### Trim spaces & strip non-printing characters — ⭐⭐⭐⭐⭐ · Grade A
`TRIM` + `CLEAN` across a selection; kills the invisible junk that breaks lookups.

### Fill blanks with the value above — ⭐⭐⭐⭐ · Grade A
Repairs "group label only on first row" exports.

### Unmerge cells and fill down the value — ⭐⭐⭐⭐ · Grade B
Turns merged report layouts back into analyzable data.

### Standardize number / date formats — ⭐⭐⭐⭐ · Grade B
Apply consistent formats to key columns after transfer.

### Change case (UPPER / lower / Proper) — ⭐⭐⭐ · Grade A
Normalize text casing across a range.

### Convert text-numbers to real numbers — ⭐⭐⭐ · Grade B
Strip `$`, `,`, `%`, leading apostrophes; re-type as numeric.

### Delete all hidden rows / columns — ⭐⭐ · Grade C
Remove filtered-out or hidden data before sharing.

---

## 3. Validate & Reconcile — catch problems before they ship

### Reconcile two datasets on a key, flag mismatches — ⭐⭐⭐⭐⭐ · Grade C
Match A vs B on an ID, report missing/extra/differing rows. High value for finance close; under-shared as a generic snippet.

### Data-validation scan (blanks / type / out-of-range) → report — ⭐⭐⭐⭐ · Grade C
Sweep a range against rules and write an exceptions list.

### Highlight duplicates — ⭐⭐⭐⭐ · Grade A
Color-flag repeated keys without deleting.

### Highlight outliers / out-of-threshold values — ⭐⭐⭐ · Grade C
Flag values past a parameterized bound (your dashboard already leans on this idea via CF).

### Flag missing required fields — ⭐⭐⭐ · Grade C
Mark rows missing a mandatory column.

---

## 4. Match & Enrich — look things up

### Refresh a reference/master table from an external file (column-mapped) — ⭐⭐⭐⭐⭐ · Grade C ✅ DONE
Pulls named columns from a master file in header order, keeps the key as text. Your `UpdateRefMaster`.

### VLOOKUP / XLOOKUP-style join in code — ⭐⭐⭐⭐ · Grade B
Fast in-memory join over large sets (dictionary-based), avoids volatile lookup formulas.

### Add mapped columns from a lookup table — ⭐⭐⭐ · Grade B
Enrich rows with attributes pulled from a reference sheet.

---

## 5. Format & Make Presentable

### Conditional highlighting by rule / parameter — ⭐⭐⭐⭐ · Grade A
Rules-driven coloring beyond native CF, controlled from config.

### Style header row + autofit columns — ⭐⭐⭐⭐ · Grade A
Bold/fill/font the header, size every column. The classic "make an export readable" macro.

### Add & style a report title (merge + center) — ⭐⭐⭐ · Grade B
Drop a formatted title block on top of a report.

### Apply a house table style (banding / borders) — ⭐⭐⭐ · Grade B
One-click brand/format standard for any table.

### Print-layout prep (landscape / fit-to-width / margins) — ⭐⭐⭐ · Grade B
Board-pack and period-end binder formatting.

### Highlight cells with comments / formulas / blanks — ⭐⭐ · Grade A
Popular audit-view helper; light value in a driven library.

### Spotlight the active row / column — ⭐⭐ · Grade A
Very popular download, but mostly cosmetic.

---

## 6. Summarize & Analyze

### Refresh all pivot tables — ⭐⭐⭐⭐ · Grade A ✅ DONE (embedded)
Loops every sheet and refreshes each pivot after import. Inside your `ImportSourceData`.

### Force full recalculation — ⭐⭐⭐ · Grade B ✅ DONE (embedded)
`CalculateFull` so downstream tabs update. Inside Import/Export.

### Build a summary / pivot from a raw range — ⭐⭐⭐⭐ · Grade B
Generate a summary table from chosen dimensions + measures.

### Copy a pivot to a static values sheet — ⭐⭐⭐ · Grade B
Freeze pivot output as values + formatting for distribution.

### Apply a consistent style to all charts — ⭐⭐ · Grade B
Enforce one visual standard across chart objects.

---

## 7. Get It Out — export & share

### Export a values-only snapshot (.xlsx) — ⭐⭐⭐⭐⭐ · Grade B ✅ DONE
Copies main tabs to a new file, flattens spills + legacy arrays + formulas to values, keeps formatting/CF, breaks external links, dates the filename. Your `ExportDashboard` — the most sophisticated routine in the file.

### Convert formulas to values — ⭐⭐⭐⭐ · Grade A ✅ DONE (embedded)
Freeze formula outputs (incl. dynamic-array spills). Inside `ExportDashboard`.

### Dated / timestamped output filename — ⭐⭐⭐⭐ · Grade B ✅ DONE (embedded)
`yyyy-mm-dd` + base name with a configurable day offset. Inside `ExportDashboard`.

### Export sheet / range / workbook to PDF — ⭐⭐⭐⭐⭐ · Grade A
One-click PDF of a range, sheet, or whole book.

### Save each worksheet as its own file / PDF — ⭐⭐⭐⭐ · Grade A
Explode a workbook into per-tab deliverables.

### Save each filtered group (or row) as its own file — ⭐⭐⭐ · Grade B
Split by a column value into separate files.

### Email a sheet / range as an attachment (Outlook) — ⭐⭐⭐⭐ · Grade A
Send the current output straight from Excel.

### Send bulk personalized emails from a list — ⭐⭐⭐ · Grade A
Mail-merge-style send driven by a table.

---

## 8. Move Around & Housekeeping — workbook management

### Unhide all worksheets — ⭐⭐⭐ · Grade A
The single most-downloaded VBA snippet in existence.

### Protect / unprotect all sheets at once — ⭐⭐⭐ · Grade A
One password across the whole book.

### Create a Table of Contents with hyperlinks — ⭐⭐⭐ · Grade A
Auto-built index linking to every tab.

### Unhide all rows & columns — ⭐⭐ · Grade B
Reset visibility across the sheet.

### Sort worksheets alphabetically — ⭐⭐ · Grade B
Reorder tabs.

### Hide all sheets except the active one — ⭐⭐ · Grade B
Focus/presentation mode.

### Add a "Back to TOC" link on every sheet — ⭐⭐ · Grade B
Companion to the TOC macro.

### Delete all empty sheets — ⭐⭐ · Grade C
Cleanup pass.

### Jump to next / previous sheet — ⭐ · Grade C
Minor navigation convenience.

---

## 9. Protect & Back Up — safety & versioning

### Timestamped backup copy — ⭐⭐⭐⭐ · Grade B
`SaveCopyAs` into a Backup folder with a datetime stamp.

### Snapshot a sheet / range before a risky operation — ⭐⭐⭐⭐ · Grade C
Auto-save a restore point before Import/Clear runs. A strong safety add for this library.

### Save a month-end versioned file — ⭐⭐⭐ · Grade B
Prior-month-end naming for finance packs.

### Move / archive a generated file to a dated folder — ⭐⭐⭐ · Grade B
Relocate outputs into `Archive\yyyy_mm\`.

---

## 10. Run Faster / Run Automatically — engine & events

### Speed wrapper (ScreenUpdating / Calculation / Events off → on) — ⭐⭐⭐⭐⭐ · Grade B ✅ DONE (embedded)
Every button wraps work in the manual-calc / no-events / no-repaint pattern with guaranteed restore in the error handler. Should be a shared helper in the library, not copy-pasted per macro.

### Robust last-row / last-column finder — ⭐⭐⭐⭐ · Grade A ✅ DONE (embedded)
`End(xlUp)` / `End(xlToLeft)` + `SpecialCells` fallback. Used throughout; worth exposing as a shared utility.

### Structured error handler + status log — ⭐⭐⭐⭐ · Grade C ✅ DONE (embedded)
`On Error GoTo Fail`, state restore, and a run summary `MsgBox`. Standardize it as the library's error pattern.

### Loop a folder of workbooks and process each — ⭐⭐⭐⭐ · Grade B
Batch-run any of the above across many files.

### Timestamp a row on edit (Worksheet_Change) — ⭐⭐⭐ · Grade A
Auto-stamp when a row is entered/changed.

---

## Notes for the build

Look at how many items are the *same verb with a different parameter*: **export** (PDF vs xlsx vs per-sheet vs per-row), **combine** (sheets vs workbooks), **clean** (blanks / dupes / spaces / case), **highlight** (dupes / outliers / missing). Those collapse into a handful of parameterized routines the config driver dispatches — one Export engine, one Combine engine, one Clean engine — rather than 60 separate macros. The ✅ embedded helpers (speed wrapper, last-row finder, error handler, type coercion) are the shared plumbing every engine should call.
