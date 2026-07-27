# Changelog

All notable changes to this project are documented here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versioning: [SemVer](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.4.0] - 2026-07-27

### Added

- **Anchored placements — put a button on one of your own sheets.** A placement code
  of the form `Sheet!Cell` (e.g. `MIC!B4`) draws that button on your tab at that cell
  instead of on a generated page. `$B$4`, unquoted names containing spaces, and the
  quoted `'MIC Data'!B4` form are all accepted. The button uses `xlMove`, so it travels
  with its cell and is never resized by a narrow column. An anchored button is not on a
  generated page and never affects page packing, column caps or category splitting.
  Finalize **blocks** on a malformed code, an engine-owned target (`Macros 1`), a
  missing sheet, or two actions aimed at one cell — including the same cell spelled two
  ways; a valid anchor is reported in a soft `Placed on your own sheets` list. A
  placement that fails at build time (a protected sheet, a target deleted after
  activation) reports a count and a repair instruction instead of taking down the build.
- **Three-state per-action availability.** The placement block is now the **Action
  catalog** (`Action | Availability | Placement`) with a validated dropdown: `Button`
  (shown and runnable), `Chains only` (no button, runs only as a chain step), `Off`
  (no button, runs nowhere). `Off` means off — the dispatcher refuses it and Finalize
  blocks activation for any chain that uses one. An unrecognised value blocks Finalize
  rather than defaulting to `Button`.
- **`usedBy` on a field** — settings belonging only to `Off` actions collect into one
  collapsed *"settings for actions you haven't turned on"* band at the bottom of Setup.
  Relegated, never omitted: Setup re-seeds from the finalized `_Config`, so omitting
  them would cost a second Finalize and discard un-finalized typing. A `Chains only`
  action still runs, so it keeps its settings in the main list.
- **`category` on an action** — optional, defaults to the package name, and decoupled
  from the module, so two packages can feed one category and a future package split
  cannot move a button in the user's mental model.
- **Capacity-packed auto-layout.** With no placement codes filled in, buttons
  distribute automatically across numbered pages: up to three category-headed columns
  per page, up to 6 buttons per column, small categories sharing a page, and a category
  never split across tabs unless it needs more columns than a page holds. Filling in
  even one code still hands total control back to the codes.
- **Legible descriptions.** Button pages are an **unmerged cell-card grid** — the
  button anchored to a content cell, its description in the cell beneath, that row
  auto-fitted. Grouped pages previously had no descriptions at all, and merged cells
  (which Excel's AutoFit ignores) are forbidden on the grid so nothing clips.
- **A machine-owned config schema marker**, with **fail-closed** migration. A v0.3.0
  config is detected and its page codes read from the column they are actually in;
  every recovered code is validated; on any conflict `Create Setup Tab` refuses with a
  report instead of rebuilding. A marker written by a newer major version refuses
  outright rather than offering a rebuild, and a damaged one asks first, spelling out
  what would be discarded and what kept, defaulting to **No**.
- **`Does the range you pick include a header row?`** on Clean — the contract's first
  use of the `bool` kind, so its `yes`/`no` encoding is now fixed. A blank cell reads as
  the field's declared default; an unrecognised value blocks Finalize rather than
  silently reading as `No` and deleting a header row.
- Contract frozen at **v1.2.0** (additive).
- **Build / tooling:** a reserved-word gate rejects a VBA keyword used as an identifier.
  The list is empirical — a word is added when Excel actually rejects one, never on
  suspicion. Fixture and harness packages are kept out of the default bundle.

### Changed

- The placement column is labelled **`Placement`** rather than `Page code`, since a
  code can now be a page code, an anchor, or `hide`.

### Removed

- **The single `Macros` tab default.** With no placement codes, v0.3.0 built one tab
  and stacked every button on it; v0.4.0 packs categories onto numbered pages instead,
  and the stale `Macros` tab is deleted on rebuild. Manual page codes are unaffected.

### Fixed

- Every range named in a Clean completion message is **sheet-qualified** — `MIC!M2:M12`,
  and `'MIC Data'!M2:M12` when the name needs quoting — so a message can no longer read
  as though the action ran on whichever sheet you happen to be looking at.
- **A failed button placement no longer leaves an orphan shape.** `Buttons.Add` creates
  the shape before it can be named, so a failure part-way through placement left a
  button carrying Excel's default name, no caption and no `OnAction` — and permanently
  unremovable, because the sweep matches on the engine's name prefix. Placement is now
  all-or-nothing: the partial shape is deleted and the original error re-raised.
- **Finalize no longer claims the Macros pages were rebuilt when none were built.**
  Every button hidden, switched `Off`, or anchored is a legitimate configuration, and
  the dialog now says so instead of sending you looking for tabs that were never meant
  to exist.
- Instructions that told you to run `Core_Run` or `Core_BuildControlPanel` from **Alt+F8**
  are corrected throughout. Excel's Macro dialog lists only `Public Sub`s taking no
  arguments, and an *optional* argument is enough to hide one; use the VBE Immediate
  window (Alt+F11, Ctrl+G).

## [0.3.0] - 2026-07-24

### Added

- **Compose Your Buttons** — a Core-owned Setup section with two capabilities: **Chains** (compose several actions into one button that runs them in order — steps authored as `Label [action.id]` tokens; each chain is an immutable `chain.<id>` with a freely-editable label; on a step failure the runner prompts Continue/Stop; per-step success dialogs are suppressed and one summary is shown) and **Grouping/layout** (assign each button a page/column code like `1A`; no codes → the single "Macros" tab as before; any code → numbered "Macros N" pages, with blank-coded buttons in page 1's "Unassigned" area and `hide` to omit). Finalize validates chains + placement and blocks activation on structural errors (unknown steps, bad codes, reserved-namespace or id collisions).
- **Setup Menu tab** — the Create Setup Tab / Finalize Setup Tab controls now live on their own dedicated tab (built by `Core_BuildControlPanel`) so they can be hidden when not in use; rebuilding reuses the tab and preserves its hidden state.
- **Finalize rebuilds the control panel** — activating a configuration now rebuilds the Macros tab automatically, so authored chains and grouping appear immediately (no manual `Core_BuildControlPanel` run).
- Contract frozen at **v1.1.0** (additive): the `chain` config entity, reserved `chain.*` namespace, engine-owned config, placement code, and the dispatch/quiet convention.
- Package completion dialogs now route through `Core_Done` so they suppress inside a chain.
- **Build / tooling:** `build/build.ps1` now fails on a module-level declaration placed after the first procedure (a VBA compile error the block-balance check could not see); `scripts/vba-syntax-check.mjs` drives the "VBA Pro" language server headlessly for a fast pre-Excel syntax gate.

### Changed

- The Create / Finalize admin controls moved off the Macros tab (and out of the grouped page-1 "Setup" column) onto the dedicated, hideable **Setup Menu** tab.

### Fixed

- Re-running **Finalize** no longer errors *"That name is already taken."* and no longer leaves orphan `Setup (2)`/`(3)` copies: the very-hidden `_Config` is unhidden before it is deleted, so the copy→rename cannot collide. (Pre-existing bug; surfaced by the multi-finalize Compose workflow.)
- `DeletePanels` now guarantees `Application.DisplayAlerts` is restored even if a sheet delete raises, so a leaked `DisplayAlerts = False` can no longer suppress later confirmation prompts.

### Removed

- The standalone `MacroLayout.bas` grouping prototype (its logic is absorbed into Core).

## [0.2.1] - 2026-07-23

### Fixed

- Setup tinted band rows (section headers + list/table label & explanation) now merge across columns B:E and wrap to the full band width instead of being cramped in column B. Merged-row heights are set explicitly since AutoFit ignores merged cells.
- A section header identical to its package banner is no longer rendered (removed the duplicate "CLEAN & STANDARDIZE" in PkgClean; Core now skips any section whose text equals the package banner).

## [0.2.0] - 2026-07-23

### Added

- **Two-tier Setup**: manifest fields declare a tier ("simple" default / "advanced"); the Setup renderer collects advanced fields under one "ADVANCED - most users can leave these as-is" band per package, so regular users can stop after the simple fields. An unknown/mis-cased tier degrades to "simple" (never hidden).
- **Suggest Import Types** helper: a button that reads the live Setup import table, opens each matched source file, samples it, and fills a suggested per-column type map (T/N/D) for review (never applied silently; ID/code columns flagged for the user to keep as text). Adds two reusable Core setup-helper services, `Core_SetupTable` / `Core_SetupSetColumn`, that read/write the live Setup sheet before Finalize.
- **Per-field help notes**: manifest fields gain an optional plain-language `help` string, rendered grey/italic on the Setup sheet.

### Changed

- Setup labels and table headers reworded for non-technical users ("Input tab (in this workbook)", "Source file (name starts with...)", "Column types"); the import note now explains the same-folder rule, the ignored `(1)`/`(2)` copy suffix, newest-file logic, and the full type-map legend.
- Macro-tab button descriptions span two rows (no longer clipped); list/table label + help render in the tinted band above the column headers so the headers sit directly on the input cells.

## [0.1.0] - 2026-07-23

### Added

- Initial scaffold (office-automation profile, source-first VBA workflow per ADR 0001).
- **Core engine** (`src/core/Core.bas`): manifest-building API (`Core_DeclarePackage`/`AddSection`/`AddField`/`AddList`/`AddTable`/`AddDefault`/`AddAction`); config readers with graceful manifest-default fallback (`Core_Cfg`/`CfgLong`/`CfgList`/`CfgTable`); shared helpers (depth-counted speed wrapper, robust + bare last-row/col, file finder, header-column finder, type coercion, structured error/log); zero-trust package discovery via `Application.Run` (no "Trust access to the VBA project object model"); Setup-sheet renderer, control-panel/button builder, and the `Core_Run "<actionId>"` dispatcher.
- **PkgImportExport** package (`src/packages/PkgImportExport.bas`): the Sample Dashboard's four buttons - Clear Input Tabs, Import Source Data, Update Ref_Master, Export Dashboard - re-expressed as a thin package on Core with no behaviour change (Import now shows one dialog instead of two).
- **PkgClean** package (`src/packages/PkgClean.bas`): Remove Duplicates, Trim & Clean Text, Fill Blanks Down, driven by an ad-hoc range picker; preserves formula cells and rejects multi-area selections.
- **Build** (`build/build.ps1`): concatenates `src/core/Core.bas` + `src/packages/Pkg*.bas` into a single-module `dist/MacrosBundle.bas`; fails on a duplicate procedure name across modules or a Sub/Function block imbalance. `dist/` is a gitignored build target.
- `docs/designs/setup-simplicity.md`: planned (deferred) two-tier Simple/Advanced Setup and a suggest-column-types helper.
