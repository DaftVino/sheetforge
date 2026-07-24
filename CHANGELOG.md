# Changelog

All notable changes to this project are documented here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versioning: [SemVer](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
