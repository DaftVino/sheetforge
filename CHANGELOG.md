# Changelog

All notable changes to this project are documented here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versioning: [SemVer](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-07-23

### Added

- Initial scaffold (office-automation profile, source-first VBA workflow per ADR 0001).
- **Core engine** (`src/core/Core.bas`): manifest-building API (`Core_DeclarePackage`/`AddSection`/`AddField`/`AddList`/`AddTable`/`AddDefault`/`AddAction`); config readers with graceful manifest-default fallback (`Core_Cfg`/`CfgLong`/`CfgList`/`CfgTable`); shared helpers (depth-counted speed wrapper, robust + bare last-row/col, file finder, header-column finder, type coercion, structured error/log); zero-trust package discovery via `Application.Run` (no "Trust access to the VBA project object model"); Setup-sheet renderer, control-panel/button builder, and the `Core_Run "<actionId>"` dispatcher.
- **PkgImportExport** package (`src/packages/PkgImportExport.bas`): the Sample Dashboard's four buttons - Clear Input Tabs, Import Source Data, Update Ref_Master, Export Dashboard - re-expressed as a thin package on Core with no behaviour change (Import now shows one dialog instead of two).
- **PkgClean** package (`src/packages/PkgClean.bas`): Remove Duplicates, Trim & Clean Text, Fill Blanks Down, driven by an ad-hoc range picker; preserves formula cells and rejects multi-area selections.
- **Build** (`build/build.ps1`): concatenates `src/core/Core.bas` + `src/packages/Pkg*.bas` into a single-module `dist/MacrosBundle.bas`; fails on a duplicate procedure name across modules or a Sub/Function block imbalance. `dist/` is a gitignored build target.
- `docs/designs/setup-simplicity.md`: planned (deferred) two-tier Simple/Advanced Setup and a suggest-column-types helper.
