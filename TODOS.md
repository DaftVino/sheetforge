# TODOs

Local task list until this repo gets a GitHub Project (repo-standards 6.5). Tasks
touching Core or more than one package go through `writing-plans` first. Verify =
`pwsh build/build.ps1`; there is no headless VBA test, so also import
`dist/MacrosBundle.bas` and run the workbook, recorded in the PR/notes.

Current: **v0.2.1** on `main` (config-driven engine + PkgImportExport + PkgClean +
build script + full Setup UX). Repo is **local-only** (no git remote yet).

## Before the remote (blocker)

- [ ] **Sanitize real store/entity specifics** before any push to
      `github.com/DaftVino/sheetforge` (currently private, nothing pushed yet).
      Scrub franchise/store/entity names and real prefixes from:
      `src/packages/PkgImportExport.bas` manifest defaults (export tabs, import
      prefixes, master columns) and `src/legacy/SetupImportExportMacros.bas`.
      Decide: neutral demo defaults vs. a separate private config. Then
      `git remote add origin ...` and push.

## Next

- [ ] **Freeze `spec/contract.md`** — the language-neutral config schema + action
      ids (field keys, kinds, tiers; action ids and semantics) so the Apps Script
      port is mechanical. Last item of the engine-architecture-draft build order
      before porting.

## Roadmap

- [ ] **Apps Script port** — implement the frozen spec on Google Sheets (Apps
      Script), same Setup layout + action ids, different engine. Only after the
      spec is frozen.
- [ ] **More packages** as needed from `docs/macro-catalog-plug-and-play.md`
      (Reconcile / Validate, Format & Present, Summarize, Workbook housekeeping,
      Backup). Each is a thin `Pkg<Name>.bas` + one line in Core's candidate list.
- [ ] Refresh `docs/code-map.md` (`/code-map`) once `src/` grows further.

## Housekeeping / minor

- [ ] Delete merged local branches: `feat/core-import-export`,
      `feat/setup-two-tier`, `feat/setup-band-polish`.
- [ ] Setup merged-band row heights are an estimate (AutoFit ignores merged
      cells). If a tinted explanation ever clips its last line, lower the
      chars-per-line constant in `Core.TintRow` (currently 115).
- [ ] Consider whether the type-map source-location (same folder, newest file)
      should become a configurable/asked option rather than a fixed rule
      (raised during UX review; deferred).
