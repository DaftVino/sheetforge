# TODOs

Local task list until this repo gets a GitHub Project (repo-standards 6.5). Tasks
touching Core or more than one package go through `writing-plans` first. Verify =
`pwsh build/build.ps1`; there is no headless VBA test, so also import
`dist/MacrosBundle.bas` and run the workbook, recorded in the PR/notes.

Current: **v0.2.1** on `main`, **Compose Your Buttons (chains + grouping,
contract v1.1.0) shipped on `feat/compose-your-buttons`, pending merge**
(config-driven engine + PkgImportExport + PkgClean + build script + full Setup
UX). Remote configured: `origin` = `github.com/DaftVino/sheetforge` (private);
`main` pushed at the sanitized tip.

## Before the remote (blocker) — DONE (2026-07-24)

- [x] **Sanitize real store/entity specifics** and publish. Scrubbed all 7
      tracked files to fully-generic demo defaults (design + plan:
      `docs/designs/sanitization-pass.md`, `-plan.md`); genericized the real
      store number `12345` → `12345`; renamed legacy module `SetupImportExportMacros` →
      `SetupImportExportMacros`. `main` fast-forwarded and pushed to
      `origin`. **Caveat (tree-only choice):** the pre-`d29b0d7` *history* still
      contains the old jargon — the remote is private, so this was accepted;
      revisit (squash/filter history) before any public release.

## Next — active

Nothing active right now — next up is Roadmap, below.

## Done (recent)

- [x] **"Compose your buttons" feature** — DONE (2026-07-24, contract **v1.1.0**).
      One Core-owned Setup section, two capabilities: **Chains** (compose several
      actions into one button that runs them in order; `Label [action.id]` step
      tokens; immutable `chain.<id>` with an editable label; Continue/Stop prompt
      on a step failure; per-step dialogs suppressed, one summary shown) and
      **Grouping/layout** (per-button `1A`-style page/column codes; no codes →
      the single "Macros" tab; any code → numbered "Macros N" pages, blank →
      page 1 "Unassigned", `hide` → omitted; Setup/Finalize admin controls always
      on page 1). Finalize validates chains + placement and blocks activation on
      structural errors. This finishes the long-pending handoff **step 7**
      (grouping integration) and absorbs `src/packages/MacroLayout.bas`'s logic
      into Core; the standalone prototype file is deleted (Task 10).

- [x] **Freeze `spec/contract.md`** — DONE (2026-07-24, contract **v1.0.0**).
      Language-neutral SoT: manifest object, field model, canonical kind
      vocabulary (`range` excluded as never-implemented, `bool` reserved), tier
      rules, config storage/read + graceful-fallback, action contract, the frozen
      registry for both packages, portability appendix, SemVer change-control.
      Reviewed via `/review` (accuracy-audited against the code; 4 fixes applied).
      Next roadmap item is the Apps Script port.

## Roadmap

- [ ] **Apps Script port** — implement the frozen spec on Google Sheets (Apps
      Script), same Setup layout + action ids, different engine. Only after the
      spec is frozen.
- [ ] **Catalog build-out** (separate track from "Compose your buttons" above) —
      grow toward the ~60 selectable macros in `docs/macro-catalog-plug-and-play.md`
      (Reconcile / Validate, Format & Present, Summarize, Workbook housekeeping,
      Backup). Only 2 of ~60 are built (PkgImportExport, PkgClean); the rest were
      always intended as user setup options, deferred by the §10 build order (now
      complete). Each is a thin `Pkg<Name>.bas` + one line in Core's candidate list.
- [ ] Refresh `docs/code-map.md` (`/code-map`) once `src/` grows further.
- [ ] **Built-page polish (deferred, design §8)** — the grouped "Macros N" pages
      are functional but plain: per-page titles/column headings/legend,
      Previous/Next page navigation, a live placement preview during Setup, and
      a visually distinct admin area (Setup/Finalize controls) on page 1.
      Explicitly out of scope for "Compose your buttons" (Task 10).

## Housekeeping / minor

- [x] Delete merged local branches (`feat/*`, `chore/sanitize-defaults`,
      `docs/freeze-contract`) — DONE 2026-07-24; only `main` remains.
- [ ] Setup merged-band row heights are an estimate (AutoFit ignores merged
      cells). If a tinted explanation ever clips its last line, lower the
      chars-per-line constant in `Core.TintRow` (currently 115).
- [ ] Consider whether the type-map source-location (same folder, newest file)
      should become a configurable/asked option rather than a fixed rule
      (raised during UX review; deferred).
