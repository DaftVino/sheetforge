# TODOs

Local task list until this repo gets a GitHub Project (repo-standards 6.5). Tasks
touching Core or more than one package go through `writing-plans` first. Verify =
`pwsh build/build.ps1`; there is no headless VBA test, so also import
`dist/MacrosBundle.bas` and run the workbook, recorded in the PR/notes.

Current: **v0.2.1** on `main` (config-driven engine + PkgImportExport + PkgClean +
build script + full Setup UX). Remote configured: `origin` =
`github.com/DaftVino/sheetforge` (private); `main` pushed at the sanitized tip.

## Before the remote (blocker) — DONE (2026-07-24)

- [x] **Sanitize real store/entity specifics** and publish. Scrubbed all 7
      tracked files to fully-generic demo defaults (design + plan:
      `docs/designs/sanitization-pass.md`, `-plan.md`); genericized the real
      store number `3176` → `12345`; renamed legacy module `FoodDashMacros` →
      `SetupImportExportMacros`. `main` fast-forwarded and pushed to
      `origin`. **Caveat (tree-only choice):** the pre-`d29b0d7` *history* still
      contains the old jargon — the remote is private, so this was accepted;
      revisit (squash/filter history) before any public release.

## Next — active

- [ ] **"Compose your buttons" feature** — spec (`brainstorming` → `writing-plans`),
      then build. ONE spec covering two linked Setup-page capabilities the user
      asked for (2026-07-24):
      - **(1) Grouping / layout** — a Setup section to assign actions to
        pages/sections. This *finishes the long-pending handoff **step 7***: wire
        the `src/packages/MacroLayout.bas` prototype — today a standalone,
        **un-built** test module (excluded from the build; buttons call a
        placeholder `ML_RunById`) — onto the real `Core_Run` + the Setup page.
        Grouping was always intended (`docs/handoff.md`); only its integration is
        unfinished, and its blocker (`Core_Run`) now exists.
      - **(2) Chaining** — a separate Setup section to compose several actions
        into ONE button that runs them **sequentially** (generalizes the
        hardcoded clear→import inside `PkgImportExport_Import`). NEW scope:
        user-defined chains are deployment-specific **config**, not manifest data
        — a new config entity Core hasn't modeled before.
      Touches Core (`Core_Run` dispatcher, button builder, Setup renderer) + the
      contract → additive **contract v1.1.0** (composite/chained-action +
      button-group concepts). Open design Qs: where chains live (config vs
      manifest), mid-chain error handling (stop vs continue), whether chained
      steps suppress their individual dialogs, adopt-vs-rewrite MacroLayout.
      **Catalog build-out (Roadmap, below) is a SEPARATE track — not this spec.**

## Done (recent)

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

## Housekeeping / minor

- [x] Delete merged local branches (`feat/*`, `chore/sanitize-defaults`,
      `docs/freeze-contract`) — DONE 2026-07-24; only `main` remains.
- [ ] Setup merged-band row heights are an estimate (AutoFit ignores merged
      cells). If a tinted explanation ever clips its last line, lower the
      chars-per-line constant in `Core.TintRow` (currently 115).
- [ ] Consider whether the type-map source-location (same folder, newest file)
      should become a configurable/asked option rather than a fixed rule
      (raised during UX review; deferred).
