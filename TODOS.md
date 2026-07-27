# TODOs

Local task list until this repo gets a GitHub Project (repo-standards 6.5). Tasks
touching Core or more than one package go through `writing-plans` first. Verify =
`pwsh build/build.ps1`; there is no headless VBA test, so also import
`dist/MacrosBundle.bas` and run the workbook, recorded in the PR/notes.

Current: **v0.4.0** (Finish Excel track 1 — packed auto-layout, three-state
availability, `Sheet!Cell` anchored placements, fail-closed config migration,
contract **v1.2.0** — owner-verified in Excel across three workbooks). Remote:
`origin` = `github.com/DaftVino/sheetforge` (private). Next is **track 2**, the
catalog build-out in waves (`docs/macro-catalog-plug-and-play.md`), and the
**Setup-surface redesign** designed in
`docs/designs/2026-07-26-setup-surface-progressive-disclosure-design.md`. The
Apps Script port is deferred to last, deliberately.

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

- [ ] **Setup-surface progressive disclosure** — designed in
      `docs/designs/2026-07-26-setup-surface-progressive-disclosure-design.md`,
      out of the v0.4.0 owner gate. Two separable parts. **Part B** (small,
      package-level, do first): move `clean.dupes`' two settings — the header
      question and the key columns — off Setup and into a prompt at run time,
      because both describe the range you just picked, not the deployment.
      **Part A** (Core, contract, migration): Action catalog to the top of Setup,
      actions default `Off`, `usedBy`-driven hide/unhide, the Chains block
      appearing on demand, and a pending-changes formula plus Refresh button —
      click-driven, because a live handler would need VBIDE trust (banned) or a
      class module (the bundle is one `.bas`). The `Off` default reverses a
      decided position and needs an **ADR** plus an outbox note. Needs
      `brainstorming` → `writing-plans` first.

- [ ] **Catalog build-out, track 2** — the macro waves in
      `docs/macro-catalog-plug-and-play.md`. Package-level, no Core changes
      expected.

## Done (recent)

- [x] **Finish Excel, track 1** — DONE (2026-07-27, `v0.4.0`, contract
      **v1.2.0**). Capacity-packed auto-layout replacing the single `Macros` tab;
      the unmerged cell-card grid so descriptions are legible; three-state
      per-action availability (`Button` / `Chains only` / `Off`) with `Off`
      refused by the dispatcher and blocking Finalize inside a chain; `usedBy`
      and `category` on the manifest; `Sheet!Cell` anchored placements onto the
      user's own tabs, with Finalize blocking a malformed, reserved, missing or
      duplicated target; a machine-owned config schema marker with fail-closed
      v0.3.0 migration, refusing outright on a newer major and asking before
      discarding on a damaged one. Owner-verified in Excel across three workbooks
      (fixture, real, and one finalized under v0.3.0). The throwaway `PkgFixture`
      (74 no-op actions) that exercised the layout was removed on the way out.

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

**Catalog build-out** — one wave per minor release, each its own plan. Per-item
verdicts (wave / Later / Deferred / Declined, with the frozen action id) are
annotated in `docs/macro-catalog-plug-and-play.md`, now the authoritative build
queue. 35 new actions across six waves; 8 items *Later*, 5 *Deferred*, 4
*Declined*. "Build all 60" is explicitly not the goal.

- [ ] **W1 — ingest→clean path (7)** `import.combine_worksheets`,
      `import.combine_workbooks`, `import.delimited_file`, `clean.remove_blanks`,
      `clean.unmerge_fill`, `clean.standardize_formats`, `clean.text_to_numbers`.
      New `PkgImport` (the two combines share one private engine) + extend
      `PkgClean`. With the shipped trim/dupes/filldown this is a complete
      raw-files→clean-dataset job the user can wire into one chain immediately.
- [ ] **W2 — Validate & Reconcile (5)** `validate.reconcile` (⭐5, the
      highest-value item in the catalog for a multi-entity close), `validate.scan`,
      `validate.highlight_duplicates`, `validate.highlight_outliers`,
      `validate.missing_fields`.
- [ ] **W3 — Format & Present + Export (8)** five `format.*` plus `export.pdf`,
      `export.per_sheet`, `export.freeze_values`.
- [ ] **W4 — Summarize & Enrich (6)** four `summarize.*` plus `enrich.join`,
      `enrich.mapped_columns`.
- [ ] **W5 — Protect & Back Up (4)** four `backup.*`. Note `backup.snapshot` is
      the unblocker for the *Deferred* destructive items — if those become wanted,
      W5 moves up rather than the safety story being skipped.
- [ ] **W6 — Workbook housekeeping (5)** five `workbook.*`.

Wave order after W1 is deliberately revisable between ships; small waves exist
so it can change.

- [ ] **Apps Script port** — implement the frozen spec on Google Sheets, same
      Setup layout + action ids, different engine. **Deferred to last on
      purpose:** the frozen contract only makes the port mechanical *after* the
      Excel feature set settles. Porting against a moving target means porting
      twice.
- [ ] Refresh `docs/code-map.md` (`/code-map`) — absent today, and `Core.bas` is
      ~77KB, so it is read in slices via Grep. Worth generating before the polish
      pass touches the renderer and the panel builder.

## Housekeeping / minor

- [x] Delete merged local branches (`feat/*`, `chore/sanitize-defaults`,
      `docs/freeze-contract`) — DONE 2026-07-24; only `main` remains.
- [ ] Setup merged-band row heights are an estimate (AutoFit ignores merged
      cells). If a tinted explanation ever clips its last line, lower the
      chars-per-line constant in `Core.TintRow` (currently 115).
- [ ] Consider whether the type-map source-location (same folder, newest file)
      should become a configurable/asked option rather than a fixed rule
      (raised during UX review; deferred).
