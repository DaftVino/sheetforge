# TODOs

Local task list until this repo gets a GitHub Project (repo-standards 6.5). Tasks
touching Core or more than one package go through `writing-plans` first. Verify =
`pwsh build/build.ps1`; there is no headless VBA test, so also import
`dist/MacrosBundle.bas` and run the workbook, recorded in the PR/notes.

Current: **v0.3.0** on `main` (Compose Your Buttons — chains + grouping,
contract **v1.1.0** — merged and owner-verified in Excel). Remote: `origin` =
`github.com/DaftVino/sheetforge` (private). Next phase is **Finish Excel**,
designed in `docs/designs/2026-07-24-finish-excel-design.md`: the built-page
polish pass first, then the catalog build-out in waves. The Apps Script port is
deferred to last, deliberately.

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

- [ ] **Built-page polish pass (`v0.4.0`, contract `v1.2.0`)** — design §2.
      Core-only. Generated `Macros Home` (legend + one card per category with
      action counts), `Macros N` pages titled by category with descriptions,
      `← Home` and Previous/Next nav, auto-layout so no-codes stops meaning one
      giant tab, tighter spacing. Plus **per-action enable/disable** on the
      placement block (which becomes `Action | Enabled | Page code`) and the
      `usedBy` field attribute, so Setup collapses settings for actions that are
      turned off instead of walling the user with them. Verified against a
      throwaway `PkgFixture` of ~60 no-op actions so the 60-button layout is
      exercised while only 7 real actions exist. Needs `writing-plans` first.

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
