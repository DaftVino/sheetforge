# TODOs

Local task list until this repo gets a GitHub Project (repo-standards 6.5). Tasks
touching Core or more than one package go through `writing-plans` first. Verify =
`pwsh build/build.ps1`; there is no headless VBA test, so also import
`dist/MacrosBundle.bas` and run the workbook, recorded in the PR/notes.

Current: **v0.17.0**, contract **v2.9.0**. **Track 2 is code-complete: 43 actions across
ten packages.** Track 1 shipped at v0.6.0-v0.9.0. Remote: `origin` =
`github.com/DaftVino/sheetforge` (private). The Apps Script port is deferred to last,
deliberately.

**The one thing blocking release is validation time, not build time.** Nothing in track 2
has run in Excel. Seven PRs are open and stacked; they must merge in order.

## Next — active

- [ ] **Merge the stack, in this order.** Each branches from the one before it, so an
      out-of-order merge does not build:
      **#20** Core refusal pass (v0.11.0, contract v2.4.0)
      → **#21** Core helper promotion (v0.12.0)
      → **#22** W2 Validate & Reconcile (v0.13.0, v2.5.0)
      → **#23** W3 Format & Export (v0.14.0, v2.6.0)
      → **#24** W4 Summarize & Enrich (v0.15.0, v2.7.0)
      → **#25** W5 Protect & Back Up (v0.16.0, v2.8.0)
      → **#26** W6 Workbook housekeeping (v0.17.0, v2.9.0)
      → **#27** the end-to-end validation document.

- [ ] **Run `docs/designs/2026-07-28-end-to-end-validation.md`.** 41 items, ordered by
      consequence rather than by release. It replaces the six per-wave gates that were never
      written, per the 2026-07-28 deferral. It has been dry-run by an agent that was allowed
      to read only the checklist and the built bundle — that pass corrected five button
      captions that did not exist and six items that could not have been executed as written.

      **§2 is the section that matters.** Seven items, each of which can destroy work rather
      than merely report a wrong number. If §2 is green, nothing in this release can lose
      your data. §3 is the wrong-numbers section. §4 and §5 degrade gracefully.

      **Start with 1.2** (`Debug ▸ Compile VBAProject`) — it discharges the only open
      assumption in the build, and everything else is invalid until it passes. Then **3.2**,
      the oldest unverified change in the tree.

- [ ] **Reskin 4 is still the likeliest real finding** (item 5.1). The palette came off a
      dark-background proof and lands on an explicitly white sheet, and nothing has ever
      checked the two together.

## After validation

- [ ] **Ship v1.0.0** once the validation passes, or fix what it finds first.
- [ ] **The Apps Script port**, against `spec/contract.md` v2.9.0. The contract is the
      language-neutral source of truth and every wave has kept it current, so the port is
      meant to be mechanical.
- [ ] **8 catalog items are marked *Later*** and 5 *Deferred*. The Deferred ones are still
      deferred: `backup.snapshot` supplies the save half of an undo story, not the restore
      half, and only one of the three blocked items ever cited undo as its blocker.

## Before the remote (blocker) — DONE (2026-07-24)

- [x] **Sanitize real store/entity specifics** and publish. Scrubbed all 7
      tracked files to fully-generic demo defaults (design + plan:
      `docs/designs/sanitization-pass.md`, `-plan.md`); genericized the real
      store number to `12345`; renamed the legacy client module to
      `SetupImportExportMacros`. `main` fast-forwarded and pushed to `origin`.
      **Caveat (tree-only choice):** the pre-`d29b0d7` *history* still
      contains the old jargon — the remote is private, so this was accepted;
      revisit (squash/filter history) before any public release.

      **2026-07-29, Task 12 step 1 — this note was itself a leak** and has been
      redacted. It quoted both retired tokens verbatim, so the audit trail
      recording their removal was the last place they survived at the tip. A
      sanitization record must name what it removed by description, never by
      value. The same pass also found `docs/macro_code_library.md` carrying six
      live references to the retired client module name — it was not one of the
      "7 tracked files" the original pass scrubbed. **The history caveat above
      is still open and now blocks going public**; see the Task 12 report.

## Next — active

- [~] **W1 manual gate — DEFERRED 2026-07-28 by owner decision.** Folded into one
      end-to-end validation on the finished workbook, after the Core pass and the
      remaining waves. Four items ran first; item 2.3 found a real defect (every source
      spaced ~15 rows apart, from an inflated `UsedRange`), fixed in `bafb905` and **not
      yet verified in Excel**. Items 1.5 and 2.4 exist to confirm that fix when the gate
      resumes. Original item text below, kept for when it does:

- [ ] ~~**Run the W1 manual gate** (`docs/designs/2026-07-28-w1-manual-gate.md`, 37
      items) and merge **PR #19**. **Start with item 2.3** — it is the only item
      that puts a colleague's unsaved workbook in the combine's path, which is
      exactly where the whole-branch review found a Critical: `Workbooks.Open` on
      an already-open file returns *the user's copy*, and closing it with alerts
      suppressed destroyed their unsaved edits while reporting success. Fixed, but
      unverified in Excel.~~

- [ ] **Finish the 8 unrun track-1 gate items.** Hardening 1.3, 2.2, 3.1, 4.1,
      4.2 · reorder 2 (clipping half), 3 · reskin 4. **Reskin 4 is the likeliest
      real finding** — the palette came off a dark-background proof and lands on
      an explicitly white sheet, and nothing has checked the two together. Two need
      their own build: 3.1 a hostile `PkgClean` declaring `core.legacy`, reorder 3
      the v0.6.0 bundle from `812d88c`. `tools/GateAudit.bas` re-runs seven
      mechanical checks in one click.

- [ ] **Core pass, before W2 — `Core_Refuse` + `Core_IsEngineSheet`.** The parked
      W1 findings compound per wave, which is the whole argument for doing this
      now rather than after:
      - `PkgClean_Refuse` and `PkgImport_Refuse` are byte-identical twins today.
        W2 makes three copies.
      - `PkgImport_IsEngineSheet` hardcodes seven sheet names Core owns as
        `Private Const`. W3's `export`/`workbook` package needs the same test.
      - A refusal sets `failed`/`failMsg` but not `failNumber`, so every chain
        summary line reads `X #0 <reason>`. Every new wave inherits it.
      - `Core_RunChain` does not abort on a failed step, it asks. A user who
        answers Continue runs the remaining steps against a stale sheet — an
        informed choice now, but only Core can change it.
      Touches Core → needs `writing-plans` first.

- [ ] **W2 — Validate & Reconcile (5).** The highest-value group in the catalog
      for a multi-entity close. Do the Core pass first.

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

- [x] **W1 — ingest→clean path (7)** — BUILT 2026-07-28 (`v0.10.0`, contract
      **v2.3.0**), PR #19 open, **manual gate not yet run**. `import.combine_worksheets`,
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
- [x] Refresh `docs/code-map.md` — DONE. It exists, is regenerated by
      `node scripts/code-map.mjs`, and is current. `Core.bas` is now ~229KB, so
      grep the map for an anchor and read the slice; never open it whole.

## Housekeeping / minor

- [x] Delete the stale `.superpowers/sdd/` workspaces — DONE 2026-07-28. The W1
      workspace is deliberately kept until PR #19 merges; it is the recovery map.

- [x] Delete merged local branches (`feat/*`, `chore/sanitize-defaults`,
      `docs/freeze-contract`) — DONE 2026-07-24; only `main` remains.
- [ ] Setup merged-band row heights are an estimate (AutoFit ignores merged
      cells). If a tinted explanation ever clips its last line, lower the
      chars-per-line constant in `Core.TintRow` (currently 115).
- [ ] Consider whether the type-map source-location (same folder, newest file)
      should become a configurable/asked option rather than a fixed rule
      (raised during UX review; deferred).
