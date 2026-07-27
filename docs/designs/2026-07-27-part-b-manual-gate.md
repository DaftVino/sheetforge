# Manual gate — Part B, per-selection settings moved to run-time prompts

**Date:** 2026-07-27
**Branch:** `feat/prompt-per-selection-settings`
**Implements:** issue #5 (Part B of `2026-07-26-setup-surface-progressive-disclosure-design.md`)
**Structural gates already green:** `pwsh build/build.ps1` → `Core + 2 package(s)`,
`Sub 62/62`, `Function 122/122`, 184 unique procedures; `node
scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` → no diagnostics.

**Every item below names the workbook, the exact click path, and the exact expected
string.** The v0.4.0 gate produced twelve items that could not be followed as written,
all from naming a value without naming where it lives (issue #3). This checklist is
written against that failure.

---

## Status — partial pass, shipped deliberately (2026-07-27)

**Run by the owner in Excel. Sections 1 and 2 pass; sections 3-6 are deferred to a
later review pass at the owner's call, and the branch ships without them.** This is
recorded here rather than in the PR alone so the next session cannot mistake an
unticked box for one nobody looked at.

| Section | State |
|---|---|
| **1. The Setup surface** | **Pass.** Item 4 asserted the wrong outcome and was corrected — see the item itself; the code was never wrong. |
| **2. The three prompts** | **Pass**, except item 5 (the completion message's interpretation line), not run. |
| **3. Cancel paths** | **Deferred.** Five items, none run. |
| **4. Key-column validation** | **Deferred.** Three items, none run. |
| **5. v0.4.0 refusal** | **Deferred.** Three items, none run. **Highest residual risk** — see below. |
| **6. Inside a chain** | **Deferred.** One item, not run. |

**What that leaves unverified.** Section 5 is the one to weigh: the `other-major`
refusal has never been reached by a genuine version bump, only by a hand-edited marker
in the v0.4.0 gate, and its neighbour — a v0.3.0-era workbook with no marker at all
still taking the legacy path — is the case that would fail *silently*. Nothing in
sections 3, 4 or 6 can corrupt data: each is a stop-early path whose failure mode is a
runtime-error dialog rather than a wrong edit. Section 2's item 5 is cosmetic wording
on a dialog whose behaviour item 4 already confirmed.

**Before the next release, run sections 3-6.** They are not cancelled, only deferred.

---

## Setup

You need **two workbooks**:

| # | Workbook | How |
|---|---|---|
| **P** | A fresh scratch copy. Tests the new behaviour on a clean install. | Import `dist/MacrosBundle.bas`, `Debug ▸ Compile VBAProject`, Alt+F8 → `Core_BuildSetup`, then `Finalize Setup Tab`. |
| **Q** | A workbook **already finalized under v0.4.0** — tests that an existing `_Config` carrying the two withdrawn keys upgrades cleanly. | Take any workbook finalized before this change. Alt+F11, right-click the `MacrosBundle` module ▸ Remove ▸ **No** to the export prompt, import the new `dist/MacrosBundle.bas`, then run `Core_BuildSetup`. |

If a copy lands in `Downloads`, unblock it first (right-click ▸ Properties ▸ Unblock) —
that is Mark-of-the-Web, not a macro-settings problem.

---

## 1. The Setup surface (workbook P)

- [x] **There is no `CLEAN & STANDARDIZE` band on the Setup tab at all.** Not an empty
      one, not a heading with nothing under it — the band is absent. This is the new
      Core guard; before it, the package drew a banner over nothing.
- [x] **The `IMPORT / EXPORT` band is still there**, with its fields unchanged. Only
      the package with zero fields loses its band.
- [x] **All three Clean actions still appear in the Action catalog**, each with a
      working `Availability` dropdown and a `Placement` cell:
      `Remove Duplicates [clean.dupes]`, `Trim & Clean Text [clean.trim]`,
      `Fill Blanks Down [clean.filldown]`. A package with no settings still has buttons.
- [x] **The other direction of the same guard — corrected 2026-07-27, it asserted the
      wrong outcome.** Set **every** `importexport.*` action to `Off` in the
      Availability column and run `Create Setup Tab`. Expected: the `IMPORT / EXPORT`
      band is **still there, with all nine fields in it**, and **no** off band appears
      at the bottom. Only the buttons go away. Set them back to `Button` afterwards.

      **This is correct, not a regression.** Relegation is driven by `usedBy`, and a
      field that declares no `usedBy` is package-wide and is *never* relegated
      (`FieldInUse`, `Core.bas:385-400`). No field in `PkgImportExport` declares one,
      so nothing can relegate, so `RenderOffBand`'s `anyOff` check draws no band
      (`Core.bas:1552-1555`) and `HasAnyTier(reg, availMap, True)` stays True, so the
      package keeps its own. **`usedBy` is exercised only by `PkgFixture` today** — the
      design doc says so outright (`2026-07-26-…-design.md:57`) and `Core.bas:2594`
      carries a comment written to stop precisely this overclaim.

      As originally written this item was unreachable with the shipped manifests, and
      the second half of the guard therefore **cannot be tested from a shipping
      package until Part A wires `usedBy` onto `PkgImportExport`'s fields**. Part A's
      gate inherits it; see the design doc's note that Part A's value rests entirely
      on those fields. The zero-field direction (item 1 above) is fully covered.

## 2. The three prompts, in order (workbook P)

Put eleven values in `M2:M12` on a scratch tab, with `M2`'s value repeated at `M5`.
Leave `M1` empty. Click **Remove Duplicates**.

- [x] **Prompt 1 — the range.** `Select the range for 'Remove Duplicates':`. Pick
      `M2:M12`.
- [x] **Prompt 2 — the header, and it names your range.** Reads
      `Does M2:M12 include a header row?` — the actual range, not a generic phrase —
      with **three** buttons: `Yes`, `No`, `Cancel`. `Yes` is the default.
- [x] **Prompt 3 — the key columns, and it names the range AND its width.** Reads
      `Numbered from the LEFT of M2:M12, which is 1 column(s) wide.` and is
      pre-filled with `1`.
- [x] **It runs.** Answer `No` then accept `1`. Expected: **ten rows remain** — `M5`
      removed because `M2` participated. Eleven remaining means the first row was
      skipped and the original symptom is back.
- [ ] **DEFERRED — The completion message states the interpretation.** It ends
      `Every row was compared, including the first.` Run it again answering `Yes` and
      it ends `The first row was treated as a header and left alone.` This is the
      point of the change: the answer is invisible in the result, so the report has
      to carry it.

## 3. Cancel is clean at all three points (workbook P) — DEFERRED

Each of these must stop with **nothing changed and no error dialog**:

- [ ] **Cancel the range picker.**
- [ ] **Cancel the header question** — press `Cancel`, not `No`.
- [ ] **Close the header question with the title-bar X.** Must behave as Cancel.
      This is why the dialog has three buttons: on a two-button `vbYesNo` MsgBox the
      X returns `vbNo`, which here would silently mean "compare every row" and delete
      a header row.
- [ ] **Cancel the key-column box.**
- [ ] **Clear the key-column box and press OK.** An empty answer is NOT a cancel — it
      means "use the default", so it proceeds with column `1`.

## 4. Key-column validation (workbook P) — DEFERRED

- [ ] **Multiple columns.** Select `M2:N12` and answer `1,2`. Both columns decide a
      duplicate — two rows are duplicates only when they match in both.
- [ ] **A column outside the range is refused before anything is edited.** Select the
      single column `M2:M12` and answer `5`. Expected:
      `Column 5 is outside M2:M12, which is 1 column(s) wide.` followed by
      `Nothing has been changed. Run Remove Duplicates again.` — and the action stops.
      **It must not raise a runtime error**, which is what would happen if this reached
      `RemoveDuplicates`, and which would read as a broken macro rather than a typo.
- [ ] **Non-numeric input falls back rather than failing.** Answer `abc`. Expected: it
      proceeds on column `1` (`ParseCols` drops non-numeric entries and defaults).

## 5. A v0.4.0 workbook is REFUSED, on purpose (workbook Q) — DEFERRED

**This section inverted after the version decision. Read it before you run it.**
`SCHEMA_VER` is now `"v2.0"`, so `CurrentSchemaMajor()` is 2 and workbook Q's marker
(`v1.2`) reads as `other-major`. A v0.4.0 config is **not migrated — it is refused.**
That is stricter than the data needs (the catalog block layout is byte-identical
across this change) and was accepted deliberately: there is no installed base, so the
false refusal costs nothing, and one version number beats two. See
`docs/adr/0002-contract-version-vs-schema-marker.md`, Rejected.

- [ ] **`Create Setup Tab` on workbook Q refuses outright**, tells you the config was
      written by a different version, and says to re-import that version's bundle.
      **No rebuild prompt, no Yes/No** — this is the `other-major` path, and it must
      not offer a choice. Your Setup tab is left exactly as it was.
- [ ] **A v0.3.0-era workbook still upgrades cleanly**, if you still have one from the
      v0.4.0 gate (workbook C). Its `_Config` has **no** marker at all, so it reads as
      `absent` and takes the legacy migration path, which is unaffected.
      **This asymmetry is correct, not a bug:** older-with-no-marker is accepted where
      newer-but-not-current is refused, because only the marker-less case has a defined
      legacy layout to migrate *from*.
- [ ] **The recovery route works.** On workbook Q, delete the whole `_Config` sheet
      (Alt+F11 → `Visible` = `-1 - xlSheetVisible`, then right-click the tab ▸ Delete),
      and run `Create Setup Tab` again. Expected: it builds a fresh Setup tab with
      defaults, no `CLEAN & STANDARDIZE` band, and no error. That is the documented
      way out for a workbook this build will not read.

**If `Create Setup Tab` silently migrates workbook Q instead of refusing it**, the
marker bump did not take effect — check `SCHEMA_VER` really is `"v2.0"` in the bundle
you imported, and that you removed the old module before importing.

## 6. Inside a chain (workbook P) — DEFERRED

- [ ] **A chain containing `clean.dupes` still prompts.** Build a two-step chain with
      `clean.dupes` in it and run it. All three prompts appear. This is correct and not
      a regression — the action was always interactive, since it has always asked for a
      range. Confirm the chain's single summary dialog still appears at the end and the
      per-step dialog is still suppressed.

---

## Version decision — settled

**Contract v2.0.0, and `SCHEMA_VER` moves with it to `"v2.0"`.** Withdrawing two keys
and changing an action's semantics is breaking by §10's own rules. The marker was kept
coupled to the contract version rather than decoupled, because **no deployment exists
yet** — the bundle is in test workbooks only, so refusing v0.4.0 configs costs nothing,
and one version number is simpler than two. `docs/adr/0002-contract-version-vs-schema-marker.md`
records the rejected alternative in full, because the analysis stops being *expensive*
rather than stopping being *true*: **revisit it before the next breaking bump if an
installed base exists by then.**

The visible consequence is section 5 — a v0.4.0 workbook is refused, not upgraded.
