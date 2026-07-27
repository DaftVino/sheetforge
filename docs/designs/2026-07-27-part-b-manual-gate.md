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

- [ ] **There is no `CLEAN & STANDARDIZE` band on the Setup tab at all.** Not an empty
      one, not a heading with nothing under it — the band is absent. This is the new
      Core guard; before it, the package drew a banner over nothing.
- [ ] **The `IMPORT / EXPORT` band is still there**, with its fields unchanged. Only
      the package with zero fields loses its band.
- [ ] **All three Clean actions still appear in the Action catalog**, each with a
      working `Availability` dropdown and a `Placement` cell:
      `Remove Duplicates [clean.dupes]`, `Trim & Clean Text [clean.trim]`,
      `Fill Blanks Down [clean.filldown]`. A package with no settings still has buttons.
- [ ] **The other direction of the same guard.** Set **every** `importexport.*` action
      to `Off` in the Availability column and run `Create Setup Tab`. Expected: the
      `IMPORT / EXPORT` band disappears too — its fields have all relegated to the
      collapsed off band at the bottom, so there is nothing left to head. Set them
      back to `Button` afterwards.

## 2. The three prompts, in order (workbook P)

Put eleven values in `M2:M12` on a scratch tab, with `M2`'s value repeated at `M5`.
Leave `M1` empty. Click **Remove Duplicates**.

- [ ] **Prompt 1 — the range.** `Select the range for 'Remove Duplicates':`. Pick
      `M2:M12`.
- [ ] **Prompt 2 — the header, and it names your range.** Reads
      `Does M2:M12 include a header row?` — the actual range, not a generic phrase —
      with **three** buttons: `Yes`, `No`, `Cancel`. `Yes` is the default.
- [ ] **Prompt 3 — the key columns, and it names the range AND its width.** Reads
      `Numbered from the LEFT of M2:M12, which is 1 column(s) wide.` and is
      pre-filled with `1`.
- [ ] **It runs.** Answer `No` then accept `1`. Expected: **ten rows remain** — `M5`
      removed because `M2` participated. Eleven remaining means the first row was
      skipped and the original symptom is back.
- [ ] **The completion message states the interpretation.** It ends
      `Every row was compared, including the first.` Run it again answering `Yes` and
      it ends `The first row was treated as a header and left alone.` This is the
      point of the change: the answer is invisible in the result, so the report has
      to carry it.

## 3. Cancel is clean at all three points (workbook P)

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

## 4. Key-column validation (workbook P)

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

## 5. A v0.4.0 workbook is REFUSED, on purpose (workbook Q)

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

## 6. Inside a chain (workbook P)

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
