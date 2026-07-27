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

## Status — sections 3-6 ARE THE RUN SHEET (2026-07-27, PR #7 merged)

Sections 1 and 2 passed and are done. **PR #7 has since merged, so sections 3-6 are now
verifying code that is already on `main` at v0.5.0.** They are no longer deferred: this
is the outstanding run, to be completed before Part A starts.

| Section | State |
|---|---|
| **1. The Setup surface** | **Pass.** Item 4 asserted the wrong outcome and was corrected — see the item itself; the code was never wrong. |
| **2. The three prompts** | **Pass**, except item 5 (the completion message's interpretation line) — folded into the run below. |
| **3. Cancel paths** | **TO RUN.** Five items. |
| **4. Key-column validation** | **TO RUN.** Three items. |
| **5. v0.4.0 refusal** | **TO RUN — do this one first.** Three items. Highest risk; see below. |
| **6. Inside a chain** | **TO RUN.** One item. |

**Run section 5 first.** It is the only one whose failure is *silent*. The `other-major`
refusal has never been reached by a genuine version bump — only by a hand-edited marker
during the v0.4.0 gate — and its neighbour, a v0.3.0-era workbook with no marker still
taking the legacy path, would fail without any visible symptom. Sections 3, 4 and 6 are
stop-early paths whose failure mode is a runtime-error dialog: annoying, but loud, and
incapable of corrupting data.

**Why now rather than later.** Part A changes this same `_Config` surface. Verifying it
first means any later failure is unambiguously Part A's doing rather than something
inherited from v0.5.0.

**Section 5 needs two old workbooks that do not exist yet — Q (v0.4.0) and C (v0.3.0).**
`dist/` is gitignored and there are no release tags, so neither the workbooks nor the old
bundles survive anywhere. **The Setup section below now builds both from their release
commits**, and the bundles are already sitting in `dist/`. Do not substitute a v0.5.0
workbook for either: a v0.5.0 config reads as `same-major`, never reaches the branch under
test, and the section would pass vacuously.

---

## Setup

### First: build the old bundles you will need

**You almost certainly do not have a v0.4.0 or v0.3.0 workbook lying around, and `dist/`
is gitignored so no old `.bas` survives either.** You have to build them. There are no
release tags, so the release commits are named directly:

```sh
# from the repo root, on any branch — these use throwaway worktrees and touch nothing
git worktree add /tmp/wt-v040 3e586cf     # chore(release): v0.4.0
git worktree add /tmp/wt-v030 cc59695     # chore: release v0.3.0

(cd /tmp/wt-v040 && pwsh build/build.ps1) && cp /tmp/wt-v040/dist/MacrosBundle.bas dist/MacrosBundle-v0.4.0.bas
(cd /tmp/wt-v030 && pwsh build/build.ps1) && cp /tmp/wt-v030/dist/MacrosBundle.bas dist/MacrosBundle-v0.3.0.bas

git worktree remove /tmp/wt-v040 --force && git worktree remove /tmp/wt-v030 --force
```

**Already done for you — both files are in `dist/` now.** They are gitignored, so rebuild
them with the above if they go missing. Verify before use, because the whole of §5 turns
on these three markers being different:

| Bundle | `SCHEMA_VER` | Reads as under v0.5.0 |
|---|---|---|
| `dist/MacrosBundle-v0.3.0.bas` | **no constant at all** | `absent` → legacy path |
| `dist/MacrosBundle-v0.4.0.bas` | `"v1.2"` | `other-major` → **refused** |
| `dist/MacrosBundle.bas` | `"v2.0"` | `same-major` → normal |

```sh
grep -c 'SCHEMA_VER' dist/MacrosBundle-v0.3.0.bas   # expect 0
grep -n 'SCHEMA_VER     As String' dist/MacrosBundle-v0.4.0.bas dist/MacrosBundle.bas
```

### Then: the three workbooks

**Q and C are INDEPENDENT workbooks built from separate lineages. Do not chain them.**
You do not start at v0.3.0 and upgrade through v0.4.0. Each starts as its own blank
workbook. Chaining is not a shortcut — it actively destroys the test: v0.4.0's Finalize
would rewrite C's marker-less `_Config` *with* a `v1.2` marker, turning C into a second
copy of Q and leaving the marker-less case untested.

Every workbook uses the **same three entry points**, in every era: `Core_BuildSetup`,
`Core_FinalizeSetup`, `Core_BuildControlPanel`. (The v0.3.0 build script's closing hint
says "run `Core_BuildControlPanel`", but `Core_BuildSetup` exists in that bundle at line
831 and is the one to use.)

Each of Q and C has the same two-phase shape: **build it old, then upgrade it.**

#### Workbook P — clean install

1. New blank workbook, save as **macro-enabled** (`.xlsm`).
2. Alt+F11 → *File ▸ Import File* → `dist/MacrosBundle.bas`.
3. *Debug ▸ Compile VBAProject*.
4. Alt+F8 → `Core_BuildSetup`.
5. Alt+F8 → `Core_FinalizeSetup`.

#### Workbook Q — finalized under v0.4.0, then upgraded → must be REFUSED

*Phase 1 — make it a genuine v0.4.0 workbook:*

1. New blank workbook, save as `Q.xlsm`.
2. Alt+F11 → *File ▸ Import File* → **`dist/MacrosBundle-v0.4.0.bas`**.
3. Alt+F8 → `Core_BuildSetup`.
4. Alt+F8 → `Core_FinalizeSetup`. **This is the step that matters** — it writes the hidden
   `_Config` stamped `v1.2`. Skip it and there is no `_Config`, the workbook reads as
   `none`, and §5 passes having tested nothing.
5. Save and close.

*Phase 2 — upgrade it to the current build:*

6. Reopen `Q.xlsm`. Alt+F11 → right-click the `MacrosBundle` module ▸ *Remove* ▸ **No** to
   the export prompt.
7. *File ▸ Import File* → `dist/MacrosBundle.bas` (the current v2.0 build).
8. Alt+F8 → `Core_BuildSetup`. **This is §5 item 1** — expect the refusal.

#### Workbook C — finalized under v0.3.0, then upgraded → must upgrade CLEANLY

Identical to Q, with one substitution: at step 2 import **`dist/MacrosBundle-v0.3.0.bas`**.
That bundle has no `SCHEMA_VER` constant at all, so its Finalize writes a `_Config` with
**no marker**, which is what makes it the legacy case. Everything else is the same, and
step 8 is **§5 item 2** — expect a clean upgrade, not a refusal.

**Finalizing is not optional for Q or C.** `Core_BuildSetup` only draws the Setup tab;
`Core_FinalizeSetup` is what copies it to the hidden `_Config` that carries the marker.

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
- [x] **The completion message states the interpretation.** It ends
      `Every row was compared, including the first.` Run it again answering `Yes` and
      it ends `The first row was treated as a header and left alone.` This is the
      point of the change: the answer is invisible in the result, so the report has
      to carry it.

## 3. Cancel is clean at all three points (workbook P)

Each of these must stop with **nothing changed and no error dialog**:

- [x] **Cancel the range picker.**
- [x] **Cancel the header question** — press `Cancel`, not `No`.
- [x] **Close the header question with the title-bar X.** Must behave as Cancel.
      This is why the dialog has three buttons: on a two-button `vbYesNo` MsgBox the
      X returns `vbNo`, which here would silently mean "compare every row" and delete
      a header row.
- [x] **Cancel the key-column box.**
- [x] **Clear the key-column box and press OK.** An empty answer is NOT a cancel — it
      means "use the default", so it proceeds with column `1`.

## 4. Key-column validation (workbook P)

- [x] **Multiple columns.** Select `M2:N12` and answer `1,2`. Both columns decide a
      duplicate — two rows are duplicates only when they match in both.
- [x] **A column outside the range is refused before anything is edited.** Select the
      single column `M2:M12` and answer `5`. Expected:
      `Column 5 is outside M2:M12, which is 1 column(s) wide.` followed by
      `Nothing has been changed. Run Remove Duplicates again.` — and the action stops.
      **It must not raise a runtime error**, which is what would happen if this reached
      `RemoveDuplicates`, and which would read as a broken macro rather than a typo.
- [x] **Non-numeric input falls back rather than failing.** Answer `abc`. Expected: it
      proceeds on column `1` (`ParseCols` drops non-numeric entries and defaults).

## 5. A v0.4.0 workbook is REFUSED, on purpose (workbook Q)  — RUN THIS FIRST

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
- [ ] **A v0.3.0-era workbook still upgrades cleanly — workbook C, buildable, no longer
      conditional.** Its `_Config` has **no** marker at all (the v0.3.0 bundle has no
      `SCHEMA_VER` constant to write one), so it reads as `absent` and takes the legacy
      path, which is unaffected.
      **This asymmetry is correct, not a bug:** older-with-no-marker is accepted where
      newer-but-not-current is refused, because only the marker-less case has a defined
      legacy layout to migrate *from*.

      **This item is now load-bearing for Part A, which it was not when it was written.**
      ADR 0003 decides to leave the legacy path untouched, on the premise that it still
      works. That premise is *this test*. If it fails, Part A's design is built on sand
      and ADR 0003 has to be reopened before any code is written — so do not tick it
      unless you actually ran it, and do not skip it because workbook C is extra effort.
      The Setup section above now tells you how to build one.
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
