# Manual gate — Part A, progressive disclosure (Availability defaults to Off)

**Date:** 2026-07-27
**Branch:** `feat/progressive-disclosure-part-a`
**Implements:** issue #6 (Part A of `2026-07-26-setup-surface-progressive-disclosure-design.md`)
and issue #10
**Decisions:** `docs/adr/0003-availability-default-and-legacy-migration.md`
**Contract:** v2.1.0
**Structural gates already green:** `pwsh build/build.ps1` → `Core + 2 package(s)`,
`Sub 62/62`, `Function 122/122`, 184 unique procedures; `node
scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` → no diagnostics;
`SCHEMA_VER` confirmed still `"v2.0"` in the built bundle.

**Every item below names the workbook, the exact click path, and the exact expected
string.** That is the standard set by the Part B gate and required by issue #3 — the
v0.4.0 gate produced twelve items that could not be followed as written, all from naming
a value without naming where it lives.

---

## Status — RUN AND PASSED, 2026-07-27

**Run in Excel by the owner on 2026-07-27** against `dist/MacrosBundle.bas` at `ecf500d`,
across workbooks A, B, C and D as specified below. All 13 items pass.

| Section | Items | State |
|---|---|---|
| **1. Fresh install** | 3/3 | Passed |
| **2. Legacy and stored configs** | 4/4 | Passed |
| **3. Fail-closed on a bad cell** | 2/2 | Passed |
| **4. Field disclosure** | 3/3 | Passed |
| **5. Runtime** | 1/1 | Passed |
| **Total** | **13/13** | **Passed** |

### Two items were unrunnable as first written — the instructions were wrong, not the code

Both were corrected mid-run and then passed. Recording them because the corrections are the
gate's real output, and issue #3 exists precisely because a checklist that cannot be followed
as written is worse than no checklist:

- **§3 item 2 said to *type* an unrecognised value.** You cannot. `AvailDropdown` applies list
  validation with `xlValidAlertStop` (`Core.bas:2120`), so Excel refuses a typed entry — the
  alert is the dropdown working. The item now specifies a **paste**, which bypasses validation
  and is also the only realistic way a bad value reaches a real deployment.
- **Every §4 item went straight from editing an `Availability` cell to `Core_BuildSetup`.**
  `Core_BuildSetup` reads the map from the **activated `_Config`** (`Core.bas:1399`), never from
  what is currently typed on the Setup tab, so a rebuild without a Finalize in between is a
  no-op and all three items looked broken while working correctly. Every §4 step is now
  **edit → `Core_FinalizeSetup` → `Core_BuildSetup`**. A stale comment at `Core.bas:1396` that
  still described the pre-Task-1 default was corrected at the same time.

§4 item 3 was also rewritten after the run: its expected result is "nothing moves and the off
band does not appear at all", which the original wording never said, and it now names the
on-screen block to look at and both dialogs `Suggest Import Types` can produce.

### The item that matters most

**§2 item 1 — a v0.3.0 legacy workbook keeps every button.** This is the regression test
for contract §5's pinned legacy default (Task 1 Step 2). §6's default moved to `off`; §5's
legacy default was deliberately pinned to `button` and does NOT follow it. If that pin were
lost or wrong, a working pre-v1.2 installation would silently lose every button on its next
rebuild. If this item fails, Task 1 Step 2 was wrong or was reverted — stop and fix that
before anything else.

### The item with no prior passing run

**§4 item 1** exercises the empty-band guard's *second* direction for the first time ever.
Part B's gate item 1.4 could not test it: relegation is driven by `usedBy`, no shipping
package declared one, so nothing could relegate. Task 4 wired `usedBy` onto
`PkgImportExport`'s nine fields, which makes it reachable. **There is no earlier passing run
to compare against** — a failure here may be a genuine Core bug rather than a Part A
regression, so investigate before assuming Part A broke it.

**Outcome: passed on 2026-07-27**, once the missing Finalize was added to the steps. This is
the empty-band guard's first recorded pass in the relegating direction; it is the baseline
every later run compares against.

### Three items exist because a whole-branch review caught real defects

§1 item 3, §2 item 4 and the corrected §4 items were added after an adversarial review of
the finished branch found two P1 defects that all five tasks, the build and the syntax
checker had passed clean:

- **`FieldInUse` asked the stored map whether an id was "unknown".** A fresh workbook's
  map is empty, so every real id looked unknown, so nothing relegated and Part A had no
  visible effect on exactly the install it was designed for. It now asks the package
  registry. → §1 item 3.
- **`ReadAvailMap`'s three exits needed three different answers, not two.** The review
  first argued the "no catalog block at all" exit should be flagged legacy like the
  absent-marker one. It was, then reverted on re-review: that exit is reached by a damaged
  *current* config as well as a pre-v1.1 one, and flagging it fails open on destructive
  actions. It fails closed, and the reasoning is written into the code. → §2 item 4, which
  records it as a judgement call the owner may want to overturn.

Both were invisible to every structural gate, and both are the *same* failure the pin in
§2 item 1 guards against. Treat that as the standing lesson: **on this change, "the build
is green" says almost nothing.**

---

## Setup

### First: build the old bundles you will need

`dist/` is gitignored, so no old `.bas` survives. There are no release tags, so the release
commits are named directly.

```sh
# from the repo root, on any branch - throwaway worktrees, touches nothing tracked
git worktree add /tmp/wt-v030 cc59695     # chore: release v0.3.0
git worktree add /tmp/wt-v050 0b7094d     # release v0.5.0

(cd /tmp/wt-v030 && pwsh build/build.ps1) && cp /tmp/wt-v030/dist/MacrosBundle.bas dist/MacrosBundle-v0.3.0.bas
(cd /tmp/wt-v050 && pwsh build/build.ps1) && cp /tmp/wt-v050/dist/MacrosBundle.bas dist/MacrosBundle-v0.5.0.bas

git worktree remove /tmp/wt-v030 --force && git worktree remove /tmp/wt-v050 --force
```

**Already done for you — both files are in `dist/` now.** Verify before use:

| Bundle | `SCHEMA_VER` | Reads as under this build | Availability default it wrote |
|---|---|---|---|
| `dist/MacrosBundle-v0.3.0.bas` | **no constant at all** | `absent` → legacy path | column did not exist |
| `dist/MacrosBundle-v0.5.0.bas` | `"v2.0"` | `same-major` → normal | `Button` |
| `dist/MacrosBundle.bas` | `"v2.0"` | `same-major` → normal | **`Off`** |

```sh
grep -c 'SCHEMA_VER' dist/MacrosBundle-v0.3.0.bas            # expect 0
grep -n 'SCHEMA_VER     As String' dist/MacrosBundle-v0.5.0.bas dist/MacrosBundle.bas
```

**v0.5.0 and this build share `"v2.0"` deliberately.** ADR 0003 keeps the marker put so
every v0.5.0 config stays readable — that is what §2 item 2 tests. If a build ever reports
`"v3.0"`, something is wrong; stop.

### Then: the three workbooks

**A, B and C are INDEPENDENT workbooks built from separate lineages. Do not chain them.**
Chaining destroys the test: the current build's Finalize would rewrite C's marker-less
`_Config` *with* a `v2.0` marker, turning C into a second copy of B and leaving the legacy
case untested.

**Workbook C must be rebuilt from scratch for this gate — do not reuse Part B's C.** That
one was upgraded during the Part B run, so its `_Config` now carries a current marker. It
would pass the legacy test vacuously. Pristineness has to be structural, not asserted.
**Save a pristine copy of C before upgrading it**, so the next gate does not need a rebuild.

Every workbook uses the same three entry points in every era: `Core_BuildSetup`,
`Core_FinalizeSetup`, `Core_BuildControlPanel`. (The v0.3.0 build script's closing hint says
`Core_BuildControlPanel`, but `Core_BuildSetup` exists in that bundle and is the one to use.)

#### Workbook A — clean install on the current build

1. New blank workbook, save as **macro-enabled**, `A.xlsm`.
2. Alt+F11 → *File ▸ Import File* → `dist/MacrosBundle.bas`.
3. *Debug ▸ Compile VBAProject*. **Do not skip this** — a keyword-as-identifier error is
   reported here and nowhere else; neither `build.ps1` nor the syntax checker can see it.
4. Alt+F8 → `Core_BuildSetup`.

Stop there. **Do not Finalize yet** — §1 is about what a fresh install shows before anything
is activated.

#### Workbook B — finalized under v0.5.0, then upgraded → every stored choice must survive

*Phase 1 — make it a genuine v0.5.0 workbook:*

1. New blank workbook, save as `B.xlsm`.
2. Alt+F11 → *File ▸ Import File* → **`dist/MacrosBundle-v0.5.0.bas`**.
3. Alt+F8 → `Core_BuildSetup`. Every `Availability` cell will read `Button` — that is
   v0.5.0's default and is correct for this phase.
4. On the Setup tab, in the **Action catalog** block, set a deliberate mix in the
   `Availability` column:
   - `Remove Duplicates [clean.dupes]` → **`Off`**
   - `Trim & Clean Text [clean.trim]` → **`Chains only`**
   - `4.  Export Dashboard (.xlsx) [importexport.export]` → **`Off`**
   - leave every other action at `Button`.
5. Alt+F8 → `Core_FinalizeSetup`. **This is the step that matters** — it writes the hidden
   `_Config` carrying those five-of-eight choices and the `v2.0` marker. Skip it and the
   workbook reads as `none` and §2 item 2 passes having tested nothing.
6. Save and close.

*Phase 2 — upgrade it to the current build:*

7. Reopen `B.xlsm`. Alt+F11 → right-click the `MacrosBundle` module ▸ *Remove* ▸ **No** to
   the export prompt.
8. *File ▸ Import File* → `dist/MacrosBundle.bas`.
9. Alt+F8 → `Core_BuildSetup`. **This is §2 item 2.**

#### Workbook C — finalized under v0.3.0, then upgraded → every button must survive

Identical shape to B, with one substitution: at step 2 import **`dist/MacrosBundle-v0.3.0.bas`**.
That bundle has no `SCHEMA_VER` constant and no `Availability` column at all, so its Finalize
writes a `_Config` with **no marker** — which is what makes it the legacy case. There is no
step 4 for C: there is nothing to set.

1. New blank workbook, save as `C.xlsm`.
2. Alt+F11 → *File ▸ Import File* → **`dist/MacrosBundle-v0.3.0.bas`**.
3. Alt+F8 → `Core_BuildSetup`.
4. Alt+F8 → `Core_FinalizeSetup`.
5. Save and close. **Now copy the file to `C-pristine.xlsm` and leave that copy alone.**
6. Reopen `C.xlsm`. Alt+F11 → right-click `MacrosBundle` ▸ *Remove* ▸ **No**.
7. *File ▸ Import File* → `dist/MacrosBundle.bas`.
8. Alt+F8 → `Core_BuildSetup`. **This is §2 item 1.**

If a copy lands in `Downloads`, unblock it first (right-click ▸ Properties ▸ Unblock) — that
is Mark-of-the-Web, not a macro-settings problem.

**Known-benign, recorded so it is not re-investigated:** the first `Core_BuildSetup` after
importing the v0.3.0 bundle may open a Windows **"Finding printer"** dialog. Cancel it; it
causes no harm and does not recur. Nothing in the toolkit touches printing — it is Excel
computing page-break metrics for a newly added sheet. Diagnosed in full in the Part B gate.

---

## 1. A fresh install shows a deliberate catalog (workbook A)

- [x] **Every `Availability` cell reads `Off`.** On the Setup tab, in the **Action catalog**
      block, all eight actions — `Remove Duplicates [clean.dupes]`,
      `Trim & Clean Text [clean.trim]`, `Fill Blanks Down [clean.filldown]`,
      `1.  Clear Input Tabs [importexport.clear]`, `2.  Import Source Data [importexport.import]`,
      `3.  Update Ref_Master [importexport.updateMaster]`,
      `4.  Export Dashboard (.xlsx) [importexport.export]`,
      `Suggest Import Types [importexport.suggestTypes]` — read exactly **`Off`**.
      Not blank, not `Button`. **This is the feature.**

      A blank cell here is a failure, not a variant: blank now blocks Finalize (§3 item 1),
      so it would mean `AvailLabel` is not rendering the token.

- [x] **No buttons are drawn.** There is no `Macros 1` tab (nor `Macros`, nor `Macros 2`)
      anywhere in the workbook. Every action is `Off`, so there is nothing to place.

- [x] **Every setting has already stepped out of the way.** There is **no**
      `IMPORT / EXPORT` band on the Setup tab. Instead there is one band headed

      ```
      SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON
      ```

      collapsed, with a `+` in the left margin. Open it: it holds an `Import / Export`
      sub-heading and all nine fields. There is no `CLEAN & STANDARDIZE` band either
      (that package has no settings at all — v0.5.0's guard, unchanged).

      **This is the whole point of Part A and it was the last thing to work.** The first
      implementation got the catalog right and this wrong: `FieldInUse` decided "unknown
      id" by asking the stored availability map, and a fresh workbook's map is empty, so
      every real action id looked unknown, so every field counted as in-use and stayed in
      the main band under a catalog that said `Off`. It now asks the package registry
      instead. If you see the nine fields in a normal `IMPORT / EXPORT` band here, that
      regression is back.

## 2. Existing configurations are untouched

- [x] **A v0.3.0 legacy workbook keeps every button (workbook C).**
      **THE LOAD-BEARING ITEM — see "The item that matters most" above.**
      After Setup step 8, every `Availability` cell in workbook C's Action catalog reads
      **`Button`**, and running `Core_FinalizeSetup` then `Core_BuildControlPanel` draws
      every button exactly as v0.3.0 did.

      Its `_Config` has **no** marker (the v0.3.0 bundle has no `SCHEMA_VER` constant to
      write one), so it reads as `absent`, `ReadAvailMap` flags it under the reserved
      `core.legacy` key, and `AvailFor` returns `Button` for every id it does not find.

      **If any cell reads `Off`, STOP.** Contract §5's pinned legacy default has been lost
      and a working installation would be stripped of its buttons. Do not work around it.

- [x] **A marked v0.5.0 config keeps every stored choice verbatim (workbook B).**
      After Setup step 9, the Action catalog reads **exactly** what was set in Setup step 4:
      `Remove Duplicates [clean.dupes]` → `Off`; `Trim & Clean Text [clean.trim]` →
      `Chains only`; `4.  Export Dashboard (.xlsx) [importexport.export]` → `Off`;
      **every other action still `Button`, NOT `Off`.**

      The five that stayed `Button` are the real test. A stored `button` still means
      `button` — the new default applies only where nothing was ever stored, which is what
      makes contract v2.1.0 additive rather than breaking. If those five come back `Off`,
      the default is being applied to stored values and the change IS breaking.

- [x] **A newly-registered action arrives Off (workbook B).**
      Simulates a package upgrade adding an action to a config that predates it.
      Alt+F11 → in the Project pane select the `_Config` sheet → *Properties* →
      set `Visible` to `-1 - xlSheetVisible`. On `_Config`, find the Action catalog block and
      **delete the entire worksheet row** for `clean.filldown` (right-click the row number ▸
      *Delete*). Re-hide `_Config` (`Visible` = `2 - xlSheetVeryHidden`). Alt+F8 →
      `Core_BuildSetup`.

      Expected: `Fill Blanks Down [clean.filldown]` reads **`Off`**, and every other action
      still reads what it read before. Under v0.5.0 it would have come back `Button`.

- [x] **A `_Config` with no Action catalog block at all reads Off, and that is
      deliberate (workbook B).**
      The third provenance, and the one where the two safe answers point in opposite
      directions. Reached directly rather than with a fourth bundle, because it is the
      same code path.

      Alt+F11 → select the `_Config` sheet → *Properties* → `Visible` =
      `-1 - xlSheetVisible`. On `_Config`, find the row whose key column reads
      `k_core_layout` and **delete that entire worksheet row and every row of its block**,
      down to (not including) the next key row. Re-hide `_Config` (`Visible` =
      `2 - xlSheetVeryHidden`). Alt+F8 → `Core_BuildSetup`.

      Expected: every `Availability` cell reads **`Off`**, and — the part that matters —
      **your `IMPORT / EXPORT` field values are still there**, unchanged, in the off band.
      Nothing you typed was lost; only the availability choices, which lived in the block
      you deleted, are gone.

      **Why Off and not Button, when §2 item 1 says the opposite for workbook C.** Two
      histories reach this exact sheet state and cannot be told apart from it: a pre-v1.1
      config written before the catalog block existed, and a current config whose block
      was damaged. Flagging it legacy was implemented, then reverted: it fails *open*,
      handing a damaged current config's destructive actions back as live buttons. Failing
      closed costs the pre-v1.1 side almost nothing — that era had no Availability column
      **and no placements**, so there are no configured buttons to strip, and its config
      *values* load either way. Workbook C is different in the one way that counts: it has
      a real placement block, so reading it as Off would strip a working install.

      **If you disagree with that trade, say so — it is a judgement call, not a
      derivation.** It was made on the fact that ADR 0002 records no installed base.

      Restore `B.xlsm` from a copy afterwards, or rebuild it — you have destroyed its
      catalog block.

## 3. A bad Availability cell fails closed (workbook A — issue #10)

Set every action back to `Button` first, then Finalize, so you are starting from a valid
config.

- [x] **A blank cell blocks Finalize.** Clear the `Availability` cell for
      `Trim & Clean Text [clean.trim]` (select it, press Delete). Alt+F8 →
      `Core_FinalizeSetup`. Expected: the report refuses to activate and contains

      ```
      MISSING or UNRECOGNISED Availability (use Button, Chains only, or Off):
        - clean.trim
      ```

      Blank used to read as `Button`. With the default at `Off`, reading a cleared cell as
      either value is wrong and unrecoverable from the result, so it blocks instead.

- [x] **An unrecognised value is no longer laundered into `Button`.** Get `maybe` into the
      `Availability` cell for `Trim & Clean Text [clean.trim]`.

      **You cannot type it.** `AvailDropdown` applies list validation with
      `xlValidAlertStop` ([Core.bas:2120](../../src/core/Core.bas#L2120)), so Excel refuses a
      typed entry with a "The value doesn't match the data validation restrictions" alert —
      that alert is the dropdown working, not a failure. **Paste instead:** type `maybe` into
      any scratch cell, copy it, and paste over the `Availability` cell. Excel's paste carries
      the source cell's (absent) validation with it and lands the value. That is also the only
      realistic way a bad value reaches a real deployment, which is what makes it the right
      thing to test.

      Alt+F8 → `Core_FinalizeSetup` — expect the same block as above, naming `clean.trim`.
      Then Alt+F8 → `Core_BuildSetup`.

      Expected: the cell comes back **blank**. **It must NOT come back as `Button`.**
      That silent rewrite was issue #10: `AvailLabel`'s `Case Else` returned `"Button"`, the
      rebuild wrote it into the cell, and the block never fired again because the evidence
      had been overwritten. Fix the cell to `Off` afterwards and Finalize to leave the
      workbook valid.

## 4. Fields appear as their actions are turned on (workbook A)

Set every action to `Button` first and Finalize, so the Setup tab is in its "all on" state.

**Finalize between every toggle and every rebuild.** `Core_BuildSetup` reads the availability
map from the **activated** `_Config` ([Core.bas:1399](../../src/core/Core.bas#L1399)), not from
what is currently typed on the Setup tab. Change an `Availability` cell and rebuild without
finalizing first and the rebuild re-seeds Setup from the old config — nothing moves, and the
item looks broken when it is working. Every step below is therefore
**edit → `Core_FinalizeSetup` → `Core_BuildSetup`**.

- [x] **All-off hides a package's band and moves its fields to the off band.**
      Set **every** `importexport.*` action to `Off` — all five: `clear`, `import`,
      `updateMaster`, `export`, `suggestTypes`. Alt+F8 → `Core_FinalizeSetup`, then Alt+F8 →
      `Core_BuildSetup`.

      Expected: the `IMPORT / EXPORT` band **is gone**, and a band headed

      ```
      SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON
      ```

      appears at the bottom, collapsed, with a `+` in the left margin. Open it: it holds an
      `Import / Export` sub-heading and **all nine fields** — `k_export_base`,
      `k_settings_sheet`, `k_export_offset`, `k_export_tabs`, `k_import`, `k_master_prefix`,
      `k_master_dest`, `k_master_sheet`, `k_master_cols`.

      **First run ever — see "The item with no prior passing run" above.** Part B's gate
      item 1.4 recorded the opposite outcome as correct *at the time*, because no field
      declared `usedBy` then. Task 4 changed that; this item supersedes it.

      Confirm nothing you typed was lost: the values in the off band are the values that
      were in the main band. Set the five actions back to `Button`, Finalize, and rebuild
      before continuing.

- [x] **Per-action disclosure — one action off relegates only its own fields.**
      Set **only** `4.  Export Dashboard (.xlsx) [importexport.export]` to `Off`; leave the
      other four `importexport.*` actions at `Button`. Alt+F8 → `Core_FinalizeSetup`, then
      Alt+F8 → `Core_BuildSetup`.

      Expected: the `IMPORT / EXPORT` band is **still there**, holding `k_import`,
      `k_master_prefix`, `k_master_dest`, `k_master_sheet` and `k_master_cols`. The off band
      holds exactly the four export fields — `k_export_base`, `k_settings_sheet`,
      `k_export_offset`, `k_export_tabs` — and nothing else.

- [x] **`Suggest Import Types` keeps `k_import` visible on its own.**
      Set `1.  Clear Input Tabs [importexport.clear]` and `2.  Import Source Data
      [importexport.import]` to `Off`, leaving `Suggest Import Types
      [importexport.suggestTypes]` at `Button`. Alt+F8 → `Core_FinalizeSetup`, then Alt+F8 →
      `Core_BuildSetup`.

      **Expected — nothing moves. The off band does not appear at all.** Two actions are now
      `Off`, but every field still names at least one action that is on, so none of them is
      relegated:

      | Field(s) | `usedBy` | Still on? |
      |---|---|---|
      | `k_export_base`, `k_settings_sheet`, `k_export_offset`, `k_export_tabs` | `export` | yes — `Button` |
      | `k_import` | `clear`, `import`, `suggestTypes` | **`suggestTypes` only** |
      | `k_master_prefix`, `k_master_dest`, `k_master_sheet`, `k_master_cols` | `updateMaster` | yes — `Button` |

      `PkgClean` declares no fields, so there is nothing else on the tab. If the
      `SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON` band appears at all, something is wrong.

      **The row that matters is `k_import`** — the one field whose only surviving action is
      `suggestTypes`. On the Setup tab it is the block labelled
      **`Import files (one row per input tab)`**, under the `IMPORT FILES` sub-heading inside
      the `IMPORT / EXPORT` band: a bold three-column header reading
      `Input tab (in this workbook)` / `Source file (name starts with...)` / `Column types`,
      then four yellow rows `IN_Data1` … `IN_Data4`.

      - **Pass:** that block is where it has always been, in the normal `IMPORT / EXPORT` band.
      - **Fail:** it has dropped into the collapsed band at the bottom of the tab. That would
        mean `FieldInUse` stopped at the two `Off` actions instead of continuing to the third
        — the "one runnable is enough" rule at
        [Core.bas:412](../../src/core/Core.bas#L412) — and the fix is in Core, not the gate.

      Then Finalize again and click **Suggest Import Types**.

      You almost certainly have no `Source File A`-style workbooks sitting in the folder, and
      that is fine — this half is a smoke test that the helper can still find `k_import`'s rows,
      not a test of type inference. Expect a dialog reading
      `Suggested type maps filled for 0 import row(s)`, followed by
      `No source file found for:` and the four `IN_Data` tabs. The four `Column types` cells
      keep the values they already had ([PkgImportExport.bas:472](../../src/packages/PkgImportExport.bas#L472)
      writes the existing map back when no file matches). **That is a pass.**

      The failure signal is the *other* dialog —
      `No import rows on the Setup tab yet` — which means `Core_SetupTable` could not read the
      block at all.

      This is why `k_import` names `suggestTypes` in its `usedBy` even though that action
      reads no config key: it is a setup helper that *writes* into those rows via
      `Core_SetupSetColumn`. Relegate them and it would be filling in a collapsed band the
      user cannot see — and the dialog explicitly tells them to go review what it wrote.

## 5. Runtime behaviour on a fresh workbook (workbook D — a throwaway)

- [x] **`Core_Run` on a workbook with no `_Config` reports the action is off.**
      Workbook A has been Finalized by now, so this needs its own workbook.
      Take a **new** blank `.xlsm`, import `dist/MacrosBundle.bas`, and do **not** run
      `Core_BuildSetup` or `Core_FinalizeSetup` — there must be no `_Config` sheet at all.
      Alt+F11 → *View ▸ Immediate Window* (Ctrl+G) → type `Core_Run "clean.trim"` and press
      Enter. (Not Alt+F8: `Core_Run` takes an argument, so Excel's Macro dialog never lists it.)

      Expected: a message box titled **`Action is off`** reading

      ```
      'clean.trim' is switched Off on the Setup tab.

      Change its Availability to run it.
      ```

      **This is a deliberate behaviour change worth confirming.** Under v0.5.0 the same call
      ran the action, because an empty map meant `Button`. It now means `Off` — fresh and
      legacy are told apart by the `core.legacy` flag, and this workbook is fresh.

---

## Out of scope, recorded so it is not re-litigated

- **Issue #8 — `Core_FinalizeSetup`'s activate is delete-then-copy**, leaving a window in
  which the workbook has no `_Config` at all. This gate runs many Finalize cycles and so
  enters that window many times; it does not test or fix it. Pre-existing, filed. The
  `C-pristine.xlsm` copy in Setup is a partial mitigation, not a fix.
- **Issue #11 — the palette reskin.** Deferred until everything else is done. The logo stays
  out of the build: a `.bas` cannot carry a PNG.
- **A fresh install shows an empty catalog with no designed empty state.** Accepted
  deliberately in ADR 0003. The panel wanted a polished catalog with descriptions, grouping
  and an empty-state prompt; that was discarded as a second design project. If first-run
  confusion is actually observed during this run, **record it here as an observation** — it
  reopens as its own work, on evidence, not as a fix folded into Part A.
