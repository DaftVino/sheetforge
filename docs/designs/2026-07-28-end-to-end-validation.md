# End-to-end validation — the finished workbook

**For the owner. This is the one you run.** It replaces the six per-wave gates that would
otherwise have been written between v0.10.0 and v0.17.0, folded into a single pass over the
finished workbook, per the deferral decision of 2026-07-28 (`bde8400`).

Written for someone who has not read the code. Every item states **the observable result
first**, then what a failure looks like, then why it matters. Tick an item only against a real
run. If you skip one, write **UNTESTED** beside it — a blank tick is worse than an honest gap,
because the next person reads it as passed.

**Ordered by consequence, not by release.** The first section can destroy your work; the
second can hand you a wrong number that looks right; the rest degrade gracefully. If you only
have an hour, the hour belongs to §2 and §3.

Source: `2026-07-28-deferred-validation-log.md`, 40 claims accumulated across seven phases.
Sequencing: `2026-07-28-track2-completion-plan-of-record.md`.

---

## 1. Setup — do this once, and get it clean before anything else

- [ ] **1.1 — The bundle imports without complaint.** `pwsh build/build.ps1`, then in Excel
  remove any old `MacrosBundle` module and import `dist/MacrosBundle.bas`.
  *Expected:* no import error. *Failure:* Excel refuses the file or reports a parse error.

- [ ] **1.2 — `Debug ▸ Compile VBAProject` is clean.** *Expected:* nothing happens, which is
  what success looks like here. *Failure:* any compile error at all.
  **This is the item that discharges the one open assumption in the whole build**:
  `Select Case LCase(<Const>)` in `Core_IsEngineSheet` is believed legal VBA and is exercised
  nowhere else in this codebase. The Node syntax checker passed it, but that checker is a
  parser and explicitly not a compiler. **If this fails, stop** — the fallback is an
  `If/ElseIf` chain, and everything below is invalid until it compiles.

- [ ] **1.3 — `Core_BuildSetup` (Alt+F8) builds the Setup tab, and `Macros 1` exists
  afterwards.** *Expected:* a Setup tab listing **43 actions**, and at least one generated
  `Macros <n>` page carrying buttons.
  *Failure:* fewer than 43 actions, or no `Macros 1`.
  **Redo this until `Macros 1` exists.** The partial W1 gate left this inconclusive, and
  several items below need a built button page.

- [ ] **1.4 — Every package is actually alive.** Count the actions on the Setup tab.
  *Expected:* **43**. *Failure:* a multiple-of-a-package short — 5 missing is `PkgValidate`,
  8 is `PkgFormat`+`PkgExport`, and so on.
  *Why:* a package can build, pass every gate, and never run, if it was not registered in
  Core's candidate list. This count is the only thing that catches it.

- [ ] **1.5 — `GateAudit_Run` reports 0 FAIL.**
  `tools/GateAudit.bas` is **not** part of the bundle and is **not** produced by the build —
  it is a separate tool checked into the repo. In the VBA editor: File ▸ Import File ▸
  `tools/GateAudit.bas`, then Alt+F8 ▸ `GateAudit_Run`.
  *Expected:* a summary reporting **0 FAIL**. SKIPs are acceptable and it says why it skipped.
  *Failure:* any FAIL — or `GateAudit_Run` not appearing in the Alt+F8 list, which means the
  import did not take.
  It re-runs seven mechanical checks in one click, including that every action id is present
  in the catalog and that no unexpected id appears.
  **Remove the module afterwards** — it is a test fixture, not part of the deployment.

---

## 2. Can it destroy your work? — run these first

These are the items where a failure costs you something you cannot get back.

- [ ] **2.1 — An already-open source workbook keeps its unsaved edits.**
  Open one of the files configured in `k_import`, type something into it, **do not save**, and
  leave it open. Run **2.  Import Source Data**.
  *Expected:* the action completes; the source is **still open**; your typed text is **still
  there**; and the completion message says a source was already open and that what was read
  was its on-screen state.
  *Failure:* the source closed, or your typing gone.
  *Why:* `Workbooks.Open` handed an already-open path returns **your copy**, and closing it
  discards unsaved work — with alerts suppressed, so Excel never warns you. **This is the
  single most valuable item in this document.** Repeat it for **3.  Update Ref_Master**.

- [ ] **2.2 — Snapshot pruning leaves a lookalike file alone.**
  Put two files in the Snapshots folder: `SNAPSHOT_notes.txt` (anything) and
  `SNAPSHOT_2020-01-01_000000.xlsm`. Run **Snapshot Before a Risky Step**
  enough times to exceed the keep count.
  *Expected:* the `.txt` **still exists**; the old dated `.xlsm` is gone.
  *Failure:* the `.txt` deleted.
  *Why:* pruning deletes files. Matching the `SNAPSHOT_` prefix alone is not proof this
  toolkit wrote the file.

- [ ] **2.3 — Archiving refuses instead of overwriting, and refuses an open file.**
  Run **Archive a Generated Output** twice at the same destination; then try it on a
  file open in Excel.
  *Expected:* a refusal both times, naming the reason, with the file left where it was.

- [ ] **2.4 — A failed cross-drive archive tells you which file exists.**
  Archive to a removable or network drive and interrupt it (disconnect, or make the
  destination read-only mid-run).
  *Expected:* the error states whether the **source** still exists and whether the
  **destination** was created.
  *Failure:* a generic error, leaving you to work out by hand whether your file survived.
  *Why:* across volumes Windows implements a move as copy-then-delete, which can half-fail.

- [ ] **2.5 — `Unhide All Worksheets` leaves `_Config` hidden.**
  Hide a sheet of your own, then run it.
  *Expected:* your sheet is visible; **`_Config` and `_Config_staged` are still hidden**; the
  message says how many were left alone.
  *Failure:* `_Config` visible.
  *Why:* it holds your settings. In front of a user, it invites an edit that breaks Setup.

- [ ] **2.6 — `Protect All Sheets` does not strand the toolkit's buttons.**
  Protect everything, then run **Finalize** / rebuild the button pages.
  *Expected:* the rebuild completes and buttons are replaced normally.
  *Failure:* an error about removing controls, or buttons stacking up duplicated.
  *Why:* Excel's Protect defaults to protecting drawing objects — and the buttons **are**
  drawing objects. The engine is forbidden from unprotecting your sheet to fix it, so if this
  is wrong the stranding is permanent.

- [ ] **2.7 — `Build Table of Contents` refuses an existing sheet and a reserved name.**
  Run it twice; then set its name to `Macros 1` and run again.
  *Expected:* the second run refuses rather than clearing; the reserved name is refused by
  name.
  *Failure:* a sheet cleared, or a sheet created that the next Finalize silently deletes.

---

## 3. Can it hand you a wrong number that looks right? — run these second

- [ ] **3.1 — A chain stops at a refused step, and the later steps do not run.**
  Build a 4-step chain starting with **Combine All Worksheets**, then three cleaning actions.
  Run it once so `Master` is created. **Run it again without deleting `Master`.**
  *Expected:* the summary shows the combine's refusal with **no error number**; it says the
  remaining 3 steps were skipped; the count reads **`0/4`**; and — the load-bearing part —
  **`Master` still holds exactly what run 1 produced.**
  *Failure:* `4/4 ok`, or `Master` visibly cleaned a second time.
  *Why:* this is the defect the whole Core pass exists to remove — a daily chain cleaning last
  period's numbers and reporting success.

- [ ] **3.2 — Combined sources are contiguous.**
  Combine several sheets, at least one with only two rows.
  *Expected:* each source's rows start immediately after the previous source's last row.
  *Failure:* gaps — a two-row source landing at row 27.
  *Why:* **this is the oldest unverified change in the tree.** It is the fix for the only
  defect the partial W1 gate found, and nothing since has confirmed it.

- [ ] **3.3 — A user sheet named `Macros Data` is combined; `Macros 1` is not.**
  *Why:* the engine-sheet test requires a numeric tail. A prefix-only test would swallow your
  own `Macros Data`.

- [ ] **3.4 — `Reconcile Two Datasets` reports a duplicate key instead of pairing rows.**
  Build A and B where one key appears **twice on A** and once on B, with different values.
  *Expected:* one **`Duplicate key`** row naming counts 2 and 1, and **no `Differs` rows for
  that key.**
  *Failure:* a `Differs` row — meaning rows were paired by position, inventing a difference
  your data does not contain.

- [ ] **3.5 — A blank key is reported, never matched.** Give A and B one row each with an empty
  key. *Expected:* two **`Empty key`** rows. *Failure:* a match, or a `Differs`.

- [ ] **3.6 — A repeated header stops the reconcile before it compares.**
  Give range A two columns both headed `Amount`.
  *Expected:* a refusal naming `Amount` and side A, and **no report sheet created.**
  *Failure:* a report that silently compared only the first `Amount`.

- [ ] **3.7 — `Join / Enrich Columns` does not merge two distinct keys.**
  Put `AB  CD` (two internal spaces) and `AB CD` (one) in the reference table as separate keys,
  and both in the target.
  *Expected:* each matches its own row.
  *Failure:* both target rows getting the same reference values.
  *Why:* the key is trimmed at the ends only. Collapsing internal whitespace would map two
  genuinely different identifiers onto one.

- [ ] **3.8 — The duplicate-key count is distinct keys, not occurrences.**
  Put one key in the reference table **four** times. *Expected:* the report says **1**
  duplicated key, not 3.

- [ ] **3.9 — `Refresh All PivotTables` names a pivot it could not refresh.**
  Delete a pivot's source range, then run it.
  *Expected:* that pivot is named as failed and is **not** counted as refreshed.
  *Failure:* a clean "all refreshed" message.

- [ ] **3.10 — `Export Dashboard` still produces the same file it always did.**
  Run it on a workbook containing a dynamic-array spill, a legacy CSE array, and ordinary
  formulas. Compare against a file exported by v0.13.0 or earlier.
  *Expected:* the same flattening — every formula frozen, spills expanded to static values,
  formatting and conditional formatting intact.
  *Failure:* a `#SPILL!`, a surviving formula, or a lost array.
  *Why:* the flatten engine moved out of this action into Core. **It is the most sophisticated
  routine in the codebase and it produces a finance deliverable.**

- [ ] **3.11 — `Export Each Tab Separately` refuses a name collision.**
  Create two sheets named `A"B` and `A<B` — both legal in Excel, both sanitising to `A-B`.
  *Expected:* a refusal naming both, and **no files written at all.**
  *Failure:* one file, silently holding the second sheet.

- [ ] **3.12 — The same run emits no `Setup.pdf` or `_Config.pdf`.**
  Fix the collision from 3.11 (rename one of the two sheets), then let **Export Each Tab
  Separately** complete.
  *Expected:* one file per **data** sheet, and **no file named after the Setup tab, `_Config`,
  or any `Macros <n>` page.**
  *Failure:* `Setup.pdf` in the output folder.

---

## 4. Does it behave? — the rest

- [ ] **4.1 — A refusal inside a chain shows no dialog of its own.** During 3.1, the only
  dialog is the final summary. *Failure:* a modal box mid-chain.
- [ ] **4.2 — A refusal on its own button still shows its dialog.** Click Combine's own button
  with `Master` present. *Expected:* the full message, no chain summary.
  *Why:* 4.1 and 4.2 are a **pair** — together they prove suppression is conditional rather
  than broken. Either alone passes on a build that broke the other.
- [ ] **4.3 — A step that FAILS still asks.** Point `k_import_folder` at a folder holding a
  zero-byte `.xlsx`; run a chain starting with Combine All Workbooks.
  *Expected:* a Continue/Stop prompt. *Failure:* the chain stops without asking.
  *Why:* the asymmetry is the design. If failures stop too, the distinction was never built.
- [ ] **4.4 — The chain summary from 3.1 contains no `#0`.**
  Re-read the summary box from item 3.1 before dismissing it.
  *Expected:* the refused step's line shows its reason with **no number** after the `X`.
  *Failure:* a line reading `X  #0  <reason>`.
  *Why:* a refusal has no error number, and `#0` beside a reason tells you the macro
  malfunctioned when it did exactly the right thing.
- [ ] **4.5 — An Import/Export refusal stops a chain.** Clear `k_export_tabs`, run a chain
  containing Export Dashboard. *Expected:* the chain stops and says so.
- [ ] **4.6 — Highlighting is undoable and scoped.** Run **Highlight Duplicates** on a range
  where some cells already carry a fill. *Expected:* pre-filled cells keep their colour and are
  reported as skipped; **only the newly marked cells are left selected**; clearing the fill on
  that selection restores the sheet exactly.
  *Failure:* a pre-existing fill overwritten, or the whole picked range selected afterwards.
- [ ] **4.7 — Every confirm dialog cancels cleanly.** Cancel each action at its confirmation.
  *Expected:* nothing changed, no sheet created, no completion message.
- [ ] **4.8 — `Add a Report Title` asks before inserting rows**, and a formula using `ROW()`
  below the title reflects its new position afterwards.
  *Why:* this is the action that **admits** it changes values. The item confirms the admission
  is the true one.
- [ ] **4.9 — A part-written export batch reports how far it got.**
  Run **Export Each Tab Separately** into a folder on a removable drive and pull the drive
  mid-run. (Alternative: start the export, then make the destination folder read-only from
  another window before it finishes.)
  *Expected:* the error states **how many files were already written** before it stopped.
  *Failure:* a bare error with no count, leaving you unsure what is in the folder.
  *Why:* the write loop is sequential and does not roll back. It does not claim to — this
  confirms it at least tells you where it stopped.
- [ ] **4.10 — `Build a Summary Table` refuses a source with a blank or duplicated header**,
  and a failure leaves no half-built sheet behind.
- [ ] **4.11 — `Save Month-End Version` names the file correctly on 1 January** and on the 1st of a
  month. *Why:* prior-month-end on 1 Jan is 31 Dec of the previous year.
- [ ] **4.12 — `Snapshot Before a Risky Step` as step 1 of a destructive chain runs quietly** and the later
  steps still run.
- [ ] **4.13 — `Unprotect All Sheets` with a wrong password names the sheets that refused** and
  does not stop halfway.
- [ ] **4.14 — A TOC link opens a sheet whose name contains an apostrophe.**
  Rename a sheet to something like `Bob's Data`, run **Build Table of Contents** (into a fresh
  name — it refuses an existing one), then click that sheet's link on the TOC.
  *Expected:* the `Bob's Data` sheet becomes active.
  *Failure:* an error box, or the click doing nothing.
  *Why:* the sheet name is embedded inside a quoted reference, so the apostrophe has to be
  doubled or the reference is malformed.
- [ ] **4.15 — `Unhide All Rows & Columns` refuses on an active chart sheet** rather than
  raising a type mismatch, and says what an active AutoFilter means for the result.

---

## 5. Appearance — carried over from track 1, still unrun

- [ ] **5.1 — Legibility on a dark Excel theme.** Switch Excel to a dark theme and look at the
  Setup tab, a generated `Macros <n>` page, and the three highlight colours.
  *Failure:* text that disappears, or a fill that swallows its own text.
  *Why:* **this is the likeliest real finding in the whole document.** The palette came off a
  dark-background proof and lands on an explicitly white sheet, and nothing has ever checked
  the two together.
- [ ] **5.2 — The remaining 7 track-1 items**: hardening 1.3, 2.2, 3.1, 4.1, 4.2 · reorder 2
  (clipping half), 3. Their full text is in `2026-07-27-core-hardening-manual-gate.md` and
  `2026-07-27-action-catalog-first-manual-gate.md` — **this document does not restate them.**
  **Two cannot be run against the bundle you just imported and need their own build first.**
  Hardening 3.1 needs a deliberately hostile `PkgClean` declaring the reserved id
  `core.legacy`: edit a copy, run `pwsh build/build.ps1`, confirm the build **fails**, discard
  the edit. Reorder 3 needs the v0.6.0 bundle rebuilt from commit `812d88c`.
  **Do these two last and rebuild from `main` afterwards** — both leave a bundle you do not
  want to keep.

---

## Result

- Items passed: ____ / 41
- Items recorded **UNTESTED**: ____
- Anything that failed, with what you saw:

**If §2 is all green, nothing in this release can destroy your work.** That is the bar worth
holding this document to; everything else is recoverable by re-running something.
