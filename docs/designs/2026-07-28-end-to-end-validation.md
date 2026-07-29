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

Source: `2026-07-28-deferred-validation-log.md`, 40 claims accumulated across seven phases,
plus the claims the Macros Home phase added as it was built — **79 items in total**.
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
  afterwards.** *Expected:* a Setup tab listing **51 actions**, and at least one generated
  `Macros <n>` page carrying buttons.
  *Failure:* fewer than 51 actions, or no `Macros 1`.
  **Redo this until `Macros 1` exists.** The partial W1 gate left this inconclusive, and
  several items below need a built button page.

- [ ] **1.4 — Every package is actually alive.** Count the actions on the Setup tab.
  *Expected:* **51**.
  *Failure:* short by a whole package's worth. Use this table to identify which — but note
  that three counts are ambiguous, so a short count narrows the suspects rather than naming
  one:

  | Missing | Package(s) it could be |
  |---|---|
  | 2 | `PkgEnrich` |
  | 3 | `PkgImport` |
  | 4 | `PkgBackup` **or** `PkgSummarize` |
  | 5 | `PkgExport`, `PkgImportExport` **or** `PkgValidate` |
  | 7 | `PkgFormat` **or** `PkgWorkbook` |
  | 9 | `PkgClean` |

  *Why:* a package can build, pass every gate, and never run, if it was not registered in
  `CoreCandidates()`. This count is the only thing that catches it. To settle an ambiguous
  count, look at which action captions are missing rather than counting again.

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
  leave it open. Run **Import Source Data**.
  *Expected:* the action completes; the source is **still open**; your typed text is **still
  there**; and the completion message says a source was already open and that what was read
  was its on-screen state.
  *Failure:* the source closed, or your typing gone.
  *Why:* `Workbooks.Open` handed an already-open path returns **your copy**, and closing it
  discards unsaved work — with alerts suppressed, so Excel never warns you. **This is the
  single most valuable item in this document.** Repeat it for **Update Ref_Master**.

- [ ] **2.2 — Snapshot pruning leaves a lookalike file alone.**
  **Setup first**, or you cannot do this item: on the Setup tab fill in
  *Folder to write backup and snapshot files into* (`k_backup_folder`) and set *Snapshots to
  keep* (`k_backup_snapshot_keep`) to **2** — it defaults to 20, and 21 runs is not a test.
  Finalize. Snapshots are written to **`<that folder>\Backup\Snapshots\`**, two levels down;
  the folders are created on the first run, so run the action once and let it build them.
  Now put two files in that `Snapshots` folder: `SNAPSHOT_notes.txt` (any content) and
  `SNAPSHOT_2020-01-01_000000.xlsm` (copy any workbook and rename it). Run **Snapshot Before a
  Risky Step** three more times.
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
  Run it twice. Then — **this is the part that is easy to get wrong** — the destination name is
  not a prompt, it is the Setup field *Name for the Table of Contents sheet*
  (`k_workbook_toc_name`, default `Contents`). Change that field to `Macros 1`, click
  **Finalize Setup Tab**, and only then run the action again. The action reads the *saved*
  config, so without the Finalize you have simply run the first test a third time.
  *Expected:* the second run refuses rather than clearing; the third refuses the reserved name,
  naming it.
  *Failure:* a sheet cleared, or a sheet created that the next Finalize silently deletes. Also
  a failure: a third run whose refusal message names `Contents` rather than `Macros 1` — that
  means the new name never reached the action and the reserved-name path was never exercised.

- [ ] **2.8 — `Add A Back To Contents Link To Every Sheet` never overwrites.**
  First run `Build Table of Contents` so a `Contents` sheet exists. Then put a value in `A1`
  of one sheet and a formula in `A1` of another, and run this action.
  *Expected:* every other sheet gets a working link in `A1`; those two are **skipped**, and
  both are named in the report **with what was found in the cell**.
  *Failure:* either cell overwritten, either sheet silently skipped without being named, or
  any row inserted anywhere.
  *Why:* this is the only action in the toolkit that writes into sheets you own. Inserting a
  row instead of skipping would move every absolute reference on the sheet, which is not a
  gentler version of overwriting.

- [ ] **2.9 — `Add A Back To Contents Link To Every Sheet` refuses with no Contents sheet.**
  Delete the `Contents` sheet, then run the action.
  *Expected:* it refuses, naming the missing sheet and telling you to run
  `Build Table of Contents` first. **No sheet is modified.**
  *Failure:* it creates the contents sheet itself, or writes links pointing at nothing.

- [ ] **2.10 — `Email A Sheet As An Attachment` writes the file first and never sends.**
  Fill in the recipient and subject on Setup, Finalize, then run it and pick a sheet.
  *Expected:* the `.xlsx` exists on disk **before** any Outlook window appears; an Outlook
  draft opens with that file attached; **nothing is sent**. Close the draft without sending.
  *Failure:* a sent message, an empty draft, or a draft with no attachment.
  *Why:* a macro that sends unattended is one wrong range away from mailing the wrong numbers
  to a client.

- [ ] **2.11 — `Email A Sheet As An Attachment` keeps the file when Outlook is unavailable.**
  Only runnable where Outlook is absent or blocked by policy. **Skip and say so if Outlook
  works here** — do not fake it.
  *Expected:* a refusal naming the **full path** of the file it already wrote, and that file
  is still on disk.
  *Failure:* the file deleted, the path missing from the message, or a raw VBA error.

- [ ] **2.12 — `Save Each Group As Its Own File` refuses the whole batch on a collision.**
  Make two rows whose grouping values differ only by a character illegal in a filename — for
  example `A/B` and `A:B` — and run it.
  *Expected:* a refusal naming both values. **Zero files written.**
  *Failure:* any file written, or one group silently overwriting another.
  *Why:* the safe-name reduction is many-to-one, so two different values can produce one
  filename. Writing until something breaks would leave a folder that looks complete.

- [ ] **2.13 — `Split One Column Into Many` refuses rather than overwriting.**
  Put data in the column immediately right of a column that needs splitting into two, then
  run it with the destination set to `SOURCE`.
  *Expected:* a refusal naming the first occupied cell. **Nothing changed.**
  *Failure:* any cell overwritten.

- [ ] **2.14 — `Highlight Comments, Formulas and Blanks` is genuinely reversible.**
  Give a cell your own fill colour first. Run the action, then run it again and accept the
  removal.
  *Expected:* after the second run your original fill is **back**, and any unrelated
  conditional formatting on the sheet is untouched.
  *Failure:* your fill lost, a residual audit colour, or an unrelated rule deleted.
  *Why:* this is the claim that makes the action safe to run at all. It uses tagged
  conditional-format rules rather than direct fills precisely so removal can tell its own
  marking from yours — if that fails, running it costs you formatting you cannot recover.

- [ ] **2.15 — `Complete Macro Setup` refuses on a workbook that has never been finalized.**
  Import into a clean workbook, run `Core_BuildSetup`, then go to Setup Menu and click
  `Complete Macro Setup` **without** finalizing.
  *Expected:* a refusal telling you to Finalize first. **Every tab still there.**
  *Failure:* anything deleted.

- [ ] **2.16 — `Complete Macro Setup` refuses when there are no macro pages.**
  Set all 51 actions to `Off`, Finalize, then click it.
  *Expected:* a refusal saying it would leave you with nothing. **Nothing deleted.**
  *Failure:* it strips the setup tabs off a workbook that has no buttons either, leaving no
  on-sheet way back.

- [ ] **2.17 — Answering No deletes nothing, and Enter alone answers No.**
  Click it on a properly finalized workbook, read the confirmation, press **Enter** without
  clicking.
  *Expected:* nothing happens; every tab still there.
  *Failure:* Enter triggering the deletion. The dialog defaults to No deliberately.

- [ ] **2.18 — After Yes, the setup tabs are gone AND every button still works.**
  Click it and answer Yes. Then click one of your macro buttons.
  *Expected:* `Setup`, `Setup Menu` and `Macros Home` are gone; the `Macros <n>` pages remain;
  **and the button you clicked actually runs.**
  *Failure:* the button doing nothing, or reporting the action is not available.
  ***This is the single most important item in this section.*** It proves `_Config` survived.
  If `_Config` were deleted, every button would still be sitting there looking normal and
  would silently do nothing — the worst possible failure shape, because nothing on screen
  says anything is wrong.

- [ ] **2.19 — B1 of every remaining page is empty, with no dead link.**
  *Expected:* the `<< Home` cell is blank on every `Macros <n>` page.
  *Failure:* text that still looks like a link, or a click that raises an error because the
  sheet it pointed at is gone.

- [ ] **2.20 — `Alt+F8` then `Core_BuildSetup` brings the Setup tab back, with your settings
  still in it.** *Expected:* Setup is rebuilt and your answers are still there — they were
  read back out of the surviving `_Config`.
  *Failure:* a blank Setup tab. That means the settings did not survive, and the confirmation
  dialog promised a way back that does not exist.

- [ ] **2.21 — `Alt+F8` lists exactly three names.**
  *Expected:* `Core_BuildSetup`, `Core_FinalizeSetup`, `Core_CompleteSetup` — and nothing
  else.
  *Failure:* **any fourth name at all.** Write down whatever it is; the name identifies the
  cause on its own.
  - `Core_SpeedOn` or `Core_SpeedOff` — running either leaves Excel with calculation and
    screen updating switched off and nothing on screen saying so, which then looks like the
    toolkit broke the workbook.
  - `Core_BuildControlPanel` — it was `Public` until a dry-run caught it, and it is the reason
    this item's failure clause is open-ended rather than naming two suspects. It rebuilds the
    `Macros <n>` pages from state that Finalize prepares, so a user who picks it out of this
    list gets a panel built from a stale registry.
  *Why:* this list is the only inventory of macro names a non-technical user ever sees, and
  `docs/setup-guide.md` tells them to pick from it. Every extra name is one they must rule out.

  **Assumption, flagged per rule 7:** that a macro can delete the sheet hosting the button
  that invoked it. Nothing in this codebase has ever exercised that. `Setup Menu` is deleted
  **last** so a failure leaves the most recoverable state. If Excel refuses, the fallback is
  `Application.OnTime` to defer the final delete past the end of the handler.

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

- [ ] **3.10 — `Export Dashboard (.xlsx)` still produces the same file it always did.**
  Run it on a workbook containing a dynamic-array spill, a legacy CSE array, and ordinary
  formulas. Compare against a file exported by v0.13.0 or earlier. **You need that older file
  before you start** — either one you already exported and kept, or one you make now:
  `git stash && git checkout v0.13.0 && pwsh build/build.ps1`, import that bundle into a
  scratch copy of the workbook, export, then `git checkout feat/macros-home` and rebuild.
  **If you have neither, record this item UNTESTED rather than guessing** — a comparison
  against nothing is not a pass.
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

- [ ] **3.13 — `Split One Column Into Many` keeps leading zeros.**
  Put `00123,North` in a cell, set the delimiter to `COMMA` and the destination to `SOURCE`,
  and run it.
  *Expected:* two cells reading **`00123`** and `North` — the zeros still there, the cell
  still text.
  *Failure:* `123`. That is the whole reason every output field is forced to text, and it is
  invisible until someone reconciles an account code that no longer matches.

- [ ] **3.14 — `Change Case` does not touch formulas or numbers.**
  Select a range containing text, a number, and a formula returning text. Run it, choose
  `Proper`.
  *Expected:* only the constant text changes. The formula **is still a formula**, the number
  is still a number, and the report states how many formula cells were skipped.
  *Failure:* a formula replaced by its result, a number turned into text, or a silent skip
  count.
  *Why:* flattening a formula here would be indistinguishable from correct output until the
  next recalculation that should have changed it and didn't.

---

## 4. Does it behave? — the rest

- [ ] **4.1 — A refusal inside a chain shows no dialog of its own.** During 3.1, the only
  dialog is the final summary. *Failure:* a modal box mid-chain.
- [ ] **4.2 — A refusal on its own button still shows its dialog.** Click Combine's own button
  with `Master` present. *Expected:* the full message, no chain summary.
  *Why:* 4.1 and 4.2 are a **pair** — together they prove suppression is conditional rather
  than broken. Either alone passes on a build that broke the other.
- [ ] **4.3 — A step that FAILS still asks.** Point `k_import_folder` at a folder holding a
  zero-byte `.xlsx`; run a chain starting with **Combine All Workbooks in a Folder**.
  *Expected:* a Continue/Stop prompt. *Failure:* the chain stops without asking.
  *Why:* the asymmetry is the design. If failures stop too, the distinction was never built.
- [ ] **4.4 — The chain summary from 3.1 contains no `#0`.**
  Re-read the summary box from item 3.1 before dismissing it.
  *Expected:* the refused step's line shows its reason with **no number** after the `X`.
  *Failure:* a line reading `X  #0  <reason>`.
  *Why:* a refusal has no error number, and `#0` beside a reason tells you the macro
  malfunctioned when it did exactly the right thing.
- [ ] **4.5 — An Import/Export refusal stops a chain.** Clear `k_export_tabs`, run a chain
  containing **Export Dashboard (.xlsx)**. *Expected:* the chain stops and says so.
- [ ] **4.6 — Highlighting is undoable and scoped.** Run **Highlight Duplicates** on a range
  where some cells already carry a fill. *Expected:* pre-filled cells keep their colour and are
  reported as skipped; **only the newly marked cells are left selected**; clearing the fill on
  that selection restores the sheet exactly.
  *Failure:* a pre-existing fill overwritten, or the whole picked range selected afterwards.
- [ ] **4.7 — Every confirm dialog cancels cleanly.** Cancel each action at its confirmation.
  *Expected:* nothing changed, no sheet created, no completion message.
- [ ] **4.8 — `Add Report Title` asks before inserting rows**, and a formula using `ROW()`
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

### New in this phase — Macros Home

- [ ] **4.16 — `Macros Home` exists after Finalize and lists all 51 actions.**
  *Expected:* a `Macros Home` tab whose Action Catalog has **51 rows**, plus the chains.
  *Failure:* the tab missing, or fewer than 51 rows.

- [ ] **4.17 — An action set to `Off` still appears on `Macros Home`, marked `Off`, with no
  page link.** Set one action to `Off` on Setup, Finalize, then read its row.
  *Expected:* the row is **present**, its Availability reads `Off`, and its PAGE cell is
  empty.
  *Failure:* the row missing entirely.
  ***This is the item that proves the contract §6 rewrite actually shipped.*** The withdrawn
  version listed only actions that became buttons. Since every action arrives `Off`, that
  version renders an empty page on a fresh import — which is exactly when someone needs to
  read what the toolkit does.

- [ ] **4.18 — Every `Macros <n>` page has a working `<< Home` link in B1, on the FIRST
  build.** Import into a clean workbook, Finalize once, and click `<< Home` on each page.
  *Expected:* it is a **link** and it opens `Macros Home`.
  *Failure:* plain bold text instead of a link, on any page.
  ***The words "first build" are the whole item.*** A guard was once added here that made
  every Home link dead text on a first build and correct on the second. Finalizing twice
  before checking would hide exactly the defect this is looking for.

- [ ] **4.19 — Finalize lands you on `Macros Home`.** *Expected:* after Finalize the active
  tab is `Macros Home`, not `Macros 1` and not Setup.

- [ ] **4.20 — With every action `Off`, Finalize still produces `Macros Home`.**
  Set all 51 to `Off` and Finalize.
  *Expected:* `Macros Home` is built and lists all 51 as `Off`. No error.
  *Failure:* an error, or a workbook with nowhere to land.
  *Why:* there are no button pages in this state, so before this phase there was no page to
  land on at all.

### New in this phase — chains

- [ ] **4.21 — A chain authored as `clean.dupes, clean.trim` finalizes and runs.**
  Type exactly that into a chain's Steps cell, give it a name, Finalize, click its button.
  *Expected:* Finalize reports no problems; the button runs both actions in order.
  *Failure:* `(no steps!)` or `[UNKNOWN]` in the Finalize report.

- [ ] **4.22 — A chain authored in the OLD format still finalizes.**
  Type `Remove Duplicates [clean.dupes]` on one line and `Trim & Clean Text [clean.trim]` on
  the next, into a chain's Steps cell. Finalize. (Those are the two captions as the old format
  actually stored them — the parser reads the id in brackets, not the caption.)
  *Expected:* accepted, and after Finalize the cell has been **rewritten** as
  `clean.dupes, clean.trim`.
  *Failure:* `[UNKNOWN]` against either step.
  *Why:* any workbook already carrying chains stores them in the old form. If this fails, the
  upgrade breaks every deployed chain at the first Finalize.

- [ ] **4.23 — A hand-typed id in the wrong case is rewritten canonical.**
  Type `CLEAN.DUPES` into a Steps cell and Finalize.
  *Expected:* the cell now reads `clean.dupes`.
  *Failure:* the shouting casing persists.
  *Why:* this normalisation was lost by a refactor once before and restored deliberately.

- [ ] **4.24 — The copy list beside Chains shows bare ids.** *Expected:* column G lists
  `clean.dupes` and the like, with the caption beside it in H, under a header saying to copy
  an **id**. *Failure:* the old `Caption [id]` token form.

### New in this phase — the remaining actions

- [ ] **4.25 — `Sort Worksheets Alphabetically` leaves the toolkit's tabs where they are.**
  Scramble several user tabs, hide one of them, and run it.
  *Expected:* user tabs in A-Z order, the hidden one **still hidden** and in its right place;
  `Setup`, `Setup Menu`, `Macros Home` and every `Macros <n>` unmoved.
  *Failure:* any engine tab moved, or a hidden sheet made visible.

- [ ] **4.26 — `Sort Worksheets Alphabetically` is correct in a workbook containing a chart
  sheet.** Add a chart sheet (not an embedded chart), scramble the user tabs, run it.
  *Expected:* the user worksheets end in A-Z order.
  *Failure:* tabs landing in wrong positions.
  *Why:* Excel numbers worksheets and chart sheets in one sequence but this action enumerates
  worksheets only. Mixing the two index spaces was a real defect in the first draft, and a
  workbook with no chart sheet cannot detect it.

- [ ] **4.27 — `Apply One Style To All Charts` styles report charts and skips engine ones.**
  *Expected:* every chart on your own sheets takes the configured style; the count is
  reported. *Failure:* an eligible chart unchanged, or a run-time error where the workbook has
  no charts at all.

- [ ] **4.28 — The first four buttons carry no leading numbers.**
  *Expected:* `Clear Input Tabs`, `Import Source Data`, `Update Ref_Master`,
  `Export Dashboard (.xlsx)` — no `1.`, `2.`, `3.`, `4.`.
  *Failure:* any surviving number.

### New in this phase — the Finalize reminder

- [ ] **4.29 — On a fresh import with everything `Off`, the reminder line is BLANK.**
  Import into a clean workbook, run `Core_BuildSetup`, and look at the row directly above the
  `Action` / `Availability` / `Placement` headers.
  *Expected:* **empty**. A `Setup Menu` link sits beside it; that is always there.
  *Failure:* the reminder text showing when nothing is switched on.
  *Why:* this is the item that proves the formula excludes blanks and matches `Off`
  correctly. If it fires here it fires forever, from first import, and means nothing.
  **Run 4.29 and 4.30 as a pair, in that order.** On its own 4.29 would pass on a build that
  never wrote the formula at all — a blank cell is also what "no formula" looks like. 4.30 is
  what proves the cell is a live formula; 4.29 is what proves it is the *right* one.

- [ ] **4.30 — Setting one action to `Button` makes the reminder appear immediately, before
  Finalize.** Change a single Availability dropdown and do not click anything else.
  *Expected:* the line appears **as soon as the dropdown changes**, telling you to click
  `Finalize Setup Tab` and then `Create Setup Tab`.
  *Failure:* nothing appears until after Finalize — which would mean it is static text, not a
  formula, and the whole point of it is gone.

- [ ] **4.31 — Setting that action back to `Off` makes the line disappear again.**
  *Failure:* the text persisting once shown.

- [ ] **4.32 — The `Setup Menu` link beside the reminder opens the Setup Menu tab.**

- [ ] **4.33 — Finalize still works with the reminder row present.**
  *Expected:* Finalize reports the same **51 actions** and completes normally.
  *Failure:* a wrong action count, or Finalize reporting it cannot find a block.
  *Why:* the reminder inserts a row above the action catalog. It carries no machine key
  precisely so nothing can mistake it for a block — this item is what proves that held.

### New in this phase — the header notes

- [ ] **4.34 — All four headers carry a readable note.** Hover each of these on the Setup tab:
  **Availability** and **Placement** in the Action Catalog, **Steps (action ids, separated by
  commas)** and **Placement** in the Chains block, and **Column types** on the import table.
  Those are the header cells' full text — the Chains one is not a bare `Steps`.
  *Expected:* a note appears on each, and the text is **not clipped** — you can read the last
  sentence.
  *Failure:* a missing note, or one cut off mid-sentence.

- [ ] **4.35 — The `Column types` note says what an unrecognised code does.**
  *Expected:* it lists `T`, `N`, `D`, `DT`, `B`, `G` **and** says that anything else is
  treated as auto-detect.
  *Why:* that behaviour is completely invisible today. A typo in this column silently becomes
  `G`, and an ID column loses its leading zeros with nothing anywhere saying why.

- [ ] **4.36 — Running `Core_BuildSetup` twice leaves ONE note per header, and does not
  error.** Run it, then run it again without doing anything else, then hover the headers.
  *Expected:* one note each, no error dialog.
  *Failure:* a run-time error on the second build, or a doubled/stale note.
  ***This is the item that proves the note helper clears before it adds.*** `AddComment`
  raises on a cell that already carries a comment, and a single build cannot expose it —
  only the second one can.

- [ ] **4.37 — Finalize still reports 51 actions after the notes exist.**
  *Failure:* a different count, which would mean a note disturbed the block the tamper check
  reads.
  *Why:* `MarkerVerdict` reads the `Availability` header string back off the sheet. The note
  attaches to that cell; the header text itself must be untouched.

  **Assumption, flagged per rule 7:** that a note anchors correctly on a header cell whose row
  is later `AutoFit`, given the row directly above every header is merged B:E. Nothing in this
  codebase has attached a note to anything before.

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

- Items passed: ____ / 79
- Items recorded **UNTESTED**: ____
- Anything that failed, with what you saw:

**If §2 is all green, nothing in this release can destroy your work.** That is the bar worth
holding this document to; everything else is recoverable by re-running something.
