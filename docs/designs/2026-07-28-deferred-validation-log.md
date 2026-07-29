# Deferred validation log

Run-time-only claims, accumulated per phase, for the **single** end-to-end validation at the
end of track 2. **This is not a checklist to work through now.** The W1 gate was deferred on
2026-07-28 (`bde8400`) because the answers are available statically; anything a careful read
or a build guard can settle is settled there instead, and only what genuinely needs Excel
lands here.

Each entry states the observable result first, then what a failure looks like, then why —
`CLAUDE.md` §7. **An entry whose expected result equals the default state proves nothing**,
so each pairs something that must happen with something that must not.

Sequencing and the reasoning for deferring:
`2026-07-28-track2-completion-plan-of-record.md` §3.5.

---

## Carried in from before track 2

These predate the phases below and are still unrun.

- **The `bafb905` used-range fix has never run in Excel.** It is the fix for the only defect
  the partial W1 gate found — every source landing ~15 rows below the last, because
  `PkgImport_AppendSheet` took its extent from `UsedRange`. Items 1.5 and 2.4 of
  `2026-07-28-w1-manual-gate.md` exist to confirm it. **Start here**; it is the oldest
  unverified change in the tree and everything since sits on top of it.
- **W1 gate item 0.3 is inconclusive** — `Macros 1` did not exist after the run. Redo it
  before anything that depends on a built button page.
- **8 track-1 gate items are unrun.** Hardening 1.3, 2.2, 3.1, 4.1, 4.2 · reorder 2
  (clipping half), 3 · reskin 4. **Reskin 4 is the likeliest real finding:** the palette came
  off a dark-background proof and lands on an explicitly white sheet, and nothing has checked
  the two together.

---

## Phase 1 — the Core refusal pass (v0.11.0, contract v2.4.0)

Everything this phase changed shows up only at run time, inside a chain. The build cannot see
a chain halt.

- **A chain stops at a refused step, and says so.** Build a 4-step chain beginning with
  `import.combine_worksheets`, run it once so `Master` is created, then run it **again**
  without deleting `Master`. Expected: the refusal appears in the summary with **no error
  number**; the summary says the remaining 3 steps were skipped; the count reads `0/4`; and —
  the load-bearing part — **`Master` still holds exactly what run 1 produced**, because the
  clean steps did not touch it. *Failure looks like:* `4/4 ok`, or `Master` visibly cleaned a
  second time. *Why:* this is the exact wrong-number-with-a-tick case the phase exists to
  remove.
- **A refusal inside a chain shows no dialog of its own.** During the run above, the only
  dialog is the final chain summary. *Failure looks like:* a modal refusal box mid-chain.
  *Why:* `Core_Refuse` suppresses its own presentation when quiet, exactly as `Core_Done`
  does.
- **A refusal standalone still shows its dialog.** Click the combine's own button with
  `Master` present. Expected: the full refusal message, and no chain summary. *Why:* this is
  the pair for the item above — together they prove suppression is **conditional**, not
  broken. Either one alone passes on a build that broke the other.
- **A step that FAILS still asks.** Point `k_import_folder` at a folder holding a zero-byte
  `.xlsx` and run a chain starting with `import.combine_workbooks`. Expected: a Continue/Stop
  prompt. *Failure looks like:* the chain stops without asking. *Why:* the asymmetry is
  settled decision 2 of the phase-1 plan; if failures stop too, the distinction was never
  implemented.
- **No summary line reads `#0`.** *Why:* a refusal has no error number, and `#0` beside a
  reason tells the user their macro malfunctioned when it did exactly what it should.
- **A user sheet named `Macros Data` is still combined, and `Macros 1` is still excluded.**
  *Why:* `Core_IsEngineSheet` replaced `PkgImport`'s hardcoded copy of the seven engine sheet
  names. This is the only thing that proves the numeric-tail rule survived the move — a
  prefix-only test would swallow `Macros Data`.
- **`Select Case LCase(<Const>)` compiles.** Excel `Debug ▸ Compile VBAProject` is clean.
  **Marked as an assumption** (`CLAUDE.md` §7): it is believed legal VBA and is exercised
  nowhere else in this codebase. The phase-1 plan's Task 3 states the `If/ElseIf` fallback.
  The Node syntax checker passed it, but that checker is a parser and explicitly not a
  compiler.
- **An already-open source workbook survives `importexport.import` and
  `importexport.updateMaster` with its unsaved edits intact.** Open one of the configured
  source files, type something into it, do NOT save, then run the action. Expected: the action
  completes, the source is **still open**, the typed text is **still there**, and the
  completion message says a source was already open and that what was read was its on-screen
  state. *Failure looks like:* the source closes, or the typed text is gone. *Why:*
  `Workbooks.Open` on an already-open path returns the user's copy, and `Close
  SaveChanges:=False` under suppressed alerts discards it without a prompt. This was the W1
  Critical for the combine; the same defect was in these two shipped actions and was found
  only by reviewing this pass's diff. **This is the highest-value item in this log** — it is
  the only one where a failure destroys the user's work rather than reporting a wrong number.
- **`importexport.*` refusals now stop a chain too.** Clear `k_export_tabs` and run a chain
  containing `importexport.export`. Expected: the chain stops, and says so. *Failure looks
  like:* the chain continues past it, or reports it as ok. *Why:* those five actions
  previously showed a message and returned **without** setting the failed flag, so a chain
  ticked them green — Task 4b. This is a behaviour change to shipped actions and is the one
  item here that touches something the owner has used before.

---

## W2 — Validate & Reconcile (v0.13.0, contract v2.5.0)

- **`validate.reconcile` reports a duplicate key instead of pairing rows.** Build A and B
  where one key appears twice on side A and once on side B, with different values. Expected:
  one `Duplicate key` row naming counts 2 and 1, and **no `Differs` rows for that key**.
  *Failure looks like:* a `Differs` row, which means rows were paired by position — a
  difference the data does not contain. *Why:* this is the wrong-number-in-a-finance-report
  case the action exists to avoid.
- **A blank key is reported, never matched.** Give A and B one row each whose key cells are
  empty. Expected: two `Empty key` rows. *Failure looks like:* a match, or a `Differs`.
- **A repeated header stops the run before comparing.** Give range A two columns both headed
  `Amount`. Expected: a refusal naming `Amount` and side A, and **no report sheet created**.
- **Colouring is undoable and scoped.** Run `validate.highlight_duplicates` on a range where
  some cells already carry a fill. Expected: the pre-filled cells keep their own colour and
  are reported as skipped; only the newly marked cells are left selected; clearing the fill
  on that selection restores the sheet exactly. *Failure looks like:* a pre-existing fill
  overwritten, or the whole picked range selected afterwards — which would strip fills the
  action deliberately did not set.
- **Every confirm dialog cancels cleanly.** Cancel each of the five actions at its
  confirmation. Expected: nothing coloured, no sheet created, no completion message.
- **Colour legibility on a dark Excel theme.** The three highlight colours against a dark
  theme. *Why:* pairs with the outstanding track-1 reskin item — same untested assumption,
  same palette origin.

---

## W3 — Format & Present, plus export (v0.14.0, contract v2.6.0)

- **`importexport.export` still produces the same file it did before the refactor.** Run it on
  a workbook containing a dynamic-array spill, a legacy CSE array, and ordinary formulas.
  Expected: byte-for-byte the same flattening as v0.13.0 — every formula frozen, spills
  expanded to static values, formatting and conditional formatting intact. *Failure looks
  like:* a `#SPILL!`, a surviving formula, or a lost array. *Why:* the flatten engine moved out
  of this shipped action into Core. **This is the highest-risk item in W3** — it is the most
  sophisticated routine in the codebase and it produces a finance deliverable.
- **`export.per_sheet` refuses a name collision instead of overwriting.** Create two sheets
  named `A"B` and `A<B` — both legal in Excel, both sanitising to `A-B`. Expected: a refusal
  naming both, and **no files written at all**. *Failure looks like:* one file, silently
  holding the second sheet.
- **`export.per_sheet` emits no `Setup.pdf` or `_Config.pdf`.** *Why:* `Core_IsEngineSheet`.
- **A partly-written batch reports how far it got.** Make the destination unwritable partway
  through. Expected: the failure names how many files were already written. *Why:* the write
  loop is sequential with no rollback, and the plan no longer claims otherwise.
- **`format.report_title` asks before inserting rows, and a formula using `ROW()` below the
  title reflects its new position afterwards.** *Why:* this is the action that admits it
  changes values; the item exists to confirm the admission is the true one.

---

## W4 — Summarize & Enrich (v0.15.0, contract v2.7.0)

- **`enrich.join` does not merge two distinct keys.** Put `AB  CD` (two internal spaces) and
  `AB CD` (one) in the reference table as separate keys, and both in the target. Expected: they
  match their own rows, independently. *Failure looks like:* both target rows getting the same
  reference values. *Why:* an earlier version normalised keys with a function that collapses
  internal whitespace. **This is the wrong-numbers-in-a-finance-deliverable case for W4.**
- **The duplicate-key count is distinct keys, not occurrences.** Put one key in the reference
  table **four** times. Expected: the report says **1** duplicated key, not 3.
- **`summarize.refresh_pivots` reports a pivot it could not refresh.** Delete a pivot's source
  range, then run it. Expected: that pivot is named as failed and is **not** counted as
  refreshed. *Failure looks like:* a clean "all refreshed" message.
- **`summarize.build_pivot` refuses a source with a duplicate or blank header**, before
  building anything.
- **A failed `build_pivot` leaves no half-built sheet behind**, and reports the original error
  rather than a cleanup error.

---

## W5 — Protect & Back Up (v0.16.0, contract v2.8.0)

- **Pruning leaves a lookalike file alone.** Put a file named `SNAPSHOT_notes.txt` and another
  named `SNAPSHOT_2020-01-01_000000.xlsm` in the Snapshots folder, then run `backup.snapshot`
  enough times to trigger pruning. Expected: the `.txt` **still exists**; the old dated `.xlsm`
  is gone. *Failure looks like:* the `.txt` deleted. *Why:* a prefix match is not proof the
  action wrote the file. **This is the destroys-user-data item for W5.**
- **`backup.archive_output` refuses rather than overwrites**, and refuses a file open in this
  Excel instance. Try both.
- **A cross-volume move that fails reports which files exist.** Archive to a different drive
  and interrupt it (remove the drive, or make the destination read-only mid-run). Expected: the
  error names whether the source still exists and whether the destination was created. *Failure
  looks like:* a generic error leaving you unsure whether your file survived.
- **`backup.month_end_version` names the file correctly on 1 January** and on the 1st of a
  month. *Why:* prior-month-end on 1 Jan is 31 Dec of the previous year.
- **`backup.snapshot` as step 1 of a destructive chain runs quietly** and the later steps still
  run. *Why:* this is the intended flagship usage.
