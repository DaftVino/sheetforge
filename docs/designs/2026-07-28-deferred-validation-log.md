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
- **`importexport.*` refusals now stop a chain too.** Clear `k_export_tabs` and run a chain
  containing `importexport.export`. Expected: the chain stops, and says so. *Failure looks
  like:* the chain continues past it, or reports it as ok. *Why:* those five actions
  previously showed a message and returned **without** setting the failed flag, so a chain
  ticked them green — Task 4b. This is a behaviour change to shipped actions and is the one
  item here that touches something the owner has used before.
