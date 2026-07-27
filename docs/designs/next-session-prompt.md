# Next session — paste this into a fresh chat

> Delete this file once the v0.4.0 gate is finished; it is a launch pad, not a document.

## Start here

Run `/orient` before anything else, including clarifying questions. Then continue the
owner's manual Excel gate — Task 10 of the track-1 plan, 24 items remaining — and when
it passes, do Task 11 (tear down the fixture, refresh docs, ship v0.4.0).

## Read first

Read nothing else until you have these, in order:

1. `docs/designs/2026-07-24-finish-excel-track1-plan.md` (139.7K) — the phase contract.
   **Never read whole.** Grep for `### Task 10` and read only that slice (~line 2168
   onward); it is the live checklist the owner is working through, and Task 11 follows
   it. Its "Before you start" preamble names the three workbooks the gate needs.
2. `.superpowers/sdd/2026-07-25-track1-owner-feedback-plan/progress.md` (31.8K) — the
   execution ledger. Every review, fix round, parked finding, deferred minor, and the
   gate results so far. Records several places a prior session was wrong and corrected.
3. `docs/designs/2026-07-25-track1-owner-feedback-plan.md` (135.7K) — **read only the
   `## Handoff log` slice** (grep for it, read from there). Entry 1 is dated 2026-07-26
   and carries the do-not-revert list. The task bodies above it are already implemented
   and reviewed; do not re-read them unless a specific task needs revisiting.
4. `spec/contract.md` (42.4K) — normative for the Apps Script port. **Read in slices**:
   §3, §6, §7, §8 and §10 were all amended by this work. Grep for the section heading
   you need rather than reading front to back.
5. `docs/code-map.md` (28.6K) — current as of `8f3a719`. Grep it for a symbol, then read
   only the lines around the anchor. This is how you avoid opening `src/core/Core.bas`,
   which is 208K.

## Branch

`feat/finish-excel-track1` — 72 commits ahead of `main`, unpushed, no PR, nothing
merged. Every piece of v0.4.0 lives here. Stay on it. Do not branch from `main`: it is
still v0.3.0 with contract v1.1.0 and has none of this work.

## Constraints

- **Do not revert `AddBtn`'s all-or-nothing error handling.** It has no handler of its
  own and runs under the caller's `On Error Resume Next`; without the fix a failure
  abandoned the rest of the procedure after `Buttons.Add` had already created the
  shape, leaving a permanently unsweepable orphan named `Button 11`.
- **Do not use `eNum` or any case-variant of a VBA keyword as an identifier.** VBA is
  case-insensitive, so `eNum` IS `Enum`. This broke `Debug ▸ Compile` and is the second
  instance after `any`. `build.ps1`'s reserved-word list is empirical: add a word only
  when Excel rejects one, never on suspicion.
- **Excel does NOT forbid `!` in a sheet name** (the set is `: \ / ? * [ ]`). Do not
  "simplify" `InStrRev` to `InStr` in `ParseAnchor` — it would silently break anchors
  onto a sheet whose name contains `!`.
- **`hide` is filtered in `GatherButtons` and nowhere else.** Do not add a second
  filter; `SplitAnchored` runs strictly downstream of that funnel.
- **No contract version bump.** v1.2.0 is unreleased.
- **TDD is formally off** (CLAUDE.md, "Off"): no VBA runner, no headless Excel. Do not
  propose adding a test suite. The gates are structural only; the owner's import is the
  real verification.
- **Never read an `.xlsm`**, and never read `src/core/Core.bas` whole.

## Exit criteria

- `pwsh build/build.ps1` exits 0 with balanced Sub/Function counts, and
  `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` reports no diagnostics
- All 24 remaining Task 10 checkboxes in `docs/designs/2026-07-24-finish-excel-track1-plan.md`
  are ticked, including the protected-sheet item
- `src/packages/PkgFixture.bas` is removed and its two registration hooks are gone
  (Task 11 Step 1)
- `CHANGELOG.md` contains a released `[0.4.0]` entry
- A PR exists with the completed checklist pasted into its body
- `docs/designs/2026-07-25-track1-owner-feedback-plan.md` has a new dated handoff entry

## Unknowns and risks

- **The protected-sheet check is the highest-value item left and has never been run.**
  It rests on `ws.Buttons.Add` raising 1004 on a `DrawingObjects`-protected sheet — the
  one assumption nothing agent-side can execute. Before the fix it took down the whole
  build rather than skipping one button.
- **Three fixes have never been verified in Excel**: the leading-apostrophe repair for
  `'Sheet Name'!B4`, the sweep-on-clear, and the four Finalize-blocking conditions.
- **A chain-placement failure seen during the gate is NOT reproducible** on current code
  and is deliberately left unattributed in the ledger. A step-by-step diagnostic showed
  every property assignment succeeding. Do not invent a cause; if it recurs, it is a new
  finding.
- **`ANCHOR_BTN_W` at 110pt** was judged acceptable once and may want a second look
  against real data. Owner's call, one-line change.
- `.vscode/` is untracked and unrelated to this work — left alone deliberately, not an
  oversight. Nothing else is uncommitted.
- GitHub issue #3 proposes a blind-dry-run gate for manual checklists, assigned to the
  owner. Not blocking, but it is the systemic fix for the six unusable checklist items
  this phase produced.
- The SDD workspace at `.superpowers/sdd/2026-07-25-track1-owner-feedback-plan/` was
  deliberately NOT deleted. Delete it only after the gate passes and the PR lands.
