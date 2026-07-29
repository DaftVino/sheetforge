# Next session — paste this into a fresh chat

> Delete this file once Tasks 9–12 are underway; it is a launch pad, not a document.

## Start here

Run `/orient` before anything else, including clarifying questions. Then finish the Macros Home
phase: Tasks 9 to 12 of the implementation plan — reconcile the validation checklist and
dry-run it blind, sweep the documentation, write the non-technical setup guide, then sanitize
the repo and hand the owner a public release. All eight code tasks are already done and
committed; **do not write more VBA unless a defect forces it**.

## Read first

Read nothing else until you have these, in order:

1. `docs/designs/2026-07-29-macros-home-phase-plan.md` (141K) — **over the slice threshold, do
   not read front to back.** The four remaining tasks are lines **2741–3088**; that slice is
   your complete specification, including the exact `reading-level.mjs` invocation, the release
   commands, and the blind-dry-run instructions. Tasks 1–8 above line 2741 are done — read them
   only if you need to know why something is the way it is.
2. `docs/designs/2026-07-28-end-to-end-validation.md` (33K) — **over the slice threshold: read
   it in slices, one section at a time, from the `^## ` anchors.** The 79-item owner-facing
   checklist Task 9 reconciles and Task 11's guide must not contradict. Sections are ordered by
   consequence; the slices Task 9 changes most are section 1 (Setup) and section 2 (Can it
   destroy your work?).
3. `scripts/reading-level.mjs` (2.9K) — read whole. The executable gate binding Task 11: it
   exits 1 above Flesch-Kincaid grade 6.0, so the owner's hard rule fails the step rather than
   being asserted. Already validated in both directions.
4. `docs/designs/2026-07-24-finish-excel-design.md` (46K) — **over the slice threshold: read
   one slice only, from the anchor `## 5. Handoff log` at line 300**, and within it only the
   **top entry**. That slice is the do-not-revert list and the account of what the plan got
   wrong during execution. Do not read the rest of the file.

## Branch

`feat/macros-home` — 20 commits, all eight code tasks, 51 actions, contract v2.10.0, build
green. **Stay on it.** Nothing is merged and no PR is open; Tasks 9-12 finish here and land as
one PR against `main`.

`main` is at v0.17.0 with 10 packages — that is track 2, landed earlier today as PR #28. Do
**not** branch from it for this work; you would lose all eight code tasks. Do not cut any new
branch at all.

## Constraints

- **The do-not-revert list is the top entry of the handoff log (manifest entry 4). Read it
  before touching `src/`.** Seven items, each a defect corrected during execution that the plan
  still describes the old way.
- **`VERSION` is still `0.17.0`. Task 10 bumps it to `0.18.0`** — that bump is what the release
  depends on, and no earlier task did it.
- **Contract v2.10.0 is frozen and complete.** Tasks 9-12 change no action id and no contract
  semantics. If you find yourself editing `spec/contract.md`, stop and ask why.
- **TDD is off.** VBA has no headless runner and Excel cannot run on CI. Do not propose a test
  suite. Verification is the build, the syntax check, deliberate injections, and the owner's
  manual run.
- **Neither the build nor the syntax checker resolves symbols.** A call to a helper that does
  not exist passes both gates and fails only at Compile in Excel. Any new call site must be
  checked by hand.
- **codex is `gpt-5.6-sol` at high effort already.** Invoke it as plain
  `codex exec --sandbox workspace-write "<prompt>" < /dev/null`. Passing `--model` or `-c` made
  a run hang at "Reading additional input from stdin" and produce nothing. Its Windows sandbox
  cannot spawn shell children, so never ask it to run a build.
- **The owner performs Task 12 step 3 (making the repo public). An agent must not.** It is
  irreversible — clones and forks survive a revert.
- **`dist/` is gitignored — never `git add -f` the bundle.** `.vscode/` stays untracked. No
  attribution footer, `Co-Authored-By`, or session URL in any commit or PR body.

## Exit criteria

- `pwsh build/build.ps1` exits 0, reporting `Core + 10 package(s)` with balanced Sub and
  Function counts
- `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` reports no diagnostics
- `node scripts/reading-level.mjs docs/setup-guide.md` exits 0 — grade at or under 6.00
- `git show origin/main:VERSION` prints 0.18.0
- `docs/setup-guide.md` and its PDF exist, and every button caption, sheet name and procedure
  name the guide types is greppable in `dist/MacrosBundle.bas`
- The blind dry-run of both the checklist and the guide is recorded in the PR body, with what
  it found
- A GitHub release v0.18.0 exists carrying `MacrosBundle-v0.18.0.bas` as an asset

## Unknowns and risks

- **`.vscode/` is untracked and stays that way.** Deliberate and long-standing across sessions;
  do not commit it and do not delete it. It is the only path `git status --porcelain` reports.
- **Nothing in this phase has ever run in Excel.** 79 unrun checklist items, up from 41. The
  oldest — the `bafb905` used-range fix — is nine phases deep and still unconfirmed. A defect
  found during Task 9's dry-run may reopen a completed task rather than only adding to this
  one.
- **Two run-time claims are marked as assumptions and neither gate can reach them:** that a
  macro can delete the sheet hosting the button that invoked it (Complete Macro Setup), and
  that an AddComment note anchors correctly on a header row that is later AutoFit. Both are on
  the checklist; if either fails the fix is in `src/core/Core.bas` and this phase reopens.
- **Layout at 51 actions is unproven.** Auto-layout packs 18 buttons per page, so 51 actions
  plus chains is roughly four pages. PkgFixture was the design's mechanism for proving it and
  the owner dropped it — the only remaining proof is the owner looking at the sheet.
- **Task 12 needs the owner twice**: to make the repository public, and to choose a LICENSE if
  none exists. Steps 1 and 2 (history scan, public-facing basics) can be done first and
  reported; then stop.
- **The guide's audience cannot be tested here.** It is written for someone who has never
  opened the VBA editor. The reading-level gate measures prose difficulty, not whether the
  steps work; only the blind dry-run approximates the real reader.
- Three issues are filed and deliberately not built — **#29** Finalize auto-running Setup,
  **#30** a worked example per action, **#31** stripping code for Off actions. Do not implement
  any of them in this phase.
