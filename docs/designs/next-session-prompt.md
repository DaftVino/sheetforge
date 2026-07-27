# Next session — paste this into a fresh chat

> Delete this file once Part A execution is underway; it is a launch pad, not a document.

## Start here

Run `/orient` before anything else, including clarifying questions. Then execute
`docs/designs/2026-07-27-part-a-implementation-plan.md` task by task using
`superpowers:subagent-driven-development`, dispatching the mechanical work to codex via
`/codex` — luna and terra for lighter tasks, sol for heavier ones. Opus reviews every
task before it is committed, and **Opus implements Tasks 1 and 3 itself**: they touch the
contract and `_Config`, which the owner reserves for Opus. Task 4 is mechanical and is the
best candidate to hand to codex outright.

The plan is written, reviewed and approved. Do not re-plan it, do not re-open its
decisions, and do not brainstorm — start at Task 1.

## Read first

Read nothing else until you have these, in order:

1. `docs/designs/2026-07-27-part-a-implementation-plan.md` (27.6K) — the phase contract:
   five tasks with the exact VBA for every step, plus the adapted verification cycle
   (TDD is off here, so it is build → syntax-check → commit, never a test suite).
2. `docs/adr/0003-availability-default-and-legacy-migration.md` (6.5K) — the binding
   decisions, including why three unanimous panel resolutions were NOT adopted. Read
   before touching any default.
3. `spec/contract.md` (46.1K) — **over the slice threshold; never read front to back.**
   Task 1 edits four sections: grep for `## 5.`, `## 6.`, `## 7.`, `## 10.` and read only
   the lines around each anchor. The `availability` row is at roughly line 258 and §5's
   "Absent is not invalid" paragraph at roughly 237-244.
4. `docs/designs/2026-07-27-part-b-manual-gate.md` (15.6K) — the precedent Task 5's gate
   must match: every item names the workbook, the click path and the exact expected
   string. It also carries the recipe for rebuilding the v0.3.0 and v0.4.0 fixture
   bundles, which Task 5 Step 1 needs.

## Branch

`feat/progressive-disclosure-part-a` — rebased onto `main`, four commits, **docs only, no
VBA changed yet**. It carries the spec, ADR 0003, the panel record and the plan. Stay on
it for all five tasks and open one PR from it at the end. Do not branch again from `main`;
Part B already landed and this branch is current with it.

## Constraints

- **TDD is formally off** (`CLAUDE.md`). Excel has no headless runner. Do not propose or
  write a test suite; the gates are the build, the syntax checker, and the owner's manual
  run.
- **Contract §5 must pin `button` literally** (Task 1 Step 2). §5 currently defers to "the
  documented default", and the moment §6's default becomes `off` that sentence silently
  flips the legacy path and strips every button from a v0.3.0 deployment. This is the
  single most important step in the plan; Task 5's gate item 2 is its regression test.
- **`SCHEMA_VER` must stay `"v2.0"`.** ADR 0003 keeps it deliberately so v0.5.0 configs
  remain readable. If a build ever reports `"v3.0"`, something is wrong.
- **Never read `src/core/Core.bas` whole** (210K) and never read an `.xlsm`. Grep
  `docs/code-map.md` for a symbol, then read only the slice around the anchor.
- **Do not reintroduce `clean_keycols` or `clean_hasheader`.** Asking every time is the
  feature, and it is now a normative rule in contract §8.
- **`AvailLabel`'s `Case Else` fails closed, it does not flip to `Off`.** It is a live bug
  (issue #10), not a cosmetic choice — the value round-trips into stored config. Task 2
  fixes it, and must land before Task 3.
- **Gate fixtures are rebuilt, never reused.** Workbook C was consumed by the Part B gate;
  reusing it would pass a legacy test vacuously.
- **No attribution footers** in any commit message or PR body.

## Exit criteria

- `pwsh build/build.ps1` exits 0 reporting `Core + 2 package(s)` with balanced Sub and
  Function counts
- `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` reports no diagnostics
- `grep -n 'SCHEMA_VER     As String' dist/MacrosBundle.bas` still shows `"v2.0"`
- `spec/contract.md` records v2.1.0 in §10, with §5 pinning `button` for the legacy case
  independently of §6
- Every checkbox in `docs/designs/2026-07-27-part-a-manual-gate.md` is either ticked
  against a real run or explicitly recorded as untested — never ticked unrun
- `VERSION` contains `0.6.0` and `CHANGELOG.md` carries a dated `## [0.6.0]` heading
- A PR exists from `feat/progressive-disclosure-part-a` whose body closes #6 and #10 and
  carries the completed gate table

## Unknowns and risks

- `.vscode/` is untracked and unrelated to this work. It has been left alone deliberately
  across several sessions; it is not an oversight and should not be committed or removed.
- Issue #8 is open and out of scope: `Core_FinalizeSetup`'s activate is delete-then-copy,
  so there is a window in which the workbook has no `_Config` at all. Part A adds Finalize
  cycles during gating but does not touch that path. Do not fix it here.
- Issue #11 (the palette reskin) is deferred until everything else is done. The logo stays
  out of the build — a `.bas` cannot carry a PNG. Do not fold it into this work.
- Nothing in Part A has been run in Excel. All five tasks are unverified behaviourally
  until the owner runs Task 5's gate; the structural gates cannot see a dialog or a
  rendered band.
- Task 5 gate item 7 exercises the empty-band guard's second direction for the first time
  ever. It was untestable before Task 4 wired `usedBy`, so it has no prior passing run to
  compare against — a failure there may be a genuine Core bug rather than a Part A
  regression.
- The `usedBy` mapping in Task 4 was derived from each action's `Core_Cfg` calls, but a
  typo in a `usedBy` id is silent by design (`FieldInUse` treats an unknown id as in-use).
  Task 4 Step 5's two greps are the only check; do not skip them.
- Leftover SDD files remain under `.superpowers/sdd/`. The directory is gitignored and two
  delete attempts were refused by the permission classifier, so the owner should remove it
  by hand. Nothing depends on it.
