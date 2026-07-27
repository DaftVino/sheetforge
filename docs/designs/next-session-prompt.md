# Next session — paste this into a fresh chat

> Delete this file once the Part B gate is done; it is a launch pad, not a document.

## Start here

Run `/orient` before anything else, including clarifying questions. Then walk the owner
through `docs/designs/2026-07-27-part-b-manual-gate.md` — Part B is implemented,
committed and pushed, but **never run in Excel**. When it passes, settle the contract
version question in ADR 0002, then ship Part B and start planning Part A (issue #6).

**Model and agent routing, owner's instruction.** Offload as much as possible to
**codex via the CLI** (`/codex`): `luna` and `terra` for lighter tasks, `sol` for
heavier ones. Use Sonnet subagents for mechanical work. **Opus reviews everything**, and
Opus implements anything critical or sensitive — the contract, the migration path, and
anything touching `_Config`.

## Read first

Read nothing else until you have these, in order:

1. `docs/designs/2026-07-27-part-b-manual-gate.md` (7.4K) — the phase contract: the
   manual checklist for work already built, plus the two-workbook setup it needs.
2. `docs/adr/0002-contract-version-vs-schema-marker.md` (4.4K) — DRAFT, unaccepted. The
   one blocking decision: contract v2.0.0 vs `SCHEMA_VER` staying `v1.2`.
3. `src/packages/PkgClean.bas` (11K) — the whole implementation under test; small
   enough to read entire.
4. `spec/contract.md` (46K) — **read in slices from anchors; never front to back.**
   Grep for the anchor, then read only the lines around it. Three amendments to find:
   `VERSION DECISION PENDING` in §8, `UNRELEASED, VERSION UNDECIDED` in §10, and the
   new no-visible-fields rule at the end of §2.
5. `docs/designs/2026-07-26-setup-surface-progressive-disclosure-design.md` (13K) — the
   design both parts come from; Part A's section is what the next plan expands.
6. `docs/code-map.md` (28K) — current as of `4170ba5`. Grep for a symbol, then read only
   the lines around the anchor; `src/core/Core.bas` is 210K and must never be read whole.

## Branch

`feat/prompt-per-selection-settings` — two commits, pushed, **no PR yet**, branched from
the merge commit of PR #4. It carries all of Part B: the run-time prompts, the Core
empty-band guard, the contract amendments, the gate checklist and the ADR draft. Stay on
it until the gate passes. `main` is v0.4.0 and does **not** have this work. Branch again
off `main` only when starting Part A, and only after Part B lands.

## Constraints

- **Do not reintroduce `clean_keycols` or `clean_hasheader` as remembered settings.**
  Asking every time is the feature — it is what keeps the answer matched to the
  selection the user just made. This is now a normative rule in contract §8.
- **Do not bump `SCHEMA_VER`.** It is `"v1.2"` and `CurrentSchemaMajor()` gates
  migration on it. Moving its major refuses every existing v0.4.0 `_Config` outright,
  for a change that alters no stored layout. This is the whole subject of ADR 0002 and
  must not be settled by editing the constant.
- **`AskHasHeader` uses `vbYesNoCancel` on purpose.** On a two-button `vbYesNo` MsgBox
  the title-bar X returns `vbNo`, which here silently means "compare every row" and
  deletes a header row. Do not simplify it to two buttons.
- **`AskKeyCols` uses `Application.InputBox`, not VBA's `InputBox`.** VBA's returns `""`
  for both Cancel and an empty OK; those must not mean the same thing.
- Carried forward from v0.4.0 and still binding: `AddBtn`'s all-or-nothing placement; no
  `eNum` or any case-variant of a VBA keyword as an identifier; `InStrRev` not `InStr`
  in `ParseAnchor`; `hide` filtered only in `GatherButtons`; `NO_PAGES_MSG` stays one
  constant.
- TDD is formally off (CLAUDE.md, "Off") — no VBA runner, no headless Excel. Do not
  propose a test suite. The gates are structural; the owner's import is the verification.
- Never read an `.xlsm`, and never read `src/core/Core.bas` whole.

## Exit criteria

- `pwsh build/build.ps1` exits 0 reporting `Core + 2 package(s)` with balanced
  `Sub`/`Function` counts, and `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`
  reports no diagnostics
- Every checkbox in `docs/designs/2026-07-27-part-b-manual-gate.md` is ticked, including
  section 5's upgrade pass on a workbook finalized under v0.4.0
- `docs/adr/0002-contract-version-vs-schema-marker.md` has `**Status:** Accepted` or
  `Rejected` — not `DRAFT` — and `spec/contract.md`'s two warning blocks are resolved to
  a decided version
- A PR exists for `feat/prompt-per-selection-settings` with the completed checklist in
  its body, and `CHANGELOG.md` carries the entry under `## [Unreleased]`
- A note is queued to `~/.daftplate/outbox/` if ADR 0002 is accepted, per CLAUDE.md's
  rule that deviating from a standard needs an ADR here **and** an outbox note

## Unknowns and risks

- **Nothing in Part B has been run in Excel.** Both structural gates pass, but they
  cannot see a dialog. The three prompts, the cancel paths, the out-of-range column
  rejection and the empty-band guard are all unverified against a real workbook.
- **The empty-band guard is the riskiest change**, because it is in Core and affects
  every package, not just PkgClean. If it is wrong, a package that *should* show a band
  loses it. Section 1 of the gate tests both directions deliberately — the zero-field
  package and the all-actions-Off package.
- **The upgrade path is untested.** A v0.4.0 `_Config` still carries `clean_keycols` and
  `clean_hasheader` rows for fields that no longer exist. Expected to be ignored
  harmlessly, but section 5 of the gate exists because "expected to be" is not evidence.
- **ADR 0002 is unaccepted and blocks release, not testing.** Do not let it hold up the
  manual gate; the code under test is identical either way.
- **`.vscode/` is untracked** and unrelated to this work — left alone deliberately
  across several sessions, not an oversight.
- **Leftover SDD files could not be deleted.** `.superpowers/sdd/` still holds
  `2026-07-24-finish-excel-track1-plan/` and about 37 loose files from an older
  workspace layout; two delete commands were refused by the permission classifier. The
  directory is gitignored, so nothing is recoverable through git once removed — the
  owner should delete it manually.
- **Part A (issue #6) is not started and should not be started casually.** It needs
  `brainstorming` → `writing-plans` → an ADR, and it carries a migration trap: flipping
  the Availability default to `Off` silently zeroes every upgraded v0.3.0 deployment
  unless the format marker is used to tell "absent because old" from "absent because
  new".
