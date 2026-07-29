# Next session — paste this into a fresh chat

> Delete this file once the Core pass is underway; it is a launch pad, not a document.

## Start here

Run `/orient` before anything else, including clarifying questions. Then complete the rest
of the planning and implementation steps **without stopping until done**, using `/goal` on a
new branch.

**The shape of the work, in this order:**

1. **Plan every remaining step first.** Do all the planning up front — do not interleave
   planning with implementation. The Core pass is already planned in full (read manifest
   entry 1); the remaining waves W2–W6 are not, and each needs its own plan.
2. **Then review the plans.** Each plan gets a review before any code is written.
3. **Then implement**, step by step, each step reviewed after implementation.

**Use codex agents for most of the work.** Dispatch implementation, mechanical edits, and
research to subagents. **Do not delegate**: reviewing, validating, or deciding on a plan or
a step — those stay with you. You hold the cross-step context a subagent cannot.

**The W1 manual gate is deferred by owner decision** (2026-07-28) and must not be revived as
a per-wave step. Everything folds into one end-to-end validation on the finished workbook at
the very end. Do not write per-wave gate documents for the owner to run; write the
verification into the build and the code instead.

## Read first

Read nothing else until you have these, in order:

1. `docs/designs/2026-07-28-core-refuse-pass-plan.md` (28K) — the phase contract for the
   next step, planned in full with three owner decisions already settled. Read whole.
2. `docs/designs/2026-07-24-finish-excel-design.md` (32K) — §3 is the whole track-2 spec,
   §3.5 the wave order, §3.6 the stopping rule. Its `## 5. Handoff log` top entry is the
   do-not-revert list and the empirical findings from W1; **read that entry in full**.
3. `docs/macro-catalog-plug-and-play.md` (20K) — the authoritative build queue with a
   per-item verdict and frozen action id for W2–W6. This is what the remaining plans are
   built from.
4. `spec/contract.md` (68K) — **over the slice threshold, never read front to back.** Grep
   `## 6.` (action contract, where the refusal outcome lands), `## 7.` (the action-id
   vocabulary table the build gate parses) and `## 8.` (frozen registry). Currently v2.3.0.
5. `src/core/Core.bas` (232K) — **never read whole.** Grep `docs/code-map.md` (36K, current
   at `bafb905`) for an anchor, then read only that slice. The Core pass touches
   `Core_Done`, `RunAction`, `Core_RunChain`, `Core_SheetExists`.
6. `src/packages/PkgImport.bas` (40K) — the newest package and the pattern for W2's; it
   carries the 8 refusal sites the Core pass migrates. **Over the slice threshold — read it
   in slices**, not whole: grep `docs/code-map.md` for `PkgImport_` anchors, or grep the
   file itself for `PkgImport_Refuse ctx` to land on the eight call sites and
   `PkgImport_AppendSheet` for the engine.

## Branch

`feat/w1-ingest-clean` — 28 commits, W1 built and **open as PR #19 against `main`, not
merged**. Its manual gate is deferred, so #19 is mergeable on review alone.

**Branch the Core pass from `main`, not from here.** If #19 has merged, `main` already has
W1. If it has not, merge or close it first — do **not** stack, and do not branch from
`feat/w1-ingest-clean`. Stacking is what stranded 11 commits last time and why #18 exists.
Do not branch from `feat/sheetforge-identity`, `feat/action-catalog-first`,
`fix/core-hardening-and-checklist-gate` or `feat/finish-excel-track1`; all are merged and
stale.

## Constraints

- **Do not revert anything in the design doc's do-not-revert list.** In particular:
  `PkgImport_AppendSheet` measures its extent per column with `End(xlUp)` from the bottom
  and **not** from `UsedRange` — a prior round made exactly that "simplification" and traded
  a column-A truncation for a bug that spaced every source 15 rows apart. Also: the combine
  never closes a workbook it did not open; `ThisWorkbook` is excluded by path **or** name;
  every VBA-syntax pattern in `build.ps1` carries `(?i)`; the id gate's comparers are
  ordinal and its id checks use `-cmatch`/`-cne`; `k_clean_formats` declares no default rows.
- **`SCHEMA_VER` stays `"v2.0"`.** If a build ever reports `v3.0`, something is wrong.
- **Action ids are immutable, two-segment `<domain>.<operation>`, never camelCase, no alias
  table.** `build.ps1` enforces this by parsing the vocabulary table out of contract §7; a
  new domain must be added there first or the build fails.
- **TDD is off** (`CLAUDE.md`) — VBA has no headless runner and Excel cannot be driven on
  CI. Do not propose a test suite. Verification is `build.ps1`, `vba-syntax-check.mjs`, and
  deliberate injections for anything meant to fail.
- **Shared helpers live ONLY in Core; cross-package calls are forbidden.** The Core pass
  exists because W1 had to duplicate two helpers for want of a Core API.
- **Pure ASCII in every `.bas`**; `Option Explicit`; max 25 line continuations; no
  module-level declaration after the first procedure; no VBA reserved word as an identifier
  (`eNum` lowercases to `enum`).
- **`dist/` is gitignored** — never `git add -f` the bundle. `.vscode/` stays untracked
  deliberately.
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.

## Exit criteria

- `pwsh build/build.ps1` exits 0 reporting `Core + 3 package(s)` or more, with balanced
  `Sub` and `Function` counts
- `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` reports no diagnostics
- `grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas` prints `1`
- `spec/contract.md` records v2.4.0 or higher, with the refusal outcome in §6
- Every wave planned in `docs/macro-catalog-plug-and-play.md` has a plan document under
  `docs/designs/`, each reviewed before its implementation began
- `PkgClean_Refuse`, `PkgImport_Refuse`, `PkgImport_IsEngineSheet` and `PkgImport_IsDigits`
  no longer exist in `src/packages/`
- One PR per phase against `main`, none stacked

## Unknowns and risks

- **The used-range fix in `bafb905` has never run in Excel.** It is the fix for the only
  defect the owner's partial gate found, and no gate item has confirmed it. Items 1.5 and
  2.4 in `docs/designs/2026-07-28-w1-manual-gate.md` exist for that purpose and are
  deferred with everything else.
- **PR #19 is open and unmerged.** Every later phase depends on it landing. If it is still
  open, decide whether to merge it on review alone before branching — the gate that would
  normally block it is deferred.
- **8 track-1 gate items are also unrun**, from an earlier phase. Reskin 4 (legibility on a
  dark Excel theme) is the likeliest real finding: the palette came off a dark-background
  proof and lands on an explicitly white sheet, and nothing has checked the two together.
- **The 34-vs-36 action count is unreconciled.** §3.6 says W1–W6 assign 34 items producing
  35 actions; grepping the catalog for wave tags yields 36. Resolve it while planning W2,
  because every remaining wave's scope is drawn from that document.
- **`Core_RunChain` asks rather than aborts on a failed step**, and a user who answers
  Continue runs later steps against stale data. The Core pass changes this for *refusals*
  only. Whether a *failure* should also stop is still open and is the owner's call.
- **`Select Case LCase(<Const>)`** is used in the Core pass's planned `Core_IsEngineSheet`
  and is not exercised anywhere in the codebase today. The plan states the fallback.
- **`.vscode/` is untracked and left alone deliberately** across many sessions. It is not an
  oversight; do not commit or remove it.
- **`node scripts/code-map.mjs` stamps the previous commit** when run before committing, so
  a freshly regenerated map can look one commit stale. Regenerate, commit, then regenerate
  again only if the stamp matters.
