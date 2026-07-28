# Next session — paste this into a fresh chat

> Delete this file once W1 is underway; it is a launch pad, not a document.

## Start here

Run `/orient` before anything else, including clarifying questions. Then build
**W1 — the ingest→clean wave, 7 actions** — in two phases, each reviewed by codex
before it is acted on.

**Phase 1 — plan.** Invoke `superpowers:writing-plans` and write the W1
implementation plan. Then **`/codex` review the plan** before writing any code;
treat a `[P1]` as blocking and resolve it in the plan, not in the implementation.

**Phase 2 — implement.** Invoke `superpowers:subagent-driven-development` and
execute the plan task by task. Dispatch mechanical work to codex — terra or sonnet
for per-module edits, sol for the contract and `build.ps1` work. **Opus writes the
contract §7/§8 edits and the `build.ps1` id checks itself**: they are the two pieces
every later wave depends on. Then **`/codex` review the finished diff**, and fix
`[P1]`s before opening the PR.

**What is frozen, and what the plan is actually for.** W1's *scope* is settled and
must not be re-opened: the seven action ids, their module assignments and the wave's
membership are frozen in §3.4 and were approved by the owner on 2026-07-28. Do not
brainstorm the wave, do not add or drop an action, do not rename an id. The plan
covers *how*: the config fields each action exposes and their defaults, the shape of
`PkgImport_CombineEngine`, the cancel and error-restore paths, and the gate script.
Two open questions in `## Unknowns and risks` are real plan input — settle them
there, with the owner, rather than mid-implementation.

Build in this order, because the first two items are what make the other five safe
to write:

1. **Contract v2.3.0** — add the nine-domain vocabulary table to §7 (`import` ·
   `clean` · `validate` · `enrich` · `format` · `summarize` · `export` ·
   `workbook` · `backup`; reserved `chain.*` and `core.*`; grandfathered
   `importexport.*`) and the seven new ids to §8. **§7 has no vocabulary table
   today** — the design assumes one exists and it does not.
2. **`build.ps1` id enforcement** — extract literal `Core_AddAction` ids, fail on
   duplicates and on any domain absent from that table.
3. **`src/packages/PkgImport.bas`** (new) — `import.combine_worksheets`,
   `import.combine_workbooks`, `import.delimited_file`. The two combines share one
   private `PkgImport_CombineEngine`.
4. **`src/packages/PkgClean.bas`** (+4) — `clean.remove_blanks`,
   `clean.unmerge_fill`, `clean.standardize_formats`, `clean.text_to_numbers`. The
   last calls the existing `Core_Coerce`; do not reimplement coercion.
5. **Manual gate + `tools/GateAudit.bas` extension** — the manifest-shaped checks
   (every action has a manifest entry, an entry point and a valid domain) automate
   the way the colour checks did. Only click-every-button stays manual.

## Read first

Read nothing else until you have these, in order:

1. `docs/designs/2026-07-24-finish-excel-design.md` (22.1K) — the phase contract.
   §3 is the whole track-2 spec and §3.4 is W1's frozen action table. Its
   `## 5. Handoff log` carries the do-not-revert list and the empirical findings
   that cost round trips last session — read the 2026-07-28 entry in full.
2. `docs/macro-catalog-plug-and-play.md` (16.3K) — the authoritative build queue,
   with a per-item verdict. Confirms which library snippets are W1 and which are
   Deferred, Declined or unassigned.
3. `docs/macro_code_library.md` (30.8K) — the owner's reference implementation, 49
   snippets; §1 and §2 cover W1. **These are test scaffolding, not package code**:
   `ActiveSheet`/`Selection`, hardcoded paths marked `' <-- set`, no speed wrapper,
   no error handler, none runtime-tested. Converting one is most of the work.
4. `spec/contract.md` (54.1K) — **over the slice threshold; never read front to
   back.** Grep for `## 7.` (naming rules, where the vocabulary table goes) and
   `## 8.` (frozen registry, where the seven ids go). Currently v2.2.0.
5. `build/build.ps1` (16.6K) — gains the id checks. Read its existing
   duplicate-procedure and block-balance checks first and match their shape and
   their failure output.
6. `src/packages/PkgClean.bas` (11.0K) — the package pattern to copy for
   `PkgImport.bas`, and the file W1 extends by four actions.

## Branch

`main` — at v0.9.0 / contract v2.2.0, clean, with all of track 1 merged. Create a
new branch for W1 off `main`: `git checkout -b feat/w1-ingest-clean`. Do **not**
branch from
`feat/sheetforge-identity`, `feat/action-catalog-first` or
`fix/core-hardening-and-checklist-gate`; all three are merged and stale.

Open **one PR against `main`**. Do not stack PRs this time: last session stacked
three, merged them back-to-back, beat GitHub's retarget, and stranded 11 commits on
base branches that had already merged. #18 exists only to undo that.

## Constraints

- **`SCHEMA_VER` stays `"v2.0"`.** W1 adds actions; it reinterprets no stored
  value. If a build ever reports `v3.0`, something is wrong.
- **Action ids are immutable, two-segment `<domain>.<operation>`, never camelCase,
  and there is no alias table** — by design, since an alias means two permanent
  names for one action and would have to be ported to Apps Script. A wrong id is
  unfixable, which is why enforcement lands before the actions.
- **The domain is a user-facing vocabulary, NOT the module.** `import.*` actions
  live in `PkgImport` while `importexport.import` stays where it is, and both merge
  onto one Home card via `category:=`.
- **`Core_AddAction` already takes `category:=`** — shipped in track 1. Do not
  re-add it.
- **Do not revert anything in the handoff log's do-not-revert list.** In
  particular: `ClearCatalogBlock` clears row by row and skips merged rows (a
  blanket `Range.ClearContents` over D:E partially clears a `TintRow` B:E merge and
  raises `1004`); `CfgSheet` adopts an orphaned `_Config_staged`; and the identity
  palette is five `Private Function`s placed after every module-level declaration.
- **Never read `src/core/Core.bas` whole** (229.3K) and never read an `.xlsm`. Grep
  `docs/code-map.md` (29.3K, current at `8886fa8`) for a symbol, then read only
  that slice.
- **TDD is off** (`CLAUDE.md`) — there is no VBA runner and you must not propose a
  test suite. The gates are `build.ps1`, `scripts/vba-syntax-check.mjs`, and the
  owner running the imported bundle.
- **Every new manual checklist passes a blind dry-run before it reaches the
  owner** — a subagent holding the checklist and the built bundle, forbidden from
  reading `src/`. Know its limit, learned the hard way: it reads VBA, it does not
  execute it, so it catches authoring defects and no runtime defects at all.
- Every action wraps its work in the speed pattern with guaranteed restore in the
  error handler; shared helpers live only in Core; no attribution footers in any
  commit message or PR body.

## Exit criteria

- `pwsh build/build.ps1` exits 0 reporting `Core + 3 package(s)` with balanced Sub
  and Function counts, and `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`
  reports no diagnostics
- `pwsh build/build.ps1` fails with a named error when a duplicate
  `Core_AddAction` id or an off-vocabulary domain is deliberately injected, and
  passes again once reverted
- A W1 implementation plan exists under `docs/designs/`, and **both** codex reviews
  are recorded in the PR body — one of the plan before implementation, one of the
  finished diff — with no unresolved `[P1]` from either
- `grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas` prints `1`
- `spec/contract.md` records **v2.3.0**, with the nine-domain vocabulary table in
  §7 and all seven new ids in §8
- All seven W1 ids appear in `dist/MacrosBundle.bas` exactly as frozen in §3.4:
  `import.combine_worksheets`, `import.combine_workbooks`, `import.delimited_file`,
  `clean.remove_blanks`, `clean.unmerge_fill`, `clean.standardize_formats`,
  `clean.text_to_numbers`
- A W1 manual gate exists under `docs/designs/`, has passed a blind dry-run, and
  every item is either ticked against a real run or explicitly recorded as untested

## Unknowns and risks

- **8 of 15 track-1 gate items are still unrun, and all need the owner:** hardening
  1.3, 2.2, 3.1, 4.1, 4.2; reorder 2 (clipping half) and 3; reskin 4 (dark Excel
  theme). Both load-bearing items — #8 crash recovery and #12 fail-closed — passed.
  Reskin 4 is the likeliest real finding: the palette came off a dark-background
  proof and lands on an explicitly white sheet, and nothing has checked the two
  together. `tools/GateAudit.bas` re-runs six mechanical checks in one click.
- **`.vscode/` is untracked and left alone deliberately** across many sessions. It
  is not an oversight; do not commit or remove it.
- **`.superpowers/sdd/` still needs deleting by hand.** The permission classifier
  refused twice. Nothing depends on it.
- **The 34-vs-36 action count is unreconciled.** §3.6 says W1–W6 assign 34 items
  producing 35 actions; grepping the catalog for wave tags yields 36. Worth
  resolving before W2; not urgent for W1, whose seven are unambiguous.
- **Three W1 items have no owner-facing config story yet.**
  `import.delimited_file`, `clean.standardize_formats` and both combines carry
  hardcoded paths, sheet names or format strings in the library. Which become
  manifest fields, and with what defaults, is a real design decision this wave has
  to make and the design doc does not settle it.
- **`clean.remove_blanks` is destructive with no preview**, and the catalog defers
  sibling items for exactly that reason. It is assigned to W1 anyway. Confirm with
  the owner whether it needs a confirmation prompt before shipping, or whether
  row-deletion bounded to the used range is considered safe enough.
- **No W1 code exists.** The wave is approved and ordered but unstarted, so nothing
  in it has been compiled or run in Excel.
