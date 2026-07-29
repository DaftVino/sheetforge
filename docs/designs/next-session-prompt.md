# Next session — paste this into a fresh chat

> Delete this file once the Macros Home phase is underway; it is a launch pad, not a document.

## Start here

Run `/orient` before anything else, including clarifying questions. Then **land the stranded
track-2 stack onto `main`**, and after that plan and implement three things the owner
identified as missing on 2026-07-29: **`Macros Home`**, the **8 catalog items marked
*Later***, and a **Finalize reminder**. Plan every step first, review each plan before writing
any code, then implement step by step with each step reviewed — the same shape as track 2.

## Read first

Read nothing else until you have these, in order:

1. `docs/designs/2026-07-28-track2-completion-plan-of-record.md` (20K) — the phase contract
   for *how* this is done: §3.5 the no-per-wave-gates policy, §4 the constraints every phase
   inherits, §4.1 the do-not-revert list. Read whole. There is no plan for the new work yet;
   writing one in this shape is the first task.
2. `docs/designs/2026-07-24-finish-excel-design.md` (40K) — **over the slice threshold, do not
   read front to back.** Read **§2.2** (lines 49–73), which specifies `Macros Home` and is the
   only spec for it; and the **top entry of `## 5. Handoff log`** (from line 302), which is the
   do-not-revert list and the account of what track 2 did and did not deliver.
3. `docs/macro-catalog-plug-and-play.md` (20K) — the authoritative build queue. The **8 items
   marked `**Later**`** (lines 48, 79, 139, 161, 183, 186, 208, 214) are this phase's new
   actions. Each needs a frozen `<domain>.<operation>` id invented and added here.
4. `spec/contract.md` (108K) — **far over the threshold; read it only in slices, from grep
   anchors, never front to back.** Grep `## 6.` for the availability model and the withdrawn-index clause, `## 7.` for the
   vocabulary table the build gate parses and the reserved sheet names, `## 8.` for the frozen
   registry each package extends. Currently v2.9.0.
5. `src/core/Core.bas` (264K) — **never read whole.** Grep `docs/code-map.md` (48K) for an
   anchor and read only that slice. This phase touches `Core_BuildControlPanel` (:3815),
   `DeleteHomePanel` (:3943), `RenderPage` (:4212) and `AvailFor` (:2906).
6. `docs/designs/2026-07-28-end-to-end-validation.md` (20K) — the 41-item checklist this phase
   appends to, and the pattern for writing an item the owner can actually run.

## Branch

`feat/w6-workbook` — the complete track-2 tree: 43 actions, ten packages, version 0.17.0,
contract v2.9.0. **This branch, not `main`.**

**`main` is stale at v0.11.0 with three packages.** PRs #20–#27 all read MERGED, but each was
based on the previous feature branch and only #20 targeted `main`, so phases 2–8 merged into
each other and stopped. Nothing is lost; the work is here and pushed.

**Land this branch on `main` first — one PR, `feat/w6-workbook` into `main` — and do not cut
any new branch until it has merged.** Branching the new phase from `main` before then loses 28
actions. Do not branch from `docs/end-to-end-validation` (its content is already in this
branch), and do not branch from `feat/core-refuse-pass`, `feat/core-helper-promotion`,
`feat/w2-validate`, `feat/w3-format-export`, `feat/w4-summarize-enrich` or `feat/w5-backup` —
all are stale intermediate points in the same stack.

## Constraints

- **`PkgFixture` is dropped.** Design §2.5 specifies it; the owner said on 2026-07-29 they do
  not want it. Do not build it and do not re-propose it.
- **`Macros Home` must explain what each macro is for, with an example** — that is the owner's
  requirement, not just an index of links. Design §2.2 describes a card per category with a
  count and a link; this phase goes further, and §2.2's version alone does not satisfy it.
  The contract currently records the index as **withdrawn** (§6) while reserving the name
  (§7). Restoring it is a contract change, so §6's withdrawal text must be rewritten rather
  than contradicted.
- **The Finalize reminder is not specified anywhere.** The owner's words: if any action is
  turned on, tell the user they need to click Finalize and then Setup on the Setup Menu to
  complete the options for any buttons that require them. Design it; do not assume it exists.
- **Action ids are immutable, two-segment `<domain>.<operation>`, lowercase, never camelCase,
  no alias table.** `build.ps1` parses the vocabulary table out of contract §7 and fails on an
  unknown domain, so **a new domain must be added to that table in the same commit or
  earlier**. The 8 *Later* items will need ids invented — they have none yet.
- **`SCHEMA_VER` stays `"v2.0"`.** If a build reports `v3.0`, something is wrong.
- **Shared helpers live ONLY in Core; cross-package calls are forbidden.** Every package
  private must be prefixed `Pkg<Name>_` — `build.ps1` fails the build otherwise.
- **A new package must be registered in TWO places**: the file in `src/packages/`, and
  `c.Add "<Name>"` in `CoreCandidates()` (`src/core/Core.bas`, grep for it). Without the
  second it builds, passes every gate, and never runs.
- **Pure ASCII in every `.bas`**; `Option Explicit`; max 25 line continuations; no module-level
  declaration after the first procedure; **no VBA reserved word as an identifier** — `eNum`
  lower-cases to `enum` and the build rejects it, which it did once already this track.
- **VBA `And`/`Or` do not short-circuit**, and `On Error GoTo 0` inside a procedure with a
  `Fail:` handler disables that handler — both caused review findings across four wave plans.
- **TDD is off.** No headless VBA runner, Excel cannot run on CI. Do not propose a test suite.
  Verification is `pwsh build/build.ps1`, `node scripts/vba-syntax-check.mjs`, and deliberate
  injections for anything meant to fail.
- **No per-wave manual gate documents.** Run-time-only claims append to
  `docs/designs/2026-07-28-end-to-end-validation.md`. Dry-run it blind before it reaches the
  owner — that pass found five buttons the last checklist named that do not exist.
- **`dist/` is gitignored** — never `git add -f` the bundle. **`.vscode/` stays untracked**,
  deliberately.
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.

## Exit criteria

- `pwsh build/build.ps1` exits 0, reporting `Core + 10 package(s)` or more with balanced
  `Sub` and `Function` counts
- `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` reports no diagnostics
- `git show origin/main:VERSION` prints 0.18.0 or higher, proving the track-2 tree landed
- `grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas` prints 1
- `grep -ohE 'Core_AddAction reg, "[a-zA-Z_.0-9]+"' src/packages/*.bas | wc -l` prints 51
- `spec/contract.md` records v2.10.0 or higher, with §6's withdrawn-index text rewritten and
  the 8 new ids in §8
- `docs/designs/2026-07-28-end-to-end-validation.md` contains new items for `Macros Home`, the
  Finalize reminder, and each new action, and every button caption it names exists in
  `dist/MacrosBundle.bas`

## Unknowns and risks

- **`main` is 28 actions behind and every PR says MERGED.** The single most likely mistake next
  session is trusting the PR list, branching from `main`, and rebuilding work that already
  exists. Verify with `git ls-tree --name-only origin/main src/packages/` — three files means
  it has not landed.
- **Nothing in track 2 has ever run in Excel.** 41 unrun checklist items, the oldest being the
  `bafb905` used-range fix, which is now eight phases deep and still unconfirmed. A defect
  found there may reopen a wave rather than only adding to this phase.
- **Every action arrives switched `Off`** (contract §6, v2.1.0; `Core.bas:2906` returns
  `AVAIL_OFF` when `_Config` is absent). A fresh import shows 43 catalog rows and no working
  buttons. This is intended, and it is also exactly what makes the Finalize reminder
  necessary — but confirm with the owner whether the *default* should also change, because
  that is a separate decision from reminding them.
- **The 8 *Later* items have no ids, no wave, and no design.** Two may not be buildable as
  specified: *Email a sheet as an attachment* needs a defined behaviour when Outlook is absent
  or blocked by policy, and the catalog says so; *Add a "Back to TOC" link on every sheet*
  writes into every user sheet, which no existing action does. Decide both before planning
  them, and expect the owner to want a say.
- **Restoring `Macros Home` reverses a decision already recorded in the contract.** §6 says the
  reference implementation built an index and withdrew it because "at the scale a real
  deployment runs — a handful of actions and one or two chains — an index is clutter." At 51
  actions that reasoning no longer holds, and saying so in §10 is part of the work.
- **Layout at 51 actions is unverified.** Auto-layout packs three columns of six per page, so
  51 actions plus chains is roughly four pages. `PkgFixture` was the design's mechanism for
  proving this and has been dropped, so the only remaining proof is the owner looking at it.
- **Nothing is uncommitted.** `git status --porcelain` shows only the untracked `.vscode/`,
  which is deliberate and long-standing.
