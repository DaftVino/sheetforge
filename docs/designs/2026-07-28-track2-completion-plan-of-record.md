# Track 2 completion — plan of record

**What this is.** The sequencing document for everything between v0.10.0 and a finished
Excel toolkit: the Core passes, waves W2–W6, and the single end-to-end validation that
replaces the per-wave manual gates. It holds the cross-phase decisions. Each phase has its
own plan document; this one says what the phases are, why they are in this order, and which
questions are already answered so no phase re-litigates them.

Written 2026-07-28, on `feat/core-refuse-pass`, branched from `origin/main` at `6b68ebc`
(the merge of PR #19 — W1 is on `main`).

---

## 1. State at the start

| Fact | Value |
|---|---|
| Released | `v0.10.0` |
| Contract | `v2.3.0` |
| Actions shipped | 15 (8 legacy/shipped + 7 from W1) |
| Packages | `PkgImportExport`, `PkgClean`, `PkgImport` |
| PR #19 (W1) | **Merged** 2026-07-28. `origin/main` carries W1. |
| W1 manual gate | **Deferred** by owner decision (`bde8400`) |
| Track-1 gate items | 8 unrun, carried forward |

**The 34-vs-36 action count is reconciled — there is no discrepancy.** `grep -c
'\*\*W[1-6]\*\*' docs/macro-catalog-plug-and-play.md` prints **35**: one legend line
(line 17, which names `**W1**` and `**W6**` while explaining the notation) plus **34**
item lines. Those 34 items declare **35** action ids, because line 199 —
*Protect / unprotect all sheets at once* — is one catalog item carrying two ids
(`workbook.protect_sheets` + `workbook.unprotect_sheets`). §3.6 of the design doc
("W1–W6 assign 34 catalog items, producing 35 new actions") is therefore exactly right,
and the "36" in the handoff came from counting bold spans rather than lines. No scope
changes. **Do not re-open this.**

---

## 2. The phases

One PR per phase, each against `main`, **none stacked**. Branch each from `main` after the
previous has merged.

| # | Phase | Plan | New actions | Version | Contract |
|---|---|---|---|---|---|
| 1 | Core pass — refusal is an outcome | `2026-07-28-core-refuse-pass-plan.md` | 0 | 0.11.0 | v2.4.0 |
| 2 | Core pass 2 — shared-helper promotion | `2026-07-28-core-helper-promotion-plan.md` | 0 | 0.12.0 | unchanged |
| 3 | W2 — Validate & Reconcile | `2026-07-28-w2-validate-plan.md` | 5 | 0.13.0 | v2.5.0 |
| 4 | W3 — Format & Present, plus export | `2026-07-28-w3-format-export-plan.md` | 8 | 0.14.0 | v2.6.0 |
| 5 | W4 — Summarize & Enrich | `2026-07-28-w4-summarize-enrich-plan.md` | 6 | 0.15.0 | v2.7.0 |
| 6 | W5 — Protect & Back Up | `2026-07-28-w5-backup-plan.md` | 4 | 0.16.0 | v2.8.0 |
| 7 | W6 — Workbook housekeeping | `2026-07-28-w6-workbook-plan.md` | 5 | 0.17.0 | v2.9.0 |
| 8 | End-to-end validation | `2026-07-28-end-to-end-validation.md` | 0 | 1.0.0 | unchanged |

28 new actions across W2–W6, plus W1's 7 = the 35 §3.6 counts.

**Ending total: 43 action ids** (15 shipped today + 28 from W2–W6). §3.6's "~48" is a
different quantity, not a larger scope: it adds **13 catalog items** marked shipped-or-
embedded to the 35 wave ids. Only 7 of those 13 are selectable actions — the other 6 are
embedded helpers with no id of their own (`Core_Coerce`, `Core_FindFile`, the speed
wrapper, the last-row finder, the error handler, the dated-filename convention). Counting
ids against ids gives 43; counting catalog items against ids gives ~48. **Neither number
is wrong and nothing is being cut** — the 8 *Later* items are unassigned in both.

One bookkeeping note: `importexport.suggestTypes` is a shipped id that appears in no
catalog entry, so the catalog alone under-counts today's actions by one.

---

## 3. Decisions made here, so no phase re-litigates them

### 3.1 Phase 2 is inserted, and why

**Decision: a second Core pass runs before W2.**

The Core refusal pass (phase 1) exists because W1 had to duplicate two helpers for want of
a Core API. Research over `src/packages/` on 2026-07-28 found that **the same thing happens
again at W2, in six more places**, and phase 1 does not touch any of them:

| What W2 needs | Where it lives today | Who else needs it |
|---|---|---|
| Prompt the user for a range | `AskRange`, `PkgClean.bas:603` — **private** | every W2/W3/W4 action that acts on a selection |
| Ask "is row 1 a header?" | `AskHasHeader`, `PkgClean.bas:626` — **private**, and its dialog title is hardcoded `"Remove Duplicates"` | W2 `validate.*`, W4 `summarize.build_pivot` |
| Parse `"1,2,4"` into column numbers | `ParseCols`, `PkgClean.bas:676` — **private** | W2 `validate.reconcile` (key columns), W4 `enrich.join` |
| Narrow a whole-column pick to the used area | `PkgClean_NarrowToUsed`, `PkgClean.bas:582` | every action taking a range |
| Validate-then-create a destination sheet | `PkgImport_DestUsable` + a two-line add/name recipe, **duplicated twice inside `PkgImport.bas` already** (`:192` and `:415`) | W2 exceptions report, W4 pivot sheets, W5 snapshot, W6 TOC |
| Resolve and test an output folder | `PkgImport_ResolveFolder` `:571`, `PkgImport_FolderExists` `:585` | W3 export, W5 backup |

**The flat namespace makes this worse than ordinary duplication.** The bundle is one
module, so `AskRange` sitting unprefixed in `PkgClean` does not merely fail to be reusable —
it *reserves the name*. `build.ps1`'s duplicate-procedure guard (`build/build.ps1:482`) fails
any build in which a second package declares its own `AskRange`. So W2 cannot copy it and
cannot re-declare it; the only two legal moves are promote it to Core or invent a differently
named twin. Copying is the move the standards forbid, and inventing a twin is the move that
produced this pass.

Doing this as its own phase rather than folding it into W2 keeps a pure refactor — no
behaviour change, no new action — in its own reviewable diff, which is the only way the
"did anything change for the user?" question has a clean answer.

### 3.2 The naming inconsistency is fixed in phase 2, not left

`PkgImport.bas` prefixes all 14 of its privates. `PkgClean.bas` prefixes 10 of 16.
`PkgImportExport.bas` prefixes **0 of 5**. In a flat namespace an unprefixed private is a
name claimed against every future package. Phase 2 resolves each one either by promoting it
to `Core_*` (generic) or by prefixing it (package-specific). After phase 2, every private in
every package file carries its package prefix, and the rule is stated in `CLAUDE.md` rule 4
terms so W3–W6 inherit it.

### 3.3 `PkgImportExport`'s missing refusal flag is fixed in phase 1

Research found that `PkgImportExport`'s refusal paths (e.g. `PkgImportExport_Clear`,
`PkgImportExport.bas:96`) show a `MsgBox` and `Exit Sub` **without ever setting
`ctx("failed")`**. Inside a chain that is reported as a green tick beside a step that did
nothing — the precise defect phase 1 exists to eliminate, present in a third package.

**Decision: fold it into phase 1 as its own task, and say so.** This is a deliberate,
stated extension of the scope the owner settled on 2026-07-28. That settlement's decision 3
("`Core_Refuse` + `Core_IsEngineSheet` + the `#0` fix, not the remaining W1 minors") was
about not absorbing *W1 minors*; this is not a W1 minor, it is a fresh instance of the
defect the pass is named after, and the edit is the same shape as phase 1's Task 4. Leaving
it would mean the pass cures the disease in two packages and ships it in the third.
Flagged rather than done silently, per `CLAUDE.md`.

### 3.3a Core keeps growing after phase 2, and the waves are no longer interchangeable

Phase 2 promotes what is **already duplicated**. Drafting W2–W6 surfaced three more helpers
that phase 2 could not have known about, because they have no second consumer until a
specific wave exists. Each is added by the wave that creates its second consumer, per phase
2's own "two consumers, or already duplicated" rule:

| Helper | Added by | Why not phase 2 |
|---|---|---|
| `Core_ResolveColumn` | **phase 2 after all** | W2 and W4 both needed `PkgClean_ResolveFormatCol`'s letter-or-header resolution, which made it three consumers — so it moved into phase 2 with `PkgClean` migrated onto it. |
| `Core_FreezeValues` | **W3** | The formula-to-values flatten lives inside `PkgImportExport_Export` today. `export.freeze_values` is its second consumer and does not exist until W3. |
| `Core_SafeFileName` | **W3** | `export.pdf` and `export.per_sheet` are two consumers on their own. |
| `Core_AlreadyOpen` | **W5** | Phase 2 explicitly declined to promote it on the grounds that W5's need was speculative. Drafting W5 showed the need is real (`backup.archive_output` moves a file that must not be open), which is exactly the condition phase 2 named for promoting it later. |

**Consequence: W3 must land before W5.** `backup.archive_output` and `backup.snapshot` call
`Core_SafeFileName`, which W3 introduces. The wave order in design §3.5 already puts W3
third and W5 fifth, so nothing has to move — but the dependency is now real rather than
incidental, and reordering the remaining waves is no longer free.

**`Core_ResolveColumn` returns a `Long` index, not a `Range`.** Three plans independently
invented three signatures for it; this is the one. Callers wanting the column itself write
`rng.Columns(idx)`; `0` means unresolved. It is the single promotion in phase 2 whose call
site changes shape rather than only its name.

### 3.4 One package per domain

Seven new modules, one per active domain: `PkgValidate`, `PkgFormat`, `PkgExport`,
`PkgSummarize`, `PkgEnrich`, `PkgBackup`, `PkgWorkbook`.

Domain and module stay decoupled in the *contract* (§7: the domain is a user-facing
vocabulary, not the module) — that decoupling exists so a future split is not a breaking
change, not to require a mismatch today. Aligning them now costs nothing and keeps
§3.3's "an engine is capped at one coherent workflow" honest: `format.*` and `export.*` do
not share an engine, so they do not share a module even though they ship in the same wave.

**Every new module needs registering in two places** (research finding,
`build/build.ps1:82` and `src/core/Core.bas:952`):

1. `src/packages/Pkg<Name>.bas` — `build.ps1` picks it up by glob, no registration needed.
2. `CoreCandidates()` at `src/core/Core.bas:952` — a hardcoded `c.Add "<Name>"` list Core
   probes via `Application.Run`. **Without this the package builds, bundles, passes every
   gate, and never runs.** It is the single easiest way to ship a silently dead wave, so
   every wave plan carries it as an explicit step with its own verification.

### 3.5 No per-wave manual gate documents

**Decision: no phase writes a gate document for the owner to work through.**

The owner deferred the W1 gate on 2026-07-28 (`bde8400`) with the reasoning stated in the
banner: *"the answers are available statically, so defects that a careful read can find
should be found that way rather than spent as owner time."* Writing W2–W6 gates would spend
exactly the time that decision reclaimed.

This **supersedes Task 5 steps 1–2 of the phase-1 plan**, which were written before the
deferral and add items to the W1 gate. What replaces them, per phase:

1. **Verification into the build.** Anything a gate item could assert statically becomes a
   `build.ps1` guard or a `GateAudit.bas` check. A guard is run on every build forever; a
   gate item is run once, by hand, if someone remembers.
2. **Verification into the code.** Anything a gate item could assert at run time becomes a
   refusal with a plain-language reason, so the failure states itself to the user instead of
   waiting to be discovered.
3. **A deliberate injection** for anything meant to fail: break it on purpose, confirm the
   build or the guard catches it, revert. Recorded in the PR body.
4. **An adversarial static read** of the built bundle, by a subagent forbidden from reading
   `src/`, for anything that is genuinely run-time-only. This keeps the blindness that makes
   the technique work while spending an agent instead of the owner.

**Phase 8 writes one end-to-end validation document, once, over the finished workbook.** It
is the only thing the owner runs, and it accumulates what each phase would have asked for.
Each phase appends its run-time-only claims to a running list in
`docs/designs/2026-07-28-deferred-validation-log.md` as it goes, so phase 8 assembles rather
than reconstructs.

**The residual risk, stated plainly rather than argued away.** The four substitutes above do
not close the run-time hole; they **postpone** it. None of them executes Excel. A build guard
proves a guard fires, not that a macro works. A refusal the user can read proves the message
exists, not that the precondition is detected correctly. A static read proves the code says
what it appears to say. So through phases 1–7 **every wave builds on behaviour that has been
reviewed but never observed**, and a defect in phase 1's chain halt would sit under five
waves before phase 8 finds it — with the `bafb905` used-range fix already in that position
today, unverified since the partial gate.

That is a deliberate, owner-made trade of early detection for owner time, and it is the right
call *given* that Excel cannot be driven on CI. It is recorded here so it stays a decision
rather than becoming an assumption. **No phase may describe its run-time behaviour as
verified.** Reviewed, yes. Verified, at phase 8.

### 3.6 What phase 8 already owes the owner

Carried forward, before any new phase adds to it:

- **The `bafb905` used-range fix has never run in Excel.** It is the fix for the only defect
  the partial W1 gate found. W1 gate items 1.5 and 2.4 exist to confirm it.
- **W1 gate item 0.3 is inconclusive** — `Macros 1` did not exist afterwards. Redo first.
- **8 track-1 gate items** are unrun. **Reskin 4 (legibility on a dark Excel theme) is the
  likeliest real finding**: the palette came off a dark-background proof and lands on an
  explicitly white sheet, and nothing has checked the two together.
- **`Select Case LCase(<Const>)`** (phase 1, `Core_IsEngineSheet`) is believed legal VBA but
  is exercised nowhere in this codebase today. Marked as an assumption; the plan states the
  `If/ElseIf` fallback.

---

## 4. Constraints every phase inherits

Copied verbatim so no phase plan has to be read alongside this one.

- **`SCHEMA_VER` stays `"v2.0"`.** If a build ever reports `v3.0`, something is wrong.
- **Action ids are immutable**, two-segment `<domain>.<operation>`, lowercase ASCII and
  underscores, each segment starting with a letter. **Never camelCase. No alias table.**
  `build.ps1` parses the vocabulary table out of contract §7 and fails on an unknown domain,
  so **a new domain must be added to that table in the same commit or earlier**.
- **Shared helpers live ONLY in Core. Cross-package calls are forbidden.**
- **Pure ASCII in every `.bas`**; `Option Explicit`; max 25 line continuations per statement;
  **no module-level declaration after the first procedure** (in the *assembled bundle* — so a
  package's `Const`/`Dim` must sit above its own first procedure); no VBA reserved word as an
  identifier (`eNum` lowercases to `enum`).
- **Every action entry point must be `Public Sub <Name>(ctx As Object)`** — one `Object`
  parameter, nothing else. `build.ps1:472` fails the build otherwise, because `Core_Run`
  dispatches through `Application.Run`.
- **TDD is off** (`CLAUDE.md`). VBA has no headless runner and Excel cannot be driven on CI.
  **Do not propose a test suite.** Verification is `pwsh build/build.ps1`, `node
  scripts/vba-syntax-check.mjs dist/MacrosBundle.bas`, and deliberate injections.
- **`dist/` is gitignored** — never `git add -f` the bundle. **`.vscode/` stays untracked**,
  deliberately, across sessions. Do not commit it and do not remove it.
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.

### 4.1 The do-not-revert list

From the design doc's handoff log. Reverting any of these re-introduces a defect that was
found the expensive way.

- **`PkgImport_AppendSheet` measures its extent per column with `End(xlUp)` from the bottom,
  not from `UsedRange`.** A prior round made exactly that "simplification" and traded a
  column-A truncation for a bug that spaced every source 15 rows apart.
- **The combine never closes a workbook it did not open**, and `ThisWorkbook` is excluded by
  path **or** name.
- **Every VBA-syntax pattern in `build.ps1` carries `(?i)`.**
- **The id gate's comparers are ordinal**; its id checks use `-cmatch` / `-cne`.
- **`k_clean_formats` declares no default rows.**

---

## 5. Open, and owned by the owner

**`Core_RunChain` asks rather than aborts on a failed step.** Phase 1 changes this for
*refusals* only, by settled decision: a refusal stops a chain, an error still asks. Whether a
**failure** should also stop is still open and is the owner's call. It is not assumed either
way by any phase here, and no plan depends on the answer.

---

## 6. Order of work

Planning is finished before implementation starts — all eight plans written and reviewed
before the first line of phase-1 code. That is deliberate: the Core API phase 2 establishes
is the vocabulary every wave plan is written in, so a wave planned after phase 2 shipped
would be planned against a moving target, and a wave planned before it would invent its own
helpers again.
