# Spec — Part A, progressive disclosure on the Setup surface

**Date:** 2026-07-27
**Branch:** `feat/progressive-disclosure-part-a` (stacked on `feat/prompt-per-selection-settings`)
**Implements:** issue #6, Part A of `2026-07-26-setup-surface-progressive-disclosure-design.md`
**Panel:** `docs/panels/2026-07-27-part-a-availability-default.md` — 6 resolutions, all 3-0
**Groundwork:** `docs/designs/2026-07-27-part-a-groundwork.md`

> **STATUS: NOT APPROVED.** Written while the owner was asleep, from a panel that
> ruled on *mechanism* only. The panel did not rule on whether Part A ships, and
> neither did I. **Section 8 lists what the owner must decide before any code is
> written.** No VBA has been changed on this branch.

---

## 1. What this changes, in one paragraph

Today a fresh install draws a button for every action a package declares, and the Setup
sheet shows every field of every package. Part A inverts the default: actions arrive
`Off`, the action catalog is the first thing the user meets, and a package's fields
surface as its actions are turned on. The hard part is not the inversion — it is that
"no stored answer" currently means four different things in four places, and one of them
is a live v0.3.0 deployment that must keep every button it already has.

## 2. The four sites, and what each becomes

| # | Site | Today | After |
|---|---|---|---|
| 1 | `ParseAvail` (`Core.bas:1928`) | blank cell → `AVAIL_BUTTON` | blank cell → `""` (unrecognised) |
| 2 | `AvailFor` (`Core.bas:2021`) | id absent from map → `AVAIL_BUTTON` | id absent → `AVAIL_OFF`, **unless the map is flagged legacy** |
| 3 | `ReadAvailMap` (`Core.bas:2004`) | marker `absent` → empty map | marker `absent` → empty map **flagged legacy** |
| 4 | `AvailLabel` (`Core.bas:1991`) | `Case Else` → `"Button"` | `Case Else` → `"Off"` |

Site 1 is smaller than the panel assumed. `ParseAvail` already returns `""` for an
unrecognised value (`Core.bas:1937`), and `ComposeReport` already blocks Finalize on it
with the message *"UNRECOGNISED Availability (use Button, Chains only, or Off)"*
(`Core.bas:2562-2580`), incrementing `structural`. **The hard-block the panel demanded
already exists.** Blank simply short-circuits at line 1928 before reaching it. So
resolution 3 is a one-line change plus a comment, not new validation machinery.

Site 2 is the whole difficulty and is addressed in §3.

## 3. Provenance: the flag, not the constant

`ReadAvailMap` must stop discarding what `MarkerVerdict` already knows. The map returned
carries a flag saying whether it came from a legacy (`absent`-marker) block, and
`AvailFor` consults the flag instead of a hardcoded constant:

- **legacy map** → an absent id means **Button**. Preserves a working v0.3.0 deployment
  exactly as it is today.
- **any other map** (fresh, or current-and-marked) → an absent id means **Off**. This is
  the new default, and it covers both a fresh install and a newly-registered action id.

A `Scripting.Dictionary` cannot carry a flag on itself, so this needs either a
module-level companion or a reserved key.

**Use a reserved key, and reserve it in the contract.** Contract §7 states no character
rule for action ids — it says only `pkg.verb`, unique — so it would be **wrong to assert
that a sentinel like `*legacy` cannot collide**. What §7 *does* establish is a convention
of reserved namespaces (`chain.*` is reserved for engine chains). Part A follows that
convention rather than relying on an unstated character rule:

> **Reserved (v2.1.0):** the action-id namespace `core.*` is reserved for engine-internal
> map sentinels; no package action id may use it.

The flag key is then `core.legacy`. `AvailFor` must exclude any reserved key from
iteration; `ComposeReport`'s loop iterates the *registry*, not the map, so it is
unaffected either way.

**This makes Part A a contract change (§7 plus the §3 presentation restatement), so it
needs a minor contract bump and an ADR.** That is on the owner's reserved list.

## 4. The legacy migration

Per panel resolutions 4, 5 and 6, all 3-0.

**Where.** `Core_FinalizeSetup` only. Never `Core_BuildSetup` — BuildSetup deletes and
re-seeds the Setup sheet from `_Config`, so writing there would mutate the only persisted
state before the user has confirmed anything.

**Order.** Stage the complete v2.0 block, validate it, and write the format marker
**last**. An interrupted write therefore leaves a block with the `Availability` header but
no marker, which `MarkerVerdict` already classifies `"tampered"` (`Core.bas:2235-2239`)
and which `ComposeReport` already refuses (`Core.bas:2557-2560`). The failure mode is a
refusal with an explanation, not a silent half-migration. **This reuses machinery rather
than inventing it, and is the single best idea the panel produced.**

**Evidence.** The legacy block is preserved as a dated backup under its own column-A key,
placed so no reader consumes it. `ReadAvailMap` bounds a block by scanning column A for
the next non-blank key, so the backup must sit under a distinct key and must not be
between `LAYOUT_KEY` and its `NextKeyRow`. Without this, a wrong conversion is
indistinguishable from user intent forever, and there is no undo.

**Notice.** One dialog at migration: existing actions preserved, new actions arrive Off.
The dialog is *not* the evidence — the retained block is.

### 4.1 A hazard the panel never saw

The activate path is **delete-then-copy** (`Core.bas:2769-2775`): the old `_Config` is
deleted, then the Setup sheet is copied and renamed. Between those two statements the
workbook has **no `_Config` at all**. A crash there loses the entire configuration.

This predates Part A and is not introduced by it, so **fixing it is out of scope here**.
But it bounds what "staged and atomic" can honestly mean: the staging in §4 is
marker-last ordering within a block, not a transaction over the sheet. Recorded so the
next person does not read "atomic" and believe more than is true. Worth its own issue.

## 5. Field disclosure

`usedBy` is declared by **no shipping package** — every `PkgImportExport` field is
package-wide, and `PkgClean` now has no fields at all. So the relegation renderer is
exercised only by `PkgFixture`, and Part A's user-visible value rests entirely on wiring
`usedBy` onto `PkgImportExport`'s nine fields.

Fields are **relegated, never omitted** — the standing rule, unchanged. Rebuilding Setup
re-seeds from the finalized `_Config`, so omitting a field would cost the user
toggle → Finalize → rebuild → fill → Finalize and would discard anything typed but not
finalized.

**Inherited gate item.** Part B's checklist item 1.4 could not be tested because nothing
shipping declares `usedBy`. Once this section lands, that test becomes reachable and
**Part A's manual gate must include it**: set every `importexport.*` action Off and
confirm the band disappears *and* the off band appears.

## 6. Explicitly out of scope

- **A designed first-run empty state.** The aesthete argued all-Off makes first run look
  empty and answered itself with "a polished catalog, clear action descriptions, sensible
  grouping, and an explicit empty-state prompt". That is a second design project. Part A
  ships a legible catalog with everything Off. Discarded at panel, recorded there.
- **Fixing the delete-then-copy activate.** §4.1. Separate issue.
- **Any test suite.** TDD is formally off (CLAUDE.md); Excel has no headless runner.

## 7. Verification

Structural, then manual — the only gates this repo has:

1. `pwsh build/build.ps1` → `Core + 2 package(s)`, balanced Sub/Function counts.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` → no diagnostics.
3. A manual gate written to the same standard as Part B's (issue #3), covering at
   minimum: a fresh install shows everything Off; a **v0.3.0 legacy workbook keeps every
   button**; the migration dialog appears once and not again; the retained legacy block
   is never rendered on Setup and never read back as config; an interrupted migration is
   refused rather than half-applied; a blank Availability cell blocks Finalize with the
   existing message; and Part B's inherited item 1.4.

**The legacy case cannot be tested without a legacy workbook.** If no v0.3.0-era workbook
still exists, §4 is untestable dead code and should not be built — see §8.

## 8. What the owner must decide

**None of these are settled, and the first one can cancel most of this spec.**

1. **Does a v0.3.0-era (marker-absent) workbook still exist anywhere?** If not, the entire
   legacy migration in §4 is dead code guarding nobody.

   **The repo appears to answer this already.** ADR 0002, dated today, records the owner's
   own finding: *"there is no installed base. The bundle exists only in test workbooks."*
   If that holds — and it is a statement about deployments, not versions — then §4
   protects a corpus of zero real users plus at most one test fixture, and the owner's own
   reasoning from that ADR applies verbatim: *simpler wins while it is free.* The
   consistent decision would then be to **refuse a legacy config outright, exactly as
   v0.4.0 configs are now refused**, or to migrate it plainly with no retained backup.

   That reverses three unanimous panel resolutions on a fact the panel never saw, so it is
   the owner's call, not mine. But **§4 should not be built until this is confirmed.**
   §2, §3 and the blank-cell rule stand either way.
2. **Does Part A ship at all before the Apps Script port?** It changes the contract's
   presentation rules (§3 restatement, per the design doc) and every deployment's default.
3. **Is the retained legacy backup permanent, or retired after N clean migrations?** The
   panel's revisit table treats clean runs as evidence for retiring it.
4. **Does `AvailLabel`'s `Case Else` really move?** It is display-only, but it is what
   seeds a cell the user then sees. Moving it and site 2 together is consistent; moving
   only site 2 would render "Button" over a value that behaves as Off.

## Handoff log

- **2026-07-27** — Spec written from the 3-0 panel. Branch created, no code changed.
  Two of the panel's revisit triggers were discharged immediately by reading Core: the
  Finalize hard-block already exists (§2), and no atomic `_Config` helper exists — the
  activate is delete-then-copy (§4.1). Panel record:
  `docs/panels/2026-07-27-part-a-availability-default.md`.
  **Blocked on §8, question 1.**
