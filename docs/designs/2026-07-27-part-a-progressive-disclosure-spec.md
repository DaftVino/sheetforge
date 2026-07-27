# Spec — Part A, progressive disclosure on the Setup surface

**Date:** 2026-07-27 (revised same day after two owner decisions)
**Branch:** `feat/progressive-disclosure-part-a` (rebased onto `main` @ v0.5.0)
**Implements:** issue #6, Part A of `2026-07-26-setup-surface-progressive-disclosure-design.md`
**Decision record:** `docs/adr/0003-availability-default-and-legacy-migration.md`
**Panel:** `docs/panels/2026-07-27-part-a-availability-default.md` — 6 resolutions 3-0, **three not adopted**
**Groundwork:** `docs/designs/2026-07-27-part-a-groundwork.md`

> **STATUS: scope settled, no code written.** The design questions are answered
> (ADR 0003). What remains before implementation is `writing-plans`, and the
> prerequisite in §7.

---

## 1. What this changes

A fresh install draws a button for every action a package declares, and Setup shows every
field of every package. Part A inverts it: actions arrive `off`, the action catalog is the
first thing the user meets, and a package's fields surface as its actions are turned on.

The difficulty is not the inversion. It is that "no stored answer" resolves in four places
and one of those cases is a legacy store that must keep every button it has.

## 2. The four sites

| # | Site | Today | After |
|---|---|---|---|
| 1 | `ParseAvail` (`Core.bas:1928`) | blank cell → `AVAIL_BUTTON` | blank cell → `""` (unrecognised) |
| 2 | `AvailFor` (`Core.bas:2021`) | id absent → `AVAIL_BUTTON` | id absent → `AVAIL_OFF`, **unless the map is legacy-flagged** |
| 3 | `ReadAvailMap` (`Core.bas:2004`) | marker `absent` → empty map | marker `absent` → empty map **flagged legacy** |
| 4 | `AvailLabel` (`Core.bas:1991`) | `Case Else` → `"Button"` | `Case Else` → `"Off"` |

**Site 1 is one line.** `ParseAvail` already returns `""` for an unrecognised value
(`Core.bas:1937`), and `ComposeReport` already refuses to activate on it with the message
*"UNRECOGNISED Availability (use Button, Chains only, or Off)"* (`Core.bas:2562-2580`),
incrementing `structural`. **The Finalize hard-block already exists** — blank merely
short-circuits before reaching it.

## 3. Provenance: a flag, not a constant

`ReadAvailMap` must stop discarding what `MarkerVerdict` already knows. The returned map
carries a flag when it came from an `absent`-marker block, and `AvailFor` consults the
flag rather than a hardcoded constant:

- **legacy map** → absent id means **`button`**. A v0.3.0 deployment is unchanged.
- **any other map** (fresh, or current-and-marked) → absent id means **`off`**. Covers both
  a fresh install and a newly-registered action id.

The flag is **not optional**. Preserving legacy behaviour requires telling legacy apart
from fresh, and both are an empty map. There is no version of this change that keeps the
legacy guarantee without carrying provenance forward.

**Key:** `core.legacy`, under a `core.*` namespace reserved in contract §7, following the
existing `chain.*` precedent. §7 states no character rule for action ids, so a sentinel
cannot be asserted collision-proof — it must be reserved. `AvailFor` excludes reserved
keys from iteration; `ComposeReport` iterates the *registry*, not the map, so it is
unaffected.

## 4. The legacy path is untouched

**No migration, no retained backup, no refusal.** Contract §5 already defines `absent` →
read with the older layout, and §10 calls the asymmetry "correct and not a defect". There
is no installed base (ADR 0002), so machinery to convert it would protect a corpus of
zero, and refusing it would be a *breaking* change costing more than doing nothing.

Full reasoning, including the three panel resolutions this overrides, is in ADR 0003.

## 5. Contract changes — v2.1.0, minor

Additive under §10's two-axes rule: **no stored value is reinterpreted.** A stored
`button` still means `button`; the default only applies where nothing was ever stored.

1. **§6, line ~258** — `availability` "Defaults to `button`" becomes "Defaults to `off`",
   with a note that this changes no stored value's meaning.
2. **§5, lines ~237-244 — the load-bearing edit.** It currently says the legacy reader
   takes *"the documented default (`button` for `availability`)"*. Once §6 says `off`,
   that sentence **silently flips legacy behaviour and destroys §4 of this spec.** §5 must
   name `button` as a fixed legacy value, pinned, not as a pointer to §6.
3. **§7** — reserve the `core.*` action-id namespace for engine-internal sentinels.
4. **§3** — restate presentation for disclosure (relegation moving toward "hidden in
   place"), with **omitting fields still non-conforming**.

> Edit 2 is the one that breaks the feature if missed, and it is in a different section
> from the change that triggers it. Do it in the same commit as edit 1.

## 6. Field disclosure

`usedBy` is declared by **no shipping package** — every `PkgImportExport` field is
package-wide, `PkgClean` has none. So the relegation renderer is exercised only by
`PkgFixture`, and Part A's visible value rests entirely on wiring `usedBy` onto
`PkgImportExport`'s nine fields.

Fields are **relegated, never omitted** — unchanged standing rule. Rebuilding Setup
re-seeds from the finalized `_Config`, so omitting a field would cost the user
toggle → Finalize → rebuild → fill → Finalize and discard anything typed but not finalized.

**Inherited gate item.** Part B's item 1.4 could not be tested because nothing shipping
declares `usedBy`. Once this lands it becomes reachable, and **Part A's gate must include
it**: set every `importexport.*` action Off, confirm the band disappears *and* the off band
appears.

## 7. Prerequisite

**Part B's gate sections 3-6 must be run first** (PR #9). They verify `_Config` behaviour
already on `main` that Part A then modifies; running them first means a later failure is
unambiguously Part A's doing rather than inherited. Section 5 especially — it is the only
one whose failure is silent, and it needs workbook Q finalized under **v0.4.0**.

## 8. Out of scope

- **A designed first-run empty state.** Discarded at panel as a second design project.
  Part A ships a legible catalog with everything Off. Reopens on observed confusion.
- **Fixing the delete-then-copy activate** (`Core.bas:2769-2775`) — issue #8.
- **Any test suite.** TDD is formally off; Excel has no headless runner.

## 9. Verification

1. `pwsh build/build.ps1` → `Core + 2 package(s)`, balanced Sub/Function counts.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` → no diagnostics.
3. A manual gate written to the Part B standard (issue #3), covering at minimum: a fresh
   install shows everything Off; **a v0.3.0 legacy workbook keeps every button** (the
   §5-edit-2 regression); a marked v0.5.0 config keeps every stored choice; a
   newly-registered action id arrives Off; a blank Availability cell blocks Finalize with
   the existing message; and Part B's inherited item 1.4.

**Item 3's legacy case needs a v0.3.0-era workbook.** If none exists the test is
unrunnable, and that must be recorded in the gate as untested rather than ticked.

## Handoff log

- **2026-07-27** — Spec written from the 3-0 panel, then **narrowed twice by owner
  decision**. Reading Core discharged two panel revisit triggers: the Finalize hard-block
  already exists (§2), and no atomic `_Config` helper does — the activate is
  delete-then-copy (issue #8). Reading the contract killed the third: refusing legacy is
  *breaking*, not simplifying, so the legacy path is now left alone (§4) and the contract
  takes a minor bump instead of a major one.
- **Scope is settled. Next: run PR #9's gate (§7), then `writing-plans`.** The §5 edit-2
  trap is the single most important thing not to lose.
