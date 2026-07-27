# Part A groundwork — where the Availability default actually lives

**Date:** 2026-07-27
**Status:** research only. **This is not a plan and decides nothing.**
**For:** issue #6, Part A of `2026-07-26-setup-surface-progressive-disclosure-design.md`
**Pipeline position:** input to `brainstorming`, which has not happened yet.

Part A "should not be started casually" — it needs brainstorming → writing-plans → an
ADR. This document exists so that the brainstorm starts from facts about the code
rather than from a guess, because the migration trap it carries is real and I could
establish its exact shape without deciding anything.

---

## The headline

**"Flip the Availability default to Off" is not a one-line change. The Button default
is baked into four sites, and they do not mean the same thing.** Three distinct
provenances currently collapse into one indistinguishable "no answer", and Part A needs
two of them to diverge.

| # | Site | Today | What "no answer" means there |
|---|---|---|---|
| 1 | `ParseAvail` (`Core.bas:1928`) | blank cell → `AVAIL_BUTTON` | The row exists; the user cleared the cell. |
| 2 | `AvailFor` (`Core.bas:2021-2025`) | id not in map → `AVAIL_BUTTON` | The action has no row at all. |
| 3 | `ReadAvailMap` (`Core.bas:2011`) | marker `absent` → **empty map** | The whole block predates the Availability column. |
| 4 | `AvailLabel` (`Core.bas:1991-1995`) | `Case Else` → `"Button"` | Display only, but it renders the seeded cell. |

Site 3 is the trap. It returns an empty map, every id then misses at site 2, and site 2
says Button. That chain is what makes a v0.3.0 workbook come back with every button
intact — the intended behaviour today, and the thing that breaks silently if site 2's
constant is simply flipped.

---

## The trap, stated precisely

`AvailFor` returning Off on a miss would make all three of these read Off:

1. **Fresh workbook, no `_Config`.** Part A *wants* Off — this is the new default.
2. **A v2.0 config where a package upgrade added a new action id.** Part A *wants* Off —
   a new action arrives switched off, which is the whole point of the feature.
3. **A v0.3.0-era `_Config`, marker absent, no Availability column at all.** Part A must
   **not** make this Off. Every button in an upgraded deployment would vanish at once,
   with no error and nothing on screen to explain it.

Cases 1 and 3 are *both* an empty map today. That is the entire difficulty. Nothing
downstream of `ReadAvailMap` can tell them apart, because the emptiness is all that
survives.

**The marker already carries the missing bit.** `MarkerVerdict` (`Core.bas:2229`) is the
single place that classifies a block, and it separates `absent` (genuine pre-v1.2
layout) from `same-major`, `other-major`, `unparseable` and `tampered`. So the
information Part A needs is present and already centralized — it is just discarded at
`Core.bas:2011`, where `absent` becomes `Exit Function` and the provenance is lost.

That suggests the shape of the answer without settling it: **the default has to travel
with the map rather than be a constant at the point of use.** Whether that is a flag on
the returned Dictionary, a second return value, or a distinct verdict passed down to
`AvailFor`, is a brainstorm question, not one to settle here.

---

## What v0.5.0's refusal already did for Part A

Worth knowing before the brainstorm, because it narrows the problem.

Since `SCHEMA_VER` moved to `"v2.0"`, a v0.4.0 config is `other-major` and is **refused,
not migrated**. So Part A does not have to carry a v0.4.0 migration at all. The live
provenances reduce to exactly two:

- **fresh** — no `_Config`, or no `LAYOUT_KEY` row (`ConfigSchemaStatus` → `"none"`)
- **legacy** — marker `absent`, v0.3.0-era (`ConfigSchemaStatus` → `"legacy"`)

Two cases, and they need opposite defaults. That is a much smaller problem than it
looked before the version decision, and it is an argument that Part A got *easier* by
shipping v0.5.0 first.

---

## The second thing Part A owes

**`usedBy` is not wired to any shipping package.** Every field in `PkgImportExport`
declares no `usedBy`, so it is package-wide and never relegates
(`FieldInUse`, `Core.bas:385-400`). `PkgClean` now declares no fields at all. So the
relegation renderer — `RenderOffBand`, plus the `HasAnyTier` guard Part B just added — is
exercised **only by `PkgFixture`** today.

Two consequences:

- Part A's user-visible value rests entirely on `PkgImportExport`'s nine fields, exactly
  as the design doc predicted at line 201. Wiring `usedBy` onto them is not a nicety; it
  is the feature.
- **Part B's gate item 1.4 is inherited by Part A's gate.** The empty-band guard's
  second direction — a package whose every field has relegated — cannot be tested from a
  shipping package until this wiring exists. It is corrected and marked as such in
  `2026-07-27-part-b-manual-gate.md`.

---

## Deliberately not decided here

- Whether the default flips for *all* actions or only for newly-registered ones.
- Whether legacy deployments are grandfathered permanently, or migrated once with a
  visible one-time notice.
- Whether `ParseAvail`'s blank-cell default (site 1) should move at all — a cleared cell
  on a *current* config is arguably a different question from an absent row.
- The contract §3 restatement the design doc flags at line 131 (presentation moving from
  "one collapsed band at the end" to "hidden in place", with omitting fields still
  non-conforming).

Each of those changes what gets built. They are brainstorm input, not findings.
