# ADR 0003: Invert the Availability default, and migrate legacy configs once

<!-- File: docs/adr/0003-availability-default-and-legacy-migration.md — one page max. Never edited after acceptance; reversals are a new ADR. -->

**Status:** Proposed — **not accepted, and no code written against it**
**Date:** 2026-07-27

Drafted overnight from a 3-0 Codex panel (`docs/panels/2026-07-27-part-a-availability-default.md`).
The panel ruled on **mechanism only**. It did not rule on whether Part A ships, and the
owner has not seen it. **This ADR is blocked on one question of fact — see Decision, note 1.**

## Context

Part A (issue #6) inverts the Setup surface: the action catalog comes first, actions
default to `Off`, and a package's fields surface as its actions are turned on.

The obstacle is that "no stored answer" currently means four different things, resolved
in four places, all of which answer `button`:

| Site | Today |
|---|---|
| `ParseAvail` (`Core.bas:1928`) | blank cell → `button` |
| `AvailFor` (`Core.bas:2021`) | id absent from map → `button` |
| `ReadAvailMap` (`Core.bas:2004`) | marker `absent` → empty map, so everything → `button` |
| `AvailLabel` (`Core.bas:1991`) | `Case Else` → `"Button"` |

Three provenances collapse into one empty map:

- **fresh** — no `_Config`. Part A wants `Off`.
- **legacy** — a v0.3.0-era `_Config` with no format marker and no Availability column.
  Must stay `button`; flipping it makes every button in a working deployment vanish at
  once, silently, with no undo.
- **current-but-missing-id** — a marked config where a package upgrade added an action id
  with no row yet. Part A wants `Off`.

Fresh and legacy are indistinguishable *at the point of use* today. They are **not**
indistinguishable at the point of *reading*: `MarkerVerdict` (`Core.bas:2229`) already
separates `absent` from `same-major`, and `ReadAvailMap` throws that bit away at its
`absent` early-exit.

v0.5.0 narrowed the problem. A v0.4.0 config is now `other-major` and refused outright
(ADR 0002), so the only live migration path in is the legacy one — exactly two
provenances, needing opposite defaults.

## Decision

1. **`ReadAvailMap` carries provenance instead of discarding it.** The returned map is
   flagged when it came from an `absent`-marker block, under a reserved key. `AvailFor`
   consults the flag: a legacy map means an absent id is `button`; any other map means an
   absent id is `off`.
2. **Reserve `core.*` in contract §7** for engine-internal map sentinels, following the
   existing `chain.*` precedent. §7 states no character rule for action ids, so a sentinel
   key cannot be asserted collision-proof — it has to be *reserved*. This makes Part A a
   contract change requiring a minor bump.
3. **Legacy migrates exactly once, at Finalize, never at `Core_BuildSetup`.** BuildSetup
   is the destructive re-seed; it must not mutate the only persisted state before the user
   confirms.
4. **The format marker is written last.** An interrupted migration leaves the
   `Availability` header with no marker, which `MarkerVerdict` already classifies
   `"tampered"` and `ComposeReport` already refuses. The failure mode becomes a refusal
   with an explanation instead of a silent half-migration, reusing machinery rather than
   inventing it.
5. **The legacy block is retained as a dated backup** under its own column-A key, outside
   any block a reader consumes. A one-time dialog is a notice, not evidence: without the
   retained block, a wrong conversion is indistinguishable from user intent forever.
6. **A blank Availability cell blocks Finalize.** `ParseAvail` returns `""` for blank as
   it already does for any unrecognised value, and the existing structural block at
   `ComposeReport` (`Core.bas:2562-2580`) refuses to activate with the message it already
   carries. Runtime keeps `off` as a defensive fallback. Fail-closed governs *runtime*; it
   does not license committing damaged input to the only store.

> **Note 1 — the blocking question, and what the repo already says about it.**
> Decisions 3, 4 and 5 exist solely to protect v0.3.0-era workbooks. **If none exists,
> they guard nobody.**
>
> The panel was not told this, but **ADR 0002 already records the answer, dated the same
> day: "there is no installed base. The bundle exists only in test workbooks."** If that
> holds for v0.3.0-era workbooks too — and it should, since it is a statement about
> deployments, not about versions — then the entire migrate-once-with-backup apparatus in
> decisions 3-5 protects a corpus of zero real users plus, at most, one test fixture
> (the Part B gate's "workbook C").
>
> **The owner's own reasoning in ADR 0002 then applies verbatim:** *"the false refusal
> this ADR was written to prevent costs nothing today… simpler wins while it is free."*
> Rebuilding a test workbook costs a rebuild; it does not cost a finance team's workday.
> On that reasoning the consistent decision is the simple one — either refuse a legacy
> config outright, exactly as v0.4.0 configs are now refused, or migrate it plainly
> without the retained-backup machinery.
>
> **I have not made that call.** It reverses three unanimous panel resolutions on a fact
> the panel never saw, and it is the owner's to take. But the ADR would be wrong to
> recommend building decisions 3-5 without putting this in front of them first.
> **Decisions 1, 2 and 6 stand either way and are unaffected.**

## Consequences

- **A v0.3.0 workbook keeps every button**, then converts once with a visible notice.
- **New actions arrive switched off** on every existing deployment, which is the feature.
- **A fresh install shows an empty catalog.** Accepted deliberately. The panel's aesthete
  wanted this answered with a designed empty state — a polished catalog, action
  descriptions, grouping, an empty-state prompt — and that was **discarded as a second
  design project smuggled into a subordinate clause**. Part A ships a legible catalog with
  everything Off. If first-run confusion is actually observed, that reopens as its own
  piece of work, on evidence.
- **The contract takes a minor bump** for the `core.*` reservation and the §3 presentation
  restatement (relegation moving from "one collapsed band at the end" to "hidden in
  place", with omitting fields still non-conforming).
- **`usedBy` must be wired onto `PkgImportExport`'s nine fields**, or Part A has no
  visible effect — no shipping package declares it today. This also makes Part B's gate
  item 1.4 testable for the first time; Part A's gate inherits it.
- **Not fixed here: the activate path is delete-then-copy** (`Core.bas:2769-2775`) and has
  a window with no `_Config` at all. Pre-existing, out of scope, and it bounds what
  "staged" can honestly mean in decision 4 — marker-last ordering within a block, not a
  transaction over the sheet. Deserves its own issue.

## Alternatives rejected

- **Grandfather legacy permanently** (a missing marker meaning "all on" forever). Rejected
  3-0: it keeps an unmarked block as executable policy indefinitely, so a later refactor
  that reads "empty" as the new default-off disables production workflows anyway — the
  failure the veteran seat had seen before. Migration makes the state explicit and
  inspectable.
- **Flip the default only for newly-registered actions.** Rejected 3-0: it leaves
  first-time users with today's overwhelming surface and makes the product's default
  depend on an arbitrary release boundary.
- **A one-time dialog as the only record of migration.** Proposed by the native seat in
  round 1 and **reversed by that same seat** in round 2: after migration the evidence of
  migration is gone, and `_Config` is the only memory that survives a rebuild.
