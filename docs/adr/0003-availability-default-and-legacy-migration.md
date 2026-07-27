# ADR 0003: Invert the Availability default, and leave the legacy path alone

<!-- File: docs/adr/0003-availability-default-and-legacy-migration.md — one page max. Never edited after acceptance; reversals are a new ADR. -->

**Status:** Proposed — decisions settled by the owner 2026-07-27, no code written yet
**Date:** 2026-07-27

Drafted from a 3-0 Codex panel (`docs/panels/2026-07-27-part-a-availability-default.md`),
then **substantially narrowed by two owner decisions the panel never had the facts for.**
The panel's migration resolutions (3, 4, 5 in its record) are **not adopted**; see
Alternatives rejected. What survives is smaller and cheaper than what the panel designed.

## Context

Part A (issue #6) inverts the Setup surface: the action catalog comes first, actions
default to `off`, and a package's fields surface as its actions are turned on.

The obstacle is that "no stored answer" is resolved in four places, all answering
`button`:

| Site | Today |
|---|---|
| `ParseAvail` (`Core.bas:1928`) | blank cell → `button` |
| `AvailFor` (`Core.bas:2021`) | id absent from map → `button` |
| `ReadAvailMap` (`Core.bas:2004`) | marker `absent` → empty map, so everything → `button` |
| `AvailLabel` (`Core.bas:1991`) | `Case Else` → `"Button"` |

Three provenances collapse into one empty map: **fresh** (no `_Config`), **legacy** (a
pre-v1.2 store with no marker and no Availability column), and **current-but-missing-id**
(a marked store where a package upgrade added an action id with no row yet). Part A wants
`off` for the first and third. The second must keep `button`.

Fresh and legacy are indistinguishable at the point of *use*, but not at the point of
*reading*: `MarkerVerdict` (`Core.bas:2229`) already separates `absent` from
`same-major`, and `ReadAvailMap` discards that bit at its `absent` early-exit.

## Decision

1. **`ReadAvailMap` carries provenance instead of discarding it.** The returned map is
   flagged when it came from an `absent`-marker block. `AvailFor` consults the flag: a
   legacy map means an absent id is `button`; any other map means an absent id is `off`.
   **The flag is unavoidable** — preserving legacy behaviour *requires* telling legacy
   apart from fresh, and an empty map cannot do that on its own.
2. **Reserve `core.*` in contract §7** for engine-internal map sentinels, following the
   existing `chain.*` precedent, and use `core.legacy` as the flag key. §7 states no
   character rule for action ids, so a sentinel cannot be asserted collision-proof — it
   has to be reserved.
3. **The legacy path is otherwise untouched.** No migration, no retained backup, no
   refusal. Contract §5 already defines `absent` → read with the older layout, and §10
   calls the resulting asymmetry "correct and not a defect". It stays conforming, it
   stays dead code, and it costs nothing.
4. **Contract §5 must pin `button` literally.** §5 currently tells the legacy reader to
   take *"the documented default (`button` for `availability`)"*. Once §6's documented
   default becomes `off`, that sentence silently flips legacy behaviour and destroys
   decision 3. **This is the load-bearing edit of the whole change** — §5 must name
   `button` as a fixed legacy value, not as a pointer to §6.
5. **A blank Availability cell blocks Finalize.** `ParseAvail` returns `""` for blank as
   it already does for any unrecognised value, and the existing structural block in
   `ComposeReport` (`Core.bas:2562-2580`) refuses to activate with the message it already
   carries. Runtime keeps `off` as a defensive fallback. Fail-closed governs *runtime*; it
   does not license committing damaged input to the only store.
6. **The contract takes a MINOR bump to v2.1.0.** Under §10's two-axes rule the flip is
   **additive**, not breaking: no stored value is reinterpreted — a stored `button` still
   means `button` — because the default only applies where nothing was ever stored. The
   Setup-surface disclosure change is engine presentation, also additive.

## Consequences

- **`SCHEMA_VER` does not move.** v0.5.0 configs stay readable, so workbooks P and Q
  survive and the outstanding gate run (PR #9) is not invalidated.
- **A v0.3.0 workbook keeps every button**, exactly as today, with no migration event.
- **New actions arrive switched off** on every existing deployment. This is the feature.
- **A fresh install shows an empty catalog.** Accepted deliberately. The panel's aesthete
  wanted this answered with a designed empty state — polished catalog, action
  descriptions, grouping, an empty-state prompt — **discarded as a second design project
  smuggled into a subordinate clause.** If first-run confusion is actually observed, that
  reopens as its own work, on evidence.
- **`usedBy` must be wired onto `PkgImportExport`'s nine fields**, or Part A has no
  visible effect — no shipping package declares it today. This also makes Part B's gate
  item 1.4 testable for the first time; Part A's gate inherits it.
- **Not fixed here: the activate path is delete-then-copy** (`Core.bas:2769-2775`), with a
  window holding no `_Config` at all. Pre-existing, filed as issue #8.

## Alternatives rejected

- **Migrate legacy once, marker-last, with a retained dated backup.** The panel's
  resolutions 3-5, unanimous. **Rejected on a fact the panel was never given:** ADR 0002
  records "there is no installed base — the bundle exists only in test workbooks", so the
  apparatus protects a corpus of zero. The owner's own reasoning from 0002 applies
  verbatim: *simpler wins while it is free.*
- **Refuse a marker-absent config outright**, as v0.4.0 configs now are. Chosen by the
  owner first, then **reversed once the price was established**: §5 states "Absent is not
  invalid" as a named rule and §10 calls the asymmetry "correct and not a defect", so
  refusing is a **breaking** change to v3.0.0. Under ADR 0002's coupling `SCHEMA_VER`
  would go to `"v3.0"` and refuse every v0.5.0 config too — invalidating workbooks P and
  Q mid-gate. It costs strictly more than doing nothing and buys only the removal of dead
  code.
- **Grandfather legacy permanently.** Rejected 3-0 by the panel as an unmarked block
  becoming executable policy forever. **Adopted anyway**, because the panel's objection
  assumed the behaviour was undocumented — contract §5 defines it explicitly, and decision
  4 pins it against drift, which is what the panel was actually worried about.
- **Flip the default only for newly-registered actions.** Rejected 3-0: it leaves
  first-time users with today's overwhelming surface and makes the default depend on an
  arbitrary release boundary.
