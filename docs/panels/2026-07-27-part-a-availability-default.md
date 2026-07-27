# Panel — flipping the Availability default without zeroing legacy deployments

**Subject:** Part A (issue #6), how to invert the Availability default from `button`
to `off` when three provenances collapse into one empty map.
**Commit judged:** `bc89e93` (branch `feat/prompt-per-selection-settings`, PR #7 open).
**Context given to the panel:** `docs/designs/2026-07-27-part-a-groundwork.md`, with the
four default sites and `MarkerVerdict`'s verdict vocabulary pasted inline.

## Cost

| | |
|---|---|
| Rounds | 2 (round 2 targeted, not a re-file) |
| Executor | Codex CLI 0.145.0, via stdin — it cannot read files on this machine |
| Models | `gpt-5.6-terra` (native, veteran), `gpt-5.6-sol` (aesthete), medium effort |
| Tokens | ~82,100 across 6 calls — R1 36,200 / R2 45,900 |
| Independence | Full. No subagents; a different model family throughout. |

Codex's own tool calls failed with `CreateProcessAsUserW failed: 5` on every attempt, as
the skill documents. Every seat therefore reasoned from the inlined source only. **That
is the main caveat on this record:** no seat verified a line number independently, so
their citations are only as good as the excerpt I pasted.

## Resolutions

| # | Question | Resolution | Vote |
|---|---|---|---|
| 1 | Grandfather legacy forever, or migrate once? | **Migrate once, visibly.** Write every legacy action explicitly as `button` into a marked v2.0 block. Do not keep "marker absent means all on" as a permanent runtime branch. | 3-0, round 1 |
| 2 | Off for all actions, or only newly-registered ones? | **All actions on a fresh install.** For a current stored config, preserve known rows and default only missing ids to Off. | 3-0, round 1 |
| 3 | Blank Availability cell on a current config? | **Finalize hard-blocks it** — inline error, refuse to save. Runtime keeps `Off` as a defensive fallback only. | 3-0 after both others moved to the aesthete |
| 4 | Where does the migration write happen? | **Finalize only**, never `Core_BuildSetup` — BuildSetup is the destructive re-seed and must not touch the only persisted state before confirmation. | 3-0, round 2 |
| 5 | What if migration is interrupted? | **Stage, validate, then commit — marker written LAST.** A partial write has no marker but does have the `Availability` header, which `MarkerVerdict` already classifies `"tampered"` and already refuses. Legacy stays authoritative until the marker lands. | 3-0, round 2 |
| 6 | Is a one-time dialog enough evidence? | **No. Retain the legacy block** as a hidden dated backup outside the active keyed block, so `ReadAvailMap` cannot consume it. A dialog leaves no evidence that a wrong conversion happened. | 3-0 after the native reversed |

## What moved

- **The native reversed on resolution 6.** Round 1 it proposed a one-time notice and
  nothing more. Round 2, pressed on the chair-injected gap — that after migration the
  evidence of migration is gone — it ruled a persistent audit record *necessary*, and
  went further than the others by specifying the backup must live outside the active
  keyed block so `ReadAvailMap` cannot read it back as config. That was the single most
  useful sentence the panel produced.
- **The native and veteran both moved to the aesthete on resolution 3.** Round 1 both
  said blank reads `Off` by the fail-closed idiom. Round 2 both accepted that fail-closed
  governs *runtime* and does not license committing damage to the only store.
- **The veteran upgraded resolution 5.** It accepted the native's marker-last ordering,
  then correctly objected that *detection is not recovery* — which turned "write the
  marker last" into "stage the whole block, validate it, then commit".
- **Chair-injected gap.** Neither the interruption case nor the post-migration evidence
  case appeared in round 1. Both were put in by the chair as the fourth-seat question,
  and both produced movement. That is the argument for having asked it.

## Chair's discards

Recorded so they are not quietly reinstated later:

- **The aesthete's empty-state programme.** It answered its own strongest objection —
  all-Off makes first run look empty — with "a polished catalog, clear action
  descriptions, sensible grouping, and an explicit empty-state prompt". That is a second
  design project in a subordinate clause, and it is the aesthete's signature failure of
  ignoring cost. Part A ships a legible catalog with everything Off; it does not also
  grow a designed empty state.
- **The veteran's "transactional/rollback mechanism".** Over-general. A worksheet has no
  transactions. Marker-last plus a retained legacy block is the achievable form and is
  what resolution 5 takes.

## Revisit triggers

Falsifiers, not opinions. A row reopens on logged evidence only.

| Trigger | Reopens | Where evidence comes from | Observations |
|---|---|---|---|
| `Core.bas` turns out to have an existing validator that deliberately normalizes blank enum cells during Finalize | 3 — blank would then have a documented meaning and the hard-block would be inconsistent with it | Read `Core_FinalizeSetup`'s validation pass before implementing | — |
| An atomic `_Config` replacement helper already exists | 5 — staging could be simpler than the panel assumed | Same read | — |
| A legacy (v0.3.0, marker-absent) workbook is confirmed *not* to exist anywhere | 1, 2, 6 — with no legacy corpus the whole migration path is dead code and should not be built | Owner's own workbook inventory | — |
| The retained legacy backup block is read back as config by any reader | 6 — the backup location is wrong and is actively harmful | A manual gate item that finalizes twice and confirms the backup never appears on Setup | — |
| A migration runs cleanly on a real legacy workbook | 5, 6 — a clean run is evidence too; several clean runs argue for retiring the backup at some later version | The Part A manual gate | — |
| First-run with everything Off is observed to confuse a real user | Chair's discard of the aesthete's empty-state programme | Owner's own first-run, and the finance team's | — |

## Open, and NOT decided by this panel

The panel ruled on mechanism. It did not rule on **whether Part A ships at all**, nor on
the contract §3 restatement the design doc flags (presentation moving from "one collapsed
band at the end" to "hidden in place", with omitting fields still non-conforming). Those
remain the owner's.
