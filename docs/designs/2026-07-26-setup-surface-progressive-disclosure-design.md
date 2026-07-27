# Design — progressive disclosure on the Setup surface, and prompt-vs-Setup for per-run options

**Date:** 2026-07-26
**Branch:** proposed for a phase *after* `feat/finish-excel-track1` ships v0.4.0
**Status:** design only. Nothing here is implemented. Requires `brainstorming` →
`writing-plans` → an ADR before any code.
**Supersedes nothing.** Extends the Setup surface described in
`docs/designs/2026-07-25-track1-owner-feedback-plan.md` and `spec/contract.md` §3/§5/§6.

**Origin.** Both parts came out of Task 10 Step 6 of the v0.4.0 owner gate, when the
owner hit `Does the range you pick include a header row?` and said: *"I thought this
was going to be a popup, not a setup cell. This is exceptionally bad placement — users
will never find this and will be frustrated by it."* That is a real defect report
against a shipped surface, not a feature request, and it splits into two independent
problems with different answers.

---

## The two problems

| | Problem | Answer |
|---|---|---|
| **A** | Every field for every action is on Setup at once, whether or not that action is even turned on. A user cannot tell which settings apply to what they are about to do. | Progressive disclosure: catalog first, actions default Off, a field appears when an action that uses it is turned on. |
| **B** | Some settings are properties of *the selection you just made*, not of the deployment. Putting them on Setup means changing them requires editing Setup and re-Finalizing between two runs of the same button. | Those belong in a prompt at run time, not on Setup at all. |

They are separable. **B** is a small package-level change; **A** is a Core change with
contract and migration consequences. Neither belongs in v0.4.0.

---

## Part A — progressive disclosure

### The owner's proposal, as stated

Keep the Setup surface structurally as it is, but:

1. Move the Action catalog to the **top** of the sheet.
2. Default every action's Availability to **Off**.
3. Hide every other option initially.
4. As an action is switched to `Button` or `Chains only`, the settings that action
   uses become visible.
5. The Chains authoring block appears only once at least one action is set to
   `Chains only`, so the user connects the block to the change they just made.

### What already exists — this is mostly assembly, not invention

- **The field↔action binding.** `usedBy` is a declared, specified, optional attribute
  on `Core_AddField` / `Core_AddList` / `Core_AddTable` (`Core.bas:244-267`), frozen
  in contract §3 at v1.2.0.
- **The exact rule.** `FieldInUse` (`Core.bas:342-355`): a field with no `usedBy` is
  package-wide and always shown; a field is relegated only when **every** action it
  names is Off; an unknown id counts as in-use, so a typo can never make a field
  unreachable. This *is* rule 4 above, already written.
- **A relegation renderer.** `Core.bas:1484` already collects "every field whose
  `usedBy` names only Off actions, from every package" into one band.
- **Proof it works.** `PkgFixture` exercises it (`fx_always` declares no `usedBy`;
  another field declares an unknown id). **Neither shipping package uses it.**
- **Buttons on a sheet.** `AddBtn` + `OnAction` is established machinery, and the
  hideable Setup Menu tab (`BuildAdminPanel`) already exists as a home for controls.

So the mechanism is paid for. What changes is its **presentation** (a collapsed band
at the end → hidden in place), its **ordering** (catalog last → catalog first), and
its **default** (everything on → everything off).

### The hard constraint: this cannot be live

Unhiding rows the moment a dropdown changes requires `Worksheet_Change`, which only
fires from the sheet's own code module. Writing there needs *"Trust access to the VBA
project object model"* — **banned outright** by CLAUDE.md rule 3, which is not
negotiable and is the reason package discovery probes `Application.Run` instead of
enumerating modules.

The usual escape, an Application-level `WithEvents` class, needs a `.cls`. Delivery is
one concatenated `.bas` (`docs/adr/0001-source-first-vba-toolkit.md`), and `src/`
contains no class module, form, or event handler of any kind — verified.

**Therefore the interaction is click-driven, not reactive:** set availability → click
Refresh → the page reorganizes. This is a real limitation and must be stated to the
user rather than discovered by them.

### Making click-driven feel intentional: the pending-changes formula

A worksheet formula recalculates the instant a dropdown changes, with no events and no
trust dependency. That gives live *feedback* even though the rebuild is manual.

**Rejected predicate.** "All actions Off → `""`, otherwise show the message." Once the
user turns on a single action the condition stays true forever, so the message persists
after the refresh it asked for. Permanent on-screen text is text people stop reading.

**Correct predicate: availability has changed since the page was last rendered** —
false immediately after a refresh, true the moment anything moves.

`Core_BuildSetup` already knows what it rendered against; it writes
`AvailLabel(AvailFor(priorAvail, id))` into each Availability cell (`Core.bas:2056`).
Have it also stamp that token into a **hidden snapshot column** on the same row. The
message cell is then a straight comparison:

```
=IF(SUMPRODUCT(--(live<>snapshot))=0,"",
    SUMPRODUCT(--(live<>snapshot))&" change(s) pending - click Refresh Options to show their settings")
```

Self-clearing, gives a count, needs nothing but recalculation.

**Owner's decision for the first test build:** the message also **links to the Setup
Menu tab**, and a Refresh button is placed **on the Setup Menu tab as well as** at the
top of Setup — deliberate redundancy, so that a failure or omission of the on-sheet
button does not block the test. **After testing, decide where the single button
lives** and remove the other. Do not ship both permanently without that decision.

### The migration trap — the one thing that must not be got wrong

`AvailFor` returns `Button` when an id is absent, **deliberately** (`Core.bas:1974-1979`),
and `ReadAvailMap` exits early returning an empty map when `MarkerVerdict` reads
`absent` (`Core.bas:1959-1972`) — which is what a v0.3.0 `_Config`, having no
Availability column at all, produces.

Flip the default to Off naively and **every existing v0.3.0 deployment upgrades to zero
buttons, silently.** That is exactly the class of failure Task 10 Step 7 exists to catch.

It is solvable, because the format marker already distinguishes the two cases:

- marker `absent` (pre-v1.2 `_Config`) → resolve missing ids to **`Button`**, unchanged;
- marker present, id genuinely unset (new deployment) → resolve to **`Off`**.

The point is that "absent" currently means one thing and would come to mean two. That
distinction has to be designed and tested, not inherited.

### Contract and process consequences

- **§3** — `usedBy` presentation changes from "one collapsed band at the end" to
  "hidden in place". The portability note (a platform that cannot collapse MUST still
  render the band, visibly and last) needs restating for hiding, and **omitting
  fields remains non-conforming**.
- **§5** — the hidden snapshot column is a new column on the Setup sheet, and
  `_Config` is a *copy* of that sheet. It interacts with the schema marker in the
  `LAYOUT_KEY` row (`Core.bas:2043-2044`) and with the v0.3.0 column-shift migration.
  Additive, but not free.
- **§6** — the Availability default flips.
- **ADR required.** The Off default reverses a decided position, and contract §3
  already considered and rejected a two-phase flow on the grounds that omitting
  fields "would force a second activate cycle and discard un-activated input".
  Per CLAUDE.md this needs an ADR here **and** a note queued to
  `~/.daftplate/outbox/`. Silent deviation is the one unforgivable move.
- **No version bump yet.** Contract v1.2.0 is unreleased on the current branch.

### Smaller things not to trip over

- A Refresh button on Setup must survive `Core_BuildSetup`'s own rebuild and must not
  be caught by `SweepPlacedButtons`, which matches the `SFPlaced_` name prefix.
- The whole design depends on Excel actually recalculating, which depends on
  `Application.Calculation` always being restored after a failed action. That is the
  speed-pattern guarantee and it is what Task 10 Step 6's third item verifies.
- Ordering the catalog first is safe for reads: `FindKeyRow` locates blocks by key,
  not position.

---

## Part B — options that belong in a prompt, not on Setup

### The rule

**A setting that describes the thing the user just picked belongs in a prompt at run
time. A setting that describes the deployment belongs on Setup.**

The owner's framing, recorded verbatim in substance: a one-off, on-the-fly macro is not
a sequence of events and has no list of things to get right. The user may reasonably
change their mind between two runs, and should not have to remember — and probably will
not remember — a stored setting for something that straightforward. By contrast
`1. Clear Input Tabs` runs in a series with other macros and carries a list of tabs the
user must enter carefully. That is Setup's job.

Retaining a prompt's answer between runs is **not** required. Asking every time is the
feature: it is what makes the answer match the selection.

### The determination, run now against both shipping packages

| Action | Fields it uses | Verdict |
|---|---|---|
| `clean.dupes` | `clean_keycols`, `clean_hasheader` | **Move both to a prompt.** |
| `clean.trim` | none | nothing to move |
| `clean.filldown` | none | nothing to move |
| `importexport.clear` | `k_import` (table of input tabs) | **Stays on Setup** — the owner's own example |
| `importexport.import` | `k_import` | stays |
| `importexport.updateMaster` | `k_master_prefix`, `k_master_dest`, `k_master_sheet`, `k_master_cols` | stays |
| `importexport.export` | `k_export_base`, `k_export_tabs`, `k_export_offset` | stays |
| `importexport.suggestTypes` | reads the live Setup sheet by design | stays |

**Only `clean.dupes` qualifies, and both of its fields move.** `clean_keycols` is the
stronger case of the two, not the weaker: its own help text defines it as *"numbered
1,2,3… from the LEFT of the range you pick"* — it is meaningless except in relation to
a selection that does not exist until run time.

The `importexport.*` actions all fail the test in the same way: their settings are
lists and tables describing this deployment's workbook and folder, they are used by
actions numbered to run in sequence, and re-entering them per run would be worse in
every respect.

### Two consequences worth knowing before agreeing

1. **PkgClean would be left with no fields at all**, so `usedBy` hide/unhide has
   nothing to demonstrate in a shipping package. Part A's value would rest entirely on
   `PkgImportExport`'s fields — which are genuinely Setup-shaped, so this is a
   presentational loss, not a functional one. But it means Part A cannot be
   demonstrated using the very field that motivated it.
2. **`clean_hasheader` is the contract's first and only use of the `bool` kind**
   (§3 and §10 both record the reservation being taken up by it). Moving it to a
   prompt returns `bool` to unused, and the reservation language in §3/§10 has to be
   revisited rather than left describing a field that no longer exists.

### How to keep the determination honest as packages are added

The scan above is a point-in-time answer for two packages. The durable version is a
test applied when an action is declared:

> Does this setting describe the selection the user is about to make, or the workbook
> the deployment lives in? If a user could reasonably want a different answer on two
> consecutive runs of the same button, it is a prompt.

Re-run the table above whenever a package is added or an action gains a field, and
record the verdict next to the action. If later testing surfaces more candidates, they
join Part B's list rather than being handled ad hoc.

---

## Sequencing

1. Finish the Task 10 gate and ship **v0.4.0** unchanged. The gate verifies the
   migration and fail-closed machinery both parts build on; redesigning first means
   building on an unverified foundation and not knowing which layer a later failure
   came from.
2. Then `brainstorming` → `writing-plans` → ADR → build, **Part B first** (contained,
   package-level, immediately felt) and **Part A second** (Core, contract, migration).
3. Re-run the Task 10 checklist against a released v0.4.0 baseline rather than against
   an unmerged branch.

## Open questions for the next phase

- Does Part A's hide/unhide use Excel row outlining (collapsible, per contract §3's
  portability reference) or plain `Rows.Hidden`? Outlining is discoverable and
  reversible by the user; hiding is cleaner but a user who unhides manually then sees
  a surface the engine will overwrite.
- Where does the single Refresh button live — Setup, Setup Menu, or both? **Deferred
  to the test build by owner decision.**
- Should the pending-changes message name the actions that changed, or only count
  them? Naming is more useful and more fragile.
- `ANCHOR_BTN_W` at 110pt is still carried forward from v0.4.0 as an owner-judgement
  item, unrelated to this design but unresolved.
