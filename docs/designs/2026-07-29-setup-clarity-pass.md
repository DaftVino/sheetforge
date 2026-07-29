# Setup clarity pass — spec and plan

**Status:** planned 2026-07-29, after the owner's first real Excel run of the v0.18.0 bundle.
**Branch:** `feat/setup-clarity` off `feat/macros-home` (PR #32 is open and unmerged).
**Scope:** three owner-reported defects on the Setup tab. No new actions, no new package, no
contract change. Contract v2.10.0 stays frozen — nothing here touches an action id, a config
key, or a schema marker.

The bundle works: the owner imported it and ran it. Everything below is about a Setup tab that
works but does not explain itself.

---

## 1. The Finalize reminder is incomplete and clipped

### What is wrong

`REMINDER_TEXT` (`src/core/Core.bas:183`) reads:

> Actions are switched on. Click 'Finalize Setup Tab' and then 'Create Setup Tab' to reveal
> their option fields.

Two defects.

**It stops one step early.** It describes the first half of the loop and then goes quiet at
exactly the moment the user has new fields in front of them and no idea that answering them is
not enough. The setup guide written in Task 11 had to spell this out precisely because the
sheet does not: the loop is Finalize → Create → answer → **Finalize again**. The reminder
teaches two of those four steps.

**It is clipped, and structurally so.** The formula is written to `VAL_COL` — column C, width
34 — and the cell immediately right of it, `VAL_COL + 1` (column D), holds the `Setup Menu`
nav link. Excel spills overflow text only across *empty* cells, so the nav link is a hard wall
at 34 characters. The text is ~130. It was never going to fit; widening C would not fix it and
would wreck every other block on the sheet.

### What it becomes

New text, ~180 characters, still using only single quotes so it remains safe inside a formula
string literal:

> Actions are switched on. Click 'Finalize Setup Tab', then 'Create Setup Tab' to show their
> option fields. Fill those in, then click 'Finalize Setup Tab' again to save them.

**Moves to `D1`.** Row 1 currently holds only `B1` = "Toolkit Setup". `E1` through `H1` are
empty, so a `D1` value overflows across 34+30+3+26+42 = 135 characters of width. That is the
whole point of choosing D1 over a wider column: the room is in the *empty neighbours*, not in
the column.

Three things make D1 safe, each verified against the code rather than assumed:

- **AutoFit cannot reach it.** `ws.Rows(firstContentRow & ":" & r).AutoFit` (line 2253) starts
  at `firstContentRow`, which is set after the reference-list header rows — row 6, not row 1.
  A row-1 write is outside that range, so the row keeps its default height and the text
  overflows sideways instead of being clipped to a wrapped height.
- **Do not set `WrapText`.** Wrapping is what turns "overflows into empty neighbours" into
  "clipped to one row height". This is the trap that produced the original defect.
- **The user is already looking at it.** `Core_BuildSetup` ends with `ws.Range("B1").Select`
  (line 2263). The reminder lands two cells from the cursor, on the row the eye starts on.

**Bold, and `RGB(150, 0, 0)`.** Not `ForgeRed()` — that helper's own comment says *"brand
accent - fills only, never behind text"*, and this is text. `RGB(150, 0, 0)` is the dark red
already used for the "Do not edit column A" warning at line 2182, so the sheet gains no new
colour.

### What must not change

- **The formula is still written after the action rows exist.** Do-not-revert item 2 of the
  2026-07-24 handoff log: an earlier draft computed the block's extent from a registry count,
  and the same `lastDataRow` bounds a `.Validation` call whose own comment warns about leaking
  outside the block. The write moves from `Cells(reminderRow, VAL_COL)` to `Cells(1, 4)` and
  stays exactly where it is in the procedure's order — after the loop that produces
  `lastDataRow`.
- **It still compares `AvailLabel(AVAIL_OFF)`, not `AVAIL_OFF`.** Do-not-revert item 3. The
  cell holds `"Off"`; the token is `"off"`.
- **`reminderRow` keeps its `Setup Menu` nav link**, and keeps writing no machine key. The row
  is reserved and keyless on purpose: blocks are located by machine key via `FindKeyRow`, so a
  key on that row would make it look like a block. Checklist item 4.32 tests that link.

---

## 2. An option band never says which macro it belongs to

### What is wrong

Enable an action, Finalize, Create, and option fields appear under a band header like
`SPLIT COLUMN` or `BACKUP & SNAPSHOT`. Nothing on that header says which macro the fields
underneath configure. The user has just chosen from 51 captions in the action catalog and is now
looking at a band named after a *topic*, not after anything they clicked.

Some are guessable. `EMAIL` is not obviously `Email A Sheet As An Attachment` rather than a
setting shared by every export, and `GENERAL` says nothing at all.

### Where the information already is

Every field carries a `usedBy` argument — the last parameter of `Core_AddField` /
`Core_AddList` / `Core_AddTable`. It is a comma-separated list of action ids, and it is what
`FieldInUse` (line 450) already reads to decide whether a field belongs in the visible area or
in the collapsed off band. `ActionCaptionFor` (line 2620) already maps an id to its caption.

So the mapping exists; it is simply never shown.

### What it becomes

The section header gains a prefix naming the macros its fields configure, before the existing
title:

```
Split One Column Into Many  -  SPLIT COLUMN
Backup Now (Timestamped Copy), Snapshot Before a Risky Step, Save Month-End Version, Archive a Generated Output  -  BACKUP & SNAPSHOT
```

Rules, each of which exists because of something in the current code:

- **The prefix is computed from the fields that will actually render**, not from every field
  the section declares. `RenderTier` emits a section header lazily — it holds the title in
  `pending` and draws it only when the first field matching *this* tier and *this* in-use side
  appears. A prefix built from all declared fields would name macros whose fields are not on
  screen. This needs a forward scan from the section to the next section, applying the same
  `tier` and `FieldInUse` tests the loop applies.
- **Ids are de-duplicated and kept in declaration order.** Several fields in one section
  usually share a `usedBy`; `BACKUP & SNAPSHOT` has one field used by four actions and another
  used by one.
- **Each id is tested on its own, not just the field it sits on.** `FieldInUse` is an **ANY**
  test — it reports a field in use when *one* of its ids is runnable. `k_backup_folder` names
  all four backup actions, so switching three on leaves the field in use while one action is
  `Off`. Naming all four there would caption a macro the user deliberately disabled and can see
  no button for, which is wrong in the direction they cannot detect. So the per-id filter is
  separate from the per-field one. `Chains only` counts as runnable: a chain can still run the
  action, so its options still have to be answerable and its name still belongs on the band.
- **A field with an empty `usedBy` contributes nothing**, so a band with no macro of its own
  keeps its bare title with no orphaned separator. **This path is unreachable today** — all 40
  field declarations name at least one action — so it is defensive only, and no checklist item
  covers it because none could be executed. `GENERAL` is **not** an example: both `GENERAL`
  bands tie every field to specific export actions and therefore do get a prefix.
- **No truncation.** The four-caption case is long, and a shortened list that silently drops a
  macro is worse than a long header: the whole point is to tell the user which macros these
  fields reach. **Note the height mechanism, because it is not AutoFit.** `TintRow` merges B:E,
  and Excel's AutoFit does not resize a row containing merged cells — which is why `TintRow`
  computes the height itself, as `(Len(txt) \ 115) + 1` lines. The longest header the toolkit can
  currently produce is the four backup captions at ~132 characters, which that estimate puts on
  two lines. It fits, but it fits by estimate rather than by measurement, so checklist item 4.39
  exists to look at exactly that case.
- **Captions, not ids.** The user picked captions in the action catalog. `ActionCaptionFor`
  falls back to the id when a lookup fails, which is the right failure: an unresolvable id
  shows as itself rather than vanishing.

### Which bands get a prefix

Only the visible ones, which falls out of the existing code without a new condition.
`showSections` is `True` in exactly one call — line 2228, the simple + in-use pass. The
advanced pass (2231) and both off-band passes (2393, 2394) pass `False` and draw no section
headers at all. So "each visible band header" is already the only place a section header
exists.

The package banner (`SECTIONHEADER(ws, r, banner)`, line 2227) and the `ADVANCED` and off-band
headers are **not** touched: a package banner covers a whole package, not one macro, and
prefixing it with every caption in the package would be noise.

---

## 3. Help text names a tab only the author recognises

### What is wrong

The placement help says:

> for example  **TabName!B4**  places it on cell B4 of the TabName tab

`TabName` is a leftover tab name from the client workbook this toolkit grew out of. To anyone else
it is three letters with no meaning, in the one example that is supposed to teach the
`Sheet!Cell` form. It appears in four user-facing strings and seven code comments.

This is the same class of leftover the Task 12 sanitization pass was looking for, found from
the other direction — not "does this leak a client name" but "would a stranger understand
this".

### What it becomes

`TabName!B4`, the owner's suggested wording, in all four user-facing strings:

| line | context |
|---|---|
| 2751 | Chains block intro — `Placement (1A, hide, or TabName!B4)` |
| 2762 | Chains `Placement` header note |
| 3021 | Action catalog block intro — the long teaching example |
| 3047 | Action catalog `Placement` header note |

`TabName` reads as a placeholder in a way `TabName` cannot, and it survives the sentence that
explains it: *"TabName!B4 puts the button on cell B4 of your own tab named TabName"*.

The seven code comments (lines 878, 882, 3685, 5216, 5539, 5567, 5568) change too. They are not
user-facing, so they are not a correctness defect — but they are client jargon in a repo that
is about to be public, and leaving them means the next person greps `TabName` and finds a
half-finished rename.

### The scan for similar wording

Every `e.g.` / `such as` / `for example` in user-facing text was listed and read. Findings:

- **`TabName!B4`** — the defect above. Fix.
- **`C:\Close\Backups`, `C:\Close\January`** (`PkgBackup.bas:45`, `PkgExport.bas:29`,
  `PkgImport.bas:35`) — `Close` is the owner's month-end-close folder convention. Not a client
  identifier, and Task 12's plan listed these as "known and fine", but as a *teaching example*
  for a stranger `C:\Reports\January` is plainly clearer. **Change**, and flagged here as the
  one judgement call in this item — it is a one-word revert if the owner disagrees.
- **`e.g. Store, Month`** (grouping-column prompt) — "Store" is an ordinary English noun doing
  ordinary work as a grouping example. **Keep.**
- **`Ref_Master`, `Master`, `Contents`** — real sheet names the toolkit creates and reads.
  **Keep**; genericizing these would make the help wrong.
- **`e.g. mm/dd/yyyy`, `$#,##0.00`, `greater_than 1000`, `contains VIP`, `1A`** — format and
  syntax examples with no place or person in them. **Keep.**

---

## Plan

Four steps. Steps 1–3 are independent and touch disjoint line ranges of `src/core/Core.bas`
(plus three package files in step 3), so they can be written in one pass and reviewed together.

- [ ] **Step 1 — the reminder (item 1).** `src/core/Core.bas`.
  1. Rewrite `REMINDER_TEXT` (line 183). Single quotes only; no double quote may enter the
     string, or the formula literal breaks.
  2. In `RenderPlacementBlock`, change the write target from `ws.Cells(reminderRow, VAL_COL)`
     to `ws.Cells(1, 4)`. Do not move the statement — it stays after the action-row loop.
  3. Set `Font.Bold = True` and `Font.Color = RGB(150, 0, 0)` on `D1`. Do not set `WrapText`.
  4. Leave the `NavLink` at `reminderRow` alone, and leave `KEY_COL` on that row empty.
  5. Extend the comment block above the formula to record why D1 and not a wider C.

- [ ] **Step 2 — band macro references (item 2).** `src/core/Core.bas`.
  1. Add a private helper that, given `reg`, the index of a section within `reg("fields")`, the
     `tier` and `availMap` and `wantInUse`, walks forward to the next section and returns the
     de-duplicated, declaration-ordered list of `usedBy` ids of the fields that would render —
     applying the same `tier` and `FieldInUse` tests `RenderTier` applies. Empty `usedBy` is
     skipped.
  2. Add a second private helper mapping that id list to a display prefix via
     `ActionCaptionFor`, returning `""` for an empty list.
  3. In `RenderTier`, when a pending section header is drawn, prefix it. Because the helper
     needs the section's *index*, the `For Each f In reg("fields")` loop needs a position
     counter, or a `For i` over the collection — the collection is 1-based, so
     `For i = 1 To reg("fields").Count`.
  4. Only the `showSections = True` path changes. The advanced and off-band passes already
     draw no section headers; do not add one.

- [ ] **Step 3 — generic examples (item 3).** `src/core/Core.bas`, `src/packages/PkgBackup.bas`,
  `src/packages/PkgExport.bas`, `src/packages/PkgImport.bas`.
  1. `TabName` → `TabName` in the four user-facing strings and the seven comments.
  2. `C:\Close\Backups` → `C:\Reports\Backups`; `C:\Close\January` → `C:\Reports\January`.
  3. Grep `TabName` and `C:\Close` to zero across `src/`.

- [ ] **Step 4 — gates and the checklist.**
  1. `pwsh build/build.ps1` — must report `Core + 10 package(s)` with balanced Sub and Function
     counts. **Every new call site checked by hand**: neither the build nor the syntax checker
     resolves symbols, so a call to a helper that does not exist passes both gates and fails
     only at `Debug ▸ Compile` in Excel. Step 2 adds two new helpers and two new call sites.
  2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` — no diagnostics.
  3. **Checklist items 4.29–4.33 all move.** They currently describe the reminder as "the row
     directly above the `Action` / `Availability` / `Placement` headers"; it is now `D1`. 4.34
     names the `Placement` note text that step 3 rewrites. Update
     `docs/designs/2026-07-28-end-to-end-validation.md`, and add an item for the band prefix
     and one for the third sentence of the reminder.
  4. `CHANGELOG.md` under `## [0.18.0]` — this lands before that release, so it is not a new
     version, and the entries join the ones already there.
  5. Re-run the blind dry-run **only if step 3's rewording changes an item's instructions**,
     which it does for 4.34.

### Risks

- **Nothing here runs in Excel either.** Same standing caveat: the gates are a concatenator and
  a parser. The reminder's move to `D1` is the change most likely to look wrong on screen while
  passing both gates, because overflow behaviour is Excel's, not the code's — it depends on
  `E1:H1` being empty, which the code guarantees today but no test enforces.
- **Step 2's forward scan must agree with `RenderTier`'s own filter.** If the two ever disagree,
  a header names a macro whose fields are not shown, which is worse than no prefix. They are
  written next to each other and must use the same `FieldInUse` call, not a reimplementation.
- **`For Each` → `For i` in `RenderTier`** is a mechanical change to a loop that already has
  three exit paths through `pending` / `havePending`. The section-skip case (a section title
  identical to the package banner, line 2296) must keep behaving the same.
