# Design — v0.4.0 track 1, the four items from the owner's Excel gate

**Date:** 2026-07-25
**Branch:** `feat/finish-excel-track1`
**Supersedes nothing.** Extends `docs/designs/2026-07-24-finish-excel-track1-plan.md`
(Tasks 1–9 complete) with the work that came out of Task 10, the owner's manual
Excel gate. Task 11 (teardown + ship) stays blocked until all four land.

**Input brief:** `docs/designs/2026-07-25-track1-owner-feedback-handoff.md`
**Execution ledger:** `.superpowers/sdd/2026-07-24-finish-excel-track1-plan/progress.md`

---

## Decisions taken before design (owner, this session)

| # | Question | Decision |
|---|---|---|
| 1 | `PkgClean_RemoveDupes` hardcodes `Header:=xlYes` | Add a Setup config field, default Yes |
| 2 | Does `TabName!B4` flip the workbook into manual mode? | **No.** Sheet placements are their own axis, like `hide` |
| 3 | How is a placed button cleaned up on rebuild? | Sweep every sheet by shape-name prefix, then re-place |
| 4 | Finalize behaviour when the named sheet does not exist | **Block activation**, naming the action and the sheet |
| 5 | Where does the workbook land when a build finishes? | First generated page, else stay put |
| 6 | Release scope | All four items in v0.4.0, one PR, one manual gate |
| 7 | Kind for the new `clean_hasheader` field | `bool` (its contract-reserved first use), not `text` |

Decision 7 was presented with its cost (~80 lines and a §3 amendment, against
~10 lines for `text`) and approved with the rest of the design.

---

## Two facts this design rests on, verified rather than assumed

1. **No consumer mutates a registry dictionary.** Every mutation of a `reg` in
   `Core.bas` is inside `Core_DeclarePackage` / `Core_AddField` /
   `Core_AddAction` (`Core.bas:146-224`), all of which run *during*
   `TryManifest`, inside `BuildRegistry` itself, against a dictionary it has
   just created. All thirteen `For Each reg In BuildRegistry()` sites only read.
   Memoizing a shared object is therefore safe. **The implementer re-verifies
   this and stops if it does not hold** — the handoff makes this a stop
   condition, and a cache handed out by reference is exactly the shape where
   being wrong is silent.
2. **The `bool` kind is contract-reserved and unused.** §3 marks it
   "**Reserved: defined but not yet used by any package.** Encoding (`yes`/`no`
   vs `TRUE`/`FALSE`) is fixed on first use and this table updated then", and
   §10 repeats the reservation. `clean_hasheader` is that first use, so fixing
   the encoding is contract-mandated work, not scope creep.

---

## Implementation order

**1 → 2 → 4 → 3**, deliberately not the owner's priority order.

Item 4 (remove `Macros Home`) runs before item 3 (sheet placement) because item
3 must write a placed-button count into the completion message and must reason
about "zero generated pages" — and item 4 is what creates both of those. Item 3
first means writing that code twice and reviewing it twice.

Items 1 and 2 stay first as the owner asked: they are small, independent of the
other two, and they are what makes item 3's behaviour diagnosable at the gate.

---

## Item 1 — Cross-tab reporting, and the duplicates header

### 1a. Sheet-qualified range references

**The defect.** `Range.Address` omits the sheet name by design, so
`PkgClean`'s three completion messages report `M2:M12` whether the action ran on
`TabName` or on the button's own tab. This fully explains the owner's symptom
without implying anything ran on the wrong sheet.

**What the sweep of the other actions found** (the handoff asked for it):

- `PkgClean_TrimClean` and `PkgClean_FillDown` address every cell through
  `rng.Cells(r, c)`, which carries the range's own parent. They cannot act on
  the wrong sheet.
- Every `PkgImportExport` message already names its configured tab or file
  explicitly (`mDest`, `mSheet`, `tlist`, `outPath`). None is ambiguous.

So the change is three lines, and the open question — *did* it act on the wrong
sheet — is settled by the owner's next run, in one click.

**The change.** A new shared helper in Core, because CLAUDE.md rule 4 puts
shared helpers only in Core and a copy in `PkgClean` would be a standards
deviation:

```
' "TabName!M2:M12", or "'TabName Data'!M2:M12" when the sheet name needs quoting.
' Range.Address omits the sheet name by design, so every message that names a
' range must qualify it here or a cross-tab run is indistinguishable from a
' same-tab one - which is exactly what the owner hit at the v0.4.0 gate.
Public Function Core_RangeRef(rng As Range) As String
```

Quoting rule, matching Excel's own: quote unless the sheet name is entirely
`A-Z`, `a-z`, `0-9`, `_` and does not begin with a digit; double any embedded
apostrophe inside the quotes.

Call sites: `PkgClean.bas:45`, `:81`, `:112`.

### 1b. `clean_hasheader`

**The defect.** `rng.RemoveDuplicates ... Header:=xlYes` is hardcoded
(`PkgClean.bas:42`). `RemoveDuplicates` treats the range's first row as a header
and excludes it, so a user selecting `M2:M12` intending all eleven values
considered gets only `M3:M12` deduped. This is very likely the owner's
"works if there is something in M1" observation.

**New field**, declared in `PkgClean_Manifest` immediately after
`clean_keycols`:

| key | kind | tier | default | label |
|---|---|---|---|---|
| `clean_hasheader` | `bool` | simple | `yes` | Does the range you pick include a header row? |

Help text states plainly that with Yes the first row is never removed, and with
No every row in the selection is considered.

`PkgClean_RemoveDupes` becomes:

```
rng.RemoveDuplicates Columns:=ParseCols(keys), _
                     Header:=IIf(Core_CfgBool(ctx, "clean_hasheader"), xlYes, xlNo)
```

The range picker prompt drops its unconditional "(include the header row)",
which is now a lie half the time.

**This is additive, not a semantics break.** The default `yes` reproduces
today's behaviour exactly, so no v1.0.0/v1.1.0 config is reinterpreted. §10's
"changing an action's semantics is breaking" does not fire.

### 1c. First use of the `bool` kind

Fixing the encoding is contract-mandated (§3, §10). Chosen encoding: **`yes` /
`no`**, canonical, lowercase in the store.

New engine machinery, all mirroring the `availability` design that already
exists for the same problem:

- **Reader.** `Public Function Core_CfgBool(ctx As Object, ByVal key As String)
  As Boolean`, alongside `Core_Cfg` / `Core_CfgLong` / `Core_CfgList` /
  `Core_CfgTable` (`Core.bas:294-311`).
- **Parser.** A private `ParseBool(ByVal raw As String, ByRef flag As Boolean)
  As Boolean` returning False for anything unrecognised, so callers fail closed.
  Accepted, case-insensitively: `yes` / `y` / `true` / `1` for true;
  `no` / `n` / `false` / `0` for false. This mirrors §6's existing
  reference-implementation tolerance clause for `availability`, and it is
  explicitly a reference-implementation convenience — a port MAY accept only
  `yes` / `no`.
- **Blank is not unrecognised.** A blank cell means "never set" and reads as the
  field's declared default, exactly as a blank `availability` reads as `button`.
  `Core_CfgBool` resolves that default by looking the field up in the context's
  own `reg` (`ctx("reg")`), not by taking a hardcoded False and not by an
  optional caller-supplied fallback — the declared default is already on the
  manifest, and two places to state it is two places to disagree. The registry
  walk this costs is free after item 2's memoization.
- **Renderer.** The Setup field renderer gains a validated `Yes` / `No`
  dropdown for `bool` fields, mirroring `AvailDropdown` — including its
  `On Error Resume Next` wrapper, because a refused dropdown must not abort the
  build and the parser guards the value regardless.
- **Finalize blocks** on an unrecognised value, naming the field and the
  accepted values. A user who typed `ye` must not be handed a silent `no`.

**Existing `IsTrue` is deliberately not reused.** It accepts only `TRUE` / `1`
and fails *open* to False for everything else, which is right for the optional
`_Packages` enable column and wrong here.

### 1d. Contract changes for item 1

- **§3** — the `bool` row loses "Reserved" and states the fixed encoding
  (`yes` / `no`, canonical lowercase), the Excel realization (validated Yes/No
  dropdown), and that a blank reads as the field's default.
- **§8**, package `clean` — add the `clean_hasheader` field row; amend
  `clean.dupes` semantics to reference it instead of "(incl. header row)".
- **§10** — the "Reserved for first use" bullet records that v1.2.0 is the first
  use and the encoding is now fixed; the v1.2.0 bullet gains the field.
- **No version bump.** v1.2.0 is unreleased.

---

## Item 2 — Performance

**The defect.** `BuildRegistry` (`Core.bas:711`) is uncached. Every call re-runs
`CoreCandidates`, one `Application.Run "Pkg<Name>_Manifest"` per candidate, and
a `PackageEnabled` worksheet scan per candidate. It is called from
`RenderOffBand` (twice), `HasAnyTier`'s per-package loops, `ActionCaptionFor`
and `ActionTokenFor` once per chain step, `StepsInline` once per step, and
`ActionIdSet` twice inside a single `ComposeReport`. At 82 actions the owner
measured the first Setup build at several minutes.

### 2a. Memoize `BuildRegistry`

Module-level cache in the declarations section (CLAUDE.md rule: module-level
`Const`/`Dim` stays in the declarations section):

```
Private mRegCache As Collection
Private mRegCacheOk As Boolean
```

`BuildRegistry` returns `mRegCache` when `mRegCacheOk` is True, otherwise builds,
stores, sets the flag, and returns.

**Invalidation** — `InvalidateRegistry` (clears both) at the top of:

- `Core_BuildSetup`
- `Core_FinalizeSetup`
- `Core_BuildControlPanel`

and **explicitly not** in `Core_Run`, so a 20-step chain builds the registry
once. The registry depends on installed modules (fixed for the session) and on
the optional `_Packages` sheet's enable/disable values, which a user can edit
between builds — the three entry points above are every place a build begins, so
an edited `_Packages` is always picked up.

**Stop condition:** the implementer re-verifies that no caller mutates a
returned `reg` (see "Two facts", above) and stops and reports if one does.

### 2b. Collapse `AvailDropdown`

`AvailDropdown` (`Core.bas:1780`) makes two COM validation calls per action row
inside `RenderPlacementBlock`'s loop — about 164 at fixture scale. Replace with
one `Validation.Delete` + `Validation.Add` over the exact block of data rows
after the loop.

Bound it to the rows actually written (first action row through the last), never
the whole column, so it cannot leak validation onto the sentinel/key rows or
below `END_KEY`. The `On Error Resume Next` tolerance and the comment explaining
why it exists are preserved verbatim.

Nothing depends on the per-cell form: the value is guarded by `ParseAvail` and
blocked at Finalize regardless of whether the dropdown was accepted.

### 2c. Verification

Both changes are behaviour-neutral, so they are verified by the same Setup build
the owner already ran, with a clock on it. No contract change.

---

## Item 4 — Remove `Macros Home`

The owner reversed their own earlier call after seeing the feature built. That
call stands and is not re-argued. "Save the work" means preserve the code.

### 4a. Move, do not delete

`BuildHomePage`, `HomeCatalog`, `HomeLegend`, `HomeLine` (`Core.bas:3340-3523`)
and the already-orphaned `AddText` (`Core.bas:3524`, whose last caller
`RenderPageLegacy` was deleted in Task 8) move verbatim to:

```
src/attic/HomeIndex.bas.txt
```

`.txt` so `build.ps1`'s `Pkg*.bas` glob cannot pick it up and neither can its
ASCII or reserved-word gates, and so nobody mistakes it for a live module.

The file is headed with: what it was, that it **never reached a release**, the
owner's reason for withdrawing it, and what re-wiring a restore would need —
the `BuildHomePage layout, manual` call at the end of the pipeline, the
`<< Home` nav link, `IsPanelName`'s carve-out, and the §6 index clause.

### 4b. The four non-obvious consequences

**Landing place.** `BuildHomePage` was the only thing in the pipeline that
activated a sheet and selected a cell. Replacement:

```
' Activate BEFORE selecting - Range.Select on a non-active sheet raises 1004,
' which is how RenderPageLegacy's bare Range("A1").Select survived only on the
' fresh-sheet path (Worksheets.Add activates). See ledger, Task 8.
Private Sub LandOnFirstPage(layout As Object)
```

Activates the lowest-numbered generated page and selects A1. Does nothing when
there are no generated pages at all — reachable now that every button can be
hidden, off, or placed on a user's own tab. Also does nothing if the target
sheet is not `xlSheetVisible`: a user may have hidden it, and activating a
hidden sheet raises.

**Dead nav link.** `RenderNavRow` (`Core.bas:2873`) drops its unconditional
`<< Home` link. `< Previous` and `Next >` still work and shift left into cells
B and D. `NavLink` itself is unchanged.

**Orphaned index.** `DeletePanels` deliberately spares a sheet named
`Macros Home` — `IsPanelName` requires a numeric tail. With the feature gone, an
index left behind by a pre-removal build would be **never deleted and never
rebuilt**, full of dead links. New `DeleteHomePanel`, called unconditionally
from `Core_BuildControlPanel`, mirroring `DeleteLegacyPanel`'s retirement of the
pre-v0.4.0 single `Macros` tab.

`IsPanelName` keeps its numeric-tail rule unchanged, but its `LOAD-BEARING`
comment block (`Core.bas:3552-3559`) is **rewritten**: it currently instructs the
next reader to preserve a carve-out that has become vestigial, and a stale
load-bearing comment is worse than none.

**Contract §6 mandates an index.** It currently reads "the engine **always**
generates an index surface, even for a single page". Left alone, the Apps Script
port gets built to a spec the reference implementation violates. Changes:

- The index becomes **optional** — a platform MAY provide one; when it does, the
  listed contents (page list, button catalog, admin link, legend) still describe
  what it contains.
- The **previous/next navigation requirement stays**, unconditionally.
- **§7 keeps `Macros Home` reserved.** A future restore wants the name, and
  reserving an unused one costs nothing.
- **§10's v1.2.0 bullet** records the withdrawal and the owner's reason.
- **No version bump.** v1.2.0 is unreleased, and this is engine presentation
  under §10's two-axis rule.

### 4c. Completion message

`Core_BuildControlPanel`'s success dialog drops "plus the '<name>' index", and
its "Built N macro page(s)" wording must stay true when N is zero.

Item 3 later adds the placed-button count to this same string. Ordering item 4
first is what keeps that a single edit to a settled sentence rather than two
edits to a moving one.

### 4d. Task 10 checklist re-cut

Deleted: every item verifying the page list, the button catalog, the legend, the
`<< Home` link, Home's own links, and the "no duplicate Home tab" case.

**Not dropped — re-aimed.** The AutoFit risk was Home's un-floored catalog. With
that gone, the primary AutoFit check becomes the card-page category header at
`Macros 1!F3`: `Clean & Standardize (cont. 2)` is 29 characters, 11pt bold, in a
30-character column, with no height floor — the most borderline wrap in the
build.

Added: sheet-qualified messages, the duplicates header field, a timed Setup
build, and the anchor-placement checks from item 3.

---

## Item 3 — `Sheet!Cell` placement

The owner's original intent, unbuilt: *"I specifically wanted the option to add a
macro to a specific page and cell, not just organize the buttons on new pages."*

### 3a. Grammar

A placement value in the Action catalog's "Page code" column (and the Chains
block's) becomes one of four forms:

| Form | Example | Meaning |
|---|---|---|
| blank | | automatic layout, or page 1's Unassigned in manual mode |
| `hide` | `hide` | button omitted entirely |
| `<digits><letters>` | `1A`, `12AB` | generated page + column (unchanged) |
| `<sheet>!<cell>` | `TabName!B4` | **new** — anchored to one cell on a user's own sheet |

**The discriminator is `!`**, which Excel forbids in a sheet name, so the two
placement forms can never be confused — including for a sheet legitimately named
`1A`.

Accepted anchor spellings: `TabName!B4`, `'TabName Data'!B4`, `TabName Data!B4` (unquoted
with spaces), `TabName!$B$4`. The sheet part may be quoted; an embedded apostrophe is
doubled inside the quotes. The cell part is a single A1 reference — optional `$`,
1–3 column letters, 1–7 row digits — that must resolve to a real cell. A colon
(a multi-cell range) is rejected.

The column header on the Setup tab is relabelled from "Page code" to
**"Placement"**, and the block's help text gains the new form. Both blocks
(actions and chains) use the same grammar, as they already do — so a **chain**
can be anchored to a user's sheet exactly like an action, and
`PlaceAnchoredButtons` receives both from the same collection.

**The relabel has a string ripple, and it is part of the work, not a side
effect.** "Page code" appears in the Action catalog header, the Chains block
header, both blocks' help text, `RenderComposeSection`'s intro,
`GateConfigMigration`'s recovery messages ("correct the page code(s) named
above", "every other saved page code is kept"), `ComposeReport`'s bad-code and
layout lines, and contract §6 prose. Every one of them is updated in the same
task. A half-relabelled surface is worse than no relabel: the ledger records
four separate rounds spent on user-facing strings that had drifted from the
behaviour they described.

**Not load-bearing:** unlike the `Availability` header — whose caption is tied
to `AVAIL_HEADER` (`Core.bas:60`) because `ConfigSchemaStatus`'s tamper test
compares against it at `Core.bas:1954` — no check anywhere reads the placement
column's caption. Verified by grep, not assumed: `"Page code"` appears only as a
written value, at `Core.bas:1656` and `Core.bas:1810`. The relabel cannot
silently disarm a migration guard.

**But it is written from two independent literals**, which is the exact coupling
the ledger closed for `AVAIL_HEADER` (Task 4 minor, folded into Task 5). Since
this task is editing both anyway, the new caption is introduced as a
`PLACE_HEADER` Const in the declarations section and both sites use it.

### 3b. Parsing

**`ParseCode` is not touched.** Its four call sites keep their page-code meaning,
and its hard-won hardening (digit cap, 3-letter cap, `VSPLIT_MULT` guard) stays
exactly as the ledger left it.

New sibling:

```
' "TabName!B4" / "'TabName Data'!$B$4" -> sheetName "TabName Data", cellRef "B4".
' Returns False on anything malformed. The "!" discriminator is safe because
' Excel forbids "!" in a sheet name, so an anchor can never be confused with a
' page code - not even for a sheet named "1A".
Private Function ParseAnchor(ByVal code As String, ByRef sheetName As String, _
                             ByRef cellRef As String) As Boolean
```

and a cheap classifier so no site has to re-derive the rule:

```
' "" | "hide" | "page" | "anchor" - the single place a placement value's form
' is decided. Every reader goes through this rather than testing for "!" itself.
Private Function CodeKind(ByVal code As String) As String
```

**All four existing `ParseCode` call sites gain an anchor branch:**

1. **`ConfigSchemaStatus`** (`Core.bas:2023`) — an anchor must not be classified
   `invalid`. Missing this is not cosmetic: the `invalid` path *clears* the
   offending rows, so every anchor would be silently deleted on a marker
   recovery.
2. **`ComposeReport`, chain codes** (`Core.bas:2130`) — validate anchors.
3. **`ComposeReport`, placement codes** (`Core.bas:2148`) — validate anchors, plus
   the four new structural checks below.
4. **`BuildManualLayout`** (`Core.bas:2923`) — anchors never reach it, because
   they are partitioned out one layer above. A defensive branch is *not* added:
   the ledger records that a per-filer duplicate of a funnel-level filter is
   exactly what produced the `hide` regression in Task 7.

### 3c. Mode: anchors are their own axis

`AnyPlacementCode` (`Core.bas:2303`) already ignores blanks and `hide`. It now
also ignores anchors, so typing `TabName!B4` on one action does **not** flip the
workbook into manual mode and dump the other 81 buttons into page 1's
Unassigned column.

This is the same carve-out shape the ledger records three fix rounds defending
for `hide` (Task 7, fix round 2 fix 5, and the `AnyCode` / `AnyPlacementCode`
split). The contract clause is written to match, so a port cannot reintroduce
the bug by reading §6 literally — which is precisely what the ledger notes
almost happened with `hide`.

### 3d. Partitioning — where anchored buttons leave the page pipeline

`GatherButtons` stays the **single funnel** and keeps dropping `hide` there. The
ledger is emphatic: *"Feed both filers ONLY this collection - nothing assembled
another way, or hide silently stops working."*

The split happens in `Core_BuildControlPanel`, strictly downstream of that
funnel, so `hide` has already been applied to both halves:

```
Set codeMap  = ReadCodeMap(CfgSheet())      ' read ONCE - unchanged
manual       = AnyPlacementCode(codeMap)    ' now also ignores anchors
Set buttons  = GatherButtons(codeMap)       ' single funnel - "hide" dropped here
SplitAnchored buttons, codeMap, paged, anchored   ' NEW - downstream of the funnel
Set layout   = CreateObject("Scripting.Dictionary")
DeletePanels
DeleteLegacyPanel
DeleteHomePanel                             ' NEW (item 4)
SweepPlacedButtons                          ' NEW - after the deletes, fewer sheets to walk
If manual Then BuildManualLayout paged, codeMap, layout Else BuildAutoLayout paged, layout
RenderAllPages layout
BuildAdminPanel
PlaceAnchoredButtons anchored, codeMap      ' NEW
LandOnFirstPage layout                      ' NEW (item 4) - replaces BuildHomePage's activation
```

Only `paged` reaches either filer. `anchored` never enters `layout`, so it never
affects page counts, column caps, category packing, or continuation pages.

### 3e. Rendering — a genuinely separate code path

The card-grid renderer owns its sheet outright: `Cells.Clear`, `Cells.UnMerge`,
`Hyperlinks.Delete`, `Cells.RowHeight`, six column widths, tab colour, interior
fill (`EnsurePanelSheet`, `Core.bas:2602`). **A user's own sheet must never be
treated that way**, so this is a different procedure, not a parameter on that
one.

```
' Anchor one button per placement onto a sheet the engine otherwise leaves
' COMPLETELY alone. Writes to no cell: no Clear, no UnMerge, no column width,
' no row height, no value, no format, no hyperlink. The only mutation to a
' user's sheet is adding (and, via SweepPlacedButtons, removing) a shape whose
' name carries this engine's reserved prefix.
Private Sub PlaceAnchoredButtons(anchored As Collection, codeMap As Object)
```

Per button: resolve the sheet and cell, then

```
AddBtn ws, PLACED_PREFIX & SafeShapeName(id), caption, onAction, _
       target.Left, target.Top, ANCHOR_BTN_W, BTN_H
```

with `Placement = xlMove` — **not** the card grid's `xlMoveAndSize`. On a
generated page the engine controls the column width and wants the button to
track it; on a user's sheet, narrowing a column must not squash their button.
`AddBtn` gains an optional trailing `place` parameter defaulting to
`xlMoveAndSize`, so all existing call sites are unchanged.

Width is the fixed constant `ANCHOR_BTN_W` (150pt), **not** the target cell's
width — a user's column may be three characters wide. Height reuses `BTN_H`.
Tunable at the owner's gate if 150pt reads wrong next to real data.

`SafeShapeName` maps any character outside `A-Za-z0-9_` to `_`, so
`clean.dupes` becomes `SFPlaced_clean_dupes`. A theoretical collision between
two ids differing only by that substitution is harmless: placement is
sweep-then-add and never looks a button up by name.

### 3f. Rebuild cleanup

```
' Delete every button this engine has ever placed on a user's sheet, across the
' whole workbook, before re-placing from the CURRENT codes. Self-healing: a
' moved cell, a cleared code, a renamed sheet, and an orphan left by a build
' that died halfway are all fixed by the same pass. DeletePanels cannot help
' here - it only deletes generated "Macros N" sheets, and an anchored button
' lives on a sheet the engine must otherwise never touch.
Private Sub SweepPlacedButtons()
```

Walks every worksheet, iterates `Shapes` **backwards** (the established idiom in
`EnsurePanelSheet`), and deletes any shape whose name starts with
`PLACED_PREFIX` (`"SFPlaced_"`). It runs on every build, including builds with
zero anchors — that is exactly how clearing a code removes its button.

The prefix is reserved in contract §7. Nothing else in the workbook may use it.

### 3g. Validation — what blocks Finalize

Per decision 4, all four are **structural** (they block activation), consistent
with how a bad page code is already treated:

| Condition | Message |
|---|---|
| Sheet does not exist | `NO SUCH SHEET for <id>: 'MICC!B4' - this workbook has no tab named "MICC".` |
| Engine-owned target | `RESERVED SHEET for <id>: 'Macros 1!B4' - the engine rebuilds that tab on every build; pick one of your own.` |
| Target cell already claimed | `DUPLICATE TARGET: <id A> and <id B> both point at TabName!B4.` |
| Malformed anchor | `BAD CODE for <id>: 'TabName!B4:C9'` (reuses the existing bad-code message) |

Engine-owned names: `Setup`, `Setup Menu`, `Macros`, any `Macros <n>`
(`IsPanelName`), `Macros Home`, `_Config`, `_Packages` — a new
`IsEngineOwnedSheet` helper, so the list lives in exactly one place.

Duplicate detection normalises through the resolved `Range.Address(True, True)`
keyed by the resolved sheet name, so `B4` and `$B$4` collide correctly. It runs
only after the sheet-exists check has passed.

`ComposeReport` also gains a **soft** (informational) list of buttons placed on
the user's own tabs, alongside the existing Unassigned and Hidden lists, in the
same `caption [id]` format both of those already use.

### 3h. Contract changes for item 3

**§6**, grouping/layout, gains:

- the anchor form in the placement grammar, with the `!` discriminator and the
  accepted spellings;
- **normative:** an anchor is **not** a placement code for the mode decision —
  the same carve-out `hide` already has, and stated in both directions so a port
  reading it literally cannot reintroduce the bug;
- **normative:** the engine MUST NOT modify the target surface beyond adding and
  removing its own control — no clearing, no resizing, no writing;
- **normative:** an implementation MUST remove its previously-placed controls on
  rebuild, so a moved or cleared placement cannot leave a live orphan;
- the four blocking conditions above.

**§7** reserves the `SFPlaced_` shape-name prefix (reference implementation) and
notes a platform reserves its own equivalent.

**§10's v1.2.0 bullet** records the anchor form. Anchors add no stored `key`,
`id`, or `kind`; they extend an existing value's grammar with a form no valid
v1.1.0 config could contain (a code with `!` was previously invalid and blocked
Finalize), so no existing config is reinterpreted. **Additive, no version bump.**

---

## Verification

Unchanged and non-negotiable: **there is no headless Excel, no VBA compiler on
the agent side, and no test suite.** TDD is formally off for this repo and is not
proposed. `pwsh build/build.ps1` and `node scripts/vba-syntax-check.mjs` are
structural only — they cannot resolve a call target, see a wrong argument count,
or (before `c404349`) catch a reserved-word identifier.

This session's own history is the proof: a `Dim ... any As Boolean` compile error
passed four review rounds and both automated gates, and only Excel found it.

- After each of the four items: `pwsh build/build.ps1`, then the same for the
  fixture bundle, then `scripts/vba-syntax-check.mjs` on both.
- `node scripts/code-map.mjs` before the PR — it was 23 source commits stale at
  the start of this session.
- **The owner's manual import of `dist/MacrosBundle.fixture.bas` is the gate.**
  The re-cut Task 10 checklist (4d) is the coverage.

Fixture build command (the comma form works from any shell since Task 9):

```
pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
```

---

## Explicitly out of scope

Deferred to the post-ship watch list, unchanged by this design: the five
procedures that walk the registry to find one action by id (one `FindAction(id)`
collapses four); manual mode's missing category concept at catalog scale;
`HasAnyTier` / `HasAdvanced` being named backwards; `"hide"` as five independent
literal comparisons where one `Const` would do; and the ~31 deferred minors in
the ledger, a third of which were already stale when triaged.

## Do not revert

The eight items under "Do NOT revert" in
`docs/designs/2026-07-25-track1-owner-feedback-handoff.md` are earned across nine
tasks and roughly twenty review rounds. Three bear directly on this design and
are called out where they apply: the `hide` filter lives in `GatherButtons` and
nowhere else (3d); nothing on the card grid may be merged (unchanged here);
and `build.ps1`'s reserved-word list stays conservative and empirical — add a
word only when Excel rejects one.
