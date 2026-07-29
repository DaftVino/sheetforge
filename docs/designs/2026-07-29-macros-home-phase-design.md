# Macros Home, the eight Later items, and finishing the Setup surface — design

**What this is.** The design for everything the owner identified as missing on 2026-07-29,
after track 2 was declared complete against its plan of record rather than against
`2026-07-24-finish-excel-design.md`. Three gaps from that correction, plus six further items
the owner raised the same day while reading the built Setup tab.

Written 2026-07-29 on `feat/w6-workbook`, which carries the whole of track 2 (43 actions, ten
packages, `VERSION` 0.17.0, contract v2.9.0). **This phase does not start until PR #28 has
merged `feat/w6-workbook` into `main`.** `main` is stale at v0.11.0 with three packages;
branching from it before then silently discards 28 actions.

Ending state: **51 actions**, ten packages, `VERSION` 0.18.0, contract **v2.10.0**,
`SCHEMA_VER` unchanged at `"v2.0"`.

---

## 1. Scope

| # | Item | Source |
|---|---|---|
| 1 | **Macros Home** — the withdrawn index, restored | design §2.2; owner 2026-07-29 |
| 2 | **The 8 catalog items marked *Later*** | `docs/macro-catalog-plug-and-play.md` |
| 3 | **A Finalize reminder** on the Setup tab | owner 2026-07-29, unspecified anywhere until now |
| 4 | **Explanatory notes** on four Setup headers | owner 2026-07-29 |
| 5 | **Drop the `1. `–`4. ` numbering** from the ImportExport captions | owner 2026-07-29 |
| 6 | **Chain steps become comma-separated action ids** | owner 2026-07-29 |
| 7 | **`Complete Macro Setup`** — a button that removes the configuration tabs | owner 2026-07-29 |
| 8 | **`README.md` and the surrounding docs** brought up to date, last | owner 2026-07-29 |
| 9 | **A non-technical setup guide**, web and PDF, at a 6th-grade reading level | owner 2026-07-29 |
| 10 | **Sanitize, make the repo public, ship a versioned release** | owner 2026-07-29 |

### Explicitly out of scope

- **`PkgFixture`** (design §2.5). The owner does not want it. Do not build it and do not
  re-propose it. The consequence is named honestly in §9: layout at 51 actions has no proof
  other than the owner looking at it.
- **A worked example per action** — deferred to issue #30. `Macros Home` ships with
  descriptions only.
- **Finalize auto-running Setup** — deferred to issue #29.
- **Stripping code for Off actions** — deferred to issue #31, which also records that the
  run-time version of it is forbidden here (VBIDE trust, `CLAUDE.md` rule 3).

---

## 2. The eight Later items

**No new package and no new domain.** All four target packages already exist, so
`CoreCandidates()` needs no edit and the build still reports `Core + 10 package(s)`. All four
domains are already `active` in contract §7's vocabulary table, so the table is untouched and
the build gate cannot trip on an unknown domain. This removes the single largest mechanical
risk the phase could have carried.

| Catalog line | Id | Package |
|---|---|---|
| 48 — Split one column into many (Text to Columns) | `clean.split_column` | `PkgClean` |
| 79 — Change case (UPPER / lower / Proper) | `clean.change_case` | `PkgClean` |
| 139 — Highlight cells with comments / formulas / blanks | `format.highlight_special` | `PkgFormat` |
| 161 — Apply a consistent style to all charts | `format.chart_style` | `PkgFormat` |
| 183 — Save each filtered group as its own file | `export.split_by_value` | `PkgExport` |
| 186 — Email a sheet / range as an attachment | `export.email` | `PkgExport` |
| 208 — Sort worksheets alphabetically | `workbook.sort_sheets` | `PkgWorkbook` |
| 214 — Add a "Back to TOC" link on every sheet | `workbook.toc_links` | `PkgWorkbook` |

Every id is two-segment, lowercase, underscore-separated, and starts each segment with a
letter. 43 + 8 = **51**.

### 2.1 `chart_style` is a `format` id, not a `summarize` one

§7 defines `summarize` as *"Aggregating: pivots, recalculation, summary sheets, charts"* and
`format` as *"Presentation only: number formats, styles, print layout. Changes no value."*
The catalog files this item under §6 Summarize & Analyze, so the obvious reading is
`summarize.chart_style`.

It is filed under `format` anyway, because restyling an existing chart changes no value and
aggregates nothing — it is the definition of `format` almost word for word. `summarize`'s
mention of charts covers *building* them from aggregates. Recording the reasoning here so the
next reader does not re-open it on the strength of the catalog's section heading.

### 2.2 `export.email` — the Outlook question the catalog raised

The catalog says this one *"needs a defined behaviour when Outlook is absent or blocked by
policy"*. Settled by the owner, 2026-07-29:

1. **Export the file first.** The action writes its attachment to disk before it touches
   Outlook.
2. **Late-bind Outlook, open a draft, and stop.** `CreateObject("Outlook.Application")`,
   a new mail item with the file attached and recipient/subject pre-filled from Setup,
   `.Display` — never `.Send`. The user presses Send.
3. **When Outlook cannot be reached**, refuse through `Core_Refuse` with a message naming
   the file already written, so the user attaches it by hand. The export is not rolled back;
   a policy block is not a reason to destroy work already done.

A macro that sends mail unattended is one wrong range away from mailing the wrong numbers to
a client. That is why step 2 stops at a draft, and it is not a preference to be tuned later.

### 2.3 `workbook.toc_links` — writing into sheets the user owns

No existing action writes into a user's own sheets, which is why this one needed a decision.
Settled by the owner, 2026-07-29:

- The link goes into **one cell per sheet**, whose address is a `text`-kind config field
  defaulting to `A1`.
- **A sheet whose target cell already has content is skipped**, not overwritten, and every
  skipped sheet is reported by name with what was found in the cell.
- Sheets the engine owns are excluded, as is the TOC sheet itself.

Rejected: inserting a new row 1 on every sheet. It overwrites nothing because it makes room
first, but it shifts every row on every sheet, breaking absolute references, named ranges and
print areas across the whole workbook with no one-click undo. Also rejected: a floating shape
per sheet, which touches no cell but needs its own name prefix and its own sweep — the
engine's `SFPlaced_` prefix is deleted on every rebuild, so anything reusing it would vanish
at the next Finalize.

---

## 3. Macros Home

The withdrawn implementation is preserved at `src/attic/HomeIndex.bas.txt` (295 lines, six
procedures). Its header enumerates the five wiring points a restore must redo, and they are
still accurate. This is a restore-and-extend, not a fresh build.

### 3.1 What it contains

Per the owner's requirement — it must explain **what each macro is for**, not merely index
links, which is more than design §2.2's card-per-category-with-a-count:

- toolkit title and a legend explaining the placement code and the availability states;
- a **page summary** — one row per generated page, with its button count and a link;
- an **action catalog** — **all 51 actions**, grouped by display category, each showing its
  caption, its full unabbreviated description, its current availability state
  (`Button` / `Chains only` / `Off`), and a link to its page when it has one;
- a link to `Setup Menu`.

### 3.2 All 51, including the Off ones — and why that is a contract change

Contract §6 currently scopes the catalog to *"every action whose `availability` placed it as
a button — an action that is `off` or `chains-only` has no button and no page, so it does not
appear in this catalog"*.

That cannot stand, because **every action arrives Off** (§6, v2.1.0; `AvailFor` returns
`AVAIL_OFF` when `_Config` is absent, `Core.bas:2906`). A conforming catalog on a fresh import
would be empty — at exactly the moment someone most needs to read what the toolkit does. §6 is
rewritten so the catalog covers every registered action and shows its state.

### 3.3 The five wiring points

1. **`BuildHomePage` runs last and activates**, replacing `LandOnFirstPage` as the build's
   landing place. This incidentally fixes the everything-is-Off case, which today has no page
   to land on at all.
2. **The nav row shifts back** — `<< Home` returns to B1, `< Previous` to D, `Next >` to F.
   **The Home link is not guarded with `Core_SheetExists`.** `Core.bas:4348` records that a
   guard was tried and, because Home builds last, it rendered dead bold text instead of links
   on every page of a first build, self-healing only on the second. Do not re-add it.
3. **`DeleteHomePanel` is removed entirely**, not bypassed. It exists only to stop a stale
   index surviving a rebuild; once the index is rebuilt every time, it deletes the thing that
   was just made. `EnsurePanelSheet` reuses and clears, which covers the stale case.
4. **`IsPanelName`'s numeric-tail carve-out is unchanged** — it is what lets `DeletePanels`
   spare `Macros Home`, and it is already correct for a restore.
5. **Contract §6 and §7** — §3.2 above, plus §7's *"deletes `Macros Home` unconditionally with
   alerts suppressed on every rebuild"* clause, which describes the withdrawal and is now
   false.

`PageButtonCount` comes back with `BuildHomePage`; it was removed because that was its only
caller.

---

## 4. The Setup surface

### 4.1 Drop the numbering (item 5)

`1.  Clear Input Tabs`, `2.  Import Source Data`, `3.  Update Ref_Master`,
`4.  Export Dashboard (.xlsx)` — the numbering is stored in the caption literals at
`PkgImportExport.bas:80-83`, not generated by any renderer, so this is a four-string package
edit.

It is not only a package edit. Those captions are also the button faces, and they are frozen
in contract §8's registry table. Swept in the same commit:

- `spec/contract.md` §8 — four title cells;
- `docs/designs/2026-07-28-end-to-end-validation.md` lines 67 and 74, which name
  `2.  Import Source Data` and `3.  Update Ref_Master` in live checklist items.

Stored chains are unaffected: `ParseStepTokens` reads only the bracketed id, so a caption
change cannot break a chain — and after item 6 below the captions are not in chain tokens at
all.

### 4.2 Chain steps become comma-separated ids (item 6)

Today the Steps cell holds `Label [action.id]` tokens, one per line, in a single wrapped cell
per chain row. The owner's objection is that this is awkward to paste and needlessly long.

- The header becomes **`Steps (action ids, separated by commas)`**.
- `StepsToTokens` writes **bare ids joined by `, `** — `clean.dupes, clean.trim`.
- `ParseStepTokens` gains `,` beside the `;` and newline it already accepts, **and keeps
  reading the bracketed form**. That tolerance already exists in the parser and costs nothing
  to keep; dropping it would block Finalize on any workbook already carrying chains until the
  user retyped them by hand.
- `RenderActionTokens` — the read-only copy-source list in column G — lists the bare id with
  the caption beside it in H, so the id stays the thing you copy.

**This is not a contract break.** §2 says the steps are *"an id-list joined by a delimiter that
cannot occur in an id (reference: newline)"*, and `contract.md:13-17` marks anything labelled
*reference* as one platform's realization rather than normative. A comma cannot occur in an id.
The reference note moves to `(reference: comma)`.

**A previous decision is being reversed deliberately.**
`2026-07-24-compose-your-buttons-design.md:70` chose the delimiter as *"newline or `;` — never
comma, since captions contain commas"*. That reasoning was sound and dies with its premise:
captions are no longer in the token.

### 4.3 Explanatory notes on four headers (item 4)

**No comment API exists anywhere in the repo** — zero hits for `AddComment`, `.Comment`,
`Validation.InputMessage`. This needs a new Core helper.

Four headers get an Excel note, because each is a value the user must type with nothing on the
sheet saying what is accepted:

| Header | Where | Why |
|---|---|---|
| `Availability` | Action Catalog | Has a dropdown, so the values are discoverable, but not their meaning — and `Off` is the default for everything on a fresh import, which nothing says. |
| `Placement` | Action Catalog **and** Chains | Four accepted forms — blank, a page code like `1A`, `hide`, an anchor like `MIC!B4` — none of them written down at the point of entry. |
| `Steps` | Chains | The most-changed field on the sheet after §4.2. |
| `Column types` | `k_import` | `T,N,D,DT,B,G`, comma-separated. **There is no validation** — `Core_Coerce`'s `Case Else` treats an unrecognised letter as auto-detect, so a typo is invisible today. |

**`Column types` forces the mechanism.** Availability, Placement and Steps are Core literals,
so Core can annotate them directly. `Column types` is not: it is package-supplied table
metadata, a pipe-delimited header string at `PkgImportExport.bas:51`. Core cannot annotate it
without knowing what packages exist, which is the layering rule. So `Core_AddTable` gains an
**optional pipe-delimited notes argument paralleling its headers**, and any package can
annotate its own columns. Optional, so no existing declaration changes shape.

**This step is built last** — owner's instruction, so the note text describes the finished
sheet rather than being written against a moving target and revised.

Assumption to confirm at run time: legacy `Range.AddComment` notes survive the Setup tab's
delete-and-rebuild cycle intact, and the note anchors correctly on an unmerged header cell
whose row is later `AutoFit`. The row above every header row is merged B:E, which is the
practical constraint on anything floating above it.

### 4.4 The Finalize reminder (item 3)

A formula-driven line above the Action Catalog block, plus a hyperlink to `Setup Menu`, firing
whenever any action's availability is not `Off`. Text instructs the user to click **Finalize
Setup Tab** and then **Create Setup Tab**, because turning an action on reveals option fields
that only appear after the config is activated and the sheet re-rendered — the chicken-and-egg
of design §2.3.

**The trade, recorded once so it is a decision and not an oversight.** The condition stays true
forever once anything is configured, so after the first Finalize the line never turns off. The
alternative — comparing the Setup sheet's availability column against the activated copy in
`_Config`, so the line means *"you have changes that are not live yet"* — was offered and
declined by the owner on 2026-07-29 in favour of the simpler formula. Not to be re-opened
inside this phase.

Implementation constraint: the reminder row carries **no machine key in column A**. Blocks are
located by key through `FindKeyRow`, never by row offset, so a keyless row inserted above the
catalog cannot shift anything's parse. `MarkerVerdict`'s tamper test reads the `Availability`
header string back — that caption must not be reworded.

---

## 5. `Complete Macro Setup` (item 7)

A third button on `Setup Menu`, placed about five rows (~75pt) below `Finalize Setup Tab`
rather than the 45pt gap between the existing two — the separation is deliberate, because this
one is not a button to click by accident. Entry point `Core_CompleteSetup`, a public no-argument
`Sub` in the same shape as `Core_BuildSetup`.

Intent, in the owner's words: the user now has all the macros they want, set up the way they
want, and removes the clutter.

### 5.1 What it deletes, and what it must not

**Deletes:** `Setup`, `Setup Menu`, `Macros Home`.

**Must not delete `_Config`.** It is read on every single action run —
`IsRunnable(ReadAvailMap(CfgSheet()), actionId)` at `Core.bas:1742` — chains resolve from it at
`Core.bas:1884`, and every package setting a user ever typed (folder paths, input tab names,
column types, the TOC sheet name, protection passwords) is read from it at
`Core.bas:549-563`. Deleting it does not tidy the workbook; it stops all 51 buttons working.
It is very hidden and is not clutter.

**Does not touch `_Packages`.** `PackageEnabled` (`Core.bas:1629`) treats an **absent**
`_Packages` sheet as *every package enabled*, so deleting one that lists a package as `FALSE`
silently turns that package back on. Left alone.

**Strips the `<< Home` link cell** from B1 of every `Macros <n>` page in the same pass, so no
page is left pointing at a tab that no longer exists. One cell per page — not a re-render.
Re-rendering the whole layout on a workbook the user has just declared finished risks the
thing they came to keep.

### 5.2 Guards

- **Refuses when no `Macros <n>` page exists.** Otherwise it would strip every engine surface
  from a workbook that has no buttons either, leaving nothing at all and no on-sheet way back.
- **Refuses when `_Config` is absent** — that means Finalize has never run, so there is
  nothing to complete.
- **Warns about unfinalized changes in the confirmation** and proceeds on Yes. The
  confirmation names every tab that will be deleted, states that `_Config` and the `Macros`
  pages survive, states that Setup can be rebuilt with `Alt+F8 ▸ Core_BuildSetup`, and tells
  the user to click No and Finalize first if they have edited Setup since their last Finalize.
  Detecting the pending edit automatically was offered and declined by the owner on
  2026-07-29.

### 5.3 The assumption this rests on

`Core_CompleteSetup` deletes `Setup Menu`, which is the sheet hosting the button that invoked
it. The handler lives in a standard module, so the code survives the sheet, and
`Application.DisplayAlerts = False` suppresses the confirmation prompt. This is believed to
work and is **marked as an assumption**, per `CLAUDE.md` rule 7 — nothing in this codebase
exercises it today. `Setup Menu` is deleted **last** so that a failure leaves the most
recoverable state. If it turns out Excel refuses, the fallback is `Application.OnTime` to
defer the final delete past the end of the handler.

After the deletions it activates the first surviving `Macros <n>` page, since the landing place
built in §3.3 is one of the tabs it just removed.

---

## 6. Contract v2.10.0

| Section | Change |
|---|---|
| §2 | Chain steps delimiter reference: newline → comma. Non-normative note only. |
| §6 | The index-surface withdrawal text rewritten — the reference implementation generates one again. The catalog clause rewritten to cover every registered action with its state, replacing the button-only scope (§3.2). |
| §7 | The *"deletes `Macros Home` unconditionally with alerts suppressed on every rebuild"* clause replaced; the name stays reserved but is now generated rather than swept. |
| §8 | Eight new action ids with their semantics; four retitled `importexport` captions; the new config fields for `export.email`, `workbook.toc_links` and `export.split_by_value`; the optional table-notes argument. |
| §10 | The change record, including why *"at the scale a real deployment runs — a handful of actions and one or two chains — an index is clutter"* stopped being true. It was written against seven actions. At 51 it no longer holds, and saying so is the honest way to reverse it. |

`SCHEMA_VER` stays `"v2.0"`. Action ids remain immutable; nothing existing is renamed. Captions
are display strings, not ids, so retitling four of them is additive.

---

## 7. Step order

Each step is planned, reviewed before any code is written, and reviewed again after.

1. **Contract v2.10.0** — ids frozen before any code can reference them; §6, §7, §2, §8, §10.
2. **The eight actions** — `PkgClean` +2, `PkgFormat` +2, `PkgExport` +2, `PkgWorkbook` +2.
3. **Chain steps → comma-separated ids** (§4.2).
4. **Drop the caption numbering** (§4.1), with the §8 and checklist sweep.
5. **Macros Home** (§3) — after the actions exist, so it is first rendered at full 51 scale.
6. **The Finalize reminder** (§4.4).
7. **`Complete Macro Setup`** (§5), and `Core_SpeedOn` / `Core_SpeedOff` become `Private`
   (§11.3) so Alt + F8 lists only the two entry points a user is meant to run.
8. **Header notes** (§4.3) — **the last code step**, per the owner, so the note text describes
   the finished sheet rather than being written against a moving target and revised.
9. **Validation checklist append + blind dry-run** (§8).
10. **Documentation sweep** — `README.md` and every doc the phase invalidated, brought up to
    date against the finished tree. **Owner's instruction, 2026-07-29: give this to a codex
    agent at medium effort.** It runs last for the same reason the header notes do — the docs
    describe the finished thing, so writing them earlier means writing them twice. Reviewed
    and verified here, not taken on trust.

    Known targets, to be confirmed by the agent rather than assumed complete: `README.md`
    (action count, package list, what the workbook looks like after import),
    `docs/architecture.md`, `docs/macro-catalog-plug-and-play.md` (the eight *Later* markers
    become built ids), `docs/quick-ref-workflow.md`, `docs/code-map.md` (regenerate with
    `node scripts/code-map.mjs`, never by hand), and `CHANGELOG.md` + `VERSION` at 0.18.0.

11. **The non-technical setup guide** (§11) — after the docs sweep, because it names buttons
    and keystrokes that must already be final.
12. **Sanitize, publish, release** (§12) — last, after everything else is done.

Then the PR.

---

## 8. Verification

TDD is off; VBA has no headless runner and Excel cannot be driven on CI. Do not propose a test
suite. Verification is:

- `pwsh build/build.ps1` — exits 0, `Core + 10 package(s)`, balanced `Sub` and `Function`
  counts.
- `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` — no diagnostics.
- `grep -ohE 'Core_AddAction reg, "[a-zA-Z_.0-9]+"' src/packages/*.bas | wc -l` → **51**.
- `grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas` → **1**.
- **Deliberate injections** for anything meant to fail — a malformed id, a duplicate id, an
  unprefixed package private — confirmed caught, reverted, recorded in the PR body.
- **Run-time-only claims append to `docs/designs/2026-07-28-end-to-end-validation.md`**. No
  per-wave gate document. New items for `Macros Home`, the Finalize reminder,
  `Complete Macro Setup`, the header notes, and each of the eight actions.
- **The checklist is dry-run blind** before it reaches the owner — a subagent with the built
  bundle and no access to `src/`. The last such pass found five buttons the checklist named
  that do not exist; every button caption this one names must be greppable in
  `dist/MacrosBundle.bas`.

---

## 9. Risks and assumptions

- **Nothing in track 2 has ever run in Excel.** 41 unrun checklist items, the oldest being the
  `bafb905` used-range fix, now eight phases deep. A defect found there may reopen a wave
  rather than only adding to this phase. This phase adds to that debt; it does not discharge
  it.
- **Layout at 51 actions is unverified.** Auto-layout packs 3 columns × 6 buttons = 18 per
  page, so 51 actions plus chains is roughly four pages. `PkgFixture` was the design's
  mechanism for proving this and is dropped, so the only remaining proof is the owner looking
  at it. `Macros Home`'s page summary at least makes the result legible at a glance.
- **Assumption — `AddComment` notes** survive Setup's delete-and-rebuild and anchor correctly
  on an `AutoFit` header row (§4.3).
- **Assumption — a macro can delete the sheet hosting the button that invoked it** (§5.3).
- **Every action still arrives Off.** A fresh import shows 51 catalog rows and no working
  buttons. That is intended, and it is what makes the reminder necessary. Whether the
  *default* should change is a separate decision the owner has not been asked to make, and no
  part of this phase assumes either answer.

---

## 10. Decision log

Every entry below was settled by the owner on 2026-07-29 and is not to be re-opened inside
this phase.

| Decision | Outcome |
|---|---|
| `PkgFixture` | Dropped. Do not build, do not re-propose. |
| Worked example per action | Deferred to issue #30. Home ships with descriptions only. |
| Macros Home scope | All 51 actions with their current state. |
| Chain step format | Bare ids, comma-separated; old bracketed form still parsed. |
| `export.email` | Writes the file, opens a draft, never sends; refuses gracefully and keeps the file. |
| `workbook.toc_links` | One configurable cell per sheet; skips and reports where occupied. |
| Header notes | Availability, Placement (both blocks), Steps, Column types. Built last. |
| Finalize reminder trigger | Fires whenever anything is not `Off`, permanent-alert trade accepted. |
| `Complete Macro Setup` deletes | `Setup`, `Setup Menu`, `Macros Home`. Not `_Config`, not `_Packages`. |
| Unfinalized-changes guard | Warn in the confirmation and proceed; no automatic detection. |
| Finalize auto-running Setup | Deferred to issue #29. |
| Stripping code for Off actions | Deferred to issue #31. |
| Printing a `build.ps1` command to the user | Rejected. It contradicts the project's mission of being usable by non-technical people. |
| Distribution | Repo goes public with versioned GitHub Releases, **after** all work is done and the repo is sanitized. |
| `Core_SpeedOn` / `Core_SpeedOff` | Become `Private`, so Alt + F8 lists only `Core_BuildSetup` and `Core_FinalizeSetup`. |

---

## 11. The non-technical setup guide

Two deliverables from one source: `docs/setup-guide.md` for the web, and a PDF generated from
it. Written for someone who has never opened the VBA editor and does not know what a module
is.

### 11.1 The hard rule

**No more than a 6th-grade reading level.** Owner's instruction, stated as a hard rule, so it
is **measured rather than asserted** — `scripts/reading-level.mjs` computes the
Flesch–Kincaid grade of the finished file, and the guide does not ship above **6.0**. Claiming
a reading level without measuring it is how a document written by someone fluent in the subject
ends up at grade 11.

What that means in practice:

- One action per step. Never two.
- Name the exact keystroke and the exact on-screen text: *"Press **Alt + F11**"*, *"Click
  **File** in the top left"*, *"Click **Import File...**"*.
- No jargon, and no jargon smuggled in as a parenthetical. Not *module*, not *macro-enabled*,
  not *dispatcher*, not *manifest*.
- Short sentences. Say what they will see after each step, so they can tell whether it worked.

### 11.2 The order it teaches

The settings cycle is explained **before** the walkthrough, because it is the one thing that
confuses everyone and it is not guessable:

> To change your settings you click **Finalize Setup Tab**, then **Create Setup Tab**. That
> makes the new questions appear. After you answer them, click **Finalize Setup Tab** again to
> save them.

Then the walkthrough, end to end: download the file → **Alt + F11** → **File** ▸
**Import File...** → pick the file → **Alt + Q** → **Alt + F8** → run **Core_BuildSetup** →
fill the yellow cells on **Setup** → set the macros you want to **Button** → **Setup Menu** ▸
**Finalize Setup Tab** → **Create Setup Tab** → answer the new questions → **Finalize Setup
Tab** again → your buttons are on the **Macros 1** page and all of them are listed on **Macros
Home** → when you are finished for good, **Setup Menu** ▸ **Complete Macro Setup**.

It closes with three short asks: you can ask for a new macro on the repo's Issues page; check
back now and then for a newer version; and if this helped, please star the repo.

### 11.3 One snag the guide has to handle

**Alt + F8 does not show a clean list.** It shows exactly four names, because Excel hides any
`Sub` that takes arguments and every action entry point takes `(ctx As Object)`:

```text
Core_SpeedOn
Core_SpeedOff
Core_BuildSetup
Core_FinalizeSetup
```

`Core_SpeedOn` and `Core_SpeedOff` are internal helpers that no user should ever run, and
running one leaves Excel with calculation and screen updating switched off, with nothing on
screen saying so — which then looks like the toolkit broke the workbook.

**Settled by the owner, 2026-07-29: both become `Private`.** The bundle is a single module, so
nothing outside Core calls them by name and no behaviour changes; Alt + F8 then lists exactly
`Core_BuildSetup` and `Core_FinalizeSetup`, which are the two things the guide tells people to
run. It narrows Core's public surface, so contract §9's platform-abstraction appendix is
checked for a reference to either name in the same step. Done in **step 7**, not step 11, so
the change is in the bundle before the guide is written against it.

### 11.4 Verification

- Flesch–Kincaid grade ≤ 6.0, measured.
- **Blind dry-run**, same discipline as the checklist: a subagent forbidden from reading
  `src/`, given only the guide and the built bundle, follows it and flags every step it cannot
  execute. A guide is exactly the artifact where blindness is the whole test.
- Every button caption, sheet name and procedure name the guide types must be greppable in
  `dist/MacrosBundle.bas`.

---

## 12. Sanitize, publish, release

**Last, and not started until steps 1–11 are done and merged.** Owner's sequencing.

The repository is private today and has no releases, so there is currently no way for a
non-technical user to obtain the bundle at all — which is also why the guide's
request-a-macro, check-for-updates and star-us asks have nothing to point at.

**There is no security reason to withhold the built `.bas`.** It is plain text; a scan of
`dist/MacrosBundle.bas` found no credentials, no URLs, no email addresses, and only three
absolute paths, all of them `e.g.` examples inside help strings (`C:\Close\January`). A `.bas`
is also a safer thing to hand someone than an `.xlsm`, which browsers, mail clients and
Excel's own Mark-of-the-Web all block by default. `dist/` stays gitignored regardless: a
release asset is not a tracked file, so nothing generated enters the source tree.

Sequence:

1. **Sanitization review of the whole repository, including git history** — not just the tip.
   Secrets, client and franchise identifiers, internal absolute paths, anything in
   `docs/designs/` written on the assumption of a private audience. A separate reviewed step,
   because making a repository public cannot be undone: clones and forks survive a revert.
2. **Confirm the public-facing basics** — a LICENSE, a README aimed at a stranger, and an
   Issues page a non-technical user can post to.
3. **Make the repository public.** Owner's action, not the agent's.
4. **Cut release `v0.18.0`** with the built bundle attached as
   **`MacrosBundle-v0.18.0.bas`** — the version in the filename is what lets a user tell two
   downloads apart, which was the owner's requirement.
5. The guide links to **`/releases/latest`**, never to a pinned version, so it does not go
   stale on the next release.
