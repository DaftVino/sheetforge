# Finish Excel — built-page polish, then catalog build-out

**Status:** design agreed 2026-07-24, ready for `writing-plans`.
**Baseline:** `main` at v0.3.0, contract v1.1.0.
**Scope:** the remaining Excel work before the Apps Script port. The port stays
deferred to last: the frozen contract only makes it mechanical *after* the Excel
feature set settles, so porting against a moving target means porting twice.

---

## 1. The two tracks, and why polish goes first

**Track 1 — built-page UI polish (`v0.4.0`).** Core-only. The owner called the
generated panels "disorganized and cluttered" at **seven** buttons. Shipping 50+
more into that layout first makes the complaint worse, and every catalog package
also adds config rows to the one Setup sheet the user scrolls. Polish is the
container; the catalog fills it.

**Track 2 — catalog build-out (`v0.5.0` onward, one minor per wave).** ~58 of the
~60 candidates in `docs/macro-catalog-plug-and-play.md` are unbuilt. Built in
waves of 5–8 actions, each wave a shippable increment.

Sequencing was decided by the owner. The risk it creates — designing a layout for
60 buttons while only 7 exist — is answered in §2.5.

---

## 2. Track 1 — the polish pass (`v0.4.0`)

### 2.1 What exists already (checked, not assumed)

Reading Core before designing changed the shape of this track:

- **`RenderPlacementBlock` (`Core.bas:1242`) is already a per-action catalog
  row.** It emits `Caption [id]`, the id hidden in `RID_COL`, and one yellow
  editable cell per action. Per-action Enabled is *one more column on an existing
  block*, not a new mechanism.
- **`RenderPage` (`Core.bas:1734`) already draws a page title and column
  headers.** So §8's "per-page titles, column headings" are partly built. The
  real gaps are different: grouped pages carry **no descriptions at all** (only
  `BuildSinglePanel` writes them, into column C), titles are meaningless
  (`"Macros - Page 1"`, `"Column A"`), there is no legend, no navigation, and no
  auto-layout.
- **`GatherButtons` (`Core.bas:1549`) is the single funnel** for both actions and
  chains, so disabled-action filtering lands in exactly one place.
- **`Core_BuildSetup` (`Core.bas:825`) deletes and rebuilds Setup**, re-seeding
  from the finalized `_Config`. This is the constraint that decides §2.3.

### 2.2 What the user sees afterwards

- **`Macros Home`** — generated first tab: toolkit title, a legend explaining the
  page/column code convention, one card per display category showing its name,
  action count and a hyperlink to its page, plus a link to `Setup Menu`.
- **`Macros N` pages** — titled with the **category name**, not `"Page 3"`.
  Columns headed by section name, not `"Column A"`. Every button followed by its
  manifest description (porting what `BuildSinglePanel` already does). A
  persistent `← Home` link in a fixed position on every page, plus Previous/Next
  page links.
- **Auto-layout.** With no codes assigned, buttons distribute by display category
  into named pages, filling each column to a maximum height before wrapping to
  the next. The 60-button wall that `BuildSinglePanel` would become stops being
  the default. Manual `1A` codes remain a complete override; `hide` still omits a
  button entirely.
- **Tighter spacing**, via the existing `BTN_W` / `BTN_H` / `ROW_GAP` /
  `COL_GAP` / `HDR_H` constants.

**Not built:** no search — a generated, well-sectioned index plus Excel's own
Ctrl+F already solves finding an action, and a worksheet search widget is
decoration until evidence says otherwise. **No Recent Actions, ever** — it needs
runtime writes that the next Finalize overwrites, so it would quietly break.
**Favorites deferred** — curation has no value at seven actions; it can arrive
later as an action-level flag on Setup with no contract break.

### 2.3 Per-action enable, and the chicken-and-egg

The placement block becomes the **Action Catalog block**:
`Action | Enabled | Page code`. `Enabled` defaults to yes for already-shipped
actions, so v0.4.0 is a no-op for a current user. It is read by a sibling of
`ReadCodeMap` keyed off `RID_COL`; `GatherButtons` skips disabled actions.
Disabled actions stay listed in the catalog block, so nothing becomes
undiscoverable — which is the owner's hard requirement.

**The wrinkle.** "Render a config field only when an enabled action consumes it"
is the obvious design and it is wrong here. Because `Core_BuildSetup` deletes and
re-seeds Setup from the *finalized* `_Config`, that flow costs
**toggle → Finalize → rebuild Setup → fill → Finalize again**, and the rebuild
discards anything typed but not yet finalized.

**Resolution.** Every field always renders, but fields belonging to no enabled
action are relegated to a collapsed band at the bottom — *"Settings for actions
you haven't turned on"* — using Excel's outline grouping. The user's working area
stays short, there is no rebuild cycle, nothing typed is lost, and it degrades
safely to a plain visible band if grouping fails.

This needs `Core_AddField` to gain an optional `usedBy:="clean.dupes"` argument.
A field that does not declare it is treated as package-wide and always renders in
place, so both existing packages keep working untouched.

### 2.4 Contract impact — additive, `v1.2.0`

New: the `enabled` action attribute · the field `usedBy` attribute · the
`category` action attribute (§3.2) · the auto-layout rule for uncoded buttons ·
the reserved `Macros Home` sheet name. No existing id, key, or behaviour changes,
so this is a minor bump, not a major one.

### 2.5 Verifying a 7-action pass against 60 actions

Pages are generated from registry data, so a throwaway **`PkgFixture` declaring
~60 no-op actions across 9 fake categories** exercises the 60-action layout,
pagination and column wrapping in a single Excel session. It never ships: it is
excluded from the bundle by simply not being in Core's candidate list, and it is
deleted once the pass is verified.

---

## 3. Track 2 — the catalog build-out

### 3.1 The ID contract, frozen in `v1.2.0`

Two segments, `<domain>.<operation>`, lowercase ASCII and underscores.

**The domain is a controlled user-facing vocabulary, not the module that
implements it.** This is the decision that matters most, because an action id is
a permanent dispatch key that the user types by hand into a Setup cell as
`Label [action.id]`. Encoding the module in it would make a future package split
a major contract break for no benefit.

Nine domains: `import` · `clean` · `validate` · `enrich` · `format` ·
`summarize` · `export` · `workbook` · `backup`. Reserved: `chain.*`.
Grandfathered: `importexport.*`.

Frozen rules:

- The operation segment is descriptive where ambiguity looms —
  `clean.remove_blank_rows`, never `clean.blank2`.
- **No new id is ever camelCase.**
- The eight shipped ids are permanent exceptions, listed literally, never renamed
  and never aliased: `clean.dupes`, `clean.trim`, `clean.filldown`,
  `importexport.clear`, `importexport.import`, `importexport.updateMaster`,
  `importexport.export`, `importexport.suggestTypes`. Two are camelCase and the
  `importexport` domain conflates two different jobs — direct evidence for the
  rule above, and both stay exactly as they are.
- Ids are immutable. Captions, help text, entry points and module assignment stay
  free to change.
- No alias or migration table. It would create two permanent names for one action
  and would have to be ported to Apps Script.

**Enforcement, in two places.** `Core_AddAction` rejects a duplicate id during
registry assembly, which protects a hand-edited bundle. `build.ps1` extracts
literal `Core_AddAction` ids and fails on duplicates *and* on any id whose domain
is absent from the contract's vocabulary table. `spec/contract.md` owns that
table. `build.ps1` has no id check at all today.

### 3.2 Display category, decoupled from module and id

**`Core_AddAction` gains an optional `category:=`**, defaulting to the package
name. Home groups by category.

This fixes a legacy incoherence for free: `importexport.import` declares
"Bring Data In" while `importexport.export` declares "Get It Out" — one module,
two Home cards, no id touched. New import work lands in a new `PkgImport` under
`import.*` ids and merges onto the *same* Home card as the legacy import action.
The user sees coherent jobs; modules stay whatever is convenient to maintain.

Packages are created only when a wave needs one. No nine empty modules up front.

### 3.3 Implementation shape

Every listed action gets its own permanent public `Sub` and manifest entry —
non-negotiable, it is how the owner's "all actions listed" requirement is met.
Each delegates to a **private, package-prefixed family engine** where variants
share real work.

An engine is capped at one coherent workflow. When it starts branching on
unrelated concerns it gets split, rather than growing into a flag maze where a
fix for one variant silently changes three others. Because the bundle is one flat
namespace, private helpers and enums are package-prefixed too
(`PkgImport_CombineEngine`, not `CombineEngine`). Cross-package sharing is
forbidden; anything genuinely shared moves into Core.

### 3.4 Wave 1 — the ingest→clean path (7 actions)

| Action | ID | Rating | Module |
|---|---|---|---|
| Combine all worksheets into one master | `import.combine_worksheets` | ⭐5 / A | `PkgImport` (new) |
| Combine all workbooks in a folder | `import.combine_workbooks` | ⭐5 / A | `PkgImport` |
| Import CSV/TXT with delimiter parsing | `import.delimited_file` | ⭐4 / A | `PkgImport` |
| Remove blank rows / columns | `clean.remove_blanks` | ⭐5 / A | `PkgClean` |
| Unmerge cells and fill down | `clean.unmerge_fill` | ⭐4 / B | `PkgClean` |
| Standardize number / date formats | `clean.standardize_formats` | ⭐4 / B | `PkgClean` |
| Convert text-numbers to real numbers | `clean.text_to_numbers` | ⭐3 / B | `PkgClean` |

The two combines share one private `PkgImport_CombineEngine` — identical
header-aligning append, differing only in source enumeration.
`clean.text_to_numbers` calls the **existing `Core_Coerce`** rather than
reimplementing coercion.

Chosen as a coherent workflow rather than a cross-category sampler because this
toolkit's distinguishing feature is chains: with the shipped `clean.trim`,
`clean.dupes` and `clean.filldown`, these seven form a complete
raw-files→clean-dataset job the user wires into one chain the day it lands. A
sampler optimises for the builder's confidence in architectural coverage; the
user experiences whether they can finish a job.

### 3.5 Wave protocol and order

5–8 actions per wave. Per wave: author in `src/` → `pwsh build/build.ps1` →
`node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` → import + Excel
Debug ▸ Compile VBAProject + **click every action, every cancel path, and every
error-restore path** → `/review` → `/ship` as a minor bump. Each wave is an
additive contract minor.

Order after wave 1, deliberately revisable between ships — small waves exist so
this can change:

1. **W2 — Validate & Reconcile.** Reconcile two datasets on a key (⭐5), the
   validation scan to an exceptions report, highlight duplicates, flag missing
   required fields. Highest-value group for a multi-entity close, and it consumes
   wave 1's output: combine → clean → reconcile is one continuous chain.
2. **W3 — Format & Present, plus export to PDF.**
3. **W4 — Summarize & Analyze, plus Match & Enrich.**
4. **W5 — Protect & Back Up.**
5. **W6 — Workbook housekeeping.**

### 3.6 The stopping rule

An action earns its place only with: a stable user story in this finance
workflow · bounded inputs and outputs · a safe cancel and error-restore path · a
hand-verification script someone can repeat. Popularity alone does not qualify;
"most-downloaded VBA snippet" is not a maintenance contract.

**Waves W1–W6 assign 34 catalog items, producing 35 new actions.** With the 13
already shipped or embedded, that brings the toolkit to ~48 listed actions. A
further 8 items are marked *Later* (in scope, no wave yet), 5 *Deferred*, and 4
*Declined*. "Build all 60" is rejected as a goal. Verdicts are annotated per item
in `docs/macro-catalog-plug-and-play.md`, which is now the authoritative build
queue.

One dependency worth noting: `backup.snapshot` (W5) is the unblocker for several
*Deferred* destructive items, since it supplies the restore point they lack. If
those items become wanted sooner, W5 moves up rather than the safety story being
skipped.

**Declined outright:**

- *Bulk personalized email* — not responsible without preview, recipient count,
  explicit confirmation and a draft-only mode.
- *Timestamp row on edit* and *spotlight active row/column* — both need event
  handlers in sheet or `ThisWorkbook` modules, architecturally incompatible with
  a single imported standard-module bundle. Faking support is worse than
  declining.
- *Jump to next/previous sheet* — no value as a button.

**Deferred pending a safety story:**

- *Delete all empty sheets*, *hide all sheets except active* — destructive with
  no preview.
- *Pull data from a closed workbook* — undefined failure mode.

---

## 4. Why the contested decisions went the way they did

Settled by a `/deliberate` panel (native / veteran / aesthete, `gpt-5.6-terra`
and `gpt-5.6-sol` at medium effort, 2 rounds, ~110k tokens). Stopped a round
early: two of three live questions resolved, and the third had not moved across
either round.

**Converged in round 1** — 8–9 job-oriented packages; a thin listed public
wrapper per action over a private family engine; per-action enable (per-package
is too coarse at ten actions per package); the single-`Macros`-tab default is
dead, replaced by a generated index plus category pages with `1A` codes demoted
to an override; waves of 5–8; "build all 58" rejected.

**The ID split (2-1, then reversed).** The real disagreement was whether an
action id is an engineering identifier or a user-facing token the customer types
by hand. Once that was put directly, the aesthete flipped to two segments and
supplied the synthesis both others missed: the first segment is a controlled
*domain vocabulary* that merely happens to align with package names today. That
buys the veteran's reorganisation-safety at the native's ergonomics with zero
migration. The native was defending a bad convention — "package belongs in the
id" is exactly the coupling that makes a package split a major break. The veteran
over-generalised: `clean.rows.remove_blank` charges real daily typing cost
against a reorganisation that may never happen.

**Navigation.** The aesthete's round-1 Macro Home with search, recents and counts
ignored cost, and it withdrew search and recents itself once asked where the
state lives. Recent Actions died on a defect all three agreed on, not on taste.

**Where this design overrides the panel.** All three wanted Setup to render a
field only when an enabled action consumes it. Given `Core_BuildSetup`'s
delete-and-reseed behaviour (§2.1), that costs a double-Finalize and can discard
typing, so §2.3 collapses the band instead of omitting the fields.

**The panel's most useful contribution** was the `PkgFixture` idea in §2.5 — the
concrete answer to designing a 60-button layout while only 7 actions exist.

---

## 5. Handoff log

### 2026-07-29 - the Macros Home phase: 8 of 12 tasks done, 51 actions, all code complete

**Track 2 landed first.** `feat/w6-workbook` merged to `main` as **PR #28** (merge commit, not
squash — the twelve commits were eight reviewed phases). The stranding described in the entry
below is fixed: `origin/main` went from v0.11.0 / 3 packages to **v0.17.0 / 10 packages**.
Verified on disk with `git ls-tree`, not taken from the PR list, which was the trap.

**Branch: `feat/macros-home`**, cut from `main` *after* that merge. 17 commits. Nothing merged
yet; no PR opened. Working tree clean apart from the deliberately-untracked `.vscode/`.

**All eight code tasks are done.** Tasks 9–12 (checklist reconcile + blind dry-run, docs sweep,
setup guide, sanitize/publish/release) remain. `VERSION` is still `0.17.0` — Task 10 bumps it.

| Metric | Before | After |
|---|---|---|
| Action ids | 43 | **51** |
| Contract | v2.9.0 | **v2.10.0** |
| Bundle | 10859 lines, Sub 123/123, Function 182/182 | **12512 lines, Sub 133/133, Function 197/197** |
| Checklist items | 41 | **79** |
| Packages | 10 | 10 (no new package, no new domain) |

There is no test count: TDD is off here. Those build metrics are the equivalent.

#### What the plan got wrong — all **do not revert**

1. **`StepsToTokens` must write through `CanonicalActionId`.** The plan said "join the raw
   ids". That silently dropped the normalisation whose own comment above `ActionTokenFor`
   records it *had already been lost once by a refactor and restored as fix round 2*. Without
   it a Steps cell typed `CLEAN.DUPES` keeps that casing forever.
2. **The reminder formula is written AFTER the action rows, never predicted before them.** The
   first draft computed the block's extent from a registry count. Correct today, but the same
   `lastDataRow` bounds a `.Validation` call whose own comment warns about leaking outside the
   block — so the first future filter added to that loop would break both silently.
3. **The reminder compares `AvailLabel(AVAIL_OFF)`, not `AVAIL_OFF`.** The cell holds `"Off"`;
   the token is `"off"`. Case-insensitive `COUNTIFS` would probably have saved it, but that is
   a claim about Excel nothing has run to confirm.
4. **`ParseStepTokens` must never comma-split a line containing a bracket.** A caption may
   contain a comma; splitting one halves a legacy `Caption [id]` token. No caption does today
   (verified, 0 of 51) — the parser must not depend on that quietly.
5. **`workbook.sort_sheets` must never read `Worksheet.Index`.** It indexes `Sheets`, which
   includes chart sheets, while the sort walks `Worksheets`. Any workbook holding a chart
   diverges the two and the sort moves the user's own tabs to wrong positions, silently.
   Checklist 4.26 exists solely to catch this and *requires a chart sheet to be present*.
6. **`export.email`'s Outlook-unreachable refusal must not end `"Nothing has been changed."`**
   It had, three lines after saying it wrote a file.
7. **`HeaderNote`'s `On Error GoTo 0` is legal only there**, because that helper has no `Fail:`
   handler. Do not copy the line into one that does.

#### Discovered — empirical, with evidence

- **codex is `gpt-5.6-sol` at `model_reasoning_effort = "high"` by default** (`~/.codex/config.toml`).
  `--model sol` is **rejected**: *"The 'sol' model is not supported when using Codex with a
  ChatGPT account."* Adding `--model` / `-c` flags made a run **hang at
  `Reading additional input from stdin...`** and produce nothing for many minutes. Invoke as
  plain `codex exec --sandbox workspace-write "<prompt>" < /dev/null`. Its Windows sandbox
  also cannot spawn shell children (`CreateProcessAsUserW failed: 5`), so it falls back to a
  node REPL — never ask it to run a build.
- **Neither `build.ps1` nor `vba-syntax-check.mjs` resolves symbols.** A typo'd call to a
  helper that does not exist passes both gates and fails only at `Debug ▸ Compile` in Excel.
  Every new call site must be checked by hand; that is how the 15 helpers `HomeCatalog` calls
  were verified.
- **The action-id gate scans every package even on a `-Packages` subset build.** It cannot be
  dodged.
- **`Core_AddAction` must carry id, caption and entry point as three literal strings before any
  line continuation.** The gate rejects anything else and says so precisely.
- **`tools/GateAudit.bas` is imported as its own module**, so module-scope `Private` puts a
  Core helper out of its reach. It calls only `Core_BuildSetup` and `Core_FinalizeSetup`.
- **Alt+F8 lists only public `Sub`s taking no arguments** — now exactly `Core_BuildSetup`,
  `Core_FinalizeSetup`, `Core_CompleteSetup`.

#### Still open

- **Tasks 9–12.** Task 9 is the checklist reconcile plus a **blind dry-run** — a subagent given
  the checklist and the bundle and forbidden `src/`. The previous such pass found five buttons
  the checklist named that do not exist.
- **Task 12 needs the owner and cannot be done by an agent.** Making the repo public is
  irreversible; the owner performs it, and picks a LICENSE if none exists.
- **Nothing in this phase has run in Excel.** 79 unrun checklist items now, up from 41. The
  oldest, the `bafb905` used-range fix, is nine phases deep and still unconfirmed.
- **Layout at 51 actions is unproven.** `PkgFixture` was the design's mechanism for proving it
  and the owner dropped it. The only remaining proof is the owner looking at the sheet.
- Three issues filed and deliberately **not** built: **#29** Finalize auto-running Setup,
  **#30** a worked example per action, **#31** stripping code for Off actions (which records
  that the run-time version is impossible here — it needs VBIDE trust, forbidden by rule 3).

**Next phase — Tasks 9 to 12.** Read manifest is in `docs/designs/next-session-prompt.md`.


### 2026-07-29 - track 2 is code-complete (43 actions), and the stack never reached main

**Eight phases planned, reviewed and merged as PRs #20-#27.** Two Core passes then W2-W6,
plus one end-to-end validation document. `feat/w6-workbook` carries the finished tree:
**43 actions across ten packages**, `VERSION` 0.17.0, contract v2.9.0, `SCHEMA_VER` still
`"v2.0"`. Build is green (`Core + 10 package(s)`, `Sub 123/123 | Function 182/182`), syntax
check clean.

**The defect this hand-off exists to flag: `main` is still at v0.11.0 with three packages.**
Every PR reads MERGED, but each was based on the previous feature branch and only #20
targeted `main`. Phases 2-8 merged into each other and stopped. This is the same stranding
that produced #18. **Nothing is lost** - `feat/w6-workbook` is complete and pushed - but the
first job of the next session is landing it, and no new branch may be cut from `main` until
it is.

**What review bought, since it is the argument for doing it again.** Planning every phase
first, reviewing each plan before writing code, and dry-running the checklist blind caught
four defects that would otherwise have shipped, two of which destroy data: `importexport`'s
import and updateMaster discarded a colleague's unsaved edits by closing a workbook they had
not opened; W5's snapshot pruning would `Kill` any file merely starting `SNAPSHOT_`; W6's TOC
skipped the reserved-name check whenever the sheet already existed, so an existing `_Config`
could reach a `Cells.Clear`; and `Core_DestUsable` accepted `Macros 3`, which builds, fills
with a month of data, and is deleted without a prompt by the next Finalize. The blind
dry-run then found five buttons the checklist named that do not exist.

**Nothing has run in Excel.** Every phase shipped reviewed, not verified, per the 2026-07-28
gate deferral. `docs/designs/2026-07-28-end-to-end-validation.md` is the 41-item consequence-
ordered checklist that discharges it, and it is still unrun.

**Scope the owner corrected on 2026-07-29.** Declaring track 2 "complete" measured the work
against the plan of record rather than against this design. Three gaps were real and are the
next phase: `Macros Home` (design section 2.2) was built during track 1 and withdrawn before
release; the **8 catalog items marked *Later*** say "in scope, no wave assigned yet" and were
never assigned; and a **Finalize-reminder** the owner expected was never specified anywhere.
`PkgFixture` (section 2.5) was also never built and is now **explicitly dropped** - the owner
does not want it.

**Next phase - Macros Home, the 8 Later items, and the Finalize reminder.** No plan document
exists yet; writing one is the first task, in the same shape as track 2: plan every step
first, review each plan before any code, then implement step by step. Contract goes to
v2.10.0 or higher. Ending catalog: **51 actions**.

**Read manifest for the next phase.** `docs/designs/next-session-prompt.md` holds it.


### 2026-07-28 — W1 built and open as PR #19; gate DEFERRED; Core pass planned

**Branch / PR.** All work on `feat/w1-ingest-clean`, 27 commits, `29a3392..bafb905`,
open as **PR #19 against `main`** and **not merged**. Nothing was squash-merged. One PR,
not stacked — the #18 lesson held.

**What shipped.** `spec/contract.md` v2.2.0 → **v2.3.0**; `build/build.ps1` gained the
action-id gate; **new** `src/packages/PkgImport.bas`; `src/packages/PkgClean.bas` +1 field
+4 actions; one line in `Core.bas`'s `CoreCandidates`; `tools/GateAudit.bas` gained an
action-catalog probe; **new** `docs/designs/2026-07-28-w1-manual-gate.md` (39 items);
**new** `docs/designs/2026-07-28-core-refuse-pass-plan.md`. Build: `Core + 3 package(s)`,
`Sub 75/75`, `Function 148/148`, 223 procedures. **Test count unchanged: none, by repo
policy** — VBA has no headless runner.

**Corrections made during execution — DO NOT REVERT:**

- **`PkgImport_AppendSheet` measures its extent per column with `End(xlUp)` from the
  bottom, taking the maximum — NOT from `UsedRange`.** *Reason:* `UsedRange` is a rectangle
  Excel inflates with formatting and with cells that once held data, so a two-row sheet can
  report itself sixteen rows deep. The function advances `destRow` by the extent it is
  given, so an inflated source pushed every LATER source down by the difference: sources
  landed ~15 rows apart and a two-row file wrote to rows 27–28. Found by the owner at gate
  item 2.3. Do **not** "simplify" this back to `UsedRange` — a prior round did exactly that
  to fix a column-A truncation, and traded a smaller bug for a bigger one.
- **The combine probes for an already-open workbook and never closes one it did not open**
  (`PkgImport_AlreadyOpen`, `weOpened`). *Reason:* `Workbooks.Open` on a file already open
  in the same instance returns **the user's own copy**, and with `DisplayAlerts` suppressed
  under `Core_SpeedOn` the "you will lose your changes" prompt never fires — so
  `Close SaveChanges:=False` silently destroyed a colleague's unsaved edits while reporting
  success. The probe also confirms `wb.FullName`, because the `Workbooks` collection is
  keyed on file *name*.
- **`ThisWorkbook` is excluded by path OR by file name, not path alone.** *Reason:* on
  OneDrive/AutoSave `ThisWorkbook.FullName` is an `https://` URL, and on a share it can be
  UNC while the configured folder is a mapped drive — path-only misses, this workbook joins
  its own source list, and the run dies on a name collision that does not exist.
- **Every VBA-syntax pattern in `build.ps1` carries `(?i)`.** *Reason:* `[regex]::Match` is
  case-sensitive in .NET, so lowercase VBA went unmatched. Fixing only the id gate opened a
  worse hole: a consistently-lowercase procedure was counted by neither balance pattern and
  became invisible to every structural guard — a silent full pass.
- **The id gate's comparers are ORDINAL and its id checks use `-cmatch`/`-cne`.** *Reason:*
  PowerShell's `-match` and hashtable keys are case-INSENSITIVE, so
  `import.combineWorksheets` passed the lowercase-only shape check and §7's no-camelCase
  rule — the one rule the gate exists to enforce — was unenforced.
- **`k_clean_formats` declares NO default rows.** *Reason:* §5 resolves an empty stored
  table to the field's declared defaults, so demo rows make
  `clean.standardize_formats`' required refusal **unreachable dead code** and its gate item
  unexecutable. Now normative in contract §8.
- **`PkgImport_DelimitedFile`'s cleanup lives one frame down** (`PkgImport_ImportDelimited`).
  *Reason:* VBA will not re-enter a handler in the same procedure, so a re-raise inside
  `Fail:` escaped and `Core_Fail` never ran.

**Discovered (empirical — each cost a round trip):**

- **`Workbooks.Open` on an already-open file returns the open copy**, not a read-only
  second one. This is the whole basis of the critical fix above.
- **A `vbYesNo` MsgBox has its title-bar X DISABLED** — it does not return `vbNo`. A gate
  item telling the owner to dismiss one with the X is unexecutable. Corrected in place.
- **`Core_RunChain` does NOT abort on a failed step** — it asks "Continue with the
  remaining steps?" (`Core.bas:1248-1252`, `AskContinue`). A refusal therefore reports as a
  green tick unless the action sets `ctx("failed")`, which is the entire premise of the
  Core pass planned next.
- **`node scripts/code-map.mjs` stamps the PREVIOUS commit** when run before committing,
  because it reads `git log -1` over indexed files. Regenerate and commit, then regenerate
  again if the stamp matters — or accept the off-by-one. Bit this session twice.
- **The blind dry-run cannot catch an instruction that is wrong about which workbook to
  use.** It holds the built bundle, so it already knows `import.combine_worksheets`
  iterates `ThisWorkbook.Worksheets`, and "make a scratch workbook" reads as harmless
  shorthand. The owner hit it on the first section. It remains an authoring gate, not a
  test.

**Still open — the W1 manual gate is DEFERRED, and needs the human.**
Owner decision 2026-07-28: fold it into one end-to-end validation on the finished workbook
once the Core pass and the remaining waves are built, rather than spending owner time on
defects a careful read should find. **4 of 39 items ran.** 0.1, 0.2, 0.4 passed
(`GateAudit_Run`: 31 PASS / 0 FAIL / 2 SKIP, all fifteen action ids present). 0.3 is
inconclusive — `Macros 1` did not exist afterwards, so redo it first. **2.3 failed and
found the used-range defect above; the fix in `bafb905` is NOT verified in Excel.** Items
**1.5** and **2.4** exist to confirm it when the gate resumes. Also still unrun: the 8
track-1 items, of which reskin 4 (dark theme) remains the likeliest real finding.

**Next phase — the Core pass**, planned in full at
`docs/designs/2026-07-28-core-refuse-pass-plan.md` (28K, 5 tasks). `Core_Refuse` +
`Core_IsEngineSheet` + the `X #0` chain-summary fix, then delete the package copies that
exist only because Core lacked them. Owner decisions already settled: Core owns the refusal
dialog; a refusal STOPS a chain while an error still asks; scope excludes the remaining W1
minors. Contract goes to v2.4.0.

**Read manifest for the next phase.** `docs/designs/2026-07-28-core-refuse-pass-plan.md`
(28K — the contract, read whole) · `spec/contract.md` (68K — **over the slice threshold**;
grep `## 6.` and `## 10.`) · `src/core/Core.bas` (232K — **never read whole**; grep
`docs/code-map.md` (36K, current at `bafb905`) for `Core_Done`, `RunAction`,
`Core_RunChain`, `Core_SheetExists`) · `src/packages/PkgImport.bas` (40K — 8 refusal sites)
· `src/packages/PkgClean.bas` (36K — 1 refusal site) ·
`docs/designs/2026-07-28-w1-manual-gate.md` (32K — gains the new items in Task 5).

### 2026-07-28 — track 1 finished and shipped (v0.7.0–v0.9.0); W1 approved, not started

**Branch / commits.** All on `main`, five PRs merged: **#14** Part A (v0.6.0),
**#15** Core hardening (v0.7.0), **#16** Setup reorder (v0.8.0), **#17** identity
reskin (v0.9.0), **#18** a re-land. `main` is at **v0.9.0 / contract v2.2.0**;
`SCHEMA_VER` deliberately still `"v2.0"`.

**#16 and #17 merged into their base branches, not `main`.** They were stacked PRs
whose bases were the branch below them, and merging all three back-to-back beat
GitHub's retarget — 11 commits stranded one level down, and #18 exists purely to
bring them over. **If you stack PRs again: merge one, wait for the retarget, then
merge the next** — or base every PR on `main` up front and accept the noisier diff.

**What shipped.** `src/core/Core.bas` (now 229.3K), `spec/contract.md` (v2.0.0 →
v2.2.0), `CLAUDE.md`, README + CHANGELOG (logo), **new** `tools/GateAudit.bas`,
**new** `docs/assets/sheetforge-logo.png`, **new** `docs/macro_code_library.md`
(owner-authored, 30.8K), and four gate documents. Build: `Core + 2 package(s)`,
`Sub 62/62`, `Function 129/129`, 191 unique procedures. Test count unchanged:
**none, by repo policy** — VBA has no headless runner.

**Closed:** #6, #10 · #8, #12, #13, #3 · #11.

**Corrections made during execution — DO NOT REVERT:**

- **`ClearCatalogBlock` clears row by row and skips merged rows** (`114dfa3`).
  *Reason:* `TintRow` merges columns B:E, and a blanket `Range.ClearContents` over
  D:E is a **partial** clear of that merge → `1004 We can't do that to a merged
  cell`, aborting the rebuild. Do not simplify it back to a range clear, and do not
  widen it to B:E — a merged row carries no id and no availability, so skipping it
  is correct, not a workaround.
- **`CfgSheet` adopts an orphaned `_Config_staged` when `_Config` is missing.**
  *Reason:* staged-then-swap alone opens a *new* way to lose data — a stage sheet
  left by an interrupted swap would be cleared by the next activate, destroying the
  only surviving copy. Verified by hand; gate item 1.2 passed.
- **`MarkerVerdict` requires positive proof of the pre-v1.2 shape** (`LooksPreV12`)
  instead of reading the visible `Availability` header back. *Reason:* a
  user-editable caption was deciding a machine-owned question; renaming it promoted
  a tampered block to the legacy path and armed every destructive action. The
  residual hole is documented as gate item 2.2, not a silence.
- **`firstContentRow` is assigned ABOVE the compose call in `Core_BuildSetup`.**
  *Reason:* Chain `Steps` cells are `WrapText` multi-line; anchoring AutoFit below
  compose clips them. That is why the rename from `firstFieldRow` exists at all.
- **Semantic colours are NOT in the identity palette.** The yellow input cells
  (`255,242,204`) and the two warning reds (`150,0,0`, `192,0,0`) were left alone.
  *Reason:* the brand red is close enough that reusing it for chrome stops a warning
  reading as a warning. `ForgeRed` appears only on the Setup sheet tab, where no
  text sits on it.
- **The palette is five `Private Function`s, not `Const`s**, placed AFTER every
  module-level declaration. *Reason:* VBA rejects `Const c As Long = RGB(...)`
  ("constant expression required"), and a procedure inside the declaration block
  fails the build with "module-level declaration after a procedure".
- **`GateAudit`'s spacing check asserts nothing.** *Reason:* it reported FAIL on a
  correct build — it counted blanks before the next KEY row, but the spacer sits
  before the next BLOCK. It cannot be repaired into an assertion: with no chains
  defined the 20 chain input rows are themselves empty, so row-blankness does not
  delimit the block. Do not re-arm it.
- **Gate helpers loop column A; they never call `Range.Find`** — see below.

**Discovered (empirical — each cost a round trip, none is recorded anywhere else):**

- **`Range.Find` returns `Nothing` on a very-hidden `_Config`.** Every argument was
  named (`LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False`), so this is not the
  inherited-`LookIn` trap. Evidence: the sheet demonstrably held `k_core_layout` at
  row 11 and a plain loop found it immediately. Core's own `FindKeyRow` loops
  column A — match it.
- **`xlSheetVeryHidden` is `2`, not `-2`** (`-1` visible, `0` hidden). A gate
  written against `-2` fails the owner at item 1 on a correct build.
- **Excel's module import never replaces a same-named module.** Importing
  `MacrosBundle.bas` over an existing `MacrosBundle` adds `MacrosBundle1`,
  duplicating ~191 procedures → `Ambiguous name detected` on compile and doubled
  Alt+F8 entries. **Always remove the old module first.**
- **`Left$(s, n)` with `n > Len(s)` returns the whole string rather than raising.**
  Verified by executing the logic in **VBScript via `cscript`** across ten cases
  including `""`, `"c"` and `"core"`. VBScript is the closest locally executable
  dialect; **onecompiler.com/vb is VB.NET and cannot compile this project** — no
  `Application`, no `Worksheet`, no `CreateObject`.
- **`IMPORT / EXPORT` is `UCase()` of the package display name, composed at run
  time.** Grepping the bundle for it finds nothing; that is not evidence it is
  unemitted. A subagent reported it as a defect and was wrong.
- **The blind dry-run reads VBA; it does not execute it.** It caught 13 + 6 + 7
  authoring defects across three gates and **zero** runtime defects — both helper
  failures above sailed straight through it. It is an authoring gate, not a test.
- **`build.ps1` catches "module-level declaration after a procedure"**, a real VBA
  compile rule and the closest thing to CI here. It caught the palette placement.

**Still open — the manual gate, 8 of 15 items unrun. All need the human.**
Passed: hardening 1.1, 1.2, 2.1 · reorder 1 · reskin 1, 2, 3, plus the spacing half
of reorder 2 confirmed by eye. Both load-bearing items (#8 crash recovery, #12
fail-closed) passed. Unrun: **hardening 1.3, 2.2, 3.1, 4.1, 4.2** · **reorder 2
(clipping half), 3** · **reskin 4 (dark Excel theme)**. Two need their own build —
3.1 a hostile `PkgClean` declaring `core.legacy`, and reorder 3 the v0.6.0 bundle
from `812d88c`. **Reskin 4 is the likeliest real finding:** the palette came off a
dark-background proof and lands on an explicitly white sheet, and nothing has
checked the two together. `tools/GateAudit.bas` re-runs six mechanical checks in
one click.

Not blocking: `.superpowers/sdd/` still needs deleting **by hand** (the permission
classifier refused twice); `.vscode/` stays untracked deliberately.

**Next phase — W1, the ingest→clean path (7 actions). Approved by the owner
2026-07-28; no code written.** Scope and ids frozen in §3.4 above. Build in this
order, because the first item is what makes the other seven safe to write:

1. **Contract v2.3.0** — add the nine-domain vocabulary table to §7 (`import` ·
   `clean` · `validate` · `enrich` · `format` · `summarize` · `export` ·
   `workbook` · `backup`; reserved `chain.*` and `core.*`; grandfathered
   `importexport.*`) and the seven new ids to §8. **§7 has no vocabulary table
   today** — §3.1 assumes one exists and it does not.
2. **`build.ps1` id enforcement** — extract literal `Core_AddAction` ids, fail on
   duplicates and on any domain absent from that table. **`build.ps1` still has no
   id check at all**, exactly as the 2026-07-24 entry recorded. Until it does,
   nothing stops a wave shipping `import.combineWorksheets` — and ids are immutable
   with no alias table by design, so a wrong one is unfixable.
3. **`PkgImport.bas`** (new) — `import.combine_worksheets`,
   `import.combine_workbooks`, `import.delimited_file`; the two combines share one
   private `PkgImport_CombineEngine`.
4. **`PkgClean.bas`** (+4) — `clean.remove_blanks`, `clean.unmerge_fill`,
   `clean.standardize_formats`, `clean.text_to_numbers`; the last calls the
   existing `Core_Coerce` rather than reimplementing coercion.
5. **Gate + `GateAudit` extension** — manifest-shaped checks (every action has a
   manifest entry, an entry point, a valid domain) automate the way the colour
   checks did. Only click-every-button stays manual.

`Core_AddAction` already takes `category:=` — §3.2 shipped in track 1; do not
re-add it.

**Read manifest.** This design (22.1K, §3 is the whole track-2 spec) ·
`docs/macro-catalog-plug-and-play.md` (16.3K, the authoritative build queue) ·
`docs/macro_code_library.md` (30.8K, **new** — the owner's reference
implementation, 49 snippets; §1 and §2 cover W1. **Its snippets are test
scaffolding, not package code**: `ActiveSheet`/`Selection`, hardcoded paths marked
`' <-- set`, no speed wrapper, no error handler, none runtime-tested — converting
one is most of the work) · `spec/contract.md` (54.1K — **over the slice threshold,
never read front to back**; grep `## 7.` and `## 8.`) · `build/build.ps1` (16.6K,
for the new id checks) · `src/packages/PkgClean.bas` (11.0K, the package pattern to
copy) · `src/core/Core.bas` (229.3K — **never read whole**; grep
`docs/code-map.md` (29.3K, current at `8886fa8`) for an anchor, then read that
slice only).

### 2026-07-24 — design agreed by the owner; no code written yet

**Branch / commits.** `docs/finish-excel-design` off `main` (`0d0cccd`), one
commit **`88d7b3f`** — *"docs: design the Finish Excel phase (polish pass, then
catalog waves)"*. **Not pushed, no PR** (docs-only; nothing to verify in Excel
yet). `main` remains at v0.3.0 / contract v1.1.0.

**What shipped.** Docs only, no `src/` change: **new** this design doc (16.7KB);
`docs/macro-catalog-plug-and-play.md` rewritten with a per-item build verdict
(16.3KB, now the authoritative build queue); `TODOS.md` refreshed — its header
was stale (claimed v0.2.1 and "pending merge") and its Roadmap replaced by the
six waves. Test count unchanged: **none, by repo policy** — VBA has no headless
runner. `build/build.ps1` was not run because no `.bas` file changed.

**Decided this session.** Sequencing (polish before catalog, owner's call).
Catalog shape: every action individually listed, thin public wrappers over
private package-prefixed family engines. Two-segment action ids whose domain is a
user-facing vocabulary decoupled from the module. Per-action enable with the
collapsed-band resolution. Wave 1 = the 7-action ingest→clean path; wave 2 =
Validate & Reconcile. 34 items across six waves → 35 new actions, ~48 listed
total; 8 *Later*, 5 *Deferred*, 4 *Declined*.

**Design decisions a next session might "fix" wrongly — DO NOT REVERT:**

- **Setup collapses settings for disabled actions; it does NOT omit them** (§2.3).
  *Reason:* `Core_BuildSetup` deletes and re-seeds Setup from the **finalized**
  `_Config`, so omit-unless-enabled costs toggle→Finalize→rebuild→fill→Finalize
  and discards anything typed but not finalized. This deliberately overrides all
  three `/deliberate` panelists.
- **Action ids stay two segments, `<domain>.<operation>`** — do not "improve" to
  `<domain>.<object>.<operation>`. *Reason:* the user types the id by hand into a
  Setup cell as `Label [action.id]`; brevity is a correctness property here, not
  a preference.
- **The domain is NOT the module name.** *Reason:* encoding the module in a
  permanent dispatch key makes a future package split a major contract break.
  This is why new `import.*` actions may live in a different module from
  `importexport.import` and still share one Home card.
- **The eight shipped ids are grandfathered literally, including the two
  camelCase ones** (`importexport.updateMaster`, `importexport.suggestTypes`). Do
  not normalize their casing, do not split the `importexport` domain, and **do
  not add an alias table** — an alias creates two permanent names for one action
  and would have to be ported to Apps Script.
- **Wave 1 is the coherent ingest→clean path, not a cross-category sampler.**
  *Reason:* owner's call after the panel split 2-1 the other way; chains make a
  related wave immediately useful as one button.
- **No search, and no Recent Actions ever.** *Reason:* Recent Actions needs
  runtime writes that the next Finalize overwrites — it is architecturally
  broken, not merely unpolished. Favorites is deferred but remains additive.

**Discovered (empirical — read Core before planning, these changed the design):**

- **`RenderPlacementBlock` (`Core.bas:1242`) is already a per-action catalog
  row** — it writes `Caption [id]`, the id hidden in `RID_COL`, and one yellow
  editable cell per action. So per-action Enabled is *one more column on an
  existing block*, not a new mechanism.
- **`RenderPage` (`Core.bas:1734`) already draws a page title and column
  headers**, so design §8's "per-page titles, column headings" are partly built.
  The real gaps: grouped pages carry **no descriptions at all** (only
  `BuildSinglePanel:1616` writes them, into column C), titles are meaningless
  (`"Macros - Page 1"` / `"Column A"`), no legend, no nav, no auto-layout.
- **`GatherButtons` (`Core.bas:1549`) is the single funnel** for both actions and
  chains — disabled-action filtering lands in exactly one place.
- **`Core_BuildSetup` (`Core.bas:825`) deletes and rebuilds the Setup sheet**,
  re-seeding from the finalized `_Config`. This is the constraint behind §2.3.
- **`build.ps1` has NO action-id check today** — only duplicate *procedure* names
  (line 91). Both the duplicate-id and unknown-domain checks are new work.
- **Shipped action ids, in full:** `importexport.clear`, `importexport.import`,
  `importexport.updateMaster`, `importexport.export`,
  `importexport.suggestTypes`, `clean.dupes`, `clean.trim`, `clean.filldown`.
- **`/deliberate` on this machine (in addition to the known stdin rule):** pass
  the **full** model id (`gpt-5.6-terra` / `gpt-5.6-sol`); codex prints its answer
  **twice** in the captured output, so strip the preamble by cutting after the
  **last** `^codex$` line; `node` importing a script from the `x:` drive needs a
  `file:///x:/...` URL, not `x:/...`; and **PowerShell here-strings (`@'…'@`)
  break in the Bash tool** — use a heredoc (this created a stray empty file
  named `clean`, since removed).
- **`docs/code-map.md` is still absent** and `Core.bas` is 76.8KB / ~125 procs.

**Still open.** Nothing blocking, and **nothing needs the human right now** — the
design is owner-approved. The Excel manual-verification gate (import + Debug ▸
Compile VBAProject + click every action, cancel path and error path) arrives at
the *end* of the polish implementation and only the owner can run it. Deferred
non-blocking Minors from the previous phase still sit in
`.superpowers/sdd/progress.md`.

**Next phase — `writing-plans` for track 1** (`v0.4.0`, Core-only, contract
bumped to v1.2.0). Track 2 gets its own plan per wave. Consider `/code-map`
first: the pass touches the Setup renderer, the panel builder and
`GatherButtons`, all inside a 77KB file.

**Read manifest.** This design (16.7KB, §2 is the whole track-1 spec) ·
`docs/macro-catalog-plug-and-play.md` (16.3KB, the build queue) ·
`spec/contract.md` (19.4KB, v1.1.0 → to be bumped to v1.2.0) · `TODOS.md`
(6.4KB) · `build/build.ps1` (8.1KB, for the new id checks) · `src/core/Core.bas`
(76.8KB — **read in slices via Grep, never whole**): `Core_AddField:88`,
`Core_AddAction:125`, `Core_BuildSetup:820`, `RenderTier:923`,
`RenderPlacementBlock:1242`, `ReadCodeMap:1269`, `Core_FinalizeSetup:1419`,
`Core_BuildControlPanel:1517`, `GatherButtons:1549`, `BuildSinglePanel:1579`,
`BuildGroupedPanels:1636`, `AddToPage:1723`, `RenderPage:1734`,
`DeletePanels:1783` · `src/packages/PkgClean.bas` (6.8KB, the package pattern to
copy for wave 1).

**Still true and easy to trip over.** No VBA compiler in CI — `build.ps1` checks
block balance, duplicate proc names and module-level declarations after the first
procedure; `scripts/vba-syntax-check.mjs` is a parser only, missing undefined
vars and type errors. The real gate is a human importing the bundle, running
Debug ▸ Compile VBAProject, and clicking. The do-not-revert lists in
`docs/designs/2026-07-24-compose-your-buttons-plan.md` (§ Handoff log) still
apply in full — in particular: module-level `Const` stays in the declarations
section, `_Config` is unhidden before `.Delete`, and admin controls stay on the
`Setup Menu` tab.
