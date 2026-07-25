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

### 2026-07-24 — design agreed, plan not yet written

**Decided this session.** Sequencing (polish before catalog). Catalog shape (all
actions individually listed, thin wrappers over private family engines).
Two-segment ids with a domain vocabulary decoupled from module names. Per-action
enable with the collapsed-band resolution. Wave 1 = the 7-action ingest→clean
path. Wave 2 = Validate & Reconcile. Target ~35–40 of 58, verdicts annotated into
the catalog doc.

**Next.** `writing-plans` for track 1 (`v0.4.0`, Core-only + contract v1.2.0).
Track 2 gets its own plan per wave.

**Read manifest for the next session.** This design · `spec/contract.md` (v1.1.0,
to be bumped) · `docs/macro-catalog-plug-and-play.md` (now annotated with build
verdicts) · `src/core/Core.bas` in slices — `Core_BuildSetup:820`,
`RenderPlacementBlock:1242`, `ReadCodeMap:1269`, `Core_BuildControlPanel:1517`,
`GatherButtons:1549`, `BuildSinglePanel:1579`, `BuildGroupedPanels:1636`,
`RenderPage:1734` · `src/packages/PkgClean.bas` (the package pattern to copy).

**Still true and easy to trip over.** No VBA compiler in CI — `build.ps1` checks
block balance, duplicate proc names and module-level declarations after the first
procedure; `scripts/vba-syntax-check.mjs` is a parser only. The real gate is a
human importing the bundle, running Debug ▸ Compile VBAProject, and clicking. The
do-not-revert lists in
`docs/designs/2026-07-24-compose-your-buttons-plan.md` (§ Handoff log) still
apply in full — in particular: module-level `Const` stays in the declarations
section, `_Config` is unhidden before `.Delete`, and admin controls stay on the
`Setup Menu` tab.
