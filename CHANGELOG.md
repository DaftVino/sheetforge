<img src="docs/assets/sheetforge-logo.png" alt="sheetforge" width="200">

# Changelog

All notable changes to this project are documented here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versioning: [SemVer](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.9.0] - 2026-07-27

### Changed

- **The generated sheets now look like sheetforge.** The Setup tab and the Macros pages
  carried two unrelated colour schemes that had drifted apart — purple on one, blue on the
  other — neither of which was the product's identity. Both are replaced by the steel-and-red
  palette from the identity proof: a red Setup tab strip, deep-navy headings, steel-grey
  help text, and a pale steel band fill. The Macros tab strip is steel rather than red, so
  the two surfaces stay distinguishable at a glance.

  **Colours that mean something were left alone.** The yellow input cells still mean "fill
  this in", and the two warning reds still mean "careful" and "this is broken". Those are
  not decoration, and the brand red is close enough to them that reusing it for chrome
  would make a warning stop reading as a warning — so the brand red appears only on the
  Setup tab strip, where no text sits on it.

  Presentation only: additive under contract §10's two-axes rule, and no stored
  configuration changes meaning.

- **The logo is on the README and this changelog.** It stays out of the build. A `.bas`
  cannot carry a PNG, and the alternatives — shipping a second file, or base64-encoding a
  2 MB asset into a 270 KB bundle — would each break the one-file-import model that is the
  whole delivery story. The README copy is resized to 38 KB so neither page pulls 2 MB.

## [0.8.0] - 2026-07-27

### Changed

- **The Action catalog is now the first thing on the Setup tab, and Chains comes second.**
  The tab opens on the block you have to act on: since v0.6.0 every action arrives `Off`,
  so nothing runs until you set an `Availability` cell — and until now that block sat
  below another package's export paths and file prefixes. Chains moved below the catalog
  for the same reason: a chain is built out of actions that are already switched on, so
  being asked to compose one first inverted the order you actually work in. The grey help
  line above the block now says so, instead of leading with Chains.

  **Nothing about your saved configuration changes.** This is presentation only. Every
  reader finds its block by name rather than by position, so a configuration saved by an
  earlier version is read exactly as before; the tab is simply re-drawn in the new order
  the next time you click **Create Setup Tab**.

  Completes the layout half of the Setup-surface redesign, whose default-`Off` and
  appear-when-used halves shipped in v0.6.0.

## [0.7.0] - 2026-07-27

### Fixed

- **Finalize can no longer lose your entire configuration if Excel dies at the wrong
  moment.** Activating used to delete the saved configuration *before* copying the new
  one into its place, leaving a moment in which the workbook held no configuration at
  all. A crash, a hang you force-quit, or any failure inside that copy destroyed
  everything you had set up — every setting, every button placement, every Availability
  choice — with no undo and no warning. Worse, the next **Create Setup Tab** read the
  hole as "never configured" and rebuilt defaults over it, so the loss was silent.
  Activating now copies first, checks the copy is complete, and only then replaces the
  old one. If anything goes wrong before that check, your saved configuration is
  untouched and a message says so. If a crash interrupts the swap itself, the next
  action recovers the configuration automatically instead of discarding it.
  **This is not a claim that activating is now atomic** — Excel offers no such
  guarantee, and the remaining gap is documented rather than papered over. (#8)
- **A hand-edited configuration can no longer talk its way onto the legacy path.** The
  engine decides whether a saved configuration predates the Availability column by
  reading a hidden format marker. When that marker was missing, it used to settle the
  question by reading a *visible column heading* back — so renaming that heading by hand
  was enough to make a current, tampered configuration look ancient, which activated
  every action as a live button, destructive ones included. The decision is now made
  from structure the newer format always writes and the older one never did, and an
  unexplained missing marker is refused rather than trusted. (#12)
- **A package can no longer take over an engine-reserved action id.** The contract has
  reserved the `core.*` namespace since v0.6.0, but Finalize never enforced it the way it
  enforces `chain.*`. A package declaring `core.legacy` would have collided with the
  engine's own provenance flag and quietly turned every action you had not explicitly
  switched on back into a button — reverting v0.6.0's Off default on that deployment.
  Finalize now refuses it by name. (#13)

### Changed

- **Contract v2.2.0** (additive). Staged-then-swap activation and orphan adoption are now
  normative (§5); deciding a format marker is "absent" requires positive structural
  evidence and may not rest on a user-editable caption (§5); refusing a package action id
  in `core.*` is normative rather than declared (§7); `_Config` and `_Config_staged` are
  reserved sheet names (§7). No stored value changes meaning and `SCHEMA_VER` stays
  `"v2.0"`, so every v0.5.0 and v0.6.0 configuration remains readable.
- **Manual checklists now pass a blind dry-run before they reach the owner.** A subagent
  holding the checklist and the built bundle, but forbidden from reading `src/`, flags
  every item that cannot be executed, cannot be judged, or would pass on a broken build.
  Added to the pipeline in `CLAUDE.md` along with an authoring rule. The v0.4.0 gate lost
  six items mid-run to this and the v0.6.0 gate still lost two; the reviews all passed
  them, because every review asked "does this match the code?" and it did. (#3)

## [0.6.0] - 2026-07-27

### Changed

- **BREAKING — an action now arrives switched off, and you turn it on.** A fresh
  install shows the Action catalog with every `Availability` cell reading `Off` and
  no buttons drawn, instead of every button at once. Set the ones you want to
  `Button`, then click **`Finalize Setup Tab`** to activate them — `Create Setup Tab`
  rebuilds from the last activated configuration and would discard the edits you
  just made. The old default handed a first-time user
  eight live buttons — several of them destructive — before they had read what any
  of them did. **Your existing deployments are unaffected**: a default only applies
  where nothing was ever stored, so every choice already saved still means exactly
  what it said. A configuration saved by v0.5.0 keeps all of it verbatim, and a
  pre-v0.4.0 configuration — which predates the Availability column entirely — still
  reads as everything on.
- **A newly-registered action arrives off on an existing deployment too.** When a
  package upgrade adds an action your saved config has never seen, it appears in the
  catalog switched `Off` rather than turning itself into a live button on your next
  rebuild.
- **Settings now follow the actions that use them.** Turn off every `Import / Export`
  action and the whole `IMPORT / EXPORT` band moves into
  `SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON` at the bottom — collapsed, kept, and
  still saved. Turn off just `Export Dashboard` and only its four export settings move;
  the master-list settings stay put. Nothing you typed is lost either way: set the
  action back to `Button` and its settings come back up. This is what
  `Clean & Standardize` losing its band in v0.5.0 was the first half of.
- Contract **v2.1.0** (additive): `availability` defaults to `off` (§6); the
  pre-v1.2.0 legacy default is pinned to `button` independently of it (§5), so moving
  one cannot move the other; reserves the `core.*` action-id namespace (§7); restates
  relegation presentation (§3).
- **Configurations saved by v0.5.0 remain readable.** The stored format marker does
  **not** move — `SCHEMA_VER` stays `"v2.0"`. Nothing about your stored data is being
  reinterpreted, so there is nothing to refuse.

### Fixed

- **A damaged `Availability` value is no longer quietly turned back into a live
  button** (#10). Typing something the dropdown does not offer — `maybe`, or `of` for
  `off` — used to block activation once, then get rewritten as `Button` by the next
  `Create Setup Tab`, which destroyed the evidence so the block never fired again. The
  cell now comes back blank and keeps blocking until you fix it. A **blank** cell blocks
  too: it used to read as `Button`, which was defensible while `Button` was the default,
  but with the default at `Off` either reading is wrong and neither is visible in the
  result. The message reads
  `MISSING or UNRECOGNISED Availability (use Button, Chains only, or Off)`.

## [0.5.0] - 2026-07-27

### Changed

- **BREAKING — `Remove Duplicates` now asks, instead of remembering.** The two
  settings it used to read from the Setup tab — which column(s) decide a duplicate,
  and whether the selected range's first row is a header — are now asked at run time,
  after you pick the range and naming it. Both describe *the range you just picked*,
  not the deployment: the key columns are numbered from the left of a selection that
  does not exist until you make it. Storing them meant editing Setup and re-Finalizing
  between two runs of the same button, and nobody discovered the setting at the moment
  they acted. Each prompt cancels cleanly, an empty column answer means "use the
  default", and a column outside the selected range is refused before anything is
  edited. The completion message states which header interpretation was applied,
  because that is not recoverable from the result.
- **A package with nothing to configure no longer draws an empty band on Setup.**
  `Clean & Standardize` now has no settings at all, and so has no section — its
  buttons are unaffected. The same applies to any package whose every setting has
  relegated to the "actions you haven't turned on" band.
- Contract **v2.0.0** (breaking): removes the `clean_keycols` and `clean_hasheader`
  fields, changes `clean.dupes` semantics, and adds two normative rules — a setting
  describing the caller's selection must be a run-time input rather than a config
  field, and a package with no visible fields must render no band.
- **Configurations saved by v0.4.0 are not readable by this version.** The stored
  format marker moves to `v2.0` with the contract, so `Create Setup Tab` refuses a
  v0.4.0 `_Config` and asks you to re-import that version's bundle rather than
  rebuilding over it. Delete the hidden `_Config` sheet to start fresh.

## [0.4.0] - 2026-07-27

### Added

- **Anchored placements — put a button on one of your own sheets.** A placement code
  of the form `Sheet!Cell` (e.g. `TabName!B4`) draws that button on your tab at that cell
  instead of on a generated page. `$B$4`, unquoted names containing spaces, and the
  quoted `'TabName Data'!B4` form are all accepted. The button uses `xlMove`, so it travels
  with its cell and is never resized by a narrow column. An anchored button is not on a
  generated page and never affects page packing, column caps or category splitting.
  Finalize **blocks** on a malformed code, an engine-owned target (`Macros 1`), a
  missing sheet, or two actions aimed at one cell — including the same cell spelled two
  ways; a valid anchor is reported in a soft `Placed on your own sheets` list. A
  placement that fails at build time (a protected sheet, a target deleted after
  activation) reports a count and a repair instruction instead of taking down the build.
- **Three-state per-action availability.** The placement block is now the **Action
  catalog** (`Action | Availability | Placement`) with a validated dropdown: `Button`
  (shown and runnable), `Chains only` (no button, runs only as a chain step), `Off`
  (no button, runs nowhere). `Off` means off — the dispatcher refuses it and Finalize
  blocks activation for any chain that uses one. An unrecognised value blocks Finalize
  rather than defaulting to `Button`.
- **`usedBy` on a field** — settings belonging only to `Off` actions collect into one
  collapsed *"settings for actions you haven't turned on"* band at the bottom of Setup.
  Relegated, never omitted: Setup re-seeds from the finalized `_Config`, so omitting
  them would cost a second Finalize and discard un-finalized typing. A `Chains only`
  action still runs, so it keeps its settings in the main list.
- **`category` on an action** — optional, defaults to the package name, and decoupled
  from the module, so two packages can feed one category and a future package split
  cannot move a button in the user's mental model.
- **Capacity-packed auto-layout.** With no placement codes filled in, buttons
  distribute automatically across numbered pages: up to three category-headed columns
  per page, up to 6 buttons per column, small categories sharing a page, and a category
  never split across tabs unless it needs more columns than a page holds. Filling in
  even one code still hands total control back to the codes.
- **Legible descriptions.** Button pages are an **unmerged cell-card grid** — the
  button anchored to a content cell, its description in the cell beneath, that row
  auto-fitted. Grouped pages previously had no descriptions at all, and merged cells
  (which Excel's AutoFit ignores) are forbidden on the grid so nothing clips.
- **A machine-owned config schema marker**, with **fail-closed** migration. A v0.3.0
  config is detected and its page codes read from the column they are actually in;
  every recovered code is validated; on any conflict `Create Setup Tab` refuses with a
  report instead of rebuilding. A marker written by a newer major version refuses
  outright rather than offering a rebuild, and a damaged one asks first, spelling out
  what would be discarded and what kept, defaulting to **No**.
- **`Does the range you pick include a header row?`** on Clean — the contract's first
  use of the `bool` kind, so its `yes`/`no` encoding is now fixed. A blank cell reads as
  the field's declared default; an unrecognised value blocks Finalize rather than
  silently reading as `No` and deleting a header row.
- Contract frozen at **v1.2.0** (additive).
- **Build / tooling:** a reserved-word gate rejects a VBA keyword used as an identifier.
  The list is empirical — a word is added when Excel actually rejects one, never on
  suspicion. Fixture and harness packages are kept out of the default bundle.

### Changed

- The placement column is labelled **`Placement`** rather than `Page code`, since a
  code can now be a page code, an anchor, or `hide`.

### Removed

- **The single `Macros` tab default.** With no placement codes, v0.3.0 built one tab
  and stacked every button on it; v0.4.0 packs categories onto numbered pages instead,
  and the stale `Macros` tab is deleted on rebuild. Manual page codes are unaffected.

### Fixed

- Every range named in a Clean completion message is **sheet-qualified** — `TabName!M2:M12`,
  and `'TabName Data'!M2:M12` when the name needs quoting — so a message can no longer read
  as though the action ran on whichever sheet you happen to be looking at.
- **A failed button placement no longer leaves an orphan shape.** `Buttons.Add` creates
  the shape before it can be named, so a failure part-way through placement left a
  button carrying Excel's default name, no caption and no `OnAction` — and permanently
  unremovable, because the sweep matches on the engine's name prefix. Placement is now
  all-or-nothing: the partial shape is deleted and the original error re-raised.
- **Finalize no longer claims the Macros pages were rebuilt when none were built.**
  Every button hidden, switched `Off`, or anchored is a legitimate configuration, and
  the dialog now says so instead of sending you looking for tabs that were never meant
  to exist.
- Instructions that told you to run `Core_Run` or `Core_BuildControlPanel` from **Alt+F8**
  are corrected throughout. Excel's Macro dialog lists only `Public Sub`s taking no
  arguments, and an *optional* argument is enough to hide one; use the VBE Immediate
  window (Alt+F11, Ctrl+G).

## [0.3.0] - 2026-07-24

### Added

- **Compose Your Buttons** — a Core-owned Setup section with two capabilities: **Chains** (compose several actions into one button that runs them in order — steps authored as `Label [action.id]` tokens; each chain is an immutable `chain.<id>` with a freely-editable label; on a step failure the runner prompts Continue/Stop; per-step success dialogs are suppressed and one summary is shown) and **Grouping/layout** (assign each button a page/column code like `1A`; no codes → the single "Macros" tab as before; any code → numbered "Macros N" pages, with blank-coded buttons in page 1's "Unassigned" area and `hide` to omit). Finalize validates chains + placement and blocks activation on structural errors (unknown steps, bad codes, reserved-namespace or id collisions).
- **Setup Menu tab** — the Create Setup Tab / Finalize Setup Tab controls now live on their own dedicated tab (built by `Core_BuildControlPanel`) so they can be hidden when not in use; rebuilding reuses the tab and preserves its hidden state.
- **Finalize rebuilds the control panel** — activating a configuration now rebuilds the Macros tab automatically, so authored chains and grouping appear immediately (no manual `Core_BuildControlPanel` run).
- Contract frozen at **v1.1.0** (additive): the `chain` config entity, reserved `chain.*` namespace, engine-owned config, placement code, and the dispatch/quiet convention.
- Package completion dialogs now route through `Core_Done` so they suppress inside a chain.
- **Build / tooling:** `build/build.ps1` now fails on a module-level declaration placed after the first procedure (a VBA compile error the block-balance check could not see); `scripts/vba-syntax-check.mjs` drives the "VBA Pro" language server headlessly for a fast pre-Excel syntax gate.

### Changed

- The Create / Finalize admin controls moved off the Macros tab (and out of the grouped page-1 "Setup" column) onto the dedicated, hideable **Setup Menu** tab.

### Fixed

- Re-running **Finalize** no longer errors *"That name is already taken."* and no longer leaves orphan `Setup (2)`/`(3)` copies: the very-hidden `_Config` is unhidden before it is deleted, so the copy→rename cannot collide. (Pre-existing bug; surfaced by the multi-finalize Compose workflow.)
- `DeletePanels` now guarantees `Application.DisplayAlerts` is restored even if a sheet delete raises, so a leaked `DisplayAlerts = False` can no longer suppress later confirmation prompts.

### Removed

- The standalone `MacroLayout.bas` grouping prototype (its logic is absorbed into Core).

## [0.2.1] - 2026-07-23

### Fixed

- Setup tinted band rows (section headers + list/table label & explanation) now merge across columns B:E and wrap to the full band width instead of being cramped in column B. Merged-row heights are set explicitly since AutoFit ignores merged cells.
- A section header identical to its package banner is no longer rendered (removed the duplicate "CLEAN & STANDARDIZE" in PkgClean; Core now skips any section whose text equals the package banner).

## [0.2.0] - 2026-07-23

### Added

- **Two-tier Setup**: manifest fields declare a tier ("simple" default / "advanced"); the Setup renderer collects advanced fields under one "ADVANCED - most users can leave these as-is" band per package, so regular users can stop after the simple fields. An unknown/mis-cased tier degrades to "simple" (never hidden).
- **Suggest Import Types** helper: a button that reads the live Setup import table, opens each matched source file, samples it, and fills a suggested per-column type map (T/N/D) for review (never applied silently; ID/code columns flagged for the user to keep as text). Adds two reusable Core setup-helper services, `Core_SetupTable` / `Core_SetupSetColumn`, that read/write the live Setup sheet before Finalize.
- **Per-field help notes**: manifest fields gain an optional plain-language `help` string, rendered grey/italic on the Setup sheet.

### Changed

- Setup labels and table headers reworded for non-technical users ("Input tab (in this workbook)", "Source file (name starts with...)", "Column types"); the import note now explains the same-folder rule, the ignored `(1)`/`(2)` copy suffix, newest-file logic, and the full type-map legend.
- Macro-tab button descriptions span two rows (no longer clipped); list/table label + help render in the tinted band above the column headers so the headers sit directly on the input cells.

## [0.1.0] - 2026-07-23

### Added

- Initial scaffold (office-automation profile, source-first VBA workflow per ADR 0001).
- **Core engine** (`src/core/Core.bas`): manifest-building API (`Core_DeclarePackage`/`AddSection`/`AddField`/`AddList`/`AddTable`/`AddDefault`/`AddAction`); config readers with graceful manifest-default fallback (`Core_Cfg`/`CfgLong`/`CfgList`/`CfgTable`); shared helpers (depth-counted speed wrapper, robust + bare last-row/col, file finder, header-column finder, type coercion, structured error/log); zero-trust package discovery via `Application.Run` (no "Trust access to the VBA project object model"); Setup-sheet renderer, control-panel/button builder, and the `Core_Run "<actionId>"` dispatcher.
- **PkgImportExport** package (`src/packages/PkgImportExport.bas`): the Sample Dashboard's four buttons - Clear Input Tabs, Import Source Data, Update Ref_Master, Export Dashboard - re-expressed as a thin package on Core with no behaviour change (Import now shows one dialog instead of two).
- **PkgClean** package (`src/packages/PkgClean.bas`): Remove Duplicates, Trim & Clean Text, Fill Blanks Down, driven by an ad-hoc range picker; preserves formula cells and rejects multi-area selections.
- **Build** (`build/build.ps1`): concatenates `src/core/Core.bas` + `src/packages/Pkg*.bas` into a single-module `dist/MacrosBundle.bas`; fails on a duplicate procedure name across modules or a Sub/Function block imbalance. `dist/` is a gitignored build target.
- `docs/designs/setup-simplicity.md`: planned (deferred) two-tier Simple/Advanced Setup and a suggest-column-types helper.
