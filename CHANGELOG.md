<img src="docs/assets/sheetforge-logo.png" alt="sheetforge" width="200">

# Changelog

All notable changes to this project are documented here. Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versioning: [SemVer](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.17.0] - 2026-07-28

### Added

- **Five housekeeping actions**: unhide every worksheet, protect and unprotect every sheet
  (two separate buttons, not one toggle), build a table of contents, and reset hidden rows and
  columns.

- **Unhide All Worksheets leaves the toolkit's own sheets hidden.** `_Config` holds your
  settings and is hidden on purpose; unhiding it puts engine internals in front of you and
  invites an edit that breaks the Setup tab. The most-downloaded version of this macro on the
  internet unhides everything.

- **Protect and Unprotect are two buttons because one toggle hides which way it will go.** A
  single button on a workbook you have not looked at in a month is a guess.

  **Protecting deliberately leaves drawing objects unprotected.** Excel's default protects
  them, and the toolkit's buttons are drawing objects - so the default would leave the engine
  unable to remove its own buttons on the next rebuild, permanently, with no way back that
  does not involve changing protection you set. Unprotect reports which sheets refused the
  password rather than stopping halfway with some sheets left open.

- **The Table of Contents is create-only.** If a sheet by that name exists it stops and says
  so, rather than clearing it and rebuilding. An earlier design recognised its own output by a
  sentence written in A1 - but a sentence you could type yourself is not proof the toolkit
  wrote the sheet, and the price of getting that wrong is a cleared sheet. Rebuilding means
  deleting the old one yourself, which is the trade.

### Changed

- **Contract v2.9.0.** Adds the `workbook` package, and a normative rule that an
  implementation creating a new destination sheet must refuse a name the engine reserves - not
  only one that already exists.

## [0.16.0] - 2026-07-28

### Added

- **Four ways to keep a copy before something goes wrong.** A timestamped copy into a Backup
  folder, a snapshot taken before a risky run, a month-end named version for finance packs, and
  moving a finished output into a dated archive folder.

- **Snapshot is the one to know about, and it is worth being precise about what it does.** It
  saves a full copy of the workbook, named and dated, into a Snapshots folder. **It does not
  restore.** There is no undo button in this release, and the catalog items that are waiting on
  one are still waiting - only "hide all sheets except the active one" was ever blocked on
  undo; the other two are blocked on not having a preview. Saying otherwise would be selling
  you a safety net that is not there.

- **Old snapshots are pruned, and the pruning checks its own handwriting.** It keeps the newest
  few and deletes the rest - but only files matching the complete name this action writes:
  the prefix, a real timestamp, and the workbook's own extension. **A file merely starting
  `SNAPSHOT_` is left alone.** Matching the prefix and deleting the result would have made your
  own `SNAPSHOT_notes.txt` collateral.

- **Archiving a file moves it, which is the only thing in this toolkit that does.** It refuses
  if the destination exists, and refuses if the file is open in this Excel. Across drives,
  Windows implements a move as copy-then-delete, so it can half-fail - **so it checks
  afterwards which of the two files exist and tells you**, rather than reporting a generic
  error and leaving you to work out whether your file still exists.

### Changed

- **Contract v2.8.0.** Adds the `backup` package. Additive: `SCHEMA_VER` stays `"v2.0"`.

## [0.15.0] - 2026-07-28

### Added

- **A join done in code, not with lookup formulas.** Point it at a target range and a
  reference sheet, name the key column in each, list the columns to pull across, and it
  writes them in. It builds an index of the reference table once and looks each key up
  directly, so it does not slow down as the data grows the way a sheet full of VLOOKUPs does.

  **Two things it does deliberately, because both are how a join quietly produces wrong
  numbers.** Keys are compared after trimming outside whitespace and nothing else - an
  earlier version also collapsed spaces *inside* the key, which can map two genuinely
  different identifiers onto one another. And when the reference table has the same key more
  than once, it reports **how many distinct keys were duplicated**, not how many extra rows it
  saw; it takes the first and tells you, rather than picking one silently.

- **Add mapped columns from a reference sheet** shares that same engine, driven from Setup
  rather than asked each run.

- **Refresh all pivots and Force full recalculation are now their own buttons.** Both were
  locked inside the import action, so the only way to refresh a pivot was to re-import
  everything. As standalone actions they can also be chain steps.

  **Refresh reports a pivot it could not refresh rather than counting it as done.** A pivot
  whose source has gone raises, and the old inline version could not tell you which one.

- **Build a summary pivot** from a range, and **Freeze a pivot to values** for distribution.
  Both refuse a source range with a blank or duplicated header, because a pivot built on one
  fails later with an error that points nowhere near the cause. If either fails partway, it
  removes the half-built sheet rather than leaving it behind, and still reports the original
  error rather than the cleanup.

### Changed

- **Contract v2.7.0.** Adds the `summarize` and `enrich` packages. Additive: `SCHEMA_VER`
  stays `"v2.0"`.

## [0.14.0] - 2026-07-28

### Added

- **Eight actions that make a finished thing presentable and then get it out of Excel.**
  Five formatting: rules-driven highlighting, header styling with autofit, a report title
  block, a house table style, and print layout for board packs. Three exporting: freeze
  formulas to values, PDF a range or sheet or workbook, and explode a workbook into one file
  per tab.

- **Convert Formulas to Values is now its own button.** It was buried inside Export Dashboard,
  where you could only get it by exporting the whole thing. It handles dynamic-array spills
  and legacy CSE arrays, not just ordinary formulas.

  **The export and the new action now run the same code**, rather than a second copy of the
  same thirty lines drifting apart from the original. It reports how many formulas it actually
  converted, and says so when some could not be - the version inside the export swallowed
  those silently.

- **Exporting refuses rather than overwrites.** Neither PDF action will write over a file that
  already exists; it stops and names the file. **Save each sheet as its own file checks the
  whole batch before writing anything**, including whether two of your sheet names would
  produce the same file name after illegal characters are replaced - two sheets called `A"B`
  and `A<B` both become `A-B`, and the second would have silently overwritten the first.

  It also skips the toolkit's own tabs, so you do not get `Setup.pdf` and `_Config.pdf`.

### Changed

- **Contract v2.6.0.** Adds the `format` and `export` packages. Additive: `SCHEMA_VER` stays
  `"v2.0"`.

- **Add a Report Title inserts rows, and the contract now says so.** The `format` domain
  promises to change no value, and this action writes a title string and pushes everything
  down - which can change what `ROW()` and `INDIRECT()` return. That was previously claimed to
  be no change at all. It is a change, it is stated, and the action asks before doing it.

## [0.13.0] - 2026-07-28

### Added

- **Five actions that check your data and tell you what is wrong with it, without changing
  any of it.** This is the highest-value group in the catalog for a multi-entity close, and
  it consumes the previous wave's output directly: combine, then clean, then reconcile is one
  continuous chain.

  **Reconcile Two Datasets** is the flagship. Pick range A, pick range B, name the key column
  in each, and it reports every key that is in one and not the other, and every column that
  differs on a key present in both. It matches on the key, never on position.

  **It refuses rather than guesses in three places, and each one is a wrong number it
  declines to produce.** A key that appears more than once on either side is reported as a
  duplicate with both counts, and its rows are never paired up - pairing same-key rows by
  position is how a reconciliation invents a difference that does not exist. A row whose key
  is entirely blank is reported on its own, and two blank keys are never treated as matching
  each other. And a column header that appears twice in either range stops the whole run
  before anything is compared, naming the header - silently comparing the first one and
  dropping the rest would produce a report that looks complete and is not.

- **Scan for Data Issues** sweeps a range against rules you configure - blank, number, date,
  or a numeric range - and writes what failed to an exceptions sheet.

- **Three actions that mark problems in place**: duplicates, values outside a bound, and
  blanks in columns you have marked mandatory.

  **They colour cells, and colouring is the whole safety story, so it is worth being precise
  about.** A cell that already has a fill of its own is left alone and counted separately -
  never overwritten. So every cell these actions touch is guaranteed to have had no fill
  beforehand, which makes "clear the fill on those cells" a complete undo that disturbs
  nothing else. Afterwards the marked cells are left selected, so the undo is one click and
  applies to exactly what changed rather than to the whole range you picked.

  **Each shows you the count and waits for a yes** before colouring anything.

- **Neither of the two new rule tables ships with example rows**, deliberately. A demo row
  would mean a deployment's first run colours real cells against a bound nobody set, before
  anyone has opened Setup. The worked examples live in the field help text instead.

### Changed

- **Contract v2.5.0.** Adds the `validate` package - four fields and five actions. Additive:
  no existing stored value changes meaning, and `SCHEMA_VER` stays `"v2.0"`.

## [0.12.0] - 2026-07-28

### Changed

- **Almost nothing changes for you in this release, and that is the point.** It moves fifteen
  helpers that were trapped inside one package each into the shared engine, so the next five
  waves of actions build on one copy instead of writing a sixth.

  The bundle is a single module, so a helper sitting privately inside `PkgClean` did not
  merely fail to be reusable - it *reserved that name* against every other package. The
  second package wanting a range picker could neither copy it nor declare its own.

- **One dialog stops lying about which action it belongs to.** The "does this range include a
  header row?" question was captioned **Remove Duplicates** no matter what was running. It had
  exactly one caller, so it had never shown; the next wave gives it three.

### Fixed

- **Creating a sheet no longer half-succeeds.** Two cases that used to slip through:

  A destination name already taken by a **chart sheet** passed validation, because the check
  only looked at worksheets. The rename then failed *after* an empty sheet had already been
  added - an engine error and a stray `Sheet1` from an action whose whole promise is that it
  changes nothing when it refuses.

  A destination named `Macros 3`, or any other name the toolkit reserves for itself, also
  passed - because no such sheet existed *yet*. It would build, fill with a month of combined
  data, and be deleted without a prompt by the next Finalize. Both are refused now, by name,
  with the reason.

### Added

- **The build now fails an unprefixed package helper.** Same reasoning as the ids gate: the
  cheapest time to catch a name that will collide is before it exists. Verified by deliberately
  injecting one and watching the build reject it.

## [0.11.0] - 2026-07-28

### Fixed

- **A chain no longer reports success for work it did not do.** This is the whole release.

  When a step *refuses* — it checked, found the world not ready, and stopped having changed
  nothing — the chain now stops there and tells you which step refused, why, and how many
  steps did not run. Before, it asked whether to continue, and answering yes ran every
  later step against whatever was already on the sheet.

  That was not a hypothetical. **Combine All Worksheets refuses on every run after the
  first**, because the sheet it creates now exists. So a daily chain of combine → clean →
  clean → clean showed a dialog, cleaned *last* period's numbers, and reported `4/4 ok`.
  A wrong number in a finance deliverable with a tick beside it is the worst thing this
  toolkit could produce, and it is the thing this release removes.

  A step that genuinely *fails* still asks. The difference is deliberate: a failure may be
  incidental to one step and you are placed to judge, but a refusal means that step's output
  does not exist, so everything after it would be working on stale data.

- **Refusals stop showing a fake error number.** A refusal is not an error, so the summary
  line no longer reads `X #0 A sheet named 'Master' already exists` — which told you your
  macro had malfunctioned when it had done exactly the right thing.

- **The Import / Export actions were not reporting their refusals at all.** All five of them
  showed a message and returned without telling the engine anything had gone wrong, so inside
  a chain each was ticked green beside a step that did nothing. They now report like every
  other action. If you have a chain containing one of these, it will now stop where it
  previously carried on.

- **Two refusal messages now name what they are talking about on their first line** — the
  folder that was not found, and how many column headers were missing. The chain summary
  quotes only the first line, so "This folder was not found:" with the path underneath told
  you nothing useful.

### Changed

- **Contract v2.4.0.** A refusal is now a distinct outcome from a failure, alongside success;
  an implementation must expose a single call that both records it and shows its reason; and a
  chain must stop at one. Classed additive: no stored value changes meaning, and `SCHEMA_VER`
  stays `"v2.0"`, so every existing configuration keeps working unchanged.

## [0.10.0] - 2026-07-28

### Added

- **Seven new actions, and together they finish a job rather than sampling features.**
  Three bring data in — combine every sheet in this workbook, combine every workbook in a
  folder, import a delimited file — and four clean it up: remove blank rows and columns,
  unmerge and fill down, standardize number formats, convert text to real numbers. With the
  three cleaning actions that already shipped, they chain into one raw-files-to-clean-dataset
  run you can wire to a single button the day you import this.

  **The two combines line columns up by their header names, not by position.** A source
  whose `Amount` sits in column D still lands under `Amount`, and a header no earlier
  source declared becomes a new column on the right. A position-based combine quietly
  scrambles a real month-end folder, which is the failure this exists to prevent.

  **Neither combine will overwrite a sheet.** If the destination name is taken, the action
  stops before touching anything and tells you which sheet is in the way. The standard
  published version of this macro opens by deleting the existing sheet; that line does not
  exist here.

  **Removing blank rows shows you the count first and waits for a yes.** It also refuses to
  delete a row that is blank inside your selection but holds data outside it, and says how
  many it held back. The restore-point action that would make this safer is four waves
  away, so the preview is the safety story.

- **The build now refuses to ship a wrong action id.** Ids are permanent — there is no
  alias table, by design, because an alias means two names for one action forever and both
  would have to be ported to Apps Script. So a wrong id is unfixable, and the only useful
  time to catch one is before it exists.

  The gate reads its vocabulary of nine domains out of `spec/contract.md` rather than
  keeping a second copy, and fails the build if that table is missing, duplicated or
  malformed — never falling back to a built-in list, because a gate that quietly disables
  itself when its input goes missing is worse than no gate. It rejects a duplicate id, a
  camelCase id, an id in a reserved namespace, a new id in the closed legacy domain, and an
  entry point that is not actually dispatchable.

### Changed

- **`spec/contract.md` is v2.3.0**: §7 gains the nine-domain vocabulary table it was always
  assumed to have, §8 gains the `import` package and the new `clean` field and actions.
  Additive throughout — `SCHEMA_VER` stays `"v2.0"` and every existing configuration keeps
  its meaning.

- **§7 no longer says an action id is `pkg.verb`.** It never really was: the first segment
  is a user-facing vocabulary word, deliberately decoupled from the module that implements
  it, so that splitting a package later is not a contract break. `import.*` actions live in
  the new `PkgImport` while `importexport.import` stays where it is, and both appear on the
  same card.

### Fixed

- **`build.ps1` matched VBA keywords case-sensitively.** VBA is not case-sensitive, so a
  module written `public sub` slipped past the duplicate-name, block-balance and
  declaration-order guards — it balanced against nothing and was invisible to all three.
  Every pattern that matches VBA syntax now ignores case. Action ids are still matched
  case-sensitively, because there the case rule is the point.

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
