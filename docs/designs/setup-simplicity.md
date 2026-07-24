# Design: Setup-sheet simplicity for non-technical users

Status: **planned, not started** (decided 2026-07-23; deferred behind PkgClean).
Owner objective: the toolkit's north star — make configuration as simple and
self-explanatory as possible for non-technical franchise-finance users.

## Why

Real-workbook test of the step-2/3 bundle (Core + PkgImportExport) passed
functionally. UX findings:

- The per-column **type map** (`T,N,D,DT,B,G`) is the one genuinely expert-level
  field. But the **defaults already work** for the existing dashboard, so regular
  users never need to touch it — the pain only appears when a power user adapts
  the toolkit to a **new workbook or a changed data schema**.
- The G/H reference lists (tabs in book / files in folder) are helpful — extend
  that "give them something to pick, don't make them type" instinct.

So the highest-leverage move is not to make the type map easier to fill, but to
**hide it from the people who shouldn't see it**, and to help the power user who
does.

## Adjustment 1 — Two-tier Setup (Simple / Advanced)

Split the Setup sheet into a **Simple** band (must-fill fields, plain-language
labels) and an **Advanced** band clearly marked *"most users can leave these
as-is."* Regular users finish the top and stop.

Assign fields to a tier:

- **Simple:** export tabs, import tab names + source prefixes, master prefix,
  master destination, master columns.
- **Advanced:** the type maps, the export date offset, the settings/CF sheet
  name, the master source-sheet name.

Core changes:

- Each manifest field carries a `tier` (`"simple"` default, `"advanced"`).
  Either an optional trailing arg on `Core_AddField`/`Core_AddList`/`Core_AddTable`
  or paired `Core_AddAdvanced*` calls (keep the 25-continuation limit in mind —
  one `.Add`-style call per line).
- `Core_BuildSetup` renders a "SIMPLE — fill these in" region, then an
  "ADVANCED — most users can leave these as-is" region, per installed package.
- `Core_FinalizeSetup` reads both tiers identically (tier is presentation only).

## Adjustment 2 — Suggest-types helper (power-user path)

A helper that, for an import row (input tab + source prefix already given), opens
the matched source file, samples the data, and writes **suggested** type codes
into the type-map cell for the user to review.

**Hard constraint — suggest, never auto-apply.** A store number like `3176`
samples as a *number* but must stay *text* or the dashboard's MATCH lookups break
(the exact failure the type maps exist to prevent). Detection cannot see this, so
the helper must fill a suggestion the user confirms/edits, never silently commit.

Two shapes to pick from at build time:

- A "Suggest column types" button that fills the type-map cell from a source
  sample (numbers/dates/text), leaving ID columns for the user to correct.
- The simpler inversion: ask only **"which columns are text?"** (by number or
  header) and default the rest to `G` (auto). Easier to fill and explain; loses
  precise `N` vs `D` typing, which the coercion mostly handles anyway.

Deliver as a Core "Setup service" a package can request (aligns with the draft §9
"reference helpers as a Core service" idea), so future packages reuse it.

## Open build-design note (affects step 6)

The current delivery is **one concatenated module** (`dist/MacrosBundle.bas`).
A single `.bas` is a single VBA module, so **every `Private` helper name must be
globally unique across Core + all packages** — this contradicts the draft §5
claim that package-private helpers can share short names freely. Options for the
step-6 build script: (a) keep one module and enforce/lint unique private names,
or (b) build to multiple modules (loses the single-import simplicity). Decide
before packages proliferate.

## Handoff log

- 2026-07-23: Decided direction (two-tier + suggest-types) after the first
  real-workbook test. Sequenced **after** PkgClean (prove-the-engine-generalizes
  step). This doc captures the decision; a full implementation plan should
  precede the work. The button-overlap layout bug found in the same test was
  fixed immediately in `Core_BuildControlPanel` (not deferred).

### 2026-07-23 — Engine + Setup-UX arc shipped (v0.1.0 -> v0.2.1)

**Branch / merges (all local; repo has NO git remote yet).** `main` head
`af54b2f`. Squash points: `08442f2` merge v0.1.0 (Core + PkgImportExport +
PkgClean + build), `d0d005a` merge v0.2.0 (two-tier + suggest-types + clarity),
`1e9e3ff` merge v0.2.1 (band merge/wrap + banner-duplicate fix). Feature
branches `feat/core-import-export`, `feat/setup-two-tier`, `feat/setup-band-polish`
are merged but **not deleted** (see TODOS.md housekeeping).

**What shipped.** `src/core/Core.bas` (44KB, the engine), `src/packages/PkgImportExport.bas`
(28KB), `src/packages/PkgClean.bas` (7KB), `build/build.ps1` (concatenate ->
single-module `dist/MacrosBundle.bas`; fails on duplicate proc names or block
imbalance), `VERSION`, `CHANGELOG.md`, `TODOS.md`, this doc. `dist/` is
**gitignored** (build target). No test count — VBA has no headless runner;
verify = `pwsh build/build.ps1` (green: Core + 2 packages, Sub 32/32, Function
55/55, 87 unique procs) plus a manual Excel import, which the user ran and
confirmed at each release.

**What the plan got wrong / do-not-revert.**

- Two-tier was scoped DOWN from this doc: the type map stays inside the import
  table (only 3 scalars are `advanced`: settings sheet, date offset, master
  source-sheet). **Do not revert** to hiding the type map as a whole field - a
  table column can't be cleanly tiered; that friction is handled by the
  suggest-types helper instead.
- Import shows ONE dialog (the clear is a silent internal step), not the
  legacy two. **Do not revert.**
- Export forces `xlCalculationAutomatic` after `Core_SpeedOn` (which sets
  Manual). **Do not revert** - the spill/legacy-array flatten path was written
  and tested under Automatic; Manual risks the #REF!/dropped-spill bugs.
- `Core_LastRowUp` (bare End(xlUp)) is used for the Import SOURCE read; the
  widened `Core_LastRow` is for CLEARING. **Do not revert** the source read to
  the widened one (it pulled trailing blank rows).
- Single-module bundle => every Private name must be globally unique; the build
  enforces this. **Do not revert** to allowing dup private names.

**What was discovered (empirical, for the next session).**

- VBA `And` does NOT short-circuit - comparing an Error-subtype Variant (`#N/A`
  cell) to a string raises err 13. Branch on `VarType`/`IsError` first (see
  `InferColumn`).
- Excel `AutoFit` IGNORES merged cells - merged tinted band rows get an explicit
  height estimate in `Core.TintRow` (chars-per-line = 115, erring tall).
- The trailing list/table block needs the `END_KEY` sentinel row or its values
  are dropped (End(xlUp) stops at the block's own key row).
- A gitleaks pre-commit hook runs on every commit (all clean so far).
- `gh` is authed as `DaftVino`; the private repo `github.com/DaftVino/sheetforge`
  exists but **nothing has been pushed** - sanitize first.
- Build/scripting language for this repo is **PowerShell** (`build/build.ps1`).

**Still open (needs human decision).** Sanitization of real store/entity names
before any push to the remote is the blocker - see TODOS.md "Before the remote".
Everything through v0.2.1 is Excel-verified; no re-verify pending.

**Next phase — Sanitize, then Freeze the contract.** Read manifest:
`TODOS.md` (2.3KB, the task list), `CLAUDE.md` (5.3KB), `src/packages/PkgImportExport.bas`
(28KB, holds the real defaults to scrub), `src/legacy/SetupImportExportMacros.bas`
(45KB - the other place real names live), then `docs/engine-architecture-draft.md`
section 8/10 for the portable-spec plan. Start with `/orient`.
