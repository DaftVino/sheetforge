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
