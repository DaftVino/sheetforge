# Project Handoff — Plug & Play Spreadsheet Macros

Use this to continue the project in VS Code (Claude Code). Read **Current status** (next) for where things stand, then paste the **Next-session prompt** at the bottom into a new session. The *Project context*, *Repo layout*, and VBA-gotchas sections below are the durable onboarding reference; *Current status*, *Build order*, and the prompt are refreshed as phases complete.

---

## Current status (2026-07-24)

**The engine + portable contract are done and published.** `main` is on a
**private** remote `github.com/DaftVino/sheetforge`.

- **Shipped (v0.1.0–v0.2.1):** `src/core/Core.bas` (config engine, Setup renderer,
  button builder, `Core_Run` dispatcher, `_Packages` registry, shared helpers) +
  `PkgImportExport` + `PkgClean`; `build/build.ps1` (Core + `Pkg*.bas` →
  single-module `dist/MacrosBundle.bas`; fails on duplicate proc names / block
  imbalance). Two-tier Setup UX (Simple/Advanced, help notes, suggest-import-types).
  Verify = `pwsh build/build.ps1` (green: Sub 32/32, Function 55/55, 87 procs) +
  a manual Excel import (owner-confirmed each release).
- **Contract frozen:** `spec/contract.md` **v1.0.0** — language-neutral SoT
  (manifest object, field model, kind vocabulary, tiers, config storage/read,
  action contract, frozen registry for both packages, portability appendix,
  SemVer change control). `/review`-audited against the code.
- **Sanitized + first publish:** all client jargon + the real store number
  scrubbed; pushed **tree-only** (pre-`d29b0d7` *history* still holds jargon —
  private repo; revisit before any public release).

**Two things intended but NOT done (not lost — deferred):**

- **Grouping/layout integration = step 7 below.** `MacroLayout.bas` is a working
  standalone prototype, still **un-built** (excluded from the bundle) and
  **un-wired** (`ML_RunById` placeholder, never connected to `Core_Run`).
- **Catalog build-out.** Only 2 of ~60 `macro-catalog-plug-and-play.md` macros
  are packages.

**Next phase = the "Compose your buttons" spec** — grouping integration + a new
chaining capability (compose actions into one sequential button), as one
Setup-page feature; additive **contract v1.1.0**. See `TODOS.md` "Next — active"
and the Next-session prompt at the bottom.

---

## Suggested repo layout

```
spreadsheet-macros/
├─ CLAUDE.md                 # paste the "Project context" section below here
├─ README.md
├─ spec/
│   └─ contract.md           # language-neutral config schema + action ids — v1.0.0, FROZEN
├─ src/
│   ├─ core/                 # Core module(s) — extracted from SetupImportExportMacros
│   ├─ packages/             # one .bas per package (PkgImportExport, PkgClean, ...)
│   │   ├─ PkgImportExport.bas
│   │   └─ MacroLayout.bas   # the grouping/layout feature (already built)
│   └─ legacy/
│       └─ SetupImportExportMacros.bas   # the working dashboard file (source of Core)
├─ dist/
│   └─ MacrosBundle.bas      # (build output) Core + enabled packages concatenated
├─ build/
│   └─ build.(ps1|py|js)     # concatenates src/core + chosen packages into dist/
├─ apps-script/              # (later) Google Sheets port mirroring spec/contract.md
└─ docs/
    ├─ engine_architecture_draft.md
    ├─ macro_catalog_plug_and_play.md
    └─ excel_macro_examples_2026.md
```

---

## Project context (paste into CLAUDE.md)

**Who:** Matt (James M. Baker), VP of Finance & Technology for a multi-entity Papa John's franchise group (70+ stores). High technical fluency — assume it. Wants dynamic, formula-driven, edge-case-aware solutions and direct pushback on shaky logic.

**Goal:** A reusable, config-driven macro toolkit so non-technical users and power users alike can apply common spreadsheet automations (clean, import/combine, reconcile, format, summarize, export) to almost any workbook without rewriting code per sheet. A user-input **Setup page** is the backbone: users fill it in and the system configures itself.

**Architecture (decided):**
- **Layered model.** One required **Core** module (config engine, Setup renderer, button/menu builder, `Core_Run` dispatcher, shared helpers, package registry) + many thin **package** modules (Clean, Export, ImportExport, Reconcile, ...). Packages depend on Core; Core depends on nothing.
- **Package contract (manifest).** Each package exposes exactly one sub `Pkg<Name>_Manifest(reg)` that declares: identity (id, name, version), config fields (key, label, type, default), and actions (actionId, caption, entry-point sub, help). Core calls manifests to render the Setup sheet and build buttons. Actions share a fixed signature `Pkg<Name>_<Verb>(ctx)` and call `Core_*` for all plumbing.
- **Dispatcher.** Buttons never point at package subs directly. `OnAction = 'Core_Run "clean.dupes"'`; Core looks up the entry point and `Application.Run`s it with a context object. Adding a package requires zero button code.
- **Discovery without VBIDE trust.** VBA can't enumerate its own modules without "Trust access to the VBA project object model" — which we deliberately avoid. Instead Core keeps a maintained `_Packages` list and probes each `Pkg<Name>_Manifest` via `Application.Run` (missing = catchable error). No trust setting required, ever.
- **Naming (mandatory — the namespace is flat).** Public entry points `Pkg<Name>_<Verb>`; manifests `Pkg<Name>_Manifest`; Core's API `Core_*`; everything else `Private`. Shared helpers (speed wrapper, last-row finder, error/log, config reader, type coercion, file finder) live ONLY in Core; packages call them.

**Distribution (decided):** Modular source during development; ship **ONE concatenated `.bas`** per deployment (Core + the packages that deployment needs). Config **toggles** enable/disable packages, so one built file can expose different action sets per site. One import for the 70-store, non-technical users. A tiny build step concatenates `src/core` + chosen `src/packages` into `dist/MacrosBundle.bas`.

**Platform (decided):** **Excel/VBA first.** Freeze the **config schema + action contract as a language-neutral spec** (`spec/contract.md`) so the later **Google Sheets (Apps Script)** port is mechanical. Sheets is secondary; others (LibreOffice Basic, Python) tertiary. Code is NOT shared across platforms (different languages/runtimes) — only the design/spec + the Setup-sheet layout are portable. Keep platform assumptions out of the contract: e.g. "files in the same folder as the workbook" (true in Excel, meaningless in Sheets/Drive) must be abstracted to a `source_location` config field the engine interprets per platform.

**Setup page** is the one surface where design matters as much as code: two tiers on one sheet (a Simple yellow-cell band for regular users + an Advanced band for power users), package-driven rendering (sections appear only for enabled packages), live validation before writing the hidden `_Config`, reference lists to cut typos, and graceful fallback to per-package defaults when unconfigured.

**VBA gotchas to respect:**
- `Option Explicit` everywhere.
- 25 line-continuation (`_`) limit per statement — build long seed lists with one `Collection.Add` per line, not a giant `Array(...)` with continuations.
- Wrap every action in the speed pattern (ScreenUpdating/Calculation/EnableEvents off → on) with guaranteed restore in the error handler.
- VBA only runs inside Excel; VS Code is for authoring/versioning/building, not executing.

---

## Current files (move these into the repo)

- **`SetupImportExportMacros.bas`** (module `SetupImportExportMacros`) → `src/legacy/`. The working dashboard automation for the Sample Dashboard workbook. Four buttons: Clear Input Tabs, Import Source Data (folder match + type coercion), Update Ref_Master (column-mapped refresh from a master file), Export Dashboard (values-only snapshot: flattens spills/legacy arrays/formulas, keeps formatting + conditional formatting, breaks external links, dated filename), plus a Setup/Finalize/BuildMacroTab config driver writing a hidden `_Config`. **~80% of Core already lives here** — this file is the extraction source, and its 4 buttons become the `PkgImportExport` package.
- **`MacroLayout.bas`** (module `MacroLayout`) → `src/packages/`. The grouping/layout feature. A Master sheet holds an Enabled flag per macro (stand-in for the future package-toggle layer); "Build Grouping Page" lists only enabled macros with a yellow **Code** cell; codes like `1A` mean tab 1 / column A (number = tab, letters = column, multi-digit tabs and `AA`+ columns supported); "Apply Layout" builds `Macros 1`, `Macros 2`... tabs placing buttons by code, stacked in list order. Buttons call a **placeholder `ML_RunById`** — the single seam to swap for `Core_Run`; **the dispatcher now exists, so wiring this is the next phase (step 7).** Re-running Build Grouping Page preserves already-entered codes.
- **`macro_catalog_plug_and_play.md`** → `docs/`. ~60 candidate macros grouped by user-facing job, each rated (⭐ project importance + A–F real-world popularity), with the ones already implemented in the dashboard marked done.
- **`engine_architecture_draft.md`** → `docs/`. The full architecture rationale, including the stress-test of the multi-file concept and the cross-platform table.
- **`excel_macro_examples_2026.md`** → `docs/`. Reference VBA snippets for common patterns.

---

## Build order (status)

Engine + contract (1–6) are **done**; 7–9 are the open scope.

1. ✓ `git init`, repo layout, `CLAUDE.md`.
2. ✓ **Extract Core** from `SetupImportExportMacros` (config engine, Setup renderer, button builder, helpers, `Core_Run` dispatcher, manifest API, `_Packages` registry).
3. ✓ **Refactor the 4 dashboard buttons** into `PkgImportExport` — no behavior change.
4. ✓ **Second package `PkgClean`** — proved the contract generalizes with zero new wiring.
5. ✓ **Freeze `spec/contract.md`** (v1.0.0) — `/review`-audited.
6. ✓ **`build/build.ps1`** — concatenate Core + `Pkg*.bas` → `dist/MacrosBundle.bas`.
7. **→ NEXT — "Compose your buttons":** one spec + build for **(a) grouping
   integration** (wire `MacroLayout`'s `ML_RunById` → `Core_Run` and surface
   grouping on the Setup page; adopt or rewrite the prototype) and **(b)
   chaining** (compose actions into one sequential button; generalizes the
   hardcoded clear→import). Additive **contract v1.1.0**. `brainstorming` →
   `writing-plans` first.
8. **Catalog build-out** (separate track): more `Pkg<Name>.bas` toward the ~60 in
   `macro-catalog-plug-and-play.md`.
9. **Apps Script port:** implement the frozen contract on Google Sheets, once the
   Excel feature set settles.

---

## Next-session prompt (paste into a new Claude Code session)

> Continuing **SheetForge** (config-driven Excel/VBA macro toolkit). Run `/orient` first.
>
> **Read manifest:**
> - `CLAUDE.md` (~5KB) — workflow gates
> - `docs/handoff.md` (this file, ~14KB) — Current status + Project context
> - `TODOS.md` (~8KB) — the **"Next — active"** item IS this task
> - `spec/contract.md` (~20KB) — frozen **contract v1.0.0**; you'll extend it to v1.1.0
> - `src/core/Core.bas` (44KB, 1127 lines — read in slices via Grep / `docs/code-map.md`): dispatcher `Core_Run` (~L601), button builder `Core_BuildControlPanel` (~L979), Setup renderer `Core_BuildSetup` (~L695)
> - `src/packages/MacroLayout.bas` (28KB) — standalone grouping prototype; `ML_RunById` (~L350) is the seam to `Core_Run`
> - `src/packages/PkgImportExport.bas` (28KB) — `PkgImportExport_Import` (~L106) is the hardcoded clear→import that chaining generalizes
> - `docs/engine-architecture-draft.md` §9 (Setup page) — design principles
>
> **Task — spec then build the "Compose your buttons" feature.** ONE feature, two linked Setup-page sections:
> 1. **Grouping/layout** — assign actions to pages/sections. Finish the pending **step 7**: wire `MacroLayout`'s `ML_RunById` onto `Core_Run` and surface grouping on the Setup page (adopt or rewrite the prototype — today it's standalone, un-built, un-integrated).
> 2. **Chaining** — compose several actions into ONE button that runs them sequentially (generalizes the hardcoded clear→import). User-defined chains are deployment-specific **config**, not manifest data — a new config entity.
>
> This touches Core (`Core_Run`, button builder, Setup renderer) + the contract → additive **contract v1.1.0**. So: **`brainstorming` first** to settle the open design questions — where chains live (config vs manifest), mid-chain error handling (stop vs continue on failure), whether chained steps suppress their individual confirmation dialogs, adopt-vs-rewrite MacroLayout, and how grouping + chaining share the Setup surface — then `writing-plans`, build in `src/`, `pwsh build/build.ps1`, `/review`, land to `main`. **Catalog build-out is a separate track — not part of this.**
