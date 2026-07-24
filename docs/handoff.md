# Project Handoff — Plug & Play Spreadsheet Macros

Use this to continue the project in VS Code (Claude Code). Read **Current status** (next) for where things stand, then paste the **Next-session prompt** at the bottom into a new session. The *Project context*, *Repo layout*, and VBA-gotchas sections below are the durable onboarding reference; *Current status*, *Build order*, and the prompt are refreshed as phases complete.

---

## Current status (2026-07-24)

**The engine + portable contract are done and published; "Compose Your
Buttons" (grouping + chaining) is now built on `feat/compose-your-buttons`,
contract v1.1.0, pending merge to `main`.** `main` remains on a **private**
remote `github.com/DaftVino/sheetforge`.

- **Shipped (v0.1.0–v0.2.1) on `main`:** `src/core/Core.bas` (config engine,
  Setup renderer, button builder, `Core_Run` dispatcher, `_Packages` registry,
  shared helpers) + `PkgImportExport` + `PkgClean`; `build/build.ps1` (Core +
  `Pkg*.bas` → single-module `dist/MacrosBundle.bas`; fails on duplicate proc
  names / block imbalance). Two-tier Setup UX (Simple/Advanced, help notes,
  suggest-import-types). Verify = `pwsh build/build.ps1` (green) + a manual
  Excel import (owner-confirmed each release).
- **Shipped (branch `feat/compose-your-buttons`, contract v1.1.0):**
  **Compose Your Buttons** — a Core-owned Setup section with **Chains**
  (compose several actions into one button run in order, `Label [action.id]`
  step tokens, immutable `chain.<id>` config entity, Continue/Stop prompt on a
  step failure, per-step dialogs suppressed via `Core_Done` in favor of one
  summary) and **Grouping/layout** (`1A`-style page/column codes per button; no
  codes → the single "Macros" tab; any code → numbered "Macros N" pages, blank
  → page 1 "Unassigned", `hide` → omitted; Setup/Finalize admin controls always
  on page 1). Finalize validates chains + placement and blocks activation on
  structural errors. This finishes the long-pending **step 7** below.
  `src/packages/MacroLayout.bas` — the standalone grouping prototype — has been
  **deleted**; its logic (`ParseCode`, `IndexToCol`, `SortedKeys`,
  `DeletePanels`, `AddText`, geometry) was absorbed into Core. Build verified
  green (`pwsh build/build.ps1`, exit 0); **manual Excel import/run is
  pending the owner** (checklist in the PR).
- **Contract frozen:** `spec/contract.md` **v1.0.0**, extended additively to
  **v1.1.0** for the `chain` config entity, reserved `chain.*` namespace,
  engine-owned config, placement code, and the dispatch/quiet convention.
  `/review`-audited against the code.
- **Sanitized + first publish:** all client jargon + the real store number
  scrubbed; pushed **tree-only** (pre-`d29b0d7` *history* still holds jargon —
  private repo; revisit before any public release).

**Deferred (not lost):**

- **Built-page polish** (design §8) — per-page titles/column headings/legend,
  Previous/Next page navigation, a live placement preview, a visually distinct
  admin area. Out of scope for "Compose your buttons"; tracked in `TODOS.md`
  Roadmap.
- **Catalog build-out.** Only 2 of ~60 `macro-catalog-plug-and-play.md` macros
  are packages.

**Next phase** — merge `feat/compose-your-buttons` to `main` once the owner
completes the manual Excel checklist, then pick up Catalog build-out or the
Apps Script port. See `TODOS.md` Roadmap.

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

Engine + contract (1–7) are **done**; 8–9 are the open scope.

1. ✓ `git init`, repo layout, `CLAUDE.md`.
2. ✓ **Extract Core** from `SetupImportExportMacros` (config engine, Setup renderer, button builder, helpers, `Core_Run` dispatcher, manifest API, `_Packages` registry).
3. ✓ **Refactor the 4 dashboard buttons** into `PkgImportExport` — no behavior change.
4. ✓ **Second package `PkgClean`** — proved the contract generalizes with zero new wiring.
5. ✓ **Freeze `spec/contract.md`** (v1.0.0) — `/review`-audited.
6. ✓ **`build/build.ps1`** — concatenate Core + `Pkg*.bas` → `dist/MacrosBundle.bas`.
7. ✓ **"Compose your buttons"** — grouping integration (assign each button a
   page/column code on the Setup page; `MacroLayout.bas`'s prototype logic
   absorbed into Core, the standalone file deleted) **and** chaining (compose
   actions into one sequential button, `Label [action.id]` steps, immutable
   `chain.<id>` config, Continue/Stop on failure). Additive **contract v1.1.0**.
   Built on `feat/compose-your-buttons`; manual Excel verification pending the
   owner.
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
> - `docs/handoff.md` (this file) — Current status + Project context
> - `TODOS.md` — Roadmap section has the next candidates (Catalog build-out, Apps Script port)
> - `spec/contract.md` — frozen **contract v1.1.0** (chains + grouping added additively over v1.0.0)
> - `src/core/Core.bas` — read in slices via Grep / `docs/code-map.md`
>
> **Status:** "Compose Your Buttons" (chains + grouping/layout) is built on
> `feat/compose-your-buttons`, `pwsh build/build.ps1` is green, and
> `src/packages/MacroLayout.bas` (the standalone grouping prototype) has been
> deleted — its logic now lives in Core. **What's outstanding:** the owner's
> manual Excel import/run of `dist/MacrosBundle.bas` (checklist recorded in the
> Task 10 PR) — do this before merging the branch to `main`.
>
> **Task — pick the next roadmap item** (see `TODOS.md` Roadmap): Catalog
> build-out (more `Pkg<Name>.bas` toward the ~60 in
> `macro-catalog-plug-and-play.md`), the Apps Script port, or the deferred
> built-page polish (design §8: per-page titles/headings/legend, page nav, a
> live placement preview, a distinct admin area). Confirm with the user which
> to pick up before starting `brainstorming`/`writing-plans`.
