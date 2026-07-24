# Project Handoff — Plug & Play Spreadsheet Macros

Use this to continue the project in VS Code (Claude Code). Move the files listed below into the repo, then paste the **Kickoff prompt** at the bottom into a new Claude Code session.

---

## Suggested repo layout

```
spreadsheet-macros/
├─ CLAUDE.md                 # paste the "Project context" section below here
├─ README.md
├─ spec/
│   └─ contract.md           # (to create) language-neutral config schema + action ids — source of truth
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
- **`MacroLayout.bas`** (module `MacroLayout`) → `src/packages/`. The grouping/layout feature. A Master sheet holds an Enabled flag per macro (stand-in for the future package-toggle layer); "Build Grouping Page" lists only enabled macros with a yellow **Code** cell; codes like `1A` mean tab 1 / column A (number = tab, letters = column, multi-digit tabs and `AA`+ columns supported); "Apply Layout" builds `Macros 1`, `Macros 2`... tabs placing buttons by code, stacked in list order. Buttons call a **placeholder `ML_RunById`** — the single seam to swap for `Core_Run` once the dispatcher exists. Re-running Build Grouping Page preserves already-entered codes.
- **`macro_catalog_plug_and_play.md`** → `docs/`. ~60 candidate macros grouped by user-facing job, each rated (⭐ project importance + A–F real-world popularity), with the ones already implemented in the dashboard marked done.
- **`engine_architecture_draft.md`** → `docs/`. The full architecture rationale, including the stress-test of the multi-file concept and the cross-platform table.
- **`excel_macro_examples_2026.md`** → `docs/`. Reference VBA snippets for common patterns.

---

## Next steps (build order)

1. `git init`; drop in the files per the layout; paste Project context into `CLAUDE.md`.
2. **Extract Core** from `SetupImportExportMacros`: config engine (`CfgSheet`/`SheetSingle`/`SheetList1`/`SheetTable3`), Setup renderer (`CreateSetupTab`/`FinalizeSetupTab`), button builder (`BuildMacroTab`), helpers (`FindSourceFile`/`CoerceVal`/`SheetExists`), plus new `Core_Run` dispatcher, `Core_DeclarePackage`/`Core_AddField`/`Core_AddAction`, and the `_Packages` registry.
3. **Refactor the 4 dashboard buttons** into `PkgImportExport` with a manifest — proving the contract on code that already works, no behavior change for dashboard users.
4. **Add a second tiny package** (`PkgClean`: RemoveDuplicates + TrimClean) to prove buttons appear with zero new wiring.
5. **Freeze `spec/contract.md`** — config field keys, action ids, action semantics — as the portable source of truth.
6. **Write `build/build.*`** to concatenate Core + selected packages into `dist/MacrosBundle.bas`.
7. Wire `MacroLayout`'s `ML_RunById` to `Core_Run`.
8. **Only then** start the Apps Script port against the spec.

Do 2–4 before touching Sheets: if the manifest → dispatcher → registry loop works for two packages in Excel, the platform is proven and porting becomes mechanical.

---

## Kickoff prompt (paste into Claude Code)

> I'm continuing a VBA project — a reusable, config-driven Excel macro toolkit (details in `CLAUDE.md` and `docs/engine_architecture_draft.md`). Please read `CLAUDE.md`, `docs/engine_architecture_draft.md`, `src/legacy/SetupImportExportMacros.bas`, and `src/packages/MacroLayout.bas` first.
>
> Then start **step 2 of the build order: extract Core** from `SetupImportExportMacros.bas` into `src/core/`. Create a `Core` module that exposes: the config engine (generalized from the `Cfg*`/`Sheet*` readers), shared helpers (speed wrapper, robust last-row/col finder, structured error+log handler, `SheetExists`, file finder, type coercion), a `_Packages` registry with `Application.Run`-based manifest probing (no "Trust access to VBA project" dependency), the manifest-building API (`Core_DeclarePackage`, `Core_AddField`, `Core_AddAction`), a Setup-sheet renderer that builds fields from installed packages' manifests, a button/menu builder, and the `Core_Run "<actionId>"` dispatcher.
>
> Keep the existing dashboard behavior working — the goal is that the four current buttons can be re-expressed as a `PkgImportExport` package (step 3) with no functional change. Follow the naming conventions (`Core_*`, `Pkg<Name>_*`, everything else `Private`), respect the VBA 25-continuation limit, and preserve the graceful fallback-to-defaults behavior. Propose the Core module's public API as a short list before writing the full implementation so I can approve it.
