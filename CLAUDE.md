# CLAUDE.md

## Project

A reusable, config-driven Excel/VBA macro toolkit: one required Core module (config engine, Setup-sheet renderer, button builder, Core_Run dispatcher, package registry) plus thin package modules (ImportExport, Clean, Reconcile). Non-technical and power users configure automations via a Setup page; each deployment ships one concatenated .bas. Runs on Excel/Windows for a multi-entity franchise finance team, with a later Apps Script port planned against a language-neutral spec.

## Key docs

- Architecture: `docs/architecture.md`; design/engine rationale: `docs/engine-architecture-draft.md`
- Conventions: `engineering-standards/repo-standards.md` in the `daftplate` repo — canonical, never copied here. Local checkout path is in `~/.claude/CLAUDE.md`.
- Workflow quick reference: `docs/quick-ref-workflow.md`
- Decisions: `docs/adr/` — read before proposing architectural changes; start with `0001-source-first-vba-toolkit.md`
- Active plans / design-on-a-sheet: `docs/designs/`

## Commands

```
verify:  pwsh build/build.ps1  → concatenates src/core/Core.bas + src/packages/Pkg*.bas into dist/MacrosBundle.bas; FAILS on a duplicate proc name across modules (single-module bundle) or a Sub/Function block imbalance. `-Packages Pkg…` builds a subset.
test:    n/a — VBA has no headless runner; verify by importing dist/MacrosBundle.bas and running the workbook, recorded in the PR
deploy:  n/a — delivered as one imported .bas, not deployed
```

Run verify (the build) before every push; it is the closest thing to CI here. Never deploy without being asked.

## Workflow rules

The non-negotiable gates live in `~/.claude/CLAUDE.md` and govern this repo. Repo-specific additions only below. Deviating from a standard requires an ADR here **and** a note queued to `~/.daftplate/outbox/`. Silent deviation is the one unforgivable move. Tasks live on this repo's GitHub Project (repo-standards §6.5), never as draft cards.

This repo is **source-first**, not export-first (`docs/adr/0001-source-first-vba-toolkit.md`): the `.bas`/`.cls` under `src/` are the source of truth, the workbook is a build/import target. The scaffolded `scripts/export-vba.ps1` and `docs/vba-runbook.md` describe the inverse flow and are **non-authoritative here**. Full rationale: `docs/handoff.md`.

1. **`src/` is the source of truth.** Modular VBA: one required `src/core/` (config engine, Setup renderer, button builder, `Core_Run` dispatcher, `_Packages` registry, helpers) + thin `src/packages/Pkg<Name>.bas` that depend on Core. Review `.bas` diffs with `/review` like any code.
2. **Build, then import — one direction.** The build concatenates `src/core` + enabled packages into `dist/MacrosBundle.bas`, imported into Excel by hand. A `.bas` edit does nothing in Excel until re-imported; there is no round-trip back to `src/`.
3. **No VBIDE trust dependency, ever.** Core discovers packages by probing `Pkg<Name>_Manifest` via `Application.Run` against a maintained `_Packages` list — never by enumerating modules (which needs "Trust access to the VBA project object model").
4. **Naming is mandatory — flat namespace.** Entry points `Pkg<Name>_<Verb>`; manifests `Pkg<Name>_Manifest`; Core's API `Core_*`; everything else `Private`. Shared helpers live ONLY in Core.
5. **VBA hygiene.** `Option Explicit` everywhere; the 25 line-continuation limit per statement (one `Collection.Add` per line, not a giant `Array(...)`); wrap every action in the speed pattern (ScreenUpdating/Calculation/EnableEvents off → on) with guaranteed restore in the error handler; quote every path (folders have spaces and capitals).
6. **Keep the contract portable.** Freeze the config schema + action ids in `spec/contract.md` as the language-neutral source of truth so the later Apps Script port is mechanical.

## Pipeline

`brainstorming` → `writing-plans` for anything touching Core or more than one package → author in `src/` → build (`dist/MacrosBundle.bas`) → `/review` over the `src/` diff → `/ship`. Verify by importing the bundle and running the workbook; record that manual run in the PR.

## Off

- `test-driven-development` — no VBA runner; Excel cannot be driven headlessly on CI. Verification is the build plus a manual run of the imported workbook.
- `/qa`, `/browse`, the design family — no web surface (the `docs/designs/` Setup-page work is design-on-a-sheet, not a web UI).
- `/land-and-deploy`, `/canary`, `/gas-deploy`, `/benchmark` — delivered as one imported `.bas`, not deployed, and nothing to benchmark.

## Context budget

- **Never read an `.xlsm`** (binary archive — wastes the window). Read the `src/` module you are changing, not the whole toolkit; **grep `docs/code-map.md` first, then read the slice** — it anchors every proc, section and module-level `Const` in `src/`, and `Core.bas` is 77KB. Refresh it with `node scripts/code-map.mjs` (a repo-local generator: the shared daftplate one indexes no `.bas` at all). Roadmap: `docs/macro-catalog-plug-and-play.md`; VBA patterns: `docs/excel-macro-examples-2026.md`; decided already: `docs/handoff.md`.
- Session start: `/orient`. Phase or session end: `/handoff`.

## Subagent defaults

- "Which package/sub does X?" → an Explore subagent over `src/`; take its answer, not its file dumps. For repetitive edits across modules, a cheaper-model subagent, one task per module — remember edits don't take effect in Excel until the bundle is rebuilt and re-imported.
