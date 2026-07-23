# CLAUDE.md

## Project

A reusable, config-driven Excel/VBA macro toolkit: one required Core module (config engine, Setup-sheet renderer, button builder, Core_Run dispatcher, package registry) plus thin package modules (ImportExport, Clean, Reconcile). Non-technical and power users configure automations via a Setup page; each deployment ships one concatenated .bas. Runs on Excel/Windows for a multi-entity franchise finance team, with a later Apps Script port planned against a language-neutral spec.

## Key docs

- Architecture: `docs/architecture.md`
- Conventions: `engineering-standards/repo-standards.md` in the `daftplate` repo — canonical, never copied here (its ADR 0001). Your local checkout path is recorded in `~/.claude/CLAUDE.md`.
- Workflow quick reference: `docs/quick-ref-workflow.md`
- Decisions: `docs/adr/` — read before proposing architectural changes
- Active plans: `docs/designs/`

## Commands

```
verify:  pwsh -File scripts/export-vba.ps1 -Check
test:    n/a
deploy:  n/a
```

Run verify before every push. Never deploy without being asked.

## Workflow rules

The non-negotiable gates live in `~/.claude/CLAUDE.md` and govern this repo. Repo-specific additions only below.

Deviating from a standard requires an ADR here **and** a note queued to `~/.daftplate/outbox/` so the rule itself gets revisited upstream. Silent deviation is the one unforgivable move. Tasks live on this repo's GitHub Project (repo-standards §6.5) — never as draft cards, which agents cannot see.

## Repo-specific constraints

1. **The VBA in `vba/` is the reviewable artifact; the `.xlsm` is the binary that carries it.** Run `pwsh -File scripts/export-vba.ps1` before every commit that touched a macro. A workbook committed without its export is a change nobody can read.
2. **The export is one-directional.** `vba/` is generated *from* the workbook. Editing a `.bas` file changes nothing until it is imported back in Excel, by hand, deliberately. Never "fix" a bug by editing the export.
3. **Excel must trust the VBA project object model** or the export silently produces nothing. The Trust Center path, the import procedure, and the never-merge-a-workbook rule are all in `docs/vba-runbook.md` — read it before your first export.
4. **Quote every path.** These folders have spaces and capitals in their names, and an unquoted path fails in a way that looks like a missing file rather than a quoting bug.

## Pipeline

`brainstorming` → `writing-plans` for anything touching more than one workbook → export (`scripts/export-vba.ps1`) → `/review` over the `vba/` diff → `/ship`.

## Off

- `test-driven-development` — VBA has no runner here and Excel cannot be driven headlessly on CI. Verification is the export plus a manual run of the workbook, recorded in the PR.
- `/qa`, `/browse`, the design family — no web surface.
- `/land-and-deploy`, `/canary`, `/gas-deploy`, `/benchmark` — the workbook is delivered by being opened, not deployed, and there is no harness to benchmark.

## Context budget

- **Never read an `.xlsm`.** It is a binary archive; a `Read` on it wastes the window and returns nothing usable. The `vba/` export is the readable form.
- Read the module you are changing, not the whole export. `docs/code-map.md` (from `/code-map`) indexes `vba/` once it grows past a few files.
- `docs/vba-runbook.md` answers the Excel-side questions once; read it instead of re-deriving the Trust Center path.
- Session start: `/orient`. Phase or session end: `/handoff`.

## Subagent defaults

- "Which macro does X?" → an Explore subagent over `vba/`; take its answer, not its file dumps.
- Repetitive edits across many modules → a cheaper-model subagent, one task per workbook, and remember the edits do not take effect until re-imported.
