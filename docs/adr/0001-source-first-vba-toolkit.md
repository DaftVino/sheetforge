# ADR 0001: Author VBA as source in-repo; the workbook is a build output, not the artifact

<!-- File: docs/adr/0001-source-first-vba-toolkit.md — one page max. Never edited after acceptance; reversals are a new ADR. -->

**Status:** Accepted
**Date:** 2026-07-23

## Context

SheetForge is a reusable, config-driven Excel/VBA macro toolkit: one required Core module (config engine, Setup-sheet renderer, button builder, `Core_Run` dispatcher, package registry) plus many thin package modules that depend on Core. It is developed in VS Code and versioned in git, then delivered to non-technical users across a 70+ store franchise group as **one concatenated `.bas`** per deployment. VBA only executes inside Excel, so authoring, review, and building all happen outside it. The repo was scaffolded from daftplate's `office-automation` profile — the closest of the eight — but that profile assumes the inverse workflow: a hand-edited `.xlsm` is the source of truth and `vba/` is a one-directional *export* generated from it (`scripts/export-vba.ps1`), which is explicitly never edited.

## Decision

We will treat the **`.bas`/`.cls` files under `src/` as the source of truth** for this repo, authored directly in VS Code and reviewed as normal code. A workbook is a **build/import target, not a tracked source artifact**: a build step concatenates `src/core` + the enabled `src/packages` into `dist/MacrosBundle.bas`, which is imported into Excel by hand. Config toggles enable/disable packages so one built file exposes different action sets per site. The **config schema + action contract is frozen as a language-neutral spec** (`spec/contract.md`) so a later Google Sheets (Apps Script) port is mechanical; code is not shared across platforms, only the spec and the Setup-sheet layout. This decision governs **this repo only**. It supersedes the `office-automation` profile's export-first assumptions for SheetForge: the profile's `scripts/export-vba.ps1`, `docs/vba-runbook.md`, and the "never edit the `.bas` / never read the `.xlsm`" constraints in `CLAUDE.md` do not apply here and have been replaced by the source-first equivalents.

## Alternatives considered

- **Follow the profile as written (export-first).** Keep a canonical `.xlsm` and treat `vba/` as a generated export. Rejected: it inverts the actual workflow — the toolkit is authored as modular source and shipped as a bundle, never round-tripped through one master workbook — and it would force "Trust access to the VBA project object model" on, which the architecture deliberately avoids (Core probes package manifests via `Application.Run`, not the VBIDE object model).
- **A different profile (`local-tool` / `app-monolith`).** Rejected: the deliverable is Excel-hosted VBA, not a Node CLI or a TS service+DB; `office-automation` is the only profile whose stack assumptions (Excel/Windows, VBA-as-software, no web/deploy) fit.
- **Build the Google Sheets port now, in parallel.** Rejected: prove the manifest → dispatcher → registry loop on two Excel packages first; port mechanically against the frozen spec afterward.

## Consequences

- **Easier:** `.bas` diffs are first-class and reviewable with `/review`; no Excel round-trip or Trust Center dependency to read a change; one import for non-technical users; the spec keeps the eventual Sheets port cheap.
- **Harder / given up:** no headless test runner — verification is a build plus a manual run of the imported workbook, recorded in the PR. The profile's `export-vba.ps1` and `vba-runbook.md` are dead weight for SheetForge until/unless an export flow is ever wanted; they are left in place but marked non-authoritative by `CLAUDE.md`. A build script (`build/build.*`, HANDOFF step 6) does not exist yet, so `verify:` is manual until it lands.
- **Revisit if:** SheetForge ever adopts a canonical workbook that is edited in Excel and round-tripped (e.g. forms/`.frx` layout work), at which point the export-first model — and a new, superseding ADR — would be warranted.

## Amendment, 2026-07-29 — the export-first files are deleted, not just marked

The Consequences section above says `export-vba.ps1` and `vba-runbook.md` "are left in place but marked non-authoritative by `CLAUDE.md`". **That is no longer true, and the decision itself is unchanged** — this amendment records only that the two files were removed before the repository was made public.

Marking them was sufficient while the repo was private and every reader had already read `CLAUDE.md`. It stops being sufficient in public, for a reason that is not about tidiness: `docs/vba-runbook.md` step 1 instructed the reader to tick **Trust access to the VBA project object model**. That is a security-relevant Excel setting, this architecture is specifically built so that it is never needed (see the first alternative considered above), and `CLAUDE.md` rule 3 forbids depending on it outright. A stranger browsing `docs/` does not read `CLAUDE.md` first, so the marking they were relying on was invisible at exactly the moment it mattered — the reader would have loosened a real setting to follow a workflow this repo rejected in this very ADR.

The `.gitignore` now also ignores `*.xlsm` / `*.xlsx` outright, which is this decision applied one step further: if a workbook is never a source, a public repo should make it impossible to commit one by accident.

Nothing else changes. Should an export flow ever genuinely be wanted, it needs a new superseding ADR, and it cannot be built on VBIDE trust.
