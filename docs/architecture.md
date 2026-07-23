# Architecture

A reusable, config-driven Excel/VBA macro toolkit: one required Core module (config engine, Setup-sheet renderer, button builder, Core_Run dispatcher, package registry) plus thin package modules (ImportExport, Clean, Reconcile). Non-technical and power users configure automations via a Setup page; each deployment ships one concatenated .bas. Runs on Excel/Windows for a multi-entity franchise finance team, with a later Apps Script port planned against a language-neutral spec.

## How it works

TODO — replace with the real shape of the system before the first release. This is a living document: update it in the same PR as any change it describes.

## Key decisions

See `docs/adr/`.
