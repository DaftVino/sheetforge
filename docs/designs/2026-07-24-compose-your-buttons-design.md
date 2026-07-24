# Design — "Compose Your Buttons" (contract v1.1.0)

Status: approved in brainstorming 2026-07-24; next step = `writing-plans`.
Scope: handoff **step 7** (grouping integration) + a new **chaining** capability,
as ONE Setup-page feature. Catalog build-out is a separate track, out of scope.

Touches Core (`Core_Run` dispatcher, button builder, Setup renderer, Finalize) +
`spec/contract.md` (additive **v1.1.0**). Deletes `src/packages/MacroLayout.bas`
(its logic is absorbed into Core).

---

## 1. The unifying idea

**A button is either an action or a chain** — both are just an id that
`Core_Run "<id>"` executes. Grouping arranges *buttons* onto pages; chaining
*creates a new kind of button*. The two features are one system:

- Chaining adds buttons (composed of existing actions).
- Grouping places buttons (actions and chains alike) onto pages.
- Both flow through the single dispatcher the contract already blesses.

Every placed button keeps `OnAction = 'Core_Run "<id>"'`. `Core_Run` learns to
recognize a **chain id** (namespaced `chain.<slug>`) and route it to the chain
runner; the action path is otherwise unchanged.

---

## 2. Chaining

### 2.1 Data model

A chain is `{ id: "chain.<slug>", label, steps: [actionId, ...] }`.

- **Deployment config**, not manifest data — authored on Setup, stored in the
  hidden `_Config`, read by Core at run time.
- **No nesting** — a chain's steps are plain actions only, never other chains.
- Steps run **in list order**, left → right.
- `id` is `chain.` + a slug derived from the Button name (lowercased, spaces →
  `_`, non-alphanumerics dropped). Collisions between two chain names that slug
  identically are reported at Finalize.

### 2.2 Authoring on Setup — the **Chains** table

Columns: `Button name | Steps | Code`.

- One chain per row. **Steps** = an ordered list of action *captions* separated
  by `,` or `→`. A **reference list of available action captions** is rendered
  alongside (same pattern as the existing "tabs in this workbook / files in this
  folder" reference columns).
- **Code** = this chain's grouping code (see §3); blank = default page.

At **Finalize**:

1. Split each Steps cell; resolve each caption to its action id via the registry
   (caption → id lookup built from `BuildRegistry()`).
2. Validate every step resolves to a real action id; report unknown captions
   (with the offending chain + text) **before** writing the config store —
   exactly like today's tab/file validation.
3. Store resolved chains as a table in `_Config` (label + comma-joined step-id
   list per row).

### 2.3 Runtime — `Core_RunChain`, reached via `Core_Run "chain.<slug>"`

The current single dispatch body is refactored into a shared inner routine:

```
RunAction(id, ctx) As Boolean      ' returns success; sets ctx error info on fail
```

used by both:

- **`Core_Run "<actionId>"`** — normal single-button path; dialogs ON (behavior
  unchanged from today).
- **`Core_RunChain "chain.<slug>"`** — the chain path:
  - Sets `ctx("quiet") = True` for each step, so steps run **silently**.
  - Runs steps in order via `RunAction`.
  - **On a step failure:** pop a mid-chain prompt —
    *"Step N ‹label› failed: ‹reason›. Continue / Stop?"* — and honor the choice
    (Continue → run remaining steps; Stop → halt).
  - **At the end:** ONE combined summary
    (`Refresh All — done. Clear ✓  Import ✓ (312 rows)  Update ✓  Export ✓`),
    or a "stopped at step N of M" summary if halted.

`Core_Run` first checks whether `<id>` names a defined chain (or carries the
`chain.` prefix) → routes to `Core_RunChain`; otherwise the action path. One
public button entry point, unchanged for callers.

### 2.4 The handler refactor (approved scope cost)

Each action's final `MsgBox` is routed through a new helper:

```
Core_Done ctx, msg      ' if ctx("quiet"): append msg to a collector on ctx and
                        ' suppress the dialog; else MsgBox exactly as today.
```

- Touches the **tail** of each handler in `PkgImportExport` and `PkgClean` —
  mechanical, **no behavior change** for standalone (non-chained) buttons.
- `importexport.import`'s existing internal silent `clear→import` is left as-is;
  it is already this pattern, now generalized to any composed sequence.
- The chain runner reads the collector to build its end-of-chain summary.

---

## 3. Grouping / layout

### 3.1 What it does

Assigns each button a **code** like `1A`: number = page, letter(s) = column.
Core builds `Macros 1`, `Macros 2`… pages, placing buttons by code and stacking
same-code buttons in list order. This is MacroLayout's model, rebuilt in Core.

### 3.2 Salvaged vs dropped (MacroLayout → Core)

**Salvaged** (logic moves into Core; the file is deleted):

- `ParseCode` — `12AB` → tab 12 / column "AB" (multi-digit tabs, `AA`+ columns).
- The yellow fill-in grid + **preserve-already-entered-codes** on rebuild.
- Column/stack placement geometry (`Macros N` tabs, columns, stacked buttons).

**Dropped:**

- Its separate hand-maintained **"Master" catalog** — Core's action registry is
  the single source of truth.
- The **`ML_RunById`** placeholder — `Core_Run` is the real dispatcher.

### 3.3 Authoring on Setup — the chicken-and-egg fix

Grouping must list chains, but chains are authored on the same sheet. To keep a
**single Setup pass** (no seed/refresh dance):

- **Actions** are known from the registry at build time → Core seeds a
  **Placement** table `Action | Code` from the registry, preserving prior codes.
- **Chains** aren't known at build time → each chain carries its **own Code**
  column in the Chains table (§2.2).

Both feed one code→placement map at Finalize.

### 3.4 Placement rules

- **No codes anywhere → today's single "Macros" tab, byte-for-byte unchanged.**
  Grouping is purely opt-in.
- Any codes present → build grouped pages. A **blank-coded button defaults to
  page 1** (nothing silently disappears — this differs from the prototype, which
  dropped blank codes).
- Core's own **Setup / Finalize** buttons are never user-groupable; they always
  land on page 1.

---

## 4. The shared Setup surface + engine-owned config

Both tables live under one **"Compose Your Buttons"** section on Setup:
the **Chains** table (§2.2), then the **Placement** table (§3.3), then the
available-actions reference list.

This introduces a concept the engine has not had: **config that no package
manifest declares.** Today every field/section on Setup comes from a package
manifest; chains and placement are authored freely by the deployment.

To reuse existing machinery instead of a parallel renderer/reader, Core registers
**a built-in pseudo-package `core`** in the registry whose manifest declares
these two tables. Then:

- the existing Setup renderer draws them,
- Finalize validates + stores them,
- `Core_CfgTable` reads them,

with only the special-case logic added on top: registry-seeding the Placement
rows, and caption→id resolution + validation for both tables.

---

## 5. Contract v1.1.0 — additive changes (minor bump)

1. **New config entity `chain`** — id, label, ordered `steps[]` of action ids;
   deployment config (not manifest data). Semantics: sequential run; quiet steps;
   mid-chain failure prompts continue/stop; one end-of-run summary.
2. **Engine-owned config** — §5 gains that config MAY originate from the engine
   itself (the `core` pseudo-package), not only package manifests.
3. **Grouping/layout as portable config** — a per-button placement code
   (`<page><column>`) + page-building semantics; blank/absent = single default
   page.
4. **Dispatch (§6) extension** — `Core_Run` resolves *either* an action id or a
   `chain.<slug>` id; adds the quiet-context convention (`ctx.quiet` + a
   done-sink) so composed steps suppress their own dialogs. Handler binding stays
   platform-specific.

All additive: no existing key, action id, kind, or semantics changes → v1.0.0
deployments remain valid (§10 SemVer rule).

---

## 6. Build / verify / cleanup

- **Delete** `src/packages/MacroLayout.bas` (never in the build; logic absorbed).
- New Core procedures (names TBD in planning; must not collide — the build fails
  on duplicate proc names): `Core_RunChain`, `RunAction` (shared inner dispatch),
  `Core_Done`, the code parser + placement builder, chain/placement config
  readers + Finalize validators, and pseudo-package registration.
- `pwsh build/build.ps1` must stay green (Sub/Function balance + no duplicate
  names). Salvaged helpers take `Core`/private names — no `ML_*` carried over.
- No headless VBA test: manual Excel import + run of (a) a defined chain and
  (b) a grouped layout, recorded in the PR.

---

## 7. Boundaries (YAGNI)

Out of scope for this feature:

- Nested chains (a chain calling a chain).
- Conditional / branching steps or per-step arguments.
- Drag-and-drop layout — codes only.
- Reordering UI beyond authoring list order.
- Catalog build-out (separate track).

Grouping and chaining ship **together** as one contract bump and one build.

---

## 8. Decisions settled in brainstorming (2026-07-24)

| Question | Decision |
|---|---|
| Grouping vs the single control panel; MacroLayout adopt/rewrite | **Rewrite** grouping into Core as an opt-in config layer sourced from the registry; salvage the prototype's code-parser / fill-in grid / code-preservation; **delete** `MacroLayout.bas`. |
| Where chains live | **Config** (deployment-specific), surfaced via a Core-owned pseudo-package. |
| Mid-chain error handling | **Ask the user** — Continue / Stop prompt on the failing step. |
| Chained-step dialogs | **Suppressed**; one combined summary at the end (via `ctx.quiet` + `Core_Done`). |
| Chain authoring format | **One row per chain**: `Button name | Steps | Code`, steps chosen against an actions reference list, validated at Finalize. |
