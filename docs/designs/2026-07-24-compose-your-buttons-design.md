# Design — "Compose Your Buttons" (contract v1.1.0)

Status: approved in brainstorming 2026-07-24; hardened by a `/deliberate` panel
(same day) — see §9 for what changed and why. Next step = `writing-plans`.

Scope: handoff **step 7** (grouping integration) + a new **chaining** capability,
as ONE Setup-page feature. Catalog build-out is a separate track, out of scope.

Touches Core (`Core_Run` dispatcher, button builder, Setup renderer, Finalize) +
`spec/contract.md` (additive **v1.1.0**). Deletes `src/packages/MacroLayout.bas`
(its logic is absorbed into Core).

---

## 1. The unifying idea

**A button is either an action or a chain** — both are just a stable **id** that
`Core_Run "<id>"` executes. Grouping arranges *buttons* onto pages; chaining
*creates a new kind of button*. The two features are one system:

- Chaining adds buttons (composed of existing actions).
- Grouping places buttons (actions and chains alike) onto pages.
- Both flow through the single dispatcher the contract already blesses.

Every placed button keeps `OnAction = 'Core_Run "<id>"'`. **`Core_Run` remains the
sole public entry point.** It recognizes a reserved **chain id** (`chain.<id>`)
and routes it to a **`Private`** chain runner; the action path is unchanged.

**Identity rule (the spine of this whole design):** everything is keyed on
**stable ids**, never on captions or labels. Captions/labels are presentation —
they get reworded, localized, and edited, and a package version can change a
caption. Persisting a caption as a key would let a harmless rename silently
rewire a finance automation. So: chain steps store **action ids**; placement
rows are keyed by **button id**; a chain's identity is an **immutable generated
id**, independent of its editable label.

---

## 2. Chaining

### 2.1 Data model

A chain is `{ id: "chain.<token>", label, steps: [actionId, ...] }`.

- **Deployment config**, not manifest data — authored on Setup, stored in the
  hidden `_Config`, read by Core at run time.
- **`id` is immutable and generated**, not derived from the label. The token is
  an opaque short id minted at first Finalize (reference: `c1`, `c2`, … or a
  slug-plus-disambiguator — exact scheme chosen in planning; the only
  requirements are *stable*, *unique*, and *not the label*). The **label is
  editable freely** without changing identity.
- **`chain.` is a reserved namespace.** No package action id may begin `chain.`;
  the build/Finalize flags any that does. `Core_Run` treats a `chain.` id as a
  chain, everything else as an action.
- **No nesting** — a chain's steps are plain action ids only, never other chains.
- Steps run **in list order** (the authoring row order).

### 2.2 Authoring on Setup — the **Chains** table

Columns (left→right), each chain one row:

| (hidden id) | Button name | Steps | Code |
| ----------- | ----------- | ----- | ---- |

- **Hidden id** — the immutable `chain.<token>`, in the machine-key column (col A
  convention), preserved across Setup rebuilds exactly like every other field's
  machine key. A row with no id is a *new* chain; Finalize mints one.
- **Button name** — the editable label.
- **Steps** — an ordered list of **`Label [action.id]` tokens** separated by an
  unambiguous delimiter (reference: newline or `;` — never comma, since captions
  contain commas). A reference panel lists every available action as a
  **copy-ready token** (`Import Source Data [importexport.import]`), so the user
  copies rather than transcribes. Only the bracketed **id** is authoritative;
  the label half is a human hint and is re-derived on rebuild.
- **Code** — this chain's grouping code (§3); blank = page 1 (§3.4).

At **Finalize** (validate-before-write, per the project's existing philosophy):

1. Parse each Steps cell into tokens; extract the bracketed **action id** from
   each.
2. Validate every id resolves to a real registered action; report unknown /
   empty-step / malformed-token rows **by chain name** before writing config.
3. Preserve/assign the immutable chain id (match the hidden-id column; mint for
   new rows; report two rows sharing an id).
4. Store resolved chains in `_Config`: per chain, its id, label, and an
   **ordered id-list** of steps (encoding defined in the contract — a delimiter
   that cannot occur in an id).

### 2.3 Runtime — the chain runner (reached only via `Core_Run "chain.<id>"`)

The current single dispatch body is refactored into a shared inner routine that
returns a **structured result**, not a bare Boolean:

```
' Runs ONE action under an isolated child context; never dialogs on its own.
Private Function RunAction(actionId, parentCtx, quiet) As StepResult
```

where `StepResult` carries `{ actionId, label, status (ok|failed), message,
errNumber, errText }`. Each step gets a **fresh child context** (its own
`actionId`, owning manifest `reg`, `log`, `quiet` flag, and a result/done sink) —
the chain context is **not** reused as the step context, so handlers never see a
false identity or a muddled log.

Two callers share `RunAction`:

- **`Core_Run "<actionId>"`** — single-button path, `quiet=False`; dialogs ON,
  behavior **unchanged from today**.
- **`Core_RunChain` (Private)** — for each step: `quiet=True` (success dialogs
  suppressed), run via `RunAction`, collect the `StepResult`.
  - **On a step failure:** the step's error boundary **restores speed/lock state
    first** (see §2.5), *then* pops the mid-chain prompt —
    *"Step N of M ‹label› failed: ‹reason›. Continue / Stop?"* — and honors the
    choice (Continue → run remaining; Stop → halt).
  - **At the end:** ONE combined summary built from the collected `StepResult`s
    (`Refresh All — done. 1 Clear ✓  2 Import ✓ (312 rows)  3 Update ✓`), or a
    "stopped at step N of M" summary if halted.

### 2.4 The quiet-suppression convention (explicit, opt-in, honestly scoped)

Suppression is a **defined convention**, not a mechanical edit to today's two
packages:

- A handler reports completion through `Core_Done ctx, msg`: if `ctx("quiet")`
  it writes `msg` to the context's result sink and **suppresses the dialog**;
  otherwise it `MsgBox`es exactly as today.
- **Suppression is guaranteed only for completion/info dialogs routed through
  `Core_Done`** (Core-owned dialogs, and packages that adopt the convention). A
  handler that calls `MsgBox` directly is **non-conforming for chaining** — the
  contract states this, and both bundled packages (`PkgImportExport`,
  `PkgClean`) are migrated to `Core_Done` as part of this work.
- **Not suppressible, by nature:** a step's own *interactive* prompts — e.g.
  `clean.*` actions prompt for a range at run time, and any "are you sure?"
  confirmation. These require human input and will still appear mid-chain.
  **Known limitation:** chains are best composed of non-interactive actions;
  chaining an action that prompts for input means the user is prompted mid-run.
  Documented, not solved here.

### 2.5 Speed/lock ownership across a chain

Core already has a **depth-counted** speed wrapper (`Core_SpeedOn`/`SpeedOff`
nest correctly). The chain design makes its use explicit:

- The chain runner opens the depth-counted wrapper once around the whole run, so
  nested per-step `Core_SpeedOn`/`SpeedOff` calls don't prematurely re-enable
  Excel between steps.
- The original user state — ScreenUpdating, **Calculation (including `xlManual`)**,
  EnableEvents — is captured once and restored on normal exit **and** on any
  step failure **before** the Continue/Stop prompt is shown, so the user is never
  asked a question while Excel is left frozen or with events off.

---

## 3. Grouping / layout

### 3.1 What it does

Assigns each button a **code** like `1A`: number = page, letter(s) = column.
Core builds `Macros 1`, `Macros 2`… pages, placing buttons by code. This is
MacroLayout's model, rebuilt in Core.

### 3.2 Salvaged vs dropped (MacroLayout → Core)

**Salvaged** (logic moves into Core; the file is deleted):

- `ParseCode` — `12AB` → page 12 / column "AB" (multi-digit pages, `AA`+ columns).
- The yellow fill-in grid + **preserve-already-entered-codes** on rebuild.
- Column/stack placement geometry (`Macros N` tabs, columns, stacked buttons).

**Dropped:**

- Its separate hand-maintained **"Master" catalog** — Core's action registry is
  the single source of truth.
- The **`ML_RunById`** placeholder — `Core_Run` is the real dispatcher.

### 3.3 Authoring on Setup — the **Placement** table (+ chains' own Code)

Grouping must list chains, but chains are authored on the same sheet. To keep a
**single Setup pass** (no seed/refresh dance):

- **Actions** are known from the registry at build time → Core seeds a
  **Placement** table, keyed by **action id** (hidden), showing `caption [id]`,
  with a yellow **Code** cell. Prior codes **and prior row order** are preserved
  on rebuild; newly-added actions **append** (existing buttons never jump).
- **Chains** aren't known at build time → each chain carries its **own Code**
  column in the Chains table (§2.2).

Both feed one **id → code** placement map at Finalize.

### 3.4 Placement rules (exact, non-contradictory)

- **No code on any button, anywhere → legacy single "Macros" tab, unchanged.**
  Grouping is purely opt-in; a deployment that ignores it sees today's behavior
  byte-for-byte.
- **Any non-blank code present → a layout is "active":** Core builds numbered
  `Macros N` pages.
  - A **blank-coded** button → **page 1**, rendered in a clearly-labeled
    **"Unassigned"** section placed *after* the assigned page-1 content (so page 1
    doesn't read as a junk drawer). Nothing silently disappears.
  - A button whose Code is the explicit marker **`hide`** → **not placed
    anywhere** (the deliberate power-user opt-out). This is the *only* way a
    button is omitted; a mere blank never omits.
- **Stacking order** within a column = the authoring table's **row order**
  (preserved across rebuilds), which is stable and user-controllable — not
  transient registry iteration order.
- Core's own **Setup / Finalize** admin buttons are never user-groupable; they
  always render in a fixed admin area on page 1.

### 3.5 Finalize review (the agreed cheap UX)

Whatever the codes, Finalize shows a **plain-language review and asks to confirm
before writing config** — this is the project's "validate + report before
activate" philosophy applied to the new tables. It lists:

- resolved chains and their steps;
- problems: unknown/empty/malformed steps, duplicate captions in a reference,
  chain-id collisions, invalid codes;
- the **pages it will build** and, per page/column, the buttons on it;
- the **Unassigned** buttons (→ page 1) and the **hidden** buttons.

Successful parsing never stands in for the user understanding what will happen.

---

## 4. The shared Setup surface + engine-owned config (done *without* a fake package)

Both tables live under one **"Compose Your Buttons"** section on Setup: the
**Chains** table (§2.2), then the **Placement** table (§3.3), then the available-
actions reference (copy-ready tokens).

This is config that **no package manifest declares** — chains and placement are
authored freely by the deployment. It must NOT be smuggled in as a `core`
pseudo-package: Core is required infrastructure, *not* a discoverable
`Pkg<Name>.bas`, and the discovery rule (§3 of CLAUDE.md — Core probes
`Pkg<Name>_Manifest` via `Application.Run` against `_Packages`, never enumerates
modules) must stay pure. A fake `core` entry in `BuildRegistry()` would corrupt
that invariant.

Instead, add an **explicit Core-owned Setup-schema hook**:

- `Core_BuildSetup` renders the "Compose Your Buttons" section from a
  **Core-internal schema** (chains table + placement table), reusing the existing
  low-level drawing helpers (`TintRow`, table renderer, yellow-cell), **outside**
  the package-discovery loop.
- `Core_FinalizeSetup` validates + stores these two tables via a dedicated Core
  step (again outside discovery).
- Dedicated Core readers load chains + placement at run time.

Package discovery and package manifests stay exclusively about packages.

---

## 5. Contract v1.1.0 — additive changes (minor bump)

Written as **language-neutral semantics**, not Excel sheet-building steps. The
Excel realizations (hidden id column, `Macros N` sheets, `Core_*` names) stay in
the platform-abstraction appendix; the reserved-namespace and identity rules are
normative.

1. **New config entity `chain`** — an **immutable id** (independent of its
   editable label) + an ordered list of **action ids** (`steps`). Encoding: an
   id-list joined by a delimiter that cannot occur in an id. Semantics:
   sequential run in list order; per-step success dialogs suppressed; mid-run
   failure prompts continue/stop; one end-of-run summary. Chains are deployment
   config, resolved and validated at activate time.
2. **Reserved id namespace** — `chain.*` is reserved for chains; no action id may
   use it. Dispatch resolves a `chain.*` id to the chain runner, any other id to
   an action.
3. **Engine-owned config** — §5 gains that config MAY originate from the engine
   itself via an explicit engine schema hook (NOT via package discovery); the
   Setup grid layout + keys remain identical across platforms.
4. **Grouping/layout as portable config** — a per-button **placement code**
   (`<page-number><column-letters>`) keyed by button id. Semantics: *no code on
   any button* = single default page; *layout active* = numbered pages, blank →
   default page's **Unassigned** area, explicit **`hide`** → omitted; stacking =
   authoring order; ordering and page derivation are deterministic.
5. **Dispatch/quiet convention (§6 extension)** — a single public dispatch entry
   resolves action-or-chain ids; a defined **quiet/result convention** lets
   composed steps suppress **completion** dialogs (guaranteed only for
   conforming handlers), while interactive prompts and errors are not
   suppressible. Handler binding stays platform-specific.

All additive: no existing key, action id, kind, or semantics changes → v1.0.0
deployments remain valid (§10 SemVer rule). The identity guarantees (2) and the
immutable-chain-id (1) exist specifically to keep a future rename from becoming a
silent semantic break.

---

## 6. Build / verify / cleanup

- **Delete** `src/packages/MacroLayout.bas` (never in the build; logic absorbed).
- New Core procedures (names finalized in planning; must not collide — the build
  fails on duplicate proc names): `Core_RunChain` (Private), `RunAction` (shared
  inner dispatch returning a structured result), `Core_Done`, the code parser +
  placement builder, the Core-owned Setup-schema render/validate/store steps,
  chain + placement config readers, and the id-mint/preserve + reserved-namespace
  checks.
- Migrate `PkgImportExport` + `PkgClean` completion dialogs to `Core_Done`
  (mechanical, no standalone-behavior change).
- `pwsh build/build.ps1` must stay green (Sub/Function balance + no duplicate
  names). Salvaged helpers take `Core`/private names — no `ML_*` carried over.
- No headless VBA test: manual Excel import + run of (a) a defined chain
  (including a stop-on-failure path) and (b) a grouped layout with an unassigned
  and a hidden button, recorded in the PR.

---

## 7. Boundaries (YAGNI — this feature)

Out of scope: nested chains; conditional/branching steps; per-step arguments;
drag-and-drop layout (codes only); reordering beyond authoring row order;
catalog build-out (separate track).

Grouping and chaining ship **together** as one contract bump and one build.

---

## 8. Deferred polish (noted, not built now)

The design panel's aesthete pushed for richer built-page UX. Worth doing later,
but gold-plating relative to "make grouping + chaining correct and portable," and
not required for a usable first version (built pages inherit the existing panel's
caption+description rendering):

- Per-page titles, column headings, a legend, and Previous/Next page navigation.
- A live visual placement preview on the Setup sheet (beyond the Finalize text
  review).
- A visually distinct administrative area for Setup/Finalize.

Filed for a follow-up polish pass once the feature is in use.

---

## 9. Decisions & revisions

### 9.1 Settled in brainstorming (2026-07-24)

| Question | Decision |
|---|---|
| Grouping vs the single control panel; MacroLayout adopt/rewrite | **Rewrite** grouping into Core as an opt-in config layer sourced from the registry; salvage the code-parser / fill-in grid / code-preservation; **delete** `MacroLayout.bas`. |
| Where chains live | **Config** (deployment-specific), via an engine-owned schema hook. |
| Mid-chain error handling | **Ask the user** — Continue / Stop prompt on the failing step. |
| Chained-step success dialogs | **Suppressed**; one combined summary at the end. |
| Chain authoring format | **One row per chain**: name + ordered steps + code, validated at Finalize. |

### 9.2 Hardened by `/deliberate` (2026-07-24) — what changed and why

A three-archetype panel (native / veteran / aesthete) reviewed the first draft
and converged unanimously on several **correctness** fixes (not preferences —
defects), all now folded in above:

- **Key on stable ids, never captions** — chain steps store action ids;
  placement rows are keyed by button id; authoring uses `Label [id]` tokens.
  (Captions are presentation; keying on them makes an edit rewire an automation.)
- **Immutable chain id, label editable separately** — a label-derived slug was a
  mutable dispatch key and a §10 major-break-in-disguise.
- **`Core_RunChain` is Private; `Core_Run` stays the sole public entry;** the
  `chain.` namespace is genuinely reserved (no action may use it).
- **Structured step results + a fresh child context per step** — not a bare
  Boolean on a reused context.
- **Explicit, depth-aware speed/lock ownership** — restore original state
  (including `xlManual`) before the failure prompt.
- **Quiet suppression as an explicit, honestly-scoped convention** — guaranteed
  only for `Core_Done`-routed completion dialogs; interactive prompts/errors are
  not suppressible; both bundled packages migrated.
- **Engine-owned config via an explicit Core hook, NOT a fake `core` package** —
  keeps the `Application.Run` package-discovery invariant pure.
- **Exact, non-contradictory layout semantics** + language-neutral contract text.

A second focused panel decided the **blank-code disposition** (also unanimous):
**blank → page 1** in a labeled **Unassigned** area, with an explicit **`hide`**
marker as the opt-out — the only rule consistent with the project's
graceful-fallback, nothing-silently-breaks intent. "Blank = omit" was rejected as
"a hostile meaning for a blank cell"; "require a code" as against opt-in fallback.

Panel cost: 6 Codex calls (2 topics × 3 archetypes, one round each), mid-tier.
