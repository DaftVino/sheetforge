# Consolidated Engine Architecture — Draft

Design for a reusable, config-driven macro system: many grouped macro **packages** plus one **setup** engine that configures whichever packages are present. Excel/VBA is the primary target; Google Sheets is secondary; other systems tertiary.

---

## 1. Testing your concept (the honest read)

Your idea — *multiple package `.bas` files + one setup `.bas` that runs setup for any package the user imports* — is directionally right, and in Excel it's closer to "just works" than in most languages. But there's one load-bearing gap it glosses over. Here's the breakdown.

**Where you're right.** VBA standard modules all share a single global namespace inside a project. Import ten `.bas` files and every `Public` sub/function in them is instantly callable from every other module — no linking, no imports, no references. So "packages that can see the setup engine, and vice versa" is free. Your instinct that this is straightforward is correct *for the wiring*.

**Where it's incomplete — the gap.** A *generic* setup engine can't do anything useful with a package it knows nothing about. Setup needs to know: what config fields does this package need? What actions does it expose? What captions/help text go on the buttons? Your current `SetupImportExportMacros.bas` "knows" all of that because it's **hardcoded** (the `k_export_tabs`, `k_import`, etc. keys are baked into `CreateSetupTab`). A generic engine can't hardcode that — the packages have to *tell* it. That means a **contract**: every package must declare itself in a shape Core understands. That contract is the piece your concept is missing, and it's the whole ballgame.

**The second gap — discovery.** Even with a contract, how does the setup engine find out *which* packages got imported? VBA cannot enumerate its own loaded modules without enabling **"Trust access to the VBA project object model"** — the exact setting your current file deliberately avoids needing. So Core can't just "look around" for packages. It needs an explicit registry it maintains itself (details in §4).

**Where it breaks if unmanaged — three failure modes.**
1. **Name collisions.** Two packages each defining `Public Sub RemoveDuplicates` → an "Ambiguous name detected" error the moment either is called unqualified. A naming convention is mandatory, not optional (§5).
2. **Helper duplication.** Every package wants the speed wrapper, the last-row finder, the error handler, the config reader. If each `.bas` carries its own copy, you get either collisions (if `Public`) or bloat and drift (if `Private`). Shared plumbing must have a single owner (§2, §5).
3. **Cross-platform.** VBA and Google Apps Script are different languages on different runtimes. "One codebase, any spreadsheet" is not achievable. What's portable is the *design*, not the code (§8).

**Verdict.** Keep your structure — it's the right shape. Add three things it's missing: a **package contract (manifest)**, a **Core-owned registry + shared helpers**, and a **portable spec** separate from the code. Those three turn "a folder of `.bas` files" into an actual plug-and-play platform.

---

## 2. The layered model

Three layers. The dependency arrow points one way: packages depend on Core; Core depends on nothing.

```
┌─────────────────────────────────────────────────────────────┐
│  LAYER 2 — CONFIG SURFACE  (the only designed UI)            │
│  Setup sheet  ·  hidden _Config  ·  _Registry / _Packages    │
│  Rendered dynamically from installed packages' manifests     │
└───────────────▲─────────────────────────────────────────────┘
                │ reads/writes
┌───────────────┴─────────────────────────────────────────────┐
│  LAYER 0 — CORE  (required, exactly one, you already own 80%)│
│  • Config engine (read _Config, fall back to defaults)       │
│  • Setup renderer (builds Setup sheet FROM manifests)        │
│  • Button/menu builder (BuildMacroTab, generalized)          │
│  • Dispatcher (Core_Run "<actionId>")                        │
│  • Shared helpers: speed wrapper, last-row/col, error+log,   │
│    file finder, type coercion                                │
│  • Registry (which packages are installed)                   │
└───────────────▲─────────────────────────────────────────────┘
                │ each package registers + calls up into Core
     ┌──────────┼──────────┬───────────────┬──────────────┐
┌────┴────┐ ┌───┴────┐ ┌───┴─────┐ ┌───────┴──────┐ ┌─────┴─────┐
│ IMPORT/ │ │ CLEAN  │ │ EXPORT  │ │ RECONCILE /  │ │  (future  │
│ EXPORT  │ │        │ │         │ │  VALIDATE    │ │  packages)│
│ (your 4 │ │        │ │         │ │              │ │           │
│ buttons)│ │        │ │         │ │              │ │           │
└─────────┘ └────────┘ └─────────┘ └──────────────┘ └───────────┘
        LAYER 1 — PACKAGES  (optional, many, thin)
```

**Core is not new work — it's an extraction.** Your `SetupImportExportMacros.bas` already contains ~80% of Core: `CfgSheet`/`SheetSingle`/`SheetList1`/`SheetTable3` (config engine), `BuildMacroTab` (button builder), `CreateSetupTab`/`FinalizeSetupTab` (Setup renderer + validation), `FindSourceFile`/`CoerceVal`/`SheetExists` (helpers), and the speed/error pattern in every button. The job is to *lift those out* of the dashboard-specific parts and make them manifest-driven.

**Packages become thin.** Your four buttons (Clear, Import, Update Ref_Master, Export) stop being the whole file and become one package — `ImportExport` — that declares a manifest and calls Core for all plumbing.

---

## 3. The package contract (the keystone)

Every package module exposes **exactly one** well-known entry point: a manifest sub named by convention. Core calls it to learn everything about the package.

```vba
' In module "PkgClean"
Public Sub PkgClean_Manifest(reg As Object)   ' reg = a Scripting.Dictionary Core hands in
    ' --- identity ---
    Core_DeclarePackage reg, id:="clean", name:="Clean & Standardize", version:="1.0.0"

    ' --- config fields this package needs on the Setup sheet ---
    '            key,               label,                              type,   default
    Core_AddField reg, "clean_range",  "Range to clean (blank = used range)", "range", ""
    Core_AddField reg, "clean_keycols","Dedupe key columns (e.g. 1,2,4)",    "text",  "1"

    ' --- actions this package exposes (become buttons) ---
    '             actionId,            caption,                 entryPointSub,          help
    Core_AddAction reg, "clean.dupes",  "Remove Duplicates",    "PkgClean_RemoveDupes", "De-dupe on the key columns."
    Core_AddAction reg, "clean.trim",   "Trim & Clean Text",    "PkgClean_TrimClean",   "Strip spaces & non-printing chars."
    Core_AddAction reg, "clean.blanks", "Fill Blanks Down",     "PkgClean_FillDown",    "Fill blank cells with the value above."
End Sub
```

Each action is an ordinary sub with a **fixed signature** so the dispatcher can call any of them the same way:

```vba
Public Sub PkgClean_RemoveDupes(ctx As Object)
    Core_SpeedOn                                   ' shared speed wrapper
    On Error GoTo Fail
    Dim rng As Range: Set rng = Core_ResolveRange(ctx, "clean_range")
    Dim keys As Variant: keys = Core_CfgList(ctx, "clean_keycols")
    rng.RemoveDuplicates Columns:=keys, Header:=xlYes
    Core_Log ctx, "Removed duplicates on " & rng.Address
    Core_SpeedOff: Exit Sub
Fail:
    Core_Fail ctx, "clean.dupes", Err               ' shared error handler + restore
End Sub
```

The point: **a package declares data (fields + actions) and implements small subs.** All ceremony — reading config, speed, errors, logging, button creation — lives in Core. Adding a package never touches Core's code.

---

## 4. Discovery without VBIDE trust

The constraint: Core can't enumerate loaded modules without the trust setting you avoid. So Core maintains its own list and *probes* by name.

**Mechanism.** A `_Packages` sheet (Core-managed) holds one package id per row. To rebuild the Setup sheet, Core walks that list and calls each package's manifest by name, tolerating absence:

```vba
' Core: run a package's manifest if that module is actually present
Private Function TryManifest(pkgId As String, reg As Object) As Boolean
    On Error Resume Next
    Application.Run "Pkg" & ProperId(pkgId) & "_Manifest", reg   ' e.g. PkgClean_Manifest
    TryManifest = (Err.Number = 0)
    On Error GoTo 0
End Function
```

`Application.Run` on a missing macro raises a catchable error, so Core can probe safely. Two ways to populate `_Packages`:

- **"Add Package" picker (recommended).** Core keeps a small candidate list of known package ids and probes each; the ones whose manifest runs cleanly are "installed" and get offered in a picker. User ticks the ones they want active. Zero trust setting required.
- **Self-registration.** The first time any action from a package runs, it appends its id to `_Packages`. Simpler, but a package is invisible on the Setup sheet until it's been run once — worse UX for setup.

Recommendation: **maintained `_Packages` list + probe.** It's the only zero-trust option that shows all installed packages *before* the user runs anything, which is what a Setup screen needs.

---

## 5. Naming & no-duplication rules

Non-negotiable, because the global namespace is flat:

- **Public entry points are prefixed by package:** `Pkg<Name>_<Verb>` (e.g. `PkgExport_ToPdf`). The manifest is always `Pkg<Name>_Manifest`. This guarantees uniqueness and makes `Application.Run` name-building deterministic.
- **Core's public API is prefixed `Core_`** (`Core_Run`, `Core_SpeedOn`, `Core_CfgList`, `Core_ResolveRange`, `Core_Log`, `Core_Fail`).
- **Everything else is `Private`.** Package-internal helpers stay inside the module and can share short names freely.
- **Shared helpers have exactly one home: Core.** Packages never re-declare the speed wrapper, last-row finder, error handler, or config reader. If a package needs it, it calls `Core_*`. This is what makes packages thin and kills drift.
- **Version key per package** (in the manifest) so Core can warn if a package expects a newer Core than is installed.

Consequence, stated plainly: **packages are not standalone.** A package `.bas` requires Core to be present. That's the deliberate trade — you give up "drop any single file in and it works alone" to get "no duplicated plumbing and no collisions." For a library that's the right trade; document Core as a hard dependency.

---

## 6. Dispatcher (generic button wiring)

Buttons should **not** point at package subs directly — that would re-hardcode the wiring every time. Instead every button's `OnAction` points at one Core entry with the action id baked in:

```vba
btn.OnAction = "'Core_Run ""clean.dupes""'"
```

```vba
Public Sub Core_Run(actionId As String)
    Dim ep As String: ep = LookupEntryPoint(actionId)      ' from the manifests
    Dim ctx As Object: Set ctx = BuildContext(actionId)     ' config + logging handle
    If ep = "" Then MsgBox "Action not installed: " & actionId, vbExclamation: Exit Sub
    Application.Run ep, ctx                                  ' calls PkgClean_RemoveDupes, etc.
End Sub
```

One wiring path. Add a package → its actions appear as buttons → each button calls `Core_Run` → Core dispatches. No button code is ever written per package.

---

## 7. Distribution model — a real fork

How the pieces reach a user's workbook. This genuinely changes the build, so it's a decision, not a default.

**Option A — many files, mix-and-match.** Ship Core + each package as separate `.bas`. Users import Core plus whichever packages they want. Maximally modular; good if you envision a catalog people pick from. Cost: N manual imports (VBIDE → Import File, once per file), plus dependency/version management on the user's side. For 70-store, non-technical users, every extra import is friction and a support call.

**Option B — modular source, single built file (recommended for your org).** Keep Core and each package as separate `.bas` *during development*. Ship **one** built `.bas` per deployment (Core + the packages that deployment needs, concatenated by a tiny build step). The Setup config *enables/disables* packages, so the same shipped file can expose different action sets per site. Users do **one** import. You keep clean modular source; they get monolithic simplicity.

**Option C — both.** Author modularly; produce both a bundled file and individual package files, so power users can cherry-pick and regular sites get the bundle.

My recommendation is **B for your 70 stores** (one import, config-gated features) with the source kept modular so A/C stay available later. This is the "modular authoring, monolithic delivery" pattern.

---

## 8. Cross-platform reality (pushback)

Blunt version: **you cannot write one set of macros that runs in Excel and Google Sheets.** They are different languages on different runtimes. Planning for "shared code" will waste effort; plan for "shared *design*, separate code."

| | Excel (primary) | Google Sheets (secondary) | LibreOffice (tertiary) | Python/xlwings (tertiary) |
|---|---|---|---|---|
| Language | VBA | Apps Script (JavaScript) | LibreOffice Basic (UNO API) | Python |
| Unit | `.bas` module | `.gs` file | Basic module | `.py` |
| Distribution | Import into VBA project | Bound/container script or Add-on | Import into Basic IDE | Package/venv |
| UI to launch | Form-control buttons / ribbon | Custom menu (`onOpen`) / sidebar | Buttons / menu | External / xlwings buttons |
| Prompt/alert | `MsgBox` | `ui.alert` / `ui.prompt` | `MsgBox` (similar) | none native |
| File access | local folder next to workbook | Google Drive (no local folder) | local folder | filesystem |
| Runtime limit | none | ~6 min consumer / 30 min Workspace | none | none |

**What actually is portable:** the *contract and the config*, because those are data and spec, not code — the config field keys, the action ids, the action semantics ("`clean.dupes` de-dupes on the key columns"), and the Setup sheet layout (both Excel and Sheets are grids). None of that is language-specific.

**Strategy.** Make the **config schema + action contract the single source of truth** — write it once as a language-neutral spec (a JSON or markdown file in this project). Then implement it per platform *against the same spec*: VBA now, Apps Script later, each rendering the same Setup sheet and honoring the same action ids and config keys. A workbook configured in Excel and one configured in Sheets would have the *same Setup layout and the same keys* — only the engine behind them differs.

**Design consequence to bake in now:** keep platform assumptions out of the *contract*. Your current file assumes "source files live in the same folder as the workbook" — true in Excel, meaningless in Sheets/Drive. In the portable spec, that becomes an abstract `source_location` field the engine interprets per platform (local path in Excel, Drive folder id in Sheets). Abstracting Excel-isms out of the contract now is what makes the Sheets port a rewrite of *engines*, not of *design*.

The upside for you: the **Setup sheet is the most portable asset in the whole project** — which lines up exactly with your read that the Setup page is the one place design carries as much weight as code. Time invested there pays off on every platform.

---

## 9. The Setup page (the one designed surface)

Principles, generalizing what your `CreateSetupTab` already does well:

- **Package-driven rendering.** Sections appear only for installed/enabled packages (from the manifests). No dead fields.
- **Two tiers on one sheet.** A "Simple" band (only the must-fill yellow cells, with plain-language labels) up top; an "Advanced" band (optional/expert fields, type maps, offsets) below or on a second sheet. Regular users finish the top and stop; power users drop into Advanced. You already do the spirit of this with yellow input cells + hidden machine keys in column A — formalize the two tiers.
- **Live validation before activate.** Keep your `FinalizeSetupTab` pattern: validate against the actual workbook (tabs exist, files match) and report before writing `_Config`. Extend it to validate per-package required fields from the manifest.
- **Reference helpers.** Keep your `PopulateReferenceLists` (tabs in this book / files in this folder) — it's a great "reduce typos" touch. Make it a Core service packages can request.
- **Graceful default fallback.** Keep the behavior where absent `_Config` = built-in defaults, so an unconfigured import still runs. Generalize it: each package's manifest supplies its own defaults.
- **Visual polish lives here and only here.** Section headers, color, status column, hidden keys — this is where the design budget goes.

---

## 10. Suggested build order

1. **Extract Core** from `SetupImportExportMacros.bas`: config engine, helpers, button builder, Setup renderer — plus new `Core_Run` dispatcher, `Core_DeclarePackage`/`Core_AddField`/`Core_AddAction`, and the `_Packages` registry.
2. **Refactor your four buttons** into a `PkgImportExport` module with a manifest. Nothing about the sample dashboard changes for the user; you've just proven the contract on real code.
3. **Add a second tiny package** (`PkgClean`: RemoveDupes + TrimClean) to prove the contract generalizes and buttons appear with zero new wiring.
4. **Freeze the config schema + action contract** as the portable spec file (JSON/markdown) in this project.
5. **Only then** start the Apps Script port, mirroring the spec.

Do steps 1–3 before worrying about Sheets at all — if the manifest/dispatcher/registry loop works cleanly for two packages in Excel, the platform is proven and porting is mechanical.

---

## Open decisions

Two forks change what I draft next; everything above holds regardless of how you answer:

1. **Distribution model** — single built file with config toggles (simplest for 70 stores) vs. many mix-and-match files vs. both.
2. **Cross-platform scope now** — Excel-first with a locked portable spec (port later) vs. design Excel + Sheets in parallel now vs. Excel-only for now.
