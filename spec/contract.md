# SheetForge portable contract

**Contract version: 1.2.0** (2026-07-24) · SemVer, independent of the
repo `VERSION` and of per-package versions.

This is the **language-neutral source of truth** for SheetForge's config schema
and action contract. The Excel/VBA build (`src/core/Core.bas` + `src/packages/`)
is the **reference implementation**; the planned Apps Script / Google Sheets port
implements the *same* contract against a different engine. Per
`docs/engine-architecture-draft.md` §8, only the **contract, the config, and the
Setup layout are portable** — the engine code is not.

**Normative rule.** Where this document and the reference code disagree, **this
document wins**, and the code is corrected to match. Where a platform cannot
realize a behavior identically, it follows the *portability note* for that item.
Behavior described here is what every platform must reproduce; anything labeled
"reference implementation" or "Excel" is one platform's realization.

---

## 1. The manifest object

Each package exposes exactly one entry point — a manifest declaration — that
tells the engine everything about the package. It declares:

| Field | Meaning |
|---|---|
| `id` | Package id — lowercase short token, globally unique (`importexport`, `clean`). |
| `name` | Human display name (`"Import / Export"`). |
| `version` | Package SemVer (`"1.0.0"`), independent of this contract's version. |
| `fields[]` | Ordered list of config fields (see §2). Render order = declaration order **within a tier** (the advanced tier is grouped into one band, §4). |
| `actions[]` | Ordered list of actions the package exposes (see §6). |

The engine discovers packages from a maintained registry and probes each one's
manifest by convention — **never** by enumerating code modules. (Reference
implementation: `Application.Run "Pkg<Name>_Manifest"`, so no "Trust access to
the VBA project object model" is required.)

---

## 2. Field model

A field is the schema for one configurable value on the Setup surface. Every
field is the following record:

```json
{
  "key":      "string — machine key, unique across ALL enabled packages",
  "label":    "string — human label shown on the Setup surface",
  "kind":     "one of the kinds enumerated in §3",
  "tier":     "\"simple\" | \"advanced\"  (see §4)",
  "default":  "string — scalar default (kinds text/int/sheet/file/bool)",
  "headers":  "string — column header(s); list = one header, table = \"|\"-delimited",
  "defaults": ["ordered seed rows for kind list/table; see §3"],
  "help":     "string — optional plain-language note shown under the field",
  "usedBy":   "string — optional comma-separated action ids this field configures (v1.2.0)"
}
```

- **`key` uniqueness is contract-critical.** The runtime config is a single flat
  key→value store shared by all enabled packages, so every `key` must be unique
  across every package a deployment enables. Keys are opaque tokens; namespacing
  them (`k_*` for importexport, `clean_*` for clean) is the convention that keeps
  them unique.
- **Order is significant.** Fields render in declaration order **within their
  tier** (the advanced tier is collected into one band, §4, so an advanced field
  declared before a later simple field still renders after it); `list`/`table`
  seed rows preserve their order (§3).
- `default` carries the scalar default for scalar kinds; `defaults[]` carries the
  seed rows for `list`/`table`. A `section` has neither.
- **`usedBy` (v1.2.0, optional).** A comma-separated list of action `id`s this
  field configures. A field that declares nothing is **package-wide** and always
  renders in the working area. A field is relegated only when **every** action it
  names is `off` (§6) — an action that is `chains-only` still runs, so it still
  needs its settings. Relegated means **moved, never omitted**: the engine still
  renders the field, still stores it, and still falls back to its default; it is
  simply placed in one collapsed "settings for actions you haven't turned on" band
  at the end of the Setup surface. An unrecognised id in `usedBy` counts as in-use,
  so a typo cannot make a field unreachable.
  **Portability note:** collapsing is a presentation affordance (reference: Excel
  row outlining). A platform that cannot collapse MUST still render the band,
  visibly and last. Omitting the fields is non-conforming — it would force a
  second activate cycle and discard un-activated input.

---

## 3. Kind vocabulary (canonical)

These are the only kinds the contract defines. Each row gives the meaning, how
the value is encoded as text in the config store, and how each platform realizes
it.

| kind | data | text encoding | Excel realization | Portability note |
|---|---|---|---|---|
| `section` | none (presentation) | — | a styled header row; no key, no value | A visual divider on any grid; carries no config. |
| `text` | string | the string verbatim | free text cell | Universal. |
| `int` | integer | decimal string (e.g. `-1`) | coerced with the platform's integer parse | Universal. |
| `sheet` | a worksheet/tab name | the tab name string | validated at Finalize to exist in the workbook | Excel worksheet ↔ Sheets tab. |
| `file` | a **filename prefix** | the prefix string | newest file in the deployment's data location whose name **starts with** the value, ignoring a trailing ` (1)`/` (2)` copy suffix | The data location is the workbook's containing folder in Excel; a Drive folder in Sheets. See §7. |
| `bool` | yes/no | **reserved** | — | **Reserved: defined but not yet used by any package.** Encoding (`yes`/`no` vs `TRUE`/`FALSE`) is fixed on first use and this table updated then. |
| `list` | ordered single values | one value per row | consecutive value cells (column C), one per row | Ordered rows on any grid. |
| `table` | ordered N-column rows | each row `"a\|b\|c"`; up to 3 columns | value cells across columns C/D/E, one row per record | Ordered rows on any grid. |

Notes:
- **`range` is intentionally NOT a kind.** An earlier design draft proposed a
  `range` field kind (and a `Core_ResolveRange` reader); neither was ever
  implemented. PkgClean instead answers "which range?" with an **ad-hoc range
  picker at action time**, so no `range` config field exists. Ports MUST follow
  this: range-scoped actions prompt for their range at run time, they do not read
  it from config.
- A `table` header string is `"|"`-delimited (not comma) because a header may
  itself contain commas (e.g. `"Column types (T,N,D,...)"`). Row values in
  `defaults[]` are `"|"`-delimited for the same reason.
- Import **column-type codes** (`T`/`N`/`D`/`DT`/`B`/`G`) are a *data format
  inside* the `importexport` package's `k_import` table (§8), **not** field
  kinds. They are defined where that field is defined.

---

## 4. Tier semantics

Every field carries a `tier`:

- **`simple`** (default) — must-fill, plain-language fields. Rendered first; a
  regular user completes these and stops.
- **`advanced`** — optional/expert fields, grouped under a single advanced band
  per package, labeled to the effect of *ADVANCED — most users can leave these
  as-is* (the reference renders that exact wording; a port supplies its own).

**Degrade rule (normative):** an unknown or mis-cased tier value MUST degrade to
`simple` (still shown), never vanish. Reference implementation lower-cases and
trims the tier and treats anything other than `advanced` as `simple`.

`section` fields belong to the `simple` pass for placement purposes.

---

## 5. Config storage & read semantics

**Two surfaces:**

- **Setup** — the one designed, user-facing grid. Machine `key`s live in a hidden
  column (reference: column A); values live in the value column(s) (reference:
  column C, plus D/E for `table`). Section headers, tier bands, and help notes
  are presentation only.
- **Config store** — a hidden, validated key→value store the engine reads at run
  time (reference: a hidden `_Config` sheet written by Finalize). Only validated
  answers reach it.

**Read + fallback (normative):**

- The engine reads each `key` from the config store. If the store is absent or
  has no entry for a `key`, the engine falls back to that field's manifest
  **`default`** (scalar) or **`defaults[]`** (list/table). An unconfigured
  deployment therefore still runs on the declared demo defaults.
- Reference read API (types returned): scalar string, integer, ordered list,
  table — `Core_Cfg` / `Core_CfgLong` / `Core_CfgList` / `Core_CfgTable`.

**Block bounds (normative for grid platforms):**

- A fixed number of input rows is drawn per `list`/`table` block (reference:
  `MAX_LIST` = 20).
- The **last** `list`/`table` block on the Setup surface MUST be bounded by a
  sentinel key row (reference: `_core_end`), so an upward "last used row" scan
  cannot swallow or drop the block's trailing rows.

**Validation before activate (normative):** Finalize validates answers against
the actual workbook (tabs exist, files match, per-package required fields
present) and reports problems *before* writing the config store.

**Portability note:** the config store being a hidden sheet is Excel-specific.
Any platform may use an equivalent key→value store, but the **Setup grid layout
and the keys are identical across platforms**. Values are exchanged as strings;
lists are ordered rows; table rows are columns spread across the value columns
(never `"|"`-joined in storage — the `"|"` form is only for `defaults[]` seeds).

### Composite actions — chains (v1.1.0)

A **chain** is deployment config (not manifest data) that composes existing
actions into one button:
- **id** — an immutable, opaque token `chain.<token>`, minted once and preserved
  across label edits. NEVER derived live from the label. `chain.*` is reserved:
  no action id may use it.
- **label** — the button caption; freely editable without changing identity.
- **steps** — an ordered list of **action ids** (not captions). Stored as an
  id-list joined by a delimiter that cannot occur in an id (reference: newline).
- **semantics** — run steps in list order; each step's completion dialog is
  suppressed; on a step failure the engine prompts the user Continue/Stop; one
  combined summary is shown at the end. Chains resolve + validate at activate time.

**Engine-owned config (v1.1.0):** config MAY originate from the engine itself
(chains, placement) via an explicit engine schema hook, not only from package
manifests. It is authored on the same Setup grid with the same key/value model,
and MUST NOT be discovered through the package-probing mechanism (§1).

**Config schema marker (v1.2.0, normative):** the engine writes a machine-owned
schema-version token into the config store alongside the engine-owned blocks,
outside any user-editable cell (reference: the hidden per-row id column, on the
placement block's key row). Readers branch on that token, **never on a visible
label** — a display header is presentation, is user-editable, and must never
decide whether saved configuration is preserved or discarded.

Migration is **fail-closed**. When the marker is absent the store predates v1.2.0
and is read with the older column layout; every value recovered that way MUST be
validated against its grammar (for placement, a valid code or `hide`), and any
conflict — an unknown marker, or a recovered value that does not validate — MUST
abort the rebuild with a report and leave the existing surface untouched. Silently
rebuilding with blanks is non-conforming.

---

## 6. Action contract

An action is one thing a package *does* (it becomes a button). Each action is:

| Field | Meaning |
|---|---|
| `id` | Dotted `pkg.verb` token, globally unique (`importexport.import`, `clean.dupes`). |
| `caption` | Button label (may carry an ordinal, e.g. `"2.  Import Source Data"`). |
| `help` | Plain-language description shown on/near the button. |
| `category` | Optional display category (v1.2.0). Defaults to the package `name`. Groups actions on the generated index and heads the columns of a button page; deliberately decoupled from the package, so two packages may feed one category and one package may span several. |
| `availability` | Deployment config, not manifest data (v1.2.0): one of `button`, `chains-only`, `off`. Defaults to `button`. |
| `semantics` | **Normative** description of the action's behavior (§8). This is what every platform reproduces. |

**Handler binding is platform-specific and NOT part of the portable contract.**
The reference implementation binds each `id` to a VBA sub name (its `entry`), but
the contract fixes only the `id` and its `semantics`; each platform maps the `id`
to its own handler.

**Dispatch (normative):** a single engine entry point resolves an `actionId` to
the platform's handler and invokes it with a **context** carrying at least: the
`actionId`, the package manifest, and a log sink. (Reference: `Core_Run
"<actionId>"`; context exposes `actionId`, `reg`, `log`.) Each handler wraps its
own work in the engine's shared speed/lock service and reports failures through
the shared error handler — reference: the handler calls `Core_SpeedOn` /
`Core_SpeedOff` around its body and `Core_Fail` in its error path; the dispatcher
additionally runs a fail-safe restore if a handler errors before cleaning up. A
package implements only its own logic plus these shared calls.

**Dispatch (v1.1.0):** the single public dispatch entry resolves either an action
id or a reserved `chain.*` id. A **quiet convention** lets composed steps suppress
their **completion** dialogs: a handler reports completion through the engine's
done-sink, which suppresses the dialog when the context is quiet. Suppression is
guaranteed only for completion dialogs routed through that sink; interactive
prompts (e.g. a range picker) and error dialogs are not suppressible.

**Grouping / layout (v1.1.0):** each button (action or chain) may carry a
**placement code** `<page-number><column-letters>` (e.g. `1A` = page 1, column A;
`12AB` = page 12, column "AB"). Semantics: if NO button anywhere has a code, a
single default button page is built (legacy behavior). If any code is present, a
layout is active: numbered pages are built; a blank-coded button goes to page 1
(the default page)'s **Unassigned** area; the explicit code **`hide`** omits a button.
Same-code buttons stack in authoring order. Core's own admin controls are never
user-placeable.

**Per-action availability (v1.2.0):** each action carries a deployment-level
`availability`, authored on the same engine-owned surface as the placement code
(§5). It has exactly three values and defaults to `button`, so a deployment that
predates v1.2.0 is unchanged.

| Value | Button drawn | Runnable directly | Runnable as a chain step |
|---|---|---|---|
| `button` (default) | yes | yes | yes |
| `chains-only` | no | **no** | yes |
| `off` | no | **no** | **no** |

- **`off` means off.** The engine MUST refuse to dispatch an `off` action, and
  Finalize MUST **block activation** when any chain references one, naming the
  chain. An implementation where a switched-off action still executes inside a
  chain is non-conforming: a user who disables a destructive action and then
  watches it run has been actively misled.
- **`chains-only`** exists so "reachable only as part of a composed button" is an
  explicit, authored state rather than a side effect of hiding something. A direct
  dispatch of a `chains-only` id MUST be refused.
- An action that is not `button` **stays listed** on the authoring surface, so it
  never becomes undiscoverable. Only `off` relegates that action's fields (§2).
- **An unrecognised value MUST block Finalize** and report the accepted values. It
  MUST NOT fall back to `button` — a user who typed `of` meaning `off` would
  otherwise be handed a live destructive button. A **blank** cell is not
  unrecognised: it means "never set" and reads as `button`.
- **Portability note:** the reference implementation presents the three values as a
  validated dropdown on the authoring grid. A platform without cell validation
  still parses the same three tokens and still blocks on anything else.

**Auto-layout (v1.2.0) — supersedes the "no codes" clause of the v1.1.0 grouping
rule above.** When **no** button carries a placement code, the engine distributes
buttons automatically. A page is a **capacity unit, not a category unit**:

- a page holds up to **three category-headed columns**; a column holds up to
  **eight** buttons, all of one category;
- small categories therefore **share** a page;
- a category needing more than one column takes consecutive columns, and a
  category that does not fit in a page's remaining columns starts a new page
  rather than straddling one — splitting a category across tabs destroys the
  grouping that putting it in one place buys;
- a category larger than a whole page continues onto further pages.

The v1.1.0 behaviour — a single default button page holding everything — is
withdrawn: it does not survive contact with a large catalog, which is what this
version exists to support. **When any code is present the v1.1.0 rule is unchanged**
and codes remain a complete override.

**Generated index (v1.2.0):** the engine **always** generates an index surface,
even for a single page, so the workbook has one shape at every scale. It lists one
entry per built page with its title, its button count and a link; a full **action
catalog** (caption, description, and the page each button is on); a link to the
admin surface; and a legend explaining the placement code and availability. Every
built page carries a persistent link back to the index and links to its
previous/next page. Reference implementation: a `Macros Home` sheet; that name is
**reserved** (§7).

**Description legibility (v1.2.0, normative for grid platforms):** an action's
`help` text is displayed with its button and MUST NOT be clipped. Reference
implementation: an unmerged cell-card grid — button anchored to a content cell,
the description in the cell directly beneath it, that row auto-fitted. Merged
cells are forbidden on this grid precisely because Excel's AutoFit ignores them.

---

## 7. Naming rules (contract-level)

- **Package id** — lowercase short token, unique.
- **Field key** — opaque token, **unique across all enabled packages** (single
  shared config namespace). Namespace by convention (`k_*`, `clean_*`).
- **Action id** — `pkg.verb`, unique.
- **Reserved:** the id namespace `chain.*` is reserved for engine chains; no package action id may use it.
- **Reserved:** the display category `Chains` (engine-generated, holds chain
  buttons) and, in the Excel reference implementation, the sheet name
  `Macros Home`. A platform substitutes its own index surface name.
- Reference-implementation binding names (`Pkg<Name>_<Verb>` entry points,
  `Pkg<Name>_Manifest`, `Core_*` API) are an Excel/VBA detail, not part of the
  portable contract.

---

## 8. Frozen registry

The concrete config + actions a conforming implementation MUST reproduce.
Defaults shown are the shipped **generic demo** values (a deployment overrides
them via its own Setup answers).

### Package `importexport` — "Import / Export" (v1.0.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_export_base` | text | simple | `My Report` | Base name of the exported file; the run date (shifted by `k_export_offset`) is prepended → `<date> <base>.xlsx`. |
| `k_settings_sheet` | sheet | advanced | `README` | Name of the hidden settings/thresholds sheet that drives conditional-formatting; kept (hidden) inside the export so colors still work. |
| `k_export_offset` | int | advanced | `-1` | Days to shift the date placed in front of the export filename (`-1` = yesterday, `0` = today). |
| `k_export_tabs` | list | simple | `Summary`, `Details`, `Trends`, `Outliers`, `Breakdown` | The finished tabs to copy into the export, **in order**. |
| `k_import` | table | simple | see below | One row per input tab. Columns: **Input tab** (tab in this workbook to paste into) · **Source file** (name-prefix, §3 `file` rule) · **Column types** (per-column type-map codes). |
| `k_master_prefix` | file | simple | `Master List Import` | Name-prefix of the master-list file used to refresh the master destination tab. |
| `k_master_dest` | sheet | simple | `Ref_Master` | Destination tab the master list is written into. |
| `k_master_sheet` | text | advanced | `Current Master` | Worksheet name to read from inside the master-list file. |
| `k_master_cols` | list | simple | `Store ID`, `Location Name or Address`, `Entity ID`, `Mgr (DM) Group`, `Mgr (Loc) Name` | Column headers to copy from the master file, **in order**. The FIRST is the key column and is always kept as text. A deployment replaces these with its own headers. |

`k_import` header: `Input tab (in this workbook)|Source file (name starts with...)|Column types`.
Demo default rows (`|`-delimited): `IN_Data1|Source File A|T,T,T,T,T,T,T,T,N,N,N,N,N,N,N` · `IN_Data2|Source File B|T,T,T,T,T,T,T,T,N,N,N,N,N,N,N` · `IN_Data3|Source File C|T,T,T,T,T,T,T,N,N,N,N` · `IN_Data4|Source File D|T,T,T,T,T,T,T,T,T,N,N,N,N,N`.

**Column-type codes** (one per column, left→right, comma-separated), used in
`k_import`'s "Column types":

| code | meaning |
|---|---|
| `T` | Text — preserves IDs, codes, leading zeros |
| `N` | Number (strips `$` `,` `%`) |
| `D` | Date |
| `DT` | Date + time |
| `B` | Yes/No (boolean) |
| `G` | General / auto-detect |

**Actions**

| id | caption | semantics |
|---|---|---|
| `importexport.clear` | `1.  Clear Input Tabs` | Delete data from row 2 down on every input tab; keep headers + formatting. |
| `importexport.import` | `2.  Import Source Data` | Clear the input tabs, then for each `k_import` row import the newest matching source file (prefix rule, §3), coercing each column per its type-map codes. One confirmation dialog. |
| `importexport.updateMaster` | `3.  Update Ref_Master` | Refresh `k_master_dest` from the newest `k_master_prefix` file: read `k_master_sheet`, copy the `k_master_cols` columns in order, keep the first column as text. |
| `importexport.export` | `4.  Export Dashboard (.xlsx)` | Write a **values-only** snapshot of the `k_export_tabs` (in order) to a dated file (§ base/offset). Flatten dynamic-array spills, legacy CSE arrays, and formulas to static values; keep formatting, filters, and conditional formatting; break external links. |
| `importexport.suggestTypes` | Suggest Import Types | Setup helper: after the input tabs + source prefixes are filled, open each matched source file, sample it, and write a **suggested** type-map for review (never applied silently; ID/code columns flagged to stay `T`). |

### Package `clean` — "Clean & Standardize" (v1.0.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `clean_keycols` | text | simple | `1` | Which column(s) decide duplicates, numbered `1,2,3…` from the LEFT of the picked range. One number or several (`1,2,4`). |

`clean` configures no target range — each action prompts for its range at run
time (see §3, `range` note).

**Actions**

| id | caption | semantics |
|---|---|---|
| `clean.dupes` | Remove Duplicates | Prompt for a range (incl. header row); remove duplicate rows using `clean_keycols`. |
| `clean.trim` | Trim & Clean Text | Prompt for a range; strip leading/trailing spaces, collapse double spaces, remove non-printing characters. |
| `clean.filldown` | Fill Blanks Down | Prompt for a range; fill each blank cell with the value from the cell above it. |

---

## 9. Platform-abstraction appendix

The Excel-isms in the reference implementation and their neutral meaning, so a
port realizes them per platform rather than re-deriving them:

- **Data location** — "the workbook's containing folder" (the folder holding the
  `.xlsm`) is where source files and the master file are matched and where
  exports are written. Neutral: *the deployment's data location*. Excel → local
  folder next to the workbook; Sheets → a Drive folder (the container's parent or
  a configured folder id).
- **File matching (`file` kind)** — select the **newest** file whose name
  **starts with** the given prefix, ignoring a trailing ` (1)`/` (2)` copy
  suffix (so re-downloads need no renaming).
- **Export snapshot (`importexport.export`)** — a values-frozen copy of the
  listed tabs to the platform's spreadsheet format, keeping formatting / filters
  / conditional formatting and breaking external links; filename is
  `<date+offset> <base>`.
- **Dated filename** — `<run date shifted by k_export_offset> <k_export_base>`
  plus the platform's spreadsheet extension.
- **Speed/lock wrapper** — each action wraps its work in the platform's
  screen-update/recalculation/event suppression, restored on exit and on error.
  Reference: the handler calls the shared `Core_SpeedOn` / `Core_SpeedOff`, and
  the dispatcher's error path also restores as a fail-safe. (Excel:
  ScreenUpdating/Calculation/EnableEvents; other platforms use their nearest
  equivalent or omit if not applicable.)

---

## 10. Change control

- This contract is versioned with SemVer, independent of the repo and package
  versions.
- **Additive** changes — a new kind, field, action, or package; a new optional
  attribute — bump the **minor** version.
- **Breaking** changes — renaming or removing a `key` or action `id`; changing a
  field's kind, a value's encoding, or an action's semantics — bump the
  **major** version.
- The contract is the source of truth going forward: the reference VBA build and
  every port MUST conform. A discovered divergence is reconciled **in favor of
  this document**.
- **Reserved for first use:** the `bool` kind (§3) — its encoding is fixed, and
  this document updated, when the first package uses it.
- **v1.1.0 (additive):** adds the `chain` config entity, the reserved `chain.*`
  id namespace, engine-owned config, the grouping/layout placement code, and the
  quiet/dispatch convention. No existing key, action id, kind, value encoding, or
  semantics changed — v1.0.0 deployments remain valid.
- **v1.2.0 (additive):** adds the optional action `category` attribute, the
  optional field `usedBy` attribute, the three-state deployment-level
  `availability` flag, the capacity-packed auto-layout rule for uncoded buttons,
  the always-generated index surface with its navigation and action catalog, the
  no-clipping rule for descriptions, the config schema marker with fail-closed
  migration, and the reserved `Chains` category + `Macros Home` name. No `key`,
  action `id`, kind, value encoding, or **action** semantics changed, so
  v1.0.0/v1.1.0 configs remain valid and read as all-actions-`button`. The one
  *behaviour* replaced is engine layout for the uncoded case (single page →
  capacity-packed pages), which §10 does not classify as breaking: it renames
  nothing and re-encodes nothing. Owner-approved; see the Provenance section of
  `docs/designs/2026-07-24-finish-excel-track1-plan.md` for the three contested
  calls and who lost them.
