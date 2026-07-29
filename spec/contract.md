# SheetForge portable contract

**Contract version: 2.9.0** (2026-07-28) · SemVer, independent of the
repo `VERSION` and of per-package versions. Full history in §10.

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
- **Precedence over tier (v1.2.0).** `usedBy`-relegation and `tier` (§4) are
  evaluated independently, but off-ness wins: a field relegated to the off band
  goes there regardless of its `tier`, and the `advanced` band (§4) contains
  only fields that are still in use. A field never renders in both bands.
- **No visible fields, no band (v2.0.0).** A package contributing no
  field to the working area MUST NOT have a configuration band rendered for it —
  neither its banner nor an empty body. This holds whether the package declares
  no fields at all or has had every one of them relegated to the off band. A
  heading over nothing reads as a defect, and the engine already applies the same
  judgement to an empty `advanced` band and to a section header that merely
  repeats its package banner. The package is still registered and its actions
  still appear in the action catalog; it simply has no configuration surface.

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
| `bool` | yes/no | `yes` / `no` (canonical tokens; a reader MUST be case-insensitive) | validated Yes/No dropdown; a **blank** cell means "never set" and reads as the field's declared default | Encoding fixed at v1.2.0 by `clean_hasheader`, the kind's first use — **since withdrawn** (§8), which does NOT reopen the reservation: an encoding is fixed once and stays fixed, or a later field would silently mean something different from what shipped. The kind is currently declared and unused. Like `availability` (§6), the tokens are abstract, not a mandated on-screen spelling: the Excel reference implementation presents and stores `Yes` / `No`. A port MAY accept only `yes` / `no`; the reference additionally tolerates `y` / `true` / `1` and `n` / `false` / `0`. An unrecognised non-blank value MUST block Finalize, naming the field — it MUST NOT fall back to either state. |
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
- **Disclosure (v2.1.0).** A field whose actions are all `off` is relegated, and a
  platform MAY present relegation as a collapsed band or as hidden-in-place. **Omitting
  the field entirely remains non-conforming** either way: the engine re-seeds its
  authoring surface from stored config, so an omitted field costs the user a
  toggle → activate → rebuild → fill → activate cycle and discards anything typed but
  not yet activated.

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

**Interaction with `usedBy` (v1.2.0):** a field relegated to the off band (§2)
is placed there regardless of its `tier` — off-ness takes precedence, so this
`advanced` band contains only fields that are still in use.

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

**Durable replacement (v2.2.0, normative):** activating MUST NOT destroy the
existing store before its replacement exists. An implementation MUST **stage the
replacement, verify it, and only then swap it into place** — write the new store
under a reserved staging name (§7), confirm it is structurally complete (reference:
its action-catalog key row is present), and only after that confirmation remove the
old store and promote the staged one. Delete-then-write is non-conforming: the store
is the only memory that survives a rebuild, there is no undo, and an interruption in
that window leaves nothing — which the next build reads as "never configured" and
overwrites with defaults, turning a crash into silent total loss.

This is **not** a claim of atomicity, and an implementation MUST NOT present it as
one; a host that offers no transaction cannot give one. The requirement is that the
irreducible window contain a *complete replacement* rather than *nothing*. It follows
that a reader MUST adopt an orphaned staged store when the live one is missing —
that adoption is what makes the surviving state recoverable rather than merely
present, and it is unambiguous because a staged store never outlives a successful
activate. Ordering *within* the store (§5's marker-last rule) is a separate concern
and is unaffected.

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

**Escape hatch when there is no usable surface (v1.2.0).** The abort above
assumes the engine's authoring surface (the Setup grid) exists **and** exposes
the offending value in a form the user can correct by hand. If either is
untrue — the surface does not exist, or it exists but the value in conflict is
not a visible, user-editable cell (for example, a tampered or unparseable
machine-owned marker) — so a conflict would otherwise dead-end the workbook
with no path forward, an implementation MAY proceed with the rebuild anyway —
but only after an explicit, non-default confirmation that names exactly what
will be discarded and what is preserved, and only on an affirmative response
to it. Proceeding without that confirmation remains non-conforming.

**How a reader decides the marker is absent (v2.2.0, normative).** "Absent" is a
claim about the store's *era*, and it is the permissive branch — it routes to the
pinned legacy defaults below. A reader MUST NOT conclude "absent" merely because the
marker cell is empty, and MUST NOT decide it from a **user-editable display caption**
such as a column header. It MUST require positive corroborating evidence that the
store predates the marker — structural evidence the newer format cannot produce, such
as the complete absence of a column that the newer format always writes. Failing that
test, an empty marker is a **removed** marker: fail closed and refuse to activate.
The marker is machine-owned precisely so this decision is not left to something the
user can retype; a reader that leans on a caption re-opens the hole the marker exists
to close. *(The reference implementation previously classified a marker-less block by
reading its visible `Availability` header back, so renaming that header by hand
promoted a tampered current block to the legacy path and activated every action as
`button`.)*

**Absent is not invalid.** An attribute or column introduced after the stored
marker's version (for example `availability`, §6, in a store marked pre-v1.2.0)
simply does not exist in that store — there is no value to validate, so its
absence is not a conflict, and the reader takes the **pinned legacy default**, which
for `availability` is `button` — fixed at this value as of v2.1.0 and deliberately
NOT a reference to §6's current default. §6's default moved to `off` in v2.1.0; this
one did not move with it, because a store written before the column existed described
a deployment in which every action was live, and reading it as `off` would silently
remove every button from a working installation. A reader MUST treat these as two
independent values. Fail-closed applies only to a value that IS present but
unreadable or fails its grammar, and to a marker the reader does not recognise at
all — never to a column this document added after the store was last written.
This is what lets §10 promise that v1.0.0/v1.1.0 configs remain valid.

---

## 6. Action contract

An action is one thing a package *does* (it becomes a button). Each action is:

| Field | Meaning |
|---|---|
| `id` | Dotted `<domain>.<operation>` token, globally unique (`import.combine_worksheets`, `clean.dupes`). The domain is a controlled user-facing vocabulary, **not** the package that implements it — see §7 for the vocabulary table and the full rule. |
| `caption` | Button label (may carry an ordinal, e.g. `"2.  Import Source Data"`). |
| `help` | Plain-language description shown on/near the button. |
| `category` | Optional display category (v1.2.0). Defaults to the package `name`. Groups actions on an index surface where a platform provides one (§6), and heads the columns of a button page; deliberately decoupled from the package, so two packages may feed one category and one package may span several. |
| `availability` | Deployment config, not manifest data (v1.2.0): one of `button`, `chains-only`, `off`. **Defaults to `off` (v2.1.0)** — an action arrives switched off and is turned on deliberately. This changes no stored value's meaning: a stored `button` still reads as `button`. The default applies only where nothing was ever stored. For the pre-v1.2.0 legacy case the default is pinned separately and does NOT follow this row — see §5. |
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

**Three outcomes, not two (v2.4.0).** An action invocation ends in exactly one of:

| outcome | meaning | what the caller sees |
|---|---|---|
| **done** | The action did its work. | A completion message. |
| **refused** | A precondition was not met, so the action deliberately did nothing. The destination already exists; the folder is missing; nothing is configured yet. | A plain-language reason. **No error number** — there is no error. |
| **failed** | Something went wrong while the action was working. | An error number and description. |

**A refusal is not an error, and the difference is load-bearing.** A refusal is the action
working correctly: it checked, found the world not ready, and stopped having changed
nothing. An implementation MUST be able to tell the two apart, because the chain rule below
treats them differently, and because a refusal reported with a fabricated error number reads
to the user as a malfunction.

**Normative:** an implementation MUST expose a single call that records a refusal and
presents its reason, and that call MUST suppress its own presentation when the invocation is
quiet (a chain step), exactly as the completion sink does. Two separate steps — record the
outcome, then show a dialog — are non-conforming: an action that does one and forgets the
other reports a success it did not achieve, which is the defect this rule exists to prevent.

**Chains stop at a refusal (v2.4.0, normative).** When a chain step **refuses**, the chain
MUST stop at that step and MUST NOT run the remaining steps. It MUST report which step
refused, its reason, and how many steps did not run.

When a chain step **fails**, the prior behavior is unchanged: the implementation asks
whether to continue.

The asymmetry is deliberate. A failure may be incidental to one step, and the user is placed
to judge. A refusal means the step's *output does not exist* — so every later step would run
against whatever was there before, silently producing a plausible, wrong result. The
reference implementation's own combine action refuses on every run after the first (its
destination sheet now exists), so under the old ask-and-continue rule a daily chain cleaned
last period's data and reported success.

**Grouping / layout (v1.1.0):** each button (action or chain) may carry a
**placement code** `<page-number><column-letters>` (e.g. `1A` = page 1, column A;
`12AB` = page 12, column "AB") — bounded to **at most seven digits and three
letters** (v1.2.0, normative; matches Excel's own column-letter grammar, whose
last real column is `XFD`) — or the explicit code **`hide`**, which omits a
button and — normative, v1.2.0 — is **not** a placement: it does not count
toward "any code is present" below in either direction. The reference
implementation heads this column **Placement** on its authoring surface
(renamed from `Page code` at v1.2.0, when the value stopped being only a page
code). The column caption is presentation, not data: no conforming reader may
key off it. Semantics: if NO
button anywhere has a **placement** code (no codes at all, or `hide` codes
only, are both "none"), a single default button page is built (legacy
behavior; **superseded for the uncoded case by Auto-layout (v1.2.0) below** —
a store with no placement codes now packs into capacity-bound pages, not one
page; see that paragraph). If any **placement** code is present, a layout is
active: numbered pages are built; a blank-coded button goes to page 1 (the
default page)'s **Unassigned** area. `hide` omits a button the same way in
either mode. Same-code buttons stack in authoring order. Core's own admin
controls are never user-placeable.

**Anchored placement (v1.2.0).** A placement value may instead take the form
`<sheet>!<cell>` (e.g. `MIC!B4`), which anchors that one button to a single cell on a
surface the user owns, rather than filing it onto a generated page. The
discriminator is `!`. The two forms can never be confused — not because the platform
forbids `!` in a surface name (Excel, for one, does not) but because a valid **page
code** cannot contain `!` under the grammar above, so even a surface legitimately
named `1A` is unambiguous. An implementation MUST split on the **last** `!`, so a
surface whose own name contains one still resolves. The sheet part may
be quoted (`'MIC Data'!B4`) or bare even when it contains spaces (`MIC Data!B4`); an
embedded apostrophe is doubled inside the quotes. The cell part is a single A1
reference with optional `$` (`MIC!$B$4`); a range (a colon) is rejected.

**Normative, authoring-surface quirk (v1.2.0).** Where the authoring surface itself
consumes a leading apostrophe — Excel treats one as a "this cell is text" prefix and
strips it from the stored value, so a user who types `'MIC Data'!B4` into a cell
stores `MIC Data'!B4` — an implementation MUST accept the resulting
opening-quote-less form and treat the orphaned closing quote as the eaten pair. Left
unhandled, the quoted spelling this clause lists **first** is unreachable through
that platform's own UI, and the user is told the surface does not exist while looking
straight at it. A platform whose cells do not eat the character has nothing to do
here. The cost is that a surface whose real name *ends* in an apostrophe cannot be
addressed by its bare name; that is accepted, and it fails loudly rather than
resolving to the wrong surface.

- **Normative: an anchor is NOT a placement code for the mode decision.** It does not
  count toward "any code is present" in either direction — the same carve-out `hide`
  already has. A store whose only non-blank values are anchors is a store with **no**
  placement codes, and auto-layout applies to every other button unchanged. Stated in
  both directions because reading this clause literally in the other direction is how
  an implementation dumps every uncoded button into Unassigned the moment one button
  is anchored.
- **Normative: the engine MUST NOT modify the target surface** beyond adding and
  removing its own control. No clearing, no resizing, no writing a value, no
  formatting, no row or column sizing. The target is a surface the user owns; the
  engine is a guest on it.
- **Normative: an implementation MUST remove its previously-placed controls on
  rebuild**, across the whole deployment, before placing from the current values — so
  a moved cell, a cleared value, a renamed sheet, or a build that died halfway cannot
  leave a live orphan behind.
  - **Normative: that removal obligation is best-effort where the surface refuses
    control deletion.** A surface may be locked against control changes by the user
    (in Excel, a sheet protected with `DrawingObjects`, which is Protect's default);
    removal there fails and a previously-placed control survives. The engine **MUST
    NOT** alter the surface's protection state — unprotecting, then reprotecting, a
    surface the user chose to lock is a larger violation than a stale control, and
    the preceding "MUST NOT modify the target surface" rule outranks this one. The
    engine **MUST** instead report the failure to the user, with the count and the
    repair (unprotect, then rebuild). A surviving control on a locked surface is
    **conforming behaviour**, not a divergence to be reconciled under §10 — do not
    "fix" it by unprotecting.
- **Blocking conditions.** Finalize MUST refuse to activate, naming the button, when:
  the named surface does not exist; the named surface is one the engine itself
  generates and rebuilds; two buttons resolve to the same target cell (two controls at
  identical coordinates means one covers the other and is unreachable); the value is
  malformed; or the cell does not resolve on the named surface.
- Both the action catalog and the chains block use this same grammar, so a **chain**
  may be anchored exactly like an action.

**Column-cap overflow (v1.2.0).** A rendered page draws at most the same three
columns Auto-layout is capped at below, and a column draws at most the same six
buttons a same-code stack is capped at there too — regardless of whether that
column was reached by an explicit placement code or by the blank-coded
Unassigned fallback above, and regardless of which page number a code names.
This is a capacity fact about the renderer, not a grouping judgement like Auto-
layout's no-straddle rule: the reference renderer has exactly three column
slots per page and six button rows per column, and a fourth column or a
seventh same-code button is something it cannot draw, not something it
chooses not to draw. When either cap would be exceeded, the excess moves to a
continuation page instead of going undrawn: a `1D` typed on a page that already
holds coded columns `1A`/`1B`/`1C` renders on the next available page, not the
page its code names; an Unassigned column that would be a page's fourth does
the same; and a same-code stack past six buttons continues in a further
column, on a continuation page if the page it started on is already at the
column cap. This is engine presentation under §10's two-axis rule (additive:
no stored `key`, id, kind, encoding, or semantics changes), so it is folded
into this v1.2.0 freeze rather than bumping the contract version.

**Per-action availability (v1.2.0):** each action carries a deployment-level
`availability`, authored on the same engine-owned surface as the placement code
(§5). It has exactly three values. It defaulted to `button` at v1.2.0, so a
deployment that predates v1.2.0 was unchanged by its introduction; **as of v2.1.0
the default is `off`**, and the pre-v1.2.0 case keeps `button` through §5's pinned
legacy default rather than through this one.

| Value | Button drawn | Runnable directly | Runnable as a chain step |
|---|---|---|---|
| `button` | yes | yes | yes |
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
  otherwise be handed a live destructive button.
- **A blank cell MUST block Finalize too (v2.1.0).** It previously read as `button`,
  which was coherent while `button` was the default. With the default at `off`
  (§6) neither reading is safe: `button` resurrects an action the user switched
  off, `off` removes one they still want, and the choice is invisible in the
  result. A blank cell is therefore not an answer, and a reader MUST NOT resolve it
  to any token. **This is the cell being blank, not the column being absent** — an
  absent column is §5's legacy case and still reads as `button`.
- **Portability note:** the reference implementation presents the three values as a
  validated dropdown on the authoring grid. A platform without cell validation
  still parses the same three tokens and still blocks on anything else.
- **`button` / `chains-only` / `off` are abstract tokens, not a mandated on-screen
  spelling.** A platform's authoring surface presents them as human-readable
  labels; the Excel reference implementation writes and offers `Button` /
  `Chains only` / `Off` (space, not hyphen, and title case) as its dropdown and
  stored values. A conforming reader MUST accept either spelling of the middle
  state (`chains-only` and `chains only`, case-insensitive) as equivalent.
- **Reference-implementation tolerance (v1.2.0).** The three canonical tokens
  above remain the only NORMATIVE values — a port MAY be strict and accept only
  them. The Excel reference implementation additionally accepts, purely as a
  convenience for a user typing the obvious thing: `yes` / `y` / `on` / `true`
  as `button`; `chain only` / `chains` / `chain` / `chains-only` as `chains-only`;
  and `no` / `n` / `false` as `off`. A port MAY implement this tolerance but is not
  required to; it does not change what gets stored (still one of the three
  canonical tokens) or the blocking behavior for anything outside this list.

**Auto-layout (v1.2.0) — supersedes the "no codes" clause of the v1.1.0 grouping
rule above.** When **no** button carries a placement code, the engine distributes
buttons automatically. A page is a **capacity unit, not a category unit**:

- a page holds up to **three category-headed columns**; a column holds up to
  **six** buttons, all of one category (sized for one screen at typical
  display scaling — see the portability note below for the arithmetic this
  answers);
- small categories therefore **share** a page;
- a category needing more than one column takes consecutive columns, and a
  category that does not fit in a page's remaining columns starts a new page
  rather than straddling one — splitting a category across tabs destroys the
  grouping that putting it in one place buys;
- a category larger than a whole page continues onto further pages.

**Ordering (v1.2.0, normative):** categories are packed in **first-appearance
order** — the order in which each category's first action is encountered while
the engine walks its package registry (§1), and, within one package's manifest,
that manifest's `actions[]` declaration order (§6). The reserved `Chains`
category (§7) is always placed last, regardless of when its chains were
authored. Two conforming implementations given identical config MUST therefore
produce identical layouts.

The v1.1.0 behaviour — a single default button page holding everything — is
withdrawn: it does not survive contact with a large catalog, which is what this
version exists to support. **When any placement code is present the v1.1.0 rule
is unchanged** (`hide` does not count — see above) and codes remain a complete
override of **which column** a button is filed into, and the order same-code
buttons stack within it. They are not, per the Column-cap overflow clause
above (v1.2.0), a guarantee of which page or column slot that produces: a
same-code stack past six buttons continues into a further column beyond the
one its code names, and that further column lands on a continuation page only
once its own page already holds three columns — otherwise it stays on the same
page its code named, alongside the column it overflowed from.

**Portability note:** the *packing rule* is normative — fill a column, then a
column-slot on the page, then start a new page; never straddle a category
across pages unless it exceeds a whole page — and so is the ordering above. The
specific capacities (three columns, six buttons) are the reference
implementation's values, sized for one screen of Excel at typical display
scaling (eight button-pairs only fit a full-HD, 100%-scaled, maximised
window; 125% scaling — common — fits about 5.7); a platform with different
real estate MAY choose its own capacities but MUST apply the same rule and
the same ordering.

**Index surface (v1.2.0, amended v1.2.0 pre-freeze):** a platform **MAY** generate
an index surface. It is **not required**: the Excel reference implementation built one
and then withdrew it before release, because at the scale a real deployment runs —
a handful of actions and one or two chains — an index is clutter rather than
navigation. **When a platform does provide one**, it lists one entry per built page
with its title, its button count and a link; a **button catalog** (caption,
description, and the page each button is on) covering every action whose
`availability` placed it as a button — an action that is `off` or `chains-only` has
no button and no page, so it does not appear in this catalog; it remains listed on
the authoring surface (§6, per-action availability), which is where it stays
discoverable; a link to the admin surface; and a legend explaining the placement
code and availability.

**Page-to-page navigation is required regardless, and is unconditional:** every
built page MUST carry links to its previous and next page. Only the index itself,
and the per-page link back to it, are optional.

The reference implementation's withdrawn index was a `Macros Home` sheet. That name
stays **reserved** (§7) against a future restore.

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
- **Action id** — `<domain>.<operation>`, unique, exactly two segments. Each segment
  **begins with a lowercase ASCII letter** and continues with lowercase ASCII letters,
  digits and underscores. (The leading-letter rule is stated because the build gate
  enforces it; without it here, the gate would be stricter than the contract it claims to
  implement.) **The domain is a controlled user-facing vocabulary, NOT the package or
  module that implements it** — an action id is a permanent dispatch key stored in the
  config (§5), so encoding the module in it would make a future package split a major
  contract break for no benefit. The operation segment is
  descriptive where ambiguity looms (`clean.remove_blank_rows`, never `clean.blank2`).
  **No new id is ever camelCase.** Ids are immutable; captions, help text, entry points
  and module assignment stay free to change. There is no alias or migration table, by
  design: an alias means two permanent names for one action and would have to be ported to
  Apps Script.

  **The vocabulary is the table below, and it is normative.** The marker comment appears on
  its own line above it — a blank line may separate the two — and is a parse anchor:
  `build/build.ps1` reads the table from this file
  and fails a build whose action ids use a domain that is not listed `active` or
  `grandfathered`. Do not reformat the table, rename its columns, or remove the marker
  without changing that parser in the same commit.

  **Machine-readable constraints, because a build gate reads this table.** The marker
  appears **exactly once** in this document. The header row is exactly
  `| domain | status | meaning |` and is followed by exactly `|---|---|---|`. Every row
  between the header and the terminating blank line is a data row of the form
  `` | `<domain>` | <status> | <meaning> | `` where `<status>` is one of `active`,
  `reserved`, `grandfathered`. **No cell may contain a `|` character** — there is no
  escaping rule and none is wanted; a pipe would corrupt the rendered table as well as the
  parse. A domain appears at most once. The parser **fails the build** on any violation
  rather than skipping the offending row: a gate that quietly parses eleven of twelve rows
  still goes green while silently permitting the twelfth domain.

<!-- action-domain-vocabulary -->

| domain | status | meaning |
|---|---|---|
| `import` | active | Bringing outside data into the workbook: files, folders, other sheets. |
| `clean` | active | Repairing data already present: blanks, duplicates, merges, text that should be numbers. |
| `validate` | active | Checking data against a rule and reporting what fails; never repairs silently. |
| `enrich` | active | Adding derived or looked-up columns from a reference table. |
| `format` | active | Presentation only: number formats, styles, print layout. Changes no value. |
| `summarize` | active | Aggregating: pivots, recalculation, summary sheets, charts. |
| `export` | active | Writing data out of the workbook: files, PDFs, values-only snapshots. |
| `workbook` | active | Housekeeping on the workbook itself: sheet visibility, protection, ordering, navigation. |
| `backup` | active | Restore points: timestamped copies, pre-operation snapshots, archives. |
| `chain` | reserved | Engine chains (section 6). No package action id may use it. |
| `core` | reserved | Engine-internal sentinels (v2.1.0, this section). No package action id may use it. |
| `importexport` | grandfathered | The pre-v1.2.0 package-named domain. Its five shipped ids are permanent; no new id may be added to it. It is direct evidence for the domain-is-not-the-module rule above. |

**The eight shipped ids are permanent exceptions**, listed literally, never renamed and
never aliased: `clean.dupes`, `clean.trim`, `clean.filldown`, `importexport.clear`,
`importexport.import`, `importexport.updateMaster`, `importexport.export`,
`importexport.suggestTypes`. Two are camelCase and the `importexport` domain conflates two
different jobs; both stay exactly as they are, and the rules above exist because of them.

**Enforcement, in two places.** The engine rejects a duplicate id during registry
assembly, which protects a hand-edited bundle. The build gate extracts literal action-id
declarations from source and fails on a duplicate, a malformed id, a reserved domain, a new
id in a grandfathered domain, an unknown domain, or an entry point no source declares as
dispatchable. A platform without a build step performs the equivalent checks at
manifest-assembly time.

- **Reserved:** the id namespace `chain.*` is reserved for engine chains; no package action id may use it.
- **Reserved (v2.1.0):** the action-id namespace `core.*` is reserved for engine-internal
  sentinels; no package action id may use it. The reference implementation uses
  `core.legacy` as a provenance flag inside the in-memory availability map, so an id
  colliding with it would be read as engine state rather than as an action.
  **Enforced (v2.2.0, normative):** an implementation MUST refuse to activate a
  configuration in which any package action id falls in `core.*`, naming the offending
  id, exactly as it already refuses `chain.*`. Declaring the reservation is not enough:
  this section states no character rule for action ids, so a sentinel living in the same
  keyspace as real ids cannot be assumed collision-proof, and a collision on
  `core.legacy` would resolve every unlisted action to `button` instead of `off` —
  silently reverting §6's v2.1.0 default on that deployment, destructive actions
  included.
- **Reserved (sheet names, v2.2.0):** in the Excel reference implementation the sheet
  names `_Config` (the activated store) and `_Config_staged` (where Finalize stages a
  replacement before swapping it into place) are engine-owned. A user sheet taking
  either name would be read as engine state — `_Config_staged` in particular is
  **adopted as the configuration** when `_Config` is missing, which is how an
  interrupted activate heals (§5). A platform substitutes its own two names but MUST
  reserve both: staged-then-swap needs somewhere to stage.
- **Reserved:** the display category `Chains` (engine-generated, holds chain
  buttons) and, in the Excel reference implementation, the sheet name
  `Macros Home`. A platform substitutes its own index surface name. The name
  stays reserved even though the reference implementation withdrew its index
  before release (§6) — and as of that withdrawal the engine **deletes
  `Macros Home` unconditionally with alerts suppressed on every rebuild**, the
  same treatment `Macros` and `Macros <n>` get below, so that a stale index left
  by an earlier build cannot survive: a package or a user relying on that name
  for anything else loses it silently.
- **Reserved (v1.2.0):** in the Excel reference implementation, the sheet name
  `Macros` (the pre-v0.4.0 single default page, retired) and every `Macros <n>`
  (the generated page tabs) — the engine deletes these unconditionally with
  alerts suppressed on every rebuild, so a package or a user relying on one of
  these names for anything else loses it silently. A platform reserves the
  equivalent generated-page names on its own surface.
- **Reserved (v1.2.0):** in the Excel reference implementation, the shape-name prefix
  `SFPlaced_`. Every control the engine anchors onto a user's own surface carries it,
  and the engine removes controls by that prefix alone on every rebuild — so anything
  else in the workbook using the prefix is deleted without warning. A platform
  reserves its own equivalent namespace for the same purpose.
- **Reserved (v2.9.0, normative):** an implementation creating a NEW destination sheet from a
  `text`-kind config field MUST refuse a value naming a sheet in this section's reserved family,
  or a sheet the platform's own generated-page naming rule claims (the Excel reference
  implementation: `Core_IsEngineSheet`), and MUST NOT check only for an EXISTING sheet of that
  name. Discovered by `workbook.build_toc` (§8): a value like `"Macros 3"` names no sheet today,
  so an exists-only check lets it through, and the sheet it creates is destroyed unconditionally
  by the very next rebuild (the `Macros`/`Macros <n>` rule above) - a data-loss trap that would
  not surface until the rebuild AFTER the one that hid it. This closes the same gap for every
  existing `text`-kind destination field (`k_import_dest` and any added by W2-W5), not only this
  wave's field, because the underlying check is one shared engine function in the reference
  implementation.
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

### Package `import` — "Bring Data In" (v1.0.0, contract v2.3.0)

A second package in the `import` domain. `importexport.import` keeps its id and its
module; both packages' actions declare the display category `Import / Export`, so the user
sees one coherent card while the modules stay whatever is convenient to maintain (§7's
domain-is-not-the-module rule, realized).

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_import_dest` | text | simple | `Master` | Name of the **new** sheet a combine action creates. `usedBy` the two combines. A conforming implementation MUST create this sheet and MUST NOT overwrite, clear or delete an existing sheet of that name — if one exists the action MUST stop before changing anything and say so, naming the sheet. Kind is `text`, deliberately **not** `sheet`: a `sheet` field is validated at activation to **exist**, and this one must not. |
| `k_import_folder` | text | simple | *(blank)* | Folder the workbook combine reads from. **Blank means the deployment's data location** (§9 — the workbook's containing folder in Excel, a Drive folder in Sheets), which is the common case and the reason the default is blank rather than a path. `usedBy` `import.combine_workbooks`. |
| `k_import_delim` | text | advanced | `comma` | Column separator for delimited files. Token vocabulary: `comma`, `tab`, `semicolon`, `pipe`, `space`; any other non-blank value is read as **the literal separator character** (its first character). Blank reads as `comma`. `usedBy` `import.delimited_file`. |
| `k_import_qualifier` | text | advanced | `"` | The character wrapping text that itself contains the separator. `"` (double quote), `'` (single quote), or the token `none` to disable qualification entirely. **A blank cell reads as the declared default `"`, not as `none`** — §5's blank-means-never-set rule resolves a blank to the manifest default, so `none` must be typed to be meant. Any other value also reads as `"`. `usedBy` `import.delimited_file`. |

**Actions**

| id | caption | semantics |
|---|---|---|
| `import.combine_worksheets` | Combine All Worksheets | Stack every data sheet in this workbook into one new sheet named `k_import_dest`, **aligning columns by header text, not by position**: a source's `Amount` column lands under the master `Amount` wherever it sits, and a header no earlier source declared is appended as a new column on the right. Row 1 of each source is its header and is not repeated. Two columns sharing a header *within one source* MUST NOT overwrite each other. Engine-owned sheets (the Setup surface, the config store and its staging sheet, the package-enable sheet, and every generated button page) MUST be excluded — and only those: a user sheet whose name merely resembles a generated one MUST still be combined. If `k_import_dest` already names an existing sheet the action MUST stop having changed nothing. A failure part-way MUST leave no partially built destination sheet behind. |
| `import.combine_workbooks` | Combine All Workbooks in a Folder | As `import.combine_worksheets`, but the sources are the first worksheet of every workbook in `k_import_folder`, excluding this workbook itself. Each source is opened read-only and closed without saving, one at a time. Same header alignment, same refusal, same partial-failure cleanup. A missing folder, and a blank folder setting in a workbook that has never been saved, MUST each be reported as themselves rather than surfacing as an engine error. |
| `import.delimited_file` | Import a Delimited File | Prompt for one delimited file; parse it with the explicit separator and text qualifier from config rather than letting the platform guess; write it to a new sheet named after the file. Same refusal rule as the combines if that sheet name is taken. The prompt MUST offer a cancel that abandons the action having changed nothing. A failure after the file is opened MUST leave neither the opened source nor a partial destination sheet behind. |

### Package `clean` — "Clean & Standardize" (v1.0.0)

**Fields**

**One field (v2.3.0).** `clean` declared no configuration fields at all from v2.0.0 to
v2.2.0. It now declares exactly one:

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_clean_formats` | table | simple | **none** | Column-to-number-format map for `clean.standardize_formats`. Two columns: **Column** (a column letter, or a header name matched against the first row of the range the caller picks) and **Number format** (a platform number-format string). A row whose Column resolves to nothing inside the picked range MUST be reported, not silently skipped. `usedBy` `clean.standardize_formats`. |

`k_clean_formats` header: `Column (letter or header name)|Number format`.

**This field declares NO default rows, and that is normative, not incidental.** §5's
fallback resolves an empty stored table to the field's declared defaults — so a field that
declared demo rows could never observe "the user configured nothing", and
`clean.standardize_formats`' required refusal below would be unreachable. It is also the
difference between an action that does nothing until told and one that silently applies a
demo date format to whatever range a first-time user picked. A conforming implementation
MUST NOT ship default rows for this field.

`k_import`'s demo rows (§8, package `importexport`) are the opposite case and stay: a
three-column type map is genuinely hard to guess from an empty grid, and its action already
refuses to run against tabs that do not exist, so a stale demo row cannot silently act on
real data. Worked examples for `k_clean_formats` live in its help text instead, where they
teach without becoming stored configuration.

**This does not reopen the rule below, it satisfies it.** A house number format is a fact
about the deployment — the same columns get the same formats every close — not a fact
about the range the caller is about to pick, which is precisely the distinction the
withdrawn fields failed. Resolution by **column letter or header name**, never by position
within the selection, is what keeps it on the deployment side of that line: a position
counted from the left of a selection that does not exist yet would be the withdrawn
`clean_keycols` under a new name.

It previously declared `clean_keycols` and `clean_hasheader`. Both were withdrawn
because each describes **the range the caller selects**, not the
deployment: the key columns are numbered from the left of a selection that does
not exist until the action runs, and whether a first row is a header is a
property of that selection. A stored answer forced a re-activation cycle between
two runs of the same action, and was not discoverable at the moment of use. They
are now **run-time inputs** of `clean.dupes` (see Actions below).

**Normative rule this establishes.** A setting that describes the selection the
caller is about to make MUST be a run-time input, not a config field. A setting
that describes the deployment MUST be a config field. Run-time inputs MUST NOT be
persisted between invocations: asking each time is what keeps the answer matched
to the selection.

`clean` configures no target range — each action prompts for its range at run
time (see §3, `range` note).

**Actions**

| id | caption | semantics |
|---|---|---|
| `clean.dupes` | Remove Duplicates | Prompt for a range; then prompt for **(a)** whether that range's first row is a header and **(b)** which column(s), numbered from the LEFT of that range, decide a duplicate. Remove duplicate rows accordingly; a header row is excluded from comparison and never removed. Each of the three prompts MUST offer a cancel that abandons the action having changed nothing, and MUST be distinguishable from an empty answer. A column number outside the selected range MUST be rejected before any edit, naming the range's width. The completion report MUST state which header interpretation was applied, since it is not recoverable from the result. |
| `clean.trim` | Trim & Clean Text | Prompt for a range; strip leading/trailing spaces, collapse double spaces, remove non-printing characters. |
| `clean.filldown` | Fill Blanks Down | Prompt for a range; fill each blank cell with the value from the cell above it. |
| `clean.remove_blanks` | Remove Blank Rows & Columns | Prompt for a range; find every entirely blank row and column inside it. **Delete only those that are also blank across the whole sheet's used extent**, so deleting a whole row can never destroy data sitting outside the caller's selection; report the ones held back and why. Before deleting anything, show a count of exactly what will go and require an explicit confirmation; declining or cancelling MUST change nothing. The completion report MUST state both counts, because neither is recoverable from the result. |
| `clean.unmerge_fill` | Unmerge & Fill Down | Prompt for a range; unmerge every merged block **wholly inside** it and write the block's value into every cell it covered. A block crossing the selection's edge MUST be left alone and reported — unmerging it would change cells the caller did not select. A merge anchor holding a formula MUST keep its formula; the cells it covered receive the computed value. |
| `clean.standardize_formats` | Standardize Formats | Prompt for a range; apply each `k_clean_formats` row's number format to that row's column **within the picked range only**. Changes formatting, never a stored value. A mapping row that resolves to no column in the picked range MUST be reported by name, not silently dropped. With no rows configured, the action MUST stop and say where to configure them (see the field's no-defaults rule above, which is what makes this state observable). |
| `clean.text_to_numbers` | Convert Text to Numbers | Prompt for a range; convert text that represents a number into a real number, using **the same coercion the import action uses** (§8, column-type code `N`) rather than a second implementation. Formula cells are never touched. Text the coercion cannot parse is left exactly as it is. The completion report MUST state how many cells changed. **Known limitation, stated because it is visible to the user:** the `N` coercion strips a trailing `%`, so `12%` becomes the number `12`, not `0.12`. |

### Package `validate` — "Validate & Reconcile" (v1.0.0, contract v2.5.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_validate_recon_dest` | text | simple | `Reconcile Exceptions` | Name of the **new** sheet `validate.reconcile` writes its exceptions report to. Created only when the run finds at least one exception; a perfect match creates no sheet. `usedBy` `validate.reconcile`. |
| `k_validate_scan_dest` | text | simple | `Scan Exceptions` | Name of the **new** sheet `validate.scan` writes its exceptions report to. Created only when a configured rule finds at least one failing cell. `usedBy` `validate.scan`. |
| `k_validate_scan_rules` | table | simple | **none** | Data-quality rules for `validate.scan`. Three columns: **Column** (a column letter, or a header name matched against the first row of the range the caller picks — a one-to-three-letter entry is always read as a letter), **Rule** (`blank`, `number`, `date`, or `range`), **Bound** (read only for `range`, as `min,max`; either side may be blank for an open-ended bound). A row whose Column resolves to nothing inside the picked range MUST be reported by name, not silently skipped. A Rule value other than `blank`/`number`/`date`/`range`, a non-blank Bound component that does not parse as a number, or a `range` rule whose Bound is wholly blank, MUST cause `validate.scan` to refuse before scanning anything, naming the offending row - never silently treated as unbounded or as a rule that never fails. `usedBy` `validate.scan`. |
| `k_validate_outliers` | table | simple | **none** | Numeric bounds for `validate.highlight_outliers`. Three columns: **Column** (letter or header name, same resolution rule as above), **Min**, **Max** (either may be blank for an open-ended bound). A row whose Column resolves to nothing inside the picked range MUST be reported by name. A non-blank Min or Max that does not parse as a number MUST cause `validate.highlight_outliers` to refuse before checking anything, naming the offending row - never silently treated as unbounded on that side. `usedBy` `validate.highlight_outliers`. |

`k_validate_scan_rules` header: `Column (letter or header name)|Rule (blank, number, date, range)|Bound (range rule only: min,max)`.
`k_validate_outliers` header: `Column (letter or header name)|Min (blank = no floor)|Max (blank = no ceiling)`.

**Both table fields declare NO default rows, deliberately — extending the `k_clean_formats`
rule (§8, package `clean`) to two more fields, for a sharper reason.** `k_clean_formats` with a
demo row risks an invisible number-format change; these two, with a demo row, risk *colouring
real cells* — `validate.scan` writes an exceptions sheet named from a rule nobody configured,
and `validate.highlight_outliers` colours cells against a bound nobody set, on a deployment's
very first run, before anyone has looked at Setup. A conforming implementation MUST NOT ship
default rows for either field. Worked examples live in each field's help text instead.

**`validate.reconcile` needs no config field beyond its destination name.** Its two ranges, and
the key column(s) within each, describe the SELECTION the caller is about to make — not the
deployment — so per §5's rule (established for `clean`, §8) they are run-time inputs, asked at
the moment the action runs, never stored on Setup.

**Actions**

| id | caption | semantics |
|---|---|---|
| `validate.reconcile` | Reconcile Two Datasets | Prompt for two ranges (Dataset A, Dataset B); row 1 of each is always its header, not asked. Prompt for the key column(s) of each, numbered from the LEFT of that range. A row whose key column(s) are ALL blank MUST be reported as its own status, `Empty key`, one row at a time, and MUST NEVER be entered into the match index - an empty key on each side MUST NOT be treated as a match. Build a key index per side from every remaining row (composite key -> the row(s) carrying it); a key present on only one side is reported `In A only` / `In B only`. A key whose count is greater than one on EITHER side MUST be reported as `Duplicate key`, with both sides' counts (zero for a side that lacks the key entirely), and its rows MUST NOT be compared column-by-column — an implementation MUST NOT pair same-key rows across sides by position or any other guess, and duplication MUST be evaluated on each side independently of whether the other side has the key at all. A non-key header name that occurs more than once in either range's header row MUST cause the action to refuse before comparing anything, naming the repeated header(s) and which side repeats them - an implementation MUST NOT silently compare only the first occurrence and drop the rest. A key matched exactly once on each side is compared across every header common to both ranges (excluding each side's own key column(s)); a numeric comparison MUST tolerate formatting differences (e.g. `100` vs `100.00`); a differing column is reported as `Differs`, one row per differing column. When there is at least one exception, the action MUST show the exception count and require an explicit confirmation before writing it to a new sheet named `k_validate_recon_dest`; declining or cancelling MUST change nothing. Neither input range is changed. |
| `validate.scan` | Scan for Data Issues | Prompt for a range and whether its first row is a header. Before any cell is read, every `k_validate_scan_rules` row MUST be validated: an unrecognised Rule value, a non-blank Bound component that does not parse as a number, or a `range` rule whose Bound is wholly blank, MUST cause the action to refuse, naming the offending row(s), rather than silently treating the row as unbounded or as never failing. Apply every remaining `k_validate_scan_rules` row to its resolved column; a cell failing its rule is written to a new sheet named `k_validate_scan_dest`, only when at least one exception exists and only after the action shows the exception count and the user explicitly confirms; declining or cancelling MUST change nothing. A rule row whose Column resolves to nothing in the picked range MUST be reported by name. The picked range is not changed. |
| `validate.highlight_duplicates` | Highlight Duplicates | Prompt for a range, whether its first row is a header, and the key column(s), numbered from the LEFT of that range. Colour every key-column cell belonging to a row whose key is shared with at least one other row in the range. A cell that already carries a fill colour of its own MUST be left alone and counted separately, never overwritten. Candidate cells MUST be deduplicated by address before the preview count is computed, so a repeated column selection cannot inflate the count. Before colouring anything, show a count of exactly what will be marked and require an explicit confirmation; declining or cancelling MUST change nothing. No value, formula, or existing formatting is changed. The undo instruction given afterward MUST refer only to the cells actually coloured, never to the whole picked range. |
| `validate.highlight_outliers` | Highlight Outliers | Prompt for a range and whether its first row is a header. Before any cell is read, every `k_validate_outliers` row MUST be validated: a non-blank Min or Max that does not parse as a number MUST cause the action to refuse, naming the offending row(s), rather than silently treating that side as unbounded. For every remaining `k_validate_outliers` row, colour every numeric, non-blank cell in its resolved column that falls outside `Min`/`Max`. Same already-filled-cell carve-out, same deduplicated preview count, same confirm-before-colouring gate, same "nothing else changes" guarantee, and same scoped undo instruction as `validate.highlight_duplicates`. A configured column that resolves to nothing in the picked range MUST be reported by name. |
| `validate.missing_fields` | Flag Missing Required Fields | Prompt for a range, whether its first row is a header, and which column(s), numbered from the LEFT of that range, are mandatory. Colour every blank cell in those columns. Same already-filled-cell carve-out, same confirm-before-colouring gate, same "nothing else changes" guarantee as the other two highlight actions. |

**Colouring is reporting, not repair.** §7 classes `validate` as "checking data against a rule
and reporting what fails; never repairs silently." The three actions above that write a fill
colour change no value, no formula, and no shape of the data — they mark, in place, exactly
what a written exceptions list marks on a separate sheet, using the sheet the data already
lives on as the report surface instead of a new one. A conforming implementation MUST make
this reversible without disturbing any other formatting: a cell that already carried a fill
colour of its own MUST NOT be recoloured, and a cell this touches is guaranteed to have carried
no fill immediately beforehand — so "clear the fill on the affected cells" always, and only,
undoes what the action did. **No action in this package may delete, move, or overwrite a value
or formula.**

### Package `format` - "Format & Present" (v1.0.0, contract v2.6.0)

Presentation only (§7): every action changes how a report looks and never a stored value.

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_format_highlight_rules` | table | simple | **none** | Rules for `format.conditional_highlight`. Three columns: **Column** (a column letter, or a header name matched against the first row of the picked range - same resolution rule as `k_clean_formats`) - **Condition** (`greater_than <value>`, `less_than <value>`, `equal_to <value>`, `contains <value>`, or `blank`, one operator plus its comparison value in one cell) - **Fill color** (one of `red`, `yellow`, `green`, `orange`, `blue`, `gray`). A row whose Column resolves to nothing inside the picked range MUST be reported, not silently skipped. Rules are evaluated top to bottom; a cell matching more than one rule takes the LAST match. `usedBy` `format.conditional_highlight`. |

`k_format_highlight_rules` header: `Column (letter or header)|Condition|Fill color`.

**This field ships with NO default rows, for the same reason `k_clean_formats` does not
(§8, package `clean`).** §5 resolves an empty stored table to the field's declared defaults,
so demo rows would make `format.conditional_highlight`'s required "nothing configured"
refusal unreachable, and would risk coloring a first-time user's own data by a rule they
never wrote.

**Actions**

| id | caption | semantics |
|---|---|---|
| `format.conditional_highlight` | Highlight by Rule | Prompt for a range, then whether its first row is a header. Read `k_format_highlight_rules`; if empty, refuse naming where to configure it. Otherwise, before changing anything, show the range and the rule count and ask to proceed. On yes: clear any existing native conditional formatting in the picked range, then for each rule, resolve its Column against the range and set a static fill color (§ above - not a live conditional-formatting rule) on every data-row cell in that column matching the rule's condition. A rule whose Column resolves to nothing MUST be reported by name, not silently dropped. Changes cell fill only; no value changes. |
| `format.style_headers` | Style Header Row | Prompt for a range. Bold and fill the range's first row, then autofit every column in the range. Overwrites any formatting already on that row. No confirmation - this is unambiguously what the action is for. |
| `format.report_title` | Add Report Title | Refuse if the active sheet is engine-owned. Prompt for title text (run-time input, not stored config - §3's "range is not a kind" precedent extended to free text); refuse if left blank. Show the row count that will be inserted and ask to proceed, noting that formulas and defined names on that sheet follow the shift automatically (native host behavior, §9) and that a formula whose result depends on its own row (`ROW()`, an `INDIRECT()` built from a row number) can return a different result after the shift. On yes: insert that many rows at row 1, merge and center the title text across the sheet's used width, and style it. This DOES write one new stored value - the title text, into the merged title cell; every value that already existed on the sheet keeps its own value and only its row position moves. |
| `format.table_style` | Apply House Table Style | Prompt for a range. Show what will happen (existing fill/border formatting in the range is replaced with a fixed banded-row + thin-border house style) and ask to proceed. On yes: apply it. No value changes. |
| `format.print_layout` | Prep for Print | Refuse if the active sheet is engine-owned. Set that sheet to landscape orientation, fit-to-one-page-wide, a fixed narrow-margin preset, and a repeating first print row. Fixed preset, no configuration. No confirmation - touches print setup only, no cell. |

### Package `export` - "Export" (v1.0.0, contract v2.6.0)

A second package in the `export` domain alongside the grandfathered `importexport.export`
(§7). Declares its own fields under new keys; it does **not** read `k_export_base`,
`k_export_offset` or `k_export_tabs` from `importexport`'s manifest - those are a different
package's declared defaults, and the reference implementation's config-read fallback is
scoped per owning package (see this wave's plan, Decision 5), so borrowing another package's
key is not just a naming collision but a live default-resolution bug waiting for an
unconfigured deployment.

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_export_folder` | text | simple | *(blank)* | Folder exported files are written to. **Blank means the deployment's data location** (§9), same convention as `k_import_folder`. `usedBy` `export.pdf, export.per_sheet`. |
| `k_export_date_offset` | int | advanced | `-1` | Days to shift the date placed in front of an exported file's name. Same semantics as `importexport`'s `k_export_offset`, under its own key (§ above). `usedBy` `export.pdf, export.per_sheet`. |
| `k_export_pdf_base` | text | simple | `Report` | Base name for `export.pdf`'s single dated PDF -> `<date> <base>.pdf`. `usedBy` `export.pdf`. |
| `k_export_persheet_format` | text | simple | `pdf` | File type `export.per_sheet` writes: `pdf` or `xlsx`. Any other stored value MUST be refused at run time, naming the two accepted tokens. `usedBy` `export.per_sheet`. |

**Actions**

| id | caption | semantics |
|---|---|---|
| `export.freeze_values` | Freeze Formulas to Values | Prompt for a range. Count its formula cells (including dynamic-array spill anchors); refuse if none. Show the count and ask to proceed, stating this cannot be undone. On yes: convert every formula in the range - including spills and legacy CSE arrays - to a static value, using the same flattening behavior `importexport.export` already applies per sheet (§ Core_FreezeValues). Promoted from being embedded inside `importexport.export` to a standalone, range-scoped action; `importexport.export`'s own behavior is unchanged. |
| `export.pdf` | Export to PDF | Resolve the export folder; refuse if the workbook has never been saved and no folder is configured, or if the configured folder does not exist. Ask whether to export the whole workbook, one sheet, or a range (workbook / sheet / range); an unrecognised answer is refused, naming the three accepted tokens. Build the dated output path (`k_export_date_offset` + `k_export_pdf_base`); refuse rather than overwrite if a file already exists at that path. For "sheet", refuse if the named sheet does not exist or is engine-owned. For "range", the picked range temporarily becomes the target sheet's print area for the export and is restored to its prior value immediately after, whether the export succeeds or fails. Writes one PDF. |
| `export.per_sheet` | Export Each Tab Separately | Validate `k_export_persheet_format` is `pdf` or `xlsx`; refuse otherwise. Resolve and validate the export folder as `export.pdf` does. Build the list of every VISIBLE, non-engine-owned sheet (§ Decision 6 - hidden sheets are excluded too, on top of the required engine-sheet exclusion); refuse if none. Check every target file path for an existing collision, AND every target path against every other target path for a collision WITHIN this batch (sanitizing is many-to-one - two different, both-legal sheet names can produce the identical file name); refuse the WHOLE batch, naming every collision, if even one exists - no file is written until every target is clear. That refusal is all-or-nothing; the write itself is not transactional - it proceeds file by file, and a failure partway through leaves the files already written in place, with the failure message stating how many succeeded before the error. Otherwise show the full file list and ask to proceed. On yes: write one file per sheet, named `<date> <sheet name>.<ext>`. The `xlsx` path flattens every formula in a full workbook copy (via `Core_FreezeValues`, same technique `importexport.export` already relies on) before isolating each sheet, so a sheet with a cross-sheet formula does not export with a broken reference. |
### Package `summarize` - "Summarize & Analyze" (v1.0.0)

**Fields**

None. Every action either takes no input (`summarize.recalculate`), operates on every sheet
(`summarize.refresh_pivots`), or prompts for its range and columns at run time
(`summarize.build_pivot`, `summarize.freeze_pivot`) - none of the four describes a fact about
the deployment, so none is a config field. Same precedent `clean` set before v2.3.0 gave it
its first field.

**Actions**

| id | caption | semantics |
|---|---|---|
| `summarize.refresh_pivots` | Refresh All PivotTables | Refresh every PivotTable on every sheet in the workbook, skipping engine-owned sheets (§7's `Core_IsEngineSheet` test). Success MUST be read from `RefreshTable`'s own documented Boolean return value, not merely inferred from the absence of a raised error. A PivotTable whose refresh raises or returns `False` MUST NOT stop the others; it MUST be counted and named (sheet, pivot, error) in the completion report rather than silently skipped or aborting the run. With no PivotTables anywhere in the workbook the action MUST refuse rather than report a vacuous success. |
| `summarize.recalculate` | Recalculate Workbook | Force a full recalculation of every formula in the workbook. This action MUST NOT refuse under any condition: it changes no stored value, so no precondition can be unmet. |
| `summarize.build_pivot` | Build a Summary Table | Prompt for a range whose first row MUST be treated as its header, then which column(s) to group by and which to sum. Write the result as a new PivotTable on a new sheet. Every value column is aggregated by SUM; a per-column aggregation choice is out of scope for this version. The destination sheet MUST NOT overwrite an existing sheet - same refusal rule as `import.combine_worksheets`. |
| `summarize.freeze_pivot` | Freeze a Pivot to Values | Prompt for a cell; resolve the PivotTable that cell belongs to, refusing by name if it belongs to none. Copy that pivot's current output - values, number formats, and column widths - to a new sheet. The action MUST NOT modify the source PivotTable (no refresh, no write to its cells). Same destination-sheet refusal rule as `summarize.build_pivot`. |

### Package `enrich` - "Match & Enrich" (v1.0.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_enrich_lookup_sheet` | sheet | simple | `Ref_Master` | The reference sheet `enrich.mapped_columns` reads from. Unlike a destination-sheet field (e.g. `k_import_dest`), this one MUST already exist - a conforming implementation MUST refuse rather than create it. `usedBy` `enrich.mapped_columns`. |
| `k_enrich_lookup_key` | text | simple | `Store ID` | A column letter or header-text spec, resolved the same way as `k_clean_formats`' Column cell (§8, package `clean`): a one-to-three-letter value is read as a letter, anything else is matched against the target's own first row. Used against BOTH the reference sheet and whatever range the caller selects - the matching column has one name for both sides. `usedBy` `enrich.mapped_columns`. |
| `k_enrich_mapped` | table | simple | none | One row per column to add. Two columns: **Reference column** (letter or header text in the reference sheet) and **New column header** (the header given to the added column). Declares NO default rows, for the same reason `k_clean_formats` does not (§8, package `clean`): §5 resolves an empty stored table to the field's declared defaults, so demo rows would make `enrich.mapped_columns`' required "nothing configured" refusal unreachable. `usedBy` `enrich.mapped_columns`. |

`k_enrich_mapped` header: `Reference column (letter or header)|New column header`.

**Actions**

| id | caption | semantics |
|---|---|---|
| `enrich.join` | Join / Enrich Columns | Prompt for a target range and a reference range, each with its header row included and treated as such - there is no "does this have a header" branch. Prompt for the ONE matching column on each side (a multi-column answer MUST be refused, not narrowed to its first entry) and which reference column(s) to copy in. Match by building an in-memory dictionary keyed on the reference's matching column, then scanning the target once - an implementation MUST NOT re-scan the reference table per target row (dictionary lookup, not nested iteration). Key comparison MUST be exact aside from leading/trailing whitespace and case; an implementation MUST NOT normalize keys in any way that could make two distinct source values compare equal (e.g. collapsing internal whitespace or stripping characters). A reference key seen more than once MUST resolve to its FIRST occurrence; the number of DISTINCT reference keys that had at least one duplicate MUST be counted and reported, never silently overwritten and never a reason to refuse the whole action. A target row matching no reference key, or whose own key is blank, MUST be left blank in the new columns and counted as unmatched, not treated as an error. New columns are written immediately to the right of the target range, headed with the reference table's own header text, and only after the implementation confirms every cell in that exact destination block is empty and the user has confirmed the write; if any cell there is already occupied, the action MUST refuse rather than overwrite it. |
| `enrich.mapped_columns` | Add Mapped Columns | Prompt only for the target range; the reference sheet, its matching column, and which columns to add come from `k_enrich_lookup_sheet` / `k_enrich_lookup_key` / `k_enrich_mapped`. MUST share the match implementation with `enrich.join` - same first-wins-and-counted-by-distinct-key duplicate-key rule, same key normalization policy, same unmatched-row rule, same append-to-the-right placement guarded by the same empty-destination check - differing only in where its five inputs come from. MUST refuse if the configured reference sheet does not exist, if the configured key does not resolve on the reference sheet, or if `k_enrich_mapped` is empty or none of its rows resolve to a real reference column. |
### Package `backup` — "Protect & Back Up" (v1.0.0, contract v2.8.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_backup_folder` | text | simple | *(blank)* | Root folder for every action in this package. **Blank means the deployment's data location** (section 9), which is the common case. Backups land in a `Backup` subfolder here, snapshots in `Backup\Snapshots`, and archived files in `Archive\yyyy_mm`. `usedBy` all four actions. |
| `k_backup_base` | text | simple | `Backup` | Base name used inside a timestamped or month-end file name. `usedBy` `backup.timestamped_copy`, `backup.month_end_version`. |
| `k_backup_offset` | int | advanced | `0` | Days to shift today's date before computing the prior month end. Same convention as `k_export_offset` (package `importexport`): a negative value looks back further than today. `usedBy` `backup.month_end_version`. |
| `k_backup_snapshot_label` | text | simple | *(blank)* | Short tag appended to a snapshot's file name, e.g. `before Import`. Blank means no tag. `usedBy` `backup.snapshot`. |
| `k_backup_snapshot_keep` | int | advanced | `20` | How many of the most recent snapshots to keep. On each new snapshot, older ones past this count are deleted automatically; nothing else in the folder is touched. `usedBy` `backup.snapshot`. |
| `k_backup_archive_prefix` | file | simple | *(blank)* | Name-prefix of the generated file this action relocates, resolved the same way every `file` field is (section 3): the newest file in the deployment's data location whose name starts with this text. `usedBy` `backup.archive_output`. |

**Actions**

| id | caption | semantics |
|---|---|---|
| `backup.timestamped_copy` | Backup Now (Timestamped Copy) | Save a full copy of this workbook, exactly as it stands at the moment the action runs (including any edits not yet saved to the main file), into `<root>\Backup\` with a datetime-stamped name. This workbook's own name, path, and unsaved-changes state MUST be unaffected. MUST refuse rather than overwrite if the computed destination name already exists. |
| `backup.snapshot` | Snapshot Before a Risky Step | As `backup.timestamped_copy`, into `<root>\Backup\Snapshots\`, named with a full timestamp and an optional configured label. After a successful save, prune that folder to the configured keep-count by file modified time, oldest first, touching only files this action's own naming pattern produced. **This action provides no restore.** A conforming implementation MUST NOT claim or imply an automatic restore capability for the file it writes. |
| `backup.month_end_version` | Save Month-End Version | As `backup.timestamped_copy`, named for the last day of the calendar month before the reference date's month (`k_backup_offset` days from today), independent of which day of that month the action runs on. |
| `backup.archive_output` | Archive a Generated Output | Locate the newest file matching `k_backup_archive_prefix` (the `file` kind rule). MUST refuse if that file is currently open in this Excel instance rather than attempting to move it. Move it into `<root>\Archive\yyyy_mm\` (the current date's year and month), creating that folder if it does not exist. MUST refuse rather than overwrite if a file of that name already exists at the destination. |
### Package `workbook` - "Workbook Housekeeping" (v1.0.0, contract v2.9.0)

**Fields**

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_workbook_toc_name` | text | simple | `Contents` | Name of the Table of Contents sheet. A conforming implementation MUST create this sheet on first run and MUST refuse on any later run where a sheet by that name already exists, WITHOUT exception for a sheet the implementation itself built on a prior run - a plain cell value cannot prove that provenance without adding new engine state, so this field is create-only by design, not merely by omission. Rebuilding after adding, removing, or renaming a tab requires deleting the previous sheet by hand first, then running the action again. Kind is `text`, not `sheet`, for the same reason as `k_import_dest` (§8): a `sheet` field is validated to already exist, and this one must not. |
| `k_workbook_password` | text | advanced | *(blank)* | Password used by `workbook.protect_sheets` and `workbook.unprotect_sheets`. Blank means protect with no password. **Not a secret** - stored as plain text in the config store, like every other field; it deters accidental edits, the same as Excel's own Protect Sheet, and provides no confidentiality. |

**Actions**

| id | caption | semantics |
|---|---|---|
| `workbook.unhide_sheets` | Unhide All Worksheets | Set every hidden or very-hidden worksheet to visible, EXCLUDING every sheet the implementation's engine owns (the Setup surface, the config store and its staging sheet, the package-enable sheet, every generated button page). The engine-owned config store MUST NOT be unhidden by this action under any configuration. |
| `workbook.protect_sheets` | Protect All Sheets | Protect every sheet NOT owned by the engine, with `k_workbook_password` (or no password if blank). An implementation MUST choose protection parameters that do not block its own future ability to remove and replace its own placed controls on a protected surface (Excel reference: `DrawingObjects:=False`) - protecting with the platform's own default MUST NOT be assumed safe. Engine-owned sheets MUST NOT be protected by this action, so a later rebuild retains its own access to them. |
| `workbook.unprotect_sheets` | Unprotect All Sheets | Unprotect every protected sheet NOT owned by the engine, using `k_workbook_password`. A sheet whose protection does not accept that password MUST be reported by name as not unprotected; the action MUST NOT stop partway and MUST attempt every remaining sheet regardless of one sheet's failure. |
| `workbook.build_toc` | Build Table of Contents | Create (§8 field semantics above) a sheet named `k_workbook_toc_name` listing every VISIBLE sheet not owned by the engine and not the Table of Contents sheet itself, each as a same-workbook hyperlink to that sheet, with an embedded apostrophe in a sheet name DOUBLED per the anchored-placement quoting rule (§6). A hidden sheet MUST NOT be listed (a link to it cannot be navigated to). The destination-name check MUST be performed unconditionally, before any other branch, and MUST refuse per the reserved-name rule (§7) and per any existing sheet of that name - including the implementation's own prior output - exactly as any other engine-created destination sheet. |
| `workbook.unhide_rows_cols` | Unhide All Rows & Columns | On the ACTIVE sheet only, set every row and every column to visible. MUST refuse if the active sheet is engine-owned (an engine-owned surface may hide rows/columns as part of its own generated layout). Because a filter-hidden row and a manually hidden row are not distinguishable, an implementation MUST warn, and obtain confirmation, before unhiding on a sheet with an active AutoFilter; outside that case no confirmation is required. |
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

- **Range-scoped PDF export (`export.pdf`, range mode)** - a temporary print-area override on
  the target sheet, restored immediately after export. Neutral: render exactly the picked
  cells to the platform's page-description format, without permanently changing the sheet's
  own print setup.
- **Per-tab explode (`export.per_sheet`)** - one output file per visible, non-engine sheet,
  named `<date+offset> <sheet name>.<ext>`. Neutral: the platform's nearest equivalent to "one
  file per tab, skip the toolkit's own surfaces."
---

## 10. Change control

- This contract is versioned with SemVer, independent of the repo and package
  versions.
- **Additive** changes — a new kind, field, action, or package; a new optional
  attribute — bump the **minor** version.
- **Breaking** changes — renaming or removing a `key` or action `id`; changing a
  field's kind, a value's encoding, or an action's semantics — bump the
  **major** version.
- **Two axes (v1.2.0, normative).** The above categories apply to **stored
  configuration**: a `key`, an action `id`, a field's `kind`, a value's
  encoding, or an action's own behaviour becoming invalid or reinterpreted is
  **breaking**. Changes to **engine presentation** — how the engine lays out,
  groups, titles, or navigates the surfaces it generates from a valid config —
  are **additive**, because every existing config keeps its meaning and keeps
  working unchanged. A platform implementing an older minor version will render
  those surfaces differently; that is expected and is not a compatibility break.
- The contract is the source of truth going forward: the reference VBA build and
  every port MUST conform. A discovered divergence is reconciled **in favor of
  this document**.
- **Reserved for first use — DISCHARGED at v1.2.0.** The `bool` kind (§3) had its
  encoding reserved until first use. `clean_hasheader` (package `clean`, §8) is that
  first use, so §3's row now states the fixed encoding and this bullet is a record,
  not an outstanding obligation. `clean_hasheader` has since been withdrawn; the
  discharge stands regardless. A reservation is discharged by the first use, not
  held open by that use continuing to exist — reopening it would let a later field
  pick a different encoding from the one already shipped and read.
- **v2.9.0 (additive):** the `workbook` domain's package, fields, and five actions (§8); one
  new reserved-name normative rule for `text`-kind destination fields (§7). The reserved-name
  rule reclassifies one previously-succeeding configuration (a destination name inside the
  reserved family, e.g. `"Macros 3"`, on a field that had no prior existing-name collision) as a
  refusal. Classed additive on the same reasoning v2.4.0 used for the chain-halt rule: the
  configuration being refused was never actually usable - it produced a sheet the very next
  rebuild destroyed unconditionally - so refusing it prevents a data-loss trap rather than
  withdrawing behaviour a deployment could rely on. `SCHEMA_VER` does not move.
- **v2.8.0 (additive):** a new package `backup` with six fields and four actions (§8). The
  `backup` domain itself is not new - section 7 already listed it `active` ahead of any package
  using it - so this bullet adds only the package, not the domain. **Additive by the two-axes
  rule:** no existing `key`, action `id`, `kind`, or encoding changes meaning; every prior stored
  configuration keeps working exactly as it did. `SCHEMA_VER` does **not** move.
- **v2.7.0 (additive):** two new packages, `summarize` (four actions, no fields) and `enrich`
  (three fields, two actions). Both domains were already listed `active` in §7's vocabulary
  table before this version - this bullet backs that listing with a shipping package, and
  changes no vocabulary. `SCHEMA_VER` does **not** move - it stays `"v2.0"` and every existing
  configuration keeps its meaning. No `key`, action `id`, or field `kind` from any earlier
  version is renamed, removed, or reinterpreted.
- **v2.6.0 (additive):** two new packages, `format` (one field, five actions) and `export`
  (four fields, three actions) - §8. `export.freeze_values` promotes logic that already
  shipped embedded inside `importexport.export`; that action's own id, fields and behavior
  are unchanged. **Additive by the two-axes rule:** no existing `key`, action `id`, `kind` or
  encoding is touched or reinterpreted; every new field and action is new surface on an
  already-`active` domain (§7, v2.3.0). `SCHEMA_VER` does **not** move - it stays `"v2.0"`.
  **One judgement call worth stating, honestly rather than softened:** `format.report_title`
  DOES write a stored value - the title text goes into the newly inserted, merged title cell
  - and it moves every value that already existed on the sheet below the insert point to a
  new row address. A formula whose RESULT depends on its own row (`ROW()`, or an
  `INDIRECT()` string built from a row number) can therefore return a different result after
  the shift, even though the formula's own stored text is untouched. This still belongs
  under the `format` domain, not `export`/`clean`, because every value that existed on the
  sheet before the action ran keeps its own value - the ONE new value is the title text
  itself, which is the entire declared purpose of the action, disclosed in its own confirm
  prompt and completion message, not a side effect discovered later. A platform whose
  row-insert does not auto-adjust formulas and defined names the way Excel's does MUST
  perform the equivalent adjustment itself, or refuse the action - silently breaking a
  reference is not a conforming realization of this semantics.
- **v2.5.0 (additive):** a new package `validate` with four fields and five actions (§8):
  `validate.reconcile`, `validate.scan`, `validate.highlight_duplicates`,
  `validate.highlight_outliers`, `validate.missing_fields`. The `validate` domain has been
  listed `active` in §7's vocabulary table since v2.3.0, before any package implemented it;
  this version is the first to populate it. **Additive by the two-axes rule:** every existing
  stored value keeps its meaning, no existing `key`, `id`, `kind` or encoding changes.
  `SCHEMA_VER` does **not** move — it stays `"v2.0"`.
- **v2.4.0 (additive):** a refusal becomes a distinct action outcome from a failure (§6); an
  implementation MUST expose a single record-and-present call for it; a chain now stops at a
  refused step instead of asking. `SCHEMA_VER` does **not** move — it stays `"v2.0"` and
  every existing configuration keeps its meaning.
  **The judgement call in classing this additive.** By the letter, a chain that used to ask
  "continue?" and now stops is a behavior change to an existing valid configuration. It is
  classed additive because no *stored value* is reinterpreted — the chain, its steps and
  their order all still mean exactly what they meant — and because the behavior being
  removed could not be relied on for a correct result: continuing past a step that produced
  no output ran the remainder against stale data and reported success. A platform
  implementing v2.3.0 will ask where a v2.4.0 platform stops; that is expected and is not a
  compatibility break.
- **v2.3.0 (additive):** the action-id domain vocabulary becomes a normative, marked table
  in §7 and the enforcement it always described becomes real; a new package `import` with
  four fields and three actions (§8); one new `clean` field and four new `clean` actions
  (§8). **Additive by the two-axes rule:** every existing stored value keeps its meaning,
  no `key`, `id`, `kind` or encoding changes, and the eight shipped ids are listed as
  permanent exceptions so the new shape rules cannot invalidate them. `SCHEMA_VER` does
  **not** move — it stays `"v2.0"` and every v0.5.0 and v0.6.0 config remains readable.
  The one thing that changes an outcome is a **build** outcome, not a runtime one: a
  source tree declaring a malformed or off-vocabulary action id now fails to build. No
  shipped source does.
  **One normative negative worth naming:** `k_clean_formats` MUST ship with no default
  rows. §5 resolves an empty stored table to the declared defaults, so demo rows would
  make `clean.standardize_formats`' required "nothing configured" refusal unreachable —
  the field's own semantics depend on the absence.
- **v2.2.0 (additive):** three hardening rules, none of which touches a stored value's
  meaning. (a) **Durable replacement (§5):** activating MUST stage, verify, then swap;
  delete-then-write is now non-conforming, and a reader MUST adopt an orphaned staged
  store. (b) **Marker absence must be earned (§5):** "absent" requires positive
  structural corroboration and MUST NOT be decided from a user-editable caption; an
  unexplained empty marker is "removed" and fails closed. (c) **`core.*` enforced
  (§7):** refusing a package action id in that namespace becomes normative rather
  than declared.
  **Additive by the two-axes rule:** every existing valid store keeps its meaning and
  keeps working. (b) is the only one that changes any outcome, and only for a store
  that is *already* non-conforming — a v1.2+ block whose machine-owned marker was
  removed by hand. That store previously activated as all-`button`; it is now refused.
  No conforming writer emits it, for the same reason §6's blank-cell case is additive:
  the marker is written on every render. `SCHEMA_VER` does not move — it stays `"v2.0"`
  and every v0.5.0 and v0.6.0 config remains readable.
- **v2.1.0 (additive):** `availability` defaults to `off` (§6); the pre-v1.2.0 legacy
  default is pinned to `button` independently (§5); reserves the `core.*` action-id
  namespace (§7); restates relegation presentation for disclosure (§3). **Additive, not
  breaking, by the two-axes rule above:** no stored value is reinterpreted — a stored
  `button` still means `button` — because a default applies only where nothing was
  stored. The stored-format marker does NOT move; `SCHEMA_VER` stays `"v2.0"` and every
  v0.5.0 config remains readable.
  **The one judgement call in that classification:** a **blank** `availability` cell
  used to read as `button` and now blocks Finalize (§6), which by the letter is
  reinterpreting a stored value. It is classed additive because **no conforming writer
  has ever emitted one** — an implementation writes the label for a resolved token on
  every render, and all three tokens have non-empty labels. A blank cell can therefore
  only arrive by hand-editing the authoring surface, which is unvalidated input under
  §5's fail-closed rules rather than stored configuration. If a port has ever written
  a blank cell, this change IS breaking for that port and it must migrate them first.
- **v2.0.0 (BREAKING):** withdraws the `clean_keycols` and `clean_hasheader` fields
  and changes `clean.dupes` to take both as run-time inputs (§8); adds §8's normative
  selection-vs-deployment rule and §2's no-visible-fields-no-band rule. Removing a
  `key` and changing an action's semantics are breaking by the rules above.
- **The stored-format marker moves with it, deliberately.** `SCHEMA_VER` goes to
  `"v2.0"`, so `CurrentSchemaMajor()` is 2 and every `_Config` written by v0.4.0
  (marker `v1.2`) reads as `other-major` and is **refused**. This is stricter than
  the data requires — the Action catalog block layout is byte-identical across this
  change — and it was accepted on the ground that **no deployment exists yet**: the
  bundle is in test workbooks only, so there is no installed base to refuse. Keeping
  one version number is simpler than decoupling the two and explaining the split in
  every port. The alternative was analysed and rejected in `docs/adr/0002`.
  **Revisit before the next breaking bump if an installed base exists by then** —
  at that point, refusing every stored config to express a change that touched no
  stored byte is a real cost, and decoupling becomes the better trade.
- **Asymmetry this creates, which is correct and not a defect.** A v1.1.0-era config
  has *no* marker, reads as `absent`, and still migrates cleanly through the legacy
  path. A v1.2.0 config has `v1.2`, reads as `other-major`, and is refused. Older is
  accepted where newer-but-not-current is not, because only the marker-less case has
  a defined legacy layout to migrate *from*.
- **v1.1.0 (additive):** adds the `chain` config entity, the reserved `chain.*`
  id namespace, engine-owned config, the grouping/layout placement code, and the
  quiet/dispatch convention. No existing key, action id, kind, value encoding, or
  semantics changed — v1.0.0 deployments remain valid.
- **v1.2.0 (additive):** adds the optional action `category` attribute, the
  optional field `usedBy` attribute, the `clean_hasheader` field and the
  now-fixed `bool` encoding (§3), the three-state deployment-level
  `availability` flag, the capacity-packed auto-layout rule for uncoded buttons,
  the column-cap overflow clause for coded layouts (a page's three-column and a
  column's six-button capacity bind a coded layout too, so an over-capacity
  column or button-stack moves to a further column - and, once its page is
  already at the three-column cap, on to a continuation page - rather than a
  page number or column a code names being silently unrenderable),
  the optional index surface with its button catalog (built, then **withdrawn before
  release** at the owner's call at the v0.4.0 manual gate — "needless clutter... most
  folks will only ever load a handful of macros or 1 or 2 chains into a workbook";
  the code is preserved unbuilt at `src/attic/HomeIndex.bas.txt` and the `Macros Home`
  name stays reserved in §7), the unconditional previous/next page navigation,
  the no-clipping rule for descriptions, the config schema marker with fail-closed
  migration, the anchored `<sheet>!<cell>` placement form with its engine-must-not-modify and
  must-sweep-on-rebuild rules, and the reserved `Chains` category + `Macros Home` name. No `key`,
  action `id`, kind, value encoding, or **action** semantics changed, so
  v1.0.0/v1.1.0 configs remain valid and read as all-actions-`button`. The
  *behaviours* replaced are engine layout for the uncoded case (single page →
  capacity-packed pages) and, for the coded case, an over-capacity column now
  overflowing instead of being silently undrawn — both engine presentation
  under the two-axis rule above, not stored-config compatibility, so both stay
  additive. The anchored placement form adds no stored `key`, `id` or `kind`; it extends an
  existing value's grammar with a form no valid v1.1.0 config could contain (a value
  containing `!` was previously invalid and blocked Finalize), so no existing config is
  reinterpreted. **Additive.** Owner-approved; see
  the Provenance section of `docs/designs/2026-07-24-finish-excel-track1-plan.md`
  for the three contested calls and who lost them.
