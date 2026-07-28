# SheetForge portable contract

**Contract version: 2.3.0** (2026-07-28) · SemVer, independent of the
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
  module that implements it** — an action id is a permanent dispatch key a user may type by
  hand (in the Excel reference implementation, into a chain step cell, where only the
  bracketed id is authoritative and any label beside it is decoration — §6), so encoding
  the module in it would make a future package split a major contract break for no benefit. The operation segment is
  descriptive where ambiguity looms (`clean.remove_blank_rows`, never `clean.blank2`).
  **No new id is ever camelCase.** Ids are immutable; captions, help text, entry points
  and module assignment stay free to change. There is no alias or migration table, by
  design: an alias means two permanent names for one action and would have to be ported to
  Apps Script.

  **The vocabulary is the table below, and it is normative.** The marker comment
  immediately above it is a parse anchor — `build/build.ps1` reads the table from this file
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
