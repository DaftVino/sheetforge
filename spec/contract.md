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
- **Precedence over tier (v1.2.0).** `usedBy`-relegation and `tier` (§4) are
  evaluated independently, but off-ness wins: a field relegated to the off band
  goes there regardless of its `tier`, and the `advanced` band (§4) contains
  only fields that are still in use. A field never renders in both bands.
- **No visible fields, no band (version pending).** A package contributing no
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

**Absent is not invalid.** An attribute or column introduced after the stored
marker's version (for example `availability`, §6, in a store marked pre-v1.2.0)
simply does not exist in that store — there is no value to validate, so its
absence is not a conflict, and the reader takes the documented default (`button`
for `availability`). Fail-closed applies only to a value that IS present but
unreadable or fails its grammar, and to a marker the reader does not recognise at
all — never to a column this document added after the store was last written.
This is what lets §10 promise that v1.0.0/v1.1.0 configs remain valid.

---

## 6. Action contract

An action is one thing a package *does* (it becomes a button). Each action is:

| Field | Meaning |
|---|---|
| `id` | Dotted `pkg.verb` token, globally unique (`importexport.import`, `clean.dupes`). |
| `caption` | Button label (may carry an ordinal, e.g. `"2.  Import Source Data"`). |
| `help` | Plain-language description shown on/near the button. |
| `category` | Optional display category (v1.2.0). Defaults to the package `name`. Groups actions on an index surface where a platform provides one (§6), and heads the columns of a button page; deliberately decoupled from the package, so two packages may feed one category and one package may span several. |
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
`<sheet>!<cell>` (e.g. `TabName!B4`), which anchors that one button to a single cell on a
surface the user owns, rather than filing it onto a generated page. The
discriminator is `!`. The two forms can never be confused — not because the platform
forbids `!` in a surface name (Excel, for one, does not) but because a valid **page
code** cannot contain `!` under the grammar above, so even a surface legitimately
named `1A` is unambiguous. An implementation MUST split on the **last** `!`, so a
surface whose own name contains one still resolves. The sheet part may
be quoted (`'TabName Data'!B4`) or bare even when it contains spaces (`TabName Data!B4`); an
embedded apostrophe is doubled inside the quotes. The cell part is a single A1
reference with optional `$` (`TabName!$B$4`); a range (a colon) is rejected.

**Normative, authoring-surface quirk (v1.2.0).** Where the authoring surface itself
consumes a leading apostrophe — Excel treats one as a "this cell is text" prefix and
strips it from the stored value, so a user who types `'TabName Data'!B4` into a cell
stores `TabName Data'!B4` — an implementation MUST accept the resulting
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
- **Action id** — `pkg.verb`, unique.
- **Reserved:** the id namespace `chain.*` is reserved for engine chains; no package action id may use it.
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

### Package `clean` — "Clean & Standardize" (v1.0.0)

**Fields**

> **⚠ VERSION DECISION PENDING — do not release this contract until resolved.**
> Removing `clean_keycols` and `clean_hasheader` is **breaking** by §10's own rule
> (removing a `key`; changing an action's semantics), which points at **v2.0.0**.
> But `SCHEMA_VER` is `"v1.2"` and `CurrentSchemaMajor()` gates migration: if the
> stored-format marker follows the contract to major 2, every existing v0.4.0
> `_Config` reads as `other-major` and `Create Setup Tab` **refuses outright** —
> a hard break for every current deployment, for a change that alters no stored
> layout at all. The Action catalog block is byte-identical either way.
> The likely resolution is that the two numbers **decouple**: `SCHEMA_VER`
> describes the stored block layout and stays `v1.2`; the contract version
> describes the published surface and goes to v2.0.0. That needs an ADR, because
> §10 currently implies they move together. **Owner decides.**

**None (version pending).** `clean` declares no configuration fields at all, and a
conforming implementation MUST therefore render no configuration band for it —
see §2's rule on a package with no visible fields.

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
- **⚠ UNRELEASED, VERSION UNDECIDED.** Withdrawing `clean_keycols` and
  `clean_hasheader` (§8) and changing `clean.dupes` to take run-time inputs is
  **breaking** by the rules above, which points at **v2.0.0**. The open question is
  whether `SCHEMA_VER` — the stored-format marker that `CurrentSchemaMajor()` gates
  migration on — follows it. It must not: the Action catalog block layout is
  unchanged, so a marker bump to major 2 would make every existing v0.4.0 `_Config`
  read as `other-major` and be **refused outright**, breaking every deployment for a
  change that alters no stored layout. That implies the two version numbers
  **decouple**, which §10 does not currently contemplate. **Needs an ADR and the
  owner's decision before release.**
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
