# Part A — Progressive Disclosure Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Invert the Availability default from `button` to `off` so a fresh install shows a deliberate catalog rather than every button at once, without changing what any existing deployment does.

**Architecture:** Four code sites currently resolve "no stored answer" to `button`. Three change. The fourth — a legacy store with no format marker — must keep `button`, and since legacy and fresh are both an empty map, `ReadAvailMap` carries provenance forward in a reserved dictionary key rather than the caller guessing. `usedBy` is then wired onto `PkgImportExport`'s nine fields so turning an action off actually hides its settings.

**Tech Stack:** VBA (Excel/Windows), `pwsh build/build.ps1`, `node scripts/vba-syntax-check.mjs`. No test runner exists.

**Spec:** `docs/designs/2026-07-27-part-a-progressive-disclosure-spec.md`
**Decisions:** `docs/adr/0003-availability-default-and-legacy-migration.md`
**Panel:** `docs/panels/2026-07-27-part-a-availability-default.md`
**Closes:** #6, #10

## Global Constraints

- **TDD is formally OFF** (`CLAUDE.md`). Excel cannot be driven headlessly and no VBA runner exists. **Do not write a test suite.** The verification cycle for every task is: edit `src/` → `pwsh build/build.ps1` → `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` → commit. Behavioural verification is the manual gate in Task 5, run by the owner.
- **`src/` is the source of truth.** Never edit `dist/MacrosBundle.bas`; it is a gitignored build artifact. A `.bas` edit does nothing in Excel until the bundle is rebuilt and re-imported by hand.
- **`Option Explicit` everywhere.** Declare every variable.
- **25 line-continuation limit per statement.** One `Collection.Add` per line; never a giant `Array(...)`.
- **Flat namespace.** Core's API is `Core_*`; everything else `Private`. Shared helpers live ONLY in Core.
- **No `eNum`, or any case-variant of a VBA keyword, as an identifier.** `Any`, `Name`, `Len` and friends are reserved; `Dim any As Boolean` is a compile error the VBE reports against the *procedure declaration line*, which neither structural gate can see.
- **The contract is the source of truth** (`spec/contract.md` §10: "the reference VBA build and every port MUST conform"). Silent deviation requires an ADR; per `CLAUDE.md` it is "the one unforgivable move".
- **Build must report `Core + 2 package(s)` with balanced Sub/Function counts.** A duplicate proc name across modules fails the build.
- **Never read an `.xlsm`.** Never read `src/core/Core.bas` whole (210KB) — grep `docs/code-map.md`, then read the slice.
- **No attribution footers** in any commit message or PR body.

---

### Task 1: Contract v2.1.0

Contract first, in its own commit, before any code uses it. Per ADR 0003 decision 2 and §10's rule that the reference build MUST conform, shipping code that occupies an unreserved namespace is a silent deviation.

**Files:**
- Modify: `spec/contract.md` §3, §5, §6, §7, §10

**Interfaces:**
- Consumes: nothing.
- Produces: the reserved action-id namespace `core.*`, relied on by Task 3's `AVAIL_LEGACY_KEY`.

- [ ] **Step 1: §6 — flip the documented default**

Find the `availability` row in §6's attribute table (~line 258):

```
| `availability` | Deployment config, not manifest data (v1.2.0): one of `button`, `chains-only`, `off`. Defaults to `button`. |
```

Replace with:

```
| `availability` | Deployment config, not manifest data (v1.2.0): one of `button`, `chains-only`, `off`. **Defaults to `off` (v2.1.0)** — an action arrives switched off and is turned on deliberately. This changes no stored value's meaning: a stored `button` still reads as `button`. The default applies only where nothing was ever stored. For the pre-v1.2.0 legacy case the default is pinned separately and does NOT follow this row — see §5. |
```

- [ ] **Step 2: §5 — pin `button` for legacy. THIS IS THE LOAD-BEARING EDIT**

§5's "Absent is not invalid" paragraph (~lines 237-244) currently points at §6's default. Once Step 1 lands, that pointer silently flips legacy behaviour to `off` and destroys the guarantee a manual gate has already verified. Find:

```
absence is not a conflict, and the reader takes the documented default (`button`
for `availability`).
```

Replace with:

```
absence is not a conflict, and the reader takes the **pinned legacy default**, which
for `availability` is `button` — fixed at this value as of v2.1.0 and deliberately
NOT a reference to §6's current default. §6's default moved to `off` in v2.1.0; this
one did not move with it, because a store written before the column existed described
a deployment in which every action was live, and reading it as `off` would silently
remove every button from a working installation. A reader MUST treat these as two
independent values.
```

- [ ] **Step 3: §7 — reserve the `core.*` namespace**

After the existing `chain.*` reservation bullet, add:

```
- **Reserved (v2.1.0):** the action-id namespace `core.*` is reserved for engine-internal
  sentinels; no package action id may use it. The reference implementation uses
  `core.legacy` as a provenance flag inside the in-memory availability map, so an id
  colliding with it would be read as engine state rather than as an action.
```

- [ ] **Step 4: §3 — restate presentation for disclosure**

Add to §3's presentation notes:

```
- **Disclosure (v2.1.0).** A field whose actions are all `off` is relegated, and a
  platform MAY present relegation as a collapsed band or as hidden-in-place. **Omitting
  the field entirely remains non-conforming** either way: the engine re-seeds its
  authoring surface from stored config, so an omitted field costs the user a
  toggle → activate → rebuild → fill → activate cycle and discards anything typed but
  not yet activated.
```

- [ ] **Step 5: §10 — record the version**

Add above the `v2.0.0` entry:

```
- **v2.1.0 (additive):** `availability` defaults to `off` (§6); the pre-v1.2.0 legacy
  default is pinned to `button` independently (§5); reserves the `core.*` action-id
  namespace (§7); restates relegation presentation for disclosure (§3). **Additive, not
  breaking, by the two-axes rule above:** no stored value is reinterpreted — a stored
  `button` still means `button` — because a default applies only where nothing was
  stored. The stored-format marker does NOT move; `SCHEMA_VER` stays `"v2.0"` and every
  v0.5.0 config remains readable.
```

- [ ] **Step 6: Verify no other section still defers to §6's default**

Run:

```sh
grep -n 'documented default\|defaults to .button\|Defaults to .button' spec/contract.md
```

Expected: only §5's newly pinned wording and §6's own row. **Any other hit is another place that will silently flip — fix it before committing.**

- [ ] **Step 7: Commit**

```bash
git add spec/contract.md
git commit -m "feat(contract)!: v2.1.0 - availability defaults to off, legacy pinned to button

Additive under section 10's two-axes rule: no stored value is
reinterpreted, because a default applies only where nothing was stored.
SCHEMA_VER does not move and v0.5.0 configs stay readable.

Section 5's legacy default is now PINNED to button rather than pointing at
section 6. Without that pin, moving section 6's default to off would
silently flip the pre-v1.2 legacy path and remove every button from a
working installation - the exact case the v0.5.0 gate verified by hand."
```

---

### Task 2: Fail closed on a blank or unrecognised Availability cell

Fixes #10. Do this **before** Task 3: the laundering bug is materially worse once `off` is the default, so close it first.

**Files:**
- Modify: `src/core/Core.bas:1918-1939` (`ParseAvail`), `src/core/Core.bas:1990-1995` (`AvailLabel`)

**Interfaces:**
- Consumes: nothing.
- Produces: `ParseAvail` and `AvailLabel` both return `""` for anything they cannot resolve. Task 3 relies on `""` propagating to `ComposeReport`'s existing structural block.

- [ ] **Step 1: `ParseAvail` — a blank cell stops meaning `button`**

Replace lines 1918-1928 (the comment block and the blank-cell early exit):

```vba
' Parse an Availability cell to one of the three tokens.
'   ""      -> ""            (v2.1.0: blank is NOT an answer; the CALLER must block)
'   invalid -> ""            (the CALLER must block Finalize)
' It must never fall back to Button on an unrecognised string: a user who typed
' "of" meaning "off" would otherwise be handed a live destructive button.
' Blank joined that rule in v2.1.0. It used to read as Button, which was correct
' while Button was the default; with the default at Off, silently reading a
' cleared cell as EITHER value is wrong - Button resurrects an action the user
' switched off, Off removes one they still want. Neither is recoverable from the
' result, so blank must block instead. ComposeReport already refuses to activate
' on "" and needs no change.
' Yes/No are accepted as synonyms so a user upgrading from a two-value habit, or
' one who types the obvious thing, is understood rather than blocked.
Private Function ParseAvail(ByVal v As String) As String
    Dim s As String
    s = LCase(Trim(v))
    If Len(s) = 0 Then ParseAvail = "": Exit Function
```

Leave the `Select Case` block below unchanged.

- [ ] **Step 2: `AvailLabel` — stop laundering corruption into "Button"**

`Case Else` currently returns `"Button"`, which is how `button` is produced at all. Add an explicit `Case AVAIL_BUTTON` before changing `Case Else`, or every action renders blank. Replace lines 1989-1996:

```vba
' The display label written into the cell for a token.
'
' NOT display-only: RenderPlacementBlock writes this result into the Setup cell
' (Core.bas ~2101) and ReadAvailMap reads that same cell back through ParseAvail
' (~2015), so the label round-trips into stored configuration.
'
' Case Else therefore fails CLOSED (issue #10). It used to return "Button", which
' meant an unrecognised stored value was silently rewritten as a live button on
' the next 'Create Setup Tab' - bypassing ComposeReport's block, which only ever
' sees "" and never saw it again once the rebuild had overwritten the cell.
' Returning "" leaves the cell blank so that block fires as designed.
Private Function AvailLabel(ByVal token As String) As String
    Select Case token
        Case AVAIL_BUTTON: AvailLabel = "Button"
        Case AVAIL_CHAIN:  AvailLabel = "Chains only"
        Case AVAIL_OFF:    AvailLabel = "Off"
        Case Else:         AvailLabel = ""        ' fail closed - see above
    End Select
End Function
```

- [ ] **Step 3: Widen `ComposeReport`'s message to cover blank**

At `src/core/Core.bas:2578` the block reads "UNRECOGNISED Availability", which reads oddly for an empty cell. Replace:

```vba
        s = s & vbCrLf & "MISSING or UNRECOGNISED Availability (use Button, Chains only, or Off):" & _
            vbCrLf & badAvail
```

- [ ] **Step 4: Build and syntax-check**

```sh
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: `Core + 2 package(s)`, balanced `Sub`/`Function` counts, `no diagnostics`.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas
git commit -m "fix(core): fail closed on a blank or unrecognised Availability cell

Closes #10.

AvailLabel's Case Else returned Button, and RenderPlacementBlock writes
that result into a Setup cell which ReadAvailMap reads back through
ParseAvail - so an unrecognised stored value was silently rewritten as a
live button on the next Create Setup Tab. That bypassed ComposeReport's
block, which only fires on \"\" and never saw it again once the rebuild had
overwritten the cell.

Blank now blocks too. While Button was the default, reading a cleared cell
as Button was defensible; with the default at Off, either reading is wrong
and unrecoverable from the result."
```

---

### Task 3: Carry provenance, and flip the default

The core change. `ReadAvailMap` stops discarding what `MarkerVerdict` already knows.

**Files:**
- Modify: `src/core/Core.bas` — module constants (~line 50), `ReadAvailMap` (2004-2017), `AvailFor` (2019-2025)

**Interfaces:**
- Consumes: `core.*` reserved by Task 1; fail-closed `ParseAvail`/`AvailLabel` from Task 2.
- Produces: `AvailFor(availMap, id)` — **signature unchanged**. All 6 call sites and all 6 `ReadAvailMap` call sites keep working untouched, because the flag travels inside the returned dictionary.

**Why a reserved key and not a parameter:** verified that nothing enumerates the map — every consumer is a keyed `.Exists(id)` or `AvailFor(map, id)` lookup, so a sentinel key can never be mistaken for an action. Threading a parameter would mean editing 12 call sites for no gain.

- [ ] **Step 1: Add the sentinel constant**

Next to `AVAIL_BUTTON` / `AVAIL_CHAIN` / `AVAIL_OFF` (~line 50):

```vba
' Provenance sentinel inside an availability map - NOT an action id. Reserved by
' contract section 7 (v2.1.0) as part of the core.* namespace, because section 7
' states no character rule for action ids and a sentinel therefore cannot be
' assumed collision-proof. Present => the map came from a pre-v1.2 block whose
' Availability column does not exist, so an absent id means Button, not Off.
Private Const AVAIL_LEGACY_KEY As String = "core.legacy"
```

- [ ] **Step 2: `ReadAvailMap` flags the legacy case instead of exiting silently**

Replace the `MarkerVerdict` early-exit at line 2011:

```vba
    If MarkerVerdict(ws, kr, detail) = "absent" Then
        ' Pre-v1.2 block: the Availability column does not exist, so there are no
        ' values to read - but the map must still say WHERE it came from. Fresh
        ' (no _Config) and legacy are both an empty map, and they need opposite
        ' defaults, so the emptiness alone cannot decide it.
        map(AVAIL_LEGACY_KEY) = AVAIL_BUTTON
        Exit Function
    End If
```

Also update the procedure's header comment (2998-2003) — it currently says "AvailFor treats absence as Button", which stops being true in general:

```vba
' id(lower) -> availability token, from the action catalog block. An id with no
' row is simply absent; what that MEANS is decided by AvailFor and depends on
' this map's provenance - see AVAIL_LEGACY_KEY. A v0.3.0 _Config (no Availability
' column at all) is flagged legacy here and still reads as everything on.
' An unrecognised cell is stored as "" so ComposeReport can block on it.
' Column layout comes from MarkerVerdict, not a decision made here - see its
' comment for why anything but "absent" reads the v1.2+ columns.
```

- [ ] **Step 3: `AvailFor` consults provenance**

Replace lines 2019-2025:

```vba
' Off unless the map says otherwise (v2.1.0). Absence used to mean Button; it now
' means Off, so a newly-registered action arrives switched off instead of
' appearing unannounced on an existing deployment.
'
' The ONE exception is a legacy map, flagged by ReadAvailMap because its stored
' block predates the Availability column. There, absence still means Button:
' the store described a deployment in which everything was live, and reading it
' as Off would silently strip every button from a working installation. Contract
' section 5 pins that default independently of section 6's for exactly this reason.
Private Function AvailFor(availMap As Object, ByVal id As String) As String
    AvailFor = AVAIL_OFF
    If availMap Is Nothing Then Exit Function
    If availMap.Exists(LCase(id)) Then
        AvailFor = CStr(availMap(LCase(id)))
    ElseIf availMap.Exists(AVAIL_LEGACY_KEY) Then
        AvailFor = AVAIL_BUTTON
    End If
End Function
```

- [ ] **Step 4: Build and syntax-check**

```sh
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -n 'SCHEMA_VER     As String' dist/MacrosBundle.bas
```

Expected: `Core + 2 package(s)`, balanced counts, `no diagnostics`, and **`SCHEMA_VER` still `"v2.0"`** — if it moved, something is wrong; ADR 0003 explicitly keeps it put so v0.5.0 configs stay readable.

- [ ] **Step 5: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core)!: Availability defaults to Off, legacy keeps Button

Contract v2.1.0. An action with no stored row now arrives Off, so a fresh
install shows a deliberate catalog and a newly-registered action does not
appear unannounced on an existing deployment.

Fresh and legacy are both an empty map and need opposite defaults, so
ReadAvailMap now flags a pre-v1.2 block under the reserved core.legacy key
rather than discarding what MarkerVerdict already knew. AvailFor's
signature is unchanged, so all twelve call sites are untouched - nothing
enumerates the map, every consumer is a keyed lookup.

SCHEMA_VER deliberately does NOT move: v0.5.0 configs stay readable and
the legacy path is left exactly as it was, verified by hand in the v0.5.0
gate."
```

---

### Task 4: Wire `usedBy` onto `PkgImportExport`

Without this Part A has no visible effect: no shipping package declares `usedBy`, so no field can ever relegate and the disclosure feature is invisible.

**Files:**
- Modify: `src/packages/PkgImportExport.bas:26-68`

**Interfaces:**
- Consumes: `AvailFor`'s new default from Task 3, via `FieldInUse`.
- Produces: nine fields bound to the actions that read them. Makes Part B's gate item 1.4 reachable for the first time.

**The mapping is verified, not guessed** — derived by extracting each action's `Core_Cfg*` calls:

| Field | usedBy | Why |
|---|---|---|
| `k_export_base` | `importexport.export` | read by `_Export` |
| `k_settings_sheet` | `importexport.export` | read by `_Export` |
| `k_export_offset` | `importexport.export` | read by `_Export` |
| `k_export_tabs` | `importexport.export` | read by `_Export` |
| `k_import` | `importexport.clear, importexport.import, importexport.suggestTypes` | read by `_Clear` and `_Import`; **`_SuggestTypes` reads AND writes it on the live Setup sheet** via `Core_SetupTable` / `Core_SetupSetColumn` |
| `k_master_prefix` | `importexport.updateMaster` | read by `_UpdateMaster` |
| `k_master_dest` | `importexport.updateMaster` | read by `_UpdateMaster` |
| `k_master_sheet` | `importexport.updateMaster` | read by `_UpdateMaster` |
| `k_master_cols` | `importexport.updateMaster` | read by `_UpdateMaster` |

> **`k_import` must name `suggestTypes` even though that action reads no config key.** It is a setup helper that writes the type map back into `k_import`'s rows on the Setup sheet. Omit it and turning off `clear` + `import` relegates `k_import` into the collapsed band while `Core_SetupSetColumn` still targets those rows.

- [ ] **Step 1: Add `usedBy` to the four export fields**

`usedBy` is an optional trailing named argument. Do not reorder the positional arguments.

```vba
    ' ---- GENERAL ----
    Core_AddSection reg, "GENERAL"
    Core_AddField reg, "k_export_base", "Name for the exported file (a date is added in front)", "text", "My Report", _
        help:="The saved file will be named like:  2026-01-15 <this text>.xlsx", _
        usedBy:="importexport.export"
    Core_AddField reg, "k_settings_sheet", "Settings / thresholds sheet name", "sheet", "README", "advanced", _
        "The hidden sheet whose cells drive the conditional-formatting thresholds; it is kept (hidden) inside the export so the colors still work.", _
        usedBy:="importexport.export"
    Core_AddField reg, "k_export_offset", "Date shift for the export file name (days)", "int", "-1", "advanced", _
        "How many days to shift the date put in front of the file name.  -1 = yesterday, 0 = today, 1 = tomorrow.", _
        usedBy:="importexport.export"
```

- [ ] **Step 2: Add `usedBy` to `k_export_tabs`**

```vba
    ' ---- EXPORT TABS ----
    Core_AddSection reg, "EXPORT TABS"
    Core_AddList reg, "k_export_tabs", "Tabs to include in the export (one per row, in order)", "Tab name", _
        help:="Each finished tab to copy into the exported .xlsx, top to bottom in this order.", _
        usedBy:="importexport.export"
```

- [ ] **Step 3: Add `usedBy` to `k_import` — all three actions**

Append the named argument to the existing `Core_AddTable reg, "k_import", ...` call:

```vba
        usedBy:="importexport.clear, importexport.import, importexport.suggestTypes"
```

`FieldInUse` splits on `,` and trims each id, so the spaces are fine.

- [ ] **Step 4: Add `usedBy` to the four REF_MASTER fields**

```vba
    ' ---- REF_MASTER ----
    Core_AddSection reg, "REF_MASTER"
    Core_AddField reg, "k_master_prefix", "Master-list file (name starts with...)", "file", "Master List Import", _
        help:="The newest file in THIS workbook's folder whose name starts with this text; used to refresh the destination tab below. A trailing (1) or (2) is ignored.", _
        usedBy:="importexport.updateMaster"
    Core_AddField reg, "k_master_dest", "Destination tab (in this workbook)", "sheet", "Ref_Master", _
        help:="The tab in THIS workbook that the master list is written into.", _
        usedBy:="importexport.updateMaster"
    Core_AddField reg, "k_master_sheet", "Sheet name inside the master file", "text", "Current Master", "advanced", _
        "The worksheet name to read from inside the master-list file.", _
        usedBy:="importexport.updateMaster"
```

Then append `usedBy:="importexport.updateMaster"` to the `Core_AddList reg, "k_master_cols", ...` call.

- [ ] **Step 5: Verify every id is real**

A typo in `usedBy` is silent by design — `FieldInUse` treats an unknown id as in-use so a typo can never hide a field forever. That makes typos invisible, so check them:

```sh
grep -o 'usedBy:="[^"]*"' src/packages/PkgImportExport.bas | grep -o 'importexport\.[a-zA-Z]*' | sort -u
grep -o 'Core_AddAction reg, "importexport\.[a-zA-Z]*"' src/packages/PkgImportExport.bas | grep -o 'importexport\.[a-zA-Z]*' | sort -u
```

Expected: the two lists are identical — `clear`, `export`, `import`, `suggestTypes`, `updateMaster`.

- [ ] **Step 6: Build and syntax-check**

```sh
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

- [ ] **Step 7: Commit**

```bash
git add src/packages/PkgImportExport.bas
git commit -m "feat(importexport): bind each field to the actions that read it

Without this Part A is invisible: no shipping package declared usedBy, so
no field could ever relegate and the relegation renderer was exercised
only by PkgFixture.

The mapping is derived from each action's Core_Cfg calls, not guessed.
k_import names suggestTypes even though that action reads no config key -
it is a setup helper that writes the type map back into k_import's rows
via Core_SetupSetColumn, so relegating those rows would break it."
```

---

### Task 5: Manual gate, docs, and release

**Files:**
- Create: `docs/designs/2026-07-27-part-a-manual-gate.md`
- Modify: `VERSION`, `CHANGELOG.md`, `docs/code-map.md`

**Interfaces:**
- Consumes: everything above.
- Produces: the checklist the owner runs; nothing depends on it in code.

- [ ] **Step 1: Rebuild the fixture workbooks**

Per the deliberation: rebuild rather than reuse. Workbook C was upgraded during Part B's gate, so if Finalize ran its `_Config` now carries a current marker and it would pass a legacy test **vacuously**. Make pristineness structural, not asserted.

```sh
git worktree add /tmp/wt-v030 cc59695
(cd /tmp/wt-v030 && pwsh build/build.ps1) && cp /tmp/wt-v030/dist/MacrosBundle.bas dist/MacrosBundle-v0.3.0.bas
git worktree remove /tmp/wt-v030 --force
grep -c 'SCHEMA_VER' dist/MacrosBundle-v0.3.0.bas   # expect 0
```

Build workbook C from that bundle per the Setup section of `docs/designs/2026-07-27-part-b-manual-gate.md`, **and save a pristine copy before upgrading it** so the next gate does not need a rebuild.

- [ ] **Step 2: Write the gate**

Create `docs/designs/2026-07-27-part-a-manual-gate.md` to the Part B standard (issue #3: every item names the workbook, the exact click path, and the exact expected string). It must cover:

1. **Fresh install shows everything Off.** New workbook, import, `Core_BuildSetup`. Every Availability cell reads `Off`. No buttons on the Macros pages.
2. **A v0.3.0 legacy workbook keeps every button.** Workbook C. **This is the §5-pin regression test — the single most important item.** If it fails, Task 1 Step 2 was wrong or was lost.
3. **A marked v0.5.0 config keeps every stored choice.** Build a workbook under v0.5.0, set a mix of Button / Chains only / Off, Finalize, then upgrade. Every choice survives verbatim.
4. **A newly-registered action arrives Off.** Hand-edit `_Config` to delete one action's row, then `Core_BuildSetup`. That action reads `Off`, not `Button`.
5. **A blank Availability cell blocks Finalize** with `MISSING or UNRECOGNISED Availability (use Button, Chains only, or Off)`.
6. **An unrecognised value is no longer laundered** (#10). Paste `maybe` into an Availability cell, Finalize (blocked), then `Core_BuildSetup`. The cell must NOT come back as `Button`.
7. **Part B's inherited item 1.4, now reachable.** Set every `importexport.*` action Off, `Core_BuildSetup`. The `IMPORT / EXPORT` band **disappears** and the off band **appears** with all nine fields. This is the first time this has ever been testable.
8. **Per-action disclosure.** Set only `importexport.export` to Off. Its four fields relegate; `k_import` and the REF_MASTER fields stay.
9. **`Suggest Import Types` still works with `clear` and `import` Off.** `k_import` must remain visible because `suggestTypes` is still on.
10. **Direct `Core_Run` on a fresh workbook.** Behaviour change worth confirming: with no `_Config`, every action is now Off, so `Core_Run` reports the action is off rather than running it.

- [ ] **Step 3: Refresh the code map**

```sh
node scripts/code-map.mjs
```

- [ ] **Step 4: VERSION and CHANGELOG**

Set `VERSION` to `0.6.0`. Add a dated `## [0.6.0]` section to `CHANGELOG.md` above the previous release, in the established voice — lead with what the user sees, then why, then the consequence.

- [ ] **Step 5: Build, syntax-check, commit**

```sh
pwsh build/build.ps1
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
git add -A
git commit -m "docs: Part A manual gate, code map, and v0.6.0"
```

- [ ] **Step 6: Open the PR**

Base `main`, head `feat/progressive-disclosure-part-a`. Body: what changed, the structural gate output, the manual gate table with results, and `Closes #6`, `Closes #10`. **Do not tick a gate item that was not actually run** — record it as untested instead. No attribution footer.

---

## Self-Review

**Spec coverage.** §2 sites 1 and 4 → Task 2; sites 2 and 3 → Task 3. §3 provenance flag → Task 3. §4 legacy untouched → satisfied by omission, with Task 1 Step 2 as its guard. §5 contract edits 1-4 → Task 1 Steps 1-5. §6 field disclosure → Task 4. §7 prerequisite → discharged, Part B's gate is 21/21. §9 verification → Task 5. **Full coverage.**

**Placeholders.** None. Every code step carries the actual VBA or shell.

**Type consistency.** `AVAIL_LEGACY_KEY` is declared once (Task 3 Step 1) and used in Tasks 3 Steps 2-3. `AvailFor`'s signature is unchanged throughout. `usedBy` ids in Task 4 match the `Core_AddAction` ids verified in Step 5.

**Known risk this plan does not remove.** Issue #8 — `Core_FinalizeSetup`'s activate is delete-then-copy, with a window in which no `_Config` exists. Part A adds Finalize cycles during gating but does not touch that path. Out of scope, filed, and Task 5 Step 1's pristine-copy habit is a partial mitigation.

## Handoff log

- **2026-07-27** — Plan written from the spec and ADR 0003, with four deliberated decisions folded in: `AvailLabel` fails closed rather than flipping to `Off` (Task 2, and it is a live bug, #10); the contract reservation ships in the same commit as nothing that uses it yet (Task 1); fixtures are rebuilt, not reused (Task 5 Step 1); the reskin is deferred entirely (issue #11). The `usedBy` mapping was derived from each action's `Core_Cfg` calls rather than guessed, which is what caught `suggestTypes`.
