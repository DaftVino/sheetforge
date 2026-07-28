# W1 — the ingest→clean wave: implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: use `superpowers:subagent-driven-development`
> to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship the seven frozen W1 actions (`import.*` ×3 in a new `PkgImport`,
`clean.*` ×4 added to `PkgClean`) behind a contract vocabulary and a build-time id
gate that make every later wave safe to write.

**Architecture:** Three layers, built bottom-up. `spec/contract.md` v2.3.0 gains the
nine-domain action-id vocabulary table it has always been assumed to have, behind a
machine-readable marker. `build/build.ps1` **parses that table** and fails a build whose
`Core_AddAction` ids are duplicated, malformed, off-vocabulary, in a reserved namespace,
or point at an entry point no source declares — so the contract genuinely owns the
vocabulary rather than merely describing it. Only then do the seven actions land: two
combines sharing one private `PkgImport_CombineEngine`, and four clean actions that reuse
Core's existing helpers rather than reimplementing them.

**Tech stack:** VBA (Excel/Windows, one concatenated `.bas`), PowerShell 5.1+ for the
build, Node for the syntax check. No test framework — see *Verification* below.

---

## Global constraints

Copied verbatim from `CLAUDE.md`, the design doc's do-not-revert list, and the
next-session prompt. Every task's requirements implicitly include this section.

- **`SCHEMA_VER` stays `"v2.0"`.** W1 adds actions; it reinterprets no stored value.
  `grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas` must print `1`.
- **Action ids are immutable**, two-segment `<domain>.<operation>`, lowercase ASCII and
  underscores, never camelCase, **no alias table**. The seven W1 ids are frozen in the
  design doc §3.4 and must appear in the bundle exactly as written there.
- **The domain is a user-facing vocabulary, NOT the module.** `import.*` ids live in
  `PkgImport` while `importexport.import` stays in `PkgImportExport`; both merge onto one
  Home card via `category:=`.
- **`Core_AddAction` already takes `category:=`** (shipped in track 1). Do not re-add it.
- **`Option Explicit` everywhere.** Max 25 line continuations per statement — one
  `Collection.Add` per line, never a giant `Array(...)`.
- **Every action wraps its work in the speed pattern** (`Core_SpeedOn` … `Core_SpeedOff`)
  with guaranteed restore in the error handler (`Core_Fail` calls `RestoreSpeed`).
- **Shared helpers live ONLY in Core.** Cross-package sharing is forbidden. New package
  privates are package-prefixed (`PkgImport_CombineEngine`, not `CombineEngine`) — design
  doc §3.3, because the bundle is one flat namespace.
- **Pure ASCII in every `.bas`.** `build.ps1` fails on a single non-ASCII byte: no em
  dashes, no smart quotes, no section signs, no arrows in VBA comments.
- **No module-level declaration after the first procedure** — a VBA compile error that
  `build.ps1` catches.
- **Do not revert** `ClearCatalogBlock`'s row-by-row merged-row-skipping clear, `CfgSheet`
  adopting an orphaned `_Config_staged`, `MarkerVerdict`'s `LooksPreV12` positive proof,
  `firstContentRow`'s assignment above the compose call, the semantic colours staying out
  of the identity palette, or the palette being five `Private Function`s after every
  module-level declaration.
- **Never read `src/core/Core.bas` whole** (229KB) and never read an `.xlsm`. Grep
  `docs/code-map.md` for an anchor, then read that slice.
- **No attribution footers** in any commit message or PR body.
- **One PR against `main`.** Do not stack PRs — the last session stranded 11 commits doing
  that, and #18 exists only to undo it.

---

## Decisions settled with the owner, 2026-07-28

The next-session prompt flagged two open questions in `## Unknowns and risks` as real plan
input. Both are settled below, along with two destructive-path decisions the library
snippets force. **These are inputs to the tasks, not open questions — do not re-litigate
them during implementation.**

| # | Question | Decision |
|---|---|---|
| 1 | How much of the three import actions' hardcoded paths/names becomes Setup config? | **Minimal band, 4 fields.** `PkgImport` declares a combine destination name, a source folder, a delimiter and a text qualifier. Everything describing the selection — which file, which sheets — is asked at run time. This is exactly what contract §8's normative deployment-vs-selection rule dictates. |
| 2 | Where does `clean.standardize_formats`' column→format map live? | **A `table` config field, `k_clean_formats`.** The same columns get the same formats every close, so the map is a deployment fact; and a format string like `$#,##0.00` is not something to retype at a prompt. |
| 3 | What safety does `clean.remove_blanks` ship with, given `backup.snapshot` is four waves away? | **Count-and-confirm.** Scan first, report exactly what will go, require an explicit Yes. This is the preview its deferred siblings lack. |
| 4 | What do the combines do when the destination sheet already exists? | **Refuse and name it.** Never delete a user's sheet. Matches the fail-closed posture Core took in v0.7.0. |

**Decision 4 has a consequence worth stating once:** the library snippet
`CombineAllSheets` opens with `ThisWorkbook.Worksheets("Master").Delete` under
`On Error Resume Next`. That line is the single most destructive thing in the W1 source
material and it must not survive conversion in any form.

---

## Deliberately deferred, and why

State these in the PR body so they are not read as oversights.

- **The legacy `importexport.*` category split.** Design doc §3.2 describes fixing
  `importexport.import` → "Bring Data In" and `importexport.export` → "Get It Out". W1's
  three new actions instead declare `category:="Import / Export"` so they merge onto the
  card that exists **today**, which is what §3.2's "merges onto the same Home card as the
  legacy import action" means in practice. Renaming the two legacy categories is a
  one-line presentation change (additive under §10's two-axes rule) that touches actions
  outside the frozen wave; it belongs to W2 or later, not here.
- **The 34-vs-36 action count.** §3.6 says W1–W6 assign 34 items producing 35 actions;
  grepping the catalog for wave tags yields 36. Unreconciled, not urgent for W1 — whose
  seven are unambiguous. Resolve before W2.
- **`.superpowers/sdd/`** still needs deleting by hand; the permission classifier refused
  twice. Nothing depends on it.
- **`.vscode/` stays untracked**, deliberately, across many sessions. Do not commit or
  remove it.

---

## Verification, since TDD is off

`CLAUDE.md` turns off `test-driven-development`: VBA has no headless runner and Excel
cannot be driven on CI. **Do not propose a test suite.** Every task's verification step
below is one of the repo's four real gates, in ascending cost:

1. `pwsh build/build.ps1` — the closest thing to CI. Structural checks plus, after Task 2,
   the id checks.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` — must report no diagnostics.
3. **A deliberate-injection test** for anything that is supposed to *fail*. A gate nobody
   has watched fail is not a gate.
4. The owner importing the bundle and clicking every button, against a checklist that has
   passed a blind dry-run.

---

## File structure

| File | Status | Responsibility after W1 |
|---|---|---|
| `spec/contract.md` | modify | v2.3.0. §7 gains the marked, machine-readable domain vocabulary table; §8 gains a `import` package section and four `clean` actions plus one `clean` field; §10 gains the v2.3.0 entry. |
| `build/build.ps1` | modify | Gains one guard block: parse the vocabulary from the contract, then check every `Core_AddAction` id for shape, domain, reserved namespace, uniqueness, and a resolvable entry point. |
| `src/packages/PkgImport.bas` | **create** | Manifest (4 fields, 3 actions), three public entry points, one private `PkgImport_CombineEngine` plus its helpers. |
| `src/packages/PkgClean.bas` | modify | +1 field, +4 actions, +their private helpers. Existing three actions untouched. |
| `src/core/Core.bas` | modify | One line: `c.Add "Import"` in `CoreCandidates()`. Nothing else. |
| `tools/GateAudit.bas` | modify | One new probe: every declared action id reaches the Setup Action catalog block. |
| `docs/designs/2026-07-28-w1-manual-gate.md` | **create** | The owner's click-every-button checklist. |
| `CHANGELOG.md`, `VERSION` | modify | v0.10.0. |

Task order is the design doc's, because the first two items are what make the other five
safe to write.

---

### Task 1: Contract v2.3.0 — the vocabulary table and the seven ids

**Owner: Opus, written directly.** Not delegated. Every later wave reads this table, and
`build.ps1` parses it byte-for-byte in Task 2 — a formatting slip here disarms the gate
that protects every future id.

**Files:**
- Modify: `spec/contract.md` — header (line 3), §7 (from line 550), §8 (from line 602), §10 (from line 712)

**Interfaces:**
- Consumes: nothing.
- Produces: the marker line `<!-- action-domain-vocabulary -->` immediately preceding a
  three-column pipe table whose first column is a backticked domain token and whose second
  column is exactly one of `active`, `reserved`, `grandfathered`. Task 2's parser depends
  on this shape and on nothing else about §7.

- [ ] **Step 1: Bump the contract version header**

`spec/contract.md` line 3 currently reads:

```markdown
**Contract version: 2.2.0** (2026-07-27) · SemVer, independent of the
```

Change to:

```markdown
**Contract version: 2.3.0** (2026-07-28) · SemVer, independent of the
```

- [ ] **Step 2: Add the vocabulary table to §7**

In §7 "Naming rules (contract-level)", the bullet currently reading

```markdown
- **Action id** — `pkg.verb`, unique.
```

is replaced by the following. Note that the old text said `pkg.verb`, which contradicted
the design doc's frozen rule that the first segment is a **user-facing domain, not the
package** — this correction is the whole reason §7 needed touching.

````markdown
- **Action id** — `<domain>.<operation>`, unique, lowercase ASCII letters, digits and
  underscores, exactly two segments. **The domain is a controlled user-facing vocabulary,
  NOT the package or module that implements it** — encoding the module would make a future
  package split a contract break for no benefit. The operation segment is descriptive
  where ambiguity looms (`clean.remove_blank_rows`, never `clean.blank2`). **No new id is
  ever camelCase.** Ids are immutable; captions, help text, entry points and module
  assignment stay free to change. There is no alias or migration table, by design: an
  alias means two permanent names for one action and would have to be ported to Apps
  Script.

  **The vocabulary is the table below, and it is normative.** The marker comment
  immediately above it is a parse anchor — `build/build.ps1` reads the table from this
  file and fails a build whose action ids use a domain that is not listed `active` or
  `grandfathered`. Do not reformat the table, rename its columns, or remove the marker
  without changing that parser in the same commit.

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
| `chain` | reserved | Engine chains (§6). No package action id may use it. |
| `core` | reserved | Engine-internal sentinels (v2.1.0, this section). No package action id may use it. |
| `importexport` | grandfathered | The pre-v1.2.0 package-named domain. Its five shipped ids are permanent; **no new id may be added to it.** It is direct evidence for the domain-is-not-the-module rule above. |

**The eight shipped ids are permanent exceptions**, listed literally, never renamed and
never aliased: `clean.dupes`, `clean.trim`, `clean.filldown`, `importexport.clear`,
`importexport.import`, `importexport.updateMaster`, `importexport.export`,
`importexport.suggestTypes`. Two are camelCase and the `importexport` domain conflates two
different jobs; both stay exactly as they are, and the rules above exist because of them.

**Enforcement, in two places.** The engine rejects a duplicate id during registry
assembly, which protects a hand-edited bundle. The build gate extracts literal action-id
declarations from source and fails on a duplicate, a malformed id, a reserved domain, an
unknown domain, or an entry point no source declares. A platform without a build step
performs the equivalent checks at manifest-assembly time.
````

The existing `chain.*` and `core.*` reservation bullets in §7 **stay exactly as they are** —
they carry the v2.1.0/v2.2.0 normative language about `core.legacy` and fail-closed
activation, which the table's one-line summary does not replace.

- [ ] **Step 3: Add the `import` package to §8**

Insert immediately before the `### Package \`clean\`` heading (currently line 649):

````markdown
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
| `k_import_qualifier` | text | advanced | `"` | The character wrapping text that itself contains the separator. `"` (double quote), `'` (single quote), or the token `none`. Blank reads as `none`. `usedBy` `import.delimited_file`. |

**Actions**

| id | caption | semantics |
|---|---|---|
| `import.combine_worksheets` | Combine All Worksheets | Stack every data sheet in this workbook into one new sheet named `k_import_dest`, **aligning columns by header text, not by position**: a source's `Amount` column lands under the master `Amount` wherever it sits, and a header no earlier source declared is appended as a new column on the right. Row 1 of each source is its header and is not repeated. Engine-owned sheets (the Setup surface, the config store and its staging sheet, the package-enable sheet, and every generated button page) MUST be excluded. If `k_import_dest` already names an existing sheet the action MUST stop having changed nothing. A failure part-way MUST leave no partially built destination sheet behind. |
| `import.combine_workbooks` | Combine All Workbooks in a Folder | As `import.combine_worksheets`, but the sources are the first worksheet of every workbook in `k_import_folder`, excluding this workbook itself. Each source is opened read-only and closed without saving, one at a time. Same header alignment, same refusal, same partial-failure cleanup. |
| `import.delimited_file` | Import a Delimited File | Prompt for one delimited file; parse it with the explicit separator and text qualifier from config rather than letting the platform guess; write it to a new sheet named after the file. Same refusal rule as the combines if that sheet name is taken. The prompt MUST offer a cancel that abandons the action having changed nothing. |
````

- [ ] **Step 4: Extend the `clean` package in §8**

The `**Fields**` block under `### Package \`clean\`` currently reads
`**None (v2.0.0).** …`. Replace **only its first paragraph** — the two paragraphs that
follow, explaining why `clean_keycols`/`clean_hasheader` were withdrawn and stating the
normative deployment-vs-selection rule, **stay verbatim**. New first paragraph:

````markdown
**One field (v2.3.0).** `clean` declared no configuration fields at all from v2.0.0 to
v2.2.0. It now declares exactly one:

| key | kind | tier | default | semantics |
|---|---|---|---|---|
| `k_clean_formats` | table | simple | see below | Column-to-number-format map for `clean.standardize_formats`. Two columns: **Column** (a column letter, or a header name matched against the first row of the range the caller picks) and **Number format** (a platform number-format string). A row whose Column resolves to nothing inside the picked range is reported, not silently skipped. `usedBy` `clean.standardize_formats`. |

`k_clean_formats` header: `Column (letter or header name)|Number format`.
Demo default rows (`|`-delimited): `A|mm/dd/yyyy` · `B|0` · `C|$#,##0.00` · `D|0.00%`.

**This does not reopen the rule below, it satisfies it.** A house number format is a fact
about the deployment — the same columns get the same formats every close — not a fact
about the range the caller is about to pick, which is precisely the distinction the
withdrawn fields failed. Resolution by **column letter or header name**, never by position
within the selection, is what keeps it on the deployment side of that line: a position
counted from the left of a selection that does not exist yet would be the withdrawn
`clean_keycols` under a new name.
````

Then append four rows to the `clean` **Actions** table:

````markdown
| `clean.remove_blanks` | Remove Blank Rows & Columns | Prompt for a range; find every entirely blank row and column inside it. **Delete only those that are also blank across the whole sheet's used extent**, so deleting a whole row can never destroy data sitting outside the caller's selection; report the ones held back and why. Before deleting anything, show a count of exactly what will go and require an explicit confirmation; declining or cancelling MUST change nothing. The completion report MUST state both counts, because neither is recoverable from the result. |
| `clean.unmerge_fill` | Unmerge & Fill Down | Prompt for a range; unmerge every merged block inside it and write the block's value into every cell it covered. A merge anchor holding a formula MUST keep its formula; the cells it covered receive the computed value. |
| `clean.standardize_formats` | Standardize Formats | Prompt for a range; apply each `k_clean_formats` row's number format to that row's column **within the picked range only**. Changes formatting, never a stored value. A mapping row that resolves to no column in the picked range MUST be reported by name, not silently dropped. With no rows configured, the action MUST stop and say where to configure them. |
| `clean.text_to_numbers` | Convert Text to Numbers | Prompt for a range; convert text that represents a number into a real number, using **the same coercion the import action uses** (§8, column-type code `N`) rather than a second implementation. Formula cells are never touched. Text the coercion cannot parse is left exactly as it is. The completion report MUST state how many cells changed. **Known limitation, stated because it is visible to the user:** the `N` coercion strips a trailing `%`, so `12%` becomes the number `12`, not `0.12`. |
````

- [ ] **Step 5: Add the v2.3.0 entry to §10**

Insert as the **first** versioned bullet in §10 (above the `v2.2.0` bullet):

````markdown
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
````

- [ ] **Step 6: Verify the table parses the way Task 2 will parse it**

```bash
grep -n -A 16 'action-domain-vocabulary' spec/contract.md
```

Expected: the marker on its own line, a blank line, the header row, the separator row, and
exactly twelve data rows — nine `active`, two `reserved`, one `grandfathered`. Confirm
every first cell is a backticked lowercase token and every second cell is one of those
three words, lowercase, with no trailing punctuation.

```bash
grep -c '^| `[a-z][a-z0-9_]*` | \(active\|reserved\|grandfathered\) |' spec/contract.md
```

Expected: `12`.

- [ ] **Step 7: Commit**

```bash
git add spec/contract.md
git commit -m "spec: contract v2.3.0 - action-id domain vocabulary, package import, four clean actions"
```

---

### Task 2: `build.ps1` — the id gate

**Owner: Opus, written directly.** Not delegated. Ids are immutable with no alias table by
design, so a wrong one is unfixable; this is the only thing standing between a wave and a
permanent mistake. It must be watched failing before it is trusted.

**Files:**
- Modify: `build/build.ps1` — insert one guard block between the reserved-word guard
  (ends line 195) and the duplicate-procedure guard (begins line 197)

**Interfaces:**
- Consumes: Task 1's `<!-- action-domain-vocabulary -->` marker and table shape.
- Produces: a build that fails with a named error on a bad action id. Tasks 3 and 4 rely on
  it passing for their new ids.

- [ ] **Step 1: Insert the guard block**

Insert after line 195 (`}` closing the `$reservedProblems` check) and before the
`# --- guard: every Sub/Function/Property name must be unique …` comment:

```powershell
# --- guard: action ids - vocabulary, shape, uniqueness, resolvable entry point ---
# Ids are immutable, two-segment <domain>.<operation>, and there is deliberately NO
# alias table (spec/contract.md section 7) - so a wrong id is unfixable and has to be
# caught before it ships, not after. Nothing else here can see it: the id is a string
# literal, valid VBA to every other check in this file and to vba-syntax-check.mjs.
#
# The vocabulary is READ FROM THE CONTRACT, not duplicated here. Section 7 says the
# contract owns that table; a hardcoded copy would make this file a second owner, and
# the two would drift the first time a wave added a domain. The parse is anchored on an
# HTML-comment marker rather than on the heading text, so re-wording section 7 cannot
# silently disarm this.
#
# It FAILS CLOSED. An unreadable or empty table is a build failure, never a fallback to
# some built-in list: a gate that quietly disables itself when its input goes missing is
# worse than no gate, because the build still goes green.
#
# Scope is every candidate source (Core + every Pkg*.bas on disk), NOT only the
# -Packages subset this run selected - the same repo-wide-invariant reasoning as the
# ASCII guard above. A bad id sitting in an unbundled package ships the day someone
# bundles it, and the point of a gate is to catch it before that day.
$contractPath = Join-Path $Root 'spec/contract.md'
if (-not (Test-Path -LiteralPath $contractPath)) {
    Write-Host "BUILD FAILED - action id validation:" -ForegroundColor Red
    Write-Host "  contract not found: $contractPath" -ForegroundColor Red
    Write-Host "  The action-id domain vocabulary is read from that file; without it no id can be checked." -ForegroundColor Red
    exit 1
}
$contractLines = Get-Content -LiteralPath $contractPath
$vocabMarker = '<!-- action-domain-vocabulary -->'
$markerAt = -1
for ($i = 0; $i -lt $contractLines.Count; $i++) {
    if ($contractLines[$i].Trim() -eq $vocabMarker) { $markerAt = $i; break }
}
if ($markerAt -lt 0) {
    Write-Host "BUILD FAILED - action id validation:" -ForegroundColor Red
    Write-Host "  marker '$vocabMarker' not found in spec/contract.md." -ForegroundColor Red
    Write-Host "  Section 7 must carry that marker on its own line directly above the domain table." -ForegroundColor Red
    exit 1
}
# Read table rows following the marker; stop at the first non-table line AFTER at least
# one row has been read (the blank line and the header/separator rows sit in between).
$domainStatus = @{}
for ($i = $markerAt + 1; $i -lt $contractLines.Count; $i++) {
    $ln = $contractLines[$i]
    if ($ln -notmatch '^\s*\|') {
        if ($domainStatus.Count -gt 0) { break } else { continue }
    }
    if ($ln -match '^\s*\|[\s\-:|]+\|\s*$') { continue }          # separator row
    if ($ln -match '^\s*\|\s*`([a-z][a-z0-9_]*)`\s*\|\s*(active|reserved|grandfathered)\s*\|') {
        $domainStatus[$matches[1]] = $matches[2]
    }
}
if ($domainStatus.Count -eq 0) {
    Write-Host "BUILD FAILED - action id validation:" -ForegroundColor Red
    Write-Host "  the domain table under '$vocabMarker' in spec/contract.md parsed to zero domains." -ForegroundColor Red
    Write-Host "  Each row must read: | ``<domain>`` | active|reserved|grandfathered | <meaning> |" -ForegroundColor Red
    exit 1
}

# The eight ids shipped before the shape rules existed. Listed literally in contract
# section 7 as permanent exceptions, never renamed and never aliased - two are camelCase
# and would fail the shape check below. This list does NOT exempt them from the
# duplicate or entry-point checks.
$legacyIds = @(
    'clean.dupes', 'clean.trim', 'clean.filldown',
    'importexport.clear', 'importexport.import', 'importexport.updateMaster',
    'importexport.export', 'importexport.suggestTypes'
)
$legacySet = @{}; $legacyIds | ForEach-Object { $legacySet[$_] = $true }

# Every Sub name declared anywhere in the candidate sources, for entry-point resolution.
$declaredSubs = @{}
foreach ($f in $asciiCandidates) {
    foreach ($line in (Get-Content -LiteralPath $f)) {
        $m = [regex]::Match($line, '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Sub\s+([A-Za-z_]\w*)')
        if ($m.Success) { $declaredSubs[$m.Groups[1].Value] = Split-Path $f -Leaf }
    }
}

# Core_AddAction reg, "<id>", "<caption>", "<entryPoint>", "<help>" [, "<category>"]
$addActionRe = 'Core_AddAction\s+\w+\s*,\s*"([^"]*)"\s*,\s*"[^"]*"\s*,\s*"([^"]*)"'
$idShapeRe = '^[a-z][a-z0-9_]*\.[a-z][a-z0-9_]*$'
$allowed = ($domainStatus.Keys | Where-Object { $domainStatus[$_] -ne 'reserved' } | Sort-Object) -join ', '
$idSeen = @{}
$idProblems = New-Object System.Collections.Generic.List[string]
foreach ($f in $asciiCandidates) {
    $leaf = Split-Path $f -Leaf
    $srcLines = Get-Content -LiteralPath $f
    for ($i = 0; $i -lt $srcLines.Count; $i++) {
        if ($srcLines[$i] -match "^\s*'") { continue }             # commented-out example
        $m = [regex]::Match($srcLines[$i], $addActionRe)
        if (-not $m.Success) { continue }
        $id = $m.Groups[1].Value
        $entry = $m.Groups[2].Value
        $where = "${leaf}:$($i + 1)"

        if ($idSeen.ContainsKey($id)) {
            $idProblems.Add("duplicate action id '$id' at $where - already declared at $($idSeen[$id]). Ids are unique across every package.")
        } else {
            $idSeen[$id] = $where
        }

        if (-not $legacySet.ContainsKey($id)) {
            if ($id -notmatch $idShapeRe) {
                $idProblems.Add("malformed action id '$id' at $where - must be exactly two segments, <domain>.<operation>, lowercase letters/digits/underscores, never camelCase.")
            }
        }

        $domain = if ($id -match '^([^.]+)\.') { $matches[1] } else { $id }
        if (-not $domainStatus.ContainsKey($domain)) {
            $idProblems.Add("unknown domain '$domain' in action id '$id' at $where - not in the vocabulary table in spec/contract.md section 7. Allowed: $allowed.")
        } elseif ($domainStatus[$domain] -eq 'reserved') {
            $idProblems.Add("reserved domain '$domain' in action id '$id' at $where - '$domain.*' is reserved for the engine in spec/contract.md section 7 and no package action may use it.")
        }

        if ($entry -eq '') {
            $idProblems.Add("action '$id' at $where declares an empty entry point.")
        } elseif (-not $declaredSubs.ContainsKey($entry)) {
            $idProblems.Add("action '$id' at $where names entry point '$entry', which no source file declares as a Sub. Core_Run would fail to dispatch it.")
        }
    }
}
if ($idProblems.Count -gt 0) {
    Write-Host "BUILD FAILED - action id validation:" -ForegroundColor Red
    $idProblems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}
```

- [ ] **Step 2: Verify it passes on the untouched tree**

```bash
pwsh build/build.ps1
```

Expected: exit 0, `Core + 2 package(s)`. If it fails, the `$legacyIds` list or the
`$addActionRe` groups are wrong — the eight shipped ids must all pass. Do not "fix" this by
widening the shape regex; widen the exception list only if the id is one of the eight.

- [ ] **Step 3: Watch it fail — duplicate id**

Temporarily append to `PkgClean_Manifest` in `src/packages/PkgClean.bas`, on the line after
the existing `clean.filldown` declaration:

```vb
    Core_AddAction reg, "clean.trim", "Injected Duplicate", "PkgClean_TrimClean", "Deliberate build-gate injection - revert."
```

```bash
pwsh build/build.ps1; echo "exit=$?"
```

Expected: `exit=1`, and a red line naming `duplicate action id 'clean.trim'` with both
locations. Then revert:

```bash
git checkout -- src/packages/PkgClean.bas
```

- [ ] **Step 4: Watch it fail — off-vocabulary domain**

Same file, same place:

```vb
    Core_AddAction reg, "tidy.everything", "Injected Bad Domain", "PkgClean_TrimClean", "Deliberate build-gate injection - revert."
```

```bash
pwsh build/build.ps1; echo "exit=$?"
```

Expected: `exit=1`, and a red line naming `unknown domain 'tidy'` and listing the allowed
domains. Then revert:

```bash
git checkout -- src/packages/PkgClean.bas
```

- [ ] **Step 5: Watch it fail — reserved namespace and missing entry point**

The reserved-domain path and the entry-point path are separate branches and neither is
exercised by Steps 3–4. Inject both at once:

```vb
    Core_AddAction reg, "core.legacy", "Injected Reserved", "PkgClean_NoSuchProc", "Deliberate build-gate injection - revert."
```

```bash
pwsh build/build.ps1; echo "exit=$?"
```

Expected: `exit=1`, **two** red lines — `reserved domain 'core'` and
`names entry point 'PkgClean_NoSuchProc', which no source file declares as a Sub`. Then
revert:

```bash
git checkout -- src/packages/PkgClean.bas
```

- [ ] **Step 6: Watch it fail closed — vocabulary unreadable**

The fail-closed claim is the one most likely to be wrong and the least likely to be
noticed, because it only matters on a day nobody is looking.

Move the marker, build, then restore from git rather than from a copy — git is the reliable
restore here and it needs no scratch file at all.

```bash
sed -i 's/<!-- action-domain-vocabulary -->/<!-- moved -->/' spec/contract.md
pwsh build/build.ps1; echo "exit=$?"
git checkout -- spec/contract.md
```

Expected: `exit=1` with `marker '<!-- action-domain-vocabulary -->' not found`, **not** a
green build. Note this step only works after Task 1 is committed; if it is not, stash
instead of checking out.

- [ ] **Step 7: Confirm the tree is clean and commit**

```bash
git status --short
```

Expected: only `build/build.ps1` modified, plus the untracked `.vscode/`. If
`src/packages/PkgClean.bas` or `spec/contract.md` still shows as modified, an injection was
not reverted — revert it before committing.

```bash
git add build/build.ps1
git commit -m "build: fail on a duplicate, malformed, reserved or off-vocabulary action id"
```

---

### Task 3: `src/packages/PkgImport.bas` — the three import actions

**Owner: delegate to codex (`terra` or `sonnet`).** Mechanical conversion against a
complete spec. Hand the subagent this task in full plus `src/packages/PkgClean.bas` as the
pattern; it must not need to read `Core.bas`.

**Files:**
- Create: `src/packages/PkgImport.bas`
- Modify: `src/core/Core.bas` — `CoreCandidates()` at line 952

**Interfaces:**
- Consumes: `Core_DeclarePackage`, `Core_AddSection`, `Core_AddField`, `Core_AddAction`,
  `Core_Cfg`, `Core_SpeedOn`, `Core_SpeedOff`, `Core_Done`, `Core_Fail`, `Core_SheetExists`,
  `Core_LastRowUp`, `Core_LastCol`, `Core_FileName`. Exact signatures below.
- Produces: three public entry points named in the manifest —
  `PkgImport_CombineWorksheets`, `PkgImport_CombineWorkbooks`, `PkgImport_DelimitedFile`.
  Task 6's gate probe reads their three ids off the Setup catalog.

**Core API this task uses — do not guess these, and do not add helpers to Core:**

```vb
Core_DeclarePackage(reg As Object, id As String, name As String, version As String)
Core_AddSection(reg As Object, title As String)
Core_AddField(reg As Object, key As String, label As String, kind As String, dflt As String, _
              Optional tier As String = "simple", Optional help As String = "", _
              Optional usedBy As String = "")
Core_AddAction(reg As Object, actionId As String, caption As String, entryPoint As String, _
               help As String, Optional category As String = "")
Core_Cfg(ctx As Object, key As String) As String        ' blank cell -> manifest default
Core_SpeedOn() / Core_SpeedOff()                        ' depth-counted; SpeedOff restores
Core_Done(ctx As Object, msg As String, Optional title As String = "Done")
Core_Fail(ctx As Object, actionId As String, e As ErrObject)   ' restores speed itself
Core_SheetExists(nm As String) As Boolean
Core_LastRowUp(ws As Worksheet, Optional col As Long = 1) As Long   ' bare End(xlUp)
Core_LastCol(ws As Worksheet, Optional row As Long = 1) As Long
Core_FileName(p As String) As String                    ' path -> file name with extension
```

- [ ] **Step 1: Register the package with Core**

`src/core/Core.bas`, `CoreCandidates()` at line 952, currently:

```vb
Private Function CoreCandidates() As Collection
    Dim c As New Collection
    c.Add "ImportExport"            ' probes PkgImportExport_Manifest
    c.Add "Clean"                   ' probes PkgClean_Manifest
    Set CoreCandidates = c
End Function
```

Add one line, keeping one `c.Add` per line (the 25-continuation rule):

```vb
Private Function CoreCandidates() As Collection
    Dim c As New Collection
    c.Add "ImportExport"            ' probes PkgImportExport_Manifest
    c.Add "Import"                  ' probes PkgImport_Manifest
    c.Add "Clean"                   ' probes PkgClean_Manifest
    Set CoreCandidates = c
End Function
```

This is the **only** change to `Core.bas` in the entire wave. Without it the module builds
into the bundle and is never probed, so its buttons never appear — and nothing else in the
build reports that.

- [ ] **Step 2: Create the module header and manifest**

Create `src/packages/PkgImport.bas` beginning with:

```vb
Attribute VB_Name = "PkgImport"
Option Explicit

'================================================================
' PkgImport - bringing outside data in. Wave 1 of the catalog
' build-out (docs/designs/2026-07-24-finish-excel-design.md 3.4).
'
' Three actions, two of which are the same job with a different
' source list, so both delegate to ONE private engine:
'   import.combine_worksheets   every data sheet in THIS workbook
'   import.combine_workbooks    the first sheet of every workbook in a folder
'   import.delimited_file       one CSV/TXT, parsed explicitly
'
' The domain "import" is a USER-FACING vocabulary, not this module's
' name. The legacy importexport.import action keeps its id and stays
' in PkgImportExport; both declare the category "Import / Export" so
' the user sees one card. See contract section 7.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgImport_Manifest(reg As Object)
    Core_DeclarePackage reg, "import", "Bring Data In", "1.0.0"

    Core_AddSection reg, "COMBINE"
    Core_AddField reg, "k_import_dest", "Name for the combined sheet this creates", "text", "Master", _
        "simple", _
        "A NEW sheet by this name is created each run. If a sheet by this name already exists, the action stops and asks you to rename or remove it - it never overwrites one.", _
        "import.combine_worksheets,import.combine_workbooks"
    Core_AddField reg, "k_import_folder", "Folder to read workbooks from", "text", "", _
        "simple", _
        "Leave blank to use the folder this workbook is saved in - that is the usual answer. Otherwise a full path, e.g. C:\Close\January.", _
        "import.combine_workbooks"

    Core_AddSection reg, "DELIMITED FILES"
    Core_AddField reg, "k_import_delim", "Column separator in CSV / TXT files", "text", "comma", _
        "advanced", _
        "One of: comma, tab, semicolon, pipe, space. Or type the separator character itself. Blank means comma.", _
        "import.delimited_file"
    Core_AddField reg, "k_import_qualifier", "Quote character around text", "text", """", _
        "advanced", _
        "The character wrapping text that itself contains the separator. Type a double quote, a single quote, or the word none. Blank means none.", _
        "import.delimited_file"

    Core_AddAction reg, "import.combine_worksheets", "Combine All Worksheets", "PkgImport_CombineWorksheets", _
        "Stacks every data sheet in this workbook into one new sheet, lining up columns by their header names. Your Setup and Macros tabs are left out.", _
        "Import / Export"
    Core_AddAction reg, "import.combine_workbooks", "Combine All Workbooks in a Folder", "PkgImport_CombineWorkbooks", _
        "Stacks the first sheet of every workbook in the folder into one new sheet, lining up columns by their header names. Each file is opened read-only and closed without saving.", _
        "Import / Export"
    Core_AddAction reg, "import.delimited_file", "Import a Delimited File", "PkgImport_DelimitedFile", _
        "Pick a CSV or TXT file; it is parsed with the separator and quote character from Setup rather than letting Excel guess, and written to a new sheet named after the file.", _
        "Import / Export"
End Sub
```

`"""" ` is a one-character VBA string holding a double quote. That is the intended default
for `k_import_qualifier` — do not "fix" it to an empty string.

- [ ] **Step 3: Write the two combine entry points**

```vb
'================================================================
' ACTION - Combine every data sheet in this workbook
'================================================================
Public Sub PkgImport_CombineWorksheets(ctx As Object)
    Dim destName As String, sources As Collection, ws As Worksheet
    Dim rowsOut As Long, sourcesOut As Long
    Const TITLE As String = "Combine All Worksheets"
    On Error GoTo Fail

    ' The destination check runs FIRST, before any prompt or any work: the
    ' cheapest failure should be the earliest one, and a refusal after the user
    ' has answered questions reads as a bug.
    destName = Trim(Core_Cfg(ctx, "k_import_dest"))
    If Not PkgImport_DestUsable(destName, TITLE) Then Exit Sub

    Set sources = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If Not PkgImport_IsEngineSheet(ws.name) Then
            If StrComp(ws.name, destName, vbTextCompare) <> 0 Then sources.Add ws
        End If
    Next ws
    If sources.Count = 0 Then
        MsgBox "There are no data sheets to combine in this workbook." & vbCrLf & vbCrLf & _
               "Every sheet here belongs to the macro toolkit itself." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgImport_CombineEngine sources, destName, rowsOut, sourcesOut
    Core_SpeedOff

    Core_Done ctx, "Combined " & sourcesOut & " sheet(s) into '" & destName & "'." & vbCrLf & _
           rowsOut & " data row(s) written." & vbCrLf & vbCrLf & _
           "Columns were lined up by their header names, so a column that sits in a " & _
           "different position on different sheets still landed in one column here.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "import.combine_worksheets", Err
End Sub

'================================================================
' ACTION - Combine the first sheet of every workbook in a folder
'================================================================
Public Sub PkgImport_CombineWorkbooks(ctx As Object)
    Dim destName As String, folderPath As String, sources As Collection
    Dim rowsOut As Long, sourcesOut As Long
    Const TITLE As String = "Combine All Workbooks in a Folder"
    On Error GoTo Fail

    destName = Trim(Core_Cfg(ctx, "k_import_dest"))
    If Not PkgImport_DestUsable(destName, TITLE) Then Exit Sub

    folderPath = PkgImport_ResolveFolder(Core_Cfg(ctx, "k_import_folder"))
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MsgBox "This folder was not found:" & vbCrLf & vbCrLf & "   " & folderPath & vbCrLf & vbCrLf & _
               "Check 'Folder to read workbooks from' on the Setup tab, Finalize, and run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    Set sources = PkgImport_WorkbookPaths(folderPath)
    If sources.Count = 0 Then
        MsgBox "No workbooks were found in:" & vbCrLf & vbCrLf & "   " & folderPath & vbCrLf & vbCrLf & _
               "This workbook itself is always skipped." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgImport_CombineEngine sources, destName, rowsOut, sourcesOut
    Core_SpeedOff

    Core_Done ctx, "Combined " & sourcesOut & " workbook(s) into '" & destName & "'." & vbCrLf & _
           rowsOut & " data row(s) written, read from:" & vbCrLf & "   " & folderPath & vbCrLf & vbCrLf & _
           "Columns were lined up by their header names. Every file was opened read-only " & _
           "and closed without saving.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "import.combine_workbooks", Err
End Sub
```

- [ ] **Step 4: Write the shared engine**

This is the piece design doc §3.4 names. A `sources` member is either a `Worksheet`
(already open, in this workbook) or a `String` path the engine opens and closes itself —
that difference **is** the difference between the two combines, which is why there is one
engine and not two.

```vb
'================================================================
' The shared combine. sources holds either Worksheet objects (open, in
' this workbook) or String paths to workbooks this engine opens and
' closes one at a time. That is the ONLY difference between the two
' combine actions - design section 3.4.
'
' Columns are aligned BY HEADER TEXT, not by position: a source's
' "Amount" column lands under the master "Amount" wherever it sits, and
' a header no earlier source declared is appended as a new column on
' the right. Row 1 of each source is its header and is never repeated.
'
' On ANY error it deletes the half-built destination sheet and re-raises,
' so a failure leaves the workbook as it found it. That cleanup is the
' reason this has its own handler rather than letting the caller's
' Core_Fail see the error directly.
'================================================================
Private Sub PkgImport_CombineEngine(sources As Collection, ByVal destName As String, _
                                    ByRef rowsOut As Long, ByRef sourcesOut As Long)
    Dim dest As Worksheet, src As Worksheet, wbSrc As Workbook
    Dim one As Variant, headerMap As Object
    Dim destRow As Long, destCols As Long
    Dim eNum As Long, eDesc As String
    On Error GoTo EngineFail

    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = 1                    ' vbTextCompare: "Amount" matches "amount"
    Set dest = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    dest.name = destName
    destRow = 2                                  ' row 1 is the master header, grown as we learn it
    destCols = 0

    For Each one In sources
        Set src = Nothing
        Set wbSrc = Nothing
        If VarType(one) = vbString Then
            Set wbSrc = Workbooks.Open(Filename:=CStr(one), ReadOnly:=True, UpdateLinks:=0)
            Set src = wbSrc.Worksheets(1)
        Else
            Set src = one
        End If

        PkgImport_AppendSheet src, dest, headerMap, destRow, destCols
        sourcesOut = sourcesOut + 1

        If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Next one

    rowsOut = destRow - 2
    Exit Sub

EngineFail:
    eNum = Err.Number
    eDesc = Err.Description
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    ' DisplayAlerts is already False under Core_SpeedOn, so this does not prompt;
    ' Core_Fail's RestoreSpeed puts it back either way.
    If Not dest Is Nothing Then dest.Delete
    On Error GoTo 0
    Err.Raise eNum, "PkgImport_CombineEngine", eDesc
End Sub

' Append one source sheet under the growing master header.
' Writes ONE column block per source column rather than cell by cell - a
' cell-at-a-time loop over a real close file is minutes, not seconds, even
' with screen updating off.
Private Sub PkgImport_AppendSheet(src As Worksheet, dest As Worksheet, headerMap As Object, _
                                  ByRef destRow As Long, ByRef destCols As Long)
    Dim lastR As Long, lastC As Long, c As Long, r As Long, target As Long
    Dim hdr As String, vals As Variant, colVals() As Variant

    lastR = Core_LastRowUp(src, 1)
    lastC = Core_LastCol(src, 1)
    If lastR < 2 Or lastC < 1 Then Exit Sub          ' empty, or a header with no data

    vals = src.Range(src.Cells(1, 1), src.Cells(lastR, lastC)).Value

    ReDim colVals(1 To lastR - 1, 1 To 1)
    For c = 1 To lastC
        hdr = Trim(CStr(vals(1, c)))
        If Len(hdr) = 0 Then hdr = "Column " & c     ' an unnamed column still needs a home
        If Not headerMap.Exists(hdr) Then
            destCols = destCols + 1
            headerMap.Add hdr, destCols
            dest.Cells(1, destCols).Value = hdr
            dest.Cells(1, destCols).Font.Bold = True
        End If
        target = headerMap(hdr)
        For r = 2 To lastR
            colVals(r - 1, 1) = vals(r, c)
        Next r
        dest.Cells(destRow, target).Resize(lastR - 1, 1).Value = colVals
    Next c
    destRow = destRow + (lastR - 1)
End Sub
```

- [ ] **Step 5: Write the delimited-file action**

```vb
'================================================================
' ACTION - Import one delimited file with explicit parsing
'================================================================
Public Sub PkgImport_DelimitedFile(ctx As Object)
    Dim picked As Variant, destName As String, dest As Worksheet, wbTxt As Workbook
    Dim rowsOut As Long, colsOut As Long
    Const TITLE As String = "Import a Delimited File"
    On Error GoTo Fail

    ' Application.GetOpenFilename returns Boolean False on Cancel, which VarType
    ' tells apart from any path string - the same reason PkgClean uses
    ' Application.InputBox rather than VBA's InputBox.
    picked = Application.GetOpenFilename( _
        FileFilter:="Delimited files (*.csv;*.txt;*.tsv),*.csv;*.txt;*.tsv", _
        title:="Pick the file to import")
    If VarType(picked) = vbBoolean Then Exit Sub                  ' user cancelled

    destName = PkgImport_SheetNameFor(CStr(picked))
    If Not PkgImport_DestUsable(destName, TITLE) Then Exit Sub

    Core_SpeedOn
    Set wbTxt = PkgImport_OpenDelimited(CStr(picked), _
                    Core_Cfg(ctx, "k_import_delim"), Core_Cfg(ctx, "k_import_qualifier"))
    Set dest = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    dest.name = destName
    rowsOut = Core_LastRowUp(wbTxt.Worksheets(1), 1)
    colsOut = Core_LastCol(wbTxt.Worksheets(1), 1)
    wbTxt.Worksheets(1).UsedRange.Copy dest.Range("A1")
    Application.CutCopyMode = False
    wbTxt.Close SaveChanges:=False
    Core_SpeedOff

    Core_Done ctx, "Imported " & Core_FileName(CStr(picked)) & " into a new sheet '" & _
           destName & "'." & vbCrLf & rowsOut & " row(s), " & colsOut & " column(s).", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "import.delimited_file", Err
End Sub
```

- [ ] **Step 6: Write the package-private helpers**

```vb
'================================================================
' Package-private helpers. Prefixed PkgImport_ because the bundle is
' one flat namespace (design section 3.3), even though these are Private.
'================================================================

' A destination name this package may create. It REFUSES rather than
' overwrites: the library snippet this replaces opened with
' ThisWorkbook.Worksheets("Master").Delete under On Error Resume Next,
' which silently destroys whatever the user had on that sheet. Nothing in
' this package deletes a sheet it did not create in the same run.
Private Function PkgImport_DestUsable(ByVal destName As String, ByVal title As String) As Boolean
    Dim bad As String, i As Long
    If Len(destName) = 0 Then
        MsgBox "No name is set for the sheet this creates." & vbCrLf & vbCrLf & _
               "Fill in 'Name for the combined sheet this creates' on the Setup tab, " & _
               "Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, title
        Exit Function
    End If
    If Len(destName) > 31 Then
        MsgBox "'" & destName & "' is " & Len(destName) & " characters. Excel sheet names " & _
               "stop at 31." & vbCrLf & vbCrLf & "Nothing has been changed.", vbExclamation, title
        Exit Function
    End If
    bad = ":\/?*[]"
    For i = 1 To Len(bad)
        If InStr(destName, Mid$(bad, i, 1)) > 0 Then
            MsgBox "'" & destName & "' contains " & Mid$(bad, i, 1) & ", which Excel does not " & _
                   "allow in a sheet name." & vbCrLf & vbCrLf & _
                   "None of these are allowed:  : \ / ? * [ ]" & vbCrLf & vbCrLf & _
                   "Nothing has been changed.", vbExclamation, title
            Exit Function
        End If
    Next i
    If Core_SheetExists(destName) Then
        MsgBox "A sheet named '" & destName & "' already exists in this workbook, and this " & _
               "action will not overwrite it." & vbCrLf & vbCrLf & _
               "Rename or delete that sheet, then run " & title & " again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, title
        Exit Function
    End If
    PkgImport_DestUsable = True
End Function

' Sheets the engine owns. Combining one of these into a data sheet would
' be meaningless at best; _Config in particular is the activated config
' store. Generated button pages are "Macros 1", "Macros 2", ... so the
' test is a prefix, not a list.
Private Function PkgImport_IsEngineSheet(ByVal nm As String) As Boolean
    Select Case LCase(Trim(nm))
        Case "setup", "setup menu", "_config", "_config_staged", "_packages", _
             "macros", "macros home"
            PkgImport_IsEngineSheet = True
        Case Else
            PkgImport_IsEngineSheet = (LCase(Left$(nm, 7)) = "macros ")
    End Select
End Function

' Blank means the folder this workbook is saved in - contract section 9's
' "data location", and the answer in the overwhelming majority of cases.
' Always returns a path ending in a separator.
Private Function PkgImport_ResolveFolder(ByVal raw As String) As String
    Dim p As String
    p = Trim(raw)
    If Len(p) = 0 Then p = ThisWorkbook.Path
    If Len(p) = 0 Then Exit Function                 ' workbook never saved
    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    PkgImport_ResolveFolder = p
End Function

' Every workbook in the folder except this one. Collects names FIRST and
' returns full paths, because calling anything that touches the file system
' inside a Dir() loop can corrupt the iterator on a network drive - the same
' reason Core_FindFile is written this way.
Private Function PkgImport_WorkbookPaths(ByVal folderPath As String) As Collection
    Dim out As Collection, names As Collection, f As String, one As Variant
    Set out = New Collection
    Set names = New Collection
    f = Dir(folderPath & "*.xls*")
    Do While Len(f) > 0
        names.Add f
        f = Dir
    Loop
    For Each one In names
        If LCase(CStr(one)) <> LCase(ThisWorkbook.name) Then out.Add folderPath & CStr(one)
    Next one
    Set PkgImport_WorkbookPaths = out
End Function

' A sheet name from a file name: drop the extension, replace the characters
' Excel forbids, and cut to 31. It does NOT make the name unique - a clash is
' reported by PkgImport_DestUsable and refused, exactly like a combine's.
Private Function PkgImport_SheetNameFor(ByVal fullPath As String) As String
    Dim nm As String, bad As String, i As Long
    nm = Core_FileName(fullPath)
    If InStrRev(nm, ".") > 1 Then nm = Left$(nm, InStrRev(nm, ".") - 1)
    bad = ":\/?*[]"
    For i = 1 To Len(bad)
        nm = Replace(nm, Mid$(bad, i, 1), "-")
    Next i
    nm = Trim(nm)
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    PkgImport_SheetNameFor = nm
End Function

' Open a delimited file with the separator and qualifier the deployment
' configured, instead of letting Excel guess from the extension. Workbooks.Open
' on a .csv applies the machine's list separator; a European-configured machine
' opening a comma file lands everything in column A, which is exactly the
' failure this action exists to prevent.
Private Function PkgImport_OpenDelimited(ByVal fullPath As String, ByVal delim As String, _
                                         ByVal qual As String) As Workbook
    Dim useTab As Boolean, useComma As Boolean, useSemi As Boolean
    Dim useSpace As Boolean, useOther As Boolean
    Dim otherCh As String, tq As Long, d As String

    d = LCase(Trim(delim))
    Select Case d
        Case "", "comma", ","
            useComma = True
        Case "tab"
            useTab = True
        Case "semicolon", ";"
            useSemi = True
        Case "space"
            useSpace = True
        Case "pipe", "|"
            useOther = True
            otherCh = "|"
        Case Else
            useOther = True
            otherCh = Left$(Trim(delim), 1)
    End Select

    Select Case Trim(qual)
        Case "'"
            tq = xlTextQualifierSingleQuote
        Case """"
            tq = xlTextQualifierDoubleQuote
        Case Else
            tq = xlTextQualifierNone
    End Select

    Workbooks.OpenText Filename:=fullPath, DataType:=xlDelimited, TextQualifier:=tq, _
        Tab:=useTab, Semicolon:=useSemi, Comma:=useComma, Space:=useSpace, _
        Other:=useOther, OtherChar:=otherCh, Local:=True
    Set PkgImport_OpenDelimited = ActiveWorkbook
End Function
```

- [ ] **Step 7: Build**

```bash
pwsh build/build.ps1
```

Expected: exit 0, `Core + 3 package(s): PkgClean.bas, PkgImport.bas, PkgImportExport.bas`,
balanced `Sub` and `Function` counts. The id gate from Task 2 now checks the three new ids;
if it names one, the id is wrong, **not** the gate.

```bash
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: no diagnostics.

- [ ] **Step 8: Confirm the three ids reached the bundle**

```bash
grep -c 'import\.combine_worksheets\|import\.combine_workbooks\|import\.delimited_file' dist/MacrosBundle.bas
```

Expected: `3`.

- [ ] **Step 9: Commit**

```bash
git add src/packages/PkgImport.bas src/core/Core.bas
git commit -m "feat(import): combine worksheets, combine workbooks, import a delimited file"
```

**`dist/` is gitignored** (`.gitignore:16`). Do not `git add -f` the bundle — it is a build
artifact, the exit criteria check it on disk, and the owner rebuilds it before importing.

---

### Task 4: `src/packages/PkgClean.bas` — the four clean actions

**Owner: delegate to codex (`terra` or `sonnet`).** Hand the subagent this task in full plus
the existing `src/packages/PkgClean.bas`. It reuses that file's own private helpers
(`AskRange`, `IsBlankVal`) — it must not duplicate them, and it must not touch the three
shipped actions.

**Files:**
- Modify: `src/packages/PkgClean.bas`

**Interfaces:**
- Consumes: everything Task 3 lists, plus `Core_AddTable`, `Core_AddDefault`,
  `Core_CfgTable`, `Core_RangeRef`, `Core_Coerce`, and PkgClean's existing private
  `AskRange(title As String) As Range` (returns `Nothing` on cancel or a multi-area
  selection).
- Produces: four public entry points — `PkgClean_RemoveBlanks`, `PkgClean_UnmergeFill`,
  `PkgClean_StandardizeFormats`, `PkgClean_TextToNumbers`.

**Additional Core API for this task:**

```vb
Core_AddTable(reg As Object, key As String, label As String, colHeaders As String, _
              Optional tier As String = "simple", Optional help As String = "", _
              Optional usedBy As String = "")      ' colHeaders is "|"-delimited, up to 3
Core_AddDefault(reg As Object, key As String, rowCsv As String)   ' one call per default row
Core_CfgTable(ctx As Object, key As String, ByRef a() As String, ByRef b() As String, _
              ByRef c() As String) As Boolean      ' False when no rows are configured
Core_RangeRef(rng As Range) As String              ' "Sheet1!B2:M400", sheet-qualified
Core_Coerce(v As Variant, code As String) As Variant   ' "N" -> Double, or v unchanged
```

- [ ] **Step 1: Add the one config field to the manifest**

`PkgClean_Manifest` currently opens with a comment block asserting the package declares no
fields. That comment is load-bearing — it records why `clean_keycols` and `clean_hasheader`
were withdrawn — so **extend it, do not replace it.** Insert immediately after the existing
`' ... Do NOT reintroduce these as remembered settings.` line:

```vb
    ' v2.3.0 adds ONE field, and it is not a walk-back of the paragraph above.
    ' A house number format is a fact about the DEPLOYMENT - the same columns get
    ' the same formats every close - not a fact about the range the caller is
    ' about to pick. Its rows address columns by LETTER or by HEADER NAME, never
    ' by position within the selection; a position counted from the left of a
    ' selection that does not exist yet would be clean_keycols under a new name.
    Core_AddTable reg, "k_clean_formats", "Column formats", _
        "Column (letter or header name)|Number format", "simple", _
        "One row per column you want formatted. Left cell: a column letter like C, or the column's header text. Right cell: the format, e.g. mm/dd/yyyy or $#,##0.00.", _
        "clean.standardize_formats"
    Core_AddDefault reg, "k_clean_formats", "A|mm/dd/yyyy"
    Core_AddDefault reg, "k_clean_formats", "B|0"
    Core_AddDefault reg, "k_clean_formats", "C|$#,##0.00"
    Core_AddDefault reg, "k_clean_formats", "D|0.00%"
```

One `Core_AddDefault` per line — never a single `Array(...)`, per the 25-continuation rule.

Then append four `Core_AddAction` lines after the existing three. **No `category:=`** — the
default is the package display name, which puts them on the same card as `clean.dupes`.

```vb
    Core_AddAction reg, "clean.remove_blanks", "Remove Blank Rows & Columns", "PkgClean_RemoveBlanks", "Pick a range; shows you how many completely blank rows and columns it found, and deletes them only after you say yes."
    Core_AddAction reg, "clean.unmerge_fill", "Unmerge & Fill Down", "PkgClean_UnmergeFill", "Pick a range; unmerges every merged block and puts the block's value into every cell it covered, so the data can be sorted and filtered."
    Core_AddAction reg, "clean.standardize_formats", "Standardize Formats", "PkgClean_StandardizeFormats", "Pick a range; applies the number formats listed in the Column formats table on Setup. Changes how values look, never the values themselves."
    Core_AddAction reg, "clean.text_to_numbers", "Convert Text to Numbers", "PkgClean_TextToNumbers", "Pick a range; turns text that represents a number into a real number, stripping $ and thousands separators. Formulas are never touched. Note that 12% becomes 12, not 0.12."
End Sub
```

- [ ] **Step 2: Write `clean.remove_blanks`**

The safety story the catalog's deferred siblings lack: scan, report, confirm, then delete —
and never delete a whole row on the strength of a partial selection.

```vb
'================================================================
' ACTION - Delete completely blank rows and columns
'
' Two counts, not one, and the difference matters. A row can be blank
' INSIDE the picked range while holding data outside it; deleting the
' whole row would then destroy something the user never selected. Those
' rows are counted, reported, and left alone. Only rows blank across the
' sheet's whole used extent are deleted.
'================================================================
Public Sub PkgClean_RemoveBlanks(ctx As Object)
    Dim rng As Range, ws As Worksheet
    Dim rowsIn As Collection, rowsDel As Collection
    Dim colsIn As Collection, colsDel As Collection
    Dim i As Long
    Const TITLE As String = "Remove Blank Rows & Columns"
    On Error GoTo Fail

    Set rng = AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                  ' user cancelled
    Set ws = rng.Parent

    Core_SpeedOn
    PkgClean_ScanBlanks rng, rowsIn, rowsDel, colsIn, colsDel
    Core_SpeedOff

    If rowsIn.Count = 0 And colsIn.Count = 0 Then
        MsgBox "There are no completely blank rows or columns inside " & Core_RangeRef(rng) & _
               "." & vbCrLf & vbCrLf & "Nothing has been changed.", vbInformation, TITLE
        Exit Sub
    End If
    If Not PkgClean_ConfirmBlanks(rng, rowsIn, rowsDel, colsIn, colsDel, TITLE) Then Exit Sub

    Core_SpeedOn
    ' Bottom-up and right-to-left: deleting row 7 renumbers everything below it,
    ' so a top-down loop skips every other match.
    For i = rowsDel.Count To 1 Step -1
        ws.Rows(rowsDel(i)).Delete
    Next i
    For i = colsDel.Count To 1 Step -1
        ws.Columns(colsDel(i)).Delete
    Next i
    Core_SpeedOff

    Core_Done ctx, "Deleted " & rowsDel.Count & " blank row(s) and " & colsDel.Count & _
           " blank column(s)." & vbCrLf & vbCrLf & _
           PkgClean_HeldBackNote(rowsIn.Count - rowsDel.Count, colsIn.Count - colsDel.Count), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "clean.remove_blanks", Err
End Sub

' Fills four collections of ABSOLUTE sheet row/column numbers:
'   rowsIn  - blank inside the picked range
'   rowsDel - blank inside AND across the sheet's used extent (safe to delete)
' and the same pair for columns. Numbers are collected ascending; the caller
' deletes in reverse.
Private Sub PkgClean_ScanBlanks(rng As Range, ByRef rowsIn As Collection, ByRef rowsDel As Collection, _
                                ByRef colsIn As Collection, ByRef colsDel As Collection)
    Dim ws As Worksheet, r As Long, c As Long
    Set ws = rng.Parent
    Set rowsIn = New Collection
    Set rowsDel = New Collection
    Set colsIn = New Collection
    Set colsDel = New Collection

    For r = rng.Row To rng.Row + rng.Rows.Count - 1
        If Application.CountA(Application.Intersect(ws.Rows(r), rng)) = 0 Then
            rowsIn.Add r
            If Application.CountA(ws.Rows(r)) = 0 Then rowsDel.Add r
        End If
    Next r
    For c = rng.Column To rng.Column + rng.Columns.Count - 1
        If Application.CountA(Application.Intersect(ws.Columns(c), rng)) = 0 Then
            colsIn.Add c
            If Application.CountA(ws.Columns(c)) = 0 Then colsDel.Add c
        End If
    Next c
End Sub

' The preview the deferred destructive items lack. vbYesNo, not vbYesNoCancel:
' here the title-bar X returns vbNo, and vbNo is the SAFE answer - it deletes
' nothing. (clean.dupes needs the third button because its vbNo is destructive.)
' Defaults to No.
Private Function PkgClean_ConfirmBlanks(rng As Range, rowsIn As Collection, rowsDel As Collection, _
                                        colsIn As Collection, colsDel As Collection, _
                                        ByVal title As String) As Boolean
    Dim msg As String, held As String
    msg = "Inside " & Core_RangeRef(rng) & " this will delete:" & vbCrLf & vbCrLf & _
          "   " & rowsDel.Count & " completely blank row(s)" & vbCrLf & _
          "   " & colsDel.Count & " completely blank column(s)" & vbCrLf & vbCrLf
    held = PkgClean_HeldBackNote(rowsIn.Count - rowsDel.Count, colsIn.Count - colsDel.Count)
    If Len(held) > 0 Then msg = msg & held & vbCrLf & vbCrLf
    msg = msg & "Whole rows and columns are removed, so everything below and to the " & _
          "right shifts up and left. This cannot be undone with Ctrl+Z." & vbCrLf & vbCrLf & _
          "Delete them?"
    PkgClean_ConfirmBlanks = (MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, title) = vbYes)
End Function

' Names what was deliberately NOT deleted. Returns "" when nothing was held back,
' so the caller can test Len() rather than compare against a sentence.
Private Function PkgClean_HeldBackNote(ByVal heldRows As Long, ByVal heldCols As Long) As String
    If heldRows = 0 And heldCols = 0 Then Exit Function
    PkgClean_HeldBackNote = _
        heldRows & " row(s) and " & heldCols & " column(s) are blank inside your " & _
        "selection but hold data outside it, so they were left alone - deleting them " & _
        "would have removed data you did not select."
End Function
```

- [ ] **Step 3: Write `clean.unmerge_fill`**

```vb
'================================================================
' ACTION - Unmerge every merged block and fill its value down
'================================================================
Public Sub PkgClean_UnmergeFill(ctx As Object)
    Dim rng As Range, c As Range, seen As Object, addrs As Collection
    Dim area As Range, i As Long, v As Variant, fml As String, hadFormula As Boolean
    Const TITLE As String = "Unmerge & Fill Down"
    On Error GoTo Fail

    Set rng = AskRange(TITLE)
    If rng Is Nothing Then Exit Sub

    Core_SpeedOn
    ' Collect the merge addresses FIRST. Unmerging inside the loop that is
    ' walking the cells changes what MergeArea reports mid-iteration.
    Set seen = CreateObject("Scripting.Dictionary")
    Set addrs = New Collection
    For Each c In rng.Cells
        If c.MergeCells Then
            If Not seen.Exists(c.MergeArea.Address) Then
                seen.Add c.MergeArea.Address, 1
                addrs.Add c.MergeArea.Address
            End If
        End If
    Next c

    For i = 1 To addrs.Count
        Set area = rng.Parent.Range(addrs(i))
        v = area.Cells(1, 1).Value                   ' the value the merge displayed
        hadFormula = area.Cells(1, 1).HasFormula
        If hadFormula Then fml = area.Cells(1, 1).Formula
        area.UnMerge
        area.Value = v
        ' Put the anchor's formula back ONLY when there was one. Restoring
        ' .Formula unconditionally would round-trip every constant through its
        ' text form, and a text cell holding "0123" comes back as the number 123.
        If hadFormula Then area.Cells(1, 1).Formula = fml
    Next i
    Core_SpeedOff

    If addrs.Count = 0 Then
        MsgBox "There are no merged cells inside " & Core_RangeRef(rng) & "." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbInformation, TITLE
        Exit Sub
    End If

    Core_Done ctx, "Unmerged " & addrs.Count & " merged block(s) in " & Core_RangeRef(rng) & _
           " and filled each one's value into every cell it covered.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "clean.unmerge_fill", Err
End Sub
```

- [ ] **Step 4: Write `clean.standardize_formats`**

```vb
'================================================================
' ACTION - Apply the configured number formats to the picked range
'================================================================
Public Sub PkgClean_StandardizeFormats(ctx As Object)
    Dim rng As Range, cols() As String, fmts() As String, spare() As String
    Dim i As Long, tgt As Range, applied As Long, missed As String
    Const TITLE As String = "Standardize Formats"
    On Error GoTo Fail

    Set rng = AskRange(TITLE)
    If rng Is Nothing Then Exit Sub

    If Not Core_CfgTable(ctx, "k_clean_formats", cols, fmts, spare) Then
        MsgBox "No column formats are set up yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Column formats' table on the Setup tab, Finalize, and run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    Core_SpeedOn
    For i = LBound(cols) To UBound(cols)
        Set tgt = PkgClean_ResolveFormatCol(rng, cols(i))
        If tgt Is Nothing Then
            missed = missed & vbCrLf & "   " & cols(i)
        ElseIf Len(Trim(fmts(i))) > 0 Then
            tgt.NumberFormat = fmts(i)
            applied = applied + 1
        End If
    Next i
    Core_SpeedOff

    ' A mapping row that matched nothing is REPORTED, never silently dropped: a
    ' renamed header is the ordinary way this table goes stale, and a silent skip
    ' looks exactly like a format that was applied.
    Core_Done ctx, "Applied " & applied & " column format(s) inside " & Core_RangeRef(rng) & _
           "." & vbCrLf & "Values were not changed - only how they are displayed." & _
           IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These rows matched no column in your " & _
               "selection and were skipped:" & missed, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "clean.standardize_formats", Err
End Sub

' Resolve one "Column" cell against the picked range. A value of one to three
' letters is read as a COLUMN LETTER; anything else is matched against the text
' in the range's first row. Never a position counted from the left of the
' selection - see the manifest comment on k_clean_formats for why.
'
' The letter reading wins a tie, so a range whose header literally reads "A"
' cannot be addressed by name. That is documented in the field's help text
' rather than resolved, because a heuristic that guessed would be worse.
' Returns Nothing when the column falls outside the picked range.
Private Function PkgClean_ResolveFormatCol(rng As Range, ByVal spec As String) As Range
    Dim s As String, absCol As Long, i As Long, hdr As String
    s = Trim(spec)
    If Len(s) = 0 Then Exit Function

    If Len(s) <= 3 And PkgClean_IsAlpha(s) Then
        On Error Resume Next
        absCol = rng.Parent.Range(s & "1").Column
        On Error GoTo 0
        If absCol > 0 Then
            If absCol >= rng.Column And absCol <= rng.Column + rng.Columns.Count - 1 Then
                Set PkgClean_ResolveFormatCol = Application.Intersect(rng.Parent.Columns(absCol), rng)
            End If
            Exit Function
        End If
    End If

    For i = 1 To rng.Columns.Count
        hdr = Trim(CStr(rng.Cells(1, i).Value))
        If StrComp(hdr, s, vbTextCompare) = 0 Then
            Set PkgClean_ResolveFormatCol = rng.Columns(i)
            Exit Function
        End If
    Next i
End Function

Private Function PkgClean_IsAlpha(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = UCase(Mid$(s, i, 1))
        If ch < "A" Or ch > "Z" Then Exit Function
    Next i
    PkgClean_IsAlpha = True
End Function
```

- [ ] **Step 5: Write `clean.text_to_numbers`**

```vb
'================================================================
' ACTION - Convert text that represents a number into a real number
'
' Calls Core_Coerce with the "N" code rather than reimplementing the
' stripping rules - design section 3.4. Core_Coerce returns the value
' UNCHANGED when it cannot parse it, so a Double back means it parsed
' and anything else means leave the cell alone.
'================================================================
Public Sub PkgClean_TextToNumbers(ctx As Object)
    Dim rng As Range, arr As Variant, r As Long, c As Long, n As Long
    Dim got As Variant
    Const TITLE As String = "Convert Text to Numbers"
    On Error GoTo Fail

    Set rng = AskRange(TITLE)
    If rng Is Nothing Then Exit Sub

    Core_SpeedOn
    arr = rng.Value
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            For c = 1 To UBound(arr, 2)
                If VarType(arr(r, c)) = vbString Then
                    If Len(arr(r, c)) > 0 Then
                        got = Core_Coerce(arr(r, c), "N")
                        If VarType(got) = vbDouble Then
                            If Not rng.Cells(r, c).HasFormula Then
                                rng.Cells(r, c).NumberFormat = "General"
                                rng.Cells(r, c).Value = got
                                n = n + 1
                            End If
                        End If
                    End If
                End If
            Next c
        Next r
    ElseIf VarType(arr) = vbString Then                  ' single-cell range
        If Len(arr) > 0 Then
            got = Core_Coerce(arr, "N")
            If VarType(got) = vbDouble Then
                If Not rng.HasFormula Then
                    rng.NumberFormat = "General"
                    rng.Value = got
                    n = 1
                End If
            End If
        End If
    End If
    Core_SpeedOff

    Core_Done ctx, "Converted " & n & " cell(s) to real numbers in " & Core_RangeRef(rng) & _
           "." & vbCrLf & vbCrLf & "Text that is not a number, and every formula cell, " & _
           "was left exactly as it was.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "clean.text_to_numbers", Err
End Sub
```

- [ ] **Step 6: Build and check the ids**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 3 package(s)`, balanced counts, no diagnostics.

```bash
grep -c 'clean\.remove_blanks\|clean\.unmerge_fill\|clean\.standardize_formats\|clean\.text_to_numbers' dist/MacrosBundle.bas
```

Expected: `4`.

- [ ] **Step 7: Commit**

```bash
git add src/packages/PkgClean.bas
git commit -m "feat(clean): remove blanks, unmerge and fill, standardize formats, text to numbers"
```

---

### Task 5: All seven ids, end to end

**Owner: Opus.** A short verification task with no new code — it exists because the exit
criteria are stated against the *bundle*, and nothing so far has checked all seven together
or checked that `SCHEMA_VER` survived.

**Files:** none modified unless something fails.

- [ ] **Step 1: Every frozen id, exactly as §3.4 froze it**

```bash
for id in import.combine_worksheets import.combine_workbooks import.delimited_file \
          clean.remove_blanks clean.unmerge_fill clean.standardize_formats clean.text_to_numbers; do
  printf '%-32s %s\n' "$id" "$(grep -c "\"$id\"" dist/MacrosBundle.bas)"
done
```

Expected: every line prints `1`. A `0` means an id was mistyped; a `2` means it was declared
twice and the Task 2 gate should already have caught it — investigate the gate, not just the
id.

- [ ] **Step 2: `SCHEMA_VER` did not move**

```bash
grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas
```

Expected: `1`. If this prints `0`, something changed the stored-format marker and the wave
is no longer additive — stop and reconcile against contract §10 before going further.

- [ ] **Step 3: The build reports what the exit criteria expect**

```bash
pwsh build/build.ps1
```

Expected, literally: `Core + 3 package(s)` and `Sub <n>/<n> | Function <m>/<m>` with each
pair equal.

- [ ] **Step 4: Refresh the code map**

`Core.bas`, `PkgClean.bas` and a new `PkgImport.bas` all changed, so every anchor after them
is stale — and a stale map is worse than none, because it is trusted.

```bash
node scripts/code-map.mjs
git add docs/code-map.md && git commit -m "docs: refresh the code map for W1"
```

---

### Task 6: The gate — `GateAudit` probe, checklist, blind dry-run

**Owner: Opus writes the checklist; delegate the `GateAudit` probe to codex (`sonnet`).**
The checklist is the deliverable the owner actually holds, and `CLAUDE.md` §7 governs how it
is written.

**Files:**
- Modify: `tools/GateAudit.bas`
- Create: `docs/designs/2026-07-28-w1-manual-gate.md`

**Interfaces:**
- Consumes: the Setup sheet built by `Core_BuildSetup` from the Task 3/4 bundle. The Action
  catalog block begins at the row whose column A holds `k_core_layout`; each action row
  carries its id in **column J** (`RID_COL`), hidden.
- Produces: nothing later in W1 consumes it.

- [ ] **Step 1: Add the catalog-coverage probe to `GateAudit`**

`tools/GateAudit.bas` is not product code — nothing in `build/build.ps1` reads `tools/`, so
it never ships. Follow the file's existing conventions exactly: every probe wrapped so it
can never raise, and a probe that cannot run reports `Skp` with a reason rather than passing.
Add a `Private Sub AuditActionCatalog()` and call it from `GateAudit_Run` alongside the
existing probes.

It asserts one thing: **every id below is present in column J of the Action catalog block.**
That is the whole "an action reached the surface" question, and it is answerable without
executing any action.

```vb
' Every action id this build should expose. The seven W1 ids plus the eight
' shipped before them - contract section 8. One Add per line: the 25-continuation
' rule, and a 15-element Array(...) would breach it.
Private Function ExpectedIds() As Collection
    Dim c As New Collection
    c.Add "importexport.clear"
    c.Add "importexport.import"
    c.Add "importexport.updateMaster"
    c.Add "importexport.export"
    c.Add "importexport.suggestTypes"
    c.Add "clean.dupes"
    c.Add "clean.trim"
    c.Add "clean.filldown"
    c.Add "clean.remove_blanks"
    c.Add "clean.unmerge_fill"
    c.Add "clean.standardize_formats"
    c.Add "clean.text_to_numbers"
    c.Add "import.combine_worksheets"
    c.Add "import.combine_workbooks"
    c.Add "import.delimited_file"
    Set ExpectedIds = c
End Function
```

And the probe itself. `KeyRow`, `NextKeyRowAt`, `Sh`, `Head`, `Say`, `Pas`, `Fal` and `Skp`
are already in the file — reuse them, do not add near-duplicates.

```vb
' Every declared action reaches the Setup Action catalog. This is the whole
' "did the action make it to the surface" question, and it is answerable
' WITHOUT running any action - which matters, because several of them delete
' rows or create sheets and a gate must not do that to get its answer.
'
' Ids live in column 10 (RID_COL, hidden) of the catalog block, which starts at
' the k_core_layout key row and ends at the next key row.
Private Sub AuditActionCatalog()
    Dim w As Worksheet, want As Collection, found As Object
    Dim rStart As Long, rEnd As Long, r As Long
    Dim id As String, one As Variant, extras As String, nFound As Long
    Head "ACTION CATALOG COVERAGE      [W1]"

    Set w = Sh("Setup")
    If w Is Nothing Then Skp "action catalog coverage", "no Setup sheet": Exit Sub
    rStart = KeyRow(w, "k_core_layout")
    If rStart = 0 Then Skp "action catalog coverage", "k_core_layout key row not found": Exit Sub
    rEnd = NextKeyRowAt(w, rStart)
    If rEnd <= rStart + 1 Then Skp "action catalog coverage", "the catalog block is empty": Exit Sub

    Set found = CreateObject("Scripting.Dictionary")
    found.CompareMode = 1
    For r = rStart + 1 To rEnd - 1
        id = Trim(CStr(w.Cells(r, 10).Value))
        If Len(id) > 0 Then
            If Not found.Exists(id) Then found.Add id, r
            nFound = nFound + 1
        End If
    Next r
    Say "  catalog rows " & (rStart + 1) & "-" & (rEnd - 1) & ", " & nFound & " id(s) present"

    Set want = ExpectedIds()
    For Each one In want
        If found.Exists(CStr(one)) Then
            Pas "catalog lists " & CStr(one)
        Else
            Fal "catalog lists " & CStr(one), "no row in column J carries that id"
        End If
    Next one

    ' An id on the sheet that no package declares is exactly as interesting as a
    ' missing one - it is how a stale build announces itself. Reported, not failed:
    ' a deployment may legitimately carry a package this list does not know about.
    For Each one In found.Keys
        If Not AuditIdExpected(want, CStr(one)) Then extras = extras & vbCrLf & "     " & CStr(one)
    Next one
    If Len(extras) > 0 Then
        Say "  UNEXPECTED ids on the sheet (stale build, or a package this audit does not know):" & extras
    Else
        Say "  no unexpected ids"
    End If
End Sub

Private Function AuditIdExpected(want As Collection, ByVal id As String) As Boolean
    Dim one As Variant
    For Each one In want
        If StrComp(CStr(one), id, vbTextCompare) = 0 Then AuditIdExpected = True: Exit Function
    Next one
End Function
```

Then add `AuditActionCatalog` to `GateAudit_Run`'s probe sequence, after `AuditOrder`.

- [ ] **Step 2: Verify the probe fails when it should**

Import the bundle into a scratch workbook, build Setup, then **delete one action row** from
the catalog block by hand and re-run `GateAudit_Run`. Expected: that id reported `FAIL`, the
other fourteen `PASS`. A probe that has only ever been seen passing is not evidence.

- [ ] **Step 3: Write the manual gate**

Create `docs/designs/2026-07-28-w1-manual-gate.md`. `CLAUDE.md` §7 governs every item:

> Each item states **the observable result first**, then what a failure looks like, then
> why. An item whose expected result equals the default state proves nothing. Claims about
> Excel's own behaviour are **assumptions** until the gate confirms them — mark them.

Cover, at minimum:

1. **Import hygiene** — remove the old `MacrosBundle` module *before* importing the new one.
   Excel's import never replaces a same-named module; it adds `MacrosBundle1`, duplicating
   ~200 procedures into `Ambiguous name detected` on compile. Empirically confirmed
   2026-07-28.
2. **Debug ▸ Compile VBAProject** is clean.
3. **`GateAudit_Run`** reports all fifteen ids present.
4. **Each of the seven actions runs from its button** and produces the stated result.
5. **Each of the seven cancel paths** leaves the workbook unchanged — the range picker's
   Cancel, `GetOpenFilename`'s Cancel, and `clean.remove_blanks`' No **and** its title-bar X.
6. **The two refusal paths**: a combine when `k_import_dest` already names a sheet
   (must refuse and change nothing), and `clean.standardize_formats` with an empty table.
7. **The partial-failure cleanup**: point `k_import_folder` at a folder containing a file
   that cannot be opened, run `import.combine_workbooks`, and confirm **no orphaned
   destination sheet is left behind**.
8. **`clean.remove_blanks`' held-back count** — select a range whose blank rows carry data
   outside the selection, and confirm those rows are reported as held back and survive.

**Mark these as assumptions**, since each is a claim about Excel rather than about this code,
and none has been executed:

- `Application.GetOpenFilename` returns Boolean `False` on Cancel (the plan's cancel test).
- `Workbooks.OpenText` leaves the opened file as `ActiveWorkbook`, which
  `PkgImport_OpenDelimited` returns.
- `Workbooks.OpenText`'s `Local:=True` coexists with an explicit `TextQualifier`.
- `Application.CountA` accepts a whole-row/whole-column `Range` without an overflow.
- `Range.UnMerge` leaves the anchor's value and formula intact and the covered cells empty.
- `dest.Delete` inside the engine's error handler does not prompt, because `Core_SpeedOn`
  already set `DisplayAlerts = False`.

- [ ] **Step 4: Blind dry-run — required before the checklist reaches the owner**

Dispatch a subagent with the checklist and the built `dist/MacrosBundle.bas`, and
**forbid it from reading `src/`**. The blindness is the mechanism: it reproduces the only
reader who matters.

Instruct it to flag every item it cannot execute (undefined term, missing precondition,
unstated setup), cannot judge (no observable expected result), or that would pass no matter
what the build did — and to ask that last one out loud: *could this item pass on a broken
build?*

**Know its limit**, and say so in the PR: it reads VBA, it does not execute it. Across three
gates it caught 13 + 6 + 7 authoring defects and **zero** runtime defects. It is an authoring
gate, not a test.

- [ ] **Step 5: Fix what the dry-run found, then commit**

```bash
git add tools/GateAudit.bas docs/designs/2026-07-28-w1-manual-gate.md
git commit -m "test: W1 manual gate, plus a GateAudit probe for action-catalog coverage"
```

---

### Task 7: Ship

**Owner: Opus.**

- [ ] **Step 1: Codex review of the finished diff**

```
/codex review
```

over the whole `feat/w1-ingest-clean` diff. Treat a `[P1]` as blocking and fix it **before**
opening the PR. Record the review in the PR body.

- [ ] **Step 2: Version and changelog**

`VERSION` → `0.10.0`. Add a `## [0.10.0] - 2026-07-28` entry to `CHANGELOG.md` in the house
style — prose that says what the user can now do and what was deliberately not done, not a
bullet list of file names. Cover the seven actions, the id gate, the four settled decisions,
and the deferred legacy-category split.

- [ ] **Step 3: Final gates, all four**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas
grep -n 'Contract version: 2.3.0' spec/contract.md
git status --short
```

Expected: build exit 0 with `Core + 3 package(s)`; no diagnostics; `1`; one match; and a tree
clean but for the untracked `.vscode/`.

- [ ] **Step 4: Open one PR against `main`**

**One PR. Do not stack.** The PR body records: both codex reviews (the plan review and the
diff review) with no unresolved `[P1]` from either; the four owner decisions; the deferred
items above; the manual-gate status, with every checklist item either ticked against a real
run or **explicitly recorded as untested**; and no attribution footer — check the tail before
submitting.

---

## Unknowns and risks carried into this wave

Restated from the design doc's handoff log, plus what this plan adds. None blocks W1.

- **8 of 15 track-1 gate items are still unrun and all need the owner:** hardening 1.3, 2.2,
  3.1, 4.1, 4.2; reorder 2 (clipping half) and 3; reskin 4 (dark Excel theme). Both
  load-bearing items — #8 crash recovery and #12 fail-closed — passed. **Reskin 4 is the
  likeliest real finding:** the palette came off a dark-background proof and lands on an
  explicitly white sheet, and nothing has checked the two together.
- **W1 adds a fourth Setup band.** `PkgImport` is the first package since v0.5.0 to declare
  fields, and `PkgClean` goes from no band to one. The progressive-disclosure work in
  `docs/designs/setup-simplicity.md` was tuned against a two-band Setup; whether four bands
  still reads well is a question for the owner's eye at the gate, not something this plan
  can settle.
- **The `%` wart in `clean.text_to_numbers`.** `Core_Coerce`'s `N` code strips a trailing
  `%`, so `12%` becomes `12`, not `0.12`. Documented in the action's help text and in
  contract §8 rather than fixed, because fixing it means either diverging from the coercion
  §3.4 says to reuse, or changing `Core_Coerce` and with it every import that already
  depends on its behaviour. Worth revisiting if the owner hits it.
- **Header alignment is case-insensitive but not whitespace- or spelling-tolerant.**
  `Amount ` and `Amount` match (both are trimmed); `Amt` and `Amount` become two columns.
  That is the honest behaviour and the completion dialog says columns were matched by name,
  but a source with inconsistent headers will produce a wider sheet than the user expects.
- **The combines write no number formats.** Values transfer, formats do not — which is
  precisely why `clean.standardize_formats` is in the same wave. Chain them.
