# Handoff — v0.4.0 track 1, after the owner's Excel verification

**Date:** 2026-07-25
**Branch:** `feat/finish-excel-track1`, **38 commits, NOT pushed, no PR.** `main` is still v0.3.0 / contract v1.1.0.
**State:** Tasks 1–9 built and reviewed. Task 10 (the owner's manual Excel gate) **partially run** — it found four items. Task 11 (teardown + ship) not started and **must not start** until these are resolved.

---

## What happened

Tasks 1–9 each had a dedicated code review, several had two or three fix rounds, and a whole-branch review plus one fix wave followed. Then the owner imported the bundle and ran the gate. It immediately earned its keep.

**A compile error, already fixed (`4626460`).** `RenderOffBand` declared `Dim reg As Object, any As Boolean`. **`Any` is a VBA reserved word** (it appears in `Declare … As Any`), so the line cannot compile — and the VBE reports the error against the *enclosing procedure's declaration line* rather than the `Dim`, which is why it presented as a syntax error in a signature that was perfectly valid. Renamed to `anyOff`.

Nothing in the repo could have caught it: `build.ps1`'s checks are structural, `scripts/vba-syntax-check.mjs` parses without flagging reserved-word identifiers, and **four review rounds read the line as valid VBA**. Excel's compiler was the only detector. A gate now exists (`c404349`) that fails the build on a reserved word used as a `Dim`/parameter identifier — demonstrated failing on a reintroduced `any` and passing on the clean tree. Its word list is deliberately conservative and empirical: it omits `name` (`Core_FindCol`) and `base` (`MintChainId`), both pre-existing and both demonstrably compiling since v0.2.0/v0.3.0. **Add a word only when Excel rejects one**, never on suspicion — over-broad, that gate fails builds that would have worked.

---

## The four open items, in the owner's words and priority order

### 1. Placement should accept a page NAME and a CELL, not only a generated page

> *"page code OR PAGE NAME — I specifically wanted the option to add a macro to a specific page and cell, not just organize the buttons on new pages."*

This is a **feature gap against the owner's original intent**, not a defect. Today a placement code is `<page-number><column-letters>` (`1A`, `12AB`) and only ever positions a button on an engine-generated `Macros N` page. The owner wants to put a button on **an existing sheet of their choosing, at a specific cell** — e.g. `TabName!B4`.

Scope this properly before coding; it is the largest of the four:
- **Contract §6** fixes the placement-code grammar and is normative for the Apps Script port. It needs a second form, and §7's naming rules may need a word about sheet-name references.
- The renderer (`RenderPage`) draws onto a grid it owns and controls entirely — column widths, row heights, `Cells.Clear`. **A user's own sheet must not be treated that way.** Placing a button on `TabName` means anchoring a shape to a cell on a sheet the engine must otherwise leave completely alone. That is a genuinely different code path from the card grid, not a parameter on it.
- Rebuild semantics need deciding: `DeletePanels` currently deletes and recreates every generated page. A button placed on a user sheet cannot be cleaned up that way — it needs to be found and removed by name, or the sheet accumulates a duplicate button on every rebuild.
- Interaction with `hide`, with auto-layout, and with `Macros Home`'s removal (item 4) all need thinking through.

### 2. An action appears not to work on a tab other than the one the button is on

> *"it seems like i am not able to run a macro on another tab than the one its on. even when referencing another tab and range (like: `TabName!M2:M12`), it still says it completed on M2:M12. Works if there is something in M1 on the same tab as the macro."*

**Diagnosed from the code, not reproduced — there is no Excel access from the agent side.** Two concrete findings, one unresolved question:

- **The completion message cannot tell you which sheet ran.** `PkgClean`'s actions report `rng.Address(False, False)`, and `Range.Address` **omits the sheet name by design**. So the dialog prints `M2:M12` whether it acted on `TabName` or on the button's own tab. This fully explains the reported symptom *without* implying the operation was wrong. **Fix first**, because it is what makes the rest diagnosable: report `rng.Parent.Name & "!" & rng.Address(False, False)` in every message that names a range, across both packages and Core.
- **`Header:=xlYes` is hardcoded** in `PkgClean_RemoveDupes`. `RemoveDuplicates` treats the range's **first row as a header and excludes it**. The prompt says "(include the header row)", but a user selecting `M2:M12` intending all eleven values considered gets only `M3:M12` deduped. **This is very likely the "works if there is something in M1" observation.** It is a design question (hardcode `xlYes`, ask, or make it a config field) — do not change behaviour without the owner's decision.
- **Unresolved:** whether the operation genuinely acted on the wrong sheet. `AskRange` uses `Application.InputBox(Type:=8)`, which returns a real `Range` carrying its own parent sheet, so it *should* be correct. Sheet-qualifying the message settles it in one run. Check the other two Clean actions and the ImportExport actions for the same class of assumption while you are there.

### 3. The first Setup build took several minutes

> *"the initial load after importing the macro (run setup) took several minutes. all operations after that were quite fast."*

**Predicted by the whole-branch review, cause already identified.** `BuildRegistry` is **uncached**: every call re-runs `CoreCandidates`, one `Application.Run "Pkg<Name>_Manifest"` per candidate, and a `PackageEnabled` worksheet scan per candidate. It is called from `RenderOffBand` (twice), `HasAnyTier`'s per-package loops, `ActionCaptionFor` and `ActionTokenFor` **once per chain step**, `StepsInline` once per step, and `ActionIdSet` **twice** inside a single `ComposeReport` — on top of the main render loops. With the 82-action fixture that is thousands of manifest rebuilds.

Memoize it behind a module-level cache. Two things to get right:
- **Invalidation.** The registry depends on installed modules (fixed per session) *and* on the optional `_Packages` sheet's enable/disable values, which a user can edit between builds. Invalidate at the top of `Core_BuildSetup`, `Core_FinalizeSetup` and `Core_BuildControlPanel` — but **not** inside `Core_Run`, so a 20-step chain builds the registry once.
- **Confirm no caller mutates the returned regs** before caching a shared object. Read every caller; if one does, stop and say so.

Also worth collapsing: `AvailDropdown` makes **two COM validation calls per action row** in `RenderPlacementBlock` — ~164 calls at fixture scale. One `Validation.Add` over the whole column after the loop would do, if nothing depends on the per-cell form.

Note the owner tested with the 82-action fixture; the shipping bundle has 8 actions, so the real-world number is far smaller. Fix it anyway — the catalog roadmap adds ~58 more actions, which is where this bites for real.

### 4. Remove `Macros Home` from the bundle, keep the work

> *"macros home tab is basically useless and shouldn't exist. its needless clutter and since most folks will only ever load a handful of macros or 1 or 2 chains into a workbook, there's really no use for it other than as a demo. save the work for later reference, but remove from the bundle."*

**This reverses the owner's own earlier decision**, which they made before they could see it — exactly what a manual gate is for. Their call stands; do not re-argue it.

*"Save the work"* means preserve the code, not delete it. Suggested: move `BuildHomePage`, `HomeCatalog`, `HomeLegend`, `HomeLine` and anything used only by them into `src/attic/HomeIndex.bas.txt` — a `.txt` extension so `build.ps1`'s `Pkg*.bas` glob and its ASCII and reserved-word gates cannot pick it up and nobody mistakes it for a live module. Head that file with what it was, that it never reached a release, why it was removed, and what re-wiring a restore would need.

Then unwire carefully — there are four non-obvious consequences:

- **`BuildHomePage` was the only thing that activated a sheet and selected a cell at the end of a build.** Removing it leaves the pipeline ending on `BuildAdminPanel` with no deliberate landing place. Choose one (activating the first generated page is the obvious answer) and make sure it cannot raise **1004** — `Range.Select` on a non-active sheet raises, so activate first, and the target may be hidden.
- **`RenderNavRow` writes a `<< Home` link on every page.** With no index that link is dead. Remove it; keep `< Previous` / `Next >`, which still work.
- **`DeletePanels` deliberately spares a sheet named `Macros Home`** — `IsPanelName` requires a numeric tail — so a rebuild could reuse it. With the feature gone, an index left behind by a pre-removal build becomes **orphaned forever**: never deleted, never rebuilt, full of dead links. Add a one-time cleanup mirroring how `DeleteLegacyPanel` retires the old single `Macros` tab.
- **`spec/contract.md` §6 currently *mandates* an index surface** ("the engine **always** generates an index surface…"). Leave that and the Apps Script port gets built to a spec the reference implementation violates. Make the index **optional** (a platform MAY provide one), keep the previous/next navigation requirement, record the withdrawal and the owner's reason in §10's v1.2.0 bullet, and **keep `Macros Home` reserved in §7** — a future restore wants the name and reserving an unused one costs nothing. **Do not bump the contract version**; v1.2.0 is unreleased.

Finally, **the Task 10 checklist has to be re-cut**: delete every item verifying the page list, the button catalog, the legend, the `<< Home` link and Home's own links, and the "no duplicate Home tab" case. **Do not silently drop the AutoFit risk** — re-aim it. The card-page header check (`Macros 1!F3`, `Clean & Standardize (cont. 2)` — 29 characters, 11pt bold, in a 30-character column, no height floor) becomes the primary AutoFit check now that Home's un-floored catalog is gone.

---

## Do NOT revert (earned across nine tasks and ~20 review rounds)

1. **Setup COLLAPSES settings for `Off` actions; it never omits them.** `Core_BuildSetup` deletes and re-seeds Setup from the *finalized* `_Config`, so omitting would cost toggle → Finalize → rebuild → fill → Finalize and discard un-finalized typing. This already overrode a unanimous three-reviewer panel. **Note:** neither shipped package declares any `usedBy`, so no band is drawn in the real bundle today — the fixture proves the mechanism.
2. **`Off` means off.** No button, refused by the dispatcher, and Finalize **blocks** on a chain referencing one. `Chains only` is the explicit state for "reachable only inside a chain". All four dispatch routes were verified closed.
3. **The schema marker lives in the hidden `RID_COL`, never in a visible header**, and migration is **fail-closed**. The marker is non-numeric (`v1.2`) because a bare `"1.2"` is coerced to a number and read back through a locale-aware `CStr`.
4. **`hide` is filtered in `GatherButtons`** — one funnel above both layout filers — so it works in auto and manual mode alike. Do not add a second filter.
5. **No `Core_SheetExists` guard on nav-link targets.** One was added, and because the target was built last it made every link dead text on a first build. The comment says so.
6. **Nothing on the card grid may be merged.** Excel's `AutoFit` ignores merged cells; that is the entire reason the grid is unmerged.
7. **Module-level `Const` stays in the declarations section**; `_Config` is unhidden before `.Delete`; admin controls stay on `Setup Menu`.
8. **`build.ps1`'s reserved-word list stays conservative.** Empirical only.

## Read manifest

- **`.superpowers/sdd/2026-07-24-finish-excel-track1-plan/progress.md`** (~660 lines) — the execution ledger: every task, review, adjudication, deferred minor and correction. **Start here.** It records several places where I was wrong and was corrected.
- `docs/designs/2026-07-24-finish-excel-track1-plan.md` — the plan; its Task 10 section is the owner's checklist.
- `spec/contract.md` — normative; amended in five tasks.
- `src/core/Core.bas` (~3300 lines) — **grep it, never read it whole.** `docs/code-map.md` is badly stale; regenerate with `node scripts/code-map.mjs` before relying on it.

## Verification reality

There is **no headless Excel, no VBA compiler available to the agent, and no test suite** (TDD is formally off for this repo — do not propose adding one). `pwsh build/build.ps1` and `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` are **structural only**: they cannot resolve call targets, see a wrong argument count, or catch a reserved-word identifier before the new gate. **The owner's manual import is the real gate**, and this session is the proof — it found a compile error that four review rounds and two automated gates all passed.

Build the fixture bundle with:
`pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas`
(the comma form works from any shell now; a space after a comma is also handled).

## Post-ship watch list

Beyond the four items: ~a third of the 31 deferred minors in the ledger were already stale when triaged; five procedures walk the registry to find one action by id (one `FindAction(id)` collapses four); manual mode has no category concept, so at catalog scale one placement code discards the auto-layout for every other button; `AddText` is orphaned; `HasAnyTier`/`HasAdvanced` are named backwards; `"hide"` is five independent literal comparisons where one `Const` would do.
