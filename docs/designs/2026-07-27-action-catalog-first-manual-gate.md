# Manual gate — Action catalog first on the Setup tab

**Date:** 2026-07-27
**Branch:** `feat/action-catalog-first` (stacked on `fix/core-hardening-and-checklist-gate`)
**Implements:** the ordering half of
`docs/designs/2026-07-26-setup-surface-progressive-disclosure-design.md`
**Structural gates green:** `pwsh build/build.ps1` → `Core + 2 package(s)`, `Sub 62/62`,
`Function 124/124`, 186 unique procedures; `node scripts/vba-syntax-check.mjs
dist/MacrosBundle.bas` → no diagnostics.

Presentation-only change: no reader is order-dependent, so no stored configuration changes
meaning. This gate exists because the structural gates cannot see a rendered sheet.

**Passed a blind dry-run** (`CLAUDE.md` pipeline), which found six defects in the first
draft — including one that made item 3 unreachable and one that would have left you with a
duplicate module and an "Ambiguous name detected" compile error. Excel-behaviour claims
that remain unverified are marked `ASSUMPTION`.

---

## Status — NOT RUN

| Item | State |
|---|---|
| 1. The Setup tab opens on the Action catalog, Chains second | Not run |
| 2. Nothing is clipped and no boundary gained or lost a blank row | Not run |
| 3. An old-order `_Config` still reads correctly | Not run |
| **Total 0/3** | **Not run** |

---

## Setup

New workbook `R.xlsm`, saved macro-enabled. Alt+F11 → *File ▸ Import File* →
`dist/MacrosBundle.bas` → *Debug ▸ Compile VBAProject* → Alt+F8 → `Core_BuildSetup`.

**Then switch actions on, or there will be nothing to order.** On a fresh build every
action arrives `Off` (v0.6.0), so every settings field is relegated to the collapsed band
and **no package settings band renders at all**. Items 1 and 2 are about where the catalog
sits *relative to the settings bands*, so those bands have to exist:

1. In the Action catalog, set **all eight** `Availability` cells to `Button`.
2. Alt+F8 → `Core_FinalizeSetup` → **Yes**.
3. Alt+F8 → `Core_BuildSetup`. A prompt appears reading
   `A 'Setup' tab already exists. Rebuild it?` — **the default button is No, so click Yes
   deliberately.**

**Re-importing the bundle later:** the module is named `MacrosBundle`, and Excel's import
never replaces a same-named module — it adds a second one as `MacrosBundle1`, duplicating
all 186 procedures, which produces `Ambiguous name detected` on compile and doubled Alt+F8
entries. **Always remove the old module first:** in the Project Explorer, right-click
`MacrosBundle` → *Remove MacrosBundle* → **No** when asked whether to export.

---

- [ ] **The Action catalog is first, Chains second, settings bands after both.**
      Scroll to the top of the Setup tab.

      Expected, top to bottom:
      1. the page header rows (`Toolkit Setup` and its instruction line),
      2. the `COMPOSE YOUR BUTTONS` heading,
      3. a grey help line beginning `Turn actions on in the Action catalog below`,
      4. the **Action catalog** — column C reads `Action` in bold, then eight
         `Caption [id]` rows with yellow `Availability` dropdowns in column D,
      5. the **Chains** block — bold `Button name` / `Steps (one token per line)` /
         `Page code` headers,
      6. only then the `IMPORT / EXPORT` band and its `GENERAL` / `EXPORT TABS` /
         `IMPORT FILES` / `REF_MASTER` sub-headings.

      Failure: Chains appears above the Action catalog, or either appears below
      `IMPORT / EXPORT`. That is the pre-change order.

- [ ] **Nothing is clipped, and no boundary gained or lost a blank row.**
      The two risks the reviews ranked highest.

      **Clipping.** The Chains block's `Steps (one token per line)` cells wrap to multiple
      lines. Type two tokens on separate lines into the first Chains row's Steps cell —
      Alt+Enter for the line break (`ASSUMPTION`: Alt+Enter is the in-cell newline on
      Windows Excel). Expected: the row is tall enough to show both lines.

      *If the row does not grow, do not fail this item yet.* A row whose height was set
      explicitly will not auto-grow regardless of where AutoFit is anchored. Widen the
      column or re-run AutoFit by hand first. **This item is about AutoFit ordering, not
      row-height locking** — it fails only if the compose rows are systematically short
      while the package-field rows below them fit their content correctly.

      **Spacing.** Expected exactly **one** blank row at each boundary: between the Chains
      block and `IMPORT / EXPORT`, and between `REF_MASTER` (the last sub-section) and
      `SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON`. Failure: two blank rows at either
      boundary, or none.

- [ ] **A configuration saved in the OLD order still reads correctly.** The one item that
      can reveal a real defect rather than a cosmetic one.

      **First build the v0.6.0 bundle.** It is not in `dist/` — that directory is
      gitignored and currently holds only v0.3.0, v0.4.0 and v0.5.0. The repo has no tags,
      so the commit is named directly: `812d88c` is the last commit with `VERSION` at
      0.6.0. In PowerShell from the repo root:

      ```powershell
      git worktree add ../wt-v060 812d88c
      pwsh ../wt-v060/build/build.ps1
      Copy-Item ../wt-v060/dist/MacrosBundle.bas dist/MacrosBundle-v0.6.0.bas
      git worktree remove ../wt-v060 --force
      ```

      **Then build the old-order workbook.** New workbook `S.xlsm`, import
      `dist/MacrosBundle-v0.6.0.bas`, `Core_BuildSetup`. Every action arrives `Off` — that
      is v0.6.0's default. Set a deliberate three-way mix **explicitly**, changing all
      eight cells:

      - `Remove Duplicates [clean.dupes]` → **`Off`** (leave it, but confirm it reads `Off`)
      - `Trim & Clean Text [clean.trim]` → **`Chains only`**
      - **the other six → `Button`** (they will be `Off` until you change them)

      `Core_FinalizeSetup` → **Yes**. That `_Config` now has Chains above the catalog.

      **Now upgrade.** Remove the `MacrosBundle` module as described in Setup, import
      `dist/MacrosBundle.bas`, then Alt+F8 → `Core_BuildSetup` → **Yes**.

      Expected: the tab comes back in the **new** order carrying the **old** answers —
      `clean.dupes` `Off`, `clean.trim` `Chains only`, the other six `Button`.

      Three ways it can fail, and they mean different things:
      - **Answers reset to all `Button`** — a reader took the legacy path. Order-dependence
        in the marker logic.
      - **Answers reset to all `Off`** — this is *not* necessarily an ordering bug. v0.7.0
        added "a missing marker must be earned" and "no catalog block fails closed"; a
        v0.6.0 `_Config` that trips either path reads as fresh. Check for a refusal message
        first, and report which one you saw.
      - **The mix survives** — pass.

      **This item cannot pass on a broken build:** the expected result is `Off` +
      `Chains only` + six `Button`, which is neither the fresh default (all `Off`) nor the
      legacy default (all `Button`). Only correct key-based reading of an old-order store
      produces it.
