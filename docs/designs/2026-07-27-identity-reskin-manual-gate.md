# Manual gate — sheetforge identity reskin

**Date:** 2026-07-27
**Branch:** `feat/sheetforge-identity` (stacked on `feat/action-catalog-first` → `fix/core-hardening-and-checklist-gate`)
**Closes:** #11
**Structural gates green:** `pwsh build/build.ps1` → `Core + 2 package(s)`, `Sub 62/62`,
`Function 129/129`, 191 unique procedures; `node scripts/vba-syntax-check.mjs
dist/MacrosBundle.bas` → no diagnostics.

Presentation only. No stored configuration changes meaning, and no behaviour changes —
which is exactly why this gate is a **looking** gate. The structural gates cannot see a
colour.

**Passed a blind dry-run** before reaching you. Excel-behaviour claims that remain
unverified are marked `ASSUMPTION`.

---

## Status — NOT RUN

| Item | State |
|---|---|
| 1. Setup tab wears the identity | Not run |
| 2. Macros pages match, and stay distinguishable | Not run |
| 3. Semantic colours survived | Not run |
| 4. Legible on a dark Excel theme | Not run |
| **Total 0/4** | **Not run** |

### The reference

`docs/designs/sheetforge-design-proof.png`. **Only its colours are in scope.** It is a
mobile app concept with a store, `$49` packs, a cart and cloud sync; none of that exists or
is being built. Do not judge these items against its layout.

| Role | Value | Where it should appear |
|---|---|---|
| `ForgeRed` | `RGB(187, 30, 32)` | the Setup **tab strip only** |
| `ForgeNavy` | `RGB(17, 30, 46)` | headings, titles, card category headers |
| `ForgeSteel` | `RGB(49, 60, 74)` | help/secondary text, the **Macros** tab strip |
| `ForgeTint` | `RGB(226, 232, 238)` | band and card header fill |
| `ForgeRule` | `RGB(214, 221, 228)` | hairline borders |

---

## Setup

New workbook, **saved as macro-enabled `P.xlsm` before you start** (item 4 may have you
reopen it; an unsaved workbook loses the whole build). Alt+F11 → *File ▸ Import File* →
`dist/MacrosBundle.bas` → *Debug ▸ Compile VBAProject* → Alt+F8 → `Core_BuildSetup`. Set
all eight `Availability` cells to `Button`, then `Core_FinalizeSetup` → **Yes** (so the
Macros pages actually draw buttons to look at).

**First, confirm you are looking at this build at all.** A stale import is the single
likeliest way to "fail" every item below:

```sh
grep -c 'ForgeRed\|ForgeNavy' dist/MacrosBundle.bas    # expect non-zero
```

If that is zero, the bundle predates the reskin — re-import before going further. (Checking
that the *old* purple literals are absent proves nothing on its own: a build where the
reskin never landed and the colours came from somewhere else would also be empty.)

**Read colours back numerically; do not judge them by eye.** Every item below gives a
number, because "is that navy or black?" and "is that the same blue as the other sheet?"
are not answerable by looking. In the Immediate window (Ctrl+G):

| What | Expression | |
|---|---|---|
| a cell's font | `?Selection.Font.Color` | select the cell first |
| a cell's fill | `?Selection.Interior.Color` | |
| a cell's bottom border | `?Selection.Borders(xlEdgeBottom).Color` | |
| **the sheet tab** | `?ActiveSheet.Tab.Color` | activate that sheet first |

These print a **packed Long**, not an RGB triple, because `RGB(r,g,b)` = `r + g*256 +
b*65536`:

| | packed |
|---|---|
| `ForgeRed` | `2105019` |
| `ForgeNavy` | `3022353` |
| `ForgeSteel` | `4865073` |
| `ForgeTint` | `15657186` |
| `ForgeRule` | `14998998` |
| yellow input cell | `13431551` |
| caution red | `150` |
| white | `16777215` |

**A note on wording:** the coloured object is the **individual sheet tab** at the bottom of
the window — the little labelled rectangle — not the whole tab bar. Excel also renders an
*active* sheet's tab as a washed-out gradient, so the colour you see is not the colour that
was set. That is precisely why every tab claim below is checked with a number.

---

- [ ] **The Setup sheet wears the identity.** Activate the Setup worksheet, then check each
      of these — the numbers are the item, the descriptions are only so you know what you
      are looking at.

      | Select this | Run | Expect |
      |---|---|---|
      | (just activate the sheet) | `?ActiveSheet.Tab.Color` | `2105019` — **red** |
      | B1, `Toolkit Setup` | `?Selection.Font.Color` | `3022353` — navy |
      | the `COMPOSE YOUR BUTTONS` header cell | `?Selection.Font.Color` | `3022353` |
      | the same cell | `?Selection.Interior.Color` | `15657186` — pale steel |
      | a grey help line under a band header | `?Selection.Font.Color` | `4865073` — steel |
      | a banded row | `?Selection.Borders(xlEdgeBottom).Color` | `14998998` — hairline |

      The sheet-tab check is the headline claim of the whole change, and it is the one you
      **cannot** do by eye: Excel washes out the active sheet's tab into a gradient, so the
      colour rendered is not the colour set.

      Failure: any number that does not match. The old purple values were `10498160`
      (`112,48,160`), `6561846` (`54,32,100`), `7882330` (`90,70,120`) and `15784413`
      (`221,217,240`); seeing one of those means the reskin did not land on that surface.

- [ ] **The Macros pages match, and the two surfaces stay apart.** Go to `Setup Menu`, then
      `Macros 1`.

      | Sheet | Select this | Run | Expect |
      |---|---|---|---|
      | `Setup Menu` | (activate it) | `?ActiveSheet.Tab.Color` | `4865073` — **steel** |
      | `Setup Menu` | B2, `Setup Menu` title | `?Selection.Font.Color` | `3022353` |
      | `Macros 1` | (activate it) | `?ActiveSheet.Tab.Color` | `4865073` |
      | `Macros 1` | a card category header cell | `?Selection.Font.Color` | `3022353` |
      | `Macros 1` | the same cell | `?Selection.Interior.Color` | `15657186` |

      Two distinct failures, and they mean opposite things:
      - **`7949855` (the old blue, `31,78,121`) anywhere** → the reskin covered only the
        Setup sheet and left the Macros pages behind.
      - **A Macros tab reading `2105019`** → it went too far. Setup must be red and the
        Macros surfaces steel. The old purple-vs-blue accidentally told you which surface
        you were on, and that distinction was meant to survive.

- [ ] **Colours that mean something still mean it.** All on the Setup sheet.

      - **Input cells are still yellow.** Select any yellow input cell →
        `?Selection.Interior.Color` → **`13431551`**. Yellow means "fill this in"; it must
        not have become steel or navy.
      - **The caution line is still its own red.** Select the cell holding the italic line
        `Do not edit column A (machine keys). Blank rows in a list are ignored.` — it is in
        column B, expected around row 3 — and run `?Selection.Font.Color` → **`150`**. If it
        prints `2105019` the brand red was folded into a warning, and the warning stops
        reading as one.
      - **The brand red appears nowhere else on the grid.** Do not sweep this by eye.
        *Home ▸ Find & Select ▸ Find ▸ Options ▸ Format ▸ Font ▸ Color* → set
        `RGB(187, 30, 32)` → **Find All**. Expect `We couldn't find anything to replace...`
        / no cells found. The brand red is a fill on the sheet tab only.

- [ ] **Legible on a dark Excel theme.** The palette came from a dark-first proof and lands
      on an explicitly white sheet; nothing before this item checks the two together.

      **Save `P.xlsm` first.** Then switch Excel to its dark theme — *File ▸ Options ▸
      General ▸ Office Theme ▸ Black* (`ASSUMPTION`: that is the path on current Excel for
      Windows; some builds put it under *File ▸ Account ▸ Office Theme*, and older ones
      offer `Dark Gray` instead of `Black`). The theme applies live — **no reopen needed**,
      and reopening an unsaved workbook would lose the build.

      Then re-take three numbers under the dark theme, because the point of this item is
      that Excel has *not* quietly re-mapped anything:

      | Select this | Run | Expect |
      |---|---|---|
      | an empty body cell on Setup | `?Selection.Interior.Color` | `16777215` — still white |
      | Setup sheet | `?ActiveSheet.Tab.Color` | `2105019` — still red |
      | `Macros 1` | `?ActiveSheet.Tab.Color` | `4865073` — still steel |

      Then look, and record what you see even if it passes: is every heading, help line and
      card header still comfortably readable against the white body? Are the two sheet tabs
      still tellable apart against a dark chrome, or do both read as "dark smudge"?

      **This is the item most likely to produce a real finding**, and it is the one item
      whose verdict is partly a judgement call. If the numbers hold but it looks wrong, say
      so — that is a legitimate fail and the palette should move.
