# W1 manual gate — the ingest→clean wave (7 actions)

> **For the owner.** This is written for someone who has not read the code. Every item
> states the observable result first, then what a failure looks like, then why it matters.
> Tick an item only against a real run. If you skip one, write **UNTESTED** next to it
> rather than leaving it blank — a blank tick is indistinguishable from a pass nobody did,
> and eight of track 1's fifteen items are still unrun for exactly that reason.
>
> **Run each section in order, and do not credit a "nothing happened" item on its own.**
> Items 1.4, 3.4, 4.2 and 4.3 all expect *no change* — and a completely broken build that
> does nothing when you click would pass every one of them. Each is only evidence once the
> positive-path item in the same section (1.1, 3.1, 4.1, 4.4) has already passed. They are
> marked **[pairs with N]** below.

**Build under test:** `dist/MacrosBundle.bas`, built from branch `feat/w1-ingest-clean`.
Rebuild it with `pwsh build/build.ps1` before you start.

**Also needed:** `tools/GateAudit.bas`, from the same branch. It is **not** part of the
bundle — nothing in the build reads `tools/` — so it is a second, separate import used
only by item 0.4. If it is missing from your checkout, skip 0.4 and mark it UNTESTED.

---

## 0. Setup — do this first, in this order

- [ ] **0.1 — Remove any existing `MacrosBundle` module BEFORE importing.**
  In the VBE (Alt+F11), right-click any module named `MacrosBundle` or `MacrosBundle1` in
  the Project pane and choose *Remove*. Say **No** when it offers to export.
  **Then** File ▸ Import File ▸ `dist/MacrosBundle.bas`.

  *Failure looks like:* a second module appears named `MacrosBundle1`, and the next step
  reports `Ambiguous name detected`.
  *Why:* Excel's import never replaces a same-named module — it adds a numbered copy,
  duplicating ~216 procedures. Confirmed by hand on 2026-07-28.

- [ ] **0.2 — Debug ▸ Compile VBAProject reports nothing.**
  The menu item greys out when compilation succeeds.

  *Failure looks like:* a dialog naming a line, with that line highlighted.
  *Why:* this is the only real compiler in the pipeline. `build.ps1` does structural checks
  and `vba-syntax-check.mjs` parses, but neither is VBA.

- [ ] **0.3 — Run `Core_BuildSetup` (Alt+F8), then Finalize.**
  The Setup tab appears. Set the seven new actions to `Button` in the Action catalog block,
  then Finalize. The Macros pages rebuild with the new buttons on them.

  *Failure looks like:* an action missing from the catalog, or Finalize refusing.
  *Why:* since v0.6.0 every action arrives `Off`, so nothing below runs until you do this.

- [ ] **0.4 — Import `tools/GateAudit.bas` and run `GateAudit_Run`.**
  The report names all fifteen action ids as present in the catalog.

  *Failure looks like:* a `FAIL` line naming an id that did not reach the Setup surface.
  *Why:* it automates the "did the action appear at all" question for every action at once,
  so the items below can be about behaviour instead.

---

## 1. `import.combine_worksheets`

Make a scratch workbook with three data sheets. Give them **the same headers in a
different column order** — e.g. Sheet1 `Store | Date | Amount`, Sheet2 `Date | Amount |
Store`. Put two or three data rows on each.

- [ ] **1.1 — It creates one new sheet named from Setup (default `Master`) holding every
  source row, with each value under its own header.**
  `Amount` values from Sheet2 land in the `Amount` column, not in the column that happened
  to be third.

  *Failure looks like:* values in the wrong columns, or a second header row in the middle
  of the data.
  *Why:* alignment is by header text, not position. This is the whole feature — a
  position-based combine silently scrambles a real month-end folder.

- [ ] **1.2 — Run it a second time. It refuses, names the existing sheet, and changes
  nothing.**
  Message: *"A sheet named 'Master' already exists in this workbook, and this action will
  not overwrite it."* The existing `Master` still holds the first run's data.

  *Failure looks like:* `Master` silently replaced, or a `Master (2)` appearing.
  *Why:* the reference snippet this replaces opened with
  `ThisWorkbook.Worksheets("Master").Delete`. That would destroy whatever you had there.
  This is the single most destructive line in the source material and it must not exist.

- [ ] **1.3 — A user sheet named `Macros Data` IS included in the combine.**
  Rename one source sheet to `Macros Data` and re-run (delete `Master` first). Its rows
  appear.

  *Failure looks like:* the sheet is silently skipped and its rows are missing.
  *Why:* the engine reserves the generated pages `Macros 1`, `Macros 2`, … A prefix test
  that skipped anything starting `Macros ` would eat a user's own data sheet without
  saying so.

- [ ] **1.4 [pairs with 1.1] — Your `Setup` and `Macros <n>` tabs are NOT in the result.**
  *Failure looks like:* Setup's key column or a page of buttons appearing as data rows.
  *Why:* engine-owned sheets are not data.

---

## 2. `import.combine_workbooks`

Put two or three workbooks in a folder. Set `Folder to read workbooks from` on Setup, or
leave it blank to use this workbook's own folder.

- [ ] **2.1 — It combines the first sheet of each file, and the source files are unchanged
  afterwards.**
  Re-open one source and confirm it is untouched and was not saved.

  *Failure looks like:* a source file modified, or Excel prompting to save one on close.
  *Why:* each is opened read-only and closed without saving.

- [ ] **2.2 — Point it at a folder that does not exist. It says so by name and changes
  nothing.**
  *Failure looks like:* a raw `Run-time error '76'` / `Path not found` dialog, or the
  generic "Action failed" engine error.
  *Why:* a wrong path in a config cell is an ordinary user mistake and should read as one.

- [ ] **2.3 — ⚠️ Have a colleague (or a second Excel window) hold one source workbook OPEN,
  then run it.**
  Record exactly what happens.

  *Expected — but NOT verified, and the reason this item exists:* Excel writes a hidden
  lock file named `~$Something.xlsx` beside an open workbook, and the folder scan matches
  `*.xls*`. If the action tries to open that lock file it will fail and abandon the whole
  combine.
  *Why this is on the list:* in a shared finance folder somebody usually has a file open.
  If this fails, it is the most likely real-world failure in the wave and the fix is one
  line (skip names starting `~$`). **Write down what you see even if it works.**

---

## 3. `import.delimited_file`

- [ ] **3.1 — Pick a CSV whose text fields contain commas inside quotes. The columns split
  correctly.**
  A field like `"Smith, John"` stays in one cell.

  *Failure looks like:* everything in column A, or `Smith` and `John` in separate columns.
  *Why:* the action parses with an explicit separator and quote character from Setup rather
  than letting Excel guess from the file extension and your Windows list-separator setting.

- [ ] **3.2 — The new sheet is named after the file.**
  *Failure looks like:* `Sheet4`, or a name with `/` or `[` in it surviving.

- [ ] **3.3 — Import the same file twice. The second run refuses by name and creates
  nothing.**
  *Failure looks like:* a second sheet appearing, or the first being overwritten.

- [ ] **3.4 [pairs with 3.1] — Press Cancel on the file picker. Nothing happens at all.**
  No new sheet, no dialog, no error.

  *Failure looks like:* an error dialog, or an empty sheet created before the cancel was
  noticed.
  *Why:* cancel must be reachable and must cost nothing.

---

## 4. `clean.remove_blanks` — the destructive one

Build a range with some entirely blank rows in the middle, and **at least one row that is
blank inside your selection but has data outside it** (e.g. select `B2:M40`, and put a
value in `A17` on an otherwise blank row 17).

- [ ] **4.1 — It shows a count BEFORE deleting anything, and names what it will not touch.**
  The dialog states how many blank rows and columns it will delete, and separately that N
  rows are blank inside your selection but hold data outside it and were left alone.

  *Failure looks like:* deletion with no dialog, or a dialog with no counts.
  *Why:* this action ships without the `backup.snapshot` restore point that its deferred
  siblings are waiting for. The count IS the safety story.

- [ ] **4.2 [pairs with 4.4] — Answer No. Nothing is deleted.**
  Every row and column is still there.

- [ ] **4.3 [pairs with 4.4] — ⚠️ Answer by clicking the title-bar X rather than a button.
  Nothing is deleted.**
  *Expected — but NOT verified:* the dialog has Yes and No only, and the X is assumed to
  return the same as No.
  *Why this is on the list:* the whole safety argument for a two-button dialog is that its
  escape hatch is the safe answer. If the X does something else, this action needs a third
  button. **This item is why the dialog is shaped the way it is — please actually try it.**

- [ ] **4.4 — Answer Yes. The blank rows go; the row with data outside your selection
  SURVIVES.**
  Row 17 in the example above is still there, still holding `A17`.

  *Failure looks like:* row 17 gone and `A17`'s value with it.
  *Why:* a whole-row delete based on a partial selection would destroy data the user never
  selected and never saw. This is the item that proves it does not.

---

## 5. `clean.unmerge_fill`

Make a range with merged blocks — some wholly inside what you will select, and **one that
crosses the edge of your selection**.

- [ ] **5.1 — Blocks inside the selection are unmerged, and every cell they covered holds
  the block's value.**
  *Failure looks like:* unmerged but only the top-left cell populated — the rest blank.
  *Why:* blank cells under a former merge break sorting and lookups, which is the reason
  to run this at all.

- [ ] **5.2 — A block crossing the selection edge is LEFT MERGED, and the report says how
  many were skipped and why.**
  *Failure looks like:* it is unmerged too, and cells outside your selection change.

- [ ] **5.3 — A merged block whose top-left cell held a FORMULA still shows that formula
  in the formula bar afterwards; the other cells hold its computed value.**
  Click the anchor cell and look at the formula bar.

  *Failure looks like:* the formula replaced by its value.
  *Why:* flattening a formula is silent and unrecoverable once saved.

---

## 6. `clean.standardize_formats`

- [ ] **6.1 — With the `Column formats` table on Setup EMPTY, it refuses and says where to
  configure it.**
  Message points at the Setup tab. Nothing on the sheet changes.

  *Failure looks like:* it runs and applies some format you never configured — most likely
  `mm/dd/yyyy` to column A.
  *Why:* this is the one item most likely to catch a regression. Core resolves an empty
  stored table to the field's *declared defaults*, so if anyone ever adds demo rows to this
  field, this refusal silently becomes unreachable and a first-time user gets a date format
  applied to whatever they picked. **If this item passes, that rule is still intact.**

- [ ] **6.2 — Fill in a row addressing a column by LETTER (e.g. `C` / `$#,##0.00`).
  Only that column's display changes; the stored values do not.**
  Click a cell and confirm the formula bar still shows the raw number.

- [ ] **6.3 — Fill in a row addressing a column by HEADER NAME. It finds the right column
  regardless of position.**

- [ ] **6.4 — Add a row naming a column that is not in your selection. It is reported by
  name in the completion message.**
  *Failure looks like:* silently skipped.
  *Why:* a renamed header is the ordinary way this table goes stale, and a silent skip
  looks exactly like a format that was applied.

---

## 7. `clean.text_to_numbers`

- [ ] **7.1 — Text like `$1,234.56` becomes a real number, right-aligned, usable in
  `=SUM()`.**
  *Failure looks like:* still left-aligned; `SUM` returns 0.

- [ ] **7.2 — Text that is not a number is left exactly as it was.**
  *Failure looks like:* a cell blanked or turned into `0`.

- [ ] **7.3 — A formula cell is not touched.**
  Put a formula in the range and confirm the formula bar still shows it.

- [ ] **7.4 — Note what `12%` becomes.**
  **Known and documented:** it becomes the number `12`, not `0.12`. The shared coercion
  strips the percent sign.

  *Why this is here rather than fixed:* the action calls the same coercion the import
  action uses, deliberately, rather than growing a second implementation. Changing it would
  change every existing import. Tick this item if `12%` → `12`; tell us if that is the
  wrong call for your data.

---

## 8. Error recovery — the one that gets blamed on the next macro

- [ ] **8.1 — Force a failure, dismiss the error, then confirm Excel is not left broken.**

  **Use this trigger:** set `Folder to read workbooks from` on Setup to a path that does
  not exist, Finalize, and run `import.combine_workbooks`.

  *Do not use "rename a `.txt` to `.xlsx`".* Excel often detects the mismatch and either
  repairs the file silently or shows a prompt that gets auto-answered because the action
  suppresses alerts while it runs — so it may never raise an error at all, and you would be
  testing nothing. A missing folder fails deterministically.

  Note that a missing folder is caught deliberately and reports as itself (item 2.2), so
  for the state check below use a **second** trigger that reaches the engine: point the
  folder at a real folder containing a zero-byte file renamed to `.xlsx`.

  Before running, note these four in the VBE Immediate window (Ctrl+G, type each and press
  Enter):

  ```
  ?Application.Calculation
  ?Application.ScreenUpdating
  ?Application.EnableEvents
  ?Application.DisplayAlerts
  ```

  After the failure, check all four again. **They must read the same as before.**

  *Failure looks like:* `Calculation` still `-4135` (manual), or `ScreenUpdating` still
  `False`.
  *Why:* a macro that leaves calculation on manual makes every later formula stop updating.
  The user blames whatever they run next, not this. This is the most expensive failure mode
  to diagnose and the cheapest to check.

- [ ] **8.2 — After that failure, no half-built sheet is left behind.**
  No `Master` (or partially populated sheet) in the workbook.

  *Failure looks like:* a sheet with some rows from the files that were readable.

- [ ] **8.3 — After a failed `import.delimited_file`, no source file is left open.**
  Check the Window menu / taskbar for a stray workbook.

---

## 9. Chain them — the reason this wave exists

**How to build a chain.** On the Setup tab, find the **Chains** block. In one chain's
`Steps` cell, put one step per line, each as `Label [action.id]` — the bracketed id is what
actually runs and the label is only the caption. Use exactly this:

```
Combine All Worksheets [import.combine_worksheets]
Remove Blank Rows & Columns [clean.remove_blanks]
Trim & Clean Text [clean.trim]
Convert Text to Numbers [clean.text_to_numbers]
```

Give the chain a name, Finalize, and its button appears on a Macros page.

**Expect to be interrupted — this is correct, not a bug.** Running this chain will stop for
input **four** times: a range picker for each of `clean.remove_blanks`, `clean.trim` and
`clean.text_to_numbers`, plus the Yes/No delete confirmation inside `clean.remove_blanks`.
Chain mode suppresses each step's *completion* message, not its questions. It is deliberate
that a destructive confirmation is never suppressed, and range-scoped actions ask for their
range at run time by design rather than storing it.

- [ ] **9.1 — Build the chain above, run it from its single button, and get a clean dataset
  at the end.**
  You answer four prompts along the way (three range pickers, one Yes/No). At the end you
  get **one** summary, not four.

  *Failure looks like:* four separate "complete" dialogs instead of one summary; the chain
  stopping part-way with no message; or a step reporting an error and the chain continuing
  as if it had worked.
  *Why:* these seven were chosen as a coherent workflow rather than a feature sampler
  precisely so this works the day it lands. If the individual actions pass but this does
  not, the wave has not delivered what it was for.

---

## Assumptions this checklist is testing, not asserting

Each of these is a claim about **Excel's own behaviour**, not about this code. None has
been executed anywhere — there is no VBA runner in this project. They are called out so a
failure reads as "the assumption was wrong" rather than "the macro is broken":

| Assumption | Item that tests it |
|---|---|
| A `vbYesNo` MsgBox's title-bar X returns `vbNo` | 4.3 |
| `Workbooks.OpenText` leaves the file as `ActiveWorkbook` | 3.1 |
| `OpenText`'s `Local:=True` coexists with an explicit `TextQualifier` | 3.1 |
| `Application.GetOpenFilename` returns `False` on Cancel | 3.4 |
| `Range.UnMerge` leaves the anchor intact and covered cells empty | 5.1, 5.3 |
| `Application.CountA` accepts a whole-row/column range | 4.1 |
| A macro's row deletion cannot be undone with Ctrl+Z (the dialog says so) | 4.4 |
| `Scripting.FileSystemObject` is available for the folder check | 2.2 |
| Deleting whole rows never renumbers columns | 4.4 |
| The folder scan's `*.xls*` glob also matches `~$` lock files | 2.3 |
| Excel's sheet-name rules match the hardcoded `: \ / ? * [ ]` and 31-character checks | 1.2, 3.2 |

---

## Result

- Items passed: ____ / 33
- Items recorded UNTESTED: ____
- Anything that failed, with what you saw:
