# Manual gate — Core hardening (#8, #12, #13) and the checklist gate (#3)

**Date:** 2026-07-27
**Branch:** `fix/core-hardening-and-checklist-gate`
**Closes:** #8, #12, #13, #3
**Contract:** v2.2.0 (§5 durable replacement, §5 marker absence, §7 `core.*` enforced,
§7 reserved sheet names)
**Structural gates green:** `pwsh build/build.ps1` → `Core + 2 package(s)`, `Sub 62/62`,
`Function 124/124`, 186 unique procedures; `node scripts/vba-syntax-check.mjs
dist/MacrosBundle.bas` → no diagnostics; `SCHEMA_VER` still `"v2.0"`.

**This checklist passed a blind dry-run** (`CLAUDE.md` pipeline) before reaching you. That
pass found 13 defects in the first draft, including one that would have failed you at the
very first item and two that would have scored green on a reverted build. They are fixed
here. Excel-behaviour claims that remain unverified are marked `ASSUMPTION`.

---

## Status — NOT RUN

| Section | Items | State |
|---|---|---|
| 1. Interrupted activate survives (#8) | 0/3 | Not run |
| 2. A removed marker fails closed (#12) | 0/2 | Not run |
| 3. `core.*` is refused (#13) | 0/1 | Not run |
| 4. Part A did not regress | 0/2 | Not run |
| **Total** | **0/8** | **Not run** |

**Two of the eight are preconditions, not evidence.** Items 1.1 and 1.3 pass unchanged on
a pre-fix build — normal activation was never broken. They are here to prove the fixture is
sound before the items that *do* discriminate. Do not read 8/8 as eight proofs; the
load-bearing items are **1.2, 2.1, 2.2, 3.1 and 4.2**.

### What is NOT covered, and why

**A real crash mid-activate is not simulated.** Nothing can force Excel to die between the
delete and the rename. Section 1 instead builds *the state a crash leaves behind* and
verifies recovery from it. The recovery path is tested; the crash is not.

**#12's fix is hardening, not a proof.** A missing marker is now "removed" unless the block
shows the pre-v1.2 shape — no third column anywhere. The documented bypass (rename the
visible `Availability` header) is closed. An attacker who *also* clears the `Placement`
header and every placement code still reaches the legacy path, because at that point no
sheet-level evidence distinguishes the block from a genuine pre-v1.2 one. **Item 2.2 proves
that hole is still open, on purpose.**

---

## Setup

One workbook, **H.xlsm**, built fresh on this build. Item 3.1 needs a second, deliberately
hostile bundle — build it when you reach it, not now.

1. New blank workbook, save as macro-enabled, `H.xlsm`.
2. Alt+F11 → *File ▸ Import File* → `dist/MacrosBundle.bas`.
3. *Debug ▸ Compile VBAProject*. **Do not skip** — a keyword-as-identifier error surfaces
   here and nowhere else. Two new procedures ship in this build (`IsCoreId`,
   `LooksPreV12`).
4. **Paste the diagnostic module.** In the VBE: *Insert ▸ Module*, and paste the whole
   block below into it. Every sheet inspection in this gate runs from Alt+F8, not the
   Immediate window — the Immediate window rejects `Dim`, so undeclared variables there
   are a trap under `Option Explicit`.

   ```vba
   Option Explicit

   ' Gate helper. Not part of the product - delete when the gate is done.
   Public Sub GateDumpSheets()
       Dim s As Worksheet, msg As String
       For Each s In ThisWorkbook.Worksheets
           msg = msg & s.Name & "   visible=" & s.Visible & vbCrLf
       Next s
       MsgBox msg, vbInformation, "Sheets"
   End Sub

   ' Prints the catalog block's marker cell and its column D header BEFORE you
   ' change anything, so the later confirmation cannot be satisfied by a cell
   ' that was always empty.
   Public Sub GateShowMarker()
       Dim w As Worksheet, c As Range
       Set w = ThisWorkbook.Worksheets("_Config")
       Set c = w.Columns(1).Find(What:="k_core_layout", LookIn:=xlValues, _
                                 LookAt:=xlWhole, MatchCase:=False)
       If c Is Nothing Then
           MsgBox "k_core_layout NOT FOUND - stop and report this.", vbCritical
           Exit Sub
       End If
       MsgBox "row=" & c.Row & vbCrLf & _
              "marker (col J) = [" & w.Cells(c.Row, 10).Value & "]" & vbCrLf & _
              "header (col D) = [" & w.Cells(c.Row, 4).Value & "]" & vbCrLf & _
              "header (col E) = [" & w.Cells(c.Row, 5).Value & "]", _
              vbInformation, "Before"
   End Sub

   Public Sub GateTamperMarker()
       Dim w As Worksheet, c As Range
       Set w = ThisWorkbook.Worksheets("_Config")
       Set c = w.Columns(1).Find(What:="k_core_layout", LookIn:=xlValues, _
                                 LookAt:=xlWhole, MatchCase:=False)
       If c Is Nothing Then MsgBox "not found", vbCritical: Exit Sub
       w.Cells(c.Row, 10).ClearContents        ' blank the hidden marker
       w.Cells(c.Row, 4).Value = "Avail"       ' rename the visible header
       MsgBox "Marker blanked, header renamed to 'Avail'.", vbInformation
   End Sub

   ' Clears column E across the whole catalog block - the second half of item 2.2.
   ' Bounds the block by the next key row rather than guessing a row count.
   Public Sub GateClearColumnE()
       Dim w As Worksheet, c As Range, i As Long, lastR As Long, n As Long
       Set w = ThisWorkbook.Worksheets("_Config")
       Set c = w.Columns(1).Find(What:="k_core_layout", LookIn:=xlValues, _
                                 LookAt:=xlWhole, MatchCase:=False)
       If c Is Nothing Then MsgBox "not found", vbCritical: Exit Sub
       lastR = w.Cells(w.Rows.Count, 1).End(xlUp).Row
       n = c.Row
       For i = c.Row + 1 To lastR
           If Len(Trim(CStr(w.Cells(i, 1).Value))) > 0 Then Exit For
           n = i
       Next i
       w.Range(w.Cells(c.Row, 5), w.Cells(n, 5)).ClearContents
       MsgBox "Cleared column E, rows " & c.Row & "-" & n & ".", vbInformation
   End Sub

   Public Sub GateRenameConfigToStaged()
       ThisWorkbook.Worksheets("_Config").Name = "_Config_staged"
       MsgBox "_Config renamed to _Config_staged.", vbInformation
   End Sub
   ```

5. Alt+F8 → `Core_BuildSetup`.
6. On the **Setup** tab, scroll to the **Action catalog** block (the rows under the
   `COMPOSE YOUR BUTTONS` heading, each reading `Caption [id]` in column C with a yellow
   `Availability` dropdown in column D). **Eight actions are listed** — three `clean.*`
   and five `importexport.*`. Set a deliberate, non-default mix:
   - `Remove Duplicates [clean.dupes]` → **`Off`**
   - `Trim & Clean Text [clean.trim]` → **`Chains only`**
   - the other **six** → **`Button`**
7. Alt+F8 → `Core_FinalizeSetup` → **Yes**.

That mix is the fixture. Every item below asks whether it survived.

**Sheet visibility values you will see:** `-1` = visible, `0` = hidden, **`2` = very
hidden**. `_Config` is stored very hidden, so it reads **`2`**. (`xlSheetVeryHidden` is 2,
not -2 — the first draft of this gate had that backwards.)

---

## 1. An interrupted activate does not lose the configuration (#8)

- [ ] **Precondition — a normal Finalize leaves no staging sheet behind.**
      Alt+F8 → `GateDumpSheets`.

      Expected: a line `_Config   visible=2`, the `Setup` tab, the `Macros` page(s), and
      **no line reading `_Config_staged`**.

      Failure: a `_Config_staged` line means the swap did not complete. Every later item in
      this section is invalid until that is understood.

      *Precondition, not evidence — this passes on a pre-fix build too.*

- [ ] **The state a crash leaves behind is recovered, not discarded.** This is the whole of
      #8. Alt+F8 → `GateRenameConfigToStaged`.

      The workbook now holds a complete configuration under the staging name and **no
      `_Config`** — exactly what a crash between the delete and the rename leaves.
      *(`ASSUMPTION`: a very-hidden sheet can be renamed without unhiding it. If
      `GateRenameConfigToStaged` raises, unhide first — `ThisWorkbook.Worksheets("_Config")
      .Visible = -1` — and record the correction here.)*

      Now Alt+F8 → `Core_BuildSetup`. **A prompt appears** because the Setup tab already
      exists, titled `Create Setup Tab` and asking whether to rebuild it — answer **Yes**.
      *(If no prompt appears at all, say so; that is itself a result, not a non-event.)*

      Expected, on the rebuilt Setup tab: **the fixture mix is back** — `clean.dupes` reads
      `Off`, `clean.trim` reads `Chains only`, the other six read `Button`. Then Alt+F8 →
      `GateDumpSheets`: `_Config` is back at `visible=2`, and `_Config_staged` is **gone**.

      Failure: every `Availability` cell reads `Off` and the staging sheet is still listed.
      That is the pre-fix behaviour — `CfgSheet` found no `_Config`, reported "never
      configured", and re-seeded Setup with defaults **on top of a config sitting right
      there**. That silent overwrite is the real data loss in #8; the missing sheet is only
      how it starts.

      **This item cannot pass on a broken build.** The expected result is the three-way
      fixture mix. All-`Off` is the fresh default and all-`Button` is the legacy default;
      neither produces `Off` + `Chains only` + six `Button`. Only genuine recovery does.

- [ ] **Precondition — Finalize still works after the recovery.**
      Alt+F8 → `Core_FinalizeSetup` → **Yes**. Expected: a dialog reading
      `Configuration activated.` Then `GateDumpSheets`: `_Config` at `2`, no
      `_Config_staged`.

      Failure: a run-time error **1004** whose text is `That name is already taken. Try a
      different one.`, or an orphaned `Setup (2)` tab. Either means the swap collided with
      something the recovery left behind. *(`ASSUMPTION`: that is Excel's own 1004 wording,
      not a string this bundle emits — the exact phrasing may vary by Excel version. The
      run-time error number is the reliable part.)*

      *Precondition, not evidence.*

---

## 2. A removed marker fails closed (#12)

- [ ] **Renaming the visible header no longer promotes a tampered block to legacy.**

      First, Alt+F8 → `GateShowMarker`. **Before you change anything**, confirm the dialog
      shows a **non-empty** `marker (col J)` (it should read `v2.0`) and
      `header (col D) = [Availability]`. If the marker is already empty or column D is not
      `Availability`, stop — the rest of this item would be meaningless, because you would
      be "confirming" a cell that was never populated.

      Then Alt+F8 → `GateTamperMarker`, and Alt+F8 → `Core_BuildSetup`.

      Expected: the rebuild **stops and shows a message** rather than proceeding. The
      message contains this sentence (one continuous sentence on screen; wrapping depends
      on the dialog width):

      ```
      The hidden format marker for the Action catalog block is missing, even though the block still has this version's third column, so saved Availability choices cannot be trusted.
      ```

      followed by a prompt about rebuilding the Action catalog block. **The Setup tab is
      not silently re-seeded.**

      Failure: the Setup tab rebuilds with no such message and **`clean.dupes` reads
      `Button`** — you set it to `Off`. That is the #12 bug: the block took the legacy path,
      where every absent id reads `Button`, and a hand-renamed header was enough to
      re-arm destructive actions.

      Why it now holds: the discriminator is no longer the caption in column D. It is
      whether column **E** is empty across the whole block (`LooksPreV12`). A v1.2+ block
      writes the `Placement` header there; a genuine v0.3.0 block never wrote that column
      at all — verified against the v0.3.0 bundle's own renderer, not assumed.

- [ ] **The residual hole is real — confirm it, do not fix it here.**
      Continuing from the tampered state, Alt+F8 → `GateClearColumnE`, then Alt+F8 →
      `Core_BuildSetup`.

      Expected: the legacy path **is** taken this time. The discriminator is
      **`clean.dupes`, which must now read `Button`** — you set it to `Off`, and only the
      legacy default turns an absent or unreadable value into `Button`. Do not judge this
      by the other actions; six of them were already `Button` and prove nothing.

      **That is the correct, documented outcome, not a failure.** With the marker gone
      *and* the third column gone, the block is indistinguishable from a genuine pre-v1.2
      one by any sheet-level evidence, and contract §5 requires a genuine pre-v1.2 store to
      read as legacy. Tick this item if `clean.dupes` reads `Button`, and record it: the
      point is that reaching this now costs the destruction of every placement code rather
      than one renamed caption.

      `H.xlsm` is deliberately corrupt from here. Close it **without saving**.

---

## 3. A package cannot occupy the `core.*` namespace (#13)

- [ ] **Finalize refuses a `core.legacy` action id.** This needs a hostile build; nothing
      that ships declares one. In `src/packages/PkgClean.bas`, immediately after the
      `clean.filldown` action line, add:

      ```vba
      Core_AddAction reg, "core.legacy", "Hostile", "PkgClean_TrimClean", "temporary - gate item 3.1"
      ```

      Then `pwsh build/build.ps1`, import the rebuilt `dist/MacrosBundle.bas` into a **new**
      workbook, Alt+F8 → `Core_BuildSetup`, then Alt+F8 → `Core_FinalizeSetup`.

      Expected: activation is refused, and the report contains

      ```
      RESERVED: action id 'core.legacy' uses core.*
      ```

      followed by a line ending `problem(s) above must be fixed before activating.` There
      is no "Activate this configuration now?" prompt — the refusal comes instead of it.

      Failure: it activates. The sentinel key is then occupied by a real action, `AvailFor`
      finds the legacy flag present on every map it builds, and **every action you did not
      explicitly list resolves to `Button` instead of `Off`** — Part A silently reverted on
      that deployment.

      **Revert before doing anything else:** `git checkout src/packages/PkgClean.bas` then
      `pwsh build/build.ps1`. Confirm the bundle is back to `Function 124/124`, 186 unique
      procedures.

---

## 4. Part A did not regress

Use a **fresh** workbook — `H.xlsm` is corrupt on purpose. Import the reverted bundle.

- [ ] **A fresh install still defaults to `Off`, and the off band holds the nine fields.**
      `Core_BuildSetup` on the new workbook. Expected: all eight `Availability` cells read
      `Off`, and no buttons exist on any tab.

      Then find the band headed `SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON` at the bottom
      of the Setup tab. It is **collapsed** — click the `+` in Excel's row-outline margin,
      to the left of the row numbers, to expand it. Expected inside: an `Import / Export`
      sub-heading and all nine fields — `k_export_base`, `k_settings_sheet`,
      `k_export_offset`, `k_export_tabs`, `k_import`, `k_master_prefix`, `k_master_dest`,
      `k_master_sheet`, `k_master_cols`.

      *This half is weak on its own — `Off` is the fresh default, so it would tick green on
      a fully reverted build.* The positive control is the next paragraph.

      **Positive control:** set `Remove Duplicates [clean.dupes]` to `Button`, Finalize,
      then `Core_BuildSetup` again and answer **Yes**. Expected: `clean.dupes` still reads
      `Button`. Failure here means config reading is broken outright and the all-`Off`
      result above was meaningless.

- [ ] **A blank `Availability` cell still blocks Finalize.** Set all eight actions to
      `Button`, Finalize, then clear the `Availability` cell for
      `Trim & Clean Text [clean.trim]` (select it, press Delete) and run
      `Core_FinalizeSetup`. Expected:

      ```
      MISSING or UNRECOGNISED Availability (use Button, Chains only, or Off):
        - clean.trim
      ```

      Failure: it activates. The `core.*` arm added to `ComposeReport`'s reserved-namespace
      loop sits directly above this check; a misplaced `ElseIf` could shadow it.
