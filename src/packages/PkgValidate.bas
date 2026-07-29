Attribute VB_Name = "PkgValidate"
Option Explicit

'================================================================
' PkgValidate - Wave 2 of the catalog build-out (docs/macro-catalog-
' plug-and-play.md section 3; docs/designs/2026-07-28-w2-validate-plan.md).
'
' Five actions. None of them repairs anything - contract section 7's
' domain table says so plainly: "Checking data against a rule and
' reporting what fails; never repairs silently." Two of the five write
' a cell colour; that is reporting IN PLACE rather than to a separate
' sheet, not a repair - see PkgValidate_MarkCell for the rule that
' keeps a colour call honest (never overwrite a cell's own fill, so
' clearing the fill afterward is always a correct undo).
'
'   validate.reconcile             match two ranges on a key, report
'                                   missing / extra / differing rows
'   validate.scan                  sweep a range against configured
'                                   rules, report exceptions to a new
'                                   sheet
'   validate.highlight_duplicates  colour-flag rows sharing a key
'   validate.highlight_outliers    colour-flag values outside a bound
'   validate.missing_fields        colour-flag rows missing a
'                                   mandatory column
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgValidate_Manifest(reg As Object)
    Core_DeclarePackage reg, "validate", "Validate & Reconcile", "1.0.0"

    Core_AddSection reg, "RECONCILE"
    Core_AddField reg, "k_validate_recon_dest", "Name for the reconciliation report sheet", "text", "Reconcile Exceptions", _
        "simple", _
        "A NEW sheet by this name is created only when there is something to report. If a sheet by this name already exists, the action stops and asks you to rename or remove it - it never overwrites one.", _
        "validate.reconcile"

    Core_AddSection reg, "SCAN"
    Core_AddField reg, "k_validate_scan_dest", "Name for the scan exceptions sheet", "text", "Scan Exceptions", _
        "simple", _
        "A NEW sheet by this name is created only when a rule below finds something. If a sheet by this name already exists, the action stops and asks you to rename or remove it.", _
        "validate.scan"
    Core_AddTable reg, "k_validate_scan_rules", "Data rules", _
        "Column (letter or header name)|Rule (blank, number, date, range)|Bound (range rule only: min,max)", "simple", _
        "One row per rule. Left cell: a column letter like C, or the column's header text - a one-to-three-letter entry is always read as a letter, so a column headed exactly A or ID must be addressed by its letter. Middle cell: blank (flag empty cells), number (flag anything not a number), date (flag anything not a date), or range (flag a number outside the Bound). Bound is only read for range, as two numbers separated by a comma - leave one side blank for an open-ended bound, e.g. ',100000' flags anything over 100000. Nothing happens until you fill this in.", _
        "validate.scan"

    Core_AddSection reg, "OUTLIERS"
    Core_AddTable reg, "k_validate_outliers", "Outlier bounds", _
        "Column (letter or header name)|Min (blank = no floor)|Max (blank = no ceiling)", "simple", _
        "One row per column to check. Left cell: a column letter or header name, same rule as above. A value below Min or above Max is flagged; leave either side blank to leave that side unchecked. Nothing happens until you fill this in.", _
        "validate.highlight_outliers"

    Core_AddAction reg, "validate.reconcile", "Reconcile Two Datasets", "PkgValidate_Reconcile", _
        "Pick two ranges and the key column(s) in each; reports rows found only in one side, rows sharing a key that is not unique on either side, rows with an empty key, and rows whose other columns differ - to a new sheet, after you confirm. Nothing in either range is changed.", _
        "Validate / Reconcile"
    Core_AddAction reg, "validate.scan", "Scan for Data Issues", "PkgValidate_Scan", _
        "Pick a range; checks it against the Data rules table on Setup and writes every row that fails a rule to a new sheet, after you confirm. Nothing in the picked range is changed.", _
        "Validate / Reconcile"
    Core_AddAction reg, "validate.highlight_duplicates", "Highlight Duplicates", "PkgValidate_HighlightDuplicates", _
        "Pick a range and the column(s) that form a key; colour-flags the key cell(s) of every row sharing a key with another row, after you confirm. Nothing is deleted or moved.", _
        "Validate / Reconcile"
    Core_AddAction reg, "validate.highlight_outliers", "Highlight Outliers", "PkgValidate_HighlightOutliers", _
        "Pick a range; colour-flags any cell in a configured column that falls outside its bound from the Outlier bounds table on Setup, after you confirm. Nothing is deleted or moved.", _
        "Validate / Reconcile"
    Core_AddAction reg, "validate.missing_fields", "Flag Missing Required Fields", "PkgValidate_MissingFields", _
        "Pick a range and the column(s) that are mandatory; colour-flags every blank cell in those columns, after you confirm. Nothing is deleted or moved.", _
        "Validate / Reconcile"
End Sub

'================================================================
' Shared helpers. Prefixed PkgValidate_ because the bundle is one flat
' namespace (CLAUDE.md rule 4), even though these are Private.
'================================================================

' Range.Value2, but normalised to always be a 2-D array. Value2 on a
' range that holds exactly ONE cell returns a scalar, not a 1x1 array -
' every array-indexed read in this module assumes 2-D, so that edge
' case is handled once, here, instead of guarded at every call site.
'
' Every engine in this module reads through THIS, not through
' rng.Cells(r, c).Value in a loop. A per-cell .Value read is a COM
' round trip; at "a real month-end dataset" scale that is the
' difference between a key index that builds in under a second and one
' that visibly grinds. Only the WRITE side (colouring cells, writing an
' exceptions sheet) still touches real Range objects, because
' formatting cannot be set from an array.
Private Function PkgValidate_Grid(rng As Range) As Variant
    Dim g(1 To 1, 1 To 1) As Variant
    If rng.Cells.Count = 1 Then
        g(1, 1) = rng.Value2
        PkgValidate_Grid = g
    Else
        PkgValidate_Grid = rng.Value2
    End If
End Function

' Column resolution (letter-or-header, against a picked range) is not
' defined in this package. It is Core_ResolveColumn, promoted to Core
' by phase 2 (2026-07-28-core-helper-promotion-plan.md, Task 2 Step 2)
' precisely because this wave makes it a second, independent consumer
' of the same shape PkgClean_StandardizeFormats already had - see
' "Decisions taken", W2 plan, decision 13. It returns a Long column
' index counted from 1 at rng's left edge (0 when the spec does not
' resolve inside rng), not a Range; every call site in this module
' converts with rng.Columns(idx) where it needs the Range back.

' RGB(...) is not a compile-time constant in VBA, so these are computed
' Functions rather than Const - each value is precomputed and commented
' with the RGB() call it equals, so it can be checked by eye. Chosen to
' be visually distinct from each other and to echo Excel's own built-in
' conditional-formatting palette, so a user recognises them as "a rule
' flagged this" rather than "someone coloured this by hand".
Private Function PkgValidate_ColorDuplicate() As Long
    PkgValidate_ColorDuplicate = 13551615         ' RGB(255, 199, 206) - Excel's own light red
End Function

Private Function PkgValidate_ColorOutlier() As Long
    PkgValidate_ColorOutlier = 10284031           ' RGB(255, 235, 156) - Excel's own light amber
End Function

Private Function PkgValidate_ColorMissing() As Long
    PkgValidate_ColorMissing = 16761036           ' RGB(204, 192, 255) - light violet, distinct from the other two
End Function

' Colours ONE cell, but only if it carries no fill of its own already -
' ColorIndex = xlColorIndexNone is Excel's "nothing set" state. A cell
' that already has a fill is left untouched and counted in skipped,
' never overwritten.
'
' This is what makes "select the range, Home > Fill Color > No Fill" a
' SAFE undo instruction rather than a guess: every cell this marks is
' guaranteed to have had No Fill immediately before, because a cell
' that did not was never touched. Restoring No Fill can therefore never
' destroy formatting the user set themselves.
Private Sub PkgValidate_MarkCell(c As Range, ByVal clr As Long, _
                                 ByRef marked As Long, ByRef skipped As Long)
    If c.Interior.ColorIndex = xlColorIndexNone Then
        c.Interior.Color = clr
        marked = marked + 1
    Else
        skipped = skipped + 1
    End If
End Sub

' How many of a candidate cell Collection WOULD be marked vs skipped,
' without touching anything - used to give the confirm dialog an exact
' count before the user commits, using the same test PkgValidate_MarkCell
' applies when it actually runs.
Private Sub PkgValidate_ClassifyMark(cells As Collection, ByRef willMark As Long, ByRef willSkip As Long)
    Dim c As Range
    willMark = 0: willSkip = 0
    For Each c In cells
        If c.Interior.ColorIndex = xlColorIndexNone Then
            willMark = willMark + 1
        Else
            willSkip = willSkip + 1
        End If
    Next c
End Sub

' Shared confirm-before-colouring dialog for the three highlight-style
' actions. vbYesNo, not vbYesNoCancel: Windows disables the title-bar X
' on a message box with no Cancel button, so there is no accidental
' dismissal that could be mistaken for consent - the same reasoning
' PkgClean_ConfirmBlanks uses for a delete. A colour is not a delete,
' but it is just as unreachable by Ctrl+Z once applied through VBA, so
' the same bar applies. Defaults to No.
Private Function PkgValidate_ConfirmColor(ByVal what As String, ByVal willMark As Long, _
                                          ByVal willSkip As Long, ByVal rngRef As String, _
                                          ByVal title As String) As Boolean
    Dim msg As String
    msg = "Inside " & rngRef & " this will colour-flag " & willMark & " cell(s) " & what & "." & vbCrLf & vbCrLf
    If willSkip > 0 Then
        msg = msg & willSkip & " more cell(s) matched but already carry a fill colour of their " & _
              "own, and are left alone." & vbCrLf & vbCrLf
    End If
    msg = msg & "No value, formula or existing formatting is changed - only cells with no fill " & _
          "colour today receive one." & vbCrLf & vbCrLf & _
          "This cannot be undone with Ctrl+Z. Afterwards, the exact cells this run colours will " & _
          "be selected for you - Home > Fill Color > No Fill on that selection reverses only " & _
          "what this run did." & vbCrLf & vbCrLf & "Continue?"
    PkgValidate_ConfirmColor = (MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, title) = vbYes)
End Function

' Shared confirm-before-writing dialog for the two report-sheet actions
' (reconcile, scan). Both already skip creating a sheet at all when
' there is nothing to report (Decision 7); this confirms the write
' itself once there IS something to report - Core_DestUsable only
' checks that the destination name is free of a collision, it is not a
' preview-and-confirm step, so writing a new sheet full of exceptions
' still needs its own explicit "yes, do this" the same way a colour
' does. vbYesNo, not vbYesNoCancel, for the same title-bar-X reasoning
' as PkgValidate_ConfirmColor. Defaults to No.
Private Function PkgValidate_ConfirmReport(ByVal what As String, ByVal title As String) As Boolean
    Dim msg As String
    msg = "This will create a new sheet and " & what & "." & vbCrLf & vbCrLf & _
          "No existing sheet, value, or formula is changed." & vbCrLf & vbCrLf & "Continue?"
    PkgValidate_ConfirmReport = (MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, title) = vbYes)
End Function

' Removes duplicate cells (by address) from a candidate Collection - a
' user can pick the same column twice via Core_AskCols, or configure
' the same column in two rule/bound rows, and without this a single
' cell would be counted, confirmed, and marked more than once, making
' the preview count larger than the number of cells actually affected.
Private Function PkgValidate_DedupeCells(cells As Collection) As Collection
    Dim seen As Object, c As Range, out As Collection
    Set seen = CreateObject("Scripting.Dictionary")
    Set out = New Collection
    For Each c In cells
        If Not seen.Exists(c.Address(True, True, xlA1, True)) Then
            seen.Add c.Address(True, True, xlA1, True), True
            out.Add c
        End If
    Next c
    Set PkgValidate_DedupeCells = out
End Function

' Colours every cell in a (already deduplicated) candidate Collection,
' applying PkgValidate_MarkCell's never-overwrite rule to each one, and
' returns a Range that is the Union of exactly the cells that were
' ACTUALLY coloured this run (Nothing if none were). Every one of the
' three highlight actions selects that Range afterward and tells the
' user to clear ITS fill, not the whole picked range - a cell this run
' skipped, because it already carried a fill of its own, was never part
' of what got coloured, so it must never be part of what "undo"
' touches either. Selecting the whole picked range instead (an earlier
' draft of this plan did) would strip fills this action deliberately
' left alone, which is not a safe undo.
Private Function PkgValidate_MarkCells(cells As Collection, ByVal clr As Long, _
                                       ByRef marked As Long, ByRef skipped As Long) As Range
    Dim c As Range, before As Long, out As Range
    marked = 0: skipped = 0
    For Each c In cells
        before = marked
        PkgValidate_MarkCell c, clr, marked, skipped
        If marked > before Then
            If out Is Nothing Then
                Set out = c
            Else
                Set out = Application.Union(out, c)
            End If
        End If
    Next c
    Set PkgValidate_MarkCells = out
End Function

'================================================================
' ACTION - Reconcile two datasets on a key
'================================================================
Public Sub PkgValidate_Reconcile(ctx As Object)
    Dim destName As String
    Dim rngA As Range, rngB As Range
    Dim keysA As String, keysB As String
    Dim colsA As Variant, colsB As Variant
    Dim outRows As Collection, refuseMsg As String
    Dim ws As Worksheet
    Const TITLE As String = "Reconcile Two Datasets"
    On Error GoTo Fail

    ' Cheapest failure first, before any prompt - same reasoning as
    ' PkgImport_CombineWorksheets: a refusal after the user has already
    ' answered four questions reads as a bug, not a precondition.
    destName = Trim(Core_Cfg(ctx, "k_validate_recon_dest"))
    If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub

    Set rngA = Core_AskRange("Dataset A (the primary side)")
    If rngA Is Nothing Then Exit Sub
    Set rngA = Core_NarrowToUsed(rngA)
    If rngA Is Nothing Then
        Core_Refuse ctx, "The range you picked for Dataset A is entirely outside the used " & _
               "area of its sheet, so there is nothing in it to reconcile." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If rngA.Rows.Count < 2 Then
        ' Row 1 is ALWAYS the header here, not asked (Decision 1, W2 plan) - a
        ' one-row pick has a header and no data underneath it.
        Core_Refuse ctx, "Dataset A (" & Core_RangeRef(rngA) & ") is only a header row - there " & _
               "is no data underneath it to reconcile." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rngB = Core_AskRange("Dataset B (the comparison side)")
    If rngB Is Nothing Then Exit Sub
    Set rngB = Core_NarrowToUsed(rngB)
    If rngB Is Nothing Then
        Core_Refuse ctx, "The range you picked for Dataset B is entirely outside the used " & _
               "area of its sheet, so there is nothing in it to reconcile." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If rngB.Rows.Count < 2 Then
        Core_Refuse ctx, "Dataset B (" & Core_RangeRef(rngB) & ") is only a header row - there " & _
               "is no data underneath it to reconcile." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(rngA, keysA, _
            "Which column(s) of Dataset A hold the key to match rows on?", TITLE) Then Exit Sub
    If Not Core_AskCols(rngB, keysB, _
            "Which column(s) of Dataset B hold the SAME key?", TITLE) Then Exit Sub
    colsA = Core_ParseCols(keysA)
    colsB = Core_ParseCols(keysB)

    Core_SpeedOn
    PkgValidate_ReconcileEngine rngA, rngB, colsA, colsB, outRows, refuseMsg
    Core_SpeedOff

    If Len(refuseMsg) > 0 Then
        Core_Refuse ctx, refuseMsg, TITLE
        Exit Sub
    End If

    If outRows.Count = 0 Then
        Core_Done ctx, "Every row matched: no rows found on only one side, no duplicate keys, " & _
               "no empty keys, and no differing columns." & vbCrLf & vbCrLf & "No report sheet " & _
               "was created - there was nothing to report.", TITLE
        Exit Sub
    End If

    If Not PkgValidate_ConfirmReport("write " & outRows.Count & " exception row(s) to '" & _
            destName & "'", TITLE) Then Exit Sub

    Set ws = Core_NewSheet(destName)
    PkgValidate_WriteReconcileReport ws, outRows

    Core_Done ctx, "Wrote " & outRows.Count & " exception row(s) to '" & destName & "'." & _
           vbCrLf & vbCrLf & "Status column meanings:" & vbCrLf & _
           "   In A only      - key exists in Dataset A, not in Dataset B" & vbCrLf & _
           "   In B only      - key exists in Dataset B, not in Dataset A" & vbCrLf & _
           "   Duplicate key  - the key is not unique on one or both sides, so those rows " & _
           "were NOT compared row-to-row" & vbCrLf & _
           "   Empty key      - the key column(s) were blank for this row, so it could not be " & _
           "matched to anything" & vbCrLf & _
           "   Differs        - the key matched exactly once on each side, but a column " & _
           "disagrees", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.reconcile", Err
End Sub

' The reconcile engine. Builds ONE Scripting.Dictionary index per side,
' keyed by the composite key (PkgValidate_KeyOfArr), NOT nested loops -
' a "for each row in A, for each row in B" pass is O(rows(A) x
' rows(B)), the shape that looks fine on a forty-row demo and then
' visibly freezes on a real month-end extract with thousands of rows on
' each side. Building both indexes once and then walking them is
' O(rows(A) + rows(B)) - plus one comparison per common header for
' every matched pair, so the full cost is O(rows(A) + rows(B) +
' matched-pairs x common-headers), not strictly O(rows(A) + rows(B));
' it is still linear in the inputs plus the width of what actually gets
' compared, never the product of the two sides' row counts.
'
' CompareMode = vbTextCompare is set on each Dictionary BEFORE the
' first .Add - Scripting.Dictionary raises if it is changed afterward -
' so key matching is case-insensitive the same way every other name
' comparison in this codebase is (e.g. PkgImport's sheet-name check
' uses StrComp ... vbTextCompare). Without it, "ABC123" and "abc123"
' would reconcile as two different accounts.
'
' A key that is NOT unique on a side is the case a naive implementation
' gets silently wrong: pairing the first A-row with the first B-row by
' POSITION looks plausible and is not what the data says. Nothing here
' pairs same-key rows by position. Any key whose count is greater than
' one on EITHER side is reported once, as "Duplicate key", with both
' counts, and none of its rows enter the row-by-row Differs comparison -
' there is no principled way to decide which A-row corresponds to which
' B-row from the key alone, so the honest answer is "not compared", not
' a guess. (A key absent from one side entirely still reports each of
' its occurrences as In A/B only, even if it is ALSO duplicated on its
' own side - that duplication does not matter to "does the other side
' have this at all".)
' refuseMsg is set (and outRows left empty) when the two ranges cannot
' be reconciled at all - currently only a duplicate NON-KEY header on
' either side (see PkgValidate_DuplicateHeaders). The caller checks
' refuseMsg BEFORE looking at outRows.
Private Sub PkgValidate_ReconcileEngine(rngA As Range, rngB As Range, colsA As Variant, colsB As Variant, _
                                        ByRef outRows As Collection, ByRef refuseMsg As String)
    Dim valsA As Variant, valsB As Variant
    Dim dictA As Object, dictB As Object
    Dim r As Long, k As Variant, cA As Collection, cB As Collection
    Dim commonHdrs As Collection
    Dim dupA As String, dupB As String
    Dim bExists As Boolean, bCount As Long
    Dim tmpColl As Collection

    Set outRows = New Collection
    refuseMsg = ""
    valsA = PkgValidate_Grid(rngA)
    valsB = PkgValidate_Grid(rngB)

    ' Duplicate non-key headers make "every header common to both
    ' ranges" ambiguous - comparing only the first occurrence and
    ' dropping the rest (what a naive header-to-column map does) can
    ' let a differing second "Amount" column vanish from the report
    ' with no trace. Refuse before anything else runs.
    dupA = PkgValidate_DuplicateHeaders(valsA, colsA)
    dupB = PkgValidate_DuplicateHeaders(valsB, colsB)
    If Len(dupA) > 0 Or Len(dupB) > 0 Then
        refuseMsg = "Dataset A and Dataset B must each have unique column headers outside the " & _
            "key column(s), so every column can be matched by name without ambiguity." & vbCrLf & vbCrLf
        If Len(dupA) > 0 Then refuseMsg = refuseMsg & "Dataset A repeats: " & dupA & vbCrLf
        If Len(dupB) > 0 Then refuseMsg = refuseMsg & "Dataset B repeats: " & dupB & vbCrLf
        refuseMsg = refuseMsg & vbCrLf & "Nothing has been changed."
        Exit Sub
    End If

    Set dictA = CreateObject("Scripting.Dictionary")
    Set dictB = CreateObject("Scripting.Dictionary")
    dictA.CompareMode = vbTextCompare
    dictB.CompareMode = vbTextCompare

    ' A row whose key column(s) are ALL blank is not a real key and
    ' MUST NEVER enter the match index - indexing it would let one
    ' blank-key row on each side pair as a false "match". Reported as
    ' its own status, one row at a time, instead.
    For r = 2 To UBound(valsA, 1)
        If PkgValidate_KeyIsBlank(valsA, r, colsA) Then
            outRows.Add Array("Empty key", "", "", "row " & r & " in A", "")
        Else
            k = PkgValidate_KeyOfArr(valsA, r, colsA)
            If Not dictA.Exists(k) Then
                Set tmpColl = New Collection
                dictA.Add k, tmpColl
            End If
            dictA(k).Add r
        End If
    Next r
    For r = 2 To UBound(valsB, 1)
        If PkgValidate_KeyIsBlank(valsB, r, colsB) Then
            outRows.Add Array("Empty key", "", "", "", "row " & r & " in B")
        Else
            k = PkgValidate_KeyOfArr(valsB, r, colsB)
            If Not dictB.Exists(k) Then
                Set tmpColl = New Collection
                dictB.Add k, tmpColl
            End If
            dictB(k).Add r
        End If
    Next r

    Set commonHdrs = PkgValidate_CommonHeadersArr(valsA, valsB, colsA, colsB)

    ' Duplication is judged on EACH side independently of whether the
    ' other side has the key at all - a key duplicated only in A must
    ' still report "Duplicate key" even when B has zero rows for it.
    ' Checking "cA.Count > 1 Or bCount > 1" BEFORE the not-in-B branch
    ' is what makes that true; testing existence first (the earlier
    ' shape) let a same-side duplicate slip through as a run of
    ' "In A only" rows whenever the other side had no match, losing the
    ' duplicate warning exactly where the source is most malformed.
    For Each k In dictA.Keys
        Set cA = dictA(k)
        bExists = dictB.Exists(k)
        bCount = 0
        If bExists Then bCount = dictB(k).Count
        If cA.Count > 1 Or bCount > 1 Then
            outRows.Add Array("Duplicate key", k, "", cA.Count & " row(s) in A", bCount & " row(s) in B")
        ElseIf Not bExists Then
            PkgValidate_AddInAOnly outRows, valsA, cA, k
        Else
            Set cB = dictB(k)
            PkgValidate_CompareRow outRows, valsA, CLng(cA(1)), valsB, CLng(cB(1)), commonHdrs, k
        End If
    Next k
    For Each k In dictB.Keys
        If Not dictA.Exists(k) Then
            Set cB = dictB(k)
            If cB.Count > 1 Then
                outRows.Add Array("Duplicate key", k, "", "0 row(s) in A", cB.Count & " row(s) in B")
            Else
                PkgValidate_AddInBOnly outRows, valsB, cB, k
            End If
        End If
    Next k
End Sub

' True when every key column of row r is blank - see the "Empty key"
' handling above.
Private Function PkgValidate_KeyIsBlank(vals As Variant, ByVal r As Long, keyCols As Variant) As Boolean
    Dim i As Long
    PkgValidate_KeyIsBlank = True
    For i = LBound(keyCols) To UBound(keyCols)
        If Not Core_IsBlank(vals(r, keyCols(i))) Then
            PkgValidate_KeyIsBlank = False
            Exit Function
        End If
    Next i
End Function

' Non-key header names that occur more than once in vals' header row
' (row 1), comma-joined for the refusal message, or "" if none repeat.
Private Function PkgValidate_DuplicateHeaders(vals As Variant, keyCols As Variant) As String
    Dim map As Object, c As Long, hdr As String, out As String
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    For c = 1 To UBound(vals, 2)
        If Not PkgValidate_InArray(c, keyCols) Then
            hdr = Trim$(CStr(vals(1, c)))
            If Len(hdr) > 0 Then
                If map.Exists(hdr) Then
                    If InStr(1, out, hdr, vbTextCompare) = 0 Then
                        out = out & IIf(Len(out) > 0, ", ", "") & hdr
                    End If
                Else
                    map.Add hdr, c
                End If
            End If
        End If
    Next c
    PkgValidate_DuplicateHeaders = out
End Function

' Builds a composite key from one or more columns of ONE row, numbered
' from the left of the range exactly like Core_AskCols/Core_ParseCols
' number them. Chr(31) (the ASCII "unit separator") joins the parts
' because it is not a character a person can type into a cell by hand,
' so a key made of "A" + "BC" can never collide with one made of "AB" +
' "C" the way joining on a comma or pipe could.
Private Function PkgValidate_KeyOfArr(vals As Variant, ByVal r As Long, keyCols As Variant) As String
    Dim i As Long, parts As String
    For i = LBound(keyCols) To UBound(keyCols)
        parts = parts & Chr$(31) & Trim$(CStr(vals(r, keyCols(i))))
    Next i
    PkgValidate_KeyOfArr = Mid$(parts, 2)     ' drop the leading separator
End Function

' Headers shared by both ranges, EXCLUDING each side's own key
' column(s) - those are already equal by construction (they are what
' matched the rows), so comparing them again would just repeat the key
' back at the user as a false "difference" whenever the two sides
' format it differently (e.g. "007" vs "7").
Private Function PkgValidate_CommonHeadersArr(valsA As Variant, valsB As Variant, colsA As Variant, colsB As Variant) As Collection
    Dim mapA As Object, mapB As Object, c As Long, hdr As String
    Dim out As New Collection, k As Variant
    Set mapA = CreateObject("Scripting.Dictionary"): mapA.CompareMode = vbTextCompare
    Set mapB = CreateObject("Scripting.Dictionary"): mapB.CompareMode = vbTextCompare
    For c = 1 To UBound(valsA, 2)
        If Not PkgValidate_InArray(c, colsA) Then
            hdr = Trim$(CStr(valsA(1, c)))
            If Len(hdr) > 0 Then If Not mapA.Exists(hdr) Then mapA.Add hdr, c
        End If
    Next c
    For c = 1 To UBound(valsB, 2)
        If Not PkgValidate_InArray(c, colsB) Then
            hdr = Trim$(CStr(valsB(1, c)))
            If Len(hdr) > 0 Then If Not mapB.Exists(hdr) Then mapB.Add hdr, c
        End If
    Next c
    For Each k In mapA.Keys
        If mapB.Exists(k) Then out.Add Array(CStr(k), CLng(mapA(k)), CLng(mapB(k)))
    Next k
    Set PkgValidate_CommonHeadersArr = out
End Function

Private Function PkgValidate_InArray(ByVal n As Long, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = n Then PkgValidate_InArray = True: Exit Function
    Next i
End Function

' Compares one matched pair of rows across every header common to both
' sides, using PkgValidate_ValuesEqual so a formatting difference like
' "100" vs "100.00" does not read as a disagreement. One output row per
' column that actually differs - a pair with no differing column adds
' nothing, which is how a full match stays silent rather than
' cluttering the report with an all-clear line for every one of
' possibly thousands of matched rows.
Private Sub PkgValidate_CompareRow(outRows As Collection, valsA As Variant, ByVal rA As Long, _
                                   valsB As Variant, ByVal rB As Long, commonHdrs As Collection, ByVal k As Variant)
    Dim pair As Variant, vA As Variant, vB As Variant
    For Each pair In commonHdrs
        vA = valsA(rA, pair(1))
        vB = valsB(rB, pair(2))
        If Not PkgValidate_ValuesEqual(vA, vB) Then
            outRows.Add Array("Differs", k, CStr(pair(0)), Trim$(CStr(vA)), Trim$(CStr(vB)))
        End If
    Next pair
End Sub

' Numeric-tolerant, case-insensitive equality for one cell pair. Two
' numbers are compared as numbers (so "100" and "100.00" agree, and so
' does ordinary floating-point noise past the ninth decimal); anything
' else is compared as trimmed text, case-insensitively, matching every
' other text comparison in this codebase.
Private Function PkgValidate_ValuesEqual(ByVal vA As Variant, ByVal vB As Variant) As Boolean
    If IsNumeric(vA) And IsNumeric(vB) Then
        PkgValidate_ValuesEqual = (Abs(CDbl(vA) - CDbl(vB)) < 0.000000001)
    Else
        PkgValidate_ValuesEqual = (StrComp(Trim$(CStr(vA)), Trim$(CStr(vB)), vbTextCompare) = 0)
    End If
End Function

Private Sub PkgValidate_AddInAOnly(outRows As Collection, valsA As Variant, cA As Collection, ByVal k As Variant)
    Dim i As Long, r As Long
    For i = 1 To cA.Count
        r = cA(i)
        outRows.Add Array("In A only", k, "", PkgValidate_RowTextArr(valsA, r), "")
    Next i
End Sub

Private Sub PkgValidate_AddInBOnly(outRows As Collection, valsB As Variant, cB As Collection, ByVal k As Variant)
    Dim i As Long, r As Long
    For i = 1 To cB.Count
        r = cB(i)
        outRows.Add Array("In B only", k, "", "", PkgValidate_RowTextArr(valsB, r))
    Next i
End Sub

' All of one row's values, comma-joined, so an "In A only" / "In B
' only" line shows the whole record rather than just the key that
' matched - the report stays a fixed 5 columns regardless of how wide
' either source range was, by folding the extra columns into one cell
' rather than widening the sheet per source.
Private Function PkgValidate_RowTextArr(vals As Variant, ByVal r As Long) As String
    Dim c As Long, out As String
    For c = 1 To UBound(vals, 2)
        If c > 1 Then out = out & ", "
        out = out & Trim$(CStr(vals(r, c)))
    Next c
    PkgValidate_RowTextArr = out
End Function

' Writes the exceptions Collection built above to a fresh sheet, one
' row per exception. This side stays plain Range.Cells writes, not a
' bulk array write like the read side - the exceptions count is bounded
' by how much is actually wrong, not by the size of either source
' range, so it is never the hot path Decision 16 is about.
Private Sub PkgValidate_WriteReconcileReport(ws As Worksheet, outRows As Collection)
    Dim i As Long, item As Variant
    ws.Range("A1:E1").Value = Array("Status", "Key", "Column", "Value in A", "Value in B")
    ws.Range("A1:E1").Font.Bold = True
    i = 2
    For Each item In outRows
        ws.Cells(i, 1).Value = item(0)
        ws.Cells(i, 2).Value = item(1)
        ws.Cells(i, 3).Value = item(2)
        ws.Cells(i, 4).Value = item(3)
        ws.Cells(i, 5).Value = item(4)
        i = i + 1
    Next item
    ws.Columns("A:E").AutoFit
End Sub

'================================================================
' ACTION - Scan a range against configured rules
'================================================================
Public Sub PkgValidate_Scan(ctx As Object)
    Dim destName As String, rng As Range, hasHeader As Boolean
    Dim cols() As String, rules() As String, bounds() As String
    Dim outRows As Collection, missed As String, badRules As String
    Dim ws As Worksheet
    Const TITLE As String = "Scan for Data Issues"
    On Error GoTo Fail

    destName = Trim(Core_Cfg(ctx, "k_validate_scan_dest"))
    If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub

    If Not Core_CfgTable(ctx, "k_validate_scan_rules", cols, rules, bounds) Then
        Core_Refuse ctx, "No data rules are set up yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Data rules' table on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Configuration errors must refuse, not fail open - an unrecognised
    ' Rule, or a Bound component that is neither blank nor a number,
    ' silently became "unbounded" (or "never fails", for an unknown
    ' rule name) in an earlier draft of this engine. Validate the whole
    ' table before reading a single cell.
    badRules = PkgValidate_ValidateScanRules(cols, rules, bounds)
    If Len(badRules) > 0 Then
        Core_Refuse ctx, "The Data rules table on Setup has problem(s):" & badRules & vbCrLf & vbCrLf & _
               "Fix the table on the Setup tab, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to scan." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub

    Core_SpeedOn
    PkgValidate_ScanEngine rng, hasHeader, cols, rules, bounds, outRows, missed
    Core_SpeedOff

    If outRows.Count = 0 Then
        Core_Done ctx, "No exceptions found inside " & Core_RangeRef(rng) & "." & vbCrLf & vbCrLf & _
               "No report sheet was created - there was nothing to report." & _
               IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These configured rule(s) matched no " & _
                   "column in your selection and were skipped:" & missed, ""), TITLE
        Exit Sub
    End If

    If Not PkgValidate_ConfirmReport("write " & outRows.Count & " exception(s) to '" & destName & _
            "'", TITLE) Then Exit Sub

    Set ws = Core_NewSheet(destName)
    PkgValidate_WriteScanReport ws, outRows

    Core_Done ctx, "Wrote " & outRows.Count & " exception(s) to '" & destName & "'." & _
           IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These configured rule(s) matched no column " & _
               "in your selection and were skipped:" & missed, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.scan", Err
End Sub

' Validates the whole k_validate_scan_rules table BEFORE any cell is
' read, so a misconfigured Setup table refuses cleanly, naming the
' offending row(s), instead of the engine silently treating an
' unrecognised rule name as "never fails" or non-numeric bound text as
' "no bound on that side".
Private Function PkgValidate_ValidateScanRules(cols() As String, rules() As String, bounds() As String) As String
    Dim i As Long, ruleName As String, msg As String, parts() As String
    For i = LBound(cols) To UBound(cols)
        ruleName = LCase$(Trim$(rules(i)))
        Select Case ruleName
            Case "blank", "number", "date"
                ' no Bound expected; nothing further to validate
            Case "range"
                If Len(Trim$(bounds(i))) = 0 Then
                    msg = msg & vbCrLf & "   Row for '" & cols(i) & "': rule 'range' needs a " & _
                        "Bound (min,max) - it is blank."
                Else
                    parts = Split(bounds(i), ",")
                    If UBound(parts) >= 0 Then
                        If Len(Trim$(parts(0))) > 0 Then
                            If Not IsNumeric(Trim$(parts(0))) Then
                                msg = msg & vbCrLf & "   Row for '" & cols(i) & "': Bound minimum '" & _
                                    Trim$(parts(0)) & "' is not a number."
                            End If
                        End If
                    End If
                    If UBound(parts) >= 1 Then
                        If Len(Trim$(parts(1))) > 0 Then
                            If Not IsNumeric(Trim$(parts(1))) Then
                                msg = msg & vbCrLf & "   Row for '" & cols(i) & "': Bound maximum '" & _
                                    Trim$(parts(1)) & "' is not a number."
                            End If
                        End If
                    End If
                End If
            Case Else
                msg = msg & vbCrLf & "   Row for '" & cols(i) & "': rule '" & rules(i) & "' is not " & _
                    "recognised - must be blank, number, date, or range."
        End Select
    Next i
    PkgValidate_ValidateScanRules = msg
End Function

' One pass per configured rule row, not one pass per cell per rule -
' the column is resolved ONCE (Core_ResolveColumn) and its values
' read ONCE in bulk (PkgValidate_Grid), so N rules over an M-column
' range costs N bulk reads of one column each, not N x (rows x M)
' individual cell reads.
Private Sub PkgValidate_ScanEngine(rng As Range, ByVal hasHeader As Boolean, cols() As String, _
                                   rules() As String, bounds() As String, ByRef outRows As Collection, _
                                   ByRef missed As String)
    Dim i As Long, r As Long, firstRow As Long, idx As Long, tgt As Range, vals As Variant, v As Variant
    Dim ruleName As String, lo As Double, hi As Double, hasLo As Boolean, hasHi As Boolean
    Dim reason As String, hdr As String
    Set outRows = New Collection
    firstRow = IIf(hasHeader, 2, 1)

    For i = LBound(cols) To UBound(cols)
        idx = Core_ResolveColumn(rng, cols(i))
        If idx = 0 Then
            missed = missed & vbCrLf & "   " & cols(i)
        Else
            Set tgt = rng.Columns(idx)
            vals = PkgValidate_Grid(tgt)
            hdr = Trim$(CStr(vals(1, 1)))
            ruleName = LCase$(Trim$(rules(i)))
            PkgValidate_ParseBound bounds(i), lo, hasLo, hi, hasHi
            For r = firstRow To UBound(vals, 1)
                v = vals(r, 1)
                reason = PkgValidate_RuleFails(ruleName, v, lo, hasLo, hi, hasHi)
                If Len(reason) > 0 Then
                    outRows.Add Array(Core_RangeRef(tgt.Cells(r, 1)), hdr, ruleName, Trim$(CStr(v)), reason)
                End If
            Next r
        End If
    Next i
End Sub

' "min,max" -> two optional Doubles. Either side blank means unbounded
' on that side. A non-numeric side is treated the same as blank
' (unbounded) rather than raising - a typo'd bound should not crash the
' scan, it should just not constrain that side.
Private Sub PkgValidate_ParseBound(ByVal spec As String, ByRef lo As Double, ByRef hasLo As Boolean, _
                                   ByRef hi As Double, ByRef hasHi As Boolean)
    Dim parts() As String
    hasLo = False: hasHi = False: lo = 0: hi = 0
    parts = Split(spec, ",")
    If UBound(parts) >= 0 Then
        If IsNumeric(Trim$(parts(0))) Then lo = CDbl(Trim$(parts(0))): hasLo = True
    End If
    If UBound(parts) >= 1 Then
        If IsNumeric(Trim$(parts(1))) Then hi = CDbl(Trim$(parts(1))): hasHi = True
    End If
End Sub

' "" means the cell passes. Any non-empty string is the reason it did
' not, and becomes the report's Reason column verbatim. A blank cell
' passes "number" and "date" - a blank is not "not a number", it is
' just blank, and flagging it here would duplicate validate.missing_fields'
' job; only the "blank" rule itself flags blanks.
Private Function PkgValidate_RuleFails(ByVal ruleName As String, ByVal v As Variant, _
                                       ByVal lo As Double, ByVal hasLo As Boolean, _
                                       ByVal hi As Double, ByVal hasHi As Boolean) As String
    Select Case ruleName
        Case "blank"
            If Core_IsBlank(v) Then PkgValidate_RuleFails = "Blank"
        Case "number"
            If Not Core_IsBlank(v) Then
                If Not IsNumeric(v) Then PkgValidate_RuleFails = "Not a number"
            End If
        Case "date"
            If Not Core_IsBlank(v) Then
                If Not IsDate(v) Then PkgValidate_RuleFails = "Not a date"
            End If
        Case "range"
            If Not Core_IsBlank(v) Then
                If IsNumeric(v) Then
                    If hasLo Then
                        If CDbl(v) < lo Then PkgValidate_RuleFails = CStr(v) & " is below the minimum " & lo
                    End If
                    If Len(PkgValidate_RuleFails) = 0 And hasHi Then
                        If CDbl(v) > hi Then PkgValidate_RuleFails = CStr(v) & " is above the maximum " & hi
                    End If
                Else
                    ' A non-numeric value under a range rule is itself
                    ' the exception - it must not pass silently just
                    ' because there is no number to compare against the
                    ' bound.
                    PkgValidate_RuleFails = "Not a number, so its range cannot be checked"
                End If
            End If
        Case Else
            ' Defence in depth only - PkgValidate_ValidateScanRules
            ' refuses an unrecognised rule name before this function is
            ' ever reached, so this branch should be unreachable.
            PkgValidate_RuleFails = "Unrecognised rule '" & ruleName & "'"
    End Select
End Function

Private Sub PkgValidate_WriteScanReport(ws As Worksheet, outRows As Collection)
    Dim i As Long, item As Variant
    ws.Range("A1:E1").Value = Array("Cell", "Column", "Rule", "Value", "Reason")
    ws.Range("A1:E1").Font.Bold = True
    i = 2
    For Each item In outRows
        ws.Cells(i, 1).Value = item(0)
        ws.Cells(i, 2).Value = item(1)
        ws.Cells(i, 3).Value = item(2)
        ws.Cells(i, 4).Value = item(3)
        ws.Cells(i, 5).Value = item(4)
        i = i + 1
    Next item
    ws.Columns("A:E").AutoFit
End Sub

'================================================================
' ACTION - Colour-flag rows sharing a duplicate key
'================================================================
Public Sub PkgValidate_HighlightDuplicates(ctx As Object)
    Dim rng As Range, hasHeader As Boolean, keys As String, cols As Variant
    Dim dupCells As Collection, dupRows As Long
    Dim willMark As Long, willSkip As Long
    Dim marked As Long, skipped As Long, markedRng As Range, undoMsg As String
    Const TITLE As String = "Highlight Duplicates"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to check." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub
    If Not Core_AskCols(rng, keys, _
            "Which column(s) form the key - two rows sharing the same value(s) here are " & _
            "duplicates?", TITLE) Then Exit Sub
    cols = Core_ParseCols(keys)

    Core_SpeedOn
    Set dupCells = PkgValidate_FindDuplicateCells(rng, hasHeader, cols, dupRows)
    Core_SpeedOff
    Set dupCells = PkgValidate_DedupeCells(dupCells)

    If dupCells.Count = 0 Then
        Core_Done ctx, "No duplicate keys found inside " & Core_RangeRef(rng) & "." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    PkgValidate_ClassifyMark dupCells, willMark, willSkip
    If Not PkgValidate_ConfirmColor("marking the key cell(s) of " & dupRows & " row(s) that " & _
            "share a key with another row", willMark, willSkip, Core_RangeRef(rng), TITLE) Then Exit Sub

    Core_SpeedOn
    Set markedRng = PkgValidate_MarkCells(dupCells, PkgValidate_ColorDuplicate(), marked, skipped)
    Core_SpeedOff

    ' The undo instruction names ONLY the cells this run coloured, never
    ' the whole picked range - a cell rng contains that already had its
    ' own fill was deliberately skipped, and "select rng, No Fill" would
    ' strip that pre-existing fill too. Selecting markedRng for the user
    ' means Home > Fill Color > No Fill on the current selection is
    ' always, and only, an undo of what this run did.
    If Not markedRng Is Nothing Then
        markedRng.Select
        undoMsg = "To remove the flags: the " & marked & " marked cell(s) are now selected - " & _
            "use Home > Fill Color > No Fill."
    End If

    Core_Done ctx, "Colour-flagged " & marked & " key cell(s) across " & dupRows & " row(s) " & _
           "sharing a duplicate key." & IIf(skipped > 0, vbCrLf & vbCrLf & skipped & " cell(s) " & _
           "already had a fill colour of their own and were left alone.", "") & vbCrLf & vbCrLf & _
           undoMsg, TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.highlight_duplicates", Err
End Sub

' Every KEY-COLUMN cell (not every cell in the row) belonging to a row
' whose composite key is shared with at least one other row in rng.
' Colouring only the key cells, not the whole row, keeps the footprint
' of a single confirmed action as small as the finding it reports - the
' row is a duplicate BECAUSE of what is in these cells, and a wider
' flag would touch formatting on columns that had nothing to do with
' it. The caller deduplicates the result (Core_AskCols permits a
' repeated column number, which would otherwise add the same cell to
' the candidate list twice and inflate the preview count).
Private Function PkgValidate_FindDuplicateCells(rng As Range, ByVal hasHeader As Boolean, cols As Variant, _
                                                ByRef dupRowCount As Long) As Collection
    Dim vals As Variant, dict As Object, r As Long, firstRow As Long, k As Variant, i As Long, j As Long
    Dim out As Collection, tmpColl As Collection
    Set out = New Collection
    vals = PkgValidate_Grid(rng)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    firstRow = IIf(hasHeader, 2, 1)
    For r = firstRow To UBound(vals, 1)
        k = PkgValidate_KeyOfArr(vals, r, cols)
        If Not dict.Exists(k) Then
            Set tmpColl = New Collection
            dict.Add k, tmpColl
        End If
        dict(k).Add r
    Next r
    dupRowCount = 0
    For Each k In dict.Keys
        If dict(k).Count > 1 Then
            dupRowCount = dupRowCount + dict(k).Count
            For i = 1 To dict(k).Count
                r = dict(k)(i)
                For j = LBound(cols) To UBound(cols)
                    out.Add rng.Cells(r, cols(j))
                Next j
            Next i
        End If
    Next k
    Set PkgValidate_FindDuplicateCells = out
End Function

'================================================================
' ACTION - Colour-flag values outside a configured bound
'================================================================
Public Sub PkgValidate_HighlightOutliers(ctx As Object)
    Dim rng As Range, hasHeader As Boolean
    Dim cols() As String, mins() As String, maxs() As String, missed As String, badBounds As String
    Dim outCells As Collection
    Dim willMark As Long, willSkip As Long
    Dim marked As Long, skipped As Long, markedRng As Range, undoMsg As String
    Const TITLE As String = "Highlight Outliers"
    On Error GoTo Fail

    If Not Core_CfgTable(ctx, "k_validate_outliers", cols, mins, maxs) Then
        Core_Refuse ctx, "No outlier bounds are set up yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Outlier bounds' table on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Same "refuse, don't fail open" rule as validate.scan - a Min or
    ' Max that is neither blank nor a number must not silently become
    ' "no bound on that side".
    badBounds = PkgValidate_ValidateOutlierBounds(cols, mins, maxs)
    If Len(badBounds) > 0 Then
        Core_Refuse ctx, "The Outlier bounds table on Setup has problem(s):" & badBounds & vbCrLf & vbCrLf & _
               "Fix the table on the Setup tab, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to check." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub

    Core_SpeedOn
    Set outCells = PkgValidate_FindOutlierCells(rng, hasHeader, cols, mins, maxs, missed)
    Core_SpeedOff
    Set outCells = PkgValidate_DedupeCells(outCells)

    If outCells.Count = 0 Then
        Core_Done ctx, "No values outside their configured bound(s) were found inside " & _
               Core_RangeRef(rng) & "." & vbCrLf & vbCrLf & "Nothing has been changed." & _
               IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These configured column(s) matched " & _
                   "nothing in your selection and were skipped:" & missed, ""), TITLE
        Exit Sub
    End If

    PkgValidate_ClassifyMark outCells, willMark, willSkip
    If Not PkgValidate_ConfirmColor("that fall outside their configured bound", willMark, _
            willSkip, Core_RangeRef(rng), TITLE) Then Exit Sub

    Core_SpeedOn
    Set markedRng = PkgValidate_MarkCells(outCells, PkgValidate_ColorOutlier(), marked, skipped)
    Core_SpeedOff

    ' See the identical comment in PkgValidate_HighlightDuplicates -
    ' the undo instruction must name only the cells this run coloured,
    ' never the whole picked range.
    If Not markedRng Is Nothing Then
        markedRng.Select
        undoMsg = "To remove the flags: the " & marked & " marked cell(s) are now selected - " & _
            "use Home > Fill Color > No Fill."
    End If

    Core_Done ctx, "Colour-flagged " & marked & " value(s) outside their configured bound." & _
           IIf(skipped > 0, vbCrLf & vbCrLf & skipped & " cell(s) already had a fill colour " & _
           "of their own and were left alone.", "") & vbCrLf & vbCrLf & undoMsg & _
           IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These configured column(s) matched nothing " & _
               "in your selection and were skipped:" & missed, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.highlight_outliers", Err
End Sub

' Validates every k_validate_outliers row BEFORE any cell is read - a
' non-blank Min or Max that is not a number must refuse, naming the
' row, instead of silently being treated as an open-ended side.
Private Function PkgValidate_ValidateOutlierBounds(cols() As String, mins() As String, maxs() As String) As String
    Dim i As Long, msg As String
    For i = LBound(cols) To UBound(cols)
        If Len(Trim$(mins(i))) > 0 Then
            If Not IsNumeric(Trim$(mins(i))) Then
                msg = msg & vbCrLf & "   Row for '" & cols(i) & "': Min '" & Trim$(mins(i)) & _
                    "' is not a number."
            End If
        End If
        If Len(Trim$(maxs(i))) > 0 Then
            If Not IsNumeric(Trim$(maxs(i))) Then
                msg = msg & vbCrLf & "   Row for '" & cols(i) & "': Max '" & Trim$(maxs(i)) & _
                    "' is not a number."
            End If
        End If
        If Len(Trim$(mins(i))) = 0 And Len(Trim$(maxs(i))) = 0 Then
            msg = msg & vbCrLf & "   Row for '" & cols(i) & "': both Min and Max are blank, so " & _
                "this row would never flag anything."
        End If
    Next i
    PkgValidate_ValidateOutlierBounds = msg
End Function

' An outlier check only applies to actual numbers - a blank or
' non-numeric cell is SKIPPED here, never flagged. Flagging blanks is
' validate.missing_fields' job; conflating the two would mean a single
' cell could be flagged twice, for two different reasons, by two
' different actions, with no way to tell which one fired. The caller
' deduplicates the result (a Min/Max row can name a column another row
' also names, or Core_AskCols-style repeats elsewhere in this package,
' which would otherwise add the same cell twice).
Private Function PkgValidate_FindOutlierCells(rng As Range, ByVal hasHeader As Boolean, cols() As String, _
                                              mins() As String, maxs() As String, ByRef missed As String) As Collection
    Dim out As New Collection
    Dim i As Long, r As Long, firstRow As Long, idx As Long, tgt As Range, vals As Variant, v As Variant
    Dim lo As Double, hi As Double, hasLo As Boolean, hasHi As Boolean
    firstRow = IIf(hasHeader, 2, 1)
    For i = LBound(cols) To UBound(cols)
        idx = Core_ResolveColumn(rng, cols(i))
        If idx = 0 Then
            missed = missed & vbCrLf & "   " & cols(i)
        Else
            Set tgt = rng.Columns(idx)
            vals = PkgValidate_Grid(tgt)
            ' Non-blank implies numeric here - PkgValidate_ValidateOutlierBounds
            ' already refused the run otherwise, so Len(...) > 0 is a
            ' safe stand-in for the IsNumeric check an earlier draft
            ' used (and CDbl below cannot raise).
            hasLo = (Len(Trim$(mins(i))) > 0)
            If hasLo Then lo = CDbl(Trim$(mins(i)))
            hasHi = (Len(Trim$(maxs(i))) > 0)
            If hasHi Then hi = CDbl(Trim$(maxs(i)))
            For r = firstRow To UBound(vals, 1)
                v = vals(r, 1)
                If IsNumeric(v) And Not Core_IsBlank(v) Then
                    ' v is confirmed numeric by the guard above, so CDbl(v)
                    ' cannot raise regardless of hasLo/hasHi's values - this
                    ' single If adds the cell at most once even when both
                    ' sides are configured.
                    If (hasLo And CDbl(v) < lo) Or (hasHi And CDbl(v) > hi) Then out.Add tgt.Cells(r, 1)
                End If
            Next r
        End If
    Next i
    Set PkgValidate_FindOutlierCells = out
End Function

'================================================================
' ACTION - Colour-flag blank cells in mandatory columns
'================================================================
Public Sub PkgValidate_MissingFields(ctx As Object)
    Dim rng As Range, hasHeader As Boolean, keys As String, cols As Variant
    Dim blankCells As Collection
    Dim willMark As Long, willSkip As Long
    Dim marked As Long, skipped As Long, markedRng As Range, undoMsg As String
    Const TITLE As String = "Flag Missing Required Fields"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to check." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub
    If Not Core_AskCols(rng, keys, _
            "Which column(s) are mandatory - a blank cell in any of them gets flagged?", TITLE) Then Exit Sub
    cols = Core_ParseCols(keys)

    Core_SpeedOn
    Set blankCells = PkgValidate_FindBlankCells(rng, hasHeader, cols)
    Core_SpeedOff
    Set blankCells = PkgValidate_DedupeCells(blankCells)

    If blankCells.Count = 0 Then
        Core_Done ctx, "No blank cells found in the mandatory column(s) inside " & _
               Core_RangeRef(rng) & "." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    PkgValidate_ClassifyMark blankCells, willMark, willSkip
    If Not PkgValidate_ConfirmColor("that are blank in a mandatory column", willMark, willSkip, _
            Core_RangeRef(rng), TITLE) Then Exit Sub

    Core_SpeedOn
    Set markedRng = PkgValidate_MarkCells(blankCells, PkgValidate_ColorMissing(), marked, skipped)
    Core_SpeedOff

    ' See the identical comment in PkgValidate_HighlightDuplicates -
    ' the undo instruction must name only the cells this run coloured,
    ' never the whole picked range.
    If Not markedRng Is Nothing Then
        markedRng.Select
        undoMsg = "To remove the flags: the " & marked & " marked cell(s) are now selected - " & _
            "use Home > Fill Color > No Fill."
    End If

    Core_Done ctx, "Colour-flagged " & marked & " blank mandatory cell(s)." & _
           IIf(skipped > 0, vbCrLf & vbCrLf & skipped & " cell(s) already had a fill colour " & _
           "of their own and were left alone.", "") & vbCrLf & vbCrLf & undoMsg, TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "validate.missing_fields", Err
End Sub

' Caller deduplicates the result - a repeated column number from
' Core_AskCols would otherwise add the same blank cell twice.
Private Function PkgValidate_FindBlankCells(rng As Range, ByVal hasHeader As Boolean, cols As Variant) As Collection
    Dim out As New Collection
    Dim vals As Variant, r As Long, firstRow As Long, i As Long
    vals = PkgValidate_Grid(rng)
    firstRow = IIf(hasHeader, 2, 1)
    For r = firstRow To UBound(vals, 1)
        For i = LBound(cols) To UBound(cols)
            If Core_IsBlank(vals(r, cols(i))) Then out.Add rng.Cells(r, cols(i))
        Next i
    Next r
    Set PkgValidate_FindBlankCells = out
End Function
