Attribute VB_Name = "PkgClean"
Option Explicit

'================================================================
' PkgClean - a small second package, built to prove Core is a
' platform: a brand-new package's buttons appear on the Macros tabs
' with ZERO changes to Core beyond adding "Clean" to Core's candidate
' list. All plumbing (speed, errors, completion dialogs) is Core_*.
'
' From v0.5.0 to v2.2.0 this package declared NO config fields and rendered
' no band on Setup at all. v2.3.0 added k_clean_formats, a deployment-level
' number-format map. Split Column later added two deployment-level settings
' for its delimiter and destination. The point it used to prove still holds:
' a package asks for Setup surface only where the answer describes the
' DEPLOYMENT, never where it describes the range the caller just picked.
' Everything else it needs describes the range the user picks, so it is
' asked for at run time. See PkgClean_Manifest.
'
' Depends on: Core (required).
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgClean_Manifest(reg As Object)
    Core_DeclarePackage reg, "clean", "Clean & Standardize", "1.0.0"

    ' From v0.5.0 to v2.2.0 this package declared no fields and rendered no band
    ' on Setup. k_clean_formats, below, was the first deployment-level setting.
    ' Both of its former settings - the key column(s) and whether the first row is a
    ' header - describe THE RANGE THE USER JUST PICKED, not the deployment, so they
    ' are asked for at run time instead. The key columns are the clearer case: they
    ' are numbered from the left of a selection that does not exist until the action
    ' runs. Storing either on Setup meant re-editing Setup and re-Finalizing between
    ' two runs of the same button, and the v0.4.0 owner gate found that nobody
    ' discovers the setting at the moment they act.
    '
    ' Asking every time is the feature, not a cost: it is what makes the answer match
    ' the selection. Do NOT reintroduce these as remembered settings.
    '
    ' v2.3.0 adds ONE field, and it is not a walk-back of the paragraph above.
    ' A house number format is a fact about the DEPLOYMENT - the same columns get
    ' the same formats every close - not a fact about the range the caller is
    ' about to pick. Its rows address columns by LETTER or by HEADER NAME, never
    ' by position within the selection; a position counted from the left of a
    ' selection that does not exist yet would be clean_keycols under a new name.
    '
    ' It declares NO default rows, deliberately. Core resolves an EMPTY stored table
    ' to the field's declared defaults, so a field with demo rows can never observe
    ' "nothing configured" - clean.standardize_formats' refusal would be unreachable
    ' code and the gate item testing it would be impossible to execute. It also means
    ' a first-time user cannot have mm/dd/yyyy applied to column A of a range they
    ' picked before filling anything in. Examples live in the help text instead.
    ' Do NOT add Core_AddDefault rows here.
    Core_AddTable reg, "k_clean_formats", "Column formats", _
        "Column (letter or header name)|Number format", "simple", _
        "One row per column you want formatted. Left cell: a column letter like C, or the column's header text - a one-to-three-letter entry is always read as a letter, so a column headed exactly A or ID must be addressed by its letter. Right cell: the format, e.g. mm/dd/yyyy or $#,##0.00. Nothing happens until you fill this in.", _
        "clean.standardize_formats"

    Core_AddSection reg, "Split Column"
    Core_AddField reg, "k_clean_split_delimiter", "Split-column delimiter", "text", _
        "COMMA", "simple", _
        "Enter one delimiter character, or use COMMA, TAB, SEMICOLON, PIPE, or SPACE.", _
        "clean.split_column"
    Core_AddField reg, "k_clean_split_destination", "Split-column destination", "text", _
        "SOURCE", "simple", _
        "Use SOURCE to replace the picked column, or enter an absolute worksheet column letter such as D. A separate destination leaves the source column unchanged.", _
        "clean.split_column"

    Core_AddAction reg, "clean.dupes", "Remove Duplicates", "PkgClean_RemoveDupes", "Pick a range, then answer two questions about it: whether its first row is a header, and which column(s) decide a duplicate."
    Core_AddAction reg, "clean.trim", "Trim & Clean Text", "PkgClean_TrimClean", "Pick a range; strips leading/trailing spaces, collapses double spaces, and removes non-printing characters."
    Core_AddAction reg, "clean.filldown", "Fill Blanks Down", "PkgClean_FillDown", "Pick a range; fills each blank cell with the value from the cell above it."
    Core_AddAction reg, "clean.remove_blanks", "Remove Blank Rows & Columns", "PkgClean_RemoveBlanks", "Pick a range; shows you how many completely blank rows and columns it found, and deletes them only after you say yes."
    Core_AddAction reg, "clean.unmerge_fill", "Unmerge & Fill Down", "PkgClean_UnmergeFill", "Pick a range; unmerges every merged block and puts the block's value into every cell it covered, so the data can be sorted and filtered."
    Core_AddAction reg, "clean.standardize_formats", "Standardize Formats", "PkgClean_StandardizeFormats", "Pick a range; applies the number formats listed in the Column formats table on Setup. Changes how values look, never the values themselves."
    Core_AddAction reg, "clean.text_to_numbers", "Convert Text to Numbers", "PkgClean_TextToNumbers", "Pick a range; turns text that represents a number into a real number, stripping $ and thousands separators. Formulas are never touched. Note that 12% becomes 12, not 0.12."
    Core_AddAction reg, "clean.split_column", "Split One Column Into Many", "PkgClean_SplitColumn", "For exports that land everything in column A"
    Core_AddAction reg, "clean.change_case", "Change Case", "PkgClean_ChangeCase", "Normalize text casing across a range. In a chain, this one stops and waits for your answer before it does anything."
End Sub

'================================================================
' ACTION - Remove duplicate rows on the chosen key column(s)
'================================================================
Public Sub PkgClean_RemoveDupes(ctx As Object)
    Dim rng As Range, keys As String, hasHeader As Boolean
    On Error GoTo Fail
    Set rng = Core_AskRange("Remove Duplicates")
    If rng Is Nothing Then Exit Sub                 ' user cancelled

    ' Both questions are asked AFTER the range, and both name it, because both are
    ' about that selection. Each returns False on Cancel and the action then stops
    ' having changed nothing - the range picker's own contract, extended.
    If Not Core_AskHasHeader(rng, hasHeader, "Remove Duplicates") Then Exit Sub
    If Not Core_AskCols(rng, keys, _
            "Which column(s) decide whether two rows are duplicates?", _
            "Remove Duplicates") Then Exit Sub

    Core_SpeedOn
    rng.RemoveDuplicates Columns:=Core_ParseCols(keys), _
                         Header:=IIf(hasHeader, xlYes, xlNo)
    Core_SpeedOff

    ' The message states what was ASSUMED as well as what was done: the header
    ' answer is invisible afterwards, and "10 rows left, not 11" is only
    ' interpretable if you know which way it was answered.
    Core_Done ctx, "Remove Duplicates complete on " & Core_RangeRef(rng) & _
           " using key column(s) " & keys & "." & vbCrLf & _
           IIf(hasHeader, "The first row was treated as a header and left alone.", _
                          "Every row was compared, including the first."), _
           "Remove Duplicates"
    Exit Sub
Fail:
    Core_Fail ctx, "clean.dupes", Err
End Sub

'================================================================
' ACTION - Trim & clean every text cell in the chosen range
'================================================================
Public Sub PkgClean_TrimClean(ctx As Object)
    Dim rng As Range, arr As Variant, r As Long, c As Long
    On Error GoTo Fail
    Set rng = Core_AskRange("Trim & Clean Text")
    If rng Is Nothing Then Exit Sub

    Dim cleaned As String
    Core_SpeedOn
    arr = rng.Value
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            For c = 1 To UBound(arr, 2)
                If VarType(arr(r, c)) = vbString Then
                    cleaned = Core_CleanText(CStr(arr(r, c)))
                    ' write only changed, non-formula cells - never flatten a formula
                    If cleaned <> CStr(arr(r, c)) Then
                        If Not rng.Cells(r, c).HasFormula Then rng.Cells(r, c).Value = cleaned
                    End If
                End If
            Next c
        Next r
    ElseIf VarType(arr) = vbString Then              ' single-cell range
        If Not rng.HasFormula Then rng.Value = Core_CleanText(CStr(arr))
    End If
    Core_SpeedOff

    Core_Done ctx, "Trimmed & cleaned text on " & Core_RangeRef(rng) & ".", _
           "Trim & Clean Text"
    Exit Sub
Fail:
    Core_Fail ctx, "clean.trim", Err
End Sub

'================================================================
' ACTION - Fill each blank cell with the value above it
'================================================================
Public Sub PkgClean_FillDown(ctx As Object)
    Dim rng As Range, arr As Variant, r As Long, c As Long
    On Error GoTo Fail
    Set rng = Core_AskRange("Fill Blanks Down")
    If rng Is Nothing Then Exit Sub

    Core_SpeedOn
    arr = rng.Value
    If IsArray(arr) Then
        For c = 1 To UBound(arr, 2)
            For r = 2 To UBound(arr, 1)
                If Core_IsBlank(arr(r, c)) Then
                    ' fill real blank cells only; leave formula cells untouched
                    If Not rng.Cells(r, c).HasFormula Then rng.Cells(r, c).Value = arr(r - 1, c)
                    arr(r, c) = arr(r - 1, c)          ' propagate so consecutive blanks cascade
                End If
            Next r
        Next c
    End If
    Core_SpeedOff

    Core_Done ctx, "Filled blanks down on " & Core_RangeRef(rng) & ".", _
           "Fill Blanks Down"
    Exit Sub
Fail:
    Core_Fail ctx, "clean.filldown", Err
End Sub

'================================================================
' ACTION - Delete completely blank rows and columns
'
' Two counts, not one, and the difference matters. A row can be blank
' INSIDE the picked range while holding data outside it; deleting the
' whole row would then destroy something the user never selected. Those
' rows are counted, reported, and left alone. Only rows blank across the
' sheet's whole used extent - and free of merged cells - are deleted.
'================================================================
Public Sub PkgClean_RemoveBlanks(ctx As Object)
    Dim rng As Range, ws As Worksheet
    Dim rowsIn As Collection, rowsDel As Collection
    Dim colsIn As Collection, colsDel As Collection
    Dim i As Long
    Const TITLE As String = "Remove Blank Rows & Columns"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                  ' user cancelled
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        MsgBox "The range you picked is entirely outside the used area of the sheet, so " & _
               "there is nothing in it to examine." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbInformation, TITLE
        Exit Sub
    End If
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
'   rowsDel - blank inside AND across the sheet's used extent AND touching no
'             merged block (safe to delete)
' and the same pair for columns. Numbers are collected ascending; the caller
' deletes in reverse.
'
' CountA = 0 is NOT on its own a licence to delete. A row that holds no values can
' still be one stripe of a merged block - the middle row of a three-row merged
' title is the everyday case - and deleting it damages or destroys that merge,
' including the part of it sitting outside the user's selection. Such a row stays
' out of rowsDel, which automatically lands it in the held-back count the caller
' reports, so the user is told rather than surprised.
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
            If Application.CountA(ws.Rows(r)) = 0 Then
                If Not PkgClean_HasMerge(ws.Rows(r)) Then rowsDel.Add r
            End If
        End If
    Next r
    For c = rng.Column To rng.Column + rng.Columns.Count - 1
        If Application.CountA(Application.Intersect(ws.Columns(c), rng)) = 0 Then
            colsIn.Add c
            If Application.CountA(ws.Columns(c)) = 0 Then
                If Not PkgClean_HasMerge(ws.Columns(c)) Then colsDel.Add c
            End If
        End If
    Next c
End Sub

' True when ANY cell in rng belongs to a merged block.
'
' Range.MergeCells is a Variant with three answers, not two: False when no cell in
' the range is merged, True when every cell is, and NULL when the range is mixed -
' which is the ordinary answer for a whole row carrying one merged block, and the
' case that matters most here. Null is tested first because CBool(Null) raises
' error 94, so a plain CBool(rng.MergeCells) would blow up on exactly the input
' this function exists to recognise.
Private Function PkgClean_HasMerge(rng As Range) As Boolean
    Dim v As Variant
    v = rng.MergeCells
    If IsNull(v) Then
        PkgClean_HasMerge = True
    Else
        PkgClean_HasMerge = CBool(v)
    End If
End Function

' The preview the deferred destructive items lack. vbYesNo, not vbYesNoCancel:
' Windows DISABLES the title-bar X on a message box that has no Cancel button, so
' this dialog cannot be dismissed at all - Esc does nothing either, and it has to be
' answered Yes or No. That is what makes two buttons safe HERE: there is no
' accidental exit that could be mistaken for consent to delete. (clean.dupes carries
' the third button for the opposite need - with only Yes and No, a user who wants to
' abandon the action has to pick one of two answers that both act.) Defaults to No.
Private Function PkgClean_ConfirmBlanks(rng As Range, rowsIn As Collection, rowsDel As Collection, _
                                        colsIn As Collection, colsDel As Collection, _
                                        ByVal title As String) As Boolean
    Dim msg As String, held As String
    msg = "Inside " & Core_RangeRef(rng) & " this will delete:" & vbCrLf & vbCrLf & _
          "   " & rowsDel.Count & " completely blank row(s)" & vbCrLf & _
          "   " & colsDel.Count & " completely blank column(s)" & vbCrLf & vbCrLf
    held = PkgClean_HeldBackNote(rowsIn.Count - rowsDel.Count, colsIn.Count - colsDel.Count)
    If Len(held) > 0 Then msg = msg & held & vbCrLf & vbCrLf
    ' Spell out what "blank" means. It is decided on CONTENT only - no values, no
    ' formulas - and nothing here inspects formatting or notes, so a row that looks
    ' empty but carries a fill colour or a comment counts as blank and goes.
    msg = msg & "Blank means no values and no formulas. Formatting or notes on an " & _
          "otherwise empty row or column go with it." & vbCrLf & vbCrLf
    msg = msg & "Whole rows and columns are removed, so everything below and to the " & _
          "right shifts up and left. This cannot be undone with Ctrl+Z." & vbCrLf & vbCrLf & _
          "Delete them?"
    PkgClean_ConfirmBlanks = (MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, title) = vbYes)
End Function

' Names what was deliberately NOT deleted. Returns "" when nothing was held back,
' so the caller can test Len() rather than compare against a sentence.
'
' Two reasons land a row here and the sentence names both, because the user cannot
' tell them apart from the count: it holds data outside the selection, or it is part
' of a merged block. Deleting either changes something outside what was picked.
Private Function PkgClean_HeldBackNote(ByVal heldRows As Long, ByVal heldCols As Long) As String
    If heldRows = 0 And heldCols = 0 Then Exit Function
    PkgClean_HeldBackNote = _
        heldRows & " row(s) and " & heldCols & " column(s) are blank inside your " & _
        "selection but were left alone - each one either holds data outside your " & _
        "selection or is part of a merged block, and deleting it would have changed " & _
        "something you did not select."
End Function

'================================================================
' ACTION - Unmerge every merged block and fill its value down
'================================================================
Public Sub PkgClean_UnmergeFill(ctx As Object)
    Dim rng As Range, c As Range, seen As Object, addrs As Collection
    Dim area As Range, i As Long, v As Variant, fml As String, hadFormula As Boolean
    Dim straddling As Long
    Const TITLE As String = "Unmerge & Fill Down"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        MsgBox "The range you picked is entirely outside the used area of the sheet, so " & _
               "there are no merged cells in it." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbInformation, TITLE
        Exit Sub
    End If

    Core_SpeedOn
    ' Collect the merge addresses FIRST. Unmerging inside the loop that is
    ' walking the cells changes what MergeArea reports mid-iteration.
    Set seen = CreateObject("Scripting.Dictionary")
    Set addrs = New Collection
    For Each c In rng.Cells
        If c.MergeCells Then
            If Not seen.Exists(c.MergeArea.Address) Then
                seen.Add c.MergeArea.Address, 1
                ' Only blocks WHOLLY INSIDE the selection. A merge that straddles the
                ' edge would otherwise be unmerged and overwritten outside the range the
                ' user picked - the one thing every action here promises not to do. They
                ' are counted and named in the report instead, so the user can widen the
                ' selection and run it again rather than wonder why nothing happened.
                If PkgClean_Contains(rng, rng.Parent.Range(c.MergeArea.Address)) Then
                    addrs.Add c.MergeArea.Address
                Else
                    straddling = straddling + 1
                End If
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
        MsgBox "There are no merged cells wholly inside " & Core_RangeRef(rng) & "." & _
               vbCrLf & vbCrLf & PkgClean_StraddleNote(straddling) & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbInformation, TITLE
        Exit Sub
    End If

    Core_Done ctx, "Unmerged " & addrs.Count & " merged block(s) in " & Core_RangeRef(rng) & _
           " and filled each one's value into every cell it covered." & _
           IIf(straddling > 0, vbCrLf & vbCrLf & PkgClean_StraddleNote(straddling), ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "clean.unmerge_fill", Err
End Sub

' Names the merged blocks that were left alone, and why. Returns "" when there were
' none, so a caller can drop the sentence entirely.
Private Function PkgClean_StraddleNote(ByVal n As Long) As String
    If n = 0 Then Exit Function
    PkgClean_StraddleNote = n & " merged block(s) cross the edge of your selection and " & _
        "were left alone - unmerging one would have changed cells outside it. Select the " & _
        "whole block and run this again."
End Function

' Is inner entirely within outer? Used to keep an edit inside the picked range.
Private Function PkgClean_Contains(outer As Range, inner As Range) As Boolean
    If outer.Parent.name <> inner.Parent.name Then Exit Function
    PkgClean_Contains = (inner.Row >= outer.Row) And _
        (inner.Row + inner.Rows.Count - 1 <= outer.Row + outer.Rows.Count - 1) And _
        (inner.Column >= outer.Column) And _
        (inner.Column + inner.Columns.Count - 1 <= outer.Column + outer.Columns.Count - 1)
End Function

'================================================================
' ACTION - Apply the configured number formats to the picked range
'================================================================
Public Sub PkgClean_StandardizeFormats(ctx As Object)
    Dim rng As Range, cols() As String, fmts() As String, spare() As String
    Dim i As Long, idx As Long, tgt As Range, applied As Long, missed As String
    Const TITLE As String = "Standardize Formats"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub

    If Not Core_CfgTable(ctx, "k_clean_formats", cols, fmts, spare) Then
        Core_Refuse ctx, "No column formats are set up yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Column formats' table on the Setup tab, Finalize, and run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    For i = LBound(cols) To UBound(cols)
        idx = Core_ResolveColumn(rng, cols(i))
        If idx = 0 Then
            Set tgt = Nothing
        Else
            Set tgt = rng.Columns(idx)
        End If
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

    Set rng = Core_AskRange(TITLE)
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

'================================================================
' ACTION - Split one selected column into text fields
'================================================================
Public Sub PkgClean_SplitColumn(ctx As Object)
    Dim src As Range, ws As Worksheet, destCell As Range, target As Range
    Dim scanRange As Range, c As Range
    Dim delimiter As String, destSpec As String, firstOccupied As String
    Dim sourceRef As String, destNote As String
    Dim destCol As Long, maxFields As Long, fields As Long, i As Long
    Dim formulaState As Variant, fieldInfo() As Variant
    Const TITLE As String = "Split One Column Into Many"
    On Error GoTo Fail

    Set src = Core_AskRange(TITLE)
    If src Is Nothing Then Exit Sub
    If src.Columns.Count <> 1 Then
        Core_Refuse ctx, "This action accepts exactly one selected column." & vbCrLf & vbCrLf & _
               "No Setup change is needed; run it again and pick one column only." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set src = Core_NarrowToUsed(src)
    If src Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to split." & vbCrLf & vbCrLf & _
               "No Setup change is needed; run it again and pick a populated column." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If
    Set ws = src.Parent
    sourceRef = Core_RangeRef(src)

    delimiter = PkgClean_Delimiter(Core_Cfg(ctx, "k_clean_split_delimiter"))
    If Len(delimiter) = 0 Then
        Core_Refuse ctx, "The split-column delimiter is not valid." & vbCrLf & vbCrLf & _
               "Set 'Split-column delimiter' on the Setup tab to one character, or to " & _
               "COMMA, TAB, SEMICOLON, PIPE, or SPACE, Finalize, and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    destSpec = Trim$(Core_Cfg(ctx, "k_clean_split_destination"))
    destCol = PkgClean_DestinationColumn(ws, destSpec, src.Column)
    If destCol = 0 Then
        Core_Refuse ctx, "The split-column destination is not valid." & vbCrLf & vbCrLf & _
               "Set 'Split-column destination' on the Setup tab to SOURCE or an absolute " & _
               "worksheet column letter, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    formulaState = src.HasFormula
    If IsNull(formulaState) Or (formulaState = True) Then
        Core_Refuse ctx, "The selected column contains one or more formulas. Splitting it " & _
               "would replace those formulas with text values." & vbCrLf & vbCrLf & _
               "No Setup change is needed; pick a values-only export column and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    For Each c In src.Cells
        fields = PkgClean_FieldCount(c.Value2, delimiter)
        If fields > maxFields Then maxFields = fields
    Next c
    If maxFields <= 1 Then
        Core_Refuse ctx, "The configured delimiter was not found in the selected column." & _
               vbCrLf & vbCrLf & "Check 'Split-column delimiter' on the Setup tab, " & _
               "Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If destCol + maxFields - 1 > ws.Columns.Count Then
        Core_Refuse ctx, "The split needs " & maxFields & " destination columns, but they " & _
               "would run past the last column on the worksheet." & vbCrLf & vbCrLf & _
               "Set 'Split-column destination' to an earlier column on the Setup tab, " & _
               "Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set destCell = ws.Cells(src.Row, destCol)
    Set target = ws.Range(destCell, _
        ws.Cells(src.Row + src.Rows.Count - 1, destCol + maxFields - 1))
    If destCol = src.Column Then
        Set scanRange = ws.Range(ws.Cells(src.Row, src.Column + 1), _
            ws.Cells(src.Row + src.Rows.Count - 1, src.Column + maxFields - 1))
    Else
        Set scanRange = target
    End If

    If Application.CountA(scanRange) > 0 Then
        For Each c In scanRange.Cells
            If Not IsEmpty(c.Value2) Or c.HasFormula Then
                firstOccupied = c.Address(False, False)
                Exit For
            End If
        Next c
    End If

    If Len(firstOccupied) > 0 Then
        Core_Refuse ctx, "Splitting would overwrite existing content or a formula at " & _
               firstOccupied & "." & vbCrLf & vbCrLf & _
               "Clear the intended output area, or change 'Split-column destination' on " & _
               "the Setup tab, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ReDim fieldInfo(0 To maxFields - 1)
    For i = 0 To maxFields - 1
        fieldInfo(i) = Array(i + 1, xlTextFormat)
    Next i

    Core_SpeedOn
    src.TextToColumns Destination:=destCell, DataType:=xlDelimited, _
        TextQualifier:=xlTextQualifierDoubleQuote, ConsecutiveDelimiter:=False, _
        Tab:=False, Semicolon:=False, Comma:=False, Space:=False, _
        Other:=True, OtherChar:=delimiter, FieldInfo:=fieldInfo
    Core_SpeedOff

    If destCol = src.Column Then
        destNote = "The selected source column was replaced by the split fields."
    Else
        destNote = "The source column was left unchanged."
    End If
    Core_Done ctx, "Split " & src.Rows.Count & " row(s) from " & sourceRef & _
           " into as many as " & maxFields & " text column(s), starting at " & _
           destCell.Address(False, False) & "." & vbCrLf & vbCrLf & destNote, TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "clean.split_column", Err
End Sub

Private Function PkgClean_Delimiter(ByVal raw As String) As String
    Dim token As String
    token = UCase$(Trim$(raw))

    Select Case token
        Case "COMMA"
            PkgClean_Delimiter = ","
        Case "TAB"
            PkgClean_Delimiter = vbTab
        Case "SEMICOLON"
            PkgClean_Delimiter = ";"
        Case "PIPE"
            PkgClean_Delimiter = "|"
        Case "SPACE"
            PkgClean_Delimiter = " "
        Case Else
            If Len(raw) = 1 Then PkgClean_Delimiter = raw
    End Select
End Function

Private Function PkgClean_DestinationColumn(ws As Worksheet, ByVal spec As String, _
                                            ByVal sourceColumn As Long) As Long
    Dim token As String, i As Long, colNum As Long
    token = UCase$(Trim$(spec))

    If token = "SOURCE" Then
        PkgClean_DestinationColumn = sourceColumn
        Exit Function
    End If
    If Len(token) = 0 Then Exit Function
    If Len(token) > 3 Then Exit Function
    If Not Core_IsAlpha(token) Then Exit Function

    For i = 1 To Len(token)
        colNum = colNum * 26 + Asc(Mid$(token, i, 1)) - Asc("A") + 1
    Next i
    If colNum <= ws.Columns.Count Then PkgClean_DestinationColumn = colNum
End Function

Private Function PkgClean_FieldCount(ByVal rawValue As Variant, _
                                     ByVal delimiter As String) As Long
    Dim s As String, ch As String, i As Long, n As Long
    Dim inQuotes As Boolean
    n = 1
    PkgClean_FieldCount = n

    If IsError(rawValue) Then Exit Function
    If IsEmpty(rawValue) Then Exit Function
    s = CStr(rawValue)

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If inQuotes Then
                If i < Len(s) Then
                    If Mid$(s, i + 1, 1) = """" Then
                        i = i + 1
                    Else
                        inQuotes = False
                    End If
                Else
                    inQuotes = False
                End If
            Else
                inQuotes = True
            End If
        Else
            If Not inQuotes Then
                If ch = delimiter Then n = n + 1
            End If
        End If
    Next i

    PkgClean_FieldCount = n
End Function

'================================================================
' ACTION - Change the case of constant text cells
'================================================================
Public Sub PkgClean_ChangeCase(ctx As Object)
    Dim rng As Range, arr As Variant, converted As String, caseMode As String
    Dim r As Long, c As Long, textCount As Long, changed As Long
    Dim formulaCount As Long
    Const TITLE As String = "Change Case"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is no text in it to change." & vbCrLf & vbCrLf & _
               "No Setup change is needed; run it again and pick a populated range." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not PkgClean_AskCase(caseMode, TITLE) Then Exit Sub

    Core_SpeedOn
    arr = rng.Value2
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            For c = 1 To UBound(arr, 2)
                If rng.Cells(r, c).HasFormula Then
                    formulaCount = formulaCount + 1
                ElseIf VarType(arr(r, c)) = vbString Then
                    textCount = textCount + 1
                    Select Case caseMode
                        Case "UPPER"
                            converted = UCase$(CStr(arr(r, c)))
                        Case "lower"
                            converted = LCase$(CStr(arr(r, c)))
                        Case "Proper"
                            converted = StrConv(CStr(arr(r, c)), vbProperCase)
                    End Select
                    If converted <> CStr(arr(r, c)) Then
                        rng.Cells(r, c).Value = converted
                        changed = changed + 1
                    End If
                End If
            Next c
        Next r
    Else
        If rng.HasFormula Then
            formulaCount = 1
        ElseIf VarType(arr) = vbString Then
            textCount = 1
            Select Case caseMode
                Case "UPPER"
                    converted = UCase$(CStr(arr))
                Case "lower"
                    converted = LCase$(CStr(arr))
                Case "Proper"
                    converted = StrConv(CStr(arr), vbProperCase)
            End Select
            If converted <> CStr(arr) Then
                rng.Value = converted
                changed = 1
            End If
        End If
    End If
    Core_SpeedOff

    Core_Done ctx, "Applied " & caseMode & " case to " & textCount & _
           " constant text cell(s) in " & Core_RangeRef(rng) & "; " & changed & _
           " cell(s) changed." & vbCrLf & vbCrLf & _
           "Text already in that case was not rewritten. Numbers, errors, blanks, and " & _
           formulaCount & " formula cell(s) were left exactly as they were.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "clean.change_case", Err
End Sub

Private Function PkgClean_AskCase(ByRef caseMode As String, _
                                  ByVal title As String) As Boolean
    Dim answer As Variant, token As String

    Do
        answer = Application.InputBox( _
            Prompt:="Enter UPPER, lower, or Proper." & vbCrLf & vbCrLf & _
                    "Only constant text cells will be changed.", _
            title:=title, Default:="Proper", Type:=2)

        If VarType(answer) = vbBoolean Then
            If answer = False Then Exit Function
        End If

        token = LCase$(Trim$(CStr(answer)))
        Select Case token
            Case "upper"
                caseMode = "UPPER"
                PkgClean_AskCase = True
                Exit Function
            Case "lower"
                caseMode = "lower"
                PkgClean_AskCase = True
                Exit Function
            Case "proper"
                caseMode = "Proper"
                PkgClean_AskCase = True
                Exit Function
            Case Else
                MsgBox "Enter UPPER, lower, or Proper, or click Cancel to stop.", _
                       vbInformation, title
        End Select
    Loop
End Function
