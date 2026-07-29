Attribute VB_Name = "PkgClean"
Option Explicit

'================================================================
' PkgClean - a small second package, built to prove Core is a
' platform: a brand-new package's buttons appear on the Macros tabs
' with ZERO changes to Core beyond adding "Clean" to Core's candidate
' list. All plumbing (speed, errors, completion dialogs) is Core_*.
'
' From v0.5.0 to v2.2.0 this package declared NO config fields and rendered
' no band on Setup at all. v2.3.0 adds exactly one - k_clean_formats, a
' deployment-level number-format map - so the band exists again, holding
' that one table and nothing else. The point it used to prove still holds:
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

    ' This package declares exactly ONE field - k_clean_formats, below - and from
    ' v0.5.0 to v2.2.0 it declared none at all and rendered no band on Setup.
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

    Core_AddAction reg, "clean.dupes", "Remove Duplicates", "PkgClean_RemoveDupes", "Pick a range, then answer two questions about it: whether its first row is a header, and which column(s) decide a duplicate."
    Core_AddAction reg, "clean.trim", "Trim & Clean Text", "PkgClean_TrimClean", "Pick a range; strips leading/trailing spaces, collapses double spaces, and removes non-printing characters."
    Core_AddAction reg, "clean.filldown", "Fill Blanks Down", "PkgClean_FillDown", "Pick a range; fills each blank cell with the value from the cell above it."
    Core_AddAction reg, "clean.remove_blanks", "Remove Blank Rows & Columns", "PkgClean_RemoveBlanks", "Pick a range; shows you how many completely blank rows and columns it found, and deletes them only after you say yes."
    Core_AddAction reg, "clean.unmerge_fill", "Unmerge & Fill Down", "PkgClean_UnmergeFill", "Pick a range; unmerges every merged block and puts the block's value into every cell it covered, so the data can be sorted and filtered."
    Core_AddAction reg, "clean.standardize_formats", "Standardize Formats", "PkgClean_StandardizeFormats", "Pick a range; applies the number formats listed in the Column formats table on Setup. Changes how values look, never the values themselves."
    Core_AddAction reg, "clean.text_to_numbers", "Convert Text to Numbers", "PkgClean_TextToNumbers", "Pick a range; turns text that represents a number into a real number, stripping $ and thousands separators. Formulas are never touched. Note that 12% becomes 12, not 0.12."
End Sub

'================================================================
' ACTION - Remove duplicate rows on the chosen key column(s)
'================================================================
Public Sub PkgClean_RemoveDupes(ctx As Object)
    Dim rng As Range, keys As String, hasHeader As Boolean
    On Error GoTo Fail
    Set rng = AskRange("Remove Duplicates")
    If rng Is Nothing Then Exit Sub                 ' user cancelled

    ' Both questions are asked AFTER the range, and both name it, because both are
    ' about that selection. Each returns False on Cancel and the action then stops
    ' having changed nothing - the range picker's own contract, extended.
    If Not AskHasHeader(rng, hasHeader) Then Exit Sub
    If Not AskKeyCols(rng, keys) Then Exit Sub

    Core_SpeedOn
    rng.RemoveDuplicates Columns:=ParseCols(keys), _
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
    Set rng = AskRange("Trim & Clean Text")
    If rng Is Nothing Then Exit Sub

    Dim cleaned As String
    Core_SpeedOn
    arr = rng.Value
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            For c = 1 To UBound(arr, 2)
                If VarType(arr(r, c)) = vbString Then
                    cleaned = CleanText(CStr(arr(r, c)))
                    ' write only changed, non-formula cells - never flatten a formula
                    If cleaned <> CStr(arr(r, c)) Then
                        If Not rng.Cells(r, c).HasFormula Then rng.Cells(r, c).Value = cleaned
                    End If
                End If
            Next c
        Next r
    ElseIf VarType(arr) = vbString Then              ' single-cell range
        If Not rng.HasFormula Then rng.Value = CleanText(CStr(arr))
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
    Set rng = AskRange("Fill Blanks Down")
    If rng Is Nothing Then Exit Sub

    Core_SpeedOn
    arr = rng.Value
    If IsArray(arr) Then
        For c = 1 To UBound(arr, 2)
            For r = 2 To UBound(arr, 1)
                If IsBlankVal(arr(r, c)) Then
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

    Set rng = AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                  ' user cancelled
    Set rng = PkgClean_NarrowToUsed(rng)
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

    Set rng = AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = PkgClean_NarrowToUsed(rng)
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
    Dim i As Long, tgt As Range, applied As Long, missed As String
    Const TITLE As String = "Standardize Formats"
    On Error GoTo Fail

    Set rng = AskRange(TITLE)
    If rng Is Nothing Then Exit Sub

    If Not Core_CfgTable(ctx, "k_clean_formats", cols, fmts, spare) Then
        PkgClean_Refuse ctx, "No column formats are set up yet."
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

'================================================================
' Package-private helpers
'================================================================

' Narrow a picked range to the part of the sheet that actually holds anything.
' Returns Nothing when the pick lies entirely outside the used area - the caller
' reports that and stops, having changed nothing.
'
' Selecting whole columns is how people pick "this table" - A:D is one click on the
' column headers - and it hands these two actions roughly four million cells. Both
' do COM work PER CELL or per row (an Intersect and a CountA each, a MergeArea read
' each), so the screen freezes for minutes, and clean.remove_blanks then asks Excel
' to delete about a million rows one at a time inside its destructive block. Nothing
' outside UsedRange can hold a value, a formula or a merge, so the intersection
' examines the same data and finishes in seconds.
'
' Everything downstream uses the NARROWED range, including the range named in the
' messages, so what the user is told matches what was actually looked at.
Private Function PkgClean_NarrowToUsed(rng As Range) As Range
    If rng Is Nothing Then Exit Function
    Set PkgClean_NarrowToUsed = Application.Intersect(rng, rng.Parent.UsedRange)
End Function

' Mark ctx as failed on a refusal - a path that shows its reason and returns WITHOUT
' doing the work. Core's RunAction decides a step succeeded by the ABSENCE of this
' flag, so a refusal that only shows a MsgBox is reported to a chain as a green tick
' beside a step that did nothing. Writes wrapped the way Core_Fail wraps its own, so
' a ctx that never received these keys cannot raise from inside a refusal path.
'
' The MsgBox stays, quiet or not: a chain that stops has to say why.
Private Sub PkgClean_Refuse(ctx As Object, ByVal reason As String)
    On Error Resume Next
    ctx("failed") = True
    ctx("failMsg") = reason
    On Error GoTo 0
End Sub

' Prompt the user to select a range. Returns Nothing if they cancel
' (InputBox Type:=8 returns Boolean False on cancel, which fails the Set).
Private Function AskRange(ByVal title As String) As Range
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select the range for '" & title & "':", _
        title:=title, Type:=8)
    On Error GoTo 0
    If Not rng Is Nothing Then
        If rng.Areas.Count > 1 Then
            MsgBox "Please select a single, continuous block (not several separate areas).", _
                   vbExclamation, title
            Set rng = Nothing
        End If
    End If
    Set AskRange = rng
End Function

' Ask whether the picked range's first row is a header. Returns False on Cancel,
' leaving hasHeader untouched and the caller responsible for stopping.
'
' vbYesNoCancel, not vbYesNo: a two-button dialog has no way out, and the title-bar
' X on a vbYesNo MsgBox returns vbNo - which here would silently choose "compare
' every row" and delete a header row. Cancel has to be reachable and distinct.
Private Function AskHasHeader(rng As Range, ByRef hasHeader As Boolean) As Boolean
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Does " & Core_RangeRef(rng) & " include a header row?" & vbCrLf & vbCrLf & _
                 "Yes  -  the first row is a title row: it is kept, and never " & _
                 "compared against the data." & vbCrLf & _
                 "No  -  every row is data: all of them are compared, including " & _
                 "the first one.", _
                 vbQuestion + vbYesNoCancel + vbDefaultButton1, "Remove Duplicates")
    If ans = vbCancel Then Exit Function
    hasHeader = (ans = vbYes)
    AskHasHeader = True
End Function

' Ask which column(s) decide a duplicate, numbered from the LEFT of the picked
' range. Returns False on Cancel.
'
' Application.InputBox, not VBA's InputBox: the latter returns "" for BOTH Cancel
' and an empty OK, and those must not mean the same thing. Application.InputBox
' returns Boolean False on Cancel, which VarType can tell apart from any string.
Private Function AskKeyCols(rng As Range, ByRef keys As String) As Boolean
    Dim v As Variant, cols As Variant, i As Long, nCols As Long
    nCols = rng.Columns.Count
    v = Application.InputBox( _
        Prompt:="Which column(s) decide whether two rows are duplicates?" & vbCrLf & vbCrLf & _
                "Numbered from the LEFT of " & Core_RangeRef(rng) & ", which is " & _
                nCols & " column(s) wide. One number, or several like 1,2,4.", _
        title:="Remove Duplicates", Default:="1", Type:=2)
    If VarType(v) = vbBoolean Then Exit Function     ' Cancel
    keys = Trim(CStr(v))
    If Len(keys) = 0 Then keys = "1"                 ' empty OK means the default

    ' Validate against the picked range's width. A column number past the right
    ' edge raises 1004 inside RemoveDuplicates, which surfaces through Core_Fail
    ' as an engine error - it reads as a broken macro rather than as the typo it
    ' is. ParseCols already drops non-numeric entries and falls back to column 1.
    cols = ParseCols(keys)
    For i = LBound(cols) To UBound(cols)
        If cols(i) < 1 Or cols(i) > nCols Then
            MsgBox "Column " & cols(i) & " is outside " & Core_RangeRef(rng) & _
                   ", which is " & nCols & " column(s) wide." & vbCrLf & vbCrLf & _
                   "Nothing has been changed. Run Remove Duplicates again.", _
                   vbExclamation, "Remove Duplicates"
            Exit Function
        End If
    Next i
    AskKeyCols = True
End Function

' "1,2,4" -> Variant array of Longs for Range.RemoveDuplicates Columns.
' Non-numeric entries are ignored; an empty result defaults to column 1.
Private Function ParseCols(ByVal s As String) As Variant
    Dim parts() As String, i As Long, c As Collection, out() As Variant, k As Long
    Set c = New Collection
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        If IsNumeric(Trim(parts(i))) Then c.Add CLng(Trim(parts(i)))
    Next i
    If c.Count = 0 Then c.Add CLng(1)
    ReDim out(0 To c.Count - 1)
    For k = 1 To c.Count
        out(k - 1) = c(k)
    Next k
    ParseCols = out
End Function

' Excel TRIM (collapse spaces) + CLEAN (strip non-printing chars).
Private Function CleanText(ByVal s As String) As String
    CleanText = Application.WorksheetFunction.Clean(Application.WorksheetFunction.Trim(s))
End Function

Private Function IsBlankVal(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankVal = True
    ElseIf VarType(v) = vbString Then
        IsBlankVal = (v = "")
    End If
End Function
