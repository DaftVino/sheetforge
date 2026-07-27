Attribute VB_Name = "PkgClean"
Option Explicit

'================================================================
' PkgClean - a small second package, built to prove Core is a
' platform: a brand-new package's buttons appear on the Macros tabs
' with ZERO changes to Core beyond adding "Clean" to Core's candidate
' list. All plumbing (speed, errors, completion dialogs) is Core_*.
'
' As of v0.5.0 this package declares NO config fields and therefore
' renders no band on Setup at all - it is now also the proof that a
' package needing no configuration costs the user no Setup surface.
' Everything it needs describes the range the user picks, so it is
' asked for at run time. See PkgClean_Manifest.
'
' Depends on: Core (required).
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgClean_Manifest(reg As Object)
    Core_DeclarePackage reg, "clean", "Clean & Standardize", "1.0.0"

    ' This package declares NO fields, deliberately, so it renders no band on Setup.
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

    Core_AddAction reg, "clean.dupes", "Remove Duplicates", "PkgClean_RemoveDupes", "Pick a range, then answer two questions about it: whether its first row is a header, and which column(s) decide a duplicate."
    Core_AddAction reg, "clean.trim", "Trim & Clean Text", "PkgClean_TrimClean", "Pick a range; strips leading/trailing spaces, collapses double spaces, and removes non-printing characters."
    Core_AddAction reg, "clean.filldown", "Fill Blanks Down", "PkgClean_FillDown", "Pick a range; fills each blank cell with the value from the cell above it."
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
' Package-private helpers
'================================================================

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
