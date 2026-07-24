Attribute VB_Name = "PkgClean"
Option Explicit

'================================================================
' PkgClean - a small second package, built to prove Core is a
' platform: a brand-new package's config field and buttons appear
' on the Setup and Macros tabs with ZERO changes to Core beyond
' adding "Clean" to Core's candidate list. All plumbing (config,
' speed, errors) is Core_*; the only package-specific need - "which
' range?" - is answered ad hoc with a range picker, so Clean needs
' no configured target sheet.
'
' Depends on: Core (required).
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgClean_Manifest(reg As Object)
    Core_DeclarePackage reg, "clean", "Clean & Standardize", "1.0.0"

    ' (No section header: the package banner already reads "CLEAN & STANDARDIZE".)
    Core_AddField reg, "clean_keycols", "Remove-duplicates key column(s)", "text", "1", _
        help:="Which column(s) decide whether two rows are duplicates, numbered 1,2,3... from the LEFT of the range you pick. One number, or several like 1,2,4."

    Core_AddAction reg, "clean.dupes", "Remove Duplicates", "PkgClean_RemoveDupes", "Pick a range (include the header row); removes duplicate rows using the key column(s) from Setup."
    Core_AddAction reg, "clean.trim", "Trim & Clean Text", "PkgClean_TrimClean", "Pick a range; strips leading/trailing spaces, collapses double spaces, and removes non-printing characters."
    Core_AddAction reg, "clean.filldown", "Fill Blanks Down", "PkgClean_FillDown", "Pick a range; fills each blank cell with the value from the cell above it."
End Sub

'================================================================
' ACTION - Remove duplicate rows on the chosen key column(s)
'================================================================
Public Sub PkgClean_RemoveDupes(ctx As Object)
    Dim rng As Range, keys As String
    On Error GoTo Fail
    Set rng = AskRange("Remove Duplicates")
    If rng Is Nothing Then Exit Sub                 ' user cancelled

    keys = Core_Cfg(ctx, "clean_keycols")
    Core_SpeedOn
    rng.RemoveDuplicates Columns:=ParseCols(keys), Header:=xlYes
    Core_SpeedOff

    Core_Done ctx, "Remove Duplicates complete on " & rng.Address(False, False) & _
           " using key column(s) " & keys & ".", "Remove Duplicates"
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

    Core_Done ctx, "Trimmed & cleaned text on " & rng.Address(False, False) & ".", _
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

    Core_Done ctx, "Filled blanks down on " & rng.Address(False, False) & ".", _
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
        Prompt:="Select the range for '" & title & "' (include the header row):", _
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
