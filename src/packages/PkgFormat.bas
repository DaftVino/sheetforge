Attribute VB_Name = "PkgFormat"
Option Explicit

'================================================================
' PkgFormat - presentation only. Every action here changes how a
' report LOOKS - fonts, fills, borders, row layout, print setup -
' and never a stored value (contract section 7: the format domain
' "changes no value"). Two actions still touch the sheet beyond a
' cell's own formatting: report_title inserts rows, and table_style
' / conditional_highlight overwrite existing formatting - each says
' so and asks before it runs.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST - Core reads this to learn the package's fields + actions.
'================================================================
Public Sub PkgFormat_Manifest(reg As Object)
    Core_DeclarePackage reg, "format", "Format & Present", "1.0.0"

    Core_AddSection reg, "CONDITIONAL HIGHLIGHT"
    Core_AddTable reg, "k_format_highlight_rules", "Highlight rules", _
        "Column (letter or header)|Condition|Fill color", "simple", _
        "One row per rule. Column: a letter like C, or the column's header text (a " & _
        "one-to-three-letter entry is always read as a letter). Condition: greater_than " & _
        "<value>, less_than <value>, equal_to <value>, contains <value>, or blank (no value " & _
        "needed) - e.g. 'greater_than 1000' or 'contains VIP'. Fill color: red, yellow, " & _
        "green, orange, blue, or gray. Rules are checked top to bottom; if more than one " & _
        "matches the same cell, the LAST match wins. Nothing happens until you fill this in.", _
        "format.conditional_highlight"

    Core_AddAction reg, "format.conditional_highlight", "Highlight by Rule", "PkgFormat_ConditionalHighlight", _
        "Pick a range; colors cells whose value matches a rule from the 'Highlight rules' " & _
        "table on Setup. Clears any existing conditional formatting in the range first, and " & _
        "asks before doing either."
    Core_AddAction reg, "format.style_headers", "Style Header Row", "PkgFormat_StyleHeaders", _
        "Pick a range; bolds and fills its first row as a header and autofits every column " & _
        "in the range. Overwrites any formatting already on that row."
    Core_AddAction reg, "format.report_title", "Add Report Title", "PkgFormat_ReportTitle", _
        "Type a title; inserts two rows at the top of the active sheet, merges and centers " & _
        "the title across the sheet's used width, and shifts everything else down. Asks first."
    Core_AddAction reg, "format.table_style", "Apply House Table Style", "PkgFormat_TableStyle", _
        "Pick a range; applies banded row fill and thin borders in the house style. " & _
        "Overwrites any fill or border formatting already in the range. Asks first."
    Core_AddAction reg, "format.print_layout", "Prep for Print", "PkgFormat_PrintLayout", _
        "Sets the active sheet to landscape, fit-to-width, narrow margins, and a repeating " & _
        "header row - the board-pack preset. Changes no cell, only how the sheet prints."
End Sub

'================================================================
' ACTION - Style the header row of a picked range, autofit columns
'================================================================
Public Sub PkgFormat_StyleHeaders(ctx As Object)
    Dim rng As Range, hdr As Range
    Const TITLE As String = "Style Header Row"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                   ' user cancelled
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to style." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    Set hdr = rng.Rows(1)

    Core_SpeedOn
    With hdr
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 73, 125)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    rng.Columns.AutoFit
    Core_SpeedOff

    Core_Done ctx, "Styled the header row and autofit " & rng.Columns.Count & " column(s) in " & _
           Core_RangeRef(rng) & "." & vbCrLf & vbCrLf & _
           "Any formatting already on that row was overwritten.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.style_headers", Err
End Sub

'================================================================
' ACTION - Insert a formatted title block above a report
'================================================================
Public Sub PkgFormat_ReportTitle(ctx As Object)
    Dim ws As Worksheet, titleText As Variant, lastCol As Long
    Const TITLE As String = "Add Report Title"
    Const ROWS_INSERTED As Long = 2
    On Error GoTo Fail

    Set ws = ActiveSheet
    If Core_IsEngineSheet(ws.name) Then
        Core_Refuse ctx, "'" & ws.name & "' belongs to the macro toolkit itself, not to a " & _
               "report." & vbCrLf & vbCrLf & "Switch to the report tab and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    titleText = Application.InputBox( _
        Prompt:="Title text for the top of '" & ws.name & "':", title:=TITLE, Type:=2)
    If VarType(titleText) = vbBoolean Then Exit Sub          ' user cancelled
    If Len(Trim(CStr(titleText))) = 0 Then
        Core_Refuse ctx, "No title text was entered." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' UsedRange.Columns.Count is a WIDTH (e.g. 3 for a used range D:F), not the
    ' absolute last column index - adding UsedRange.Column (the FIRST used
    ' column) recovers the actual last column so the merge spans the real
    ' data instead of always starting at A.
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastCol < 1 Then lastCol = 1

    If MsgBox("This inserts " & ROWS_INSERTED & " row(s) at the top of '" & ws.name & _
              "' and shifts everything currently there down by " & ROWS_INSERTED & " row(s)." & _
              vbCrLf & vbCrLf & _
              "Formulas and defined names on this sheet follow the shift automatically - " & _
              "that is Excel's own row-insert behavior, not something this macro adds. A " & _
              "formula whose RESULT depends on its own row (ROW(), or an INDIRECT() built " & _
              "from a row number) can return a different result after the shift, even " & _
              "though its formula text does not change." & vbCrLf & vbCrLf & _
              "The title text you type becomes a new stored value in the inserted title " & _
              "cell; every other cell keeps its own value and only moves." & vbCrLf & vbCrLf & _
              "Insert the title block?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    ws.Rows("1:" & ROWS_INSERTED).Insert Shift:=xlShiftDown
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Merge
        .Value = CStr(titleText)
        .Font.Bold = True
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Rows(1).RowHeight = 24
    Core_SpeedOff

    Core_Done ctx, "Inserted a title block on '" & ws.name & "'." & vbCrLf & vbCrLf & _
           "Everything that was already on this sheet kept its own value and moved down " & _
           ROWS_INSERTED & " row(s). The title text itself is a new stored value, in the " & _
           "merged cell above it, and any formula whose result depends on row position " & _
           "(ROW(), an INDIRECT() built from a row number) may now return a different " & _
           "result.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.report_title", Err
End Sub

'================================================================
' ACTION - Apply the house table style (banded rows + borders)
'================================================================
Public Sub PkgFormat_TableStyle(ctx As Object)
    Dim rng As Range, r As Long, bandColor As Long
    Const TITLE As String = "Apply House Table Style"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                   ' user cancelled
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to style." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If MsgBox("This replaces any fill and border formatting already inside " & _
              Core_RangeRef(rng) & " with the house table style (banded rows + thin " & _
              "borders). It does not change any value." & vbCrLf & vbCrLf & "Apply it?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    bandColor = RGB(237, 241, 245)
    rng.Interior.ColorIndex = xlColorIndexNone
    For r = 1 To rng.Rows.Count
        If r Mod 2 = 0 Then rng.Rows(r).Interior.Color = bandColor
    Next r
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With
    rng.Rows(1).Font.Bold = True
    Core_SpeedOff

    Core_Done ctx, "Applied the house table style to " & Core_RangeRef(rng) & "." & vbCrLf & _
           vbCrLf & "Any fill or border formatting already there was replaced. No value " & _
           "changed.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.table_style", Err
End Sub

'================================================================
' ACTION - Color cells that match a configured rule
'================================================================
Public Sub PkgFormat_ConditionalHighlight(ctx As Object)
    Dim cols() As String, conds() As String, colors() As String
    Dim rng As Range, hasHeader As Boolean, startRow As Long
    Dim i As Long, tgt As Range, matched As Long, missed As String
    Const TITLE As String = "Highlight by Rule"
    On Error GoTo Fail

    If Not Core_CfgTable(ctx, "k_format_highlight_rules", cols, conds, colors) Then
        Core_Refuse ctx, "No highlight rules are configured yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Highlight rules' table on the Setup tab, Finalize, and run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                   ' user cancelled
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to check." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub    ' user cancelled

    If MsgBox("This clears any existing conditional formatting inside " & Core_RangeRef(rng) & _
              " and colors cells against " & (UBound(cols) - LBound(cols) + 1) & " rule(s) " & _
              "from Setup." & vbCrLf & vbCrLf & "Apply it?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    startRow = IIf(hasHeader, 2, 1)

    Core_SpeedOn
    rng.FormatConditions.Delete
    For i = startRow To rng.Rows.Count
        rng.Rows(i).Interior.ColorIndex = xlColorIndexNone
    Next i

    For i = LBound(cols) To UBound(cols)
        Set tgt = PkgFormat_ResolveCol(rng, cols(i))
        If tgt Is Nothing Then
            missed = missed & vbCrLf & "   " & cols(i)
        Else
            matched = matched + PkgFormat_ApplyRule(tgt, startRow, conds(i), colors(i))
        End If
    Next i
    Core_SpeedOff

    Core_Done ctx, "Colored " & matched & " cell(s) inside " & Core_RangeRef(rng) & "." & _
           vbCrLf & "Existing conditional formatting in the range was cleared first." & _
           IIf(Len(missed) > 0, vbCrLf & vbCrLf & "These rows matched no column in your " & _
               "selection and were skipped:" & missed, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.conditional_highlight", Err
End Sub

' Parse ONE "Condition" cell ("greater_than 1000", "contains VIP", "blank")
' and color every matching cell in colRng from startRow down. colRng spans
' the SAME rows as the picked range (PkgFormat_ResolveCol guarantees this),
' so startRow indexes it directly. Returns how many cells it colored.
'
' An unrecognised operator or color colors nothing and returns 0 - the row
' is silently skipped rather than raising, so one typo does not abort the
' whole rule table; the help text spells out the exact five operator tokens
' and six color names for exactly this reason.
Private Function PkgFormat_ApplyRule(colRng As Range, ByVal startRow As Long, _
                                     ByVal condSpec As String, ByVal colorName As String) As Long
    Dim op As String, val As String, sp As Long, i As Long, c As Range, hit As Boolean
    Dim fillColor As Long
    condSpec = Trim(condSpec)
    sp = InStr(condSpec, " ")
    If sp > 0 Then
        op = LCase(Left$(condSpec, sp - 1))
        val = Trim(Mid$(condSpec, sp + 1))
    Else
        op = LCase(condSpec)
        val = ""
    End If
    fillColor = PkgFormat_ColorFromName(colorName)
    If fillColor = -1 Then Exit Function

    For i = startRow To colRng.Rows.Count
        Set c = colRng.Cells(i, 1)
        hit = False
        Select Case op
            Case "blank"
                hit = Core_IsBlank(c.Value)
            Case "equal_to"
                hit = (StrComp(Trim(CStr(c.Value)), val, vbTextCompare) = 0)
            Case "contains"
                hit = (InStr(1, CStr(c.Value), val, vbTextCompare) > 0)
            Case "greater_than"
                ' VBA's And does NOT short-circuit - both operands are always
                ' evaluated, so a single `IsNumeric(...) And CDbl(...)`
                ' expression still calls CDbl on a non-numeric cell and
                ' raises. Nested Ifs guarantee CDbl only runs once both
                ' sides are confirmed numeric.
                If IsNumeric(c.Value) Then
                    If IsNumeric(val) Then
                        hit = (CDbl(c.Value) > CDbl(val))
                    End If
                End If
            Case "less_than"
                If IsNumeric(c.Value) Then
                    If IsNumeric(val) Then
                        hit = (CDbl(c.Value) < CDbl(val))
                    End If
                End If
        End Select
        If hit Then
            c.Interior.Color = fillColor
            PkgFormat_ApplyRule = PkgFormat_ApplyRule + 1
        End If
    Next i
End Function

' Named colors only - a non-technical user typing a color name is more
' reliable than one typing a hex code or a VBA color constant. Returns -1
' for anything else; the caller reads that as "skip this rule".
Private Function PkgFormat_ColorFromName(ByVal nm As String) As Long
    Select Case LCase(Trim(nm))
        Case "red":    PkgFormat_ColorFromName = RGB(255, 199, 206)
        Case "yellow": PkgFormat_ColorFromName = RGB(255, 235, 156)
        Case "green":  PkgFormat_ColorFromName = RGB(198, 239, 206)
        Case "orange": PkgFormat_ColorFromName = RGB(252, 213, 180)
        Case "blue":   PkgFormat_ColorFromName = RGB(189, 215, 238)
        Case "gray", "grey": PkgFormat_ColorFromName = RGB(217, 217, 217)
        Case Else: PkgFormat_ColorFromName = -1
    End Select
End Function

' Resolve one "Column" cell against the picked range, returned as a range
' spanning the SAME ROWS as rng (one column wide) so a caller can index it
' by row number regardless of where the header row is.
'
' Same letter-or-header logic as PkgClean_ResolveFormatCol - a DELIBERATE
' duplicate, not an oversight. See this wave's plan, Decision 8: the
' helper-promotion pass declined to promote the original with "each is one
' package's own job", a call made before this was a second consumer. Not
' re-litigated here; flagged instead for whichever wave adds a third.
Private Function PkgFormat_ResolveCol(rng As Range, ByVal spec As String) As Range
    Dim s As String, absCol As Long, i As Long, hdr As String
    s = Trim(spec)
    If Len(s) = 0 Then Exit Function

    If Len(s) <= 3 And Core_IsAlpha(s) Then
        On Error Resume Next
        absCol = rng.Parent.Range(s & "1").Column
        On Error GoTo 0
        If absCol > 0 Then
            If absCol >= rng.Column And absCol <= rng.Column + rng.Columns.Count - 1 Then
                Set PkgFormat_ResolveCol = Application.Intersect(rng.Parent.Columns(absCol), rng)
            End If
            Exit Function
        End If
    End If

    For i = 1 To rng.Columns.Count
        hdr = Trim(CStr(rng.Cells(1, i).Value))
        If StrComp(hdr, s, vbTextCompare) = 0 Then
            Set PkgFormat_ResolveCol = rng.Columns(i)
            Exit Function
        End If
    Next i
End Function

'================================================================
' ACTION - Board-pack print preset: landscape, fit-to-width, narrow
' margins, repeating header row. Fixed preset, no configuration -
' the whole point is one consistent look every close.
'================================================================
Public Sub PkgFormat_PrintLayout(ctx As Object)
    Dim ws As Worksheet
    Const TITLE As String = "Prep for Print"
    On Error GoTo Fail

    Set ws = ActiveSheet
    If Core_IsEngineSheet(ws.name) Then
        Core_Refuse ctx, "'" & ws.name & "' belongs to the macro toolkit itself, not to a " & _
               "report." & vbCrLf & vbCrLf & "Switch to the report tab and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        ' Narrow Resume Next to this ONE statement (some sheet states reject
        ' PrintTitleRows) and restore the Sub's own Fail: routing right
        ' after - never On Error GoTo 0, which would leave every statement
        ' below unhandled inside a procedure that has a Fail: label.
        On Error Resume Next
        .PrintTitleRows = "$1:$1"
        On Error GoTo Fail
    End With
    Core_SpeedOff

    Core_Done ctx, "'" & ws.name & "' is set to landscape, fit-to-width, narrow margins, " & _
           "and a repeating header row (row 1)." & vbCrLf & vbCrLf & _
           "No cell was changed - only how this sheet prints.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "format.print_layout", Err
End Sub
