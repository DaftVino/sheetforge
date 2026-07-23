Attribute VB_Name = "MacroLayout"
Option Explicit

'================================================================
' MacroLayout - Grouping / Layout builder  (standalone test module)
'
' PURPOSE
'   Lets you decide how ENABLED macros are presented in the final
'   output. Each enabled macro gets a CODE like "1A":
'       - the NUMBER = which output TAB it lands on   (1, 2, 3...)
'       - the LETTER = which COLUMN on that tab        (A, B, C...)
'   Macros that share a code are grouped together and stacked in
'   the order they appear on the grouping page.
'
'   Example: mark 10 macros "1A" -> all 10 sit in Tab 1, Column A.
'   Mark 10 "2A" and 10 "2B" -> Tab 2 gets Column A (10) + Column B (10).
'
' HOW TO USE (first run: press Alt+F8 -> run "ML_BuildControl")
'   1. "Seed / Reset Master"   - writes the catalog of macros with an
'                                Enabled (TRUE/FALSE) flag you can edit.
'   2. "Build Grouping Page"   - lists ONLY the enabled macros, each with
'                                a yellow CODE cell for you to fill in.
'   3. "Apply Layout"          - reads the codes and builds the "Macros 1",
'                                "Macros 2"... tabs with the buttons placed
'                                by tab number + column letter.
'
' NOTE: the buttons currently call a placeholder (ML_RunById) that just
' reports which macro would run. Swap that for the real Core dispatcher
' (Core_Run) once the engine is wired. Nothing else here needs to change.
'================================================================

' ---- sheet names ----
Private Const CONTROL_SHEET As String = "Layout Control"
Private Const MASTER_SHEET  As String = "Macro Master"
Private Const LAYOUT_SHEET  As String = "Grouping Page"
Private Const PANEL_BASE    As String = "Macros "      ' output tabs: "Macros 1", "Macros 2"...

' ---- master sheet columns ----
Private Const M_ENABLED As Long = 1      ' A
Private Const M_ID      As Long = 2      ' B (hidden)
Private Const M_NAME    As Long = 3      ' C
Private Const M_GROUP   As Long = 4      ' D
Private Const M_DESC    As Long = 5      ' E
Private Const M_FIRST   As Long = 2      ' first data row

' ---- grouping-page columns ----
Private Const G_CODE  As Long = 1        ' A (yellow input)
Private Const G_NAME  As Long = 2        ' B
Private Const G_GROUP As Long = 3        ' C
Private Const G_DESC  As Long = 4        ' D
Private Const G_ID    As Long = 6        ' F (hidden)
Private Const G_FIRST As Long = 6        ' first data row

' ---- button / column layout geometry (points) ----
Private Const BTN_W   As Double = 210
Private Const BTN_H   As Double = 30
Private Const ROW_GAP As Double = 8
Private Const COL_GAP As Double = 24
Private Const HDR_H   As Double = 22

'================================================================
' Build the control panel (run this once via Alt+F8)
'================================================================
Public Sub ML_BuildControl()
    Dim ws As Worksheet, t As Double, l As Double
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set ws = EnsureSheet(CONTROL_SHEET, True)
    On Error Resume Next
    ws.Buttons.Delete
    ws.Shapes.SelectAll: Selection.Delete
    On Error GoTo Fail
    ws.Cells.Clear
    ws.Tab.Color = RGB(31, 78, 121)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 30
    ws.Columns("C").ColumnWidth = 90

    With ws.Range("B2")
        .Value = "Macro Layout - Control Panel"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(31, 78, 121)
    End With
    ws.Range("B3").Value = "Run 1 -> 2 -> 3. Edit Enabled flags on '" & MASTER_SHEET & "', codes on '" & LAYOUT_SHEET & "'."
    ws.Range("B3").Font.Italic = True

    l = ws.Range("B5").Left
    t = ws.Range("B5").Top

    ML_AddBtnAction ws, "1.  Seed / Reset Master", "ML_SeedMaster", l, t, 220, 34
    ws.Range("C5").Value = "Writes the macro catalog with an Enabled (TRUE/FALSE) column. Edit which are on."
    t = t + 34 + 14
    ML_AddBtnAction ws, "2.  Build Grouping Page", "ML_BuildGroupingPage", l, t, 220, 34
    ws.Range("C8").Value = "Lists ONLY enabled macros, each with a yellow CODE cell (e.g. 1A). Keeps codes you already entered."
    t = t + 34 + 14
    ML_AddBtnAction ws, "3.  Apply Layout", "ML_ApplyLayout", l, t, 220, 34
    ws.Range("C11").Value = "Reads the codes and (re)builds the 'Macros 1', 'Macros 2'... tabs with buttons placed by tab # + column letter."

    ws.Range("C5,C8,C11").Font.Size = 10
    ws.Range("C5,C8,C11").WrapText = True
    ws.Range("C5,C8,C11").VerticalAlignment = xlTop

    ws.Activate: ws.Range("A1").Select
    Application.ScreenUpdating = True
    MsgBox "Control panel ready. Use buttons 1 -> 2 -> 3.", vbInformation, "Macro Layout"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "ML_BuildControl error: " & Err.Description, vbExclamation, "Macro Layout"
End Sub

'================================================================
' 1 - Seed / reset the master catalog
'================================================================
Public Sub ML_SeedMaster()
    Dim ws As Worksheet, lines As Collection, i As Long, r As Long
    Dim parts() As String
    On Error GoTo Fail

    If SheetExists(MASTER_SHEET) Then
        If MsgBox("'" & MASTER_SHEET & "' already exists. Reset it to the full catalog?" & vbCrLf & _
                  "(Your Enabled edits will be lost.)", vbQuestion + vbYesNo, "Seed Master") = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Set ws = EnsureSheet(MASTER_SHEET, False)
    ws.Cells.Clear
    ws.Tab.Color = RGB(112, 48, 160)

    ws.Cells(1, M_ENABLED).Value = "Enabled"
    ws.Cells(1, M_ID).Value = "id"
    ws.Cells(1, M_NAME).Value = "Macro"
    ws.Cells(1, M_GROUP).Value = "Group"
    ws.Cells(1, M_DESC).Value = "Description"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, M_DESC)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, M_DESC)).Interior.Color = RGB(221, 217, 240)

    Set lines = SeedLines()
    r = M_FIRST
    For i = 1 To lines.Count
        parts = Split(lines(i), "|")
        ws.Cells(r, M_ENABLED).Value = True
        ws.Cells(r, M_ID).Value = parts(0)
        ws.Cells(r, M_NAME).Value = parts(1)
        ws.Cells(r, M_GROUP).Value = parts(2)
        If UBound(parts) >= 3 Then ws.Cells(r, M_DESC).Value = parts(3)
        r = r + 1
    Next i

    ' TRUE/FALSE dropdown on the Enabled column
    With ws.Range(ws.Cells(M_FIRST, M_ENABLED), ws.Cells(r - 1, M_ENABLED)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="TRUE,FALSE"
    End With

    ws.Columns(M_ENABLED).ColumnWidth = 10
    ws.Columns(M_ID).Hidden = True
    ws.Columns(M_NAME).ColumnWidth = 40
    ws.Columns(M_GROUP).ColumnWidth = 24
    ws.Columns(M_DESC).ColumnWidth = 60
    ws.Cells(M_FIRST, M_ENABLED).Select
    ActiveWindow.FreezePanes = False

    ws.Activate
    Application.ScreenUpdating = True
    MsgBox (r - M_FIRST) & " macros written to '" & MASTER_SHEET & "'." & vbCrLf & _
           "Set Enabled = FALSE for any you don't want, then run '2. Build Grouping Page'.", _
           vbInformation, "Seed Master"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "ML_SeedMaster error: " & Err.Description, vbExclamation, "Seed Master"
End Sub

'================================================================
' 2 - Build the grouping page (only enabled macros)
'================================================================
Public Sub ML_BuildGroupingPage()
    Dim wsM As Worksheet, ws As Worksheet
    Dim r As Long, last As Long, outR As Long, n As Long
    Dim prior As Object, id As String
    On Error GoTo Fail

    If Not SheetExists(MASTER_SHEET) Then
        MsgBox "Run '1. Seed / Reset Master' first.", vbExclamation, "Grouping Page": Exit Sub
    End If
    Set wsM = ThisWorkbook.Worksheets(MASTER_SHEET)

    ' remember codes already entered (so toggling Enabled doesn't lose them)
    Set prior = ReadPriorCodes()

    Application.ScreenUpdating = False
    Set ws = EnsureSheet(LAYOUT_SHEET, False)
    ws.Cells.Clear
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo Fail
    ws.Tab.Color = RGB(0, 128, 96)

    ws.Cells(1, 2).Value = "Grouping Page - arrange the enabled macros"
    ws.Cells(1, 2).Font.Size = 16: ws.Cells(1, 2).Font.Bold = True: ws.Cells(1, 2).Font.Color = RGB(0, 110, 80)
    ws.Cells(2, 2).Value = "Enter a CODE in column A for each macro:  NUMBER = tab,  LETTER = column.  e.g. 1A, 2B."
    ws.Cells(2, 2).Font.Italic = True
    ws.Cells(3, 2).Value = "Same code = grouped together, stacked top-to-bottom in this list order. Blank = not placed. Then run '3. Apply Layout'."
    ws.Cells(3, 2).Font.Italic = True: ws.Cells(3, 2).Font.Color = RGB(150, 0, 0)

    ws.Cells(5, G_CODE).Value = "Code"
    ws.Cells(5, G_NAME).Value = "Macro"
    ws.Cells(5, G_GROUP).Value = "Group"
    ws.Cells(5, G_DESC).Value = "Description"
    ws.Range(ws.Cells(5, G_CODE), ws.Cells(5, G_DESC)).Font.Bold = True
    ws.Range(ws.Cells(5, G_CODE), ws.Cells(5, G_DESC)).Interior.Color = RGB(221, 240, 232)

    last = wsM.Cells(wsM.Rows.Count, M_ID).End(xlUp).Row
    outR = G_FIRST
    For r = M_FIRST To last
        If IsEnabled(wsM.Cells(r, M_ENABLED).Value) Then
            id = CStr(wsM.Cells(r, M_ID).Value)
            ws.Cells(outR, G_NAME).Value = wsM.Cells(r, M_NAME).Value
            ws.Cells(outR, G_GROUP).Value = wsM.Cells(r, M_GROUP).Value
            ws.Cells(outR, G_DESC).Value = wsM.Cells(r, M_DESC).Value
            ws.Cells(outR, G_ID).Value = id
            If prior.Exists(id) Then ws.Cells(outR, G_CODE).Value = prior(id)
            Yellow ws.Cells(outR, G_CODE)
            outR = outR + 1
        End If
    Next r
    n = outR - G_FIRST

    ws.Columns(G_CODE).ColumnWidth = 10
    ws.Columns(G_NAME).ColumnWidth = 40
    ws.Columns(G_GROUP).ColumnWidth = 24
    ws.Columns(G_DESC).ColumnWidth = 60
    ws.Columns(G_ID).Hidden = True
    ws.Cells(G_FIRST, G_CODE).Select

    ws.Activate
    Application.ScreenUpdating = True
    If n = 0 Then
        MsgBox "No enabled macros found. Set some Enabled = TRUE on '" & MASTER_SHEET & "'.", vbExclamation, "Grouping Page"
    Else
        MsgBox n & " enabled macros listed. Fill the yellow Code cells, then run '3. Apply Layout'.", vbInformation, "Grouping Page"
    End If
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "ML_BuildGroupingPage error: " & Err.Description, vbExclamation, "Grouping Page"
End Sub

'================================================================
' 3 - Apply the layout: build the Macros N tabs from the codes
'================================================================
Public Sub ML_ApplyLayout()
    Dim ws As Worksheet, tabs As Object, cols As Object, items As Collection
    Dim r As Long, last As Long
    Dim code As String, id As String, nm As String
    Dim tabNum As Long, colIndex As Long, colLetter As String
    Dim tabKeys As Variant, colKeys As Variant, ti As Long, ci As Long, ii As Long
    Dim placed As Long, skipped As Long, report As String
    Dim panel As Worksheet, leftBase As Double, topBase As Double, colLeft As Double, topPos As Double
    Dim it As Variant
    On Error GoTo Fail

    If Not SheetExists(LAYOUT_SHEET) Then
        MsgBox "Run '2. Build Grouping Page' first.", vbExclamation, "Apply Layout": Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(LAYOUT_SHEET)

    ' ---- read the grouping page into nested dictionaries ----
    Set tabs = CreateObject("Scripting.Dictionary")   ' tabNum -> (colIndex -> Collection of Array(name,id,letter))
    last = ws.Cells(ws.Rows.Count, G_ID).End(xlUp).Row
    For r = G_FIRST To last
        nm = CStr(ws.Cells(r, G_NAME).Value)
        If Len(Trim(nm)) > 0 Then
            code = CStr(ws.Cells(r, G_CODE).Value)
            id = CStr(ws.Cells(r, G_ID).Value)
            If ParseCode(code, tabNum, colIndex, colLetter) Then
                If Not tabs.Exists(tabNum) Then tabs.Add tabNum, CreateObject("Scripting.Dictionary")
                Set cols = tabs(tabNum)
                If Not cols.Exists(colIndex) Then cols.Add colIndex, New Collection
                cols(colIndex).Add Array(nm, id, colLetter)
                placed = placed + 1
            Else
                skipped = skipped + 1
                report = report & "  - " & nm & "  (code """ & code & """)" & vbCrLf
            End If
        End If
    Next r

    If placed = 0 Then
        MsgBox "No valid codes found. Enter codes like 1A on the grouping page first.", vbExclamation, "Apply Layout"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    DeletePanels                       ' clear any prior Macros N tabs

    ' ---- build each tab ----
    tabKeys = SortedKeys(tabs)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = tabKeys(ti)
        Set panel = EnsureSheet(PANEL_BASE & tabNum, False)
        panel.Cells.Clear
        On Error Resume Next
        panel.Buttons.Delete
        panel.Shapes.SelectAll: Selection.Delete
        On Error GoTo Fail
        panel.Tab.Color = RGB(68, 114, 196)

        leftBase = panel.Range("B4").Left
        topBase = panel.Range("B4").Top
        ML_AddText panel, leftBase, panel.Range("B2").Top, 400, 24, _
                   "Macros - Tab " & tabNum, 15, True, RGB(31, 78, 121)

        Set cols = tabs(tabNum)
        colKeys = SortedKeys(cols)
        For ci = LBound(colKeys) To UBound(colKeys)
            colIndex = colKeys(ci)
            colLeft = leftBase + (colIndex - 1) * (BTN_W + COL_GAP)
            Set items = cols(colIndex)
            ML_AddText panel, colLeft, topBase, BTN_W, HDR_H, _
                       "Column " & IndexToCol(colIndex), 11, True, RGB(90, 90, 90)
            topPos = topBase + HDR_H + ROW_GAP
            For ii = 1 To items.Count
                it = items(ii)
                ML_AddBtnRun panel, CStr(it(0)), CStr(it(1)), colLeft, topPos, BTN_W, BTN_H
                topPos = topPos + BTN_H + ROW_GAP
            Next ii
        Next ci

        panel.Range("A1").Select
    Next ti

    ThisWorkbook.Worksheets(PANEL_BASE & CLng(tabKeys(LBound(tabKeys)))).Activate
    Application.ScreenUpdating = True

    Dim msg As String
    msg = placed & " macro(s) placed across " & tabs.Count & " tab(s)."
    If skipped > 0 Then msg = msg & vbCrLf & vbCrLf & skipped & " skipped (blank/invalid code):" & vbCrLf & report
    MsgBox msg, vbInformation, "Apply Layout"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "ML_ApplyLayout error: " & Err.Description, vbExclamation, "Apply Layout"
End Sub

'================================================================
' Placeholder dispatcher - swap for Core_Run when the engine is wired
'================================================================
Public Sub ML_RunById(ByVal id As String)
    MsgBox "This button is wired to macro:" & vbCrLf & vbCrLf & "   " & id & vbCrLf & vbCrLf & _
           "(Placeholder - connect to the real Core dispatcher later.)", vbInformation, "Run Macro"
End Sub

'================================================================
' Code parsing:  "12AB" -> tabNum 12, column index for "AB"
'================================================================
Private Function ParseCode(ByVal code As String, ByRef tabNum As Long, _
                           ByRef colIndex As Long, ByRef colLetter As String) As Boolean
    Dim s As String, i As Long, ch As String, digits As String, letters As String
    tabNum = 0: colIndex = 0: colLetter = ""
    s = UCase(Trim(code))
    If Len(s) = 0 Then Exit Function
    i = 1
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch Else Exit Do
        i = i + 1
    Loop
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch >= "A" And ch <= "Z" Then letters = letters & ch Else Exit Do
        i = i + 1
    Loop
    If i <= Len(s) Then Exit Function            ' leftover chars -> invalid
    If Len(digits) = 0 Or Len(letters) = 0 Then Exit Function
    tabNum = CLng(digits)
    colLetter = letters
    colIndex = ColToIndex(letters)
    If tabNum < 1 Or colIndex < 1 Then Exit Function
    ParseCode = True
End Function

Private Function ColToIndex(ByVal letters As String) As Long
    Dim i As Long, n As Long
    For i = 1 To Len(letters)
        n = n * 26 + (Asc(Mid(letters, i, 1)) - Asc("A") + 1)
    Next i
    ColToIndex = n
End Function

Private Function IndexToCol(ByVal idx As Long) As String
    Dim n As Long, r As Long, s As String
    n = idx
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(Asc("A") + r) & s
        n = (n - 1) \ 26
    Loop
    IndexToCol = s
End Function

'================================================================
' Helpers
'================================================================
Private Function SortedKeys(d As Object) As Variant
    Dim k As Variant, arr() As Long, n As Long, i As Long, j As Long, t As Long
    n = d.Count
    If n = 0 Then SortedKeys = Array(): Exit Function
    ReDim arr(0 To n - 1)
    i = 0
    For Each k In d.Keys
        arr(i) = CLng(k): i = i + 1
    Next k
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If arr(j) > arr(j + 1) Then
                t = arr(j): arr(j) = arr(j + 1): arr(j + 1) = t
            End If
        Next j
    Next i
    SortedKeys = arr
End Function

Private Function ReadPriorCodes() As Object
    Dim d As Object, ws As Worksheet, r As Long, last As Long, id As String, code As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ReadPriorCodes = d
    If Not SheetExists(LAYOUT_SHEET) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(LAYOUT_SHEET)
    last = ws.Cells(ws.Rows.Count, G_ID).End(xlUp).Row
    For r = G_FIRST To last
        id = CStr(ws.Cells(r, G_ID).Value)
        code = Trim(CStr(ws.Cells(r, G_CODE).Value))
        If Len(id) > 0 And Len(code) > 0 Then
            If Not d.Exists(id) Then d.Add id, code
        End If
    Next r
End Function

Private Function IsEnabled(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        IsEnabled = v
    Else
        IsEnabled = (UCase(Trim(CStr(v))) = "TRUE" Or Trim(CStr(v)) = "1")
    End If
End Function

Private Sub DeletePanels()
    Dim i As Long, nm As String
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        nm = ThisWorkbook.Worksheets(i).Name
        If IsPanelName(nm) Then ThisWorkbook.Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = True
End Sub

Private Function IsPanelName(ByVal nm As String) As Boolean
    Dim tail As String
    If LCase(Left(nm, Len(PANEL_BASE))) <> LCase(PANEL_BASE) Then Exit Function
    tail = Mid(nm, Len(PANEL_BASE) + 1)
    If Len(tail) = 0 Then Exit Function
    IsPanelName = IsNumeric(tail)
End Function

Private Function EnsureSheet(ByVal nm As String, ByVal atFront As Boolean) As Worksheet
    If SheetExists(nm) Then
        Set EnsureSheet = ThisWorkbook.Worksheets(nm)
    ElseIf atFront Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        EnsureSheet.Name = nm
    Else
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = nm
    End If
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(nm)
    SheetExists = Not s Is Nothing
    On Error GoTo 0
End Function

Private Sub Yellow(rng As Range)
    rng.Interior.Color = RGB(255, 242, 204)
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(200, 200, 200)
End Sub

' button that runs a named sub with no argument (control panel)
Private Sub ML_AddBtnAction(ws As Worksheet, cap As String, action As String, _
                            l As Double, t As Double, w As Double, h As Double)
    Dim b As Button
    Set b = ws.Buttons.Add(l, t, w, h)
    b.Caption = cap
    b.OnAction = action
    b.Font.Size = 11: b.Font.Bold = True
End Sub

' button that runs ML_RunById "<id>"  (placed macro)
Private Sub ML_AddBtnRun(ws As Worksheet, cap As String, id As String, _
                         l As Double, t As Double, w As Double, h As Double)
    Dim b As Button
    Set b = ws.Buttons.Add(l, t, w, h)
    b.Caption = cap
    b.OnAction = "'ML_RunById " & Chr(34) & id & Chr(34) & "'"
    b.Font.Size = 10
End Sub

Private Function ML_AddText(ws As Worksheet, l As Double, t As Double, w As Double, h As Double, _
                            txt As String, sz As Single, bold As Boolean, colr As Long) As Shape
    Dim sh As Shape
    Set sh = ws.Shapes.AddTextbox(1, l, t, w, h)   ' 1 = msoTextOrientationHorizontal
    sh.Fill.Visible = msoFalse
    sh.Line.Visible = msoFalse
    With sh.TextFrame2.TextRange
        .Text = txt
        .Font.Size = sz
        .Font.Bold = bold
        .Font.Fill.ForeColor.RGB = colr
    End With
    Set ML_AddText = sh
End Function

'================================================================
' Catalog seed  (id | name | group | description)
' One statement per line (avoids VBA's 25-continuation limit).
'================================================================
Private Function SeedLines() As Collection
    Dim s As New Collection
    s.Add "importexport.import|Import files from folder (typed)|Bring Data In|Match files by prefix, clear target, paste, coerce types"
    s.Add "importexport.combineSheets|Combine all sheets into one|Bring Data In|Stack every tab into one master table"
    s.Add "importexport.combineBooks|Combine all workbooks in folder|Bring Data In|Append every file in a folder into one table"
    s.Add "importexport.csv|Import CSV / TXT (delimited)|Bring Data In|Parse a delimited file with type handling"
    s.Add "importexport.splitcols|Split column into many|Bring Data In|Text-to-columns on a selection"
    s.Add "importexport.closedbook|Pull from a closed workbook|Bring Data In|Read a range without opening the source"
    s.Add "clean.clearrows|Clear data rows, keep headers|Clean & Standardize|Wipe row 2 down, keep headers + formatting"
    s.Add "clean.blankrows|Remove blank rows / columns|Clean & Standardize|Delete empty rows/cols in the used range"
    s.Add "clean.dupes|Remove duplicates|Clean & Standardize|De-dupe on chosen key columns"
    s.Add "clean.trim|Trim & clean text|Clean & Standardize|Strip spaces and non-printing characters"
    s.Add "clean.filldown|Fill blanks down|Clean & Standardize|Fill blank cells with the value above"
    s.Add "clean.unmerge|Unmerge & fill down|Clean & Standardize|Unmerge cells and repeat the value"
    s.Add "clean.formats|Standardize number / date formats|Clean & Standardize|Apply consistent formats to key columns"
    s.Add "clean.case|Change case (UPPER/lower/Proper)|Clean & Standardize|Normalize text casing"
    s.Add "clean.textnum|Text-numbers to real numbers|Clean & Standardize|Strip $ , % and re-type as numeric"
    s.Add "recon.match|Reconcile two datasets on a key|Validate & Reconcile|Flag missing/extra/differing rows"
    s.Add "recon.validate|Data-validation scan|Validate & Reconcile|Sweep a range for blanks/type/range errors"
    s.Add "recon.dupehighlight|Highlight duplicates|Validate & Reconcile|Color repeated keys without deleting"
    s.Add "recon.outliers|Highlight outliers|Validate & Reconcile|Flag values past a threshold"
    s.Add "recon.missing|Flag missing required fields|Validate & Reconcile|Mark rows missing a mandatory column"
    s.Add "enrich.refmaster|Refresh reference master (mapped)|Match & Enrich|Pull named columns from a master file"
    s.Add "enrich.join|Lookup / join in code|Match & Enrich|Fast dictionary join over large sets"
    s.Add "enrich.mapcols|Add mapped columns|Match & Enrich|Enrich rows from a lookup table"
    s.Add "format.cf|Conditional highlighting|Format & Present|Rule-driven coloring from config"
    s.Add "format.headers|Style header row + autofit|Format & Present|Bold/fill header, size every column"
    s.Add "format.title|Add & style a report title|Format & Present|Merge + center a formatted title block"
    s.Add "format.tablestyle|Apply house table style|Format & Present|Banding/borders standard for any table"
    s.Add "format.print|Print-layout prep|Format & Present|Landscape, fit-to-width, tight margins"
    s.Add "sum.pivotrefresh|Refresh all pivot tables|Summarize & Analyze|Loop every sheet and refresh pivots"
    s.Add "sum.recalc|Force full recalculation|Summarize & Analyze|CalculateFull so downstream tabs update"
    s.Add "sum.summary|Build summary / pivot|Summarize & Analyze|Summary table from chosen dims + measures"
    s.Add "sum.pivotstatic|Copy pivot to static values|Summarize & Analyze|Freeze pivot output as values + formats"
    s.Add "sum.charts|Style all charts|Summarize & Analyze|Apply one visual standard to chart objects"
    s.Add "export.snapshot|Export values-only snapshot (.xlsx)|Get It Out|Flatten to values, keep formatting/CF, dated file"
    s.Add "export.tovalues|Convert formulas to values|Get It Out|Freeze formula outputs incl. spills"
    s.Add "export.pdf|Export to PDF|Get It Out|One-click PDF of a range/sheet/workbook"
    s.Add "export.persheet|Save each sheet as own file/PDF|Get It Out|Explode a workbook into per-tab files"
    s.Add "export.email|Email sheet / range (Outlook)|Get It Out|Send current output as an attachment"
    s.Add "book.unhideall|Unhide all worksheets|Workbook Housekeeping|Reveal every hidden tab at once"
    s.Add "book.protectall|Protect / unprotect all sheets|Workbook Housekeeping|One password across the book"
    s.Add "book.toc|Create TOC with links|Workbook Housekeeping|Auto-built index linking every tab"
    s.Add "book.sortsheets|Sort worksheets alphabetically|Workbook Housekeeping|Reorder tabs A-Z"
    s.Add "safe.backup|Timestamped backup copy|Protect & Back Up|SaveCopyAs into a Backup folder"
    s.Add "safe.snapshot|Snapshot before a risky op|Protect & Back Up|Auto restore-point before Clear/Import"
    s.Add "safe.archive|Archive file to dated folder|Protect & Back Up|Move output into Archive\yyyy_mm"
    s.Add "eng.lastrow|Last-row / col finder (util)|Run Faster|Robust End(xlUp) + SpecialCells fallback"
    s.Add "eng.timestamp|Timestamp row on edit|Run Faster|Auto-stamp when a row is entered/changed"
    s.Add "eng.batchfolder|Loop a folder of workbooks|Run Faster|Batch-run actions across many files"
    Set SeedLines = s
End Function
