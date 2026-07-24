Attribute VB_Name = "MacrosBundle"
Option Explicit

'============================================================================
'  MacrosBundle - Core + enabled packages, concatenated into ONE module.
'  GENERATED ARTIFACT - do not edit. Source of truth:
'    src/core/Core.bas | src/packages/PkgImportExport.bas | src/packages/PkgClean.bas
'  Import this ONE .bas, then run Core_BuildControlPanel (Alt+F8).
'============================================================================

'########################## src/core/Core.bas ##########################

'================================================================
' Core - the required engine every package depends on.
'
' Extracted from SetupImportExportMacros (src/legacy/SetupImportExportMacros.bas):
' the config engine, Setup renderer, control-panel/button builder and
' shared helpers are lifted here and generalized so a package no longer
' hardcodes any of them. A package is now a thin module that:
'   1. describes itself in a manifest  (Pkg<Name>_Manifest reg)
'   2. implements small action subs     (Pkg<Name>_<Verb> ctx)
' Core does everything else - Setup sheet, buttons, config, speed, errors.
'
' Contract (never break these):
'   * Everything Public is Core_*.  Everything else is Private.
'   * Discovery is zero-trust: Core probes "Pkg<Name>_Manifest" with
'     Application.Run, never enumerates modules (no "Trust access to the
'     VBA project object model").
'   * Two objects flow through the system, both Scripting.Dictionary:
'       reg  - a package's manifest (id/name/version + fields + actions)
'       ctx  - per-action context (actionId + its package's reg + log)
'
' Object model
'   reg("id"), reg("name"), reg("version")
'   reg("fields")  = Collection of field dictionaries, in render order
'   reg("actions") = Collection of action dictionaries
'   field: key, label, kind, default, headers, defaults(Collection)
'          kind = "section" | "text" | "int" | "sheet" | "file" |
'                 "bool" | "list" | "table"
'   action: id, caption, entry, help
'   ctx("actionId"), ctx("reg"), ctx("log")
'================================================================

' ---- sheet names ----
Private Const SETUP_SHEET    As String = "Setup"
Private Const CONFIG_SHEET   As String = "_Config"     ' hidden, written by Finalize
Private Const MACRO_SHEET    As String = "Macros"
Private Const PACKAGES_SHEET As String = "_Packages"   ' optional enable/disable layer

' ---- Setup / _Config layout ----
Private Const KEY_COL   As Long = 1       ' column A: machine keys (hidden)
Private Const VAL_COL   As Long = 3       ' column C: scalar value / list value / 1st table col
Private Const MAX_LIST  As Long = 20      ' input rows drawn per list / table block
Private Const END_KEY   As String = "_core_end"   ' sentinel key row: bounds the last list/table block

' ---- speed wrapper saved state (module-level, depth-counted) ----
Private mSpeedDepth  As Long
Private mSpeedActive As Boolean
Private mSaveScreen  As Boolean
Private mSaveEvents  As Boolean
Private mSaveAlerts  As Boolean
Private mSaveCalc    As XlCalculation

'================================================================
' 1. MANIFEST-BUILDING API  (called inside Pkg<Name>_Manifest reg)
'================================================================
Public Sub Core_DeclarePackage(reg As Object, ByVal id As String, _
                               ByVal name As String, ByVal version As String)
    reg.RemoveAll
    reg.Add "id", id
    reg.Add "name", name
    reg.Add "version", version
    reg.Add "fields", New Collection
    reg.Add "actions", New Collection
End Sub

Public Sub Core_AddSection(reg As Object, ByVal title As String)
    reg("fields").Add NewField("", title, "section", "")
End Sub

Public Sub Core_AddField(reg As Object, ByVal key As String, ByVal label As String, _
                         ByVal kind As String, ByVal dflt As String)
    reg("fields").Add NewField(key, label, kind, dflt)
End Sub

Public Sub Core_AddList(reg As Object, ByVal key As String, _
                        ByVal label As String, ByVal colHeader As String)
    Dim f As Object
    Set f = NewField(key, label, "list", "")
    f("headers") = colHeader
    reg("fields").Add f
End Sub

' colHeaders is a "|"-delimited string (NOT comma - a header may itself contain
' commas, e.g. "Type map (T,N,D,...)"). Up to 3 columns are read (C, D, E).
Public Sub Core_AddTable(reg As Object, ByVal key As String, _
                         ByVal label As String, ByVal colHeaders As String)
    Dim f As Object
    Set f = NewField(key, label, "table", "")
    f("headers") = colHeaders
    reg("fields").Add f
End Sub

' Seed a fallback row for a list ("value") or table ("a|b|c") field.
' Call once per default row so nothing hits VBA's 25-continuation limit.
Public Sub Core_AddDefault(reg As Object, ByVal key As String, ByVal rowCsv As String)
    Dim f As Object
    Set f = FindField(reg, key)
    If Not f Is Nothing Then f("defaults").Add rowCsv
End Sub

Public Sub Core_AddAction(reg As Object, ByVal actionId As String, ByVal caption As String, _
                          ByVal entryPoint As String, ByVal help As String)
    Dim a As Object
    Set a = CreateObject("Scripting.Dictionary")
    a.Add "id", actionId
    a.Add "caption", caption
    a.Add "entry", entryPoint
    a.Add "help", help
    reg("actions").Add a
End Sub

Private Function NewField(ByVal key As String, ByVal label As String, _
                          ByVal kind As String, ByVal dflt As String) As Object
    Dim f As Object
    Set f = CreateObject("Scripting.Dictionary")
    f.Add "key", key
    f.Add "label", label
    f.Add "kind", kind
    f.Add "default", dflt
    f.Add "headers", ""
    f.Add "defaults", New Collection
    Set NewField = f
End Function

Private Function FindField(reg As Object, ByVal key As String) As Object
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If LCase(CStr(f("key"))) = LCase(key) Then
                Set FindField = f
                Exit Function
            End If
        End If
    Next f
End Function

'================================================================
' 2. CONFIG-READ API  (called inside action subs, given ctx)
'    Value comes from the active _Config sheet; if that is absent or
'    the cell is blank, the package's manifest default is returned.
'================================================================
Public Function Core_Cfg(ctx As Object, ByVal key As String) As String
    Core_Cfg = ScalarValue(ctx("reg"), key)
End Function

Public Function Core_CfgLong(ctx As Object, ByVal key As String) As Long
    Dim s As String
    s = Core_Cfg(ctx, key)
    ' A non-blank but non-numeric cell must not silently become 0; fall back to
    ' the field's manifest default (matches the legacy typed-wrapper defaults).
    If Not IsNumeric(s) Then s = FieldDefaultReg(ctx("reg"), key)
    If IsNumeric(s) Then Core_CfgLong = CLng(s)
End Function

Public Function Core_CfgList(ctx As Object, ByVal key As String) As Collection
    Set Core_CfgList = ListValues(ctx("reg"), key)
End Function

Public Function Core_CfgTable(ctx As Object, ByVal key As String, _
                              ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    Core_CfgTable = TableValues(ctx("reg"), key, a, b, c)
End Function

' ---- value resolution: sheet read, else manifest default ----
Private Function ScalarValue(reg As Object, ByVal key As String) As String
    Dim v As String
    v = ReadScalar(CfgSheet(), key)
    If v = "" Then v = FieldDefaultReg(reg, key)
    ScalarValue = v
End Function

Private Function ListValues(reg As Object, ByVal key As String) As Collection
    Dim col As Collection
    Set col = ReadList(CfgSheet(), key)
    If col.Count = 0 Then Set col = DefaultListReg(reg, key)
    Set ListValues = col
End Function

Private Function TableValues(reg As Object, ByVal key As String, _
                             ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    If ReadTable3(CfgSheet(), key, a, b, c) Then
        TableValues = True
        Exit Function
    End If
    TableValues = DefaultTable3Reg(reg, key, a, b, c)
End Function

Private Function FieldDefaultReg(reg As Object, ByVal key As String) As String
    Dim f As Object
    Set f = FindField(reg, key)
    If Not f Is Nothing Then FieldDefaultReg = CStr(f("default"))
End Function

Private Function DefaultListReg(reg As Object, ByVal key As String) As Collection
    Dim f As Object, v As Variant, col As New Collection
    Set DefaultListReg = col
    Set f = FindField(reg, key)
    If f Is Nothing Then Exit Function
    For Each v In f("defaults")
        col.Add CStr(v)
    Next v
End Function

Private Function DefaultTable3Reg(reg As Object, ByVal key As String, _
                                  ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    Dim f As Object, n As Long, i As Long, parts() As String
    Set f = FindField(reg, key)
    If f Is Nothing Then Exit Function
    n = f("defaults").Count
    If n = 0 Then Exit Function
    ReDim a(1 To n)
    ReDim b(1 To n)
    ReDim c(1 To n)
    For i = 1 To n
        parts = Split(CStr(f("defaults")(i)), "|")
        a(i) = PartAt(parts, 0)
        b(i) = PartAt(parts, 1)
        c(i) = PartAt(parts, 2)
    Next i
    DefaultTable3Reg = True
End Function

' ---- raw sheet readers (config source = _Config, or any given sheet) ----
Private Function CfgSheet() As Worksheet
    On Error Resume Next
    Set CfgSheet = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0
End Function

Private Function FindKeyRow(ws As Worksheet, ByVal key As String) As Long
    Dim lastR As Long, r As Long
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, KEY_COL).End(xlUp).Row
    For r = 1 To lastR
        If LCase(Trim(CStr(ws.Cells(r, KEY_COL).Value))) = LCase(key) Then
            FindKeyRow = r
            Exit Function
        End If
    Next r
End Function

' A list/table block runs from the key row until the next row that
' carries ANY key in column A (prefix-independent, unlike the legacy
' "k_" test), or the end of the used keys.
Private Function NextKeyRow(ws As Worksheet, ByVal afterRow As Long) As Long
    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.Rows.Count, KEY_COL).End(xlUp).Row
    For r = afterRow + 1 To lastR
        If Trim(CStr(ws.Cells(r, KEY_COL).Value)) <> "" Then
            NextKeyRow = r
            Exit Function
        End If
    Next r
    NextKeyRow = lastR + 1
End Function

Private Function ReadScalar(ws As Worksheet, ByVal key As String) As String
    Dim r As Long
    If ws Is Nothing Then Exit Function
    r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    ReadScalar = Trim(CStr(ws.Cells(r, VAL_COL).Value))
End Function

Private Function ReadList(ws As Worksheet, ByVal key As String) As Collection
    Dim r As Long, nextK As Long, i As Long, v As String, col As New Collection
    Set ReadList = col
    If ws Is Nothing Then Exit Function
    r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    nextK = NextKeyRow(ws, r)
    For i = r + 1 To nextK - 1
        v = Trim(CStr(ws.Cells(i, VAL_COL).Value))
        If v <> "" Then col.Add v
    Next i
End Function

Private Function ReadTable3(ws As Worksheet, ByVal key As String, _
                            ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    Dim r As Long, nextK As Long, i As Long, n As Long, k As Long
    If ws Is Nothing Then Exit Function
    r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    nextK = NextKeyRow(ws, r)
    For i = r + 1 To nextK - 1
        If Trim(CStr(ws.Cells(i, VAL_COL).Value)) <> "" Then n = n + 1
    Next i
    If n = 0 Then Exit Function
    ReDim a(1 To n)
    ReDim b(1 To n)
    ReDim c(1 To n)
    For i = r + 1 To nextK - 1
        If Trim(CStr(ws.Cells(i, VAL_COL).Value)) <> "" Then
            k = k + 1
            a(k) = Trim(CStr(ws.Cells(i, VAL_COL).Value))
            b(k) = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
            c(k) = Trim(CStr(ws.Cells(i, VAL_COL + 2).Value))
        End If
    Next i
    ReadTable3 = True
End Function

'================================================================
' 3. SHARED HELPERS  (Public so packages reuse them; single home here)
'================================================================

' ---- speed wrapper: disable, do work, guaranteed restore ----
Public Sub Core_SpeedOn()
    If mSpeedDepth = 0 Then
        mSaveScreen = Application.ScreenUpdating
        mSaveEvents = Application.EnableEvents
        mSaveAlerts = Application.DisplayAlerts
        mSaveCalc = Application.Calculation
        mSpeedActive = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    End If
    mSpeedDepth = mSpeedDepth + 1
End Sub

Public Sub Core_SpeedOff()
    If mSpeedDepth > 0 Then mSpeedDepth = mSpeedDepth - 1
    If mSpeedDepth <= 0 Then RestoreSpeed
End Sub

' Restore app state. Uses the saved values when Core_SpeedOn ran; otherwise
' forces sane normals so a failure can never leave Excel frozen.
Private Sub RestoreSpeed()
    On Error Resume Next
    If mSpeedActive Then
        Application.Calculation = mSaveCalc
        Application.DisplayAlerts = mSaveAlerts
        Application.EnableEvents = mSaveEvents
        Application.ScreenUpdating = mSaveScreen
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    mSpeedDepth = 0
    mSpeedActive = False
    On Error GoTo 0
End Sub

' ---- structured error + log ----
Public Sub Core_Log(ctx As Object, ByVal msg As String)
    ctx("log") = ctx("log") & msg & vbCrLf
End Sub

Public Sub Core_Fail(ctx As Object, ByVal actionId As String, e As ErrObject)
    Dim m As String
    m = "#" & e.Number & "  " & e.Description
    RestoreSpeed
    MsgBox "Action '" & actionId & "' failed:" & vbCrLf & vbCrLf & m, _
           vbExclamation, "Macro Error"
End Sub

' ---- worksheet / range utilities ----
Public Function Core_SheetExists(ByVal nm As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(nm)
    Core_SheetExists = Not s Is Nothing
    On Error GoTo 0
End Function

' Robust last data row: End(xlUp) on a column, widened by the used-range
' last cell (SpecialCells), which catches trailing formatted/blank rows.
Public Function Core_LastRow(ws As Worksheet, Optional ByVal col As Long = 1) As Long
    Dim r As Long, ur As Long
    If ws Is Nothing Then Exit Function
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    On Error Resume Next
    ur = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    On Error GoTo 0
    If ur > r Then r = ur
    Core_LastRow = r
End Function

' Data extent by a SINGLE column only (bare End(xlUp), NOT widened by
' SpecialCells). Use when you want exactly the block one column defines - e.g.
' reading a source export - where a widened last cell would pull in trailing
' formatted-but-blank rows. Core_LastRow is the widened variant for clearing.
Public Function Core_LastRowUp(ws As Worksheet, Optional ByVal col As Long = 1) As Long
    If ws Is Nothing Then Exit Function
    Core_LastRowUp = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Function Core_LastCol(ws As Worksheet, Optional ByVal row As Long = 1) As Long
    If ws Is Nothing Then Exit Function
    Core_LastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
End Function

Public Function Core_FindCol(ByVal hdr As Variant, ByVal name As String) As Long
    Dim k As Long
    For k = LBound(hdr, 2) To UBound(hdr, 2)
        If Trim(LCase(CStr(hdr(1, k)))) = Trim(LCase(name)) Then
            Core_FindCol = k
            Exit Function
        End If
    Next k
End Function

' ---- file finder: newest file in the workbook folder whose base name
'      = prefix (optionally followed by a " (n)" suffix). Excludes this
'      workbook. Collects names first, then stats - calling FileDateTime
'      inside the Dir() loop can corrupt the iterator on network drives.
Public Function Core_FindFile(ByVal prefix As String) As String
    Dim folder As String, f As String, baseName As String, suffix As String
    Dim names() As String, n As Long, i As Long
    Dim bestFile As String, bestTime As Date, tStamp As Date
    If Trim(prefix) = "" Then Exit Function
    folder = ThisWorkbook.Path
    If Right(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator

    ReDim names(1 To 2000)
    n = 0
    f = Dir(folder & "*.xls*")
    Do While Len(f) > 0
        n = n + 1
        names(n) = f
        f = Dir
    Loop

    For i = 1 To n
        f = names(i)
        If LCase(f) <> LCase(ThisWorkbook.name) Then
            baseName = f
            If InStrRev(baseName, ".") > 0 Then baseName = Left(baseName, InStrRev(baseName, ".") - 1)
            If LCase(Left(baseName, Len(prefix))) = LCase(prefix) Then
                suffix = Trim(Mid(baseName, Len(prefix) + 1))
                If suffix = "" Or suffix Like "(*)" Then
                    tStamp = FileDateTime(folder & f)
                    If bestFile = "" Or tStamp > bestTime Then
                        bestFile = folder & f
                        bestTime = tStamp
                    End If
                End If
            End If
        End If
    Next i
    Core_FindFile = bestFile
End Function

Public Function Core_FileName(ByVal p As String) As String
    If InStrRev(p, "\") > 0 Then
        Core_FileName = Mid(p, InStrRev(p, "\") + 1)
    Else
        Core_FileName = p
    End If
End Function

' Coerce a text value from a source export to a target type.
'   T=text  N=number(strips $ , %)  D/DT=date  B=boolean  G=general(auto)
Public Function Core_Coerce(ByVal v As Variant, ByVal code As String) As Variant
    Dim s As String
    If IsEmpty(v) Then Core_Coerce = "": Exit Function
    If VarType(v) = vbString Then
        If v = "" Then Core_Coerce = "": Exit Function
    End If

    Select Case UCase(code)
        Case "T"
            Core_Coerce = CStr(v)
        Case "N"
            If IsNumeric(v) Then
                Core_Coerce = CDbl(v)
            Else
                s = Replace(Replace(Replace(Replace(CStr(v), "$", ""), ",", ""), " ", ""), "%", "")
                If IsNumeric(s) Then Core_Coerce = CDbl(s) Else Core_Coerce = v
            End If
        Case "D", "DT"
            If IsDate(v) Then Core_Coerce = CDate(v) Else Core_Coerce = v
        Case "B"
            Core_Coerce = (UCase(Trim(CStr(v))) = "TRUE" Or Trim(CStr(v)) = "1")
        Case Else                     ' "G" general / auto
            s = Replace(Replace(CStr(v), ",", ""), "$", "")
            If IsNumeric(s) Then
                Core_Coerce = CDbl(s)
            ElseIf IsDate(v) Then
                Core_Coerce = CDate(v)
            Else
                Core_Coerce = v
            End If
    End Select
End Function

'================================================================
' 4. REGISTRY + DISPATCHER
'    Zero-trust discovery: probe a maintained candidate list by name.
'    Add one package (its PascalCase module suffix) per line below as
'    packages are built - e.g.  c.Add "ImportExport"  ->  probes
'    "PkgImportExport_Manifest".
'================================================================
Private Function CoreCandidates() As Collection
    Dim c As New Collection
    c.Add "ImportExport"            ' probes PkgImportExport_Manifest
    c.Add "Clean"                   ' probes PkgClean_Manifest
    Set CoreCandidates = c
End Function

' Run a package's manifest if that module is actually present.
Private Function TryManifest(ByVal pkgName As String, reg As Object) As Boolean
    On Error Resume Next
    Application.Run "Pkg" & pkgName & "_Manifest", reg
    TryManifest = (Err.Number = 0)
    On Error GoTo 0
End Function

' The installed + enabled packages, each as a fully-built reg.
Private Function BuildRegistry() As Collection
    Dim names As Collection, nm As Variant, reg As Object, out As New Collection
    Set names = CoreCandidates()
    For Each nm In names
        Set reg = CreateObject("Scripting.Dictionary")
        If TryManifest(CStr(nm), reg) Then
            If reg.Exists("id") Then
                If PackageEnabled(CStr(reg("id"))) Then out.Add reg
            End If
        End If
    Next nm
    Set BuildRegistry = out
End Function

' A package is active unless the optional _Packages sheet lists it as FALSE.
Private Function PackageEnabled(ByVal id As String) As Boolean
    Dim ws As Worksheet, r As Long, last As Long
    PackageEnabled = True
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PACKAGES_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        If LCase(Trim(CStr(ws.Cells(r, 1).Value))) = LCase(id) Then
            PackageEnabled = IsTrue(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
End Function

' Button dispatcher. Every button's OnAction is  'Core_Run "<actionId>"'
Public Sub Core_Run(ByVal actionId As String)
    Dim reg As Object, ctx As Object, entry As String
    On Error GoTo Fail
    Set reg = FindActionOwner(actionId, entry)
    If reg Is Nothing Then
        MsgBox "Action not installed: " & actionId, vbExclamation, "Core"
        Exit Sub
    End If
    Set ctx = NewCtx(actionId, reg)
    Application.Run entry, ctx
    Exit Sub
Fail:
    ' Safety net: a package sub that used Core_SpeedOn but raised without its
    ' own Core_Fail handler would otherwise leave Excel frozen. Restore + report.
    RestoreSpeed
    MsgBox "Action '" & actionId & "' failed:" & vbCrLf & vbCrLf & _
           "#" & Err.Number & "  " & Err.Description, vbExclamation, "Macro Error"
End Sub

Private Function FindActionOwner(ByVal actionId As String, ByRef entry As String) As Object
    Dim reg As Object, a As Object
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If LCase(CStr(a("id"))) = LCase(actionId) Then
                entry = CStr(a("entry"))
                Set FindActionOwner = reg
                Exit Function
            End If
        Next a
    Next reg
End Function

Private Function NewCtx(ByVal actionId As String, reg As Object) As Object
    Dim ctx As Object
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.Add "actionId", actionId
    ctx.Add "reg", reg
    ctx.Add "log", ""
    Set NewCtx = ctx
End Function

'================================================================
' 5. SETUP RENDERER  (builds the Setup sheet FROM the manifests)
'================================================================
Public Sub Core_BuildSetup()
    Dim ws As Worksheet, reg As Object, f As Object, r As Long, nPkgs As Long
    Dim firstFieldRow As Long
    On Error GoTo Fail

    If Core_SheetExists(SETUP_SHEET) Then
        If MsgBox("A '" & SETUP_SHEET & "' tab already exists. Rebuild it?" & vbCrLf & _
                  "(It will be re-seeded from the current/active settings.)", _
                  vbQuestion + vbYesNo, "Create Setup Tab") = vbNo Then Exit Sub
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(SETUP_SHEET).Delete
        Application.DisplayAlerts = True
    End If

    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
    ws.name = SETUP_SHEET
    ws.Tab.Color = RGB(112, 48, 160)
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Columns("A").ColumnWidth = 9
    ws.Columns("B").ColumnWidth = 50
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 34
    ws.Columns("E").ColumnWidth = 30
    ws.Columns("F").ColumnWidth = 3
    ws.Columns("G").ColumnWidth = 26
    ws.Columns("H").ColumnWidth = 42

    r = 1
    ws.Cells(r, 2).Value = "Toolkit Setup"
    ws.Cells(r, 2).Font.Size = 16
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Color = RGB(112, 48, 160)
    r = r + 1
    ws.Cells(r, 2).Value = "Fill in the YELLOW cells, then click  'Finalize Setup Tab'  on the Macros tab."
    ws.Cells(r, 2).Font.Italic = True
    r = r + 1
    ws.Cells(r, 2).Value = "Do not edit column A (machine keys). Blank rows in a list are ignored."
    ws.Cells(r, 2).Font.Italic = True
    ws.Cells(r, 2).Font.Color = RGB(150, 0, 0)
    r = r + 1
    ws.Cells(r, 7).Value = "Reference: tabs in THIS workbook"
    ws.Cells(r, 7).Font.Bold = True
    ws.Cells(r, 8).Value = "Reference: Excel files in THIS folder"
    ws.Cells(r, 8).Font.Bold = True
    PopulateReferenceLists ws, r + 1
    r = r + 2

    firstFieldRow = r
    For Each reg In BuildRegistry()
        nPkgs = nPkgs + 1
        r = SectionHeader(ws, r, UCase(CStr(reg("name"))))
        For Each f In reg("fields")
            Select Case f("kind")
                Case "section"
                    r = SectionHeader(ws, r, CStr(f("label")))
                Case "list"
                    r = RenderList(ws, r, reg, f)
                Case "table"
                    r = RenderTable(ws, r, reg, f)
                Case Else
                    r = RenderScalar(ws, r, reg, f)
            End Select
        Next f
        r = r + 1
    Next reg

    ' Sentinel key in column A. The config readers bound a list/table block by
    ' scanning column A for the next non-blank cell; without a trailing key the
    ' LAST list/table on the sheet has no lower bound (End(xlUp) stops at its own
    ' key row, whose value rows are blank in column A) and its values are dropped.
    ws.Cells(r, KEY_COL).Value = END_KEY

    ' Long labels in column B are set to wrap (below); auto-size the field rows
    ' so a wrapped label is fully visible instead of being clipped by the value
    ' cell in column C.
    ws.Rows(firstFieldRow & ":" & r).AutoFit

    ws.Columns("A").Font.Color = RGB(190, 190, 190)
    ws.Columns("A").Hidden = True
    ws.Activate
    ws.Range("B1").Select
    Application.ScreenUpdating = True

    If nPkgs = 0 Then
        MsgBox "'" & SETUP_SHEET & "' created, but no packages are installed yet." & vbCrLf & _
               "Import a Pkg* module and add it to Core's candidate list.", _
               vbInformation, "Create Setup Tab"
    Else
        MsgBox "'" & SETUP_SHEET & "' tab created and pre-filled." & vbCrLf & vbCrLf & _
               "Edit the yellow cells, then run  'Finalize Setup Tab'.", _
               vbInformation, "Create Setup Tab"
    End If
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Core_BuildSetup error: " & Err.Description, vbExclamation, "Create Setup Tab"
End Sub

Private Function RenderScalar(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    ws.Cells(r, KEY_COL).Value = f("key")
    LabelCell ws.Cells(r, 2), CStr(f("label"))
    ws.Cells(r, VAL_COL).Value = ScalarValue(reg, CStr(f("key")))
    Yellow ws.Cells(r, VAL_COL)
    RenderScalar = r + 1
End Function

' Write a field label into column B, wrapped + top-aligned, so a long label
' stays fully visible next to the occupied value cell in column C.
Private Sub LabelCell(rng As Range, ByVal txt As String)
    rng.Value = txt
    rng.WrapText = True
    rng.VerticalAlignment = xlTop
End Sub

Private Function RenderList(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    Dim vals As Collection, i As Long
    ws.Cells(r, KEY_COL).Value = f("key")
    LabelCell ws.Cells(r, 2), CStr(f("label"))
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, VAL_COL).Value = f("headers")
    ws.Cells(r, VAL_COL).Font.Bold = True
    r = r + 1
    Set vals = ListValues(reg, CStr(f("key")))
    For i = 1 To MAX_LIST
        If i <= vals.Count Then ws.Cells(r, VAL_COL).Value = vals(i)
        Yellow ws.Cells(r, VAL_COL)
        r = r + 1
    Next i
    RenderList = r
End Function

Private Function RenderTable(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    Dim heads() As String, a() As String, b() As String, c() As String
    Dim i As Long, j As Long, nHead As Long, hasData As Boolean
    heads = Split(CStr(f("headers")), "|")
    nHead = UBound(heads)
    If nHead > 2 Then nHead = 2                 ' 3 columns read (C, D, E)
    ws.Cells(r, KEY_COL).Value = f("key")
    LabelCell ws.Cells(r, 2), CStr(f("label"))
    ws.Cells(r, 2).Font.Bold = True
    For j = 0 To nHead
        ws.Cells(r, VAL_COL + j).Value = Trim(heads(j))
        ws.Cells(r, VAL_COL + j).Font.Bold = True
    Next j
    r = r + 1
    hasData = TableValues(reg, CStr(f("key")), a, b, c)
    For i = 1 To MAX_LIST
        If hasData Then
            If i <= UBound(a) Then
                ws.Cells(r, VAL_COL).Value = a(i)
                ws.Cells(r, VAL_COL + 1).Value = b(i)
                ws.Cells(r, VAL_COL + 2).Value = c(i)
            End If
        End If
        For j = 0 To nHead
            Yellow ws.Cells(r, VAL_COL + j)
        Next j
        r = r + 1
    Next i
    RenderTable = r
End Function

'================================================================
' 6. FINALIZE  (validate the Setup sheet, then activate -> _Config)
'================================================================
Public Sub Core_FinalizeSetup()
    Dim ws As Worksheet, cfg As Worksheet, reg As Object, f As Object
    Dim rpt As String, warn As Long, v As String, msg As String
    On Error GoTo Fail

    If Not Core_SheetExists(SETUP_SHEET) Then
        MsgBox "No '" & SETUP_SHEET & "' tab found. Run 'Create Setup Tab' first.", _
               vbExclamation, "Finalize Setup Tab"
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(SETUP_SHEET)

    ' --- validate scalar fields against this workbook (before activating) ---
    For Each reg In BuildRegistry()
        For Each f In reg("fields")
            v = ReadScalar(ws, CStr(f("key")))
            Select Case f("kind")
                Case "sheet"
                    If v <> "" Then
                        If Not Core_SheetExists(v) Then
                            rpt = rpt & "- '" & f("label") & "': tab '" & v & "' not found." & vbCrLf
                            warn = warn + 1
                        End If
                    End If
                Case "file"
                    If v <> "" Then
                        If Core_FindFile(v) = "" Then
                            rpt = rpt & "- '" & f("label") & "': no file matches prefix '" & v & "'." & vbCrLf
                            warn = warn + 1
                        End If
                    End If
            End Select
        Next f
    Next reg

    If warn = 0 Then
        msg = "All checks passed."
    Else
        msg = warn & " item(s) to note:" & vbCrLf & vbCrLf & rpt
    End If
    If MsgBox(msg & vbCrLf & "Activate this configuration now?", _
              vbQuestion + vbYesNo, "Finalize Setup Tab") = vbNo Then Exit Sub

    ' --- activate: copy Setup -> hidden _Config ---
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(CONFIG_SHEET).Delete
    On Error GoTo Fail
    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set cfg = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    cfg.name = CONFIG_SHEET
    ws.Activate                        ' a very-hidden sheet cannot be the active sheet
    cfg.Visible = xlSheetVeryHidden
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Configuration activated." & vbCrLf & vbCrLf & _
           "The buttons now use your Setup answers instead of the built-in defaults.", _
           vbInformation, "Finalize Setup Tab"
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Core_FinalizeSetup error: " & Err.Description, vbExclamation, "Finalize Setup Tab"
End Sub

'================================================================
' 7. CONTROL PANEL  (builds the Macros tab + wires every button)
'================================================================
Public Sub Core_BuildControlPanel()
    Dim ws As Worksheet, entries As New Collection, e As Variant
    Dim i As Long, leftPos As Double, topPos As Double
    Dim btnW As Double, btnH As Double, gap As Double, helpRow As Long
    On Error GoTo Fail

    Application.ScreenUpdating = False
    If Core_SheetExists(MACRO_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(MACRO_SHEET)
        On Error Resume Next
        ws.Buttons.Delete
        On Error GoTo Fail
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.name = MACRO_SHEET
    End If

    ws.Tab.Color = RGB(31, 78, 121)
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 47      ' wide enough to hold the btnW-point button
    ws.Columns("C").ColumnWidth = 85      ' description column, sits clear of the button

    With ws.Range("B2")
        .Value = "Macro Toolkit  -  Control Panel"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    With ws.Range("B3")
        .Value = "Run the action buttons for daily work. First-time setup: Create Setup Tab -> Finalize Setup Tab."
        .Font.Italic = True
    End With

    PanelEntries entries          ' each item = Array(caption, onAction, help)

    ' Cadence: btnH + gap = 45pt = 3 rows at 15pt, so each description row lines
    ' up with its button. Force the row heights so the alignment holds regardless
    ' of the sheet's default.
    btnW = 220: btnH = 32: gap = 13
    leftPos = ws.Range("B5").Left
    topPos = ws.Range("B5").Top
    ws.Rows("5:" & (5 + entries.Count * 3 + 2)).RowHeight = 15
    helpRow = 5
    i = 0
    For Each e In entries
        i = i + 1
        AddBtn ws, "btn" & i, CStr(e(0)), CStr(e(1)), leftPos, topPos, btnW, btnH
        ws.Cells(helpRow, 3).Value = CStr(e(2))
        ws.Cells(helpRow, 3).Font.Size = 10
        ws.Cells(helpRow, 3).VerticalAlignment = xlTop
        ws.Cells(helpRow, 3).WrapText = True
        topPos = topPos + btnH + gap
        helpRow = helpRow + 3
    Next e

    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True
    MsgBox "Macros tab built with " & entries.Count & " button(s).", vbInformation, "Setup Complete"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Core_BuildControlPanel error: " & Err.Description, vbExclamation, "Setup"
End Sub

' Daily action buttons first (from the manifests), Core's own setup buttons last.
Private Sub PanelEntries(entries As Collection)
    Dim reg As Object, a As Object, onAct As String
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            onAct = "'Core_Run " & Chr(34) & CStr(a("id")) & Chr(34) & "'"
            entries.Add Array(CStr(a("caption")), onAct, CStr(a("help")))
        Next a
    Next reg
    entries.Add Array("Create Setup Tab", "Core_BuildSetup", _
                      "Builds a pre-filled 'Setup' tab from the installed packages. Use it to reuse these macros in another workbook.")
    entries.Add Array("Finalize Setup Tab", "Core_FinalizeSetup", _
                      "Validates your Setup answers and activates them. The action buttons then use your settings instead of the built-in defaults.")
End Sub

'================================================================
' Setup-tab drawing helpers
'================================================================
Private Sub Yellow(rng As Range)
    rng.Interior.Color = RGB(255, 242, 204)
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(200, 200, 200)
End Sub

Private Function SectionHeader(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).Interior.Color = RGB(221, 217, 240)
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Color = RGB(54, 32, 100)
    SectionHeader = r + 1
End Function

Private Sub PopulateReferenceLists(ws As Worksheet, ByVal startRow As Long)
    Dim wsx As Worksheet, gr As Long, hr As Long
    Dim fdr As String, fn As String
    gr = startRow
    For Each wsx In ThisWorkbook.Worksheets
        ws.Cells(gr, 7).Value = wsx.name
        ws.Cells(gr, 7).Font.Color = RGB(120, 120, 120)
        gr = gr + 1
        If gr > startRow + 250 Then Exit For
    Next wsx
    hr = startRow
    fdr = ThisWorkbook.Path
    If Right(fdr, 1) <> Application.PathSeparator Then fdr = fdr & Application.PathSeparator
    fn = Dir(fdr & "*.xls*")
    Do While Len(fn) > 0
        ws.Cells(hr, 8).Value = fn
        ws.Cells(hr, 8).Font.Color = RGB(120, 120, 120)
        hr = hr + 1
        If hr > startRow + 250 Then Exit Do
        fn = Dir
    Loop
End Sub

Private Sub AddBtn(ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal action As String, _
                   ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double)
    Dim btn As Button
    Set btn = ws.Buttons.Add(l, t, w, h)
    btn.name = nm
    btn.Caption = cap
    btn.OnAction = action
    btn.Font.Size = 11
    btn.Font.Bold = True
End Sub

'================================================================
' Small utilities
'================================================================
Private Function PartAt(ByRef parts() As String, ByVal idx As Long) As String
    If idx >= LBound(parts) And idx <= UBound(parts) Then PartAt = Trim(parts(idx))
End Function

Private Function IsTrue(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        IsTrue = v
    Else
        IsTrue = (UCase(Trim(CStr(v))) = "TRUE" Or Trim(CStr(v)) = "1")
    End If
End Function


'#################### src/packages/PkgImportExport.bas ####################

'================================================================
' PkgImportExport - the original Sample Dashboard's four
' buttons (Clear, Import, Update Ref_Master, Export), re-expressed as
' a thin package on Core. No plumbing lives here: config, speed, error
' handling, file finding and type coercion all call Core_*.
'
' Depends on: Core (required). Discovered zero-trust by Core probing
' PkgImportExport_Manifest via Application.Run.
'
' Behavior is identical to the legacy SetupImportExportMacros module, with ONE
' intentional change: clicking "Import" no longer pops a separate
' "Cleared" dialog before "Import complete" - the clear is now a silent
' internal step (shared ClearTabs helper). The standalone "Clear Input
' Tabs" button still shows its own confirmation.
'================================================================

'================================================================
' MANIFEST - Core reads this to learn the package's fields + actions.
'================================================================
Public Sub PkgImportExport_Manifest(reg As Object)
    Core_DeclarePackage reg, "importexport", "Import / Export", "1.0.0"

    ' ---- GENERAL ----
    Core_AddSection reg, "GENERAL"
    Core_AddField reg, "k_settings_sheet", "Settings / CF sheet name (kept hidden inside the export)", "sheet", "README"
    Core_AddField reg, "k_export_base", "Export file BASE name (the date is added in front)", "text", "My Report"
    Core_AddField reg, "k_export_offset", "Export date offset in days  (-1 = yesterday, 0 = today)", "int", "-1"

    ' ---- EXPORT TABS ----
    Core_AddSection reg, "EXPORT TABS  (copied into the export, in this order)"
    Core_AddList reg, "k_export_tabs", "List each tab name to export, one per row:", "Tab name  v"
    Core_AddDefault reg, "k_export_tabs", "Sample Recap (Daily)"
    Core_AddDefault reg, "k_export_tabs", "Sample Category"
    Core_AddDefault reg, "k_export_tabs", "Sample Detail"
    Core_AddDefault reg, "k_export_tabs", "Inv Value Outliers"
    Core_AddDefault reg, "k_export_tabs", "Floater Finder"

    ' ---- IMPORT FILES ----
    Core_AddSection reg, "IMPORT FILES"
    Core_AddTable reg, "k_import", "For each input tab: its name, the source-file name prefix, and the column type map.", "Input tab  v|Source file name prefix|Type map (T,N,D,DT,B,G)"
    Core_AddDefault reg, "k_import", "IN_Day|Target Inventory Data (day)|T,T,T,T,T,T,T,T,N,N,N,N,N,N,N"
    Core_AddDefault reg, "k_import", "IN_PTD|Target Inventory Data (period to date)|T,T,T,T,T,T,T,T,N,N,N,N,N,N,N"
    Core_AddDefault reg, "k_import", "IN_InvValue|Inventory Value Data|T,T,T,T,T,T,T,N,N,N,N"
    Core_AddDefault reg, "k_import", "IN_Clock|Clock in and out Combine|T,T,T,T,T,T,T,T,T,N,N,N,N,N"

    ' ---- REF_MASTER ----
    Core_AddSection reg, "REF_MASTER"
    Core_AddField reg, "k_master_prefix", "Master-list file name prefix", "file", "Master Store List"
    Core_AddField reg, "k_master_sheet", "Source sheet name inside that file", "text", "Current Master"
    Core_AddField reg, "k_master_dest", "Destination tab in THIS workbook", "sheet", "Ref_Master"
    Core_AddList reg, "k_master_cols", "Source column headers to copy, in order (1st = Store, kept as text):", "Header  v"
    Core_AddDefault reg, "k_master_cols", "Store #"
    Core_AddDefault reg, "k_master_cols", "Location Name"
    Core_AddDefault reg, "k_master_cols", "Entity"
    Core_AddDefault reg, "k_master_cols", "DO Name"
    Core_AddDefault reg, "k_master_cols", "GM Name"

    ' ---- ACTIONS (become buttons) ----
    Core_AddAction reg, "importexport.clear", "1.  Clear Input Tabs", "PkgImportExport_Clear", "Deletes data (row 2 down) on the input tabs. Keeps headers + formatting."
    Core_AddAction reg, "importexport.import", "2.  Import Source Data", "PkgImportExport_Import", "Clears the input tabs, then imports the matching source files from this folder (auto-handles (1)/(2) name suffixes)."
    Core_AddAction reg, "importexport.updateMaster", "3.  Update Ref_Master", "PkgImportExport_UpdateMaster", "Refreshes the Ref_Master tab from the master-list file in this folder."
    Core_AddAction reg, "importexport.export", "4.  Export Dashboard (.xlsx)", "PkgImportExport_Export", "Saves a values-only copy of the main tabs (formatting + filters + conditional formatting kept)."
End Sub

'================================================================
' ACTION 1 - Clear input tab data (keep headers + formatting)
'================================================================
Public Sub PkgImportExport_Clear(ctx As Object)
    Dim tabNames() As String, prefixes() As String, typeMaps() As String
    Dim tlist As String
    On Error GoTo Fail

    If Not Core_CfgTable(ctx, "k_import", tabNames, prefixes, typeMaps) Then
        MsgBox "No input tabs are configured.", vbExclamation, "Clear Input Tabs"
        Exit Sub
    End If

    Core_SpeedOn
    tlist = ClearTabs(tabNames)
    Core_SpeedOff

    MsgBox "Cleared data (row 2 and below) on:" & vbCrLf & _
           "  " & tlist & vbCrLf & vbCrLf & _
           "Headers and formatting were kept.", vbInformation, "Clear Input Tabs"
    Exit Sub
Fail:
    Core_Fail ctx, "importexport.clear", Err
End Sub

'================================================================
' ACTION 2 - Clear, then import each source file into its tab
'================================================================
Public Sub PkgImportExport_Import(ctx As Object)
    Dim tabNames() As String, prefixes() As String, typeMaps() As String
    Dim i As Long, ws As Worksheet, srcPath As String
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim arr As Variant, codes() As String, code As String
    Dim r As Long, cc As Long, sLastRow As Long, sLastCol As Long
    Dim logTxt As String
    Dim sh As Worksheet, pt As PivotTable
    On Error GoTo Fail

    If Not Core_CfgTable(ctx, "k_import", tabNames, prefixes, typeMaps) Then
        MsgBox "No import files are configured.", vbExclamation, "Import Source Data"
        Exit Sub
    End If

    Core_SpeedOn
    Application.AskToUpdateLinks = False

    ClearTabs tabNames                     ' silent internal clear (no dialog)

    For i = LBound(tabNames) To UBound(tabNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(tabNames(i))
        On Error GoTo Fail
        If ws Is Nothing Then
            logTxt = logTxt & "X   " & tabNames(i) & "   <- TAB NOT FOUND in this workbook" & vbCrLf
            GoTo NextTab
        End If
        srcPath = Core_FindFile(prefixes(i))
        If srcPath = "" Then
            logTxt = logTxt & "X   " & tabNames(i) & "   <- NO FILE matching """ & prefixes(i) & "*""" & vbCrLf
        Else
            Set wbSrc = Application.Workbooks.Open(fileName:=srcPath, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, CorruptLoad:=xlRepairFile)
            Set wsSrc = wbSrc.Worksheets(1)
            sLastRow = Core_LastRowUp(wsSrc, 1)   ' bare End(xlUp): match legacy source read
            sLastCol = Core_LastCol(wsSrc, 1)
            codes = Split(typeMaps(i), ",")
            If sLastRow >= 2 Then
                arr = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(sLastRow, sLastCol)).Value
                For cc = 1 To sLastCol
                    If (cc - 1) <= UBound(codes) Then code = Trim(codes(cc - 1)) Else code = "G"
                    ' Text columns must be formatted as Text BEFORE writing, else Excel
                    ' turns numeric-looking strings (store "12345") back into numbers and
                    ' breaks the dashboard's type-strict MATCH lookups.
                    If code = "T" Then _
                        ws.Range(ws.Cells(2, cc), ws.Cells(sLastRow, cc)).NumberFormat = "@"
                    For r = 1 To UBound(arr, 1)
                        arr(r, cc) = Core_Coerce(arr(r, cc), code)
                    Next r
                Next cc
                ws.Range(ws.Cells(2, 1), ws.Cells(sLastRow, sLastCol)).Value = arr
                logTxt = logTxt & "OK  " & tabNames(i) & "   <- " & Core_FileName(srcPath) & _
                         "   (" & (sLastRow - 1) & " rows)" & vbCrLf
            Else
                logTxt = logTxt & "!   " & tabNames(i) & "   <- " & Core_FileName(srcPath) & "   (no data rows)" & vbCrLf
            End If
            wbSrc.Close SaveChanges:=False
            Set wbSrc = Nothing
        End If
NextTab:
    Next i

    ' Refresh pivots and force a full recalc so downstream tabs update.
    Application.Calculation = xlCalculationAutomatic
    For Each sh In ThisWorkbook.Worksheets
        For Each pt In sh.PivotTables
            On Error Resume Next
            pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
            pt.RefreshTable
            On Error GoTo Fail
        Next pt
    Next sh
    Application.CalculateFull

    Application.AskToUpdateLinks = True
    Core_SpeedOff
    MsgBox "Import complete (input tabs were cleared first):" & vbCrLf & vbCrLf & logTxt, _
           vbInformation, "Import Source Data"
    Exit Sub
Fail:
    Core_Fail ctx, "importexport.import", Err
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.AskToUpdateLinks = True
End Sub

'================================================================
' ACTION 3 - Update Ref_Master from the Master Store List
'================================================================
Public Sub PkgImportExport_UpdateMaster(ctx As Object)
    Dim srcPath As String, wbSrc As Workbook, wsSrc As Worksheet, wsDst As Worksheet
    Dim hdr As Variant, lastRow As Long, lastCol As Long, dLast As Long
    Dim j As Long, outR As Long, n As Long, k As Long
    Dim cols As Collection, nCols As Long, srcCol() As Long, missing As String
    Dim mPrefix As String, mSheet As String, mDest As String
    On Error GoTo Fail

    mPrefix = Core_Cfg(ctx, "k_master_prefix")
    mSheet = Core_Cfg(ctx, "k_master_sheet")
    mDest = Core_Cfg(ctx, "k_master_dest")

    srcPath = Core_FindFile(mPrefix)
    If srcPath = "" Then
        MsgBox "No file matching """ & mPrefix & "*"" was found in this folder.", _
               vbExclamation, "Update Ref_Master"
        Exit Sub
    End If
    If Not Core_SheetExists(mDest) Then
        MsgBox "Destination tab '" & mDest & "' was not found in this workbook.", _
               vbExclamation, "Update Ref_Master"
        Exit Sub
    End If
    Set wsDst = ThisWorkbook.Worksheets(mDest)
    Set cols = Core_CfgList(ctx, "k_master_cols")
    nCols = cols.Count
    If nCols = 0 Then
        MsgBox "No Ref_Master columns are configured.", vbExclamation, "Update Ref_Master"
        Exit Sub
    End If
    ReDim srcCol(1 To nCols)

    Core_SpeedOn
    Application.AskToUpdateLinks = False

    Set wbSrc = Application.Workbooks.Open(fileName:=srcPath, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, CorruptLoad:=xlRepairFile)
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(mSheet)
    On Error GoTo Fail
    If wsSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
        Application.AskToUpdateLinks = True
        Core_SpeedOff
        MsgBox "Sheet '" & mSheet & "' not found in the master file.", _
               vbExclamation, "Update Ref_Master"
        Exit Sub
    End If

    lastCol = Core_LastCol(wsSrc, 1)
    hdr = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Value

    For k = 1 To nCols
        srcCol(k) = Core_FindCol(hdr, CStr(cols(k)))
        If srcCol(k) = 0 Then missing = missing & vbCrLf & "   - " & cols(k)
    Next k
    If Len(missing) > 0 Then
        wbSrc.Close SaveChanges:=False
        Application.AskToUpdateLinks = True
        Core_SpeedOff
        MsgBox "Could not find these column header(s) on sheet '" & mSheet & "':" & missing, _
               vbExclamation, "Update Ref_Master"
        Exit Sub
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, srcCol(1)).End(xlUp).Row

    dLast = Core_LastRow(wsDst, 1)
    If dLast >= 2 Then wsDst.Range(wsDst.Cells(2, 1), wsDst.Cells(dLast, nCols)).ClearContents

    ' Force the Store column to Text so store numbers stay text and keep
    ' matching the dashboard's text-based store keys.
    wsDst.Columns(1).NumberFormat = "@"

    outR = 2
    For j = 2 To lastRow
        If Trim(CStr(wsSrc.Cells(j, srcCol(1)).Value)) <> "" Then
            wsDst.Cells(outR, 1).Value = CStr(wsSrc.Cells(j, srcCol(1)).Value)   ' Store (text)
            For k = 2 To nCols
                wsDst.Cells(outR, k).Value = wsSrc.Cells(j, srcCol(k)).Value
            Next k
            outR = outR + 1
        End If
    Next j
    n = outR - 2

    wbSrc.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.AskToUpdateLinks = True
    Core_SpeedOff
    MsgBox mDest & " updated from """ & Core_FileName(srcPath) & """ (sheet '" & mSheet & "')." & vbCrLf & _
           n & " rows written.", vbInformation, "Update Ref_Master"
    Exit Sub
Fail:
    Core_Fail ctx, "importexport.updateMaster", Err
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.AskToUpdateLinks = True
End Sub

'================================================================
' ACTION 4 - Export the main tabs as a values-only .xlsx
'================================================================
Public Sub PkgImportExport_Export(ctx As Object)
    Dim mainTabs As Variant, t As Variant
    Dim newWb As Workbook, ws As Worksheet
    Dim fRange As Range, ar As Range
    Dim keepD As Object, sh As Worksheet, i As Long
    Dim lnks As Variant, li As Long
    Dim outPath As String, fileName As String
    Dim setSh As String, expTabs As Collection
    Dim cel As Range, sp As Range, sv As Variant
    On Error GoTo Fail

    Set expTabs = Core_CfgList(ctx, "k_export_tabs")
    If expTabs.Count = 0 Then
        MsgBox "No export tabs are configured.", vbExclamation, "Export Dashboard"
        Exit Sub
    End If
    mainTabs = CollectionToArray(expTabs)
    setSh = Core_Cfg(ctx, "k_settings_sheet")

    ' every visible main tab must exist
    For Each t In mainTabs
        If Not Core_SheetExists(CStr(t)) Then
            MsgBox "Missing tab: " & t & vbCrLf & "Export cancelled.", vbExclamation, "Export Dashboard"
            Exit Sub
        End If
    Next t
    If Not Core_SheetExists(setSh) Then
        MsgBox "Settings sheet '" & setSh & "' not found. It must stay in" & vbCrLf & _
               "the export (hidden) for the conditional-formatting thresholds.", _
               vbExclamation, "Export Dashboard"
        Exit Sub
    End If

    Core_SpeedOn
    ' Export's spill/legacy-array flattening was written and tested under
    ' AUTOMATIC calculation (see the ordering fixes below). Core_SpeedOn sets
    ' Manual, so restore Automatic here to keep this delicate path identical to
    ' the legacy behavior. ScreenUpdating/Events/Alerts stay off from SpeedOn.
    Application.Calculation = xlCalculationAutomatic
    Application.PrintCommunication = False
    Application.CalculateFull

    ' Copy EVERY worksheet to a new workbook so all references resolve
    ' internally (no blanks when we freeze). Copying sheets does not bring
    ' the VBA project, so the new book has no macros and no external links.
    ThisWorkbook.Worksheets.Copy
    Set newWb = Application.ActiveWorkbook
    newWb.Application.CalculateFull

    ' Flatten dynamic-array SPILLS and legacy CSE ARRAYS to static values
    ' first, then freeze ordinary formulas.
    For Each ws In newWb.Worksheets
        Set fRange = Nothing
        On Error Resume Next
        Set fRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo Fail
        If Not fRange Is Nothing Then
            For Each cel In fRange.Cells
                If cel.HasSpill Then
                    Set sp = cel.SpillingToRange
                    sv = sp.Value
                    cel.ClearContents
                    sp.Value = sv
                ElseIf cel.HasArray Then
                    Set sp = cel.CurrentArray
                    sv = sp.Value
                    sp.ClearContents
                    sp.Value = sv
                End If
            Next cel
        End If
        Set fRange = Nothing
        On Error Resume Next
        Set fRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo Fail
        If Not fRange Is Nothing Then
            For Each ar In fRange.Areas
                On Error Resume Next
                ar.Value = ar.Value
                On Error GoTo Fail
            Next ar
        End If
    Next ws

    ' Delete every sheet EXCEPT the main tabs + the settings sheet.
    Set keepD = CreateObject("Scripting.Dictionary")
    For Each t In mainTabs
        keepD(LCase(CStr(t))) = 1
    Next t
    keepD(LCase(setSh)) = 1
    For i = newWb.Worksheets.Count To 1 Step -1
        Set sh = newWb.Worksheets(i)
        If Not keepD.Exists(LCase(sh.name)) Then
            sh.Visible = xlSheetVisible       ' a hidden sheet can't always be deleted
            sh.Delete
        End If
    Next i
    ' Keep the settings sheet but hidden (its cells still feed the CF rules).
    newWb.Worksheets(setSh).Visible = xlSheetVeryHidden
    ' Restore the main-tab order (deletes can shuffle it).
    For i = UBound(mainTabs) To LBound(mainTabs) Step -1
        newWb.Worksheets(CStr(mainTabs(i))).Move Before:=newWb.Worksheets(1)
    Next i
    ' Defensive: break any stray external links (there should be none now).
    On Error Resume Next
    lnks = newWb.LinkSources(xlLinkTypeExcelLinks)
    If Not IsEmpty(lnks) Then
        For li = LBound(lnks) To UBound(lnks)
            newWb.BreakLink name:=CStr(lnks(li)), Type:=xlLinkTypeExcelLinks
        Next li
    End If
    On Error GoTo Fail
    ' tidy each visible view
    For Each ws In newWb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End If
    Next ws
    newWb.Worksheets(CStr(mainTabs(LBound(mainTabs)))).Activate

    fileName = Format(Date + Core_CfgLong(ctx, "k_export_offset"), "yyyy-mm-dd") & _
               " " & Core_Cfg(ctx, "k_export_base") & ".xlsx"
    outPath = ThisWorkbook.Path
    If Right(outPath, 1) <> Application.PathSeparator Then outPath = outPath & Application.PathSeparator
    outPath = outPath & fileName
    newWb.SaveAs fileName:=outPath, FileFormat:=51   ' 51 = .xlsx (no macros)
    newWb.Close SaveChanges:=False

    Application.PrintCommunication = True
    Core_SpeedOff
    MsgBox "Exported (all rows, all formatting + conditional formatting kept) to:" & vbCrLf & vbCrLf & _
           outPath, vbInformation, "Export Dashboard"
    Exit Sub
Fail:
    Core_Fail ctx, "importexport.export", Err
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    Application.PrintCommunication = True
End Sub

'================================================================
' Package-private helpers
'================================================================

' Clear row 2 downward on each named tab that exists (no speed management,
' no dialog - the caller wraps this). Returns the comma list of tab names.
Private Function ClearTabs(ByRef tabNames() As String) As String
    Dim i As Long, ws As Worksheet, lastRow As Long, lastCol As Long, tlist As String
    For i = LBound(tabNames) To UBound(tabNames)
        tlist = tlist & IIf(tlist = "", "", ", ") & tabNames(i)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(tabNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            lastCol = Core_LastCol(ws, 1)
            lastRow = Core_LastRow(ws, 1)
            If lastRow >= 2 Then
                ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
            End If
        End If
    Next i
    ClearTabs = tlist
End Function

' Collection of tab names -> 0-based Variant array (LBound/UBound friendly).
Private Function CollectionToArray(col As Collection) As Variant
    Dim arr() As String, i As Long
    If col.Count = 0 Then CollectionToArray = Array(): Exit Function
    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        arr(i - 1) = CStr(col(i))
    Next i
    CollectionToArray = arr
End Function


'#################### src/packages/PkgClean.bas ####################

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

    Core_AddSection reg, "CLEAN & STANDARDIZE"
    Core_AddField reg, "clean_keycols", "Remove-duplicates key column(s), e.g. 1 or 1,2,4 (numbered from the left of the range you pick)", "text", "1"

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

    MsgBox "Remove Duplicates complete on " & rng.Address(False, False) & _
           " using key column(s) " & keys & ".", vbInformation, "Remove Duplicates"
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

    Core_SpeedOn
    arr = rng.Value
    If IsArray(arr) Then
        For r = 1 To UBound(arr, 1)
            For c = 1 To UBound(arr, 2)
                If VarType(arr(r, c)) = vbString Then arr(r, c) = CleanText(CStr(arr(r, c)))
            Next c
        Next r
        rng.Value = arr
    ElseIf VarType(arr) = vbString Then              ' single-cell range
        rng.Value = CleanText(CStr(arr))
    End If
    Core_SpeedOff

    MsgBox "Trimmed & cleaned text on " & rng.Address(False, False) & ".", _
           vbInformation, "Trim & Clean Text"
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
                If IsBlankVal(arr(r, c)) Then arr(r, c) = arr(r - 1, c)
            Next r
        Next c
        rng.Value = arr
    End If
    Core_SpeedOff

    MsgBox "Filled blanks down on " & rng.Address(False, False) & ".", _
           vbInformation, "Fill Blanks Down"
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
