Attribute VB_Name = "Core"
Option Explicit

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

' ---- grouped control-panel geometry (ported from MacroLayout; Task 9) ----
Private Const PANEL_BASE As String = "Macros "   ' grouped output tabs: "Macros 1", "Macros 2"...
Private Const BTN_W   As Double = 210
Private Const BTN_H   As Double = 30
Private Const ROW_GAP As Double = 8
Private Const COL_GAP As Double = 24
Private Const HDR_H   As Double = 22

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
    reg("fields").Add NewField("", title, "section", "", "simple", "")
End Sub

' tier is "simple" (default) or "advanced". Advanced fields are grouped under an
' "ADVANCED - most users can leave these as-is" band so regular users can stop
' after the simple fields. help is an optional plain-language note shown in grey
' under the field (pass via help:= to skip tier).
Public Sub Core_AddField(reg As Object, ByVal key As String, ByVal label As String, _
                         ByVal kind As String, ByVal dflt As String, _
                         Optional ByVal tier As String = "simple", _
                         Optional ByVal help As String = "")
    reg("fields").Add NewField(key, label, kind, dflt, tier, help)
End Sub

Public Sub Core_AddList(reg As Object, ByVal key As String, _
                        ByVal label As String, ByVal colHeader As String, _
                        Optional ByVal tier As String = "simple", _
                        Optional ByVal help As String = "")
    Dim f As Object
    Set f = NewField(key, label, "list", "", tier, help)
    f("headers") = colHeader
    reg("fields").Add f
End Sub

' colHeaders is a "|"-delimited string (NOT comma - a header may itself contain
' commas, e.g. "Type map (T,N,D,...)"). Up to 3 columns are read (C, D, E).
Public Sub Core_AddTable(reg As Object, ByVal key As String, _
                         ByVal label As String, ByVal colHeaders As String, _
                         Optional ByVal tier As String = "simple", _
                         Optional ByVal help As String = "")
    Dim f As Object
    Set f = NewField(key, label, "table", "", tier, help)
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
                          ByVal kind As String, ByVal dflt As String, _
                          ByVal tier As String, ByVal help As String) As Object
    Dim f As Object
    Set f = CreateObject("Scripting.Dictionary")
    f.Add "key", key
    f.Add "label", label
    f.Add "kind", kind
    f.Add "default", dflt
    f.Add "headers", ""
    f.Add "defaults", New Collection
    f.Add "help", help
    ' Normalize the tier so an unknown or mis-cased value degrades to "simple"
    ' (still shown) rather than matching neither render pass and vanishing.
    Dim t As String
    t = LCase(Trim(tier))
    If t <> "advanced" Then t = "simple"
    f.Add "tier", t
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

' ---- setup-helper API: read/write the LIVE Setup sheet before Finalize ----
' For "setup helper" actions (e.g. suggest column types) that operate on what
' the user is editing on the Setup tab, not the activated _Config. These encode
' the Setup layout so a package never hardcodes column positions.

' Read a 3-column table block (tab/prefix/type map style) from the Setup sheet.
Public Function Core_SetupTable(ByVal key As String, _
                                ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    Core_SetupTable = ReadTable3(SetupSheet(), key, a, b, c)
End Function

' Write one column of a block on the Setup sheet: colOffset 0=C, 1=D, 2=E.
' Values are applied, in order, to the block's populated rows (those with a
' non-blank first column) - the same rows Core_SetupTable returns.
Public Sub Core_SetupSetColumn(ByVal key As String, ByVal colOffset As Long, values As Collection)
    Dim ws As Worksheet, r As Long, nextK As Long, i As Long, n As Long
    Set ws = SetupSheet()
    If ws Is Nothing Then Exit Sub
    r = FindKeyRow(ws, key)
    If r = 0 Then Exit Sub
    nextK = NextKeyRow(ws, r)
    For i = r + 1 To nextK - 1
        If Trim(CStr(ws.Cells(i, VAL_COL).Value)) <> "" Then
            n = n + 1
            If n <= values.Count Then ws.Cells(i, VAL_COL + colOffset).Value = values(n)
        End If
    Next i
End Sub

Private Function SetupSheet() As Worksheet
    On Error Resume Next
    Set SetupSheet = ThisWorkbook.Worksheets(SETUP_SHEET)
    On Error GoTo 0
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

    On Error Resume Next
    ctx("failed") = True
    ctx("failNumber") = e.Number
    ctx("failMsg") = e.Description
    On Error GoTo 0

    Core_Log ctx, "FAILED " & actionId & ": " & m
    If Not IsQuiet(ctx) Then _
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

Private Function IsChainId(ByVal id As String) As Boolean
    IsChainId = (LCase(Left$(id, 6)) = "chain.")
End Function

Private Function IsQuiet(ctx As Object) As Boolean
    On Error Resume Next
    IsQuiet = (ctx.Exists("quiet"))
    If IsQuiet Then IsQuiet = (ctx("quiet") = True)
End Function

' Completion sink. Handlers call this instead of MsgBox for their final summary.
' When the context is quiet (a chain step), the message is captured on the ctx
' and the dialog is suppressed; otherwise it shows exactly as a direct MsgBox.
Public Sub Core_Done(ctx As Object, ByVal msg As String, Optional ByVal title As String = "Done")
    On Error Resume Next
    ctx("doneMsg") = msg
    ctx("done") = True
    On Error GoTo 0
    If Not IsQuiet(ctx) Then MsgBox msg, vbInformation, title
End Sub

' Run one action under an isolated child context. Returns a result dictionary;
' never shows its own error/success dialog (callers decide). quiet=True suppresses
' the handler's Core_Done completion dialog and captures its message instead.
Private Function RunAction(ByVal actionId As String, ByVal quiet As Boolean) As Object
    Dim res As Object, reg As Object, ctx As Object, entry As String
    Set res = CreateObject("Scripting.Dictionary")
    res("id") = actionId
    res("label") = actionId
    res("ok") = False
    res("found") = False
    res("message") = ""
    res("errNumber") = 0
    res("errText") = ""
    res("reported") = False

    On Error GoTo StepFail
    Set reg = FindActionOwner(actionId, entry)
    If reg Is Nothing Then
        On Error GoTo 0
        Set RunAction = res            ' found stays False
        Exit Function
    End If
    res("found") = True
    res("label") = ActionCaption(reg, actionId)

    Set ctx = NewCtx(actionId, reg)
    ctx("quiet") = quiet
    ctx("done") = False
    ctx("doneMsg") = ""

    Application.Run entry, ctx
    On Error GoTo 0
    If CtxFlag(ctx, "failed") Then         ' handler swallowed its own error
        res("ok") = False
        res("errNumber") = CtxNum(ctx, "failNumber")
        res("errText") = CtxStr(ctx, "failMsg")
        res("reported") = Not quiet        ' standalone: Core_Fail already showed it
    Else
        res("ok") = True
        res("message") = CtxStr(ctx, "doneMsg")
    End If
    Set RunAction = res
    Exit Function

StepFail:
    RestoreSpeed                       ' fail-safe: never leave Excel frozen
    res("ok") = False
    res("errNumber") = Err.Number
    res("errText") = Err.Description
    Set RunAction = res
End Function

' Caption for an action id from its owning manifest (for result labels/summaries).
Private Function ActionCaption(reg As Object, ByVal actionId As String) As String
    Dim a As Object
    ActionCaption = actionId
    For Each a In reg("actions")
        If LCase(CStr(a("id"))) = LCase(actionId) Then
            ActionCaption = CStr(a("caption")): Exit Function
        End If
    Next a
End Function

' Safe ctx accessors - avoid Scripting.Dictionary missing-key errors.
Private Function CtxFlag(ctx As Object, ByVal k As String) As Boolean
    On Error Resume Next
    If ctx.Exists(k) Then CtxFlag = (ctx(k) = True)
End Function

Private Function CtxStr(ctx As Object, ByVal k As String) As String
    On Error Resume Next
    If ctx.Exists(k) Then CtxStr = CStr(ctx(k))
End Function

Private Function CtxNum(ctx As Object, ByVal k As String) As Long
    On Error Resume Next
    If ctx.Exists(k) Then CtxNum = CLng(ctx(k))
End Function

' Button dispatcher. Every button's OnAction is  'Core_Run "<actionId>"'
Public Sub Core_Run(ByVal actionId As String)
    Dim res As Object
    If IsChainId(actionId) Then Core_RunChain actionId: Exit Sub
    Set res = RunAction(actionId, False)
    If Not res("found") Then
        MsgBox "Action not installed: " & actionId, vbExclamation, "Core"
    ElseIf Not res("ok") And Not res("reported") Then
        MsgBox "Action '" & actionId & "' failed:" & vbCrLf & vbCrLf & _
               "#" & res("errNumber") & "  " & res("errText"), vbExclamation, "Macro Error"
    End If
End Sub

' Run a chain: each step quiet; prompt Continue/Stop on failure; one end summary.
' No outer Core_SpeedOn/Off here - each step self-manages its own speed via its
' handler; wrapping the whole run would corrupt the depth-counted speed state
' when a step fails (RunAction's StepFail path hard-restores via RestoreSpeed).
Private Sub Core_RunChain(ByVal chainId As String)
    Dim chains As Collection, ch As Object, target As Object
    Dim stepId As Variant, res As Object, n As Long, ok As Long, summary As String
    Dim stopped As Boolean, total As Long
    On Error GoTo RunnerFail

    Set chains = ReadChains(CfgSheet())
    For Each ch In chains
        If LCase(CStr(ch("id"))) = LCase(chainId) Then Set target = ch: Exit For
    Next ch
    If target Is Nothing Then
        MsgBox "Chain not found: " & chainId, vbExclamation, "Core": Exit Sub
    End If
    total = target("steps").Count

    For Each stepId In target("steps")
        n = n + 1
        Set res = RunAction(CStr(stepId), True)      ' quiet; each step owns its speed
        If Not res("found") Then
            summary = summary & n & ". " & stepId & "  X not installed" & vbCrLf
            If Not AskContinue(n, CStr(stepId), "action not installed") Then stopped = True: Exit For
        ElseIf res("ok") Then
            ok = ok + 1
            summary = summary & n & ". " & res("label") & "  " & ChrW$(&H2713) & _
                      IIf(Len(res("message")) > 0, "  (" & FirstLine(CStr(res("message"))) & ")", "") & vbCrLf
        Else
            summary = summary & n & ". " & res("label") & "  X  #" & res("errNumber") & _
                      " " & res("errText") & vbCrLf
            If Not AskContinue(n, CStr(res("label")), CStr(res("errText"))) Then stopped = True: Exit For
        End If
    Next stepId

    Dim head As String
    head = target("label") & IIf(stopped, "  - stopped at step " & n, "  - done") & _
           "  (" & ok & "/" & total & " ok)"
    MsgBox head & vbCrLf & vbCrLf & summary, _
           IIf(stopped, vbExclamation, vbInformation), "Chain"
    Exit Sub
RunnerFail:
    RestoreSpeed
    MsgBox "Chain '" & chainId & "' error at step " & n & ":" & vbCrLf & vbCrLf & _
           "#" & Err.Number & "  " & Err.Description, vbExclamation, "Chain"
End Sub

' "Step N <label> failed: <reason>. Continue with the rest?"  Yes=continue.
Private Function AskContinue(ByVal n As Long, ByVal label As String, ByVal reason As String) As Boolean
    AskContinue = (MsgBox("Step " & n & "  '" & label & "'  failed:" & vbCrLf & vbCrLf & _
        reason & vbCrLf & vbCrLf & "Continue with the remaining steps?", _
        vbExclamation + vbYesNo, "Chain - step failed") = vbYes)
End Function

Private Function FirstLine(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, vbCr): If p = 0 Then p = InStr(s, vbLf)
    If p > 0 Then FirstLine = Left$(s, p - 1) Else FirstLine = s   ' NOT IIf: eager Left$(s,-1) would raise
End Function

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
    Dim banner As String
    For Each reg In BuildRegistry()
        nPkgs = nPkgs + 1
        banner = UCase(CStr(reg("name")))
        r = SectionHeader(ws, r, banner)
        r = RenderTier(ws, r, reg, "simple", True, banner)   ' simple fields, with their sections
        If HasAdvanced(reg) Then
            r = SectionHeader(ws, r, "ADVANCED  -  most users can leave these as-is")
            r = RenderTier(ws, r, reg, "advanced", False, banner)  ' advanced, collected flat
        End If
        r = r + 1
    Next reg

    r = RenderComposeSection(ws, r)

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
    ws.Columns(RID_COL).Hidden = True
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

' Render every field of the given tier. When showSections is True, a section
' header is emitted just before its first matching field (so a section whose
' fields are all in the other tier is not drawn). When False (the advanced
' pass), section headers are skipped and the fields render flat under the one
' ADVANCED band.
Private Function RenderTier(ws As Worksheet, ByVal r As Long, reg As Object, _
                            ByVal tier As String, ByVal showSections As Boolean, _
                            ByVal bannerText As String) As Long
    Dim f As Object, pending As String, havePending As Boolean
    For Each f In reg("fields")
        If f("kind") = "section" Then
            ' Skip a section header that just repeats the package banner (a
            ' description identical to the header above it reads as a mistake).
            If UCase(Trim(CStr(f("label")))) <> bannerText Then
                pending = CStr(f("label"))
                havePending = True
            End If
        ElseIf CStr(f("tier")) = tier Then
            If showSections And havePending Then
                r = SectionHeader(ws, r, pending)
                havePending = False
            End If
            r = RenderField(ws, r, reg, f)
        End If
    Next f
    RenderTier = r
End Function

Private Function RenderField(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    Select Case f("kind")
        Case "list"
            RenderField = RenderList(ws, r, reg, f)
        Case "table"
            RenderField = RenderTable(ws, r, reg, f)
        Case Else
            RenderField = RenderScalar(ws, r, reg, f)
    End Select
End Function

Private Function HasAdvanced(reg As Object) As Boolean
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If CStr(f("tier")) = "advanced" Then
                HasAdvanced = True
                Exit Function
            End If
        End If
    Next f
End Function

Private Function RenderScalar(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    ws.Cells(r, KEY_COL).Value = f("key")
    LabelCell ws.Cells(r, 2), CStr(f("label"))
    ws.Cells(r, VAL_COL).Value = ScalarValue(reg, CStr(f("key")))
    Yellow ws.Cells(r, VAL_COL)
    r = r + 1
    r = RenderHelp(ws, r, CStr(f("help")))
    RenderScalar = r
End Function

' A plain-language note under a field, grey + italic + wrapped. Its column A and
' column C stay blank, so it is invisible to the config block readers (which key
' off column A) and the list/table readers (which key off column C). Returns the
' next row (unchanged when there is no help text).
Private Function RenderHelp(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    If Trim(txt) = "" Then RenderHelp = r: Exit Function
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Italic = True
    ws.Cells(r, 2).Font.Size = 9
    ws.Cells(r, 2).Font.Color = RGB(120, 120, 120)
    ws.Cells(r, 2).WrapText = True
    ws.Cells(r, 2).VerticalAlignment = xlTop
    RenderHelp = r + 1
End Function

' Write a field label into column B, wrapped + top-aligned, so a long label
' stays fully visible next to the occupied value cell in column C.
Private Sub LabelCell(rng As Range, ByVal txt As String)
    rng.Value = txt
    rng.WrapText = True
    rng.VerticalAlignment = xlTop
End Sub

' A full-width tinted band row: columns B:E merged, text wrapped. Merged cells are
' ignored by AutoFit, so the row height is set from a line estimate (erring tall
' to avoid clipping). Column A and column C are inside the merge with no separate
' value, so the row stays invisible to the config block readers. Returns r + 1.
Private Function TintRow(ws As Worksheet, ByVal r As Long, ByVal txt As String, _
                         ByVal sz As Single, ByVal bold As Boolean, _
                         ByVal ital As Boolean, ByVal colr As Long) As Long
    Dim lines As Long
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Merge
        .Interior.Color = RGB(221, 217, 240)
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    With ws.Cells(r, 2)
        .Value = txt
        .Font.Size = sz
        .Font.Bold = bold
        .Font.Italic = ital
        .Font.Color = colr
    End With
    lines = (Len(txt) \ 115) + 1                 ' ~115 chars/line across B:E; round up
    ws.Rows(r).RowHeight = lines * (sz + 4)
    TintRow = r + 1
End Function

' For list/table fields: draw the label and the explanation in the section-tinted
' band ABOVE the column-header row, so the headers sit directly on top of the
' input cells (no gap). The band rows are merged B:E; they are invisible to the
' config block readers. Returns the row for the header/key row.
Private Function BlockIntro(ws As Worksheet, ByVal r As Long, _
                            ByVal label As String, ByVal help As String) As Long
    r = TintRow(ws, r, label, 11, True, False, RGB(54, 32, 100))
    If Trim(help) <> "" Then r = TintRow(ws, r, help, 9, False, True, RGB(90, 70, 120))
    BlockIntro = r
End Function

Private Function RenderList(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    Dim vals As Collection, i As Long
    r = BlockIntro(ws, r, CStr(f("label")), CStr(f("help")))   ' label + help above the header
    ws.Cells(r, KEY_COL).Value = f("key")
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
    r = BlockIntro(ws, r, CStr(f("label")), CStr(f("help")))   ' label + help above the headers
    ws.Cells(r, KEY_COL).Value = f("key")
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
' 5b. CORE-OWNED SETUP SCHEMA  (chains + placement; NOT a package)
'   Rendered by Core_BuildSetup outside the package-discovery loop.
'================================================================
Private Const CHAINS_KEY As String = "k_core_chains"   ' block sentinel (col A)
Private Const LAYOUT_KEY As String = "k_core_layout"   ' block sentinel (col A)
Private Const RID_COL    As Long = 10                  ' col J: hidden per-row id

Private Function RenderComposeSection(ws As Worksheet, ByVal r As Long) As Long
    r = SectionHeader(ws, r, "COMPOSE YOUR BUTTONS")
    r = TintRow(ws, r, "Build multi-step buttons (Chains) and arrange buttons " & _
                "onto pages (Page code). Leave everything blank to keep one Macros tab.", _
                10, False, True, RGB(90, 90, 90))
    r = r + 1
    r = RenderChainsBlock(ws, r)
    r = r + 1
    r = RenderPlacementBlock(ws, r)
    RenderComposeSection = r
End Function

' "Label [id]" per line -> Collection of ids (only the bracketed id is authoritative)
Private Function ParseStepTokens(ByVal cell As String) As Collection
    Dim out As New Collection, lines() As String, ln As String, i As Long
    Dim lo As Long, hi As Long, id As String
    Set ParseStepTokens = out
    cell = Replace(cell, ";", vbLf)
    lines = Split(cell, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim(lines(i))
        If Len(ln) > 0 Then
            lo = InStrRev(ln, "["): hi = InStrRev(ln, "]")
            If lo > 0 And hi > lo Then
                id = Trim(Mid(ln, lo + 1, hi - lo - 1))
            Else
                id = ln            ' bare id with no brackets is tolerated
            End If
            If Len(id) > 0 Then out.Add id
        End If
    Next i
End Function

' "Import Source Data [importexport.import]" for an action id (falls back to the id).
Private Function ActionTokenFor(ByVal actionId As String) As String
    Dim reg As Object, a As Object
    ActionTokenFor = actionId & " [" & actionId & "]"
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If LCase(CStr(a("id"))) = LCase(actionId) Then
                ActionTokenFor = CStr(a("caption")) & " [" & CStr(a("id")) & "]"
                Exit Function
            End If
        Next a
    Next reg
End Function

Private Function StepsToTokens(steps As Collection) As String
    Dim s As String, v As Variant
    For Each v In steps
        s = s & ActionTokenFor(CStr(v)) & vbLf
    Next v
    StepsToTokens = Left$(s, Len(s) - IIf(Len(s) > 0, 1, 0))
End Function

Private Sub RenderActionTokens(ws As Worksheet, ByVal startRow As Long)
    Dim reg As Object, a As Object, gr As Long
    ws.Cells(startRow - 1, 7).Value = "Available actions - copy a token into Steps:"
    ws.Cells(startRow - 1, 7).Font.Bold = True
    gr = startRow
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            ws.Cells(gr, 7).Value = CStr(a("caption")) & " [" & CStr(a("id")) & "]"
            ws.Cells(gr, 7).Font.Color = RGB(120, 120, 120)
            gr = gr + 1
        Next a
    Next reg
End Sub

' chain.<slug>, unique within `used` (a dictionary of lowercased ids already taken).
Private Function MintChainId(ByVal label As String, used As Object) As String
    Dim base As String, s As String, ch As String, i As Long, cand As String, n As Long
    For i = 1 To Len(label)
        ch = LCase(Mid(label, i, 1))
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            base = base & ch
        ElseIf Len(base) > 0 And Right$(base, 1) <> "_" Then
            base = base & "_"
        End If
    Next i
    base = Trim(Replace(base, "__", "_"))
    Do While Len(base) > 0 And Right$(base, 1) = "_"
        base = Left$(base, Len(base) - 1)
    Loop
    If Len(base) = 0 Then base = "chain"
    cand = "chain." & base
    Do While used.Exists(LCase(cand))
        n = n + 1
        cand = "chain." & base & "_" & n
    Loop
    MintChainId = cand
End Function

Private Function RenderChainsBlock(ws As Worksheet, ByVal r As Long) As Long
    Dim existing As Collection, ch As Object, i As Long, tokens As String
    r = BlockIntro(ws, r, "Chains  -  one button that runs several actions in order", _
                   "Button name + Steps (paste tokens from the Actions list, one per line) + Page code.")
    ws.Cells(r, KEY_COL).Value = CHAINS_KEY
    ws.Cells(r, VAL_COL).Value = "Button name"
    ws.Cells(r, VAL_COL + 1).Value = "Steps (one token per line)"
    ws.Cells(r, VAL_COL + 2).Value = "Page code"
    ws.Range(ws.Cells(r, VAL_COL), ws.Cells(r, VAL_COL + 2)).Font.Bold = True
    r = r + 1

    Set existing = ReadChains(CfgSheet())          ' re-seed from activated config
    For i = 1 To MAX_LIST
        If i <= existing.Count Then
            Set ch = existing(i)
            ws.Cells(r, VAL_COL).Value = ch("label")
            ws.Cells(r, VAL_COL + 1).Value = StepsToTokens(ch("steps"))   ' Label [id] per line
            ws.Cells(r, VAL_COL + 2).Value = ch("code")
            ws.Cells(r, RID_COL).Value = ch("id")   ' preserve immutable id
        End If
        Yellow ws.Cells(r, VAL_COL)
        Yellow ws.Cells(r, VAL_COL + 1)
        Yellow ws.Cells(r, VAL_COL + 2)
        ws.Cells(r, VAL_COL + 1).WrapText = True
        r = r + 1
    Next i
    ' Copy-ready action tokens in G/H for authoring Steps.
    RenderActionTokens ws, r - MAX_LIST
    RenderChainsBlock = r
End Function

' Read the chains block from ws (CfgSheet at run time; SetupSheet for validation).
' Each item: {id, label, steps(Collection of ids), code}. IDs are read from the
' hidden RID column when present; a row with a label but no id has none yet.
Private Function ReadChains(ws As Worksheet) As Collection
    Dim out As New Collection, kr As Long, nk As Long, i As Long
    Dim label As String, ch As Object
    Set ReadChains = out
    If ws Is Nothing Then Exit Function
    kr = FindKeyRow(ws, CHAINS_KEY)
    If kr = 0 Then Exit Function
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1
        label = Trim(CStr(ws.Cells(i, VAL_COL).Value))
        If Len(label) > 0 Then
            Set ch = CreateObject("Scripting.Dictionary")
            ch("label") = label
            ch("id") = Trim(CStr(ws.Cells(i, RID_COL).Value))
            ch("code") = Trim(CStr(ws.Cells(i, VAL_COL + 2).Value))
            Set ch("steps") = ParseStepTokens(CStr(ws.Cells(i, VAL_COL + 1).Value))
            out.Add ch
        End If
    Next i
End Function

' One row per registry action: hidden id in RID_COL, "caption [id]" (grey) in C,
' a yellow Page-code cell in D. Prior codes are preserved from CfgSheet() by id.
Private Function RenderPlacementBlock(ws As Worksheet, ByVal r As Long) As Long
    Dim prior As Object, reg As Object, a As Object, id As String
    r = BlockIntro(ws, r, "Actions on pages  -  where each action button goes", _
                   "Page code: number = page, letters = column (e.g. 1A). Blank = page 1. " & _
                   "Type  hide  to leave a button off entirely.")
    ws.Cells(r, KEY_COL).Value = LAYOUT_KEY
    ws.Cells(r, VAL_COL).Value = "Action"
    ws.Cells(r, VAL_COL + 1).Value = "Page code"
    ws.Range(ws.Cells(r, VAL_COL), ws.Cells(r, VAL_COL + 1)).Font.Bold = True
    r = r + 1

    Set prior = ReadCodeMap(CfgSheet())            ' preserve prior codes by id
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            id = CStr(a("id"))
            ws.Cells(r, VAL_COL).Value = CStr(a("caption")) & " [" & id & "]"
            ws.Cells(r, VAL_COL).Font.Color = RGB(90, 90, 90)
            ws.Cells(r, RID_COL).Value = id
            If prior.Exists(LCase(id)) Then ws.Cells(r, VAL_COL + 1).Value = prior(LCase(id))
            Yellow ws.Cells(r, VAL_COL + 1)
            r = r + 1
        Next a
    Next reg
    RenderPlacementBlock = r
End Function

' id(lower) -> code, merging the placement (actions) block and the chains block.
Private Function ReadCodeMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String
    Dim chains As Collection, ch As Object
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadCodeMap = map
    If ws Is Nothing Then Exit Function
    ' actions
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr > 0 Then
        nk = NextKeyRow(ws, kr)
        For i = kr + 1 To nk - 1
            id = Trim(CStr(ws.Cells(i, RID_COL).Value))
            If Len(id) > 0 Then map(LCase(id)) = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
        Next i
    End If
    ' chains
    Set chains = ReadChains(ws)
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 Then map(LCase(CStr(ch("id")))) = CStr(ch("code"))
    Next ch
End Function

' Mint chain.<slug> into any chains-block row that has a Button name but no id.
' Runs on the LIVE Setup sheet AFTER the user confirms, so Finalize's copy persists them.
Private Sub AssignChainIds(ws As Worksheet)
    Dim kr As Long, nk As Long, i As Long, used As Object, label As String, id As String
    Set used = CreateObject("Scripting.Dictionary")
    kr = FindKeyRow(ws, CHAINS_KEY)
    If kr = 0 Then Exit Sub
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1                 ' register existing ids
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(id) > 0 Then used(LCase(id)) = True
    Next i
    For i = kr + 1 To nk - 1                 ' mint for named rows lacking an id
        label = Trim(CStr(ws.Cells(i, VAL_COL).Value))
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(label) > 0 And Len(id) = 0 Then
            id = MintChainId(label, used)
            ws.Cells(i, RID_COL).Value = id
            used(LCase(id)) = True
        End If
    Next i
End Sub

' Build the review text + count STRUCTURAL problems (things that must be fixed
' before activating). Validates chain steps against the registry, action + chain
' ids against the reserved chain.* namespace, and every code against ParseCode.
' Unassigned/hidden button lists and the layout note are SOFT (informational).
Private Function ComposeReport(ws As Worksheet, ByRef structural As Long) As String
    Dim s As String, chains As Collection, ch As Object, ids As Object, stepId As Variant
    Dim tabNum As Long, colIx As Long, colL As String, code As String
    Dim map As Object, k As Variant, unassigned As String, hidden As String
    Dim regX As Object, aX As Object
    Dim ch2 As Object, actIds As Object, seenIds As Object, cid As String
    Set ids = ActionIdSet()                        ' lowercased set of valid action ids
    Set chains = ReadChains(ws)

    s = "CHAINS:" & vbCrLf
    If chains.Count = 0 Then s = s & "  (none)" & vbCrLf
    For Each ch In chains
        s = s & "  - " & ch("label") & "  ->  "
        If ch("steps").Count = 0 Then
            s = s & "(no steps!)": structural = structural + 1
        Else
            For Each stepId In ch("steps")
                s = s & CStr(stepId)
                If Not ids.Exists(LCase(CStr(stepId))) Then
                    s = s & "[UNKNOWN]": structural = structural + 1
                End If
                s = s & "  "
            Next stepId
        End If
        code = CStr(ch("code"))
        If Len(code) > 0 And LCase(code) <> "hide" Then
            If Not ParseCode(code, tabNum, colIx, colL) Then _
                s = s & "  {bad code '" & code & "'}": structural = structural + 1
        End If
        s = s & vbCrLf
    Next ch

    ' Placement codes: validity + collect unassigned/hidden.
    Set map = ReadCodeMap(ws)
    For Each k In map.Keys
        code = CStr(map(k))
        If LCase(code) = "hide" Then
            hidden = hidden & "  - " & k & vbCrLf
        ElseIf Len(code) = 0 Then
            unassigned = unassigned & "  - " & k & vbCrLf
        ElseIf Not ParseCode(code, tabNum, colIx, colL) Then
            s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf: structural = structural + 1
        End If
    Next k

    ' (a) no package action id may live in the reserved chain.* namespace
    For Each regX In BuildRegistry()
        For Each aX In regX("actions")
            If IsChainId(CStr(aX("id"))) Then
                s = s & "RESERVED: action id '" & aX("id") & "' uses chain.*" & vbCrLf
                structural = structural + 1
            End If
        Next aX
    Next regX
    ' (b) each existing chain id must start chain., not collide with an action id, and be unique
    Set actIds = ActionIdSet()
    Set seenIds = CreateObject("Scripting.Dictionary")
    For Each ch2 In ReadChains(ws)
        cid = CStr(ch2("id"))
        If Len(cid) > 0 Then
            If Not IsChainId(cid) Then _
                s = s & "BAD chain id '" & cid & "' (must start chain.)" & vbCrLf: structural = structural + 1
            If actIds.Exists(LCase(cid)) Then _
                s = s & "COLLISION: chain id '" & cid & "' matches an action id" & vbCrLf: structural = structural + 1
            If seenIds.Exists(LCase(cid)) Then _
                s = s & "DUPLICATE chain id '" & cid & "'" & vbCrLf: structural = structural + 1
            seenIds(LCase(cid)) = True
        End If
    Next ch2

    If AnyCode(map) Then
        s = s & vbCrLf & "A layout is ACTIVE (numbered pages will be built)." & vbCrLf
        If Len(unassigned) > 0 Then s = s & "Unassigned -> page 1:" & vbCrLf & unassigned
        If Len(hidden) > 0 Then s = s & "Hidden (not placed):" & vbCrLf & hidden
    Else
        s = s & vbCrLf & "No page codes -> a single 'Macros' tab (current behavior)." & vbCrLf
    End If
    ComposeReport = s
End Function

Private Function ActionIdSet() As Object
    Dim d As Object, reg As Object, a As Object
    Set d = CreateObject("Scripting.Dictionary")
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            d(LCase(CStr(a("id")))) = True
        Next a
    Next reg
    Set ActionIdSet = d
End Function

Private Function AnyCode(map As Object) As Boolean
    Dim k As Variant
    For Each k In map.Keys
        If Len(Trim(CStr(map(k)))) > 0 Then AnyCode = True: Exit Function
    Next k
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

    ' --- Compose Your Buttons: validate chains + placement ---
    Dim composeRep As String, structural As Long
    composeRep = ComposeReport(ws, structural)
    msg = msg & vbCrLf & vbCrLf & String(40, "-") & vbCrLf & composeRep
    If structural > 0 Then
        MsgBox msg & vbCrLf & structural & " problem(s) above must be fixed before activating.", _
               vbExclamation, "Finalize Setup Tab"
        Exit Sub
    End If

    If MsgBox(msg & vbCrLf & "Activate this configuration now?", _
              vbQuestion + vbYesNo, "Finalize Setup Tab") = vbNo Then Exit Sub

    AssignChainIds ws          ' mint ids now, after confirm, before the copy persists them

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
' Legacy single "Macros" tab when no page codes exist; grouped "Macros N" pages
' when a layout is active. Prior panels are cleared before either mode draws.
Public Sub Core_BuildControlPanel()
    Dim buttons As Collection, codeMap As Object
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Set buttons = GatherButtons()
    Set codeMap = ReadCodeMap(CfgSheet())
    DeletePanels                         ' remove any prior "Macros N" pages
    If AnyCode(codeMap) Then
        DeleteLegacyPanel                ' grouped mode: drop the stale single "Macros" tab
        BuildGroupedPanels buttons, codeMap
    Else
        BuildSinglePanel buttons         ' legacy single "Macros" tab (reused/cleared)
    End If
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "Core_BuildControlPanel error: " & Err.Description, vbExclamation, "Setup"
End Sub

' Drop the single legacy "Macros" tab when switching into grouped mode.
Private Sub DeleteLegacyPanel()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(MACRO_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' Each item: Array(id, caption, onAction, help). Actions from the registry,
' chains from activated config. Admin buttons are added by the callers.
Private Function GatherButtons() As Collection
    Dim out As New Collection, reg As Object, a As Object, chains As Collection, ch As Object
    Dim onAct As String
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            onAct = "'Core_Run " & Chr(34) & CStr(a("id")) & Chr(34) & "'"
            out.Add Array(CStr(a("id")), CStr(a("caption")), onAct, CStr(a("help")))
        Next a
    Next reg
    Set chains = ReadChains(CfgSheet())
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 Then
            onAct = "'Core_Run " & Chr(34) & CStr(ch("id")) & Chr(34) & "'"
            out.Add Array(CStr(ch("id")), CStr(ch("label")), onAct, _
                          "Chain: " & StepsInline(ch("steps")))
        End If
    Next ch
    Set GatherButtons = out
End Function

Private Function StepsInline(steps As Collection) As String
    Dim s As String, v As Variant
    For Each v In steps
        s = s & IIf(Len(s) > 0, " > ", "") & CStr(v)
    Next v
    StepsInline = s
End Function

' Legacy layout: one "Macros" tab, action + chain buttons, then Core's two admin
' buttons appended with the same cadence. Matches the pre-grouping look.
Private Sub BuildSinglePanel(buttons As Collection)
    Dim ws As Worksheet, drawList As New Collection, e As Variant
    Dim i As Long, leftPos As Double, topPos As Double
    Dim btnW As Double, btnH As Double, gap As Double, helpRow As Long

    Set ws = EnsurePanelSheet(MACRO_SHEET, True)

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

    ' Daily action + chain buttons first, then Core's own setup buttons last.
    For Each e In buttons
        drawList.Add Array(CStr(e(1)), CStr(e(2)), CStr(e(3)))
    Next e
    drawList.Add Array("Create Setup Tab", "Core_BuildSetup", _
                       "Builds a pre-filled 'Setup' tab from the installed packages. Use it to reuse these macros in another workbook.")
    drawList.Add Array("Finalize Setup Tab", "Core_FinalizeSetup", _
                       "Validates your Setup answers and activates them. The action buttons then use your settings instead of the built-in defaults.")

    ' Cadence: btnH + gap = 45pt = 3 rows at 15pt, so each description row lines
    ' up with its button. Force the row heights so the alignment holds regardless
    ' of the sheet's default.
    btnW = 220: btnH = 32: gap = 13
    leftPos = ws.Range("B5").Left
    topPos = ws.Range("B5").Top
    ws.Rows("5:" & (5 + drawList.Count * 3 + 2)).RowHeight = 15
    helpRow = 5
    i = 0
    For Each e In drawList
        i = i + 1
        AddBtn ws, "btn" & i, CStr(e(0)), CStr(e(1)), leftPos, topPos, btnW, btnH
        ' Description spans 2 rows (~button height) so a long, wrapped description
        ' shows in full instead of being clipped to one 15pt row.
        With ws.Range(ws.Cells(helpRow, 3), ws.Cells(helpRow + 1, 3))
            .Merge
            .Value = CStr(e(2))
            .Font.Size = 10
            .WrapText = True
            .VerticalAlignment = xlCenter
        End With
        topPos = topPos + btnH + gap
        helpRow = helpRow + 3
    Next e

    ws.Activate
    ws.Range("A1").Select
    MsgBox "Macros tab built with " & buttons.Count & " button(s).", vbInformation, "Setup Complete"
End Sub

' Grouped layout: place each button onto a "Macros N" page by its code. Blank code
' -> page 1 Unassigned column (9999); "hide" -> omitted; unparseable -> Unassigned.
' Admin buttons always land in a reserved "Setup" column (9998) on page 1.
Private Sub BuildGroupedPanels(buttons As Collection, codeMap As Object)
    Dim pages As Object, e As Variant, id As String, code As String
    Dim tabNum As Long, colIx As Long, colL As String, unassigned As New Collection
    Set pages = CreateObject("Scripting.Dictionary")
    For Each e In buttons
        id = LCase(CStr(e(0)))
        code = ""
        If codeMap.Exists(id) Then code = Trim(CStr(codeMap(id)))
        If LCase(code) = "hide" Then GoTo NextBtn
        If Len(code) = 0 Then
            unassigned.Add e
        ElseIf ParseCode(code, tabNum, colIx, colL) Then
            AddToPage pages, tabNum, colIx, e
        Else
            unassigned.Add e            ' unparseable degrades to Unassigned (already blocked at Finalize)
        End If
NextBtn:
    Next e
    If Not pages.Exists(1) Then pages.Add 1, CreateObject("Scripting.Dictionary")   ' force page 1
    For Each e In unassigned
        AddToPage pages, 1, 9999, e
    Next e
    ' admin controls always on page 1, in a reserved "Setup" column
    AddToPage pages, 1, 9998, Array("", "Create Setup Tab", "Core_BuildSetup", "")
    AddToPage pages, 1, 9998, Array("", "Finalize Setup Tab", "Core_FinalizeSetup", "")
    Dim tabKeys As Variant, ti As Long, firstPanel As Worksheet, panel As Worksheet
    tabKeys = SortedKeys(pages)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        Set panel = RenderPage(pages, CLng(tabKeys(ti)))
        If firstPanel Is Nothing Then Set firstPanel = panel
    Next ti
    If Not firstPanel Is Nothing Then firstPanel.Activate
    MsgBox "Built " & pages.Count & " page(s).", vbInformation, "Setup Complete"
End Sub

' Get-or-add a panel sheet, clear its shapes (backward loop) + cells, and format.
Private Function EnsurePanelSheet(ByVal nm As String, ByVal atFront As Boolean) As Worksheet
    Dim ws As Worksheet, si As Long
    If Core_SheetExists(nm) Then
        Set ws = ThisWorkbook.Worksheets(nm)
    ElseIf atFront Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.name = nm
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = nm
    End If
    For si = ws.Shapes.Count To 1 Step -1     ' remove old buttons + textboxes safely
        ws.Shapes(si).Delete
    Next si
    ws.Cells.Clear
    ws.Cells.UnMerge
    ws.Tab.Color = RGB(31, 78, 121)
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 47
    ws.Columns("C").ColumnWidth = 85
    Set EnsurePanelSheet = ws
End Function

' Insert one button-array into the nested pages dictionary: tabNum -> (colIx -> Collection).
Private Sub AddToPage(pages As Object, ByVal tabNum As Long, ByVal colIx As Long, ByVal btn As Variant)
    Dim cols As Object, items As Collection
    If Not pages.Exists(tabNum) Then pages.Add tabNum, CreateObject("Scripting.Dictionary")
    Set cols = pages(tabNum)
    If Not cols.Exists(colIx) Then cols.Add colIx, New Collection
    Set items = cols(colIx)
    items.Add btn                 ' btn = Array(id, caption, onAction, help)
End Sub

' Draw one "Macros N" page: title, per-column headers (Setup/Unassigned/Column X),
' and the stacked buttons. Columns 9998 = "Setup", 9999 = "Unassigned".
Private Function RenderPage(pages As Object, ByVal tabNum As Long) As Worksheet
    Dim panel As Worksheet, cols As Object, colKeys As Variant, ci As Long
    Dim colIx As Long, items As Collection, ii As Long, it As Variant
    Dim leftBase As Double, topBase As Double, colLeft As Double, topPos As Double, hdr As String
    Set panel = EnsurePanelSheet(PANEL_BASE & tabNum, False)
    leftBase = panel.Range("B4").Left: topBase = panel.Range("B4").Top
    AddText panel, leftBase, panel.Range("B2").Top, 400, 24, "Macros - Page " & tabNum, 15, True, RGB(31, 78, 121)
    Set cols = pages(tabNum)
    colKeys = SortedKeys(cols)
    For ci = LBound(colKeys) To UBound(colKeys)
        colIx = colKeys(ci)
        colLeft = leftBase + ci * (BTN_W + COL_GAP)          ' ci is 0-based across sorted columns
        Set items = cols(colIx)
        If colIx = 9998 Then
            hdr = "Setup"
        ElseIf colIx = 9999 Then
            hdr = "Unassigned"
        Else
            hdr = "Column " & IndexToCol(colIx)
        End If
        AddText panel, colLeft, topBase, BTN_W, HDR_H, hdr, 11, True, RGB(90, 90, 90)
        topPos = topBase + HDR_H + ROW_GAP
        For ii = 1 To items.Count
            it = items(ii)
            AddBtn panel, "b" & tabNum & "_" & colIx & "_" & ii, CStr(it(1)), CStr(it(2)), colLeft, topPos, BTN_W, BTN_H
            topPos = topPos + BTN_H + ROW_GAP
        Next ii
    Next ci
    panel.Range("A1").Select
    Set RenderPage = panel
End Function

' Frameless textbox label (ported from MacroLayout ML_AddText).
Private Function AddText(ws As Worksheet, ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, _
                         ByVal txt As String, ByVal sz As Single, ByVal bold As Boolean, ByVal colr As Long) As Shape
    Dim sh As Shape
    Set sh = ws.Shapes.AddTextbox(1, l, t, w, h)
    sh.Fill.Visible = msoFalse
    sh.Line.Visible = msoFalse
    With sh.TextFrame2.TextRange
        .Text = txt
        .Font.Size = sz
        .Font.Bold = bold
        .Font.Fill.ForeColor.RGB = colr
    End With
    Set AddText = sh
End Function

' Remove every "Macros N" grouped page (numeric tail; the legacy "Macros" is left).
Private Sub DeletePanels()
    Dim i As Long, nm As String
    Application.DisplayAlerts = False
    On Error Resume Next
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        nm = ThisWorkbook.Worksheets(i).Name
        If IsPanelName(nm) Then ThisWorkbook.Worksheets(i).Delete
    Next i
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Private Function IsPanelName(ByVal nm As String) As Boolean
    Dim tail As String
    If LCase(Left(nm, Len(PANEL_BASE))) <> LCase(PANEL_BASE) Then Exit Function
    tail = Mid(nm, Len(PANEL_BASE) + 1)
    If Len(tail) = 0 Then Exit Function
    IsPanelName = IsNumeric(tail)
End Function

'================================================================
' Setup-tab drawing helpers
'================================================================
Private Sub Yellow(rng As Range)
    rng.Interior.Color = RGB(255, 242, 204)
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(200, 200, 200)
End Sub

Private Function SectionHeader(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    SectionHeader = TintRow(ws, r, txt, 11, True, False, RGB(54, 32, 100))
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

' ---- page-code parsing / column math (ported from MacroLayout; Task 9 reuses) ----

' "12AB" -> tabNum 12, colIndex for "AB". Returns False on malformed input.
' HARDENED: rejects page 0, guards CLng overflow on the digit run.
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
    If i <= Len(s) Then Exit Function
    If Len(digits) = 0 Or Len(letters) = 0 Then Exit Function
    If Len(digits) > 7 Then Exit Function        ' overflow guard (CLng-safe page range)
    tabNum = CLng(digits)
    If tabNum < 1 Then Exit Function             ' pages are 1-based; "0A" is invalid
    colLetter = letters
    For i = 1 To Len(letters)
        colIndex = colIndex * 26 + (Asc(Mid(letters, i, 1)) - Asc("A") + 1)
    Next i
    ParseCode = True
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
