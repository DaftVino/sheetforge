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
Private Const ADMIN_SHEET    As String = "Setup Menu"   ' Create/Finalize controls; hideable
Private Const PACKAGES_SHEET As String = "_Packages"   ' optional enable/disable layer

Private Const CHAIN_CAT      As String = "Chains"      ' reserved display category for chain buttons

' ---- per-action availability (contract v1.2.0) ----
' Stored in the action catalog block as the display label ("Button" / "Chains
' only" / "Off") and parsed to these tokens. Blank means never set and reads as
' Button, which is what makes v0.4.0 a no-op for a v0.3.0 deployment.
Private Const AVAIL_BUTTON   As String = "button"
Private Const AVAIL_CHAIN    As String = "chains only"
Private Const AVAIL_OFF      As String = "off"

' Column header caption for the action catalog's Availability column.
' RenderPlacementBlock writes this literal as the visible header; MarkerVerdict's
' tamper test (case-insensitively) reads it back to decide whether an absent
' schema marker means "pre-v1.2 layout" (recoverable) or "this v1.2-shaped block
' was tampered with" (must not be trusted). The two must stay tied to the same
' Const - if the caption is reworded in only one place, the tamper test silently
' stops firing, a tampered store degrades to "absent", ReadAvailMap returns an
' empty map, and every action the user switched Off goes live again.
Private Const AVAIL_HEADER   As String = "Availability"

' Config-format marker, written into the hidden RID column of the action catalog's
' key row - a machine-owned cell, not a visible header. A display label is
' presentation, is user-editable, and must never decide whether saved config is
' preserved or discarded (contract section 5). Task 4 builds the migration on it;
' it is declared and read here because ReadAvailMap needs it.
' The "v" prefix is load-bearing: a bare "1.2" is stored as General and Excel
' coerces it to the NUMBER 1.2, which is then read back through a locale-aware
' CStr (a comma-decimal locale yields "1,2"; some locales read it as a date).
' A non-numeric token cannot be coerced, so the marker survives the round trip
' byte-for-byte regardless of the workbook's locale.
Private Const SCHEMA_VER     As String = "v1.2"

' ---- button page geometry (v0.4.0: an unmerged CELL-CARD grid) ----
' Cards live in worksheet columns B, D and F, with C and E as narrow gutters.
' NOTHING on this grid may be merged: Excel's AutoFit ignores merged cells (which
' is why the Setup band rows have to estimate their height instead), and AutoFit
' on the description rows is the entire reason a description cannot clip.
Private Const PANEL_BASE As String = "Macros "   ' grouped output tabs: "Macros 1", "Macros 2"...
Private Const HOME_SHEET As String = "Macros Home"  ' generated index; RESERVED name (contract 7)
Private Const CARD_COLS   As Long = 3            ' card columns per page -> B, D, F
Private Const CARD_W      As Double = 30         ' content column width, in characters
Private Const GUTTER_W    As Double = 3          ' spacer column width, in characters
Private Const CARD_HDR_ROW As Long = 3           ' category header row; cards start at 4
                                                 ' (rows 2/4 blank chrome spacers were removed -
                                                 ' fix round 1: title now row 2, header row 3)
Private Const BTN_H      As Double = 28          ' fix round 1: 26 clipped a 2-line wrapped caption
Private Const BTN_ROW_H  As Double = 32          ' button row height (BTN_H + padding)
Private Const DESC_MIN_H As Double = 24          ' floor after AutoFit
' No DESC_MAX_H: contract 6 forbids clipping a description, and a manual clamp on
' an AutoFit row is exactly how Excel clips. AutoFit gets the honest height, full
' stop, for an action's AUTHORED help. A chain's help is engine-SYNTHESISED
' (GatherButtons: "Chain: " & StepsInline(...)), not authored, so contract 6 does
' not bind it - CardHelp (carry-over C, Task 7) bounds it before it reaches the
' card grid instead of letting a pathological 20-step chain draw one tall row
' that inflates every card column beside it; HomeCatalog shows the full recipe.
Private Const AUTO_COL_MAX As Long = 6           ' buttons per card column before a new column
                                                 ' (fix round 1: was 8; 8 only fits full-HD/100%)
Private Const CHAIN_CARD_STEPS As Long = 4       ' step captions shown on a card before "+N more"
' Fix round 1, fix 6: CHAIN_CARD_STEPS alone bounds step COUNT, not length - four
' long captions (an action caption may carry an ordinal and a full phrase) sail
' through untouched and can still reproduce the ~170pt row carry-over C exists to
' prevent. CHAIN_CARD_CHARS bounds the composed width directly; CardHelp stops
' adding steps once either cap is hit, whichever comes first.
Private Const CHAIN_CARD_CHARS As Long = 120     ' composed-text length cap before "+N more" cuts in
' Fix round 2, fix 6: separates a real colIx (from ParseCode, or the 9998/9999
' reserved columns) from a SYNTHETIC colIx minted by SplitTallManualColumns for
' a same-code stack's vertical overflow (piece colIx = origColIx * VSPLIT_MULT +
' pieceIndex). Large enough that no realistic placement code's column (up to a
' 3-letter code, max 18278) collides with it; SplitTallManualColumns itself
' guards the multiply against Long overflow for anything larger.
Private Const VSPLIT_MULT As Long = 100000
' ---- Setup / _Config layout ----
Private Const KEY_COL   As Long = 1       ' column A: machine keys (hidden)
Private Const VAL_COL   As Long = 3       ' column C: scalar value / list value / 1st table col
Private Const MAX_LIST  As Long = 20      ' input rows drawn per list / table block
Private Const END_KEY   As String = "_core_end"   ' sentinel key row: bounds the last list/table block
Private Const CHAINS_KEY As String = "k_core_chains"   ' Compose section: chains block sentinel (col A)
Private Const LAYOUT_KEY As String = "k_core_layout"   ' Compose section: placement block sentinel (col A)
Private Const RID_COL    As Long = 10                  ' Compose section: col J, hidden per-row id

' ---- speed wrapper saved state (module-level, depth-counted) ----
Private mSpeedDepth  As Long
Private mSpeedActive As Boolean
Private mSaveScreen  As Boolean
Private mSaveEvents  As Boolean
Private mSaveAlerts  As Boolean
Private mSaveCalc    As XlCalculation

' Fix round 1, fix 2: Core_BuildControlPanel's own Fail: handler shows its
' error and then returns NORMALLY (a Sub can't propagate an error out), so a
' caller has no way to tell success from failure by the call alone.
' Core_FinalizeSetup needs to, because it shows its own unconditional
' "activated and rebuilt" message right after calling it. Reset to False at
' the very top of every Core_BuildControlPanel call (not just on entry to the
' module) so a stale True from a PRIOR successful call can never leak into a
' later call that fails before reaching its own success line - and only set
' True once the whole pipeline (through BuildHomePage) has actually finished.
Private mLastBuildOk As Boolean

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
    reg("fields").Add NewField("", title, "section", "", "simple", "", "")
End Sub

' tier is "simple" (default) or "advanced". Advanced fields are grouped under an
' "ADVANCED - most users can leave these as-is" band so regular users can stop
' after the simple fields. help is an optional plain-language note shown in grey
' under the field (pass via help:= to skip tier). usedBy is an optional
' comma-separated list of action ids this field configures; a field whose every
' named action is Off moves into the collapsed band at the bottom of Setup (never
' omitted - see Core_BuildSetup). A "Chains only" action still runs, so it still
' keeps its fields in the working area.
Public Sub Core_AddField(reg As Object, ByVal key As String, ByVal label As String, _
                         ByVal kind As String, ByVal dflt As String, _
                         Optional ByVal tier As String = "simple", _
                         Optional ByVal help As String = "", _
                         Optional ByVal usedBy As String = "")
    reg("fields").Add NewField(key, label, kind, dflt, tier, help, usedBy)
End Sub

Public Sub Core_AddList(reg As Object, ByVal key As String, _
                        ByVal label As String, ByVal colHeader As String, _
                        Optional ByVal tier As String = "simple", _
                        Optional ByVal help As String = "", _
                        Optional ByVal usedBy As String = "")
    Dim f As Object
    Set f = NewField(key, label, "list", "", tier, help, usedBy)
    f("headers") = colHeader
    reg("fields").Add f
End Sub

' colHeaders is a "|"-delimited string (NOT comma - a header may itself contain
' commas, e.g. "Type map (T,N,D,...)"). Up to 3 columns are read (C, D, E).
Public Sub Core_AddTable(reg As Object, ByVal key As String, _
                         ByVal label As String, ByVal colHeaders As String, _
                         Optional ByVal tier As String = "simple", _
                         Optional ByVal help As String = "", _
                         Optional ByVal usedBy As String = "")
    Dim f As Object
    Set f = NewField(key, label, "table", "", tier, help, usedBy)
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

' category is the DISPLAY grouping used by the generated index and by the column
' headers on a button page. It defaults to the package's display name, so an
' action that declares nothing keeps today's grouping. It is deliberately
' decoupled from the module: two packages may feed one category, and one package
' may span several, so a future package split never moves a button in the user's
' mental model.
Public Sub Core_AddAction(reg As Object, ByVal actionId As String, ByVal caption As String, _
                          ByVal entryPoint As String, ByVal help As String, _
                          Optional ByVal category As String = "")
    Dim a As Object, cat As String
    Set a = CreateObject("Scripting.Dictionary")
    a.Add "id", actionId
    a.Add "caption", caption
    a.Add "entry", entryPoint
    a.Add "help", help
    cat = Trim(category)
    If Len(cat) = 0 Then cat = CStr(reg("name"))
    a.Add "category", cat
    reg("actions").Add a
End Sub

Private Function NewField(ByVal key As String, ByVal label As String, _
                          ByVal kind As String, ByVal dflt As String, _
                          ByVal tier As String, ByVal help As String, _
                          ByVal usedBy As String) As Object
    Dim f As Object
    Set f = CreateObject("Scripting.Dictionary")
    f.Add "key", key
    f.Add "label", label
    f.Add "kind", kind
    f.Add "default", dflt
    f.Add "headers", ""
    f.Add "defaults", New Collection
    f.Add "help", help
    ' Comma-separated action ids this field configures; "" = package-wide (always
    ' shown). Stored lowercased so FieldInUse can compare without re-normalizing.
    f.Add "usedBy", LCase(Trim(usedBy))
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

' Does this field belong to the user's working area, given the availability map?
' True when the field is package-wide (no usedBy) or when at least ONE action it
' names is not Off - a "Chains only" action still runs, so it still needs its
' settings. False means "relegate to the collapsed band" - never "omit":
' Core_BuildSetup deletes and re-seeds Setup from the FINALIZED _Config, so
' omitting a field would cost toggle -> Finalize -> rebuild -> fill -> Finalize
' and would discard anything typed but not yet finalized (design section 2.3).
' An unknown id counts as in-use, so a typo in usedBy cannot hide a field forever.
Private Function FieldInUse(f As Object, availMap As Object) As Boolean
    Dim ids() As String, i As Long, one As String
    FieldInUse = True
    If f Is Nothing Then Exit Function
    If Not f.Exists("usedBy") Then Exit Function
    If Len(Trim(CStr(f("usedBy")))) = 0 Then Exit Function
    ids = Split(CStr(f("usedBy")), ",")
    For i = LBound(ids) To UBound(ids)
        one = LCase(Trim(ids(i)))
        If Len(one) > 0 Then
            If Not availMap.Exists(one) Then Exit Function     ' unknown id -> in use
            If IsRunnable(availMap, one) Then Exit Function    ' one runnable is enough
        End If
    Next i
    FieldInUse = False
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
    c.Add "Fixture"                 ' TEMPORARY (v0.4.0 verification) - remove with PkgFixture
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

    ' Defence in depth: Finalize blocks a chain that references an Off action, so
    ' this should be unreachable through a validated config. It exists because
    ' "switched off but it ran anyway" is the one failure mode of this feature
    ' that a finance user cannot recover from. Placed under the error trap (not
    ' above it): ReadAvailMap does real worksheet I/O (FindKeyRow/NextKeyRow walk
    ' the sheet with .End(xlUp) and read cells), and a protected or corrupt
    ' _Config must land on StepFail like any other failure here, not escape the
    ' procedure. Resolves the owner early (instead of at the un-gated lookup
    ' below) purely so the "Off" result carries the real caption, not the raw id -
    ' Core_RunChain's Continue/Stop prompt is exactly the dialog this gate exists
    ' to protect, and it must not show a finance user a dotted id.
    If Not IsRunnable(ReadAvailMap(CfgSheet()), actionId) Then
        Set reg = FindActionOwner(actionId, entry)
        res("found") = True
        If Not reg Is Nothing Then res("label") = ActionCaption(reg, actionId)
        res("ok") = False
        res("errNumber") = 0
        res("errText") = "Action '" & actionId & "' is switched Off on the Setup tab."
        On Error GoTo 0
        Set RunAction = res
        Exit Function
    End If

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
    res("found") = True                ' a real error, not "action not installed"
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

' The single public dispatch entry. Resolves an action id or a reserved chain.* id.
' A DIRECT invocation of anything not available as a button is refused here: with
' the button gone the only ways to arrive are a stale shape or Alt+F8, and neither
' should be able to run something the user switched off. Chain steps do not pass
' through here - Core_RunChain calls RunAction, which independently refuses Off.
Public Sub Core_Run(ByVal actionId As String)
    Dim res As Object, availMap As Object, t As String
    If IsChainId(actionId) Then Core_RunChain actionId: Exit Sub

    ' Reading the availability map is a worksheet read (FindKeyRow/NextKeyRow walk
    ' the sheet). Core_Run is a button's OnAction with no other error handler in
    ' the call chain above it, so an unexpected failure here (a protected or
    ' corrupt _Config) must fail CLOSED - refuse to run and say why - rather than
    ' surface a raw VBA runtime-error dialog and break into the IDE in front of a
    ' finance user.
    On Error GoTo AvailReadFail
    Set availMap = ReadAvailMap(CfgSheet())
    t = AvailFor(availMap, actionId)
    On Error GoTo 0

    If t <> AVAIL_BUTTON Then
        If t = AVAIL_CHAIN Then
            MsgBox "'" & actionId & "' is set to 'Chains only' on the Setup tab, so it " & _
                   "runs only as a step inside a Chain." & vbCrLf & vbCrLf & _
                   "Change its Availability to 'Button' to run it on its own.", _
                   vbInformation, "Not available on its own"
        Else
            MsgBox "'" & actionId & "' is switched Off on the Setup tab." & vbCrLf & vbCrLf & _
                   "Change its Availability to run it.", vbExclamation, "Action is off"
        End If
        Exit Sub
    End If
    Set res = RunAction(actionId, False)
    If Not res("found") And res("errNumber") = 0 Then
        MsgBox "Action not installed: " & actionId, vbExclamation, "Core"
    ElseIf Not res("ok") And Not res("reported") Then
        MsgBox "Action '" & actionId & "' failed:" & vbCrLf & vbCrLf & _
               "#" & res("errNumber") & "  " & res("errText"), vbExclamation, "Macro Error"
    End If
    Exit Sub

AvailReadFail:
    MsgBox "Could not confirm '" & actionId & "' is switched on, so it was not run." & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Rebuild the Setup tab if the problem continues.", vbExclamation, "Core"
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

' Clears the Action catalog block's cells in _Config so a rebuild's prior-value
' reads (ReadCodeMap/ReadAvailMap) find nothing for the cleared rows, instead
' of carrying an unreadable value forward unchanged. idsToClear Is Nothing
' clears the WHOLE block (marker, every id, every Availability/Page-code
' cell) - nothing in it can be trusted. A Dictionary clears only the
' Page-code cell of a row whose lower-cased id is a key in it; every other
' row is untouched - used when exactly which values are bad is known.
Private Sub ClearCatalogBlock(cfgWs As Worksheet, idsToClear As Object)
    Dim cKr As Long, cNk As Long, ci As Long, cid As String
    If cfgWs Is Nothing Then Exit Sub
    cKr = FindKeyRow(cfgWs, LAYOUT_KEY)
    If cKr = 0 Then Exit Sub
    cNk = NextKeyRow(cfgWs, cKr)
    If idsToClear Is Nothing Then
        cfgWs.Range(cfgWs.Cells(cKr, RID_COL), cfgWs.Cells(cNk - 1, RID_COL)).ClearContents
        cfgWs.Range(cfgWs.Cells(cKr, VAL_COL + 1), cfgWs.Cells(cNk - 1, VAL_COL + 2)).ClearContents
    Else
        For ci = cKr + 1 To cNk - 1
            cid = LCase(Trim(CStr(cfgWs.Cells(ci, RID_COL).Value)))
            If Len(cid) > 0 Then
                If idsToClear.Exists(cid) Then cfgWs.Cells(ci, VAL_COL + 1).ClearContents
            End If
        Next ci
    End If
End Sub

' The one guarded-rebuild confirmation, shared by "invalid" and "damaged" -
' they used to differ only in clear scope and wording, and that duplication is
' exactly where divergent disclosures crept in. The Off-gating disclosure is
' UNCONDITIONAL, not just for a full clear: a targeted clear still touches
' VAL_COL + 1, which IS the Availability column in a genuine v1.2+ block, so a
' doubly-tampered store (marker missing AND the header renamed, landing on
' "invalid" instead of "damaged") could still have live Off state in the very
' cells being cleared - "invalid never touches Availability" was a
' prediction, not a fact, so the disclosure no longer depends on it.
' Returns True if the user confirmed (clear, and any Setup-tab delete, already
' done); False if they declined (nothing touched). On True with the tab
' having existed, THIS function deletes it, so Core_BuildSetup's own "tab
' already exists, rebuild it?" prompt never fires twice over one choice.
Private Function ConfirmAndClear(ByVal problem As String, ByVal discarded As String, _
                                  idsToClear As Object) As Boolean
    Dim setupPresent As Boolean, msg As String
    setupPresent = Core_SheetExists(SETUP_SHEET)
    msg = problem & vbCrLf & vbCrLf & _
        "The only way forward is to rebuild the Action catalog." & vbCrLf & vbCrLf & _
        "DISCARDED: " & discarded & vbCrLf & _
        "This takes effect the moment you click Yes, BEFORE the rebuild finishes: any " & _
        "action you had switched Off will be able to run again immediately, and will stay " & _
        "that way until you switch it Off again and click Finalize Setup Tab." & vbCrLf & _
        "KEPT: every other saved setting in this workbook."
    If setupPresent Then
        msg = msg & vbCrLf & vbCrLf & _
            "The current Setup tab will also be rebuilt, since its saved answers cannot be " & _
            "verified either."
    End If
    msg = msg & vbCrLf & vbCrLf & "Rebuild the Setup tab now?"
    If MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, _
              "Create Setup Tab - saved configuration unreadable") = vbNo Then
        MsgBox "Nothing has been changed.", vbInformation, "Create Setup Tab"
        ConfirmAndClear = False
        Exit Function
    End If
    ClearCatalogBlock CfgSheet(), idsToClear
    If setupPresent Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(SETUP_SHEET).Delete
        Application.DisplayAlerts = True
    End If
    ConfirmAndClear = True
End Function

' The fail-closed config migration gate (contract section 5), extracted whole
' out of Core_BuildSetup - a renderer has no business owning the recovery
' policy. True = safe to proceed (no problem, or the user just confirmed a
' guarded rebuild - clear and any Setup-tab delete already done). False = the
' user has already been told why nothing can proceed, or declined; workbook
' untouched either way. schemaState's meaning is ConfigSchemaStatus's to
' document; the one thing worth repeating here is WHY "invalid" and "damaged"
' get different treatment even though both are unreadable: "invalid"'s bad
' value is a VISIBLE, user-editable Page-code cell that ComposeReport does not
' block on, so fixing it on the Setup tab genuinely works when the tab exists.
' "damaged"'s bad value is the HIDDEN marker cell itself, which the user
' cannot see or edit, and which - since _Config is a straight ws.Copy of the
' Setup tab - the live Setup tab almost always carries too, tripping
' ComposeReport's own check. "Fix it on the Setup tab" is never a real
' instruction for "damaged", so it always gets the guarded rebuild instead,
' tab present or not.
Private Function GateConfigMigration() As Boolean
    Dim schemaState As String, schemaDetail As String, schemaBadIds As Object
    schemaState = ConfigSchemaStatus(schemaDetail, schemaBadIds)

    Select Case schemaState
        Case "none", "ok", "legacy"
            GateConfigMigration = True

        Case "unknown"
            MsgBox schemaDetail & vbCrLf & vbCrLf & _
                   "To carry on, import that version's MacrosBundle.bas again. Your settings " & _
                   "are safe in the meantime - nothing has been changed.", _
                   vbExclamation, "Create Setup Tab"
            GateConfigMigration = False

        Case "invalid"
            If Core_SheetExists(SETUP_SHEET) Then
                MsgBox schemaDetail & vbCrLf & vbCrLf & _
                       "The Setup tab is still open with your other answers on it. Correct " & _
                       "the page code(s) named above there, then click 'Finalize Setup Tab'. " & _
                       "Nothing has been changed." & vbCrLf & vbCrLf & _
                       "This is the faster fix. If you delete the Setup tab instead, 'Create " & _
                       "Setup Tab' can still recover: it will clear just the unreadable " & _
                       "code(s) and rebuild around them, keeping everything else - but " & _
                       "correcting the cell(s) here directly avoids that extra step.", _
                       vbExclamation, "Create Setup Tab"
                GateConfigMigration = False
            Else
                GateConfigMigration = ConfirmAndClear(schemaDetail, _
                    "the page code(s) named above (those actions return to automatic " & _
                    "placement); every other saved page code is kept.", schemaBadIds)
            End If

        Case "damaged"
            GateConfigMigration = ConfirmAndClear(schemaDetail, _
                "every saved page-code and Availability choice in the Action catalog (every " & _
                "action returns as a plain Button with automatic placement).", Nothing)
    End Select
End Function

Public Sub Core_BuildSetup()
    Dim ws As Worksheet, reg As Object, f As Object, r As Long, nPkgs As Long
    Dim firstFieldRow As Long, availMap As Object, bandFirst As Long, bandLast As Long
    On Error GoTo Fail

    If Not GateConfigMigration() Then Exit Sub

    If Core_SheetExists(SETUP_SHEET) Then
        ' Fix wave: "never discard un-finalized typing" is a design principle
        ' (ConfirmAndClear already honours it with an explicit DISCARDED/KEPT
        ' breakdown and vbDefaultButton2). This prompt deletes the sheet on Yes,
        ' so anything typed since the last Finalize is lost - say so plainly and
        ' default to No instead of Yes.
        If MsgBox("A '" & SETUP_SHEET & "' tab already exists. Rebuild it?" & vbCrLf & vbCrLf & _
                  "DISCARDED: anything you have typed on the Setup tab since the last 'Finalize " & _
                  "Setup Tab' - the rebuild re-seeds from the last ACTIVATED settings, not from " & _
                  "what is currently on screen." & vbCrLf & _
                  "KEPT: every setting from your last Finalize.", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Create Setup Tab") = vbNo Then Exit Sub
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
    ws.Cells(r, 2).Value = "Fill in the YELLOW cells, then click  'Finalize Setup Tab'  on the Setup Menu tab."
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

    ' The availability map comes from the ACTIVATED config, so a rebuild reflects
    ' the last Finalize. A fresh workbook has no _Config, so the map is empty and
    ' AvailFor reports Button for everything - the v0.4.0 default.
    Set availMap = ReadAvailMap(CfgSheet())

    firstFieldRow = r
    Dim banner As String
    For Each reg In BuildRegistry()
        nPkgs = nPkgs + 1
        banner = UCase(CStr(reg("name")))
        r = SectionHeader(ws, r, banner)
        r = RenderTier(ws, r, reg, "simple", True, banner, availMap, True)   ' simple fields, with their sections
        If HasAdvanced(reg, availMap) Then
            r = SectionHeader(ws, r, "ADVANCED  -  most users can leave these as-is")
            r = RenderTier(ws, r, reg, "advanced", False, banner, availMap, True)  ' advanced, collected flat
        End If
        r = r + 1
    Next reg

    r = RenderComposeSection(ws, r)

    ' The off band goes BELOW Compose on purpose: the action catalog is where the
    ' user turns things on and off, so it must never sit under a collapsed section.
    r = r + 1
    r = RenderOffBand(ws, r, availMap, bandFirst, bandLast)

    ' Sentinel key in column A. The config readers bound a list/table block by
    ' scanning column A for the next non-blank cell; without a trailing key the
    ' LAST list/table on the sheet has no lower bound (End(xlUp) stops at its own
    ' key row, whose value rows are blank in column A) and its values are dropped.
    ' It stays last, after the off band, because the off band may contain the
    ' sheet's final list/table block.
    ws.Cells(r, KEY_COL).Value = END_KEY

    ' Long labels in column B are set to wrap (below); auto-size the field rows
    ' so a wrapped label is fully visible instead of being clipped by the value
    ' cell in column C.
    ws.Rows(firstFieldRow & ":" & r).AutoFit

    ' Group AFTER AutoFit: AutoFit sizes rows and would otherwise fight the
    ' collapse, and a collapsed row's height is irrelevant anyway.
    CollapseRows ws, bandFirst, bandLast

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

' Render every field of the given tier whose in-use state matches wantInUse. When
' showSections is True, a section header is emitted just before its first matching
' field (so a section whose fields are all in the other tier is not drawn). When
' False (the advanced pass and the off-band pass), section headers are skipped and
' the fields render flat under the one band.
Private Function RenderTier(ws As Worksheet, ByVal r As Long, reg As Object, _
                            ByVal tier As String, ByVal showSections As Boolean, _
                            ByVal bannerText As String, availMap As Object, _
                            ByVal wantInUse As Boolean) As Long
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
            If FieldInUse(f, availMap) = wantInUse Then
                If showSections And havePending Then
                    r = SectionHeader(ws, r, pending)
                    havePending = False
                End If
                r = RenderField(ws, r, reg, f)
            End If
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

' Is there an advanced field that belongs in the working area? An advanced field
' whose actions are all Off goes to the off band instead, so drawing the ADVANCED
' header for it would leave an empty band.
Private Function HasAdvanced(reg As Object, availMap As Object) As Boolean
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If CStr(f("tier")) = "advanced" Then
                If FieldInUse(f, availMap) Then
                    HasAdvanced = True
                    Exit Function
                End If
            End If
        End If
    Next f
End Function

' Does this package have ANY non-section field on the requested side of the
' in-use split? Used to skip both an empty ADVANCED band and an empty off band.
Private Function HasAnyTier(reg As Object, availMap As Object, ByVal wantInUse As Boolean) As Boolean
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If FieldInUse(f, availMap) = wantInUse Then
                HasAnyTier = True
                Exit Function
            End If
        End If
    Next f
End Function

' Every field whose usedBy names only Off actions, from every package, collected
' under one band at the very bottom of the field area and collapsed with Excel's
' row outline.
'
' They are RENDERED, never omitted. DO NOT "improve" this to skip them:
' Core_BuildSetup deletes and re-seeds Setup from the FINALIZED _Config, so
' omit-unless-enabled costs toggle -> Finalize -> rebuild -> fill -> Finalize and
' discards anything typed but not yet finalized (design section 2.3, which
' deliberately overrides all three panelists of an earlier review round).
'
' Returns the next free row; reports the collapsible span through bandFirst /
' bandLast so the caller can group AFTER its AutoFit pass. bandFirst = 0 means
' nothing was relegated and no band was drawn.
Private Function RenderOffBand(ws As Worksheet, ByVal r As Long, availMap As Object, _
                               ByRef bandFirst As Long, ByRef bandLast As Long) As Long
    ' anyOff, NOT "any": Any is a VBA reserved word (it appears in Declare ... As
    ' Any), so "Dim any As Boolean" is a compile error the VBE reports against this
    ' procedure's declaration line. Neither build.ps1's structural checks nor the
    ' syntax parser can see it - it was found by the owner's Debug > Compile.
    Dim reg As Object, anyOff As Boolean
    bandFirst = 0
    bandLast = 0
    RenderOffBand = r
    For Each reg In BuildRegistry()
        If HasAnyTier(reg, availMap, False) Then anyOff = True
    Next reg
    If Not anyOff Then Exit Function           ' nothing off: draw no band at all

    r = SectionHeader(ws, r, "SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON")
    r = TintRow(ws, r, "These belong to actions set to 'Off' in the action catalog above. " & _
                "They are kept and still saved, so nothing you typed is lost - set the " & _
                "action back to 'Button' and they move up into the main list. Click the + " & _
                "in the left margin to open this section.", _
                9, False, True, RGB(90, 70, 120))
    bandFirst = r
    For Each reg In BuildRegistry()
        If HasAnyTier(reg, availMap, False) Then
            r = TintRow(ws, r, CStr(reg("name")), 10, True, False, RGB(90, 70, 120))
            r = RenderTier(ws, r, reg, "simple", False, UCase(CStr(reg("name"))), availMap, False)
            r = RenderTier(ws, r, reg, "advanced", False, UCase(CStr(reg("name"))), availMap, False)
        End If
    Next reg
    bandLast = r - 1
    RenderOffBand = r
End Function

' Group + collapse the off band, with the band's own header row left visible above
' it (SummaryRow = xlAbove). Degrades to a plain VISIBLE band if outlining is
' refused - the fields are rendered and readable by Finalize either way, because
' the config readers key off column A and never care whether a row is hidden.
Private Sub CollapseRows(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    If firstRow < 1 Or lastRow < firstRow Then Exit Sub
    On Error Resume Next
    ws.Outline.SummaryRow = xlAbove
    ws.Rows(firstRow & ":" & lastRow).Group
    ws.Outline.ShowLevels RowLevels:=1
    On Error GoTo 0
End Sub

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
'   Constants CHAINS_KEY / LAYOUT_KEY / RID_COL are declared in the module
'   declarations section at the top (VBA requires module-level Const there;
'   a Const between procedures is a compile error).
'================================================================
Private Function RenderComposeSection(ws As Worksheet, ByVal r As Long) As Long
    r = SectionHeader(ws, r, "COMPOSE YOUR BUTTONS")
    r = TintRow(ws, r, "Build multi-step buttons (Chains) and arrange buttons " & _
                "onto pages (Page code). Leave every code blank for the automatic, " & _
                "category-packed layout.", _
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

' Caption for an action id (falls back to the id itself). Shared registry lookup
' factored out of ActionTokenFor so a caption-only consumer (StepsInline) doesn't
' have to strip the " [id]" suffix back off.
Private Function ActionCaptionFor(ByVal actionId As String) As String
    Dim reg As Object, a As Object
    ActionCaptionFor = actionId
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If LCase(CStr(a("id"))) = LCase(actionId) Then
                ActionCaptionFor = CStr(a("caption"))
                Exit Function
            End If
        Next a
    Next reg
End Function

' "Import Source Data [importexport.import]" for an action id (falls back to the id).
' Own registry lookup (not ActionCaptionFor + actionId) so the bracketed id is
' the registry's canonical a("id"), not whatever casing/spelling the caller
' passed in - a hand-typed non-canonical id in a Steps cell is normalised back
' to canonical when the Setup tab is rebuilt (fix round 2: a prior refactor
' bracketed the caller's raw actionId instead, losing that normalisation).
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

' "Caption [id]" for a HIDDEN or UNASSIGNED entry in ComposeReport, matching
' the format the OFF/CHAINS ONLY availability lists just above them already
' use ("caption [id]", built from the registry) - fix round 1, fix 5: the
' hidden list used to print the bare lowercased map key, e.g.
' "importexport.suggesttypes", while everything around it names the button
' too; fix round 2, fix 2 closed the same gap in the unassigned list right
' below it, so one report no longer mixes two formats. ActionTokenFor alone
' only knows registry actions; a hidden/unassigned CHAIN (chain.*) is not in
' the registry, so it is resolved against the chains collection ComposeReport
' already read, falling back to "id [id]" only if neither lookup finds it
' (should not happen for a live id, but matches ActionTokenFor's own fallback
' shape).
Private Function HiddenTokenFor(ByVal id As String, chains As Collection) As String
    Dim ch As Object
    If IsChainId(id) Then
        HiddenTokenFor = id & " [" & id & "]"
        For Each ch In chains
            If LCase(CStr(ch("id"))) = LCase(id) Then
                HiddenTokenFor = CStr(ch("label")) & " [" & CStr(ch("id")) & "]"
                Exit Function
            End If
        Next ch
    Else
        HiddenTokenFor = ActionTokenFor(id)
    End If
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
    Dim base As String, ch As String, i As Long, cand As String, n As Long
    For i = 1 To Len(label)
        ch = LCase(Mid(label, i, 1))
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            base = base & ch
        ElseIf Len(base) > 0 And Right$(base, 1) <> "_" Then
            base = base & "_"
        End If
    Next i
    base = Trim(base)                     ' loop above never forms "__"; strip any trailing "_" below
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
    Dim existing As Collection, ch As Object, i As Long
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

' Parse an Availability cell to one of the three tokens.
'   ""      -> AVAIL_BUTTON  (never set; a pre-v0.4.0 config reads as today)
'   invalid -> ""            (the CALLER must block Finalize)
' It must never fall back to Button on an unrecognised string: a user who typed
' "of" meaning "off" would otherwise be handed a live destructive button.
' Yes/No are accepted as synonyms so a user upgrading from a two-value habit, or
' one who types the obvious thing, is understood rather than blocked.
Private Function ParseAvail(ByVal v As String) As String
    Dim s As String
    s = LCase(Trim(v))
    If Len(s) = 0 Then ParseAvail = AVAIL_BUTTON: Exit Function
    Select Case s
        Case AVAIL_BUTTON, "yes", "y", "on", "true"
            ParseAvail = AVAIL_BUTTON
        Case AVAIL_CHAIN, "chain only", "chains", "chain", "chains-only"
            ParseAvail = AVAIL_CHAIN
        Case AVAIL_OFF, "no", "n", "false"
            ParseAvail = AVAIL_OFF
        Case Else
            ParseAvail = ""                  ' unrecognised - caller must block
    End Select
End Function

' The display label written into the cell for a token.
Private Function AvailLabel(ByVal token As String) As String
    Select Case token
        Case AVAIL_CHAIN: AvailLabel = "Chains only"
        Case AVAIL_OFF:   AvailLabel = "Off"
        Case Else:        AvailLabel = "Button"
    End Select
End Function

' id(lower) -> availability token, from the action catalog block. An id with no
' row is simply absent, and AvailFor treats absence as Button - so a v0.3.0
' _Config (which has no Availability column at all) reads as everything on.
' An unrecognised cell is stored as "" so ComposeReport can block on it.
' Column layout comes from MarkerVerdict, not a decision made here - see its
' comment for why anything but "absent" reads the v1.2+ columns.
Private Function ReadAvailMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String, detail As String
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadAvailMap = map
    If ws Is Nothing Then Exit Function
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr = 0 Then Exit Function
    If MarkerVerdict(ws, kr, detail) = "absent" Then Exit Function   ' pre-v1.2 block: all Button
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(id) > 0 Then map(LCase(id)) = ParseAvail(CStr(ws.Cells(i, VAL_COL + 1).Value))
    Next i
End Function

' Button unless the map says otherwise. Absence means Button, deliberately: an
' unconfigured or pre-v1.2 deployment must show every button.
Private Function AvailFor(availMap As Object, ByVal id As String) As String
    AvailFor = AVAIL_BUTTON
    If availMap Is Nothing Then Exit Function
    If availMap.Exists(LCase(id)) Then AvailFor = CStr(availMap(LCase(id)))
End Function

' Fail closed on both: an invalid token ("") draws no button and does not run.
Private Function HasButton(availMap As Object, ByVal id As String) As Boolean
    HasButton = (AvailFor(availMap, id) = AVAIL_BUTTON)
End Function

Private Function IsRunnable(availMap As Object, ByVal id As String) As Boolean
    Dim t As String
    t = AvailFor(availMap, id)
    IsRunnable = (t = AVAIL_BUTTON Or t = AVAIL_CHAIN)
End Function

' A validated dropdown removes free-typing from a three-value column. Wrapped in
' On Error Resume Next because validation can be refused (a protected sheet, some
' locales) and a refused dropdown must not abort the build - ParseAvail still
' guards the value, and Finalize still blocks on anything unrecognised.
Private Sub AvailDropdown(cell As Range)
    On Error Resume Next
    cell.Validation.Delete
    cell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, Formula1:="Button,Chains only,Off"
    cell.Validation.IgnoreBlank = True
    cell.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub

' One row per registry action: hidden id in RID_COL, "caption [id]" (grey) in C, a
' yellow Availability cell with a validated dropdown in D, a yellow Page-code cell
' in E. Prior answers are preserved from CfgSheet() by id. Unavailable actions STAY
' listed here - that is the owner's hard requirement: turning a button off must
' never make the action undiscoverable.
Private Function RenderPlacementBlock(ws As Worksheet, ByVal r As Long) As Long
    Dim prior As Object, priorAvail As Object, reg As Object, a As Object, id As String
    ' Fix round 3, fix 3: "hide" used to read as manual-layout-only ("leave one
    ' button off A MANUAL LAYOUT") - true only before fix round 3's fix 1, which
    ' made GatherButtons drop a hide-coded button before either filer runs, so
    ' it now works the same whether or not any placement code is set.
    r = BlockIntro(ws, r, "Action catalog  -  what each action is available as, and where it goes", _
                   "Availability: 'Button' shows a button; 'Chains only' hides the button but a " & _
                   "Chain can still run it; 'Off' hides it and stops it running anywhere. " & _
                   "Page code: leave every code blank for the automatic layout; fill in even one " & _
                   "placement code and codes take over - number = page, letters = column (e.g. 1A). " & _
                   "Type  hide  to leave any one button off, whether or not you use placement codes.")
    ws.Cells(r, KEY_COL).Value = LAYOUT_KEY
    ws.Cells(r, VAL_COL).Value = "Action"
    ws.Cells(r, VAL_COL + 1).Value = AVAIL_HEADER
    ws.Cells(r, VAL_COL + 2).Value = "Page code"
    ws.Range(ws.Cells(r, VAL_COL), ws.Cells(r, VAL_COL + 2)).Font.Bold = True
    ' The machine-owned schema marker, in the hidden id column of the key row -
    ' NOT in the visible header. See Task 4 / contract section 5. Force Text so a
    ' General-formatted cell cannot coerce the marker to a number (belt-and-braces:
    ' SCHEMA_VER's "v" prefix already blocks numeric coercion on its own).
    ws.Cells(r, RID_COL).NumberFormat = "@"
    ws.Cells(r, RID_COL).Value = SCHEMA_VER
    r = r + 1

    Set prior = ReadCodeMap(CfgSheet())            ' preserve prior codes by id
    Set priorAvail = ReadAvailMap(CfgSheet())      ' preserve prior availability by id
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            id = CStr(a("id"))
            ws.Cells(r, VAL_COL).Value = CStr(a("caption")) & " [" & id & "]"
            ws.Cells(r, VAL_COL).Font.Color = RGB(90, 90, 90)
            ws.Cells(r, RID_COL).Value = id
            ws.Cells(r, VAL_COL + 1).Value = AvailLabel(AvailFor(priorAvail, id))
            If prior.Exists(LCase(id)) Then ws.Cells(r, VAL_COL + 2).Value = prior(LCase(id))
            Yellow ws.Cells(r, VAL_COL + 1)
            Yellow ws.Cells(r, VAL_COL + 2)
            AvailDropdown ws.Cells(r, VAL_COL + 1)
            r = r + 1
        Next a
    Next reg
    RenderPlacementBlock = r
End Function

' The config-format marker for the action catalog block, or "" when absent
' (meaning the block predates v1.2.0 and its page code is one column to the left).
' Deliberately NOT read from the visible header: RID_COL is hidden, is already how
' this codebase carries identity, and is copied into _Config with everything else.
Private Function BlockSchema(ws As Worksheet, ByVal keyRow As Long) As String
    If ws Is Nothing Or keyRow < 1 Then Exit Function
    BlockSchema = Trim(CStr(ws.Cells(keyRow, RID_COL).Value))
End Function

' Extracts the major version number from a schema marker of the form
' "v<major>.<minor>" (e.g. "v1.2" -> 1), and from a three-(or more-)part form
' like "v1.2.0" - spec/contract.md section 10 spells its own versions that way,
' so a future SCHEMA_VER bump is likely to look like it. ALSO accepts a bare
' "<major>.<minor>" with no "v" - the marker the interim da5de80 build wrote,
' before the "v" prefix (and the NumberFormat "@" guard added in 8b00a79)
' existed: a General-formatted cell coerced that bare text into the NUMBER 1.2,
' which reads back through CStr as "1.2" on an en-US locale. That data is
' genuinely v1.2-shaped - only its stamp is missing the "v" - so it is
' deliberately understood here rather than rejected as damaged.
' Returns False for anything that fits none of these: a present marker that is
' not a version number at all is corruption, and callers must not guess a
' major version for it (contract section 5 - fail closed, do not guess).
' HARDENED: guards CLng overflow on the major digit run, same bound as ParseCode.
Private Function SchemaMajor(ByVal marker As String, ByRef major As Long) As Boolean
    Dim s As String, dot As Long, digits As String, rest As String, i As Long, ch As String
    Dim prevDot As Boolean
    major = 0
    s = Trim(marker)
    If Len(s) = 0 Then Exit Function
    If LCase(Left(s, 1)) = "v" Then s = Mid(s, 2)      ' normal form; bare form skips this
    dot = InStr(s, ".")
    If dot < 2 Then Exit Function                      ' need >=1 digit before the first dot
    digits = Left(s, dot - 1)
    rest = Mid(s, dot + 1)                              ' minor, or "minor.patch[.more]"
    If Len(rest) = 0 Then Exit Function
    For i = 1 To Len(digits)
        ch = Mid(digits, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    If Len(digits) > 7 Then Exit Function               ' overflow guard (CLng-safe range)
    ' rest must be one or more digit groups separated by single dots ("2",
    ' "2.0", "2.0.1", ...) - anything else (letters, "..", a leading/trailing
    ' dot) is not a version number.
    prevDot = True                                       ' rest must not START with "."
    For i = 1 To Len(rest)
        ch = Mid(rest, i, 1)
        If ch = "." Then
            If prevDot Then Exit Function                ' leading "." or ".."
            prevDot = True
        ElseIf ch >= "0" And ch <= "9" Then
            prevDot = False
        Else
            Exit Function
        End If
    Next i
    If prevDot Then Exit Function                        ' rest must not END with "."
    major = CLng(digits)
    SchemaMajor = True
End Function

' The parsed major version of THIS BUILD's own SCHEMA_VER constant. SCHEMA_VER
' is a hardcoded literal controlled entirely by this codebase, so a parse
' failure here is a bug in this build, not a workbook problem. Raising instead
' of silently returning 0 matters: every caller compares a saved marker's major
' against this one, and a silent 0 would misclassify every real v1.x store as
' "unknown" and lock every workbook out with no recovery - exactly the failure
' mode this task exists to prevent, just self-inflicted instead of triggered by
' a stale workbook. Callers already run under On Error GoTo Fail, so this
' surfaces as an unmistakable "... error: SCHEMA_VER does not parse..." message
' instead of a wrong, silent comparison.
Private Function CurrentSchemaMajor() As Long
    Dim m As Long
    If Not SchemaMajor(SCHEMA_VER, m) Then
        Err.Raise vbObjectError + 1024, "CurrentSchemaMajor", _
            "SCHEMA_VER ('" & SCHEMA_VER & "') does not parse as a version number. " & _
            "This is a bug in this build, not a problem with the workbook."
    End If
    CurrentSchemaMajor = m
End Function

' THE single place that decides what an action-catalog marker means. Every
' caller that used to make its own marker decision (ReadAvailMap, ReadCodeMap,
' ConfigSchemaStatus, ComposeReport) now calls this instead of re-deriving the
' same verdict, so Finalize and a rebuild can never disagree about the same
' marker - there is exactly one place left that decides. Returns exactly one of:
'   "absent"      - no marker at all: a genuine pre-v1.2 block (contract
'                    section 5, "absent is not invalid"). The Availability
'                    column does not exist, and the code lives one column to
'                    the left of where a v1.2+ block keeps it.
'   "same-major"  - marker present, same major version as SCHEMA_VER (whatever
'                    the minor): the columns are exactly where a v1.2+ block
'                    keeps them, so it is fully readable.
'   "other-major" - marker present, parses, but a genuinely different major
'                    version: a real different build wrote this.
'   "unparseable" - marker present but does not parse as a version at all:
'                    corruption, not version skew.
'   "tampered"    - marker ABSENT, but the visible header at VAL_COL + 1
'                    already reads "Availability" - that column only exists in
'                    a v1.2+ block, so the marker was removed or edited by
'                    hand rather than never having existed.
' For a reader (ReadAvailMap/ReadCodeMap), only "absent" vs "anything else"
' matters: "absent" means legacy columns, every other verdict means the v1.2+
' column layout is where this block actually put its data - even "tampered"
' and "unparseable", because the DATA is still very likely in the v1.2+
' columns; only the marker cell itself is untrustworthy. The gates
' (ConfigSchemaStatus/ComposeReport) are the ones that care about the full
' verdict and refuse to ACTIVATE on anything but "absent"/"same-major".
' detail is filled with a plain factual explanation for every verdict except
' "absent"/"same-major" - callers append their own call to action.
Private Function MarkerVerdict(ws As Worksheet, ByVal keyRow As Long, ByRef detail As String) As String
    Dim marker As String, mMajor As Long
    detail = ""
    If ws Is Nothing Or keyRow < 1 Then MarkerVerdict = "absent": Exit Function
    marker = BlockSchema(ws, keyRow)
    If Len(marker) = 0 Then
        If LCase(Trim(CStr(ws.Cells(keyRow, VAL_COL + 1).Value))) = LCase(AVAIL_HEADER) Then
            detail = "The hidden format marker for the Action catalog block is missing, even " & _
                     "though its layout looks like this version's, so saved Availability " & _
                     "choices cannot be trusted."
            MarkerVerdict = "tampered"
        Else
            MarkerVerdict = "absent"
        End If
        Exit Function
    End If
    If Not SchemaMajor(marker, mMajor) Then
        detail = "This workbook's saved configuration format stamp ('" & marker & "') does not " & _
                 "look like a version number, so it cannot be trusted."
        MarkerVerdict = "unparseable"
        Exit Function
    End If
    If mMajor = CurrentSchemaMajor() Then
        MarkerVerdict = "same-major"
    Else
        detail = "This workbook's saved configuration was written by a different version of " & _
                 "this toolkit (format '" & marker & "'; this build understands '" & SCHEMA_VER & "')."
        MarkerVerdict = "other-major"
    End If
End Function

' Classify the ACTIVATED config before anything destructive touches the Setup
' sheet - MarkerVerdict's verdicts, translated to this module's action
' vocabulary, plus the one thing MarkerVerdict cannot tell on its own: whether
' an "absent" marker's recovered legacy page codes actually parse. Returns:
'   "none"    - no config yet (a fresh workbook); nothing to lose
'   "ok"      - MarkerVerdict "same-major": readable as-is
'   "legacy"  - MarkerVerdict "absent", every recovered page code validates
'   "invalid" - MarkerVerdict "absent", but a recovered code does not parse
'   "unknown" - MarkerVerdict "other-major"
'   "damaged" - MarkerVerdict "unparseable" or "tampered"
' detail is filled with a human-readable explanation for every status except
' "none"/"ok"/"legacy". badIds is ALWAYS set to a (possibly empty) Dictionary;
' for "invalid" it holds the lower-cased ids whose recovered code failed
' ParseCode - the same set detail lists in prose - so a caller that wants to
' clear ONLY the unreadable rows (not the whole block) has a single source of
' truth instead of re-parsing detail's text.
Private Function ConfigSchemaStatus(ByRef detail As String, ByRef badIds As Object) As String
    Dim ws As Worksheet, kr As Long, nk As Long, i As Long
    Dim id As String, code As String, tn As Long, cx As Long, cl As String
    Dim verdict As String
    detail = ""
    Set badIds = CreateObject("Scripting.Dictionary")
    Set ws = CfgSheet()
    If ws Is Nothing Then ConfigSchemaStatus = "none": Exit Function
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr = 0 Then ConfigSchemaStatus = "none": Exit Function

    verdict = MarkerVerdict(ws, kr, detail)
    Select Case verdict
        Case "same-major"
            ConfigSchemaStatus = "ok"
        Case "other-major"
            ConfigSchemaStatus = "unknown"
        Case "unparseable", "tampered"
            ConfigSchemaStatus = "damaged"
        Case "absent"
            ' Legacy: validate every page code recovered from the OLD column
            ' before letting anything overwrite the sheet holding them.
            nk = NextKeyRow(ws, kr)
            For i = kr + 1 To nk - 1
                id = Trim(CStr(ws.Cells(i, RID_COL).Value))
                If Len(id) > 0 Then
                    code = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
                    If Len(code) > 0 And LCase(code) <> "hide" Then
                        If Not ParseCode(code, tn, cx, cl) Then
                            detail = detail & "  - " & id & ": '" & code & "'" & vbCrLf
                            badIds(LCase(id)) = True
                        End If
                    End If
                End If
            Next i
            If Len(detail) > 0 Then
                detail = "These saved page codes could not be read from the older " & _
                         "configuration format:" & vbCrLf & detail
                ConfigSchemaStatus = "invalid"
            Else
                ConfigSchemaStatus = "legacy"
            End If
    End Select
End Function

' id(lower) -> code, merging the action catalog block and the chains block.
' v0.4.0 inserted Availability between Action and Page code, so a pre-v1.2
' block keeps its code one column to the left. Column layout comes from
' MarkerVerdict, not a decision made here - see its comment for why anything
' but "absent" reads the v1.2+ columns. Callers that could DESTROY the
' recovered values (Core_BuildSetup) must consult ConfigSchemaStatus first;
' this reader is permissive on purpose so the report can describe what it found.
Private Function ReadCodeMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String, codeCol As Long
    Dim chains As Collection, ch As Object, detail As String
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadCodeMap = map
    If ws Is Nothing Then Exit Function
    ' actions
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr > 0 Then
        If MarkerVerdict(ws, kr, detail) = "absent" Then
            codeCol = VAL_COL + 1                ' legacy (pre-v1.2) layout
        Else
            codeCol = VAL_COL + 2
        End If
        nk = NextKeyRow(ws, kr)
        For i = kr + 1 To nk - 1
            id = Trim(CStr(ws.Cells(i, RID_COL).Value))
            If Len(id) > 0 Then map(LCase(id)) = Trim(CStr(ws.Cells(i, codeCol).Value))
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
    Dim availMap As Object, offList As String, badAvail As String
    Dim chainOnly As String, nOff As Long, t As String
    Dim layoutKR As Long, verdict As String, vDetail As String
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
            hidden = hidden & "  - " & HiddenTokenFor(CStr(k), chains) & vbCrLf
        ElseIf Len(code) = 0 Then
            ' Fix round 2, fix 2: was the bare lowercased key - same
            ' inconsistency fix round 1's fix 5 removed from the HIDDEN list
            ' just above, now closed here too so both lists in one report
            ' agree on "caption [id]".
            unassigned = unassigned & "  - " & HiddenTokenFor(CStr(k), chains) & vbCrLf
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

    ' Availability. An unrecognised cell and an Off action used by a chain are both
    ' STRUCTURAL - they block activation. "Chains only" is reported for confidence,
    ' not as a problem: it is an authored state, not an accident.
    '
    ' Format-marker integrity, contract section 5's fail-closed requirement -
    ' delegated entirely to MarkerVerdict (see its comment) so this can never
    ' disagree with what a rebuild would decide about the same marker. Only
    ' "same-major"/"absent" pass; every other verdict blocks Finalize. This
    ' function's job is to report the problem and name the button that fixes
    ' it - NOT to predict what that rebuild will or will not preserve, which
    ' belongs to the rebuild's own confirmation (the one place that actually
    ' knows, and the one the user acts on).
    layoutKR = FindKeyRow(ws, LAYOUT_KEY)
    verdict = MarkerVerdict(ws, layoutKR, vDetail)
    ' "other-major" and "unparseable"/"tampered" share nothing but a header verb -
    ' their real recoveries are opposite. GateConfigMigration REFUSES on
    ' "other-major" ("unknown") and tells the user to re-import that version's
    ' bundle; it lets "unparseable"/"tampered" ("damaged") through to
    ' ConfirmAndClear's guarded rebuild. Naming "Create Setup Tab" for
    ' other-major would send the user to a button that then refuses them.
    Select Case verdict
        Case "other-major"
            s = s & vbCrLf & "ACTION CATALOG FORMAT MISMATCH: " & vDetail & " Nothing has been " & _
                "activated. Re-import that version's MacrosBundle.bas to continue - this build " & _
                "will not rebuild over a newer saved format." & vbCrLf
            structural = structural + 1
        Case "unparseable"
            s = s & vbCrLf & "ACTION CATALOG FORMAT MISMATCH: " & vDetail & " Nothing has been " & _
                "activated. Click 'Create Setup Tab' to rebuild the Action catalog block." & vbCrLf
            structural = structural + 1
        Case "tampered"
            s = s & vbCrLf & "ACTION CATALOG FORMAT MARKER MISSING: " & vDetail & " Nothing has " & _
                "been activated. Click 'Create Setup Tab' to rebuild the Action catalog block." & vbCrLf
            structural = structural + 1
    End Select
    Set availMap = ReadAvailMap(ws)
    For Each regX In BuildRegistry()
        For Each aX In regX("actions")
            t = AvailFor(availMap, CStr(aX("id")))
            If Len(t) = 0 Then
                badAvail = badAvail & "  - " & CStr(aX("id")) & vbCrLf
                structural = structural + 1
            ElseIf t = AVAIL_OFF Then
                nOff = nOff + 1
                offList = offList & "  - " & CStr(aX("caption")) & " [" & CStr(aX("id")) & "]" & vbCrLf
            ElseIf t = AVAIL_CHAIN Then
                chainOnly = chainOnly & "  - " & CStr(aX("caption")) & " [" & CStr(aX("id")) & "]" & vbCrLf
            End If
        Next aX
    Next regX
    If Len(badAvail) > 0 Then
        s = s & vbCrLf & "UNRECOGNISED Availability (use Button, Chains only, or Off):" & _
            vbCrLf & badAvail
    End If
    For Each ch In chains
        For Each stepId In ch("steps")
            If ids.Exists(LCase(CStr(stepId))) Then
                If AvailFor(availMap, CStr(stepId)) = AVAIL_OFF Then
                    s = s & "OFF ACTION IN CHAIN: '" & CStr(ch("label")) & "' uses " & _
                        CStr(stepId) & ", which is switched Off." & vbCrLf
                    structural = structural + 1
                End If
            End If
        Next stepId
    Next ch
    If nOff > 0 Then
        ' Deliberately does NOT claim the settings have already collapsed: that
        ' only happens for a field that declares usedBy (neither shipped package
        ' does yet), and only on the NEXT "Create Setup Tab" - Finalize never
        ' redraws Setup. Overclaiming here would describe a fixture-only feature
        ' as if it already applied to a real deployment.
        s = s & vbCrLf & "OFF (" & nOff & " action(s); no button, cannot run. Any field that " & _
            "declares this action in its usedBy will collapse into the settings band the next " & _
            "time 'Create Setup Tab' rebuilds Setup):" & vbCrLf & offList
    End If
    If Len(chainOnly) > 0 Then
        s = s & vbCrLf & "CHAINS ONLY (no button; runs only inside a Chain):" & vbCrLf & chainOnly
    End If

    ' Fix round 3, fix 2 (live path - this runs from Core_FinalizeSetup, before
    ' the user decides to activate): was AnyCode(map), which counts "hide" as a
    ' layout-activating code - the exact bug fix round 2's fix 5 exists to
    ' prevent, just at a different call site. A hide-only store would have shown
    ' "A layout is ACTIVE" here while the engine actually runs auto-layout.
    ' AnyPlacementCode agrees with what Core_BuildControlPanel's
    ' BuildAutoLayout/BuildManualLayout choice (Task 8) actually does. Both
    ' messages describe v0.4.0's real behaviour, not the pre-auto-layout
    ' "single Macros tab" default this replaced.
    If AnyPlacementCode(map) Then
        ' Fix wave: codes choose where a button STARTS, not a guarantee of the
        ' page/column it ends up on - contract section 6's Column-cap overflow
        ' clause can still relocate it once a page or column is full.
        s = s & vbCrLf & "A layout is ACTIVE - your placement codes choose where each button starts, " & _
            "but a full page or column can still move it to a continuation one." & vbCrLf
        If Len(unassigned) > 0 Then s = s & "Unassigned -> page 1's Unassigned column (or a continuation page, once the page has no free column slot left):" & vbCrLf & unassigned
    Else
        s = s & vbCrLf & "No placement codes -> categories will be packed onto numbered pages automatically." & vbCrLf
    End If
    ' "hide" is honoured in EITHER mode now (fix 1), so this is no longer
    ' conditional on AnyPlacementCode.
    If Len(hidden) > 0 Then s = s & "Hidden (not placed, in either mode):" & vbCrLf & hidden
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

' Fix round 2, fix 5: answers "is there a REAL placement code", not "is there
' any non-blank cell at all" - the latter would be True for a SINGLE "hide"
' entry with every other code left blank, and a user who hides one button
' would then flip the whole workbook into manual mode (every other,
' still-uncoded button reads as manually placed, landing in page 1's
' Unassigned column, and HomeLegend would tell them their page codes control
' a layout they never authored). "hide" omits a button; it does not place
' one, so it must not count as a placement here. Core_BuildControlPanel (Task
' 8) uses THIS helper, and only this one, both to choose BuildAutoLayout vs
' BuildManualLayout and to set BuildHomePage's isManual.
Private Function AnyPlacementCode(map As Object) As Boolean
    Dim k As Variant, v As String
    For Each k In map.Keys
        v = Trim(CStr(map(k)))
        If Len(v) > 0 And LCase(v) <> "hide" Then AnyPlacementCode = True: Exit Function
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
        ' Scoped deliberately: this only covers the scalar sheet/file field checks
        ' just above. The compose/placement report appended below can still add
        ' its own "problem(s) above must be fixed" line beneath this - the two
        ' must never be allowed to read as one verdict.
        msg = "Sheet and file checks passed."
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
    ' Remove any prior _Config. It is stored xlSheetVeryHidden, and a very-hidden
    ' sheet can resist .Delete (the delete then silently fails under On Error
    ' Resume Next, _Config survives, and the rename below collides -> "That name is
    ' already taken", orphaning a "Setup (2)" copy on every retry). Unhiding it
    ' first makes the delete reliable.
    Dim old As Worksheet
    Set old = Nothing
    On Error Resume Next
    Set old = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo Fail
    If Not old Is Nothing Then
        old.Visible = xlSheetVisible
        old.Delete
    End If
    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set cfg = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    cfg.name = CONFIG_SHEET
    ws.Activate                        ' a very-hidden sheet cannot be the active sheet
    cfg.Visible = xlSheetVeryHidden
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Core_BuildControlPanel True        ' rebuild the Macros pages so chains/placement appear now
    ' Fix round 1, fix 2: Core_BuildControlPanel cannot raise past its own Fail:
    ' handler (it shows its own error and returns normally), so without this
    ' check the message below would unconditionally claim success even when the
    ' rebuild just failed - right after the panels were already deleted, i.e.
    ' on a workbook with no buttons at all. mLastBuildOk is set True only once
    ' the whole rebuild pipeline finishes; it is reset False at the top of
    ' every Core_BuildControlPanel call, so it can never be stale here.
    If mLastBuildOk Then
        MsgBox "Configuration activated." & vbCrLf & vbCrLf & _
               "The Macros pages and 'Macros Home' index have been rebuilt to match your Setup answers.", _
               vbInformation, "Finalize Setup Tab"
    Else
        MsgBox "Configuration activated, but rebuilding the Macros pages failed - see the error " & _
               "shown just now." & vbCrLf & vbCrLf & _
               "Your Setup answers were saved. Click 'Finalize Setup Tab' again to retry the rebuild.", _
               vbExclamation, "Finalize Setup Tab"
    End If
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Core_FinalizeSetup error: " & Err.Description, vbExclamation, "Finalize Setup Tab"
End Sub

'================================================================
' 7. CONTROL PANEL  (builds the Macros pages + Macros Home index)
'================================================================
' Auto-layout by display category when no button carries a placement code; the
' manual code layout when any button does ("hide" alone does not count - see
' AnyPlacementCode). Prior "Macros N" panels, and the pre-v0.4.0 single
' "Macros" tab, are cleared before either mode draws. Home is always built,
' even for a single page, so the workbook has one shape at every scale.
' silent:=True suppresses the "built N ..." confirmation (used when Finalize
' rebuilds the panel itself and shows its own combined message).
Public Sub Core_BuildControlPanel(Optional ByVal silent As Boolean = False)
    Dim buttons As Collection, codeMap As Object, layout As Object, manual As Boolean
    mLastBuildOk = False    ' reset FIRST, every call - see the flag's own declaration comment
    On Error GoTo Fail
    Core_SpeedOn
    Set codeMap = ReadCodeMap(CfgSheet())   ' read ONCE: GatherButtons' hide filter and the
                                             ' mode decision below must see the same map
    manual = AnyPlacementCode(codeMap)      ' computed BEFORE GatherButtons - IsHideCoded reads
                                             ' codeMap(k) only after confirming .Exists(k), so
                                             ' this ordering no longer matters for correctness,
                                             ' but keeping it first documents the dependency
    Set buttons = GatherButtons(codeMap)    ' single funnel for both filers - also where "hide"
                                             ' is dropped, mode-independently (see GatherButtons)
    Set layout = CreateObject("Scripting.Dictionary")   ' neither filer creates it
    DeletePanels                            ' remove any prior "Macros N" pages (spares
                                             ' "Macros Home" by design - see IsPanelName)
    DeleteLegacyPanel                       ' v0.4.0: the single "Macros" tab is retired in BOTH modes
    If manual Then
        BuildManualLayout buttons, codeMap, layout
    Else
        BuildAutoLayout buttons, layout
    End If
    RenderAllPages layout
    BuildAdminPanel                         ' Create/Finalize controls on their own hideable tab
    BuildHomePage layout, manual            ' LAST: activates Macros Home and selects A1
    mLastBuildOk = True     ' the whole pipeline finished - only set here, never in Fail:
    Core_SpeedOff
    If Not silent Then _
        MsgBox "Built " & layout.Count & " macro page(s) with " & buttons.Count & _
               " button(s), plus the '" & HOME_SHEET & "' index.", _
               vbInformation, "Setup Complete"
    Exit Sub
Fail:
    RestoreSpeed
    ' Matches Core_FinalizeSetup's own wording for the same failure (its only
    ' caller) - there is no "re-running the build" control, and on a virgin
    ' build a failure here leaves no 'Setup Menu' tab either, so naming the real
    ' recovery (Finalize again) instead of a nonexistent one.
    MsgBox "Core_BuildControlPanel error: " & Err.Description & vbCrLf & vbCrLf & _
           "Your settings are safe. Click 'Finalize Setup Tab' again to retry the rebuild.", _
           vbExclamation, "Setup"
End Sub

' Drop the pre-v0.4.0 single "Macros" tab. Called unconditionally by
' Core_BuildControlPanel now (both modes supersede it), so this is also how an
' upgrading workbook's stale "Macros" tab is removed the first time it rebuilds
' under this version.
Private Sub DeleteLegacyPanel()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(MACRO_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' Each item: Array(id, caption, onAction, help, category). Actions from the
' registry, chains from activated config. This is the SINGLE funnel for both, so
' the availability filter lives here and nowhere else. Admin buttons are never in
' this list - they live on the Setup Menu tab. Takes codeMap from the caller
' (Task 8) rather than reading its own copy, so Core_BuildControlPanel's mode
' decision (AnyPlacementCode) and this hide filter always see the same map.
' Fix round 3, fix 1 (BLOCKING, live path): also the single funnel for the
' "hide" placement code, and for the same reason - BuildAutoLayout never sees
' codeMap at all (it only distributes by category), so a "hide"-coded button
' would otherwise be filtered ONLY inside BuildManualLayout. Fix round 2's fix
' 5 correctly stopped a lone "hide" from flipping a workbook into manual mode
' (AnyPlacementCode ignores it) - but that means a hide-only store routes to
' the AUTO filer, and the auto filer has no "hide" concept of its own, so the
' hidden button would be drawn, catalogued, and counted. Filtering it here,
' before either filer ever sees the button, makes "hide" mode-independent,
' matching contract section 6 ("hide omits a button") without qualification.
' Feed both filers ONLY this collection - nothing assembled another way - or
' "hide" silently stops working.
Private Function GatherButtons(codeMap As Object) As Collection
    Dim out As New Collection, reg As Object, a As Object, chains As Collection, ch As Object
    Dim onAct As String, availMap As Object
    Set availMap = ReadAvailMap(CfgSheet())
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If HasButton(availMap, CStr(a("id"))) And Not IsHideCoded(codeMap, CStr(a("id"))) Then
                onAct = "'Core_Run " & Chr(34) & CStr(a("id")) & Chr(34) & "'"
                out.Add Array(CStr(a("id")), CStr(a("caption")), onAct, _
                              CStr(a("help")), CStr(a("category")))
            End If
        Next a
    Next reg
    Set chains = ReadChains(CfgSheet())
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 And Not IsHideCoded(codeMap, CStr(ch("id"))) Then
            onAct = "'Core_Run " & Chr(34) & CStr(ch("id")) & Chr(34) & "'"
            out.Add Array(CStr(ch("id")), CStr(ch("label")), onAct, _
                          "Chain: " & StepsInline(ch("steps")), CHAIN_CAT)
        End If
    Next ch
    Set GatherButtons = out
End Function

' Test .Exists BEFORE reading the item. Scripting.Dictionary's default Item
' property CREATES a missing key (with an Empty value) on read, and VBA's And
' is not short-circuiting, so "codeMap.Exists(k) And LCase(codeMap(k)) = ..."
' evaluates codeMap(k) even when .Exists(k) is False, silently inserting a
' phantom key. Harmless when codeMap was a throwaway local, but Task 8 shares
' ONE codeMap between Core_BuildControlPanel's mode decision and every
' GatherButtons/filer call that follows it, so a phantom key inserted here
' would corrupt that shared map for everything reading it afterwards. The If
' below only reads codeMap(k) once .Exists(k) is already known True.
Private Function IsHideCoded(codeMap As Object, ByVal id As String) As Boolean
    Dim k As String
    k = LCase(id)
    If Not codeMap.Exists(k) Then Exit Function
    IsHideCoded = LCase(Trim(CStr(codeMap(k)))) = "hide"
End Function

' Chain help text only - fix round 1: was raw action ids ("importexport.updateMaster"),
' which are long and unbroken and blow past a card's width in a few steps. Captions
' read as "Clear Sheets > Import Data > Update Master" for a finance user. The
' Setup-tab authoring format (StepsToTokens, "Label [id]") is untouched.
Private Function StepsInline(steps As Collection) As String
    Dim s As String, v As Variant
    For Each v In steps
        s = s & IIf(Len(s) > 0, " > ", "") & ActionCaptionFor(CStr(v))
    Next v
    StepsInline = s
End Function

' Create/Finalize controls on their own tab, so they can be hidden when not in
' use and never clutter the action pages. Reused (not re-created) on rebuild, so
' a user who hides this tab keeps it hidden. Does not activate the sheet.
Private Sub BuildAdminPanel()
    Dim ws As Worksheet, leftPos As Double, topPos As Double
    Set ws = EnsurePanelSheet(ADMIN_SHEET, False)
    With ws.Range("B2")
        .Value = "Setup Menu"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(31, 78, 121)
    End With
    ' Fix round 2, fix 1: same reasoning as Macros Home's B2 - EnsurePanelSheet's
    ' row-height reset pins row 2 at StandardHeight, which does not grow for a
    ' 16pt title, so set it literally right after the title is written.
    ' Correction (Task 9): 21pt gave a 16pt font ~1pt of headroom on paper but a
    ' reviewer's arithmetic showed that is too tight in practice; bumped to 23
    ' for real headroom - a slightly tall title row costs nothing visually.
    ws.Rows(2).RowHeight = 23
    With ws.Range("B3")
        ' Fix wave: this is the ONLY sheet-surface route to Create/Finalize now
        ' that the single "Macros" tab is retired, and Excel raises "Reference
        ' isn't valid" on a hyperlink to a hidden sheet (Macros Home's own link
        ' to this tab). Hiding stays fine - it's approved behaviour - but the
        ' invitation must say how to get back.
        .Value = "First-time and re-configuration controls. Right-click this tab -> Hide when you're " & _
                 "done - to come back, right-click any visible tab -> Unhide, or press Alt+F8."
        .Font.Italic = True
    End With
    ws.Rows("5:12").RowHeight = 15
    leftPos = ws.Range("B5").Left
    topPos = ws.Range("B5").Top
    AddBtn ws, "btnCreateSetup", "Create Setup Tab", "Core_BuildSetup", leftPos, topPos, 220, 32
    With ws.Range(ws.Cells(5, 3), ws.Cells(6, 3))
        .Merge: .Value = "Builds a pre-filled 'Setup' tab from the installed packages."
        .Font.Size = 10: .WrapText = True: .VerticalAlignment = xlCenter
    End With
    topPos = topPos + 45
    AddBtn ws, "btnFinalizeSetup", "Finalize Setup Tab", "Core_FinalizeSetup", leftPos, topPos, 220, 32
    With ws.Range(ws.Cells(8, 3), ws.Cells(9, 3))
        .Merge: .Value = "Validates your Setup answers, activates them, and rebuilds the Macros pages and 'Macros Home' index."
        .Font.Size = 10: .WrapText = True: .VerticalAlignment = xlCenter
    End With
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
    ' Fix round 1, fix 9: Cells.Clear should already drop hyperlinks, but that
    ' behaviour has historically been version-flaky in Excel, and Macros Home
    ' accumulates one hyperlink per catalogued button - cheap insurance before a
    ' human imports this.
    ws.Hyperlinks.Delete
    ' Fix round 1, fix 3: Cells.Clear does not touch row height, so a row a
    ' PRIOR build's AutoFit made tall (a long description, a long chain
    ' recipe) stays that tall even after every value on it is cleared.
    ' Nothing downstream resets it wholesale - RenderPage only sets the button
    ' rows it draws and AutoFits only the description rows it uses; HomeCatalog
    ' AutoFits only the rows it writes - so a page or Home that shrinks on
    ' rebuild (fewer buttons, shorter descriptions) keeps the taller rows left
    ' behind by the bigger build. Reset to the sheet's standard height here,
    ' before anything else sizes a specific row, so every subsequent
    ' RowHeight/AutoFit call in this build starts from a clean sheet.
    ws.Cells.RowHeight = ws.StandardHeight
    ws.Tab.Color = RGB(31, 78, 121)
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 47
    ws.Columns("C").ColumnWidth = 85
    Set EnsurePanelSheet = ws
End Function

' The layout structure every panel builder fills and RenderPage consumes:
'   layout(tabNum)("title")        page title             (String)
'   layout(tabNum)("cols")(colIx)  buttons in that column (Collection of Arrays)
'   layout(tabNum)("hdrs")(colIx)  that column's header   (String)
' A page is created lazily by the first button filed onto it, so callers must file
' a button BEFORE setting its title or column header.
Private Sub AddToPage(layout As Object, ByVal tabNum As Long, ByVal colIx As Long, ByVal btn As Variant)
    Dim page As Object, cols As Object, items As Collection
    If Not layout.Exists(tabNum) Then
        Set page = CreateObject("Scripting.Dictionary")
        page.Add "title", PANEL_BASE & "- Page " & tabNum
        page.Add "cols", CreateObject("Scripting.Dictionary")
        page.Add "hdrs", CreateObject("Scripting.Dictionary")
        layout.Add tabNum, page
    End If
    Set page = layout(tabNum)
    Set cols = page("cols")
    If Not cols.Exists(colIx) Then cols.Add colIx, New Collection
    Set items = cols(colIx)
    items.Add btn                 ' btn = Array(id, caption, onAction, help, category)
End Sub

' No-ops when the page does not exist yet - file a button first (see AddToPage).
Private Sub SetPageTitle(layout As Object, ByVal tabNum As Long, ByVal title As String)
    If layout.Exists(tabNum) Then layout(tabNum)("title") = title
End Sub

Private Sub SetColHeader(layout As Object, ByVal tabNum As Long, ByVal colIx As Long, ByVal hdr As String)
    If layout.Exists(tabNum) Then layout(tabNum)("hdrs")(colIx) = hdr
End Sub

Private Function PageButtonCount(page As Object) As Long
    Dim cols As Object, k As Variant
    Set cols = page("cols")
    For Each k In cols.Keys
        PageButtonCount = PageButtonCount + cols(k).Count
    Next k
End Function

' B/D/F carry content, C/E are gutters. Column A stays narrow as a margin.
Private Sub SetupCardGrid(ws As Worksheet)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = CARD_W
    ws.Columns("C").ColumnWidth = GUTTER_W
    ws.Columns("D").ColumnWidth = CARD_W
    ws.Columns("E").ColumnWidth = GUTTER_W
    ws.Columns("F").ColumnWidth = CARD_W
End Sub

' Card content columns are B, D, F -> 2, 4, 6. slot is 1-based; 0 means off-grid.
Private Function CardCol(ByVal slot As Long) As Long
    If slot < 1 Or slot > CARD_COLS Then Exit Function
    CardCol = slot * 2
End Function

' Row of a card pair's button, 1-based pair index within a column. The SOLE
' place this arithmetic is written. RenderPage's loop and FitDescRows' start
' row both derive from this function (start row = CardBtnRow(1) + 1) instead
' of each independently re-deriving CARD_HDR_ROW's offset, so the two cannot
' drift apart if CARD_HDR_ROW - or this formula - ever changes (fix round 2:
' the previous pairing held only because CARD_HDR_ROW happened to be odd).
Private Function CardBtnRow(ByVal pairIx As Long) As Long
    CardBtnRow = CARD_HDR_ROW + 1 + (pairIx - 1) * 2
End Function

' Anchor a button to a cell so it lines up with the description directly beneath.
Private Sub AddBtnInCell(ws As Worksheet, ByVal nm As String, ByVal cap As String, _
                         ByVal action As String, anchor As Range)
    AddBtn ws, nm, cap, action, anchor.Left, anchor.Top + 1, anchor.Width, BTN_H
End Sub

' AutoFit each description row, then floor it. AutoFit works here ONLY because
' nothing on the card grid is merged. Floor only, no ceiling (fix round 1,
' contract 6): the floor keeps a one-line card from looking broken next to a
' three-line neighbour; a ceiling would clip, which is exactly what contract 6
' forbids for a description.
Private Sub FitDescRows(ws As Worksheet, ByVal firstDescRow As Long, ByVal lastRow As Long)
    Dim r As Long
    For r = firstDescRow To lastRow Step 2
        ws.Rows(r).AutoFit
        If ws.Rows(r).RowHeight < DESC_MIN_H Then ws.Rows(r).RowHeight = DESC_MIN_H
    Next r
End Sub

' Draw one "Macros N" page as an unmerged card grid: nav row, page title, one
' category-headed card column per layout column, each button followed by its
' description in the cell below it.
Private Function RenderPage(layout As Object, ByVal tabNum As Long, _
                            ByVal prevSheet As String, ByVal nextSheet As String) As Worksheet
    Dim panel As Worksheet, page As Object, cols As Object, hdrs As Object
    Dim colKeys As Variant, si As Long, slot As Long, colIx As Long
    Dim items As Collection, ii As Long, it As Variant
    Dim cc As Long, bRow As Long, dRow As Long, lastRow As Long, hdr As String, skipped As Long
    Set page = layout(tabNum)
    Set cols = page("cols")
    Set hdrs = page("hdrs")
    Set panel = EnsurePanelSheet(PANEL_BASE & tabNum, False)
    SetupCardGrid panel
    RenderNavRow panel, prevSheet, nextSheet
    ' fix round 1: rows 2 and 4 used to be blank chrome spacers; removed, so the
    ' title now sits directly under the nav row and the category header directly
    ' under the title (CARD_HDR_ROW moved from 5 to 3 to match).
    With panel.Cells(CARD_HDR_ROW - 1, 2)
        .Value = CStr(page("title"))
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With

    colKeys = SortedKeys(cols)
    lastRow = CARD_HDR_ROW     ' honest starting point: nothing drawn past the header yet
    slot = 0
    For si = LBound(colKeys) To UBound(colKeys)
        colIx = CLng(colKeys(si))
        slot = slot + 1
        cc = CardCol(slot)
        ' Fix round 1, fix 9: both filers now cap every page at CARD_COLS columns
        ' by construction (BuildAutoLayout's slotsUsed check; BuildManualLayout's
        ' CapManualColumns), so this branch is unreachable-by-invariant - but a
        ' silent discard here is exactly how the original defect happened, so
        ' while that invariant is fresh, count what would be skipped and surface
        ' it on the page itself rather than trust the invariant silently forever.
        If cc = 0 Then
            skipped = skipped + 1
            GoTo NextSlot        ' more columns than card slots: never draw off-grid
        End If
        Set items = cols(colIx)
        hdr = ""
        If hdrs.Exists(colIx) Then hdr = CStr(hdrs(colIx))
        If Len(hdr) > 0 Then
            With panel.Cells(CARD_HDR_ROW, cc)
                .Value = hdr
                .Font.Bold = True
                .Font.Size = 11
                .Font.Color = RGB(54, 32, 100)
                .Interior.Color = RGB(221, 217, 240)
                .WrapText = True    ' fix round 1: a long category used to spill into the gutter
            End With
        End If
        For ii = 1 To items.Count
            it = items(ii)
            bRow = CardBtnRow(ii)
            dRow = bRow + 1
            panel.Rows(bRow).RowHeight = BTN_ROW_H
            AddBtnInCell panel, "b" & tabNum & "_" & colIx & "_" & ii, _
                         CStr(it(1)), CStr(it(2)), panel.Cells(bRow, cc)
            With panel.Cells(dRow, cc)
                ' Carry-over C: bounded for a chain (CardHelp), untouched for an
                ' authored action - the full chain recipe stays in HomeCatalog.
                ' Fix round 2, fix 3: gated on the button's id (it(0)), not its
                ' category (it(4)) - see CardHelp's own comment.
                .Value = CardHelp(CStr(it(3)), CStr(it(0)))
                .Font.Size = 9
                .Font.Color = RGB(110, 110, 110)
                .WrapText = True
                .VerticalAlignment = xlTop
                ' fix round 1: a faint bottom rule so a button+description pair
                ' reads as one card instead of a button floating over grey text.
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .Color = RGB(221, 221, 226)
                End With
            End With
            If dRow > lastRow Then lastRow = dRow
        Next ii
NextSlot:
    Next si
    If skipped > 0 Then
        With panel.Cells(CARD_HDR_ROW - 1, 4)
            .Value = "WARNING: " & skipped & " column(s) did not fit on this page and were not drawn. " & _
                     "Both filers are supposed to cap every page at " & CARD_COLS & " columns - this means " & _
                     "that invariant broke. Report it."
            .Font.Bold = True
            .Font.Color = RGB(192, 0, 0)
            .WrapText = True
        End With
    End If
    ' Fix round 2, fix 1: the title (col B) and the skipped-column warning
    ' (col D) share this one physical row, so a single RowHeight governs both.
    ' EnsurePanelSheet's reset (fix round 1, fix 3) pins every row at
    ' StandardHeight, which - unlike Excel's implicit grow-to-fit-font on a
    ' fresh/never-sized row - does NOT grow for a 15pt bold title, so it must
    ' be set explicitly here, same as the header row just below gets an
    ' explicit AutoFit. Literal, not AutoFit: a title is one line at a known
    ' font size, so a literal height is fully deterministic and stays immune
    ' to the AutoFit-under-ScreenUpdating=False risk already flagged for
    ' FitDescRows/HomeCatalog. The warning is unreachable by invariant today
    ' (both filers cap every page at CARD_COLS columns by construction) but
    ' its message text is effectively fixed-length (only skipped's digit
    ' count varies), so a generous literal height is exactly as deterministic
    ' as the title's - sized only after both possibilities are known, which is
    ' why this runs after the skipped check rather than right after the title
    ' is written.
    If skipped > 0 Then
        panel.Rows(CARD_HDR_ROW - 1).RowHeight = 130   ' room for the wrapped warning sentence
    Else
        ' Correction (Task 9): 20 left a 15pt bold title with no real headroom;
        ' bumped to 22 so it reliably clears instead of clipping.
        panel.Rows(CARD_HDR_ROW - 1).RowHeight = 22    ' clears a single 15pt bold title line
    End If
    ' Explicit AutoFit for the header row: it renders correctly on a fresh
    ' sheet only because Excel's implicit auto-height flag is still live.
    ' RenderNavRow already sets row 1's RowHeight explicitly, which clears
    ' that flag; the header row gets the same explicit treatment here so a
    ' two-line category name never depends on an implicit sheet default.
    panel.Rows(CARD_HDR_ROW).AutoFit
    ' Start row derived from CardBtnRow(1), not CARD_HDR_ROW directly - see
    ' CardBtnRow's comment for why.
    FitDescRows panel, CardBtnRow(1) + 1, lastRow
    Set RenderPage = panel
End Function

' Row 1: a persistent Home link in a fixed position on every page, plus Previous /
' Next. Cells rather than shapes, because a cell hyperlink needs no shape
' bookkeeping and EnsurePanelSheet's Cells.Clear removes it cleanly on rebuild.
' ASCII arrows only - build.ps1 writes UTF-8 with no BOM on the assumption that
' the SOURCE is ASCII, so a literal arrow glyph in the .bas is an import hazard.
' Fix round 1, fix 1: a prior draft guarded this with Core_SheetExists(HOME_SHEET)
' and fell back to plain text when the index sheet did not exist yet. That guard
' is actively harmful and is deliberately NOT here - do not re-add it. It fires
' at RENDER time, but BuildHomePage runs LAST in the pipeline (Task 8), so on a
' virgin workbook's first build every page would render before Macros Home
' exists, drawing plain text instead of a link on every page's B1 - while
' BuildHomePage's own B3 tells the user "Every page has a << Home link back
' here." The failure is also self-healing on the SECOND build (IsPanelName spares
' Macros Home from DeletePanels, so it survives and the guard passes next time),
' which makes it invisible to anyone who only tests a rebuild, not a first build.
' It was never needed anyway: Hyperlinks.Add does not validate SubAddress, and
' the target sheet always exists by the time the macro returns and the user can
' click anything - so write the link unconditionally.
Private Sub RenderNavRow(ws As Worksheet, ByVal prevSheet As String, ByVal nextSheet As String)
    NavLink ws.Cells(1, 2), "<< Home", HOME_SHEET
    If Len(prevSheet) > 0 Then NavLink ws.Cells(1, 4), "< Previous", prevSheet
    If Len(nextSheet) > 0 Then NavLink ws.Cells(1, 6), "Next >", nextSheet
    ws.Rows(1).RowHeight = 18
End Sub

' One in-workbook hyperlink cell. The SubAddress sheet name is always quoted,
' because "Macros Home" and "Setup Menu" both contain a space.
Private Sub NavLink(cell As Range, ByVal caption As String, ByVal sheetName As String)
    cell.Parent.Hyperlinks.Add Anchor:=cell, Address:="", _
                               SubAddress:="'" & sheetName & "'!A1", _
                               TextToDisplay:=caption
    cell.Font.Size = 10
    cell.Font.Bold = True
End Sub

' Render every page in page order, threading each one's Previous/Next sheet names
' so the nav row is drawn in the same pass.
Private Sub RenderAllPages(layout As Object)
    Dim tabKeys As Variant, ti As Long, prevNm As String, nextNm As String
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        prevNm = ""
        nextNm = ""
        If ti > LBound(tabKeys) Then prevNm = PANEL_BASE & CLng(tabKeys(ti - 1))
        If ti < UBound(tabKeys) Then nextNm = PANEL_BASE & CLng(tabKeys(ti + 1))
        RenderPage layout, CLng(tabKeys(ti)), prevNm, nextNm
    Next ti
End Sub

' Manual layout: file each button onto a "Macros N" page by its code. Blank code
' -> page 1 Unassigned column (9999); unparseable -> Unassigned (already blocked
' at Finalize). "hide" never reaches here at all (see the fix round 3, fix 1
' comment below) - GatherButtons omits a hide-coded button before either filer
' runs. Files only - rendering is RenderAllPages.
' Fix round 3, fix 1: no longer checks for "hide" here - GatherButtons (the
' live-path funnel both filers are fed from) now drops a hide-coded button
' before it ever reaches either filer, so this Sub can never see one. Removed
' rather than left as a dead branch, so nobody reads it as evidence "hide" is
' still handled per-filer and re-adds a per-filer check that can never fire.
Private Sub BuildManualLayout(buttons As Collection, codeMap As Object, layout As Object)
    Dim e As Variant, id As String, code As String
    Dim tabNum As Long, colIx As Long, colL As String, unassigned As New Collection
    For Each e In buttons
        id = LCase(CStr(e(0)))
        code = ""
        If codeMap.Exists(id) Then code = Trim(CStr(codeMap(id)))
        If Len(code) = 0 Then
            unassigned.Add e
        ElseIf ParseCode(code, tabNum, colIx, colL) Then
            AddToPage layout, tabNum, colIx, e
        Else
            unassigned.Add e            ' unparseable degrades to Unassigned
        End If
    Next e
    For Each e In unassigned
        AddToPage layout, 1, 9999, e     ' blank-coded buttons go to page 1's Unassigned column
    Next e
    ' Fix round 2, fix 6: AUTO_COL_MAX bounds a same-code stack in BuildAutoLayout
    ' but nothing bounded one here - a single "1A" on 25 buttons used to draw one
    ' column 25 deep, running off-screen with no warning, while Macros Home
    ' reported the page as fine. Must run BEFORE CapManualColumns below, so the
    ' horizontal capper sees (and, if needed, relocates) the extra columns this
    ' produces - see that Sub's own comment for why splitting first and letting
    ' the existing horizontal machinery handle any resulting overflow needs no
    ' new page-overflow logic of its own.
    SplitTallManualColumns layout
    ' Carry-over A: a user can type 1A/1B/1C/1D (or fill 3 coded columns and still
    ' land blank-coded buttons in Unassigned/9999 on page 1) - that is 4+ distinct
    ' columns on one page. RenderPage only draws CARD_COLS (3) card slots and
    ' silently drops anything past slot 3, so cap here rather than let the
    ' renderer discard what the user explicitly placed. Must run BEFORE
    ' LabelManualPages below, since it moves columns between pages and
    ' LabelManualPages needs to see the FINAL column layout to title pages and
    ' head columns correctly.
    CapManualColumns layout
    ' Fix round 1, fix 5: BuildAutoLayout has always called its own labeller
    ' (LabelAutoPages) - this filer used to leave LabelManualPages for the caller
    ' to remember separately, which made the two filers asymmetric and left a
    ' silent failure mode (skip the call, get "Macros - Page N" titles that
    ' propagate into the index's category column). Self-label here so Task 8
    ' cannot miss it either way.
    LabelManualPages layout
End Sub

' Title and head a hand-built layout. Column headers keep the code-based names the
' user chose. A page is titled by category only when EVERY button on it agrees - a
' hand-built page may deliberately mix categories, and guessing a title for a
' mixed page would be worse than the generic one.
Private Sub LabelManualPages(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim cols As Object, colKeys As Variant, ci As Long, colIx As Long
    Dim items As Collection, ii As Long, it As Variant
    Dim cat As String, one As String, mixed As Boolean
    Dim origColIx As Long, contTag As String
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        cat = ""
        mixed = False
        colKeys = SortedKeys(cols)
        For ci = LBound(colKeys) To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            ' Fix round 2, fix 6: a colIx >= VSPLIT_MULT is a vertical-overflow
            ' piece SplitTallManualColumns minted for an oversized same-code
            ' stack - decode its original column straight back out of the number
            ' (see that Sub's comment): the piece may since have been relocated
            ' to a different page by CapManualColumns above, and a header set
            ' before that move would not have travelled with it.
            ' Fix round 3, fix 8: colIx also encodes WHICH piece (colIx Mod
            ' VSPLIT_MULT = the pieceIndex SplitTallManualColumns minted it
            ' with), so a stack split into 4+ pieces gets distinctly numbered
            ' headers instead of several identical "(cont.)" columns on one page.
            If colIx >= VSPLIT_MULT Then
                origColIx = colIx \ VSPLIT_MULT
                contTag = " (cont. " & (colIx Mod VSPLIT_MULT) & ")"
            Else
                origColIx = colIx
                contTag = ""
            End If
            If origColIx = 9998 Then
                SetColHeader layout, tabNum, colIx, "Setup" & contTag
            ElseIf origColIx = 9999 Then
                SetColHeader layout, tabNum, colIx, "Unassigned" & contTag
            Else
                SetColHeader layout, tabNum, colIx, "Column " & IndexToCol(origColIx) & contTag
            End If
            Set items = cols(colIx)
            For ii = 1 To items.Count
                it = items(ii)
                one = CStr(it(4))
                If Len(cat) = 0 Then
                    cat = one
                ElseIf LCase(cat) <> LCase(one) Then
                    mixed = True
                End If
            Next ii
        Next ci
        If Not mixed And Len(cat) > 0 Then SetPageTitle layout, tabNum, cat
    Next ti
End Sub

' Fix round 2, fix 6: AUTO_COL_MAX bounds a same-code stack in BuildAutoLayout;
' nothing bounded one here, so a single "1A" on (say) 25 buttons drew one column
' 25 deep, running off-screen with no warning, while Macros Home reported the
' page as fine. Split any column over AUTO_COL_MAX BEFORE CapManualColumns runs
' (BuildManualLayout calls this first): the split pieces are just more columns
' on the page by the time the existing horizontal cap-and-overflow machinery
' looks at it, so a page that gains a 4th+ column this way overflows to a
' continuation page exactly like an explicitly-coded 4th column already does -
' no new page-overflow logic needed, only more columns for the old logic to see.
'
' A split piece's colIx encodes its origin (origColIx * VSPLIT_MULT +
' pieceIndex) rather than getting a header written here, because
' CapManualColumns may relocate a piece to a different page afterwards, and a
' header set on the page it started on would not travel with it (fix round 1's
' fix 8 deliberately removed that carry-across-a-move machinery as dead code).
' LabelManualPages decodes the header straight back out of the colIx instead,
' which survives being moved because AddToPage copies the colIx unchanged.
Private Sub SplitTallManualColumns(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object, cols As Object
    Dim colKeys As Variant, ci As Long, colIx As Long, items As Collection
    Dim needCols As Long, k As Long, n As Long, taken As Long, pieceColIx As Long

    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        colKeys = SortedKeys(cols)      ' snapshot before this page's cols dict grows below
        For ci = LBound(colKeys) To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            Set items = cols(colIx)
            If items.Count <= AUTO_COL_MAX Then GoTo NextCol
            needCols = -Int(-items.Count / AUTO_COL_MAX)     ' ceiling division
            ' Guard against Long overflow for a pathological many-letter code -
            ' leave it as one tall, unsplit column rather than corrupt colIx math.
            ' (2147483647 is VBA Long's max; safe for any realistic 1-3 letter code.)
            If colIx > (2147483647 - needCols) \ VSPLIT_MULT Then GoTo NextCol
            taken = AUTO_COL_MAX                ' first AUTO_COL_MAX items stay at colIx
            For k = 2 To needCols
                pieceColIx = colIx * VSPLIT_MULT + k
                For n = 1 To AUTO_COL_MAX
                    taken = taken + 1
                    If taken > items.Count Then Exit For
                    AddToPage layout, tabNum, pieceColIx, items(taken)
                Next n
            Next k
            ' items IS cols(colIx) (Collections are reference types), so trim the
            ' overflow back out of it now that it has been copied into the pieces
            ' above - AddToPage only ever appends, it never removes.
            Do While items.Count > AUTO_COL_MAX
                items.Remove items.Count
            Loop
NextCol:
        Next ci
    Next ti
End Sub

' Carry-over A (manual side): fold any page's columns beyond CARD_COLS onto
' continuation page(s) appended after the highest page number already in use,
' CARD_COLS columns at a time, instead of letting RenderPage's slot cap (see
' CardCol) silently discard them. Nothing the user typed is lost - the moved
' column keeps its items, and the normal << Home / Next nav still reaches it;
' its header is assigned afterwards by LabelManualPages, not carried across the
' move (see that fix's comment immediately below for why). Only manual layouts
' need this: capacity-packed BuildAutoLayout never files more than CARD_COLS
' columns onto a page by construction (see BuildAutoLayout's slotsUsed check).
' Fix round 2, fix 7: the paragraph above used to claim the moved column "keeps
' ... its header (suffixed "(cont.)")" - stale the moment fix round 1's fix 8
' (immediately below) deleted the code that did that, and exactly the class of
' comment that caused fix round 1's fix 1 (a maintainer reads a comment
' describing code that is not there, and "fixes" it back in). Corrected above,
' not just noted here.
' Fix round 1, fix 8: an earlier draft carried header-preserving machinery here
' (read page("hdrs"), re-tag the moved column "(cont.)", remove it from the old
' page's hdrs) that was dead on arrival: this runs from BuildManualLayout, before
' any header is ever set, so hdrs is always empty at this point and every branch
' of that machinery was unreachable. Deleted rather than fixed-in-place - a moved
' column instead gets a plain, correct header ("Column D", not "Column D
' (cont.)") for free from LabelManualPages, which (fix 5) now always runs
' immediately after this, over the FINAL layout including whatever this moved.
Private Sub CapManualColumns(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim cols As Object, colKeys As Variant, ci As Long, colIx As Long
    Dim maxTab As Long, newTab As Long, chunkPos As Long
    Dim items As Collection, ii As Long

    tabKeys = SortedKeys(layout)
    maxTab = 0
    For ti = LBound(tabKeys) To UBound(tabKeys)
        If CLng(tabKeys(ti)) > maxTab Then maxTab = CLng(tabKeys(ti))
    Next ti

    ' Snapshot the pre-existing pages before appending continuation pages, so the
    ' loop below never revisits (or re-splits) a page it just created.
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        colKeys = SortedKeys(cols)
        If UBound(colKeys) < LBound(colKeys) + CARD_COLS Then GoTo NextTab   ' <= CARD_COLS cols: fine
        chunkPos = CARD_COLS                    ' forces a fresh page on the first overflow column
        For ci = LBound(colKeys) + CARD_COLS To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            If chunkPos >= CARD_COLS Then
                maxTab = maxTab + 1
                newTab = maxTab
                chunkPos = 0
            End If
            chunkPos = chunkPos + 1
            Set items = cols(colIx)
            For ii = 1 To items.Count
                AddToPage layout, newTab, colIx, items(ii)
            Next ii
            cols.Remove colIx
        Next ci
NextTab:
    Next ti
End Sub

' Carry-over C: a chain's help is Core-synthesised ("Chain: " & StepsInline(...),
' GatherButtons) and with MAX_LIST steps can run to ~500 characters - roughly 14
' wrapped lines on a card's width - which AutoFit then applies to the WHOLE row,
' inflating it across all three card columns. Contract 6's no-clipping rule binds
' an action's AUTHORED help; a chain's help is engine-generated, so Core is free
' to bound it here. Cards get the first few step captions plus a "+N more"
' pointer; HomeCatalog (Macros Home) is untouched and always shows the full
' recipe, so the complete step list stays visible somewhere. Only chain buttons
' (category = CHAIN_CAT) are touched - authored action help passes through as-is.
' CHAIN_CARD_STEPS lives in the module's declarations section (top of file), next
' to AUTO_COL_MAX, not here - a Const after the first procedure is a compile error.
' Fix round 2, fix 3: gates on the button's id (IsChainId, "chain.*"), NOT on
' its category string. Category = CHAIN_CAT ("Chains") used to be the gate, but
' CHAIN_CAT is reserved by contract section 7, not enforced in code - Core_AddAction
' accepts any category a package declares, so a package could pass
' category:="Chains" and have its AUTHORED help decorated (and truncated) here,
' which section 6 forbids outright. Chains only ever reach GatherButtons through
' the ReadChains branch, and every chain id is minted "chain.*" (MintChainId),
' so the id namespace is the one thing an ordinary action can never spoof.
Private Function CardHelp(ByVal helpText As String, ByVal btnId As String) As String
    Dim body As String, parts() As String, n As Long, i As Long, shown As String
    Dim usedSteps As Long, candidate As String
    If Not IsChainId(btnId) Then
        CardHelp = helpText
        Exit Function
    End If
    body = helpText
    If Left$(body, 7) = "Chain: " Then body = Mid$(body, 8)
    parts = Split(body, " > ")
    n = UBound(parts) - LBound(parts) + 1
    ' Fix round 1, fix 6: BOTH caps must clear for the text to pass through as-is
    ' - a short step count with long captions still needs bounding (see the
    ' Consts' comment above).
    If n <= CHAIN_CARD_STEPS And Len(helpText) <= CHAIN_CARD_CHARS Then
        CardHelp = helpText
        Exit Function
    End If
    For i = LBound(parts) To UBound(parts)
        If usedSteps >= CHAIN_CARD_STEPS Then Exit For
        candidate = shown & IIf(Len(shown) > 0, " > ", "") & parts(i)
        ' Always keep at least one step, even a lone long one, so the card never
        ' shows zero steps and just a "+N more" tail.
        If usedSteps > 0 And Len(candidate) > CHAIN_CARD_CHARS Then Exit For
        shown = candidate
        usedSteps = usedSteps + 1
    Next i
    ' Fix round 2, fix 2: the fast-path test above can fail (Len(helpText) alone
    ' pushed it over CHAIN_CARD_CHARS because of the "Chain: " prefix) while the
    ' loop still gets through every step before either cap actually bites - e.g.
    ' 4 steps of 27 characters, or 3 of 38. usedSteps = n then, and appending a
    ' "+0 more" tail would hand back a string LONGER than the input while
    ' claiming there is more to see, when there is not. If nothing was actually
    ' left out, return the original text unchanged.
    If usedSteps >= n Then
        CardHelp = helpText
        Exit Function
    End If
    CardHelp = "Chain: " & shown & "  +" & (n - usedSteps) & " more - see Macros Home"
End Function

' Auto-layout: no button carries a page code, so distribute buttons by display
' category. A page is a CAPACITY unit, not a category unit: up to CARD_COLS
' category-headed columns per page, up to AUTO_COL_MAX buttons per column. Small
' categories share a page. A category that does not fit in the remaining columns
' starts a new page rather than straddling one - splitting a category across tabs
' destroys the grouping that putting it in one place buys. A category larger than
' a whole page continues onto further pages. Replaces the pre-v0.4.0 single-
' "Macros"-tab default (contract v1.2.0), which did not survive contact with a
' large catalog.
Private Sub BuildAutoLayout(buttons As Collection, layout As Object)
    Dim cats As Object, seen As New Collection, order As Collection
    Dim e As Variant, cat As String, ci As Long, items As Collection
    Dim needCols As Long, tabNum As Long, slotsUsed As Long
    Dim taken As Long, k As Long, n As Long, firstColOfCat As Boolean

    Set cats = CreateObject("Scripting.Dictionary")
    ' Fix round 1, fix 7: a default Dictionary is BinaryCompare, so "Reconcile"
    ' and "reconcile" from two different packages (an explicitly supported section 6
    ' scenario - two packages feeding one category) would become two adjacent
    ' columns, even though LabelAutoPages de-dupes case-insensitively and would
    ' title the page with the category named once above both. Must be set BEFORE
    ' the first Add - VBA raises an error changing CompareMode on a non-empty
    ' Dictionary.
    cats.CompareMode = vbTextCompare
    For Each e In buttons
        cat = Trim(CStr(e(4)))
        If Len(cat) = 0 Then cat = "Other"       ' defensive: Core_AddAction never emits blank
        If Not cats.Exists(cat) Then
            cats.Add cat, New Collection
            seen.Add cat                         ' first-appearance order, stable
        End If
        cats(cat).Add e
    Next e

    Set order = CategoryOrder(cats, seen)
    tabNum = 1
    slotsUsed = 0
    For ci = 1 To order.Count
        cat = CStr(order(ci))
        Set items = cats(cat)
        needCols = -Int(-items.Count / AUTO_COL_MAX)        ' ceiling division
        ' Start a fresh page rather than straddle. Contract section 6 states the
        ' no-straddle rule unconditionally - it does not carve out an exception
        ' for a category needing more columns than a page has. A category bigger
        ' than a whole page still can't avoid straddling SOME page boundary, but
        ' it should do so at a page it starts fresh on, not one it arrives at
        ' partially filled: the per-k check inside the loop below (slotsUsed >=
        ' CARD_COLS) already forces a new page once the current one fills, so
        ' removing the "needCols <= CARD_COLS" guard here just makes that
        ' fresh-page start happen up front instead of mid-category.
        If slotsUsed > 0 And slotsUsed + needCols > CARD_COLS Then
            tabNum = tabNum + 1
            slotsUsed = 0
        End If
        taken = 0
        firstColOfCat = True
        For k = 1 To needCols
            If slotsUsed >= CARD_COLS Then
                tabNum = tabNum + 1
                slotsUsed = 0
            End If
            slotsUsed = slotsUsed + 1
            For n = 1 To AUTO_COL_MAX
                taken = taken + 1
                If taken > items.Count Then Exit For
                AddToPage layout, tabNum, slotsUsed, items(taken)
            Next n
            If firstColOfCat Then
                SetColHeader layout, tabNum, slotsUsed, cat
                firstColOfCat = False
            Else
                ' Fix round 3, fix 8 (found here too, not just in the manual
                ' split): a category needing exactly CARD_COLS columns (13-18
                ' buttons) fits on one page, so k=2 and k=3 would both be
                ' unsuffixed "(cont.)" - two identical headers on the same
                ' page. k is already this category's 1-based column counter, so
                ' number the header with it instead.
                SetColHeader layout, tabNum, slotsUsed, cat & " (cont. " & k & ")"
            End If
        Next k
    Next ci
    LabelAutoPages layout
End Sub

' First-appearance order, with the reserved Chains category moved to the end: a
' composed button reads as a summary of the actions above it, so it belongs last.
Private Function CategoryOrder(cats As Object, seen As Collection) As Collection
    Dim out As New Collection, i As Long, cat As String
    Set CategoryOrder = out
    For i = 1 To seen.Count
        cat = CStr(seen(i))
        If LCase(cat) <> LCase(CHAIN_CAT) Then out.Add cat
    Next i
    For i = 1 To seen.Count
        cat = CStr(seen(i))
        If LCase(cat) = LCase(CHAIN_CAT) Then out.Add cat
    Next i
End Function

' A packed page may hold up to CARD_COLS categories, so title it by what is
' actually on it: one category -> its name; several -> the distinct names joined.
Private Sub LabelAutoPages(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim hdrs As Object, hk As Variant, seen As Object, title As String, one As String
    Dim contAt As Long
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set hdrs = page("hdrs")
        Set seen = CreateObject("Scripting.Dictionary")
        title = ""
        For Each hk In hdrs.Keys
            one = CStr(hdrs(hk))
            ' Fix round 3, fix 8: a "(cont.)" header is now numbered ("(cont. 2)",
            ' "(cont. 3)", ...), so a literal Replace(" (cont.)", "") no longer
            ' matches - truncate at the fixed " (cont." anchor instead, which is
            ' the same regardless of the number that follows it.
            contAt = InStr(one, " (cont.")
            If contAt > 0 Then one = Left$(one, contAt - 1)
            If Not seen.Exists(LCase(one)) Then
                seen(LCase(one)) = True
                title = title & IIf(Len(title) > 0, "   |   ", "") & one
            End If
        Next hk
        If Len(title) > 70 Then title = Left$(title, 67) & "..."
        If Len(title) > 0 Then SetPageTitle layout, tabNum, title
    Next ti
End Sub

' The generated index - ALWAYS built, even for a single page, so the workbook has
' one shape at every scale (owner's call). Three parts: a page summary, the full
' action catalog, and the legend. Built last and activated, so any rebuild lands
' the user here. Fix round 1, fix 4: totalButtons is no longer a caller-supplied
' parameter - it had undefined semantics (does it count hide-coded buttons? the
' filers drop those before layout is ever built, so the only honest source is
' layout itself), and the natural thing for Task 8 to pass, buttons.Count from
' GatherButtons, would over-count against both the page list above and the
' catalog below. Summing PageButtonCount over every page is the SAME quantity
' the page-summary rows above already show per page, so there is exactly one
' definition of "how many buttons are on this workbook" and no way for a caller
' to get it wrong. isManual (fix 3) reports which filer built this layout, so
' HomeLegend can describe the truth instead of a generic capability.
Private Sub BuildHomePage(layout As Object, ByVal isManual As Boolean)
    Dim ws As Worksheet, r As Long, tabKeys As Variant, ti As Long, tabNum As Long
    Dim page As Object, totalButtons As Long
    Set ws = EnsurePanelSheet(HOME_SHEET, True)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 62
    ws.Columns("E").ColumnWidth = 16

    With ws.Range("B2")
        .Value = "Macro Toolkit"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    ' Fix round 2, fix 1: EnsurePanelSheet's row-height reset pins row 2 at
    ' StandardHeight, which does not grow for an 18pt title the way a
    ' never-explicitly-sized row implicitly would - set literally, right after
    ' the title is written, so it clears the font instead of clipping it.
    ' Correction (Task 9): 23 was possibly already short for an 18pt title
    ' (needs ~23.25); bumped to 25 for real headroom.
    ws.Rows(2).RowHeight = 25
    With ws.Range("B3")
        .Value = "Pick a page below, or find an action in the catalog. Every page has a  << Home  link back here."
        .Font.Italic = True
    End With

    r = 5
    ws.Cells(r, 2).Value = "PAGE"
    ws.Cells(r, 3).Value = "BUTTONS"
    ws.Cells(r, 4).Value = "CATEGORIES ON IT"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Font.Bold = True
        .Interior.Color = RGB(221, 217, 240)
    End With
    r = r + 1

    tabKeys = SortedKeys(layout)
    If UBound(tabKeys) < LBound(tabKeys) Then
        ws.Cells(r, 2).Value = "No actions are available as buttons yet."
        ws.Cells(r, 2).Font.Italic = True
        r = r + 1
    Else
        For ti = LBound(tabKeys) To UBound(tabKeys)
            tabNum = CLng(tabKeys(ti))
            Set page = layout(tabNum)
            NavLink ws.Cells(r, 2), PANEL_BASE & tabNum, PANEL_BASE & tabNum
            ws.Cells(r, 3).Value = PageButtonCount(page)
            totalButtons = totalButtons + PageButtonCount(page)   ' fix 4: one definition, summed here
            ws.Cells(r, 4).Value = CStr(page("title"))
            ws.Cells(r, 4).Font.Color = RGB(90, 90, 90)
            r = r + 1
        Next ti
    End If

    r = r + 1
    NavLink ws.Cells(r, 2), "Setup Menu  -  first-time setup / re-configure", ADMIN_SHEET
    r = r + 2
    r = HomeCatalog(ws, r, layout)
    r = r + 1
    r = HomeLegend(ws, r, totalButtons, isManual)

    ' Fix round 1, fix 4: Activate raises 1004 on a hidden sheet, and hiding a
    ' generated tab is a supported habit here - BuildAdminPanel explicitly
    ' invites the user to hide Setup Menu, so a user hiding Macros Home too is
    ' not misuse. Unhide before activating; a rebuild should always be able to
    ' bring the user back to the index, hidden or not.
    ws.Visible = xlSheetVisible
    ws.Activate                     ' Range.Select needs its sheet active
    ws.Range("A1").Select
End Sub

' Every button, with its full description and the page it is on. This is the one
' screen that answers "what can this workbook do?" - and unlike a card, a cell
' here is as wide as it needs to be, so nothing is abbreviated. Deliberately reads
' it(3) directly (not CardHelp) - this is where a bounded chain card points a
' user for the full recipe, so it must never itself be bounded.
Private Function HomeCatalog(ws As Worksheet, ByVal r As Long, layout As Object) As Long
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim cols As Object, colKeys As Variant, ci As Long, colIx As Long
    Dim items As Collection, ii As Long, it As Variant, firstRow As Long
    ' Contract's term (section 6): "button catalog", not "all actions" - this
    ' lists only actions that became buttons (off / chains-only ones are
    ' absent), and chains, which are not actions, are present.
    ws.Cells(r, 2).Value = "BUTTON CATALOG"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    ws.Cells(r, 2).Value = "ACTION"
    ws.Cells(r, 3).Value = "PAGE"
    ws.Cells(r, 4).Value = "WHAT IT DOES"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = RGB(240, 240, 245)
    End With
    r = r + 1
    firstRow = r
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        colKeys = SortedKeys(cols)
        For ci = LBound(colKeys) To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            Set items = cols(colIx)
            For ii = 1 To items.Count
                it = items(ii)
                ws.Cells(r, 2).Value = CStr(it(1))
                NavLink ws.Cells(r, 3), PANEL_BASE & tabNum, PANEL_BASE & tabNum
                With ws.Cells(r, 4)
                    .Value = CStr(it(3))
                    .Font.Size = 10
                    .Font.Color = RGB(90, 90, 90)
                    .WrapText = True
                    .VerticalAlignment = xlTop
                End With
                r = r + 1
            Next ii
        Next ci
    Next ti
    ' Nothing here is merged, so AutoFit gives every description its full height.
    If r > firstRow Then ws.Rows(firstRow & ":" & (r - 1)).AutoFit
    HomeCatalog = r
End Function

' The legend. This is the only place the page-code convention and Availability are
' explained where a daily user will actually see them. Fix round 1, fix 3: this is
' now mode-aware - it used to always claim "laid out automatically" and invite the
' user to "take charge of the layout", which is false and actively misleading the
' moment even one Page code exists (that code is already in charge). isManual
' reports which filer actually built this layout (BuildAutoLayout vs
' BuildManualLayout), so the first two lines describe the truth for THIS build,
' not a generic capability. Also folds in the related minor: the mechanism line
' used to promise a blank-coded button "lands on page 1", which CapManualColumns
' can now falsify (it relocates to a continuation page once page 1 is already at
' the column cap) - corrected below so the legend matches where the button
' actually went.
Private Function HomeLegend(ws As Worksheet, ByVal r As Long, ByVal totalButtons As Long, _
                            ByVal isManual As Boolean) As Long
    ws.Cells(r, 2).Value = "HOW THESE PAGES ARE BUILT"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    ' Fix wave: two corrections. (1) A Page code chooses where a button STARTS,
    ' not a guarantee of the page/column it lands on - the column cap (contract
    ' section 6) can still relocate it. (2) A chain's Page code lives in the
    ' Chains block, not the action catalog (which enumerates registry actions
    ' only) - the old wording sent a chain-author to the wrong place.
    If isManual Then
        r = HomeLine(ws, r, totalButtons & " button(s). Your Page codes choose where each button starts - " & _
                             "for an action, the code in the action catalog on the Setup tab; for a chain, " & _
                             "the code in the Chains block. A full page or column can still move a button to " & _
                             "a continuation one.")
        r = HomeLine(ws, r, "To go back to automatic packing, clear every Page code - with none set, categories " & _
                             "are packed onto pages for you again.")
    Else
        r = HomeLine(ws, r, totalButtons & " button(s), laid out automatically - categories packed onto pages, up to " & _
                             CARD_COLS & " per page.")
        r = HomeLine(ws, r, "To take charge of the layout, fill in 'Page code' - for an action, in the action " & _
                             "catalog on the Setup tab; for a chain, in the Chains block:")
    End If
    r = HomeLine(ws, r, "    1A = page 1, column A.    12AB = page 12, column AB.    hide = leave that button off.")
    ' Fix round 3, fix 3: "hide" places nothing, so it does not switch this
    ' workbook to manual mode by itself - a store with only "hide" codes still
    ' packs automatically (fix round 2's fix 5 / this line used to say ANY code
    ' switches the layout, which was true for "hide" too and is not).
    r = HomeLine(ws, r, "As soon as ONE placement code (not counting hide) is filled in, the codes take " & _
                         "over; a blank-coded button lands in page 1's Unassigned area, or a continuation " & _
                         "page once the page has no free column slot left. hide always omits a button, in " & _
                         "either mode.")
    r = HomeLine(ws, r, "Availability on the Setup tab: 'Button' = shown here. 'Chains only' = no button, but a Chain can run it.")
    r = HomeLine(ws, r, "'Off' = no button and it will not run at all; Finalize refuses to activate a Chain that uses one.")
    HomeLegend = r
End Function

Private Function HomeLine(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Size = 10
    ws.Cells(r, 2).Font.Color = RGB(90, 90, 90)
    HomeLine = r + 1
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

' LOAD-BEARING: "Macros Home" starts with PANEL_BASE ("Macros ") but has a
' non-numeric tail, so it deliberately reads False here and DeletePanels never
' touches it. That is what lets a rebuild reuse (not recreate) the index sheet -
' EnsurePanelSheet finds it already there and clears+rewrites it in place. A
' future tidy-up that "fixes" this to also match Macros Home would delete the
' index every rebuild, and Fix round 1's whole point (a stale index never
' survives a rebuild, because EnsurePanelSheet's Cells.Clear wipes it before
' BuildHomePage rewrites it) requires that the sheet still be there to clear.
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
    ' fix round 1: explicit, not relied-on-by-default. The card grid grows
    ' description rows AFTER buttons are placed, which reflows every button
    ' below the first pair - that only works because a Forms button anchors and
    ' resizes with its underlying cells, which is xlMoveAndSize. Make it explicit
    ' rather than leaving the whole card-grid layout resting on Excel's default.
    btn.Placement = xlMoveAndSize
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
    ' Fix round 3, fix 9: the digit run was capped, but the letter run never
    ' was. Two real problems, not just an absurd-input curiosity: (1) a bare-
    ' minimum 4-letter run ("AAAA" = 18279) already exceeds VSPLIT_MULT
    ' (100000)'s assumption that every real colIx from ParseCode stays well
    ' under it, so a long enough letter run would get misdecoded by
    ' LabelManualPages as a vertical-overflow split piece; (2) a sufficiently
    ' long letter run overflows the colIndex accumulation below and raises a
    ' runtime error instead of this function returning False, so a typo like
    ' this could crash Finalize's own validation pass instead of being reported
    ' as a bad code. Capping at 3 letters matches Excel's own column-letter
    ' grammar (real columns never exceed 3 letters, "XFD") and keeps every
    ' accepted colIndex at or under 18278 - safely below VSPLIT_MULT.
    If Len(letters) > 3 Then Exit Function
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
