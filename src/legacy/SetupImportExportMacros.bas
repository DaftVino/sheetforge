Attribute VB_Name = "SetupImportExportMacros"
Option Explicit

'================================================================
' Sample Dashboard - Automation Macros
' Buttons (built by BuildMacroTab) call the Public subs below.
'
' Source files are matched by NAME PREFIX in the same folder as
' this workbook, so trailing " (1)", " (2)" etc. are handled.
'
' Per-tab column TYPE maps coerce the text-based source exports
' back to the exact types each dashboard tab currently uses, so
' downstream formulas keep working.
'   Codes:  T=text  N=number(strips $ , %)  D=date  DT=date+time
'           B=boolean   G=general(auto-detect)
'
' --- REUSABLE SETUP (added) -----------------------------------
' "Create Setup Tab"  builds a visible 'Setup' sheet (pre-filled)
'   listing the export tabs, the import tab/file/type triples, the
'   Ref_Master mapping, and the export filename/date.
' "Finalize Setup Tab" validates those answers and copies them to a
'   hidden '_Config' sheet that the other macros read at run time.
' If '_Config' is absent (never finalized), every macro falls back
' to the original hardcoded values, so behavior is unchanged out of
' the box. No "Trust access to the VBA project" setting is required.
'================================================================

Private Const MACRO_SHEET As String = "Macros"
Private Const SETTINGS_SHEET As String = "README"      ' default; overridable in Setup
Private Const SETUP_SHEET As String = "Setup"
Private Const CONFIG_SHEET As String = "_Config"       ' hidden, written by Finalize

' ---- config keys (column A of Setup / _Config) ----
Private Const K_SETTINGS As String = "k_settings_sheet"
Private Const K_EXPBASE  As String = "k_export_base"
Private Const K_EXPOFF   As String = "k_export_offset"
Private Const K_EXPTABS  As String = "k_export_tabs"
Private Const K_IMPORT   As String = "k_import"
Private Const K_MPREFIX  As String = "k_master_prefix"
Private Const K_MSHEET   As String = "k_master_sheet"
Private Const K_MDEST    As String = "k_master_dest"
Private Const K_MCOLS    As String = "k_master_cols"
Private Const K_END      As String = "k_end"

Private Const MAX_LIST As Long = 20        ' rows per list block (export tabs, import files)
Private Const MAX_MCOLS As Long = 10       ' rows for Ref_Master column list

' ---- Configuration for the input tabs (now config-driven) ----
' Reads the active '_Config' sheet; if absent, uses the original
' hardcoded four-tab setup so nothing changes until you run Setup.
Private Sub InputConfig(ByRef tabNames() As String, _
                        ByRef prefixes() As String, _
                        ByRef typeMaps() As String)
    Dim a() As String, b() As String, c() As String, i As Long, n As Long
    If SheetTable3(CfgSheet(), K_IMPORT, a, b, c) Then
        n = UBound(a)
        ReDim tabNames(1 To n)
        ReDim prefixes(1 To n)
        ReDim typeMaps(1 To n)
        For i = 1 To n
            tabNames(i) = a(i)
            prefixes(i) = b(i)
            typeMaps(i) = c(i)
        Next i
        Exit Sub
    End If

    ' ---- fallback: original hardcoded defaults ----
    ReDim tabNames(1 To 4)
    ReDim prefixes(1 To 4)
    ReDim typeMaps(1 To 4)

    tabNames(1) = "IN_Data1"
    prefixes(1) = "Source File A"
    typeMaps(1) = "T,T,T,T,T,T,T,T,N,N,N,N,N,N,N"

    tabNames(2) = "IN_Data2"
    prefixes(2) = "Source File B"
    typeMaps(2) = "T,T,T,T,T,T,T,T,N,N,N,N,N,N,N"

    tabNames(3) = "IN_Data3"
    prefixes(3) = "Source File C"
    typeMaps(3) = "T,T,T,T,T,T,T,N,N,N,N"

    ' IN_Data4 is kept ALL TEXT for its identifier/date columns: a downstream
    ' step parses the time text with DATEVALUE/TIMEVALUE, and the key column
    ' must stay text to match the master tab. Only the numeric columns are N.
    tabNames(4) = "IN_Data4"
    prefixes(4) = "Source File D"
    typeMaps(4) = "T,T,T,T,T,T,T,T,T,N,N,N,N,N"
End Sub

'================================================================
' BUTTON 1 - Clear input tab data (keep headers + formatting)
'================================================================
Public Sub ClearInputTabs()
    Dim tabNames() As String, prefixes() As String, typeMaps() As String
    Dim i As Long, ws As Worksheet
    Dim lastRow As Long, lastCol As Long, urRow As Long
    Dim tlist As String

    On Error GoTo Fail
    InputConfig tabNames, prefixes, typeMaps

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For i = LBound(tabNames) To UBound(tabNames)
        tlist = tlist & IIf(tlist = "", "", ", ") & tabNames(i)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(tabNames(i))
        On Error GoTo Fail
        If Not ws Is Nothing Then
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            urRow = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
            If urRow > lastRow Then lastRow = urRow
            If lastRow >= 2 Then
                ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
            End If
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Cleared data (row 2 and below) on:" & vbCrLf & _
           "  " & tlist & vbCrLf & vbCrLf & _
           "Headers and formatting were kept.", vbInformation, "Clear Input Tabs"
    Exit Sub
Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "ClearInputTabs error: " & Err.Description, vbExclamation, "Clear Input Tabs"
End Sub

'================================================================
' BUTTON 2 - Clear, then import each source file into its tab
'================================================================
Public Sub ImportSourceData()
    Dim tabNames() As String, prefixes() As String, typeMaps() As String
    Dim i As Long, ws As Worksheet, srcPath As String
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim arr As Variant, codes() As String, code As String
    Dim r As Long, c As Long, sLastRow As Long, sLastCol As Long
    Dim logTxt As String

    On Error GoTo Fail
    InputConfig tabNames, prefixes, typeMaps

    ' Step 1: clear the input tabs so we paste into blanks
    ClearInputTabs

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    For i = LBound(tabNames) To UBound(tabNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(tabNames(i))
        On Error GoTo Fail
        If ws Is Nothing Then
            logTxt = logTxt & "X   " & tabNames(i) & "   <- TAB NOT FOUND in this workbook" & vbCrLf
            GoTo NextTab
        End If
        srcPath = FindSourceFile(prefixes(i))
        If srcPath = "" Then
            logTxt = logTxt & "X   " & tabNames(i) & "   <- NO FILE matching """ & prefixes(i) & "*""" & vbCrLf
        Else
            Set wbSrc = Application.Workbooks.Open(fileName:=srcPath, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, CorruptLoad:=xlRepairFile)
            Set wsSrc = wbSrc.Worksheets(1)
            sLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            sLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
            codes = Split(typeMaps(i), ",")
            If sLastRow >= 2 Then
                arr = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(sLastRow, sLastCol)).Value
                For c = 1 To sLastCol
                    If (c - 1) <= UBound(codes) Then code = Trim(codes(c - 1)) Else code = "G"
                    ' Text columns must be formatted as Text BEFORE writing, otherwise
                    ' Excel converts numeric-looking strings (e.g. store "12345") back to
                    ' numbers, which breaks the dashboard's type-strict MATCH lookups
                    ' (e.g. totals, percentages, key metrics).
                    If code = "T" Then _
                        ws.Range(ws.Cells(2, c), ws.Cells(sLastRow, c)).NumberFormat = "@"
                    For r = 1 To UBound(arr, 1)
                        arr(r, c) = CoerceVal(arr(r, c), code)
                    Next r
                Next c
                ws.Range(ws.Cells(2, 1), ws.Cells(sLastRow, sLastCol)).Value = arr
                logTxt = logTxt & "OK  " & tabNames(i) & "   <- " & FileNameOnly(srcPath) & _
                         "   (" & (sLastRow - 1) & " rows)" & vbCrLf
            Else
                logTxt = logTxt & "!   " & tabNames(i) & "   <- " & FileNameOnly(srcPath) & "   (no data rows)" & vbCrLf
            End If
            wbSrc.Close SaveChanges:=False
            Set wbSrc = Nothing
        End If
NextTab:
    Next i

    Application.Calculation = xlCalculationAutomatic

    ' Refresh pivot tables and force a full recalc so downstream tabs
    ' (the summary tabs, etc.) update from the new data.
    Dim sh As Worksheet, pt As PivotTable
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
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Import complete (input tabs were cleared first):" & vbCrLf & vbCrLf & logTxt, _
           vbInformation, "Import Source Data"
    Exit Sub
Fail:
    Dim eMsg As String
    eMsg = "#" & Err.Number & "  " & Err.Description
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "ImportSourceData error: " & eMsg, vbExclamation, "Import Source Data"
End Sub

'================================================================
' BUTTON 3 - Update Ref_Master from the Master List Import
'================================================================
Public Sub UpdateRefMaster()
    Dim srcPath As String, wbSrc As Workbook, wsSrc As Worksheet, wsDst As Worksheet
    Dim hdr As Variant, lastRow As Long, lastCol As Long, dLast As Long
    Dim j As Long, outR As Long, n As Long, k As Long
    Dim cols As Collection, nCols As Long, srcCol() As Long, missing As String
    Dim mPrefix As String, mSheet As String, mDest As String

    On Error GoTo Fail
    mPrefix = MasterPrefix()
    mSheet = MasterSheetName()
    mDest = MasterDestTab()

    srcPath = FindSourceFile(mPrefix)
    If srcPath = "" Then
        MsgBox "No file matching """ & mPrefix & "*"" was found in this folder.", _
               vbExclamation, "Update Ref_Master"
        Exit Sub
    End If
    If Not SheetExists(mDest) Then
        MsgBox "Destination tab '" & mDest & "' was not found in this workbook.", _
               vbExclamation, "Update Ref_Master"
        Exit Sub
    End If
    Set wsDst = ThisWorkbook.Worksheets(mDest)
    Set cols = GetMasterCols()
    nCols = cols.Count
    ReDim srcCol(1 To nCols)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    Set wbSrc = Application.Workbooks.Open(fileName:=srcPath, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, CorruptLoad:=xlRepairFile)
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(mSheet)
    On Error GoTo Fail
    If wsSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
        GoTo Restore
    End If

    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    hdr = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Value

    For k = 1 To nCols
        srcCol(k) = FindCol(hdr, CStr(cols(k)))
        If srcCol(k) = 0 Then missing = missing & vbCrLf & "   - " & cols(k)
    Next k
    If Len(missing) > 0 Then
        wbSrc.Close SaveChanges:=False
        MsgBox "Could not find these column header(s) on sheet '" & mSheet & "':" & missing, _
               vbExclamation, "Update Ref_Master"
        GoTo Restore
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, srcCol(1)).End(xlUp).Row

    dLast = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).Row
    If dLast >= 2 Then wsDst.Range(wsDst.Cells(2, 1), wsDst.Cells(dLast, nCols)).ClearContents

    ' Force the first (key) column to Text so codes/IDs keep leading zeros
    ' and keep matching the workbook's text-based keys.
    wsDst.Columns(1).NumberFormat = "@"

    outR = 2
    For j = 2 To lastRow
        If Trim(CStr(wsSrc.Cells(j, srcCol(1)).Value)) <> "" Then
            wsDst.Cells(outR, 1).Value = CStr(wsSrc.Cells(j, srcCol(1)).Value)   ' Store (kept as text)
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
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox mDest & " updated from """ & FileNameOnly(srcPath) & """ (sheet '" & mSheet & "')." & vbCrLf & _
           n & " rows written.", vbInformation, "Update Ref_Master"
    Exit Sub

Restore:
    Application.Calculation = xlCalculationAutomatic
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Sheet '" & mSheet & "' not found in the master file.", _
           vbExclamation, "Update Ref_Master"
    Exit Sub
Fail:
    Dim eMsg As String
    eMsg = "#" & Err.Number & "  " & Err.Description
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "UpdateRefMaster error: " & eMsg, vbExclamation, "Update Ref_Master"
End Sub

'================================================================
' BUTTON 4 - Export the main tabs as a values-only .xlsx
'================================================================
Public Sub ExportDashboard()
    Dim mainTabs As Variant, t As Variant
    Dim newWb As Workbook, ws As Worksheet
    Dim fRange As Range, ar As Range
    Dim keepD As Object, sh As Worksheet, i As Long
    Dim lnks As Variant, li As Long
    Dim outPath As String, fileName As String
    Dim setSh As String
    On Error GoTo Fail
    mainTabs = GetExportTabs()
    setSh = SettingsSheetName()
    ' every visible main tab must exist
    For Each t In mainTabs
        If Not SheetExists(CStr(t)) Then
            MsgBox "Missing tab: " & t & vbCrLf & "Export cancelled.", vbExclamation, "Export Dashboard"
            Exit Sub
        End If
    Next t
    If Not SheetExists(setSh) Then
        MsgBox "Settings sheet '" & setSh & "' not found. It must stay in" & vbCrLf & _
               "the export (hidden) for the conditional-formatting thresholds.", _
               vbExclamation, "Export Dashboard"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.PrintCommunication = False
    Application.CalculateFull
    ' Copy EVERY worksheet to a brand-new workbook. All tabs present =>
    ' all references resolve internally => no blanks when we freeze.
    ' (Copying the sheet collection does NOT bring the VBA project, so the
    '  new workbook has no macros and no external links.)
    ThisWorkbook.Worksheets.Copy
    Set newWb = Application.ActiveWorkbook
    newWb.Application.CalculateFull        ' make sure every value is current

    ' Flatten dynamic-array SPILLS and legacy CTRL+SHIFT+ENTER ARRAYS to static
    ' values FIRST, then freeze ordinary formulas. Freezing a block that holds
    ' a spill dropped every spilled row; freezing a block that holds a legacy
    ' array errored out and left those formulas un-frozen (they became #REF!
    ' once the input tabs were deleted - the Trends bug).
    Dim cel As Range, sp As Range, sv As Variant
    For Each ws In newWb.Worksheets
        Set fRange = Nothing
        On Error Resume Next
        Set fRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo Fail
        If Not fRange Is Nothing Then
            For Each cel In fRange.Cells
                If cel.HasSpill Then                  ' dynamic-array spill (e.g. a FILTER formula)
                    Set sp = cel.SpillingToRange
                    sv = sp.Value
                    cel.ClearContents
                    sp.Value = sv
                ElseIf cel.HasArray Then              ' legacy CTRL+SHIFT+ENTER array (e.g. a TRANSPOSE array)
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
    fileName = Format(Date + ExportDateOffset(), "yyyy-mm-dd") & " " & ExportBaseName() & ".xlsx"
    outPath = ThisWorkbook.Path
    If Right(outPath, 1) <> Application.PathSeparator Then outPath = outPath & Application.PathSeparator
    outPath = outPath & fileName
    newWb.SaveAs fileName:=outPath, FileFormat:=51   ' 51 = .xlsx (no macros)
    newWb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.PrintCommunication = True
    MsgBox "Exported (all rows, all formatting + conditional formatting kept) to:" & vbCrLf & vbCrLf & _
           outPath, vbInformation, "Export Dashboard"
    Exit Sub
Fail:
    Dim eMsg As String
    eMsg = "#" & Err.Number & "  " & Err.Description
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.PrintCommunication = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "ExportDashboard error: " & eMsg, vbExclamation, "Export Dashboard"
End Sub

'================================================================
' BUTTON 5 - Create the Setup tab (pre-filled by scraping)
'================================================================
Public Sub CreateSetupTab()
    Dim ws As Worksheet, r As Long, i As Long
    Dim ets As Variant
    Dim itn() As String, ipf() As String, itm() As String
    Dim mcols As Collection
    On Error GoTo Fail

    If SheetExists(SETUP_SHEET) Then
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
    ws.Cells(r, 2).Value = "Dashboard Setup"
    ws.Cells(r, 2).Font.Size = 16: ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Color = RGB(112, 48, 160)
    r = r + 1
    ws.Cells(r, 2).Value = "Fill in the YELLOW cells, then click  '6. Finalize Setup Tab'  on the Macros tab."
    ws.Cells(r, 2).Font.Italic = True
    r = r + 1
    ws.Cells(r, 2).Value = "Do not edit column A (machine keys). Blank rows in a list are ignored."
    ws.Cells(r, 2).Font.Italic = True: ws.Cells(r, 2).Font.Color = RGB(150, 0, 0)
    r = r + 1
    ws.Cells(r, 7).Value = "Reference: tabs in THIS workbook"
    ws.Cells(r, 7).Font.Bold = True
    ws.Cells(r, 8).Value = "Reference: Excel files in THIS folder"
    ws.Cells(r, 8).Font.Bold = True
    PopulateReferenceLists ws, r + 1
    r = r + 1

    ' ---- GENERAL ----
    r = SectionHeader(ws, r, "GENERAL")
    r = WriteSetting(ws, r, K_SETTINGS, "Settings / CF sheet name (kept hidden inside the export)", SettingsSheetName())
    r = WriteSetting(ws, r, K_EXPBASE, "Export file BASE name (the date is added in front)", ExportBaseName())
    r = WriteSetting(ws, r, K_EXPOFF, "Export date offset in days  (-1 = yesterday, 0 = today)", CStr(ExportDateOffset()))
    r = r + 1

    ' ---- EXPORT TABS ----
    r = SectionHeader(ws, r, "EXPORT TABS  (copied into the export, in this order)")
    ws.Cells(r, 1).Value = K_EXPTABS
    ws.Cells(r, 2).Value = "List each tab name to export, one per row (up to " & MAX_LIST & "):"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Value = "Tab name  v"
    ws.Cells(r, 3).Font.Bold = True
    r = r + 1
    ets = GetExportTabs()
    For i = 1 To MAX_LIST
        If (i - 1) <= UBound(ets) - LBound(ets) Then ws.Cells(r, 3).Value = ets(LBound(ets) + i - 1)
        Yellow ws.Cells(r, 3)
        r = r + 1
    Next i
    r = r + 1

    ' ---- IMPORT FILES ----
    r = SectionHeader(ws, r, "IMPORT FILES  (Button 2)")
    ws.Cells(r, 1).Value = K_IMPORT
    ws.Cells(r, 2).Value = "For each input tab: its name, the source-file name prefix, and the column type map."
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Value = "Input tab  v": ws.Cells(r, 3).Font.Bold = True
    ws.Cells(r, 4).Value = "Source file name prefix": ws.Cells(r, 4).Font.Bold = True
    ws.Cells(r, 5).Value = "Type map (T,N,D,DT,B,G)": ws.Cells(r, 5).Font.Bold = True
    r = r + 1
    InputConfig itn, ipf, itm
    For i = 1 To MAX_LIST
        If i <= UBound(itn) Then
            ws.Cells(r, 3).Value = itn(i)
            ws.Cells(r, 4).Value = ipf(i)
            ws.Cells(r, 5).Value = itm(i)            ' type maps are letters -> stay text naturally
        End If
        Yellow ws.Cells(r, 3): Yellow ws.Cells(r, 4): Yellow ws.Cells(r, 5)
        r = r + 1
    Next i
    r = r + 1

    ' ---- REF_MASTER ----
    r = SectionHeader(ws, r, "REF_MASTER  (Button 3)")
    r = WriteSetting(ws, r, K_MPREFIX, "Master-list file name prefix", MasterPrefix())
    r = WriteSetting(ws, r, K_MSHEET, "Source sheet name inside that file", MasterSheetName())
    r = WriteSetting(ws, r, K_MDEST, "Destination tab in THIS workbook", MasterDestTab())
    ws.Cells(r, 1).Value = K_MCOLS
    ws.Cells(r, 2).Value = "Source column headers to copy, in order (1st = Store, kept as text):"
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 3).Value = "Header  v": ws.Cells(r, 3).Font.Bold = True
    r = r + 1
    Set mcols = GetMasterCols()
    For i = 1 To MAX_MCOLS
        If i <= mcols.Count Then ws.Cells(r, 3).Value = mcols(i)
        Yellow ws.Cells(r, 3)
        r = r + 1
    Next i

    ' ---- END ----
    ws.Cells(r, 1).Value = K_END
    ws.Cells(r, 2).Value = "(end of config - nothing below this is read)"
    ws.Cells(r, 2).Font.Italic = True: ws.Cells(r, 2).Font.Color = RGB(150, 150, 150)

    ws.Columns("A").Font.Color = RGB(190, 190, 190)
    ws.Columns("A").Hidden = True
    ws.Activate
    ws.Range("B1").Select
    Application.ScreenUpdating = True
    MsgBox "'" & SETUP_SHEET & "' tab created and pre-filled." & vbCrLf & vbCrLf & _
           "Edit the yellow cells, then run  '6. Finalize Setup Tab'.", vbInformation, "Create Setup Tab"
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "CreateSetupTab error: " & Err.Description, vbExclamation, "Create Setup Tab"
End Sub

'================================================================
' BUTTON 6 - Validate the Setup tab and activate it (-> _Config)
'================================================================
Public Sub FinalizeSetupTab()
    Dim ws As Worksheet, cfg As Worksheet
    Dim rpt As String, warn As Long, i As Long
    Dim setSh As String, et As Collection
    Dim a() As String, b() As String, c() As String
    Dim md As String, mp As String
    On Error GoTo Fail

    If Not SheetExists(SETUP_SHEET) Then
        MsgBox "No '" & SETUP_SHEET & "' tab found. Run '5. Create Setup Tab' first.", _
               vbExclamation, "Finalize Setup Tab"
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(SETUP_SHEET)

    ' --- validate (against the Setup tab, before activating) ---
    setSh = SheetSingle(ws, K_SETTINGS, "")
    If setSh = "" Then
        rpt = rpt & "- Settings / CF sheet name is blank." & vbCrLf: warn = warn + 1
    ElseIf Not SheetExists(setSh) Then
        rpt = rpt & "- Settings sheet '" & setSh & "' not found in this workbook." & vbCrLf: warn = warn + 1
    End If

    Set et = SheetList1(ws, K_EXPTABS)
    If et.Count = 0 Then
        rpt = rpt & "- No export tabs are listed." & vbCrLf: warn = warn + 1
    Else
        For i = 1 To et.Count
            If Not SheetExists(CStr(et(i))) Then
                rpt = rpt & "- Export tab '" & et(i) & "' not found in this workbook." & vbCrLf: warn = warn + 1
            End If
        Next i
    End If

    If SheetTable3(ws, K_IMPORT, a, b, c) Then
        For i = 1 To UBound(a)
            If Not SheetExists(a(i)) Then
                rpt = rpt & "- Import tab '" & a(i) & "' not found in this workbook." & vbCrLf: warn = warn + 1
            End If
            If FindSourceFile(b(i)) = "" Then
                rpt = rpt & "- No file matches prefix '" & b(i) & "' in this folder (that tab will be skipped on import)." & vbCrLf: warn = warn + 1
            End If
        Next i
    Else
        rpt = rpt & "- No import files are listed." & vbCrLf: warn = warn + 1
    End If

    md = SheetSingle(ws, K_MDEST, "")
    If md <> "" Then
        If Not SheetExists(md) Then
            rpt = rpt & "- Ref_Master destination tab '" & md & "' not found." & vbCrLf: warn = warn + 1
        End If
    End If
    mp = SheetSingle(ws, K_MPREFIX, "")
    If mp <> "" Then
        If FindSourceFile(mp) = "" Then
            rpt = rpt & "- No file matches master prefix '" & mp & "' (Update Ref_Master will warn)." & vbCrLf: warn = warn + 1
        End If
    End If

    Dim msg As String
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
           "The Clear / Import / Update / Export buttons now use your Setup answers.", _
           vbInformation, "Finalize Setup Tab"
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "FinalizeSetupTab error: " & Err.Description, vbExclamation, "Finalize Setup Tab"
End Sub

'================================================================
' SETUP - build the Macros tab and wire the buttons
'================================================================
Public Sub BuildMacroTab()
    Dim ws As Worksheet
    Dim leftPos As Double, topPos As Double, btnW As Double, btnH As Double, gap As Double

    On Error GoTo Fail
    Application.ScreenUpdating = False

    If SheetExists(MACRO_SHEET) Then
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
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 92

    With ws.Range("B2")
        .Value = "Sample Dashboard  -  Control Panel"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    With ws.Range("B3")
        .Value = "Daily run: 2 -> 3 -> 4.   First-time setup / reuse: 5 -> 6."
        .Font.Italic = True
    End With

    btnW = 220: btnH = 38: gap = 16
    leftPos = ws.Range("B5").Left
    topPos = ws.Range("B5").Top

    AddBtn ws, "btnClear", "1.  Clear Input Tabs", "ClearInputTabs", leftPos, topPos, btnW, btnH
    ws.Range("C5").Value = "Deletes data (row 2 down) on the input tabs. Keeps headers + formatting."
    topPos = topPos + btnH + gap

    AddBtn ws, "btnImport", "2.  Import Source Data", "ImportSourceData", leftPos, topPos, btnW, btnH
    ws.Range("C8").Value = "Clears the input tabs, then imports the matching source files from this folder (auto-handles (1)/(2) name suffixes)."
    topPos = topPos + btnH + gap

    AddBtn ws, "btnRef", "3.  Update Ref_Master", "UpdateRefMaster", leftPos, topPos, btnW, btnH
    ws.Range("C11").Value = "Refreshes the Ref_Master tab from the master-list file in this folder."
    topPos = topPos + btnH + gap

    AddBtn ws, "btnExport", "4.  Export Dashboard (.xlsx)", "ExportDashboard", leftPos, topPos, btnW, btnH
    ws.Range("C14").Value = "Saves a values-only copy of the main tabs (formatting + filters + conditional formatting kept)."
    topPos = topPos + btnH + gap

    AddBtn ws, "btnCreateSetup", "5.  Create Setup Tab", "CreateSetupTab", leftPos, topPos, btnW, btnH
    ws.Range("C17").Value = "Builds a pre-filled 'Setup' tab where you list the export tabs, import files, and Ref_Master mapping. Use this to reuse these macros in another workbook."
    topPos = topPos + btnH + gap

    AddBtn ws, "btnFinalizeSetup", "6.  Finalize Setup Tab", "FinalizeSetupTab", leftPos, topPos, btnW, btnH
    ws.Range("C20").Value = "Validates your Setup answers and activates them. Buttons 1-4 then use your settings instead of the built-in defaults."

    ws.Range("C5,C8,C11,C14,C17,C20").Font.Size = 10
    ws.Range("C5,C8,C11,C14,C17,C20").VerticalAlignment = xlTop
    ws.Range("C5,C8,C11,C14,C17,C20").WrapText = True

    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True
    MsgBox "Macros tab built with 6 buttons.", vbInformation, "Setup Complete"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "BuildMacroTab error: " & Err.Description, vbExclamation, "Setup"
End Sub

'================================================================
' Config readers  (active config = the hidden _Config sheet;
' if it is absent every wrapper returns the original default)
'================================================================
Private Function CfgSheet() As Worksheet
    On Error Resume Next
    Set CfgSheet = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0
End Function

Private Function CfgFindRow(ws As Worksheet, key As String) As Long
    Dim lastR As Long, r As Long
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastR
        If LCase(Trim(CStr(ws.Cells(r, 1).Value))) = LCase(key) Then
            CfgFindRow = r
            Exit Function
        End If
    Next r
End Function

Private Function CfgNextKeyRow(ws As Worksheet, afterRow As Long) As Long
    Dim lastR As Long, r As Long, v As String
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = afterRow + 1 To lastR
        v = LCase(Trim(CStr(ws.Cells(r, 1).Value)))
        If Left(v, 2) = "k_" Then
            CfgNextKeyRow = r
            Exit Function
        End If
    Next r
    CfgNextKeyRow = lastR + 1
End Function

' single value (column C of the key's row), else default
Private Function SheetSingle(ws As Worksheet, key As String, ByVal dflt As String) As String
    Dim r As Long, v As String
    If ws Is Nothing Then SheetSingle = dflt: Exit Function
    r = CfgFindRow(ws, key)
    If r = 0 Then SheetSingle = dflt: Exit Function
    v = Trim(CStr(ws.Cells(r, 3).Value))
    If v = "" Then SheetSingle = dflt Else SheetSingle = v
End Function

' one-column list (column C) under a key, bounded by the next key row
Private Function SheetList1(ws As Worksheet, key As String) As Collection
    Dim r As Long, nextK As Long, i As Long, v As String
    Dim col As New Collection
    Set SheetList1 = col
    If ws Is Nothing Then Exit Function
    r = CfgFindRow(ws, key)
    If r = 0 Then Exit Function
    nextK = CfgNextKeyRow(ws, r)
    For i = r + 1 To nextK - 1
        v = Trim(CStr(ws.Cells(i, 3).Value))
        If v <> "" Then col.Add v
    Next i
End Function

' three-column table (C,D,E) under a key; returns 1-based arrays
Private Function SheetTable3(ws As Worksheet, key As String, _
                             ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    Dim r As Long, nextK As Long, i As Long, n As Long, k As Long
    If ws Is Nothing Then Exit Function
    r = CfgFindRow(ws, key)
    If r = 0 Then Exit Function
    nextK = CfgNextKeyRow(ws, r)
    For i = r + 1 To nextK - 1
        If Trim(CStr(ws.Cells(i, 3).Value)) <> "" Then n = n + 1
    Next i
    If n = 0 Then Exit Function
    ReDim a(1 To n): ReDim b(1 To n): ReDim c(1 To n)
    For i = r + 1 To nextK - 1
        If Trim(CStr(ws.Cells(i, 3).Value)) <> "" Then
            k = k + 1
            a(k) = Trim(CStr(ws.Cells(i, 3).Value))
            b(k) = Trim(CStr(ws.Cells(i, 4).Value))
            c(k) = Trim(CStr(ws.Cells(i, 5).Value))
        End If
    Next i
    SheetTable3 = True
End Function

' ---- typed wrappers used by the macros (active config or default) ----
Private Function SettingsSheetName() As String
    SettingsSheetName = SheetSingle(CfgSheet(), K_SETTINGS, SETTINGS_SHEET)
End Function
Private Function ExportBaseName() As String
    ExportBaseName = SheetSingle(CfgSheet(), K_EXPBASE, "My Report")
End Function
Private Function ExportDateOffset() As Long
    Dim s As String
    s = SheetSingle(CfgSheet(), K_EXPOFF, "-1")
    If IsNumeric(s) Then ExportDateOffset = CLng(s) Else ExportDateOffset = -1
End Function
Private Function MasterPrefix() As String
    MasterPrefix = SheetSingle(CfgSheet(), K_MPREFIX, "Master List Import")
End Function
Private Function MasterSheetName() As String
    MasterSheetName = SheetSingle(CfgSheet(), K_MSHEET, "Current Master")
End Function
Private Function MasterDestTab() As String
    MasterDestTab = SheetSingle(CfgSheet(), K_MDEST, "Ref_Master")
End Function
Private Function GetExportTabs() As Variant
    Dim col As Collection, i As Long, arr() As String
    Set col = SheetList1(CfgSheet(), K_EXPTABS)
    If col.Count = 0 Then
        GetExportTabs = Array("Summary", "Details", "Trends", _
                              "Outliers", "Breakdown")
    Else
        ReDim arr(0 To col.Count - 1)
        For i = 1 To col.Count
            arr(i - 1) = col(i)
        Next i
        GetExportTabs = arr
    End If
End Function
Private Function GetMasterCols() As Collection
    Dim col As Collection
    Set col = SheetList1(CfgSheet(), K_MCOLS)
    If col.Count = 0 Then
        Set col = New Collection
        col.Add "Store ID": col.Add "Location Name or Address": col.Add "Entity ID"
        col.Add "Mgr (DM) Group": col.Add "Mgr (Loc) Name"
    End If
    Set GetMasterCols = col
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
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).Interior.Color = RGB(221, 217, 240)
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Color = RGB(54, 32, 100)
    SectionHeader = r + 1
End Function

Private Function WriteSetting(ws As Worksheet, ByVal r As Long, key As String, _
                              label As String, val As String) As Long
    ws.Cells(r, 1).Value = key
    ws.Cells(r, 2).Value = label
    ws.Cells(r, 3).Value = val
    Yellow ws.Cells(r, 3)
    WriteSetting = r + 1
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

'================================================================
' Helpers (unchanged)
'================================================================
Private Sub AddBtn(ws As Worksheet, nm As String, cap As String, action As String, _
                   l As Double, t As Double, w As Double, h As Double)
    Dim btn As Button
    Set btn = ws.Buttons.Add(l, t, w, h)
    btn.name = nm
    btn.Caption = cap
    btn.OnAction = action
    btn.Font.Size = 11
    btn.Font.Bold = True
End Sub

' Find newest file in this folder whose base name = prefix (optionally
' followed by a " (n)" suffix). Excludes this workbook itself.
Private Function FindSourceFile(ByVal prefix As String) As String
    Dim folder As String, f As String, baseName As String, suffix As String
    Dim names() As String, n As Long, i As Long
    Dim bestFile As String, bestTime As Date, tStamp As Date
    If Trim(prefix) = "" Then Exit Function
    folder = ThisWorkbook.Path
    If Right(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator

    ' First collect every workbook name. Do NOT call FileDateTime (or any other
    ' file function) inside the Dir() loop - that can corrupt the Dir iterator on
    ' some drives (notably network/cloud drives) and make later files invisible.
    ReDim names(1 To 2000)
    n = 0
    f = Dir(folder & "*.xls*")
    Do While Len(f) > 0
        n = n + 1
        names(n) = f
        f = Dir
    Loop

    ' Now evaluate the collected names (safe to stat here).
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
    FindSourceFile = bestFile
End Function

' Return just the file name from a full path (no Dir() call needed).
Private Function FileNameOnly(ByVal p As String) As String
    If InStrRev(p, "\") > 0 Then
        FileNameOnly = Mid(p, InStrRev(p, "\") + 1)
    Else
        FileNameOnly = p
    End If
End Function

' Match a header name to its column index (1-based), case-insensitive.
Private Function FindCol(ByVal hdr As Variant, ByVal name As String) As Long
    Dim k As Long
    For k = LBound(hdr, 2) To UBound(hdr, 2)
        If Trim(LCase(CStr(hdr(1, k)))) = Trim(LCase(name)) Then
            FindCol = k
            Exit Function
        End If
    Next k
    FindCol = 0
End Function

' Coerce a text value from the source export to the target type.
Private Function CoerceVal(ByVal v As Variant, ByVal code As String) As Variant
    Dim s As String
    If IsEmpty(v) Then CoerceVal = "": Exit Function
    If VarType(v) = vbString Then
        If v = "" Then CoerceVal = "": Exit Function
    End If

    Select Case UCase(code)
        Case "T"
            CoerceVal = CStr(v)
        Case "N"
            If IsNumeric(v) Then
                CoerceVal = CDbl(v)
            Else
                s = Replace(Replace(Replace(Replace(CStr(v), "$", ""), ",", ""), " ", ""), "%", "")
                If IsNumeric(s) Then CoerceVal = CDbl(s) Else CoerceVal = v
            End If
        Case "D", "DT"
            If IsDate(v) Then CoerceVal = CDate(v) Else CoerceVal = v
        Case "B"
            CoerceVal = (UCase(Trim(CStr(v))) = "TRUE" Or Trim(CStr(v)) = "1")
        Case Else  ' "G" general / auto
            s = Replace(Replace(CStr(v), ",", ""), "$", "")
            If IsNumeric(s) Then
                CoerceVal = CDbl(s)
            ElseIf IsDate(v) Then
                CoerceVal = CDate(v)
            Else
                CoerceVal = v
            End If
    End Select
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(nm)
    SheetExists = Not s Is Nothing
    On Error GoTo 0
End Function
