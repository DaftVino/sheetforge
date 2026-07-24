Attribute VB_Name = "PkgImportExport"
Option Explicit

'================================================================
' PkgImportExport - the original Food Performance Dashboard's four
' buttons (Clear, Import, Update Ref_Master, Export), re-expressed as
' a thin package on Core. No plumbing lives here: config, speed, error
' handling, file finding and type coercion all call Core_*.
'
' Depends on: Core (required). Discovered zero-trust by Core probing
' PkgImportExport_Manifest via Application.Run.
'
' Behavior is identical to the legacy FoodDashMacros module, with ONE
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
    Core_AddField reg, "k_export_base", "Export file BASE name (the date is added in front)", "text", "Food Performance Dash"
    Core_AddField reg, "k_settings_sheet", "Settings / CF sheet name (kept hidden inside the export)", "sheet", "README", "advanced"
    Core_AddField reg, "k_export_offset", "Export date offset in days  (-1 = yesterday, 0 = today)", "int", "-1", "advanced"

    ' ---- EXPORT TABS ----
    Core_AddSection reg, "EXPORT TABS  (copied into the export, in this order)"
    Core_AddList reg, "k_export_tabs", "List each tab name to export, one per row:", "Tab name  v"
    Core_AddDefault reg, "k_export_tabs", "Food Recap (Daily)"
    Core_AddDefault reg, "k_export_tabs", "Category Var"
    Core_AddDefault reg, "k_export_tabs", "Var Detail"
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
    Core_AddField reg, "k_master_dest", "Destination tab in THIS workbook", "sheet", "Ref_Master"
    Core_AddField reg, "k_master_sheet", "Source sheet name inside that file", "text", "Current Master", "advanced"
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
                    ' turns numeric-looking strings (store "3176") back into numbers and
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
