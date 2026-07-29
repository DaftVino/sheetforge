Attribute VB_Name = "PkgSummarize"
Option Explicit

'================================================================
' PkgSummarize - turning a cleaned dataset into pivots and summaries.
' Wave 4 of the catalog build-out (docs/macro-catalog-plug-and-play.md,
' section 6).
'
' Two of these four actions already exist, buried inside
' PkgImportExport_Import's tail end (PkgImportExport.bas ~178-188):
' refreshing every pivot and forcing a full recalculation after an
' import. This package gives them their own action ids so they can run
' standalone or as their own chain step - summarize.refresh_pivots and
' summarize.recalculate. The legacy import KEEPS its own two-line copy
' rather than calling these; see the wave plan's decision table for why
' this is deliberate, not an oversight.
'
' Declares NO config fields. Every action either takes no input at all
' (recalculate), scans every sheet with no per-run choice
' (refresh_pivots), or prompts for its range/columns at run time
' (build_pivot, freeze_pivot) - none of the four describes a fact about
' the deployment, so none belongs on Setup. Same precedent PkgClean set
' before v2.3.0 gave it its first field.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgSummarize_Manifest(reg As Object)
    Core_DeclarePackage reg, "summarize", "Summarize & Analyze", "1.0.0"

    Core_AddAction reg, "summarize.refresh_pivots", "Refresh All PivotTables", "PkgSummarize_RefreshPivots", _
        "Refresh every PivotTable on every sheet in this workbook. Your own Setup tabs are left out.", _
        "Summarize & Analyze"
    Core_AddAction reg, "summarize.recalculate", "Recalculate Workbook", "PkgSummarize_Recalculate", _
        "Force a full recalculation of every formula in the workbook, the same as pressing Ctrl+Alt+F9.", _
        "Summarize & Analyze"
    Core_AddAction reg, "summarize.build_pivot", "Build a Summary Table", "PkgSummarize_BuildPivot", _
        "Pick a range (with its header row), then which column(s) to group by and which to total. Writes a new PivotTable to a new sheet.", _
        "Summarize & Analyze"
    Core_AddAction reg, "summarize.freeze_pivot", "Freeze a Pivot to Values", "PkgSummarize_FreezePivot", _
        "Click any cell inside a PivotTable, then copy its current layout to a new sheet as static values and formats. The original PivotTable is never changed.", _
        "Summarize & Analyze"
End Sub

'================================================================
' ACTION - Refresh every PivotTable on every sheet
'
' A bad pivot (source range deleted, external connection unreachable) can
' either raise on .RefreshTable, or complete without raising and still
' return False - RefreshTable is documented to return a Boolean, and that
' Boolean is the signal this code trusts, not merely the absence of a
' raised error. One bad pivot does NOT stop the others - see the wave
' plan's Decision 11 for why aborting or staying silent are both wrong
' here.
'================================================================
Public Sub PkgSummarize_RefreshPivots(ctx As Object)
    Dim sh As Worksheet, pt As PivotTable
    Dim total As Long, refreshed As Long, failedNames As String
    Dim thisOk As Boolean, thisErrNum As Long, thisErrDesc As String
    Const TITLE As String = "Refresh All PivotTables"
    On Error GoTo Fail

    ' Count first, without touching anything - lets the refusal below fire
    ' before any confirmation is shown, and lets the confirmation name a
    ' real count instead of asking to "refresh everything" blind.
    For Each sh In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(sh.name) Then total = total + sh.PivotTables.Count
    Next sh

    If total = 0 Then
        Core_Refuse ctx, "There are no PivotTables anywhere in this workbook to refresh." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If MsgBox("This refreshes " & total & " PivotTable(s) across this workbook now." & _
              vbCrLf & vbCrLf & "Proceed?", vbOKCancel + vbQuestion, TITLE) <> vbOK Then Exit Sub

    Core_SpeedOn
    For Each sh In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(sh.name) Then
            For Each pt In sh.PivotTables
                ' MissingItemsLimit is not supported on every pivot cache (an
                ' OLAP source raises on it) - its own error is swallowed here
                ' and MUST NOT be attributed to RefreshTable below, which is
                ' the actual success/failure signal for this pivot.
                On Error Resume Next
                Err.Clear
                pt.PivotCache.MissingItemsLimit = xlMissingItemsNone
                Err.Clear
                On Error GoTo Fail

                ' RefreshTable RETURNS a Boolean - that is the documented
                ' success signal. Still guarded with Resume Next because a
                ' genuinely broken external connection raises rather than
                ' returning False.
                On Error Resume Next
                Err.Clear
                thisOk = pt.RefreshTable
                thisErrNum = Err.Number
                thisErrDesc = Err.Description
                Err.Clear
                On Error GoTo Fail

                If thisErrNum <> 0 Then thisOk = False
                If thisOk Then
                    refreshed = refreshed + 1
                Else
                    failedNames = failedNames & vbCrLf & "   " & sh.name & " ! " & pt.name & _
                                  IIf(thisErrNum <> 0, "  (#" & thisErrNum & " " & thisErrDesc & ")", _
                                      "  (RefreshTable reported failure)")
                End If
            Next pt
        End If
    Next sh
    Core_SpeedOff

    Core_Done ctx, "Refreshed " & refreshed & " of " & total & " PivotTable(s)." & _
           IIf(Len(failedNames) > 0, vbCrLf & vbCrLf & "These did not refresh:" & failedNames, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "summarize.refresh_pivots", Err
End Sub

'================================================================
' ACTION - Force a full recalculation
'
' Never REFUSES: recalculating is always valid to run and changes no
' stored value, only formula results, so there is no precondition to
' fail. It still confirms before running - a full recalculation on a
' large finance workbook can take a noticeable while, and a click on the
' wrong button should still be cancellable at no cost, same as every
' other action here. A cancelled confirmation exits quietly, the same as
' every "user cancelled" exit elsewhere in this wave - it is not a
' refusal and is never reported as one. Exists as its own action so it
' can be its own chain step - summarize.refresh_pivots and the legacy
' import both want "and then recalculate" as their last step.
'================================================================
Public Sub PkgSummarize_Recalculate(ctx As Object)
    Const TITLE As String = "Recalculate Workbook"
    On Error GoTo Fail

    If MsgBox("This recalculates every formula in the workbook now. On a large workbook this " & _
              "can take a while." & vbCrLf & vbCrLf & "Proceed?", vbOKCancel + vbQuestion, TITLE) _
            <> vbOK Then Exit Sub

    Core_SpeedOn
    Application.CalculateFull
    Core_SpeedOff

    Core_Done ctx, "Every formula in the workbook has been recalculated.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "summarize.recalculate", Err
End Sub

'================================================================
' ACTION - Build a PivotTable from a picked range
'
' Row 1 of the picked range is ALWAYS its header - a PivotTable source
' needs field names and Excel itself requires them, so there is no
' "does this have a header" question to ask (Decision 5).
'
' Every chosen measure is SUMMED. A per-measure aggregation choice
' would need a second answer per column, which is the config DSL this
' plan was told not to build (Decision 8) - anyone who needs a
' different aggregation can change the field's function by hand on the
' finished PivotTable.
'================================================================
Public Sub PkgSummarize_BuildPivot(ctx As Object)
    Dim srcRng As Range, srcRef As String, rowKeys As String, valKeys As String
    Dim rowColsArr As Variant, valColsArr As Variant, i As Long
    Dim destName As String, destWs As Worksheet
    Dim pc As PivotCache, pt As PivotTable
    Dim fld As PivotField, hdr As String
    Dim seenHdr As Object, colScan As Long, hdrScan As String
    Dim origNum As Long, origDesc As String, origSrc As String
    Const TITLE As String = "Build a Summary Table"
    On Error GoTo Fail

    Set srcRng = Core_AskRange(TITLE & " - your data, including its header row")
    If srcRng Is Nothing Then Exit Sub                        ' user cancelled
    Set srcRng = Core_NarrowToUsed(srcRng)
    If srcRng Is Nothing Then
        Core_Refuse ctx, "Your selection needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    ElseIf srcRng.Rows.Count < 2 Then
        Core_Refuse ctx, "Your selection needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Every PivotTable field is resolved by header text (PivotFields(hdr)),
    ' which is ambiguous when two columns share a header - Excel itself would
    ' already have silently renamed the second one internally - and a blank
    ' header cannot become a field at all. Both would otherwise surface as a
    ' raw, unexplained COM error deep inside CreatePivotTable/AddDataField
    ' below. Refuse cleanly, up front, instead.
    Set seenHdr = CreateObject("Scripting.Dictionary")
    seenHdr.CompareMode = 1
    For colScan = 1 To srcRng.Columns.Count
        hdrScan = Trim$(CStr(srcRng.Cells(1, colScan).Value))
        If Len(hdrScan) = 0 Then
            Core_Refuse ctx, "Column " & colScan & " of your selection has a blank header." & _
                   vbCrLf & vbCrLf & "Every column needs header text before it can be summarized." & _
                   vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
            Exit Sub
        End If
        If seenHdr.Exists(hdrScan) Then
            Core_Refuse ctx, "Two columns in your selection are both headed '" & hdrScan & "'." & _
                   vbCrLf & vbCrLf & "Give each column its own header, then run this again." & _
                   vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
            Exit Sub
        End If
        seenHdr.Add hdrScan, True
    Next colScan

    If Not Core_AskCols(srcRng, rowKeys, _
            "Which column(s) do you want to GROUP BY (e.g. Store, Month)? " & _
            "Numbered from the left of your selection.", TITLE) Then Exit Sub
    rowColsArr = Core_ParseCols(rowKeys)

    If Not Core_AskCols(srcRng, valKeys, _
            "Which column(s) do you want SUMMED (e.g. Amount)? " & _
            "Numbered from the left of your selection.", TITLE) Then Exit Sub
    valColsArr = Core_ParseCols(valKeys)

    ' Derived from the picked range's own reference, not a fixed "Summary" -
    ' Decision 9: a fixed name run twice in one session collides with itself.
    destName = Core_SafeSheetName(Core_RangeRef(srcRng) & " Summary")
    If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub

    If MsgBox("This creates a new sheet '" & destName & "' with a PivotTable, grouped by " & _
              (UBound(rowColsArr) - LBound(rowColsArr) + 1) & " column(s) and summing " & _
              (UBound(valColsArr) - LBound(valColsArr) + 1) & " column(s)." & vbCrLf & vbCrLf & _
              "Proceed?", vbOKCancel + vbQuestion, TITLE) <> vbOK Then Exit Sub

    ' A fully qualified string reference, not the Range object itself -
    ' PivotCaches.Create is documented to accept SourceData as a Range, but a
    ' plain Range argument has a known history of an unexpected type-mismatch
    ' error on some Excel builds; the string form is the form Microsoft's own
    ' examples use.
    srcRef = "'" & srcRng.Worksheet.name & "'!" & srcRng.Address(RowAbsolute:=True, ColumnAbsolute:=True)

    Core_SpeedOn
    Set destWs = Core_NewSheet(destName)
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=srcRef)
    Set pt = pc.CreatePivotTable(TableDestination:=destWs.Range("A1"), TableName:="Pivot_" & destWs.Index)

    For i = LBound(rowColsArr) To UBound(rowColsArr)
        hdr = CStr(srcRng.Cells(1, CLng(rowColsArr(i))).Value)
        Set fld = pt.PivotFields(hdr)
        fld.Orientation = xlRowField
    Next i
    For i = LBound(valColsArr) To UBound(valColsArr)
        hdr = CStr(srcRng.Cells(1, CLng(valColsArr(i))).Value)
        pt.AddDataField pt.PivotFields(hdr), "Sum of " & hdr, xlSum
    Next i
    Core_SpeedOff

    Core_Done ctx, "Built '" & destName & "' from " & Core_RangeRef(srcRng) & "," & vbCrLf & _
           "grouped by " & (UBound(rowColsArr) - LBound(rowColsArr) + 1) & " column(s), summing " & _
           (UBound(valColsArr) - LBound(valColsArr) + 1) & " column(s).", TITLE
    Exit Sub
Fail:
    ' Not transactional the easy way - Excel needs the sheet to exist before
    ' PivotCaches.Create/CreatePivotTable can target it - so a failure after
    ' Core_NewSheet leaves a partial or empty sheet unless it is cleaned up
    ' here. Snapshot the ORIGINAL failure before attempting cleanup, since
    ' the cleanup itself runs under its own Resume Next and could otherwise
    ' overwrite Err with a cleanup-time error instead of the real one.
    origNum = Err.Number
    origDesc = Err.Description
    origSrc = Err.Source
    Core_SpeedOff
    If Not destWs Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        destWs.Delete                                  ' best-effort - remove the partial sheet
        Application.DisplayAlerts = True
    End If
    On Error Resume Next
    Err.Raise origNum, origSrc, origDesc                ' restore the ORIGINAL failure for Core_Fail
    Core_Fail ctx, "summarize.build_pivot", Err
End Sub

'================================================================
' ACTION - Copy a PivotTable's current output to a new sheet as values
'
' Reads the pivot from wherever the user clicks - Range.PivotTable
' raises when the cell is not part of one, refused here rather than
' surfacing as an engine error.
'
' Copies TableRange2 (the WHOLE visible pivot, report-filter area
' included) as values + number formats + column widths, so what lands
' on the new sheet looks like what was on screen, not just its numbers
' (Decision 10). It never calls .RefreshTable and never writes back to
' the source pivot's own cells - only .Copy reads from it. Field
' interactivity (expand/collapse, the field list, any GetPivotData
' formula elsewhere pointing at it) is NOT preserved - a static values
' sheet cannot have any of that, which is the entire point of freezing
' it, stated so it does not read as an oversight.
'================================================================
Public Sub PkgSummarize_FreezePivot(ctx As Object)
    Dim sel As Range, pt As PivotTable, destName As String, destWs As Worksheet
    Dim srcRng As Range, c As Long
    Dim origNum As Long, origDesc As String, origSrc As String
    Const TITLE As String = "Freeze a Pivot to Values"
    On Error GoTo Fail

    Set sel = Core_AskRange(TITLE & " - click any cell inside the PivotTable to freeze")
    If sel Is Nothing Then Exit Sub                           ' user cancelled

    On Error Resume Next
    Set pt = sel.Cells(1, 1).PivotTable
    On Error GoTo Fail
    If pt Is Nothing Then
        Core_Refuse ctx, "The cell you picked is not part of a PivotTable." & vbCrLf & vbCrLf & _
               "Click a cell inside the PivotTable you want to freeze, then run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    destName = Core_SafeSheetName(pt.name & " (values)")
    If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub

    If MsgBox("This copies '" & pt.name & "' to a new sheet '" & destName & "' as static " & _
              "values and formats. The original PivotTable is not changed." & vbCrLf & vbCrLf & _
              "Proceed?", vbOKCancel + vbQuestion, TITLE) <> vbOK Then Exit Sub

    Core_SpeedOn
    Set srcRng = pt.TableRange2
    Set destWs = Core_NewSheet(destName)
    srcRng.Copy
    destWs.Range("A1").PasteSpecial Paste:=xlPasteValues
    destWs.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    ' Column widths are not part of xlPasteFormats - copied separately so the
    ' frozen copy reads the way the pivot looked on screen, not just the way
    ' its cells are formatted.
    For c = 1 To srcRng.Columns.Count
        destWs.Columns(c).ColumnWidth = srcRng.Columns(c).ColumnWidth
    Next c
    Core_SpeedOff

    Core_Done ctx, "Copied '" & pt.name & "' to '" & destName & "' as values." & vbCrLf & _
           "The original PivotTable was not changed.", TITLE
    Exit Sub
Fail:
    ' Same non-transactional shape as summarize.build_pivot: the sheet must
    ' exist before the two PasteSpecial calls can target it, so a failure
    ' between Core_NewSheet and the end of the column-width loop leaves a
    ' partial sheet unless removed here. Snapshot the ORIGINAL failure before
    ' attempting cleanup - the cleanup runs under its own Resume Next and
    ' could otherwise overwrite Err with a cleanup-time error instead of the
    ' real one.
    origNum = Err.Number
    origDesc = Err.Description
    origSrc = Err.Source
    Core_SpeedOff
    Application.CutCopyMode = False
    If Not destWs Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        destWs.Delete                                  ' best-effort - remove the partial sheet
        Application.DisplayAlerts = True
    End If
    On Error Resume Next
    Err.Raise origNum, origSrc, origDesc                ' restore the ORIGINAL failure for Core_Fail
    Core_Fail ctx, "summarize.freeze_pivot", Err
End Sub
