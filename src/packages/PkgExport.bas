Attribute VB_Name = "PkgExport"
Option Explicit

'================================================================
' PkgExport - getting finished work OUT of the workbook: frozen
' values, PDFs, one file per tab. All three write outside the
' workbook (a value overwrite in place, or a new file on disk), so
' every action here either refuses on a real collision or asks
' before doing the irreversible part - matching the safety bar
' PkgClean set for in-workbook deletes.
'
' export.freeze_values shares its flattening engine, Core_FreezeValues,
' with importexport.export (PkgImportExport.bas) - see the 2026-07-28
' W3 plan for why the engine lives in Core rather than being copied.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgExport_Manifest(reg As Object)
    Core_DeclarePackage reg, "export", "Export", "1.0.0"

    Core_AddSection reg, "GENERAL"
    Core_AddField reg, "k_export_folder", "Folder to write exported files to", "text", "", _
        "simple", _
        "Leave blank to use the folder this workbook is saved in - that is the usual " & _
        "answer. Otherwise a full path, e.g. C:\Reports\January.", _
        "export.pdf,export.per_sheet,export.split_by_value,export.email"
    Core_AddField reg, "k_export_date_offset", "Date shift for exported file names (days)", "int", "-1", _
        "advanced", _
        "How many days to shift the date put in front of an exported file's name. -1 = " & _
        "yesterday, 0 = today, 1 = tomorrow.", _
        "export.pdf,export.per_sheet,export.split_by_value,export.email"

    Core_AddSection reg, "PDF EXPORT"
    Core_AddField reg, "k_export_pdf_base", "Base name for the exported PDF", "text", "Report", _
        "simple", _
        "The saved PDF will be named like:  2026-01-15 Report.pdf", _
        "export.pdf"

    Core_AddSection reg, "PER-TAB EXPORT"
    Core_AddField reg, "k_export_persheet_format", "File type for per-tab export", "text", "pdf", _
        "simple", _
        "pdf or xlsx. Each visible, non-toolkit sheet is written as its own file named " & _
        "after the sheet.", _
        "export.per_sheet"

    Core_AddSection reg, "EMAIL"
    Core_AddField reg, "k_export_email_to", "Email recipient", "text", "", _
        "simple", _
        "The address pre-filled in the Outlook draft. The action displays the draft and " & _
        "never sends it.", _
        "export.email"
    Core_AddField reg, "k_export_email_subject", "Email subject", "text", "Sheet export", _
        "simple", _
        "The subject pre-filled in the Outlook draft.", _
        "export.email"

    Core_AddAction reg, "export.freeze_values", "Freeze Formulas to Values", "PkgExport_FreezeValues", _
        "Pick a range; converts every formula in it (including dynamic-array spills) into a " & _
        "static value. Shows how many formulas it found and asks before doing it - this " & _
        "cannot be undone with Ctrl+Z."
    Core_AddAction reg, "export.pdf", "Export to PDF", "PkgExport_Pdf", _
        "Save the whole workbook, one sheet, or a range you pick, as a dated PDF. Refuses " & _
        "rather than overwriting a same-named file that is already there."
    Core_AddAction reg, "export.per_sheet", "Export Each Tab Separately", "PkgExport_PerSheet", _
        "Save every visible, non-toolkit sheet as its own dated PDF or .xlsx file. Shows " & _
        "the full file list and asks before writing anything; refuses the whole batch up " & _
        "front - naming every collision, whether with an existing file or another sheet in " & _
        "the same batch - rather than writing until something breaks. Once writing starts, " & _
        "a failure partway through leaves the files already written; it is not rolled back."
    Core_AddAction reg, "export.split_by_value", "Save Each Group As Its Own File", "PkgExport_SplitByValue", _
        "Split by a column value into separate files. " & _
        "In a chain, this one stops and waits for your answer before it does anything."
    Core_AddAction reg, "export.email", "Email A Sheet As An Attachment", "PkgExport_Email", _
        "Write one sheet to an .xlsx file, attach it to an Outlook draft with the recipient " & _
        "and subject from Setup, and display the draft without sending it. " & _
        "In a chain, this one stops and waits for your answer before it does anything."
End Sub

'================================================================
' ACTION - Freeze formulas (incl. spills) in a picked range to values
'================================================================
Public Sub PkgExport_FreezeValues(ctx As Object)
    Dim rng As Range, cnt As Long, converted As Long, failCount As Long
    Const TITLE As String = "Freeze Formulas to Values"
    On Error GoTo Fail

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub                   ' user cancelled
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to freeze." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    cnt = 0
    ' Narrow Resume Next to this ONE statement and restore this Sub's own
    ' Fail: routing immediately after - never On Error GoTo 0. GoTo 0 here
    ' would leave every statement below (including the destructive freeze)
    ' unhandled for the rest of the procedure, which is exactly the bug
    ' this action shipped with in an earlier draft of this plan.
    On Error Resume Next
    cnt = rng.SpecialCells(xlCellTypeFormulas).Cells.Count
    On Error GoTo Fail
    If cnt = 0 Then
        Core_Refuse ctx, "There are no formulas inside " & Core_RangeRef(rng) & " to freeze." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If MsgBox("This converts " & cnt & " formula(s) - including any dynamic-array spills - " & _
              "inside " & Core_RangeRef(rng) & " into static values." & vbCrLf & vbCrLf & _
              "This cannot be undone with Ctrl+Z." & vbCrLf & vbCrLf & "Freeze them?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    ' Core_FreezeValues reads live formula results, so it needs Automatic
    ' calculation exactly as importexport.export always has - Core_SpeedOn
    ' sets Manual, so it is put back before the freeze; Core_SpeedOff
    ' restores whatever was there before this action ran.
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    ' This is a user-PICKED range, already confirmed above - reevalUsedRange
    ' stays False (the default) so pass 2 cannot silently widen past what
    ' the user was shown and asked to convert (see Core_FreezeValues's
    ' header, Task 2, difference #2).
    converted = Core_FreezeValues(rng, False, failCount)
    Core_SpeedOff

    If failCount > 0 Then
        Core_Done ctx, "Converted " & converted & " formula(s) inside " & Core_RangeRef(rng) & _
               " to static values." & vbCrLf & vbCrLf & failCount & " formula(s) could NOT " & _
               "be converted and are still live formulas - re-run this action, or convert " & _
               "them by hand (select the cell(s), Ctrl+C, Paste Special > Values)." & vbCrLf & _
               vbCrLf & "This cannot be undone with Ctrl+Z for the ones that DID convert.", TITLE
    Else
        Core_Done ctx, "Converted " & converted & " formula(s) inside " & Core_RangeRef(rng) & _
               " to static values." & vbCrLf & vbCrLf & "This cannot be undone with Ctrl+Z.", TITLE
    End If
    Exit Sub
Fail:
    Core_Fail ctx, "export.freeze_values", Err
End Sub

'================================================================
' ACTION - Export the workbook, one sheet, or a range, to a dated PDF
'================================================================
Public Sub PkgExport_Pdf(ctx As Object)
    Dim scope As Variant, folderPath As String, baseName As String, outPath As String
    Dim ws As Worksheet, rng As Range, savedArea As String, sheetName As Variant
    Dim printAreaChanged As Boolean
    Const TITLE As String = "Export to PDF"
    On Error GoTo Fail

    folderPath = Core_ResolveFolder(Core_Cfg(ctx, "k_export_folder"))
    If Len(folderPath) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to " & _
               "export into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write " & _
               "exported files to' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_FolderExists(folderPath) Then
        Core_Refuse ctx, "This folder was not found: " & folderPath & vbCrLf & vbCrLf & _
               "Check 'Folder to write exported files to' on the Setup tab, Finalize, and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    scope = Application.InputBox( _
        Prompt:="Export the whole workbook, one sheet, or a range you pick?" & vbCrLf & _
                vbCrLf & "Type:  workbook,  sheet,  or  range.", _
        title:=TITLE, Default:="sheet", Type:=2)
    If VarType(scope) = vbBoolean Then Exit Sub               ' user cancelled
    scope = LCase(Trim(CStr(scope)))
    If scope <> "workbook" And scope <> "sheet" And scope <> "range" Then
        Core_Refuse ctx, "'" & scope & "' is not one of the accepted answers: workbook, " & _
               "sheet, or range." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' k_export_pdf_base is free text with no length limit of its own (Task
    ' 1, field table). 100 leaves comfortable room for the "yyyy-mm-dd "
    ' prefix (11 chars) and ".pdf" suffix (4 chars) inside a typical
    ' folder's remaining budget under Windows' MAX_PATH, so an overlong
    ' base name is truncated here instead of failing SaveAs as a generic,
    ' unexplained macro error.
    baseName = Core_SafeFileName(Core_Cfg(ctx, "k_export_pdf_base"), 100)
    outPath = folderPath & Format(Date + Core_CfgLong(ctx, "k_export_date_offset"), "yyyy-mm-dd") & _
              " " & baseName & ".pdf"
    If Len(Dir(outPath)) > 0 Then
        Core_Refuse ctx, "A file already exists at:" & vbCrLf & vbCrLf & "   " & outPath & _
               vbCrLf & vbCrLf & "This action will not overwrite it. Rename or delete that " & _
               "file, or change the base name / date offset on Setup, then run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Select Case scope
        Case "workbook"
            Core_SpeedOn
            ThisWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPath, _
                Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=False
            Core_SpeedOff

        Case "sheet"
            sheetName = Application.InputBox( _
                Prompt:="Sheet to export:", title:=TITLE, Default:=ActiveSheet.name, Type:=2)
            If VarType(sheetName) = vbBoolean Then Exit Sub   ' user cancelled
            If Not Core_SheetExists(CStr(sheetName)) Then
                Core_Refuse ctx, "There is no sheet named '" & sheetName & "'." & vbCrLf & _
                       vbCrLf & "Nothing has been changed.", TITLE
                Exit Sub
            End If
            Set ws = ThisWorkbook.Worksheets(CStr(sheetName))
            If Core_IsEngineSheet(ws.name) Then
                Core_Refuse ctx, "'" & ws.name & "' belongs to the macro toolkit itself, " & _
                       "not to a report." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
                Exit Sub
            End If
            Core_SpeedOn
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPath, _
                Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=False
            Core_SpeedOff

        Case "range"
            Set rng = Core_AskRange(TITLE)
            If rng Is Nothing Then Exit Sub                   ' user cancelled
            Set rng = Core_NarrowToUsed(rng)
            If rng Is Nothing Then
                Core_Refuse ctx, "The range you picked is entirely outside the used area " & _
                       "of the sheet, so there is nothing in it to export." & vbCrLf & _
                       vbCrLf & "Nothing has been changed.", TITLE
                Exit Sub
            End If
            Set ws = rng.Parent
            ' The picked range becomes the sheet's Print Area only for the
            ' duration of this export, then is put back exactly as it was -
            ' contract section 6 forbids an engine touching a surface the
            ' user owns beyond what it is explicitly asked to do, and a
            ' stray Print Area left behind would silently change what a
            ' later manual Ctrl+P prints.
            savedArea = ws.PageSetup.PrintArea
            printAreaChanged = True
            Core_SpeedOn
            ws.PageSetup.PrintArea = rng.Address
            ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=outPath, _
                Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=False
            ws.PageSetup.PrintArea = savedArea
            printAreaChanged = False
            Core_SpeedOff
    End Select

    Core_Done ctx, "Exported the " & scope & " to:" & vbCrLf & vbCrLf & outPath, TITLE
    Exit Sub
Fail:
    On Error Resume Next
    If printAreaChanged And Not ws Is Nothing Then ws.PageSetup.PrintArea = savedArea
    Core_Fail ctx, "export.pdf", Err
End Sub

'================================================================
' ACTION - Explode the workbook into one dated file per visible tab
'================================================================
Public Sub PkgExport_PerSheet(ctx As Object)
    Dim fmt As String, ext As String, folderPath As String, dateStr As String
    Dim ws As Worksheet, targets As Collection, paths As Collection
    Dim i As Long, existing As String, listMsg As String, n As Long
    Dim seen As Object, dupMsg As String
    Const TITLE As String = "Export Each Tab Separately"
    On Error GoTo Fail

    fmt = LCase(Trim(Core_Cfg(ctx, "k_export_persheet_format")))
    If fmt <> "pdf" And fmt <> "xlsx" Then
        Core_Refuse ctx, "'" & fmt & "' is not one of the accepted file types for per-tab " & _
               "export: pdf or xlsx." & vbCrLf & vbCrLf & "Fix 'File type for per-tab " & _
               "export' on the Setup tab, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    ext = "." & fmt

    folderPath = Core_ResolveFolder(Core_Cfg(ctx, "k_export_folder"))
    If Len(folderPath) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to " & _
               "export into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write " & _
               "exported files to' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_FolderExists(folderPath) Then
        Core_Refuse ctx, "This folder was not found: " & folderPath & vbCrLf & vbCrLf & _
               "Check 'Folder to write exported files to' on the Setup tab, Finalize, and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Every VISIBLE, non-toolkit sheet is a candidate. Hidden sheets are
    ' skipped too, on top of the required engine-sheet skip: a hidden
    ' settings/lookup tab is not itself a deliverable, and exporting it by
    ' name would put internal thresholds straight into a file meant to
    ' leave the building.
    dateStr = Format(Date + Core_CfgLong(ctx, "k_export_date_offset"), "yyyy-mm-dd")
    Set targets = New Collection
    Set paths = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Not Core_IsEngineSheet(ws.name) Then
                targets.Add ws
                ' No maxLen cap here (contrast Task 11's k_export_pdf_base
                ' call) - ws.name is already bounded to 31 characters by
                ' Excel itself, so there is nothing unbounded to cap.
                paths.Add folderPath & dateStr & " " & Core_SafeFileName(ws.name) & ext
            End If
        End If
    Next ws

    If targets.Count = 0 Then
        Core_Refuse ctx, "There are no visible, non-toolkit sheets to export." & vbCrLf & _
               vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Detect duplicate target paths WITHIN this batch before checking the
    ' disk. Core_SafeFileName's sanitization is many-to-one (every illegal
    ' character becomes "-"), so two different, both-legal sheet names -
    ' e.g. "A""B" and "A<B" - can sanitize to the identical file name. The
    ' existing-file check below only catches files already on disk; a
    ' collision created BY THIS BATCH would otherwise let the second write
    ' silently overwrite the first, since Core_SpeedOn disables
    ' DisplayAlerts and Excel's own overwrite prompt never fires. Refuse by
    ' name before anything is written, same as the existing-file check.
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For i = 1 To paths.Count
        If seen.Exists(CStr(paths(i))) Then
            dupMsg = dupMsg & vbCrLf & "   '" & targets(i).name & "' and '" & _
                seen(CStr(paths(i))) & "' both sanitize to: " & Core_FileName(CStr(paths(i)))
        Else
            seen.Add CStr(paths(i)), targets(i).name
        End If
    Next i
    If Len(dupMsg) > 0 Then
        Core_Refuse ctx, "These sheet names sanitize to the same file name, so this batch " & _
               "would overwrite its own output:" & dupMsg & vbCrLf & vbCrLf & "Rename one of " & _
               "each pair so the exported files stay distinct, then run this again." & vbCrLf & _
               vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Check EVERY target file for a PRE-EXISTING collision before writing
    ' ANY of them. This upfront refusal is all-or-nothing: nothing is
    ' written if even one target collides. The write itself, below, is NOT
    ' transactional - see the Fail: handler for what happens if it fails
    ' partway through.
    For i = 1 To paths.Count
        If Len(Dir(CStr(paths(i)))) > 0 Then existing = existing & vbCrLf & "   " & paths(i)
    Next i
    If Len(existing) > 0 Then
        Core_Refuse ctx, "These file(s) already exist and this action will not overwrite " & _
               "them:" & existing & vbCrLf & vbCrLf & "Rename or delete them, or change the " & _
               "date offset on Setup, then run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    For i = 1 To targets.Count
        listMsg = listMsg & vbCrLf & "   " & Core_FileName(CStr(paths(i)))
    Next i
    If MsgBox("This writes " & targets.Count & " file(s) to:" & vbCrLf & "   " & folderPath & _
              vbCrLf & vbCrLf & listMsg & vbCrLf & vbCrLf & "Continue?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    ' Both preflight checks above are all-or-nothing (nothing is written if
    ' any collision exists). The write loop itself is NOT: it proceeds file
    ' by file, and there is no rollback if a later file fails - the Fail:
    ' handler below reports how many succeeded before the error, using the
    ' same `n` this loop increments, so a rerun does not have to guess.
    Core_SpeedOn
    If fmt = "pdf" Then
        For i = 1 To targets.Count
            targets(i).ExportAsFixedFormat Type:=xlTypePDF, Filename:=paths(i), _
                Quality:=xlQualityStandard, IncludeDocProperties:=False, OpenAfterPublish:=False
            n = n + 1
        Next i
    Else
        For i = 1 To targets.Count
            PkgExport_OneSheetWorkbook targets(i).name, CStr(paths(i))
            n = n + 1
        Next i
    End If
    Core_SpeedOff

    Core_Done ctx, "Wrote " & n & " file(s) to:" & vbCrLf & "   " & folderPath, TITLE
    Exit Sub
Fail:
    If n > 0 Then
        Err.Description = Err.Description & vbCrLf & vbCrLf & n & " file(s) were already " & _
            "written to " & folderPath & " before this failed, and were NOT removed - this " & _
            "batch is not transactional. Check the folder before re-running."
    End If
    Core_Fail ctx, "export.per_sheet", Err
End Sub

'================================================================
' ACTION - Split a picked range into one values-only file per group
'================================================================
Public Sub PkgExport_SplitByValue(ctx As Object)
    Dim rng As Range, targetRange As Range
    Dim outWb As Workbook, outWs As Worksheet, openWb As Workbook
    Dim groups As Object, groupLabels As Object, seenPaths As Object
    Dim groupOrder As Collection, groupRows As Collection, paths As Collection
    Dim splitValue As Variant, sourceValues As Variant, outputValues As Variant
    Dim hasHeader As Boolean, keys As String
    Dim splitCol As Long, firstDataRow As Long, headerRows As Long
    Dim rowIndex As Long, colIndex As Long, outputIndex As Long
    Dim i As Long, j As Long, n As Long
    Dim folderPath As String, dateStr As String, groupKey As String
    Dim groupLabel As String, fileName As String, outPath As String
    Dim duplicateMsg As String, existing As String, listMsg As String
    Dim openBooks As String, openErr As String
    Dim errNum As Long, errDesc As String
    Const TITLE As String = "Save Each Group As Its Own File"
    On Error GoTo Fail

    folderPath = Core_ResolveFolder(Core_Cfg(ctx, "k_export_folder"))
    If Len(folderPath) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to " & _
               "export into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write " & _
               "exported files to' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_FolderExists(folderPath) Then
        Core_Refuse ctx, "This folder was not found: " & folderPath & vbCrLf & vbCrLf & _
               "Check 'Folder to write exported files to' on the Setup tab, Finalize, and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set rng = Core_AskRange(TITLE)
    If rng Is Nothing Then Exit Sub
    Set rng = Core_NarrowToUsed(rng)
    If rng Is Nothing Then
        Core_Refuse ctx, "The range you picked is entirely outside the used area of the " & _
               "sheet, so there is nothing in it to export." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Core_IsEngineSheet(rng.Parent.Name) Then
        Core_Refuse ctx, "'" & rng.Parent.Name & "' belongs to the macro toolkit itself, " & _
               "not to source data that can be split into exported files." & vbCrLf & _
               vbCrLf & "Pick a range on a non-toolkit sheet and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskHasHeader(rng, hasHeader, TITLE) Then Exit Sub
    firstDataRow = 1
    If hasHeader Then
        firstDataRow = 2
        headerRows = 1
    End If
    If rng.Rows.Count < firstDataRow Then
        Core_Refuse ctx, "The picked range contains a header but no data rows to split." & _
               vbCrLf & vbCrLf & "Pick a range that includes at least one data row and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(rng, keys, _
        "Choose exactly one column whose values should define the output groups.", _
        TITLE) Then Exit Sub
    splitCol = Core_ResolveColumn(rng, keys)
    If splitCol = 0 Then
        Core_Refuse ctx, "'" & keys & "' does not identify exactly one column inside the " & _
               "picked range." & vbCrLf & vbCrLf & "Run this again and choose one valid " & _
               "column name, letter, or position." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If rng.Cells.CountLarge = 1 Then
        ReDim sourceValues(1 To 1, 1 To 1)
        sourceValues(1, 1) = rng.Value2
    Else
        sourceValues = rng.Value2
    End If

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbBinaryCompare
    Set groupLabels = CreateObject("Scripting.Dictionary")
    groupLabels.CompareMode = vbBinaryCompare
    Set groupOrder = New Collection

    For rowIndex = firstDataRow To rng.Rows.Count
        splitValue = sourceValues(rowIndex, splitCol)
        If IsError(splitValue) Then
            Core_Refuse ctx, "The split column contains an error value at " & _
                   rng.Cells(rowIndex, splitCol).Address(False, False) & "." & _
                   vbCrLf & vbCrLf & "Fix that cell or choose another split column, then " & _
                   "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
            Exit Sub
        End If

        If Core_IsBlank(splitValue) Then
            groupKey = "Blank:"
            groupLabel = "(blank)"
        Else
            groupKey = TypeName(splitValue) & ":" & CStr(splitValue)
            groupLabel = CStr(splitValue)
        End If

        If Not groups.Exists(groupKey) Then
            Set groupRows = New Collection
            groups.Add groupKey, groupRows
            groupLabels.Add groupKey, groupLabel
            groupOrder.Add groupKey
        End If
        Set groupRows = groups.Item(groupKey)
        groupRows.Add rowIndex
    Next rowIndex

    dateStr = Format(Date + Core_CfgLong(ctx, "k_export_date_offset"), "yyyy-mm-dd")
    Set paths = New Collection
    Set seenPaths = CreateObject("Scripting.Dictionary")
    seenPaths.CompareMode = vbTextCompare

    For i = 1 To groupOrder.Count
        groupKey = CStr(groupOrder(i))
        groupLabel = CStr(groupLabels.Item(groupKey))
        fileName = Core_SafeFileName(dateStr & " " & groupLabel, 120) & ".xlsx"
        outPath = folderPath & fileName
        paths.Add outPath

        If seenPaths.Exists(fileName) Then
            duplicateMsg = duplicateMsg & vbCrLf & "   '" & groupLabel & "' and '" & _
                CStr(seenPaths.Item(fileName)) & "' both produce: " & fileName
        Else
            seenPaths.Add fileName, groupLabel
        End If

        If Len(Dir(outPath)) > 0 Then existing = existing & vbCrLf & "   " & outPath

        Set openWb = Nothing
        openErr = vbNullString
        Set openWb = PkgExport_ProbeOpen(outPath, openErr)
        If Len(openErr) > 0 Then
            openBooks = openBooks & vbCrLf & "   " & openErr
        ElseIf Not openWb Is Nothing Then
            openBooks = openBooks & vbCrLf & "   " & openWb.FullName
        End If
    Next i

    If Len(duplicateMsg) > 0 Then
        Core_Refuse ctx, "These distinct groups produce the same safe file name, so this " & _
               "batch would overwrite its own output:" & duplicateMsg & vbCrLf & vbCrLf & _
               "Change or shorten one of each pair of group values, then run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Len(openBooks) > 0 Then
        Core_Refuse ctx, "These intended output names are already open in Excel:" & _
               openBooks & vbCrLf & vbCrLf & "Close the named workbook(s), then run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Len(existing) > 0 Then
        Core_Refuse ctx, "These file(s) already exist and this action will not overwrite " & _
               "them:" & existing & vbCrLf & vbCrLf & "Rename or delete them, or change " & _
               "the folder or date offset on Setup, then run this again." & vbCrLf & _
               vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    For i = 1 To paths.Count
        listMsg = listMsg & vbCrLf & "   " & Core_FileName(CStr(paths(i)))
    Next i
    If MsgBox("This writes " & paths.Count & " values-only .xlsx file(s) to:" & vbCrLf & _
              "   " & folderPath & vbCrLf & vbCrLf & listMsg & vbCrLf & vbCrLf & _
              "The source range will not be changed. Continue?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    For i = 1 To groupOrder.Count
        groupKey = CStr(groupOrder(i))
        Set groupRows = groups.Item(groupKey)
        ReDim outputValues(1 To groupRows.Count + headerRows, 1 To rng.Columns.Count)

        If hasHeader Then
            For colIndex = 1 To rng.Columns.Count
                outputValues(1, colIndex) = sourceValues(1, colIndex)
            Next colIndex
        End If
        For j = 1 To groupRows.Count
            outputIndex = j + headerRows
            For colIndex = 1 To rng.Columns.Count
                outputValues(outputIndex, colIndex) = _
                    sourceValues(CLng(groupRows(j)), colIndex)
            Next colIndex
        Next j

        Set outWb = Workbooks.Add(xlWBATWorksheet)
        Set outWs = outWb.Worksheets(1)
        outWs.Name = "Data"
        Set targetRange = outWs.Cells(1, 1).Resize(groupRows.Count + headerRows, _
                                                     rng.Columns.Count)
        targetRange.Value2 = outputValues

        For colIndex = 1 To rng.Columns.Count
            outWs.Columns(colIndex).ColumnWidth = rng.Columns(colIndex).ColumnWidth
        Next colIndex

        Application.CutCopyMode = False
        outWb.SaveAs Filename:=CStr(paths(i)), FileFormat:=51
        n = n + 1
        outWb.Close SaveChanges:=False
        Set outWb = Nothing
    Next i
    Core_SpeedOff

    Core_Done ctx, "Wrote " & n & " group file(s) to:" & vbCrLf & _
           "   " & folderPath & vbCrLf & vbCrLf & _
           "The source range was not changed.", TITLE
    Exit Sub
Fail:
    If n > 0 Then
        Err.Description = Err.Description & vbCrLf & vbCrLf & n & " group file(s) were " & _
            "already written to " & folderPath & " before this failed. They were not " & _
            "removed; check the folder before re-running."
    End If
    errNum = Err.Number
    errDesc = Err.Description
    If Not outWb Is Nothing Then PkgExport_CloseNoSave outWb
    Err.Number = errNum
    Err.Description = errDesc
    Core_Fail ctx, "export.split_by_value", Err
End Sub

'================================================================
' ACTION - Write one values-only sheet and display it on an Outlook draft
'================================================================
Public Sub PkgExport_Email(ctx As Object)
    Dim recipient As String, subjectText As String
    Dim folderPath As String, outPath As String, dateStr As String
    Dim sheetName As Variant, pickedSheet As Object, ws As Worksheet
    Dim outlookApp As Object, mailItem As Object
    Dim openWb As Workbook, openErr As String
    Dim attachmentWritten As Boolean
    Const TITLE As String = "Email A Sheet As An Attachment"
    On Error GoTo Fail

    recipient = Trim(Core_Cfg(ctx, "k_export_email_to"))
    If Len(recipient) = 0 Then
        Core_Refuse ctx, "No email recipient is configured." & vbCrLf & vbCrLf & _
               "Fill in 'Email recipient' on the Setup tab, Finalize, and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    subjectText = Trim(Core_Cfg(ctx, "k_export_email_subject"))
    If Len(subjectText) = 0 Then
        Core_Refuse ctx, "No email subject is configured." & vbCrLf & vbCrLf & _
               "Fill in 'Email subject' on the Setup tab, Finalize, and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    folderPath = Core_ResolveFolder(Core_Cfg(ctx, "k_export_folder"))
    If Len(folderPath) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to " & _
               "export into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write " & _
               "exported files to' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not Core_FolderExists(folderPath) Then
        Core_Refuse ctx, "This folder was not found: " & folderPath & vbCrLf & vbCrLf & _
               "Check 'Folder to write exported files to' on the Setup tab, Finalize, and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    sheetName = Application.InputBox( _
        Prompt:="Sheet to export and attach:", _
        title:=TITLE, Default:=ActiveSheet.Name, Type:=2)
    If VarType(sheetName) = vbBoolean Then Exit Sub

    If Not Core_SheetExists(CStr(sheetName)) Then
        Core_Refuse ctx, "There is no sheet named '" & CStr(sheetName) & "'." & _
               vbCrLf & vbCrLf & "Choose an existing worksheet and run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set pickedSheet = ThisWorkbook.Sheets(CStr(sheetName))
    If Not TypeOf pickedSheet Is Worksheet Then
        Core_Refuse ctx, "'" & CStr(sheetName) & "' is not a worksheet and cannot be " & _
               "exported as an .xlsx attachment." & vbCrLf & vbCrLf & _
               "Choose a worksheet and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    Set ws = pickedSheet

    If Core_IsEngineSheet(ws.Name) Then
        Core_Refuse ctx, "'" & ws.Name & "' belongs to the macro toolkit itself, not to " & _
               "a report that can be emailed." & vbCrLf & vbCrLf & _
               "Choose a non-toolkit worksheet and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    dateStr = Format(Date + Core_CfgLong(ctx, "k_export_date_offset"), "yyyy-mm-dd")
    outPath = folderPath & Core_SafeFileName(dateStr & " " & ws.Name, 120) & ".xlsx"

    Set openWb = PkgExport_ProbeOpen(outPath, openErr)
    If Len(openErr) > 0 Then
        Core_Refuse ctx, "Excel already has a workbook open under the intended attachment " & _
               "name:" & vbCrLf & vbCrLf & "   " & openErr & vbCrLf & vbCrLf & _
               "Close that workbook, then run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If
    If Not openWb Is Nothing Then
        Core_Refuse ctx, "The intended attachment workbook is already open in Excel:" & _
               vbCrLf & vbCrLf & "   " & openWb.FullName & vbCrLf & vbCrLf & _
               "Close it, then run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Len(Dir(outPath)) > 0 Then
        Core_Refuse ctx, "A file already exists at:" & vbCrLf & vbCrLf & _
               "   " & outPath & vbCrLf & vbCrLf & _
               "This action will not overwrite it. Rename or delete that file, or change " & _
               "the folder or date offset on Setup, then run this again." & vbCrLf & _
               vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If MsgBox("This first writes the attachment file to:" & vbCrLf & vbCrLf & _
              "   " & outPath & vbCrLf & vbCrLf & _
              "It then opens an Outlook draft addressed to " & recipient & "." & _
              vbCrLf & vbCrLf & "The draft will be displayed and will not be sent. Continue?", _
              vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Core_SpeedOn
    PkgExport_OneSheetWorkbook ws.Name, outPath
    attachmentWritten = True
    Core_SpeedOff

    Set outlookApp = Nothing
    On Error Resume Next
    Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo Fail

    If outlookApp Is Nothing Then
        Err.Clear
        Core_Refuse ctx, "The attachment file was written to:" & vbCrLf & vbCrLf & _
               "   " & outPath & vbCrLf & vbCrLf & _
               "Outlook could not be reached. It may not be installed, or access may be " & _
               "blocked by policy. Attach this file by hand instead." & vbCrLf & vbCrLf & _
               "The file above was kept so you can attach it yourself. Nothing was emailed.", _
               TITLE
        Exit Sub
    End If

    Set mailItem = outlookApp.CreateItem(0)
    With mailItem
        .To = recipient
        .Subject = subjectText
        .Attachments.Add outPath
        .Display
    End With

    Core_Done ctx, "Wrote the attachment to:" & vbCrLf & vbCrLf & _
           outPath & vbCrLf & vbCrLf & _
           "An Outlook draft has been displayed. It has not been sent.", TITLE
    Exit Sub
Fail:
    If attachmentWritten Then
        Err.Description = Err.Description & vbCrLf & vbCrLf & _
            "The attachment file was already written and was not deleted. Its full path is:" & _
            vbCrLf & outPath
    End If
    Core_Fail ctx, "export.email", Err
End Sub

'================================================================
' Write ONE sheet out as its own .xlsx. Copies the WHOLE workbook first
' and freezes every formula in the copy (Core_FreezeValues) before
' deleting every sheet but the target - exactly the technique
' importexport.export already relies on. A sheet copied out ALONE, with
' its cross-sheet formulas still live, would show #REF! in every cell
' that pointed at a sheet no longer there; freezing first is what
' importexport.export already proved works, so this reuses it rather
' than risking a cleverer, untested shortcut.
'================================================================
Private Sub PkgExport_OneSheetWorkbook(ByVal sheetName As String, ByVal outPath As String)
    Dim newWb As Workbook, ws As Worksheet, i As Long
    Dim savedCalc As XlCalculation, errNum As Long, errDesc As String
    On Error GoTo CleanFail

    savedCalc = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

    ThisWorkbook.Worksheets.Copy
    Set newWb = Application.ActiveWorkbook
    newWb.Application.CalculateFull

    ' reevalUsedRange:=True - this is a whole-sheet flatten, same as
    ' importexport.export's own call (Task 2, Step 3); see
    ' Core_FreezeValues's header, difference #2, for why.
    For Each ws In newWb.Worksheets
        Core_FreezeValues ws.UsedRange, True
    Next ws

    For i = newWb.Worksheets.Count To 1 Step -1
        If StrComp(newWb.Worksheets(i).name, sheetName, vbTextCompare) <> 0 Then
            newWb.Worksheets(i).Visible = xlSheetVisible
            newWb.Worksheets(i).Delete
        End If
    Next i

    newWb.SaveAs Filename:=outPath, FileFormat:=51        ' 51 = .xlsx, no macros
    newWb.Close SaveChanges:=False
    Application.Calculation = savedCalc
    Exit Sub

CleanFail:
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    Application.Calculation = savedCalc
    On Error GoTo 0
    Err.Raise errNum, "PkgExport_OneSheetWorkbook", errDesc
End Sub

Private Sub PkgExport_CloseNoSave(ByVal wb As Workbook)
    On Error Resume Next
    wb.Close SaveChanges:=False
End Sub

Private Function PkgExport_ProbeOpen(ByVal outPath As String, _
                                     ByRef collision As String) As Workbook
    collision = vbNullString
    On Error GoTo NameCollision
    Set PkgExport_ProbeOpen = Core_AlreadyOpen(outPath)
    Exit Function
NameCollision:
    collision = Err.Description
    Err.Clear
End Function
