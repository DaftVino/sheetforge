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
        "answer. Otherwise a full path, e.g. C:\Close\January.", _
        "export.pdf,export.per_sheet"
    Core_AddField reg, "k_export_date_offset", "Date shift for exported file names (days)", "int", "-1", _
        "advanced", _
        "How many days to shift the date put in front of an exported file's name. -1 = " & _
        "yesterday, 0 = today, 1 = tomorrow.", _
        "export.pdf,export.per_sheet"

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
