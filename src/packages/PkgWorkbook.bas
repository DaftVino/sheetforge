Attribute VB_Name = "PkgWorkbook"
Option Explicit

'================================================================
' PkgWorkbook - housekeeping on the workbook itself: sheet
' visibility, protection, and navigation. Wave 6 of the catalog
' build-out (docs/designs/2026-07-28-w6-workbook-plan.md).
'
' Five actions from FOUR catalog items - "Protect / unprotect all
' sheets at once" is one catalog item that deliberately produces
' two ids, because a single toggle button hides which way it will
' go (docs/macro-catalog-plug-and-play.md line 199):
'   workbook.unhide_sheets       every worksheet visible again
'   workbook.protect_sheets      lock every sheet this deployment owns
'   workbook.unprotect_sheets    unlock every sheet this deployment owns
'   workbook.build_toc           an index sheet linking to every visible tab
'   workbook.unhide_rows_cols    reset row/column visibility, active sheet
'
' Every action here walks all sheets or changes sheet state, and the
' engine owns several sheets outright. Every action excludes them
' via Core_IsEngineSheet rather than a local copy of their names -
' see the plan's Decisions table for why each exclusion exists.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgWorkbook_Manifest(reg As Object)
    Core_DeclarePackage reg, "workbook", "Workbook Housekeeping", "1.0.0"

    Core_AddSection reg, "TABLE OF CONTENTS"
    Core_AddField reg, "k_workbook_toc_name", "Name for the Table of Contents sheet", "text", "Contents", _
        "simple", _
        "A sheet by this name is created the first run. On a later run, if a sheet by this " & _
        "name already exists - including this action's own sheet from a prior run - the " & _
        "action stops and asks for a different name; it never overwrites an existing sheet, " & _
        "so to rebuild the Table of Contents, delete the old one by hand first, then run this " & _
        "again. The same refusal applies to a name the toolkit itself reserves (Setup, the " & _
        "config sheets, or a generated Macros page). Kind is text, not sheet, because a sheet " & _
        "field is validated to already EXIST and this one must not.", _
        "workbook.build_toc"

    Core_AddSection reg, "PROTECTION"
    Core_AddField reg, "k_workbook_password", "Password for Protect / Unprotect All Sheets", "text", "", _
        "advanced", _
        "Leave blank to protect with no password. This is stored in plain text on the hidden " & _
        "config sheet, the same as every other Setup answer - it is not a secret, and storing " & _
        "it here does not make a sheet's contents confidential. It only stops accidental edits " & _
        "from the ribbon, the same protection Excel's own Protect Sheet provides on its own.", _
        "workbook.protect_sheets,workbook.unprotect_sheets"

    Core_AddAction reg, "workbook.unhide_sheets", "Unhide All Worksheets", "PkgWorkbook_UnhideSheets", _
        "Makes every hidden or very-hidden worksheet visible again. Setup and the toolkit's " & _
        "own internal sheets - including the hidden config store - are never touched.", _
        "Workbook Housekeeping"

    Core_AddAction reg, "workbook.protect_sheets", "Protect All Sheets", "PkgWorkbook_ProtectSheets", _
        "Locks every sheet you own against editing, with the password above (or no password " & _
        "if it is blank). Setup and the toolkit's own internal sheets are left unprotected, so " & _
        "a later rebuild can still manage its own buttons on them.", _
        "Workbook Housekeeping"

    Core_AddAction reg, "workbook.unprotect_sheets", "Unprotect All Sheets", "PkgWorkbook_UnprotectSheets", _
        "Unlocks every sheet you own, using the password above. A sheet locked with a " & _
        "different password is reported by name rather than stopping the rest.", _
        "Workbook Housekeeping"

    Core_AddAction reg, "workbook.build_toc", "Build Table of Contents", "PkgWorkbook_BuildToc", _
        "Creates one sheet listing every visible tab, each linked to jump straight to it. " & _
        "Refuses if a sheet by that name already exists - delete it by hand first to rebuild.", _
        "Workbook Housekeeping"

    Core_AddAction reg, "workbook.unhide_rows_cols", "Unhide All Rows & Columns", "PkgWorkbook_UnhideRowsCols", _
        "Resets row and column visibility on the active sheet. Warns first if an AutoFilter " & _
        "on that sheet would also lose its filtered-out rows.", _
        "Workbook Housekeeping"
End Sub

'================================================================
' ACTION - Unhide every worksheet, except the toolkit's own
'================================================================
Public Sub PkgWorkbook_UnhideSheets(ctx As Object)
    Dim ws As Worksheet, n As Long
    Const TITLE As String = "Unhide All Worksheets"
    On Error GoTo Fail

    Core_SpeedOn
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) Then
            If ws.Visible <> xlSheetVisible Then
                ws.Visible = xlSheetVisible
                n = n + 1
            End If
        End If
    Next ws
    Core_SpeedOff

    ' The toolkit's own sheets - Setup, the hidden config store and its staging
    ' sheet, the package-enable sheet, every generated Macros page - are hidden
    ' or shown for the toolkit's own reasons and are excluded unconditionally,
    ' the same exclusion every other action in this file uses. Unhiding _Config
    ' specifically would put a finance user in front of the engine's own working
    ' state and invite them to edit it - the single most important correctness
    ' point in this action.
    If n = 0 Then
        Core_Done ctx, "Nothing was hidden. Every sheet you own is already visible." & vbCrLf & vbCrLf & _
               "The toolkit's own internal sheets are never touched by this action, hidden or " & _
               "not.", TITLE
    Else
        Core_Done ctx, "Unhid " & n & " sheet(s)." & vbCrLf & vbCrLf & _
               "The toolkit's own internal sheets - Setup, the config store, and any generated " & _
               "Macros page - were left exactly as they were.", TITLE
    End If
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.unhide_sheets", Err
End Sub

'================================================================
' ACTION - Protect every sheet this deployment owns
'================================================================
Public Sub PkgWorkbook_ProtectSheets(ctx As Object)
    Dim targets As Collection, ws As Worksheet, pwd As String
    Dim doneCount As Long, failed As Collection, msg As String
    Const TITLE As String = "Protect All Sheets"
    On Error GoTo Fail

    Set targets = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) Then targets.Add ws
    Next ws
    If targets.Count = 0 Then
        Core_Refuse ctx, "There are no sheets to protect - every sheet in this workbook " & _
               "belongs to the toolkit itself." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    pwd = Core_Cfg(ctx, "k_workbook_password")
    msg = "This will protect " & targets.Count & " sheet(s) against editing" & _
          IIf(Len(pwd) = 0, ", with no password", ", with the password set on Setup") & "." & _
          vbCrLf & vbCrLf & _
          "The toolkit's own sheets - Setup, the config store, and any generated Macros page - " & _
          "are never protected by this action, so a later rebuild can still update them: " & _
          "protecting a generated page with Excel's own default would block the object-model " & _
          "deletion a rebuild needs to remove its own buttons." & vbCrLf & vbCrLf & "Continue?"
    If MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub

    Set failed = New Collection
    Core_SpeedOn
    For Each ws In targets
        On Error Resume Next
        Err.Clear
        ' DrawingObjects:=False, deliberately NOT Excel's own default (True). A
        ' sheet protected with DrawingObjects:=True blocks a future rebuild from
        ' removing its own anchored controls on it - and contract section 6
        ' already forbids the engine unprotecting a user's sheet to fix that
        ' itself. This is the one moment a PERSON, not the engine, is choosing to
        ' protect, so it is the one place that choice can still be made safely.
        ws.Protect Password:=pwd, DrawingObjects:=False, Contents:=True, Scenarios:=True
        If Err.Number <> 0 Then
            failed.Add ws.name
            Err.Clear
        Else
            doneCount = doneCount + 1
        End If
        ' Restore the procedure's own handler, NOT "On Error GoTo 0". GoTo 0
        ' would disable error handling entirely for the rest of this Sub - so a
        ' later, genuinely unexpected error (not this per-sheet Protect failure,
        ' which is already caught above) would bypass Core_Fail and leave
        ' Core_SpeedOn's ScreenUpdating/Events/Calculation changes unrestored,
        ' the same failure Core_Fail's own handler exists to prevent.
        On Error GoTo Fail
    Next ws
    Core_SpeedOff

    Core_Done ctx, "Protected " & doneCount & " of " & targets.Count & " sheet(s)." & _
           PkgWorkbook_FailedNote(failed, "already protected, or otherwise could not be locked"), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.protect_sheets", Err
End Sub

'================================================================
' Package-private helpers. Prefixed PkgWorkbook_ because the bundle
' is one flat namespace, even though these are Private.
'================================================================

' Names which target sheets did not get the requested change, so a partial
' result is STATED rather than silently rounded up to "done". Returns "" when
' every target succeeded, so a caller can test Len() rather than compare text.
' Shared by protect_sheets and unprotect_sheets - both hit the same shape of
' partial failure (one sheet's password or state does not cooperate) and both
' need the same "here is exactly what did not happen" honesty.
Private Function PkgWorkbook_FailedNote(failed As Collection, ByVal reason As String) As String
    Dim i As Long, names As String
    If failed.Count = 0 Then Exit Function
    For i = 1 To failed.Count
        If i > 1 Then names = names & ", "
        names = names & failed(i)
    Next i
    PkgWorkbook_FailedNote = vbCrLf & vbCrLf & failed.Count & " sheet(s) were skipped (" & _
        reason & "): " & names
End Function

'================================================================
' ACTION - Unprotect every sheet this deployment owns
'================================================================
Public Sub PkgWorkbook_UnprotectSheets(ctx As Object)
    Dim ws As Worksheet, pwd As String
    Dim doneCount As Long, failed As Collection
    Const TITLE As String = "Unprotect All Sheets"
    On Error GoTo Fail

    pwd = Core_Cfg(ctx, "k_workbook_password")
    Set failed = New Collection
    Core_SpeedOn
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) Then
            If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
                On Error Resume Next
                Err.Clear
                ' Password:=pwd is passed EXPLICITLY, even when pwd is "". Calling
                ' Unprotect with NO argument shows Excel's OWN interactive
                ' password box when the sheet carries one - which would surface
                ' mid-chain, breaking the quiet convention (contract section 6).
                ' An explicit argument, blank or not, never prompts: a mismatch
                ' raises error 1004 instead, which this catches per sheet.
                ' Assumption about Excel's own behaviour - see the plan's
                ' Unknowns.
                ws.Unprotect Password:=pwd
                If Err.Number <> 0 Then
                    failed.Add ws.name
                    Err.Clear
                Else
                    doneCount = doneCount + 1
                End If
                ' Restore the procedure's own handler, NOT "On Error GoTo 0" -
                ' see the identical comment in workbook.protect_sheets (Task 3)
                ' for why GoTo 0 here would leave a later real error unhandled.
                On Error GoTo Fail
            End If
        End If
    Next ws
    Core_SpeedOff

    If doneCount = 0 And failed.Count = 0 Then
        Core_Done ctx, "Nothing was protected, so there was nothing to unprotect.", TITLE
    Else
        Core_Done ctx, "Unprotected " & doneCount & " sheet(s)." & _
               PkgWorkbook_FailedNote(failed, "the password did not match"), TITLE
    End If
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.unprotect_sheets", Err
End Sub

'================================================================
' ACTION - Build a Table of Contents sheet
'================================================================
Public Sub PkgWorkbook_BuildToc(ctx As Object)
    Dim destName As String, ws As Worksheet, dest As Worksheet
    Dim rowOut As Long, n As Long
    Const TITLE As String = "Build Table of Contents"
    On Error GoTo Fail

    destName = Trim(Core_Cfg(ctx, "k_workbook_toc_name"))
    If Len(destName) = 0 Then
        Core_Refuse ctx, "No name is set for the Table of Contents sheet." & vbCrLf & vbCrLf & _
               "Fill 'Name for the Table of Contents sheet' in on the Setup tab, Finalize, and " & _
               "run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Core_DestUsable is called unconditionally, on EVERY run, whether or not a
    ' sheet by this name currently exists - it refuses both a name in the
    ' reserved family (Core_IsEngineSheet, checked even when no sheet by that
    ' name exists yet - see the plan's Decision 8) and any existing sheet,
    ' INCLUDING this action's own sheet from a prior run (Decision 7: there is
    ' no reliable way to tell "my own prior output" apart from "a user's sheet
    ' that happens to share the name" without adding new engine state, so this
    ' action does not try - it refuses either way, the same as any other
    ' destination-name field). An earlier draft special-cased an existing
    ' sheet onto a refresh-in-place branch and called Core_DestUsable only on
    ' the create branch; that let the reserved-name check be skipped whenever
    ' a sheet already occupied the name. Calling it first, before any branch,
    ' means there is no branch left for that check to miss.
    If Not Core_DestUsable(ctx, destName, TITLE) Then Exit Sub

    Core_SpeedOn
    Set dest = Core_NewSheet(destName)
    dest.Range("A1").Value = "Table of Contents"
    dest.Range("A1").Font.Bold = True
    rowOut = 2
    For Each ws In ThisWorkbook.Worksheets
        If Not Core_IsEngineSheet(ws.name) And StrComp(ws.name, destName, vbTextCompare) <> 0 Then
            ' A hidden or very-hidden sheet is skipped, not merely un-linked: a
            ' hyperlink's SubAddress cannot navigate to a hidden sheet, so a
            ' link to one is a dead click that reads as a bug in this action
            ' rather than as the sheet being hidden on purpose.
            If ws.Visible = xlSheetVisible Then
                ' An embedded apostrophe in a sheet name must be DOUBLED inside
                ' the quoted SubAddress - the same doubling rule spec/contract.md
                ' section 6 already documents for anchored placements. An
                ' un-doubled apostrophe here breaks the quoting and the link
                ' either fails to create or points at the wrong sheet.
                dest.Hyperlinks.Add Anchor:=dest.Cells(rowOut, 1), Address:="", _
                    SubAddress:="'" & Replace(ws.name, "'", "''") & "'!A1", TextToDisplay:=ws.name
                rowOut = rowOut + 1
                n = n + 1
            End If
        End If
    Next ws
    dest.Columns("A").AutoFit
    Core_SpeedOff

    If n = 0 Then
        Core_Done ctx, "Built '" & destName & "', but there were no other sheets to list - " & _
               "every other sheet is either hidden or belongs to the toolkit itself.", TITLE
    Else
        Core_Done ctx, "Built '" & destName & "' with " & n & " link(s)." & vbCrLf & vbCrLf & _
               "A hidden sheet is not listed - linking to one that cannot be navigated to would " & _
               "defeat the point of hiding it. This action does not refresh an existing sheet: " & _
               "to rebuild after adding, removing, or renaming a tab, delete '" & destName & _
               "' by hand first, then run this again.", TITLE
    End If
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.build_toc", Err
End Sub

'================================================================
' ACTION - Reset row/column visibility on the active sheet
'================================================================
Public Sub PkgWorkbook_UnhideRowsCols(ctx As Object)
    Dim ws As Worksheet
    Dim rowsHidden As Long, colsHidden As Long
    Const TITLE As String = "Unhide All Rows & Columns"
    On Error GoTo Fail

    ' TypeName is checked BEFORE assigning to a Worksheet-typed variable. A
    ' chart sheet (or any other non-worksheet sheet type) as the active sheet
    ' would make "Set ws = ActiveSheet" itself raise a generic type-mismatch
    ' error - a confusing crash rather than the plain refusal this action
    ' gives every other unsupported case.
    If TypeName(ActiveSheet) <> "Worksheet" Then
        Core_Refuse ctx, "The active sheet is not a worksheet (for example, a chart sheet), " & _
               "so there are no rows or columns to reset." & vbCrLf & vbCrLf & _
               "Switch to the worksheet you want to reset, then run this again." & vbCrLf & _
               vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If
    Set ws = ActiveSheet
    If Core_IsEngineSheet(ws.name) Then
        Core_Refuse ctx, "'" & ws.name & "' is one of the toolkit's own sheets, not a work " & _
               "sheet." & vbCrLf & vbCrLf & _
               "Some of the toolkit's own rows and columns are hidden on purpose, as part of " & _
               "its own layout - unhiding them here would expose the toolkit's internal " & _
               "working state, not something you can use." & vbCrLf & vbCrLf & _
               "Switch to the sheet you want to reset, then run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Counted across the sheet's USED range only, for the same reason
    ' Core_NarrowToUsed narrows a pick there: a per-row/per-column COM read
    ' across a whole 1,048,576-row sheet freezes the screen for no benefit.
    ' The unhide itself still clears the WHOLE sheet below, so a hidden column
    ' sitting outside the used range is still unhidden - it is just not
    ' counted in the number reported here.
    rowsHidden = PkgWorkbook_CountHidden(ws.UsedRange.Rows)
    colsHidden = PkgWorkbook_CountHidden(ws.UsedRange.Columns)
    If rowsHidden = 0 And colsHidden = 0 Then
        Core_Done ctx, "Nothing was hidden on '" & ws.name & "'.", TITLE
        Exit Sub
    End If

    ' An AutoFilter hides rows through the SAME Range.Hidden property as a
    ' manual Hide - there is no way to tell the two apart, and so no way to
    ' unhide only one kind. Warned explicitly, and only here, rather than
    ' silently changing what the user sees as "the filtered view": without an
    ' active filter, unhiding is not a surprise and this action just does it,
    ' the same as workbook.unhide_sheets.
    If ws.AutoFilterMode Then
        If MsgBox("'" & ws.name & "' has an active AutoFilter." & vbCrLf & vbCrLf & _
               "Unhiding rows will make every filtered-out row visible again - it will look " & _
               "like the filter changed, because the filter's own hidden rows are reset along " & _
               "with any manually hidden ones. There is no way to unhide one kind and not the " & _
               "other." & vbCrLf & vbCrLf & "Re-apply the filter afterward if you want it back." & _
               vbCrLf & vbCrLf & "Continue?", _
               vbExclamation + vbYesNo + vbDefaultButton2, TITLE) <> vbYes Then Exit Sub
    End If

    Core_SpeedOn
    ws.Rows.EntireRow.Hidden = False
    ws.Columns.EntireColumn.Hidden = False
    Core_SpeedOff

    Core_Done ctx, "Unhid " & rowsHidden & " row(s) and " & colsHidden & " column(s) on '" & _
           ws.name & "'.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "workbook.unhide_rows_cols", Err
End Sub

' Counts hidden rows or columns inside a Rows/Columns collection, one COM read
' per entry. Bounded to the caller's used range - see the comment at the call
' site for why the whole sheet is never walked this way.
Private Function PkgWorkbook_CountHidden(rng As Range) As Long
    Dim one As Range, n As Long
    For Each one In rng
        If one.Hidden Then n = n + 1
    Next one
    PkgWorkbook_CountHidden = n
End Function
