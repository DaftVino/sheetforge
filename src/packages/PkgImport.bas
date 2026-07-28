Attribute VB_Name = "PkgImport"
Option Explicit

'================================================================
' PkgImport - bringing outside data in. Wave 1 of the catalog
' build-out (docs/designs/2026-07-24-finish-excel-design.md 3.4).
'
' Three actions, two of which are the same job with a different
' source list, so both delegate to ONE private engine:
'   import.combine_worksheets   every data sheet in THIS workbook
'   import.combine_workbooks    the first sheet of every workbook in a folder
'   import.delimited_file       one CSV/TXT, parsed explicitly
'
' The domain "import" is a USER-FACING vocabulary, not this module's
' name. The legacy importexport.import action keeps its id and stays
' in PkgImportExport; both declare the category "Import / Export" so
' the user sees one card. See contract section 7.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgImport_Manifest(reg As Object)
    Core_DeclarePackage reg, "import", "Bring Data In", "1.0.0"

    Core_AddSection reg, "COMBINE"
    Core_AddField reg, "k_import_dest", "Name for the combined sheet this creates", "text", "Master", _
        "simple", _
        "A NEW sheet by this name is created each run. If a sheet by this name already exists, the action stops and asks you to rename or remove it - it never overwrites one.", _
        "import.combine_worksheets,import.combine_workbooks"
    Core_AddField reg, "k_import_folder", "Folder to read workbooks from", "text", "", _
        "simple", _
        "Leave blank to use the folder this workbook is saved in - that is the usual answer. Otherwise a full path, e.g. C:\Close\January.", _
        "import.combine_workbooks"

    Core_AddSection reg, "DELIMITED FILES"
    Core_AddField reg, "k_import_delim", "Column separator in CSV / TXT files", "text", "comma", _
        "advanced", _
        "One of: comma, tab, semicolon, pipe, space. Or type the separator character itself. Blank means comma.", _
        "import.delimited_file"
    Core_AddField reg, "k_import_qualifier", "Quote character around text", "text", """", _
        "advanced", _
        "The character wrapping text that itself contains the separator. Type a double quote, a single quote, or the word none. Leaving this blank means a double quote, not none - type none to turn quoting off.", _
        "import.delimited_file"

    Core_AddAction reg, "import.combine_worksheets", "Combine All Worksheets", "PkgImport_CombineWorksheets", _
        "Stacks every data sheet in this workbook into one new sheet, lining up columns by their header names. Your Setup and Macros tabs are left out.", _
        "Import / Export"
    Core_AddAction reg, "import.combine_workbooks", "Combine All Workbooks in a Folder", "PkgImport_CombineWorkbooks", _
        "Stacks the first sheet of every workbook in the folder into one new sheet, lining up columns by their header names. Each file is opened read-only and closed without saving.", _
        "Import / Export"
    Core_AddAction reg, "import.delimited_file", "Import a Delimited File", "PkgImport_DelimitedFile", _
        "Pick a CSV or TXT file; it is parsed with the separator and quote character from Setup rather than letting Excel guess, and written to a new sheet named after the file.", _
        "Import / Export"
End Sub

'================================================================
' ACTION - Combine every data sheet in this workbook
'================================================================
Public Sub PkgImport_CombineWorksheets(ctx As Object)
    Dim destName As String, sources As Collection, ws As Worksheet
    Dim rowsOut As Long, sourcesOut As Long
    Const TITLE As String = "Combine All Worksheets"
    On Error GoTo Fail

    ' The destination check runs FIRST, before any prompt or any work: the
    ' cheapest failure should be the earliest one, and a refusal after the user
    ' has answered questions reads as a bug.
    destName = Trim(Core_Cfg(ctx, "k_import_dest"))
    If Not PkgImport_DestUsable(destName, TITLE) Then Exit Sub

    Set sources = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If Not PkgImport_IsEngineSheet(ws.name) Then
            If StrComp(ws.name, destName, vbTextCompare) <> 0 Then sources.Add ws
        End If
    Next ws
    If sources.Count = 0 Then
        MsgBox "There are no data sheets to combine in this workbook." & vbCrLf & vbCrLf & _
               "Every sheet here belongs to the macro toolkit itself." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgImport_CombineEngine sources, destName, rowsOut, sourcesOut
    Core_SpeedOff

    Core_Done ctx, "Combined " & sourcesOut & " sheet(s) into '" & destName & "'." & vbCrLf & _
           rowsOut & " data row(s) written." & vbCrLf & vbCrLf & _
           "Columns were lined up by their header names, so a column that sits in a " & _
           "different position on different sheets still landed in one column here.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "import.combine_worksheets", Err
End Sub

'================================================================
' ACTION - Combine the first sheet of every workbook in a folder
'================================================================
Public Sub PkgImport_CombineWorkbooks(ctx As Object)
    Dim destName As String, folderPath As String, sources As Collection
    Dim rowsOut As Long, sourcesOut As Long
    Const TITLE As String = "Combine All Workbooks in a Folder"
    On Error GoTo Fail

    destName = Trim(Core_Cfg(ctx, "k_import_dest"))
    If Not PkgImport_DestUsable(destName, TITLE) Then Exit Sub

    folderPath = PkgImport_ResolveFolder(Core_Cfg(ctx, "k_import_folder"))
    If Len(folderPath) = 0 Then
        ' Blank config AND an unsaved workbook: there is no "this workbook's folder" yet.
        ' Caught explicitly, because otherwise it reaches Dir() as an empty path and
        ' surfaces as an engine error rather than the one-sentence fix it actually is.
        MsgBox "This workbook has never been saved, so it has no folder to read from." & _
               vbCrLf & vbCrLf & "Save it first, or set 'Folder to read workbooks from' " & _
               "on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If
    If Not PkgImport_FolderExists(folderPath) Then
        MsgBox "This folder was not found:" & vbCrLf & vbCrLf & "   " & folderPath & vbCrLf & vbCrLf & _
               "Check 'Folder to read workbooks from' on the Setup tab, Finalize, and run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    Set sources = PkgImport_WorkbookPaths(folderPath)
    If sources.Count = 0 Then
        MsgBox "No workbooks were found in:" & vbCrLf & vbCrLf & "   " & folderPath & vbCrLf & vbCrLf & _
               "This workbook itself is always skipped." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgImport_CombineEngine sources, destName, rowsOut, sourcesOut
    Core_SpeedOff

    Core_Done ctx, "Combined " & sourcesOut & " workbook(s) into '" & destName & "'." & vbCrLf & _
           rowsOut & " data row(s) written, read from:" & vbCrLf & "   " & folderPath & vbCrLf & vbCrLf & _
           "Columns were lined up by their header names. Every file was opened read-only " & _
           "and closed without saving.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "import.combine_workbooks", Err
End Sub

'================================================================
' The shared combine. sources holds either Worksheet objects (open, in
' this workbook) or String paths to workbooks this engine opens and
' closes one at a time. That is the ONLY difference between the two
' combine actions - design section 3.4.
'
' Columns are aligned BY HEADER TEXT, not by position: a source's
' "Amount" column lands under the master "Amount" wherever it sits, and
' a header no earlier source declared is appended as a new column on
' the right. Row 1 of each source is its header and is never repeated.
'
' On ANY error it deletes the half-built destination sheet and re-raises,
' so a failure leaves the workbook as it found it. That cleanup is the
' reason this has its own handler rather than letting the caller's
' Core_Fail see the error directly.
'================================================================
Private Sub PkgImport_CombineEngine(sources As Collection, ByVal destName As String, _
                                    ByRef rowsOut As Long, ByRef sourcesOut As Long)
    Dim dest As Worksheet, src As Worksheet, wbSrc As Workbook
    Dim one As Variant, headerMap As Object
    Dim destRow As Long, destCols As Long
    Dim errNum As Long, errDesc As String
    On Error GoTo EngineFail

    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = 1                    ' vbTextCompare: "Amount" matches "amount"
    Set dest = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    dest.name = destName
    destRow = 2                                  ' row 1 is the master header, grown as we learn it
    destCols = 0

    For Each one In sources
        Set src = Nothing
        Set wbSrc = Nothing
        If VarType(one) = vbString Then
            Set wbSrc = Workbooks.Open(Filename:=CStr(one), ReadOnly:=True, UpdateLinks:=0)
            Set src = wbSrc.Worksheets(1)
        Else
            Set src = one
        End If

        ' Count only sources that actually contributed rows. Reporting "combined 9
        ' sheets" when four were empty is a wrong number in the one dialog the user
        ' reads to decide whether it worked.
        If PkgImport_AppendSheet(src, dest, headerMap, destRow, destCols) Then
            sourcesOut = sourcesOut + 1
        End If

        If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Next one

    rowsOut = destRow - 2
    Exit Sub

EngineFail:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    ' DisplayAlerts is already False under Core_SpeedOn, so this does not prompt;
    ' Core_Fail's RestoreSpeed puts it back either way.
    If Not dest Is Nothing Then dest.Delete
    On Error GoTo 0
    Err.Raise errNum, "PkgImport_CombineEngine", errDesc
End Sub

' Append one source sheet under the growing master header.
' Writes ONE column block per source column rather than cell by cell - a
' cell-at-a-time loop over a real close file is minutes, not seconds, even
' with screen updating off.
Private Function PkgImport_AppendSheet(src As Worksheet, dest As Worksheet, headerMap As Object, _
                                       ByRef destRow As Long, ByRef destCols As Long) As Boolean
    Dim lastR As Long, lastC As Long, c As Long, r As Long, target As Long
    Dim hdr As String, vals As Variant, colVals() As Variant
    Dim ur As Range, seenHere As Object

    ' Per-SOURCE header tally, not the cross-source map: two sheets each having one
    ' "Amount" column must still align onto one destination column.
    Set seenHere = CreateObject("Scripting.Dictionary")
    seenHere.CompareMode = 1

    ' Extent comes from the WHOLE sheet, not from column A and row 1. A source whose
    ' first column has a gap would be truncated by End(xlUp) on column A, and a trailing
    ' column with a blank header would be dropped by End(xlToLeft) on row 1 - silently,
    ' and the completion dialog would still report success. UsedRange overstates on a
    ' sheet with stray formatting; overstating costs blank rows, understating loses data.
    Set ur = src.UsedRange
    If ur Is Nothing Then Exit Function
    lastR = ur.Row + ur.Rows.Count - 1
    lastC = ur.Column + ur.Columns.Count - 1
    If lastR < 2 Or lastC < 1 Then Exit Function     ' empty, or a header with no data

    vals = src.Range(src.Cells(1, 1), src.Cells(lastR, lastC)).Value

    ReDim colVals(1 To lastR - 1, 1 To 1)
    For c = 1 To lastC
        hdr = Trim(CStr(vals(1, c)))
        If Len(hdr) = 0 Then hdr = "Column " & c     ' an unnamed column still needs a home
        ' Two columns sharing a header WITHIN one source both map to the same destination
        ' column, so the second overwrites the first. Disambiguate rather than lose data:
        ' the duplicate becomes its own column, named so the user can see what happened.
        If seenHere.Exists(hdr) Then
            seenHere(hdr) = seenHere(hdr) + 1
            hdr = hdr & " (" & seenHere(hdr) & ")"
        Else
            seenHere.Add hdr, 1
        End If
        If Not headerMap.Exists(hdr) Then
            destCols = destCols + 1
            headerMap.Add hdr, destCols
            dest.Cells(1, destCols).Value = hdr
            dest.Cells(1, destCols).Font.Bold = True
        End If
        target = headerMap(hdr)
        For r = 2 To lastR
            colVals(r - 1, 1) = vals(r, c)
        Next r
        dest.Cells(destRow, target).Resize(lastR - 1, 1).Value = colVals
    Next c
    destRow = destRow + (lastR - 1)
    PkgImport_AppendSheet = True
End Function

'================================================================
' ACTION - Import one delimited file with explicit parsing
'================================================================
Public Sub PkgImport_DelimitedFile(ctx As Object)
    Dim picked As Variant, destName As String, dest As Worksheet, wbTxt As Workbook
    Dim rowsOut As Long, colsOut As Long, src As Worksheet
    Dim errNum As Long, errDesc As String
    Const TITLE As String = "Import a Delimited File"
    On Error GoTo Fail

    ' Application.GetOpenFilename returns Boolean False on Cancel, which VarType
    ' tells apart from any path string - the same reason PkgClean uses
    ' Application.InputBox rather than VBA's InputBox.
    picked = Application.GetOpenFilename( _
        FileFilter:="Delimited files (*.csv;*.txt;*.tsv),*.csv;*.txt;*.tsv", _
        title:="Pick the file to import")
    If VarType(picked) = vbBoolean Then Exit Sub                  ' user cancelled

    destName = PkgImport_SheetNameFor(CStr(picked))
    If Not PkgImport_DestUsable(destName, TITLE) Then Exit Sub

    Core_SpeedOn
    Set wbTxt = PkgImport_OpenDelimited(CStr(picked), _
                    Core_Cfg(ctx, "k_import_delim"), Core_Cfg(ctx, "k_import_qualifier"))
    Set dest = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    dest.name = destName
    Set src = wbTxt.Worksheets(1)
    ' Dimensions come from UsedRange, the same thing being copied. Reading them off
    ' column A / row 1 instead would misreport any file whose first column or header
    ' cell is empty - a correct import with a wrong number next to it.
    rowsOut = src.UsedRange.Rows.Count
    colsOut = src.UsedRange.Columns.Count
    src.UsedRange.Copy dest.Range("A1")
    Application.CutCopyMode = False
    wbTxt.Close SaveChanges:=False
    Set wbTxt = Nothing
    Core_SpeedOff

    Core_Done ctx, "Imported " & Core_FileName(CStr(picked)) & " into a new sheet '" & _
           destName & "'." & vbCrLf & rowsOut & " row(s), " & colsOut & " column(s).", TITLE
    Exit Sub
Fail:
    ' Same cleanup contract as the combines, and for the same reason: without it a
    ' failure after OpenText leaves the source file open in a second Excel window AND a
    ' half-written sheet in this workbook. Core_Fail restores application state only -
    ' it knows nothing about either object.
    '
    ' The error is captured FIRST and re-signalled afterwards because On Error Resume
    ' Next cleanup overwrites Err: without this, Core_Fail would report whatever the
    ' last cleanup statement did instead of the real cause. Re-raising under a local
    ' On Error GoTo keeps the report inside this action - the shipped contract is that
    ' every entry point reports through Core_Fail and lets no error escape to the
    ' dispatcher.
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    If Not wbTxt Is Nothing Then wbTxt.Close SaveChanges:=False
    If Not dest Is Nothing Then dest.Delete       ' alerts already off under Core_SpeedOn
    On Error GoTo Report
    Err.Clear
    Err.Raise errNum, "PkgImport_DelimitedFile", errDesc
    Exit Sub
Report:
    Core_Fail ctx, "import.delimited_file", Err
End Sub

'================================================================
' Package-private helpers. Prefixed PkgImport_ because the bundle is
' one flat namespace (design section 3.3), even though these are Private.
'================================================================

' A destination name this package may create. It REFUSES rather than
' overwrites: the library snippet this replaces opened with
' ThisWorkbook.Worksheets("Master").Delete under On Error Resume Next,
' which silently destroys whatever the user had on that sheet. Nothing in
' this package deletes a sheet it did not create in the same run.
Private Function PkgImport_DestUsable(ByVal destName As String, ByVal title As String) As Boolean
    Dim bad As String, i As Long
    If Len(destName) = 0 Then
        MsgBox "No name is set for the sheet this creates." & vbCrLf & vbCrLf & _
               "Fill in 'Name for the combined sheet this creates' on the Setup tab, " & _
               "Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, title
        Exit Function
    End If
    If Len(destName) > 31 Then
        MsgBox "'" & destName & "' is " & Len(destName) & " characters. Excel sheet names " & _
               "stop at 31." & vbCrLf & vbCrLf & "Nothing has been changed.", vbExclamation, title
        Exit Function
    End If
    bad = ":\/?*[]"
    For i = 1 To Len(bad)
        If InStr(destName, Mid$(bad, i, 1)) > 0 Then
            MsgBox "'" & destName & "' contains " & Mid$(bad, i, 1) & ", which Excel does not " & _
                   "allow in a sheet name." & vbCrLf & vbCrLf & _
                   "None of these are allowed:  : \ / ? * [ ]" & vbCrLf & vbCrLf & _
                   "Nothing has been changed.", vbExclamation, title
            Exit Function
        End If
    Next i
    If Core_SheetExists(destName) Then
        MsgBox "A sheet named '" & destName & "' already exists in this workbook, and this " & _
               "action will not overwrite it." & vbCrLf & vbCrLf & _
               "Rename or delete that sheet, then run " & title & " again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, title
        Exit Function
    End If
    PkgImport_DestUsable = True
End Function

' Sheets the engine owns. Combining one of these into a data sheet would
' be meaningless at best; _Config in particular is the activated config
' store. Generated button pages are "Macros 1", "Macros 2", ... so the
' test is a prefix, not a list.
Private Function PkgImport_IsEngineSheet(ByVal nm As String) As Boolean
    Dim tail As String
    Select Case LCase(Trim(nm))
        Case "setup", "setup menu", "_config", "_config_staged", "_packages", _
             "macros", "macros home"
            PkgImport_IsEngineSheet = True
            Exit Function
    End Select
    ' Generated button pages are "Macros 1", "Macros 2", ... The contract reserves THOSE,
    ' not the prefix: a plain Left$(nm, 7) = "macros " test also swallows a user's own
    ' "Macros Data" sheet and silently drops it from every combine. Require a numeric
    ' tail so only the generated names match.
    If LCase(Left$(Trim(nm), 7)) = "macros " Then
        tail = Trim(Mid$(Trim(nm), 8))
        If Len(tail) > 0 Then PkgImport_IsEngineSheet = PkgImport_IsDigits(tail)
    End If
End Function

Private Function PkgImport_IsDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    PkgImport_IsDigits = True
End Function

' Blank means the folder this workbook is saved in - contract section 9's
' "data location", and the answer in the overwhelming majority of cases.
' Returns a path ending in a separator, or "" when the workbook has never been
' saved and no folder was configured (the caller reports that case by name).
Private Function PkgImport_ResolveFolder(ByVal raw As String) As String
    Dim p As String
    p = Trim(raw)
    If Len(p) = 0 Then p = ThisWorkbook.Path
    If Len(p) = 0 Then Exit Function                 ' workbook never saved
    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    PkgImport_ResolveFolder = p
End Function

' Folder existence, via FileSystemObject rather than Dir(path, vbDirectory).
' Dir's behaviour with a trailing separator is not something to bet a refusal dialog
' on, and it also returns a non-empty result for a FILE, so a path pointing at a
' workbook would read as a valid folder. Late-bound: no reference to add, and no
' VBIDE trust dependency (CLAUDE.md rule 3).
Private Function PkgImport_FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object, p As String
    p = folderPath
    If Len(p) > 3 Then
        If Right$(p, 1) = Application.PathSeparator Then p = Left$(p, Len(p) - 1)
    End If
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then PkgImport_FolderExists = fso.FolderExists(p)
    On Error GoTo 0
End Function

' Every workbook in the folder except this one. Collects names FIRST and
' returns full paths, because calling anything that touches the file system
' inside a Dir() loop can corrupt the iterator on a network drive - the same
' reason Core_FindFile is written this way.
Private Function PkgImport_WorkbookPaths(ByVal folderPath As String) As Collection
    Dim out As Collection, names As Collection, f As String, one As Variant
    Set out = New Collection
    Set names = New Collection
    f = Dir(folderPath & "*.xls*")
    Do While Len(f) > 0
        names.Add f
        f = Dir
    Loop
    For Each one In names
        If LCase(CStr(one)) <> LCase(ThisWorkbook.name) Then out.Add folderPath & CStr(one)
    Next one
    Set PkgImport_WorkbookPaths = out
End Function

' A sheet name from a file name: drop the extension, replace the characters
' Excel forbids, and cut to 31. It does NOT make the name unique - a clash is
' reported by PkgImport_DestUsable and refused, exactly like a combine's.
Private Function PkgImport_SheetNameFor(ByVal fullPath As String) As String
    Dim nm As String, bad As String, i As Long
    nm = Core_FileName(fullPath)
    If InStrRev(nm, ".") > 1 Then nm = Left$(nm, InStrRev(nm, ".") - 1)
    bad = ":\/?*[]"
    For i = 1 To Len(bad)
        nm = Replace(nm, Mid$(bad, i, 1), "-")
    Next i
    nm = Trim(nm)
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    PkgImport_SheetNameFor = nm
End Function

' Open a delimited file with the separator and qualifier the deployment
' configured, instead of letting Excel guess from the extension. Workbooks.Open
' on a .csv applies the machine's list separator; a European-configured machine
' opening a comma file lands everything in column A, which is exactly the
' failure this action exists to prevent.
Private Function PkgImport_OpenDelimited(ByVal fullPath As String, ByVal delim As String, _
                                         ByVal qual As String) As Workbook
    Dim useTab As Boolean, useComma As Boolean, useSemi As Boolean
    Dim useSpace As Boolean, useOther As Boolean
    Dim otherCh As String, tq As Long, d As String

    d = LCase(Trim(delim))
    Select Case d
        Case "", "comma", ","
            useComma = True
        Case "tab"
            useTab = True
        Case "semicolon", ";"
            useSemi = True
        Case "space"
            useSpace = True
        Case "pipe", "|"
            useOther = True
            otherCh = "|"
        Case Else
            useOther = True
            otherCh = Left$(Trim(delim), 1)
    End Select

    ' Blank never reaches here: Core_Cfg resolves an empty cell to the manifest default,
    ' which is a double quote. "none" therefore has to be TYPED to be meant, and the
    ' field's help text says so. Anything unrecognised falls back to a double quote -
    ' the overwhelmingly common convention, and the safe guess for a stray value.
    Select Case LCase(Trim(qual))
        Case "none"
            tq = xlTextQualifierNone
        Case "'"
            tq = xlTextQualifierSingleQuote
        Case Else
            tq = xlTextQualifierDoubleQuote
    End Select

    Workbooks.OpenText Filename:=fullPath, DataType:=xlDelimited, TextQualifier:=tq, _
        Tab:=useTab, Semicolon:=useSemi, Comma:=useComma, Space:=useSpace, _
        Other:=useOther, OtherChar:=otherCh, Local:=True
    Set PkgImport_OpenDelimited = ActiveWorkbook
End Function
