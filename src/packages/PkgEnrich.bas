Attribute VB_Name = "PkgEnrich"
Option Explicit

'================================================================
' PkgEnrich - matching rows against a reference table and copying
' columns in, without a volatile lookup formula. Wave 4 of the catalog
' build-out (docs/macro-catalog-plug-and-play.md, section 4).
'
' Two actions, both the same match-and-copy operation with different
' input sources, so both delegate to ONE private engine:
'   enrich.join            everything picked at run time
'   enrich.mapped_columns  the reference sheet, its key, and which
'                          columns to add all come from Setup; only the
'                          target range is still picked at run time
'
' That is the same split PkgImport's two combine actions use for
' PkgImport_CombineEngine (PkgImport.bas:160-179) - one engine, two
' different sources for the same five inputs.
'
' The engine is dictionary-based (Scripting.Dictionary, late-bound),
' never a nested search, because this runs over a consolidated dataset
' that can be tens of thousands of rows: building the reference
' dictionary once is O(reference rows), then matching every target row
' is O(target rows) with an O(1) lookup each, and each matched row
' writes up to mappedColumns cells into the output array - so the full
' shape is O(reference + target*mappedColumns), not strictly
' O(reference + target). The point that matters stays true either way:
' a nested loop over both tables would be O(target rows * reference
' rows), minutes instead of seconds on a real month-end file.
'
' A reference key seen more than once resolves to its FIRST occurrence
' and is counted, never refused and never silently overwritten - see
' the plan's decision table for why. A target row matching nothing is
' left blank and counted, never an error - a partial match is the
' ordinary outcome at month-end, not a defect.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
Public Sub PkgEnrich_Manifest(reg As Object)
    Core_DeclarePackage reg, "enrich", "Match & Enrich", "1.0.0"

    Core_AddSection reg, "MAPPED COLUMNS"
    Core_AddField reg, "k_enrich_lookup_sheet", "Reference sheet to pull from", "sheet", "Ref_Master", _
        "simple", _
        "The sheet enrich.mapped_columns reads from. It must already exist here - unlike a destination sheet, this one is never created.", _
        "enrich.mapped_columns"
    Core_AddField reg, "k_enrich_lookup_key", "Matching column (in both sheets)", "text", "Store ID", _
        "simple", _
        "A column letter (e.g. B) or the exact header text in the reference sheet's first row. The range you pick when you run this must have a column with the same letter or header text - that is the column matched against.", _
        "enrich.mapped_columns"
    Core_AddTable reg, "k_enrich_mapped", "Columns to add", _
        "Reference column (letter or header)|New column header", "simple", _
        "One row per column to copy in from the reference sheet. Left cell: a column letter or header text in the reference sheet. Right cell: the header to give the new column here. Nothing happens until you fill this in.", _
        "enrich.mapped_columns"

    Core_AddAction reg, "enrich.join", "Join / Enrich Columns", "PkgEnrich_Join", _
        "Pick a range to enrich and a reference table (each including its header row); matches rows by a key column and copies in the reference columns you choose. Everything is picked at run time - nothing here is configured on Setup.", _
        "Match & Enrich"
    Core_AddAction reg, "enrich.mapped_columns", "Add Mapped Columns", "PkgEnrich_MappedColumns", _
        "Pick a range; adds the columns configured above from the reference sheet, matched on the key column. The mapping is set up once above and runs the same way every time.", _
        "Match & Enrich"
End Sub

'================================================================
' ACTION - Ad hoc join: every input picked at run time
'================================================================
Public Sub PkgEnrich_Join(ctx As Object)
    Dim targetRng As Range, lookupRng As Range, destPreview As Range
    Dim tKeys As String, lKeys As String, vKeys As String
    Dim targetKeyArr As Variant, lookupKeyArr As Variant, valueColsArr As Variant
    Dim newHeaders() As String, i As Long
    Dim matched As Long, unmatched As Long, dupKeys As Long, appendedCols As Long
    Const TITLE As String = "Join / Enrich Columns"
    On Error GoTo Fail

    Set targetRng = Core_AskRange(TITLE & " - the data to enrich (include its header row)")
    If targetRng Is Nothing Then Exit Sub                     ' user cancelled
    Set targetRng = Core_NarrowToUsed(targetRng)
    If targetRng Is Nothing Then
        Core_Refuse ctx, "Your selection needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    ElseIf targetRng.Rows.Count < 2 Then
        Core_Refuse ctx, "Your selection needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(targetRng, tKeys, _
            "Which ONE column holds the value to match on? Row 1 is treated as the header.", _
            TITLE) Then Exit Sub
    targetKeyArr = Core_ParseCols(tKeys)
    If UBound(targetKeyArr) > LBound(targetKeyArr) Then
        Core_Refuse ctx, "Only one column is supported as the match key; you named " & _
               (UBound(targetKeyArr) - LBound(targetKeyArr) + 1) & "." & vbCrLf & vbCrLf & _
               "Enter a single column number and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Set lookupRng = Core_AskRange(TITLE & " - the reference table (include its header row)")
    If lookupRng Is Nothing Then Exit Sub
    Set lookupRng = Core_NarrowToUsed(lookupRng)
    If lookupRng Is Nothing Then
        Core_Refuse ctx, "The reference table needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    ElseIf lookupRng.Rows.Count < 2 Then
        Core_Refuse ctx, "The reference table needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(lookupRng, lKeys, _
            "Which ONE column in the reference table is the matching key?", TITLE) Then Exit Sub
    lookupKeyArr = Core_ParseCols(lKeys)
    If UBound(lookupKeyArr) > LBound(lookupKeyArr) Then
        Core_Refuse ctx, "Only one column is supported as the match key; you named " & _
               (UBound(lookupKeyArr) - LBound(lookupKeyArr) + 1) & "." & vbCrLf & vbCrLf & _
               "Enter a single column number and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_AskCols(lookupRng, vKeys, _
            "Which column(s) in the reference table do you want copied in? " & _
            "Numbered from the left of the reference table.", TITLE) Then Exit Sub
    valueColsArr = Core_ParseCols(vKeys)

    ' New columns are named after the reference table's OWN header text - this
    ' action asks nothing about naming, which is exactly the difference from
    ' enrich.mapped_columns, where the header text is configured once.
    ReDim newHeaders(LBound(valueColsArr) To UBound(valueColsArr))
    For i = LBound(valueColsArr) To UBound(valueColsArr)
        newHeaders(i) = CStr(lookupRng.Cells(1, CLng(valueColsArr(i))).Value)
    Next i

    ' Preview the EXACT block the engine is about to write and refuse if any
    ' cell in it already holds something - "never overwrites existing
    ' columns" is only true if this is checked before the write, not assumed
    ' from wherever the write happens to land.
    Set destPreview = targetRng.Parent.Range( _
        targetRng.Parent.Cells(targetRng.Row, targetRng.Column + targetRng.Columns.Count), _
        targetRng.Parent.Cells(targetRng.Row + targetRng.Rows.Count - 1, _
            targetRng.Column + targetRng.Columns.Count + UBound(valueColsArr) - LBound(valueColsArr)))
    If Application.WorksheetFunction.CountA(destPreview) > 0 Then
        Core_Refuse ctx, "The column(s) immediately to the right of " & Core_RangeRef(targetRng) & _
               " already contain data - adding new columns there would overwrite it." & _
               vbCrLf & vbCrLf & "Clear that area, or pick a target range with empty columns to " & _
               "its right, then run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Explicit confirmation before the write - this is the one place in the
    ' action where cancelling still costs nothing.
    If MsgBox("This adds " & (UBound(valueColsArr) - LBound(valueColsArr) + 1) & _
              " new column(s) to " & Core_RangeRef(targetRng) & ", immediately to its right." & _
              vbCrLf & vbCrLf & "Proceed?", vbOKCancel + vbQuestion, TITLE) <> vbOK Then Exit Sub

    Core_SpeedOn
    PkgEnrich_JoinEngine targetRng, CLng(targetKeyArr(LBound(targetKeyArr))), _
        lookupRng, CLng(lookupKeyArr(LBound(lookupKeyArr))), _
        valueColsArr, newHeaders, matched, unmatched, dupKeys, appendedCols
    Core_SpeedOff

    Core_Done ctx, "Added " & appendedCols & " column(s) to " & Core_RangeRef(targetRng) & "." & _
           vbCrLf & matched & " row(s) matched, " & unmatched & " row(s) did not." & _
           IIf(dupKeys > 0, vbCrLf & vbCrLf & dupKeys & " key(s) in the reference table were " & _
               "not unique; the FIRST row for each was used.", ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "enrich.join", Err
End Sub

'================================================================
' ACTION - Add the columns configured on Setup, matched by key
'================================================================
Public Sub PkgEnrich_MappedColumns(ctx As Object)
    Dim lookupSheetName As String, lookupKeySpec As String
    Dim lookupWs As Worksheet, lookupRng As Range, targetRng As Range, destPreview As Range
    Dim refCols() As String, newHdrs() As String, spare() As String
    Dim lookupKeyCol As Long, targetKeyCol As Long, colIdx As Long
    Dim valueCols() As Long, headers() As String, i As Long, n As Long
    Dim missing As String
    Dim matched As Long, unmatched As Long, dupKeys As Long, appendedCols As Long
    Const TITLE As String = "Add Mapped Columns"
    On Error GoTo Fail

    lookupSheetName = Trim(Core_Cfg(ctx, "k_enrich_lookup_sheet"))
    If Not Core_SheetExists(lookupSheetName) Then
        Core_Refuse ctx, "The reference sheet '" & lookupSheetName & "' does not exist in this " & _
               "workbook." & vbCrLf & vbCrLf & "Fix 'Reference sheet to pull from' on the Setup " & _
               "tab, Finalize, and run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If
    Set lookupWs = ThisWorkbook.Worksheets(lookupSheetName)
    Set lookupRng = Core_NarrowToUsed(lookupWs.UsedRange)
    If lookupRng Is Nothing Then
        Core_Refuse ctx, "'" & lookupSheetName & "' has no data rows to match against." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    ElseIf lookupRng.Rows.Count < 2 Then
        Core_Refuse ctx, "'" & lookupSheetName & "' has no data rows to match against." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    lookupKeySpec = Trim(Core_Cfg(ctx, "k_enrich_lookup_key"))
    lookupKeyCol = Core_ResolveColumn(lookupRng, lookupKeySpec)
    If lookupKeyCol = 0 Then
        Core_Refuse ctx, "'" & lookupKeySpec & "' is not a column in '" & lookupSheetName & "'." & _
               vbCrLf & vbCrLf & "Fix 'Matching column' on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    If Not Core_CfgTable(ctx, "k_enrich_mapped", refCols, newHdrs, spare) Then
        Core_Refuse ctx, "No columns are set up to add yet." & vbCrLf & vbCrLf & _
               "Fill in the 'Columns to add' table on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ReDim valueCols(0 To UBound(refCols) - LBound(refCols))
    ReDim headers(0 To UBound(refCols) - LBound(refCols))
    n = 0
    For i = LBound(refCols) To UBound(refCols)
        colIdx = Core_ResolveColumn(lookupRng, refCols(i))
        If colIdx = 0 Then
            missing = missing & vbCrLf & "   " & refCols(i)
        Else
            valueCols(n) = colIdx
            headers(n) = newHdrs(i)
            n = n + 1
        End If
    Next i
    If n = 0 Then
        Core_Refuse ctx, "None of the configured reference columns exist in '" & lookupSheetName & _
               "'." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If
    ReDim Preserve valueCols(0 To n - 1)
    ReDim Preserve headers(0 To n - 1)

    Set targetRng = Core_AskRange(TITLE & " - the data to enrich (include its header row)")
    If targetRng Is Nothing Then Exit Sub                     ' user cancelled
    Set targetRng = Core_NarrowToUsed(targetRng)
    If targetRng Is Nothing Then
        Core_Refuse ctx, "Your selection needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    ElseIf targetRng.Rows.Count < 2 Then
        Core_Refuse ctx, "Your selection needs a header row and at least one data row." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    targetKeyCol = Core_ResolveColumn(targetRng, lookupKeySpec)
    If targetKeyCol = 0 Then
        Core_Refuse ctx, "'" & lookupKeySpec & "' is not a column in your selection." & vbCrLf & vbCrLf & _
               "The matching column is the same setting for both sides - pick a range here whose " & _
               "header includes it." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Preview the EXACT block the engine is about to write and refuse if any
    ' cell in it already holds something - same append-only guarantee as
    ' enrich.join, checked the same way.
    Set destPreview = targetRng.Parent.Range( _
        targetRng.Parent.Cells(targetRng.Row, targetRng.Column + targetRng.Columns.Count), _
        targetRng.Parent.Cells(targetRng.Row + targetRng.Rows.Count - 1, _
            targetRng.Column + targetRng.Columns.Count + UBound(valueCols) - LBound(valueCols)))
    If Application.WorksheetFunction.CountA(destPreview) > 0 Then
        Core_Refuse ctx, "The column(s) immediately to the right of " & Core_RangeRef(targetRng) & _
               " already contain data - adding new columns there would overwrite it." & _
               vbCrLf & vbCrLf & "Clear that area, or pick a target range with empty columns to " & _
               "its right, then run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Explicit confirmation before the write.
    If MsgBox("This adds " & (UBound(valueCols) - LBound(valueCols) + 1) & " new column(s) to " & _
              Core_RangeRef(targetRng) & ", immediately to its right, from '" & lookupSheetName & _
              "'." & vbCrLf & vbCrLf & "Proceed?", vbOKCancel + vbQuestion, TITLE) <> vbOK Then Exit Sub

    Core_SpeedOn
    PkgEnrich_JoinEngine targetRng, targetKeyCol, lookupRng, lookupKeyCol, valueCols, headers, _
        matched, unmatched, dupKeys, appendedCols
    Core_SpeedOff

    Core_Done ctx, "Added " & appendedCols & " column(s) to " & Core_RangeRef(targetRng) & _
           " from '" & lookupSheetName & "'." & vbCrLf & matched & " row(s) matched, " & _
           unmatched & " row(s) did not." & _
           IIf(dupKeys > 0, vbCrLf & vbCrLf & dupKeys & " key(s) in '" & lookupSheetName & _
               "' were not unique; the FIRST row for each was used.", "") & _
           IIf(Len(missing) > 0, vbCrLf & vbCrLf & "These configured reference columns were not " & _
               "found and were skipped:" & missing, ""), TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "enrich.mapped_columns", Err
End Sub

'================================================================
' Shared join engine. Both actions call this - see the file header for
' why one engine, and what "first wins" and "unmatched stays blank"
' mean for a finance dataset.
'
' Reads both ranges into arrays and writes the result as ONE block,
' never cell by cell - the reasoning PkgImport_AppendSheet's own header
' comment gives applies here just as hard (do-not-revert list, plan-of-
' record section 4.1): a per-cell COM write over tens of thousands of rows is
' minutes, not seconds, even with screen updating off.
'
' Both ranges are assumed to have a header row in row 1 - the caller's
' job, not this engine's; see Decision 5.
'================================================================
Private Sub PkgEnrich_JoinEngine(targetRng As Range, ByVal targetKeyCol As Long, _
                                 lookupRng As Range, ByVal lookupKeyCol As Long, _
                                 valueCols As Variant, newHeaders As Variant, _
                                 ByRef matched As Long, ByRef unmatched As Long, _
                                 ByRef dupKeys As Long, ByRef appendedCols As Long)
    Dim tArr As Variant, lArr As Variant, outArr As Variant
    Dim dict As Object, dupSeen As Object, key As String
    Dim nRows As Long, mRows As Long, nCols As Long, c As Long, r As Long
    Dim hitRow As Variant, srcCol As Long
    Dim destWs As Worksheet, firstNewCol As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1                            ' vbTextCompare - same rule as PkgImport's headerMap
    Set dupSeen = CreateObject("Scripting.Dictionary")
    dupSeen.CompareMode = 1                         ' tracks which keys' duplicate has already been counted

    tArr = targetRng.Value
    lArr = lookupRng.Value
    nRows = UBound(tArr, 1)                         ' row 1 is the header, both sides always have one
    mRows = UBound(lArr, 1)
    nCols = UBound(valueCols) - LBound(valueCols) + 1

    ' Build the reference dictionary once. Row 1 is the header and is skipped.
    '
    ' KEY NORMALIZATION: Trim$ only - leading/trailing whitespace is stripped
    ' because it is invisible and never distinguishes two real identifiers, but
    ' nothing else about the text is touched. Core_CleanText is deliberately
    ' NOT used here even though it is available: it collapses repeated internal
    ' spaces and strips control characters, which would let two DIFFERENT
    ' finance identifiers (e.g. "AB  12" and "AB 12", or a store code with an
    ' embedded tab from a bad export) resolve to the SAME dictionary key and
    ' silently pair with the wrong reference row - confidently wrong output
    ' over a consolidated dataset, the one failure mode a join cannot have.
    ' Case-insensitivity is handled separately, by the dictionary's own
    ' CompareMode above, not by rewriting the key text. Nothing in the
    ' contract authorizes a fuzzier match than this.
    '
    ' DUPLICATE KEYS: the row already in the dictionary wins - the first one
    ' seen, top to bottom. dupKeys counts DISTINCT keys that turned out to be
    ' duplicated, not total repeated occurrences - a key appearing 4 times is
    ' ONE duplicate key, not three - so dupSeen marks a key the first time its
    ' second occurrence is found, and every occurrence after that no longer
    ' increments dupKeys. Nothing here refuses on a duplicate: a reference
    ' table with a repeated key is common, not exceptional, and the count is
    ' how the user finds out rather than getting a wrong number with no
    ' signal at all.
    For r = 2 To mRows
        key = Trim$(CStr(lArr(r, lookupKeyCol)))
        If Len(key) > 0 Then
            If dict.Exists(key) Then
                If Not dupSeen.Exists(key) Then
                    dupKeys = dupKeys + 1
                    dupSeen.Add key, True
                End If
            Else
                dict.Add key, r
            End If
        End If
    Next r

    ' Build the output block in memory: row 1 = the new headers, rows 2..nRows
    ' = the matched reference values, or blank for a row with no match.
    ReDim outArr(1 To nRows, 1 To nCols)
    For c = 1 To nCols
        outArr(1, c) = CStr(newHeaders(LBound(newHeaders) + c - 1))
    Next c
    For r = 2 To nRows
        key = Trim$(CStr(tArr(r, targetKeyCol)))    ' same normalization policy as the reference side
        If Len(key) > 0 Then
            If dict.Exists(key) Then
                hitRow = dict(key)
                matched = matched + 1
                For c = 1 To nCols
                    srcCol = CLng(valueCols(LBound(valueCols) + c - 1))
                    outArr(r, c) = lArr(hitRow, srcCol)
                Next c
            Else
                unmatched = unmatched + 1           ' outArr row stays blank for this row's new columns
            End If
        Else
            unmatched = unmatched + 1               ' an empty target key can never match anything
        End If
    Next r

    ' One write, immediately to the right of the target range. The CALLER
    ' (PkgEnrich_Join / PkgEnrich_MappedColumns) is responsible for previewing
    ' this exact destination block and refusing before this engine is ever
    ' invoked if any cell in it is already occupied - this engine has no ctx
    ' to refuse with, so it only writes, it never checks.
    Set destWs = targetRng.Parent
    firstNewCol = targetRng.Column + targetRng.Columns.Count
    destWs.Range(destWs.Cells(targetRng.Row, firstNewCol), _
                 destWs.Cells(targetRng.Row + nRows - 1, firstNewCol + nCols - 1)).Value = outArr
    appendedCols = nCols
End Sub
