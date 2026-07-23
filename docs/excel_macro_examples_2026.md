# Excel VBA Macro Examples for Common Spreadsheet Automation Patterns

This document collects practical VBA macro examples that map to the most common Excel automation patterns still widely used in 2026, including data cleanup, report formatting, navigation helpers, pivot/chart automation, file administration, and recurring finance-style reporting routines.[cite:16][cite:17][cite:21][cite:22][cite:23][cite:30]

## How to use these examples

Each example is designed as a starting point rather than a drop-in universal solution. Sheet names, ranges, folder paths, file naming rules, and workbook structure should be parameterized for the actual environment before deployment, especially in recurring reporting workflows.[cite:22][cite:25]

## Data cleanup

### Fill blanks with the value above

This pattern is commonly used after exports where repeated dimension values only appear once and subsequent rows are blank. Microsoft Q&A and practical Excel tutorials both reflect this as a frequent cleanup scenario.[cite:17][cite:26]

```vb
Sub FillBlanksWithAbove()
    Dim rng As Range
    Dim blanks As Range

    Set rng = Selection

    On Error Resume Next
    Set blanks = rng.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    If blanks Is Nothing Then Exit Sub

    blanks.FormulaR1C1 = "=R[-1]C"
    rng.Value = rng.Value
End Sub
```

**What it does:** Finds blank cells in the selected range, fills each blank with the value from the cell above, then converts formulas back to hard values.

### Split text into columns

This is useful when a CSV-style export lands in one column and must be separated into fields before analysis or loading into a pivot workflow.

```vb
Sub SplitTextByComma()
    With Selection
        .TextToColumns Destination:=.Cells(1, 1), _
            DataType:=xlDelimited, _
            Comma:=True, _
            TextQualifier:=xlTextQualifierDoubleQuote
    End With
End Sub
```

**What it does:** Splits the selected column into multiple columns using commas as delimiters.

### Remove duplicates from a range

Removing duplicates is a standard cleanup step before creating unique lists, summary sheets, or pivot sources. ExcelCampus specifically highlights duplicate-removal workflows as a common recurring task.[cite:20][cite:23]

```vb
Sub RemoveDuplicatesSelectedRange()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("A1:D" & lastRow).RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlYes
End Sub
```

**What it does:** Removes duplicate rows across columns A:D while keeping the header row.

### Standardize formats

Formatting normalization is common when imported data mixes text, dates, and inconsistent number formatting. VBA formatting examples are frequently used to force consistent display after data transfer.[cite:28]

```vb
Sub StandardizeFormats()
    With ActiveSheet
        .Columns("A:A").NumberFormat = "mm/dd/yyyy"
        .Columns("B:B").NumberFormat = "0"
        .Columns("C:C").NumberFormat = "$#,##0.00"
        .Columns("D:D").NumberFormat = "0.00%"
    End With
End Sub
```

**What it does:** Applies consistent date, integer, currency, and percentage formats to key columns.

### Convert formulas to values

This is commonly done before archiving reports, sending static packs, or freezing outputs that should not recalculate.

```vb
Sub ConvertFormulasToValues()
    With Selection
        .Value = .Value
    End With
End Sub
```

**What it does:** Replaces formulas in the selected range with their current results.

## Report formatting

### Add a report title and style it

Everyday macro examples published by ExcelCampus include automated title formatting as a practical time-saver in recurring reports.[cite:6]

```vb
Sub AddStyledReportTitle()
    With Range("A1")
        .Value = "Month-End Operating Report"
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = RGB(217, 225, 242)
    End With
    Range("A1:F1").Merge
    Range("A1").HorizontalAlignment = xlCenter
End Sub
```

**What it does:** Writes a title in A1, merges the heading area, and applies a basic report-title style.

### Format headers and resize columns

This is a common pattern in recorded macros and report-finishing routines because many exports arrive with raw, unreadable layouts.[cite:1][cite:6]

```vb
Sub FormatHeadersAndResize()
    Dim lastCol As Long

    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column

    With Range(Cells(1, 1), Cells(1, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(48, 84, 150)
        .Font.Color = RGB(255, 255, 255)
    End With

    Cells.EntireColumn.AutoFit
End Sub
```

**What it does:** Styles the first row as headers and autofits all columns.

### Apply print layout settings

Print-prep automation is still useful for board packs, operations packets, and period-end binders.

```vb
Sub PrepPrintLayout()
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
    End With
End Sub
```

**What it does:** Sets a landscape, fit-to-width print layout with tighter margins.

## Navigation helpers

### Create a table of contents sheet

ExcelCampus highlights a table-of-contents macro as one of the highest-utility workbook navigation helpers for larger files.[cite:6]

```vb
Sub CreateTableOfContents()
    Dim ws As Worksheet
    Dim toc As Worksheet
    Dim i As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("TOC").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set toc = Worksheets.Add(Before:=Worksheets(1))
    toc.Name = "TOC"
    toc.Range("A1").Value = "Sheet Name"

    i = 2
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "TOC" Then
            toc.Hyperlinks.Add Anchor:=toc.Cells(i, 1), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            i = i + 1
        End If
    Next ws

    toc.Columns("A:A").AutoFit
End Sub
```

**What it does:** Rebuilds a TOC worksheet with hyperlinks to every sheet in the workbook.

### Add a back-to-TOC link on every sheet

This complements the table-of-contents pattern and is especially useful in workbooks with many report tabs.

```vb
Sub AddBackToTOCLinks()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "TOC" Then
            ws.Hyperlinks.Add Anchor:=ws.Range("A1"), _
                Address:="", _
                SubAddress:="'TOC'!A1", _
                TextToDisplay:="Back to TOC"
        End If
    Next ws
End Sub
```

**What it does:** Adds a hyperlink in cell A1 of each non-TOC sheet that returns to the TOC tab.

### Jump to the next worksheet

Navigation shortcuts are simple but useful in operational workbooks with many repeated tabs.

```vb
Sub GoToNextSheet()
    Dim idx As Long
    idx = ActiveSheet.Index

    If idx < Worksheets.Count Then
        Worksheets(idx + 1).Activate
    End If
End Sub
```

**What it does:** Activates the next worksheet to the right.

## Chart and pivot automation

### Refresh all pivot tables in the workbook

Pivot refresh routines are one of the most common finance and reporting macros because source data changes frequently while summary views must stay current.[cite:27][cite:30]

```vb
Sub RefreshAllPivotTables()
    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
End Sub
```

**What it does:** Loops every sheet and refreshes each pivot table.

### Copy pivot table values and formatting to a static sheet

Pivot-table.com documents a macro pattern for copying pivot outputs as values and formatting to a new worksheet, which is useful when distributing static reports.[cite:21]

```vb
Sub CopyPivotToStaticSheet()
    Dim pt As PivotTable
    Dim wsNew As Worksheet

    Set pt = ActiveCell.PivotTable
    Set wsNew = Worksheets.Add

    pt.TableRange2.Copy
    wsNew.Range("A1").PasteSpecial xlPasteValues
    wsNew.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub
```

**What it does:** Copies the active pivot table to a new worksheet as static values plus formatting.

### Apply a consistent chart style to all charts on the active sheet

This pattern is useful where recurring reporting expects a fixed visual standard.

```vb
Sub FormatChartsOnActiveSheet()
    Dim co As ChartObject

    For Each co In ActiveSheet.ChartObjects
        With co.Chart
            .ChartStyle = 201
            .HasTitle = True
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
            .Legend.Position = xlLegendPositionBottom
        End With
    Next co
End Sub
```

**What it does:** Loops through each chart object on the active sheet and applies a consistent style.

## File safety and admin tasks

### Back up the workbook with a timestamp

Saving timestamped copies is a standard protective macro pattern for version control and auditability. YouTube examples show this approach for month-end-oriented naming conventions.[cite:22]

```vb
Sub BackupWorkbookTimestamped()
    Dim backupPath As String
    Dim fileName As String

    backupPath = ThisWorkbook.Path & "\Backup\"
    If Dir(backupPath, vbDirectory) = "" Then MkDir backupPath

    fileName = Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) _
        & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"

    ThisWorkbook.SaveCopyAs backupPath & fileName
End Sub
```

**What it does:** Saves a timestamped backup copy of the current workbook into a Backup subfolder.

### Save a month-end versioned file

Prior-month-end file naming is a real-world reporting pattern documented in Excel VBA examples focused on finance-style packs.[cite:22]

```vb
Sub SaveAsPriorMonthEnd()
    Dim dt As Date
    Dim newName As String
    Dim targetPath As String

    dt = DateSerial(Year(Date), Month(Date), 0)
    targetPath = ThisWorkbook.Path & "\"
    newName = "MonthEnd_Report_" & Format(dt, "yyyy-mm-dd") & ".xlsx"

    ThisWorkbook.SaveAs Filename:=targetPath & newName, FileFormat:=xlOpenXMLWorkbook
End Sub
```

**What it does:** Saves the workbook as a non-macro file named with the prior month-end date.

### Move a generated file to an archive folder

Folder and file movement automation is a common pattern mentioned by practitioners for recurring workflow administration.[cite:3]

```vb
Sub MoveFileToArchive()
    Dim sourceFile As String
    Dim archiveFolder As String
    Dim fileName As String

    sourceFile = ThisWorkbook.FullName
    archiveFolder = ThisWorkbook.Path & "\Archive\"
    fileName = Dir(sourceFile)

    If Dir(archiveFolder, vbDirectory) = "" Then MkDir archiveFolder

    Name sourceFile As archiveFolder & fileName
End Sub
```

**What it does:** Moves the current workbook file into an Archive folder in the same directory.

## Finance and recurring reporting routines

### Month-end report formatting routine

This is a practical composite macro that combines several patterns used in close and reporting cycles: title formatting, header styling, autofit, and print layout.[cite:1][cite:6][cite:22]

```vb
Sub FormatMonthEndReport()
    Call AddStyledReportTitle
    Call FormatHeadersAndResize
    Call PrepPrintLayout
End Sub
```

**What it does:** Runs a small formatting stack that prepares a recurring month-end report for review or distribution.

### Normalize source data before building pivots

Data normalization before pivoting often includes filling blanks, standardizing formats, and removing duplicates so summaries do not misstate results.[cite:17][cite:23][cite:28][cite:30]

```vb
Sub NormalizeDataBeforePivot()
    Call FillBlanksWithAbove
    Call StandardizeFormats
    Call RemoveDuplicatesSelectedRange
End Sub
```

**What it does:** Chains common cleanup steps before pivot tables are refreshed or rebuilt.

### Refresh and export routine

A common reporting pattern is to refresh pivots, convert formulas to values, and produce a static output for circulation.

```vb
Sub RefreshAndExportRoutine()
    Call RefreshAllPivotTables
    Call ConvertFormulasToValues
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=ThisWorkbook.Path & "\Report_" & Format(Date, "yyyymmdd") & ".pdf"
End Sub
```

**What it does:** Refreshes summaries, freezes selected formula outputs, and exports the workbook to PDF.

### Recurring reporting pack archival movement

This pattern is useful when a reporting pack is generated in one folder and then moved into a dated archive after completion.

```vb
Sub ArchiveReportingPack()
    Dim reportFolder As String
    Dim archiveFolder As String
    Dim reportName As String

    reportFolder = ThisWorkbook.Path & "\"
    archiveFolder = ThisWorkbook.Path & "\Archive\" & Format(Date, "yyyy_mm") & "\"
    reportName = Dir(ThisWorkbook.FullName)

    If Dir(ThisWorkbook.Path & "\Archive\", vbDirectory) = "" Then MkDir ThisWorkbook.Path & "\Archive\"
    If Dir(archiveFolder, vbDirectory) = "" Then MkDir archiveFolder

    Name reportFolder & reportName As archiveFolder & reportName
End Sub
```

**What it does:** Moves the current workbook into a monthly archive subfolder.

## Implementation notes

These examples are intentionally simple and readable. In production, the main upgrades are usually input validation, dynamic range detection, error handling, logging, configuration constants, and replacing `Selection`-based logic with explicit worksheet/range references to reduce operator risk.[cite:22][cite:25]

Security also matters more in 2026 because Excel macro controls are tighter for files originating from the internet, which makes signed macros, trusted locations, and controlled internal distribution more important for real deployment.[cite:9][cite:10]
