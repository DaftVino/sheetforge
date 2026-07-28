Attribute VB_Name = "GateAudit"
Option Explicit

'================================================================
' GateAudit - automated portion of the v0.7.0 / v0.8.0 / v0.9.0 manual gates.
'
' NOT product code. Nothing in build/build.ps1 reads tools/, so this never ships.
' Import it (File > Import File), run GateAudit_Run from Alt+F8, read the report.
'
' It replaces the eyeball-and-Immediate-window steps in three gates with one click:
'   hardening 1.1   no staging sheet leaked, _Config very hidden
'   reorder   1     Action catalog above Chains, both above the package bands
'   reorder   2     exactly one blank row at each block boundary
'   reskin    1     Setup sheet colours
'   reskin    2     Macros pages match, and the tab strips stay distinct
'   reskin    3     semantic colours survived, brand red confined to the tab
'
' It CANNOT cover what needs a human or a different build: the crash-recovery
' flow, any refusal dialog, the dark-theme look, the core.* hostile build, or the
' v0.6.0 old-order upgrade. Those stay manual and are listed at the end of the
' report so nothing silently drops off.
'
' EVERY probe is wrapped so this can never raise. A probe that cannot run reports
' SKIP with the reason - it never counts as a pass.
'================================================================

Private Const C_RED     As Long = 2105019     ' RGB(187, 30, 32)
Private Const C_NAVY    As Long = 3022353     ' RGB( 17, 30, 46)
Private Const C_STEEL   As Long = 4865073     ' RGB( 49, 60, 74)
Private Const C_TINT    As Long = 15657186    ' RGB(226,232,238)
Private Const C_YELLOW  As Long = 13431551    ' RGB(255,242,204)
Private Const C_CAUTION As Long = 150         ' RGB(150,  0,  0)
Private Const C_WHITE   As Long = 16777215    ' RGB(255,255,255)

Private mRep  As String
Private mPass As Long
Private mFail As Long
Private mSkip As Long

Public Sub GateAudit_Run()
    mRep = "": mPass = 0: mFail = 0: mSkip = 0

    Say "SHEETFORGE GATE AUDIT"
    Say "Workbook: " & ThisWorkbook.Name
    Say String(66, "=")

    AuditSheets
    AuditOrder
    AuditSpacing
    AuditSetupColours
    AuditPanelColours
    AuditBrandRedConfined

    Say String(66, "=")
    Say "PASS " & mPass & "    FAIL " & mFail & "    SKIP " & mSkip
    Say ""
    Say "STILL MANUAL - this macro cannot check these:"
    Say "  hardening 1.2/1.3  crash recovery, and Finalize after it"
    Say "  hardening 2.1/2.2  marker refusal dialog, and the residual hole"
    Say "  hardening 3.1      core.* refusal (needs a hostile build)"
    Say "  hardening 4.1/4.2  fresh-install default, blank cell blocks Finalize"
    Say "  reorder   3        v0.6.0 old-order config upgrade"
    Say "  reskin    4        legibility on a dark Excel theme"

    WriteReport
End Sub

'---------------- checks ----------------

Private Sub AuditSheets()
    Dim w As Worksheet
    Head "SHEETS"

    Set w = Sh("Setup")
    If w Is Nothing Then
        Skp "Setup sheet exists", "not found - run Core_BuildSetup first"
    Else
        Pas "Setup sheet exists"
    End If

    Set w = Sh("_Config")
    If w Is Nothing Then
        Skp "_Config exists", "not found - run Core_FinalizeSetup first"
    Else
        Pas "_Config exists"
        Num "_Config visibility (2 = very hidden)", CLng(w.Visible), 2
    End If

    Set w = Sh("_Config_staged")
    If w Is Nothing Then
        Pas "no _Config_staged left behind        [hardening 1.1]"
    Else
        Fal "no _Config_staged left behind", "it is present - the swap did not complete"
    End If
End Sub

Private Sub AuditOrder()
    Dim w As Worksheet, rCat As Long, rChn As Long, rEnd As Long, rPkg As Long
    Head "BLOCK ORDER ON SETUP        [reorder 1]"

    Set w = Sh("Setup")
    If w Is Nothing Then Skp "block order", "no Setup sheet": Exit Sub

    rCat = KeyRow(w, "k_core_layout")
    rChn = KeyRow(w, "k_core_chains")
    rEnd = KeyRow(w, "_core_end")
    rPkg = FirstPackageKeyRow(w)

    Say "  rows: catalog=" & rCat & "  chains=" & rChn & _
        "  first package key=" & rPkg & "  _core_end=" & rEnd

    If rCat = 0 Or rChn = 0 Then
        Skp "catalog above chains", "a key row was not found"
    ElseIf rCat < rChn Then
        Pas "Action catalog is above Chains"
    Else
        Fal "Action catalog is above Chains", "catalog row " & rCat & " is below chains row " & rChn
    End If

    If rChn = 0 Or rPkg = 0 Then
        Skp "compose above the package bands", "no package field key found"
    ElseIf rChn < rPkg Then
        Pas "Chains is above the package settings bands"
    Else
        Fal "Chains is above the package settings bands", "chains row " & rChn & " is below package row " & rPkg
    End If

    If rEnd = 0 Then
        Fal "_core_end sentinel present", "not found"
    ElseIf rEnd > rCat And rEnd > rChn And rEnd > rPkg Then
        Pas "_core_end is the last key row"
    Else
        Fal "_core_end is the last key row", "it sits above another key row"
    End If
End Sub

Private Sub AuditSpacing()
    Dim w As Worksheet, rChn As Long, nk As Long, i As Long, blanks As Long
    Head "BLOCK SPACING               [reorder 2, spacing half]"

    Set w = Sh("Setup")
    If w Is Nothing Then Skp "spacing", "no Setup sheet": Exit Sub

    rChn = KeyRow(w, "k_core_chains")
    If rChn = 0 Then Skp "one blank row after the Chains block", "chains key not found": Exit Sub

    nk = NextKeyRowAt(w, rChn)
    If nk = 0 Then Skp "one blank row after the Chains block", "no following key row": Exit Sub

    ' Count trailing wholly-empty rows between the last populated chains row and
    ' whatever block starts next.
    For i = nk - 1 To rChn + 1 Step -1
        If Application.WorksheetFunction.CountA(w.Rows(i)) = 0 Then
            blanks = blanks + 1
        Else
            Exit For
        End If
    Next i

    Say "  blank rows before the next block: " & blanks
    If blanks = 1 Then
        Pas "exactly one blank row after the Chains block"
    Else
        Fal "exactly one blank row after the Chains block", blanks & " found, expected 1"
    End If
End Sub

Private Sub AuditSetupColours()
    Dim w As Worksheet, r As Long, c As Range
    Head "SETUP COLOURS               [reskin 1 and 3]"

    Set w = Sh("Setup")
    If w Is Nothing Then Skp "Setup colours", "no Setup sheet": Exit Sub

    Num "Setup sheet tab = ForgeRed", TabColour(w), C_RED
    Num "B1 title font = ForgeNavy", CLng(w.Range("B1").Font.Color), C_NAVY

    r = RowContaining(w, 2, "COMPOSE YOUR BUTTONS")
    If r = 0 Then
        Skp "band header colours", "COMPOSE YOUR BUTTONS row not found in column B"
    Else
        Num "band header font = ForgeNavy", CLng(w.Cells(r, 2).Font.Color), C_NAVY
        Num "band header fill = ForgeTint", CLng(w.Cells(r, 2).Interior.Color), C_TINT
    End If

    r = RowContaining(w, 2, "Do not edit column A")
    If r = 0 Then
        Skp "caution line still its own red", "caution row not found"
    Else
        Num "caution font = 150, NOT the brand red", CLng(w.Cells(r, 2).Font.Color), C_CAUTION
    End If

    Set c = FirstCellWithFill(w, C_YELLOW)
    If c Is Nothing Then
        Fal "yellow input cells survived", "no cell with the yellow fill found"
    Else
        Pas "yellow input cells survived (first at " & c.Address(False, False) & ")"
    End If
End Sub

Private Sub AuditPanelColours()
    Dim w As Worksheet
    Head "MACROS PAGES                [reskin 2]"

    Set w = Sh("Setup Menu")
    If w Is Nothing Then
        Skp "Setup Menu colours", "sheet not found - run Core_FinalizeSetup"
    Else
        Num "Setup Menu tab = ForgeSteel", TabColour(w), C_STEEL
        Num "Setup Menu B2 font = ForgeNavy", CLng(w.Range("B2").Font.Color), C_NAVY
    End If

    Set w = Sh("Macros 1")
    If w Is Nothing Then
        Skp "Macros 1 colours", "sheet not found - Finalize with at least one Button action"
    Else
        Num "Macros 1 tab = ForgeSteel", TabColour(w), C_STEEL
    End If

    ' The two surfaces must not converge on one colour.
    Dim a As Worksheet, b As Worksheet
    Set a = Sh("Setup"): Set b = Sh("Macros 1")
    If a Is Nothing Or b Is Nothing Then
        Skp "Setup and Macros tabs stay distinct", "a sheet is missing"
    ElseIf TabColour(a) <> TabColour(b) Then
        Pas "Setup and Macros tab colours are distinct"
    Else
        Fal "Setup and Macros tab colours are distinct", "both read " & TabColour(a)
    End If
End Sub

Private Sub AuditBrandRedConfined()
    Dim w As Worksheet, cell As Range, n As Long
    Head "BRAND RED CONFINED          [reskin 3]"

    Set w = Sh("Setup")
    If w Is Nothing Then Skp "brand red confined to the tab", "no Setup sheet": Exit Sub

    On Error Resume Next
    For Each cell In w.UsedRange
        If cell.Font.Color = C_RED Then n = n + 1
    Next cell
    On Error GoTo 0

    Say "  cells on Setup using the brand red as a font colour: " & n
    If n = 0 Then
        Pas "brand red appears on no Setup cell (tab strip only)"
    Else
        Fal "brand red appears on no Setup cell", n & " cell(s) use it as a font colour"
    End If
End Sub

'---------------- plumbing (never raises) ----------------

Private Function Sh(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set Sh = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
End Function

Private Function TabColour(w As Worksheet) As Long
    On Error Resume Next
    TabColour = CLng(w.Tab.Color)
    On Error GoTo 0
End Function

Private Function KeyRow(w As Worksheet, ByVal key As String) As Long
    Dim i As Long, lastR As Long
    On Error Resume Next
    lastR = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastR
        If LCase(Trim(CStr(w.Cells(i, 1).Value))) = LCase(key) Then
            KeyRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

Private Function NextKeyRowAt(w As Worksheet, ByVal fromRow As Long) As Long
    Dim i As Long, lastR As Long
    On Error Resume Next
    lastR = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    For i = fromRow + 1 To lastR
        If Len(Trim(CStr(w.Cells(i, 1).Value))) > 0 Then
            NextKeyRowAt = i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

' First key row in column A that is a package field key - i.e. starts "k_" but is
' not one of Core's own k_core_* keys.
Private Function FirstPackageKeyRow(w As Worksheet) As Long
    Dim i As Long, lastR As Long, v As String
    On Error Resume Next
    lastR = w.Cells(w.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastR
        v = LCase(Trim(CStr(w.Cells(i, 1).Value)))
        If Left$(v, 2) = "k_" And Left$(v, 7) <> "k_core_" Then
            FirstPackageKeyRow = i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

Private Function RowContaining(w As Worksheet, ByVal col As Long, ByVal txt As String) As Long
    Dim i As Long, lastR As Long
    On Error Resume Next
    lastR = w.UsedRange.Row + w.UsedRange.Rows.Count - 1
    For i = 1 To lastR
        If InStr(1, CStr(w.Cells(i, col).Value), txt, vbTextCompare) > 0 Then
            RowContaining = i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function

Private Function FirstCellWithFill(w As Worksheet, ByVal clr As Long) As Range
    Dim cell As Range
    On Error Resume Next
    For Each cell In w.UsedRange
        If cell.Interior.Color = clr Then
            Set FirstCellWithFill = cell
            Exit For
        End If
    Next cell
    On Error GoTo 0
End Function

Private Sub Num(ByVal label As String, ByVal actual As Long, ByVal expected As Long)
    If actual = expected Then
        Pas label & "  (" & actual & ")"
    Else
        Fal label, "got " & actual & ", expected " & expected
    End If
End Sub

Private Sub Pas(ByVal label As String)
    mPass = mPass + 1
    Say "  PASS  " & label
End Sub

Private Sub Fal(ByVal label As String, ByVal why As String)
    mFail = mFail + 1
    Say "  FAIL  " & label & "  -  " & why
End Sub

Private Sub Skp(ByVal label As String, ByVal why As String)
    mSkip = mSkip + 1
    Say "  SKIP  " & label & "  -  " & why
End Sub

Private Sub Head(ByVal t As String)
    Say ""
    Say t
End Sub

Private Sub Say(ByVal s As String)
    mRep = mRep & s & vbCrLf
End Sub

Private Sub WriteReport()
    Dim w As Worksheet, lines() As String, i As Long
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("_GateAudit").Delete
    Application.DisplayAlerts = True
    Set w = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    w.Name = "_GateAudit"
    lines = Split(mRep, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        w.Cells(i + 1, 1).Value = "'" & lines(i)
    Next i
    w.Columns(1).ColumnWidth = 100
    w.Activate
    On Error GoTo 0
    MsgBox "PASS " & mPass & "    FAIL " & mFail & "    SKIP " & mSkip & vbCrLf & vbCrLf & _
           "Full report is on the _GateAudit sheet.", vbInformation, "Gate audit"
End Sub
