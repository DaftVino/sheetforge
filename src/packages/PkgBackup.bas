Attribute VB_Name = "PkgBackup"
Option Explicit

'================================================================
' PkgBackup - restore points and archival. Wave 5 of the catalog
' build-out (docs/macro-catalog-plug-and-play.md, section 9).
'
' Four actions, three of which are the same underlying move - copy
' THIS workbook, exactly as it stands right now, to a path that
' does not yet exist - with a different destination-naming rule
' each:
'   backup.timestamped_copy    a manual "back this up right now"
'   backup.snapshot             an automatic pre-risky-step restore point
'   backup.month_end_version    named for the last closed calendar month
' The fourth MOVES a file that already exists rather than copying
' one:
'   backup.archive_output       relocate a generated output into Archive\yyyy_mm
' It is NOT this package's only destructive filesystem operation -
' backup.snapshot's own pruning step below also deletes files,
' validated against the exact name shape this package writes before
' anything is Kill'd.
'
' backup.snapshot writes a restore point but this package has no
' restore action. See the plan this shipped from
' (docs/designs/2026-07-28-w5-backup-plan.md) for why that is
' deliberate and which catalog items it does NOT unblock.
'
' Depends on: Core (required). No cross-package calls.
'================================================================

'================================================================
' MANIFEST
'================================================================
' Fields and sections only. Tasks 3-6 each insert ONE Core_AddAction
' call into this Sub, immediately above End Sub, in the same task
' that adds that action's handler - so this Sub never declares an
' action this build's entry-point gate cannot find a handler for.
Public Sub PkgBackup_Manifest(reg As Object)
    Core_DeclarePackage reg, "backup", "Protect & Back Up", "1.0.0"

    Core_AddSection reg, "BACKUP & SNAPSHOT"
    Core_AddField reg, "k_backup_folder", "Folder to write backup and snapshot files into", "text", "", _
        "simple", _
        "Leave blank to use the folder this workbook is saved in - that is the usual answer. " & _
        "Otherwise a full path, e.g. C:\Reports\Backups. Backups go in a 'Backup' subfolder here, " & _
        "snapshots in 'Backup\Snapshots', and archived files in 'Archive\yyyy_mm'.", _
        "backup.timestamped_copy,backup.snapshot,backup.month_end_version,backup.archive_output"
    Core_AddField reg, "k_backup_base", "Base name for a backup file", "text", "Backup", _
        "simple", _
        "Used to build the saved file's name, e.g. 'Backup 2026-07-28_143059.xlsm'.", _
        "backup.timestamped_copy,backup.month_end_version"
    Core_AddField reg, "k_backup_offset", "Date shift for month-end (days)", "int", "0", _
        "advanced", _
        "Days to shift today's date before finding the calendar month it falls in, then " & _
        "stepping back one month for the prior month end. 0 = use today as-is. Same " & _
        "convention as Export Dashboard's date shift.", _
        "backup.month_end_version"
    Core_AddField reg, "k_backup_snapshot_label", "Snapshot label", "text", "", _
        "simple", _
        "A short tag added to a snapshot's file name, e.g. 'before Import'. Leave blank for no tag.", _
        "backup.snapshot"
    Core_AddField reg, "k_backup_snapshot_keep", "Snapshots to keep", "int", "20", _
        "advanced", _
        "How many of the most recent snapshots to keep. Older ones are deleted " & _
        "automatically each time a new snapshot is taken. Nothing else in that folder is touched.", _
        "backup.snapshot"

    Core_AddSection reg, "ARCHIVE"
    Core_AddField reg, "k_backup_archive_prefix", "Generated output to archive (name starts with...)", "file", "", _
        "simple", _
        "Name-prefix of the file this action relocates. Matched in the folder this workbook " & _
        "is saved in, same rule as any other 'file' field: the newest file whose name starts " & _
        "with this text.", _
        "backup.archive_output"

    Core_AddAction reg, "backup.timestamped_copy", "Backup Now (Timestamped Copy)", "PkgBackup_TimestampedCopy", _
        "Saves a full copy of this workbook, exactly as it stands right now, into a 'Backup' " & _
        "subfolder with a datetime stamp in the file name. This workbook itself is untouched.", _
        "Protect & Back Up"
    Core_AddAction reg, "backup.snapshot", "Snapshot Before a Risky Step", "PkgBackup_Snapshot", _
        "Saves a full copy of this workbook as a restore point, meant to run right before a " & _
        "step that changes data. There is no automatic restore - open the snapshot file " & _
        "directly to recover from it.", _
        "Protect & Back Up"
    Core_AddAction reg, "backup.month_end_version", "Save Month-End Version", "PkgBackup_MonthEndVersion", _
        "Saves a full copy of this workbook named for the most recently closed calendar " & _
        "month, for finance packs that keep one file per close.", _
        "Protect & Back Up"
    Core_AddAction reg, "backup.archive_output", "Archive a Generated Output", "PkgBackup_ArchiveOutput", _
        "Moves the newest file matching the configured prefix into an 'Archive\yyyy_mm' " & _
        "subfolder. Refuses if the file is open in Excel or a file of that name is already " & _
        "archived there.", _
        "Protect & Back Up"
End Sub

'================================================================
' Package-private helpers
'================================================================

' Creates folderPath (and any missing parent) if it does not already exist. Neither
' MkDir nor FileSystemObject.CreateFolder can create more than one new level at a
' time - CreateFolder("A\B") fails if A does not already exist - so a fresh
' deployment's very first backup run needs each level built up one at a time, or an
' action that is supposed to OWN this folder fails on the folder it owns.
Private Sub PkgBackup_EnsureFolder(ByVal folderPath As String)
    Dim fso As Object, sep As String, p As String, parts() As String, i As Long, built As String
    sep = Application.PathSeparator
    p = folderPath
    If Right$(p, 1) = sep Then p = Left$(p, Len(p) - 1)
    If Core_FolderExists(p) Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    parts = Split(p, sep)
    built = parts(LBound(parts))                  ' e.g. "C:" - the drive root, always exists
    For i = LBound(parts) + 1 To UBound(parts)
        built = built & sep & parts(i)
        If Not fso.FolderExists(built) Then fso.CreateFolder built
    Next i
End Sub

' This workbook's own file extension, INCLUDING the dot, e.g. ".xlsm". A backup or
' snapshot must be written in the SAME format the live workbook is in - a macro-
' enabled book copied out as .xlsx would silently drop every macro, which defeats
' the point of a restore point.
Private Function PkgBackup_Ext() As String
    Dim nm As String, dot As Long
    nm = ThisWorkbook.name
    dot = InStrRev(nm, ".")
    If dot > 0 Then PkgBackup_Ext = Mid$(nm, dot) Else PkgBackup_Ext = ".xlsm"
End Function

' NOTE: a package-private "user-typed label reduced to something safe inside a
' Windows file name" helper used to live here. It is gone - the actions below call
' Core_SafeFileName (added to Core by 2026-07-28-w3-format-export-plan.md, Task 2,
' Step 2) instead, so this package does not carry a second copy of the identical
' job. See "Escalations to Core" above and "Unknowns and risks" below for the one
' behavior gap that removal opened up (no length cap on the token) and how each
' call site closes it.

' TRUE if nm matches this package's OWN file-name shape exactly:
' "SNAPSHOT_" & yyyy-mm-dd & "_" & HHnnss & optional "_<label>" & ext. A prefix
' match alone is NOT proof of ownership - a user's own "SNAPSHOT_notes.txt" starts
' with the same nine characters - so every field is checked: the literal prefix, a
' timestamp that actually parses as a real calendar date and clock time, and the
' extension this workbook itself uses. Anything that fails any one of those checks
' is left alone, never treated as this package's file.
Private Function PkgBackup_IsSnapshotName(ByVal nm As String, ByVal ext As String) As Boolean
    Const PFX As String = "SNAPSHOT_"
    Const STAMPLEN As Long = 17                     ' "yyyy-mm-dd_HHnnss"
    Dim stamp As String, datePart As String, timePart As String

    If Len(nm) < Len(PFX) + STAMPLEN + Len(ext) Then Exit Function
    If StrComp(Left$(nm, Len(PFX)), PFX, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(Right$(nm, Len(ext)), ext, vbTextCompare) <> 0 Then Exit Function

    stamp = Mid$(nm, Len(PFX) + 1, STAMPLEN)
    If Mid$(stamp, 5, 1) <> "-" Or Mid$(stamp, 8, 1) <> "-" Or Mid$(stamp, 11, 1) <> "_" Then Exit Function
    datePart = Left$(stamp, 10)                      ' "yyyy-mm-dd"
    timePart = Mid$(stamp, 12, 2) & ":" & Mid$(stamp, 14, 2) & ":" & Mid$(stamp, 16, 2)
    If Not IsDate(datePart) Then Exit Function        ' rejects an impossible calendar date, e.g. 2026-02-30
    If Not IsDate(timePart) Then Exit Function         ' rejects an impossible clock time, e.g. 25:99:99

    PkgBackup_IsSnapshotName = True
End Function

' Deletes the oldest files matching this package's own SNAPSHOT_ name shape in
' folderPath (see PkgBackup_IsSnapshotName), keeping the newest keepCount by file
' MODIFIED time (not by parsing the name - the label segment makes lexical order
' unreliable). Only files that pass the full shape check are ever candidates for
' deletion; a bare "SNAPSHOT_*" prefix match is never enough on its own.
'
' A single locked or in-use old snapshot must not fail the run that just wrote a
' GOOD new one, so Kill runs under Resume Next - a prune that skips one file this
' time is a smaller problem than a snapshot action that reports failure after it
' already did the part that matters.
Private Sub PkgBackup_PruneSnapshots(ByVal folderPath As String, ByVal keepCount As Long)
    Dim f As String, names() As String, times() As Date, n As Long, i As Long, j As Long
    Dim tmpS As String, tmpD As Date, ext As String
    Const MAXFILES As Long = 500
    If keepCount < 1 Then keepCount = 1

    ext = PkgBackup_Ext()
    ReDim names(1 To MAXFILES)
    ReDim times(1 To MAXFILES)
    n = 0
    f = Dir(folderPath & "SNAPSHOT_*")
    Do While Len(f) > 0
        If PkgBackup_IsSnapshotName(f, ext) Then
            If n >= MAXFILES Then Exit Do            ' hard cap: never scan without bound
            n = n + 1
            names(n) = folderPath & f
            times(n) = FileDateTime(folderPath & f)
        End If
        f = Dir
    Loop
    If n <= keepCount Then Exit Sub

    ' Insertion sort, oldest first. n is small (bounded by keepCount plus however
    ' many snapshots ran since the last prune), so O(n^2) costs nothing here and
    ' needs no extra library.
    For i = 2 To n
        tmpS = names(i): tmpD = times(i)
        j = i - 1
        Do While j >= 1
            If times(j) <= tmpD Then Exit Do
            names(j + 1) = names(j): times(j + 1) = times(j)
            j = j - 1
        Loop
        names(j + 1) = tmpS: times(j + 1) = tmpD
    Next i

    On Error Resume Next
    For i = 1 To n - keepCount
        Kill names(i)
    Next i
    On Error GoTo 0
End Sub

'================================================================
' ACTION - Save a timestamped backup copy right now
'================================================================
Public Sub PkgBackup_TimestampedCopy(ctx As Object)
    Dim root As String, destFolder As String, destPath As String, baseName As String
    Const TITLE As String = "Backup Now (Timestamped Copy)"
    On Error GoTo Fail

    root = Core_ResolveFolder(Core_Cfg(ctx, "k_backup_folder"))
    If Len(root) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to back up " & _
               "into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write backup and " & _
               "snapshot files into' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' 60 is this caller's cap, not a property of the helper - it bounds the one
    ' free-typed segment of this path so an unusually long configured base name
    ' cannot push the whole path past MAX_PATH, and so a base name containing a
    ' path separator or ".." cannot place the copy outside the Backup folder this
    ' action promises.
    baseName = Core_SafeFileName(Core_Cfg(ctx, "k_backup_base"), 60)
    If Len(baseName) = 0 Then baseName = "Backup"
    destFolder = root & "Backup" & Application.PathSeparator
    destPath = destFolder & baseName & " " & Format(Now, "yyyy-mm-dd_HHnnss") & PkgBackup_Ext()

    ' Checked BEFORE Core_SpeedOn: the cheapest failure should be the earliest one,
    ' and a refusal after the workbook copy has already started reads as a bug.
    If Len(Dir(destPath)) > 0 Then
        Core_Refuse ctx, "A backup file already exists at:" & vbCrLf & vbCrLf & "   " & destPath & _
               vbCrLf & vbCrLf & "This can happen if this action runs twice in the same second. " & _
               "Wait a moment and run it again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgBackup_EnsureFolder destFolder
    ThisWorkbook.SaveCopyAs destPath
    Core_SpeedOff

    Core_Done ctx, "Saved a timestamped backup copy to:" & vbCrLf & vbCrLf & "   " & destPath & _
           vbCrLf & vbCrLf & "This workbook itself was not touched - its name, path, and any " & _
           "unsaved edits are exactly as they were before this ran.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "backup.timestamped_copy", Err
End Sub

'================================================================
' ACTION - Save a restore-point snapshot before a risky step
'================================================================
' NO RESTORE. This writes a full copy of the workbook and prunes old copies; it does
' NOT read one back. See docs/designs/2026-07-28-w5-backup-plan.md for why that is
' deliberate for this wave and which catalog items it does NOT unblock as a result -
' do not extend this action's completion message to imply a restore capability it
' does not have.
Public Sub PkgBackup_Snapshot(ctx As Object)
    Dim root As String, destFolder As String, destPath As String
    Dim label As String, keepCount As Long
    Const TITLE As String = "Snapshot Before a Risky Step"
    On Error GoTo Fail

    root = Core_ResolveFolder(Core_Cfg(ctx, "k_backup_folder"))
    If Len(root) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to save a " & _
               "snapshot into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to write " & _
               "backup and snapshot files into' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' 40 is this caller's cap, not a property of the helper. The rest of the name
    ' ("SNAPSHOT_" + a 17-character timestamp + the extension) is fixed, so bounding
    ' the one free-typed segment is what keeps the whole path inside MAX_PATH.
    label = Core_SafeFileName(Core_Cfg(ctx, "k_backup_snapshot_label"), 40)
    destFolder = root & "Backup" & Application.PathSeparator & "Snapshots" & Application.PathSeparator
    destPath = destFolder & "SNAPSHOT_" & Format(Now, "yyyy-mm-dd_HHnnss") & _
               IIf(Len(label) = 0, "", "_" & label) & PkgBackup_Ext()

    If Len(Dir(destPath)) > 0 Then
        Core_Refuse ctx, "A snapshot file already exists at:" & vbCrLf & vbCrLf & "   " & destPath & _
               vbCrLf & vbCrLf & "This can happen if this action runs twice in the same second. " & _
               "Wait a moment and run it again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgBackup_EnsureFolder destFolder
    ThisWorkbook.SaveCopyAs destPath
    keepCount = Core_CfgLong(ctx, "k_backup_snapshot_keep")
    PkgBackup_PruneSnapshots destFolder, keepCount
    Core_SpeedOff

    Core_Done ctx, "Snapshot saved:" & vbCrLf & vbCrLf & "   " & destPath & vbCrLf & vbCrLf & _
           "This is a full copy of the workbook exactly as it stood the moment this ran, " & _
           "including any edits not yet saved to the main file." & vbCrLf & vbCrLf & _
           "There is no automatic restore for this snapshot. To use it, close this workbook " & _
           "and open the snapshot file directly, or copy what you need from it by hand.", TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "backup.snapshot", Err
End Sub

'================================================================
' ACTION - Save a version named for the prior closed month
'================================================================
Public Sub PkgBackup_MonthEndVersion(ctx As Object)
    Dim root As String, destFolder As String, destPath As String, baseName As String
    Dim refDate As Date, priorMonthEnd As Date
    Const TITLE As String = "Save Month-End Version"
    On Error GoTo Fail

    root = Core_ResolveFolder(Core_Cfg(ctx, "k_backup_folder"))
    If Len(root) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so it has no folder to save a " & _
               "month-end version into." & vbCrLf & vbCrLf & "Save it first, or set 'Folder to " & _
               "write backup and snapshot files into' on the Setup tab." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' "Prior month end" is computed from the CALENDAR MONTH of the reference date,
    ' not by counting back a fixed number of days - the 1st and the 28th of the
    ' same month must give the SAME answer (the month that closed before this one),
    ' and a fixed-day-count shortcut ("today minus 30") does not: it lands in a
    ' different month depending on where in the current month "today" falls, and it
    ' gets shorter months wrong. DateSerial(y, m, 1) - 1 depends only on refDate's
    ' YEAR and MONTH, never its day, and needs no January special case:
    ' DateSerial(y, 1, 1) - 1 rolls back to 31 December of y-1 on its own.
    refDate = Date + Core_CfgLong(ctx, "k_backup_offset")
    priorMonthEnd = DateSerial(Year(refDate), Month(refDate), 1) - 1

    ' 60 is this caller's cap, not a property of the helper - see Task 3's identical
    ' use of Core_SafeFileName for k_backup_base for the reasoning (MAX_PATH and no
    ' path separators/".." escaping the Backup folder).
    baseName = Core_SafeFileName(Core_Cfg(ctx, "k_backup_base"), 60)
    If Len(baseName) = 0 Then baseName = "Backup"
    destFolder = root & "Backup" & Application.PathSeparator
    destPath = destFolder & Format(priorMonthEnd, "yyyy-mm-dd") & " " & baseName & _
               " Month End" & PkgBackup_Ext()

    If Len(Dir(destPath)) > 0 Then
        Core_Refuse ctx, "A month-end version already exists at:" & vbCrLf & vbCrLf & "   " & destPath & _
               vbCrLf & vbCrLf & "This action never overwrites one. Rename or remove the " & _
               "existing file, then run this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgBackup_EnsureFolder destFolder
    ThisWorkbook.SaveCopyAs destPath
    Core_SpeedOff

    Core_Done ctx, "Saved the month-end version for " & Format(priorMonthEnd, "mmmm yyyy") & _
           " to:" & vbCrLf & vbCrLf & "   " & destPath, TITLE
    Exit Sub
Fail:
    Core_Fail ctx, "backup.month_end_version", Err
End Sub

'================================================================
' ACTION - Relocate a generated output into a dated archive folder
'================================================================
Public Sub PkgBackup_ArchiveOutput(ctx As Object)
    Dim prefix As String, srcPath As String
    Dim root As String, destFolder As String, destPath As String
    Dim openWb As Workbook, fso As Object
    Dim moveErrNum As Long, moveErrDesc As String
    Dim srcStillThere As Boolean, destNowThere As Boolean, destBytes As Long
    Const TITLE As String = "Archive a Generated Output"
    On Error GoTo Fail

    If Len(ThisWorkbook.Path) = 0 Then
        Core_Refuse ctx, "This workbook has never been saved, so there is no folder to find a " & _
               "generated file in." & vbCrLf & vbCrLf & "Save this workbook first, then run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    prefix = Trim(Core_Cfg(ctx, "k_backup_archive_prefix"))
    If Len(prefix) = 0 Then
        Core_Refuse ctx, "No file prefix is set for 'Generated output to archive'." & vbCrLf & vbCrLf & _
               "Fill it in on the Setup tab, Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' Core_FindFile is the SAME "file" kind lookup k_master_prefix already uses:
    ' newest match in THIS workbook's own folder, per contract section 3. It is not
    ' scoped by k_backup_folder - that field only says where the ARCHIVED COPY
    ' lands, not where the source is searched for.
    srcPath = Core_FindFile(prefix)
    If Len(srcPath) = 0 Then
        Core_Refuse ctx, "No file starting with '" & prefix & "' was found in:" & vbCrLf & vbCrLf & _
               "   " & ThisWorkbook.Path & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' A file cannot be safely moved while Excel holds it open: Name and
    ' FileSystemObject.MoveFile both fail against a locked file, and failing
    ' partway through a move can leave neither a clean source nor a clean
    ' destination. Refused here, before anything on disk changes - see
    ' Core_AlreadyOpen's own header for why the probe is by file name. This only
    ' catches THIS Excel instance; a different process or a different application
    ' holding the file is not visible to it, which is exactly why the state check
    ' after the move below exists.
    Set openWb = Core_AlreadyOpen(srcPath)
    If Not openWb Is Nothing Then
        Core_Refuse ctx, "'" & Core_FileName(srcPath) & "' is currently open in Excel, and this " & _
               "action cannot move a file that is open." & vbCrLf & vbCrLf & "Close it, then run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    ' ThisWorkbook.Path is already proven non-blank above, so Core_ResolveFolder
    ' cannot return "" here even with k_backup_folder left blank.
    root = Core_ResolveFolder(Core_Cfg(ctx, "k_backup_folder"))
    destFolder = root & "Archive" & Application.PathSeparator & Format(Date, "yyyy_mm") & _
                 Application.PathSeparator
    destPath = destFolder & Core_FileName(srcPath)

    If Len(Dir(destPath)) > 0 Then
        Core_Refuse ctx, "A file named '" & Core_FileName(srcPath) & "' already exists in:" & _
               vbCrLf & vbCrLf & "   " & destFolder & vbCrLf & vbCrLf & "This action never " & _
               "overwrites one. Rename or remove the existing file, then run this again." & _
               vbCrLf & vbCrLf & "Nothing has been changed.", TITLE
        Exit Sub
    End If

    Core_SpeedOn
    PkgBackup_EnsureFolder destFolder
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' NOT atomic: across volumes this is copy-then-delete, not one step, so a
    ' failure partway through can leave a deleted source and an incomplete or
    ' absent destination. The error (if any) is captured here rather than left to
    ' jump straight to Fail, because the state check right after this needs to run
    ' either way - "no error raised" is not by itself proof the move went cleanly.
    On Error Resume Next
    fso.MoveFile srcPath, destPath                 ' not the Name statement: MoveFile
                                                     ' works across drives, Name does not
    moveErrNum = Err.Number
    moveErrDesc = Err.Description
    Err.Clear
    On Error GoTo Fail
    Core_SpeedOff

    srcStillThere = (Len(Dir(srcPath)) > 0)
    destNowThere = (Len(Dir(destPath)) > 0)
    If destNowThere Then destBytes = FileLen(destPath)

    If moveErrNum = 0 And Not srcStillThere And destNowThere And destBytes > 0 Then
        Core_Done ctx, "Archived:" & vbCrLf & vbCrLf & "   " & srcPath & vbCrLf & "to" & vbCrLf & _
               "   " & destPath, TITLE
        Exit Sub
    End If

    Err.Raise vbObjectError + 1030, "PkgBackup_ArchiveOutput", _
        "The move did not finish cleanly (this is a copy-then-delete across drives, not one " & _
        "atomic step, so a partial result is possible)." & _
        IIf(moveErrNum <> 0, " Underlying error: " & moveErrDesc & ".", "") & _
        " Source (" & srcPath & "): " & IIf(srcStillThere, "still present.", "gone.") & _
        " Destination (" & destPath & "): " & _
        IIf(destNowThere, IIf(destBytes > 0, "present, " & destBytes & " bytes.", "present but 0 bytes."), "not created.") & _
        " Check both locations by hand before running this again."
    Exit Sub
Fail:
    Core_Fail ctx, "backup.archive_output", Err
End Sub
