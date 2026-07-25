Attribute VB_Name = "PkgFixture"
Option Explicit

'================================================================
' PkgFixture - TEST HARNESS, NEVER SHIPPED.
'
' Declares 74 no-op actions across 9 display categories so the v0.4.0 packed
' auto-layout can be exercised at catalog scale while only 8 real actions exist
' (design section 2.5). Category sizes are deliberately uneven so one Excel
' session covers: several small categories packed onto one page, a part-full
' column, an exactly-full column, a wrap to a second column, a category
' needing a full CARD_COLS (3) columns on one page, and a category larger
' than a whole page.
'
' Task 9 correction: the brief was written when a card column held 8 buttons.
' Task 7 lowered AUTO_COL_MAX to 6 (contract section 6 amended to match - the
' screen-fit arithmetic showed 8 only fit a full-HD, 100%-scaled, maximised
' window, while a common 125% display scaling fits about 5.7). The category
' sizes below are recomputed against AUTO_COL_MAX = 6, CARD_COLS = 3, not the
' brief's original 8/9/12/26 - see task-9-report.md for the worked-out paging
' prediction against these exact numbers plus the two real categories.
'
' Note "CLEAN & STANDARDIZE" below: it deliberately COLLIDES with PkgClean's
' "Clean & Standardize" display name, so two different modules feed one
' category. That is not a mistake - it is the one test that two packages can
' feed a single category, the whole point of decoupling category from module,
' and it must be verified, not assumed. The case difference is ALSO
' deliberate (fix round 1): the categories Dictionary in BuildAutoLayout is
' set vbTextCompare specifically so two packages spelling a category with
' different casing still merge, and a byte-identical spelling would have
' passed that test even under a plain binary compare, proving nothing about
' the case-insensitive path. Bonus: because PkgClean already contributes 3
' items, this fixture's 5 pushes the MERGED category to 8 items - past
' AUTO_COL_MAX, so the collision also exercises a wrap caused by two packages
' combining, not just a same-package category that happens to be large. The
' merged column's header still reads "Clean & Standardize" once (the
' first-seen casing, from PkgClean, which probes before PkgFixture) - not
' twice under two different spellings.
'
' Every action binds to the SAME entry point. Nothing in the contract requires
' a distinct sub per action - the ctx carries actionId - so 74 actions cost
' one sub.
'
' Help-string lengths (fix round 1) deliberately vary instead of sharing one
' boilerplate sentence: action 1 is empty (tests the card grid's height floor
' and the index catalog's un-floored AutoFit on the very first card), and the
' rest cycle through short / typical-sentence / long-paragraph bands so a
' single Excel session shows several different description heights on several
' different pages, not just one. See FixtureHelpText below.
'
' Excluded from the default bundle by build/build.ps1; build it deliberately with
'   pwsh build/build.ps1 -Packages PkgImportExport,PkgClean,PkgFixture -Out dist/MacrosBundle.fixture.bas
' DELETE this module, its CoreCandidates entry and its build.ps1 exclusion once
' the polish pass is verified.
'================================================================

Public Sub PkgFixture_Manifest(reg As Object)
    Dim cats(1 To 9) As String, counts(1 To 9) As Long
    Dim ci As Long, k As Long, n As Long

    Core_DeclarePackage reg, "fixture", "Fixture (test only)", "0.0.1"

    ' Sizes recomputed for AUTO_COL_MAX = 6 (was 8 in the brief). Each comment
    ' names the packing case this category's size is chosen to exercise -
    ' see task-9-report.md for the full page-by-page prediction.
    cats(1) = "Ingest & Import":      counts(1) = 3    ' part-full column (1 col, 3/6)
    cats(2) = "CLEAN & STANDARDIZE":  counts(2) = 5    ' collides (case-insensitively) with PkgClean's "Clean & Standardize" (3) -> merged 8, wraps to 2 cols
    cats(3) = "Validate":             counts(3) = 5    ' small, packs onto a page with others
    cats(4) = "Reconcile":            counts(4) = 9    ' wrap to a second (partial) column: 6 + 3
    cats(5) = "Report":               counts(5) = 14   ' needs a full CARD_COLS (3) columns: 6 + 6 + 2
    cats(6) = "Export":               counts(6) = 4    ' small, packs onto a page with others
    cats(7) = "Period Close":         counts(7) = 6    ' exactly-full column (1 col, 6/6, no continuation)
    cats(8) = "Admin & Housekeeping": counts(8) = 2    ' small, packs onto a page with others
    cats(9) = "Bulk Fixture Set":     counts(9) = 26   ' larger than a whole page (needs 5 cols > CARD_COLS)

    n = 0
    For ci = 1 To 9
        For k = 1 To counts(ci)
            n = n + 1
            Core_AddAction reg, "fixture.a" & n, "Fixture " & n & " - " & cats(ci), _
                           "PkgFixture_Noop", FixtureHelpText(n, cats(ci)), cats(ci)
        Next k
    Next ci

    ' Fields that exercise the collapsed off band. fx_* keys cannot collide with
    ' the shipped k_* / clean_* namespaces.
    Core_AddSection reg, "FIXTURE SETTINGS"
    Core_AddField reg, "fx_always", "Always shown (declares no usedBy)", "text", "always", _
                  "simple", "Package-wide: stays in place whatever is turned off."
    Core_AddField reg, "fx_a1", "Used by Fixture 1 only", "text", "one", _
                  "simple", "Collapses into the off band when Fixture 1 is set to Off.", _
                  "fixture.a1"
    Core_AddField reg, "fx_a2a3", "Used by Fixture 2 or Fixture 3", "text", "two", _
                  "simple", "Stays up while EITHER action is runnable; collapses only when both are Off.", _
                  "fixture.a2,fixture.a3"
    Core_AddField reg, "fx_chain", "Used by Fixture 7 (test 'Chains only')", "text", "chain", _
                  "simple", "Fixture 7 set to 'Chains only' still runs, so this must STAY in the main list.", _
                  "fixture.a7"
    Core_AddField reg, "fx_typo", "Names an action id that does not exist", "text", "safe", _
                  "simple", "An unknown usedBy id must count as in use, never hide the field.", _
                  "fixture.no_such_action"
    Core_AddField reg, "fx_adv", "Advanced, used by Fixture 4", "int", "7", _
                  "advanced", "An advanced field must move into the off band too.", _
                  "fixture.a4"
    Core_AddList reg, "fx_list", "List used by Fixture 5", "Value", _
                 "simple", "A list block must survive being collapsed.", "fixture.a5"
    Core_AddTable reg, "fx_table", "Table used by Fixture 6", "First|Second|Third", _
                  "simple", "A table block must survive being collapsed.", "fixture.a6"
    Core_AddDefault reg, "fx_list", "seed one"
    Core_AddDefault reg, "fx_list", "seed two"
    Core_AddDefault reg, "fx_table", "a1|b1|c1"
End Sub

' The single entry point for all 74 fixture actions. Routing through Core_Done
' (not MsgBox) means chain-step suppression is exercised too.
Public Sub PkgFixture_Noop(ctx As Object)
    Core_Done ctx, "Fixture action ran: " & CStr(ctx("actionId")), "Fixture"
End Sub

' Fix round 1: one shared boilerplate sentence left the card grid's
' AutoFit-with-a-floor behavior (and HomeCatalog's un-floored AutoFit)
' exercised at exactly one description length. Action 1 is the sole
' deliberately-empty description (a bare card row must still hold its floor
' height; the catalog row must still line up, not collapse). Every other
' action cycles through three bands - short, a typical one-line sentence, and
' a long multi-sentence paragraph meant to wrap several lines in a narrow
' card column - so several different pages show several different heights in
' one session instead of one uniform block.
Private Function FixtureHelpText(ByVal n As Long, ByVal catName As String) As String
    Dim bandIx As Long
    If n = 1 Then
        FixtureHelpText = ""             ' the one deliberately-empty description
        Exit Function
    End If
    bandIx = (n - 2) Mod 3
    Select Case bandIx
        Case 0                            ' short: a few words
            FixtureHelpText = "No-op fixture action " & n & "."
        Case 1                            ' typical: one plain sentence
            FixtureHelpText = "No-op fixture action " & n & " in category " & catName & _
                               " - exists only to exercise the layout, not to do real work."
        Case Else                         ' long: several sentences, meant to wrap 4-5 lines
            FixtureHelpText = "No-op fixture action " & n & " in category " & catName & _
                               ". This description is deliberately long, spanning several " & _
                               "sentences instead of one, so the card grid's autofit-with-a-floor " & _
                               "behavior gets exercised at real length instead of just a single " & _
                               "one-liner: on a card page this should wrap across several lines " & _
                               "without ever clipping, and on the Macros Home catalog page - which " & _
                               "has no per-row height floor - it should also wrap cleanly rather " & _
                               "than collapsing or overlapping the row beneath it."
    End Select
End Function
