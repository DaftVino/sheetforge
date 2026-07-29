Attribute VB_Name = "Core"
Option Explicit

'================================================================
' Core - the required engine every package depends on.
'
' Extracted from SetupImportExportMacros (src/legacy/SetupImportExportMacros.bas):
' the config engine, Setup renderer, control-panel/button builder and
' shared helpers are lifted here and generalized so a package no longer
' hardcodes any of them. A package is now a thin module that:
'   1. describes itself in a manifest  (Pkg<Name>_Manifest reg)
'   2. implements small action subs     (Pkg<Name>_<Verb> ctx)
' Core does everything else - Setup sheet, buttons, config, speed, errors.
'
' Contract (never break these):
'   * Everything Public is Core_*.  Everything else is Private.
'   * Discovery is zero-trust: Core probes "Pkg<Name>_Manifest" with
'     Application.Run, never enumerates modules (no "Trust access to the
'     VBA project object model").
'   * Two objects flow through the system, both Scripting.Dictionary:
'       reg  - a package's manifest (id/name/version + fields + actions)
'       ctx  - per-action context (actionId + its package's reg + log)
'
' Object model
'   reg("id"), reg("name"), reg("version")
'   reg("fields")  = Collection of field dictionaries, in render order
'   reg("actions") = Collection of action dictionaries
'   field: key, label, kind, default, headers, defaults(Collection)
'          kind = "section" | "text" | "int" | "sheet" | "file" |
'                 "bool" | "list" | "table"
'   action: id, caption, entry, help
'   ctx("actionId"), ctx("reg"), ctx("log")
'================================================================

' ---- sheet names ----
Private Const SETUP_SHEET    As String = "Setup"
Private Const CONFIG_SHEET   As String = "_Config"     ' hidden, written by Finalize
' Where Finalize stages the new config before swapping it into place (issue #8).
' Only ever exists DURING an activate; if you find one in a workbook, an activate was
' interrupted and CfgSheet will adopt it. Must stay <= 31 chars (Excel sheet-name limit).
Private Const CONFIG_STAGE   As String = "_Config_staged"
Private Const MACRO_SHEET    As String = "Macros"
Private Const ADMIN_SHEET    As String = "Setup Menu"   ' Create/Finalize controls; hideable
Private Const PACKAGES_SHEET As String = "_Packages"   ' optional enable/disable layer

Private Const CHAIN_CAT      As String = "Chains"      ' reserved display category for chain buttons

' ---- per-action availability (contract v1.2.0; default flipped v2.1.0) ----
' Stored in the action catalog block as the display label ("Button" / "Chains
' only" / "Off") and parsed to these tokens. Blank no longer reads as Button:
' as of v2.1.0 it is not an answer at all and blocks Finalize - see ParseAvail.
Private Const AVAIL_BUTTON   As String = "button"
Private Const AVAIL_CHAIN    As String = "chains only"
Private Const AVAIL_OFF      As String = "off"

' Provenance sentinel inside an availability map - NOT an action id. Reserved by
' contract section 7 (v2.1.0) as part of the core.* namespace, because section 7
' states no character rule for action ids and a sentinel therefore cannot be
' assumed collision-proof. Present => the map came from a pre-v1.2 block whose
' Availability column does not exist, so an absent id means Button, not Off.
Private Const AVAIL_LEGACY_KEY As String = "core.legacy"

' Column header caption for the action catalog's Availability column.
' RenderPlacementBlock writes this literal as the visible header; MarkerVerdict's
' tamper test (case-insensitively) reads it back to decide whether an absent
' schema marker means "pre-v1.2 layout" (recoverable) or "this v1.2-shaped block
' was tampered with" (must not be trusted). The two must stay tied to the same
' Const - if the caption is reworded in only one place, the tamper test silently
' stops firing, a tampered store degrades to "absent", ReadAvailMap returns an
' empty map, and every action the user switched Off goes live again.
Private Const AVAIL_HEADER   As String = "Availability"

' Column header caption for the action catalog's and the Chains block's placement
' column. Introduced as a Const in v0.4.0 because the caption was written from
' two independent literals - the exact coupling closed for AVAIL_HEADER above.
'
' Unlike AVAIL_HEADER this one is NOT load-bearing, and that was checked rather
' than assumed: no check anywhere reads this caption back. The old caption
' ('Page code') appeared only as a written value, at two sites. Kept as a
' Const anyway, because two literals for one visible string is how they drift.
'
' Relabelled from the old caption ('Page code') because a placement value is no
' longer only a page code: it may also be a "Sheet!Cell" anchor onto a user's
' own tab.
Private Const PLACE_HEADER   As String = "Placement"

' Config-format marker, written into the hidden RID column of the action catalog's
' key row - a machine-owned cell, not a visible header. A display label is
' presentation, is user-editable, and must never decide whether saved config is
' preserved or discarded (contract section 5). Task 4 builds the migration on it;
' it is declared and read here because ReadAvailMap needs it.
' The "v" prefix is load-bearing: a bare "1.2" is stored as General and Excel
' coerces it to the NUMBER 1.2, which is then read back through a locale-aware
' CStr (a comma-decimal locale yields "1,2"; some locales read it as a date).
' A non-numeric token cannot be coerced, so the marker survives the round trip
' byte-for-byte regardless of the workbook's locale.
'
' v2.0 (contract v2.0.0): PkgClean's two per-selection fields were withdrawn, which
' is a BREAKING contract change (a removed key). The stored BLOCK LAYOUT is in fact
' unchanged - same three columns, same key row - so bumping the major here refuses
' every config written by v0.4.0, which is stricter than the data requires.
'
' That was accepted deliberately, by the owner, on the ground that no deployment
' exists yet: the bundle is only in test workbooks, so there is no installed base
' to refuse. Keeping this marker coupled to the contract version is simpler than
' the alternative, which was to decouple the two numbers permanently and explain
' the split in every port. See docs/adr/0002 (rejected, with that reasoning).
'
' If an installed base ever exists, revisit before the next breaking bump: at that
' point refusing every stored config to express a change that touched no stored
' byte is a real cost, and ADR 0002's decoupling becomes the better trade.
'
' Note the asymmetry this creates and do not mistake it for a bug: a v0.3.0 config
' has NO marker, reads as "absent", and still migrates cleanly. A v0.4.0 config has
' "v1.2", reads as "other-major", and is refused. Older is accepted, newer-but-not-
' current is not, because only the marker-less case has a defined legacy layout.
Private Const SCHEMA_VER     As String = "v2.0"

' ---- button page geometry (v0.4.0: an unmerged CELL-CARD grid) ----
' Cards live in worksheet columns B, D and F, with C and E as narrow gutters.
' NOTHING on this grid may be merged: Excel's AutoFit ignores merged cells (which
' is why the Setup band rows have to estimate their height instead), and AutoFit
' on the description rows is the entire reason a description cannot clip.
Private Const PANEL_BASE As String = "Macros "   ' grouped output tabs: "Macros 1", "Macros 2"...
Private Const HOME_SHEET As String = "Macros Home"  ' generated action index; RESERVED by contract 7
Private Const CARD_COLS   As Long = 3            ' card columns per page -> B, D, F
Private Const CARD_W      As Double = 30         ' content column width, in characters
Private Const GUTTER_W    As Double = 3          ' spacer column width, in characters
Private Const CARD_HDR_ROW As Long = 3           ' category header row; cards start at 4
                                                 ' (rows 2/4 blank chrome spacers were removed -
                                                 ' fix round 1: title now row 2, header row 3)
Private Const BTN_H      As Double = 28          ' fix round 1: 26 clipped a 2-line wrapped caption
Private Const BTN_ROW_H  As Double = 32          ' button row height (BTN_H + padding)
Private Const DESC_MIN_H As Double = 24          ' floor after AutoFit
' No DESC_MAX_H: contract 6 forbids clipping a description, and a manual clamp on
' an AutoFit row is exactly how Excel clips. AutoFit gets the honest height, full
' stop, for an action's AUTHORED help. A chain's help is engine-SYNTHESISED
' (GatherButtons: "Chain: " & StepsInline(...)), not authored, so contract 6 does
' not bind it - CardHelp (carry-over C, Task 7) bounds it before it reaches the
' card grid instead of letting a pathological 20-step chain draw one tall row
' that inflates every card column beside it; the full recipe stays authored on
' the Setup tab's Chains block.
Private Const AUTO_COL_MAX As Long = 6           ' buttons per card column before a new column
                                                 ' (fix round 1: was 8; 8 only fits full-HD/100%)
Private Const CHAIN_CARD_STEPS As Long = 4       ' step captions shown on a card before "+N more"
' Fix round 1, fix 6: CHAIN_CARD_STEPS alone bounds step COUNT, not length - four
' long captions (an action caption may carry an ordinal and a full phrase) sail
' through untouched and can still reproduce the ~170pt row carry-over C exists to
' prevent. CHAIN_CARD_CHARS bounds the composed width directly; CardHelp stops
' adding steps once either cap is hit, whichever comes first.
Private Const CHAIN_CARD_CHARS As Long = 120     ' composed-text length cap before "+N more" cuts in
' Fix round 2, fix 6: separates a real colIx (from ParseCode, or the 9998/9999
' reserved columns) from a SYNTHETIC colIx minted by SplitTallManualColumns for
' a same-code stack's vertical overflow (piece colIx = origColIx * VSPLIT_MULT +
' pieceIndex). Large enough that no realistic placement code's column (up to a
' 3-letter code, max 18278) collides with it; SplitTallManualColumns itself
' guards the multiply against Long overflow for anything larger.
Private Const VSPLIT_MULT As Long = 100000
' ---- Sheet!Cell anchored buttons (v0.4.0) ----
' Shape-name prefix for every button this engine places on a user's OWN sheet.
' Contract section 7 RESERVES it normatively ("Reserved (v1.2.0): ... the shape-name
' prefix SFPlaced_ ... anything else in the workbook using the prefix is deleted
' without warning"). That reservation HAS LANDED - it is a documented guarantee, not
' a convention this engine merely happens to keep - so do not weaken, rename or
' shorten this prefix without amending spec/contract.md section 7 in the same change.
' SweepPlacedButtons finds and removes buttons by this prefix alone - it is the
' only handle the engine has on a sheet it must otherwise never touch.
Private Const PLACED_PREFIX As String = "SFPlaced_"
' Width of an anchored button, in points. Deliberately NOT the target cell's own
' width: a user's column may be three characters wide. Height reuses BTN_H.
' 110pt is the owner's call at the v0.4.0 gate (the design proposed 150pt), sized
' to stay close to one default-width column so a button is not intrusive sitting
' next to real data.
Private Const ANCHOR_BTN_W  As Double = 110
' ---- Setup / _Config layout ----
Private Const KEY_COL   As Long = 1       ' column A: machine keys (hidden)
Private Const VAL_COL   As Long = 3       ' column C: scalar value / list value / 1st table col
Private Const MAX_LIST  As Long = 20      ' input rows drawn per list / table block
Private Const END_KEY   As String = "_core_end"   ' sentinel key row: bounds the last list/table block
Private Const CHAINS_KEY As String = "k_core_chains"   ' Compose section: chains block sentinel (col A)
Private Const LAYOUT_KEY As String = "k_core_layout"   ' Compose section: placement block sentinel (col A)
Private Const RID_COL    As Long = 10                  ' Compose section: col J, hidden per-row id
Private Const REMINDER_TEXT As String = "Actions are switched on. Click 'Finalize Setup Tab', then 'Create Setup Tab' to show their option fields. Fill those in, then click 'Finalize Setup Tab' again to save them."

' ---- speed wrapper saved state (module-level, depth-counted) ----
Private mSpeedDepth  As Long
Private mSpeedActive As Boolean
Private mSaveScreen  As Boolean
Private mSaveEvents  As Boolean
Private mSaveAlerts  As Boolean
Private mSaveCalc    As XlCalculation

' Fix round 1, fix 2: Core_BuildControlPanel's own Fail: handler shows its
' error and then returns NORMALLY (a Sub can't propagate an error out), so a
' caller has no way to tell success from failure by the call alone.
' Core_FinalizeSetup needs to, because it shows its own unconditional
' "activated and rebuilt" message right after calling it. Reset to False at
' the very top of every Core_BuildControlPanel call (not just on entry to the
' module) so a stale True from a PRIOR successful call can never leak into a
' later call that fails before reaching its own success line - and only set
' True once the whole pipeline (through the generated home page) has actually finished.
Private mLastBuildOk As Boolean

' Final-review fix wave, fix 1: the "N button(s) placed / N could NOT be placed"
' note Core_BuildControlPanel composes, carried out to Core_FinalizeSetup.
' Finalize calls Core_BuildControlPanel SILENT and prints its own success message,
' so a note composed only for the non-silent MsgBox is invisible on the NORMAL user
' path - and that is exactly the path an anchor onto a protected, renamed or deleted
' sheet fails on. The skip was therefore reported to nobody: no button on the sheet,
' no button on any generated page (SplitAnchored keeps an anchor out of "paged"),
' and a plain "Configuration activated" dialog. Reset to "" at the very top of every
' Core_BuildControlPanel call, alongside mLastBuildOk, so a note from a PRIOR
' successful build can never leak into a later call that fails before composing one.
Private mLastPlacedNote As String

' The zero-pages sentence, in ONE place because two callers say it: the direct-run
' MsgBox in Core_BuildControlPanel and Core_FinalizeSetup's own success dialog. Two
' literals would be two places to disagree.
Private Const NO_PAGES_MSG As String = "No macro pages were built - every button is " & _
    "hidden, switched Off, or anchored to one of your own sheets."

' Same class of bug as mLastPlacedNote above, and found the same way - by a human
' running the gate. Finalize calls Core_BuildControlPanel SILENT, so the zero-pages
' sentence reached nobody on the normal path: a user who switches every action Off
' through Finalize was told "The Macros pages have been rebuilt to match your Setup
' answers" and then found no pages at all. Zero pages is a legitimate, deliberately
' reachable configuration, so the dialog has to say the true thing about it. Reset
' False at the very top of every Core_BuildControlPanel call, alongside mLastBuildOk,
' so a verdict from a PRIOR build can never leak into a later one.
Private mLastNoPages As Boolean

' ---- memoized package registry (v0.4.0 performance) ----
' BuildRegistry re-runs CoreCandidates, one Application.Run "Pkg<Name>_Manifest"
' per candidate, and a _Packages worksheet scan per candidate. It is called from
' RenderOffBand (twice), HasAnyTier's per-package loops, ActionCaptionFor and
' ActionTokenFor ONCE PER CHAIN STEP, StepsInline once per step, and ActionIdSet
' TWICE inside a single ComposeReport - on top of the main render loops. At the
' 82-action fixture that is thousands of manifest rebuilds, and the owner
' measured the first Setup build in minutes.
'
' Handing the SAME Collection to every caller is safe because no consumer mutates
' a reg. Every mutation of a reg dictionary in this module is inside
' Core_DeclarePackage / Core_AddSection / Core_AddField / Core_AddList /
' Core_AddTable / Core_AddDefault / Core_AddAction, all of which run DURING
' TryManifest, inside BuildRegistry itself, against a dictionary it has just
' created. All thirteen "For Each reg In BuildRegistry()" sites only read.
'
' IF YOU ADD A CONSUMER THAT MUTATES A reg, this cache becomes a silent
' cross-contamination bug: the mutation persists into every later caller in the
' same build. Deep-copy in that consumer, or call InvalidateRegistry after it.
' That includes code OUTSIDE this module: NewCtx stores the cached reg straight
' into every action context as ctx("reg"), so a package handler holding a ctx holds
' a live pointer into this cache. Core_Cfg / Core_CfgLong / Core_CfgList /
' Core_CfgTable / Core_CfgBool only read through it, and a package must do the same
' - ctx("reg") is READ-ONLY to a handler. It is the one route by which code this
' module does not own could corrupt the cache.
' Do not "just try it" - nothing in this repo's tooling can see the failure.
Private mRegCache   As Collection
Private mRegCacheOk As Boolean

'================================================================
' IDENTITY PALETTE  (v0.9.0 - issue #11)
'================================================================
' Sampled from docs/designs/sheetforge-design-proof.png. That proof is a DARK-first
' mobile concept and NONE of its product surface is in scope - it is a colour
' reference only. These are its values adapted to the surfaces Core actually draws,
' which sit on a WHITE sheet (Core_BuildSetup sets Cells.Interior white).
'
' This replaces TWO unrelated identities that had drifted apart: purple on the Setup
' tab (112,48,160 / 54,32,100 / 90,70,120) and blue on the Macros pages (31,78,121).
' One palette now covers both.
'
' Functions, not Const: VBA rejects `Const c As Long = RGB(...)` ("constant
' expression required") because RGB is a function call, and hard-coding the packed
' Long instead (FORGE_NAVY would be 3022353) is unreadable and unreviewable.
'
' SEMANTIC colours are deliberately NOT in here and were NOT changed - they carry
' meaning rather than identity, and folding them into a brand palette would erase
' the distinction: the yellow input cells (255,242,204) that mean "fill this in",
' and the two warning reds (150,0,0 caution, 192,0,0 error). The brand red is close
' enough to those that reusing it for chrome would make a warning unreadable AS a
' warning - so the brand red is confined to the Setup tab strip, where no text sits
' on it.
'
' Contrast on white, all well past WCAG AA for body text:
'   ForgeNavy  15.1:1     ForgeSteel  10.6:1
' and on ForgeTint, the band fill:
'   ForgeNavy  13.0:1     ForgeSteel   9.1:1
Private Function ForgeRed() As Long        ' brand accent - fills only, never behind text
    ForgeRed = RGB(187, 30, 32)
End Function
Private Function ForgeNavy() As Long       ' headings and titles
    ForgeNavy = RGB(17, 30, 46)
End Function
Private Function ForgeSteel() As Long      ' secondary/help text, and the Macros tab strip
    ForgeSteel = RGB(49, 60, 74)
End Function
Private Function ForgeTint() As Long       ' band + card header fill
    ForgeTint = RGB(226, 232, 238)
End Function
Private Function ForgeRule() As Long       ' hairline borders
    ForgeRule = RGB(214, 221, 228)
End Function

'================================================================
' 1. MANIFEST-BUILDING API  (called inside Pkg<Name>_Manifest reg)
'================================================================
Public Sub Core_DeclarePackage(reg As Object, ByVal id As String, _
                               ByVal name As String, ByVal version As String)
    reg.RemoveAll
    reg.Add "id", id
    reg.Add "name", name
    reg.Add "version", version
    reg.Add "fields", New Collection
    reg.Add "actions", New Collection
End Sub

Public Sub Core_AddSection(reg As Object, ByVal title As String)
    reg("fields").Add NewField("", title, "section", "", "simple", "", "")
End Sub

' tier is "simple" (default) or "advanced". Advanced fields are grouped under an
' "ADVANCED - most users can leave these as-is" band so regular users can stop
' after the simple fields. help is an optional plain-language note shown in grey
' under the field (pass via help:= to skip tier). usedBy is an optional
' comma-separated list of action ids this field configures; a field whose every
' named action is Off moves into the collapsed band at the bottom of Setup (never
' omitted - see Core_BuildSetup). A "Chains only" action still runs, so it still
' keeps its fields in the working area.
Public Sub Core_AddField(reg As Object, ByVal key As String, ByVal label As String, _
                         ByVal kind As String, ByVal dflt As String, _
                         Optional ByVal tier As String = "simple", _
                         Optional ByVal help As String = "", _
                         Optional ByVal usedBy As String = "")
    reg("fields").Add NewField(key, label, kind, dflt, tier, help, usedBy)
End Sub

Public Sub Core_AddList(reg As Object, ByVal key As String, _
                        ByVal label As String, ByVal colHeader As String, _
                        Optional ByVal tier As String = "simple", _
                        Optional ByVal help As String = "", _
                        Optional ByVal usedBy As String = "")
    Dim f As Object
    Set f = NewField(key, label, "list", "", tier, help, usedBy)
    f("headers") = colHeader
    reg("fields").Add f
End Sub

' colHeaders is a "|"-delimited string (NOT comma - a header may itself contain
' commas, e.g. "Type map (T,N,D,...)"). Up to 3 columns are read (C, D, E).
Public Sub Core_AddTable(reg As Object, ByVal key As String, _
                         ByVal label As String, ByVal colHeaders As String, _
                         Optional ByVal tier As String = "simple", _
                         Optional ByVal help As String = "", _
                         Optional ByVal usedBy As String = "", _
                         Optional ByVal colNotes As String = "")
    Dim f As Object
    Set f = NewField(key, label, "table", "", tier, help, usedBy)
    f("headers") = colHeaders
    f("notes") = colNotes
    reg("fields").Add f
End Sub

' Seed a fallback row for a list ("value") or table ("a|b|c") field.
' Call once per default row so nothing hits VBA's 25-continuation limit.
Public Sub Core_AddDefault(reg As Object, ByVal key As String, ByVal rowCsv As String)
    Dim f As Object
    Set f = FindField(reg, key)
    If Not f Is Nothing Then f("defaults").Add rowCsv
End Sub

' category is the DISPLAY grouping used by the column headers on a button page -
' and, in auto-layout, by which page a button lands on. (It also fed the generated
' index's category column, but that index was withdrawn in v0.4.0; see
' src/attic/HomeIndex.bas.txt. The button pages are its only live consumer now.)
' It defaults to the package's display name, so an
' action that declares nothing keeps today's grouping. It is deliberately
' decoupled from the module: two packages may feed one category, and one package
' may span several, so a future package split never moves a button in the user's
' mental model.
Public Sub Core_AddAction(reg As Object, ByVal actionId As String, ByVal caption As String, _
                          ByVal entryPoint As String, ByVal help As String, _
                          Optional ByVal category As String = "")
    Dim a As Object, cat As String
    Set a = CreateObject("Scripting.Dictionary")
    a.Add "id", actionId
    a.Add "caption", caption
    a.Add "entry", entryPoint
    a.Add "help", help
    cat = Trim(category)
    If Len(cat) = 0 Then cat = CStr(reg("name"))
    a.Add "category", cat
    reg("actions").Add a
End Sub

Private Function NewField(ByVal key As String, ByVal label As String, _
                          ByVal kind As String, ByVal dflt As String, _
                          ByVal tier As String, ByVal help As String, _
                          ByVal usedBy As String) As Object
    Dim f As Object
    Set f = CreateObject("Scripting.Dictionary")
    f.Add "key", key
    f.Add "label", label
    f.Add "kind", kind
    f.Add "default", dflt
    f.Add "headers", ""
    f.Add "defaults", New Collection
    f.Add "help", help
    ' Comma-separated action ids this field configures; "" = package-wide (always
    ' shown). Stored lowercased so FieldInUse can compare without re-normalizing.
    f.Add "usedBy", LCase(Trim(usedBy))
    ' Normalize the tier so an unknown or mis-cased value degrades to "simple"
    ' (still shown) rather than matching neither render pass and vanishing.
    Dim t As String
    t = LCase(Trim(tier))
    If t <> "advanced" Then t = "simple"
    f.Add "tier", t
    Set NewField = f
End Function

Private Function FindField(reg As Object, ByVal key As String) As Object
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If LCase(CStr(f("key"))) = LCase(key) Then
                Set FindField = f
                Exit Function
            End If
        End If
    Next f
End Function

' Does this field belong to the user's working area, given the availability map?
' True when the field is package-wide (no usedBy) or when at least ONE action it
' names is not Off - a "Chains only" action still runs, so it still needs its
' settings. False means "relegate to the collapsed band" - never "omit":
' Core_BuildSetup deletes and re-seeds Setup from the FINALIZED _Config, so
' omitting a field would cost toggle -> Finalize -> rebuild -> fill -> Finalize
' and would discard anything typed but not yet finalized (design section 2.3).
' An id that no installed package declares counts as in-use, so a typo in usedBy
' cannot hide a field forever.
'
' "Unknown" is decided against the REGISTRY, not against availMap (v2.1.0). It
' used to be "absent from availMap", which was equivalent while absence meant
' Button - both said "in use". Once absence means Off they diverge, and reading
' unknown from the map broke disclosure exactly where it matters most: a fresh
' workbook has an EMPTY map, so every real id looked unknown, so every field with
' a usedBy stayed in the main band while its action rendered Off. Absence is now
' resolved by AvailFor via IsRunnable, which already knows fresh from legacy.
Private Function FieldInUse(f As Object, availMap As Object) As Boolean
    Dim ids() As String, i As Long, one As String, entry As String
    FieldInUse = True
    If f Is Nothing Then Exit Function
    If Not f.Exists("usedBy") Then Exit Function
    If Len(Trim(CStr(f("usedBy")))) = 0 Then Exit Function
    ids = Split(CStr(f("usedBy")), ",")
    For i = LBound(ids) To UBound(ids)
        one = LCase(Trim(ids(i)))
        If Len(one) > 0 Then
            If FindActionOwner(one, entry) Is Nothing Then Exit Function  ' unknown id -> in use
            If IsRunnable(availMap, one) Then Exit Function    ' one runnable is enough
        End If
    Next i
    FieldInUse = False
End Function

'================================================================
' 2. CONFIG-READ API  (called inside action subs, given ctx)
'    Value comes from the active _Config sheet; if that is absent or
'    the cell is blank, the package's manifest default is returned.
'================================================================
Public Function Core_Cfg(ctx As Object, ByVal key As String) As String
    Core_Cfg = ScalarValue(ctx("reg"), key)
End Function

Public Function Core_CfgLong(ctx As Object, ByVal key As String) As Long
    Dim s As String
    s = Core_Cfg(ctx, key)
    ' A non-blank but non-numeric cell must not silently become 0; fall back to
    ' the field's manifest default (matches the legacy typed-wrapper defaults).
    If Not IsNumeric(s) Then s = FieldDefaultReg(ctx("reg"), key)
    If IsNumeric(s) Then Core_CfgLong = CLng(s)
End Function

Public Function Core_CfgList(ctx As Object, ByVal key As String) As Collection
    Set Core_CfgList = ListValues(ctx("reg"), key)
End Function

Public Function Core_CfgTable(ctx As Object, ByVal key As String, _
                              ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    Core_CfgTable = TableValues(ctx("reg"), key, a, b, c)
End Function

' Reads a "bool" field (contract section 3; clean_hasheader is its first use).
'
' A BLANK cell means "never set" and resolves to the field's DECLARED manifest
' default - not a hardcoded False and not an optional caller-supplied fallback.
' That fallback is already free: Core_Cfg -> ScalarValue returns
' FieldDefaultReg(ctx("reg"), key) when the stored value is empty. The declared
' default is on the manifest; two places to state it would be two places to
' disagree.
'
' An unrecognised NON-blank value fails closed to False. It cannot normally get
' this far - Core_FinalizeSetup blocks activation on one, naming the field - but
' a reader that guessed would be exactly the "user typed 'ye' and got a silent
' no" failure the availability parser exists to prevent.
Public Function Core_CfgBool(ctx As Object, ByVal key As String) As Boolean
    Dim flag As Boolean
    If ParseBool(Core_Cfg(ctx, key), flag) Then Core_CfgBool = flag
End Function

' ---- setup-helper API: read/write the LIVE Setup sheet before Finalize ----
' For "setup helper" actions (e.g. suggest column types) that operate on what
' the user is editing on the Setup tab, not the activated _Config. These encode
' the Setup layout so a package never hardcodes column positions.

' Read a 3-column table block (tab/prefix/type map style) from the Setup sheet.
Public Function Core_SetupTable(ByVal key As String, _
                                ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    Core_SetupTable = ReadTable3(SetupSheet(), key, a, b, c)
End Function

' Write one column of a block on the Setup sheet: colOffset 0=C, 1=D, 2=E.
' Values are applied, in order, to the block's populated rows (those with a
' non-blank first column) - the same rows Core_SetupTable returns.
Public Sub Core_SetupSetColumn(ByVal key As String, ByVal colOffset As Long, values As Collection)
    Dim ws As Worksheet, r As Long, nextK As Long, i As Long, n As Long
    Set ws = SetupSheet()
    If ws Is Nothing Then Exit Sub
    r = FindKeyRow(ws, key)
    If r = 0 Then Exit Sub
    nextK = NextKeyRow(ws, r)
    For i = r + 1 To nextK - 1
        If Trim(CStr(ws.Cells(i, VAL_COL).Value)) <> "" Then
            n = n + 1
            If n <= values.Count Then ws.Cells(i, VAL_COL + colOffset).Value = values(n)
        End If
    Next i
End Sub

Private Function SetupSheet() As Worksheet
    On Error Resume Next
    Set SetupSheet = ThisWorkbook.Worksheets(SETUP_SHEET)
    On Error GoTo 0
End Function

' ---- value resolution: sheet read, else manifest default ----
Private Function ScalarValue(reg As Object, ByVal key As String) As String
    Dim v As String
    v = ReadScalar(CfgSheet(), key)
    If v = "" Then v = FieldDefaultReg(reg, key)
    ScalarValue = v
End Function

Private Function ListValues(reg As Object, ByVal key As String) As Collection
    Dim col As Collection
    Set col = ReadList(CfgSheet(), key)
    If col.Count = 0 Then Set col = DefaultListReg(reg, key)
    Set ListValues = col
End Function

Private Function TableValues(reg As Object, ByVal key As String, _
                             ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    If ReadTable3(CfgSheet(), key, a, b, c) Then
        TableValues = True
        Exit Function
    End If
    TableValues = DefaultTable3Reg(reg, key, a, b, c)
End Function

Private Function FieldDefaultReg(reg As Object, ByVal key As String) As String
    Dim f As Object
    Set f = FindField(reg, key)
    If Not f Is Nothing Then FieldDefaultReg = CStr(f("default"))
End Function

Private Function DefaultListReg(reg As Object, ByVal key As String) As Collection
    Dim f As Object, v As Variant, col As New Collection
    Set DefaultListReg = col
    Set f = FindField(reg, key)
    If f Is Nothing Then Exit Function
    For Each v In f("defaults")
        col.Add CStr(v)
    Next v
End Function

Private Function DefaultTable3Reg(reg As Object, ByVal key As String, _
                                  ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    Dim f As Object, n As Long, i As Long, parts() As String
    Set f = FindField(reg, key)
    If f Is Nothing Then Exit Function
    n = f("defaults").Count
    If n = 0 Then Exit Function
    ReDim a(1 To n)
    ReDim b(1 To n)
    ReDim c(1 To n)
    For i = 1 To n
        parts = Split(CStr(f("defaults")(i)), "|")
        a(i) = PartAt(parts, 0)
        b(i) = PartAt(parts, 1)
        c(i) = PartAt(parts, 2)
    Next i
    DefaultTable3Reg = True
End Function

' ---- raw sheet readers (config source = _Config, or any given sheet) ----
' The ONE accessor for the activated config - and, since issue #8, the one place
' that recovers an interrupted activate.
'
' Finalize activates by copying Setup to CONFIG_STAGE, verifying it, then deleting
' the old _Config and renaming the stage into place. A crash in the gap between that
' delete and that rename leaves a workbook holding a COMPLETE config under the stage
' name and no _Config at all. Adopting it here means every reader recovers without
' knowing anything about staging - and, critically, that Core_BuildSetup does not
' read "no config yet" and re-seed Setup with defaults ON TOP of a config that is
' sitting right there. That silent overwrite is the actual data loss in #8; the
' missing sheet is only how it starts.
'
' Adoption is safe because a stage sheet never outlives a SUCCESSFUL activate: the
' success path renames it to CONFIG_SHEET, and Finalize deletes a stale one before
' every copy. Its existence alongside a missing _Config is therefore proof of an
' interrupted swap, not ambiguity.
'
' The repair is best-effort: if the rename is refused (protected workbook), the sheet
' is still returned, so reading keeps working even when healing cannot.
Private Function CfgSheet() As Worksheet
    On Error Resume Next
    Set CfgSheet = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0
    If Not CfgSheet Is Nothing Then Exit Function

    On Error Resume Next
    Set CfgSheet = ThisWorkbook.Worksheets(CONFIG_STAGE)
    On Error GoTo 0
    If CfgSheet Is Nothing Then Exit Function
    On Error Resume Next
    CfgSheet.name = CONFIG_SHEET
    CfgSheet.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Function

Private Function FindKeyRow(ws As Worksheet, ByVal key As String) As Long
    Dim lastR As Long, r As Long
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.Rows.Count, KEY_COL).End(xlUp).Row
    For r = 1 To lastR
        If LCase(Trim(CStr(ws.Cells(r, KEY_COL).Value))) = LCase(key) Then
            FindKeyRow = r
            Exit Function
        End If
    Next r
End Function

' A list/table block runs from the key row until the next row that
' carries ANY key in column A (prefix-independent, unlike the legacy
' "k_" test), or the end of the used keys.
Private Function NextKeyRow(ws As Worksheet, ByVal afterRow As Long) As Long
    Dim lastR As Long, r As Long
    lastR = ws.Cells(ws.Rows.Count, KEY_COL).End(xlUp).Row
    For r = afterRow + 1 To lastR
        If Trim(CStr(ws.Cells(r, KEY_COL).Value)) <> "" Then
            NextKeyRow = r
            Exit Function
        End If
    Next r
    NextKeyRow = lastR + 1
End Function

Private Function ReadScalar(ws As Worksheet, ByVal key As String) As String
    Dim r As Long
    If ws Is Nothing Then Exit Function
    r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    ReadScalar = Trim(CStr(ws.Cells(r, VAL_COL).Value))
End Function

Private Function ReadList(ws As Worksheet, ByVal key As String) As Collection
    Dim r As Long, nextK As Long, i As Long, v As String, col As New Collection
    Set ReadList = col
    If ws Is Nothing Then Exit Function
    r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    nextK = NextKeyRow(ws, r)
    For i = r + 1 To nextK - 1
        v = Trim(CStr(ws.Cells(i, VAL_COL).Value))
        If v <> "" Then col.Add v
    Next i
End Function

Private Function ReadTable3(ws As Worksheet, ByVal key As String, _
                            ByRef a() As String, ByRef b() As String, ByRef c() As String) As Boolean
    Dim r As Long, nextK As Long, i As Long, n As Long, k As Long
    If ws Is Nothing Then Exit Function
    r = FindKeyRow(ws, key)
    If r = 0 Then Exit Function
    nextK = NextKeyRow(ws, r)
    For i = r + 1 To nextK - 1
        If Trim(CStr(ws.Cells(i, VAL_COL).Value)) <> "" Then n = n + 1
    Next i
    If n = 0 Then Exit Function
    ReDim a(1 To n)
    ReDim b(1 To n)
    ReDim c(1 To n)
    For i = r + 1 To nextK - 1
        If Trim(CStr(ws.Cells(i, VAL_COL).Value)) <> "" Then
            k = k + 1
            a(k) = Trim(CStr(ws.Cells(i, VAL_COL).Value))
            b(k) = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
            c(k) = Trim(CStr(ws.Cells(i, VAL_COL + 2).Value))
        End If
    Next i
    ReadTable3 = True
End Function

'================================================================
' 3. SHARED HELPERS  (Public so packages reuse them; single home here)
'================================================================

' ---- speed wrapper: disable, do work, guaranteed restore ----
Private Sub Core_SpeedOn()
    If mSpeedDepth = 0 Then
        mSaveScreen = Application.ScreenUpdating
        mSaveEvents = Application.EnableEvents
        mSaveAlerts = Application.DisplayAlerts
        mSaveCalc = Application.Calculation
        mSpeedActive = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    End If
    mSpeedDepth = mSpeedDepth + 1
End Sub

Private Sub Core_SpeedOff()
    If mSpeedDepth > 0 Then mSpeedDepth = mSpeedDepth - 1
    If mSpeedDepth <= 0 Then RestoreSpeed
End Sub

' Restore app state. Uses the saved values when Core_SpeedOn ran; otherwise
' forces sane normals so a failure can never leave Excel frozen.
Private Sub RestoreSpeed()
    On Error Resume Next
    If mSpeedActive Then
        Application.Calculation = mSaveCalc
        Application.DisplayAlerts = mSaveAlerts
        Application.EnableEvents = mSaveEvents
        Application.ScreenUpdating = mSaveScreen
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    mSpeedDepth = 0
    mSpeedActive = False
    On Error GoTo 0
End Sub

' ---- structured error + log ----
Public Sub Core_Log(ctx As Object, ByVal msg As String)
    ctx("log") = ctx("log") & msg & vbCrLf
End Sub

Public Sub Core_Fail(ctx As Object, ByVal actionId As String, e As ErrObject)
    Dim m As String
    m = "#" & e.Number & "  " & e.Description
    RestoreSpeed

    On Error Resume Next
    ctx("failed") = True
    ctx("failNumber") = e.Number
    ctx("failMsg") = e.Description
    On Error GoTo 0

    Core_Log ctx, "FAILED " & actionId & ": " & m
    If Not IsQuiet(ctx) Then _
        MsgBox "Action '" & actionId & "' failed:" & vbCrLf & vbCrLf & m, _
               vbExclamation, "Macro Error"
End Sub

' ---- worksheet / range utilities ----
Public Function Core_SheetExists(ByVal nm As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(nm)
    Core_SheetExists = Not s Is Nothing
    On Error GoTo 0
End Function

' Is this sheet one the engine owns? Packages ask before treating a sheet as user
' data - combining the config store into a master would be meaningless at best.
'
' It lives in Core because the names ARE Core's: the seven constants below are
' Private Const in this module, and a package that hardcodes its own copy is one
' Core rename away from silently folding the Setup sheet into a user's data.
'
' The generated button pages are PANEL_BASE & <n> - "Macros 1", "Macros 2". The
' test requires a NUMERIC tail, not just the prefix: a plain prefix test also
' swallows a user's own sheet called "Macros Data". The accepted trade is that a
' user sheet literally named "Macros 999" is treated as the engine's, which beats
' enumerating the generated pages to find out how many exist.
'
' No Trim: trimming first means a user's own " Setup " is read as the engine's and
' silently dropped. The engine's own sheets carry no stray spaces.
Public Function Core_IsEngineSheet(ByVal nm As String) As Boolean
    Dim tail As String
    Select Case LCase(nm)
        Case LCase(SETUP_SHEET), LCase(ADMIN_SHEET), LCase(CONFIG_SHEET), _
             LCase(CONFIG_STAGE), LCase(PACKAGES_SHEET), LCase(MACRO_SHEET), _
             LCase(HOME_SHEET)
            Core_IsEngineSheet = True
            Exit Function
    End Select
    If LCase(Left$(nm, Len(PANEL_BASE))) = LCase(PANEL_BASE) Then
        tail = Mid$(nm, Len(PANEL_BASE) + 1)
        If Len(tail) > 0 Then Core_IsEngineSheet = Core_IsDigits(tail)
    End If
End Function

' Every character is 0-9. Empty is False.
Public Function Core_IsDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    Core_IsDigits = True
End Function

' The workbook at this path IF it is already open in this Excel instance, else
' Nothing. The caller uses what comes back IN PLACE and does NOT close it.
'
' This exists because Workbooks.Open is not a question - it is an answer. Handed a
' path that is already open it returns the open workbook, ReadOnly:=True and all,
' complete with whatever the user has typed and not yet saved. Closing that with
' SaveChanges:=False discards their work, and under Core_SpeedOn DisplayAlerts is
' False, so the prompt Excel would normally raise ("reopening will discard your
' changes") never reaches them. They would see the action report success and lose an
' afternoon's edits in the same second.
'
' Probed by FILE NAME, because that is how the Workbooks collection is keyed - there
' is no lookup by path. The name alone is not proof of identity, so the path is
' confirmed afterwards: a configured source folder routinely points somewhere other
' than this workbook's folder, and a close folder full of "January.xlsx" is the
' everyday shape of this data. A same-name / different-path collision is raised
' rather than guessed at - Excel will not hold two workbooks of one name open, so
' the action cannot proceed either way, and saying so beats silently reading last
' month's file.
'
' Lives in Core because two packages need it. PkgImport found this the expensive
' way at the W1 gate; PkgImportExport's import and updateMaster had the identical
' defect and were only found by reviewing this pass's diff.
Public Function Core_AlreadyOpen(ByVal fullPath As String) As Workbook
    Dim wb As Workbook, nm As String, openPath As String
    nm = Core_FileName(fullPath)
    If Len(nm) = 0 Then Exit Function

    On Error Resume Next
    Set wb = Workbooks(nm)
    On Error GoTo 0
    If wb Is Nothing Then Exit Function          ' not open: the caller opens it itself

    On Error Resume Next
    openPath = wb.FullName
    On Error GoTo 0
    If StrComp(openPath, fullPath, vbTextCompare) = 0 Then
        Set Core_AlreadyOpen = wb
        Exit Function
    End If
    Err.Raise vbObjectError + 1025, "Core_AlreadyOpen", _
        "A different workbook named '" & nm & "' is already open (" & openPath & "). " & _
        "Excel cannot have two workbooks of the same name open at once. Close that one, " & _
        "then run this again."
End Function

' "TabName!M2:M12", or "'TabName Data'!M2:M12" when the sheet name needs quoting.
' Range.Address omits the sheet name BY DESIGN, so every message that names a
' range must qualify it here or a cross-tab run is indistinguishable from a
' same-tab one - which is exactly what the owner hit at the v0.4.0 gate: three
' PkgClean actions reported "M2:M12" whether they had acted on TabName or on the
' button's own tab.
' Lives in Core, not PkgClean: CLAUDE.md rule 4 puts shared helpers only in Core,
' and a copy in a package would be a standards deviation.
Public Function Core_RangeRef(rng As Range) As String
    Dim nm As String
    If rng Is Nothing Then Exit Function
    nm = rng.Parent.Name
    If NeedsSheetQuotes(nm) Then nm = "'" & Replace(nm, "'", "''") & "'"
    Core_RangeRef = nm & "!" & rng.Address(False, False)
End Function

'---- user prompts ----
' These live in Core because a package that asks the user for a range is not doing
' anything package-specific, and because the bundle is ONE module: an unprefixed
' Private AskRange in PkgClean reserves that name against every other package
' (build.ps1's duplicate-procedure guard), so the second package to want one cannot
' copy it and cannot re-declare it. Promoting is the only move that is not a
' standards deviation.

' Prompt the user to select a range. Returns Nothing if they cancel
' (InputBox Type:=8 returns Boolean False on cancel, which fails the Set).
Public Function Core_AskRange(ByVal title As String) As Range
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox( _
        Prompt:="Select the range for '" & title & "':", _
        title:=title, Type:=8)
    On Error GoTo 0
    If Not rng Is Nothing Then
        If rng.Areas.Count > 1 Then
            MsgBox "Please select a single, continuous block (not several separate areas).", _
                   vbExclamation, title
            Set rng = Nothing
        End If
    End If
    Set Core_AskRange = rng
End Function

' Ask whether the picked range's first row is a header. Returns False on Cancel,
' leaving hasHeader untouched and the caller responsible for stopping.
'
' vbYesNoCancel, not vbYesNo: a two-button dialog has no way out, and the title-bar
' X on a vbYesNo MsgBox returns vbNo - which for a caller that deletes rows would
' silently choose "compare every row" and delete a header row. Cancel has to be
' reachable and distinct.
'
' title is a parameter because this used to hardcode "Remove Duplicates" while it
' had exactly one caller. Every other caller would have shown a dialog captioned
' with an action it was not running.
Public Function Core_AskHasHeader(rng As Range, ByRef hasHeader As Boolean, _
                                  Optional ByVal title As String = "Header row") As Boolean
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Does " & Core_RangeRef(rng) & " include a header row?" & vbCrLf & vbCrLf & _
                 "Yes  -  the first row is a title row: it is kept, and never " & _
                 "compared against the data." & vbCrLf & _
                 "No  -  every row is data: all of them are compared, including " & _
                 "the first one.", _
                 vbQuestion + vbYesNoCancel + vbDefaultButton1, title)
    If ans = vbCancel Then Exit Function
    hasHeader = (ans = vbYes)
    Core_AskHasHeader = True
End Function

' Ask for one or more column numbers, counted from the LEFT of rng. Returns False
' on Cancel or on a number outside the range's width; keys comes back as the raw
' "1,2,4" string, which Core_ParseCols turns into numbers.
'
' Application.InputBox, not VBA's InputBox: the latter returns "" for BOTH Cancel
' and an empty OK, and those must not mean the same thing. Application.InputBox
' returns Boolean False on Cancel, which VarType can tell apart from any string.
'
' question and title are parameters for the same reason Core_AskHasHeader's title
' is: this was written for Remove Duplicates and said so in both.
Public Function Core_AskCols(rng As Range, ByRef keys As String, _
                             ByVal question As String, ByVal title As String) As Boolean
    Dim v As Variant, cols As Variant, i As Long, nCols As Long
    nCols = rng.Columns.Count
    v = Application.InputBox( _
        Prompt:=question & vbCrLf & vbCrLf & _
                "Numbered from the LEFT of " & Core_RangeRef(rng) & ", which is " & _
                nCols & " column(s) wide. One number, or several like 1,2,4.", _
        title:=title, Default:="1", Type:=2)
    If VarType(v) = vbBoolean Then Exit Function     ' Cancel
    keys = Trim(CStr(v))
    If Len(keys) = 0 Then keys = "1"                 ' empty OK means the default

    ' Validate against the picked range's width. A column number past the right edge
    ' raises 1004 deep inside whatever the caller does with it, which surfaces through
    ' Core_Fail as an engine error - it reads as a broken macro rather than as the typo
    ' it is. Core_ParseCols already drops non-numeric entries and falls back to 1.
    cols = Core_ParseCols(keys)
    For i = LBound(cols) To UBound(cols)
        If cols(i) < 1 Or cols(i) > nCols Then
            MsgBox "Column " & cols(i) & " is outside " & Core_RangeRef(rng) & _
                   ", which is " & nCols & " column(s) wide." & vbCrLf & vbCrLf & _
                   "Nothing has been changed. Run " & title & " again.", _
                   vbExclamation, title
            Exit Function
        End If
    Next i
    Core_AskCols = True
End Function

' "1,2,4" -> Variant array of Longs, 0-based. Non-numeric entries are ignored;
' an empty result defaults to column 1.
Public Function Core_ParseCols(ByVal s As String) As Variant
    Dim parts() As String, i As Long, c As Collection, out() As Variant, k As Long
    Set c = New Collection
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        If IsNumeric(Trim(parts(i))) Then c.Add CLng(Trim(parts(i)))
    Next i
    If c.Count = 0 Then c.Add CLng(1)
    ReDim out(0 To c.Count - 1)
    For k = 1 To c.Count
        out(k - 1) = c(k)
    Next k
    Core_ParseCols = out
End Function

' Excel's own rule: a sheet name is written bare only when it is entirely
' A-Z / a-z / 0-9 / "_" and does not begin with a digit. Anything else is
' quoted, and an embedded apostrophe is doubled inside the quotes.
Private Function NeedsSheetQuotes(ByVal nm As String) As Boolean
    Dim i As Long, ch As String
    If Len(nm) = 0 Then NeedsSheetQuotes = True: Exit Function
    ch = Left$(nm, 1)
    If ch >= "0" And ch <= "9" Then NeedsSheetQuotes = True: Exit Function
    For i = 1 To Len(nm)
        ch = Mid$(nm, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") _
                Or (ch >= "0" And ch <= "9") Or ch = "_") Then
            NeedsSheetQuotes = True
            Exit Function
        End If
    Next i
End Function

' Robust last data row: End(xlUp) on a column, widened by the used-range
' last cell (SpecialCells), which catches trailing formatted/blank rows.
Public Function Core_LastRow(ws As Worksheet, Optional ByVal col As Long = 1) As Long
    Dim r As Long, ur As Long
    If ws Is Nothing Then Exit Function
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    On Error Resume Next
    ur = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    On Error GoTo 0
    If ur > r Then r = ur
    Core_LastRow = r
End Function

' Data extent by a SINGLE column only (bare End(xlUp), NOT widened by
' SpecialCells). Use when you want exactly the block one column defines - e.g.
' reading a source export - where a widened last cell would pull in trailing
' formatted-but-blank rows. Core_LastRow is the widened variant for clearing.
Public Function Core_LastRowUp(ws As Worksheet, Optional ByVal col As Long = 1) As Long
    If ws Is Nothing Then Exit Function
    Core_LastRowUp = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Function Core_LastCol(ws As Worksheet, Optional ByVal row As Long = 1) As Long
    If ws Is Nothing Then Exit Function
    Core_LastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
End Function

' Every character is A-Z (case-insensitively). Empty is False. Used to tell a
' column LETTER spec apart from a header NAME spec.
Public Function Core_IsAlpha(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = UCase(Mid$(s, i, 1))
        If ch < "A" Or ch > "Z" Then Exit Function
    Next i
    Core_IsAlpha = True
End Function

' A column inside rng, named either by its SPREADSHEET LETTER ("C") or by its
' HEADER TEXT ("Amount"). Returns the column's position WITHIN rng, counting from
' 1 at rng's left edge - or 0 when the spec does not resolve inside rng, so the
' caller reports a bad column rather than acting on the wrong one.
'
' Returns an INDEX, not a Range, even though the PkgClean function it is promoted
' from returned a Range. The index is the primitive: a caller wanting the column
' itself writes rng.Columns(idx), and a caller reading a bulk .Value2 array - which
' is how every engine here reads a real dataset - needs the number and would only
' convert a Range straight back into one. Going the other way is not symmetrical,
' and 0-means-no beats testing Is Nothing on a function that also has a "found it
' but it is outside your range" answer.
'
' The letter reading wins a tie, so a range whose header literally reads "A"
' cannot be addressed by name. That is documented in the consuming field's help
' text rather than resolved, because a heuristic that guessed would be worse.
'
' Header comparison is case- and space-insensitive via StrComp vbTextCompare:
' a user typing "amount" for a column headed "Amount " must not be told their
' column does not exist.
Public Function Core_ResolveColumn(rng As Range, ByVal spec As String) As Long
    Dim s As String, absCol As Long, i As Long, hdr As String
    s = Trim(spec)
    If Len(s) = 0 Then Exit Function

    If Len(s) <= 3 And Core_IsAlpha(s) Then
        On Error Resume Next
        absCol = rng.Parent.Range(s & "1").Column
        On Error GoTo 0
        If absCol > 0 Then
            If absCol >= rng.Column And absCol <= rng.Column + rng.Columns.Count - 1 Then
                Core_ResolveColumn = absCol - rng.Column + 1
            End If
            Exit Function
        End If
    End If

    For i = 1 To rng.Columns.Count
        hdr = Trim(CStr(rng.Cells(1, i).Value))
        If StrComp(hdr, s, vbTextCompare) = 0 Then
            Core_ResolveColumn = i
            Exit Function
        End If
    Next i
End Function

' Excel TRIM (collapse runs of spaces) + CLEAN (strip non-printing characters).
Public Function Core_CleanText(ByVal s As String) As String
    Core_CleanText = Application.WorksheetFunction.Clean(Application.WorksheetFunction.Trim(s))
End Function

' Empty, or the empty string. NOT the same as IsEmpty: a cell holding "" is blank
' to a user and is not Empty to VBA.
Public Function Core_IsBlank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        Core_IsBlank = True
    ElseIf VarType(v) = vbString Then
        Core_IsBlank = (v = "")
    End If
End Function

' Narrow a picked range to the part of the sheet that actually holds anything.
' Returns Nothing when the pick lies entirely outside the used area - the caller
' reports that and stops, having changed nothing.
'
' Selecting whole columns is how people pick "this table" - A:D is one click on the
' column headers - and it hands an action roughly four million cells. Anything doing
' COM work per cell or per row then freezes the screen for minutes. Nothing outside
' UsedRange can hold a value, a formula or a merge, so the intersection examines the
' same data and finishes in seconds.
'
' Callers must use the NARROWED range everywhere downstream, including in the range
' they name in their messages, so what the user is told matches what was looked at.
Public Function Core_NarrowToUsed(rng As Range) As Range
    If rng Is Nothing Then Exit Function
    Set Core_NarrowToUsed = Application.Intersect(rng, rng.Parent.UsedRange)
End Function

Public Function Core_FindCol(ByVal hdr As Variant, ByVal name As String) As Long
    Dim k As Long
    For k = LBound(hdr, 2) To UBound(hdr, 2)
        If Trim(LCase(CStr(hdr(1, k)))) = Trim(LCase(name)) Then
            Core_FindCol = k
            Exit Function
        End If
    Next k
End Function

' ---- file finder: newest file in the workbook folder whose base name
'      = prefix (optionally followed by a " (n)" suffix). Excludes this
'      workbook. Collects names first, then stats - calling FileDateTime
'      inside the Dir() loop can corrupt the iterator on network drives.
Public Function Core_FindFile(ByVal prefix As String) As String
    Dim folder As String, f As String, baseName As String, suffix As String
    Dim names() As String, n As Long, i As Long
    Dim bestFile As String, bestTime As Date, tStamp As Date
    If Trim(prefix) = "" Then Exit Function
    folder = ThisWorkbook.Path
    If Right(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator

    ReDim names(1 To 2000)
    n = 0
    f = Dir(folder & "*.xls*")
    Do While Len(f) > 0
        n = n + 1
        names(n) = f
        f = Dir
    Loop

    For i = 1 To n
        f = names(i)
        If LCase(f) <> LCase(ThisWorkbook.name) Then
            baseName = f
            If InStrRev(baseName, ".") > 0 Then baseName = Left(baseName, InStrRev(baseName, ".") - 1)
            If LCase(Left(baseName, Len(prefix))) = LCase(prefix) Then
                suffix = Trim(Mid(baseName, Len(prefix) + 1))
                If suffix = "" Or suffix Like "(*)" Then
                    tStamp = FileDateTime(folder & f)
                    If bestFile = "" Or tStamp > bestTime Then
                        bestFile = folder & f
                        bestTime = tStamp
                    End If
                End If
            End If
        End If
    Next i
    Core_FindFile = bestFile
End Function

Public Function Core_FileName(ByVal p As String) As String
    If InStrRev(p, "\") > 0 Then
        Core_FileName = Mid(p, InStrRev(p, "\") + 1)
    Else
        Core_FileName = p
    End If
End Function

' Characters Excel forbids in a sheet name. Declared once, here, because it was
' written out twice inside PkgImport alone - once to reject a name and once to
' repair one - and the two copies could drift apart silently.
Public Function Core_BadSheetChars() As String
    Core_BadSheetChars = ":\/?*[]"
End Function

' A destination name this deployment may create. It REFUSES rather than overwrites:
' the library snippet this replaces opened with Worksheets("Master").Delete under
' On Error Resume Next, which silently destroys whatever the user had there.
' Nothing in this toolkit deletes a sheet it did not create in the same run.
'
' Refuses via Core_Refuse, so a chain STOPS here rather than running its remaining
' steps against a sheet this one did not produce (contract section 6).
' fieldLabel is the Setup field the caller reads its name from, quoted back to the
' user verbatim. It is a parameter and not generic text because the message it
' appears in is the one that has to be actionable: "fill in the destination name"
' sends a user hunting down a Setup tab, and "fill in 'Name for the combined sheet
' this creates'" does not. The version this replaces hardcoded one caller's label.
Public Function Core_DestUsable(ctx As Object, ByVal destName As String, _
                                ByVal title As String, _
                                Optional ByVal fieldLabel As String = "") As Boolean
    Dim bad As String, i As Long, whichField As String, s As Object
    whichField = fieldLabel
    If Len(whichField) = 0 Then whichField = "the destination sheet name"
    If Len(destName) = 0 Then
        Core_Refuse ctx, "No name is set for the sheet this creates." & vbCrLf & vbCrLf & _
               "Fill in '" & whichField & "' on the Setup tab, Finalize, and run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", title
        Exit Function
    End If
    If Len(destName) > 31 Then
        Core_Refuse ctx, "'" & destName & "' is " & Len(destName) & " characters. Excel " & _
               "sheet names stop at 31." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", title
        Exit Function
    End If
    bad = Core_BadSheetChars()
    For i = 1 To Len(bad)
        If InStr(destName, Mid$(bad, i, 1)) > 0 Then
            Core_Refuse ctx, "'" & destName & "' contains " & Mid$(bad, i, 1) & ", which " & _
                   "Excel does not allow in a sheet name." & vbCrLf & vbCrLf & _
                   "None of these are allowed:  : \ / ? * [ ]" & vbCrLf & vbCrLf & _
                   "Nothing has been changed.", title
            Exit Function
        End If
    Next i
    ' A name the ENGINE owns is refused even though no such sheet exists yet.
    ' Contract section 7 reserves "Macros", every "Macros <n>", "Macros Home",
    ' "_Config" and "_Config_staged", and the engine DELETES those names
    ' unconditionally with alerts suppressed on every rebuild. So a user who sets
    ' the destination to "Macros 3" gets a sheet that builds, fills with a month of
    ' combined data, and is destroyed without a prompt by the next Finalize. The
    ' exists-check above cannot catch it: the name is free right up until the moment
    ' it matters. Core_IsEngineSheet already knows the whole list.
    If Core_IsEngineSheet(destName) Then
        Core_Refuse ctx, "'" & destName & "' is a name the macro toolkit uses for its own " & _
               "sheets, so it cannot be used for your data." & vbCrLf & vbCrLf & _
               "A sheet by this name would be deleted without warning the next time the " & _
               "Setup tab is finalized. Pick a different name." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", title
        Exit Function
    End If

    ' ThisWorkbook.Sheets, NOT Core_SheetExists. Core_SheetExists looks only at
    ' Worksheets, so a CHART sheet holding this name passes the check and then
    ' Core_NewSheet's rename raises 1004 - after Worksheets.Add has already put an
    ' empty sheet in the workbook. The user gets an engine error and a stray "Sheet1"
    ' from an action whose entire promise is that it changes nothing when it refuses.
    ' Sheets covers every sheet type; a name is taken or it is not.
    On Error Resume Next
    Set s = ThisWorkbook.Sheets(destName)
    On Error GoTo 0
    If Not s Is Nothing Then
        Core_Refuse ctx, "A sheet named '" & destName & "' already exists in this workbook, " & _
               "and this action will not overwrite it." & vbCrLf & vbCrLf & _
               "Rename or delete that sheet, then run " & title & " again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", title
        Exit Function
    End If
    Core_DestUsable = True
End Function

' Add a new sheet at the END of the workbook and name it. Call Core_DestUsable
' FIRST - this does no validating, deliberately, so that the refusal and the
' creation are not tangled into one call whose failure mode is ambiguous.
'
' Written once here because the same two lines were repeated verbatim at two
' places inside PkgImport, and four more waves each need a destination sheet.
'
' The rename is trapped and rolled back. Add-then-rename is two operations, and
' the rename is the one that can fail on a name Core_DestUsable approved - a sheet
' created by another macro in the same second, a name Excel rejects for a reason
' the validator does not model. Without the rollback the failure leaves a stray
' empty "Sheet1" behind and the error surfaces from an action that reported nothing.
' The original error is re-raised, so the caller's Core_Fail still reports the real
' cause rather than a swallowed one.
Public Function Core_NewSheet(ByVal destName As String) As Worksheet
    ' errNum, not eNum: "eNum" lower-cases to the VBA reserved word "enum", which
    ' Excel refuses to compile and build.ps1's reserved-word guard rejects.
    Dim ws As Worksheet, errNum As Long, errDesc As String
    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

    On Error Resume Next
    ws.name = destName
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0

    If errNum <> 0 Then
        Application.DisplayAlerts = False
        On Error Resume Next
        ws.Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        Err.Raise errNum, "Core_NewSheet", _
            "Could not name the new sheet '" & destName & "': " & errDesc
        Exit Function
    End If
    Set Core_NewSheet = ws
End Function

' Any string reduced to something Excel will accept as a sheet name: forbidden
' characters replaced with "-", trimmed, cut to 31. It does NOT make the name
' unique - a clash is Core_DestUsable's job, and is refused rather than resolved.
Public Function Core_SafeSheetName(ByVal raw As String) As String
    Dim nm As String, bad As String, i As Long
    nm = raw
    bad = Core_BadSheetChars()
    For i = 1 To Len(bad)
        nm = Replace(nm, Mid$(bad, i, 1), "-")
    Next i
    nm = Trim(nm)
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    Core_SafeSheetName = nm
End Function

' Blank means the folder this workbook is saved in - contract section 9's "data
' location", and the answer in the overwhelming majority of cases. Returns a path
' ending in a separator, or "" when the workbook has never been saved and no folder
' was configured (the caller reports that case by name).
Public Function Core_ResolveFolder(ByVal raw As String) As String
    Dim p As String
    p = Trim(raw)
    If Len(p) = 0 Then p = ThisWorkbook.Path
    If Len(p) = 0 Then Exit Function                 ' workbook never saved
    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    Core_ResolveFolder = p
End Function

' Folder existence, via FileSystemObject rather than Dir(path, vbDirectory).
' Dir's behaviour with a trailing separator is not something to bet a refusal dialog
' on, and it also returns a non-empty result for a FILE, so a path pointing at a
' workbook would read as a valid folder. Late-bound: no reference to add, and no
' VBIDE trust dependency (CLAUDE.md rule 3).
Public Function Core_FolderExists(ByVal folderPath As String) As Boolean
    Dim fso As Object, p As String
    p = folderPath
    If Len(p) > 3 Then
        If Right$(p, 1) = Application.PathSeparator Then p = Left$(p, Len(p) - 1)
    End If
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then Core_FolderExists = fso.FolderExists(p)
    On Error GoTo 0
End Function

'---- export engine: freeze formulas (incl. spills) to values ----
' Flatten dynamic-array SPILLS and legacy CSE ARRAYS in rng to static values,
' then freeze every ordinary formula left. Two passes, in this order: a
' spill's anchor cell is itself one of the "formula" cells the second pass
' would try to freeze while it is still spilling, so flattening the spill
' first turns it into an ordinary value block the second pass can walk past
' safely.
'
' Extracted from importexport.export (2026-07-28 W3 plan) - NOT moved
' unchanged. Three deliberate differences from that shipped loop, called
' out explicitly because this is a shipped action used for a finance
' deliverable:
'
'   1. This function returns the number of formula cells it ACTUALLY
'      converted (counted area by area, in pass 2), and reports any area
'      whose conversion raised an error via the `failedCount` out-param
'      instead of silently swallowing it. The original loop had no caller
'      that needed a count; export.freeze_values does (Task 10), so this
'      is new surface, not a behavior-preserving copy of what already ran.
'
'   2. Pass 2 searches the exact `rng` the caller passed in, UNLESS
'      `reevalUsedRange` is True, in which case it re-reads
'      rng.Worksheet.UsedRange fresh before pass 2. The original always
'      called ws.UsedRange fresh before EACH of its two passes
'      (PkgImportExport.bas, pre-refactor :373 and :392) - two separate
'      property reads - so if pass 1 changed the sheet's used bounds, pass
'      2 saw the new bounds. A caller that hands this function a whole
'      sheet's used range (importexport.export's own refactor below, and
'      export.per_sheet's xlsx path, Task 12) passes
'      reevalUsedRange:=True to keep that exact original behavior.
'      export.freeze_values (Task 10) hands this function a user-PICKED,
'      already-confirmed range and must NOT have pass 2 silently widen
'      that to the sheet's current UsedRange - it leaves reevalUsedRange
'      at its default, False.
'
'   3. The error handling around the one risky statement (`ar.Value =
'      ar.Value`) is `On Error Resume Next` / `On Error GoTo 0`, scoped to
'      THIS function - it is not the calling Sub's `On Error GoTo Fail`.
'      This is safe and behaviorally equivalent for every caller: VBA
'      error handling is scoped per procedure, so a callee's own `On
'      Error` statements never alter or disable a caller's error handler.
'      (This is a different situation from the bug fixed in
'      PkgExport_FreezeValues and PkgFormat_PrintLayout below, where a Sub
'      disabled its OWN `Fail:` routing - this function never had a
'      `Fail:` label to disable.)
'
' Everything else - the two-pass order, the spill/array flatten logic
' itself - is copied verbatim.
'
' PRECONDITION: Application.Calculation must be xlCalculationAutomatic when
' this runs - it reads live formula results via .Value. This was written and
' tested that way inside importexport.export; Core_SpeedOn sets Manual, so a
' caller inside that wrapper must restore Automatic first, exactly as
' importexport.export always has. This function does not manage Calculation
' itself, so two callers running back to back never fight over restoring it.
'
' reevalUsedRange: when True, pass 2 searches the CURRENT
' rng.Worksheet.UsedRange instead of the exact `rng` passed in - see
' difference #2 above. Only the whole-sheet callers pass True.
'
' failedCount: ByRef out-parameter. On return, the number of formula cells
' whose conversion to a static value raised an error and were left as live
' formulas. A caller that reports a completion count MUST check this and
' say so when it is nonzero - see export.freeze_values (Task 10).
Public Function Core_FreezeValues(rng As Range, _
                                  Optional ByVal reevalUsedRange As Boolean = False, _
                                  Optional ByRef failedCount As Long = 0) As Long
    Dim fRange As Range, ar As Range, cel As Range, sp As Range, sv As Variant
    Dim target As Range, converted As Long, failed As Long
    If rng Is Nothing Then Exit Function
    Set target = rng

    On Error Resume Next
    Set fRange = target.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not fRange Is Nothing Then
        For Each cel In fRange.Cells
            If cel.HasSpill Then
                Set sp = cel.SpillingToRange
                sv = sp.Value
                cel.ClearContents
                sp.Value = sv
            ElseIf cel.HasArray Then
                Set sp = cel.CurrentArray
                sv = sp.Value
                sp.ClearContents
                sp.Value = sv
            End If
        Next cel
    End If

    If reevalUsedRange Then
        On Error Resume Next
        Set target = rng.Worksheet.UsedRange
        On Error GoTo 0
        If target Is Nothing Then Set target = rng
    End If

    Set fRange = Nothing
    On Error Resume Next
    Set fRange = target.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not fRange Is Nothing Then
        For Each ar In fRange.Areas
            On Error Resume Next
            Err.Clear
            ar.Value = ar.Value
            If Err.Number = 0 Then
                converted = converted + ar.Cells.Count
            Else
                failed = failed + ar.Cells.Count
                Err.Clear
            End If
            On Error GoTo 0
        Next ar
    End If

    failedCount = failed
    Core_FreezeValues = converted
End Function

' Any string reduced to something Windows will accept as a FILE NAME - not a
' sheet name. Core_SafeSheetName's forbidden-character list is narrower than
' the filesystem's: Windows additionally forbids  " < > |  on top of Excel's
' own  : \ / ? * [ ]  - so a sheet name that is perfectly legal in Excel (one
' containing a quote, say) can still be an illegal file name. That is why
' this does not simply call Core_SafeSheetName and stop.
'
' Added in the W3 plan (2026-07-28) alongside export.pdf and
' export.per_sheet, and flagged there for W5: backup.timestamped_copy will
' need the identical logic. Promote-on-second-use already happened here, by
' writing it once in Core rather than in export.pdf's own package.
'
' Does NOT make the name unique - a caller writing to disk still checks for an
' existing file itself (Core_DestUsable's sheet precedent: refuse rather than
' resolve a collision by guessing).
'
' maxLen caps the result; 0 means no cap. It is a parameter rather than a fixed
' number because the right cap depends on what the caller is building. A whole
' path is bounded by MAX_PATH, but a caller assembling
' "<folder><date> <label> <timestamp>.xlsx" needs to bound its own middle
' segment, and only it knows how much room the rest takes. W5's
' backup.snapshot passes 40 for exactly this reason.
Public Function Core_SafeFileName(ByVal raw As String, _
                                  Optional ByVal maxLen As Long = 0) As String
    Dim nm As String, bad As String, i As Long
    nm = Trim(raw)
    bad = Core_BadSheetChars() & Chr(34) & "<>|"
    For i = 1 To Len(bad)
        nm = Replace(nm, Mid$(bad, i, 1), "-")
    Next i
    If maxLen > 0 Then
        If Len(nm) > maxLen Then nm = Left$(nm, maxLen)
    End If
    Core_SafeFileName = Trim(nm)
End Function

' Coerce a text value from a source export to a target type.
'   T=text  N=number(strips $ , %)  D/DT=date  B=boolean  G=general(auto)
Public Function Core_Coerce(ByVal v As Variant, ByVal code As String) As Variant
    Dim s As String
    If IsEmpty(v) Then Core_Coerce = "": Exit Function
    If VarType(v) = vbString Then
        If v = "" Then Core_Coerce = "": Exit Function
    End If

    Select Case UCase(code)
        Case "T"
            Core_Coerce = CStr(v)
        Case "N"
            If IsNumeric(v) Then
                Core_Coerce = CDbl(v)
            Else
                s = Replace(Replace(Replace(Replace(CStr(v), "$", ""), ",", ""), " ", ""), "%", "")
                If IsNumeric(s) Then Core_Coerce = CDbl(s) Else Core_Coerce = v
            End If
        Case "D", "DT"
            If IsDate(v) Then Core_Coerce = CDate(v) Else Core_Coerce = v
        Case "B"
            Core_Coerce = (UCase(Trim(CStr(v))) = "TRUE" Or Trim(CStr(v)) = "1")
        Case Else                     ' "G" general / auto
            s = Replace(Replace(CStr(v), ",", ""), "$", "")
            If IsNumeric(s) Then
                Core_Coerce = CDbl(s)
            ElseIf IsDate(v) Then
                Core_Coerce = CDate(v)
            Else
                Core_Coerce = v
            End If
    End Select
End Function

'================================================================
' 4. REGISTRY + DISPATCHER
'    Zero-trust discovery: probe a maintained candidate list by name.
'    Add one package (its PascalCase module suffix) per line below as
'    packages are built - e.g.  c.Add "ImportExport"  ->  probes
'    "PkgImportExport_Manifest".
'================================================================
Private Function CoreCandidates() As Collection
    Dim c As New Collection
    c.Add "ImportExport"            ' probes PkgImportExport_Manifest
    c.Add "Import"                  ' probes PkgImport_Manifest
    c.Add "Clean"                   ' probes PkgClean_Manifest
    c.Add "Validate"                ' probes PkgValidate_Manifest
    c.Add "Format"                  ' probes PkgFormat_Manifest
    c.Add "Export"                  ' probes PkgExport_Manifest
    c.Add "Summarize"               ' probes PkgSummarize_Manifest
    c.Add "Enrich"                  ' probes PkgEnrich_Manifest
    c.Add "Backup"                  ' probes PkgBackup_Manifest
    c.Add "Workbook"                 ' probes PkgWorkbook_Manifest
    Set CoreCandidates = c
End Function

' Run a package's manifest if that module is actually present.
Private Function TryManifest(ByVal pkgName As String, reg As Object) As Boolean
    On Error Resume Next
    Application.Run "Pkg" & pkgName & "_Manifest", reg
    TryManifest = (Err.Number = 0)
    On Error GoTo 0
End Function

' The installed + enabled packages, each as a fully-built reg. MEMOIZED - see
' mRegCache's declaration for why sharing one Collection is safe, and for the
' one thing that would make it unsafe.
Private Function BuildRegistry() As Collection
    Dim names As Collection, nm As Variant, reg As Object, out As Collection
    If mRegCacheOk Then
        Set BuildRegistry = mRegCache
        Exit Function
    End If
    ' Explicit New, not "Dim out As New Collection": an auto-instantiating local
    ' re-creates itself on the next reference after being set to Nothing, which
    ' is a surprise nobody needs sitting next to a cache.
    Set out = New Collection
    Set names = CoreCandidates()
    For Each nm In names
        Set reg = CreateObject("Scripting.Dictionary")
        If TryManifest(CStr(nm), reg) Then
            If reg.Exists("id") Then
                If PackageEnabled(CStr(reg("id"))) Then out.Add reg
            End If
        End If
    Next nm
    Set mRegCache = out
    mRegCacheOk = True
    Set BuildRegistry = out
End Function

' Drop the memoized registry. Called at the top of every entry point that BEGINS
' a build - Core_BuildSetup, Core_FinalizeSetup, Core_BuildControlPanel - and
' deliberately NOT from Core_Run.
'
' The registry depends on the installed modules (fixed for the session) and on
' the optional _Packages sheet's enable/disable values, which a user can only
' have edited BETWEEN builds. Those three entry points are every place a build
' begins, so an edited _Packages is always picked up. Invalidating in Core_Run
' would rebuild the registry once per chain STEP, which is most of what this
' cache exists to stop.
Private Sub InvalidateRegistry()
    Set mRegCache = Nothing
    mRegCacheOk = False
End Sub

' A package is active unless the optional _Packages sheet lists it as FALSE.
Private Function PackageEnabled(ByVal id As String) As Boolean
    Dim ws As Worksheet, r As Long, last As Long
    PackageEnabled = True
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PACKAGES_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To last
        If LCase(Trim(CStr(ws.Cells(r, 1).Value))) = LCase(id) Then
            PackageEnabled = IsTrue(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
End Function

Private Function IsChainId(ByVal id As String) As Boolean
    IsChainId = (LCase(Left$(id, 6)) = "chain.")
End Function

' The SECOND reserved action-id namespace (contract section 7, v2.1.0), enforced at
' Finalize exactly like chain.* - see ComposeReport (a).
'
' Not cosmetic symmetry. AVAIL_LEGACY_KEY ("core.legacy") lives INSIDE the
' availability dictionary alongside real action ids, so a package declaring an
' action with that exact id occupies the sentinel key. AvailFor would then find the
' legacy key present on every map it builds and resolve EVERY absent id to Button
' instead of Off - silently undoing the v2.1.0 default on that deployment, for
' destructive actions included. Section 7 states no character rule for action ids,
' so the sentinel cannot be asserted collision-proof on its own (ADR 0003 decision
' 2); the namespace has to be enforced rather than merely declared.
'
' Left$ with a length past the end of the string returns the whole string rather
' than raising, so a 1- or 2-character id is safe here - same as IsChainId above.
Private Function IsCoreId(ByVal id As String) As Boolean
    IsCoreId = (LCase(Left$(id, 5)) = "core.")
End Function

Private Function IsQuiet(ctx As Object) As Boolean
    On Error Resume Next
    IsQuiet = (ctx.Exists("quiet"))
    If IsQuiet Then IsQuiet = (ctx("quiet") = True)
End Function

' Completion sink. Handlers call this instead of MsgBox for their final summary.
' When the context is quiet (a chain step), the message is captured on the ctx
' and the dialog is suppressed; otherwise it shows exactly as a direct MsgBox.
Public Sub Core_Done(ctx As Object, ByVal msg As String, Optional ByVal title As String = "Done")
    On Error Resume Next
    ctx("doneMsg") = msg
    ctx("done") = True
    On Error GoTo 0
    If Not IsQuiet(ctx) Then MsgBox msg, vbInformation, title
End Sub

' Refusal sink. An action calls this INSTEAD of MsgBox when a precondition is not
' met - the destination already exists, the folder is missing, nothing is
' configured yet - and then exits having changed nothing.
'
' It records the outcome AND shows the dialog, in one call, deliberately. The
' pattern it replaces was two calls: set a flag, then MsgBox. An action that did
' one and forgot the other reported a success it never achieved, and a chain
' printed a tick beside it. One call cannot be half-done.
'
' A refusal is NOT an error: failNumber stays 0, and Core_RunChain reads the
' separate "refused" flag to decide it must stop rather than ask. See contract
' section 6.
Public Sub Core_Refuse(ctx As Object, ByVal msg As String, _
                       Optional ByVal title As String = "Cannot run this yet")
    On Error Resume Next
    ctx("failed") = True          ' not-ok, so RunAction reports it as such
    ctx("refused") = True         ' ...but a refusal, so the chain stops instead of asking
    ctx("failNumber") = 0         ' explicit: there is no error number to show
    ctx("failMsg") = msg
    On Error GoTo 0

    Core_Log ctx, "REFUSED: " & msg
    If Not IsQuiet(ctx) Then MsgBox msg, vbExclamation, title
End Sub

' Run one action under an isolated child context. Returns a result dictionary;
' never shows its own error/success dialog (callers decide). quiet=True suppresses
' the handler's Core_Done completion dialog and captures its message instead.
Private Function RunAction(ByVal actionId As String, ByVal quiet As Boolean) As Object
    Dim res As Object, reg As Object, ctx As Object, entry As String
    Set res = CreateObject("Scripting.Dictionary")
    res("id") = actionId
    res("label") = actionId
    res("ok") = False
    res("found") = False
    res("message") = ""
    res("errNumber") = 0
    res("errText") = ""
    res("reported") = False
    ' Core_RunChain reads res("refused") unconditionally, and reading a missing key
    ' from a Scripting.Dictionary raises. This procedure has three early returns -
    ' action-not-installed, StepFail, and the normal path - so the key has to exist
    ' before any of them can be taken. Not cosmetic placement.
    res("refused") = False

    On Error GoTo StepFail

    ' Defence in depth: Finalize blocks a chain that references an Off action, so
    ' this should be unreachable through a validated config. It exists because
    ' "switched off but it ran anyway" is the one failure mode of this feature
    ' that a finance user cannot recover from. Placed under the error trap (not
    ' above it): ReadAvailMap does real worksheet I/O (FindKeyRow/NextKeyRow walk
    ' the sheet with .End(xlUp) and read cells), and a protected or corrupt
    ' _Config must land on StepFail like any other failure here, not escape the
    ' procedure. Resolves the owner early (instead of at the un-gated lookup
    ' below) purely so the "Off" result carries the real caption, not the raw id -
    ' Core_RunChain's Continue/Stop prompt is exactly the dialog this gate exists
    ' to protect, and it must not show a finance user a dotted id.
    If Not IsRunnable(ReadAvailMap(CfgSheet()), actionId) Then
        Set reg = FindActionOwner(actionId, entry)
        res("found") = True
        If Not reg Is Nothing Then res("label") = ActionCaption(reg, actionId)
        res("ok") = False
        res("errNumber") = 0
        res("errText") = "Action '" & actionId & "' is switched Off on the Setup tab."
        On Error GoTo 0
        Set RunAction = res
        Exit Function
    End If

    Set reg = FindActionOwner(actionId, entry)
    If reg Is Nothing Then
        On Error GoTo 0
        Set RunAction = res            ' found stays False
        Exit Function
    End If
    res("found") = True
    res("label") = ActionCaption(reg, actionId)

    Set ctx = NewCtx(actionId, reg)
    ctx("quiet") = quiet
    ctx("done") = False
    ctx("doneMsg") = ""

    Application.Run entry, ctx
    On Error GoTo 0
    If CtxFlag(ctx, "failed") Then         ' handler swallowed its own error
        res("ok") = False
        res("refused") = CtxFlag(ctx, "refused")
        res("errNumber") = CtxNum(ctx, "failNumber")
        res("errText") = CtxStr(ctx, "failMsg")
        res("reported") = Not quiet        ' standalone: the sink already showed it
    Else
        res("ok") = True
        res("message") = CtxStr(ctx, "doneMsg")
    End If
    Set RunAction = res
    Exit Function

StepFail:
    RestoreSpeed                       ' fail-safe: never leave Excel frozen
    res("ok") = False
    res("found") = True                ' a real error, not "action not installed"
    res("errNumber") = Err.Number
    res("errText") = Err.Description
    Set RunAction = res
End Function

' Caption for an action id from its owning manifest (for result labels/summaries).
Private Function ActionCaption(reg As Object, ByVal actionId As String) As String
    Dim a As Object
    ActionCaption = actionId
    For Each a In reg("actions")
        If LCase(CStr(a("id"))) = LCase(actionId) Then
            ActionCaption = CStr(a("caption")): Exit Function
        End If
    Next a
End Function

' Safe ctx accessors - avoid Scripting.Dictionary missing-key errors.
Private Function CtxFlag(ctx As Object, ByVal k As String) As Boolean
    On Error Resume Next
    If ctx.Exists(k) Then CtxFlag = (ctx(k) = True)
End Function

Private Function CtxStr(ctx As Object, ByVal k As String) As String
    On Error Resume Next
    If ctx.Exists(k) Then CtxStr = CStr(ctx(k))
End Function

Private Function CtxNum(ctx As Object, ByVal k As String) As Long
    On Error Resume Next
    If ctx.Exists(k) Then CtxNum = CLng(ctx(k))
End Function

' The single public dispatch entry. Resolves an action id or a reserved chain.* id.
' A DIRECT invocation of anything not available as a button is refused here: with
' the button gone the only ways to arrive are a stale shape, the VBE's Immediate
' window, or Application.Run - and none should be able to run something the user
' switched off. (NOT Alt+F8: Excel's Macro dialog lists only Public Subs taking no
' arguments, so this one never appears there.) Chain steps do not pass
' through here - Core_RunChain calls RunAction, which independently refuses Off.
Public Sub Core_Run(ByVal actionId As String)
    Dim res As Object, availMap As Object, t As String
    If IsChainId(actionId) Then Core_RunChain actionId: Exit Sub

    ' Reading the availability map is a worksheet read (FindKeyRow/NextKeyRow walk
    ' the sheet). Core_Run is a button's OnAction with no other error handler in
    ' the call chain above it, so an unexpected failure here (a protected or
    ' corrupt _Config) must fail CLOSED - refuse to run and say why - rather than
    ' surface a raw VBA runtime-error dialog and break into the IDE in front of a
    ' finance user.
    On Error GoTo AvailReadFail
    Set availMap = ReadAvailMap(CfgSheet())
    t = AvailFor(availMap, actionId)
    On Error GoTo 0

    If t <> AVAIL_BUTTON Then
        If t = AVAIL_CHAIN Then
            MsgBox "'" & actionId & "' is set to 'Chains only' on the Setup tab, so it " & _
                   "runs only as a step inside a Chain." & vbCrLf & vbCrLf & _
                   "Change its Availability to 'Button' to run it on its own.", _
                   vbInformation, "Not available on its own"
        Else
            MsgBox "'" & actionId & "' is switched Off on the Setup tab." & vbCrLf & vbCrLf & _
                   "Change its Availability to run it.", vbExclamation, "Action is off"
        End If
        Exit Sub
    End If
    Set res = RunAction(actionId, False)
    If Not res("found") And res("errNumber") = 0 Then
        MsgBox "Action not installed: " & actionId, vbExclamation, "Core"
    ElseIf Not res("ok") And Not res("reported") Then
        ' FailCode, not a second copy of "#" & errNumber. This dialog is only
        ' reachable today when errNumber is non-zero, but it used to hardcode the
        ' "#" the same way Core_RunChain's summary did - and that copy is exactly
        ' where the "X #0 <reason>" line came from. Reimplementing the format at a
        ' second call site is how the defect got in; sharing the helper is how it
        ' stays out when a future caller reaches here with a refusal.
        MsgBox "Action '" & actionId & "' failed:" & vbCrLf & vbCrLf & _
               Trim$(FailCode(res)) & "  " & res("errText"), vbExclamation, "Macro Error"
    End If
    Exit Sub

AvailReadFail:
    MsgBox "Could not confirm '" & actionId & "' is switched on, so it was not run." & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Rebuild the Setup tab if the problem continues.", vbExclamation, "Core"
End Sub

' Run a chain: each step quiet; prompt Continue/Stop on failure; one end summary.
' No outer Core_SpeedOn/Off here - each step self-manages its own speed via its
' handler; wrapping the whole run would corrupt the depth-counted speed state
' when a step fails (RunAction's StepFail path hard-restores via RestoreSpeed).
Private Sub Core_RunChain(ByVal chainId As String)
    Dim chains As Collection, ch As Object, target As Object
    Dim stepId As Variant, res As Object, n As Long, ok As Long, summary As String
    Dim stopped As Boolean, total As Long, refusedAt As Boolean
    On Error GoTo RunnerFail

    Set chains = ReadChains(CfgSheet())
    For Each ch In chains
        If LCase(CStr(ch("id"))) = LCase(chainId) Then Set target = ch: Exit For
    Next ch
    If target Is Nothing Then
        MsgBox "Chain not found: " & chainId, vbExclamation, "Core": Exit Sub
    End If
    total = target("steps").Count

    For Each stepId In target("steps")
        n = n + 1
        Set res = RunAction(CStr(stepId), True)      ' quiet; each step owns its speed
        If Not res("found") Then
            summary = summary & n & ". " & stepId & "  X not installed" & vbCrLf
            If Not AskContinue(n, CStr(stepId), "action not installed") Then stopped = True: Exit For
        ElseIf res("ok") Then
            ok = ok + 1
            summary = summary & n & ". " & res("label") & "  " & ChrW$(&H2713) & _
                      IIf(Len(res("message")) > 0, "  (" & FirstLine(CStr(res("message"))) & ")", "") & vbCrLf
        Else
            summary = summary & n & ". " & res("label") & "  X" & FailCode(res) & _
                      " " & res("errText") & vbCrLf
            If res("refused") Then
                ' A refusal means this step produced NOTHING. Continuing would run every
                ' later step against whatever was already there - which is exactly how a
                ' daily chain cleans last period's data and reports success. Contract
                ' section 6: a chain stops at a refusal, and only asks on a failure.
                stopped = True
                refusedAt = True
                Exit For
            End If
            If Not AskContinue(n, CStr(res("label")), CStr(res("errText"))) Then stopped = True: Exit For
        End If
    Next stepId

    Dim head As String
    head = target("label") & IIf(stopped, "  - stopped at step " & n, "  - done") & _
           "  (" & ok & "/" & total & " ok)"
    If refusedAt Then
        head = head & vbCrLf & "Step " & n & " could not run, so the remaining " & _
               (total - n) & " step(s) were skipped."
    End If
    MsgBox head & vbCrLf & vbCrLf & summary, _
           IIf(stopped, vbExclamation, vbInformation), "Chain"
    Exit Sub
RunnerFail:
    RestoreSpeed
    MsgBox "Chain '" & chainId & "' error at step " & n & ":" & vbCrLf & vbCrLf & _
           "#" & Err.Number & "  " & Err.Description, vbExclamation, "Chain"
End Sub

' "Step N <label> failed: <reason>. Continue with the rest?"  Yes=continue.
Private Function AskContinue(ByVal n As Long, ByVal label As String, ByVal reason As String) As Boolean
    AskContinue = (MsgBox("Step " & n & "  '" & label & "'  failed:" & vbCrLf & vbCrLf & _
        reason & vbCrLf & vbCrLf & "Continue with the remaining steps?", _
        vbExclamation + vbYesNo, "Chain - step failed") = vbYes)
End Function

' "  #1004" for a real error, "" for a refusal. A refusal has no error number, and
' a summary line reading "X #0 A sheet named 'Master' already exists" tells the user
' their macro malfunctioned when it did exactly what it should.
Private Function FailCode(res As Object) As String
    On Error Resume Next
    If CLng(res("errNumber")) <> 0 Then FailCode = "  #" & res("errNumber")
End Function

Private Function FirstLine(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, vbCr): If p = 0 Then p = InStr(s, vbLf)
    If p > 0 Then FirstLine = Left$(s, p - 1) Else FirstLine = s   ' NOT IIf: eager Left$(s,-1) would raise
End Function

Private Function FindActionOwner(ByVal actionId As String, ByRef entry As String) As Object
    Dim reg As Object, a As Object
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If LCase(CStr(a("id"))) = LCase(actionId) Then
                entry = CStr(a("entry"))
                Set FindActionOwner = reg
                Exit Function
            End If
        Next a
    Next reg
End Function

Private Function NewCtx(ByVal actionId As String, reg As Object) As Object
    Dim ctx As Object
    Set ctx = CreateObject("Scripting.Dictionary")
    ctx.Add "actionId", actionId
    ctx.Add "reg", reg
    ctx.Add "log", ""
    Set NewCtx = ctx
End Function

'================================================================
' 5. SETUP RENDERER  (builds the Setup sheet FROM the manifests)
'================================================================

' Clears the Action catalog block's cells in _Config so a rebuild's prior-value
' reads (ReadCodeMap/ReadAvailMap) find nothing for the cleared rows, instead
' of carrying an unreadable value forward unchanged. idsToClear Is Nothing
' clears the WHOLE block (marker, every id, every Availability/Placement
' cell) - nothing in it can be trusted. A Dictionary clears only the
' Placement cell of a row whose lower-cased id is a key in it; every other
' row is untouched - used when exactly which values are bad is known.
Private Sub ClearCatalogBlock(cfgWs As Worksheet, idsToClear As Object)
    Dim cKr As Long, cNk As Long, ci As Long, cid As String
    If cfgWs Is Nothing Then Exit Sub
    cKr = FindKeyRow(cfgWs, LAYOUT_KEY)
    If cKr = 0 Then Exit Sub
    cNk = NextKeyRow(cfgWs, cKr)
    If idsToClear Is Nothing Then
        ' ROW BY ROW, skipping merged rows - NOT a blanket Range.ClearContents over
        ' the whole span. TintRow merges columns B:E, and clearing D:E is a PARTIAL
        ' clear of that merge, which raises 1004 "We can't do that to a merged cell"
        ' and aborts the rebuild.
        '
        ' The span is layout-key-row .. next-key-row - 1, so what falls inside it
        ' depends on which block happens to follow the catalog. v0.9.0 put Chains
        ' immediately after it, and RenderChainsBlock opens with BlockIntro -> two
        ' merged TintRows - so from that release this fired every time a damaged
        ' config was cleared. It was reachable before that too, whenever the off band
        ' rendered (its banner is also a TintRow); with every action switched on there
        ' is no off band, only a blank spacer before END_KEY, which is why it went
        ' unseen. Found by the v0.7.0 manual gate, item 2.1.
        '
        ' Do not "fix" this by reordering the blocks again or by widening the clear to
        ' B:E: a merged row carries no id and no availability, so it is never data
        ' this needs to clear. Skipping it is correct, not a workaround.
        For ci = cKr To cNk - 1
            If Not cfgWs.Cells(ci, 2).MergeCells Then
                cfgWs.Cells(ci, RID_COL).ClearContents
                cfgWs.Cells(ci, VAL_COL + 1).ClearContents
                cfgWs.Cells(ci, VAL_COL + 2).ClearContents
            End If
        Next ci
    Else
        For ci = cKr + 1 To cNk - 1
            cid = LCase(Trim(CStr(cfgWs.Cells(ci, RID_COL).Value)))
            If Len(cid) > 0 Then
                If idsToClear.Exists(cid) Then cfgWs.Cells(ci, VAL_COL + 1).ClearContents
            End If
        Next ci
    End If
End Sub

' The one guarded-rebuild confirmation, shared by "invalid" and "damaged" -
' they used to differ only in clear scope and wording, and that duplication is
' exactly where divergent disclosures crept in. The Off-gating disclosure is
' UNCONDITIONAL, not just for a full clear: a targeted clear still touches
' VAL_COL + 1, which IS the Availability column in a genuine v1.2+ block, so a
' doubly-tampered store (marker missing AND the header renamed, landing on
' "invalid" instead of "damaged") could still have live Off state in the very
' cells being cleared - "invalid never touches Availability" was a
' prediction, not a fact, so the disclosure no longer depends on it.
' Returns True if the user confirmed (clear, and any Setup-tab delete, already
' done); False if they declined (nothing touched). On True with the tab
' having existed, THIS function deletes it, so Core_BuildSetup's own "tab
' already exists, rebuild it?" prompt never fires twice over one choice.
Private Function ConfirmAndClear(ByVal problem As String, ByVal discarded As String, _
                                  idsToClear As Object) As Boolean
    Dim setupPresent As Boolean, msg As String
    setupPresent = Core_SheetExists(SETUP_SHEET)
    msg = problem & vbCrLf & vbCrLf & _
        "The only way forward is to rebuild the Action catalog." & vbCrLf & vbCrLf & _
        "DISCARDED: " & discarded & vbCrLf & _
        "This takes effect the moment you click Yes, BEFORE the rebuild finishes: any " & _
        "action you had switched Off will be able to run again immediately, and will stay " & _
        "that way until you switch it Off again and click Finalize Setup Tab." & vbCrLf & _
        "KEPT: every other saved setting in this workbook."
    If setupPresent Then
        msg = msg & vbCrLf & vbCrLf & _
            "The current Setup tab will also be rebuilt, since its saved answers cannot be " & _
            "verified either."
    End If
    msg = msg & vbCrLf & vbCrLf & "Rebuild the Setup tab now?"
    If MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2, _
              "Create Setup Tab - saved configuration unreadable") = vbNo Then
        MsgBox "Nothing has been changed.", vbInformation, "Create Setup Tab"
        ConfirmAndClear = False
        Exit Function
    End If
    ClearCatalogBlock CfgSheet(), idsToClear
    If setupPresent Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(SETUP_SHEET).Delete
        Application.DisplayAlerts = True
    End If
    ConfirmAndClear = True
End Function

' The fail-closed config migration gate (contract section 5), extracted whole
' out of Core_BuildSetup - a renderer has no business owning the recovery
' policy. True = safe to proceed (no problem, or the user just confirmed a
' guarded rebuild - clear and any Setup-tab delete already done). False = the
' user has already been told why nothing can proceed, or declined; workbook
' untouched either way. schemaState's meaning is ConfigSchemaStatus's to
' document; the one thing worth repeating here is WHY "invalid" and "damaged"
' get different treatment even though both are unreadable: "invalid"'s bad
' value is a VISIBLE, user-editable Placement cell that ComposeReport does not
' block on, so fixing it on the Setup tab genuinely works when the tab exists.
' "damaged"'s bad value is the HIDDEN marker cell itself, which the user
' cannot see or edit, and which - since _Config is a straight ws.Copy of the
' Setup tab - the live Setup tab almost always carries too, tripping
' ComposeReport's own check. "Fix it on the Setup tab" is never a real
' instruction for "damaged", so it always gets the guarded rebuild instead,
' tab present or not.
Private Function GateConfigMigration() As Boolean
    Dim schemaState As String, schemaDetail As String, schemaBadIds As Object
    schemaState = ConfigSchemaStatus(schemaDetail, schemaBadIds)

    Select Case schemaState
        Case "none", "ok", "legacy"
            GateConfigMigration = True

        Case "unknown"
            MsgBox schemaDetail & vbCrLf & vbCrLf & _
                   "To carry on, import that version's MacrosBundle.bas again. Your settings " & _
                   "are safe in the meantime - nothing has been changed.", _
                   vbExclamation, "Create Setup Tab"
            GateConfigMigration = False

        Case "invalid"
            If Core_SheetExists(SETUP_SHEET) Then
                MsgBox schemaDetail & vbCrLf & vbCrLf & _
                       "The Setup tab is still open with your other answers on it. Correct " & _
                       "the placement code(s) named above there, then click 'Finalize Setup Tab'. " & _
                       "Nothing has been changed." & vbCrLf & vbCrLf & _
                       "This is the faster fix. If you delete the Setup tab instead, 'Create " & _
                       "Setup Tab' can still recover: it will clear just the unreadable " & _
                       "code(s) and rebuild around them, keeping everything else - but " & _
                       "correcting the cell(s) here directly avoids that extra step.", _
                       vbExclamation, "Create Setup Tab"
                GateConfigMigration = False
            Else
                GateConfigMigration = ConfirmAndClear(schemaDetail, _
                    "the placement code(s) named above (those actions return to automatic " & _
                    "placement); every other saved placement is kept.", schemaBadIds)
            End If

        Case "damaged"
            GateConfigMigration = ConfirmAndClear(schemaDetail, _
                "every saved placement and Availability choice in the Action catalog (every " & _
                "action returns as a plain Button with automatic placement).", Nothing)
    End Select
End Function

Public Sub Core_BuildSetup()
    Dim ws As Worksheet, reg As Object, f As Object, r As Long, nPkgs As Long
    Dim firstContentRow As Long, availMap As Object, bandFirst As Long, bandLast As Long
    On Error GoTo Fail
    InvalidateRegistry              ' a build begins here - see InvalidateRegistry

    If Not GateConfigMigration() Then Exit Sub

    If Core_SheetExists(SETUP_SHEET) Then
        ' Fix wave: "never discard un-finalized typing" is a design principle
        ' (ConfirmAndClear already honours it with an explicit DISCARDED/KEPT
        ' breakdown and vbDefaultButton2). This prompt deletes the sheet on Yes,
        ' so anything typed since the last Finalize is lost - say so plainly and
        ' default to No instead of Yes.
        If MsgBox("A '" & SETUP_SHEET & "' tab already exists. Rebuild it?" & vbCrLf & vbCrLf & _
                  "DISCARDED: anything you have typed on the Setup tab since the last 'Finalize " & _
                  "Setup Tab' - the rebuild re-seeds from the last ACTIVATED settings, not from " & _
                  "what is currently on screen." & vbCrLf & _
                  "KEPT: every setting from your last Finalize.", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Create Setup Tab") = vbNo Then Exit Sub
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(SETUP_SHEET).Delete
        Application.DisplayAlerts = True
    End If

    Application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
    ws.name = SETUP_SHEET
    ws.Tab.Color = ForgeRed()
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Columns("A").ColumnWidth = 9
    ws.Columns("B").ColumnWidth = 50
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 34
    ws.Columns("E").ColumnWidth = 30
    ws.Columns("F").ColumnWidth = 3
    ws.Columns("G").ColumnWidth = 26
    ws.Columns("H").ColumnWidth = 42

    r = 1
    ws.Cells(r, 2).Value = "Toolkit Setup"
    ws.Cells(r, 2).Font.Size = 16
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Color = ForgeNavy()
    r = r + 1
    ws.Cells(r, 2).Value = "Fill in the YELLOW cells, then click  'Finalize Setup Tab'  on the Setup Menu tab."
    ws.Cells(r, 2).Font.Italic = True
    r = r + 1
    ws.Cells(r, 2).Value = "Do not edit column A (machine keys). Blank rows in a list are ignored."
    ws.Cells(r, 2).Font.Italic = True
    ws.Cells(r, 2).Font.Color = RGB(150, 0, 0)
    r = r + 1
    ws.Cells(r, 7).Value = "Reference: tabs in THIS workbook"
    ws.Cells(r, 7).Font.Bold = True
    ws.Cells(r, 8).Value = "Reference: Excel files in THIS folder"
    ws.Cells(r, 8).Font.Bold = True
    PopulateReferenceLists ws, r + 1
    r = r + 2

    ' The availability map comes from the ACTIVATED config, so a rebuild reflects
    ' the last Finalize - NOT what is currently typed on the Setup tab. Editing an
    ' Availability cell and rebuilding without finalizing first is a no-op here, by
    ' design (design section 2.3): the rebuild re-seeds Setup from _Config.
    ' A fresh workbook has no _Config, so the map is empty and AvailFor reports Off
    ' for every id - the v2.1.0 default (ADR 0003).
    Set availMap = ReadAvailMap(CfgSheet())

    ' AutoFit starts HERE, above Compose, not at the first package field. Compose
    ' now renders first (below) and its Chain "Steps" cells are WrapText multi-line;
    ' anchoring AutoFit after it would clip them. Named firstContentRow rather than
    ' firstFieldRow because it is no longer the first field.
    firstContentRow = r

    ' Action catalog FIRST, then Chains, then the per-package settings bands. The
    ' catalog is the surface a first-time user has to act on before anything else
    ' works - as of v0.6.0 every action arrives Off, so a Setup tab that opens on
    ' someone else's export paths buries the one block that turns the tool on.
    ' Settings for actions you have not enabled yet are the least urgent thing on
    ' the sheet, which is also why the off band collapses at the bottom.
    r = RenderComposeSection(ws, r)
    r = r + 1

    Dim banner As String
    For Each reg In BuildRegistry()
        nPkgs = nPkgs + 1
        banner = UCase(CStr(reg("name")))
        ' A package with nothing to configure draws no band. Skipping the header is
        ' the same shape as HasAdvanced guarding the ADVANCED one just below, and
        ' the same judgement v0.2.1 made when it stopped rendering a section header
        ' identical to its banner: a heading over nothing reads as a mistake.
        ' PkgClean reaches this state now that both of its settings describe the
        ' range the user picks and are asked for at run time instead. The package
        ' is still counted in nPkgs and its actions still appear in the Action
        ' catalog - it simply has no configuration surface.
        If HasAnyTier(reg, availMap, True) Then
            r = SectionHeader(ws, r, banner)
            r = RenderTier(ws, r, reg, "simple", True, banner, availMap, True)   ' simple fields, with their sections
            If HasAdvanced(reg, availMap) Then
                r = SectionHeader(ws, r, "ADVANCED  -  most users can leave these as-is")
                r = RenderTier(ws, r, reg, "advanced", False, banner, availMap, True)  ' advanced, collected flat
            End If
            r = r + 1
        End If
    Next reg

    ' The off band stays LAST: settings whose actions are all Off are the least
    ' urgent thing on the sheet, and the action catalog that switches them on is now
    ' at the very top rather than immediately above them.
    r = RenderOffBand(ws, r, availMap, bandFirst, bandLast)

    ' Sentinel key in column A. The config readers bound a list/table block by
    ' scanning column A for the next non-blank cell; without a trailing key the
    ' LAST list/table on the sheet has no lower bound (End(xlUp) stops at its own
    ' key row, whose value rows are blank in column A) and its values are dropped.
    ' It stays last, after the off band, because the off band may contain the
    ' sheet's final list/table block.
    ws.Cells(r, KEY_COL).Value = END_KEY

    ' Long labels in column B are set to wrap (below); auto-size the field rows
    ' so a wrapped label is fully visible instead of being clipped by the value
    ' cell in column C.
    ws.Rows(firstContentRow & ":" & r).AutoFit

    ' Group AFTER AutoFit: AutoFit sizes rows and would otherwise fight the
    ' collapse, and a collapsed row's height is irrelevant anyway.
    CollapseRows ws, bandFirst, bandLast

    ws.Columns("A").Font.Color = RGB(190, 190, 190)
    ws.Columns("A").Hidden = True
    ws.Columns(RID_COL).Hidden = True
    ws.Activate
    ws.Range("B1").Select
    Application.ScreenUpdating = True

    If nPkgs = 0 Then
        MsgBox "'" & SETUP_SHEET & "' created, but no packages are installed yet." & vbCrLf & _
               "Import a Pkg* module and add it to Core's candidate list.", _
               vbInformation, "Create Setup Tab"
    Else
        MsgBox "'" & SETUP_SHEET & "' tab created and pre-filled." & vbCrLf & vbCrLf & _
               "Edit the yellow cells, then run  'Finalize Setup Tab'.", _
               vbInformation, "Create Setup Tab"
    End If
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Core_BuildSetup error: " & Err.Description, vbExclamation, "Create Setup Tab"
End Sub

' Return the action ids, in declaration order, for fields that will actually render
' under this section. Re-applying RenderTier's tier and FieldInUse tests prevents a
' header from naming an action whose fields are in another tier or the off band;
' trusting the section's declaration would let the scan drift from the lazy header.
Private Function SectionActionIds(reg As Object, ByVal sectionIndex As Long, _
                                  ByVal tier As String, availMap As Object, _
                                  ByVal wantInUse As Boolean) As Collection
    ' Explicit New, not "Dim result As New Collection" - the same call this file
    ' already argues for at BuildRegistry: an auto-instantiating local silently
    ' re-creates itself on the next reference after being set to Nothing.
    Dim result As Collection, seen As Object, f As Object
    Dim ids() As String, i As Long, j As Long, actionId As String

    Set result = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For i = sectionIndex + 1 To reg("fields").Count
        Set f = reg("fields")(i)
        If f("kind") = "section" Then Exit For
        If CStr(f("tier")) = tier Then
            If FieldInUse(f, availMap) = wantInUse Then
                If f.Exists("usedBy") Then
                    If Len(Trim(CStr(f("usedBy")))) > 0 Then
                        ids = Split(CStr(f("usedBy")), ",")
                        For j = LBound(ids) To UBound(ids)
                            actionId = Trim(ids(j))
                            If Len(actionId) > 0 Then
                                If Not seen.Exists(actionId) Then
                                    seen.Add actionId, True
                                    result.Add actionId
                                End If
                            End If
                        Next j
                    End If
                End If
            End If
        End If
    Next i
    Set SectionActionIds = result
End Function

' Turn an ordered id list into the complete caption prefix. An empty list returns
' an empty string so deployment-wide sections keep their original title unchanged.
Private Function SectionActionPrefix(actionIds As Collection) As String
    Dim actionId As Variant, prefix As String
    For Each actionId In actionIds
        If Len(prefix) > 0 Then prefix = prefix & ", "
        prefix = prefix & ActionCaptionFor(CStr(actionId))
    Next actionId
    If Len(prefix) > 0 Then SectionActionPrefix = prefix & "  -  "
End Function

' Render every field of the given tier whose in-use state matches wantInUse. When
' showSections is True, a section header is emitted just before its first matching
' field (so a section whose fields are all in the other tier is not drawn). When
' False (the advanced pass and the off-band pass), section headers are skipped and
' the fields render flat under the one band.
Private Function RenderTier(ws As Worksheet, ByVal r As Long, reg As Object, _
                            ByVal tier As String, ByVal showSections As Boolean, _
                            ByVal bannerText As String, availMap As Object, _
                            ByVal wantInUse As Boolean) As Long
    Dim f As Object, pending As String, havePending As Boolean
    Dim i As Long, pendingIndex As Long, actionIds As Collection
    For i = 1 To reg("fields").Count
        Set f = reg("fields")(i)
        If f("kind") = "section" Then
            ' Skip a section header that just repeats the package banner (a
            ' description identical to the header above it reads as a mistake).
            If UCase(Trim(CStr(f("label")))) <> bannerText Then
                pending = CStr(f("label"))
                pendingIndex = i
                havePending = True
            End If
        ElseIf CStr(f("tier")) = tier Then
            If FieldInUse(f, availMap) = wantInUse Then
                If showSections And havePending Then
                    Set actionIds = SectionActionIds(reg, pendingIndex, tier, _
                                                     availMap, wantInUse)
                    r = SectionHeader(ws, r, _
                                      SectionActionPrefix(actionIds) & pending)
                    havePending = False
                End If
                r = RenderField(ws, r, reg, f)
            End If
        End If
    Next i
    RenderTier = r
End Function

Private Function RenderField(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    Select Case f("kind")
        Case "list"
            RenderField = RenderList(ws, r, reg, f)
        Case "table"
            RenderField = RenderTable(ws, r, reg, f)
        Case Else
            RenderField = RenderScalar(ws, r, reg, f)
    End Select
End Function

' Is there an advanced field that belongs in the working area? An advanced field
' whose actions are all Off goes to the off band instead, so drawing the ADVANCED
' header for it would leave an empty band.
Private Function HasAdvanced(reg As Object, availMap As Object) As Boolean
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If CStr(f("tier")) = "advanced" Then
                If FieldInUse(f, availMap) Then
                    HasAdvanced = True
                    Exit Function
                End If
            End If
        End If
    Next f
End Function

' Does this package have ANY non-section field on the requested side of the
' in-use split? Used to skip both an empty ADVANCED band and an empty off band.
Private Function HasAnyTier(reg As Object, availMap As Object, ByVal wantInUse As Boolean) As Boolean
    Dim f As Object
    For Each f In reg("fields")
        If f("kind") <> "section" Then
            If FieldInUse(f, availMap) = wantInUse Then
                HasAnyTier = True
                Exit Function
            End If
        End If
    Next f
End Function

' Every field whose usedBy names only Off actions, from every package, collected
' under one band at the very bottom of the field area and collapsed with Excel's
' row outline.
'
' They are RENDERED, never omitted. DO NOT "improve" this to skip them:
' Core_BuildSetup deletes and re-seeds Setup from the FINALIZED _Config, so
' omit-unless-enabled costs toggle -> Finalize -> rebuild -> fill -> Finalize and
' discards anything typed but not yet finalized (design section 2.3, which
' deliberately overrides all three panelists of an earlier review round).
'
' Returns the next free row; reports the collapsible span through bandFirst /
' bandLast so the caller can group AFTER its AutoFit pass. bandFirst = 0 means
' nothing was relegated and no band was drawn.
Private Function RenderOffBand(ws As Worksheet, ByVal r As Long, availMap As Object, _
                               ByRef bandFirst As Long, ByRef bandLast As Long) As Long
    ' anyOff, NOT "any": Any is a VBA reserved word (it appears in Declare ... As
    ' Any), so "Dim any As Boolean" is a compile error the VBE reports against this
    ' procedure's declaration line. Neither build.ps1's structural checks nor the
    ' syntax parser can see it - it was found by the owner's Debug > Compile.
    Dim reg As Object, anyOff As Boolean
    bandFirst = 0
    bandLast = 0
    RenderOffBand = r
    For Each reg In BuildRegistry()
        If HasAnyTier(reg, availMap, False) Then anyOff = True
    Next reg
    If Not anyOff Then Exit Function           ' nothing off: draw no band at all

    r = SectionHeader(ws, r, "SETTINGS FOR ACTIONS YOU HAVEN'T TURNED ON")
    r = TintRow(ws, r, "These belong to actions set to 'Off' in the action catalog above. " & _
                "They are kept and still saved, so nothing you typed is lost - set the " & _
                "action back to 'Button' and they move up into the main list. Click the + " & _
                "in the left margin to open this section.", _
                9, False, True, ForgeSteel())
    bandFirst = r
    For Each reg In BuildRegistry()
        If HasAnyTier(reg, availMap, False) Then
            r = TintRow(ws, r, CStr(reg("name")), 10, True, False, ForgeSteel())
            r = RenderTier(ws, r, reg, "simple", False, UCase(CStr(reg("name"))), availMap, False)
            r = RenderTier(ws, r, reg, "advanced", False, UCase(CStr(reg("name"))), availMap, False)
        End If
    Next reg
    bandLast = r - 1
    RenderOffBand = r
End Function

' Group + collapse the off band, with the band's own header row left visible above
' it (SummaryRow = xlAbove). Degrades to a plain VISIBLE band if outlining is
' refused - the fields are rendered and readable by Finalize either way, because
' the config readers key off column A and never care whether a row is hidden.
Private Sub CollapseRows(ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    If firstRow < 1 Or lastRow < firstRow Then Exit Sub
    On Error Resume Next
    ws.Outline.SummaryRow = xlAbove
    ws.Rows(firstRow & ":" & lastRow).Group
    ws.Outline.ShowLevels RowLevels:=1
    On Error GoTo 0
End Sub

Private Function RenderScalar(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    ws.Cells(r, KEY_COL).Value = f("key")
    LabelCell ws.Cells(r, 2), CStr(f("label"))
    If LCase(CStr(f("kind"))) = "bool" Then
        ' Write the DISPLAY spelling, exactly as RenderPlacementBlock writes
        ' AvailLabel rather than the raw token - otherwise a seeded lowercase
        ' "yes" would sit in a cell whose dropdown offers "Yes", which reads as
        ' a value the validation does not recognise.
        ws.Cells(r, VAL_COL).Value = BoolLabel(ScalarValue(reg, CStr(f("key"))))
        Yellow ws.Cells(r, VAL_COL)
        BoolDropdown ws.Cells(r, VAL_COL)
    Else
        ws.Cells(r, VAL_COL).Value = ScalarValue(reg, CStr(f("key")))
        Yellow ws.Cells(r, VAL_COL)
    End If
    r = r + 1
    r = RenderHelp(ws, r, CStr(f("help")))
    RenderScalar = r
End Function

' A plain-language note under a field, grey + italic + wrapped. Its column A and
' column C stay blank, so it is invisible to the config block readers (which key
' off column A) and the list/table readers (which key off column C). Returns the
' next row (unchanged when there is no help text).
Private Function RenderHelp(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    If Trim(txt) = "" Then RenderHelp = r: Exit Function
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Italic = True
    ws.Cells(r, 2).Font.Size = 9
    ws.Cells(r, 2).Font.Color = RGB(120, 120, 120)
    ws.Cells(r, 2).WrapText = True
    ws.Cells(r, 2).VerticalAlignment = xlTop
    RenderHelp = r + 1
End Function

' Write a field label into column B, wrapped + top-aligned, so a long label
' stays fully visible next to the occupied value cell in column C.
Private Sub LabelCell(rng As Range, ByVal txt As String)
    rng.Value = txt
    rng.WrapText = True
    rng.VerticalAlignment = xlTop
End Sub

' A full-width tinted band row: columns B:E merged, text wrapped. Merged cells are
' ignored by AutoFit, so the row height is set from a line estimate (erring tall
' to avoid clipping). Column A and column C are inside the merge with no separate
' value, so the row stays invisible to the config block readers. Returns r + 1.
Private Function TintRow(ws As Worksheet, ByVal r As Long, ByVal txt As String, _
                         ByVal sz As Single, ByVal bold As Boolean, _
                         ByVal ital As Boolean, ByVal colr As Long) As Long
    Dim lines As Long
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Merge
        .Interior.Color = ForgeTint()
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    With ws.Cells(r, 2)
        .Value = txt
        .Font.Size = sz
        .Font.Bold = bold
        .Font.Italic = ital
        .Font.Color = colr
    End With
    lines = (Len(txt) \ 115) + 1                 ' ~115 chars/line across B:E; round up
    ws.Rows(r).RowHeight = lines * (sz + 4)
    TintRow = r + 1
End Function

' For list/table fields: draw the label and the explanation in the section-tinted
' band ABOVE the column-header row, so the headers sit directly on top of the
' input cells (no gap). The band rows are merged B:E; they are invisible to the
' config block readers. Returns the row for the header/key row.
Private Function BlockIntro(ws As Worksheet, ByVal r As Long, _
                            ByVal label As String, ByVal help As String) As Long
    r = TintRow(ws, r, label, 11, True, False, ForgeNavy())
    If Trim(help) <> "" Then r = TintRow(ws, r, help, 9, False, True, ForgeSteel())
    BlockIntro = r
End Function

' Attach a hover note to a header cell. Clears any existing note first: Setup is
' deleted and rebuilt on every Core_BuildSetup, but Setup Menu is REUSED, and
' AddComment raises on a cell that already carries one.
Private Sub HeaderNote(cell As Range, ByVal txt As String)
    On Error Resume Next
    cell.ClearComments
    On Error GoTo 0
    cell.AddComment txt
    cell.Comment.Shape.TextFrame.AutoSize = True
End Sub

Private Function RenderList(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    Dim vals As Collection, i As Long
    r = BlockIntro(ws, r, CStr(f("label")), CStr(f("help")))   ' label + help above the header
    ws.Cells(r, KEY_COL).Value = f("key")
    ws.Cells(r, VAL_COL).Value = f("headers")
    ws.Cells(r, VAL_COL).Font.Bold = True
    r = r + 1
    Set vals = ListValues(reg, CStr(f("key")))
    For i = 1 To MAX_LIST
        If i <= vals.Count Then ws.Cells(r, VAL_COL).Value = vals(i)
        Yellow ws.Cells(r, VAL_COL)
        r = r + 1
    Next i
    RenderList = r
End Function

Private Function RenderTable(ws As Worksheet, ByVal r As Long, reg As Object, f As Object) As Long
    Dim heads() As String, notes() As String, a() As String, b() As String, c() As String
    Dim i As Long, j As Long, nHead As Long, hasData As Boolean
    heads = Split(CStr(f("headers")), "|")
    notes = Split(CStr(f("notes")), "|")
    nHead = UBound(heads)
    If nHead > 2 Then nHead = 2                 ' 3 columns read (C, D, E)
    r = BlockIntro(ws, r, CStr(f("label")), CStr(f("help")))   ' label + help above the headers
    ws.Cells(r, KEY_COL).Value = f("key")
    For j = 0 To nHead
        ws.Cells(r, VAL_COL + j).Value = Trim(heads(j))
        ws.Cells(r, VAL_COL + j).Font.Bold = True
        If j <= UBound(notes) Then
            If Len(Trim(notes(j))) > 0 Then HeaderNote ws.Cells(r, VAL_COL + j), Trim(notes(j))
        End If
    Next j
    r = r + 1
    hasData = TableValues(reg, CStr(f("key")), a, b, c)
    For i = 1 To MAX_LIST
        If hasData Then
            If i <= UBound(a) Then
                ws.Cells(r, VAL_COL).Value = a(i)
                ws.Cells(r, VAL_COL + 1).Value = b(i)
                ws.Cells(r, VAL_COL + 2).Value = c(i)
            End If
        End If
        For j = 0 To nHead
            Yellow ws.Cells(r, VAL_COL + j)
        Next j
        r = r + 1
    Next i
    RenderTable = r
End Function

'================================================================
' 5b. CORE-OWNED SETUP SCHEMA  (chains + placement; NOT a package)
'   Rendered by Core_BuildSetup outside the package-discovery loop.
'   Constants CHAINS_KEY / LAYOUT_KEY / RID_COL are declared in the module
'   declarations section at the top (VBA requires module-level Const there;
'   a Const between procedures is a compile error).
'================================================================
' Action catalog FIRST, Chains second. The catalog is the only block a new
' deployment MUST touch - every action arrives Off as of v0.6.0, so nothing runs
' until someone sets an Availability cell. Chains build on actions that are already
' switched on, so asking for them first inverted the dependency: the reader was
' invited to compose buttons out of actions they had not enabled yet.
'
' Order here is presentation only. Every reader locates its block by key
' (FindKeyRow/NextKeyRow), so a _Config written in the old order still reads
' correctly and no migration is needed - the sheet is simply re-emitted in the new
' order at the next Finalize.
Private Function RenderComposeSection(ws As Worksheet, ByVal r As Long) As Long
    r = SectionHeader(ws, r, "COMPOSE YOUR BUTTONS")
    r = TintRow(ws, r, "Turn actions on in the Action catalog below - each one is " & _
                "Off until you pick Button or Chains only. Then build multi-step " & _
                "buttons (Chains) and arrange them onto pages (Placement). Leave " & _
                "every code blank for the automatic, category-packed layout.", _
                10, False, True, RGB(90, 90, 90))
    r = r + 1
    r = RenderPlacementBlock(ws, r)
    r = r + 1
    r = RenderChainsBlock(ws, r)
    RenderComposeSection = r
End Function

' "Label [id]" or a bare id per line, and bare ids may be comma-separated.
' Only the bracketed id is authoritative. A line carrying a bracket is never
' comma-split because a caption may contain a comma.
Private Function ParseStepTokens(ByVal cell As String) As Collection
    Dim out As New Collection, lines() As String, parts() As String
    Dim ln As String, i As Long, j As Long
    Dim lo As Long, hi As Long, id As String
    Set ParseStepTokens = out
    cell = Replace(cell, ";", vbLf)
    lines = Split(cell, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim(lines(i))
        If Len(ln) > 0 Then
            lo = InStrRev(ln, "["): hi = InStrRev(ln, "]")
            If lo > 0 And hi > lo Then
                id = Trim(Mid(ln, lo + 1, hi - lo - 1))
                If Len(id) > 0 Then out.Add id
            ElseIf lo = 0 And hi = 0 Then
                parts = Split(ln, ",")
                For j = LBound(parts) To UBound(parts)
                    id = Trim(parts(j))
                    If Len(id) > 0 Then out.Add id
                Next j
            Else
                id = ln            ' preserve malformed bracket-bearing input
                If Len(id) > 0 Then out.Add id
            End If
        End If
    Next i
End Function

' Caption for an action id (falls back to the id itself). Shared registry lookup
' factored out of ActionTokenFor so a caption-only consumer (StepsInline) doesn't
' have to strip the " [id]" suffix back off.
Private Function ActionCaptionFor(ByVal actionId As String) As String
    Dim reg As Object, a As Object
    ActionCaptionFor = actionId
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If LCase(CStr(a("id"))) = LCase(actionId) Then
                ActionCaptionFor = CStr(a("caption"))
                Exit Function
            End If
        Next a
    Next reg
End Function

' "Import Source Data [importexport.import]" for an action id (falls back to the id).
' Own registry lookup (not ActionCaptionFor + actionId) so the bracketed id is
' the registry's canonical a("id"), not whatever casing/spelling the caller
' passed in - a hand-typed non-canonical id in a Steps cell is normalised back
' to canonical when the Setup tab is rebuilt (fix round 2: a prior refactor
' bracketed the caller's raw actionId instead, losing that normalisation).
Private Function ActionTokenFor(ByVal actionId As String) As String
    Dim reg As Object, a As Object
    ActionTokenFor = actionId & " [" & actionId & "]"
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If LCase(CStr(a("id"))) = LCase(actionId) Then
                ActionTokenFor = CStr(a("caption")) & " [" & CStr(a("id")) & "]"
                Exit Function
            End If
        Next a
    Next reg
End Function

' The canonical-cased id for a hand-typed one, so a Steps cell reading
' "CLEAN.DUPES" is rewritten "clean.dupes" when Setup is rebuilt. This is the
' SAME normalisation ActionTokenFor performs above and it exists separately only
' because a Steps cell now holds bare ids rather than "Caption [id]" tokens. The
' note above records that a prior refactor already lost this normalisation once
' and it came back as fix round 2; writing StepsToTokens against the caller's raw
' string re-introduced exactly that defect in a new place. Unresolvable input is
' returned unchanged - Finalize's own report is what flags an unknown id.
Private Function CanonicalActionId(ByVal actionId As String) As String
    Dim reg As Object, a As Object
    CanonicalActionId = actionId
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If LCase(CStr(a("id"))) = LCase(actionId) Then
                CanonicalActionId = CStr(a("id"))
                Exit Function
            End If
        Next a
    Next reg
End Function

' "Caption [id]" for a HIDDEN or UNASSIGNED entry in ComposeReport, matching
' the format the OFF/CHAINS ONLY availability lists just above them already
' use ("caption [id]", built from the registry) - fix round 1, fix 5: the
' hidden list used to print the bare lowercased map key, e.g.
' "importexport.suggesttypes", while everything around it names the button
' too; fix round 2, fix 2 closed the same gap in the unassigned list right
' below it, so one report no longer mixes two formats. ActionTokenFor alone
' only knows registry actions; a hidden/unassigned CHAIN (chain.*) is not in
' the registry, so it is resolved against the chains collection ComposeReport
' already read, falling back to "id [id]" only if neither lookup finds it
' (should not happen for a live id, but matches ActionTokenFor's own fallback
' shape).
Private Function HiddenTokenFor(ByVal id As String, chains As Collection) As String
    Dim ch As Object
    If IsChainId(id) Then
        HiddenTokenFor = id & " [" & id & "]"
        For Each ch In chains
            If LCase(CStr(ch("id"))) = LCase(id) Then
                HiddenTokenFor = CStr(ch("label")) & " [" & CStr(ch("id")) & "]"
                Exit Function
            End If
        Next ch
    Else
        HiddenTokenFor = ActionTokenFor(id)
    End If
End Function

Private Function StepsToTokens(steps As Collection) As String
    Dim s As String, v As Variant
    For Each v In steps
        If Len(s) > 0 Then s = s & ", "
        s = s & CanonicalActionId(CStr(v))
    Next v
    StepsToTokens = s
End Function

Private Sub RenderActionTokens(ws As Worksheet, ByVal startRow As Long)
    Dim reg As Object, a As Object, gr As Long
    ws.Cells(startRow - 1, 7).Value = "Available actions - copy an id into Steps:"
    ws.Cells(startRow - 1, 7).Font.Bold = True
    gr = startRow
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            ws.Cells(gr, 7).Value = CStr(a("id"))
            ws.Cells(gr, 8).Value = CStr(a("caption"))
            ws.Range(ws.Cells(gr, 7), ws.Cells(gr, 8)).Font.Color = RGB(120, 120, 120)
            gr = gr + 1
        Next a
    Next reg
End Sub

' chain.<slug>, unique within `used` (a dictionary of lowercased ids already taken).
Private Function MintChainId(ByVal label As String, used As Object) As String
    Dim base As String, ch As String, i As Long, cand As String, n As Long
    For i = 1 To Len(label)
        ch = LCase(Mid(label, i, 1))
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            base = base & ch
        ElseIf Len(base) > 0 And Right$(base, 1) <> "_" Then
            base = base & "_"
        End If
    Next i
    base = Trim(base)                     ' loop above never forms "__"; strip any trailing "_" below
    Do While Len(base) > 0 And Right$(base, 1) = "_"
        base = Left$(base, Len(base) - 1)
    Loop
    If Len(base) = 0 Then base = "chain"
    cand = "chain." & base
    Do While used.Exists(LCase(cand))
        n = n + 1
        cand = "chain." & base & "_" & n
    Loop
    MintChainId = cand
End Function

Private Function RenderChainsBlock(ws As Worksheet, ByVal r As Long) As Long
    Dim existing As Collection, ch As Object, i As Long
    r = BlockIntro(ws, r, "Chains  -  one button that runs several actions in order", _
                   "Button name + Steps (paste action ids from the Actions list, separated by commas) + Placement (1A, hide, or TabName!B4).")
    ws.Cells(r, KEY_COL).Value = CHAINS_KEY
    ws.Cells(r, VAL_COL).Value = "Button name"
    ws.Cells(r, VAL_COL + 1).Value = "Steps (action ids, separated by commas)"
    ws.Cells(r, VAL_COL + 2).Value = PLACE_HEADER
    ws.Range(ws.Cells(r, VAL_COL), ws.Cells(r, VAL_COL + 2)).Font.Bold = True
    HeaderNote ws.Cells(r, VAL_COL + 1), _
               "Enter action ids separated by commas, copied from the list to the right. " & _
               "The older 'Caption [id]' form still works."
    HeaderNote ws.Cells(r, VAL_COL + 2), _
               "Placement forms: blank = automatic; 1A = page 1, column A; " & _
               "hide = leave it off the pages; TabName!B4 = put the button at cell B4 on your own tab named TabName."
    r = r + 1

    Set existing = ReadChains(CfgSheet())          ' re-seed from activated config
    For i = 1 To MAX_LIST
        If i <= existing.Count Then
            Set ch = existing(i)
            ws.Cells(r, VAL_COL).Value = ch("label")
            ws.Cells(r, VAL_COL + 1).Value = StepsToTokens(ch("steps"))   ' bare ids, comma-separated
            ws.Cells(r, VAL_COL + 2).Value = ch("code")
            ws.Cells(r, RID_COL).Value = ch("id")   ' preserve immutable id
        End If
        Yellow ws.Cells(r, VAL_COL)
        Yellow ws.Cells(r, VAL_COL + 1)
        Yellow ws.Cells(r, VAL_COL + 2)
        ws.Cells(r, VAL_COL + 1).WrapText = True
        r = r + 1
    Next i
    ' Copy-ready action ids and captions in G/H for authoring Steps.
    RenderActionTokens ws, r - MAX_LIST
    RenderChainsBlock = r
End Function

' Read the chains block from ws (CfgSheet at run time; SetupSheet for validation).
' Each item: {id, label, steps(Collection of ids), code}. IDs are read from the
' hidden RID column when present; a row with a label but no id has none yet.
Private Function ReadChains(ws As Worksheet) As Collection
    Dim out As New Collection, kr As Long, nk As Long, i As Long
    Dim label As String, ch As Object
    Set ReadChains = out
    If ws Is Nothing Then Exit Function
    kr = FindKeyRow(ws, CHAINS_KEY)
    If kr = 0 Then Exit Function
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1
        label = Trim(CStr(ws.Cells(i, VAL_COL).Value))
        If Len(label) > 0 Then
            Set ch = CreateObject("Scripting.Dictionary")
            ch("label") = label
            ch("id") = Trim(CStr(ws.Cells(i, RID_COL).Value))
            ch("code") = Trim(CStr(ws.Cells(i, VAL_COL + 2).Value))
            Set ch("steps") = ParseStepTokens(CStr(ws.Cells(i, VAL_COL + 1).Value))
            out.Add ch
        End If
    Next i
End Function

' Parse an Availability cell to one of the three tokens.
'   ""      -> ""            (v2.1.0: blank is NOT an answer; the CALLER must block)
'   invalid -> ""            (the CALLER must block Finalize)
' It must never fall back to Button on an unrecognised string: a user who typed
' "of" meaning "off" would otherwise be handed a live destructive button.
' Blank joined that rule in v2.1.0. It used to read as Button, which was correct
' while Button was the default; with the default at Off, silently reading a
' cleared cell as EITHER value is wrong - Button resurrects an action the user
' switched off, Off removes one they still want. Neither is recoverable from the
' result, so blank must block instead. ComposeReport already refuses to activate
' on "" and needs no change.
' Yes/No are accepted as synonyms so a user upgrading from a two-value habit, or
' one who types the obvious thing, is understood rather than blocked.
Private Function ParseAvail(ByVal v As String) As String
    Dim s As String
    s = LCase(Trim(v))
    If Len(s) = 0 Then ParseAvail = "": Exit Function
    Select Case s
        Case AVAIL_BUTTON, "yes", "y", "on", "true"
            ParseAvail = AVAIL_BUTTON
        Case AVAIL_CHAIN, "chain only", "chains", "chain", "chains-only"
            ParseAvail = AVAIL_CHAIN
        Case AVAIL_OFF, "no", "n", "false"
            ParseAvail = AVAIL_OFF
        Case Else
            ParseAvail = ""                  ' unrecognised - caller must block
    End Select
End Function

' "yes"/"y"/"true"/"1" -> True; "no"/"n"/"false"/"0" -> False. Returns False for
' anything else, so every caller fails CLOSED.
'
' Deliberately NOT IsTrue: that one accepts only TRUE/1 and fails OPEN to False
' for everything else, which is right for the optional _Packages enable column
' and wrong here.
'
' The spellings beyond "yes"/"no" are a REFERENCE-IMPLEMENTATION convenience
' (contract section 3), mirroring the same tolerance clause section 6 already
' grants availability. A port MAY accept only "yes" / "no".
Private Function ParseBool(ByVal raw As String, ByRef flag As Boolean) As Boolean
    Select Case LCase(Trim(raw))
        Case "yes", "y", "true", "1"
            flag = True
            ParseBool = True
        Case "no", "n", "false", "0"
            flag = False
            ParseBool = True
    End Select
End Function

' Display spelling for a bool cell, mirroring AvailLabel. An UNRECOGNISED value
' is handed back untouched rather than normalised - Finalize has to be able to
' see and block on what the user actually typed.
Private Function BoolLabel(ByVal raw As String) As String
    Dim flag As Boolean
    If ParseBool(raw, flag) Then
        BoolLabel = IIf(flag, "Yes", "No")
    Else
        BoolLabel = raw
    End If
End Function

' A validated dropdown removes free-typing from a two-value cell. Wrapped in
' On Error Resume Next for the same reason AvailDropdown is: validation can be
' refused (a protected sheet, some locales) and a refused dropdown must not abort
' the build - ParseBool still guards the value, and Finalize still blocks on
' anything unrecognised.
Private Sub BoolDropdown(rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                       Operator:=xlBetween, Formula1:="Yes,No"
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub

' The display label written into the cell for a token.
'
' NOT display-only: RenderPlacementBlock writes this result into the Setup cell
' (Core.bas ~2101) and ReadAvailMap reads that same cell back through ParseAvail
' (~2015), so the label round-trips into stored configuration.
'
' Case Else therefore fails CLOSED (issue #10). It used to return "Button", which
' meant an unrecognised stored value was silently rewritten as a live button on
' the next 'Create Setup Tab' - bypassing ComposeReport's block, which only ever
' sees "" and never saw it again once the rebuild had overwritten the cell.
' Returning "" leaves the cell blank so that block fires as designed.
Private Function AvailLabel(ByVal token As String) As String
    Select Case token
        Case AVAIL_BUTTON: AvailLabel = "Button"
        Case AVAIL_CHAIN: AvailLabel = "Chains only"
        Case AVAIL_OFF:   AvailLabel = "Off"
        Case Else:         AvailLabel = ""        ' fail closed - see above
    End Select
End Function

' id(lower) -> availability token, from the action catalog block. An id with no
' row is simply absent; what that MEANS is decided by AvailFor and depends on
' this map's provenance - see AVAIL_LEGACY_KEY. A v0.3.0 _Config (no Availability
' column at all) is flagged legacy here and still reads as everything on.
' An unrecognised cell is stored as "" so ComposeReport can block on it.
' Column layout comes from MarkerVerdict, not a decision made here - see its
' comment for why anything but "absent" reads the v1.2+ columns.
'
' THE THREE EXITS BELOW ARE THREE DIFFERENT ANSWERS - do not collapse them:
'   ws Is Nothing  no _Config at all      -> FRESH. Off. The case Off exists for.
'   kr = 0         _Config, but no block  -> Off, deliberately. Two very different
'                  histories reach it and they are indistinguishable from the sheet:
'                  a pre-v1.1 config written before the block existed, and a CURRENT
'                  config whose block was deleted or corrupted. Flagging it legacy
'                  was tried and reverted: it fails OPEN, turning a damaged current
'                  config's destructive actions back into live buttons. The pre-v1.1
'                  side costs nothing to fail closed - that era had no Availability
'                  column AND no placements, so there are no configured buttons to
'                  strip, and its config VALUES (separate key rows, read by Core_Cfg)
'                  load either way, which is what contract section 10's "v1.0.0/v1.1.0
'                  remain valid" promise is about. The user re-picks defaults that
'                  never existed; nobody loses a stored choice.
'   marker absent  block, no marker       -> LEGACY. Button. Real placements exist
'                  here and reading them as Off WOULD strip a working install.
Private Function ReadAvailMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String, detail As String
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadAvailMap = map
    If ws Is Nothing Then Exit Function
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr = 0 Then Exit Function     ' no catalog block at all: NOT flagged legacy - see below
    If MarkerVerdict(ws, kr, detail) = "absent" Then
        ' Pre-v1.2 block: the Availability column does not exist, so there are no
        ' values to read - but the map must still say WHERE it came from. Fresh
        ' (no _Config) and legacy are both an empty map, and they need opposite
        ' defaults, so the emptiness alone cannot decide it.
        map(AVAIL_LEGACY_KEY) = AVAIL_BUTTON
        Exit Function
    End If
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(id) > 0 Then map(LCase(id)) = ParseAvail(CStr(ws.Cells(i, VAL_COL + 1).Value))
    Next i
End Function

' Off unless the map says otherwise (v2.1.0). Absence used to mean Button; it now
' means Off, so a newly-registered action arrives switched off instead of
' appearing unannounced on an existing deployment.
'
' The ONE exception is a legacy map, flagged by ReadAvailMap because its stored
' block predates the Availability column. There, absence still means Button:
' the store described a deployment in which everything was live, and reading it
' as Off would silently strip every button from a working installation. Contract
' section 5 pins that default independently of section 6's for exactly this reason.
Private Function AvailFor(availMap As Object, ByVal id As String) As String
    AvailFor = AVAIL_OFF
    If availMap Is Nothing Then Exit Function
    If availMap.Exists(LCase(id)) Then
        AvailFor = CStr(availMap(LCase(id)))
    ElseIf availMap.Exists(AVAIL_LEGACY_KEY) Then
        AvailFor = AVAIL_BUTTON
    End If
End Function

' Fail closed on both: an invalid token ("") draws no button and does not run.
Private Function HasButton(availMap As Object, ByVal id As String) As Boolean
    HasButton = (AvailFor(availMap, id) = AVAIL_BUTTON)
End Function

Private Function IsRunnable(availMap As Object, ByVal id As String) As Boolean
    Dim t As String
    t = AvailFor(availMap, id)
    IsRunnable = (t = AVAIL_BUTTON Or t = AVAIL_CHAIN)
End Function

' A validated dropdown removes free-typing from a three-value column. Wrapped in
' On Error Resume Next because validation can be refused (a protected sheet, some
' locales) and a refused dropdown must not abort the build - ParseAvail still
' guards the value, and Finalize still blocks on anything unrecognised.
'
' v0.4.0: takes the whole block of action rows at once, not one cell at a time.
' Per-cell it made TWO COM validation calls per row inside RenderPlacementBlock's
' loop - about 164 at the 82-action fixture scale. Nothing depends on the
' per-cell form: the value is guarded by ParseAvail and blocked at Finalize
' whether or not the dropdown was accepted.
Private Sub AvailDropdown(rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                       Operator:=xlBetween, Formula1:="Button,Chains only,Off"
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    On Error GoTo 0
End Sub

' One row per registry action: hidden id in RID_COL, "caption [id]" (grey) in C, a
' yellow Availability cell with a validated dropdown in D, a yellow Placement cell
' in E. Prior answers are preserved from CfgSheet() by id. Unavailable actions STAY
' listed here - that is the owner's hard requirement: turning a button off must
' never make the action undiscoverable.
Private Function RenderPlacementBlock(ws As Worksheet, ByVal r As Long) As Long
    Dim prior As Object, priorAvail As Object, reg As Object, a As Object, id As String
    Dim firstDataRow As Long, lastDataRow As Long, reminderRow As Long
    Dim availabilityRange As String, f As String
    ' Fix round 3, fix 3: "hide" used to read as manual-layout-only ("leave one
    ' button off A MANUAL LAYOUT") - true only before fix round 3's fix 1, which
    ' made GatherButtons drop a hide-coded button before either filer runs, so
    ' it now works the same whether or not any placement code is set.
    r = BlockIntro(ws, r, "Action catalog  -  what each action is available as, and where it goes", _
                   "Availability: 'Button' shows a button; 'Chains only' hides the button but a " & _
                   "Chain can still run it; 'Off' hides it and stops it running anywhere. " & _
                   "Placement: leave every code blank for the automatic layout; fill in even one " & _
                   "placement code and codes take over - number = page, letters = column (e.g. 1A). " & _
                   "Or put a button on one of your OWN tabs with  Sheet!Cell  - " & _
                   "for example  TabName!B4  places it on cell B4 of your own tab named TabName, and " & _
                   "does not change where any other button goes. " & _
                   "Type  hide  to leave any one button off, whether or not you use placement codes.")
    ' Reserve this row now and write its formula AFTER the action rows exist.
    ' The formula spans the availability column of the whole block, so it needs
    ' the block's real extent - and predicting that from a registry count would
    ' drift silently the moment the loop below gained any filter, taking the
    ' bounded validation range at the end of this procedure with it. Writing the
    ' formula last costs nothing: Excel does not care what order cells are
    ' filled in, and lastDataRow stays derived from the loop that produces it.
    ' KEY_COL intentionally stays EMPTY on this row: blocks are located by
    ' machine key via FindKeyRow, never by row offset, so a keyless row cannot
    ' shift any block's parse - and a key here would make it look like a block.
    reminderRow = r
    NavLink ws.Cells(reminderRow, VAL_COL + 1), "Setup Menu", ADMIN_SHEET
    r = r + 1
    ws.Cells(r, KEY_COL).Value = LAYOUT_KEY
    ws.Cells(r, VAL_COL).Value = "Action"
    ws.Cells(r, VAL_COL + 1).Value = AVAIL_HEADER
    ws.Cells(r, VAL_COL + 2).Value = PLACE_HEADER
    ws.Range(ws.Cells(r, VAL_COL), ws.Cells(r, VAL_COL + 2)).Font.Bold = True
    HeaderNote ws.Cells(r, VAL_COL + 1), _
               "Every macro starts Off. Button puts it on a Macros page. " & _
               "Chains only makes it usable inside a chain but gives it no button of its own."
    HeaderNote ws.Cells(r, VAL_COL + 2), _
               "Placement forms: blank = automatic; 1A = page 1, column A; " & _
               "hide = leave it off the pages; TabName!B4 = put the button at cell B4 on your own tab named TabName."
    ' The machine-owned schema marker, in the hidden id column of the key row -
    ' NOT in the visible header. See Task 4 / contract section 5. Force Text so a
    ' General-formatted cell cannot coerce the marker to a number (belt-and-braces:
    ' SCHEMA_VER's "v" prefix already blocks numeric coercion on its own).
    ws.Cells(r, RID_COL).NumberFormat = "@"
    ws.Cells(r, RID_COL).Value = SCHEMA_VER
    r = r + 1
    firstDataRow = r

    Set prior = ReadCodeMap(CfgSheet())            ' preserve prior codes by id
    Set priorAvail = ReadAvailMap(CfgSheet())      ' preserve prior availability by id
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            id = CStr(a("id"))
            ws.Cells(r, VAL_COL).Value = CStr(a("caption")) & " [" & id & "]"
            ws.Cells(r, VAL_COL).Font.Color = RGB(90, 90, 90)
            ws.Cells(r, RID_COL).Value = id
            ws.Cells(r, VAL_COL + 1).Value = AvailLabel(AvailFor(priorAvail, id))
            If prior.Exists(LCase(id)) Then ws.Cells(r, VAL_COL + 2).Value = prior(LCase(id))
            Yellow ws.Cells(r, VAL_COL + 1)
            Yellow ws.Cells(r, VAL_COL + 2)
            r = r + 1
        Next a
    Next reg
    lastDataRow = r - 1

    ' The reminder, written now that the block's real extent is known. It fires
    ' whenever any action is not Off and is blank otherwise, so it appears the
    ' moment a dropdown changes - before Finalize, which is the point of using a
    ' formula rather than static text. The "<>" test excludes blanks: on a fresh
    ' import every cell reads Off and the line must stay empty.
    '
    ' D1 gives the reminder empty E1:H1 to overflow across. Column D on the
    ' reminder row holds the Setup Menu nav link, so column C has no empty
    ' neighbour there; widening C would distort every block without removing
    ' that overflow wall. AutoFit starts at firstContentRow (row 6), leaving a
    ' row-1 write at its default height and preventing the clipping this move fixes.
    '
    ' Compares against AvailLabel(AVAIL_OFF) - the "Off" actually written into
    ' the cell - not the AVAIL_OFF token "off". Excel's COUNTIFS text matching
    ' is believed case-insensitive, so the token would probably have worked, but
    ' "probably" is a claim about Excel that nothing here has run to confirm.
    ' Comparing against the literal the cell contains removes the question.
    If lastDataRow >= firstDataRow Then
        availabilityRange = ws.Range(ws.Cells(firstDataRow, VAL_COL + 1), _
                                     ws.Cells(lastDataRow, VAL_COL + 1)).Address( _
                                     RowAbsolute:=True, ColumnAbsolute:=True)
        f = "=IF(COUNTIFS(" & availabilityRange & ",""<>""," & availabilityRange & _
            ",""<>" & AvailLabel(AVAIL_OFF) & """)>0,""" & REMINDER_TEXT & ""","""")"
    Else
        f = "="""""
    End If
    ws.Cells(1, 4).Formula = f
    ws.Cells(1, 4).Font.Bold = True
    ws.Cells(1, 4).Font.Color = RGB(150, 0, 0)

    ' ONE validation call over the exact block of action rows. Bounded to the
    ' rows actually written - never the whole column - so it cannot leak
    ' validation onto the key/marker row above it or onto anything below the
    ' END_KEY sentinel.
    If lastDataRow >= firstDataRow Then
        AvailDropdown ws.Range(ws.Cells(firstDataRow, VAL_COL + 1), _
                               ws.Cells(lastDataRow, VAL_COL + 1))
    End If
    RenderPlacementBlock = r
End Function

' The config-format marker for the action catalog block, or "" when absent
' (meaning the block predates v1.2.0 and its page code is one column to the left).
' Deliberately NOT read from the visible header: RID_COL is hidden, is already how
' this codebase carries identity, and is copied into _Config with everything else.
Private Function BlockSchema(ws As Worksheet, ByVal keyRow As Long) As String
    If ws Is Nothing Or keyRow < 1 Then Exit Function
    BlockSchema = Trim(CStr(ws.Cells(keyRow, RID_COL).Value))
End Function

' Extracts the major version number from a schema marker of the form
' "v<major>.<minor>" (e.g. "v1.2" -> 1), and from a three-(or more-)part form
' like "v1.2.0" - spec/contract.md section 10 spells its own versions that way,
' so a future SCHEMA_VER bump is likely to look like it. ALSO accepts a bare
' "<major>.<minor>" with no "v" - the marker the interim da5de80 build wrote,
' before the "v" prefix (and the NumberFormat "@" guard added in 8b00a79)
' existed: a General-formatted cell coerced that bare text into the NUMBER 1.2,
' which reads back through CStr as "1.2" on an en-US locale. That data is
' genuinely v1.2-shaped - only its stamp is missing the "v" - so it is
' deliberately understood here rather than rejected as damaged.
' Returns False for anything that fits none of these: a present marker that is
' not a version number at all is corruption, and callers must not guess a
' major version for it (contract section 5 - fail closed, do not guess).
' HARDENED: guards CLng overflow on the major digit run, same bound as ParseCode.
Private Function SchemaMajor(ByVal marker As String, ByRef major As Long) As Boolean
    Dim s As String, dot As Long, digits As String, rest As String, i As Long, ch As String
    Dim prevDot As Boolean
    major = 0
    s = Trim(marker)
    If Len(s) = 0 Then Exit Function
    If LCase(Left(s, 1)) = "v" Then s = Mid(s, 2)      ' normal form; bare form skips this
    dot = InStr(s, ".")
    If dot < 2 Then Exit Function                      ' need >=1 digit before the first dot
    digits = Left(s, dot - 1)
    rest = Mid(s, dot + 1)                              ' minor, or "minor.patch[.more]"
    If Len(rest) = 0 Then Exit Function
    For i = 1 To Len(digits)
        ch = Mid(digits, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    If Len(digits) > 7 Then Exit Function               ' overflow guard (CLng-safe range)
    ' rest must be one or more digit groups separated by single dots ("2",
    ' "2.0", "2.0.1", ...) - anything else (letters, "..", a leading/trailing
    ' dot) is not a version number.
    prevDot = True                                       ' rest must not START with "."
    For i = 1 To Len(rest)
        ch = Mid(rest, i, 1)
        If ch = "." Then
            If prevDot Then Exit Function                ' leading "." or ".."
            prevDot = True
        ElseIf ch >= "0" And ch <= "9" Then
            prevDot = False
        Else
            Exit Function
        End If
    Next i
    If prevDot Then Exit Function                        ' rest must not END with "."
    major = CLng(digits)
    SchemaMajor = True
End Function

' The parsed major version of THIS BUILD's own SCHEMA_VER constant. SCHEMA_VER
' is a hardcoded literal controlled entirely by this codebase, so a parse
' failure here is a bug in this build, not a workbook problem. Raising instead
' of silently returning 0 matters: every caller compares a saved marker's major
' against this one, and a silent 0 would misclassify every real v1.x store as
' "unknown" and lock every workbook out with no recovery - exactly the failure
' mode this task exists to prevent, just self-inflicted instead of triggered by
' a stale workbook. Callers already run under On Error GoTo Fail, so this
' surfaces as an unmistakable "... error: SCHEMA_VER does not parse..." message
' instead of a wrong, silent comparison.
Private Function CurrentSchemaMajor() As Long
    Dim m As Long
    If Not SchemaMajor(SCHEMA_VER, m) Then
        Err.Raise vbObjectError + 1024, "CurrentSchemaMajor", _
            "SCHEMA_VER ('" & SCHEMA_VER & "') does not parse as a version number. " & _
            "This is a bug in this build, not a problem with the workbook."
    End If
    CurrentSchemaMajor = m
End Function

' THE single place that decides what an action-catalog marker means. Every
' caller that used to make its own marker decision (ReadAvailMap, ReadCodeMap,
' ConfigSchemaStatus, ComposeReport) now calls this instead of re-deriving the
' same verdict, so Finalize and a rebuild can never disagree about the same
' marker - there is exactly one place left that decides. Returns exactly one of:
'   "absent"      - no marker at all: a genuine pre-v1.2 block (contract
'                    section 5, "absent is not invalid"). The Availability
'                    column does not exist, and the code lives one column to
'                    the left of where a v1.2+ block keeps it.
'   "same-major"  - marker present, same major version as SCHEMA_VER (whatever
'                    the minor): the columns are exactly where a v1.2+ block
'                    keeps them, so it is fully readable.
'   "other-major" - marker present, parses, but a genuinely different major
'                    version: a real different build wrote this.
'   "unparseable" - marker present but does not parse as a version at all:
'                    corruption, not version skew.
'   "tampered"    - marker ABSENT, and the block does NOT have the pre-v1.2
'                    shape (see LooksPreV12): it still carries content at
'                    VAL_COL + 2, a column that only exists in a v1.2+ block,
'                    so the marker was removed or edited by hand rather than
'                    never having existed. This is the DEFAULT for a missing
'                    marker as of v2.2.0 - "absent" now has to be earned.
' For a reader (ReadAvailMap/ReadCodeMap), only "absent" vs "anything else"
' matters: "absent" means legacy columns, every other verdict means the v1.2+
' column layout is where this block actually put its data - even "tampered"
' and "unparseable", because the DATA is still very likely in the v1.2+
' columns; only the marker cell itself is untrustworthy. The gates
' (ConfigSchemaStatus/ComposeReport) are the ones that care about the full
' verdict and refuse to ACTIVATE on anything but "absent"/"same-major".
' detail is filled with a plain factual explanation for every verdict except
' "absent"/"same-major" - callers append their own call to action.
' Does this marker-less block actually HAVE the pre-v1.2 shape? (issue #12)
'
' This used to ask whether the visible header at VAL_COL + 1 still read
' "Availability", and call the block "tampered" only if it did. That fails OPEN in
' the one direction that matters: rename that header by hand and a v1.2+ block whose
' marker was deleted is classified "absent", which sends it down the legacy path and
' activates every action - destructive ones included - as Button. A user-editable
' visible caption was the sole discriminator for a machine-owned decision, which is
' exactly what contract section 5 says the marker exists to avoid.
'
' The burden is now inverted: a missing marker means "tampered" UNLESS the block
' positively corroborates the pre-v1.2 layout. The corroborating evidence is a THIRD
' COLUMN THAT WAS NEVER WRITTEN. A pre-v1.2 block has two columns - Action at
' VAL_COL, "Page code" at VAL_COL + 1 - and never touches VAL_COL + 2. A v1.2+ block
' inserted Availability at VAL_COL + 1 and pushed Placement to VAL_COL + 2, so ANY
' content at VAL_COL + 2 proves the v1.2+ layout and no rename can hide it. Verified
' against the v0.3.0 bundle's own renderer, not assumed.
'
' Content, not caption, on purpose: the only way to defeat this is to clear
' VAL_COL + 2 across the entire block, which destroys every placement code in the
' store. That is not a disguise - it leaves nothing worth misreading. And that is the
' honest limit of this function: a block hand-edited into an exact pre-v1.2 shape IS
' a pre-v1.2 shape as far as any sheet-level evidence goes. The marker is the only
' real authority; this only decides what to do once it is gone.
Private Function LooksPreV12(ws As Worksheet, ByVal keyRow As Long) As Boolean
    Dim nk As Long, i As Long
    nk = NextKeyRow(ws, keyRow)
    For i = keyRow To nk - 1
        If Len(Trim(CStr(ws.Cells(i, VAL_COL + 2).Value))) > 0 Then Exit Function
    Next i
    LooksPreV12 = True
End Function

Private Function MarkerVerdict(ws As Worksheet, ByVal keyRow As Long, ByRef detail As String) As String
    Dim marker As String, mMajor As Long
    detail = ""
    If ws Is Nothing Or keyRow < 1 Then MarkerVerdict = "absent": Exit Function
    marker = BlockSchema(ws, keyRow)
    If Len(marker) = 0 Then
        If LooksPreV12(ws, keyRow) Then
            MarkerVerdict = "absent"
        Else
            detail = "The hidden format marker for the Action catalog block is missing, even " & _
                     "though the block still has this version's third column, so saved " & _
                     "Availability choices cannot be trusted."
            MarkerVerdict = "tampered"
        End If
        Exit Function
    End If
    If Not SchemaMajor(marker, mMajor) Then
        detail = "This workbook's saved configuration format stamp ('" & marker & "') does not " & _
                 "look like a version number, so it cannot be trusted."
        MarkerVerdict = "unparseable"
        Exit Function
    End If
    If mMajor = CurrentSchemaMajor() Then
        MarkerVerdict = "same-major"
    Else
        detail = "This workbook's saved configuration was written by a different version of " & _
                 "this toolkit (format '" & marker & "'; this build understands '" & SCHEMA_VER & "')."
        MarkerVerdict = "other-major"
    End If
End Function

' Classify the ACTIVATED config before anything destructive touches the Setup
' sheet - MarkerVerdict's verdicts, translated to this module's action
' vocabulary, plus the one thing MarkerVerdict cannot tell on its own: whether
' an "absent" marker's recovered legacy page codes actually parse. Returns:
'   "none"    - no config yet (a fresh workbook); nothing to lose
'   "ok"      - MarkerVerdict "same-major": readable as-is
'   "legacy"  - MarkerVerdict "absent", every recovered page code validates
'   "invalid" - MarkerVerdict "absent", but a recovered code does not parse
'   "unknown" - MarkerVerdict "other-major"
'   "damaged" - MarkerVerdict "unparseable" or "tampered"
' detail is filled with a human-readable explanation for every status except
' "none"/"ok"/"legacy". badIds is ALWAYS set to a (possibly empty) Dictionary;
' for "invalid" it holds the lower-cased ids whose recovered code failed
' ParseCode OR ParseAnchor, whichever CodeKind routed it to - the same set
' detail lists in prose - so a caller that wants to clear ONLY the unreadable
' rows (not the whole block) has a single source of truth instead of
' re-parsing detail's text.
Private Function ConfigSchemaStatus(ByRef detail As String, ByRef badIds As Object) As String
    Dim ws As Worksheet, kr As Long, nk As Long, i As Long
    Dim id As String, code As String, tn As Long, cx As Long, cl As String
    Dim aSheet As String, aCell As String
    Dim verdict As String
    detail = ""
    Set badIds = CreateObject("Scripting.Dictionary")
    Set ws = CfgSheet()
    If ws Is Nothing Then ConfigSchemaStatus = "none": Exit Function
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr = 0 Then ConfigSchemaStatus = "none": Exit Function

    verdict = MarkerVerdict(ws, kr, detail)
    Select Case verdict
        Case "same-major"
            ConfigSchemaStatus = "ok"
        Case "other-major"
            ConfigSchemaStatus = "unknown"
        Case "unparseable", "tampered"
            ConfigSchemaStatus = "damaged"
        Case "absent"
            ' Legacy: validate every page code recovered from the OLD column
            ' before letting anything overwrite the sheet holding them.
            nk = NextKeyRow(ws, kr)
            For i = kr + 1 To nk - 1
                id = Trim(CStr(ws.Cells(i, RID_COL).Value))
                If Len(id) > 0 Then
                    code = Trim(CStr(ws.Cells(i, VAL_COL + 1).Value))
                    ' Route through CodeKind so an anchor is never classified
                    ' "invalid". That path CLEARS the offending rows, so getting
                    ' this wrong would silently delete every anchor on a marker
                    ' recovery. A legacy store predates anchors, so this branch
                    ' is defensive rather than reachable today - which is exactly
                    ' why it has to be right, because nothing will exercise it.
                    Select Case CodeKind(code)
                        Case "page"
                            If Not ParseCode(code, tn, cx, cl) Then
                                detail = detail & "  - " & id & ": '" & code & "'" & vbCrLf
                                badIds(LCase(id)) = True
                            End If
                        Case "anchor"
                            If Not ParseAnchor(code, aSheet, aCell) Then
                                detail = detail & "  - " & id & ": '" & code & "'" & vbCrLf
                                badIds(LCase(id)) = True
                            End If
                    End Select
                End If
            Next i
            If Len(detail) > 0 Then
                detail = "These saved placements could not be read from the older " & _
                         "configuration format:" & vbCrLf & detail
                ConfigSchemaStatus = "invalid"
            Else
                ConfigSchemaStatus = "legacy"
            End If
    End Select
End Function

' id(lower) -> code, merging the action catalog block and the chains block.
' v0.4.0 inserted Availability between Action and Page code, so a pre-v1.2
' block keeps its code one column to the left. Column layout comes from
' MarkerVerdict, not a decision made here - see its comment for why anything
' but "absent" reads the v1.2+ columns. Callers that could DESTROY the
' recovered values (Core_BuildSetup) must consult ConfigSchemaStatus first;
' this reader is permissive on purpose so the report can describe what it found.
Private Function ReadCodeMap(ws As Worksheet) As Object
    Dim map As Object, kr As Long, nk As Long, i As Long, id As String, codeCol As Long
    Dim chains As Collection, ch As Object, detail As String
    Set map = CreateObject("Scripting.Dictionary")
    Set ReadCodeMap = map
    If ws Is Nothing Then Exit Function
    ' actions
    kr = FindKeyRow(ws, LAYOUT_KEY)
    If kr > 0 Then
        If MarkerVerdict(ws, kr, detail) = "absent" Then
            codeCol = VAL_COL + 1                ' legacy (pre-v1.2) layout
        Else
            codeCol = VAL_COL + 2
        End If
        nk = NextKeyRow(ws, kr)
        For i = kr + 1 To nk - 1
            id = Trim(CStr(ws.Cells(i, RID_COL).Value))
            If Len(id) > 0 Then map(LCase(id)) = Trim(CStr(ws.Cells(i, codeCol).Value))
        Next i
    End If
    ' chains
    Set chains = ReadChains(ws)
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 Then map(LCase(CStr(ch("id")))) = CStr(ch("code"))
    Next ch
End Function

' Mint chain.<slug> into any chains-block row that has a Button name but no id.
' Runs on the LIVE Setup sheet AFTER the user confirms, so Finalize's copy persists them.
Private Sub AssignChainIds(ws As Worksheet)
    Dim kr As Long, nk As Long, i As Long, used As Object, label As String, id As String
    Set used = CreateObject("Scripting.Dictionary")
    kr = FindKeyRow(ws, CHAINS_KEY)
    If kr = 0 Then Exit Sub
    nk = NextKeyRow(ws, kr)
    For i = kr + 1 To nk - 1                 ' register existing ids
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(id) > 0 Then used(LCase(id)) = True
    Next i
    For i = kr + 1 To nk - 1                 ' mint for named rows lacking an id
        label = Trim(CStr(ws.Cells(i, VAL_COL).Value))
        id = Trim(CStr(ws.Cells(i, RID_COL).Value))
        If Len(label) > 0 And Len(id) = 0 Then
            id = MintChainId(label, used)
            ws.Cells(i, RID_COL).Value = id
            used(LCase(id)) = True
        End If
    Next i
End Sub

' Build the review text + count STRUCTURAL problems (things that must be fixed
' before activating). Validates chain steps against the registry, action + chain
' ids against the reserved chain.* namespace, and every code against ParseCode.
' Unassigned/hidden button lists and the layout note are SOFT (informational).
Private Function ComposeReport(ws As Worksheet, ByRef structural As Long) As String
    Dim s As String, chains As Collection, ch As Object, ids As Object, stepId As Variant
    Dim tabNum As Long, colIx As Long, colL As String, code As String
    Dim map As Object, k As Variant, unassigned As String, hidden As String
    Dim regX As Object, aX As Object
    Dim ch2 As Object, actIds As Object, seenIds As Object, cid As String
    Dim availMap As Object, offList As String, badAvail As String
    Dim chainOnly As String, nOff As Long, t As String
    Dim layoutKR As Long, verdict As String, vDetail As String
    Dim aSheet As String, aCell As String
    Dim aKey As String, placedList As String, seenTargets As Object
    Set ids = ActionIdSet()                        ' lowercased set of valid action ids
    Set chains = ReadChains(ws)
    Set seenTargets = CreateObject("Scripting.Dictionary")

    s = "CHAINS:" & vbCrLf
    If chains.Count = 0 Then s = s & "  (none)" & vbCrLf
    For Each ch In chains
        s = s & "  - " & ch("label") & "  ->  "
        If ch("steps").Count = 0 Then
            s = s & "(no steps!)": structural = structural + 1
        Else
            For Each stepId In ch("steps")
                s = s & CStr(stepId)
                If Not ids.Exists(LCase(CStr(stepId))) Then
                    s = s & "[UNKNOWN]": structural = structural + 1
                End If
                s = s & "  "
            Next stepId
        End If
        code = CStr(ch("code"))
        Select Case CodeKind(code)
            Case "page"
                If Not ParseCode(code, tabNum, colIx, colL) Then _
                    s = s & "  {bad code '" & code & "'}": structural = structural + 1
            Case "anchor"
                If Not ParseAnchor(code, aSheet, aCell) Then _
                    s = s & "  {bad code '" & code & "'}": structural = structural + 1
        End Select
        s = s & vbCrLf
    Next ch

    ' Placement codes: validity + collect unassigned/hidden.
    Set map = ReadCodeMap(ws)
    For Each k In map.Keys
        code = CStr(map(k))
        Select Case CodeKind(code)
            Case "hide"
                hidden = hidden & "  - " & HiddenTokenFor(CStr(k), chains) & vbCrLf
            Case ""
                ' Fix round 2, fix 2: was the bare lowercased key - same
                ' inconsistency fix round 1's fix 5 removed from the HIDDEN list
                ' just above, now closed here too so both lists in one report
                ' agree on "caption [id]".
                unassigned = unassigned & "  - " & HiddenTokenFor(CStr(k), chains) & vbCrLf
            Case "page"
                If Not ParseCode(code, tabNum, colIx, colL) Then _
                    s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf: structural = structural + 1
            Case "anchor"
                If Not ParseAnchor(code, aSheet, aCell) Then
                    s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf
                    structural = structural + 1
                ElseIf IsEngineOwnedSheet(aSheet) Then
                    ' Checked before sheet-exists on purpose: an engine-owned tab
                    ' (e.g. "Macros 1") legitimately does not exist yet on a brand
                    ' new deployment's FIRST Finalize - Core_BuildControlPanel is
                    ' what creates it, and that runs AFTER Finalize. Sheet-exists
                    ' first would tell the user "no tab named Macros 1", and the
                    ' natural (wrong) response - go create one - is exactly what
                    ' DeletePanels then destroys on the next build. IsEngineOwnedSheet
                    ' inspects a string only and has no precondition on the sheet
                    ' existing, so it is safe to run first.
                    s = s & "RESERVED SHEET for " & k & ": '" & code & "' - the engine " & _
                        "rebuilds that tab on every build; pick one of your own." & vbCrLf
                    structural = structural + 1
                ElseIf Not Core_SheetExists(aSheet) Then
                    s = s & "NO SUCH SHEET for " & k & ": '" & code & "' - this workbook " & _
                        "has no tab named " & Chr(34) & aSheet & Chr(34) & "." & vbCrLf
                    structural = structural + 1
                Else
                    ' Normalise through the resolved absolute address so "B4" and
                    ' "$B$4" collide correctly. Runs only after the sheet-exists
                    ' check has passed, so an empty key here means a cell part
                    ' that parses but does not resolve (a row past the sheet's
                    ' last) - a bad code, not a duplicate.
                    aKey = LCase(aSheet) & "!" & AnchorAddrKey(aSheet, aCell)
                    If Right$(aKey, 1) = "!" Then
                        s = s & "BAD CODE for " & k & ": '" & code & "'" & vbCrLf
                        structural = structural + 1
                    ElseIf seenTargets.Exists(aKey) Then
                        ' Two buttons drawn at identical coordinates means one
                        ' physically covers the other and is unclickable. Blocking
                        ' is consistent with how a bad page code is already
                        ' treated: every placement error here is structural.
                        s = s & "DUPLICATE TARGET: " & CStr(seenTargets(aKey)) & " and " & _
                            k & " both point at " & code & "." & vbCrLf
                        structural = structural + 1
                    Else
                        seenTargets(aKey) = CStr(k)
                        placedList = placedList & "  - " & HiddenTokenFor(CStr(k), chains) & _
                                     "  ->  " & code & vbCrLf
                    End If
                End If
        End Select
    Next k

    ' (a) no package action id may live in EITHER reserved namespace (contract
    '     section 7): chain.*, or core.* as of v2.1.0. The core.* arm is not
    '     symmetry - see IsCoreId for what a "core.legacy" action would do to the
    '     availability map.
    For Each regX In BuildRegistry()
        For Each aX In regX("actions")
            If IsChainId(CStr(aX("id"))) Then
                s = s & "RESERVED: action id '" & aX("id") & "' uses chain.*" & vbCrLf
                structural = structural + 1
            ElseIf IsCoreId(CStr(aX("id"))) Then
                s = s & "RESERVED: action id '" & aX("id") & "' uses core.*" & vbCrLf
                structural = structural + 1
            End If
        Next aX
    Next regX
    ' (b) each existing chain id must start chain., not collide with an action id, and be unique
    Set actIds = ActionIdSet()
    Set seenIds = CreateObject("Scripting.Dictionary")
    For Each ch2 In ReadChains(ws)
        cid = CStr(ch2("id"))
        If Len(cid) > 0 Then
            If Not IsChainId(cid) Then _
                s = s & "BAD chain id '" & cid & "' (must start chain.)" & vbCrLf: structural = structural + 1
            If actIds.Exists(LCase(cid)) Then _
                s = s & "COLLISION: chain id '" & cid & "' matches an action id" & vbCrLf: structural = structural + 1
            If seenIds.Exists(LCase(cid)) Then _
                s = s & "DUPLICATE chain id '" & cid & "'" & vbCrLf: structural = structural + 1
            seenIds(LCase(cid)) = True
        End If
    Next ch2

    ' Availability. An unrecognised cell and an Off action used by a chain are both
    ' STRUCTURAL - they block activation. "Chains only" is reported for confidence,
    ' not as a problem: it is an authored state, not an accident.
    '
    ' Format-marker integrity, contract section 5's fail-closed requirement -
    ' delegated entirely to MarkerVerdict (see its comment) so this can never
    ' disagree with what a rebuild would decide about the same marker. Only
    ' "same-major"/"absent" pass; every other verdict blocks Finalize. This
    ' function's job is to report the problem and name the button that fixes
    ' it - NOT to predict what that rebuild will or will not preserve, which
    ' belongs to the rebuild's own confirmation (the one place that actually
    ' knows, and the one the user acts on).
    layoutKR = FindKeyRow(ws, LAYOUT_KEY)
    verdict = MarkerVerdict(ws, layoutKR, vDetail)
    ' "other-major" and "unparseable"/"tampered" share nothing but a header verb -
    ' their real recoveries are opposite. GateConfigMigration REFUSES on
    ' "other-major" ("unknown") and tells the user to re-import that version's
    ' bundle; it lets "unparseable"/"tampered" ("damaged") through to
    ' ConfirmAndClear's guarded rebuild. Naming "Create Setup Tab" for
    ' other-major would send the user to a button that then refuses them.
    Select Case verdict
        Case "other-major"
            s = s & vbCrLf & "ACTION CATALOG FORMAT MISMATCH: " & vDetail & " Nothing has been " & _
                "activated. Re-import that version's MacrosBundle.bas to continue - this build " & _
                "will not rebuild over a newer saved format." & vbCrLf
            structural = structural + 1
        Case "unparseable"
            s = s & vbCrLf & "ACTION CATALOG FORMAT MISMATCH: " & vDetail & " Nothing has been " & _
                "activated. Click 'Create Setup Tab' to rebuild the Action catalog block." & vbCrLf
            structural = structural + 1
        Case "tampered"
            s = s & vbCrLf & "ACTION CATALOG FORMAT MARKER MISSING: " & vDetail & " Nothing has " & _
                "been activated. Click 'Create Setup Tab' to rebuild the Action catalog block." & vbCrLf
            structural = structural + 1
    End Select
    Set availMap = ReadAvailMap(ws)
    For Each regX In BuildRegistry()
        For Each aX In regX("actions")
            t = AvailFor(availMap, CStr(aX("id")))
            If Len(t) = 0 Then
                badAvail = badAvail & "  - " & CStr(aX("id")) & vbCrLf
                structural = structural + 1
            ElseIf t = AVAIL_OFF Then
                nOff = nOff + 1
                offList = offList & "  - " & CStr(aX("caption")) & " [" & CStr(aX("id")) & "]" & vbCrLf
            ElseIf t = AVAIL_CHAIN Then
                chainOnly = chainOnly & "  - " & CStr(aX("caption")) & " [" & CStr(aX("id")) & "]" & vbCrLf
            End If
        Next aX
    Next regX
    If Len(badAvail) > 0 Then
        s = s & vbCrLf & "MISSING or UNRECOGNISED Availability (use Button, Chains only, or Off):" & _
            vbCrLf & badAvail
    End If
    For Each ch In chains
        For Each stepId In ch("steps")
            If ids.Exists(LCase(CStr(stepId))) Then
                If AvailFor(availMap, CStr(stepId)) = AVAIL_OFF Then
                    s = s & "OFF ACTION IN CHAIN: '" & CStr(ch("label")) & "' uses " & _
                        CStr(stepId) & ", which is switched Off." & vbCrLf
                    structural = structural + 1
                End If
            End If
        Next stepId
    Next ch
    If nOff > 0 Then
        ' Deliberately does NOT claim the settings have already collapsed: that
        ' only happens for a field that declares usedBy (neither shipped package
        ' does yet), and only on the NEXT "Create Setup Tab" - Finalize never
        ' redraws Setup. Overclaiming here would describe a fixture-only feature
        ' as if it already applied to a real deployment.
        s = s & vbCrLf & "OFF (" & nOff & " action(s); no button, cannot run. Any field that " & _
            "declares this action in its usedBy will collapse into the settings band the next " & _
            "time 'Create Setup Tab' rebuilds Setup):" & vbCrLf & offList
    End If
    If Len(chainOnly) > 0 Then
        s = s & vbCrLf & "CHAINS ONLY (no button; runs only inside a Chain):" & vbCrLf & chainOnly
    End If

    ' Fix round 3, fix 2 (live path - this runs from Core_FinalizeSetup, before
    ' the user decides to activate): was AnyCode(map), which counts "hide" as a
    ' layout-activating code - the exact bug fix round 2's fix 5 exists to
    ' prevent, just at a different call site. A hide-only store would have shown
    ' "A layout is ACTIVE" here while the engine actually runs auto-layout.
    ' AnyPlacementCode agrees with what Core_BuildControlPanel's
    ' BuildAutoLayout/BuildManualLayout choice (Task 8) actually does. Both
    ' messages describe v0.4.0's real behaviour, not the pre-auto-layout
    ' "single Macros tab" default this replaced.
    If AnyPlacementCode(map) Then
        ' Fix wave: codes choose where a button STARTS, not a guarantee of the
        ' page/column it ends up on - contract section 6's Column-cap overflow
        ' clause can still relocate it once a page or column is full.
        s = s & vbCrLf & "A layout is ACTIVE - your placement codes choose where each button starts, " & _
            "but a full page or column can still move it to a continuation one." & vbCrLf
        If Len(unassigned) > 0 Then s = s & "Unassigned -> page 1's Unassigned column (or a continuation page, once the page has no free column slot left):" & vbCrLf & unassigned
    Else
        s = s & vbCrLf & "No placement codes -> categories will be packed onto numbered pages automatically." & vbCrLf
    End If
    ' "hide" is honoured in EITHER mode now (fix 1), so this is no longer
    ' conditional on AnyPlacementCode.
    If Len(hidden) > 0 Then s = s & "Hidden (not placed, in either mode):" & vbCrLf & hidden
    If Len(placedList) > 0 Then _
        s = s & "Placed on your own sheets (not on a generated page):" & vbCrLf & placedList
    ComposeReport = s
End Function

Private Function ActionIdSet() As Object
    Dim d As Object, reg As Object, a As Object
    Set d = CreateObject("Scripting.Dictionary")
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            d(LCase(CStr(a("id")))) = True
        Next a
    Next reg
    Set ActionIdSet = d
End Function

' Fix round 2, fix 5: answers "is there a REAL placement code", not "is there
' any non-blank cell at all" - the latter would be True for a SINGLE "hide"
' entry with every other code left blank, and a user who hides one button
' would then flip the whole workbook into manual mode (every other,
' still-uncoded button reads as manually placed, landing in page 1's
' Unassigned column, and the Setup tab's own help text would tell them their
' placement codes control a layout they never authored). "hide" omits a button; it
' does not place one, so it must not count as a placement here.
' Core_BuildControlPanel (Task 8) uses THIS helper, and only this one, both to
' choose BuildAutoLayout vs BuildManualLayout.
' v0.4.0: a "Sheet!Cell" ANCHOR gets the same carve-out, and for the same reason.
' An anchor does not place a button on a generated page at all - it puts one on a
' sheet the engine otherwise never touches - so it is not a page-layout decision.
' Without this, typing "TabName!B4" on ONE action would flip the workbook into manual
' mode and dump the other 81 buttons into page 1's Unassigned column. This is the
' third carve-out of the same shape; contract section 6 now carries the matching
' clause ("an anchor is NOT a placement code for the mode decision"), stated in
' both directions, so a port reading section 6 literally cannot reintroduce the
' bug - which the ledger records almost happening with "hide".
Private Function AnyPlacementCode(map As Object) As Boolean
    Dim k As Variant
    For Each k In map.Keys
        If CodeKind(CStr(map(k))) = "page" Then AnyPlacementCode = True: Exit Function
    Next k
End Function

'================================================================
' 6. FINALIZE  (validate the Setup sheet, then activate -> _Config)
'================================================================
Public Sub Core_FinalizeSetup()
    Dim ws As Worksheet, cfg As Worksheet, reg As Object, f As Object
    Dim rpt As String, warn As Long, v As String, msg As String
    Dim boolBad As String, boolBadN As Long, bTmp As Boolean
    On Error GoTo Fail
    InvalidateRegistry              ' a build begins here - see InvalidateRegistry

    If Not Core_SheetExists(SETUP_SHEET) Then
        MsgBox "No '" & SETUP_SHEET & "' tab found. Run 'Create Setup Tab' first.", _
               vbExclamation, "Finalize Setup Tab"
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(SETUP_SHEET)

    ' --- validate scalar fields against this workbook (before activating) ---
    For Each reg In BuildRegistry()
        For Each f In reg("fields")
            v = ReadScalar(ws, CStr(f("key")))
            ' LCase, matching RenderScalar's own LCase(CStr(f("kind"))) = "bool"
            ' test. Without it this Select fails OPEN in the one direction that
            ' matters: a package declaring kind "Bool" would still get the Yes/No
            ' dropdown from the renderer but NO Finalize validation, so the "ye"
            ' typo this block exists to catch would sail through. Normalising the
            ' selector rather than one Case keeps the three arms symmetrical; the
            ' sheet/file arms only add soft warnings, so nothing newly blocks.
            Select Case LCase(CStr(f("kind")))
                Case "sheet"
                    If v <> "" Then
                        If Not Core_SheetExists(v) Then
                            rpt = rpt & "- '" & f("label") & "': tab '" & v & "' not found." & vbCrLf
                            warn = warn + 1
                        End If
                    End If
                Case "file"
                    If v <> "" Then
                        If Core_FindFile(v) = "" Then
                            rpt = rpt & "- '" & f("label") & "': no file matches prefix '" & v & "'." & vbCrLf
                            warn = warn + 1
                        End If
                    End If
                Case "bool"
                    ' STRUCTURAL, not a soft note: a user who typed "ye" meaning
                    ' "yes" must not be silently handed a "no". Contract section 3.
                    If v <> "" Then
                        If Not ParseBool(v, bTmp) Then
                            boolBad = boolBad & "  - '" & f("label") & "': '" & v & "'" & vbCrLf
                            boolBadN = boolBadN + 1
                        End If
                    End If
            End Select
        Next f
    Next reg

    If warn = 0 Then
        ' Scoped deliberately: this only covers the scalar sheet/file field checks
        ' just above. The compose/placement report appended below can still add
        ' its own "problem(s) above must be fixed" line beneath this - the two
        ' must never be allowed to read as one verdict.
        msg = "Sheet and file checks passed."
    Else
        msg = warn & " item(s) to note:" & vbCrLf & vbCrLf & rpt
    End If

    ' --- Compose Your Buttons: validate chains + placement ---
    Dim composeRep As String, structural As Long
    composeRep = ComposeReport(ws, structural)
    If boolBadN > 0 Then
        composeRep = composeRep & vbCrLf & "UNRECOGNISED Yes/No value (use Yes or No):" & _
                     vbCrLf & boolBad
        structural = structural + boolBadN
    End If
    msg = msg & vbCrLf & vbCrLf & String(40, "-") & vbCrLf & composeRep
    If structural > 0 Then
        MsgBox msg & vbCrLf & structural & " problem(s) above must be fixed before activating.", _
               vbExclamation, "Finalize Setup Tab"
        Exit Sub
    End If

    If MsgBox(msg & vbCrLf & "Activate this configuration now?", _
              vbQuestion + vbYesNo, "Finalize Setup Tab") = vbNo Then Exit Sub

    AssignChainIds ws          ' mint ids now, after confirm, before the copy persists them

    ' --- activate: copy Setup -> hidden _Config ---
    '
    ' COPY FIRST, VERIFY, THEN SWAP (issue #8). This used to delete the old _Config
    ' and then copy Setup into its place, leaving a window - between the delete and
    ' the rename - in which the workbook had NO _Config at all. A crash, a hang the
    ' user force-quits, or any failure inside ws.Copy in that window destroyed the
    ' entire stored configuration: every setting, every placement, every Availability
    ' choice, with no undo and no warning. _Config is the only memory that survives a
    ' rebuild.
    '
    ' The order is now copy -> verify -> delete -> rename, so the destructive step is
    ' reached only once a complete, checked replacement already exists on disk. If the
    ' copy or the verification fails, the old _Config is still there, untouched.
    '
    ' This does NOT make the activate atomic - Excel offers no transaction - and it is
    ' not claimed to. A crash in the (now much smaller) gap leaves the new config under
    ' the stage name instead of no config at all, which CfgSheet adopts on the next
    ' read. Total, silent loss becomes a state that heals itself. ADR 0003 decision
    ' 4's marker-last ordering within the block is unaffected by any of this.
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim old As Worksheet, staged As Worksheet
    ' CfgSheet, not a raw Worksheets() lookup: it performs the orphan-stage recovery
    ' described above, so by the time we clear a leftover stage below, that stage is
    ' guaranteed redundant rather than the only surviving copy.
    Set old = CfgSheet()

    Set staged = Nothing
    On Error Resume Next
    Set staged = ThisWorkbook.Worksheets(CONFIG_STAGE)
    On Error GoTo Fail
    If Not staged Is Nothing Then
        staged.Visible = xlSheetVisible
        staged.Delete
        Set staged = Nothing
    End If

    ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set staged = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    staged.name = CONFIG_STAGE

    ' Verify BEFORE anything destructive. A truncated or empty copy that went on to
    ' replace a good _Config would be the same total loss by a slower route, so the
    ' copy has to prove it carries the action catalog block before the old one dies.
    If FindKeyRow(staged, LAYOUT_KEY) = 0 Then
        staged.Delete
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        MsgBox "Activation stopped: the copy of the Setup tab did not come through " & _
               "complete - its Action catalog block is missing." & vbCrLf & vbCrLf & _
               "Your saved configuration has NOT been changed and nothing was lost. " & _
               "Click 'Finalize Setup Tab' again to retry.", _
               vbExclamation, "Finalize Setup Tab"
        Exit Sub
    End If

    ' Now the swap. Remove the prior _Config: it is stored xlSheetVeryHidden, and a
    ' very-hidden sheet can resist .Delete (the delete then silently fails under
    ' On Error Resume Next, _Config survives, and the rename below collides -> "That
    ' name is already taken", orphaning a "Setup (2)" copy on every retry). Unhiding
    ' it first makes the delete reliable. That reasoning is unchanged from before
    ' issue #8 - it simply happens later now.
    If Not old Is Nothing Then
        old.Visible = xlSheetVisible
        old.Delete
    End If
    Set cfg = staged
    cfg.name = CONFIG_SHEET
    ws.Activate                        ' a very-hidden sheet cannot be the active sheet
    cfg.Visible = xlSheetVeryHidden
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Core_BuildControlPanel True        ' rebuild the Macros pages so chains/placement appear now
    ' Fix round 1, fix 2: Core_BuildControlPanel cannot raise past its own Fail:
    ' handler (it shows its own error and returns normally), so without this
    ' check the message below would unconditionally claim success even when the
    ' rebuild just failed - right after the panels were already deleted, i.e.
    ' on a workbook with no buttons at all. mLastBuildOk is set True only once
    ' the whole rebuild pipeline finishes; it is reset False at the top of
    ' every Core_BuildControlPanel call, so it can never be stale here.
    ' Final-review fix wave, fix 1: mLastPlacedNote appended on the SUCCESS branch
    ' only. Core_BuildControlPanel composes it every call but shows it itself only
    ' when non-silent, and the call above is silent - so without this append the
    ' anchored-button counts, including the "could NOT be placed" one, reach nobody
    ' on the normal user path. Not appended to the failure branch: that branch is
    ' already telling the user the rebuild did not finish, and a placement tally
    ' from a half-finished build would only muddy it.
    If mLastBuildOk Then
        ' "The Macros pages have been rebuilt" is FALSE when none were built, and
        ' zero pages is a legitimate configuration a user reaches on purpose (every
        ' action Off, hidden, or anchored). Saying it anyway sends them looking for
        ' tabs that were never meant to exist - see mLastNoPages' declaration comment.
        If mLastNoPages Then
            MsgBox "Configuration activated." & vbCrLf & vbCrLf & _
                   NO_PAGES_MSG & mLastPlacedNote, _
                   vbInformation, "Finalize Setup Tab"
        Else
            MsgBox "Configuration activated." & vbCrLf & vbCrLf & _
                   "The Macros pages have been rebuilt to match your Setup answers." & _
                   mLastPlacedNote, _
                   vbInformation, "Finalize Setup Tab"
        End If
    Else
        MsgBox "Configuration activated, but rebuilding the Macros pages failed - see the error " & _
               "shown just now." & vbCrLf & vbCrLf & _
               "Your Setup answers were saved. Click 'Finalize Setup Tab' again to retry the rebuild.", _
               vbExclamation, "Finalize Setup Tab"
    End If
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Core_FinalizeSetup error: " & Err.Description, vbExclamation, "Finalize Setup Tab"
End Sub

' Strip the setup surfaces once the workbook is configured, while preserving the
' activated config and every generated macro page.
Public Sub Core_CompleteSetup()
    Dim ws As Worksheet, target As Worksheet, ans As VbMsgBoxResult
    Dim nPages As Long
    Const TITLE As String = "Complete Macro Setup"
    On Error GoTo Fail

    If Not Core_SheetExists(CONFIG_SHEET) Then
        MsgBox "Setup has never been finalized, so there is nothing to complete." & vbCrLf & vbCrLf & _
               "Click 'Finalize Setup Tab' first." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    For Each ws In ThisWorkbook.Worksheets
        If IsPanelName(ws.name) Then
            nPages = nPages + 1
            If target Is Nothing Then Set target = ws
        End If
    Next ws

    If nPages = 0 Then
        MsgBox "There are no macro pages in this workbook yet, so removing the setup tabs " & _
               "would leave you with nothing at all." & vbCrLf & vbCrLf & _
               "Switch on the macros you want, click 'Finalize Setup Tab', then run this " & _
               "again." & vbCrLf & vbCrLf & "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
    End If

    ans = MsgBox("This removes the setup tabs from this workbook:" & vbCrLf & _
                 "    Setup" & vbCrLf & "    Setup Menu" & vbCrLf & "    Macros Home" & vbCrLf & vbCrLf & _
                 "Your macro buttons and all your settings are kept." & vbCrLf & vbCrLf & _
                 "IF YOU HAVE CHANGED ANYTHING ON THE SETUP TAB SINCE YOU LAST CLICKED " & _
                 "'Finalize Setup Tab', CLICK NO NOW, FINALIZE FIRST, THEN COME BACK. " & _
                 "Unfinalized changes will be lost." & vbCrLf & vbCrLf & _
                 "To get the Setup tab back later: press Alt+F8, choose Core_BuildSetup, " & _
                 "click Run." & vbCrLf & vbCrLf & "Remove the setup tabs now?", _
                 vbYesNo + vbExclamation + vbDefaultButton2, TITLE)
    If ans <> vbYes Then Exit Sub

    Core_SpeedOn
    target.Activate
    For Each ws In ThisWorkbook.Worksheets
        If IsPanelName(ws.name) Then
            ws.Cells(1, 2).Hyperlinks.Delete
            ws.Cells(1, 2).Clear
        End If
    Next ws

    If Core_SheetExists(HOME_SHEET) Then ThisWorkbook.Worksheets(HOME_SHEET).Delete
    If Core_SheetExists(SETUP_SHEET) Then ThisWorkbook.Worksheets(SETUP_SHEET).Delete
    If Core_SheetExists(ADMIN_SHEET) Then ThisWorkbook.Worksheets(ADMIN_SHEET).Delete
    Core_SpeedOff
    MsgBox "Setup is complete. Your macro buttons are ready to use." & vbCrLf & vbCrLf & _
           "Your settings are saved and hidden. To change them later: press Alt+F8, " & _
           "choose Core_BuildSetup, click Run.", vbInformation, TITLE
    Exit Sub
Fail:
    RestoreSpeed
    MsgBox "Could not finish: " & Err.Description, vbCritical, TITLE
End Sub

'================================================================
' 7. CONTROL PANEL  (builds the Macros pages)
'================================================================
' Auto-layout by display category when no button carries a placement code; the
' manual code layout when any button does ("hide" alone does not count - see
' AnyPlacementCode). Prior "Macros N" panels, and the pre-v0.4.0 single
' "Macros" tab, are cleared before either mode draws.
' The generated index is built last and is the build's landing place.
' silent:=True suppresses the "built N ..." confirmation (used when Finalize
' rebuilds the panel itself and shows its own combined message).
'
' PRIVATE deliberately, and it must stay that way. Excel's Alt+F8 dialog is the
' only list of macro names a non-technical user ever sees, and the setup guide
' tells them to pick from it - so every Public Sub callable with no arguments is
' a name that reader has to rule out. An Optional-only parameter list does not
' reliably keep a Sub out of that dialog, so Public here risked showing a FOURTH
' name beside Core_BuildSetup, Core_FinalizeSetup and Core_CompleteSetup. Worse
' than the clutter: this rebuilds the Macros pages from state that Finalize sets
' up, so running it on its own produces a panel built from a stale registry.
' Its only caller is Core_FinalizeSetup, inside this module. Nothing external
' calls it - not a button's OnAction, not tools/GateAudit.bas.
Private Sub Core_BuildControlPanel(Optional ByVal silent As Boolean = False)
    Dim buttons As Collection, codeMap As Object, layout As Object, manual As Boolean
    Dim paged As Collection, anchored As Collection, nPlaced As Long, nSkipped As Long
    Dim placedNote As String
    mLastBuildOk = False    ' reset FIRST, every call - see the flag's own declaration comment
    mLastPlacedNote = ""    ' same reason, same place - see the variable's declaration comment
    mLastNoPages = False    ' same reason, same place - see the variable's declaration comment
    InvalidateRegistry              ' a build begins here - see InvalidateRegistry
    On Error GoTo Fail
    Core_SpeedOn
    Set codeMap = ReadCodeMap(CfgSheet())   ' read ONCE: GatherButtons' hide filter and the
                                             ' mode decision below must see the same map
    manual = AnyPlacementCode(codeMap)      ' computed BEFORE GatherButtons - IsHideCoded reads
                                             ' codeMap(k) only after confirming .Exists(k), so
                                             ' this ordering no longer matters for correctness,
                                             ' but keeping it first documents the dependency
    Set buttons = GatherButtons(codeMap)    ' single funnel for both filers - also where "hide"
                                             ' is dropped, mode-independently (see GatherButtons)
    SplitAnchored buttons, codeMap, paged, anchored   ' downstream of the funnel - "hide"
                                             ' has already been applied to both halves
    Set layout = CreateObject("Scripting.Dictionary")   ' neither filer creates it
    DeletePanels                            ' remove any prior "Macros N" pages (spares
                                             ' "Macros Home" by design - see IsPanelName)
    DeleteLegacyPanel                       ' v0.4.0: the single "Macros" tab is retired in BOTH modes
    SweepPlacedButtons                      ' after the deletes: fewer sheets left to walk
    If manual Then
        BuildManualLayout paged, codeMap, layout
    Else
        BuildAutoLayout paged, layout
    End If
    RenderAllPages layout
    BuildAdminPanel                         ' Create/Finalize controls on their own hideable tab
    nPlaced = PlaceAnchoredButtons(anchored, codeMap, nSkipped)
    BuildHomePage layout, manual            ' LAST: activates the build's deliberate landing place
    mLastBuildOk = True     ' the whole pipeline finished - only set here, never in Fail:
    Core_SpeedOff
    ' Final-review fix wave, fix 1: composed UNCONDITIONALLY, outside "If Not
    ' silent". Core_FinalizeSetup - the normal user path - calls this Sub silent and
    ' prints its own message, so a note built only for the MsgBox below was thrown
    ' away with the local on exactly the path where it matters most. mLastPlacedNote
    ' carries it out; the non-silent MsgBox below still reads the local, so a direct
    ' run of Core_BuildControlPanel is unchanged and the note is never doubled (that
    ' path shows it, Finalize's path is not running). "Direct" means the VBE's
    ' Immediate window, NOT Alt+F8: Excel's Macro dialog lists only Public Subs
    ' taking no arguments, and this one's Optional silent parameter is enough to
    ' hide it.
    If nPlaced > 0 Then _
        placedNote = vbCrLf & vbCrLf & nPlaced & " button(s) placed on your own sheets."
    ' A skipped anchor usually leaves that action with NO button anywhere - not on a
    ' generated page (SplitAnchored kept it out of "paged") and not on the user's
    ' sheet. That has to be said out loud, with a repair instruction.
    '
    ' Final-review fix wave, fix 5: the branches need DIFFERENT sentences. On a
    ' missing sheet/cell there is genuinely no button anywhere and the repair is to
    ' correct the placement. On a PROTECTED sheet neither statement holds:
    ' SweepPlacedButtons failed on that sheet too, so a button an earlier build left
    ' there is still sitting on it, still wired to its old OnAction - an accepted
    ' residual, because the engine must never unprotect a user's sheet - and the
    ' repair is to unprotect, not to retype the placement.
    If nSkipped > 0 Then _
        placedNote = placedNote & vbCrLf & vbCrLf & nSkipped & " anchored button(s) could NOT " & _
                     "be placed." & vbCrLf & _
                     "- If the sheet or cell no longer exists, that action has no button " & _
                     "anywhere: fix the placement on the Setup tab and Finalize again." & vbCrLf & _
                     "- If the sheet is protected, unprotect it and Finalize again. Any button " & _
                     "an earlier build left there is still on it and may be out of date."
    mLastPlacedNote = placedNote
    ' Carried out for the same reason mLastPlacedNote is, and set BEFORE the silent
    ' check so Finalize's dialog sees it on the very path that suppresses the MsgBox
    ' below - see the flag's declaration comment.
    mLastNoPages = (layout.Count = 0)
    If Not silent Then
        ' "Built 0 macro page(s)" is true but reads as a failure, and zero pages
        ' is a real, reachable state: every button can be hidden, switched Off,
        ' or anchored onto one of the user's own tabs.
        '
        ' Says "anchored to", NOT "placed on": with nPlaced = 0 and nSkipped > 0
        ' nothing was actually placed anywhere, and "placed on one of your own
        ' sheets" would then be a flat lie in the one state the user most needs
        ' the truth. "Anchored" is true whether or not the anchor resolved, and
        ' placedNote - composed just above, and appended to both branches below -
        ' gives the count and the repair.
        If layout.Count = 0 Then
            MsgBox NO_PAGES_MSG & placedNote, vbInformation, "Setup Complete"
        Else
            ' paged.Count, NOT buttons.Count: an anchored button is not on a
            ' generated page and must not be counted against one. This is the
            ' same class of miscount the ledger records for the home page's
            ' totalButtons - one honest definition, not the nearest handy number.
            MsgBox "Built " & layout.Count & " macro page(s) with " & paged.Count & _
                   " button(s)." & placedNote, vbInformation, "Setup Complete"
        End If
    End If
    Exit Sub
Fail:
    RestoreSpeed
    ' Matches Core_FinalizeSetup's own wording for the same failure (its only
    ' caller) - there is no "re-running the build" control, and on a virgin
    ' build a failure here leaves no 'Setup Menu' tab either, so naming the real
    ' recovery (Finalize again) instead of a nonexistent one.
    MsgBox "Core_BuildControlPanel error: " & Err.Description & vbCrLf & vbCrLf & _
           "Your settings are safe. Click 'Finalize Setup Tab' again to retry the rebuild.", _
           vbExclamation, "Setup"
End Sub

' Drop the pre-v0.4.0 single "Macros" tab. Called unconditionally by
' Core_BuildControlPanel now (both modes supersede it), so this is also how an
' upgrading workbook's stale "Macros" tab is removed the first time it rebuilds
' under this version.
Private Sub DeleteLegacyPanel()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(MACRO_SHEET).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' Each item: Array(id, caption, onAction, help, category). Actions from the
' registry, chains from activated config. This is the SINGLE funnel for both, so
' the availability filter lives here and nowhere else. Admin buttons are never in
' this list - they live on the Setup Menu tab. Takes codeMap from the caller
' (Task 8) rather than reading its own copy, so Core_BuildControlPanel's mode
' decision (AnyPlacementCode) and this hide filter always see the same map.
' Fix round 3, fix 1 (BLOCKING, live path): also the single funnel for the
' "hide" placement code, and for the same reason - BuildAutoLayout never sees
' codeMap at all (it only distributes by category), so a "hide"-coded button
' would otherwise be filtered ONLY inside BuildManualLayout. Fix round 2's fix
' 5 correctly stopped a lone "hide" from flipping a workbook into manual mode
' (AnyPlacementCode ignores it) - but that means a hide-only store routes to
' the AUTO filer, and the auto filer has no "hide" concept of its own, so the
' hidden button would be drawn, catalogued, and counted. Filtering it here,
' before either filer ever sees the button, makes "hide" mode-independent,
' matching contract section 6 ("hide omits a button") without qualification.
' Feed both filers ONLY this collection - nothing assembled another way - or
' "hide" silently stops working.
Private Function GatherButtons(codeMap As Object) As Collection
    Dim out As New Collection, reg As Object, a As Object, chains As Collection, ch As Object
    Dim onAct As String, availMap As Object
    Set availMap = ReadAvailMap(CfgSheet())
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            If HasButton(availMap, CStr(a("id"))) And Not IsHideCoded(codeMap, CStr(a("id"))) Then
                onAct = "'Core_Run " & Chr(34) & CStr(a("id")) & Chr(34) & "'"
                out.Add Array(CStr(a("id")), CStr(a("caption")), onAct, _
                              CStr(a("help")), CStr(a("category")))
            End If
        Next a
    Next reg
    Set chains = ReadChains(CfgSheet())
    For Each ch In chains
        If Len(CStr(ch("id"))) > 0 And Not IsHideCoded(codeMap, CStr(ch("id"))) Then
            onAct = "'Core_Run " & Chr(34) & CStr(ch("id")) & Chr(34) & "'"
            out.Add Array(CStr(ch("id")), CStr(ch("label")), onAct, _
                          "Chain: " & StepsInline(ch("steps")), CHAIN_CAT)
        End If
    Next ch
    Set GatherButtons = out
End Function

' Partition the SINGLE funnel's output into the buttons the page pipeline draws
' and the buttons anchored onto a user's own sheet.
'
' Runs strictly DOWNSTREAM of GatherButtons, so "hide" has already been applied to
' both halves. DO NOT re-filter "hide" here, and DO NOT assemble either
' collection any other way - GatherButtons' own comment records what happened the
' last time a filer was fed something assembled separately: "hide" silently
' stopped working, on the live path, and it took a BLOCKING fix round to find.
'
' Only "paged" reaches either layout filer, so an anchored button never affects
' page counts, column caps, category packing or continuation pages.
Private Sub SplitAnchored(buttons As Collection, codeMap As Object, _
                          ByRef paged As Collection, ByRef anchored As Collection)
    Dim e As Variant, id As String, code As String
    Set paged = New Collection
    Set anchored = New Collection
    For Each e In buttons
        id = LCase(CStr(e(0)))
        code = ""
        If codeMap.Exists(id) Then code = Trim(CStr(codeMap(id)))
        If CodeKind(code) = "anchor" Then
            anchored.Add e
        Else
            paged.Add e
        End If
    Next e
End Sub

' Test .Exists BEFORE reading the item. Scripting.Dictionary's default Item
' property CREATES a missing key (with an Empty value) on read, and VBA's And
' is not short-circuiting, so "codeMap.Exists(k) And LCase(codeMap(k)) = ..."
' evaluates codeMap(k) even when .Exists(k) is False, silently inserting a
' phantom key. Harmless when codeMap was a throwaway local, but Task 8 shares
' ONE codeMap between Core_BuildControlPanel's mode decision and every
' GatherButtons/filer call that follows it, so a phantom key inserted here
' would corrupt that shared map for everything reading it afterwards. The If
' below only reads codeMap(k) once .Exists(k) is already known True.
Private Function IsHideCoded(codeMap As Object, ByVal id As String) As Boolean
    Dim k As String
    k = LCase(id)
    If Not codeMap.Exists(k) Then Exit Function
    IsHideCoded = LCase(Trim(CStr(codeMap(k)))) = "hide"
End Function

' Chain help text only - fix round 1: was raw action ids ("importexport.updateMaster"),
' which are long and unbroken and blow past a card's width in a few steps. Captions
' read as "Clear Sheets > Import Data > Update Master" for a finance user. The
' Setup-tab authoring format (StepsToTokens, "Label [id]") is untouched.
Private Function StepsInline(steps As Collection) As String
    Dim s As String, v As Variant
    For Each v In steps
        s = s & IIf(Len(s) > 0, " > ", "") & ActionCaptionFor(CStr(v))
    Next v
    StepsInline = s
End Function

' Create/Finalize controls on their own tab, so they can be hidden when not in
' use and never clutter the action pages. Reused (not re-created) on rebuild, so
' a user who hides this tab keeps it hidden. Does not activate the sheet.
Private Sub BuildAdminPanel()
    Dim ws As Worksheet, leftPos As Double, topPos As Double
    Set ws = EnsurePanelSheet(ADMIN_SHEET, False)
    With ws.Range("B2")
        .Value = "Setup Menu"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = ForgeNavy()
    End With
    ' Fix round 2, fix 1: EnsurePanelSheet's row-height reset pins row 2 at
    ' StandardHeight, which does not grow for a 16pt title, so set it literally
    ' right after the title is written (same reasoning as a page's title row).
    ' Correction (Task 9): 21pt gave a 16pt font ~1pt of headroom on paper but a
    ' reviewer's arithmetic showed that is too tight in practice; bumped to 23
    ' for real headroom - a slightly tall title row costs nothing visually.
    ws.Rows(2).RowHeight = 23
    With ws.Range("B3")
        ' Hiding stays fine - it is approved behaviour - but the invitation must
        ' say how to get back without depending on the home-page link.
        .Value = "First-time and re-configuration controls. Right-click this tab -> Hide when you're " & _
                 "done - to come back, right-click any visible tab -> Unhide, or press Alt+F8."
        .Font.Italic = True
    End With
    ws.Rows("5:14").RowHeight = 15
    leftPos = ws.Range("B5").Left
    topPos = ws.Range("B5").Top
    AddBtn ws, "btnCreateSetup", "Create Setup Tab", "Core_BuildSetup", leftPos, topPos, 220, 32
    With ws.Range(ws.Cells(5, 3), ws.Cells(6, 3))
        .Merge: .Value = "Builds a pre-filled 'Setup' tab from the installed packages."
        .Font.Size = 10: .WrapText = True: .VerticalAlignment = xlCenter
    End With
    topPos = topPos + 45
    AddBtn ws, "btnFinalizeSetup", "Finalize Setup Tab", "Core_FinalizeSetup", leftPos, topPos, 220, 32
    With ws.Range(ws.Cells(8, 3), ws.Cells(9, 3))
        .Merge: .Value = "Validates your Setup answers, activates them, and rebuilds the Macros pages."
        .Font.Size = 10: .WrapText = True: .VerticalAlignment = xlCenter
    End With
    topPos = topPos + 75
    AddBtn ws, "btnCompleteSetup", "Complete Macro Setup", "Core_CompleteSetup", leftPos, topPos, 220, 32
    With ws.Range(ws.Cells(13, 3), ws.Cells(14, 3))
        .Merge: .Value = "Removes Setup, Setup Menu, and Macros Home while keeping macro pages and settings. " & _
                         "Press Alt+F8, then run Core_BuildSetup to bring Setup back."
        .Font.Size = 10: .WrapText = True: .VerticalAlignment = xlCenter
    End With
End Sub

' The generated index - ALWAYS built, even for a single page, so the workbook has
' one shape at every scale. Three parts: a page summary, the full action catalog,
' and the legend. Built last and activated, so any rebuild lands the user here.
Private Sub BuildHomePage(layout As Object, ByVal isManual As Boolean)
    Dim ws As Worksheet, r As Long, tabKeys As Variant, ti As Long, tabNum As Long
    Dim page As Object, totalButtons As Long
    Set ws = EnsurePanelSheet(HOME_SHEET, True)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 40
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 62
    ws.Columns("E").ColumnWidth = 16

    With ws.Range("B2")
        .Value = "Macro Toolkit"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    ws.Rows(2).RowHeight = 25
    With ws.Range("B3")
        .Value = "Pick a page below, or find an action in the catalog. Every page has a  << Home  link back here."
        .Font.Italic = True
    End With

    r = 5
    ws.Cells(r, 2).Value = "PAGE"
    ws.Cells(r, 3).Value = "BUTTONS"
    ws.Cells(r, 4).Value = "CATEGORIES ON IT"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Font.Bold = True
        .Interior.Color = RGB(221, 217, 240)
    End With
    r = r + 1

    tabKeys = SortedKeys(layout)
    If UBound(tabKeys) < LBound(tabKeys) Then
        ws.Cells(r, 2).Value = "No actions are available as buttons yet."
        ws.Cells(r, 2).Font.Italic = True
        r = r + 1
    Else
        For ti = LBound(tabKeys) To UBound(tabKeys)
            tabNum = CLng(tabKeys(ti))
            Set page = layout(tabNum)
            NavLink ws.Cells(r, 2), PANEL_BASE & tabNum, PANEL_BASE & tabNum
            ws.Cells(r, 3).Value = PageButtonCount(page)
            totalButtons = totalButtons + PageButtonCount(page)
            ws.Cells(r, 4).Value = CStr(page("title"))
            ws.Cells(r, 4).Font.Color = RGB(90, 90, 90)
            r = r + 1
        Next ti
    End If

    r = r + 1
    NavLink ws.Cells(r, 2), "Setup Menu  -  first-time setup / re-configure", ADMIN_SHEET
    r = r + 2
    r = HomeCatalog(ws, r, layout)
    r = r + 1
    r = HomeLegend(ws, r, totalButtons, isManual)

    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range("A1").Select
End Sub

' Every registered action, with its full description, availability, and generated
' page when it has one. Chains retain their existing generated-button treatment.
Private Function HomeCatalog(ws As Worksheet, ByVal r As Long, layout As Object) As Long
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim cols As Object, colKeys As Variant, ci As Long, colIx As Long
    Dim items As Collection, ii As Long, it As Variant, firstRow As Long
    Dim pageById As Object, reg As Object, a As Object, availMap As Object
    Dim chains As Collection, ch As Object, actionId As String, avail As String

    Set pageById = CreateObject("Scripting.Dictionary")
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        colKeys = SortedKeys(cols)
        For ci = LBound(colKeys) To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            Set items = cols(colIx)
            For ii = 1 To items.Count
                it = items(ii)
                pageById(LCase(CStr(it(0)))) = tabNum
            Next ii
        Next ci
    Next ti

    ws.Cells(r, 2).Value = "ACTION CATALOG"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    ws.Cells(r, 2).Value = "ACTION"
    ws.Cells(r, 3).Value = "PAGE"
    ws.Cells(r, 4).Value = "WHAT IT DOES"
    ws.Cells(r, 5).Value = "AVAILABILITY"
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Font.Bold = True
        .Font.Size = 10
        .Interior.Color = RGB(240, 240, 245)
    End With
    r = r + 1
    firstRow = r

    Set availMap = ReadAvailMap(CfgSheet())
    For Each reg In BuildRegistry()
        For Each a In reg("actions")
            actionId = CStr(a("id"))
            ws.Cells(r, 2).Value = CStr(a("caption"))
            If pageById.Exists(LCase(actionId)) Then
                tabNum = CLng(pageById(LCase(actionId)))
                NavLink ws.Cells(r, 3), PANEL_BASE & tabNum, PANEL_BASE & tabNum
            End If
            With ws.Cells(r, 4)
                .Value = CStr(a("help"))
                .Font.Size = 10
                .Font.Color = RGB(90, 90, 90)
                .WrapText = True
                .VerticalAlignment = xlTop
            End With
            avail = AvailLabel(AvailFor(availMap, actionId))
            If Len(avail) = 0 Then avail = "Off"
            ws.Cells(r, 5).Value = avail
            r = r + 1
        Next a
    Next reg

    Set chains = ReadChains(CfgSheet())
    For Each ch In chains
        actionId = CStr(ch("id"))
        If Len(actionId) > 0 And pageById.Exists(LCase(actionId)) Then
            ws.Cells(r, 2).Value = CStr(ch("label"))
            tabNum = CLng(pageById(LCase(actionId)))
            NavLink ws.Cells(r, 3), PANEL_BASE & tabNum, PANEL_BASE & tabNum
            With ws.Cells(r, 4)
                .Value = "Chain: " & StepsInline(ch("steps"))
                .Font.Size = 10
                .Font.Color = RGB(90, 90, 90)
                .WrapText = True
                .VerticalAlignment = xlTop
            End With
            ws.Cells(r, 5).Value = "Button"
            r = r + 1
        End If
    Next ch

    If r > firstRow Then ws.Rows(firstRow & ":" & (r - 1)).AutoFit
    HomeCatalog = r
End Function

' The legend explains the page-code convention and Availability where a daily
' user will see them. It is mode-aware so it describes this build truthfully.
Private Function HomeLegend(ws As Worksheet, ByVal r As Long, ByVal totalButtons As Long, _
                            ByVal isManual As Boolean) As Long
    ws.Cells(r, 2).Value = "HOW THESE PAGES ARE BUILT"
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1
    If isManual Then
        r = HomeLine(ws, r, totalButtons & " button(s). Your Page codes choose where each button starts - " & _
                             "for an action, the code in the action catalog on the Setup tab; for a chain, " & _
                             "the code in the Chains block. A full page or column can still move a button to " & _
                             "a continuation one.")
        r = HomeLine(ws, r, "To go back to automatic packing, clear every Page code - with none set, categories " & _
                             "are packed onto pages for you again.")
    Else
        r = HomeLine(ws, r, totalButtons & " button(s), laid out automatically - categories packed onto pages, up to " & _
                             CARD_COLS & " per page.")
        r = HomeLine(ws, r, "To take charge of the layout, fill in 'Page code' - for an action, in the action " & _
                             "catalog on the Setup tab; for a chain, in the Chains block:")
    End If
    r = HomeLine(ws, r, "    1A = page 1, column A.    12AB = page 12, column AB.    hide = leave that button off.")
    r = HomeLine(ws, r, "As soon as ONE placement code (not counting hide) is filled in, the codes take " & _
                         "over; a blank-coded button lands in page 1's Unassigned area, or a continuation " & _
                         "page once the page has no free column slot left. hide always omits a button, in " & _
                         "either mode.")
    r = HomeLine(ws, r, "Availability on the Setup tab: 'Button' = shown here. 'Chains only' = no button, but a Chain can run it.")
    r = HomeLine(ws, r, "'Off' = no button and it will not run at all; Finalize refuses to activate a Chain that uses one.")
    HomeLegend = r
End Function

Private Function HomeLine(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Size = 10
    ws.Cells(r, 2).Font.Color = RGB(90, 90, 90)
    HomeLine = r + 1
End Function

' Frameless textbox label (ported from MacroLayout ML_AddText).
Private Function AddText(ws As Worksheet, ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, _
                         ByVal txt As String, ByVal sz As Single, ByVal bold As Boolean, ByVal colr As Long) As Shape
    Dim sh As Shape
    Set sh = ws.Shapes.AddTextbox(1, l, t, w, h)
    sh.Fill.Visible = msoFalse
    sh.Line.Visible = msoFalse
    With sh.TextFrame2.TextRange
        .Text = txt
        .Font.Size = sz
        .Font.Bold = bold
        .Font.Fill.ForeColor.RGB = colr
    End With
    Set AddText = sh
End Function

Private Function PageButtonCount(page As Object) As Long
    Dim cols As Object, k As Variant
    Set cols = page("cols")
    For Each k In cols.Keys
        PageButtonCount = PageButtonCount + cols(k).Count
    Next k
End Function

' Get-or-add a panel sheet, clear its shapes (backward loop) + cells, and format.
Private Function EnsurePanelSheet(ByVal nm As String, ByVal atFront As Boolean) As Worksheet
    Dim ws As Worksheet, si As Long
    If Core_SheetExists(nm) Then
        Set ws = ThisWorkbook.Worksheets(nm)
    ElseIf atFront Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.name = nm
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = nm
    End If
    For si = ws.Shapes.Count To 1 Step -1     ' remove old buttons + textboxes safely
        ws.Shapes(si).Delete
    Next si
    ws.Cells.Clear
    ws.Cells.UnMerge
    ' Fix round 1, fix 9: Cells.Clear should already drop hyperlinks, but that
    ' behaviour has historically been version-flaky in Excel, and a generated
    ' page accumulates a nav hyperlink per rebuild (it used to be one per
    ' catalogued button, on the withdrawn Macros Home index) - cheap insurance
    ' before a human imports this.
    ws.Hyperlinks.Delete
    ' Fix round 1, fix 3: Cells.Clear does not touch row height, so a row a
    ' PRIOR build's AutoFit made tall (a long description, a long chain
    ' recipe) stays that tall even after every value on it is cleared.
    ' Nothing downstream resets it wholesale - RenderPage only sets the button
    ' rows it draws and AutoFits only the description rows it uses - so a page
    ' that shrinks on rebuild (fewer buttons, shorter descriptions) keeps the
    ' taller rows left behind by the bigger build. Reset to the sheet's standard height here,
    ' before anything else sizes a specific row, so every subsequent
    ' RowHeight/AutoFit call in this build starts from a clean sheet.
    ws.Cells.RowHeight = ws.StandardHeight
    ws.Tab.Color = ForgeSteel()
    ws.Cells.Interior.Color = RGB(255, 255, 255)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 47
    ws.Columns("C").ColumnWidth = 85
    Set EnsurePanelSheet = ws
End Function

' The layout structure every panel builder fills and RenderPage consumes:
'   layout(tabNum)("title")        page title             (String)
'   layout(tabNum)("cols")(colIx)  buttons in that column (Collection of Arrays)
'   layout(tabNum)("hdrs")(colIx)  that column's header   (String)
' A page is created lazily by the first button filed onto it, so callers must file
' a button BEFORE setting its title or column header.
Private Sub AddToPage(layout As Object, ByVal tabNum As Long, ByVal colIx As Long, ByVal btn As Variant)
    Dim page As Object, cols As Object, items As Collection
    If Not layout.Exists(tabNum) Then
        Set page = CreateObject("Scripting.Dictionary")
        page.Add "title", PANEL_BASE & "- Page " & tabNum
        page.Add "cols", CreateObject("Scripting.Dictionary")
        page.Add "hdrs", CreateObject("Scripting.Dictionary")
        layout.Add tabNum, page
    End If
    Set page = layout(tabNum)
    Set cols = page("cols")
    If Not cols.Exists(colIx) Then cols.Add colIx, New Collection
    Set items = cols(colIx)
    items.Add btn                 ' btn = Array(id, caption, onAction, help, category)
End Sub

' No-ops when the page does not exist yet - file a button first (see AddToPage).
Private Sub SetPageTitle(layout As Object, ByVal tabNum As Long, ByVal title As String)
    If layout.Exists(tabNum) Then layout(tabNum)("title") = title
End Sub

Private Sub SetColHeader(layout As Object, ByVal tabNum As Long, ByVal colIx As Long, ByVal hdr As String)
    If layout.Exists(tabNum) Then layout(tabNum)("hdrs")(colIx) = hdr
End Sub

' B/D/F carry content, C/E are gutters. Column A stays narrow as a margin.
Private Sub SetupCardGrid(ws As Worksheet)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = CARD_W
    ws.Columns("C").ColumnWidth = GUTTER_W
    ws.Columns("D").ColumnWidth = CARD_W
    ws.Columns("E").ColumnWidth = GUTTER_W
    ws.Columns("F").ColumnWidth = CARD_W
End Sub

' Card content columns are B, D, F -> 2, 4, 6. slot is 1-based; 0 means off-grid.
Private Function CardCol(ByVal slot As Long) As Long
    If slot < 1 Or slot > CARD_COLS Then Exit Function
    CardCol = slot * 2
End Function

' Row of a card pair's button, 1-based pair index within a column. The SOLE
' place this arithmetic is written. RenderPage's loop and FitDescRows' start
' row both derive from this function (start row = CardBtnRow(1) + 1) instead
' of each independently re-deriving CARD_HDR_ROW's offset, so the two cannot
' drift apart if CARD_HDR_ROW - or this formula - ever changes (fix round 2:
' the previous pairing held only because CARD_HDR_ROW happened to be odd).
Private Function CardBtnRow(ByVal pairIx As Long) As Long
    CardBtnRow = CARD_HDR_ROW + 1 + (pairIx - 1) * 2
End Function

' Anchor a button to a cell so it lines up with the description directly beneath.
Private Sub AddBtnInCell(ws As Worksheet, ByVal nm As String, ByVal cap As String, _
                         ByVal action As String, anchor As Range)
    AddBtn ws, nm, cap, action, anchor.Left, anchor.Top + 1, anchor.Width, BTN_H
End Sub

' AutoFit each description row, then floor it. AutoFit works here ONLY because
' nothing on the card grid is merged. Floor only, no ceiling (fix round 1,
' contract 6): the floor keeps a one-line card from looking broken next to a
' three-line neighbour; a ceiling would clip, which is exactly what contract 6
' forbids for a description.
Private Sub FitDescRows(ws As Worksheet, ByVal firstDescRow As Long, ByVal lastRow As Long)
    Dim r As Long
    For r = firstDescRow To lastRow Step 2
        ws.Rows(r).AutoFit
        If ws.Rows(r).RowHeight < DESC_MIN_H Then ws.Rows(r).RowHeight = DESC_MIN_H
    Next r
End Sub

' Draw one "Macros N" page as an unmerged card grid: nav row, page title, one
' category-headed card column per layout column, each button followed by its
' description in the cell below it.
Private Function RenderPage(layout As Object, ByVal tabNum As Long, _
                            ByVal prevSheet As String, ByVal nextSheet As String) As Worksheet
    Dim panel As Worksheet, page As Object, cols As Object, hdrs As Object
    Dim colKeys As Variant, si As Long, slot As Long, colIx As Long
    Dim items As Collection, ii As Long, it As Variant
    Dim cc As Long, bRow As Long, dRow As Long, lastRow As Long, hdr As String, skipped As Long
    Set page = layout(tabNum)
    Set cols = page("cols")
    Set hdrs = page("hdrs")
    Set panel = EnsurePanelSheet(PANEL_BASE & tabNum, False)
    SetupCardGrid panel
    RenderNavRow panel, prevSheet, nextSheet
    ' fix round 1: rows 2 and 4 used to be blank chrome spacers; removed, so the
    ' title now sits directly under the nav row and the category header directly
    ' under the title (CARD_HDR_ROW moved from 5 to 3 to match).
    With panel.Cells(CARD_HDR_ROW - 1, 2)
        .Value = CStr(page("title"))
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = ForgeNavy()
    End With

    colKeys = SortedKeys(cols)
    lastRow = CARD_HDR_ROW     ' honest starting point: nothing drawn past the header yet
    slot = 0
    For si = LBound(colKeys) To UBound(colKeys)
        colIx = CLng(colKeys(si))
        slot = slot + 1
        cc = CardCol(slot)
        ' Fix round 1, fix 9: both filers now cap every page at CARD_COLS columns
        ' by construction (BuildAutoLayout's slotsUsed check; BuildManualLayout's
        ' CapManualColumns), so this branch is unreachable-by-invariant - but a
        ' silent discard here is exactly how the original defect happened, so
        ' while that invariant is fresh, count what would be skipped and surface
        ' it on the page itself rather than trust the invariant silently forever.
        If cc = 0 Then
            skipped = skipped + 1
            GoTo NextSlot        ' more columns than card slots: never draw off-grid
        End If
        Set items = cols(colIx)
        hdr = ""
        If hdrs.Exists(colIx) Then hdr = CStr(hdrs(colIx))
        If Len(hdr) > 0 Then
            With panel.Cells(CARD_HDR_ROW, cc)
                .Value = hdr
                .Font.Bold = True
                .Font.Size = 11
                .Font.Color = ForgeNavy()
                .Interior.Color = ForgeTint()
                .WrapText = True    ' fix round 1: a long category used to spill into the gutter
            End With
        End If
        For ii = 1 To items.Count
            it = items(ii)
            bRow = CardBtnRow(ii)
            dRow = bRow + 1
            panel.Rows(bRow).RowHeight = BTN_ROW_H
            AddBtnInCell panel, "b" & tabNum & "_" & colIx & "_" & ii, _
                         CStr(it(1)), CStr(it(2)), panel.Cells(bRow, cc)
            With panel.Cells(dRow, cc)
                ' Carry-over C: bounded for a chain (CardHelp), untouched for an
                ' authored action - the full recipe stays authored on the Setup
                ' tab's Chains block.
                ' Fix round 2, fix 3: gated on the button's id (it(0)), not its
                ' category (it(4)) - see CardHelp's own comment.
                .Value = CardHelp(CStr(it(3)), CStr(it(0)))
                .Font.Size = 9
                .Font.Color = RGB(110, 110, 110)
                .WrapText = True
                .VerticalAlignment = xlTop
                ' fix round 1: a faint bottom rule so a button+description pair
                ' reads as one card instead of a button floating over grey text.
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlHairline
                    .Color = ForgeRule()
                End With
            End With
            If dRow > lastRow Then lastRow = dRow
        Next ii
NextSlot:
    Next si
    If skipped > 0 Then
        With panel.Cells(CARD_HDR_ROW - 1, 4)
            .Value = "WARNING: " & skipped & " column(s) did not fit on this page and were not drawn. " & _
                     "Both filers are supposed to cap every page at " & CARD_COLS & " columns - this means " & _
                     "that invariant broke. Report it."
            .Font.Bold = True
            .Font.Color = RGB(192, 0, 0)
            .WrapText = True
        End With
    End If
    ' Fix round 2, fix 1: the title (col B) and the skipped-column warning
    ' (col D) share this one physical row, so a single RowHeight governs both.
    ' EnsurePanelSheet's reset (fix round 1, fix 3) pins every row at
    ' StandardHeight, which - unlike Excel's implicit grow-to-fit-font on a
    ' fresh/never-sized row - does NOT grow for a 15pt bold title, so it must
    ' be set explicitly here, same as the header row just below gets an
    ' explicit AutoFit. Literal, not AutoFit: a title is one line at a known
    ' font size, so a literal height is fully deterministic and stays immune
    ' to the AutoFit-under-ScreenUpdating=False risk already flagged for
    ' FitDescRows. The warning is unreachable by invariant today
    ' (both filers cap every page at CARD_COLS columns by construction) but
    ' its message text is effectively fixed-length (only skipped's digit
    ' count varies), so a generous literal height is exactly as deterministic
    ' as the title's - sized only after both possibilities are known, which is
    ' why this runs after the skipped check rather than right after the title
    ' is written.
    If skipped > 0 Then
        panel.Rows(CARD_HDR_ROW - 1).RowHeight = 130   ' room for the wrapped warning sentence
    Else
        ' Correction (Task 9): 20 left a 15pt bold title with no real headroom;
        ' bumped to 22 so it reliably clears instead of clipping.
        panel.Rows(CARD_HDR_ROW - 1).RowHeight = 22    ' clears a single 15pt bold title line
    End If
    ' Explicit AutoFit for the header row: it renders correctly on a fresh
    ' sheet only because Excel's implicit auto-height flag is still live.
    ' RenderNavRow already sets row 1's RowHeight explicitly, which clears
    ' that flag; the header row gets the same explicit treatment here so a
    ' two-line category name never depends on an implicit sheet default.
    panel.Rows(CARD_HDR_ROW).AutoFit
    ' Start row derived from CardBtnRow(1), not CARD_HDR_ROW directly - see
    ' CardBtnRow's comment for why.
    FitDescRows panel, CardBtnRow(1) + 1, lastRow
    Set RenderPage = panel
End Function

' Row 1: page-to-page navigation in a fixed position on every page - Previous /
' Next. Cells rather than shapes, because a cell hyperlink needs no shape
' bookkeeping and EnsurePanelSheet's Cells.Clear removes it cleanly on rebuild.
' ASCII arrows only - build.ps1 writes UTF-8 with no BOM on the assumption that
' the SOURCE is ASCII, so a literal arrow glyph in the .bas is an import hazard.
' Home / Previous / Next. Do NOT guard a
' NavLink target with Core_SheetExists. One was added once, it fired at RENDER
' time against a sheet built last, and on a virgin workbook's first build it drew
' plain bold text instead of a link on every single page - then self-healed on
' the second build, so it was invisible to anyone testing a rebuild. Hyperlinks.Add
' does not validate SubAddress, and the target always exists by the time the macro
' returns and the user can click anything. Write the link unconditionally.
Private Sub RenderNavRow(ws As Worksheet, ByVal prevSheet As String, ByVal nextSheet As String)
    NavLink ws.Cells(1, 2), "<< Home", HOME_SHEET
    If Len(prevSheet) > 0 Then NavLink ws.Cells(1, 4), "< Previous", prevSheet
    If Len(nextSheet) > 0 Then NavLink ws.Cells(1, 6), "Next >", nextSheet
    ws.Rows(1).RowHeight = 18
End Sub

' One in-workbook hyperlink cell. The SubAddress sheet name is always quoted,
' because "Macros 1" and "Setup Menu" both contain a space.
Private Sub NavLink(cell As Range, ByVal caption As String, ByVal sheetName As String)
    cell.Parent.Hyperlinks.Add Anchor:=cell, Address:="", _
                               SubAddress:="'" & sheetName & "'!A1", _
                               TextToDisplay:=caption
    cell.Font.Size = 10
    cell.Font.Bold = True
End Sub

' Render every page in page order, threading each one's Previous/Next sheet names
' so the nav row is drawn in the same pass.
Private Sub RenderAllPages(layout As Object)
    Dim tabKeys As Variant, ti As Long, prevNm As String, nextNm As String
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        prevNm = ""
        nextNm = ""
        If ti > LBound(tabKeys) Then prevNm = PANEL_BASE & CLng(tabKeys(ti - 1))
        If ti < UBound(tabKeys) Then nextNm = PANEL_BASE & CLng(tabKeys(ti + 1))
        RenderPage layout, CLng(tabKeys(ti)), prevNm, nextNm
    Next ti
End Sub

' Manual layout: file each button onto a "Macros N" page by its code. Blank code
' -> page 1 Unassigned column (9999); unparseable -> Unassigned (already blocked
' at Finalize). "hide" never reaches here at all (see the fix round 3, fix 1
' comment below) - GatherButtons omits a hide-coded button before either filer
' runs. Files only - rendering is RenderAllPages.
' Fix round 3, fix 1: no longer checks for "hide" here - GatherButtons (the
' live-path funnel both filers are fed from) now drops a hide-coded button
' before it ever reaches either filer, so this Sub can never see one. Removed
' rather than left as a dead branch, so nobody reads it as evidence "hide" is
' still handled per-filer and re-adds a per-filer check that can never fire.
Private Sub BuildManualLayout(buttons As Collection, codeMap As Object, layout As Object)
    Dim e As Variant, id As String, code As String
    Dim tabNum As Long, colIx As Long, colL As String, unassigned As New Collection
    For Each e In buttons
        id = LCase(CStr(e(0)))
        code = ""
        If codeMap.Exists(id) Then code = Trim(CStr(codeMap(id)))
        If Len(code) = 0 Then
            unassigned.Add e
        ElseIf ParseCode(code, tabNum, colIx, colL) Then
            AddToPage layout, tabNum, colIx, e
        Else
            unassigned.Add e            ' unparseable degrades to Unassigned
        End If
    Next e
    For Each e In unassigned
        AddToPage layout, 1, 9999, e     ' blank-coded buttons go to page 1's Unassigned column
    Next e
    ' Fix round 2, fix 6: AUTO_COL_MAX bounds a same-code stack in BuildAutoLayout
    ' but nothing bounded one here - a single "1A" on 25 buttons used to draw one
    ' column 25 deep, running off-screen with no warning, while Macros Home
    ' reported the page as fine. Must run BEFORE CapManualColumns below, so the
    ' horizontal capper sees (and, if needed, relocates) the extra columns this
    ' produces - see that Sub's own comment for why splitting first and letting
    ' the existing horizontal machinery handle any resulting overflow needs no
    ' new page-overflow logic of its own.
    SplitTallManualColumns layout
    ' Carry-over A: a user can type 1A/1B/1C/1D (or fill 3 coded columns and still
    ' land blank-coded buttons in Unassigned/9999 on page 1) - that is 4+ distinct
    ' columns on one page. RenderPage only draws CARD_COLS (3) card slots and
    ' silently drops anything past slot 3, so cap here rather than let the
    ' renderer discard what the user explicitly placed. Must run BEFORE
    ' LabelManualPages below, since it moves columns between pages and
    ' LabelManualPages needs to see the FINAL column layout to title pages and
    ' head columns correctly.
    CapManualColumns layout
    ' Fix round 1, fix 5: BuildAutoLayout has always called its own labeller
    ' (LabelAutoPages) - this filer used to leave LabelManualPages for the caller
    ' to remember separately, which made the two filers asymmetric and left a
    ' silent failure mode (skip the call and the pages keep their "Macros - Page N"
    ' placeholder titles where a real category name belongs - which back when the
    ' generated index still existed also propagated into its category column).
    ' Self-label here so Task 8 cannot miss it either way.
    LabelManualPages layout
End Sub

' Title and head a hand-built layout. Column headers keep the code-based names the
' user chose. A page is titled by category only when EVERY button on it agrees - a
' hand-built page may deliberately mix categories, and guessing a title for a
' mixed page would be worse than the generic one.
Private Sub LabelManualPages(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim cols As Object, colKeys As Variant, ci As Long, colIx As Long
    Dim items As Collection, ii As Long, it As Variant
    Dim cat As String, one As String, mixed As Boolean
    Dim origColIx As Long, contTag As String
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        cat = ""
        mixed = False
        colKeys = SortedKeys(cols)
        For ci = LBound(colKeys) To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            ' Fix round 2, fix 6: a colIx >= VSPLIT_MULT is a vertical-overflow
            ' piece SplitTallManualColumns minted for an oversized same-code
            ' stack - decode its original column straight back out of the number
            ' (see that Sub's comment): the piece may since have been relocated
            ' to a different page by CapManualColumns above, and a header set
            ' before that move would not have travelled with it.
            ' Fix round 3, fix 8: colIx also encodes WHICH piece (colIx Mod
            ' VSPLIT_MULT = the pieceIndex SplitTallManualColumns minted it
            ' with), so a stack split into 4+ pieces gets distinctly numbered
            ' headers instead of several identical "(cont.)" columns on one page.
            If colIx >= VSPLIT_MULT Then
                origColIx = colIx \ VSPLIT_MULT
                contTag = " (cont. " & (colIx Mod VSPLIT_MULT) & ")"
            Else
                origColIx = colIx
                contTag = ""
            End If
            If origColIx = 9998 Then
                SetColHeader layout, tabNum, colIx, "Setup" & contTag
            ElseIf origColIx = 9999 Then
                SetColHeader layout, tabNum, colIx, "Unassigned" & contTag
            Else
                SetColHeader layout, tabNum, colIx, "Column " & IndexToCol(origColIx) & contTag
            End If
            Set items = cols(colIx)
            For ii = 1 To items.Count
                it = items(ii)
                one = CStr(it(4))
                If Len(cat) = 0 Then
                    cat = one
                ElseIf LCase(cat) <> LCase(one) Then
                    mixed = True
                End If
            Next ii
        Next ci
        If Not mixed And Len(cat) > 0 Then SetPageTitle layout, tabNum, cat
    Next ti
End Sub

' Fix round 2, fix 6: AUTO_COL_MAX bounds a same-code stack in BuildAutoLayout;
' nothing bounded one here, so a single "1A" on (say) 25 buttons drew one column
' 25 deep, running off-screen with no warning, while Macros Home reported the
' page as fine. Split any column over AUTO_COL_MAX BEFORE CapManualColumns runs
' (BuildManualLayout calls this first): the split pieces are just more columns
' on the page by the time the existing horizontal cap-and-overflow machinery
' looks at it, so a page that gains a 4th+ column this way overflows to a
' continuation page exactly like an explicitly-coded 4th column already does -
' no new page-overflow logic needed, only more columns for the old logic to see.
'
' A split piece's colIx encodes its origin (origColIx * VSPLIT_MULT +
' pieceIndex) rather than getting a header written here, because
' CapManualColumns may relocate a piece to a different page afterwards, and a
' header set on the page it started on would not travel with it (fix round 1's
' fix 8 deliberately removed that carry-across-a-move machinery as dead code).
' LabelManualPages decodes the header straight back out of the colIx instead,
' which survives being moved because AddToPage copies the colIx unchanged.
Private Sub SplitTallManualColumns(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object, cols As Object
    Dim colKeys As Variant, ci As Long, colIx As Long, items As Collection
    Dim needCols As Long, k As Long, n As Long, taken As Long, pieceColIx As Long

    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        colKeys = SortedKeys(cols)      ' snapshot before this page's cols dict grows below
        For ci = LBound(colKeys) To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            Set items = cols(colIx)
            If items.Count <= AUTO_COL_MAX Then GoTo NextCol
            needCols = -Int(-items.Count / AUTO_COL_MAX)     ' ceiling division
            ' Guard against Long overflow for a pathological many-letter code -
            ' leave it as one tall, unsplit column rather than corrupt colIx math.
            ' (2147483647 is VBA Long's max; safe for any realistic 1-3 letter code.)
            If colIx > (2147483647 - needCols) \ VSPLIT_MULT Then GoTo NextCol
            taken = AUTO_COL_MAX                ' first AUTO_COL_MAX items stay at colIx
            For k = 2 To needCols
                pieceColIx = colIx * VSPLIT_MULT + k
                For n = 1 To AUTO_COL_MAX
                    taken = taken + 1
                    If taken > items.Count Then Exit For
                    AddToPage layout, tabNum, pieceColIx, items(taken)
                Next n
            Next k
            ' items IS cols(colIx) (Collections are reference types), so trim the
            ' overflow back out of it now that it has been copied into the pieces
            ' above - AddToPage only ever appends, it never removes.
            Do While items.Count > AUTO_COL_MAX
                items.Remove items.Count
            Loop
NextCol:
        Next ci
    Next ti
End Sub

' Carry-over A (manual side): fold any page's columns beyond CARD_COLS onto
' continuation page(s) appended after the highest page number already in use,
' CARD_COLS columns at a time, instead of letting RenderPage's slot cap (see
' CardCol) silently discard them. Nothing the user typed is lost - the moved
' column keeps its items, and the normal Previous / Next nav still reaches it;
' its header is assigned afterwards by LabelManualPages, not carried across the
' move (see that fix's comment immediately below for why). Only manual layouts
' need this: capacity-packed BuildAutoLayout never files more than CARD_COLS
' columns onto a page by construction (see BuildAutoLayout's slotsUsed check).
' Fix round 2, fix 7: the paragraph above used to claim the moved column "keeps
' ... its header (suffixed "(cont.)")" - stale the moment fix round 1's fix 8
' (immediately below) deleted the code that did that, and exactly the class of
' comment that caused fix round 1's fix 1 (a maintainer reads a comment
' describing code that is not there, and "fixes" it back in). Corrected above,
' not just noted here.
' Fix round 1, fix 8: an earlier draft carried header-preserving machinery here
' (read page("hdrs"), re-tag the moved column "(cont.)", remove it from the old
' page's hdrs) that was dead on arrival: this runs from BuildManualLayout, before
' any header is ever set, so hdrs is always empty at this point and every branch
' of that machinery was unreachable. Deleted rather than fixed-in-place - a moved
' column instead gets a plain, correct header ("Column D", not "Column D
' (cont.)") for free from LabelManualPages, which (fix 5) now always runs
' immediately after this, over the FINAL layout including whatever this moved.
Private Sub CapManualColumns(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim cols As Object, colKeys As Variant, ci As Long, colIx As Long
    Dim maxTab As Long, newTab As Long, chunkPos As Long
    Dim items As Collection, ii As Long

    tabKeys = SortedKeys(layout)
    maxTab = 0
    For ti = LBound(tabKeys) To UBound(tabKeys)
        If CLng(tabKeys(ti)) > maxTab Then maxTab = CLng(tabKeys(ti))
    Next ti

    ' Snapshot the pre-existing pages before appending continuation pages, so the
    ' loop below never revisits (or re-splits) a page it just created.
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set cols = page("cols")
        colKeys = SortedKeys(cols)
        If UBound(colKeys) < LBound(colKeys) + CARD_COLS Then GoTo NextTab   ' <= CARD_COLS cols: fine
        chunkPos = CARD_COLS                    ' forces a fresh page on the first overflow column
        For ci = LBound(colKeys) + CARD_COLS To UBound(colKeys)
            colIx = CLng(colKeys(ci))
            If chunkPos >= CARD_COLS Then
                maxTab = maxTab + 1
                newTab = maxTab
                chunkPos = 0
            End If
            chunkPos = chunkPos + 1
            Set items = cols(colIx)
            For ii = 1 To items.Count
                AddToPage layout, newTab, colIx, items(ii)
            Next ii
            cols.Remove colIx
        Next ci
NextTab:
    Next ti
End Sub

' Carry-over C: a chain's help is Core-synthesised ("Chain: " & StepsInline(...),
' GatherButtons) and with MAX_LIST steps can run to ~500 characters - roughly 14
' wrapped lines on a card's width - which AutoFit then applies to the WHOLE row,
' inflating it across all three card columns. Contract 6's no-clipping rule binds
' an action's AUTHORED help; a chain's help is engine-generated, so Core is free
' to bound it here. Cards get the first few step captions plus a "+N more"
' pointer; the full recipe stays authored on the Setup tab's Chains block, so
' the complete step list stays visible somewhere. Only chain buttons
' (category = CHAIN_CAT) are touched - authored action help passes through as-is.
' CHAIN_CARD_STEPS lives in the module's declarations section (top of file), next
' to AUTO_COL_MAX, not here - a Const after the first procedure is a compile error.
' Fix round 2, fix 3: gates on the button's id (IsChainId, "chain.*"), NOT on
' its category string. Category = CHAIN_CAT ("Chains") used to be the gate, but
' CHAIN_CAT is reserved by contract section 7, not enforced in code - Core_AddAction
' accepts any category a package declares, so a package could pass
' category:="Chains" and have its AUTHORED help decorated (and truncated) here,
' which section 6 forbids outright. Chains only ever reach GatherButtons through
' the ReadChains branch, and every chain id is minted "chain.*" (MintChainId),
' so the id namespace is the one thing an ordinary action can never spoof.
Private Function CardHelp(ByVal helpText As String, ByVal btnId As String) As String
    Dim body As String, parts() As String, n As Long, i As Long, shown As String
    Dim usedSteps As Long, candidate As String
    If Not IsChainId(btnId) Then
        CardHelp = helpText
        Exit Function
    End If
    body = helpText
    If Left$(body, 7) = "Chain: " Then body = Mid$(body, 8)
    parts = Split(body, " > ")
    n = UBound(parts) - LBound(parts) + 1
    ' Fix round 1, fix 6: BOTH caps must clear for the text to pass through as-is
    ' - a short step count with long captions still needs bounding (see the
    ' Consts' comment above).
    If n <= CHAIN_CARD_STEPS And Len(helpText) <= CHAIN_CARD_CHARS Then
        CardHelp = helpText
        Exit Function
    End If
    For i = LBound(parts) To UBound(parts)
        If usedSteps >= CHAIN_CARD_STEPS Then Exit For
        candidate = shown & IIf(Len(shown) > 0, " > ", "") & parts(i)
        ' Always keep at least one step, even a lone long one, so the card never
        ' shows zero steps and just a "+N more" tail.
        If usedSteps > 0 And Len(candidate) > CHAIN_CARD_CHARS Then Exit For
        shown = candidate
        usedSteps = usedSteps + 1
    Next i
    ' Fix round 2, fix 2: the fast-path test above can fail (Len(helpText) alone
    ' pushed it over CHAIN_CARD_CHARS because of the "Chain: " prefix) while the
    ' loop still gets through every step before either cap actually bites - e.g.
    ' 4 steps of 27 characters, or 3 of 38. usedSteps = n then, and appending a
    ' "+0 more" tail would hand back a string LONGER than the input while
    ' claiming there is more to see, when there is not. If nothing was actually
    ' left out, return the original text unchanged.
    If usedSteps >= n Then
        CardHelp = helpText
        Exit Function
    End If
    ' The pointer names the Setup tab's Chains block, not the withdrawn
    ' "Macros Home" index it used to name (v0.4.0) - a "+N more" tail that
    ' points at a sheet the build no longer creates is worse than no pointer.
    CardHelp = "Chain: " & shown & "  +" & (n - usedSteps) & " more - see the " & SETUP_SHEET & " tab"
End Function

' Auto-layout: no button carries a placement code, so distribute buttons by display
' category. A page is a CAPACITY unit, not a category unit: up to CARD_COLS
' category-headed columns per page, up to AUTO_COL_MAX buttons per column. Small
' categories share a page. A category that does not fit in the remaining columns
' starts a new page rather than straddling one - splitting a category across tabs
' destroys the grouping that putting it in one place buys. A category larger than
' a whole page continues onto further pages. Replaces the pre-v0.4.0 single-
' "Macros"-tab default (contract v1.2.0), which did not survive contact with a
' large catalog.
Private Sub BuildAutoLayout(buttons As Collection, layout As Object)
    Dim cats As Object, seen As New Collection, order As Collection
    Dim e As Variant, cat As String, ci As Long, items As Collection
    Dim needCols As Long, tabNum As Long, slotsUsed As Long
    Dim taken As Long, k As Long, n As Long, firstColOfCat As Boolean

    Set cats = CreateObject("Scripting.Dictionary")
    ' Fix round 1, fix 7: a default Dictionary is BinaryCompare, so "Reconcile"
    ' and "reconcile" from two different packages (an explicitly supported section 6
    ' scenario - two packages feeding one category) would become two adjacent
    ' columns, even though LabelAutoPages de-dupes case-insensitively and would
    ' title the page with the category named once above both. Must be set BEFORE
    ' the first Add - VBA raises an error changing CompareMode on a non-empty
    ' Dictionary.
    cats.CompareMode = vbTextCompare
    For Each e In buttons
        cat = Trim(CStr(e(4)))
        If Len(cat) = 0 Then cat = "Other"       ' defensive: Core_AddAction never emits blank
        If Not cats.Exists(cat) Then
            cats.Add cat, New Collection
            seen.Add cat                         ' first-appearance order, stable
        End If
        cats(cat).Add e
    Next e

    Set order = CategoryOrder(cats, seen)
    tabNum = 1
    slotsUsed = 0
    For ci = 1 To order.Count
        cat = CStr(order(ci))
        Set items = cats(cat)
        needCols = -Int(-items.Count / AUTO_COL_MAX)        ' ceiling division
        ' Start a fresh page rather than straddle. Contract section 6 states the
        ' no-straddle rule unconditionally - it does not carve out an exception
        ' for a category needing more columns than a page has. A category bigger
        ' than a whole page still can't avoid straddling SOME page boundary, but
        ' it should do so at a page it starts fresh on, not one it arrives at
        ' partially filled: the per-k check inside the loop below (slotsUsed >=
        ' CARD_COLS) already forces a new page once the current one fills, so
        ' removing the "needCols <= CARD_COLS" guard here just makes that
        ' fresh-page start happen up front instead of mid-category.
        If slotsUsed > 0 And slotsUsed + needCols > CARD_COLS Then
            tabNum = tabNum + 1
            slotsUsed = 0
        End If
        taken = 0
        firstColOfCat = True
        For k = 1 To needCols
            If slotsUsed >= CARD_COLS Then
                tabNum = tabNum + 1
                slotsUsed = 0
            End If
            slotsUsed = slotsUsed + 1
            For n = 1 To AUTO_COL_MAX
                taken = taken + 1
                If taken > items.Count Then Exit For
                AddToPage layout, tabNum, slotsUsed, items(taken)
            Next n
            If firstColOfCat Then
                SetColHeader layout, tabNum, slotsUsed, cat
                firstColOfCat = False
            Else
                ' Fix round 3, fix 8 (found here too, not just in the manual
                ' split): a category needing exactly CARD_COLS columns (13-18
                ' buttons) fits on one page, so k=2 and k=3 would both be
                ' unsuffixed "(cont.)" - two identical headers on the same
                ' page. k is already this category's 1-based column counter, so
                ' number the header with it instead.
                SetColHeader layout, tabNum, slotsUsed, cat & " (cont. " & k & ")"
            End If
        Next k
    Next ci
    LabelAutoPages layout
End Sub

' First-appearance order, with the reserved Chains category moved to the end: a
' composed button reads as a summary of the actions above it, so it belongs last.
Private Function CategoryOrder(cats As Object, seen As Collection) As Collection
    Dim out As New Collection, i As Long, cat As String
    Set CategoryOrder = out
    For i = 1 To seen.Count
        cat = CStr(seen(i))
        If LCase(cat) <> LCase(CHAIN_CAT) Then out.Add cat
    Next i
    For i = 1 To seen.Count
        cat = CStr(seen(i))
        If LCase(cat) = LCase(CHAIN_CAT) Then out.Add cat
    Next i
End Function

' A packed page may hold up to CARD_COLS categories, so title it by what is
' actually on it: one category -> its name; several -> the distinct names joined.
Private Sub LabelAutoPages(layout As Object)
    Dim tabKeys As Variant, ti As Long, tabNum As Long, page As Object
    Dim hdrs As Object, hk As Variant, seen As Object, title As String, one As String
    Dim contAt As Long
    tabKeys = SortedKeys(layout)
    For ti = LBound(tabKeys) To UBound(tabKeys)
        tabNum = CLng(tabKeys(ti))
        Set page = layout(tabNum)
        Set hdrs = page("hdrs")
        Set seen = CreateObject("Scripting.Dictionary")
        title = ""
        For Each hk In hdrs.Keys
            one = CStr(hdrs(hk))
            ' Fix round 3, fix 8: a "(cont.)" header is now numbered ("(cont. 2)",
            ' "(cont. 3)", ...), so a literal Replace(" (cont.)", "") no longer
            ' matches - truncate at the fixed " (cont." anchor instead, which is
            ' the same regardless of the number that follows it.
            contAt = InStr(one, " (cont.")
            If contAt > 0 Then one = Left$(one, contAt - 1)
            If Not seen.Exists(LCase(one)) Then
                seen(LCase(one)) = True
                title = title & IIf(Len(title) > 0, "   |   ", "") & one
            End If
        Next hk
        If Len(title) > 70 Then title = Left$(title, 67) & "..."
        If Len(title) > 0 Then SetPageTitle layout, tabNum, title
    Next ti
End Sub

' Remove every "Macros N" grouped page (numeric tail; the legacy "Macros" is left).
Private Sub DeletePanels()
    Dim i As Long, nm As String
    Application.DisplayAlerts = False
    On Error Resume Next
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        nm = ThisWorkbook.Worksheets(i).Name
        If IsPanelName(nm) Then ThisWorkbook.Worksheets(i).Delete
    Next i
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

' Delete every button this engine has EVER placed on a user's sheet, across the
' whole workbook, before re-placing from the CURRENT codes.
'
' Self-healing by construction: a moved cell, a cleared code, a renamed sheet,
' and an orphan left by a build that died halfway are all fixed by the same pass.
' It runs on EVERY build, including builds with zero anchors - that is exactly
' how clearing a code removes its button.
'
' DeletePanels cannot help here: it only deletes generated "Macros N" sheets, and
' an anchored button lives on a sheet the engine must otherwise never touch.
'
' Shapes are iterated BACKWARDS - the established idiom in EnsurePanelSheet -
' because deleting one shifts every later index down.
Private Sub SweepPlacedButtons()
    Dim ws As Worksheet, si As Long
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For si = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(si).Name, Len(PLACED_PREFIX)) = PLACED_PREFIX Then _
                ws.Shapes(si).Delete
        Next si
    Next ws
    On Error GoTo 0
End Sub

' Resolve an anchor's sheet part to a Worksheet (Nothing when it does not exist)
' and hand back its cell part. The single place an anchor becomes objects, so the
' Finalize checks and the placer cannot disagree about what "TabName!B4" resolves to.
Private Function AnchorSheet(ByVal code As String, ByRef cellRef As String) As Worksheet
    Dim sh As String
    cellRef = ""
    If Not ParseAnchor(code, sh, cellRef) Then Exit Function
    If Not Core_SheetExists(sh) Then Exit Function
    Set AnchorSheet = ThisWorkbook.Worksheets(sh)
End Function

' Anchor one button per placement onto a sheet the engine otherwise leaves
' COMPLETELY alone. Writes to no cell: no Clear, no UnMerge, no column width, no
' row height, no value, no format, no hyperlink. The ONLY mutation to a user's
' sheet is adding - and, via SweepPlacedButtons, removing - a shape whose name
' carries this engine's reserved prefix.
'
' This is deliberately a separate procedure, NOT a parameter on RenderPage. The
' card-grid renderer owns its sheet outright: EnsurePanelSheet does Cells.Clear,
' Cells.UnMerge, Hyperlinks.Delete, Cells.RowHeight, six column widths, tab
' colour and interior fill. A user's own sheet must never be treated that way.
'
' Returns how many buttons were actually placed, and hands back how many were
' SKIPPED via "skipped", for the completion message.
'
' A failure skips that one button rather than raising, and is reported through
' the skipped count. A build must not die on its last step because one anchor
' went stale.
'
' Both counts are needed because a skipped anchor leaves that action with NO
' button anywhere: it is CodeKind "anchor", so SplitAnchored kept it out of
' "paged" and no generated page carries it either. Silently discarding
' anchored.Count - n would make the action unreachable from the workbook surface
' with nothing on screen saying so.
'
' EVERY failure mode here is a live one, and the sweep does NOT cover them:
'   - the target sheet was renamed or deleted after Finalize -> AnchorSheet
'     returns Nothing;
'   - the cell reference no longer resolves -> ws.Range(aCell) raises;
'   - the sheet is PROTECTED -> ws.Buttons.Add raises 1004, because Sheet
'     Protect's DEFAULT includes DrawingObjects, and protected sheets are
'     routine in a multi-entity finance workbook. SweepPlacedButtons cannot
'     delete the old button on that sheet either, so the stale one stays put.
' The AddBtn call is therefore wrapped in its own On Error Resume Next. Do NOT
' re-tighten that to cover only the Range resolution: this procedure has no
' error handler of its own, so an unguarded raise inside AddBtn propagates to
' Core_BuildControlPanel's "On Error GoTo Fail" and aborts the ENTIRE rebuild -
' mLastBuildOk stays False, the home page never builds, and the user is told the
' whole build failed because one anchor pointed at a protected tab.
'
' Task 9's Finalize checks reduce how often this fires, but cannot remove it:
' Finalize validates at Finalize time, and a sheet can be renamed, deleted or
' protected at any point afterwards.
Private Function PlaceAnchoredButtons(anchored As Collection, codeMap As Object, _
                                      ByRef skipped As Long) As Long
    Dim e As Variant, id As String, code As String, aCell As String
    Dim ws As Worksheet, target As Range, n As Long
    For Each e In anchored
        id = LCase(CStr(e(0)))
        code = ""
        If codeMap.Exists(id) Then code = Trim(CStr(codeMap(id)))
        Set ws = AnchorSheet(code, aCell)
        If Not ws Is Nothing Then
            Set target = Nothing
            On Error Resume Next
            Set target = ws.Range(aCell)
            On Error GoTo 0
            If Not target Is Nothing Then
                On Error Resume Next
                AddBtn ws, PLACED_PREFIX & SafeShapeName(CStr(e(0))), _
                       CStr(e(1)), CStr(e(2)), _
                       target.Left, target.Top, ANCHOR_BTN_W, BTN_H, xlMove
                If Err.Number = 0 Then n = n + 1 Else skipped = skipped + 1
                Err.Clear
                On Error GoTo 0
            Else
                skipped = skipped + 1
            End If
        Else
            skipped = skipped + 1
        End If
    Next e
    PlaceAnchoredButtons = n
End Function

' "Macros Home" starts with PANEL_BASE ("Macros ") but has a non-numeric tail, so
' it reads False here and DeletePanels never touches it.
'
' This is load-bearing: it lets a rebuild reuse rather than recreate the
' index sheet, so EnsurePanelSheet finds it already there, clears it, and
' rewrites it in place.
'
' The generated home index depends on this carve-out, and "delete exactly the
' pages this engine generates, and nothing else a user happened to name 'Macros
' Something'" is the correct rule on its own merits.
Private Function IsPanelName(ByVal nm As String) As Boolean
    Dim tail As String
    If LCase(Left(nm, Len(PANEL_BASE))) <> LCase(PANEL_BASE) Then Exit Function
    tail = Mid(nm, Len(PANEL_BASE) + 1)
    If Len(tail) = 0 Then Exit Function
    IsPanelName = IsNumeric(tail)
End Function

' Every sheet name this engine owns and rebuilds, in ONE place so a future
' generated surface is added once rather than four times.
'
' An anchor pointing at one of these is blocked at Finalize: the engine deletes
' or clears these tabs on every build, so a button placed there would vanish -
' silently - on the very next rebuild, which is the worst possible way for a
' placement to fail.
Private Function IsEngineOwnedSheet(ByVal nm As String) As Boolean
    Select Case LCase(Trim(nm))
        Case LCase(SETUP_SHEET), LCase(ADMIN_SHEET), LCase(MACRO_SHEET), _
             LCase(HOME_SHEET), LCase(CONFIG_SHEET), LCase(PACKAGES_SHEET)
            IsEngineOwnedSheet = True
        Case Else
            IsEngineOwnedSheet = IsPanelName(Trim(nm))  ' every generated "Macros <n>"
    End Select
End Function

' The resolved ABSOLUTE address of an anchor's cell ("$B$4"), or "" when the cell
' part does not resolve on that sheet. Normalising through Range.Address is what
' makes "B4" and "$B$4" collide correctly in the duplicate-target check - two
' spellings of one cell are one target.
' Callers must have confirmed the sheet exists first; a missing sheet yields ""
' here, which the caller reports as a bad code rather than as a duplicate.
Private Function AnchorAddrKey(ByVal sheetName As String, ByVal cellRef As String) As String
    Dim ws As Worksheet, rng As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then Set rng = ws.Range(cellRef)
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    AnchorAddrKey = rng.Address(True, True)
End Function

'================================================================
' Setup-tab drawing helpers
'================================================================
Private Sub Yellow(rng As Range)
    rng.Interior.Color = RGB(255, 242, 204)
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = RGB(200, 200, 200)
End Sub

Private Function SectionHeader(ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    SectionHeader = TintRow(ws, r, txt, 11, True, False, ForgeNavy())
End Function

Private Sub PopulateReferenceLists(ws As Worksheet, ByVal startRow As Long)
    Dim wsx As Worksheet, gr As Long, hr As Long
    Dim fdr As String, fn As String
    gr = startRow
    For Each wsx In ThisWorkbook.Worksheets
        ws.Cells(gr, 7).Value = wsx.name
        ws.Cells(gr, 7).Font.Color = RGB(120, 120, 120)
        gr = gr + 1
        If gr > startRow + 250 Then Exit For
    Next wsx
    hr = startRow
    fdr = ThisWorkbook.Path
    If Right(fdr, 1) <> Application.PathSeparator Then fdr = fdr & Application.PathSeparator
    fn = Dir(fdr & "*.xls*")
    Do While Len(fn) > 0
        ws.Cells(hr, 8).Value = fn
        ws.Cells(hr, 8).Font.Color = RGB(120, 120, 120)
        hr = hr + 1
        If hr > startRow + 250 Then Exit Do
        fn = Dir
    Loop
End Sub

Private Sub AddBtn(ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal action As String, _
                   ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, _
                   Optional ByVal place As XlPlacement = xlMoveAndSize)
    ' errNum/errDesc, NOT eNum/eDesc: VBA identifiers are case-INSENSITIVE, so
    ' "eNum" is the reserved keyword "Enum" (Public Enum ... End Enum) and the
    ' line cannot compile. The VBE reports that against this declaration line,
    ' which is why it reads as a syntax error in a Dim that looks perfectly
    ' valid. Second instance of this class here after "any"; build.ps1's
    ' reserved-word gate now covers both.
    Dim btn As Button, errNum As Long, errDesc As String
    ' ALL-OR-NOTHING. Buttons.Add can succeed and a later property assignment
    ' still raise - and a half-built button is worse than no button at all: it
    ' carries Excel's default name ("Button 11"), so SweepPlacedButtons cannot
    ' find it (that matches on PLACED_PREFIX), it has no caption to identify it
    ' and no OnAction so it does nothing when clicked, and it therefore survives
    ' every later rebuild forever. Found at the v0.4.0 gate: unclickable orphan
    ' buttons accumulating on a user's own sheet, one per failed placement.
    ' Delete the partial shape, then re-raise the ORIGINAL error so the caller
    ' still counts a skip and reports it.
    On Error GoTo Abort
    Set btn = ws.Buttons.Add(l, t, w, h)
    btn.name = nm
    btn.Caption = cap
    btn.OnAction = action
    btn.Font.Size = 11
    btn.Font.Bold = True
    ' fix round 1: explicit, not relied-on-by-default. The card grid grows
    ' description rows AFTER buttons are placed, which reflows every button
    ' below the first pair - that only works because a Forms button anchors and
    ' resizes with its underlying cells, which is xlMoveAndSize. Make it explicit
    ' rather than leaving the whole card-grid layout resting on Excel's default.
    ' v0.4.0: xlMoveAndSize stays the DEFAULT, so every existing call site is
    ' unchanged. An anchored button on a user's own sheet passes xlMove instead:
    ' on a generated page the engine controls the column width and wants the
    ' button to track it, but on a user's sheet, narrowing a column must not
    ' squash their button.
    btn.Placement = place
    Exit Sub
Abort:
    errNum = Err.Number
    errDesc = Err.Description
    If Not btn Is Nothing Then
        On Error Resume Next
        btn.Delete
        On Error GoTo 0
    End If
    Err.Raise errNum, "AddBtn", errDesc
End Sub

' Map any character outside A-Za-z0-9_ to "_", so the action id "clean.dupes"
' composes into the legal shape name "SFPlaced_clean_dupes".
' A theoretical collision between two ids differing only by that substitution is
' harmless: placement is sweep-then-add and never looks a button up by name.
Private Function SafeShapeName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") _
           Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            out = out & ch
        Else
            out = out & "_"
        End If
    Next i
    SafeShapeName = out
End Function

'================================================================
' Small utilities
'================================================================
Private Function PartAt(ByRef parts() As String, ByVal idx As Long) As String
    If idx >= LBound(parts) And idx <= UBound(parts) Then PartAt = Trim(parts(idx))
End Function

Private Function IsTrue(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        IsTrue = v
    Else
        IsTrue = (UCase(Trim(CStr(v))) = "TRUE" Or Trim(CStr(v)) = "1")
    End If
End Function

' ---- page-code parsing / column math (ported from MacroLayout; Task 9 reuses) ----

' "12AB" -> tabNum 12, colIndex for "AB". Returns False on malformed input.
' HARDENED: rejects page 0, guards CLng overflow on the digit run.
Private Function ParseCode(ByVal code As String, ByRef tabNum As Long, _
                           ByRef colIndex As Long, ByRef colLetter As String) As Boolean
    Dim s As String, i As Long, ch As String, digits As String, letters As String
    tabNum = 0: colIndex = 0: colLetter = ""
    s = UCase(Trim(code))
    If Len(s) = 0 Then Exit Function
    i = 1
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch Else Exit Do
        i = i + 1
    Loop
    Do While i <= Len(s)
        ch = Mid(s, i, 1)
        If ch >= "A" And ch <= "Z" Then letters = letters & ch Else Exit Do
        i = i + 1
    Loop
    If i <= Len(s) Then Exit Function
    If Len(digits) = 0 Or Len(letters) = 0 Then Exit Function
    If Len(digits) > 7 Then Exit Function        ' overflow guard (CLng-safe page range)
    ' Fix round 3, fix 9: the digit run was capped, but the letter run never
    ' was. Two real problems, not just an absurd-input curiosity: (1) a bare-
    ' minimum 4-letter run ("AAAA" = 18279) already exceeds VSPLIT_MULT
    ' (100000)'s assumption that every real colIx from ParseCode stays well
    ' under it, so a long enough letter run would get misdecoded by
    ' LabelManualPages as a vertical-overflow split piece; (2) a sufficiently
    ' long letter run overflows the colIndex accumulation below and raises a
    ' runtime error instead of this function returning False, so a typo like
    ' this could crash Finalize's own validation pass instead of being reported
    ' as a bad code. Capping at 3 letters matches Excel's own column-letter
    ' grammar (real columns never exceed 3 letters, "XFD") and keeps every
    ' accepted colIndex at or under 18278 - safely below VSPLIT_MULT.
    If Len(letters) > 3 Then Exit Function
    tabNum = CLng(digits)
    If tabNum < 1 Then Exit Function             ' pages are 1-based; "0A" is invalid
    colLetter = letters
    For i = 1 To Len(letters)
        colIndex = colIndex * 26 + (Asc(Mid(letters, i, 1)) - Asc("A") + 1)
    Next i
    ParseCode = True
End Function

' "" | "hide" | "page" | "anchor" - the SINGLE place a placement value's FORM is
' classified. Every reader goes through this rather than testing for "!" itself.
' (The "hide" FILTER is a separate concern applied one funnel above, in
' GatherButtons via IsHideCoded - this does not own that funnel and does not
' replace it.)
'
' The "!" discriminator is safe for two reasons, and NEITHER is "Excel forbids !
' in a sheet name" - it does not. The forbidden set is : \ / ? * [ ], so a sheet
' really can be called "Sales!Q1".
'   1. A valid PAGE CODE can never contain "!": ParseCode's grammar is digits
'      then letters with a trailing-junk guard, so "1A!" is already rejected.
'      That is what makes the two forms unconfusable - even for a sheet named "1A".
'   2. ParseAnchor splits on the LAST "!" (InStrRev), so "Sales!Q1!B4" and
'      "'Sales!Q1'!B4" both resolve to sheet "Sales!Q1", cell "B4".
' Do NOT "simplify" that InStrRev to InStr - it would silently break every anchor
' onto a sheet whose name legitimately contains "!".
Private Function CodeKind(ByVal code As String) As String
    Dim s As String
    s = Trim(code)
    If Len(s) = 0 Then CodeKind = "": Exit Function
    If LCase(s) = "hide" Then CodeKind = "hide": Exit Function
    If InStr(1, s, "!") > 0 Then CodeKind = "anchor": Exit Function
    CodeKind = "page"
End Function

' "TabName!B4" / "'TabName Data'!$B$4" / "TabName Data!B4" -> sheetName "TabName Data",
' cellRef "B4". Returns False on anything malformed.
'
' The sheet part may be quoted; an embedded apostrophe is doubled inside the
' quotes, matching Excel's own spelling. The cell part is a SINGLE A1 reference.
' A colon (a multi-cell range) is rejected: this anchors one button, not a block.
'
' Deliberately a SIBLING of ParseCode, not a branch inside it. ParseCode's four
' call sites keep their page-code meaning and its hard-won hardening - the digit
' cap, the 3-letter cap, the VSPLIT_MULT guard - stays exactly as the ledger left
' it.
Private Function ParseAnchor(ByVal code As String, ByRef sheetName As String, _
                             ByRef cellRef As String) As Boolean
    Dim s As String, p As Long, sh As String, cl As String, canon As String
    sheetName = ""
    cellRef = ""
    s = Trim(code)
    p = InStrRev(s, "!")
    If p < 2 Then Exit Function              ' no "!", or nothing before it
    If p = Len(s) Then Exit Function         ' nothing after it
    sh = Left$(s, p - 1)
    cl = Mid$(s, p + 1)
    If Left$(sh, 1) = "'" Then
        If Len(sh) < 3 Then Exit Function
        If Right$(sh, 1) <> "'" Then Exit Function
        sh = Mid$(sh, 2, Len(sh) - 2)
        sh = Replace(sh, "''", "'")
    ElseIf Right$(sh, 1) = "'" And Len(sh) > 1 Then
        ' EXCEL ATE THE OPENING QUOTE. Typing  'TabName Data'!B4  into a cell stores
        ' the value  TabName Data'!B4  - a leading apostrophe is Excel's "treat this
        ' as text" prefix, consumed by the cell and stripped from .Value. It
        ' still shows in the formula bar, which is exactly what makes this so
        ' baffling to diagnose: the sheet plainly exists, the placement plainly
        ' looks right, and the engine reports NO SUCH SHEET for a name carrying
        ' an invisible trailing quote.
        ' Quoting a name that contains spaces is the natural instinct and the
        ' spelling this contract lists first, so repair it rather than refusing
        ' it. The branch above still handles a genuinely-quoted value, which is
        ' what a port whose cells do not eat apostrophes will deliver.
        ' Cost, accepted at the v0.4.0 gate: a sheet whose real name ENDS in an
        ' apostrophe can no longer be anchored by its bare name. Legal in Excel,
        ' vanishingly rare, and it fails loudly (NO SUCH SHEET at Finalize)
        ' rather than placing a button somewhere wrong.
        sh = Left$(sh, Len(sh) - 1)
        sh = Replace(sh, "''", "'")
    End If
    sh = Trim(sh)
    If Len(sh) = 0 Then Exit Function
    If Not IsA1Ref(cl, canon) Then Exit Function
    sheetName = sh
    cellRef = canon
    ParseAnchor = True
End Function

' A single A1-style cell reference: optional "$", 1-3 column letters, optional
' "$", 1-7 row digits, and nothing else. Rejects a colon (a range), a
' whole-column reference, and any trailing character.
' Row bounds beyond the digit cap are NOT checked here - a 7-digit row that
' exceeds the sheet's real last row simply fails to resolve when the caller asks
' for the Range, which is the one place that knows the platform's limit.
' On success, canonical receives the normalised form: upper-cased letters (as
' already accumulated below), no "$", and CLng-round-tripped digits so a
' leading-zero row like "0004" comes back "4" - form-valid must mean
' syntactically resolvable, not merely "matches the grammar", and a caller
' handing "B0004" to Range would find that out the hard way instead.
Private Function IsA1Ref(ByVal s As String, Optional ByRef canonical As String) As Boolean
    Dim t As String, i As Long, ch As String, letters As String, digits As String
    t = UCase(Trim(s))
    If Len(t) = 0 Then Exit Function
    i = 1
    If Mid$(t, i, 1) = "$" Then i = i + 1
    Do While i <= Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "A" And ch <= "Z" Then letters = letters & ch Else Exit Do
        i = i + 1
    Loop
    If i <= Len(t) Then
        If Mid$(t, i, 1) = "$" Then i = i + 1
    End If
    Do While i <= Len(t)
        ch = Mid$(t, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch Else Exit Do
        i = i + 1
    Loop
    If i <= Len(t) Then Exit Function                     ' trailing junk, or a ":"
    If Len(letters) = 0 Or Len(letters) > 3 Then Exit Function
    If Len(digits) = 0 Or Len(digits) > 7 Then Exit Function
    If CLng(digits) < 1 Then Exit Function
    canonical = letters & CStr(CLng(digits))
    IsA1Ref = True
End Function

Private Function IndexToCol(ByVal idx As Long) As String
    Dim n As Long, r As Long, s As String
    n = idx
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr(Asc("A") + r) & s
        n = (n - 1) \ 26
    Loop
    IndexToCol = s
End Function

Private Function SortedKeys(d As Object) As Variant
    Dim k As Variant, arr() As Long, n As Long, i As Long, j As Long, t As Long
    n = d.Count
    If n = 0 Then SortedKeys = Array(): Exit Function
    ReDim arr(0 To n - 1)
    i = 0
    For Each k In d.Keys
        arr(i) = CLng(k): i = i + 1
    Next k
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If arr(j) > arr(j + 1) Then
                t = arr(j): arr(j) = arr(j + 1): arr(j + 1) = t
            End If
        Next j
    Next i
    SortedKeys = arr
End Function
