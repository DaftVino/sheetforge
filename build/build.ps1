#Requires -Version 5.1
<#
.SYNOPSIS
  Build dist/MacrosBundle.bas from src/core/Core.bas + enabled src/packages/Pkg*.bas.

.DESCRIPTION
  The delivery model is "modular authoring, monolithic delivery": Core and each
  package are separate .bas during development, but a deployment imports ONE .bas.
  A single imported .bas is a single VBA standard module, so this script:

    1. FAILS if any candidate source file (Core.bas + every Pkg*.bas on disk,
       not only the -Packages subset this run selected) contains a byte
       outside ASCII - the VBE imports a .bas as ANSI, so a stray non-ASCII
       byte is silently mis-decoded on import.
    2. Concatenates Core + every Pkg*.bas into one module body.
    3. Strips each file's own `Attribute VB_Name` and `Option Explicit` and
       writes exactly one of each at the top (module name = -ModuleName).
    4. FAILS if two source files declare the same Sub/Function/Property name -
       a single module needs globally-unique Public AND Private names.
    5. Sanity-checks Sub/Function block balance in the assembled content.
    6. FAILS on a module-level declaration (Const/Dim/Type/Enum/Declare/module
       var) that appears after the first procedure - VBA requires the whole
       declarations section to precede all procedures, and such a stray
       declaration is a compile error in Excel (this script does not compile VBA).

  Every check above runs on the assembled content BEFORE dist/MacrosBundle.bas
  is written, so a failing build never leaves a broken (or stale-looking, but
  actually freshly-overwritten) artifact on disk - only ever the last
  successful build's output, or nothing at all.

  NOTE: these are structural checks, NOT a VBA compiler. They cannot catch every
  compile error (undefined vars, type mismatches). The final gate is still a
  manual import + Debug > Compile VBAProject in Excel.

  Packages are discovered by the Pkg*.bas naming convention; non-package modules
  in src/packages are excluded. Which packages are actually ACTIVE at run time
  is a separate concern - Core's candidate list and the optional _Packages
  sheet gate that; a bundled package can ship disabled.

.PARAMETER Packages
  Optional explicit, ordered list of package module names (with or without .bas)
  to include instead of every Pkg*.bas. Lets a deployment ship a subset.

.EXAMPLE
  pwsh build/build.ps1
  pwsh build/build.ps1 -Packages PkgImportExport
#>
[CmdletBinding()]
param(
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    [string]$Out = $null,
    [string[]]$Packages = $null,
    [string]$ModuleName = 'MacrosBundle'
)

# `-Packages A,B,C` only auto-splits into an array when PowerShell's OWN parser
# evaluates the call (e.g. from inside another .ps1, or dot/&-invoked in a live
# session). `pwsh <script> -Packages A,B,C` from any OTHER shell - a bash
# prompt, or even a fresh `pwsh.exe` launched from a Windows shortcut - hands
# the child process raw OS argv, so "A,B,C" arrives as ONE string bound to the
# first array element; there is no invocation shell where the comma form is
# guaranteed to work without this. Split it ourselves instead of trusting the
# caller's shell: safe under this repo's Pkg<Name>.bas PascalCase naming rule,
# since no real package name can itself contain a comma. Trim each piece too -
# "-Packages PkgA, PkgB" (a space after the comma, a very natural way to type
# a list) would otherwise leave a leading space baked into the file name
# ("Package not found: ... PkgB.bas"), about the least diagnosable form that
# error can take.
if ($Packages) {
    $Packages = $Packages | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
}

$ErrorActionPreference = 'Stop'
$nl = "`r`n"

$corePath = Join-Path $Root 'src/core/Core.bas'
$pkgDir = Join-Path $Root 'src/packages'
if (-not $Out) { $Out = Join-Path $Root 'dist/MacrosBundle.bas' }

if (-not (Test-Path -LiteralPath $corePath)) { throw "Core not found: $corePath" }

# --- resolve the ordered source list: Core first, then packages ---
if ($Packages) {
    $pkgFiles = foreach ($p in $Packages) {
        $file = if ($p -match '\.bas$') { $p } else { "$p.bas" }
        $full = Join-Path $pkgDir $file
        if (-not (Test-Path -LiteralPath $full)) { throw "Package not found: $full" }
        (Get-Item -LiteralPath $full).FullName
    }
} else {
    $pkgFiles = Get-ChildItem -LiteralPath $pkgDir -Filter 'Pkg*.bas' |
                Sort-Object Name | ForEach-Object { $_.FullName }
}
$sources = @($corePath) + @($pkgFiles)

# --- guard: every candidate source file must be pure ASCII ---
# Fix round 3, fixes 5/6/7. The VBE imports a .bas as ANSI/the system codepage,
# not UTF-8, so any non-ASCII byte - a smart quote or em dash pasted from a
# doc, or (as caught here once) a stray section-sign in a comment - gets
# silently mis-decoded on import (a section-sign becomes "Â§").
#
# Runs BEFORE assembly, on every Pkg*.bas discovered on disk PLUS Core.bas -
# NOT only the files -Packages selected for this run (fix 7): "content is
# ASCII" is a repo-wide invariant a module must satisfy before it ships in ANY
# build, so `pwsh build/build.ps1 -Packages PkgImportExport` must still catch a
# bad byte sitting in PkgClean.bas even though this particular build excludes
# it - that module ships the day someone bundles it, and the whole point of a
# gate is to catch it before that day, not on it.
#
# Reports every match on a line, not just the first (fix 6), and - because
# this scan happens before $Out is touched at all - a failure here can never
# leave a poisoned dist/MacrosBundle.bas on disk (fix 5): the artifact is
# either absent (first build) or whatever the last SUCCESSFUL build left
# (unchanged), never a freshly-written file containing the bad byte.
$asciiCandidates = @($corePath) + @(
    Get-ChildItem -LiteralPath $pkgDir -Filter 'Pkg*.bas' | Sort-Object Name | ForEach-Object { $_.FullName }
)
$asciiProblems = New-Object System.Collections.Generic.List[string]
foreach ($f in $asciiCandidates) {
    $leaf = Split-Path $f -Leaf
    $srcLines = Get-Content -LiteralPath $f
    for ($i = 0; $i -lt $srcLines.Count; $i++) {
        foreach ($m in [regex]::Matches($srcLines[$i], '[^\x00-\x7F]')) {
            $ch = $m.Value
            $asciiProblems.Add("non-ASCII character in ${leaf}:$($i + 1), column $($m.Index + 1): '$ch' (U+$('{0:X4}' -f [int][char]$ch))")
        }
    }
}
if ($asciiProblems.Count -gt 0) {
    Write-Host "BUILD FAILED - source validation:" -ForegroundColor Red
    $asciiProblems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}

# --- guard: no VBA reserved word used as a Dim/parameter identifier ---
# Added after `Dim reg As Object, any As Boolean` shipped in a v0.4.0 bundle and
# failed the owner's Debug > Compile on first import. `Any` is reserved (it
# appears in `Declare ... As Any`), and the VBE reports the error against the
# ENCLOSING procedure's declaration line rather than the Dim - so the signature
# looks like the culprit and the real cause is a line further down.
#
# Nothing else here can see this class of defect: the duplicate-name and
# block-balance checks below are structural, scripts/vba-syntax-check.mjs parses
# without flagging reserved-word identifiers, and it reads as valid VBA to a
# human. Excel's own compiler was the only detector, which in this repo means a
# human's manual import - the most expensive place to find anything.
#
# The list is deliberately CONSERVATIVE and EMPIRICAL. It holds only words VBA
# hard-rejects as an identifier, so a false positive would have to be code
# nobody would write (`Dim to As Long`). It deliberately OMITS words that are
# reserved only contextually and demonstrably compile as identifiers in this
# codebase - `name` (Core_FindCol's parameter) and `base` (MintChainId), both
# shipped since v0.2.0/v0.3.0. Add a word here when Excel rejects one, not on
# suspicion: over-broad, this gate fails a build that would have worked.
$vbaReserved = @(
    'any','as','byref','byval','call','case','const','declare','dim','do','each','else','elseif',
    'end','enum','eqv','erase','event','exit','false','for','friend','function','get','global','gosub',
    'goto','if','imp','implements','in','is','let','like','loop','me','mod','new','next','not',
    'nothing','null','on','option','optional','or','paramarray','preserve','private','property',
    'public','put','raiseevent','redim','rem','resume','return','select','set','static','step',
    'stop','sub','then','to','true','typeof','until','wend','while','with','withevents','xor'
)
$resWords = @{}; $vbaReserved | ForEach-Object { $resWords[$_] = $true }
$reservedProblems = New-Object System.Collections.Generic.List[string]
foreach ($f in $asciiCandidates) {
    $leaf = Split-Path $f -Leaf
    $srcLines = Get-Content -LiteralPath $f
    for ($i = 0; $i -lt $srcLines.Count; $i++) {
        $ln = $srcLines[$i]
        if ($ln -match "^\s*'") { continue }          # comment line
        $names = @()
        # Dim / Static declarations: take the identifier before each `As`
        if ($ln -match '^\s*(?:Dim|Static)\s+(.+)$') {
            foreach ($part in ($matches[1] -split ',')) {
                if ($part -match '^\s*([A-Za-z_]\w*)') { $names += $matches[1] }
            }
        }
        # parameter lists on a Sub/Function declaration
        if ($ln -match '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+\w+\s*\((.*)$') {
            foreach ($part in ($matches[1] -split ',')) {
                if ($part -match '(?:ByVal\s+|ByRef\s+)?([A-Za-z_]\w*)\s+As\b') { $names += $matches[1] }
            }
        }
        foreach ($nm in $names) {
            if ($resWords[$nm.ToLower()]) {
                $reservedProblems.Add("VBA reserved word '$nm' used as an identifier in ${leaf}:$($i + 1) - Excel will refuse to compile this. Rename it. Line: $($ln.Trim())")
            }
        }
    }
}
if ($reservedProblems.Count -gt 0) {
    Write-Host "BUILD FAILED - source validation:" -ForegroundColor Red
    $reservedProblems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}

# --- guard: action ids - vocabulary, shape, uniqueness, resolvable entry point ---
# Ids are immutable, two-segment <domain>.<operation>, and there is deliberately NO
# alias table (spec/contract.md section 7) - so a wrong id is unfixable and has to be
# caught before it ships, not after. Nothing else here can see it: the id is a string
# literal, valid VBA to every other check in this file and to vba-syntax-check.mjs.
#
# The vocabulary is READ FROM THE CONTRACT, not duplicated here. Section 7 says the
# contract owns that table; a hardcoded copy would make this file a second owner, and
# the two would drift the first time a wave added a domain. The parse is anchored on an
# HTML-comment marker rather than on the heading text, so re-wording section 7 cannot
# silently disarm this.
#
# It FAILS CLOSED. An unreadable, malformed or empty table is a build failure, never a
# fallback to some built-in list: a gate that quietly disables itself when its input goes
# missing is worse than no gate, because the build still goes green.
#
# CASE SENSITIVITY - the one thing to get right here. PowerShell's -match and -ne are
# case-INSENSITIVE by default, which silently defeats every lowercase rule below. Two
# different rules apply and the distinction is deliberate:
#   * VBA SOURCE SYNTAX is matched case-INSENSITIVELY (-match), because VBA itself is:
#     `public sub Foo(ctx as object)` is a valid handler and must be found.
#   * ACTION ID VALUES and CONTRACT TABLE TEXT are matched case-SENSITIVELY (-cmatch,
#     -cne), because the id rule IS a case rule. Verified: with plain -match,
#     "import.combineWorksheets" passes the lowercase-only shape check, which would make
#     contract section 7's "no new id is ever camelCase" unenforceable - the single rule
#     this gate exists to enforce. Do not "simplify" the c-operators away.
$contractPath = Join-Path $Root 'spec/contract.md'
if (-not (Test-Path -LiteralPath $contractPath)) {
    Write-Host "BUILD FAILED - action id validation:" -ForegroundColor Red
    Write-Host "  contract not found: $contractPath" -ForegroundColor Red
    Write-Host "  The action-id domain vocabulary is read from that file; without it no id can be checked." -ForegroundColor Red
    exit 1
}
$contractLines = Get-Content -LiteralPath $contractPath
$vocabMarker = '<!-- action-domain-vocabulary -->'
$vocabProblems = New-Object System.Collections.Generic.List[string]

# EXACTLY one marker. Two markers (a copied section, a half-finished edit) would silently
# make the first one win, and which table is authoritative would depend on file order.
$markerLines = @()
for ($i = 0; $i -lt $contractLines.Count; $i++) {
    if ($contractLines[$i].Trim() -ceq $vocabMarker) { $markerLines += ($i + 1) }
}
if ($markerLines.Count -ne 1) {
    Write-Host "BUILD FAILED - action id validation:" -ForegroundColor Red
    if ($markerLines.Count -eq 0) {
        Write-Host "  marker '$vocabMarker' not found in spec/contract.md." -ForegroundColor Red
        Write-Host "  Section 7 must carry that marker on its own line above the domain table (a blank line between the two is fine)." -ForegroundColor Red
    } else {
        Write-Host "  marker '$vocabMarker' appears $($markerLines.Count) times in spec/contract.md (lines $($markerLines -join ', ')). It must appear exactly once." -ForegroundColor Red
    }
    exit 1
}
$markerAt = $markerLines[0] - 1

# Walk the table. EVERY line between the marker and the terminating blank must be
# recognised - header, separator, or a well-formed data row. An unrecognised row is a
# BUILD FAILURE, not a skip: silently parsing 11 of 12 rows leaves the build green while
# the 12th domain quietly stops being enforced, which is the exact failure this gate
# exists to prevent.
#
# ONLY a genuinely blank line ends the table. This previously broke on any non-pipe
# line once the header had been seen, which made a stray prose line mid-table a SILENT
# truncation: verified by inserting one after the `validate` row, which cut the
# vocabulary to three domains and then failed the build claiming the shipped
# `importexport` ids had an unknown domain - the wrong error entirely, pointing at
# innocent source. Worse, the same line placed after the last USED domain truncates the
# vocabulary with the build still green. Anything else after the header is malformed and
# is reported as such.
# Ordinal (case-sensitive) dictionary, NOT @{} or [ordered]@{}. PowerShell hashtable
# keys are case-INSENSITIVE, so `$domainStatus.ContainsKey('Import')` would be true for
# a stored 'import' - quietly accepting an id whose domain is mis-cased, which is the
# same camelCase hole the -c operators close everywhere else in this block.
$domainStatus = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
$sawHeader = $false
$sawSeparator = $false
for ($i = $markerAt + 1; $i -lt $contractLines.Count; $i++) {
    $ln = $contractLines[$i]
    $lineNo = $i + 1
    if ($ln -notmatch '^\s*\|') {
        if (-not $sawHeader) { continue }                # blank line between marker and table
        if ([string]::IsNullOrWhiteSpace($ln)) { break } # the terminating blank: table ends here
        $vocabProblems.Add("contract.md:${lineNo}: a non-table line inside the domain table. The table runs from the header row to the first BLANK line, so nothing else may sit between them - move this line below the table (leaving a blank line after the last row). Found: $($ln.Trim())")
        break
    }
    if (-not $sawHeader) {
        if ($ln.Trim() -cne '| domain | status | meaning |') {
            $vocabProblems.Add("contract.md:${lineNo}: expected the header row '| domain | status | meaning |', found: $($ln.Trim())")
            break
        }
        $sawHeader = $true
        continue
    }
    if (-not $sawSeparator) {
        if ($ln.Trim() -cne '|---|---|---|') {
            $vocabProblems.Add("contract.md:${lineNo}: expected the separator row '|---|---|---|', found: $($ln.Trim())")
            break
        }
        $sawSeparator = $true
        continue
    }
    # The trailing [^|]* is what rejects a stray pipe inside the meaning cell: it
    # corrupts the rendered table, so it must not parse cleanly here either.
    if ($ln -cmatch '^\s*\|\s*`([a-z][a-z0-9_]*)`\s*\|\s*(active|reserved|grandfathered)\s*\|[^|]*\|\s*$') {
        $dom = $matches[1]
        if ($domainStatus.ContainsKey($dom)) {
            $vocabProblems.Add("contract.md:${lineNo}: domain '$dom' is listed more than once.")
        } else {
            $domainStatus[$dom] = $matches[2]
        }
    } else {
        $vocabProblems.Add("contract.md:${lineNo}: unrecognised vocabulary row. Expected | ``<domain>`` | active/reserved/grandfathered | <meaning with no pipe> |  --  found: $($ln.Trim())")
    }
}
if (-not $sawHeader -or -not $sawSeparator) {
    $vocabProblems.Add("contract.md: the domain table under '$vocabMarker' has no header and separator row.")
}
if ($domainStatus.Count -eq 0) {
    $vocabProblems.Add("contract.md: the domain table under '$vocabMarker' parsed to zero domains.")
}
if ($vocabProblems.Count -gt 0) {
    Write-Host "BUILD FAILED - action id vocabulary (spec/contract.md section 7):" -ForegroundColor Red
    $vocabProblems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}

# The eight ids shipped before the shape rules existed. Listed literally in contract
# section 7 as permanent exceptions, never renamed and never aliased - two are camelCase
# and would fail the shape check below. This list does NOT exempt them from the
# duplicate or entry-point checks.
$legacyIds = @(
    'clean.dupes', 'clean.trim', 'clean.filldown',
    'importexport.clear', 'importexport.import', 'importexport.updateMaster',
    'importexport.export', 'importexport.suggestTypes'
)
# Ordinal, for the same reason as $domainStatus: a case-insensitive set makes
# $legacySet.ContainsKey('Clean.Dupes') true, which skips the shape check on a
# mis-cased id and then misattributes the resulting duplicate error to the real,
# permanent declaration - pointing a future maintainer at breaking an immutable id.
$legacySet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$legacyIds | ForEach-Object { [void]$legacySet.Add($_) }

# Entry points that Core_Run can actually dispatch. Core_Run reaches a handler through
# Application.Run, which cannot call a Private or Friend procedure - so "a Sub of that
# name exists somewhere" is NOT the question. The question is whether a PUBLIC Sub taking
# a single ctx argument exists. A Private Sub with the right name would sail through a
# name-only check and fail at the button, which is the most expensive place to find it.
# (In VBA an unqualified `Sub Foo` is Public, so both spellings count.)
#
# IgnoreCase is passed EXPLICITLY. [regex]::Match is case-SENSITIVE in .NET, unlike
# PowerShell's -match operator - so without this, `public sub Foo(ctx as object)` is not
# found and this guard reports "no dispatchable handler" for a handler that is really
# there. VBA is case-insensitive, so matching its syntax case-sensitively is simply wrong.
#
# Every other structural guard in this file was made case-insensitive in the same commit,
# and that was not optional. Fixing only this one opened a hole: a procedure written
# `public sub Foo(ctx as object)` ... `end sub` - consistently lowercase - was counted by
# NEITHER the Sub-open nor the End-Sub pattern, so it balanced, and it was simultaneously
# invisible to the duplicate-name and module-declaration guards. It shipped with zero
# structural validation, having previously been stopped (confusingly, but stopped) right
# here. VBA is case-insensitive; every pattern in this file that matches VBA syntax now
# carries (?i) or IgnoreCase, and any new one must too.
#
# $dispatchable itself stays a case-INSENSITIVE hashtable on purpose, because VBA resolves
# procedure names case-insensitively and so does Application.Run - an entry point declared
# `PkgFoo_Bar` and referenced as `pkgfoo_bar` really does dispatch.
$vbaIC = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
$dispatchable = @{}
foreach ($f in $asciiCandidates) {
    foreach ($line in (Get-Content -LiteralPath $f)) {
        $m = [regex]::Match($line, '^\s*(?:Public\s+)?(?:Static\s+)?Sub\s+([A-Za-z_]\w*)\s*\(\s*(?:ByRef\s+|ByVal\s+)?\w+\s+As\s+Object\s*\)', $vbaIC)
        if ($m.Success) { $dispatchable[$m.Groups[1].Value] = Split-Path $f -Leaf }
    }
}

# Truncate a VBA line at its first comment apostrophe that is OUTSIDE a string literal,
# and return the code portion. VBA has no escape character: a quote inside a string
# literal is written by DOUBLING it, so toggling the state on every `"` gets it right -
# the doubled pair toggles off then straight back on, leaving the walk inside the string.
#
# Needed because the Core_AddAction detector below tests the whole line rather than only
# its start, and Core.bas mentions the name in a TRAILING comment. Without this the gate
# would fail the build on a comment, and - the other way round - a commented-out
# declaration would be validated as if it were a real one.
function Get-VbaCodePortion([string]$line) {
    $inStr = $false
    for ($j = 0; $j -lt $line.Length; $j++) {
        $ch = $line[$j]
        if ($ch -eq '"') { $inStr = -not $inStr }
        elseif ($ch -eq "'" -and -not $inStr) { return $line.Substring(0, $j) }
    }
    return $line
}

# Core_AddAction reg, "<id>", "<caption>", "<entryPoint>", "<help>" [, "<category>"]
#
# The trailing `\s*,` is load-bearing: without it `"Handler" & suffix,` matches as far as
# the closing quote and the gate validates "Handler" as though it were the whole entry
# point, checking a name the running code never registers. Requiring the delimiter makes
# a built-up entry point fail to parse, which the unparseable-call check below then
# reports. Every real call has a help argument after the entry point, so the comma is
# always there.
$addActionRe = 'Core_AddAction\s+\w+\s*,\s*"([^"]*)"\s*,\s*"[^"]*"\s*,\s*"([^"]*)"\s*,'
# Core_AddAction's OWN declaration in Core.bas contains the name in code, not as a call.
# The detector below matches anywhere on the line, so the declaration needs an explicit
# carve-out - (?i) because VBA is case-insensitive, like every other syntax pattern here.
$addActionDeclRe = '(?i)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+Core_AddAction\b'
$idShapeRe = '^[a-z][a-z0-9_]*\.[a-z][a-z0-9_]*$'
$allowed = ($domainStatus.Keys | Where-Object { $domainStatus[$_] -ne 'reserved' } | Sort-Object) -join ', '
# Ordinal: two ids differing only in case are two DIFFERENT ids by section 7's rules,
# and each must be judged on its own rather than one masking the other.
$idSeen = [System.Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
$idProblems = New-Object System.Collections.Generic.List[string]
foreach ($f in $asciiCandidates) {
    $leaf = Split-Path $f -Leaf
    $srcLines = Get-Content -LiteralPath $f
    for ($i = 0; $i -lt $srcLines.Count; $i++) {
        # Judge the CODE only. A whole-line comment collapses to "" and falls out below;
        # a trailing comment can neither trip the detector nor be read as a declaration.
        $code = Get-VbaCodePortion $srcLines[$i]
        if ($code -notmatch '\S') { continue }                     # blank, or all comment
        $m = [regex]::Match($code, $addActionRe, $vbaIC)
        if (-not $m.Success) {
            # A CALL to Core_AddAction that does not parse is a failure, not a skip.
            # Otherwise a wave that wraps its arguments differently - a line continuation
            # before the entry point, `Call Core_AddAction(...)`, a named argument -
            # registers an action this gate never looks at, and the build goes green
            # having checked nothing. Fail loudly and make the author either reformat the
            # call or extend this regex deliberately.
            #
            # Matched ANYWHERE in the code portion, not anchored at statement position.
            # The old `^\s*(?:Call\s+)?Core_AddAction\b` anchor was a bypass, verified:
            # `If True Then Core_AddAction reg, _` with the arguments on a continuation
            # line registered `tidy.bogus` and the build exited 0. The two things the
            # anchor was really protecting against - Core.bas's own declaration and its
            # trailing-comment mention - are handled precisely instead, by
            # $addActionDeclRe and by Get-VbaCodePortion.
            if ($code -match '(?i)Core_AddAction' -and $code -notmatch $addActionDeclRe) {
                $idProblems.Add("${leaf}:$($i + 1): a Core_AddAction call this gate cannot parse. Put the whole call at the start of its own line, with the id, caption and entry point as three literal strings followed by a comma (a continuation may follow that comma). Line: $($srcLines[$i].Trim())")
            }
            continue
        }
        $id = $m.Groups[1].Value
        $entry = $m.Groups[2].Value
        $where = "${leaf}:$($i + 1)"

        if ($idSeen.ContainsKey($id)) {
            $idProblems.Add("duplicate action id '$id' at $where - already declared at $($idSeen[$id]). Ids are unique across every package.")
        } else {
            $idSeen[$id] = $where
        }

        if (-not $legacySet.Contains($id)) {
            # -cnotmatch, NOT -notmatch. See the case-sensitivity note above.
            if ($id -cnotmatch $idShapeRe) {
                $idProblems.Add("malformed action id '$id' at $where - must be exactly two segments, <domain>.<operation>, each starting with a lowercase letter and continuing with lowercase letters/digits/underscores. Never camelCase.")
            }
        }

        $domain = if ($id -match '^([^.]+)\.') { $matches[1] } else { $id }
        if (-not $domainStatus.ContainsKey($domain)) {
            $idProblems.Add("unknown domain '$domain' in action id '$id' at $where - not in the vocabulary table in spec/contract.md section 7. Allowed: $allowed.")
        } elseif ($domainStatus[$domain] -eq 'reserved') {
            $idProblems.Add("reserved domain '$domain' in action id '$id' at $where - '$domain.*' is reserved for the engine in spec/contract.md section 7 and no package action may use it.")
        } elseif ($domainStatus[$domain] -eq 'grandfathered' -and -not $legacySet.Contains($id)) {
            # 'grandfathered' means "these exact shipped ids are permanent", NOT "this
            # domain is open for business". Without this branch a wave could add
            # importexport.some_new_id and pass, which is precisely what section 7's
            # "no new id may be added to it" forbids.
            $idProblems.Add("new action id '$id' at $where uses the grandfathered domain '$domain'. That domain is closed: only its shipped ids are permitted, and no new id may be added to it (spec/contract.md section 7). Pick an active domain.")
        }

        if ($entry -eq '') {
            $idProblems.Add("action '$id' at $where declares an empty entry point.")
        } elseif (-not $dispatchable.ContainsKey($entry)) {
            $idProblems.Add("action '$id' at $where names entry point '$entry', for which no source declares a dispatchable handler. Core_Run calls it through Application.Run, so it must be a PUBLIC (or unqualified) Sub taking one Object argument: Public Sub $entry(ctx As Object). A Private or Friend Sub, or a different signature, cannot be dispatched.")
        }
    }
}
if ($idProblems.Count -gt 0) {
    Write-Host "BUILD FAILED - action id validation:" -ForegroundColor Red
    $idProblems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}

# --- guard: every Sub/Function/Property name must be unique across all modules ---
# NOTE: Property Get/Let/Set sharing a name would false-positive here; this
# codebase uses none. Add kind to the key if that changes.
$procRe = '(?i)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)'
$seen = @{}
$dupes = New-Object System.Collections.Generic.List[string]
foreach ($src in $sources) {
    $leaf = Split-Path $src -Leaf
    foreach ($line in (Get-Content -LiteralPath $src)) {
        $m = [regex]::Match($line, $procRe)
        if ($m.Success) {
            $name = $m.Groups[1].Value
            if ($seen.ContainsKey($name)) {
                $dupes.Add("  $name  (in $($seen[$name]) and $leaf)")
            } else {
                $seen[$name] = $leaf
            }
        }
    }
}
if ($dupes.Count -gt 0) {
    Write-Host "BUILD FAILED - duplicate procedure name(s); a single-module bundle needs unique names:" -ForegroundColor Red
    $dupes | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    exit 1
}

# --- assemble the bundle ---
$sb = [System.Text.StringBuilder]::new()
[void]$sb.Append("Attribute VB_Name = ""$ModuleName""$nl")
[void]$sb.Append("Option Explicit$nl$nl")
[void]$sb.Append("'============================================================================$nl")
[void]$sb.Append("'  $ModuleName - Core + enabled packages, concatenated into ONE module.$nl")
[void]$sb.Append("'  GENERATED by build/build.ps1 - do not edit. Source of truth:$nl")
foreach ($src in $sources) {
    $rel = $src.Substring($Root.Length).TrimStart('\', '/') -replace '\\', '/'
    [void]$sb.Append("'    $rel$nl")
}
[void]$sb.Append("'  Import this ONE .bas, then run Core_BuildSetup (Alt+F8) to create the$nl")
[void]$sb.Append("'  Setup tab. Core_BuildControlPanel is NOT in the Alt+F8 list - it takes an$nl")
[void]$sb.Append("'  optional argument, and Excel hides any macro that takes arguments. To run$nl")
[void]$sb.Append("'  it directly: Alt+F11, Ctrl+G, type the name, Enter.$nl")
[void]$sb.Append("'============================================================================$nl$nl")

foreach ($src in $sources) {
    $leaf = Split-Path $src -Leaf
    [void]$sb.Append("'########################## $leaf ##########################$nl")
    $droppedAttr = $false
    foreach ($line in (Get-Content -LiteralPath $src)) {
        if (-not $droppedAttr -and $line -match '^\s*Attribute\s+VB_Name\s*=') { $droppedAttr = $true; continue }
        if ($line -match '^\s*Option\s+Explicit\s*$') { continue }
        [void]$sb.Append($line + $nl)
    }
    [void]$sb.Append($nl)
}

# --- validate the assembled content BEFORE writing anything (fix round 3, fix 5) ---
# Previously wrote $Out first and validated the file it had just written - so a
# failing build left a freshly-timestamped dist/MacrosBundle.bas on disk,
# containing whatever was wrong, for someone to import instead of reading the
# red text below. Validating the in-memory StringBuilder content first means a
# failure here writes nothing: $Out is left exactly as the last SUCCESSFUL
# build produced it (or absent, on a first build) - never overwritten with a
# broken artifact.
$text = $sb.ToString()
$attr = ([regex]::Matches($text, '(?m)^Attribute VB_Name')).Count
$opt = ([regex]::Matches($text, '(?m)^Option Explicit')).Count
$subO = ([regex]::Matches($text, '(?im)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Sub\s')).Count
$subC = ([regex]::Matches($text, '(?im)^\s*End Sub\b')).Count
$fnO = ([regex]::Matches($text, '(?im)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Function\s')).Count
$fnC = ([regex]::Matches($text, '(?im)^\s*End Function\b')).Count

$problems = New-Object System.Collections.Generic.List[string]
if ($attr -ne 1) { $problems.Add("expected 1 Attribute VB_Name, found $attr") }
if ($opt -ne 1) { $problems.Add("expected 1 Option Explicit, found $opt") }
if ($subO -ne $subC) { $problems.Add("Sub/End Sub imbalance: $subO / $subC") }
if ($fnO -ne $fnC) { $problems.Add("Function/End Function imbalance: $fnO / $fnC") }

# NOTE: the assembled $text is not re-scanned for ASCII here - the guard above
# already scanned every byte of every candidate source file that could
# possibly end up in $text, before assembly ever ran, and assembly only adds
# hardcoded-ASCII boilerplate (plus -ModuleName, which Attribute VB_Name/
# Core_BuildControlPanel's own naming conventions keep to identifier
# characters in practice) - re-scanning here would just repeat the same
# findings under a different line number.

# --- guard: no module-level declaration after the first procedure ---
# VBA requires the declarations section (Const/Dim/Type/Enum/Declare/module vars)
# to precede ALL procedures. A module-level declaration between or after procedures
# is a compile error in Excel ("variable not defined" on first use). This build
# does not compile VBA, so catch that structural error here.
$declLines = [regex]::Split($text, "\r?\n")
$inProc = $false; $sawProc = $false
$procOpen = '(?i)^\s*(?:Public\s+|Private\s+|Friend\s+|Global\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s'
$procClose = '(?i)^\s*End\s+(?:Sub|Function|Property)\b'
$moduleDecl = '(?i)^\s*(?:Public\s+|Private\s+|Global\s+)?(?:Const|Dim|Type|Enum|Declare)\b|^\s*(?:Public|Private|Global)\s+(?:WithEvents\s+)?\w+\s+As\b'
for ($i = 0; $i -lt $declLines.Count; $i++) {
    $ln = $declLines[$i]
    if ($ln -match $procOpen) { $inProc = $true; $sawProc = $true; continue }
    if ($ln -match $procClose) { $inProc = $false; continue }
    if ($sawProc -and -not $inProc -and $ln -match $moduleDecl) {
        $problems.Add("module-level declaration after a procedure (VBA compile error) at line $($i + 1): $($ln.Trim())")
    }
}

if ($problems.Count -gt 0) {
    Write-Host "BUILD FAILED - output validation:" -ForegroundColor Red
    $problems | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    exit 1
}

# --- write (UTF-8 no BOM; content asserted ASCII above, so import-safe) ---
# Only reached once every check above has passed - see the fix round 3, fix 5
# comment above the balance checks for why the write moved here.
$outDir = Split-Path $Out -Parent
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
[System.IO.File]::WriteAllText($Out, $text, (New-Object System.Text.UTF8Encoding($false)))

$lines = ($text -split "`n").Count
$pkgNames = ($pkgFiles | ForEach-Object { Split-Path $_ -Leaf }) -join ', '
Write-Host "Built $Out" -ForegroundColor Green
Write-Host "  Core + $($pkgFiles.Count) package(s): $pkgNames"
Write-Host "  $lines lines | Sub $subO/$subC | Function $fnO/$fnC | unique procedures: $($seen.Count)"
Write-Host "  Import this ONE .bas, then run Core_BuildSetup (Alt+F8)."
exit 0
