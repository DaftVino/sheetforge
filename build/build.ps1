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
    # Fixture/harness packages are never part of a shipped bundle. They stay
    # excluded from the default glob rather than relying on someone remembering,
    # and are still buildable on purpose via -Packages PkgFixture.
    $excludeByDefault = @('PkgFixture.bas')
    $pkgFiles = Get-ChildItem -LiteralPath $pkgDir -Filter 'Pkg*.bas' |
                Where-Object { $excludeByDefault -notcontains $_.Name } |
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
    'end','eqv','erase','event','exit','false','for','friend','function','get','global','gosub',
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

# --- guard: every Sub/Function/Property name must be unique across all modules ---
# NOTE: Property Get/Let/Set sharing a name would false-positive here; this
# codebase uses none. Add kind to the key if that changes.
$procRe = '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)'
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
$subO = ([regex]::Matches($text, '(?m)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Sub\s')).Count
$subC = ([regex]::Matches($text, '(?m)^\s*End Sub\b')).Count
$fnO = ([regex]::Matches($text, '(?m)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?Function\s')).Count
$fnC = ([regex]::Matches($text, '(?m)^\s*End Function\b')).Count

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
$procOpen = '^\s*(?:Public\s+|Private\s+|Friend\s+|Global\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s'
$procClose = '^\s*End\s+(?:Sub|Function|Property)\b'
$moduleDecl = '^\s*(?:Public\s+|Private\s+|Global\s+)?(?:Const|Dim|Type|Enum|Declare)\b|^\s*(?:Public|Private|Global)\s+(?:WithEvents\s+)?\w+\s+As\b'
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
