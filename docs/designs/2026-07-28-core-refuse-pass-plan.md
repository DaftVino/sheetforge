# Core pass — a refusal becomes a first-class outcome

> **For agentic workers:** REQUIRED SUB-SKILL: use `superpowers:subagent-driven-development`
> to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Give Core the two things W1 had to fake — `Core_Refuse` and
`Core_IsEngineSheet` — and make a chain stop when a step refuses, so the toolkit stops
reporting a green tick for work it did not do.

> ## Amended 2026-07-28, after `bde8400` and a re-read of the packages
>
> This plan was written before the owner deferred the W1 gate. Two changes; nothing else
> in it moves, and none of the three settled decisions is re-opened.
>
> 1. **Task 5 steps 1 and 2 are withdrawn.** They add gate items to
>    `docs/designs/2026-07-28-w1-manual-gate.md` for the owner to run. The owner deferred
>    that gate on 2026-07-28 precisely to stop spending owner time on what a careful read
>    can find. **Task 5' replaces them** below: the same claims, verified statically and
>    logged for the single end-to-end validation at the end of the track.
> 2. **Task 4b is added** — `PkgImportExport`'s refusal paths never set `ctx("failed")` at
>    all, so a chain ticks them green. This is a third instance of the exact defect the
>    pass is named after, found on 2026-07-28. Scope extension, stated rather than silent;
>    reasoning in `2026-07-28-track2-completion-plan-of-record.md` §3.3.
>
> Sequencing for everything after this phase:
> `docs/designs/2026-07-28-track2-completion-plan-of-record.md`.

**Architecture:** Three Core additions and one behaviour change, then delete the package
copies that exist only because Core lacked them. `Core_Refuse` mirrors `Core_Done`: it sets
the outcome on `ctx` **and** owns the dialog, suppressed when the context is quiet.
`RunAction` carries a new `refused` flag out to `Core_RunChain`, which halts on it rather
than asking. `Core_IsEngineSheet` answers "is this sheet mine?" from Core's own constants
instead of a package's hardcoded copy of them.

**Tech stack:** VBA (Excel/Windows, one concatenated `.bas`), PowerShell for the build,
Node for the syntax check. No test framework — see *Verification*.

---

## Why this runs before W2, not after

Every item here is a parked W1 finding, and each one **compounds per wave**:

| Parked finding | Cost of waiting |
|---|---|
| `PkgClean_Refuse` and `PkgImport_Refuse` are byte-identical twins | W2 makes a third copy, W3 a fourth |
| `PkgImport_IsEngineSheet` hardcodes seven sheet names Core owns as `Private Const` | W3's `export`/`workbook` package needs the same test and will copy them again. If Core ever renames one, the stale copy silently folds the Setup sheet into a user's data |
| A refusal sets `failed`/`failMsg` but not `failNumber` | Every chain summary line reads `X #0 <reason>`, and every new wave inherits it |
| `Core_RunChain` asks Continue/Stop on a refusal | `import.combine_worksheets` refuses on **every run after the first**, so the daily chain shows a dialog, reports `4/4 ok`, and cleans the *previous* sheet — a wrong number in a finance deliverable with a tick beside it |

The last row is the reason this is not cosmetic.

---

## Decisions settled with the owner, 2026-07-28

**Inputs to the tasks, not open questions. Do not re-litigate.**

| # | Question | Decision |
|---|---|---|
| 1 | Should `Core_Refuse` own the dialog, or only set flags? | **Core owns both.** One call per refusal, suppressed when quiet, exactly like `Core_Done`. It also makes "set the flags but forget the dialog" impossible — the bug shape that produced this finding. |
| 2 | Should a refused chain step stop the chain, or keep asking? | **A refusal stops the chain.** An error still asks. A refusal means a precondition was not met, so every later step would run against data the refused step did not produce. |
| 3 | How much scope? | **`Core_Refuse` + `Core_IsEngineSheet` + the `#0` fix.** Not the remaining W1 minors — that would turn a focused pass into a grab bag. |

---

## Global constraints

- **`SCHEMA_VER` stays `"v2.0"`.** Nothing here reinterprets a stored value.
- **`Option Explicit`**; every variable declared. **Pure ASCII** in every `.bas` — the build
  fails on one non-ASCII byte. Max 25 line continuations per statement. No module-level
  declaration after the first procedure. No VBA reserved word as an identifier (`eNum`
  lowercases to `enum`).
- **Shared helpers live ONLY in Core.** Cross-package calls are forbidden. That rule is the
  whole point of this pass — do not "fix" a duplicate by calling across packages.
- **Never read `src/core/Core.bas` whole** (229KB). Grep `docs/code-map.md` for an anchor
  and read that slice.
- **`dist/` is gitignored.** Never `git add -f` the bundle.
- **The three shipped `clean` actions keep their behaviour.** `clean.dupes`, `clean.trim` and
  `clean.filldown` are not in scope; the `PkgClean` site Task 4 migrates belongs to
  `clean.standardize_formats`, a W1 action.
  **Amended:** the five `importexport` actions ARE now in scope, for their refusal *flag*
  only (Task 4b). Their wording does not change. The original constraint said they were
  untouched; that was written before the missing `ctx("failed")` was found.
- **No attribution footer, `Co-Authored-By`, or session URL** in any commit or PR body.

---

## Verification, since TDD is off

`CLAUDE.md` turns off `test-driven-development`: VBA has no headless runner and Excel cannot
be driven on CI. **Do not propose a test suite.** Every task verifies with:

1. `pwsh build/build.ps1` — must exit 0, `Core + 3 package(s)`, balanced counts.
2. `node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas` — no diagnostics.
3. **A deliberate injection** for anything meant to *fail*.
4. **A static adversarial read of the built bundle** (Task 5 Step 2), by a subagent forbidden
   from reading `src/`.

**This pass changes engine behaviour that only shows up at run time, and this phase does not
observe it.** The build cannot see a chain halting. The original version of this section said
Task 5's gate items were "the only real proof, and they must be run" — **that is withdrawn**,
because the owner deferred that gate on 2026-07-28. The claims are logged in Task 5 Step 1
and observed once, at phase 8, over the finished workbook.

**So this phase ships reviewed, not verified.** Say exactly that in the PR body. Anything
stronger is a claim the work does not support.

---

## File structure

| File | Change |
|---|---|
| `spec/contract.md` | v2.4.0. §6 gains the refusal outcome and the chain-halts rule; §10 gains the entry, including the one judgement call in classing it additive. |
| `src/core/Core.bas` | **+`Core_Refuse`**, **+`Core_IsEngineSheet`**, `RunAction` carries `refused`, `Core_RunChain` halts on it, new private `FailCode`. |
| `src/packages/PkgImport.bas` | 8 refusal sites migrate; delete `PkgImport_Refuse`, `PkgImport_IsEngineSheet`, `PkgImport_IsDigits`. |
| `src/packages/PkgClean.bas` | 1 refusal site migrates; delete `PkgClean_Refuse`. |
| `src/packages/PkgImportExport.bas` | **Task 4b:** every refusal path routes through `Core_Refuse` instead of returning silently. |
| `docs/designs/2026-07-28-deferred-validation-log.md` | **New.** The run-time-only claims, for phase 8. Replaces the withdrawn gate items. |
| `CHANGELOG.md`, `VERSION` | v0.11.0. |

---

### Task 1: Contract v2.4.0

**Owner: Opus, written directly.** The chain-halt rule is a behaviour change to the engine's
own semantics, and §10's additive/breaking call needs stating carefully rather than asserted.

**Files:** Modify `spec/contract.md` — header, §6, §10.

**Interfaces:** Produces the normative text Tasks 2–4 implement. Consumes nothing.

- [ ] **Step 1: Bump the version header**

Line 3: `**Contract version: 2.3.0** (2026-07-28)` → `**Contract version: 2.4.0** (2026-07-28)`.

- [ ] **Step 2: Add the refusal outcome to §6**

§6 currently describes an action succeeding or failing. Add a third outcome. Place it after
the existing dispatch/quiet text:

````markdown
**Three outcomes, not two (v2.4.0).** An action invocation ends in exactly one of:

| outcome | meaning | what the caller sees |
|---|---|---|
| **done** | The action did its work. | A completion message. |
| **refused** | A precondition was not met, so the action deliberately did nothing. The destination already exists; the folder is missing; nothing is configured yet. | A plain-language reason. **No error number** — there is no error. |
| **failed** | Something went wrong while the action was working. | An error number and description. |

**A refusal is not an error, and the difference is load-bearing.** A refusal is the action
working correctly: it checked, found the world not ready, and stopped having changed
nothing. An implementation MUST be able to tell the two apart, because §6's chain rule below
treats them differently, and because a refusal reported with a fabricated error number reads
to the user as a malfunction.

**Normative:** an implementation MUST expose a single call that records a refusal and
presents its reason, and that call MUST suppress its own presentation when the invocation is
quiet (a chain step), exactly as the completion sink does. Two separate steps — record the
outcome, then show a dialog — are non-conforming: an action that does one and forgets the
other reports a success it did not achieve, which is the defect this rule exists to prevent.

**Chains stop at a refusal (v2.4.0, normative).** When a chain step **refuses**, the chain
MUST stop at that step and MUST NOT run the remaining steps. It MUST report which step
refused, its reason, and how many steps did not run.

When a chain step **fails**, the prior behaviour is unchanged: the implementation asks
whether to continue.

The asymmetry is deliberate. A failure may be incidental to one step, and the user is placed
to judge. A refusal means the step's *output does not exist* — so every later step would run
against whatever was there before, silently producing a plausible, wrong result. The
reference implementation's own combine action refuses on every run after the first (its
destination sheet now exists), so under the old ask-and-continue rule a daily chain cleaned
last period's data and reported success.
````

- [ ] **Step 3: Add the §10 entry, including the judgement call**

Insert as the first versioned bullet in §10:

````markdown
- **v2.4.0 (additive):** a refusal becomes a distinct action outcome from a failure (§6); an
  implementation MUST expose a single record-and-present call for it; a chain now stops at a
  refused step instead of asking. `SCHEMA_VER` does **not** move — it stays `"v2.0"` and
  every existing configuration keeps its meaning.
  **The judgement call in classing this additive.** By the letter, a chain that used to ask
  "continue?" and now stops is a behaviour change to an existing valid configuration. It is
  classed additive because no *stored value* is reinterpreted — the chain, its steps and
  their order all still mean exactly what they meant — and because the behaviour being
  removed could not be relied on for a correct result: continuing past a step that produced
  no output ran the remainder against stale data and reported success. A platform
  implementing v2.3.0 will ask where a v2.4.0 platform stops; that is expected and is not a
  compatibility break.
````

- [ ] **Step 4: Verify**

```bash
grep -n 'Contract version: 2.4.0' spec/contract.md
grep -c 'refused' spec/contract.md
```

Then re-run the vocabulary parser check from the W1 plan's Task 1 Step 6 — §7's table must
still parse to exactly 12 domains, since `build.ps1` reads it on every build:

```bash
pwsh build/build.ps1
```

Expected: exit 0. A contract edit that breaks the table fails the build immediately.

- [ ] **Step 5: Commit**

```bash
git add spec/contract.md
git commit -m "spec: contract v2.4.0 - a refusal is a distinct outcome, and a chain stops at one"
```

---

### Task 2: Core — `Core_Refuse`, the `refused` flag, and the chain halt

**Owner: Opus, written directly.** This is Core, and it changes the dispatcher's result
contract. Every package and every future wave depends on getting it right.

**Files:** Modify `src/core/Core.bas` — near `Core_Done` (line ~1058), `RunAction` (~1069),
`Core_RunChain` (~1223).

**Interfaces:**
- Consumes: `IsQuiet`, `Core_Log`, `CtxFlag`, `CtxStr`, `CtxNum` (all existing, private).
- Produces: `Core_Refuse(ctx As Object, msg As String, [title As String])`, and
  `res("refused")` on the dictionary `RunAction` returns. Tasks 3 and 4 depend on both.

- [ ] **Step 1: Confirm what `ctx` already carries**

Before writing, read `NewCtx` (grep `docs/code-map.md` for its anchor) and confirm whether
the action id is on the context. `Core_Refuse` logs the id the way `Core_Fail` does; if the
id is not on `ctx`, log without it rather than adding a parameter — do not change
`Core_Fail`'s shape to match.

- [ ] **Step 2: Add `Core_Refuse` immediately after `Core_Done`**

It belongs beside `Core_Done` because it is the same kind of thing: a sink an action calls
instead of `MsgBox`, which knows about quiet.

```vb
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
```

- [ ] **Step 3: Carry `refused` out of `RunAction`**

In `RunAction`, initialise the flag **with the other `res` keys at the very top**, beside
`res("ok") = False` and `res("found") = False`:

```vb
    res("refused") = False
```

This placement is not cosmetic. `Core_RunChain` reads `res("refused")` unconditionally, and
reading a missing key from a `Scripting.Dictionary` raises. `RunAction` has three early
returns — action-not-installed, `StepFail`, and the normal path — so the key must exist
before any of them can be taken.

and set it in the `CtxFlag(ctx, "failed")` branch (around line 1122):

```vb
    If CtxFlag(ctx, "failed") Then         ' handler swallowed its own error
        res("ok") = False
        res("refused") = CtxFlag(ctx, "refused")
        res("errNumber") = CtxNum(ctx, "failNumber")
        res("errText") = CtxStr(ctx, "failMsg")
        res("reported") = Not quiet        ' standalone: the sink already showed it
```

Leave `StepFail` alone: an error caught there is a failure, not a refusal, and `refused`
correctly stays `False`.

- [ ] **Step 4: Stop the chain on a refusal, and fix the `#0` line**

In `Core_RunChain`'s `Else` branch (around line 1248), replace:

```vb
        Else
            summary = summary & n & ". " & res("label") & "  X  #" & res("errNumber") & _
                      " " & res("errText") & vbCrLf
            If Not AskContinue(n, CStr(res("label")), CStr(res("errText"))) Then stopped = True: Exit For
        End If
```

with:

```vb
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
```

and add the private helper beside `FirstLine`:

```vb
' "  #1004" for a real error, "" for a refusal. A refusal has no error number, and
' a summary line reading "X #0 A sheet named 'Master' already exists" tells the user
' their macro malfunctioned when it did exactly what it should.
Private Function FailCode(res As Object) As String
    On Error Resume Next
    If CLng(res("errNumber")) <> 0 Then FailCode = "  #" & res("errNumber")
End Function
```

- [ ] **Step 5: Say why the chain stopped**

The summary head currently reads `- stopped at step N`, which does not distinguish "you
chose to stop" from "it refused". Extend it so the user knows the remaining steps did not
run. In the `head` assignment (around line 1256), and using a `refusedAt` flag you set in
Step 4's `If res("refused")` branch:

```vb
    head = target("label") & IIf(stopped, "  - stopped at step " & n, "  - done") & _
           "  (" & ok & "/" & total & " ok)"
    If refusedAt Then
        head = head & vbCrLf & "Step " & n & " could not run, so the remaining " & _
               (total - n) & " step(s) were skipped."
    End If
```

Declare `refusedAt As Boolean` with the other locals at the top of `Core_RunChain` — **not**
mid-procedure, and not at module level.

- [ ] **Step 6: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 3 package(s)`, balanced counts, no diagnostics. The `Sub` count
rises by one (`Core_Refuse`) and `Function` by one (`FailCode`).

- [ ] **Step 7: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): Core_Refuse, and a chain that stops when a step refuses"
```

---

### Task 3: Core — `Core_IsEngineSheet`

**Owner: delegate to codex (`sonnet`).** Mechanical, and the constants are all in one place.

**Files:** Modify `src/core/Core.bas` — add beside `Core_SheetExists` (line ~782).

**Interfaces:**
- Consumes: the module-level constants `SETUP_SHEET`, `CONFIG_SHEET`, `CONFIG_STAGE`,
  `MACRO_SHEET`, `ADMIN_SHEET`, `PACKAGES_SHEET`, `HOME_SHEET`, `PANEL_BASE` (lines 36–44,
  124–125).
- Produces: `Core_IsEngineSheet(nm As String) As Boolean`. Task 4 deletes the package copy.

- [ ] **Step 1: Add the function**

```vb
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
        If Len(tail) > 0 Then Core_IsEngineSheet = IsAllDigits(tail)
    End If
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function
```

**A `Select Case` on `LCase(<Const>)` is legal VBA** — `Case` accepts expressions, not only
literals. If the build or a compile disagrees, fall back to an `If/ElseIf` chain rather than
re-hardcoding the strings, which is the entire defect being fixed.

- [ ] **Step 2: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, no diagnostics. Note `IsAllDigits` must not collide with any existing
name — the build's duplicate-procedure guard will say so if it does; rename to
`CoreIsAllDigits` if it fires.

- [ ] **Step 3: Commit**

```bash
git add src/core/Core.bas
git commit -m "feat(core): Core_IsEngineSheet, so a package stops hardcoding Core's sheet names"
```

---

### Task 4: Migrate both packages onto the Core API

**Owner: delegate to codex (`terra` or `sonnet`).** Nine mechanical call-site edits and four
deletions. One task, not two: the edit shape is identical in both files and splitting it
costs two reviews for one change.

**Files:** Modify `src/packages/PkgImport.bas`, `src/packages/PkgClean.bas`.

**Interfaces:** Consumes `Core_Refuse` (Task 2) and `Core_IsEngineSheet` (Task 3).

- [ ] **Step 1: Migrate the nine refusal sites**

Each site currently reads (the wording varies; the shape does not):

```vb
        PkgImport_Refuse ctx, "This folder was not found: " & folderPath
        MsgBox "This folder was not found:" & vbCrLf & vbCrLf & "   " & folderPath & vbCrLf & vbCrLf & _
               "Check 'Folder to read workbooks from' on the Setup tab, Finalize, and run " & _
               "this again." & vbCrLf & vbCrLf & "Nothing has been changed.", vbExclamation, TITLE
        Exit Sub
```

and becomes one call, keeping **the MsgBox's full wording** — not the shorter one-line
reason that was passed to `_Refuse`:

```vb
        Core_Refuse ctx, "This folder was not found:" & vbCrLf & vbCrLf & "   " & folderPath & _
               vbCrLf & vbCrLf & "Check 'Folder to read workbooks from' on the Setup tab, " & _
               "Finalize, and run this again." & vbCrLf & vbCrLf & _
               "Nothing has been changed.", TITLE
        Exit Sub
```

**Keep the long wording.** It is what the user reads standalone, and it is now also what the
chain summary quotes — `FirstLine` trims it for the summary line, so the first sentence must
stand alone. Check each message's first line does.

The eight sites in `PkgImport.bas`: no data sheets to combine · workbook never saved ·
folder not found · no workbooks found · and the four inside `PkgImport_DestUsable` (name
empty, over 31 characters, illegal character, sheet already exists). The one site in
`PkgClean.bas`: no column formats configured.

- [ ] **Step 2: Delete the four now-dead privates**

- `PkgImport_Refuse` (`PkgImport.bas:475`)
- `PkgClean_Refuse` (`PkgClean.bas:594`)
- `PkgImport_IsEngineSheet` and `PkgImport_IsDigits`

Replace the one call to `PkgImport_IsEngineSheet` in `PkgImport_CombineWorksheets` with
`Core_IsEngineSheet`. **Move the explanatory comments that live on the deleted functions
into Core** only if Task 3 did not already carry them — it did, so simply delete. Do not
leave an orphaned comment block behind.

- [ ] **Step 3: Confirm nothing still references the deleted names**

```bash
grep -n "PkgImport_Refuse\|PkgClean_Refuse\|PkgImport_IsEngineSheet\|PkgImport_IsDigits" src/packages/*.bas
```

Expected: no output.

- [ ] **Step 4: Build, syntax-check, and confirm the counts moved the right way**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 3 package(s)`, no diagnostics. Net `Sub`/`Function` counts fall by
four from Task 3's build (two `Sub`s and two `Function`s deleted).

- [ ] **Step 5: Commit**

```bash
git add src/packages/PkgImport.bas src/packages/PkgClean.bas
git commit -m "refactor: move refusal and engine-sheet knowledge out of the packages into Core"
```

---

### Task 4b: `PkgImportExport` stops ticking green for work it did not do

**Owner: delegate to codex (`sonnet`).** Same edit shape as Task 4.

**Added 2026-07-28.** A scope extension over the settled decision 3, stated rather than
silent. Decision 3 ruled out absorbing *the remaining W1 minors*; this is not one. It is a
third instance of the defect this pass is named after, found while inventorying the packages,
and it is the only one left after Task 4.

**The defect:** `PkgImportExport`'s refusal paths — `PkgImportExport_Clear` at
`PkgImportExport.bas:96`, `PkgImportExport_UpdateMaster` at `:224`, and their siblings — show
a `MsgBox` and `Exit Sub` **without ever setting `ctx("failed")`**. `RunAction` decides a step
succeeded by the absence of that flag (`Core.bas:1122`), so inside a chain each of these is
reported as a green tick beside a step that did nothing, and the steps after it run against
whatever was already on the sheet. Leaving it would mean this pass cures the disease in two
packages and ships it in the third.

**Files:** Modify `src/packages/PkgImportExport.bas`.

**Interfaces:** Consumes `Core_Refuse` (Task 2).

- [ ] **Step 1: Find every refusal site**

```bash
grep -n 'MsgBox' src/packages/PkgImportExport.bas
```

A site qualifies when **all three** are true: it shows a `MsgBox`, it then `Exit Sub`s, and
it has done no work. A `MsgBox` that reports a *completed* run is `Core_Done`'s job and is
already correct — **do not touch those**. A `MsgBox` inside `Core_Fail`'s path is an error,
not a refusal, and is also already correct.

- [ ] **Step 2: Migrate each one**

```vb
' was:
        MsgBox "<the full existing wording>", vbExclamation, TITLE
        Exit Sub
' becomes:
        Core_Refuse ctx, "<the same full existing wording>", TITLE
        Exit Sub
```

**Keep every message's wording exactly as it is.** These are shipped strings the owner has
read; this task changes what the engine records, not what the user sees standalone. The one
new consequence is that the first line now also appears in a chain summary, so **check each
message's first line stands alone as a reason** — `FirstLine` trims it for the summary.

- [ ] **Step 3: Confirm no refusal path still returns silently**

Build an explicit inventory of exit sites and walk it. A count cannot do this job — a
`grep -c` after context flags collapses everything to one integer, and a fixed context window
cannot associate an `Exit Sub` with a sink call that spans four continued lines above it.

```bash
grep -n 'Exit Sub\|^Public Sub\|^Private Sub\|Core_Refuse\|Core_Done\|Core_Fail' \
     src/packages/PkgImportExport.bas
```

That prints one line per procedure boundary, per sink call, and per exit, **in file order**.
Read it top to bottom: for each `Exit Sub`, the nearest preceding sink must be on the same
side of the last procedure boundary. An `Exit Sub` with no sink between it and the procedure
header is a silent return, and is the defect.

The `Exit Sub` immediately before a `Fail:` label is the procedure's normal end and needs no
sink — the handler below it calls `Core_Fail`.

- [ ] **Step 4: Build and syntax-check**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
```

Expected: exit 0, `Core + 3 package(s)`, no diagnostics. Counts are unchanged — this task
adds and deletes no procedure.

- [ ] **Step 5: Commit**

```bash
git add src/packages/PkgImportExport.bas
git commit -m "fix(importexport): a refusal must not be reported to a chain as a success"
```

---

### Task 5: Gate the new behaviour, and refresh the map

**Owner: Opus.** **Steps 1 and 2 are the amended ones** — see the banner at the top of this
file. They no longer write gate items for the owner.

**Files:** Create `docs/designs/2026-07-28-deferred-validation-log.md`; modify
`CHANGELOG.md`, `VERSION`, `docs/code-map.md`.

- [ ] **Step 1: Log the run-time-only claims, do not gate them**

The build cannot see a chain halt. But the owner deferred the W1 gate on 2026-07-28 to stop
spending owner time on what a careful read can find, so these do **not** become a checklist to
work through now. They go into a running log that phase 8 assembles into one end-to-end
validation over the finished workbook.

Create `docs/designs/2026-07-28-deferred-validation-log.md` with this preamble:

````markdown
# Deferred validation log

Run-time-only claims, accumulated per phase, for the **single** end-to-end validation at the
end of track 2. **This is not a checklist to work through now.** The W1 gate was deferred on
2026-07-28 (`bde8400`) because the answers are available statically; anything a careful read
or a build guard can settle is settled there instead, and only what genuinely needs Excel
lands here.

Each entry states the observable result first, then what a failure looks like, then why —
`CLAUDE.md` §7. **An entry whose expected result equals the default state proves nothing**,
so each pairs something that must happen with something that must not.
````

then append this phase's entries:

- **A chain stops at a refused step, and says so.** Run a 4-step chain beginning with
  `import.combine_worksheets` a *second* time without deleting `Master`. Expected: the
  refusal appears in the summary with no error number; the summary says the remaining 3 steps
  were skipped; the count reads `0/4`; and — the load-bearing part — **`Master` still holds
  exactly what run 1 produced**, because the clean steps did not touch it. *Failure looks
  like:* `4/4 ok`, or `Master` visibly cleaned again. *Why:* the exact wrong-number-with-a-tick
  case this pass exists to remove.
- **A refusal inside a chain shows no dialog of its own.** During the run above, the only
  dialog is the final chain summary. *Failure:* a modal refusal box mid-chain.
- **A refusal standalone still shows its dialog.** Click the combine's own button with
  `Master` present. Expected: the full refusal message, and no chain summary. *Why:* the pair
  for the item above — together they prove suppression is conditional, not broken. Either
  alone passes on a build that broke the other.
- **A step that FAILS still asks.** Point `k_import_folder` at a folder holding a zero-byte
  `.xlsx` and run a chain starting with `import.combine_workbooks`. Expected: Continue/Stop.
  *Failure:* the chain stops without asking. *Why:* the asymmetry is settled decision 2; if
  failures stop too, the distinction was never implemented.
- **A user sheet named `Macros Data` is still combined, and `Macros 1` is still excluded.**
  *Why:* `Core_IsEngineSheet` replaced the package's hardcoded copy; this is the only thing
  that proves the numeric-tail rule survived the move.
- **`Select Case LCase(<Const>)` compiles.** Excel `Debug ▸ Compile VBAProject` is clean.
  **Marked as an assumption** (`CLAUDE.md` §7): it is believed legal VBA and is exercised
  nowhere else in this codebase. Task 3 states the `If/ElseIf` fallback.

- [ ] **Step 2: Static verification instead of a blind dry-run**

What a checklist would have caught here, the bundle can be read for. Dispatch a subagent given
**only** `dist/MacrosBundle.bas` and this step, **forbidden from reading `src/`** — the
blindness is the mechanism. Ask it to answer, quoting bundle line numbers:

1. Does every `Core_Refuse` call site pass a message whose **first line stands alone** as a
   reason? (`FirstLine` trims it for the chain summary.)
2. Is there any path that sets `ctx("failed")` without also setting `ctx("refused")` or going
   through `Core_Fail`? Name it.
3. Does `RunAction` initialise `res("refused")` **before** every one of its early returns?
   Reading a missing key from a `Scripting.Dictionary` raises.
4. Can any summary line still render `#0`? Trace `FailCode` for `errNumber = 0`.
5. Does `Core_Refuse` suppress its dialog on exactly the same condition `Core_Done` does?

**Fix what it finds before opening the PR.** Its limit is known and real: it reads VBA, it
does not execute it. Anything it cannot settle stays in the Step 1 log.

- [ ] **Step 3: Version, changelog, code map**

`VERSION` → `0.11.0`. A `## [0.11.0]` entry in the house style — prose about what changes for
the user, led by the chain no longer reporting success for work it did not do. Then, **last**:

```bash
node scripts/code-map.mjs
```

- [ ] **Step 4: Final gates**

```bash
pwsh build/build.ps1 && node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas
grep -c 'SCHEMA_VER *As String = "v2.0"' dist/MacrosBundle.bas
git status --short
```

Expected: exit 0; no diagnostics; `1`; a tree clean but for the untracked `.vscode/`.
Confirm the code map's stamp equals `git log -1 --format=%h -- '*.bas' '*.cls' '*.frm' '*.mjs' '*.js' '*.ps1'`.

- [ ] **Step 5: `/codex` review the diff, fix any `[P1]`, then one PR against `main`**

Do not stack it on the W1 branch. If #19 has merged, branch from `main`; if it has not,
branch from `main` anyway and expect a small conflict in `PkgImport.bas` — the W1 branch is
the one that should land first.

---

## Unknowns and risks

- **This pass cannot be fully verified by the build.** The chain halt shows up only at run
  time. Task 5's static read narrows it; what it cannot settle ships on the strength of the
  code review and waits for phase 8. **That is a deliberate, owner-endorsed trade** — the
  alternative was owner time this phase was told not to spend — but it means the halt is
  *reviewed*, not *observed*, until then. Do not describe it as verified in the PR.
- **`Select Case LCase(<Const>)`** is believed legal VBA but is not exercised anywhere in
  this codebase today. Task 3 states the fallback. Logged as an assumption in Task 5 Step 1.
- **W1's PR #19 merged on 2026-07-28** (`6b68ebc`). `origin/main` carries W1, so this phase
  branches cleanly from `main` and the `PkgImport.bas` conflict this section used to warn
  about cannot occur.
- **`Core_Refuse`'s dialog title defaults to "Cannot run this yet".** Every W1 site passes
  its own `TITLE`, so the default is only for future callers. If it reads badly on the Setup
  surface, it is one string.
- **Not in scope, still parked:** the build gate's string-literal blind spot, `build.ps1`'s
  guard-scope inconsistency, `Core_SheetExists` not seeing chart sheets, and the two packages
  disagreeing about whether private helpers carry a package prefix.
