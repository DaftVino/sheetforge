# Contributing

Two audiences use this page, and only one of them writes code.

## If you just use SheetForge in Excel

**You do not need to know anything about programming to be useful here.** The most valuable
contribution is telling us what you wanted the toolkit to do and could not work out how to do.

- **Something is broken or confusing** → [open a bug report](https://github.com/DaftVino/sheetforge/issues/new?template=bug-report.yml).
  Say what you clicked, what happened, and what you expected. "The button did nothing" is a
  perfectly good bug report. A screenshot helps more than a description of a screenshot.
- **You want a macro that does not exist** → [open a feature request](https://github.com/DaftVino/sheetforge/issues/new?template=feature-request.yml).
  Describe the job in plain words — *"I paste four exports together every Monday and delete the
  header rows"* — not a design. Working out how is our side of it.
- **A step in the [setup guide](docs/setup-guide.md) did not match what you saw** → that is a bug
  in the guide, and we want it. The guide is written for someone who has never opened the VBA
  editor, and the only way to know it works is people telling us where it did not.

Do not report a **security** concern in a public issue — see [SECURITY.md](SECURITY.md).

## If you want to change the code

Read [`CLAUDE.md`](CLAUDE.md) first. It is short, and it is the actual working agreement:
architecture, naming rules, and the non-obvious constraints. Then
[`docs/architecture.md`](docs/architecture.md) and
[`docs/adr/`](docs/adr/) for why things are the way they are.

The rules most likely to trip you up:

1. **`src/` is the source of truth.** The workbook is a build/import target
   ([ADR 0001](docs/adr/0001-source-first-vba-toolkit.md)). There is no round trip: editing a
   `.bas` does nothing in Excel until you rebuild and re-import.
2. **Build, then import — one direction.** `pwsh build/build.ps1` concatenates `src/core` plus the
   enabled packages into `dist/MacrosBundle.bas`. That build must pass before you push; it is the
   closest thing to CI here.
3. **Never depend on VBIDE trust.** Core finds packages by calling `Pkg<Name>_Manifest` through
   `Application.Run` against a maintained list — never by enumerating modules, which would need
   *Trust access to the VBA project object model*. This is not negotiable.
4. **Naming is enforced by the build.** Entry points are `Pkg<Name>_<Verb>`, manifests are
   `Pkg<Name>_Manifest`, Core's API is `Core_*`, and **every package private needs the
   `Pkg<Name>_` prefix too** — the bundle is one module, so an unprefixed private reserves that
   name against every other package.
5. **Shared helpers live only in Core.** A second copy in a package is a standards deviation.

### Verifying a change

There is no test suite and there will not be one: VBA only executes inside Excel, and Excel cannot
be driven headlessly on CI. What exists instead:

```
pwsh build/build.ps1                                   # concatenate + structural gates
node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas # parser, not a compiler
node scripts/reading-level.mjs docs/setup-guide.md      # if you touched the guide
```

**CI runs the first and the third; the second only works on your machine.** The syntax check drives
the language server inside the *VBA Pro* (`vba-lsp`) VS Code extension, so it needs that extension
installed — install it, or point `VBA_LSP_SERVER` at a server binary. It is not optional just
because CI cannot run it: run it before every push.

**None of the three resolves symbols.** A call to a helper that does not exist passes all of them
and fails only at `Debug ▸ Compile` inside Excel, so check any new call site by hand.

Then actually run it: import `dist/MacrosBundle.bas` into a scratch workbook, run
`Core_BuildSetup`, and exercise what you changed. **Say in the pull request what you ran and what
you saw** — including if you could not run it. An honest "not run in Excel" is worth more than an
implied pass.

If your change affects behaviour a user could notice, add an item to
[`docs/designs/2026-07-28-end-to-end-validation.md`](docs/designs/2026-07-28-end-to-end-validation.md).
Each item states the observable result first, then what failure looks like, then why it matters —
and an item whose expected result equals the default state proves nothing, so it does not belong.

### A note on how this repo is built

Much of SheetForge was written with AI assistance, under the review discipline in `CLAUDE.md`:
plan first, review the plan, then implement step by step, and dry-run every owner-facing checklist
against a reader who has not seen the code. The `docs/designs/` directory is that record, kept
deliberately — it is where the reasoning lives, including the parts that turned out wrong.
Contributions are welcome from humans and from people working with agents; the bar is the same
either way, and it is the review, not the author.
