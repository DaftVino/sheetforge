# Security

SheetForge ships as VBA source you import into your own Excel workbook and run there. That is
worth being plain about: **you are pasting code into a file that probably holds real numbers**, so
this page says what the code can reach, what it will never do, and how to tell us if something
looks wrong.

## What you are running

One `.bas` file. No installer, no add-in, no service, no network calls, no telemetry. Nothing
phones home, and nothing here has a server side to be breached.

**The source is the shipped artifact.** `dist/MacrosBundle.bas` is a plain concatenation of the
files under `src/` — nothing minified, obfuscated, encoded, or compiled. Every line the toolkit
runs is a line you can read in this repo, and you can rebuild the exact file yourself with
`pwsh build/build.ps1` and diff it against the release asset.

## What it can touch

Enabled actions can, on your machine and with your permissions:

- **Read and write files** in the folders you configure on the Setup tab — imports, exports,
  backups, snapshots, PDFs.
- **Delete files**, in exactly one place: snapshot pruning keeps the most recent N snapshots and
  deletes older ones. It only deletes files matching the name shape it wrote itself, in the
  snapshot folder it owns.
- **Move files**, when archiving a generated output.
- **Change your workbook** — add and remove sheets, write cells, apply formats, insert rows,
  protect and unprotect sheets.
- **Open an Outlook draft** with a file attached, if you enable the email action. **It never
  sends.** It writes the attachment to disk first, opens a draft for you to review, and stops.

It does **not** need, ask for, or work around *Trust access to the VBA project object model*. That
setting stays off. The engine discovers its packages by calling known procedure names, never by
enumerating your VBA project — a deliberate architectural constraint, not an accident.

Macros must be enabled for any of this to run, which is Excel's decision point, not ours. If you
do not trust the file, read it first: that is why it ships as source.

## Reporting something

**For a security concern, do not open a public issue.** Use
[GitHub's private vulnerability reporting](https://github.com/DaftVino/sheetforge/security/advisories/new)
on this repository, which is visible only to the maintainers.

Worth reporting privately:

- an action that writes, moves, or deletes a file outside the folders it was configured with
- the email action sending anything, rather than opening a draft
- an action that destroys data with no confirmation and no way back
- anything that reaches the network

Ordinary bugs, wrong numbers, and confusing wording belong in
[Issues](https://github.com/DaftVino/sheetforge/issues) — those are not security reports and are
much more useful in the open.

## Supported versions

The latest release on the [Releases page](https://github.com/DaftVino/sheetforge/releases/latest)
is the only supported version. Fixes go into the next release; there are no backports. Upgrading
is re-importing one file, and your saved answers survive it.

## What we do not claim

This is a small toolkit maintained by a working finance team, not audited software.

**Its verification is honest but limited.** Every release passes a build gate, a VBA parser, and a
reading-level gate in CI, and the maintainer runs a written checklist by hand in Excel. What does
not exist is an automated test suite — VBA cannot be executed headlessly, so there is no CI job
anywhere that runs a single line of this code. The checklist and the manual pass are the
substitute, and they are recorded in each pull request. Read
`docs/designs/2026-07-28-end-to-end-validation.md` if you want to see exactly what is checked and
what is still unproven.

**Back up your workbook before running an action for the first time.** The toolkit has a backup
action for precisely this reason, and every destructive action asks first — but a file you care
about deserves a copy regardless.
