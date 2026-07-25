#!/usr/bin/env node
// Builds docs/code-map.md: a symbol index with line anchors, so an agent can
// grep the map and read one slice instead of a whole file.
// Usage: node scripts/code-map.mjs [--out <path>] [--min-bytes <n>]
//
// Why this lives here and not in daftplate: the shared scripts/code-map.mjs
// indexes .js/.ts/.html only. This repo's source of truth is VBA (.bas), which
// the shared scanner cannot see at all, so it would emit an empty map for the
// one file that most needs indexing (src/core/Core.bas, ~78KB / ~125 procs).
// Output format is deliberately identical to the shared tool's, so this can be
// deleted the day .bas support lands upstream. See
// ~/.daftplate/outbox/2026-07-24-code-map-vba-support.md.
//
// There is no parser here and there will not be one - this repo has no
// dependencies. The scanner is a line-oriented regex pass, and its contract is
// asymmetric: a missed symbol costs one grep, a mislocated symbol sends a
// reader to the wrong slice of a 78KB file. When a pattern is ambiguous, it
// does not match. `Declare` statements and multi-name `Dim a, b` lines are
// known, accepted misses.
import { mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { dirname, extname, join, relative, sep } from 'node:path';

// ---------------------------------------------------------------- VBA scanner

// An unlabelled band: `'=========`. Its name comes from the NEXT comment line.
// Tested before the labelled rule below, and that order is load-bearing: a bare
// run of `=` also satisfies the labelled pattern (rule / one `=` / rule), which
// named every band in Core.bas `=` on the first run.
const VBA_BANNER = /^[ \t]*'[ \t]*[-=*#_]{3,}[ \t]*\r?$/;
// A labelled rule carrying its name inline: `' ---- sheet names ----`. The
// label must hold something that is not itself rule punctuation.
const VBA_RULE = /^[ \t]*'[ \t]*[-=*#]{3,}[ \t]*([^-=*#\s][^\r\n]*?)[ \t]*[-=*#]{3,}[ \t]*\r?$/;
const VBA_COMMENT_TEXT = /^[ \t]*'[ \t]*(.*?)[ \t\r]*$/;

// `Public Sub Foo(`, `Private Function Bar(`, `Friend Property Get Baz(`.
// `End Sub` / `Exit Function` cannot match: the keyword is not at line start.
const VBA_PROC =
  /^[ \t]*(?:(Public|Private|Friend)[ \t]+)?(?:Static[ \t]+)?(Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+([A-Za-z_]\w*)/i;

// Module-level declarations, recognised ONLY above the first procedure - that
// is what makes them module-level. Anything matching this inside a proc body is
// a local and is not a navigation target.
const VBA_DECL =
  /^[ \t]*(Public|Private|Global|Dim)[ \t]+(?:WithEvents[ \t]+)?(Const[ \t]+)?([A-Za-z_]\w*)/i;
const VBA_BLOCK = /^[ \t]*(?:(?:Public|Private)[ \t]+)?(Type|Enum)[ \t]+([A-Za-z_]\w*)/i;

export function scanVba(text) {
  const entries = [];
  const lines = text.split(/\r?\n/);
  let inProcs = false;

  for (let i = 0; i < lines.length; i += 1) {
    const raw = lines[i];

    if (VBA_BANNER.test(raw)) {
      for (let j = i + 1; j <= i + 2 && j < lines.length; j += 1) {
        if (VBA_BANNER.test(lines[j])) continue;
        const name = (VBA_COMMENT_TEXT.exec(lines[j]) ?? [])[1];
        if (name) entries.push({ line: i + 1, kind: 'section', name });
        break;
      }
      continue;
    }
    const ruled = VBA_RULE.exec(raw);
    if (ruled) {
      entries.push({ line: i + 1, kind: 'section', name: ruled[1] });
      continue;
    }
    if (/^[ \t]*'/.test(raw)) continue;

    const proc = VBA_PROC.exec(raw);
    if (proc) {
      inProcs = true;
      const scope = (proc[1] ?? 'Public').toLowerCase();
      const type = proc[2].replace(/[ \t]+/g, ' ').toLowerCase();
      entries.push({ line: i + 1, kind: `${scope} ${type}`, name: proc[3], call: true });
      continue;
    }

    const block = VBA_BLOCK.exec(raw);
    if (block) {
      entries.push({ line: i + 1, kind: block[1].toLowerCase(), name: block[2] });
      continue;
    }

    if (inProcs) continue;
    const decl = VBA_DECL.exec(raw);
    if (decl) {
      entries.push({ line: i + 1, kind: decl[2] ? 'const' : 'module var', name: decl[3] });
    }
  }
  return entries;
}

// ------------------------------------------------- JS / PowerShell scanners

const JS_DECLARATIONS = [
  { kind: 'function', re: /^[ \t]*(?:export[ \t]+)?(?:default[ \t]+)?(?:async[ \t]+)?function[ \t*]+([A-Za-z_$][\w$]*)/ },
  { kind: 'class', re: /^[ \t]*(?:export[ \t]+)?(?:default[ \t]+)?class[ \t]+([A-Za-z_$][\w$]*)/ },
  {
    kind: 'function',
    re: /^[ \t]*(?:export[ \t]+)?(?:const|let|var)[ \t]+([A-Za-z_$][\w$]*)[ \t]*=[ \t]*(?:async[ \t]+)?(?:function\b|\([^)]*\)[ \t]*=>|[A-Za-z_$][\w$]*[ \t]*=>)/,
  },
];

// Banner before rule, and the rule's label must not be bare punctuation - the
// same ordering trap as VBA_BANNER / VBA_RULE above.
const HASH_BANNER = /^[ \t]*(?:\/\/|#)[ \t]*[-=*#]{3,}[ \t]*\r?$/;
const HASH_RULE = /^[ \t]*(?:\/\/|#)[ \t]*[-=*#]{3,}[ \t]*([^-=*#\s][^\r\n]*?)[ \t]*[-=*#]{3,}[ \t]*\r?$/;
const HASH_COMMENT_TEXT = /^[ \t]*(?:\/\/|#)[ \t]*(.*?)[ \t\r]*$/;

function scanLineOriented(text, declarations, commentStart) {
  const entries = [];
  const lines = text.split(/\r?\n/);
  for (let i = 0; i < lines.length; i += 1) {
    const raw = lines[i];
    if (HASH_BANNER.test(raw)) {
      for (let j = i + 1; j <= i + 2 && j < lines.length; j += 1) {
        if (HASH_BANNER.test(lines[j])) continue;
        const name = (HASH_COMMENT_TEXT.exec(lines[j]) ?? [])[1];
        if (name) entries.push({ line: i + 1, kind: 'section', name });
        break;
      }
      continue;
    }
    const ruled = HASH_RULE.exec(raw);
    if (ruled) {
      entries.push({ line: i + 1, kind: 'section', name: ruled[1] });
      continue;
    }
    if (commentStart.test(raw)) continue;
    for (const { kind, re } of declarations) {
      const match = re.exec(raw);
      if (match) {
        entries.push({ line: i + 1, kind, name: match[1], call: true });
        break;
      }
    }
  }
  return entries;
}

const scanJs = (text) => scanLineOriented(text, JS_DECLARATIONS, /^[ \t]*(?:\/\/|\*|\/\*)/);

const PS_DECLARATIONS = [
  { kind: 'function', re: /^[ \t]*function[ \t]+([A-Za-z_][\w-]*)/i },
];
const scanPs = (text) => scanLineOriented(text, PS_DECLARATIONS, /^[ \t]*#/);

// ------------------------------------------------------------------- render

const SCANNERS = new Map([
  ['.bas', scanVba],
  ['.cls', scanVba],
  ['.frm', scanVba],
  ['.mjs', scanJs],
  ['.js', scanJs],
  ['.ps1', scanPs],
]);

export const INDEXABLE = new Set(SCANNERS.keys());

// The same extension list as a git pathspec, so the freshness stamp and the
// check printed in the header can never drift from what is actually indexed.
const PATHSPEC = [...INDEXABLE].map((ext) => `*${ext}`);

const kb = (bytes) => `${(bytes / 1024).toFixed(1)} KB`;
// A banner comment can contain anything; backticks and pipes break the row.
const md = (s) => s.replace(/[`|]/g, '\\$&');

export function renderMap(files, meta = {}) {
  const out = [
    '# Code map',
    '',
    `Generated by \`scripts/code-map.mjs\` on ${meta.date ?? 'unknown date'}, indexing`,
    `source as of commit \`${meta.commit ?? 'unknown'}\`.`,
    '',
    'Do not edit by hand; your edits are overwritten on the next run. Grep this',
    'file for a symbol, then read only the lines around the anchor. Reading an',
    'indexed file whole defeats the purpose - `src/core/Core.bas` alone would',
    'cost a fifth of a context window.',
    '',
    '**Before trusting an anchor, check that the map is current:**',
    '',
    '```',
    `git log -1 --format=%h -- ${PATHSPEC.map((p) => `'${p}'`).join(' ')}`,
    '```',
    '',
    'If that prints anything else, indexed source has changed since the map was',
    'generated and an anchor may point at the wrong slice - regenerate with',
    '`/code-map` rather than working around it.',
    '',
    'The stamp is the last commit that touched an indexed file, deliberately not',
    '`HEAD`: a map cannot name the commit that adds it, so stamping `HEAD` would',
    'report every freshly-committed map as stale.',
    '',
    'VBA procedure entries carry their scope (`public sub`, `private function`),',
    'because the flat-namespace rule in `CLAUDE.md` is enforceable only against',
    'that: everything `Public` must be `Core_*` or `Pkg<Name>_*`.',
    '',
  ];

  for (const { path, bytes, lines, entries } of files) {
    out.push(`## \`${path}\``, '', `${kb(bytes)} · ${lines} lines`, '');
    if (!entries.length) {
      out.push('_no indexable symbols_', '');
      continue;
    }
    for (const { line, kind, name, call } of entries) {
      const label = call ? `\`${md(name)}()\`` : `**${md(name)}**`;
      out.push(`- \`${path}:${line}\` — ${label} (${kind})`);
    }
    out.push('');
  }
  return `${out.join('\n')}\n`;
}

// ------------------------------------------------------------------ collect

const EXCLUDED_DIRS = new Set(['.git', 'node_modules', '.claude', 'dist']);
const DEFAULT_MIN_BYTES = 2048;

function walkFiles(dir, root = dir, acc = []) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      if (EXCLUDED_DIRS.has(entry.name)) continue;
      walkFiles(join(dir, entry.name), root, acc);
    } else if (entry.isFile()) {
      acc.push(relative(root, join(dir, entry.name)).split(sep).join('/'));
    }
  }
  return acc;
}

const git = (repoDir, args) => {
  const r = spawnSync('git', ['-C', repoDir, ...args], { encoding: 'utf8' });
  return r.status === 0 ? r.stdout : null;
};

// Repo-relative paths git does not ignore, or null outside a git repo. Without
// this, a stale copy under an ignored directory is indexed beside the real file
// and every symbol grep returns two plausible anchors.
function gitVisible(repoDir) {
  const out = git(repoDir, ['ls-files', '-z', '--cached', '--others', '--exclude-standard']);
  return out === null ? null : new Set(out.split('\0').filter(Boolean));
}

// Short sha of the last commit that touched an indexed file. Not `rev-parse
// HEAD`: the map is committed after it is generated, so a HEAD stamp is one
// commit behind the moment it lands and the freshness check then fails on the
// very commit that generated the map.
export function sourceCommit(repoDir) {
  return git(repoDir, ['log', '-1', '--format=%h', '--', ...PATHSPEC])?.trim() || 'unknown';
}

export function collectFiles(repoDir, opts = {}) {
  const minBytes = opts.minBytes ?? DEFAULT_MIN_BYTES;
  const visible = gitVisible(repoDir);

  // Size-filter BEFORE reading: statSync is cheap, running four regexes over
  // every line of a file only to discard it is not.
  return walkFiles(repoDir)
    .filter((rel) => INDEXABLE.has(extname(rel).toLowerCase()))
    .filter((rel) => !visible || visible.has(rel))
    .map((rel) => ({ rel, bytes: statSync(join(repoDir, rel)).size }))
    .filter(({ bytes }) => bytes >= minBytes)
    .map(({ rel, bytes }) => {
      const text = readFileSync(join(repoDir, rel), 'utf8');
      return {
        path: rel,
        bytes,
        lines: text.split(/\r?\n/).length,
        entries: SCANNERS.get(extname(rel).toLowerCase())(text),
      };
    })
    .sort((a, b) => b.bytes - a.bytes || a.path.localeCompare(b.path));
}

export function buildCodeMap(repoDir, opts = {}) {
  const files = collectFiles(repoDir, opts);
  const path = join(repoDir, opts.out ?? 'docs/code-map.md');
  mkdirSync(dirname(path), { recursive: true });
  const meta = {
    commit: opts.commit ?? sourceCommit(repoDir),
    date: opts.date ?? new Date().toISOString().slice(0, 10),
  };
  writeFileSync(path, renderMap(files, meta), 'utf8');
  return {
    path,
    commit: meta.commit,
    files: files.length,
    symbols: files.reduce((n, f) => n + f.entries.length, 0),
  };
}

// Positional repo path is optional (defaults to cwd), so the parser has to skip
// the VALUE of a flag as well as the flag itself - otherwise `--out docs/x.md`
// leaves `docs/x.md` looking like the repo path.
const VALUED_FLAGS = new Set(['--out', '--min-bytes']);

export function parseArgs(args) {
  const flags = {};
  let repo;
  for (let i = 0; i < args.length; i += 1) {
    if (VALUED_FLAGS.has(args[i])) {
      flags[args[i].slice(2)] = args[i + 1];
      i += 1;
    } else if (!args[i].startsWith('--') && repo === undefined) {
      repo = args[i];
    }
  }
  return { repo, flags };
}

function main(argv) {
  const { repo = process.cwd(), flags } = parseArgs(argv.slice(2));
  const flag = (key) => flags[key];
  const minBytesArg = flag('min-bytes');
  const { path, files, symbols, commit } = buildCodeMap(repo, {
    out: flag('out'),
    minBytes: minBytesArg === undefined ? undefined : Number(minBytesArg),
  });
  console.log(`${path} — ${files} file(s), ${symbols} symbol(s), source commit ${commit}`);
  return 0;
}

process.exitCode = main(process.argv);
