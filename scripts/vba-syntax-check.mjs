#!/usr/bin/env node
// Headless VBA syntax check via the "VBA Pro" (vba-lsp) VS Code extension's
// language server. A fast pre-Excel gate that catches SYNTAX errors (malformed
// statements, unbalanced parens, bad blocks).
//
//   node scripts/vba-syntax-check.mjs dist/MacrosBundle.bas [more.bas ...]
//
// NOT a VBA compiler: it is an ANTLR-based parser, so it does NOT catch every
// compile error (undefined variables, type mismatches, a module-level Const
// misplaced after a procedure). `build/build.ps1` catches the structural rules
// it can; the final gate is still a manual import + Debug > Compile VBAProject
// in Excel. Exit code is non-zero if any file has an ERROR-severity diagnostic.
//
// Requires the extension installed locally (VS Code / Insiders / Cursor):
//   notisdataanalytics.vba-lsp  (Marketplace: "VBA Pro")
import { spawn } from 'node:child_process';
import { readFileSync, existsSync, readdirSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import path from 'node:path';
import os from 'node:os';

function findServer() {
  if (process.env.VBA_LSP_SERVER && existsSync(process.env.VBA_LSP_SERVER)) return process.env.VBA_LSP_SERVER;
  const roots = [
    path.join(os.homedir(), '.vscode', 'extensions'),
    path.join(os.homedir(), '.vscode-insiders', 'extensions'),
    path.join(os.homedir(), '.cursor', 'extensions'),
    path.join(os.homedir(), '.vscode-server', 'extensions'),
  ];
  const hits = [];
  for (const root of roots) {
    if (!existsSync(root)) continue;
    for (const name of readdirSync(root)) {
      if (!/vba-lsp/i.test(name)) continue;
      const p = path.join(root, name, 'dist', 'server', 'out', 'server.js');
      if (existsSync(p)) hits.push({ name, p });
    }
  }
  if (!hits.length) return null;
  hits.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  return hits[hits.length - 1].p; // highest version
}

function checkFile(server, file) {
  return new Promise((resolve) => {
    const target = path.resolve(file);
    const text = readFileSync(target, 'utf8');
    const uri = pathToFileURL(target).href;
    const rootUri = pathToFileURL(path.dirname(target)).href;
    const srv = spawn(process.execPath, [server, '--stdio'], { stdio: ['pipe', 'pipe', 'pipe'] });

    let buf = Buffer.alloc(0);
    const diagnostics = [];
    let nextId = 1;
    const pending = new Map();
    const send = (o) => { const s = JSON.stringify({ jsonrpc: '2.0', ...o }); srv.stdin.write(`Content-Length: ${Buffer.byteLength(s)}\r\n\r\n${s}`); };
    const request = (method, params) => { const id = nextId++; send({ id, method, params }); return new Promise((r) => pending.set(id, r)); };

    srv.stdout.on('data', (chunk) => {
      buf = Buffer.concat([buf, chunk]);
      for (;;) {
        const sep = buf.indexOf('\r\n\r\n');
        if (sep < 0) return;
        const m = /Content-Length:\s*(\d+)/i.exec(buf.slice(0, sep).toString());
        if (!m) { buf = buf.slice(sep + 4); continue; }
        const len = +m[1];
        if (buf.length < sep + 4 + len) return;
        const body = buf.slice(sep + 4, sep + 4 + len).toString();
        buf = buf.slice(sep + 4 + len);
        let msg; try { msg = JSON.parse(body); } catch { continue; }
        if (msg.method === 'textDocument/publishDiagnostics' && msg.params?.uri === uri) {
          diagnostics.length = 0; diagnostics.push(...(msg.params.diagnostics || []));
        } else if (msg.method && msg.id !== undefined) {
          let result = null;
          if (msg.method === 'workspace/configuration') {
            result = (msg.params.items || []).map(() => ({
              maxDocumentLines: 100000, maxNumberOfProblems: 1000,
              doWarnOptionExplicitMissing: true, environment: { os: 'Win64', version: 'Vba7' },
            }));
          }
          send({ id: msg.id, result });
        } else if (msg.id !== undefined && pending.has(msg.id)) {
          pending.get(msg.id)(msg); pending.delete(msg.id);
        }
      }
    });

    (async () => {
      await request('initialize', {
        processId: process.pid, rootUri,
        capabilities: { textDocument: { publishDiagnostics: {} }, workspace: { configuration: true } },
        workspaceFolders: [{ uri: rootUri, name: 'vba' }],
      });
      send({ method: 'initialized', params: {} });
      send({ method: 'textDocument/didOpen', params: { textDocument: { uri, languageId: 'vba', version: 1, text } } });
      await new Promise((r) => setTimeout(r, Number(process.env.VBA_LSP_WAIT_MS || 6000)));
      srv.kill();
      resolve(diagnostics);
    })();
  });
}

const files = process.argv.slice(2);
if (!files.length) { console.error('usage: node scripts/vba-syntax-check.mjs <file.bas> [...]'); process.exit(2); }
const server = findServer();
if (!server) { console.error('vba-lsp server not found. Install the "VBA Pro" extension, or set VBA_LSP_SERVER.'); process.exit(2); }

const sevName = { 1: 'ERROR', 2: 'WARN', 3: 'INFO', 4: 'HINT' };
let totalErrors = 0;
for (const f of files) {
  const diags = await checkFile(server, f);
  const errs = diags.filter((d) => d.severity === 1);
  totalErrors += errs.length;
  if (!diags.length) { console.log(`OK    ${f} — no diagnostics`); continue; }
  console.log(`${errs.length ? 'FAIL' : 'warn'}  ${f}`);
  for (const d of diags.sort((a, b) => a.range.start.line - b.range.start.line)) {
    console.log(`  ${sevName[d.severity] || d.severity}  L${d.range.start.line + 1}:${d.range.start.character + 1}  ${d.message}`);
  }
}
process.exit(totalErrors ? 1 : 0);
