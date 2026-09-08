#!/usr/bin/env node
/**
 * Verifies every relative require() in the backend resolves to a file that is actually present.
 *
 * This exists because the Dockerfile used to copy backend modules into the image one by one.
 * Adding contact-resolver.js without adding a matching COPY line produced an image that built
 * cleanly, passed the health check, and then threw MODULE_NOT_FOUND the first time a user tried
 * to send mail. A static check turns that into a build failure instead.
 *
 * Static on purpose: it never require()s anything, so it has no side effects (no DB files
 * created, no network, no credentials needed).
 *
 * Run: node scripts/check-requires.js [rootDir]
 */
const fs = require('fs');
const path = require('path');

const root = path.resolve(process.argv[2] || path.join(__dirname, '..'));
const SKIP_DIRS = new Set(['node_modules', '.git', 'frontend', 'public', 'scripts', 'Kode Final', '.github']);
const REQUIRE_RE = /require\(\s*['"](\.[^'"]+)['"]\s*\)/g;

function jsFilesIn(dir) {
  const out = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.isDirectory()) {
      if (SKIP_DIRS.has(entry.name)) continue;
      out.push(...jsFilesIn(path.join(dir, entry.name)));
    } else if (entry.name.endsWith('.js')) {
      out.push(path.join(dir, entry.name));
    }
  }
  return out;
}

/** Mirrors Node's resolution for relative specifiers, minus package.json "main". */
function resolves(fromFile, spec) {
  const base = path.resolve(path.dirname(fromFile), spec);
  const candidates = [base, base + '.js', base + '.json', path.join(base, 'index.js')];
  return candidates.some((candidate) => {
    try { return fs.statSync(candidate).isFile(); } catch { return false; }
  });
}

const files = jsFilesIn(root);
const missing = [];

for (const file of files) {
  const source = fs.readFileSync(file, 'utf8');
  for (const match of source.matchAll(REQUIRE_RE)) {
    if (!resolves(file, match[1])) {
      missing.push({ file: path.relative(root, file), spec: match[1] });
    }
  }
}

const scanned = files.length;
if (missing.length === 0) {
  console.log(`[check-requires] OK - ${scanned} file(s) scanned, all relative requires resolve.`);
  process.exit(0);
}

console.error(`[check-requires] FAILED - ${missing.length} unresolved require(s) across ${scanned} file(s):`);
for (const item of missing) {
  console.error(`  ${item.file}  ->  require('${item.spec}')`);
}
console.error('');
console.error('If this fails inside a Docker build, the module exists in the repo but was not');
console.error('copied into the image. Check the COPY lines in the Dockerfile.');
process.exit(1);
