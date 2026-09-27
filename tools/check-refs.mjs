/**
 * Undefined-function check.
 *
 * `node --check` only validates syntax. A call to a function that does not
 * exist parses perfectly and fails only when that line runs — which is how a
 * half-applied edit once shipped a run loop that called `refreshRail()` with no
 * such function, silently discarding every successful report.
 *
 * This scans each module for bare calls `name(...)` and fails if `name` is not
 * declared, imported, a parameter-style local, or a known global. It is
 * deliberately simple and errs toward flagging; add genuine globals to KNOWN.
 *
 * Run: node tools/check-refs.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const KNOWN = new Set([
  // language and browser globals
  'if', 'for', 'while', 'switch', 'catch', 'function', 'return', 'typeof', 'await', 'new', 'super',
  'require', 'import', 'String', 'Number', 'Boolean', 'Array', 'Object', 'Date', 'Math', 'JSON',
  'Promise', 'Set', 'Map', 'Error', 'RegExp', 'Symbol', 'Blob', 'URL', 'fetch', 'setTimeout',
  'clearTimeout', 'setInterval', 'clearInterval', 'queueMicrotask', 'requestAnimationFrame',
  'isNaN', 'isFinite', 'parseInt', 'parseFloat', 'encodeURIComponent', 'decodeURIComponent',
  'structuredClone', 'AbortController', 'FileReader', 'confirm', 'alert', 'escape',
  'Intl', 'TextEncoder', 'TextDecoder', 'indexedDB', 'Infinity',
  // keywords that can precede "(" and are not calls
  'constructor', 'in', 'of', 'async', 'get', 'set', 'static', 'yield', 'delete', 'void', 'instanceof',
]);

const files = [];
(function walk(d) {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) walk(p);
    else if (p.endsWith('.js')) files.push(p);
  }
})(path.join(root, 'js'));

let problems = 0;
for (const f of files) {
  const src = fs.readFileSync(f, 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, (m) => m.replace(/[^\n]/g, ' '))   // block comments, keeping line numbers
    .replace(/(^|[^:])\/\/.*$/gm, '$1')         // line comments
    .replace(/`(?:\\.|[^`\\])*`/g, (m) => m.replace(/[^$\{\}\n]/g, ' ')) // template text, keep ${...}
    .replace(/'(?:\\.|[^'\\])*'|"(?:\\.|[^"\\])*"/g, '""')                // string literals
    .replace(/\\[A-Za-z]/g, '  ');                                        // regex escapes such as \b( and \s(

  const declared = new Set(KNOWN);
  for (const m of src.matchAll(/\bfunction\s*\*?\s*([A-Za-z_$][\w$]*)/g)) declared.add(m[1]);
  for (const m of src.matchAll(/\b(?:const|let|var|class)\s+([A-Za-z_$][\w$]*)/g)) declared.add(m[1]);
  for (const m of src.matchAll(/\b(?:const|let|var)\s*\{([^}]*)\}/g)) {
    m[1].split(',').forEach((x) => { const n = x.split(':').pop().split('=')[0].trim(); if (n) declared.add(n); });
  }
  for (const m of src.matchAll(/\bimport\s*\{([^}]*)\}/g)) {
    m[1].split(',').forEach((x) => { const n = x.trim().split(/\s+as\s+/).pop().trim(); if (n) declared.add(n); });
  }
  for (const m of src.matchAll(/\bimport\s+\*\s+as\s+(\w+)/g)) declared.add(m[1]);
  // Declarations are also harvested from the RAW source. Nested template
  // literals can defeat the simple template-blanking above and hide a real
  // declaration; reading the raw text as well only ever makes the check more
  // permissive, never less correct.
  const rawSrc = fs.readFileSync(f, 'utf8');
  for (const text of [src, rawSrc]) {
    for (const m of text.matchAll(/\bfunction\s*\*?\s*([A-Za-z_$][\w$]*)/g)) declared.add(m[1]);
    for (const m of text.matchAll(/\b(?:const|let|var|class)\s+([A-Za-z_$][\w$]*)/g)) declared.add(m[1]);
    // array destructuring: const [a, b] / for (const [a, b] of ...)
    for (const m of text.matchAll(/\b(?:const|let|var)\s*\[([^\]]*)\]/g)) {
      m[1].split(',').forEach((x) => { const n = x.replace(/=.*$/, '').replace(/\.\.\./, '').trim(); if (/^[A-Za-z_$][\w$]*$/.test(n)) declared.add(n); });
    }
    // object destructuring
    for (const m of text.matchAll(/\b(?:const|let|var)\s*\{([^}]*)\}/g)) {
      m[1].split(',').forEach((x) => { const n = x.split(':').pop().split('=')[0].replace(/\.\.\./, '').trim(); if (/^[A-Za-z_$][\w$]*$/.test(n)) declared.add(n); });
    }
    // later declarators in one statement: const W = 760, H = 190
    for (const m of text.matchAll(/\b(?:const|let|var)\s+[^;\n]*/g)) {
      for (const d of m[0].matchAll(/,\s*([A-Za-z_$][\w$]*)\s*=(?!=|>)/g)) declared.add(d[1]);
    }
    // destructured arrow / function parameters: ([k, v]) => and ({ a, b }) =>
    for (const m of text.matchAll(/\(\s*[\[{]([^\]}]*)[\]}]\s*\)\s*=>/g)) {
      m[1].split(',').forEach((x) => { const n = x.split(':').pop().split('=')[0].replace(/\.\.\./, '').trim(); if (/^[A-Za-z_$][\w$]*$/.test(n)) declared.add(n); });
    }
  }

  // Any identifier inside an arrow parameter list is a parameter, however it
  // is destructured: ([ref, used], i) => , ({ a, b: c }) =>
  for (const text of [src, rawSrc]) {
    for (const m of text.matchAll(/\(([^()]*)\)\s*=>/g)) {
      for (const id of m[1].matchAll(/[A-Za-z_$][\w$]*/g)) declared.add(id[0]);
    }
  }

  // catch (e) binds a name too
  for (const m of src.matchAll(/\bcatch\s*\(\s*([A-Za-z_$][\w$]*)/g)) declared.add(m[1]);
  // for (const x of ...) / for (let i = ...)
  for (const m of src.matchAll(/\bfor\s*\(\s*(?:const|let|var)\s+([A-Za-z_$][\w$]*)/g)) declared.add(m[1]);
  // parameters: (a, b) => and function x(a, b)
  for (const m of src.matchAll(/\(([^()]*)\)\s*=>/g)) m[1].split(',').forEach((x) => declared.add(x.replace(/[{}[\]=].*$/, '').trim()));
  for (const m of src.matchAll(/function\s*[\w$]*\s*\(([^)]*)\)/g)) m[1].split(',').forEach((x) => declared.add(x.replace(/[{}[\]=].*$/, '').trim()));
  for (const m of src.matchAll(/\b([A-Za-z_$][\w$]*)\s*=>/g)) declared.add(m[1]);

  // Bare interpolations: `${name}`. Template text is blanked above, which also
  // blanked these, so they are scanned in the ORIGINAL source. This is how a
  // print template shipped still interpolating ${chartHTML} after the variable
  // itself had been removed.
  const raw = fs.readFileSync(f, 'utf8');
  raw.split('\n').forEach((line, i) => {
    for (const m of line.matchAll(/\$\{\s*([A-Za-z_$][\w$]*)\s*\}/g)) {
      if (declared.has(m[1])) continue;
      console.log(`  ${path.relative(root, f)}:${i + 1}  interpolates "\${${m[1]}}" which is never defined or imported`);
      problems++;
    }
  });

  const lines = src.split('\n');
  lines.forEach((line, i) => {
    // bare calls only: not obj.method(), not new X()
    for (const m of line.matchAll(/(^|[^.\w$])([A-Za-z_$][\w$]*)\s*\(/g)) {
      const name = m[2];
      if (declared.has(name)) continue;
      console.log(`  ${path.relative(root, f)}:${i + 1}  calls "${name}()" which is never defined or imported`);
      problems++;
    }
  });
}

if (problems) {
  console.log(`\n${problems} undefined call${problems === 1 ? '' : 's'} found.`);
  process.exit(1);
}
console.log(`Reference check passed: every bare call in ${files.length} modules resolves.`);
