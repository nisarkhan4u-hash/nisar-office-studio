'use strict';
// Writes THIRD-PARTY-NOTICES.md: every library shipped inside the app with its licence text.
// Run after changing dependencies: npm run notices
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const tree = JSON.parse(execSync('npm ls --omit=dev --all --json --long', { cwd: root, maxBuffer: 64e6 }).toString());
const found = new Map();
(function walk(node) {
  for (const [name, dep] of Object.entries(node.dependencies || {})) {
    const key = `${name}@${dep.version}`;
    if (dep.path && !found.has(key)) found.set(key, { name, version: dep.version, dir: dep.path });
    walk(dep);
  }
})(tree);

function licenceOf(dir) {
  const pkg = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8'));
  let id = pkg.license || pkg.licenses || 'see package';
  if (typeof id === 'object') id = [].concat(id).map(l => l.type || l).join(' OR ');
  const file = fs.readdirSync(dir).find(f => /^(licen[cs]e|copying)(\.|$)/i.test(f));
  return { id, text: file ? fs.readFileSync(path.join(dir, file), 'utf8').trim() : '', homepage: pkg.homepage || '' };
}

const out = [
  '# Third-party notices',
  '',
  'Nisar Office Studio is released under the MIT licence (see LICENSE). It includes the open-source software listed below, each under its own licence.',
  '',
  '## Electron',
  '',
  'The app runs on Electron (MIT licence, https://www.electronjs.org), which includes Chromium and Node.js. Their licences are installed with the app as `LICENSE.electron.txt` and `LICENSES.chromium.html`.',
  '',
  '## qpdf 12.4.2',
  '',
  'Apache License 2.0, https://qpdf.sourceforge.io. Used for PDF passwords and optimisation. The program and its notice are in the `resources/qpdf` folder of the app.',
  '',
  '```',
  fs.readFileSync(path.join(root, 'vendor', 'qpdf', 'LICENSE-NOTICE.txt'), 'utf8').trim(),
  '```',
  '',
];
for (const p of [...found.values()].sort((a, b) => a.name.localeCompare(b.name))) {
  const l = licenceOf(p.dir);
  out.push(`## ${p.name} ${p.version}`, '', `Licence: ${l.id}${/ OR /.test(l.id) && /MIT/.test(l.id) ? ' (used under the MIT licence)' : ''}${l.homepage ? ` · ${l.homepage}` : ''}`, '');
  if (l.text) out.push('```', l.text.replace(/```/g, "'''"), '```', '');
}
fs.writeFileSync(path.join(root, 'THIRD-PARTY-NOTICES.md'), out.join('\n'));
console.log(`THIRD-PARTY-NOTICES.md: ${found.size} packages`);
