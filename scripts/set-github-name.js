'use strict';
// Puts your GitHub user name into the docs, README and package.json (replaces YOUR-GITHUB-NAME).
// Run: npm run github-name -- <your-github-user-name>
const fs = require('fs');
const path = require('path');

const name = process.argv[2];
if (!name || !/^[A-Za-z0-9-]+$/.test(name)) {
  console.error('Usage: npm run github-name -- <your-github-user-name>');
  process.exit(1);
}
const root = path.join(__dirname, '..');
const files = ['README.md', 'docs/privacy-policy.html', 'docs/store-listing.md', 'docs/DISTRIBUTION.md', 'docs/index.html'];
for (const rel of files) {
  const file = path.join(root, rel);
  if (!fs.existsSync(file)) continue;
  const text = fs.readFileSync(file, 'utf8');
  if (text.includes('YOUR-GITHUB-NAME')) {
    fs.writeFileSync(file, text.split('YOUR-GITHUB-NAME').join(name));
    console.log(`updated ${rel}`);
  }
}
const pkgFile = path.join(root, 'package.json');
const pkg = JSON.parse(fs.readFileSync(pkgFile, 'utf8'));
pkg.homepage = `https://github.com/${name}/nisar-office-studio`;
pkg.repository = { type: 'git', url: `https://github.com/${name}/nisar-office-studio.git` };
pkg.bugs = { url: `https://github.com/${name}/nisar-office-studio/issues` };
fs.writeFileSync(pkgFile, JSON.stringify(pkg, null, 2) + '\n');
console.log('updated package.json');
