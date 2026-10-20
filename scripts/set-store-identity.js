'use strict';
// Copies the three Product identity values from Partner Center into package.json (build.appx).
// Run: node scripts/set-store-identity.js "<Package/Identity/Name>" "<Package/Identity/Publisher>" "<PublisherDisplayName>"
const fs = require('fs');
const path = require('path');

const [identityName, publisher, publisherDisplayName] = process.argv.slice(2);
if (!identityName || !publisher || !publisherDisplayName) {
  console.error('Usage: node scripts/set-store-identity.js "<Package/Identity/Name>" "<Package/Identity/Publisher>" "<Package/Properties/PublisherDisplayName>"');
  console.error('Copy the values from Partner Center > your app > Product management > Product identity.');
  process.exit(1);
}
if (!/^CN=/.test(publisher)) {
  console.error('The publisher must start with "CN=" exactly as Partner Center shows it.');
  process.exit(1);
}
const file = path.join(__dirname, '..', 'package.json');
const pkg = JSON.parse(fs.readFileSync(file, 'utf8'));
Object.assign(pkg.build.appx, { identityName, publisher, publisherDisplayName });
fs.writeFileSync(file, JSON.stringify(pkg, null, 2) + '\n');
console.log('package.json updated. Now run: npm run dist:store');
