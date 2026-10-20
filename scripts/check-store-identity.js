'use strict';
// Stops "npm run dist:store" when package.json still has the placeholder Store identity.
const pkg = require('../package.json');
const a = pkg.build.appx;
if (/REPLACE-WITH/.test(`${a.identityName} ${a.publisher}`) && !process.env.ALLOW_TEST_IDENTITY) {
  console.error('\nThe Microsoft Store identity is not set yet.');
  console.error('Reserve the app name in Partner Center, then run:');
  console.error('  npm run store-identity -- "<Package/Identity/Name>" "<Package/Identity/Publisher>" "<PublisherDisplayName>"');
  console.error('(To build a test package anyway, set ALLOW_TEST_IDENTITY=1.)\n');
  process.exit(1);
}
