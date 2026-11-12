// Writes package.json's version into src/version.js (run by `npm version`, its "version" script): the stand-in fonts'
// default address on jsDelivr names the release.
import fs from 'node:fs';

const { version } = JSON.parse(fs.readFileSync('package.json', 'utf8'));
const file = 'src/version.js';
fs.writeFileSync(file, fs.readFileSync(file, 'utf8').replace(/VERSION = '[^']*'/, `VERSION = '${version}'`));
console.log(`src/version.js: ${version}`);
