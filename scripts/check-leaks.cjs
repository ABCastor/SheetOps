const fs = require('node:fs');
const {execFileSync} = require('node:child_process');
const files = process.argv.slice(2).length ? process.argv.slice(2) : execFileSync('git', ['ls-files', '--cached', '--others', '--exclude-standard', '-z'], {encoding:'utf8'}).split('\0').filter(Boolean);
const patterns = [
  {name:'absolute home path', re: /(?<![\w/:])\/(?:Users|home)\/[A-Za-z0-9._-]+(?:\/[^\s"'<>`]*)?/g},
  {name:'email address', re: /[A-Za-z0-9._%+-]+@([A-Za-z0-9.-]+\.[A-Za-z]{2,})/g},
];
let failed = false;
for (const file of files) {
  if (!/\.(?:[cm]?[jt]s|json|md|ya?ml|txt|gs|html|svg)$/.test(file) || !fs.existsSync(file)) continue;
  const content = fs.readFileSync(file, 'utf8');
  for (const {name,re} of patterns) for (const match of content.matchAll(re)) {
    if (name === 'email address' && /^(?:example\.(?:com|org|net)|users\.noreply\.github\.com)$/i.test(match[1])) continue;
    const line = content.slice(0, match.index).split('\n').length;
    // Report the location, never the matched value.
    console.error(`${file}:${line}: ${name}`);failed = true;
  }
}
if (failed) process.exit(1);
console.log('Leak check passed.');
