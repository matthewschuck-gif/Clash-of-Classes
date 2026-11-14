// Stamps every Apps Script file with a short fingerprint of its content, and writes 00_Version.gs
// with the list. Admin › System compares the live project's stamps with GitHub's to say which files
// still need pasting (02_Router / 03_Web clashVersion, 00_Version.gs).
//   node tools/stamp-files.js          stamp + write 00_Version.gs
//   node tools/stamp-files.js --check  exit 1 if any stamp is out of date (the tests run this)
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const dir = path.join(__dirname, '..', 'apps-script');
const check = process.argv.includes('--check');
const GS_STAMP = /\n?\/\/ file-stamp[^\n]*\nvar FILE_STAMP_[A-Za-z0-9_]+ = '[0-9a-f]+';\n?$/;
const HTML_STAMP = /\n?<!-- file-stamp: [0-9a-f]+ -->\n?$/;

const varName = f => 'FILE_STAMP_' + f.replace(/\.gs$/, '').replace(/[^A-Za-z0-9]/g, '_');
const body = (f, s) => s.replace(/\r\n/g, '\n').replace(f.endsWith('.gs') ? GS_STAMP : HTML_STAMP, '').replace(/\s+$/, '');
const hash = s => crypto.createHash('sha1').update(s.replace(/\s+/g, '')).digest('hex').slice(0, 10);

const files = fs.readdirSync(dir).filter(f => /\.(gs|html)$/.test(f) && f !== '00_Version.gs').sort();
const stamps = {};
let stale = [];
files.forEach(f => {
  const p = path.join(dir, f), raw = fs.readFileSync(p, 'utf8'), crlf = raw.includes('\r\n');
  const b = body(f, raw), h = hash(b);
  stamps[f] = h;
  const out = f.endsWith('.gs') ? b + '\n\n// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one\nvar ' + varName(f) + " = '" + h + "';\n"
    : b + '\n<!-- file-stamp: ' + h + ' -->\n';
  if (out !== raw.replace(/\r\n/g, '\n')) { stale.push(f); if (!check) fs.writeFileSync(p, crlf ? out.replace(/\n/g, '\r\n') : out); }
});
const build = new Date().toISOString().slice(0, 16).replace('T', ' ');
const listing = JSON.stringify(stamps, null, 0).replace(/,"/g, ',\n  "').replace('{"', '{\n  "').replace(/}$/, '\n}');
const versionFile = `/**
 * Clash of Classes -- Apps Script backend
 * File 00: which version of every file this is (written by tools/stamp-files.js; do not edit).
 * Admin › System compares these with the stamps the live files carry, and with the newest list on
 * GitHub, to show which files still need pasting.
 */
const FILE_STAMPS_ = ${listing};
`;
const vp = path.join(dir, '00_Version.gs');
const old = fs.existsSync(vp) ? fs.readFileSync(vp, 'utf8').replace(/\r\n/g, '\n') : '';
if (old.replace(/\s+/g, '') !== versionFile.replace(/\s+/g, '')) { stale.push('00_Version.gs'); if (!check) fs.writeFileSync(vp, versionFile); }
if (check) {
  if (stale.length) { console.error('Out of date stamps: ' + stale.join(', ') + '\nRun: node tools/stamp-files.js'); process.exit(1); }
  console.log('  ok   every file stamp is current (' + files.length + ' files)');
} else console.log('stamped ' + stale.length + ' file(s); ' + files.length + ' files, build ' + build);
