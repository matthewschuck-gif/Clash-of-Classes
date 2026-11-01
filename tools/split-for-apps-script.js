// One-time split: the old single-file site (index.html + styles.css) → Apps Script HTML files
// in apps-script/. After this runs, apps-script/*.html are the source of truth for the site,
// and the root index.html becomes a thin frame for clashofclasses.org (see tools/wrapper.html).
//
//   node tools/split-for-apps-script.js
//
// Output: Index.html (head + page markup + includes), Styles.html, SiteCss.html, App.html
// (the main script), Widgets.html (Freudenfreude Ticker + Growth Leaderboard).
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const out = path.join(root, 'apps-script');
const src = fs.readFileSync(path.join(root, 'index.html'), 'utf8').replace(/^﻿/, '').replace(/\r\n/g, '\n');
const lines = src.split('\n');

const find = (re, from = 0) => { for (let i = from; i < lines.length; i++) if (re.test(lines[i])) return i; throw new Error('not found: ' + re); };
const styleStart = find(/^<style>\s*$/);
const styleEnd = find(/^<\/style>\s*$/, styleStart);
const cssLink = find(/href="\/styles\.css"/);
const appStart = find(/^<script type="module">/);
const appEnd = find(/^<\/script>\s*$/, appStart);
const widgetsStart = appEnd + 1;
const widgetsEnd = find(/^<\/script>\s*$/, find(/^<script>/, find(/^<\/script>\s*$/, widgetsStart + 1) + 1));

const head = lines.slice(0, styleStart)
  .filter(l => !/rel="manifest"|rel="apple-touch-icon"|rel="canonical"|http-equiv=|<!DOCTYPE|^<!-- CoC v/.test(l));
const body = lines.slice(cssLink + 1, appStart);
const tail = lines.slice(widgetsEnd + 1);

const index = [
  '<!DOCTYPE html>',
  '<!-- Clash of Classes, served by Apps Script. Page markup; styles and scripts are included below. -->',
  ...head.slice(head.findIndex(l => /^<html/.test(l))),
  "<?!= include('Styles'); ?>",
  "<?!= include('SiteCss'); ?>",
  ...body,
  "<?!= include('App'); ?>",
  "<?!= include('Widgets'); ?>",
  ...tail,
].join('\n');

fs.writeFileSync(path.join(out, 'Index.html'), index);
fs.writeFileSync(path.join(out, 'Styles.html'), lines.slice(styleStart, styleEnd + 1).join('\n') + '\n');
fs.writeFileSync(path.join(out, 'SiteCss.html'), '<style>\n' + fs.readFileSync(path.join(root, 'styles.css'), 'utf8').replace(/\r\n/g, '\n') + '\n</style>\n');
fs.writeFileSync(path.join(out, 'App.html'), lines.slice(appStart, appEnd + 1).join('\n') + '\n');
fs.writeFileSync(path.join(out, 'Widgets.html'), lines.slice(widgetsStart, widgetsEnd + 1).join('\n') + '\n');
for (const f of ['Index', 'Styles', 'SiteCss', 'App', 'Widgets']) {
  console.log(f + '.html', fs.statSync(path.join(out, f + '.html')).size, 'bytes');
}
