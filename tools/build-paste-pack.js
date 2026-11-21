// Builds paste-pack.html (repo root) and ../PASTE-PACK-CLASH.html: every Apps Script file with a Copy
// button and a done checkbox (remembered in this browser for this build only), plus the steps.
//   node tools/stamp-files.js && node tools/build-paste-pack.js
const fs = require('fs');
const path = require('path');
const dir = path.join(__dirname, '..', 'apps-script');
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// Files made since the last full paste (they need + in the editor first).
const NEW = ['60_StaffAccess.gs', '61_Hardening.gs', '00_Version.gs', '58_Cairns.gs', '59_MapWeek.gs', 'SiteArcade.html', '34_SheetEditor.gs', '35_SiteContent.gs', '36_TV.gs', '37_Season.gs', '38_Avalanche.gs', '39_VirtualDough.gs', '40_Service.gs', '41_Morning.gs', '42_Family.gs', '43_Rollover.gs', '44_Reveal.gs', '45_SelfCheck.gs', '46_Plinko.gs', '47_Families.gs', '48_Arcade.gs', '49_TabOrder.gs', '50_Weather.gs', '51_Staff.gs', '52_Legacy.gs', '53_FamilyMore.gs', '54_Rally.gs', '55_Badges.gs', '56_Shoutouts.gs', '57_Leagues.gs', 'Recap.html', 'Plinko.html', 'FamilyDoor.html', 'Arcade.html', 'Announce.html', 'Hall.html', 'Reveal.html',
  'SheetEditor.html', 'Dough.html', 'Family.html', 'Morning.html', 'ServiceAdmin.html'];
// ONLY=a.gs,b.html → a pack of just those files (an update since the last full paste).
const ONLY = (process.env.ONLY || '').split(',').map(x => x.trim()).filter(Boolean);
const files = fs.readdirSync(dir).filter(f => /\.(gs|html)$/.test(f) && (!ONLY.length || ONLY.includes(f))).sort((a, b) => (a.endsWith('.gs') ? 0 : 1) - (b.endsWith('.gs') ? 0 : 1) || a.localeCompare(b));
const built = new Date().toISOString().slice(0, 16).replace('T', ' ');
const rows = files.map((f, i) => {
  const src = fs.readFileSync(path.join(dir, f), 'utf8').replace(/\r\n/g, '\n');
  const name = f.replace(/\.(gs|html)$/, ''), kind = f.endsWith('.gs') ? 'Script' : 'HTML';
  return `<section><header><label><input type="checkbox" data-k="${esc(f)}"> <b>${i + 1}. ${esc(name)}</b></label>
<span class="k">${kind}${NEW.includes(f) ? ' · <em>NEW if you don\'t have it: + › ' + kind + ', name it ' + esc(name) + '</em>' : ''}</span>
<span class="n">${src.split('\n').length} lines</span><button data-i="${i}">Copy</button></header><textarea readonly id="t${i}">${esc(src)}</textarea></section>`;
}).join('\n');
const html = `<!doctype html><html><head><meta charset="utf-8"><title>Clash paste pack</title>
<meta name="viewport" content="width=device-width,initial-scale=1"><style>
body{font:15px/1.55 system-ui,sans-serif;margin:0;background:#14082a;color:#f5f0ff}main{max-width:980px;margin:0 auto;padding:20px 16px 80px}
h1{margin:0 0 6px}p,li{color:#cbb6ef}li{margin:6px 0}li b,p b{color:#fff}code{background:#2a1a48;padding:1px 6px;border-radius:6px;color:#fff}
.what{background:#1d0f36;border:1px solid #3b2a5a;border-radius:12px;padding:12px 16px}
section{border:1px solid #3b2a5a;border-radius:12px;margin:10px 0;background:#1d0f36}header{display:flex;gap:12px;align-items:center;padding:10px 14px;flex-wrap:wrap}
header label{cursor:pointer;font-size:16px}.k{color:#b09ec8;font-size:13px}.k em{color:#ffe100;font-style:normal;font-weight:700}.n{margin-left:auto;color:#8f7fa8;font-size:12px}
button{font:700 14px system-ui;padding:7px 16px;border-radius:999px;border:0;background:#ffe100;color:#1a0b2e;cursor:pointer}button.ok{background:#86efac}
textarea{display:none}section.done{opacity:.5}#bar{position:sticky;top:0;background:#14082a;padding:10px 0;font-weight:700;z-index:1}
</style></head><body><main><h1>Clash of Classes paste pack</h1>
<p>Built ${built}. <b>This round: paste every file</b> (each now carries a version stamp). From now on Admin › 🩺 System lists exactly which files are newer on GitHub, so later you only paste those.</p>
<ol>
<li>Open the <b>Clash of Classes Points</b> sheet › <b>Extensions › Apps Script</b>.</li>
<li>Files marked <b>NEW</b> that aren't in your editor yet: make them first with <b>+</b> (Script or HTML, exact name, no extension).</li>
<li>For each file: <b>Copy</b> → click that file in the editor and check its name at the top → <b>Ctrl+A</b>, <b>Ctrl+V</b> → tick the box here.</li>
<li><b>Ctrl+S</b>. No orange dots should be left. A red syntax error means the file it names (or the one just below it) got the wrong paste: repaste that one.</li>
<li>In the sheet (reload it): <b>Clash of Classes › Run setup</b>. It adds the Avalanche, Dough Teachers, Virtual Dough, Service Submissions and Squad Spotlight tabs and the new Settings.</li>
<li><b>Deploy › Manage deployments</b> › pencil on <b>AKfycbyRCm58…</b> › Version: <b>New version</b> › Deploy (access stays <b>Anyone</b>). Then the same on the My Climb deployment.</li>
<li><b>Dough Teachers</b> tab (or Admin › 📋 Sheet): a row per teacher: tick Active, their name, squad and password (e.g. room number). Untick the example row.</li>
<li><b>Settings</b>: check <b>SEASON_END</b> (2027-05-28, the last school day of May; fix it if the calendar differs). Set <b>BASH_REVEAL</b> to a June date to keep the winner secret until then.</li>
<li>Check: Admin › 🩺 System should say <b>✅ Everything is live</b>.</li>
</ol>
<p class="what">For staff: <code>clashofclasses.org/dough</code> Virtual Trait Dough · <code>clashofclasses.org/morning</code> Morning Clash · <code>clashofclasses.org/family</code> family update · students vote at <code>clashofclasses.org/vote</code>.</p>
<div id="bar">Done: <span id="cnt">0</span> / ${files.length}</div>
${rows}</main><script>
var BUILD=${JSON.stringify(built)};
function key(k){return 'clash-paste-'+BUILD+'-'+k}
function count(){var b=document.querySelectorAll('input[data-k]');var n=0;b.forEach(function(x){if(x.checked)n++;x.closest('section').classList.toggle('done',x.checked)});document.getElementById('cnt').textContent=n}
document.querySelectorAll('input[data-k]').forEach(function(x){try{x.checked=localStorage.getItem(key(x.dataset.k))==='1'}catch(e){}x.onchange=function(){try{localStorage.setItem(key(x.dataset.k),x.checked?'1':'')}catch(e){}count()}});
document.querySelectorAll('button[data-i]').forEach(function(b){b.onclick=function(){var t=document.getElementById('t'+b.dataset.i);navigator.clipboard.writeText(t.value).then(function(){b.textContent='Copied ✓';b.classList.add('ok');setTimeout(function(){b.textContent='Copy';b.classList.remove('ok')},2500)},function(){t.style.display='block';t.select();document.execCommand('copy');b.textContent='Copied ✓'})}});
count();
</script></body></html>`;
fs.writeFileSync(path.join(__dirname, '..', 'paste-pack.html'), html);
fs.writeFileSync(path.join(__dirname, '..', '..', 'PASTE-PACK-CLASH.html'), html);
console.log('paste pack: ' + files.length + ' files, ' + Math.round(html.length / 1024) + ' KB');
