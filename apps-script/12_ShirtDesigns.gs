/**
 * Clash of Classes -- Apps Script backend
 * File 12: Mayhem shirt designs -- Generals upload, staff approve, each squad votes live.
 *
 * Shirt Designs tab: ID | Squad | Title | Uploaded By | Uploaded At | File | Approved | Votes | Winner
 *   File: the Drive file (an upload from Generals HQ fills it in; staff can paste any Drive
 *   image link or ID). The images live in a private Drive folder (SHIRT_DESIGN_FOLDER_ID, made
 *   on the first upload); the site sends them itself, so nothing is shared.
 *   Approved: nothing shows to students until staff tick it.
 *   Votes is filled in as votes come in; tick Winner on the design each squad picked.
 * Shirt Votes tab: Design | Squad | 800# | At   (one row per student; a changed vote replaces it)
 *
 * Students vote on My Squad with their 800#, for their own squad's designs only, while
 * SHIRT_VOTING is ON. The page refreshes the counts every few seconds.
 */

const DESIGN_HEADERS_ = ['ID', 'Squad', 'Title', 'Uploaded By', 'Uploaded At', 'File', 'Approved', 'Votes', 'Winner'];
const DESIGN_VOTE_HEADERS_ = ['Design', 'Squad', '800#', 'At'];
const MAX_DESIGN_BYTES_ = 1500000;

function ensureShirtDesignTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Shirt Designs', DESIGN_HEADERS_], ['Shirt Votes', DESIGN_VOTE_HEADERS_]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
}

function driveIdFrom_(v) {
  const m = String(v || '').match(/[-\w]{25,}/);
  return m ? m[0] : '';
}

function designFolder_() {
  const id = driveIdFrom_(getSetting_('SHIRT_DESIGN_FOLDER_ID'));
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* fall through and make one */ } }
  const folder = DriveApp.createFolder('Clash of Classes — Mayhem Shirt Designs');
  setSettingValue_('SHIRT_DESIGN_FOLDER_ID', folder.getId());
  return folder;
}

/** A General uploads a design for their squad (an image the page already shrank). */
function clashGeneralsUploadDesign(token, title, dataUrl) {
  const me = hqSession_(token);
  title = String(title || '').trim().slice(0, 80);
  if (!title) throw new Error('Give the design a name.');
  const m = String(dataUrl || '').match(/^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/=]+)$/);
  if (!m) throw new Error('Pick a PNG or JPG picture.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > MAX_DESIGN_BYTES_) throw new Error('That picture is too big. Try a smaller one.');
  const mine = sheetRows_('Shirt Designs').filter(function (r) { return squadOf_(r['Squad']) === me.squad; });
  if (mine.length >= (Number(getSetting_('SHIRT_DESIGNS_PER_SQUAD')) || 6)) throw new Error('Your squad already has the most designs allowed. Ask staff to remove one.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const file = designFolder_().createFile(Utilities.newBlob(bytes, m[1], me.squad + ' - ' + title + (m[1] === 'image/png' ? '.png' : '.jpg')));
    const id = 'D' + Date.now().toString(36);
    getSheet_('Shirt Designs').appendRow([id, me.squad, title, me.name || me.id,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), file.getId(), false, 0, false]);
  } finally {
    lock.releaseLock();
  }
  return clashGeneralsHq(token);
}

// Designs with their vote counts: [{ id, squad, title, by, approved, votes, winner, file }].
function designsWithVotes_() {
  const votes = sheetRows_('Shirt Votes');
  return sheetRows_('Shirt Designs').filter(function (r) { return String(r['ID'] || '').trim(); }).map(function (r) {
    return { id: String(r['ID']), row: r._row, squad: squadOf_(r['Squad']), title: String(r['Title'] || ''), by: String(r['Uploaded By'] || ''),
      approved: ticked_(r['Approved']), winner: ticked_(r['Winner']), file: driveIdFrom_(r['File']),
      votes: votes.filter(function (v) { return String(v['Design']) === String(r['ID']); }).length };
  });
}

/** For My Squad: this squad's approved designs, the counts, and what this student picked. */
function clashShirtDesigns(id) {
  const sid = id800_(id);
  const me = sid ? sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return id800_(r['800#']) === sid; })[0] : null;
  if (!me) return JSON.stringify({ designs: [], open: false });
  const squad = squadOf_(me['Squad']);
  const mine = sheetRows_('Shirt Votes').filter(function (v) { return id800_(v['800#']) === sid; })[0];
  const designs = designsWithVotes_().filter(function (d) { return d.approved && d.squad === squad; })
    .map(function (d) { return { id: d.id, title: d.title, by: d.by, votes: d.votes, winner: d.winner }; });
  return JSON.stringify({ squad: squad, designs: designs, voted: mine ? String(mine['Design']) : '',
    open: String(getSetting_('SHIRT_VOTING')).toUpperCase() === 'ON' });
}

/** A design's picture as a data URL (approved ones for everyone; a General's own squad's too). */
function clashShirtImage(designId, token) {
  const d = designsWithVotes_().filter(function (x) { return x.id === String(designId); })[0];
  if (!d || !d.file) throw new Error('That design isn\'t available.');
  if (!d.approved) {
    let ok = false;
    try { ok = hqSession_(token).squad === d.squad; } catch (e) { ok = false; }
    if (!ok) throw new Error('That design isn\'t approved yet.');
  }
  const cache = CacheService.getScriptCache();
  const hit = cache.get('design-' + d.id);
  if (hit) return hit;
  const blob = DriveApp.getFileById(d.file).getBlob();
  const url = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  if (url.length < 95000) cache.put('design-' + d.id, url, 21600);
  return url;
}

/** A student's vote (or a changed vote) for one of their own squad's designs. */
function clashShirtVote(id, designId) {
  if (String(getSetting_('SHIRT_VOTING')).toUpperCase() !== 'ON') throw new Error('Shirt voting is closed.');
  const cache = CacheService.getScriptCache();
  const used = Number(cache.get('clash-shirtvotes')) || 0;
  if (used >= 1500) throw new Error('The site is busy. Try again in a few minutes.');
  cache.put('clash-shirtvotes', String(used + 1), 600);
  const sid = id800_(id);
  const me = sheetRows_(getSetting_('STUDENT_BOARD_TAB') || 'Student Board').filter(function (r) { return id800_(r['800#']) === sid; })[0];
  if (!me) throw new Error('That 800 number isn\'t on the list.');
  const squad = squadOf_(me['Squad']);
  const d = designsWithVotes_().filter(function (x) { return x.id === String(designId); })[0];
  if (!d || !d.approved || d.squad !== squad) throw new Error('You can only vote for your own squad\'s designs.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_('Shirt Votes');
    const old = sheetRows_('Shirt Votes').filter(function (v) { return id800_(v['800#']) === sid; })[0];
    const row = [d.id, squad, sid, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')];
    if (old) sh.getRange(old._row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    // Keep the Votes column current so staff see the counts in the Sheet.
    const votesCol = DESIGN_HEADERS_.indexOf('Votes') + 1;
    designsWithVotes_().filter(function (x) { return x.squad === squad; }).forEach(function (x) { getSheet_('Shirt Designs').getRange(x.row, votesCol).setValue(x.votes); });
  } finally {
    lock.releaseLock();
  }
  return clashShirtDesigns(sid);
}

// For Generals HQ: every design from their squad, approved or not, with counts.
function hqDesigns_(squad) {
  return designsWithVotes_().filter(function (d) { return d.squad === squad; })
    .map(function (d) { return { id: d.id, title: d.title, by: d.by, approved: d.approved, votes: d.votes, winner: d.winner }; });
}
