/**
 * Clash of Classes -- Apps Script backend
 * File 58: Cairn shout-outs become squad points on the events tab, by themselves.
 *
 * The Cairn (bit.ly/ems-cairn, the Belay Line shout-out form) writes to the incident sheet's
 * "Cairn Shoutouts" tab: Timestamp | Shoutout For | What They Did | Category | Trait |
 * Nominator Grade | Nominator Squad | Nominator 800# | Submitted By | Status.
 * Staff already read each one: Status "Posted to Ticker" means a person said yes. That is the
 * approval, so nothing new waits here. Every hour (the background tick), each approved Cairn not
 * paid yet gives CAIRN_POINTS to the squad of the student who WROTE it (the form tells them "this
 * is how your squad gets credit"; Nominator Squad, else their 800# on the Students tab / Student
 * Board). One events row per squad per day: "Cairn shout-outs: Gold (3)", type cairn. No names go
 * on the events tab (the shout-out's text stays on the incident sheet and the ticker).
 * Cairn Points tab (private: keys only): Key | Cairn Date | Squad | Points | Sent.
 * Settings: CAIRN_EVENTS ON/OFF, CAIRN_POINTS (2), CAIRN_SHEET, CAIRN_TAB, CAIRN_APPROVED
 * (words in Status that mean yes; blank = every Cairn counts), CAIRN_FROM (yyyy-mm-dd: only
 * Cairns on or after it; blank = all).
 */

const CAIRN_LEDGER_HEADERS_ = ['Key', 'Cairn Date', 'Squad', 'Points', 'Sent'];

function ensureCairnTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Cairn Points')) return ss.getSheetByName('Cairn Points');
  const sh = ss.insertSheet('Cairn Points');
  sh.getRange(1, 1, 1, CAIRN_LEDGER_HEADERS_.length).setValues([CAIRN_LEDGER_HEADERS_]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function cairnKey_(r) {
  const t = r['Timestamp'];
  const when = Object.prototype.toString.call(t) === '[object Date]' ? t.toISOString() : String(t || '').trim();
  return when + '|' + id800_(r['Nominator 800#']) + '|' + String(r['What They Did'] || '').trim().slice(0, 40).toLowerCase();
}

/** The approved Cairns not paid yet: [{ key, date, squad }] (plus how many were skipped and why). */
function cairnsToPay_() {
  const tab = String(getSetting_('CAIRN_TAB') || 'Cairn Shoutouts');
  const rows = openRows_(getSetting_('CAIRN_SHEET') || '1VUDvmkqG6F7-1PQWWc3_FWQxGz6bj_tBRdi4UuAWqP0', tab);
  const yesWords = listOf_(getSetting_('CAIRN_APPROVED') == null ? 'posted, approved' : getSetting_('CAIRN_APPROVED'));
  const from = String(getSetting_('CAIRN_FROM') || '').slice(0, 10);
  const paid = {};
  sheetRows_('Cairn Points').forEach(function (r) { paid[String(r['Key'])] = true; });
  const look = squadLookup_(); // 27_FollowUpStats.gs
  const out = { pay: [], waiting: 0, noSquad: 0 };
  rows.forEach(function (r) {
    if (!String(r['What They Did'] || r['Shoutout For'] || '').trim()) return;
    const status = String(r['Status'] || '').toLowerCase();
    if (yesWords.length && !yesWords.some(function (w) { return status.indexOf(w) > -1; })) { out.waiting++; return; }
    const date = ymdOfCell_(r['Timestamp']) || todayYmd_();
    if (from && date < from) return;
    const key = cairnKey_(r);
    if (paid[key]) return;
    let squad = squadOf_(r['Nominator Squad']);
    if (SQUAD_KEYS_.indexOf(squad) === -1) squad = look.byId[id800_(r['Nominator 800#'])] || '';
    if (!squad) { out.noSquad++; return; }
    out.pay.push({ key: key, date: date, squad: squad });
  });
  return out;
}

/** Pays the approved Cairns. Returns a line for the menu / Admin › System. */
function payCairns_() {
  if (String(getSetting_('CAIRN_EVENTS') || 'ON').toUpperCase() === 'OFF') return 'CAIRN_EVENTS is OFF.';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureCairnTab_();
    const each = settingNum_('CAIRN_POINTS', 2), c = cairnsToPay_();
    if (!c.pay.length) return 'No new approved Cairns.' + (c.waiting ? ' ' + c.waiting + ' waiting for staff.' : '');
    const groups = {};
    c.pay.forEach(function (x) { const k = x.date + '|' + x.squad; (groups[k] = groups[k] || []).push(x); });
    const ev = getSheet_(TABS.EVENTS), led = getSheet_('Cairn Points'), today = todayYmd_(), lines = [];
    Object.keys(groups).sort().forEach(function (k) {
      const list = groups[k], date = list[0].date, s = list[0].squad, pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' };
      pts[s] = list.length * each;
      ev.appendRow(['', 'Cairn shout-outs: ' + titleSquad_(s) + ' (' + list.length + ')', 'cairn', date,
        list.length + ' Cairn shout-out' + (list.length === 1 ? '' : 's') + ' approved for the ticker', '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
      list.forEach(function (x) { led.appendRow([x.key, date, s, each, today]); });
      lines.push(titleSquad_(s) + ' +' + pts[s] + ' (' + date + ')');
    });
    return 'Cairns paid: ' + lines.join(', ') + '.' + (c.noSquad ? ' ' + c.noSquad + ' had no squad (check the 800#).' : '');
  } finally {
    lock.releaseLock();
  }
}

// Every hour from backgroundTick (03_Web.gs).
function maybePayCairns_() {
  const props = PropertiesService.getScriptProperties(), hour = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH');
  if (props.getProperty('cairn-run') === hour) return;
  props.setProperty('cairn-run', hour);
  payCairns_();
}

function menuPayCairns_() { SpreadsheetApp.getUi().alert('Cairn shout-outs', payCairns_(), SpreadsheetApp.getUi().ButtonSet.OK); }

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_58_Cairns = 'd5ecf0f4c7';
