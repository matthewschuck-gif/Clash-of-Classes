/**
 * Clash of Classes -- Apps Script backend
 * File 3: the site itself, served by this project (no more GitHub Pages for the app).
 *
 * clashofclasses.org is now a tiny GitHub page that shows this web app full-screen inside it,
 * so the address students know stays the same. Everything else -- the page, the data, the
 * admin sign-in -- lives here and in this Sheet.
 *
 * Page files: Index (markup), Styles + SiteCss (CSS), App (the main script), Widgets (Ticker
 * and Growth board). The page talks to the functions below with google.script.run.
 *
 * Admin password: stored only as a salted SHA-256 hash in Script Properties (ADMIN_PW_HASH /
 * ADMIN_PW_SALT), set from the Sheet menu or Admin › Settings on the site. It is never sent to
 * the page and is no longer kept in the site data. Until one is set, the password that was
 * already in the site data still works once and is then saved scrambled.
 *
 * Saving: only with an admin session (about 6 hours). The save MERGES the events tab: rows
 * other projects added after the page loaded (Trait Dough, Mindset, Trail Journal, STAR
 * growth...) are kept instead of being overwritten by the admin's older copy.
 */

const SESSION_HOURS_ = 6;
const LEGACY_DEFAULT_PW_ = 'ephrata2026'; // the site's built-in password before this move

// ── PAGE ────────────────────────────────────────────────────────────
function serveSite_(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.config = JSON.stringify({ pollMs: Math.max(3, Number(getSetting_('POLL_SECONDS')) || 5) * 1000 });
  return t.evaluate()
    .setTitle(getSetting_('SITE_TITLE') || 'Clash of Classes — Ephrata Middle School')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover')
    // clashofclasses.org shows this page inside a frame, so framing must be allowed.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// ── DATA ────────────────────────────────────────────────────────────
// The site data with the password taken out (anyone can open the site).
function publicLoad_() {
  const res = loadAppData_('coc');
  if (res.value) delete res.value.password;
  return res;
}

function clashLoad() { return JSON.stringify(publicLoad_()); }

function clashMeta() { return JSON.stringify(getAppDataMeta_('coc')); }

/**
 * Admin save. sync = { seen: ['id|name', …] the page had loaded, metrics: the numbers it
 * loaded }. Events the page never saw are kept; events it saw and dropped were deleted on
 * purpose. The metrics tab is only rewritten when the page changed the numbers.
 */
function clashSave(token, valueJson, syncJson) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const value = JSON.parse(valueJson);
    const sync = syncJson ? JSON.parse(syncJson) : {};
    delete value.password;
    const key = function (e) { return e.id + '|' + e.name; };
    const seen = {};
    (sync.seen || []).forEach(function (k) { seen[k] = true; });
    const mine = {};
    const events = Array.isArray(value.events) ? value.events : [];
    events.forEach(function (e) { mine[key(e)] = true; });
    const added = readEventsTab_().filter(function (e) { return !seen[key(e)] && !mine[key(e)]; });
    // A row another project added can share an id with one the page just made: renumber it.
    const ids = {};
    let maxId = 100;
    events.forEach(function (e) { ids[e.id] = true; maxId = Math.max(maxId, Number(e.id) || 0); });
    added.forEach(function (e) { maxId = Math.max(maxId, Number(e.id) || 0); });
    added.forEach(function (e) {
      if (ids[e.id]) e.id = ++maxId;
      ids[e.id] = true;
    });
    value.events = events.concat(added);
    if (value.nextId !== undefined && Number(value.nextId) <= maxId) value.nextId = maxId + 1;
    const unchanged = !!sync.metrics && sameMetrics_(sync.metrics, value.metrics);
    const res = saveAppData_('coc', value, { skipMetricsTab: unchanged });
    return JSON.stringify({ saved: true, keptFromOthers: added.length, updated_at: res.updated_at });
  } finally {
    lock.releaseLock();
  }
}

function sameMetrics_(a, b) {
  if (!a || !b) return false;
  return ['attendance', 'referral', 'grades'].every(function (k) {
    return ['BLACK', 'GOLD', 'GREY', 'PURPLE'].every(function (s) {
      return (Number(a[k] && a[k][s]) || 0) === (Number(b[k] && b[k][s]) || 0);
    });
  });
}

/** A student's poll vote: no sign-in, changes only the poll's vote counts. */
function clashVote(question, optIndex, prevIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const res = loadAppData_('coc');
    const v = res.value || {};
    const poll = v.poll || {};
    const options = poll.options || [];
    optIndex = Number(optIndex);
    if (!poll.active || poll.question !== question) throw new Error('This poll is closed.');
    if (!(optIndex >= 0 && optIndex < options.length)) throw new Error('Unknown choice.');
    poll.votes = poll.votes || {};
    prevIndex = Number(prevIndex);
    if (prevIndex >= 0 && prevIndex !== optIndex) poll.votes[prevIndex] = Math.max(0, (Number(poll.votes[prevIndex]) || 1) - 1);
    if (prevIndex !== optIndex) poll.votes[optIndex] = (Number(poll.votes[optIndex]) || 0) + 1;
    v.poll = poll;
    saveAppData_('coc', v, { skipMetricsTab: true, skipEventsTab: true });
    return JSON.stringify(poll.votes);
  } finally {
    lock.releaseLock();
  }
}

/** Community Culture Stats: counts only from the Trail Journal Data Sheet, no names or text. */
function clashTrailJournalStats() {
  const id = String(getSetting_('TJ_SHEET_ID') || '').trim();
  const out = { reflections: [], panels: [] };
  if (!id) return JSON.stringify(out);
  const book = SpreadsheetApp.openById(id);
  const pick = function (tab, cols) {
    const sh = book.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return [];
    const data = sh.getDataRange().getValues();
    const h = data[0].map(function (x) { return String(x).trim(); });
    const idx = cols.map(function (c) { return h.indexOf(c); });
    return data.slice(1).map(function (r) {
      const o = {};
      cols.forEach(function (c, i) {
        let val = idx[i] > -1 ? r[idx[i]] : '';
        if (Object.prototype.toString.call(val) === '[object Date]') val = val.toISOString();
        o[c] = val;
      });
      return o;
    });
  };
  out.reflections = pick('reflections', ['insight_level', 'created_at']);
  out.panels = pick('panel_sessions', ['clash_squad', 'status']);
  return JSON.stringify(out);
}

// ── ADMIN SIGN-IN ───────────────────────────────────────────────────
function clashLogin(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('clash-login-fails')) || 0;
  if (fails >= 10) throw new Error('Too many tries. Wait 10 minutes.');
  if (!checkPassword_(String(pw || ''))) {
    cache.put('clash-login-fails', String(fails + 1), 600);
    throw new Error('Incorrect password.');
  }
  const token = Utilities.getUuid();
  cache.put('clash-admin-' + token, '1', SESSION_HOURS_ * 3600);
  return JSON.stringify({ token: token });
}

function requireAdmin_(token) {
  if (!token || !CacheService.getScriptCache().get('clash-admin-' + token)) {
    throw new Error('Please sign in as admin again.');
  }
}

function clashSetPassword(token, pw) {
  requireAdmin_(token);
  if (String(pw || '').length < 4) throw new Error('Use at least 4 characters.');
  setPassword_(String(pw));
  return JSON.stringify({ ok: true });
}

function checkPassword_(pw) {
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty('ADMIN_PW_HASH');
  if (hash) return hashPw_(pw, props.getProperty('ADMIN_PW_SALT')) === hash;
  // Not set yet: the password the site used before (in the site data, or the built-in one).
  // The first correct sign-in saves it scrambled, and the site data copy is then dropped.
  let legacy = '';
  try { legacy = (loadAppData_('coc').value || {}).password || ''; } catch (e) { /* no data yet */ }
  const ok = pw && (pw === legacy || pw === LEGACY_DEFAULT_PW_);
  if (ok) setPassword_(legacy || pw);
  return !!ok;
}

function setPassword_(pw) {
  const salt = Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ADMIN_PW_SALT', salt);
  props.setProperty('ADMIN_PW_HASH', hashPw_(pw, salt));
}

function hashPw_(pw, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

// ── SHEET MENU ──────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Clash of Classes')
    .addItem('Run setup (tabs + Settings)', 'setupSpreadsheet')
    .addItem('Set admin password…', 'menuSetPassword_')
    .addItem('Open the site', 'menuOpenSite_')
    .addSeparator()
    .addItem('Turn on challenge bonuses', 'menuTurnOnChallenges_')
    .addItem('Turn off challenge bonuses', 'menuTurnOffChallenges_')
    .addToUi();
}

function menuSetPassword_() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Set the Clash admin password', 'Type the new password (at least 4 characters). It is stored scrambled; nobody can read it back.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pw = r.getResponseText();
  if (pw.length < 4) { ui.alert('Too short: use at least 4 characters.'); return; }
  setPassword_(pw);
  ui.alert('Admin password saved.');
}

function menuOpenSite_() {
  const url = getSetting_('SITE_URL') || ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(
    '<p><a href="' + url + '" target="_blank">Open Clash of Classes</a></p>').setWidth(300).setHeight(80), 'Clash of Classes');
}
