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
  if (e && e.parameter && e.parameter.page === 'me') return serveMe_(); // My Climb, the student dashboard (16_Me.gs)
  const t = HtmlService.createTemplateFromFile('Index');
  t.config = JSON.stringify({ pollMs: Math.max(3, Number(getSetting_('POLL_SECONDS')) || 5) * 1000 });
  // The first load rides inside the page, so it shows without a second trip to the server.
  let boot = 'null';
  try { boot = cachedPublicLoad_().replace(/</g, '\\u003c'); } catch (err) { console.warn('Boot data: ' + err.message); }
  t.boot = boot;
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
function publicLoad_(raw) {
  const res = loadAppData_('coc');
  if (res.value) {
    delete res.value.password;
    maybeCheckThaws_(); // thaws a freeze whose goals are met (14_Thaw.gs), at most every 5 minutes
    maybeGameTick_(); // map claims and engine unlocks (15_Game.gs), at most every 10 minutes
    res.value.studentLink = String(getSetting_('STUDENT_LINK') || '');
    res.value.apps = programApps_(); // Trait Dough, Student of the Month, Mindset web apps
    res.value.freezes = activeFreezes_(); // read-only, from the Point Freezes tab
    try { res.value.schoolCalendar = schoolCalendar_(); } catch (e) { /* the page falls back to its built-in list */ }
    if (!raw && Array.isArray(res.value.events)) res.value.events = res.value.events.map(publicEvent_);
  }
  return res;
}

// The program web apps the site shows inside its pages (Settings; links from the EMS Hub).
function programApps_() {
  const u = function (k) { const v = String(getSetting_(k) || '').trim(); return /^https:\/\//i.test(v) ? v : ''; };
  return { embed: String(getSetting_('EMBED_APPS') || 'ON').toUpperCase() !== 'OFF', traitdough: u('TRAIT_DOUGH_APP'), sotm: u('SOTM_APP'), mindset: u('MINDSET_APP') };
}

// The site is public. Points rows other tools write can hold a student's name (Trait Dough,
// Mindset, Trail Journal). Behavior-related types (PRIVATE_EVENT_TYPES) show only their label,
// with no name or note. Other rows named "Something: First Last" show "First L." when
// TV_SHOW_NAMES is ON, otherwise just "Something". Staff signed in as admin get the full rows.
function publicEvent_(e) {
  const out = Object.assign({}, e);
  const priv = String(getSetting_('PRIVATE_EVENT_TYPES') || 'trail_journal, trailback_panel, camp_mountaineer')
    .split(',').map(function (s) { return s.trim(); }).filter(String);
  if (priv.indexOf(String(e.type)) > -1) {
    const labels = parsePairsSetting_('EVENT_LABELS');
    out.name = (labels[e.type] || String(e.type).replace(/_/g, ' ')).replace(/^[^\w]+\s*/, '');
    out.note = '';
    return out;
  }
  const m = String(e.name || '').match(/^(.*?(?:turn-in|winner|Summit|Growth))\s*[:—-]\s*(.+)$/i);
  if (m) {
    const parts = m[2].trim().split(/\s+/);
    const short = parts.length > 1 ? parts[0] + ' ' + parts[parts.length - 1].charAt(0) + '.' : parts[0];
    out.name = String(getSetting_('TV_SHOW_NAMES')).toUpperCase() === 'ON' ? m[1] + ': ' + short : m[1];
  }
  return out;
}

/** The site's data. With a valid admin session the rows come back unfiltered. */
function clashLoad(token) {
  let admin = false;
  try { if (token) { requireAdmin_(token); admin = true; } } catch (e) { admin = false; }
  return admin ? JSON.stringify(publicLoad_(true)) : cachedPublicLoad_();
}

// ── SPEED: cache the public data until the Sheet changes ────────────
// Everyone who opens the site gets the same public data, so it's built once and kept in the
// script cache, keyed by when the Sheet last changed (the same stamp the page polls). New points
// change the stamp, so nobody sees old numbers. Big values are split across cache keys.
const CACHE_CHUNK_ = 90000;
function cachePutBig_(key, text, seconds) {
  const cache = CacheService.getScriptCache(), parts = {}, n = Math.ceil(text.length / CACHE_CHUNK_) || 1;
  for (let i = 0; i < n; i++) parts[key + ':' + i] = text.slice(i * CACHE_CHUNK_, (i + 1) * CACHE_CHUNK_);
  parts[key] = String(n);
  try { cache.putAll(parts, seconds); } catch (e) { console.warn('Cache put: ' + e.message); }
}
function cacheGetBig_(key) {
  const cache = CacheService.getScriptCache(), n = Number(cache.get(key));
  if (!n) return null;
  const keys = [];
  for (let i = 0; i < n; i++) keys.push(key + ':' + i);
  const got = cache.getAll(keys);
  if (keys.some(function (k) { return got[k] == null; })) return null;
  return keys.map(function (k) { return got[k]; }).join('');
}
function dataStamp_() { try { return getAppDataMeta_('coc').updated_at; } catch (e) { return String(Math.floor(Date.now() / 60000)); } }

function cachedPublicLoad_(fresh) {
  const key = 'pub-load-' + dataStamp_();
  if (!fresh) { const hit = cacheGetBig_(key); if (hit) return hit; }
  const text = JSON.stringify(publicLoad_(false));
  cachePutBig_(key, text, 300);
  return text;
}

// Every 5 minutes (installed by Run setup): map claims, thaw checks, and a warm cache, so no
// visitor waits for them.
function backgroundTick() {
  noteRun_('backgroundTick'); // 22_Week.gs: for Admin › System
  try { maybeCheckThaws_(); } catch (e) { noteError_('thaw check', e); }
  try { maybeGameTick_(); } catch (e) { noteError_('map claims', e); }
  try { flushUsage_(); snapshotStages_(); } catch (e) { noteError_('usage + creature snapshot', e); }
  try { cachedPublicLoad_(true); cachedGame_(true); } catch (e) { noteError_('warm cache', e); }
}

function installBackgroundTick_() {
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'backgroundTick') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('backgroundTick').timeBased().everyMinutes(5).create();
  } catch (e) { console.warn('Background trigger: ' + e.message); }
}

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
    delete value.freezes; // lives on the Point Freezes tab, not in the saved blob
    delete value.schoolCalendar; // lives on the School Calendar tab
    delete value.studentLink; // lives on Settings
    const key = function (e) { return e.id + '|' + e.name; };
    const seen = {};
    (sync.seen || []).forEach(function (k) { seen[k] = true; });
    const events = Array.isArray(value.events) ? value.events : [];
    // A page that loaded before signing in holds the public (shortened) names: put the real
    // name and note back on any row the admin didn't change, so a save never rewrites them.
    const stored = readEventsTab_();
    const storedById = {};
    stored.forEach(function (s) { storedById[s.id] = s; });
    events.forEach(function (e) {
      const s = storedById[e.id];
      if (!s) return;
      const p = publicEvent_(s);
      if (e.name === p.name && e.name !== s.name) e.name = s.name;
      if (e.note === p.note && e.note !== s.note) e.note = s.note;
    });
    const mine = {};
    events.forEach(function (e) { mine[key(e)] = true; });
    const wasSeen = function (s) { return seen[key(s)] || seen[s.id + '|' + publicEvent_(s).name]; };
    const added = stored.filter(function (e) { return !wasSeen(e) && !mine[key(e)]; });
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
    applyFreezes_(null, value.events, []); // an admin-added event during a freeze counts 0 too
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
    .addSeparator()
    .addItem('Today\'s poll announcement', 'menuPollAnnouncement_')
    .addItem('Turn on daily poll points', 'menuTurnOnPolls_')
    .addItem('Turn off daily poll points', 'menuTurnOffPolls_')
    .addSeparator()
    .addItem('Turn on Squad Duel bonuses', 'menuTurnOnDuels_')
    .addItem('Turn off Squad Duel bonuses', 'menuTurnOffDuels_')
    .addSeparator()
    .addItem('Send shirt sale points', 'menuSendShirtPoints_')
    .addItem('Send Squad Night attendance points', 'menuSendSquadNightPoints_')
    .addSeparator()
    .addItem('Check thaw goals now', 'menuCheckThaws_')
    .addSeparator()
    .addItem('Import student roster', 'menuImportRoster_')
    .addItem('Turn on game awards', 'menuTurnOnGame_')
    .addItem('Turn off game awards', 'menuTurnOffGame_')
    .addItem('Use the building map (replaces Map Zones)', 'menuBuildingMap_')
    .addItem('Today\'s morning announcement', 'menuAnnouncement_')
    .addItem('Turn on the weekly digest email', 'menuTurnOnDigest_')
    .addItem('Turn off the weekly digest email', 'menuTurnOffDigest_')
    .addItem('Send me a preview of the weekly digest', 'menuPreviewDigest_')
    .addItem('Show or hide the Opponent', 'menuToggleMonster_')
    .addItem('Sync days off from the district calendar', 'menuSyncSchoolCalendar_')
    .addItem('Turn on daily calendar sync', 'menuTurnOnCalendarSync_')
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
