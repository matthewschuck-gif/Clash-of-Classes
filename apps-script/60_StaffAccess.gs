/**
 * Clash of Classes -- Apps Script backend
 * File 60: staff sign-in by Google account, roles, and the admin action log.
 *
 * WHO CAN SIGN IN: staff open the STAFF link (the "Anyone within EASD" deployment, the same one
 * as My Climb, without ?page=me) signed in with their school Google account and click "Sign in
 * with my school account". The server reads Session.getActiveUser() -- the page never says who
 * it is -- and lets them in only when:
 *   - the email ends in STAFF_EMAIL_DOMAIN (easdpa.org; student accounts are gomounts.org), and
 *   - it is on the Staff Access tab, Active is ticked, and its Role is in EDIT_ROLES.
 * The Sheet's owner (the account the web app runs as) is always an Admin, so nobody gets locked
 * out. Taking someone off the tab (or unticking Active) ends their session within 5 minutes.
 *
 * ROLES (Staff Access › Role): Admin, Counselor, Secretary, Teacher. EDIT_ROLES says which of
 * them can use Admin on this site (default: Admin only).
 *
 * THE SHARED PASSWORD: still works while PASSWORD_SIGN_IN is ON so nothing breaks on the day you
 * paste this. Its actions are logged as "Shared admin password". Turn it OFF once your admins are
 * on the Staff Access tab. The old built-in password from before the move no longer works.
 *
 * ADMIN ACTIONS LOG (tab Admin Actions): every admin call that changes something -- saves, point
 * changes, wheel spins, approvals, sheet edits, sign-ins and refused sign-ins -- with the time,
 * the person, their role and how they signed in. Read-only calls (loading a panel) are not logged.
 * Rows wait in the cache and are written every 5 minutes by backgroundTick, so logging never
 * makes the public site's cached data go stale.
 */

const STAFF_ACCESS_TAB_ = 'Staff Access';
const STAFF_ACCESS_HEADERS_ = ['Email', 'Name', 'Role', 'Active', 'Added', 'Last Sign-in'];
const ADMIN_ACTIONS_TAB_ = 'Admin Actions';
const ADMIN_ACTIONS_HEADERS_ = ['Time', 'Who', 'Email', 'Role', 'Signed in with', 'Action', 'Detail'];
const STAFF_ROLES_ = ['Admin', 'Counselor', 'Secretary', 'Teacher'];
const ACTION_BUFFER_KEY_ = 'clash-action-buffer';

// Admin calls that only read. Everything else that passes requireAdmin_ is logged.
const READ_ONLY_ADMIN_CALLS_ = ['clashLoad', 'clashMapLoad', 'clashHealth', 'clashVersion', 'clashAnnouncement',
  'clashSheetLoad', 'clashSheetTabs', 'clashSheetActions', 'clashServiceList', 'clashSelfCheck', 'clashFamilySwitches',
  'clashPlanned', 'clashApprovals', 'clashFairness', 'clashTeamLeague', 'clashDoorCheck', 'clashFamilyDoorFind',
  'clashReveal', 'clashWeekly', 'clashMe', 'clashMePage_', 'clashAccessInfo', 'clashActionLog'];

// Plain-language names for the log. Anything not listed shows its function name.
const ACTION_LABELS_ = {
  clashSave: 'Saved site data (points, events, wheel, settings)', clashSetPassword: 'Changed the shared admin password',
  clashMapSave: 'Saved the map', clashSheetSave: 'Edited a Sheet tab from the site', clashSheetAction: 'Ran a Sheet action',
  clashAvalancheStart: 'Started an Avalanche', clashServiceDecide: 'Decided a service project', clashServiceFile: 'Filed a service project',
  clashPlinkoDrop: 'Plinko drop', clashPlinkoGrant: 'Granted Plinko drops', clashFamilyDoorCheckIn: 'Family Night check-in',
  clashFamilyShoutDecide: 'Decided a family shout-out', clashFamilySwitch: 'Turned a family piece on/off', clashWrapSend: 'Sent the Friday wrap-up',
  clashPlannedPost: 'Posted a planned event', clashApprove: 'Approved a banner or shout-out', clashStartRivalry: 'Started Rivalry Week',
  clashDoorAdmit: 'Admitted at the Mayhem door', clashMeSaveCreature: 'Saved a creature (preview)', clashMePickBoost: 'Picked a boost (preview)',
  clashMeDrawTickets: 'Drew tickets (preview)', clashMeKeepTickets: 'Kept tickets (preview)', clashMeSaveLook: 'Saved a look (preview)',
};

function ensureStaffAccessTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [[STAFF_ACCESS_TAB_, STAFF_ACCESS_HEADERS_], [ADMIN_ACTIONS_TAB_, ADMIN_ACTIONS_HEADERS_]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[0] === STAFF_ACCESS_TAB_) {
      const owner = ownerEmail_();
      if (owner) sh.appendRow([owner, 'Sheet owner', 'Admin', true, todayYmd_(), '']);
      try {
        sh.getRange(2, 3, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STAFF_ROLES_, true).build());
        sh.getRange(2, 4, 500, 1).insertCheckboxes();
      } catch (e) { /* plain cells still work */ }
    }
  });
}

function ownerEmail_() {
  try { return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { return ''; }
}

function activeEmail_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { return ''; }
}

function editRoles_() {
  const list = String(getSetting_('EDIT_ROLES') || 'Admin').split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(String);
  if (list.indexOf('admin') === -1) list.push('admin');
  return list;
}

function passwordSignInOn_() { return String(getSetting_('PASSWORD_SIGN_IN') || 'ON').toUpperCase() !== 'OFF'; }

function staffDomainOk_(email) {
  const domains = String(getSetting_('STAFF_EMAIL_DOMAIN') || 'easdpa.org').split(',').map(function (s) { return s.trim().toLowerCase().replace(/^@/, ''); }).filter(String);
  const at = email.lastIndexOf('@');
  return at > 0 && domains.indexOf(email.slice(at + 1)) > -1;
}

// The Staff Access list, cached for 5 minutes (so removing someone takes effect quickly).
function staffAccessMap_(fresh) {
  const cache = CacheService.getScriptCache();
  if (!fresh) { const hit = cache.get('clash-staff-access'); if (hit) { try { return JSON.parse(hit); } catch (e) { /* rebuild */ } } }
  const map = {};
  sheetRows_(STAFF_ACCESS_TAB_).forEach(function (r) {
    const email = String(r['Email'] || '').trim().toLowerCase();
    if (!email) return;
    const active = r['Active'] === true || /^(true|yes|y|x|1|on)$/i.test(String(r['Active']).trim());
    map[email] = { name: String(r['Name'] || '').trim(), role: String(r['Role'] || '').trim(), active: active, row: r._row };
  });
  try { cache.put('clash-staff-access', JSON.stringify(map), 300); } catch (e) { /* fine */ }
  return map;
}

// Who this email is allowed to be right now, or null.
function staffRecord_(email) {
  email = String(email || '').toLowerCase();
  if (!email) return null;
  const owner = ownerEmail_();
  const rec = staffAccessMap_()[email];
  if (email === owner) return { email: email, name: (rec && rec.name) || 'Sheet owner', role: 'Admin' };
  if (!staffDomainOk_(email) || !rec || !rec.active) return null;
  const role = STAFF_ROLES_.filter(function (x) { return x.toLowerCase() === rec.role.toLowerCase(); })[0];
  if (!role || editRoles_().indexOf(role.toLowerCase()) === -1) return null;
  return { email: email, name: rec.name || email, role: role };
}

function newAdminSession_(who) {
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('clash-admin-' + token, JSON.stringify(who), SESSION_HOURS_ * 3600);
  return token;
}

// The session behind a token, or null. Google sessions are re-checked against the list.
function adminSession_(token) {
  if (!token || typeof token !== 'string' || token.length > 200) return null;
  const raw = CacheService.getScriptCache().get('clash-admin-' + token);
  if (!raw) return null;
  let who;
  if (raw === '1') who = { name: 'Shared admin password', email: '', role: 'Admin', via: 'password' }; // a session from before this file
  else { try { who = JSON.parse(raw); } catch (e) { return null; } }
  if (who.via === 'password' && !passwordSignInOn_()) return null;
  if (who.via === 'google') {
    const now = staffRecord_(who.email);
    if (!now) return null;
    who.role = now.role; who.name = now.name;
  }
  return who;
}

// ── SIGN-IN ─────────────────────────────────────────────────────────
/** Staff sign-in with the school Google account (works on the "Anyone within EASD" link). */
function clashStaffSignIn() {
  const email = activeEmail_();
  if (!email) {
    const link = String(getSetting_('STAFF_LINK') || getSetting_('STUDENT_LINK') || '').trim();
    throw new Error('NO_GOOGLE: Open the staff link signed in with your school Google account' + (link ? ': ' + link : '.'));
  }
  const who = staffRecord_(email);
  if (!who) {
    logAdminAction_({ name: '', email: email, role: '', via: 'google' }, 'Refused sign-in', 'Not on the Staff Access tab (or not Active, or role not in EDIT_ROLES)');
    throw new Error('NOT_ON_LIST: ' + email + ' is not on the Clash Staff Access list. Ask Mr. Schuck to add you.');
  }
  who.via = 'google';
  const token = newAdminSession_(who);
  noteSignIn_(email);
  logAdminAction_(who, 'Signed in', '');
  return JSON.stringify({ token: token, name: who.name, role: who.role, via: 'google' });
}

function noteSignIn_(email) {
  try {
    const rec = staffAccessMap_()[email];
    const sh = getSheetSafe_(STAFF_ACCESS_TAB_);
    if (rec && sh) {
      const day = todayYmd_(), key = 'clash-seen-signin-' + email + '-' + day, cache = CacheService.getScriptCache();
      if (!cache.get(key)) { sh.getRange(rec.row, 6).setValue(day); cache.put(key, '1', 21600); } // once a day, so it barely touches the Sheet
    }
  } catch (e) { /* never block a sign-in */ }
}

/** Who is signed in with this token (for the admin header and the Access card). */
function clashAccessInfo(token) {
  const who = requireAdmin_(token);
  const rows = sheetRows_(STAFF_ACCESS_TAB_).map(function (r) {
    return { email: String(r['Email'] || ''), name: String(r['Name'] || ''), role: String(r['Role'] || ''), active: r['Active'] === true || String(r['Active']).toUpperCase() === 'TRUE', last: String(r['Last Sign-in'] || '') };
  });
  return JSON.stringify({ me: { name: who.name, role: who.role, via: who.via }, passwordSignIn: passwordSignInOn_(), editRoles: STAFF_ROLES_.filter(function (r) { return editRoles_().indexOf(r.toLowerCase()) > -1; }), staff: who.role === 'Admin' ? rows : [] });
}

// ── ADMIN ACTIONS LOG ───────────────────────────────────────────────
// The admin function that called requireAdmin_, read from the call stack.
function adminCaller_() {
  try {
    const names = [], re = /at ([A-Za-z0-9_$]+) \(/g, stack = String(new Error().stack || '');
    let m;
    while ((m = re.exec(stack))) names.push(m[1]);
    const i = names.indexOf('requireAdmin_');
    let fn = i > -1 ? names[i + 1] : '';
    if (fn === 'meId_') fn = names[i + 2] || fn;
    if (fn && fn.slice(-1) === '_' && names[i + 2] && /^clash/.test(names[i + 2])) fn = names[i + 2]; // a helper → the public call
    return fn || '';
  } catch (e) { return ''; }
}

function logAdminAction_(who, action, detail) {
  try {
    const row = [new Date().toISOString(), who.name || '', who.email || '', who.role || '', who.via === 'password' ? 'Shared password' : who.via === 'google' ? 'Google account' : '', String(action || '').slice(0, 200), String(detail || '').slice(0, 500)];
    const lock = LockService.getScriptLock();
    let locked = false;
    try { locked = lock.tryLock ? lock.tryLock(500) : (lock.waitLock(500), true); } catch (e) { locked = false; }
    const cache = CacheService.getScriptCache();
    let buf = [];
    try { buf = JSON.parse(cache.get(ACTION_BUFFER_KEY_) || '[]'); } catch (e) { buf = []; }
    buf.push(row);
    if (buf.length > 300) buf = buf.slice(-300);
    cache.put(ACTION_BUFFER_KEY_, JSON.stringify(buf), 21600);
    if (locked) lock.releaseLock();
  } catch (e) { console.warn('Admin action log: ' + e.message); }
}

/** Writes waiting admin actions to the Admin Actions tab (from backgroundTick). */
function flushAdminActions_() {
  const cache = CacheService.getScriptCache();
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return; }
  let buf = [];
  try { buf = JSON.parse(cache.get(ACTION_BUFFER_KEY_) || '[]'); } catch (e) { buf = []; }
  if (buf.length) cache.remove(ACTION_BUFFER_KEY_);
  lock.releaseLock();
  if (!buf.length) return;
  let sh = getSheetSafe_(ADMIN_ACTIONS_TAB_);
  if (!sh) { ensureStaffAccessTabs_(); sh = getSheetSafe_(ADMIN_ACTIONS_TAB_); }
  const tz = Session.getScriptTimeZone();
  const rows = buf.map(function (r) { const out = r.slice(); out[0] = Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd HH:mm'); return out; });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, ADMIN_ACTIONS_HEADERS_.length).setValues(rows);
}

/** The latest admin actions, for Admin › System (Admins only). */
function clashActionLog(token) {
  const who = requireAdmin_(token);
  if (who.role !== 'Admin') return JSON.stringify({ rows: [] });
  let waiting = [];
  try { waiting = JSON.parse(CacheService.getScriptCache().get(ACTION_BUFFER_KEY_) || '[]'); } catch (e) { waiting = []; }
  const tz = Session.getScriptTimeZone();
  const fromCache = waiting.map(function (r) { return [Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd HH:mm')].concat(r.slice(1)); });
  const sh = getSheetSafe_(ADMIN_ACTIONS_TAB_);
  let saved = [];
  if (sh && sh.getLastRow() > 1) {
    const n = Math.min(40, sh.getLastRow() - 1);
    saved = sh.getRange(sh.getLastRow() - n + 1, 1, n, ADMIN_ACTIONS_HEADERS_.length).getValues()
      .map(function (r) { return r.map(function (v) { return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm') : String(v); }); });
  }
  const all = saved.concat(fromCache).reverse().slice(0, 40);
  return JSON.stringify({ rows: all.map(function (r) { return { time: r[0], who: r[1] || r[2], role: r[3], via: r[4], action: r[5], detail: r[6] }; }) });
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_60_StaffAccess = 'ac9c312cea';
