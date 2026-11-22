/**
 * Clash of Classes -- Apps Script backend
 * File 65: Preview as a student (Admin › System › "Preview as a student").
 *
 * Staff already signed in to Admin pick a made-up sample student on a squad (nothing saves) or
 * type a real 800# (read-only look at exactly what that student sees). The page gets a short-lived
 * preview ticket (2 hours, in the cache) instead of the admin session or the 800# in the address.
 * The ticket opens, inside a pop-up on the site:
 *   - My Climb (?page=me&pv=…, with any tab: Me, Tickets, Hope, My Squad, Map, Game, Arcade)
 *   - Generals HQ for the squad (?page=hq&pv=…; that one is the real HQ, so it saves)
 *   - Summit Party minigame tryout (?page=summit-play&pv=…&mode=tryout): any minigame, alone, on repeat
 *   - Summit Party demo game: the board and the squad's Chromebook side by side; the computer plays
 *     the other three squads. A demo game never writes to the Summit tabs and never earns points,
 *     and it works even while Settings › SUMMIT_PARTY is OFF (63_SummitParty.gs, g.demo).
 * Turn it off with Settings › STUDENT_PREVIEW = OFF.
 */

const PV_HOURS_ = 2;

function previewOn_() { return String(getSetting_('STUDENT_PREVIEW') || 'ON').toUpperCase() !== 'OFF'; }

/** Admin › System: make a preview ticket. asId = 'sample-GOLD' (a made-up student) or a real 800#. */
function clashPreviewTicket(token, asId) {
  const who = requireAdmin_(token, { noLog: true });
  if (!previewOn_()) throw new Error('Student preview is off (Settings › STUDENT_PREVIEW).');
  let squad = sampleSquad_(asId), id = '', name = 'Sample student';
  if (!squad) {
    id = id800_(asId);
    if (id.length < 5) throw new Error('Type a student 800#, or pick a sample student.');
    const row = boardRows_().filter(function (r) { return id800_(r['800#']) === id; })[0];
    if (!row) throw new Error('That 800# is not on the Student Board.');
    squad = squadOf_(row['Squad']);
    if (SQUAD_KEYS_.indexOf(squad) === -1) throw new Error('That student is not on a squad yet.');
    name = String(row['Name'] || 'Student');
    logAdminAction_(who, 'Previewed a student', id);
  }
  const ticket = Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('pv-' + ticket, JSON.stringify({ token: token, asId: id ? 'view-' + id : 'sample-' + squad, squad: squad, name: name, real: !!id }), PV_HOURS_ * 3600);
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { url = ''; }
  return JSON.stringify({ ticket: ticket, url: url || String(getSetting_('STAFF_LINK') || getSetting_('STUDENT_LINK') || ''), squad: squad, name: name, real: !!id,
    summitOn: spOn_(), draw: spDrawOn_() });
}

function previewTicket_(ticket) {
  const t = String(ticket || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 64);
  const raw = t ? CacheService.getScriptCache().get('pv-' + t) : null;
  if (!raw) throw new Error('This preview has run out. Open it again from Admin › System.');
  const pv = JSON.parse(raw);
  if (!adminSession_(pv.token)) throw new Error('Please sign in to Admin again, then open the preview from Admin › System.');
  return pv;
}

/** The preview pages: who to show. My Climb uses token + id exactly like its own staff preview. */
function clashPreviewOpen(ticket) {
  const pv = previewTicket_(ticket);
  return JSON.stringify({ token: pv.token, id: pv.asId, squad: pv.squad, name: pv.name, real: pv.real });
}

/** Generals HQ for the previewed student's squad (the real HQ: what you do there is saved). */
function clashPreviewHq(ticket) {
  const pv = previewTicket_(ticket);
  return JSON.stringify({ token: hqPreviewToken_(pv.squad), squad: pv.squad, name: 'Staff preview' });
}

/** Summit Party minigame tryout: the art and minigames, and which ones can be tried alone. */
function clashPreviewSummit(ticket) {
  const pv = previewTicket_(ticket), v = spView_(null);
  v.on = true;
  v.squad = pv.squad;
  v.tryout = Object.keys(SP_MINIGAMES_).filter(function (k) { return k !== 'chase' && k !== 'sketch'; });
  return JSON.stringify(v);
}

/** Summit Party demo game: the previewed squad plays; the computer plays the other three. No points. */
function clashSummitDemo(ticket, rounds, layout) {
  const pv = previewTicket_(ticket);
  return spCreate_('Preview (' + pv.name + ')', Math.max(1, Math.min(10, Math.round(Number(rounds) || 3))), layout, { mine: pv.squad });
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_65_Preview = '';

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_65_Preview = 'a7912929d4';
