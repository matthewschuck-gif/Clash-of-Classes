/**
 * Clash of Classes -- Apps Script backend
 * File 46: Plinko. A big board for the TV or an assembly: clashofclasses.org/plinko (?page=plinko).
 *
 * PLINKO SLOTS tab (edit any time, like the Wheel): Order | Label | Kind | Value | Color
 *   One row per slot along the bottom, left to right (PLINKO_ROWS + 1 slots; 8 rows = 9 slots).
 *   Kind: points (Value points to the squad) · mystery (one of the Values, e.g. 5|10|25|50, picked at
 *   random) · nothing (no points: "So close!"). Color: any CSS color, blank = the board's colors.
 *   Like a real board, the middle slots catch most drops and the edges are rare (8 rows: the two
 *   edge slots together come up about 1 drop in 128), so put the big prizes at the edges.
 * PLINKO DROPS tab: Squad | Earned | Reason | Used | Slot | Prize | Points
 *   A row = one drop a squad has earned. Staff add drops in Admin › 📋 Sheet, or with the + Drop
 *   button on the board (admin password). The board shows each squad's waiting drops.
 * The drop: staff (signed in on the board) press a squad's Drop button. The server decides the path
 * (left or right at every row, 50/50, like a real board) so the result is fair and saved before the
 * ball moves; the page animates that exact path. The prize goes on the events tab (type plinko).
 * Practice drops (no prize) can run any time.
 */

const PLINKO_SLOT_HEADERS_ = ['Order', 'Label', 'Kind', 'Value', 'Color'];
const PLINKO_DROP_HEADERS_ = ['Squad', 'Earned', 'Reason', 'Used', 'Slot', 'Prize', 'Points'];
const PLINKO_SEED_ = [
  [1, 'JACKPOT 100', 'points', 100, '#ffd23f'],
  [2, '40', 'points', 40, '#b583ff'],
  [3, '15', 'points', 15, '#7dd3fc'],
  [4, '5', 'points', 5, '#9aa5b8'],
  [5, 'MYSTERY', 'mystery', '5|10|25|50', '#4ade80'],
  [6, '5', 'points', 5, '#9aa5b8'],
  [7, '15', 'points', 15, '#7dd3fc'],
  [8, '40', 'points', 40, '#b583ff'],
  [9, 'JACKPOT 100', 'points', 100, '#ffd23f'],
];

function ensurePlinkoTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [['Plinko Slots', PLINKO_SLOT_HEADERS_, PLINKO_SEED_], ['Plinko Drops', PLINKO_DROP_HEADERS_, null]].forEach(function (t) {
    if (ss.getSheetByName(t[0])) return;
    const sh = ss.insertSheet(t[0]);
    sh.getRange(1, 1, 1, t[1].length).setValues([t[1]]).setFontWeight('bold').setBackground('#490e6f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (t[2]) sh.getRange(2, 1, t[2].length, t[1].length).setValues(t[2]);
  });
}

function plinkoRows_() { return Math.max(4, Math.min(12, settingNum_('PLINKO_ROWS', 8))); }

/** The slots, padded or trimmed to rows + 1. */
function plinkoSlots_() {
  const n = plinkoRows_() + 1;
  const list = sheetRows_('Plinko Slots').filter(function (r) { return String(r['Label'] || '').trim() || String(r['Value'] || '').trim(); })
    .sort(function (a, b) { return (Number(a['Order']) || 0) - (Number(b['Order']) || 0); })
    .map(function (r) {
      const kind = String(r['Kind'] || 'points').trim().toLowerCase();
      return { label: String(r['Label'] || r['Value'] || '').trim(), kind: /myst/.test(kind) ? 'mystery' : /noth|none|zero/.test(kind) ? 'nothing' : 'points',
        value: String(r['Value'] == null ? '' : r['Value']).trim(), color: String(r['Color'] || '').trim() };
    });
  // No slots on the tab yet (before Run setup): the starting board.
  if (!list.length) PLINKO_SEED_.forEach(function (r) { list.push({ label: r[1], kind: r[2], value: String(r[3]), color: r[4] }); });
  while (list.length < n) list.push({ label: '5', kind: 'points', value: '5', color: '' });
  return list.slice(0, n);
}

function plinkoWaiting_() {
  const out = { BLACK: 0, GOLD: 0, GREY: 0, PURPLE: 0 };
  sheetRows_('Plinko Drops').forEach(function (r) { const s = squadOf_(r['Squad']); if (out[s] !== undefined && !String(r['Used'] || '').trim()) out[s]++; });
  return out;
}

/** The board for the page (public: slots, waiting drops, the last few drops). */
function clashPlinkoBoard() {
  const recent = sheetRows_('Plinko Drops').filter(function (r) { return String(r['Used'] || '').trim(); })
    .sort(function (a, b) { return String(b['Used']).localeCompare(String(a['Used'])); }).slice(0, 6)
    .map(function (r) { return { squad: squadOf_(r['Squad']), prize: String(r['Prize'] || ''), points: Number(r['Points']) || 0 }; });
  return JSON.stringify({ rows: plinkoRows_(), slots: plinkoSlots_(), waiting: plinkoWaiting_(), recent: recent });
}

// Left or right at every row, 50/50. The slot = how many times it went right.
function plinkoPath_(rows) { const p = []; for (let i = 0; i < rows; i++) p.push(Math.random() < 0.5 ? 0 : 1); return p; }

function plinkoPrize_(slot) {
  if (slot.kind === 'nothing') return { label: slot.label || 'So close!', points: 0 };
  if (slot.kind === 'mystery') {
    const opts = String(slot.value || '5|10|25').split(/[|,]/).map(Number).filter(function (n) { return n > 0; });
    const pts = opts.length ? opts[Math.floor(Math.random() * opts.length)] : 10;
    return { label: 'Mystery: ' + pts, points: pts };
  }
  const pts = Math.max(0, Number(slot.value) || Number(String(slot.label).replace(/\D/g, '')) || 0);
  return { label: slot.label || String(pts), points: pts };
}

/** Staff drop one of a squad's waiting drops (practice = no prize, no drop used). */
function clashPlinkoDrop(token, squad, practice) {
  requireAdmin_(token);
  squad = squadOf_(squad);
  const rows = plinkoRows_(), path = plinkoPath_(rows);
  const slotIdx = path.reduce(function (a, b) { return a + b; }, 0), slot = plinkoSlots_()[slotIdx];
  if (practice) return JSON.stringify({ path: path, slot: slotIdx, prize: plinkoPrize_(slot).label, points: 0, practice: true });
  if (SQUAD_KEYS_.indexOf(squad) === -1) throw new Error('Pick a squad.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const drop = sheetRows_('Plinko Drops').filter(function (r) { return squadOf_(r['Squad']) === squad && !String(r['Used'] || '').trim(); })[0];
    if (!drop) throw new Error(titleSquad_(squad) + ' has no drop waiting. Add one with + Drop.');
    const prize = plinkoPrize_(slot), sh = getSheet_('Plinko Drops');
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    const set = function (h, v) { const c = headerCol_(sh, h); if (c) sh.getRange(drop._row, c).setValue(v); };
    set('Used', stamp); set('Slot', slotIdx + 1); set('Prize', prize.label); set('Points', prize.points);
    if (prize.points > 0) {
      const pts = { BLACK: '', GOLD: '', GREY: '', PURPLE: '' }; pts[squad] = prize.points;
      getSheet_('events').appendRow(['', 'Plinko: ' + prize.label, 'plinko', todayYmd_(), String(drop['Reason'] || ''), '', pts.BLACK, pts.GOLD, pts.GREY, pts.PURPLE]);
    }
    return JSON.stringify({ path: path, slot: slotIdx, prize: prize.label, points: prize.points, squad: squad, waiting: plinkoWaiting_() });
  } finally { lock.releaseLock(); }
}

/** Staff give a squad drops (from the board). */
function clashPlinkoGrant(token, squad, reason, n) {
  requireAdmin_(token);
  squad = squadOf_(squad);
  if (SQUAD_KEYS_.indexOf(squad) === -1) throw new Error('Pick a squad.');
  n = Math.max(1, Math.min(10, Number(n) || 1));
  ensurePlinkoTabs_();
  const sh = getSheet_('Plinko Drops'), today = todayYmd_();
  for (let i = 0; i < n; i++) sh.appendRow([squad, today, String(reason || 'Staff').slice(0, 120), '', '', '', '']);
  return JSON.stringify({ waiting: plinkoWaiting_() });
}

/** Plinko odds for each slot (for staff placing prizes): [{ label, percent }]. */
function plinkoOdds_() {
  const rows = plinkoRows_(), slots = plinkoSlots_(), total = Math.pow(2, rows);
  const choose = function (n, k) { let r = 1; for (let i = 1; i <= k; i++) r = r * (n - i + 1) / i; return r; };
  return slots.map(function (s, k) { return { label: s.label, percent: Math.round(10000 * choose(rows, k) / total) / 100 }; });
}

function servePlinko_() {
  return HtmlService.createHtmlOutputFromFile('Plinko').setTitle('Plinko · Clash of Classes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// file-stamp (tools/stamp-files.js): lets Admin › System tell whether this file is the newest one
var FILE_STAMP_46_Plinko = 'ad5ea7680f';
