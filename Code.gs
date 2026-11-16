// ============================================================
//  Medical Center Dashboard - Code.gs  (v6: Today settings, Sync Now, editable labels, Reports access rules)
// ============================================================
var _CL = null;   // center list (loaded from Settings, editable in the dashboard)
var TZ = 'Asia/Dhaka';
var DEFAULT_PASSWORD = 'admin';
var CACHE_TTL = 21600;
var SPREADSHEET_ID = '11YYc6-98vtihGSDmymEOVZFx57jcRaZpBsNnO8dySUs';   // only needed if this script is NOT created from Extensions > Apps Script inside the sheet

var REPORT_DEF = {
  agentCol: 'J',        // Agent Wise: agent name
  agentDateCol: 'AP',   // Agent Wise: date used for the From / To filter
  regCol: 'AP',         // Registration Monthly: MMR1 date
  wafidCol: 'AO',       // Wafid Online Monthly: Wafid Online date
  centerCol: 'AR',      // center name column
  statusCol: 'AM',      // Status filter: FINAL STATUS column
  labels: {
    agent: 'Agent Wise Report', reg: 'Registration Monthly Summary', wafid: 'Wafid Online Monthly Summary',
    title_agent: '', title_reg: '', title_wafid: '',
    agentHead: 'Agent', monthHead: 'Year-Month', grand: 'Grand Total',
    allCenters: 'All Centers', allAgents: 'All Agents', all: 'All',
    center: 'Center', from: 'From', to: 'To', agentLbl: 'Agent', status: 'Status',
    search: 'Search', print: 'Print', pdf: 'PDF A4', excel: 'Excel',
    empty: 'Select filters and press Search to see the report', noData: 'No data'
  }
};

var DEF = {
  title: 'MEDICAL CENTER MANAGEMENT DASHBOARD',
  refresh: 10,          // seconds (Settings > General, minimum 5)
  firstRow: 6,
  centerList: ['AJBAHUR','AL MAHMUD','AMIR JAHAN','ANAS','PHOENIX','TANJIM','TAQDIR'],
  cols: {total:'G', wafid:'AO', transfer:'AT', phys:'AS', gender:'F', country:'H', center:'B', wafidCenter:'AR'},
  group: {name:'AL JAHAN GROUP',
    address:'MK Tower, 5th Floor, 12 Progoti Shoroni, Block-J, Baridhara, Dhaka-1212, Bangladesh',
    phone:'+88 01888 000 666', email:'info@aljahanint.com', web:'https://www.aljahanint.com', logo:''},
  centers: {
    'AJBAHUR': {name:'Ajbahur Medical Center',
      address:'Shah Amanat Tower (1st Floor), Kaptai Raster Matha, Mohora, Chandgaon, Chattogram',
      phone:'+8801810020464', email:'ajbahurmedicalcenter@gmail.com', web:'https://ajbahurmedicalcenter.com', logo:''},
    'AL MAHMUD': {name:'Al Mahmud Medical Center',
      address:'Grameen Banglar Akshay Tower, 51 Madani Avenue, Natun Bazar (100 Feet Road), Vatara, Dhaka',
      phone:'+880248814712', email:'almahmudmedicalcenter@gmail.com', web:'https://almahmudmedicalcenter.com', logo:''},
    'AMIR JAHAN': {name:'Amir Jahan Medical Center',
      address:'Anchor Tower, 2nd Floor, 108 Bir Uttam C R Datta Road, (Beside of Bangla Vision TV) Kawran Bazar, Dhaka',
      phone:'+8801714177881', email:'amirjmc3@gmail.com', web:'http://www.amirjahanmedicalcenter.com', logo:''},
    'ANAS': {name:'Anas Medical Center',
      address:'60/E/1, Dewan Complex (4th Floor) Purana Paltan, Dhaka-1000, Bangladesh',
      phone:'+8801791375444', email:'anasmedicalcentre@gmail.com', web:'https://anasmedicalcenter.com', logo:''},
    'PHOENIX': {name:'Phoenix Medical Center',
      address:'Grameen Banglar Akshay Tower, 51 Madani Avenue, Natun Bazar (100 Feet Road), Vatara, Dhaka',
      phone:'+8801744164775', email:'medicalcenterphoenix@gmail.com', web:'http://phoenixmedicalbd.com', logo:''},
    'TANJIM': {name:'Tanjim Medical Center',
      address:'Pervez Tower (1st Floor), South side of Rangmahal, Sajbarkhila, Sherpur City, Sherpur-2100.',
      phone:'+8801810020465', email:'tanjimmedicalcenter@gmail.com', web:'https://tanjimmedicalcenter.com', logo:''},
    'TAQDIR': {name:'Taqdir Medical Center',
      address:'Najma Plaza (2nd Floor), Chankhar Dokan, Ashikathi 02 No. Union, (Cumilla Road) Chandpur Sadar, Chandpur.',
      phone:'+8801810020454', email:'taqdirmedicalcenter@gmail.com', web:'https://www.taqdirmedicalcenter.com', logo:''}
  },
  labels: {
    today: 'TODAY',
    menu: ['Dashboard','Reports','Settings'],
    cards: ['TOTAL FILE','WAFID ONLINE COMPLETED','SOFTWARE TRANSFER COMPLETE','FILE COMPLETE','TRANSFER PENDING','FILE PENDING'],
    charts: {g:'Gender Report', w:'Wafid Online Report', cw:'Country Wise Report', f:'File Complete & Pending', co:'Center Wise Report'}
  },
  reports: REPORT_DEF,
  colors: {on:false, accent:'#1565ff', sidebar:'#0a2f6b', header:'#0a2a63', bg:'#eaf2fd',
    cards:['#1565ff','#15a34a','#6d28d9','#0d9488','#ef4444','#f59e0b'],
    chart:['#1565ff','#16a34a','#f59e0b','#8b5cf6','#14b8a6','#ec4899','#1e3a8a','#ef4444'],
    rp:{bg:'#070d24', side:'#0f3fb8', menu:'#2f6bff', fcard:'#0d1b4c', rcard:'#0d1b4c', tab:'#142a66', tabon:'#2f6bff', field:'#091338', head:'#14306e', grand:'#14306e', text:'#ffffff'}}
};

// ---------- Web app / menu ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Medical Center Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Dashboard')
    .addItem('Open Dashboard', 'openDash')
    .addItem('Sync now', 'syncMenu')
    .addToUi();
}
function openDash() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('Index').setWidth(1400).setHeight(800), 'Dashboard');
}
function syncMenu() {
  repBump_();
  refreshAll();
  try { SpreadsheetApp.getActive().toast('Synced', 'Dashboard', 5); } catch (e) { Logger.log('Synced'); }
}

// ---------- Password (salted SHA-256, default "admin") ----------
function hashPw_(pw) {
  var p = PropertiesService.getScriptProperties();
  var salt = p.getProperty('PW_SALT');
  if (!salt) { salt = Utilities.getUuid(); p.setProperty('PW_SALT', salt); }
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pw, Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}
function ok(pw) {
  if (!pw) return false;
  var h = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH');
  if (!h) return pw === DEFAULT_PASSWORD;
  return hashPw_(pw) === h;
}
function checkPassword(pw) { return ok(pw); }
function isDefaultPassword() {
  var c = CacheService.getScriptCache(), v = c.get('DPW5');
  if (v === null) { v = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH') ? '0' : '1'; c.put('DPW5', v, CACHE_TTL); }
  return v === '1';
}
function changePassword(oldPw, newPw) {
  if (!ok(oldPw)) throw new Error('Wrong password');
  newPw = String(newPw || '');
  if (newPw.length < 6) throw new Error('New password must be at least 6 characters');
  if (newPw === DEFAULT_PASSWORD) throw new Error('Choose a different password');
  PropertiesService.getScriptProperties().setProperty('ADMIN_HASH', hashPw_(newPw));
  CacheService.getScriptCache().remove('DPW5');
  return true;
}

// ---------- Settings ----------
function mergeDef_(d, o) {
  if (o === undefined || o === null) return d;
  if (Array.isArray(d)) return d.map(function (x, i) { return (Array.isArray(o) && o[i] !== undefined) ? o[i] : x; });
  if (d && typeof d === 'object') { var r = {}; for (var k in d) r[k] = mergeDef_(d[k], o[k]); return r; }
  return o;
}
var BLANK_ENT = {name:'', address:'', phone:'', email:'', web:'', logo:''};
function cleanList_(a, def) {
  if (!Array.isArray(a) || !a.length) return def.slice();
  var seen = {}, out = [];
  a.forEach(function (x) {
    var k = String(x || '').toUpperCase().replace(/\s+/g, ' ').replace(/[^A-Z0-9 \-_.]/g, '').trim();
    if (k && k !== 'ALL' && !seen[k] && out.length < 30) { seen[k] = 1; out.push(k); }
  });
  return out.length ? out : def.slice();
}
// blank text -> default text (arrays and objects handled one by one)
function fillBlank_(d, o) {
  if (Array.isArray(d)) return d.map(function (x, i) { return fillBlank_(x, Array.isArray(o) ? o[i] : undefined); });
  if (d && typeof d === 'object') {
    var r = {};
    for (var k in d) r[k] = fillBlank_(d[k], (o && typeof o === 'object') ? o[k] : undefined);
    return r;
  }
  var v = String(o === undefined || o === null ? '' : o).replace(/\s+/g, ' ').trim().slice(0, 120);
  return v || d;
}
// Editable names for every other text of the pages: {key: text}. A blank text is dropped (the page then uses its default).
function cleanUi_(u) {
  var o = {}, n = 0;
  if (u && typeof u === 'object') {
    for (var k in u) {
      if (!/^[A-Za-z0-9_]{1,40}$/.test(k) || n >= 500) continue;
      var v = String(u[k] === undefined || u[k] === null ? '' : u[k]).replace(/\s+/g, ' ').trim().slice(0, 120);
      if (v) { o[k] = v; n++; }
    }
  }
  return o;
}
// TODAY cards: which Excel column each card reads, and which value counts as "complete".
//   total / wafid : rows whose date in that column is today
//   transfer / file : among today's rows, the column equals the value (COMPLETE)
// If nothing was saved yet, the General > Sheet columns are used.
function todayClean_(t, cols) {
  t = (t && typeof t === 'object') ? t : {};
  var dflt = {total: cols.wafid, wafid: cols.wafid, transfer: cols.transfer, file: cols.phys}, o = {};
  ['total', 'wafid', 'transfer', 'file'].forEach(function (k) {
    var e = (t[k] && typeof t[k] === 'object') ? t[k] : {};
    var col = String(e.col || '').toUpperCase().replace(/[^A-Z]/g, '').slice(0, 3) || dflt[k];
    var val = '';
    if (k === 'transfer' || k === 'file') val = String(e.val === undefined || e.val === null ? '' : e.val).replace(/\s+/g, ' ').trim().slice(0, 60) || 'COMPLETE';
    o[k] = {col: col, val: val};
  });
  return o;
}
function normVal_(v) { return String(v === undefined || v === null ? '' : v).trim().toUpperCase(); }
function CL_() { if (!_CL) _CL = getSettings().centerList; return _CL; }
function getSettings() {
  var cc = CacheService.getScriptCache().get('SET9');
  if (cc) { try { return JSON.parse(cc); } catch (e) {} }
  var o = getSettings_();
  try { CacheService.getScriptCache().put('SET9', JSON.stringify(o), CACHE_TTL); } catch (e) {}
  return o;
}
function getSettings_() {
  var raw = getBig_('SETB');
  if (!raw) raw = PropertiesService.getScriptProperties().getProperty('SETTINGS');
  var s = {}; try { s = raw ? JSON.parse(raw) : {}; } catch (e) { s = {}; }
  var o = mergeDef_(DEF, s);
  o.refresh = Math.max(5, Number(o.refresh) || DEF.refresh);
  o.firstRow = Math.max(1, Number(o.firstRow) || DEF.firstRow);
  o.centerList = cleanList_(s.centerList, DEF.centerList);
  var sc = (s.centers && typeof s.centers === 'object') ? s.centers : {}, cs = {};
  o.centerList.forEach(function (k) { cs[k] = mergeDef_(DEF.centers[k] || BLANK_ENT, sc[k]); });
  o.centers = cs;
  o.reports = repClean_(o.reports);
  o.title = String(o.title || '').replace(/\s+/g, ' ').trim() || DEF.title;
  o.labels = fillBlank_(DEF.labels, s.labels);
  o.today = todayClean_(s.today, o.cols);
  o.ui = cleanUi_(s.ui);
  return o;
}
function saveConfig(pw, cfg) {
  if (!ok(pw)) throw new Error('Wrong password');
  cfg = cfg || {};
  var old = getSettings_();
  var list = cleanList_(cfg.centerList, old.centerList);
  var ce = (cfg.centers && typeof cfg.centers === 'object') ? cfg.centers : {};
  var s = mergeDef_(DEF, cfg);
  s.centerList = list;
  s.centers = {};
  list.forEach(function (k) { s.centers[k] = mergeDef_(DEF.centers[k] || BLANK_ENT, ce[k]); });
  s.refresh = Math.min(3600, Math.max(5, Number(s.refresh) || DEF.refresh));
  s.firstRow = Math.max(1, Math.floor(Number(s.firstRow) || DEF.firstRow));
  for (var k in s.cols) {
    s.cols[k] = String(s.cols[k] || DEF.cols[k]).toUpperCase().replace(/[^A-Z]/g, '') || DEF.cols[k];
  }
  s.reports = repClean_(s.reports);
  s.title = String(s.title || '').replace(/\s+/g, ' ').trim() || DEF.title;
  s.labels = fillBlank_(DEF.labels, cfg.labels);
  s.today = todayClean_(cfg.today, s.cols);
  s.ui = cleanUi_(cfg.ui);
  var hex = /^#[0-9a-fA-F]{6}$/;
  if (s.colors && s.colors.rp) for (var q in DEF.colors.rp) if (!hex.test(String(s.colors.rp[q]))) s.colors.rp[q] = DEF.colors.rp[q];
  var json = jsonA_(s);
  if (json.length > 40000) throw new Error('Settings too large (' + json.length + ' characters). Shorten some text or logo links.');
  setBig_('SETB', json);
  PropertiesService.getScriptProperties().deleteProperty('SETTINGS');
  CacheService.getScriptCache().remove('SET9');
  _CL = null;
  if (JSON.stringify(old.cols) !== JSON.stringify(s.cols) || old.firstRow !== s.firstRow ||
      JSON.stringify(old.centerList) !== JSON.stringify(s.centerList) ||
      JSON.stringify(old.today) !== JSON.stringify(s.today)) clearTabs_(old.centerList.concat(s.centerList));
  return getSettings();
}
// Restores titles, columns, centers, labels and colors. Password and profile are NOT changed.
function resetSettings(pw) {
  if (!ok(pw)) throw new Error('Wrong password');
  var old = getSettings_();
  delBig_('SETB');
  PropertiesService.getScriptProperties().deleteProperty('SETTINGS');
  CacheService.getScriptCache().remove('SET9');
  _CL = null;
  clearTabs_(old.centerList.concat(DEF.centerList));
  return getSettings();
}
// ---------- Per-user prefs ----------
function normC_(c) {
  c = String(c || 'ALL').toUpperCase();
  return (c === 'ALL' || CL_().indexOf(c) > -1) ? c : 'ALL';
}
// Themes: only 'colorful' and 'dark'. Old 'black' becomes 'dark'; anything else (white, darkgray, system, unknown) becomes 'colorful'.
function normTheme_(t) {
  t = String(t || '').toLowerCase();
  return (t === 'dark' || t === 'black') ? 'dark' : 'colorful';
}
function saveTheme(t) { PropertiesService.getUserProperties().setProperty('THEME', normTheme_(t)); }
function saveCenter(c) { c = normC_(c); PropertiesService.getUserProperties().setProperty('CENTER', c); return c; }

// One call at start: settings + data together. A sheet error never blanks the page: settings still load and the error is returned as text.
function getInit(center, theme) {
  var c = center ? normC_(center) : '', th = theme || '';
  if (!c || !th) {
    var up = PropertiesService.getUserProperties().getProperties();
    if (!c) c = normC_(up.CENTER);
    if (!th) th = up.THEME || 'colorful';
  }
  th = normTheme_(th);
  var out = {
    boot: {s: getSettings(), theme: th, center: c, centers: CL_(), profile: getProfile(), defaultPw: isDefaultPassword()},
    data: null, error: ''
  };
  try { out.data = getData(c); } catch (e) { out.error = String(e && e.message || e); }
  return out;
}

// ---------- Data (fast: per-center results are pre-computed and cached) ----------
function getSS_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) {}
  if (!ss) {
    var id = SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty('SS_ID');
    if (id) ss = SpreadsheetApp.openById(id);
  }
  if (!ss) throw new Error('Spreadsheet not found. Open the sheet > Extensions > Apps Script, or put the Sheet ID in SPREADSHEET_ID at the top of Code.gs');
  return ss;
}
function ci(l) {
  var n = 0; l = String(l).toUpperCase();
  for (var i = 0; i < l.length; i++) n = n * 26 + l.charCodeAt(i) - 64;
  return n - 1;
}
var MON_ = {jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12};
function p2_(x) { return ('0' + x).slice(-2); }
// Any date -> 'yyyy-MM-dd' (Bangladesh time for real dates). Only the date counts, never the time.
// Understands: real date cells, 2026-11-13, 13/11/2026, 13-11-2026 10:30, 13-Nov-2026, 13 November 2026, Nov 13, 2026, 13/11/26, Excel serial numbers.
function dk(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) return Utilities.formatDate(new Date(Math.round((v - 25569) * 86400000)), 'UTC', 'yyyy-MM-dd');
    return '';
  }
  v = String(v === undefined || v === null ? '' : v).trim();
  if (!v) return '';
  var y = v.match(/(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (y) return y[1] + '-' + p2_(y[2]) + '-' + p2_(y[3]);
  var m = v.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
  if (m) return m[3] + '-' + p2_(m[2]) + '-' + p2_(m[1]);
  var a = v.match(/(\d{1,2})(?:st|nd|rd|th)?[\s\-\/.,]*([A-Za-z]{3,9})[\s\-\/.,]*(\d{4})/);
  if (a && MON_[a[2].slice(0, 3).toLowerCase()]) return a[3] + '-' + p2_(MON_[a[2].slice(0, 3).toLowerCase()]) + '-' + p2_(a[1]);
  var b = v.match(/([A-Za-z]{3,9})[\s\-\/.,]*(\d{1,2})(?:st|nd|rd|th)?[\s\-\/.,]*(\d{4})/);
  if (b && MON_[b[1].slice(0, 3).toLowerCase()]) return b[3] + '-' + p2_(MON_[b[1].slice(0, 3).toLowerCase()]) + '-' + p2_(b[2]);
  var s = v.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2})(?!\d)/);
  if (s) return '20' + s[3] + '-' + p2_(s[2]) + '-' + p2_(s[1]);
  return '';
}
function sn(v) {
  v = String(v).trim().toUpperCase();
  if (!v) return 'UNKNOWN';
  var cl = CL_();
  for (var i = 0; i < cl.length; i++) if (v.indexOf(cl[i]) > -1) return cl[i];
  return v;
}
function normName_(s) { return String(s || '').replace(/\s+/g, ' ').trim().toUpperCase(); }
function centerOf_(sheetName) {
  var n = normName_(sheetName);
  return CL_().indexOf(n) > -1 ? n : null;
}
function findSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) if (normName_(all[i].getName()) === name) return all[i];
  return null;
}
function emptyTab_(today) {
  return {day: today, total:0, wafid:0, transfer:0, phys:0, male:0, female:0,
          t:{total:0, wafid:0, transfer:0, phys:0}, country:{}, byB:{}, byW:{}, miss:false, ts: Date.now()};
}
// Reads the needed columns with the fewest possible calls (nearby columns are fetched together). cols = {key: 'AO', ...}
function readCols_(sh, cols, firstRow, r) {
  var keys = Object.keys(cols), lastC = sh.getLastColumn(), list = [], k, i;
  for (i = 0; i < keys.length; i++) {
    k = keys[i];
    var x = ci(cols[k]);
    if (x < lastC) list.push({k: k, i: x});
  }
  list.sort(function (a, b) { return a.i - b.i; });
  var groups = [];
  list.forEach(function (o) {
    var g = groups[groups.length - 1];
    if (g && o.i - g.end <= 10) { g.end = o.i; g.items.push(o); }
    else groups.push({start: o.i, end: o.i, items: [o]});
  });
  var col = {};
  groups.forEach(function (g) {
    var v = sh.getRange(firstRow, g.start + 1, r, g.end - g.start + 1).getValues();
    g.items.forEach(function (o) { col[o.k] = {v: v, off: o.i - g.start}; });
  });
  return col;
}
// TODAY cards (Settings > General > TODAY section). Everything is counted in ONE pass over the rows:
//   TOTAL FILE (Today)              = rows whose date in the "total" column is today
//   WAFID ONLINE COMPLETED (Today)  = rows whose date in the "wafid" column is today
//   SOFTWARE TRANSFER COMPLETE      = today's rows where the "transfer" column = its value (COMPLETE)
//   FILE COMPLETE                   = today's rows where the "file" column = its value (COMPLETE)
//   TRANSFER PENDING / FILE PENDING = TOTAL FILE (Today) minus the COMPLETE count (done in the page)
// Only the date is compared (never the time), in Bangladesh time.
function computeTab_(name, S, today, ss) {
  var R = emptyTab_(today);
  var sh = findSheet_(ss || getSS_(), name);
  if (!sh) { R.miss = true; return R; }
  var r = sh.getLastRow() - S.firstRow + 1;
  if (r < 1) return R;
  var TD = S.today || todayClean_(null, S.cols);
  var spec = {}, k;
  for (k in S.cols) spec[k] = S.cols[k];
  spec.tdTotal = TD.total.col; spec.tdWafid = TD.wafid.col; spec.tdTr = TD.transfer.col; spec.tdFi = TD.file.col;
  var col = readCols_(sh, spec, S.firstRow, r);
  var trv = normVal_(TD.transfer.val) || 'COMPLETE', fiv = normVal_(TD.file.val) || 'COMPLETE';
  function g(c, i) { return c ? c.v[i][c.off] : ''; }
  // memo: the same date / center text repeats on thousands of rows, so each one is converted only once
  var dm = {}, sm = {};
  function dkm(v) {
    var key = (v instanceof Date) ? 't' + v.getTime() : 's' + v, x = dm[key];
    if (x === undefined) { x = dk(v); dm[key] = x; }
    return x;
  }
  function snm(v) {
    var key = String(v), x = sm[key];
    if (x === undefined) { x = sn(v); sm[key] = x; }
    return x;
  }
  function isToday(v) { return v !== '' && v !== null && v !== undefined && dkm(v) === today; }
  var cT = col.total, cTr = col.transfer, cPh = col.phys, cW = col.wafid,
      cB = col.center, cWC = col.wafidCenter, cG = col.gender, cC = col.country,
      tT = col.tdTotal, tW = col.tdWafid, tTr = col.tdTr, tFi = col.tdFi;
  for (var i = 0; i < r; i++) {
    if (String(g(cT, i)).trim() === '') continue;
    var tr = String(g(cTr, i)).trim().toUpperCase() === 'COMPLETE';
    var ph = String(g(cPh, i)).trim().toUpperCase() === 'COMPLETE';
    var wv = g(cW, i);
    var wd = String(wv).trim() !== '';
    var b = snm(g(cB, i));
    var y = R.byB[b] = R.byB[b] || {total:0, wafid:0, transfer:0, phys:0};
    R.total++; y.total++;
    if (wd) { R.wafid++; y.wafid++; var q = snm(g(cWC, i)); R.byW[q] = (R.byW[q] || 0) + 1; }
    if (tr) { R.transfer++; y.transfer++; }
    if (ph) { R.phys++; y.phys++; }
    // today
    if (isToday(g(tT, i))) {
      R.t.total++;
      if (normVal_(g(tTr, i)) === trv) R.t.transfer++;
      if (normVal_(g(tFi, i)) === fiv) R.t.phys++;
    }
    if (isToday(g(tW, i))) R.t.wafid++;
    var gd = String(g(cG, i)).trim().toUpperCase().charAt(0);
    if (gd === 'M') R.male++; else if (gd === 'F') R.female++;
    var c = String(g(cC, i)).trim().toUpperCase();
    if (c) R.country[c] = (R.country[c] || 0) + 1;
  }
  return R;
}
function storeTab_(name, R, persist) {
  var json = JSON.stringify(R), cache = CacheService.getScriptCache();
  if (json.length < 90000) cache.put('T7_' + name, json, CACHE_TTL);
  if (persist && json.length < 8800) {
    PropertiesService.getScriptProperties().setProperty('TAB7_' + name, json);
    cache.put('PW_' + name, '1', 60);
  }
}
function refreshTab_(name, S, ss) {
  var start = Date.now(), cache = CacheService.getScriptCache();
  cache.put('LC_' + name, String(start), 120);
  var R = computeTab_(name, S || getSettings(), dk(new Date()), ss);
  storeTab_(name, R, !cache.get('PW_' + name));
  return R;
}
function getTab_(name, pre, fast) {
  var cache = CacheService.getScriptCache(), today = dk(new Date()), R = null;
  var c = pre !== undefined ? pre : cache.get('T7_' + name);
  if (c) { try { R = JSON.parse(c); } catch (e) { R = null; } }
  if (!R) {
    var p = PropertiesService.getScriptProperties().getProperty('TAB7_' + name);
    if (p) { try { R = JSON.parse(p); cache.put('T7_' + name, p, CACHE_TTL); } catch (e) { R = null; } }
  }
  // fast mode: triggers keep data fresh. Without it, data older than 20 seconds is rebuilt on demand.
  var stale = !fast && R && (Date.now() - (R.ts || 0) > 20000);
  if (!R || (R.day !== today && !fast) || stale) R = refreshTab_(name);
  else if (R.day !== today) R.t = {total:0, wafid:0, transfer:0, phys:0};   // new day: show 0 until the trigger rebuilds (max 10 min)
  return R;
}
function addMap_(a, b) { for (var k in b) a[k] = (a[k] || 0) + b[k]; }
function getData(center) {
  center = normC_(center);
  var names = center === 'ALL' ? CL_() : [center];
  var R = {total:0, wafid:0, transfer:0, phys:0, male:0, female:0,
           t:{total:0, wafid:0, transfer:0, phys:0}, country:{}, tabs:{}, byB:{}, byW:{}, miss:[], ts:0};
  var keys = names.map(function (n) { return 'T7_' + n; });
  var hit = CacheService.getScriptCache().getAll(keys);          // one cache call for all centers
  var fast = PropertiesService.getScriptProperties().getProperty('FAST') === '1';
  var oldest = Date.now();
  names.forEach(function (n) {
    var p = getTab_(n, hit['T7_' + n] || null, fast), f;
    ['total','wafid','transfer','phys','male','female'].forEach(function (x) { R[x] += p[x]; });
    for (f in R.t) R.t[f] += p.t[f];
    addMap_(R.country, p.country); addMap_(R.byW, p.byW);
    for (var b in p.byB) {
      var y = R.byB[b] = R.byB[b] || {total:0, wafid:0, transfer:0, phys:0};
      for (f in y) y[f] += p.byB[b][f];
    }
    R.tabs[n] = {total:p.total, wafid:p.wafid, transfer:p.transfer, phys:p.phys};
    if (p.miss) R.miss.push(n);
    if (p.ts && p.ts < oldest) oldest = p.ts;
  });
  R.ts = oldest;
  return R;
}

// ---------- Fast updates: triggers ----------
function clearTabs_(names) {
  var cache = CacheService.getScriptCache(), p = PropertiesService.getScriptProperties();
  repBump_(names || CL_());
  (names || CL_()).forEach(function (n) {
    cache.remove('T7_' + n); cache.remove('PW_' + n); p.deleteProperty('TAB7_' + n);
  });
}
function refreshAll() {
  var S = getSettings(), ss = getSS_();
  CL_().forEach(function (n) { refreshTab_(n, S, ss); });
  repWarm_(S, ss);
}
// Runs when a sheet is edited (installed trigger): rebuilds just that center
function onSheetEdit(e) {
  var t0 = Date.now();
  try {
    var name = null;
    if (e && e.range) {
      name = centerOf_(e.range.getSheet().getName());
      if (!name) return;                       // not a center sheet
      repBump_([name]);                        // report rows of this center must be re-read
    } else if (e && e.changeType === 'EDIT') {
      return;                                  // plain edits are handled by the onEdit trigger
    }
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) return;
    try {
      if (name) {
        var lc = Number(CacheService.getScriptCache().get('LC_' + name) || 0);
        if (lc > t0) return;                   // a newer rebuild already includes this edit
        refreshTab_(name);
      } else {
        repBump_();
        refreshAll();
      }
    } finally { lock.releaseLock(); }
  } catch (x) { /* ignore */ }
}
function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var h = t.getHandlerFunction();
    if (h === 'onSheetEdit' || h === 'refreshAll') ScriptApp.deleteTrigger(t);
  });
  var ss = getSS_(), p = PropertiesService.getScriptProperties();
  p.setProperty('SS_ID', ss.getId());
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('refreshAll').timeBased().everyMinutes(10).create();
  p.setProperty('FAST', '1');
}
// Sync Now button: reads the Excel sheet again right now and returns fresh numbers for the selected center.
function syncNow(center) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      repBump_();                                            // Reports rows are re-read on the next Search
      var S = getSettings(), ss = getSS_();
      CL_().forEach(function (n) { refreshTab_(n, S, ss); });
    } finally { lock.releaseLock(); }
  }
  return getData(center);
}

// ---------- Diagnose: run this from the Apps Script editor and read View > Logs ----------
function diagnose() {
  var S = getSettings_(), ss = getSS_();
  Logger.log('Spreadsheet: ' + ss.getName() + ' | first row: ' + S.firstRow);
  Logger.log('Columns: ' + JSON.stringify(S.cols));
  CL_().forEach(function (n) {
    var sh = findSheet_(ss, n);
    if (!sh) { Logger.log(n + ': SHEET NOT FOUND'); return; }
    var t0 = Date.now(), R = computeTab_(n, S, dk(new Date()), ss);
    Logger.log(n + ': rows=' + sh.getLastRow() + ' cols=' + sh.getLastColumn() +
      ' total=' + R.total + ' wafid=' + R.wafid + ' transfer=' + R.transfer + ' phys=' + R.phys +
      ' today=' + R.t.wafid + ' (' + (Date.now() - t0) + ' ms)');
  });
}

// ============================================================
//  PROFILE (sidebar photo, name, designation) + chunked property storage
// ============================================================
// ---- small helpers: store long text in several Script Properties ----
function jsonA_(o) {
  return JSON.stringify(o).replace(/[\u007f-\uffff]/g, function (c) {
    return '\\u' + ('0000' + c.charCodeAt(0).toString(16)).slice(-4);
  });
}
function setBig_(k, str) {
  var p = PropertiesService.getScriptProperties();
  var n = Math.max(1, Math.ceil(str.length / 8000));
  var old = Number(p.getProperty(k + '_N') || 0), o = {}, i;
  for (i = 0; i < n; i++) o[k + '_' + i] = str.substr(i * 8000, 8000);
  o[k + '_N'] = String(n);
  p.setProperties(o);
  for (i = n; i < old; i++) p.deleteProperty(k + '_' + i);
}
function getBig_(k) {
  var p = PropertiesService.getScriptProperties();
  var n = Number(p.getProperty(k + '_N') || 0), s = '';
  if (!n) return '';
  for (var i = 0; i < n; i++) s += p.getProperty(k + '_' + i) || '';
  return s;
}
function delBig_(k) {
  var p = PropertiesService.getScriptProperties();
  var n = Number(p.getProperty(k + '_N') || 0);
  for (var i = 0; i < n; i++) p.deleteProperty(k + '_' + i);
  p.deleteProperty(k + '_N');
}

// ---- profile: {name, desig, company, photo(data URL)} ----
function getProfile() {
  var cache = CacheService.getScriptCache(), v = cache.get('PROF1');
  if (v) { try { return JSON.parse(v); } catch (e) {} }
  var raw = PropertiesService.getScriptProperties().getProperty('PROFILE'), o = {};
  try { o = raw ? JSON.parse(raw) : {}; } catch (e) { o = {}; }
  var r = { name: o.name || '', desig: o.desig || '', company: o.company || '', photo: getBig_('PHOTO') };
  var j = JSON.stringify(r);
  if (j.length < 90000) { try { cache.put('PROF1', j, CACHE_TTL); } catch (e) {} }
  return r;
}

// pr.photo: undefined = keep current photo, '' = remove photo, 'data:image/...' = new photo
function saveProfile(pw, pr) {
  if (!ok(pw)) throw new Error('Wrong password');
  pr = pr || {};
  var name = String(pr.name || '').replace(/\s+/g, ' ').trim().slice(0, 60);
  var desig = String(pr.desig || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  var company = String(pr.company || '').replace(/\s+/g, ' ').trim().slice(0, 80);
  if (pr.photo !== undefined && pr.photo !== null) {
    var ph = String(pr.photo);
    if (ph === '') {
      delBig_('PHOTO');
    } else {
      if (ph.length > 70000) throw new Error('Photo is too large. Please choose a smaller image.');
      if (!/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(ph)) throw new Error('Unsupported photo data.');
      setBig_('PHOTO', ph);
    }
  }
  PropertiesService.getScriptProperties().setProperty('PROFILE', jsonA_({ name: name, desig: desig, company: company }));
  CacheService.getScriptCache().remove('PROF1');
  return getProfile();
}

// ============================================================
//  REPORTS (Agent Wise / Registration Monthly / Wafid Online Monthly)
//  Everything here is editable in Settings > Reports (columns, names) and Theme & colors (Step 5).
//  Speed: each center sheet is read once, reduced to a compact row list and cached.
//  A Search then only counts that list in memory (no sheet reads).
// ============================================================
function repClean_(r) {
  r = mergeDef_(REPORT_DEF, r || {});
  ['agentCol', 'agentDateCol', 'regCol', 'wafidCol', 'centerCol', 'statusCol'].forEach(function (k) {
    r[k] = String(r[k] || '').toUpperCase().replace(/[^A-Z]/g, '') || REPORT_DEF[k];
  });
  for (var k in REPORT_DEF.labels) {
    var v = r.labels[k];
    r.labels[k] = String(v === undefined || v === null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, 120) || REPORT_DEF.labels[k];
  }
  delete r.defFrom;
  return r;
}

var REP_MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function repFast_() { return PropertiesService.getScriptProperties().getProperty('FAST') === '1'; }
// Cache of compact rows per center. The key holds a version (changes when that sheet is edited) and the column settings.
function repVer_(names) {
  var keys = names.map(function (n) { return 'RRV_' + n; }), got = CacheService.getScriptCache().getAll(keys), o = {};
  names.forEach(function (n) { o[n] = got['RRV_' + n] || '0'; });
  return o;
}
function repBump_(names) {
  var c = CacheService.getScriptCache(), o = {}, t = String(Date.now());
  (names || CL_()).forEach(function (n) { o['RRV_' + n] = t; });
  c.putAll(o, CACHE_TTL);
}
function repBase_(name, ver, S, R) {
  return 'RR4_' + name + '_' + ver + '_' + [S.firstRow, S.cols.total, R.agentCol, R.agentDateCol, R.regCol, R.wafidCol, R.centerCol, R.statusCol].join('');
}
// Reads only the needed columns of one sheet and returns {ce, ag, st, rows}
// rows = "center,agent,status,regDate,wafidDate,agentDate;..." (dates as yyyymmdd, ids point into ce / ag / st)
function repExtract_(sh, S, R) {
  var O = {ce: [], ag: [], st: [], rows: ''};
  var r = sh.getLastRow() - S.firstRow + 1;
  if (r < 1) return O;
  var want = {total: ci(S.cols.total), center: ci(R.centerCol), agent: ci(R.agentCol), status: ci(R.statusCol),
              reg: ci(R.regCol), wafid: ci(R.wafidCol), ad: ci(R.agentDateCol)};
  var lastC = sh.getLastColumn(), idx = [], k;
  for (k in want) if (want[k] < lastC && idx.indexOf(want[k]) < 0) idx.push(want[k]);
  idx.sort(function (a, b) { return a - b; });
  var groups = [];
  idx.forEach(function (x) {
    var g = groups[groups.length - 1];
    if (g && x - g.end <= 10) g.end = x; else groups.push({start: x, end: x});
  });
  groups.forEach(function (g) { g.v = sh.getRange(S.firstRow, g.start + 1, r, g.end - g.start + 1).getValues(); });
  var m = {};
  for (k in want) groups.forEach(function (g) { if (want[k] >= g.start && want[k] <= g.end) m[k] = {v: g.v, off: want[k] - g.start}; });
  function at(key, i) { var c = m[key]; return c ? c.v[i][c.off] : ''; }
  var cm = {}, am = {}, sm = {}, dm = {}, out = [];
  function id(map, arr, key, label) {
    var x = map['k' + key];
    if (x === undefined) { x = arr.length; arr.push(label); map['k' + key] = x; }
    return x;
  }
  function dn(v) {
    if (v === '' || v === null || v === undefined) return '';
    var key = (v instanceof Date) ? 't' + v.getTime() : 's' + v, x = dm[key];
    if (x === undefined) { var d = dk(v); x = d ? d.replace(/-/g, '') : ''; dm[key] = x; }
    return x;
  }
  for (var i = 0; i < r; i++) {
    if (String(at('total', i)).trim() === '') continue;                       // not a file row
    var c = sn(at('center', i));
    var ag = String(at('agent', i)).replace(/\s+/g, ' ').trim();
    var st = String(at('status', i)).replace(/\s+/g, ' ').trim();
    out.push([id(cm, O.ce, c, c),
              ag ? id(am, O.ag, ag.toUpperCase(), ag) : '',
              st ? id(sm, O.st, st.toUpperCase(), st) : '',
              dn(at('reg', i)), dn(at('wafid', i)), dn(at('ad', i))].join(','));
  }
  O.rows = out.join(';');
  return O;
}
function repStore_(base, O) {
  var s = JSON.stringify({ce: O.ce, ag: O.ag, st: O.st}) + '\n' + O.rows, cache = CacheService.getScriptCache(), o = {}, n = 0, ttl = repFast_() ? CACHE_TTL : 120;
  for (var p = 0; p < s.length; p += 60000) { o[base + '_' + n] = s.substr(p, 60000); n++; }
  o[base + '_n'] = String(n);
  try { cache.putAll(o, ttl); } catch (e) { /* too big for cache: next search reads the sheet again */ }
}
function repLoad_(name, ver, S, R, ss) {
  var cache = CacheService.getScriptCache(), base = repBase_(name, ver, S, R), n = Number(cache.get(base + '_n') || 0), s = '';
  if (n) {
    var keys = [], i;
    for (i = 0; i < n; i++) keys.push(base + '_' + i);
    var got = cache.getAll(keys), ok = true;
    for (i = 0; i < n; i++) { if (got[keys[i]] === undefined) { ok = false; break; } s += got[keys[i]]; }
    if (ok) {
      var nl = s.indexOf('\n');
      try { var H = JSON.parse(s.slice(0, nl)); return {ce: H.ce, ag: H.ag, st: H.st, rows: s.slice(nl + 1)}; } catch (e) { /* rebuild */ }
    }
  }
  var sh = findSheet_(ss, name);
  if (!sh) return null;
  var O = repExtract_(sh, S, R);
  repStore_(base, O);
  return O;
}
function repAll_(S, R, ss) {
  var names = CL_(), ver = repVer_(names), out = [];
  names.forEach(function (n) { var O = repLoad_(n, ver[n], S, R, ss); if (O) out.push(O); });
  return out;
}
// Called by the 10-minute trigger so the first Search is fast too
function repWarm_(S, ss) {
  try { repAll_(S, repClean_(S.reports), ss); } catch (e) { /* ignore */ }
}

// type: 'agent' | 'reg' | 'wafid'   f: {center, from, to, agent, status}  (dates yyyy-mm-dd, blank = no limit)   token: Reports login token
function getReports(type, f, token) {
  rptRequire_(token);
  f = f || {};
  var S = getSettings(), R = repClean_(S.reports), ss = getSS_();
  var sets = repAll_(S, R, ss);
  var di = type === 'wafid' ? 4 : (type === 'agent' ? 5 : 3);          // position of the date inside a row
  var from = f.from ? String(f.from).replace(/-/g, '') : '', to = f.to ? String(f.to).replace(/-/g, '') : '';
  var cf = (f.center && f.center !== 'ALL') ? normName_(f.center) : '';
  var cols = cf ? [cf] : CL_().slice(), colSet = {};
  cols.forEach(function (c) { colSet[c] = 1; });
  var ag0 = String(f.agent || '').toUpperCase(), sf = String(f.status || '').toUpperCase();
  var rows = {}, agents = {}, statuses = {}, colT = {}, grand = 0;
  sets.forEach(function (O) {
    if (!O.rows) return;
    var list = O.rows.split(';');
    for (var i = 0; i < list.length; i++) {
      var p = list[i].split(','), d = p[di];
      if (!d || (from && d < from) || (to && d > to)) continue;
      var c = O.ce[+p[0]]; if (!colSet[c]) continue;
      var ag = p[1] === '' ? '' : O.ag[+p[1]], st = p[2] === '' ? '' : O.st[+p[2]];
      if (ag) agents[ag.toUpperCase()] = ag;
      if (st) statuses[st.toUpperCase()] = st;
      if (sf && sf !== 'ALL' && st.toUpperCase() !== sf) continue;
      if (ag0 && ag0 !== 'ALL' && ag.toUpperCase() !== ag0) continue;
      var key, label;
      if (type === 'agent') { if (!ag) continue; key = ag.toUpperCase(); label = ag; }
      else { key = d.slice(0, 6); label = key.slice(0, 4) + '-' + REP_MON[+key.slice(4) - 1]; }
      var o = rows[key] = rows[key] || {label: label, n: {}, total: 0};
      o.n[c] = (o.n[c] || 0) + 1; o.total++; colT[c] = (colT[c] || 0) + 1; grand++;
    }
  });
  return {
    type: type, centers: cols, grand: grand, colT: colT,
    rows: Object.keys(rows).sort().map(function (k) { return rows[k]; }),
    agents: Object.keys(agents).sort().map(function (k) { return agents[k]; }),
    statuses: Object.keys(statuses).sort().map(function (k) { return statuses[k]; }),
    ts: Utilities.formatDate(new Date(), TZ, 'dd MMM yyyy, hh:mm a')
  };
}
// Agent and status lists for the filter boxes (all dates). Also warms the cache before the first Search.
function rptOptions(token) {
  rptRequire_(token);
  var S = getSettings(), R = repClean_(S.reports), sets = repAll_(S, R, getSS_()), ag = {}, st = {};
  sets.forEach(function (O) {
    O.ag.forEach(function (a) { ag[a.toUpperCase()] = a; });
    O.st.forEach(function (a) { st[a.toUpperCase()] = a; });
  });
  return {
    agents: Object.keys(ag).sort().map(function (k) { return ag[k]; }),
    statuses: Object.keys(st).sort().map(function (k) { return st[k]; })
  };
}


/* =====================================================================
   REPORTS ACCESS
   - Lock Reports page (tick): Reports needs a login. Untick: everyone can open Reports.
   - Gmail addresses on the access list open Reports with NO password
     (recognised automatically when Google tells the script who is visiting,
      otherwise they confirm with a code sent to their Gmail).
   - The admin (script owner account, or anyone who enters the admin password) opens Reports without the Reports password.
   - Everyone else must enter the one Reports password.
   - The list, the Reports password and the lock are set by the admin in Settings > Reports.
   - A login is remembered only for the browser session (the page keeps the token in sessionStorage),
     so closing the browser asks for the login again.
   ===================================================================== */

var RPT_TTL = 21600; // login stays valid 6 hours (CacheService maximum)

function rptProps_() { return PropertiesService.getScriptProperties(); }
function rptCache_() { return CacheService.getScriptCache(); }
function rptUsers_() { try { return JSON.parse(rptProps_().getProperty('RPT_USERS') || '{}'); } catch (e) { return {}; } }
function rptLocked_() { return rptProps_().getProperty('RPT_LOCK') === '1'; }
function rptEmail_(e) { return String(e || '').trim().toLowerCase(); }
function rptValid_(e) { return /^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$/.test(e); }
function rptHash_(salt, pw) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}
function rptMutate_(fn) {
  var l = LockService.getScriptLock(); l.waitLock(10000);
  try { var u = rptUsers_(); var r = fn(u); rptProps_().setProperty('RPT_USERS', JSON.stringify(u)); return r; }
  finally { l.releaseLock(); }
}
function rptAdmin_(pw) { if (!checkPassword(pw)) throw new Error('Admin password required'); }
function rptPwSet_() { return !!rptProps_().getProperty('RPT_PWH'); }
function rptPwVer_() { return rptProps_().getProperty('RPT_PWV') || '0'; }
function rptAdmVer_() { return rptProps_().getProperty('ADMIN_HASH') || 'default'; }
// Google account that owns the spreadsheet = the admin
function rptOwner_() {
  var c = rptCache_(), v = c.get('rown');
  if (v !== null) return v === '-' ? '' : v;
  var e = '';
  try { var o = getSS_().getOwner(); e = o ? rptEmail_(o.getEmail()) : ''; } catch (x) { e = ''; }
  try { c.put('rown', e || '-', RPT_TTL); } catch (x2) {}
  return e;
}

/* ---- login sessions ---- */
function rptIssue_(who, ver) {
  var t = Utilities.getUuid() + Utilities.getUuid();
  rptCache_().put('rt_' + t, who + '|' + ver, RPT_TTL);
  return { ok: true, token: t };
}
function rptWho_(token) {
  if (!token) return null;
  var raw = rptCache_().get('rt_' + token); if (!raw) return null;
  var i = raw.lastIndexOf('|'), who = raw.slice(0, i), ver = raw.slice(i + 1);
  if (who === '*adm') return ver === rptAdmVer_() ? who : null;        // admin login dies when the admin password changes
  if (who === '*pw') return ver === rptPwVer_() ? who : null;          // password login dies when the password changes
  var u = rptUsers_()[who];
  return (u && String(u.v || 0) === ver) ? who : null;                  // removed from the list = no access
}
/* First line of getReports */
function rptRequire_(token) {
  if (!rptLocked_()) return '';
  var who = rptWho_(token);
  if (!who) throw new Error('Reports login required');
  return who;
}
function rptFail_(key) { var c = rptCache_(), k = 'rf_' + key; c.put(k, String(+(c.get(k) || 0) + 1), 600); }
function rptBlocked_(key, max) { return +(rptCache_().get('rf_' + key) || 0) >= (max || 5); }

/* ---- public functions called from the page ---- */
// One call when Reports Summary is clicked.
//   lock off                         -> { locked: false }  (Reports opens directly)
//   token still valid (same session) -> { locked: true, valid: true }
//   listed Gmail that Google can identify -> a token straight away
//   otherwise                        -> { locked: true, hasPw }  (the page shows the Reports Login popup)
function rptStatus(token, adminPw) {
  if (!rptLocked_()) return { locked: false };
  if (token && rptWho_(token)) return { locked: true, valid: true };
  // admin: already logged in to Settings in this page (admin password), or the owner Google account
  if (adminPw && ok(adminPw)) { var ta = rptIssue_('*adm', rptAdmVer_()); return { locked: true, token: ta.token, auto: true }; }
  var em = '';
  try { em = rptEmail_(Session.getActiveUser().getEmail()); } catch (e) { em = ''; }
  if (em && em === rptOwner_()) { var to = rptIssue_('*adm', rptAdmVer_()); return { locked: true, token: to.token, auto: true }; }
  var u = rptUsers_();
  if (em && u[em]) { var t = rptIssue_(em, u[em].v || 0); return { locked: true, token: t.token, auto: true }; }
  return { locked: true, hasPw: rptPwSet_() };
}

// Reports password (for everyone who is not the admin and not on the Gmail list). Minimum length: 1 character.
function rptLogin(password) {
  password = String(password || '');
  if (!password) return { ok: false, error: 'Enter the password' };
  if (rptBlocked_('pw', 10)) return { ok: false, error: 'Too many attempts. Try again in 10 minutes.' };
  if (ok(password)) { rptCache_().remove('rf_pw'); return rptIssue_('*adm', rptAdmVer_()); }   // the admin password also opens Reports
  var p = rptProps_(), h = p.getProperty('RPT_PWH'), s = p.getProperty('RPT_PWS');
  if (!h) return { ok: false, error: 'No Reports password has been set yet. Ask the admin.' };
  if (rptHash_(s, password) !== h) { rptFail_('pw'); return { ok: false, error: 'Wrong password' }; }
  rptCache_().remove('rf_pw');
  return rptIssue_('*pw', rptPwVer_());
}

// Listed Gmail without automatic recognition: confirm with a code sent to that Gmail (no password needed)
function rptSendCode(email) {
  email = rptEmail_(email);
  var c = rptCache_();
  if (!rptUsers_()[email]) return { ok: false, error: 'This Gmail is not on the access list.' };
  if (c.get('rs_' + email)) return { ok: false, error: 'A code was just sent. Wait one minute.' };
  var code = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 900000 + 100000);
  c.put('ro_' + email, code, 600); c.put('ra_' + email, '0', 600); c.put('rs_' + email, '1', 60);
  MailApp.sendEmail(email, 'Reports verification code',
    'Your Reports verification code is ' + code + '\n\nIt is valid for 10 minutes. If you did not ask for it, ignore this email.');
  return { ok: true };
}
function rptVerifyCode(email, code) {
  email = rptEmail_(email);
  var c = rptCache_(), real = c.get('ro_' + email), u = rptUsers_()[email];
  if (!u) return { ok: false, error: 'This Gmail is not on the access list.' };
  if (!real) return { ok: false, error: 'Code expired. Get a new code.' };
  var a = +(c.get('ra_' + email) || 0) + 1;
  if (a > 5) { c.remove('ro_' + email); return { ok: false, error: 'Too many wrong codes. Get a new code.' }; }
  c.put('ra_' + email, String(a), 600);
  if (String(code || '').trim() !== real) return { ok: false, error: 'Wrong code' };
  c.remove('ro_' + email);
  return rptIssue_(email, u.v || 0);
}

function rptLogout(token) { if (token) rptCache_().remove('rt_' + token); return true; }

/* ---- admin (needs the Settings admin password) ---- */
function rptAdminGet(adminPw) {
  rptAdmin_(adminPw);
  var u = rptUsers_();
  return { locked: rptLocked_(), hasPw: rptPwSet_(), users: Object.keys(u).sort().map(function (e) { return { email: e }; }) };
}

// newPw: blank = keep the current Reports password
function rptAdminSave(adminPw, locked, emails, newPw) {
  rptAdmin_(adminPw);
  var list = [], seen = {};
  (emails || []).forEach(function (e) {
    e = rptEmail_(e);
    if (!rptValid_(e)) throw new Error('Invalid email: ' + e);
    if (!seen[e]) { seen[e] = 1; list.push(e); }
  });
  newPw = String(newPw || '');
  if (newPw.length > 100) throw new Error('Reports password is too long (maximum 100 characters)');
  rptMutate_(function (u) {
    var keep = {};
    list.forEach(function (e) { keep[e] = u[e] || { v: Date.now() }; });
    Object.keys(u).forEach(function (k) { delete u[k]; });
    Object.keys(keep).forEach(function (k) { u[k] = keep[k]; });
  });
  if (newPw) {
    var salt = Utilities.getUuid(), p = rptProps_();
    p.setProperties({ RPT_PWS: salt, RPT_PWH: rptHash_(salt, newPw), RPT_PWV: String(Date.now()) });
  }
  rptProps_().setProperty('RPT_LOCK', locked ? '1' : '0');
  return rptAdminGet(adminPw);
}
