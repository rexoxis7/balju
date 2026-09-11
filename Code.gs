/*******************************************************************
 *  스테이지랩 발주장부 — 서버 (Google Apps Script)
 *  ---------------------------------------------------------------
 *  설치 순서
 *   1) 이 파일 전체를 Apps Script 편집기의 Code.gs 에 붙여넣기
 *   2) 아래 [설정] 두 줄을 원하는 값으로 수정
 *   3) 편집기 위쪽 함수 선택에서 setup 을 고르고 ▶ 실행 (권한 승인)
 *   4) 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 배포
 *   5) 나온 웹 앱 URL을 앱의 "서버 주소"에 입력
 *******************************************************************/

// ───────────── [설정] ─────────────
const ADMIN_NAME = '사장님';                 // 관리자 아이디(이름)
const ADMIN_INITIAL_PASSWORD = 'stagelab2026'; // 관리자 최초 비밀번호 (첫 로그인 후 반드시 변경)
const SESSION_DAYS = 30;                    // 로그인 유지 기간(일)
const TZ = 'Asia/Seoul';                    // 날짜 기준 시간대
const PEPPER = 'SL-' + 'change-this-to-any-random-text-7f3a9'; // 비밀번호 해시 비밀값 (한 번 정한 뒤 바꾸지 마세요)
// ──────────────────────────────────

const SHEET_ORDERS = '발주';
const SHEET_STAFF = '직원';
const SHEET_SESSIONS = '세션';
const SHEET_SETTINGS = '설정';
const PHOTO_FOLDER = '발주장부_사진';

const ORDER_COLS = ['id','item','colorName','qty','unit','supplier','orderer','orderDate','receiveDate','receivedBy','receivedAt',
  'purpose','customer','vehicles','code','thumb','photoId','delivery','note','done','doneAt','doneBy','deleted','createdBy','createdAt','updatedAt','log','notePhotoIds','noteThumbs'];
const STAFF_COLS = ['name','role','salt','hash','active','mustChange','createdAt'];
const SESSION_COLS = ['token','name','role','createdAt','expiresAt','lastSeen','device'];
const CLIENT_EDITABLE = ['item','colorName','qty','unit','supplier','orderer','orderDate','receiveDate','purpose','customer','vehicles','code','thumb','delivery','note'];

const DEFAULT_OPTIONS = {
  suppliers:['커버플릭','비닐프로그','비닐프로그-CYS/터미낙스','오퍼(오토벨벳)','오토모스트','TECKWRAP'],
  customers:['RP','천안LX서북점','고양몬스터디테일','픽스디테일링','순정디테일링'],
  items:['스타빌','윈드쉴드','파이온','고광택PET','컬러PPF','무광PPF','하이무광PPF','미들무광PPF','블랙PPF','열차단PPF','3x','2x','글로우'],
  units:['롤','m','ea']
};

/* ═══════════════ 최초 설정 (편집기에서 1회 실행) ═══════════════ */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_ORDERS, ORDER_COLS);
  ensureSheet_(ss, SHEET_STAFF, STAFF_COLS);
  ensureSheet_(ss, SHEET_SESSIONS, SESSION_COLS);
  const st = ensureSheet_(ss, SHEET_SETTINGS, ['key','value']);
  if (!getSetting_('options')) setSetting_('options', JSON.stringify(DEFAULT_OPTIONS));
  if (!getSetting_('photoFolderId')) {
    const folder = DriveApp.createFolder(PHOTO_FOLDER);
    setSetting_('photoFolderId', folder.getId());
  }
  const staff = readAll_(SHEET_STAFF, STAFF_COLS);
  if (!staff.some(s => s.name === ADMIN_NAME)) {
    const salt = randomHex_(16);
    appendRow_(SHEET_STAFF, STAFF_COLS, { name: ADMIN_NAME, role: 'admin', salt: salt, hash: hashPassword_(ADMIN_INITIAL_PASSWORD, salt), active: 'Y', mustChange: 'Y', createdAt: nowIso_() });
  }
  Logger.log('설정 완료. 관리자: ' + ADMIN_NAME + ' / 최초 비밀번호: ' + ADMIN_INITIAL_PASSWORD);
}

/* ═══════════════ HTTP 진입점 ═══════════════ */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'ping') return json_({ ok: true, app: 'stagelab-balju', v: 3 });
  let html;
  try { html = HtmlService.createHtmlOutputFromFile('index').getContent(); }
  catch (err) { return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:20px">index.html 파일이 없습니다. Apps Script 편집기에서 파일 추가 → HTML → 이름 "index" 로 만든 뒤 앱 화면 코드를 붙여넣으세요.</p>'); }
  html = html.replace('%%API_URL%%', ScriptApp.getService().getUrl());
  return HtmlService.createHtmlOutput(html).setTitle('발주장부')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, code: 'bad_request', message: '요청 형식 오류' }); }
  try {
    const action = String(req.action || '');
    if (action === 'ping') return json_({ ok: true });
    if (action === 'login') return json_(login_(req));
    const session = getSession_(req.token);
    if (!session) return json_({ ok: false, code: 'unauthorized', message: '로그인이 필요합니다' });
    const handler = HANDLERS[action];
    if (!handler) return json_({ ok: false, code: 'unknown_action', message: '알 수 없는 요청: ' + action });
    const data = handler(req, session);
    return json_({ ok: true, data: data });
  } catch (err) {
    const msg = (err && err.message) || String(err);
    const code = msg.indexOf('FORBIDDEN') === 0 ? 'forbidden' : 'error';
    return json_({ ok: false, code: code, message: msg.replace(/^FORBIDDEN:?\s*/, '') });
  }
}

const HANDLERS = {
  logout: (r, s) => { deleteSession_(r.token); return { ok: true }; },
  logoutAll: (r, s) => { requireAdmin_(s); clearSheet_(SHEET_SESSIONS, SESSION_COLS); const ep = randomHex_(4); setSetting_('sessionEpoch', ep); CacheService.getScriptCache().put('epoch', ep, 21600); return { ok: true }; },
  me: (r, s) => ({ name: s.name, role: s.role, mustChange: s.mustChange }),
  changePassword: (r, s) => changePassword_(s, r.oldPassword, r.newPassword),
  bootstrap: (r, s) => ({ orders: listOrders_(), options: getOptions_(), staff: staffNames_(), me: { name: s.name, role: s.role, mustChange: s.mustChange }, serverTime: nowIso_() }),
  list: (r, s) => ({ orders: listOrders_() }),
  save: (r, s) => saveOrder_(r.order || {}, s),
  receive: (r, s) => patchOrder_(r.id, { receiveDate: today_(), receivedBy: s.name, receivedAt: nowIso_(), done: 'N', doneAt: '', doneBy: '' }, s.name + ' 입고 확인', s),
  unreceive: (r, s) => patchOrder_(r.id, { receiveDate: '', receivedBy: '', receivedAt: '', done: 'N', doneAt: '', doneBy: '' }, '입고 취소', s),
  done: (r, s) => patchOrder_(r.id, { done: 'Y', doneAt: nowIso_(), doneBy: s.name }, r.label || '사용완료', s),
  undone: (r, s) => patchOrder_(r.id, { done: 'N', doneAt: '', doneBy: '' }, '완료 취소', s),
  setVehicles: (r, s) => patchOrder_(r.id, { vehicles: JSON.stringify(Array.isArray(r.vehicles) ? r.vehicles : []) }, '차량 기록: ' + (r.vehicles || []).join(', '), s),
  softDelete: (r, s) => bulkPatch_(r.ids, { deleted: 'Y' }, '삭제(휴지통)', s),
  restore: (r, s) => bulkPatch_(r.ids, { deleted: 'N' }, '복구', s),
  purge: (r, s) => { requireAdmin_(s); return purge_(r.ids); },
  deleteAll: (r, s) => { requireAdmin_(s); const ids = listOrders_().filter(o => o.deleted !== 'Y').map(o => o.id); return bulkPatch_(ids, { deleted: 'Y' }, '전체 삭제(휴지통)', s); },
  purgeTrash: (r, s) => { requireAdmin_(s); const ids = listOrders_().filter(o => o.deleted === 'Y').map(o => o.id); return purge_(ids); },
  setOptions: (r, s) => { const o = sanitizeOptions_(r.options); setSetting_('options', JSON.stringify(o)); return o; },
  staffList: (r, s) => { requireAdmin_(s); return readAll_(SHEET_STAFF, STAFF_COLS).map(x => ({ name: x.name, role: x.role, active: x.active === 'Y', mustChange: x.mustChange === 'Y', createdAt: x.createdAt })); },
  staffAdd: (r, s) => { requireAdmin_(s); return staffAdd_(r.name, r.password, r.role === 'admin' ? 'admin' : 'staff'); },
  staffReset: (r, s) => { requireAdmin_(s); return staffSetPassword_(r.name, r.password, true); },
  staffToggle: (r, s) => { requireAdmin_(s); return staffToggle_(r.name, s); },
  photoPut: (r, s) => photoPut_(r.id, r.dataUrl),
  photoGet: (r, s) => photoGet_(r.id),
  notePhotoAdd: (r, s) => notePhotoAdd_(r.id, r.dataUrl, r.thumb),
  notePhotoDel: (r, s) => notePhotoDel_(r.id, r.photoId),
  notePhotoGet: (r, s) => notePhotoGet_(r.id, r.photoId)
};

/* ═══════════════ 인증 ═══════════════ */
function login_(req) {
  const name = String(req.name || '').trim();
  const password = String(req.password || '');
  if (!name || !password) return { ok: false, code: 'bad_request', message: '이름과 비밀번호를 입력하세요' };
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + name;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) return { ok: false, code: 'locked', message: '5회 실패로 15분간 잠겼습니다' };
  const staff = readAll_(SHEET_STAFF, STAFF_COLS).find(x => x.name === name);
  if (!staff || staff.active !== 'Y' || hashPassword_(password, staff.salt) !== staff.hash) {
    cache.put(failKey, String(fails + 1), 900);
    return { ok: false, code: 'bad_credentials', message: '이름 또는 비밀번호가 맞지 않습니다' + (fails >= 2 ? ' (' + (5 - fails - 1) + '회 남음)' : '') };
  }
  cache.remove(failKey);
  const token = randomHex_(32);
  const exp = new Date(Date.now() + SESSION_DAYS * 864e5).toISOString();
  appendRow_(SHEET_SESSIONS, SESSION_COLS, { token: token, name: name, role: staff.role, createdAt: nowIso_(), expiresAt: exp, lastSeen: nowIso_(), device: String(req.device || '').slice(0, 80) });
  cache.put('tok_' + token, JSON.stringify({ name: name, role: staff.role, exp: exp, mustChange: staff.mustChange === 'Y', ep: epoch_() }), 3600);
  return { ok: true, data: { token: token, name: name, role: staff.role, mustChange: staff.mustChange === 'Y', expiresAt: exp } };
}
function getSession_(token) {
  token = String(token || '');
  if (!/^[0-9a-f]{64}$/.test(token)) return null;
  const cache = CacheService.getScriptCache();
  const c = cache.get('tok_' + token);
  if (c) { const s = JSON.parse(c); if (new Date(s.exp) > new Date() && s.ep === epoch_()) return s; }
  const rows = readAll_(SHEET_SESSIONS, SESSION_COLS);
  const row = rows.find(x => x.token === token);
  if (!row) return null;
  if (new Date(row.expiresAt) < new Date()) { deleteSession_(token); return null; }
  const staff = readAll_(SHEET_STAFF, STAFF_COLS).find(x => x.name === row.name);
  if (!staff || staff.active !== 'Y') { deleteSession_(token); return null; }
  const s = { name: row.name, role: staff.role, exp: row.expiresAt, mustChange: staff.mustChange === 'Y', ep: epoch_() };
  cache.put('tok_' + token, JSON.stringify(s), 3600);
  return s;
}
function epoch_() { const cache = CacheService.getScriptCache(); let ep = cache.get('epoch'); if (!ep) { ep = getSetting_('sessionEpoch') || '0'; cache.put('epoch', ep, 21600); } return ep; }
function deleteSession_(token) {
  CacheService.getScriptCache().remove('tok_' + token);
  const sh = sheet_(SHEET_SESSIONS); const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) if (vals[i][0] === token) sh.deleteRow(i + 1);
}
function requireAdmin_(s) { if (s.role !== 'admin') throw new Error('FORBIDDEN: 관리자만 할 수 있습니다'); }
function changePassword_(s, oldPw, newPw) {
  newPw = String(newPw || '');
  if (newPw.length < 6) throw new Error('새 비밀번호는 6자 이상이어야 합니다');
  const staff = readAll_(SHEET_STAFF, STAFF_COLS).find(x => x.name === s.name);
  if (!staff || hashPassword_(String(oldPw || ''), staff.salt) !== staff.hash) throw new Error('현재 비밀번호가 맞지 않습니다');
  staffSetPassword_(s.name, newPw, false);
  invalidateUserSessions_(s.name);
  return { ok: true };
}
function staffAdd_(name, password, role) {
  name = String(name || '').trim(); password = String(password || '');
  if (!name) throw new Error('이름을 입력하세요');
  if (password.length < 4) throw new Error('임시 비밀번호는 4자 이상');
  if (readAll_(SHEET_STAFF, STAFF_COLS).some(x => x.name === name)) throw new Error('이미 있는 이름입니다');
  const salt = randomHex_(16);
  appendRow_(SHEET_STAFF, STAFF_COLS, { name: name, role: role, salt: salt, hash: hashPassword_(password, salt), active: 'Y', mustChange: 'Y', createdAt: nowIso_() });
  return { ok: true };
}
function staffSetPassword_(name, password, mustChange) {
  password = String(password || ''); if (password.length < 4) throw new Error('비밀번호는 4자 이상');
  const sh = sheet_(SHEET_STAFF); const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) if (vals[i][0] === name) {
    const salt = randomHex_(16);
    sh.getRange(i + 1, 3, 1, 2).setValues([[salt, hashPassword_(password, salt)]]);
    sh.getRange(i + 1, 6).setValue(mustChange ? 'Y' : 'N');
    invalidateUserSessions_(name);
    return { ok: true };
  }
  throw new Error('직원을 찾을 수 없습니다');
}
function staffToggle_(name, s) {
  if (name === s.name) throw new Error('본인 계정은 비활성화할 수 없습니다');
  const sh = sheet_(SHEET_STAFF); const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) if (vals[i][0] === name) {
    const now = vals[i][4] === 'Y' ? 'N' : 'Y'; sh.getRange(i + 1, 5).setValue(now);
    if (now === 'N') invalidateUserSessions_(name);
    return { active: now === 'Y' };
  }
  throw new Error('직원을 찾을 수 없습니다');
}
function invalidateUserSessions_(name) {
  const sh = sheet_(SHEET_SESSIONS); const vals = sh.getDataRange().getValues(); const cache = CacheService.getScriptCache();
  for (let i = vals.length - 1; i >= 1; i--) if (vals[i][1] === name) { cache.remove('tok_' + vals[i][0]); sh.deleteRow(i + 1); }
}
function staffNames_() { return readAll_(SHEET_STAFF, STAFF_COLS).filter(x => x.active === 'Y').map(x => x.name); }

/* ═══════════════ 발주 ═══════════════ */
function listOrders_() { return readAll_(SHEET_ORDERS, ORDER_COLS); }
function saveOrder_(o, s) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = sheet_(SHEET_ORDERS); const idx = indexById_(sh);
    const clean = {};
    CLIENT_EDITABLE.forEach(k => { let v = o[k]; if (k === 'vehicles') v = JSON.stringify(Array.isArray(v) ? v : []); clean[k] = v == null ? '' : String(v); });
    if (!clean.item) throw new Error('품목이 필요합니다');
    const now = nowIso_();
    if (o.id && idx[o.id]) {
      const row = idx[o.id]; const cur = rowToObj_(sh.getRange(row, 1, 1, ORDER_COLS.length).getValues()[0], ORDER_COLS);
      // 입고 정보는 서버가 관리: 클라이언트가 입고일을 새로 넣으면 확인자=본인
      if (clean.receiveDate && !cur.receiveDate) { clean.receivedBy = s.name; clean.receivedAt = now; }
      if (!clean.receiveDate) { clean.receivedBy = ''; clean.receivedAt = ''; clean.done = 'N'; clean.doneAt = ''; clean.doneBy = ''; }
      clean.updatedAt = now; clean.log = pushLog_(cur.log, s.name, '수정');
      writeFields_(sh, row, clean);
      return rowToObj_(sh.getRange(row, 1, 1, ORDER_COLS.length).getValues()[0], ORDER_COLS);
    }
    const id = randomHex_(8);
    const rec = Object.assign({}, clean, { id: id, receivedBy: clean.receiveDate ? s.name : '', receivedAt: clean.receiveDate ? now : '', photoId: '', done: 'N', doneAt: '', doneBy: '', deleted: 'N', createdBy: s.name, createdAt: now, updatedAt: now, log: pushLog_('', s.name, '등록') });
    appendRow_(SHEET_ORDERS, ORDER_COLS, rec);
    return rec;
  } finally { lock.releaseLock(); }
}
function patchOrder_(id, patch, what, s) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = sheet_(SHEET_ORDERS); const idx = indexById_(sh); const row = idx[id];
    if (!row) throw new Error('기록을 찾을 수 없습니다');
    const cur = rowToObj_(sh.getRange(row, 1, 1, ORDER_COLS.length).getValues()[0], ORDER_COLS);
    patch.updatedAt = nowIso_(); patch.log = pushLog_(cur.log, s.name, what);
    writeFields_(sh, row, patch);
    return rowToObj_(sh.getRange(row, 1, 1, ORDER_COLS.length).getValues()[0], ORDER_COLS);
  } finally { lock.releaseLock(); }
}
function bulkPatch_(ids, patch, what, s) {
  ids = Array.isArray(ids) ? ids : []; if (!ids.length) return { count: 0 };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SHEET_ORDERS); const idx = indexById_(sh); let n = 0;
    ids.forEach(id => { const row = idx[id]; if (!row) return;
      const cur = rowToObj_(sh.getRange(row, 1, 1, ORDER_COLS.length).getValues()[0], ORDER_COLS);
      const p = Object.assign({}, patch, { updatedAt: nowIso_(), log: pushLog_(cur.log, s.name, what) });
      writeFields_(sh, row, p); n++; });
    return { count: n };
  } finally { lock.releaseLock(); }
}
function purge_(ids) {
  ids = Array.isArray(ids) ? ids : []; if (!ids.length) return { count: 0 };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(SHEET_ORDERS); const vals = sh.getDataRange().getValues(); const set = {}; ids.forEach(i => set[i] = true); let n = 0;
    const pi = ORDER_COLS.indexOf('photoId');
    for (let i = vals.length - 1; i >= 1; i--) if (set[vals[i][0]]) {
      const fid = vals[i][pi]; if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} }
      sh.deleteRow(i + 1); n++;
    }
    return { count: n };
  } finally { lock.releaseLock(); }
}
function pushLog_(logStr, who, what) {
  let l = []; try { l = JSON.parse(logStr || '[]'); } catch (e) { l = []; }
  l.push({ t: nowIso_(), w: who, m: String(what || '').slice(0, 120) });
  return JSON.stringify(l.slice(-40));
}

/* ═══════════════ 사진 (Drive) ═══════════════ */
function photoPut_(id, dataUrl) {
  const m = /^data:(image\/[a-z]+);base64,(.+)$/i.exec(String(dataUrl || ''));
  if (!m) throw new Error('사진 형식 오류');
  if (m[2].length > 2.5e6) throw new Error('사진이 너무 큽니다');
  const folder = DriveApp.getFolderById(getSetting_('photoFolderId'));
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + '.jpg');
  const sh = sheet_(SHEET_ORDERS); const idx = indexById_(sh); const row = idx[id]; if (!row) throw new Error('기록을 찾을 수 없습니다');
  const pi = ORDER_COLS.indexOf('photoId') + 1; const old = sh.getRange(row, pi).getValue();
  if (old) { try { DriveApp.getFileById(old).setTrashed(true); } catch (e) {} }
  const file = folder.createFile(blob);
  sh.getRange(row, pi).setValue(file.getId());
  return { photoId: file.getId() };
}
function photoGet_(id) {
  const sh = sheet_(SHEET_ORDERS); const idx = indexById_(sh); const row = idx[id]; if (!row) return { dataUrl: '' };
  const fid = sh.getRange(row, ORDER_COLS.indexOf('photoId') + 1).getValue(); if (!fid) return { dataUrl: '' };
  try { const b = DriveApp.getFileById(fid).getBlob(); return { dataUrl: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) }; }
  catch (e) { return { dataUrl: '' }; }
}

/* ═══════════════ 특이사항 사진 (여러 장, Drive 저장) ═══════════════ */
function noteLists_(sh, row) {
  const ci = ORDER_COLS.indexOf('notePhotoIds') + 1, ct = ORDER_COLS.indexOf('noteThumbs') + 1;
  let ids = [], ths = [];
  try { ids = JSON.parse(sh.getRange(row, ci).getValue() || '[]'); } catch (e) { ids = []; }
  try { ths = JSON.parse(sh.getRange(row, ct).getValue() || '[]'); } catch (e) { ths = []; }
  if (!Array.isArray(ids)) ids = []; if (!Array.isArray(ths)) ths = [];
  return { ci: ci, ct: ct, ids: ids, ths: ths };
}
function notePhotoAdd_(id, dataUrl, thumb) {
  const m = /^data:(image\/[a-z]+);base64,(.+)$/i.exec(String(dataUrl || ''));
  if (!m) throw new Error('사진 형식 오류');
  if (m[2].length > 2.5e6) throw new Error('사진이 너무 큽니다');
  thumb = String(thumb || ''); if (thumb.length > 2600) thumb = '';
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = sheet_(SHEET_ORDERS); ensureCols_(sh, ORDER_COLS.length);
    const idx = indexById_(sh); const row = idx[id]; if (!row) throw new Error('기록을 찾을 수 없습니다');
    const L = noteLists_(sh, row);
    if (L.ids.length >= 20) throw new Error('특이사항 사진은 20장까지입니다');
    const folder = DriveApp.getFolderById(getSetting_('photoFolderId'));
    const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + '-memo-' + (L.ids.length + 1) + '.jpg'));
    L.ids.push(file.getId()); L.ths.push(thumb);
    // 시트 한 칸은 5만 자까지만 담을 수 있어, 넘칠 것 같으면 새 미리보기만 비워 둔다(원본 사진은 드라이브에 안전히 저장됨)
    if (JSON.stringify(L.ths).length > 45000) { L.ths[L.ths.length - 1] = ''; }
    sh.getRange(row, L.ci).setNumberFormat('@').setValue(JSON.stringify(L.ids));
    sh.getRange(row, L.ct).setNumberFormat('@').setValue(JSON.stringify(L.ths));
    return { notePhotoIds: L.ids, noteThumbs: L.ths };
  } finally { lock.releaseLock(); }
}
function notePhotoDel_(id, pid) {
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = sheet_(SHEET_ORDERS); ensureCols_(sh, ORDER_COLS.length);
    const idx = indexById_(sh); const row = idx[id]; if (!row) throw new Error('기록을 찾을 수 없습니다');
    const L = noteLists_(sh, row);
    const i = L.ids.indexOf(String(pid || ''));
    if (i >= 0) {
      try { DriveApp.getFileById(L.ids[i]).setTrashed(true); } catch (e) {}
      L.ids.splice(i, 1); L.ths.splice(i, 1);
      sh.getRange(row, L.ci).setNumberFormat('@').setValue(JSON.stringify(L.ids));
      sh.getRange(row, L.ct).setNumberFormat('@').setValue(JSON.stringify(L.ths));
    }
    return { notePhotoIds: L.ids, noteThumbs: L.ths };
  } finally { lock.releaseLock(); }
}
function notePhotoGet_(id, pid) {
  const sh = sheet_(SHEET_ORDERS); ensureCols_(sh, ORDER_COLS.length);
  const idx = indexById_(sh); const row = idx[id]; if (!row) return { dataUrl: '' };
  const L = noteLists_(sh, row);
  if (L.ids.indexOf(String(pid || '')) < 0) return { dataUrl: '' };
  try { const b = DriveApp.getFileById(String(pid)).getBlob(); return { dataUrl: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) }; }
  catch (e) { return { dataUrl: '' }; }
}

/* ═══════════════ 옵션 ═══════════════ */
function getOptions_() { try { const o = JSON.parse(getSetting_('options') || '{}'); return sanitizeOptions_(o); } catch (e) { return DEFAULT_OPTIONS; } }
function sanitizeOptions_(o) { const out = {}; Object.keys(DEFAULT_OPTIONS).forEach(k => { const a = (o && Array.isArray(o[k])) ? o[k] : DEFAULT_OPTIONS[k]; out[k] = a.map(x => String(x).trim()).filter(Boolean).slice(0, 100); if (!out[k].length) out[k] = DEFAULT_OPTIONS[k]; }); return out; }
function getSetting_(key) { const rows = readAll_(SHEET_SETTINGS, ['key','value']); const r = rows.find(x => x.key === key); return r ? r.value : ''; }
function setSetting_(key, value) { const sh = sheet_(SHEET_SETTINGS); const vals = sh.getDataRange().getValues(); for (let i = 1; i < vals.length; i++) if (vals[i][0] === key) { sh.getRange(i + 1, 2).setValue(value); return; } sh.appendRow([key, value]); }

/* ═══════════════ 시트 유틸 ═══════════════ */
function sheet_(name) { const ss = SpreadsheetApp.getActiveSpreadsheet(); const sh = ss.getSheetByName(name); if (!sh) throw new Error('시트가 없습니다: ' + name + ' — setup 을 먼저 실행하세요'); return sh; }
function ensureSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  if (sh.getLastRow() === 0) { sh.appendRow(cols); sh.setFrozenRows(1); sh.getRange(1, 1, 1, cols.length).setFontWeight('bold'); }
  sh.getRange(1, 1, Math.max(sh.getMaxRows(), 2000), Math.max(cols.length, sh.getMaxColumns())).setNumberFormat('@');
  return sh;
}
function ensureCols_(sh, n) { if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns()); }
function readAll_(name, cols) {
  const sh = sheet_(name); ensureCols_(sh, cols.length); const last = sh.getLastRow(); if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  return vals.filter(r => r[0] !== '' && r[0] != null).map(r => rowToObj_(r, cols));
}
function rowToObj_(r, cols) { const o = {}; cols.forEach((c, i) => { let v = r[i]; if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd'); o[c] = v == null ? '' : v; }); return o; }
function appendRow_(name, cols, obj) { const sh = sheet_(name); const row = cols.map(c => obj[c] == null ? '' : obj[c]); sh.appendRow(row); const r = sh.getLastRow(); sh.getRange(r, 1, 1, cols.length).setNumberFormat('@').setValues([row]); }
function writeFields_(sh, row, patch) { Object.keys(patch).forEach(k => { const ci = ORDER_COLS.indexOf(k); if (ci >= 0) sh.getRange(row, ci + 1).setNumberFormat('@').setValue(patch[k] == null ? '' : String(patch[k])); }); }
function indexById_(sh) { const last = sh.getLastRow(); const idx = {}; if (last < 2) return idx; const ids = sh.getRange(2, 1, last - 1, 1).getValues(); ids.forEach((r, i) => { if (r[0]) idx[r[0]] = i + 2; }); return idx; }
function clearSheet_(name, cols) { const sh = sheet_(name); sh.clearContents(); sh.appendRow(cols); }

/* ═══════════════ 공통 ═══════════════ */
function hashPassword_(pw, salt) {
  let h = sha256_(salt + PEPPER + pw);
  for (let i = 0; i < 1500; i++) h = sha256_(h + salt);
  return h;
}
function sha256_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function randomHex_(bytes) { let s = ''; while (s.length < bytes * 2) s += Utilities.getUuid().replace(/-/g, ''); return s.slice(0, bytes * 2); }
function nowIso_() { return new Date().toISOString(); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
