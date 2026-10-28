/**
 * @OnlyCurrentDoc
 */
/**
 * 매물장 응대 데스크 — 구글 시트 연결용 Apps Script
 *
 * 이 코드는 고치지 말고 그대로 붙여넣으세요.
 * 비밀번호는 구글 시트의 「설정」 탭에 적습니다.
 *   B1 칸: 사무실 비밀번호 (모든 직원: 매물 보기·검색·응대·새 매물 접수)
 *   B2 칸: 열람 비밀번호   (일부 직원: 연락처·현관 비밀번호·원문 보기, 삭제)
 * 비밀번호를 바꾸면 바로 적용됩니다. (다시 배포할 필요 없음)
 */

/** 처음 한 번: 위쪽 함수 선택 칸에서 '처음실행'을 고르고 ▶ 실행 → 권한 허용 */
function 처음실행() {
  const ss = SpreadsheetApp.getActive();
  const pw = passwords_();
  Logger.log('연결 확인: ' + ss.getName());
  Logger.log('매물 ' + (ss.getSheetByName('매물').getLastRow() - 1) + '건, 손님 ' + (ss.getSheetByName('손님').getLastRow() - 1) + '건');
  Logger.log('일정 ' + (ss.getSheetByName('일정') ? (ss.getSheetByName('일정').getLastRow() - 1) + '건' : '탭 없음 (첫 일정 저장 때 자동 생성)'));
  Logger.log('사무실 비밀번호: ' + (pw.office ? '입력됨' : '비어 있음 → 설정 탭 B1에 입력하세요'));
  Logger.log('열람 비밀번호: ' + (pw.full ? '입력됨' : '비어 있음 → 설정 탭 B2에 입력하세요'));
}

const SHEETS = { listings: '매물', clients: '손님', schedule: '일정' };
const SKIP = ['설정', '사용법'];
// 위 목록에 없는 탭 이름도 그대로 받아서 저장 (앞으로 탭이 늘어나도 다시 배포할 필요 없음)
function sheetName_(key) {
  if (SHEETS[key]) return SHEETS[key];
  key = String(key || '');
  if (!key || SKIP.indexOf(key) >= 0) return null;
  return key;
}
// 열람 비밀번호가 있어야 보이는 칸
const SECRET = {
  '매물': ['연락처', '내놓은분', '집보는법', '원문'],
  '손님': ['연락처', '원문']
};

function passwords_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('설정');
  if (!sh) return { office: '', full: '' };
  return {
    office: String(sh.getRange('B1').getDisplayValue()).trim(),
    full: String(sh.getRange('B2').getDisplayValue()).trim()
  };
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function level_(key) {
  key = String(key || '').trim();
  if (!key) return 0;
  const pw = passwords_();
  if (pw.full && key === pw.full) return 2;
  if (pw.office && key === pw.office) return 1;
  return 0;
}

function readSheet_(name, lv) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return [];
  const v = sh.getDataRange().getDisplayValues();
  const head = v[0];
  const hide = lv >= 2 ? [] : (SECRET[name] || []);
  return v.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { if (hide.indexOf(h) < 0) o[h] = r[i]; });
    if (lv < 2) o._locked = hide.filter(function (h) { const i = head.indexOf(h); return i >= 0 && r[i]; }).join(',');
    return o;
  });
}

function doGet(e) {
  const lv = level_(e && e.parameter && e.parameter.key);
  if (!lv) return out_({ ok: false, error: 'password' });
  const res = { ok: true, level: lv, version: 2 };
  Object.keys(SHEETS).forEach(function (k) { res[k] = readSheet_(SHEETS[k], lv); });
  // 그 밖의 탭은 tabs 아래로
  const known = Object.keys(SHEETS).map(function (k) { return SHEETS[k]; }).concat(SKIP);
  res.tabs = {};
  SpreadsheetApp.getActive().getSheets().forEach(function (sh) {
    const n = sh.getName(); if (known.indexOf(n) < 0) res.tabs[n] = readSheet_(n, lv);
  });
  return out_(res);
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  const lv = level_(b.key);
  if (!lv) return out_({ ok: false, error: 'password' });
  const name = sheetName_(b.sheet);
  if (!name) return out_({ ok: false, error: 'sheet' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      const keys = Object.keys(b.rec || { ID: '' });
      sh.getRange(1, 1, 1, keys.length).setValues([keys]);
    }
    const values = sh.getDataRange().getValues();
    const head = values[0].map(String);
    let row = -1;
    for (let i = 1; i < values.length; i++) if (String(values[i][0]) === String(b.id)) { row = i + 1; break; }

    if (b.action === 'delete') {
      if (lv < 2 && name !== SHEETS.schedule) return out_({ ok: false, error: 'level' });   // 매물·손님 삭제는 열람 비밀번호가 있어야 (일정은 누구나)
      if (row > 0) sh.deleteRow(row);
      return out_({ ok: true });
    }
    const rec = b.rec || {};
    delete rec._locked;
    // 사무실 비밀번호로는 기존 매물의 비밀 칸을 덮어쓰지 못함 (새 매물 접수는 가능)
    if (lv < 2 && row > 0) (SECRET[name] || []).forEach(function (h) { delete rec[h]; });

    Object.keys(rec).forEach(function (h) {
      if (head.indexOf(h) < 0) { head.push(h); sh.getRange(1, head.length).setValue(h); }
    });
    const cur = row > 0 ? sh.getRange(row, 1, 1, head.length).getValues()[0] : head.map(function () { return ''; });
    head.forEach(function (h, i) { if (h in rec) cur[i] = rec[h]; });
    const target = row > 0 ? row : sh.getLastRow() + 1;
    const rng = sh.getRange(target, 1, 1, head.length);
    rng.setNumberFormat('@');           // 날짜·전화번호가 자동 변환되지 않도록 텍스트로 저장
    rng.setValues([cur.map(function (x) { return x == null ? '' : String(x); })]);
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
