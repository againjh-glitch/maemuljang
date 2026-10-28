/**
 * 매물장 응대 데스크 — 구글 시트 연결용 Apps Script
 *
 * ■ 비밀번호 두 개 (이 파일 안에서만 바꿉니다. 깃허브에는 실제 비밀번호를 올리지 마세요)
 *   PASSWORD      : 사무실 비밀번호. 매물 보기·검색·손님 응대·새 매물 접수 가능.
 *                   연락처, 현관 비밀번호, 내놓은 분 이름, 원문은 서버에서 아예 보내지 않습니다.
 *   FULL_PASSWORD : 열람 비밀번호. 위 정보까지 모두 보고 수정할 수 있습니다.
 *
 * ■ 바꾼 뒤에는 반드시: 배포 > 배포 관리 > 연필 > 버전: 새 버전 > 배포
 *
 * ■ 처음 배포: 배포 > 새 배포 > 유형: 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자
 *   구글 시트 자체는 관리자만 볼 수 있게 공유하지 마세요(시트에는 모든 정보가 있습니다).
 */
const PASSWORD = '여기에-사무실-비밀번호'; 0114
const FULL_PASSWORD = '여기에-열람-비밀번호';yw0114

const SHEETS = { listings: '매물', clients: '손님' };
// 열람 비밀번호가 있어야 보이는 칸
const SECRET = {
  '매물': ['연락처', '내놓은분', '집보는법', '원문'],
  '손님': ['연락처', '원문']
};

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function level_(key) {
  key = String(key || '');
  const bad = /^여기에-/;
  if (key && key === FULL_PASSWORD && !bad.test(FULL_PASSWORD)) return 2;
  if (key && key === PASSWORD && !bad.test(PASSWORD)) return 1;
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
  const lv = level_(e.parameter.key);
  if (!lv) return out_({ ok: false, error: 'password' });
  return out_({ ok: true, level: lv, listings: readSheet_(SHEETS.listings, lv), clients: readSheet_(SHEETS.clients, lv) });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  const lv = level_(b.key);
  if (!lv) return out_({ ok: false, error: 'password' });
  const name = SHEETS[b.sheet];
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
      if (lv < 2) return out_({ ok: false, error: 'level' });   // 삭제는 열람 비밀번호가 있어야
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
