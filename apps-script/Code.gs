/**
 * 매물장 응대 데스크 — 구글 시트 연결용 Apps Script
 * 1) 아래 PASSWORD 를 사무실 비밀번호로 바꾸세요 (직원들이 사이트 설정에 입력할 값).
 * 2) 배포 > 새 배포 > 유형: 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자 → 배포 → 웹 앱 URL 복사.
 * 3) 이 코드를 고친 뒤에는 배포 > 배포 관리 > 수정(연필) > 버전: 새 버전 으로 다시 배포해야 반영됩니다.
 * 이 파일의 비밀번호는 깃허브에 올리지 마세요. (깃허브에는 '여기에-비밀번호' 그대로 둡니다)
 */
const PASSWORD = '여기에-비밀번호';
const SHEETS = { listings: '매물', clients: '손님' };

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function readSheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return [];
  const v = sh.getDataRange().getDisplayValues();
  const head = v[0];
  return v.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

function doGet(e) {
  if ((e.parameter.key || '') !== PASSWORD) return out_({ ok: false, error: 'password' });
  return out_({ ok: true, listings: readSheet_(SHEETS.listings), clients: readSheet_(SHEETS.clients) });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.key !== PASSWORD) return out_({ ok: false, error: 'password' });
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
      if (row > 0) sh.deleteRow(row);
      return out_({ ok: true });
    }
    // 새 항목이 생기면 열을 오른쪽에 추가
    Object.keys(b.rec).forEach(function (h) {
      if (head.indexOf(h) < 0) { head.push(h); sh.getRange(1, head.length).setValue(h); }
    });
    const cur = row > 0 ? sh.getRange(row, 1, 1, head.length).getValues()[0] : head.map(function () { return ''; });
    head.forEach(function (h, i) { if (h in b.rec) cur[i] = b.rec[h]; });
    const target = row > 0 ? row : sh.getLastRow() + 1;
    const rng = sh.getRange(target, 1, 1, head.length);
    rng.setNumberFormat('@');           // 날짜·전화번호가 자동 변환되지 않도록 텍스트로 저장
    rng.setValues([cur.map(function (x) { return x == null ? '' : String(x); })]);
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
