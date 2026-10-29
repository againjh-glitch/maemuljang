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
 *
 * 텔레그램 알림 (선택) — 새 매물이 등록되면 알림방으로 메시지
 *   B3 칸: 텔레그램 봇 토큰 (@BotFather 에서 받은 값)
 *   B4 칸: 알림방 번호 (비워 두고 '텔레그램_방찾기' 를 실행하면 자동으로 채워짐)
 *   B5 칸: 알림에 집주인 연락처 넣기 (O / 비우면 안 넣음)
 *   B6 칸: 사이트 주소 (알림 맨 아래 링크)
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
  const tg = tgSettings_();
  Logger.log('텔레그램 알림: ' + (tg.token && tg.chat ? '켜짐' : tg.token ? '방 번호 없음 → 텔레그램_방찾기 실행' : '꺼짐 (설정 탭 B3에 봇 토큰을 넣으면 켜짐)'));
}

/* ---------- 텔레그램 알림 ---------- */
function tgSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('설정');
  if (!sh) return {};
  const v = sh.getRange('B3:B6').getDisplayValues().map(function (r) { return String(r[0]).trim(); });
  return { token: v[0], chat: v[1], withPhone: /^(o|O|0|예|포함)$/.test(v[2]), site: v[3] };
}
function tgSend_(text, chat) {
  const tg = tgSettings_();
  if (!tg.token || !(chat || tg.chat)) return { ok: false, description: '설정 탭 B3(봇 토큰)·B4(방 번호)를 확인하세요' };
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + tg.token + '/sendMessage', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ chat_id: chat || tg.chat, text: text, disable_web_page_preview: true })
  });
  return JSON.parse(res.getContentText());
}
/** 봇을 알림방에 초대한 뒤 실행: 방 번호를 찾아 설정 탭 B4에 적습니다 */
function 텔레그램_방찾기() {
  const tg = tgSettings_();
  if (!tg.token) { Logger.log('설정 탭 B3에 봇 토큰을 먼저 넣으세요.'); return; }
  const res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + tg.token + '/getUpdates', { muteHttpExceptions: true }).getContentText());
  if (!res.ok) { Logger.log('토큰이 맞지 않습니다: ' + res.description); return; }
  const chats = {};
  (res.result || []).forEach(function (u) {
    const c = (u.message || u.my_chat_member || u.channel_post || u.edited_message || {}).chat;
    if (c && c.type !== 'private') chats[c.id] = c.title || c.id;
    if (c && c.type === 'private') chats[c.id] = (c.first_name || '') + ' (개인 대화)';
  });
  const ids = Object.keys(chats);
  if (!ids.length) { Logger.log('찾은 방이 없습니다. 봇을 알림방에 초대한 뒤, 그 방에 아무 메시지나 하나 보내고 다시 실행하세요.'); return; }
  ids.forEach(function (id) { Logger.log('찾은 방: ' + chats[id] + ' → 번호 ' + id); });
  const group = ids.filter(function (id) { return !/개인 대화/.test(chats[id]); });
  const pick = group.length ? group[group.length - 1] : ids[0];
  SpreadsheetApp.getActive().getSheetByName('설정').getRange('B4').setNumberFormat('@').setValue(String(pick));
  Logger.log('설정 탭 B4에 "' + chats[pick] + '" 방 번호를 적었습니다. 다른 방이면 B4를 위 번호로 바꾸세요.');
}
/** 알림방으로 시험 메시지 보내기 */
function 텔레그램_테스트() {
  const r = tgSend_('✅ 매물장 알림 연결 테스트입니다. 새 매물이 등록되면 이 방으로 알려 드립니다.');
  Logger.log(r.ok ? '보냈습니다. 텔레그램 알림방을 확인하세요.' : '보내지 못했습니다: ' + r.description);
}
function man_(n) { n = Math.round(+n || 0); if (n >= 10000) { const e = Math.floor(n / 10000), r = n % 10000; return e + '억' + (r ? ' ' + r.toLocaleString() : ''); } return n.toLocaleString(); }
function dealsText_(s) {
  return String(s || '').replace(/(매매|전세|월세)\s*([0-9]+)(?:\/([0-9]+))?/g, function (_, t, p, r) { return t + ' ' + man_(p) + (r ? '/' + r : ''); });
}
function notifyNewListing_(r) {
  const tg = tgSettings_(); if (!tg.token || !tg.chat) return;
  const g = function (k) { return String(r[k] || '').trim(); };
  const coop = g('상대부동산') || /공동중개/.test(g('특징'));
  const spec = [g('연식') && g('연식') + '년', g('평수') && g('평수') + '평', (g('방') || g('화장실')) && ('방' + (g('방') || '?') + '/욕' + (g('화장실') || '?')), g('향'),
    g('엘베') === 'O' ? '엘베O' : g('엘베') === 'X' ? '엘베X' : '', g('주차') === 'O' || g('주차') === '0' ? '주차O' : g('주차') === 'X' ? '주차X' : ''].filter(String).join(' · ');
  const lines = [
    '🏠 새 매물' + (coop ? ' [공동중개]' : '') + ' — ' + (g('종류') || ''),
    '[' + dealsText_(g('거래조건')) + ']',
    [g('단지·건물'), g('주소')].filter(String).join(' ') + (g('동네') ? ' (' + g('동네') + ')' : ''),
    spec,
    g('특징') ? '특징: ' + g('특징') : '',
    g('메모') ? '메모: ' + g('메모') : '',
    coop && g('상대부동산') ? '상대 부동산: ' + g('상대부동산') + ' ' + g('상대부동산연락처') : '',
    tg.withPhone && g('연락처') ? '연락처: ' + g('연락처') : '',
    g('네이버링크') ? '네이버: ' + g('네이버링크') : '',
    tg.site ? '👉 ' + tg.site : ''
  ].filter(String);
  try { tgSend_(lines.join('\n')); } catch (e) { /* 알림 실패해도 저장은 그대로 */ }
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
  let newListing = null;
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
    if (row < 0 && name === SHEETS.listings) newListing = rec;   // 새 매물만 알림
  } finally {
    lock.releaseLock();
  }
  if (newListing) notifyNewListing_(newListing);
  return out_({ ok: true });
}
