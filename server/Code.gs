/**
 * 아침 브리핑 앱 — Google Sheets + Apps Script 웹앱
 * - 시트: 대시보드 / 브리핑 / 메모 / 댓글 / 일정 / 할일 / 지표 / 설정 / 항목 / 씨앗
 * - doGet  : ?k=<키>            → 아이폰·PC용 앱 화면
 *            ?api=pull&key=<키> → 매일 루틴이 읽는 JSON
 * - doPost : {key, action:"push"|"import", ...} → 매일 루틴이 쓰는 JSON
 * - 메모·브리핑 "확인" 표시: 확인시각(seenAt) 이후 Claude 댓글이 달리거나(브리핑은 재작성 포함) 하면 다시 미확인.
 * - 항목: 브리핑을 주제별 카드로 쪼갠 것. [확인]하면 "오늘"에서 빠지고 "관리"의 카테고리(금융·부동산…)로 쌓인다.
 */
const TZ = 'Asia/Seoul';

const SECTIONS = [
  ['sec_indicators', '관심 지표'], ['sec_schedule', '일정'], ['sec_todo', '할 일'],
  ['sec_tracking', '고정 추적'], ['sec_decisions', '결정사항'], ['sec_analysis', '분석'],
  ['sec_thoughts', '생각·용어'], ['sec_replies', '댓글 답변'], ['sec_etc', '기타'],
  ['sec_questions', '되묻는 질문'],
];

const SCHEMA = {
  briefs: { sheet: '브리핑', key: 'date', bools: ['pinned'], cols: [
    ['date', '날짜'], ['header', '헤더'], ...SECTIONS,
    ['memoCount', '메모수'], ['commentCount', '댓글수'], ['pinned', '고정'], ['createdAt', '작성시각'], ['seenAt', '확인시각']] },
  memos: { sheet: '메모', key: 'id', bools: ['pinned', 'mpin', 'hidden'], cols: [
    ['id', 'id'], ['day', '날짜'], ['text', '내용'], ['status', '상태'], ['category', '카테고리'],
    ['pinned', '고정'], ['createdAt', '작성시각'], ['updatedAt', '수정시각'], ['sentAt', '처리시각'], ['source', '출처'],
    ['seenAt', '확인시각'], ['domain', '관리'], ['folder', '폴더'], ['mpin', '관리고정'], ['hidden', '숨김'], ['sort', '관리순서']] },
  items: { sheet: '항목', key: 'id', bools: ['pinned', 'mpin', 'hidden'], cols: [
    ['id', 'id'], ['date', '날짜'], ['section', '섹션'], ['title', '제목'], ['body', '내용'], ['domain', '관리'],
    ['createdAt', '작성시각'], ['seenAt', '확인시각'], ['pinned', '고정'], ['memoIds', '메모id'],
    ['folder', '폴더'], ['mpin', '관리고정'], ['hidden', '숨김'], ['sort', '관리순서']] },
  comments: { sheet: '댓글', key: 'id', bools: ['seen'], cols: [
    ['id', 'id'], ['targetType', '대상'], ['targetId', '대상id'], ['by', '작성자'], ['text', '내용'],
    ['createdAt', '작성시각'], ['seen', '확인']] },
  schedules: { sheet: '일정', key: 'id', bools: ['done'], cols: [
    ['id', 'id'], ['title', '제목'], ['date', '날짜'], ['time', '시간'], ['done', '완료'], ['source', '출처'],
    ['memoId', '메모id'], ['note', '비고'], ['createdAt', '작성시각']] },
  todos: { sheet: '할일', key: 'id', bools: ['done'], cols: [
    ['id', 'id'], ['text', '내용'], ['done', '완료'], ['source', '출처'], ['memoId', '메모id'],
    ['createdAt', '작성시각'], ['doneAt', '완료시각']] },
  watchlist: { sheet: '지표', key: 'symbol', bools: [], cols: [
    ['order', '순서'], ['label', '이름'], ['symbol', '심볼'], ['enabled', '포함'], ['prompt', '프롬프트'], ['price', '현재가'],
    ['d1', '전일비'], ['d5', '5일'], ['comment', '코멘트'], ['updatedAt', '갱신시각']] },
  seeds: { sheet: '씨앗', key: 'id', bools: [], cols: [
    ['id', 'id'], ['date', '날짜'], ['name', '영역'], ['count', '개수'], ['updatedAt', '수정시각']] },
  settings: { sheet: '설정', key: 'key', bools: [], cols: [['key', '키'], ['value', '값']] },
};

const CATEGORIES = ['일정', '할 일', '분석', '결정', '생각', '기타'];
const DEFAULT_DOMAINS = ['금융', '부동산', '인사이트', '전공', '사업', '기타'];
const TABS = ['today', 'memo', 'plan', 'history', 'watch', 'manage'];
const SCHEMA_VERSION = 'v6';
const DEFAULT_SEEDS = ['전공', '영어', '금융', '부동산', '연애', '통제'];
const SEED_MAX = 3;
// 화면에서 직접 고칠 수 있는 칸 (앞으로 화면 기능을 늘려도 서버를 다시 붙여넣지 않도록 넓게 둔다)
const EDITABLE = {
  memos: ['text', 'day', 'category', 'domain', 'folder', 'mpin', 'hidden', 'sort', 'pinned', 'seenAt', 'status'],
  items: ['section', 'title', 'body', 'domain', 'folder', 'mpin', 'hidden', 'sort', 'pinned', 'seenAt', 'memoIds', 'date'],
  schedules: ['title', 'date', 'time', 'done', 'note'],
  todos: ['text', 'done'],
};

/* 한 번의 실행 안에서 같은 시트를 여러 번 읽지 않도록 캐시 (쓰기 때마다 비움) */
const RC_ = {}, HC_ = {};
function invalidate_(name) { delete RC_[name]; }

/* ───────────── 공통 유틸 ───────────── */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function nowIso_() { return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function newId_(prefix) { return (prefix || '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16); }
function apiKey_() {
  const k = PropertiesService.getScriptProperties().getProperty('API_KEY');
  return k || (typeof INSTALL_KEY !== 'undefined' ? INSTALL_KEY : '');
}
/** 설치 스크립트로 배포한 경우 첫 접속 때 시트 구성·이관을 자동으로 한다. */
function ensureSetup_() {
  if (ss_().getSheetByName('설정') && PropertiesService.getScriptProperties().getProperty('API_KEY')) {
    const cache = CacheService.getScriptCache();
    if (cache.get('schema') === SCHEMA_VERSION) return;
    ensureSheet_('items'); ensureSheet_('seeds'); ensureColumns_('watchlist'); ensureColumns_('memos'); ensureColumns_('briefs'); ensureColumns_('items');
    cache.put('schema', SCHEMA_VERSION, 21600);
    return;
  }
  withLock_(() => {
    if (ss_().getSheetByName('설정') && PropertiesService.getScriptProperties().getProperty('API_KEY')) return;
    setup();
  });
}
function checkKey_(k) {
  const key = apiKey_();
  if (!key || String(k || '') !== key) throw new Error('인증 키가 올바르지 않습니다.');
}
function toBool_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1; }

function cellToValue_(v, field) {
  if (v instanceof Date) {
    if (field === 'time') return Utilities.formatDate(v, TZ, 'HH:mm');
    if (field === 'date' || field === 'day') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
    return v.toISOString();
  }
  return v;
}

/** 스키마에 새로 추가된 시트가 없으면 머리글만 넣어 만든다. */
function ensureSheet_(name) {
  const s = SCHEMA[name];
  if (ss_().getSheetByName(s.sheet)) return;
  const sh = ss_().insertSheet(s.sheet);
  const labels = s.cols.map(c => c[1]);
  sh.getRange(1, 1, 1, labels.length).setValues([labels]).setFontWeight('bold').setBackground('#eef2f7');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), labels.length).setNumberFormat('@');
}

/** 스키마에 새로 추가된 열이 시트에 없으면 맨 끝에 붙인다. */
function ensureColumns_(name) {
  const sh = ss_().getSheetByName(SCHEMA[name].sheet);
  if (!sh) return;
  const header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  SCHEMA[name].cols.forEach(([, label]) => {
    if (header.indexOf(label) < 0) {
      const col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(label).setFontWeight('bold').setBackground('#eef2f7');
      sh.getRange(2, col, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
      header.push(label);
      delete HC_[name];
    }
  });
}

function sheetOf_(name) {
  const s = SCHEMA[name];
  const sh = ss_().getSheetByName(s.sheet);
  if (!sh) throw new Error('시트가 없습니다: ' + s.sheet + ' (메뉴 > 브리핑 앱 > 초기 설정 실행)');
  return sh;
}

/** 헤더 이름으로 열을 찾으므로 사용자가 시트에서 열 순서를 바꿔도 동작한다. */
function colMap_(name, sh) {
  if (HC_[name]) return HC_[name];
  const s = SCHEMA[name];
  const header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  const map = {};
  s.cols.forEach(([field, label]) => {
    const i = header.indexOf(label);
    if (i >= 0) map[field] = i;
  });
  return (HC_[name] = { map, width: header.length });
}

function readAll_(name) {
  if (RC_[name]) return RC_[name].map(o => Object.assign({}, o));
  const s = SCHEMA[name];
  const sh = sheetOf_(name);
  const { map } = colMap_(name, sh);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const out = [];
  values.forEach((row, i) => {
    const o = { _row: i + 2 };
    Object.keys(map).forEach(f => {
      let v = cellToValue_(row[map[f]], f);
      if (s.bools.indexOf(f) >= 0) v = toBool_(v);
      else if (v === null || v === undefined) v = '';
      o[f] = v;
    });
    if (String(o[s.key] || '') !== '') out.push(o);
  });
  RC_[name] = out;
  return out.map(o => Object.assign({}, o));
}

function rowFrom_(name, sh, obj, base) {
  const { map, width } = colMap_(name, sh);
  const row = base ? base.slice() : new Array(width).fill('');
  Object.keys(obj).forEach(f => {
    if (f in map && f !== '_row') row[map[f]] = obj[f] === undefined || obj[f] === null ? '' : obj[f];
  });
  return row;
}

function append_(name, objs) {
  if (!objs.length) return;
  const sh = sheetOf_(name);
  const rows = objs.map(o => rowFrom_(name, sh, o));
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  invalidate_(name);
}

function patchRow_(name, rowNum, patch) {
  const sh = sheetOf_(name);
  const width = sh.getLastColumn();
  const base = sh.getRange(rowNum, 1, 1, width).getValues()[0];
  sh.getRange(rowNum, 1, 1, width).setValues([rowFrom_(name, sh, patch, base)]);
  invalidate_(name);
}

function findBy_(name, field, value) {
  return readAll_(name).find(o => String(o[field]) === String(value)) || null;
}

function upsert_(name, obj) {
  const k = SCHEMA[name].key;
  const found = findBy_(name, k, obj[k]);
  if (found) patchRow_(name, found._row, obj); else append_(name, [obj]);
}

function deleteBy_(name, field, value) {
  const sh = sheetOf_(name);
  readAll_(name).filter(o => String(o[field]) === String(value))
    .map(o => o._row).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
  invalidate_(name);
}

function strip_(o) { const c = Object.assign({}, o); delete c._row; return c; }

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function domains_() {
  const v = getSetting_('관리카테고리');
  const list = v ? v.split(',').map(x => x.trim()).filter(Boolean) : [];
  return list.length ? list : DEFAULT_DOMAINS.slice();
}
function jsonSetting_(k, dflt) { try { const v = getSetting_(k); return v ? JSON.parse(v) : dflt; } catch (e) { return dflt; } }
function tabs_() {
  const t = jsonSetting_('탭설정', {});
  const order = (t.order || []).filter(x => TABS.indexOf(x) >= 0);
  TABS.forEach(x => { if (order.indexOf(x) < 0) order.push(x); });
  return { order, hidden: (t.hidden || []).filter(x => TABS.indexOf(x) >= 0) };
}
function folders_() { return jsonSetting_('관리폴더', {}); }
function seedList_() { const l = jsonSetting_('씨앗목록', null); return Array.isArray(l) && l.length ? l : DEFAULT_SEEDS.slice(); }
function seeds_() { return readAll_('seeds').map(x => ({ date: String(x.date), name: String(x.name), count: Math.max(0, Math.min(SEED_MAX, +x.count || 0)) })).filter(x => x.count); }
function getSetting_(k) { const r = findBy_('settings', 'key', k); return r ? String(r.value) : ''; }
function setSetting_(k, v) { upsert_('settings', { key: k, value: v }); }

/* ───────────── 초기 설정 ───────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('브리핑 앱')
    .addItem('초기 설정 (시트 만들기·키 발급)', 'setup')
    .addItem('앱 주소·키 다시 보기', 'showKey')
    .addToUi();
}

function setup() {
  const ss = ss_();
  Object.keys(SCHEMA).forEach(name => {
    const s = SCHEMA[name];
    let sh = ss.getSheetByName(s.sheet);
    if (!sh) sh = ss.insertSheet(s.sheet);
    if (sh.getLastRow() === 0) {
      const labels = s.cols.map(c => c[1]);
      sh.getRange(1, 1, 1, labels.length).setValues([labels]).setFontWeight('bold').setBackground('#eef2f7');
      sh.setFrozenRows(1);
      const body = n => sh.getRange(2, n, sh.getMaxRows() - 1, 1);
      s.cols.forEach(([f], i) => {
        if (s.bools.indexOf(f) >= 0) {
          body(i + 1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
        } else if (['order', 'price', 'd1', 'd5', 'memoCount', 'commentCount'].indexOf(f) < 0) {
          body(i + 1).setNumberFormat('@');
        }
      });
      sh.getRange(2, 1, sh.getMaxRows() - 1, labels.length).setVerticalAlignment('top');
      sh.getRange(1, 1, sh.getMaxRows(), labels.length).setWrap(true);
    }
  });
  const briefSh = ss.getSheetByName('브리핑');
  briefSh.setColumnWidths(3, SECTIONS.length, 280);
  ss.getSheetByName('메모').setColumnWidth(3, 360);
  ss.getSheetByName('일정').setColumnWidth(2, 260);
  ss.getSheetByName('할일').setColumnWidth(2, 320);

  if (!getSetting_('지표공통프롬프트')) setSetting_('지표공통프롬프트', '');

  if (typeof MIGRATION !== 'undefined' && !getSetting_('이관완료')) {
    importData_(MIGRATION);
    setSetting_('이관완료', nowIso_());
  }
  buildDashboard_();

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('API_KEY')) props.setProperty('API_KEY',
    typeof INSTALL_KEY !== 'undefined' ? INSTALL_KEY : Utilities.getUuid().replace(/-/g, ''));
  const blank = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);
  showKey();
}

function showKey() {
  const key = apiKey_();
  const url = ScriptApp.getService().getUrl();
  const msg = '인증 키: ' + key + '\n\n' +
    (url ? '앱 주소(아이폰 홈 화면·PC 즐겨찾기에 추가): ' + url + '?k=' + key
         : '아직 웹 앱으로 배포되지 않았습니다. 배포 > 새 배포 > 웹 앱 후 이 메뉴를 다시 열어 주세요.');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* 편집기에서 실행 시 로그로만 */ }
}

function buildDashboard_() {
  const ss = ss_();
  let sh = ss.getSheetByName('대시보드');
  if (!sh) sh = ss.insertSheet('대시보드', 0);
  sh.clear();
  const cells = [
    ['A1', '📋 아침 브리핑 대시보드'],
    ['A3', '최근 브리핑'], ['B3', '=IFERROR(INDEX(SORT(FILTER(브리핑!A2:B, 브리핑!A2:A<>""), 1, FALSE), 1, 1), "")'],
    ['C3', '=IFERROR(INDEX(SORT(FILTER(브리핑!A2:B, 브리핑!A2:A<>""), 1, FALSE), 1, 2), "")'],
    ['A4', '처리 대기 메모'], ['B4', '=COUNTIF(메모!D2:D, "new")'],
    ['A5', '새 댓글'], ['B5', '=COUNTIFS(댓글!D2:D, "user", 댓글!G2:G, FALSE)'],
    ['A6', '미완료 일정'], ['B6', '=COUNTIFS(일정!B2:B, "<>", 일정!E2:E, FALSE)'],
    ['A7', '미완료 할 일'], ['B7', '=COUNTIFS(할일!B2:B, "<>", 할일!C2:C, FALSE)'],
    ['A9', '관심 지표'],
    ['A10', '=IFERROR(SORT(FILTER({지표!A2:A, 지표!B2:B, 지표!E2:E, 지표!F2:F, 지표!G2:G}, 지표!B2:B<>""), 1, TRUE), "지표 없음")'],
    ['G9', '다가오는 일정'],
    ['G10', '=IFERROR(SORT(FILTER({일정!C2:C, 일정!D2:D, 일정!B2:B}, 일정!B2:B<>"", 일정!E2:E=FALSE), 1, TRUE), "일정 없음")'],
    ['G22', '할 일'],
    ['G23', '=IFERROR(FILTER(할일!B2:B, 할일!B2:B<>"", 할일!C2:C=FALSE), "할 일 없음")'],
    ['A22', '오늘의 되묻는 질문'],
    ['A23', '=IFERROR(INDEX(SORT(FILTER(브리핑!A2:L, 브리핑!A2:A<>""), 1, FALSE), 1, 12), "")'],
  ];
  cells.forEach(([a1, v]) => {
    const r = sh.getRange(a1);
    if (String(v).charAt(0) === '=') r.setFormula(v); else r.setValue(v);
  });
  sh.getRange('A1').setFontSize(16).setFontWeight('bold');
  ['A3:A7', 'A9', 'G9', 'G22', 'A22'].forEach(a => sh.getRange(a).setFontWeight('bold'));
  sh.getRange('A23:E23').merge().setWrap(true).setVerticalAlignment('top');
  sh.setRowHeight(23, 120);
  sh.setColumnWidth(1, 120); sh.setColumnWidth(9, 280);
}

/* ───────────── 웹 진입점 ───────────── */

function doGet(e) {
  const p = (e && e.parameter) || {};
  ensureSetup_();
  if (p.api) {
    try {
      checkKey_(p.key);
      if (p.api === 'pull') return json_(apiPull_());
      throw new Error('알 수 없는 api: ' + p.api);
    } catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
  }
  if (!apiKey_() || p.k !== apiKey_()) {
    return HtmlService.createHtmlOutput('<p style="font:16px sans-serif;padding:24px">앱 주소 끝에 ?k=인증키 를 붙여서 여세요.</p>')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  const t = HtmlService.createTemplateFromFile('Index');
  t.key = p.k;
  return t.evaluate().setTitle('아침 브리핑')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function doPost(e) {
  try {
    ensureSetup_();
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    checkKey_(body.key);
    if (body.action === 'push') return json_(withLock_(() => apiPush_(body)));
    if (body.action === 'ui') return json_(uiDispatch_(body));
    if (body.action === 'import') return json_(withLock_(() => ({ ok: true, counts: importData_(body) })));
    throw new Error('알 수 없는 action: ' + body.action);
  } catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────────── 루틴용 API ───────────── */

function commentsIndex_() {
  const idx = {};
  readAll_('comments').forEach(c => {
    const k = c.targetType + ':' + c.targetId;
    (idx[k] = idx[k] || []).push(strip_(c));
  });
  Object.keys(idx).forEach(k => idx[k].sort((a, b) => String(a.createdAt).localeCompare(String(b.createdAt))));
  return idx;
}

function attach_(type, o, idx) {
  const id = type === 'brief' ? o.date : o.id;
  const comments = idx[type + ':' + id] || [];
  return Object.assign(strip_(o), {
    comments, hasNewComment: comments.some(c => c.by === 'user' && !c.seen),
  });
}

function watchlist_() {
  return readAll_('watchlist').map(w => Object.assign(strip_(w), { enabled: !(w.enabled === false || w.enabled === 'FALSE') }))
    .sort((a, b) => (Number(a.order) || 0) - (Number(b.order) || 0));
}

function apiPull_() {
  const idx = commentsIndex_();
  const memos = readAll_('memos').map(m => attach_('memo', m, idx));
  const briefs = readAll_('briefs').map(b => attach_('brief', b, idx))
    .sort((a, b) => String(b.date).localeCompare(String(a.date)));
  return {
    ok: true,
    today: today_(),
    newMemos: memos.filter(m => m.status === 'new'),
    memosWithNewComments: memos.filter(m => m.hasNewComment),
    pinnedMemos: memos.filter(m => m.pinned),
    briefsWithNewComments: briefs.filter(b => b.hasNewComment),
    pinnedBriefs: briefs.filter(b => b.pinned),
    recentBriefs: briefs.slice(0, 3),
    schedules: readAll_('schedules').filter(s => !s.done).map(strip_)
      .sort((a, b) => String(a.date).localeCompare(String(b.date))),
    todos: readAll_('todos').filter(t => !t.done).map(strip_),
    watchlist: { items: watchlist_().filter(w => w.enabled), prompt: getSetting_('지표공통프롬프트'),
      all: watchlist_(), updatedAt: getSetting_('지표수정시각') },
    domains: domains_(),
    itemsWithNewComments: readAll_('items').map(i => attach_('item', i, idx)).filter(i => i.hasNewComment),
    memosNoDomain: memos.filter(m => !m.domain).slice(0, 50).map(m => ({ id: m.id, day: m.day, text: m.text })),
    pinnedItems: readAll_('items').map(i => attach_('item', i, idx)).filter(i => i.pinned),
    seedList: seedList_(),
    seedsRecent: seeds_().filter(x => x.date >= Utilities.formatDate(new Date(Date.now() - 14 * 864e5), TZ, 'yyyy-MM-dd'))
      .sort((a, b) => a.date.localeCompare(b.date)),
  };
}

/** 지표 목록 통째 교체. enabled는 기존 값을 심볼 기준으로 유지 */
function replaceWatchlist_(items, prompt, updatedAt) {
  const old = {};
  readAll_('watchlist').forEach(w => { old[w.symbol] = w; });
  const sh = sheetOf_('watchlist');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent(); invalidate_('watchlist');
  append_('watchlist', (items || []).filter(it => it && it.symbol).map((it, i) => {
    const o = old[it.symbol] || {};
    const en = it.enabled !== undefined ? it.enabled !== false : !(o.enabled === false || o.enabled === 'FALSE');
    return { order: i + 1, label: it.label || it.symbol, symbol: it.symbol, prompt: it.prompt || '', enabled: en,
      price: o.price || '', d1: o.d1 || '', d5: o.d5 || '', comment: o.comment || '', updatedAt: o.updatedAt || '' };
  }));
  setSetting_('지표공통프롬프트', prompt || '');
  setSetting_('지표수정시각', updatedAt || nowIso_());
}

/**
 * body = {
 *   brief: {date, header, sections:{sec_indicators:"…", …}, memoCount, commentCount},
 *   indicators: [{symbol, price, d1, d5, comment}],
 *   schedules: [{title, date, time, memoId, note}],
 *   todos: [{text, memoId}],
 *   memoUpdates: [{id, status, category}],
 *   comments: [{targetType:"memo"|"brief", targetId, text}],
 *   seenCommentIds: ["…"]
 * }
 */
function apiPush_(body) {
  const now = nowIso_();
  const counts = { brief: 0, indicators: 0, schedules: 0, todos: 0, memos: 0, comments: 0, seen: 0 };

  if (body.brief && body.brief.date) {
    const b = body.brief, row = { date: b.date, header: b.header || '', createdAt: now,
      memoCount: b.memoCount || 0, commentCount: b.commentCount || 0 };
    SECTIONS.forEach(([f]) => { row[f] = (b.sections && b.sections[f]) || ''; });
    const found = findBy_('briefs', 'date', b.date);
    if (found) patchRow_('briefs', found._row, row);
    else append_('briefs', [Object.assign(row, { pinned: false })]);
    counts.brief = 1;
  }

  if (body.watchlist && body.watchlist.items) {
    replaceWatchlist_(body.watchlist.items, body.watchlist.prompt, body.watchlist.updatedAt);
    counts.watchlist = body.watchlist.items.length;
  }

  if (body.indicators && body.indicators.length) {
    const wl = readAll_('watchlist');
    body.indicators.forEach(ind => {
      const r = wl.find(w => w.symbol === ind.symbol);
      if (!r) return;
      patchRow_('watchlist', r._row, { price: ind.price == null ? '' : ind.price, d1: ind.d1 == null ? '' : ind.d1,
        d5: ind.d5 == null ? '' : ind.d5, comment: ind.comment || '', updatedAt: now });
      counts.indicators++;
    });
  }

  if (body.schedules && body.schedules.length) {
    const open = readAll_('schedules');
    const add = body.schedules.filter(s => s.title && s.date &&
      !open.some(o => o.title === s.title && o.date === s.date))
      .map(s => ({ id: s.id || newId_('s'), title: s.title, date: s.date, time: s.time || '', done: false,
        source: 'claude', memoId: s.memoId || '', note: s.note || '', createdAt: now }));
    append_('schedules', add); counts.schedules = add.length;
  }

  if (body.todos && body.todos.length) {
    const cur = readAll_('todos');
    const add = body.todos.filter(t => t.text && !cur.some(c => !c.done && c.text === t.text))
      .map(t => ({ id: t.id || newId_('t'), text: t.text, done: false, source: 'claude',
        memoId: t.memoId || '', createdAt: now, doneAt: '' }));
    append_('todos', add); counts.todos = add.length;
  }

  if (body.memoUpdates && body.memoUpdates.length) {
    const memos = readAll_('memos'), doms = domains_();
    body.memoUpdates.forEach(u => {
      const m = memos.find(x => x.id === u.id);
      if (!m) return;
      const patch = {};
      if (u.status) { patch.status = u.status; patch.sentAt = now; }
      if (u.category && !m.category && CATEGORIES.indexOf(u.category) >= 0) patch.category = u.category;
      if (u.domain && !m.domain && doms.indexOf(u.domain) >= 0) patch.domain = u.domain;
      if (Object.keys(patch).length) { patchRow_('memos', m._row, patch); counts.memos++; }
    });
  }

  // 브리핑 항목: 같은 날짜로 다시 보내면 아직 확인·댓글 없는 항목만 새것으로 바꾼다
  if (body.items && body.items.length && body.brief && body.brief.date) {
    const date = body.brief.date, doms = domains_(), cmt = commentsIndex_();
    const sh = sheetOf_('items');
    readAll_('items').filter(i => i.date === date && !i.seenAt && !(cmt['item:' + i.id] || []).length)
      .map(i => i._row).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
    invalidate_('items');
    const add = body.items.filter(it => it && (it.title || it.body)).map(it => ({
      id: newId_('i'), date, section: it.section || '', title: it.title || '', body: it.body || '',
      domain: doms.indexOf(it.domain) >= 0 ? it.domain : '기타', createdAt: now, seenAt: '', pinned: false,
      memoIds: (Array.isArray(it.memoIds) ? it.memoIds : String(it.memoIds || '').split(',')).map(x => String(x).trim()).filter(Boolean).join(',') }));
    append_('items', add); counts.items = add.length;
  }

  if (body.comments && body.comments.length) {
    const add = body.comments.filter(c => c.text && c.targetId)
      .map(c => ({ id: newId_('c'), targetType: ['brief', 'item'].indexOf(c.targetType) >= 0 ? c.targetType : 'memo',
        targetId: c.targetId, by: 'claude', text: c.text, createdAt: now, seen: true }));
    append_('comments', add); counts.comments = add.length;
  }

  if (body.seenCommentIds && body.seenCommentIds.length) {
    const ids = body.seenCommentIds.map(String);
    readAll_('comments').filter(c => ids.indexOf(String(c.id)) >= 0 && !c.seen)
      .forEach(c => { patchRow_('comments', c._row, { seen: true }); counts.seen++; });
  }

  return { ok: true, counts };
}

/** 기존 메모장 데이터 등을 id 기준으로 병합(upsert) */
function importData_(d) {
  const counts = {};
  const plan = [['memos', d.memos], ['comments', d.comments], ['schedules', d.schedules],
    ['todos', d.todos], ['briefs', d.briefs], ['watchlist', d.watchlist], ['items', d.items]];
  plan.forEach(([name, rows]) => {
    if (!rows || !rows.length) return;
    const k = SCHEMA[name].key;
    const existing = {};
    if (name === 'items') ensureSheet_('items');
    readAll_(name).forEach(o => { existing[String(o[k])] = o._row; });
    const fresh = [];
    rows.forEach(r => {
      if (existing[String(r[k])]) patchRow_(name, existing[String(r[k])], r); else fresh.push(r);
    });
    append_(name, fresh);
    counts[name] = rows.length;
  });
  if (d.watchlistPrompt !== undefined) setSetting_('지표공통프롬프트', d.watchlistPrompt || '');
  return counts;
}

/* ───────────── 앱 화면용 함수 (google.script.run 또는 외부 앱의 POST action:"ui") ───────────── */

const UI_FUNCS = { uiSetSeed, uiSaveSeedList, uiRenameSeed, uiLoad, uiAddMemo, uiEditMemo, uiDeleteMemo, uiTogglePin, uiAddComment, uiDeleteComment,
  uiSetSeen, uiSetDomain, uiSaveDomains, uiSetEntry, uiSetEntries, uiDeleteItem, uiSaveTabs, uiSaveFolders, uiRenameFolder, uiDeleteFolder, uiSaveSchedule, uiSetDone, uiDelete, uiAddTodo, uiSaveWatchlist, uiSearch, uiQuotes };

function uiDispatch_(body) {
  const fn = UI_FUNCS[body.fn];
  if (!fn) throw new Error('알 수 없는 fn: ' + body.fn);
  return { ok: true, data: fn.apply(null, [body.key].concat(body.args || [])) };
}

/* 🌱 씨앗: 날짜·영역마다 0~SEED_MAX 개. 0이면 행을 지운다. 빠른 응답을 위해 씨앗만 돌려준다 */
function uiSetSeed(key, date, name, count) {
  checkKey_(key);
  date = String(date || ''); name = String(name || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !name) throw new Error('날짜·영역을 확인해 주세요.');
  const n = Math.max(0, Math.min(SEED_MAX, Math.round(+count || 0))), id = date + '|' + name;
  withLock_(() => { if (n) upsert_('seeds', { id, date, name, count: n, updatedAt: nowIso_() }); else deleteBy_('seeds', 'id', id); });
  return { seeds: seeds_(), seedList: seedList_() };
}
function uiSaveSeedList(key, list) {
  checkKey_(key);
  const clean = (list || []).map(x => String(x).trim()).filter(Boolean).filter((x, i, a) => a.indexOf(x) === i).slice(0, 20);
  if (!clean.length) throw new Error('영역이 하나는 있어야 해요.');
  withLock_(() => setSetting_('씨앗목록', JSON.stringify(clean)));
  return { seeds: seeds_(), seedList: seedList_() };
}
function uiRenameSeed(key, from, to) {
  checkKey_(key);
  from = String(from || '').trim(); to = String(to || '').trim();
  if (!from || !to || from === to) return { seeds: seeds_(), seedList: seedList_() };
  withLock_(() => {
    const l = seedList_(); if (l.indexOf(to) >= 0) throw new Error('이미 있는 이름이에요.');
    setSetting_('씨앗목록', JSON.stringify(l.map(x => x === from ? to : x)));
    readAll_('seeds').filter(r => String(r.name) === from)
      .forEach(r => patchRow_('seeds', r._row, { id: String(r.date) + '|' + to, name: to }));
  });
  return { seeds: seeds_(), seedList: seedList_() };
}

function uiLoad(key) {
  checkKey_(key);
  const idx = commentsIndex_();
  const briefs = readAll_('briefs').map(b => attach_('brief', b, idx))
    .sort((a, b) => String(b.date).localeCompare(String(a.date))).slice(0, 90);
  const memos = readAll_('memos').map(m => attach_('memo', m, idx))
    .sort((a, b) => String(b.day).localeCompare(String(a.day)) || String(b.createdAt).localeCompare(String(a.createdAt)))
    .slice(0, 300);
  const schedules = readAll_('schedules').map(strip_)
    .sort((a, b) => (String(a.date) + a.time).localeCompare(String(b.date) + b.time));
  const todos = readAll_('todos').map(strip_)
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  const items = readAll_('items').map(i => attach_('item', i, idx))
    .sort((a, b) => String(b.date).localeCompare(String(a.date)) || String(a.createdAt).localeCompare(String(b.createdAt)))
    .slice(0, 800);
  return { today: today_(), sections: SECTIONS, categories: CATEGORIES, domains: domains_(), tabs: tabs_(), folders: folders_(),
    briefs, memos, items, schedules, todos, seeds: seeds_(), seedList: seedList_(), seedMax: SEED_MAX,
    watchlist: { items: watchlist_(), prompt: getSetting_('지표공통프롬프트') } };
}

/** analyze=false 면 "관리 노트"(status=note)로 저장해 아침 분석에서 빠진다 */
function uiAddMemo(key, text, category, domain, analyze, folder) {
  checkKey_(key);
  text = String(text || '').trim();
  if (!text) throw new Error('내용이 비어 있습니다.');
  const now = nowIso_();
  // 직접 쓴 메모는 쓰는 순간 본 것으로 둔다. Claude 답글이 달리면 다시 미확인이 된다.
  withLock_(() => append_('memos', [{ id: newId_('m'), day: today_(), text, status: analyze === false ? 'note' : 'new', folder: folder || '',
    category: CATEGORIES.indexOf(category) >= 0 ? category : '', pinned: false, createdAt: now, source: 'app', seenAt: now,
    domain: domains_().indexOf(domain) >= 0 ? domain : '' }]));
  return uiLoad(key);
}

function uiEditMemo(key, id, text, category, day, domain) {
  checkKey_(key);
  withLock_(() => {
    const m = findBy_('memos', 'id', id);
    if (!m) throw new Error('메모를 찾을 수 없습니다.');
    const patch = { updatedAt: nowIso_() };
    if (text !== null && text !== undefined) {
      const t = String(text).trim();
      if (t && t !== String(m.text)) { patch.text = t; patch.status = 'new'; } // 내용이 바뀌면 다음 브리핑에서 다시 정리
    }
    if (category !== null && category !== undefined) patch.category = CATEGORIES.indexOf(category) >= 0 ? category : '';
    if (day && /^\d{4}-\d{2}-\d{2}$/.test(day)) patch.day = day;
    if (domain !== null && domain !== undefined) patch.domain = domains_().indexOf(domain) >= 0 ? domain : '';
    patchRow_('memos', m._row, patch);
  });
  return uiLoad(key);
}

function uiDeleteMemo(key, id) {
  checkKey_(key);
  withLock_(() => {
    deleteBy_('memos', 'id', id);
    readAll_('comments').filter(c => c.targetType === 'memo' && c.targetId === id)
      .map(c => c._row).sort((a, b) => b - a).forEach(r => sheetOf_('comments').deleteRow(r));
    invalidate_('comments');
  });
  return uiLoad(key);
}

function uiTogglePin(key, type, id) {
  checkKey_(key);
  withLock_(() => {
    const name = type === 'brief' ? 'briefs' : 'memos';
    const o = findBy_(name, type === 'brief' ? 'date' : 'id', id);
    if (o) patchRow_(name, o._row, { pinned: !o.pinned });
  });
  return uiLoad(key);
}

function uiAddComment(key, type, id, text) {
  checkKey_(key);
  text = String(text || '').trim();
  if (!text) throw new Error('댓글이 비어 있습니다.');
  withLock_(() => append_('comments', [{ id: newId_('c'), targetType: ['brief', 'item'].indexOf(type) >= 0 ? type : 'memo',
    targetId: id, by: 'user', text, createdAt: nowIso_(), seen: false }]));
  return uiLoad(key);
}

function uiDeleteComment(key, id) {
  checkKey_(key);
  withLock_(() => deleteBy_('comments', 'id', id));
  return uiLoad(key);
}

/** 확인 표시: type = "memo" | "brief", ids = 메모 id 또는 브리핑 날짜 목록, seen = false 이면 미확인으로 되돌림 */
function uiSetSeen(key, type, ids, seen) {
  checkKey_(key);
  const name = type === 'brief' ? 'briefs' : type === 'item' ? 'items' : 'memos', k = type === 'brief' ? 'date' : 'id';
  ids = (Array.isArray(ids) ? ids : [ids]).map(String);
  withLock_(() => {
    const at = seen === false ? '' : nowIso_();
    readAll_(name).filter(o => ids.indexOf(String(o[k])) >= 0)
      .forEach(o => patchRow_(name, o._row, { seenAt: at }));
  });
  return uiLoad(key);
}

/** 관리 카테고리 지정: type = "item" | "memo" */
function uiSetDomain(key, type, id, domain) {
  checkKey_(key);
  const name = type === 'item' ? 'items' : 'memos';
  withLock_(() => {
    const o = findBy_(name, 'id', id);
    if (o) patchRow_(name, o._row, { domain: domains_().indexOf(domain) >= 0 ? domain : '' });
  });
  return uiLoad(key);
}

/** 관리 화면용 필드 변경: type = "item" | "memo", patch 에서 domain/folder/mpin/hidden 만 받는다 */
function cleanPatch_(name, patch) {
  const p = {}, allow = EDITABLE[name] || [], bools = SCHEMA[name].bools;
  Object.keys(patch || {}).forEach(f => {
    if (allow.indexOf(f) < 0) return;
    let v = patch[f];
    if (bools.indexOf(f) >= 0 || f === 'done') v = !!v;
    else if (f === 'domain') v = domains_().indexOf(v) >= 0 ? v : '';
    else if (f === 'memoIds') v = (Array.isArray(v) ? v : String(v || '').split(',')).map(x => String(x).trim()).filter(Boolean)
      .filter((x, i, a) => a.indexOf(x) === i).join(',');
    else if (f === 'sort') v = (v === '' || v === null || v === undefined || isNaN(Number(v))) ? '' : Number(v);
    else v = v === null || v === undefined ? '' : String(v);
    p[f] = v;
  });
  return p;
}
const TYPE_SHEET = { item: 'items', memo: 'memos', schedule: 'schedules', todo: 'todos' };

function uiSetEntry(key, type, id, patch) {
  checkKey_(key);
  const name = TYPE_SHEET[type] || 'memos', p = cleanPatch_(name, patch);
  withLock_(() => { const o = findBy_(name, 'id', id); if (o && Object.keys(p).length) patchRow_(name, o._row, p); });
  return uiLoad(key);
}

/** 여러 개를 한 번에: list = [{type, id, patch}] (답변 붙이기·떼기, 순서 저장 등) */
function uiSetEntries(key, list) {
  checkKey_(key);
  withLock_(() => (list || []).slice(0, 300).forEach(x => {
    const name = TYPE_SHEET[x.type] || 'memos', p = cleanPatch_(name, x.patch);
    const o = findBy_(name, 'id', x.id); if (o && Object.keys(p).length) patchRow_(name, o._row, p);
  }));
  return uiLoad(key);
}

function uiDeleteItem(key, id) {
  checkKey_(key);
  withLock_(() => {
    deleteBy_('items', 'id', id);
    readAll_('comments').filter(c => c.targetType === 'item' && c.targetId === id)
      .map(c => c._row).sort((a, b) => b - a).forEach(r => sheetOf_('comments').deleteRow(r));
    invalidate_('comments');
  });
  return uiLoad(key);
}

/** 탭 순서·숨김 저장 (폰·PC 공통) */
function uiSaveTabs(key, cfg) {
  checkKey_(key);
  const order = (cfg && cfg.order || []).filter(x => TABS.indexOf(x) >= 0);
  TABS.forEach(x => { if (order.indexOf(x) < 0) order.push(x); });
  const hidden = (cfg && cfg.hidden || []).filter(x => TABS.indexOf(x) >= 0);
  if (hidden.length >= TABS.length) throw new Error('탭이 하나는 보여야 해요.');
  withLock_(() => setSetting_('탭설정', JSON.stringify({ order, hidden })));
  return uiLoad(key);
}

/** 카테고리 안의 폴더 목록 저장 */
function uiSaveFolders(key, domain, list) {
  checkKey_(key);
  const clean = (list || []).map(x => String(x).trim()).filter(Boolean).filter((x, i, a) => a.indexOf(x) === i).slice(0, 30);
  withLock_(() => { const f = folders_(); f[domain] = clean; setSetting_('관리폴더', JSON.stringify(f)); });
  return uiLoad(key);
}

function uiRenameFolder(key, domain, from, to) {
  checkKey_(key);
  to = String(to || '').trim();
  if (!to) throw new Error('폴더 이름을 적어 주세요.');
  withLock_(() => {
    const f = folders_(); f[domain] = (f[domain] || []).map(x => x === from ? to : x); setSetting_('관리폴더', JSON.stringify(f));
    ['memos', 'items'].forEach(n => readAll_(n).filter(o => o.domain === domain && o.folder === from)
      .forEach(o => patchRow_(n, o._row, { folder: to })));
  });
  return uiLoad(key);
}

/** 폴더 삭제: 안의 기록은 지우지 않고 "폴더 없음"으로 */
function uiDeleteFolder(key, domain, name) {
  checkKey_(key);
  withLock_(() => {
    const f = folders_(); f[domain] = (f[domain] || []).filter(x => x !== name); setSetting_('관리폴더', JSON.stringify(f));
    ['memos', 'items'].forEach(n => readAll_(n).filter(o => o.domain === domain && o.folder === name)
      .forEach(o => patchRow_(n, o._row, { folder: '' })));
  });
  return uiLoad(key);
}

/** 관리 카테고리 목록 저장 (순서대로) */
function uiSaveDomains(key, list) {
  checkKey_(key);
  const clean = (list || []).map(x => String(x).replace(/,/g, ' ').trim()).filter(Boolean)
    .filter((x, i, a) => a.indexOf(x) === i).slice(0, 20);
  if (!clean.length) throw new Error('카테고리가 하나 이상 있어야 해요.');
  withLock_(() => setSetting_('관리카테고리', clean.join(',')));
  return uiLoad(key);
}

function uiSaveSchedule(key, s) {
  checkKey_(key);
  if (!s || !s.title || !/^\d{4}-\d{2}-\d{2}$/.test(s.date)) throw new Error('제목과 날짜를 확인해 주세요.');
  withLock_(() => {
    if (s.id) {
      const o = findBy_('schedules', 'id', s.id);
      if (o) { patchRow_('schedules', o._row, { title: s.title, date: s.date, time: s.time || '' }); return; }
    }
    append_('schedules', [{ id: newId_('s'), title: s.title, date: s.date, time: s.time || '', done: false,
      source: 'user', memoId: '', note: '', createdAt: nowIso_() }]);
  });
  return uiLoad(key);
}

function uiSetDone(key, type, id, done) {
  checkKey_(key);
  withLock_(() => {
    const name = type === 'todo' ? 'todos' : 'schedules';
    const o = findBy_(name, 'id', id);
    if (!o) return;
    const patch = { done: !!done };
    if (name === 'todos') patch.doneAt = done ? nowIso_() : '';
    patchRow_(name, o._row, patch);
  });
  return uiLoad(key);
}

function uiDelete(key, type, id) {
  checkKey_(key);
  withLock_(() => deleteBy_(type === 'todo' ? 'todos' : 'schedules', 'id', id));
  return uiLoad(key);
}

function uiAddTodo(key, text) {
  checkKey_(key);
  text = String(text || '').trim();
  if (!text) throw new Error('할 일이 비어 있습니다.');
  withLock_(() => append_('todos', [{ id: newId_('t'), text, done: false, source: 'user', memoId: '',
    createdAt: nowIso_(), doneAt: '' }]));
  return uiLoad(key);
}

/** 지표 목록 저장: 순서는 배열 순서. 기존 시세 값은 심볼 기준으로 유지. */
function uiSaveWatchlist(key, items, prompt) {
  checkKey_(key);
  withLock_(() => {
    const old = {};
    readAll_('watchlist').forEach(w => { old[w.symbol] = w; });
    const sh = sheetOf_('watchlist');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent(); invalidate_('watchlist');
    const rows = (items || []).filter(it => it && it.symbol).map((it, i) => {
      const o = old[it.symbol] || {};
      return { order: i + 1, label: it.label || it.symbol, symbol: String(it.symbol).trim(), prompt: it.prompt || '',
        enabled: it.enabled === false ? false : true,
        price: o.price || '', d1: o.d1 || '', d5: o.d5 || '', comment: o.comment || '', updatedAt: o.updatedAt || '' };
    });
    append_('watchlist', rows);
    setSetting_('지표공통프롬프트', prompt || '');
    setSetting_('지표수정시각', nowIso_());
  });
  return uiLoad(key);
}

/* ───────────── 지표: 종목 검색·실시간 시세 ───────────── */

/** 네이버 증권 자동완성으로 종목을 찾아 야후 심볼로 바꿔 준다. (예: 한미 → 한미반도체 042700.KS) */
function uiSearch(key, q) {
  checkKey_(key);
  q = String(q || '').trim();
  if (!q) return [];
  const url = 'https://ac.stock.naver.com/ac?target=stock,index&q=' + encodeURIComponent(q);
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (res.getResponseCode() !== 200) return [];
  const items = (JSON.parse(res.getContentText()).items || []);
  const US = ['NASDAQ', 'NYSE', 'AMEX'];
  return items.map(it => {
    let symbol = '';
    if (it.typeCode === 'KOSPI') symbol = it.code + '.KS';
    else if (it.typeCode === 'KOSDAQ') symbol = it.code + '.KQ';
    else if (US.indexOf(it.typeCode) >= 0 || it.nationCode === 'USA') symbol = it.code;
    return symbol ? { label: it.name, symbol, market: it.typeName || it.typeCode } : null;
  }).filter(Boolean).slice(0, 12);
}

/** 야후 파이낸스에서 현재가·전일비·5일 등락을 가져온다. */
function uiQuotes(key, symbols) {
  checkKey_(key);
  symbols = (symbols || []).filter(Boolean).slice(0, 40);
  if (!symbols.length) return {};
  const reqs = symbols.map(sym => ({
    url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(sym) + '?range=7d&interval=1d',
    muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124 Safari/537.36' },
  }));
  const out = {};
  UrlFetchApp.fetchAll(reqs).forEach((res, i) => {
    try {
      const r = JSON.parse(res.getContentText()).chart.result[0], m = r.meta;
      const closes = (r.indicators.quote[0].close || []).filter(x => x !== null);
      const p = m.regularMarketPrice, pv = m.chartPreviousClose || m.previousClose || closes[closes.length - 2];
      out[symbols[i]] = { price: p, d1: pv ? (p - pv) / pv * 100 : null, d5: closes[0] ? (p - closes[0]) / closes[0] * 100 : null,
        currency: m.currency || '' };
    } catch (e) { out[symbols[i]] = null; }
  });
  return out;
}

/** 편집기에서 한 번 실행해 "외부 사이트 접속" 권한을 승인하는 용도 (종목 검색·시세에 필요) */
function authorize() {
  const code = UrlFetchApp.fetch('https://ac.stock.naver.com/ac?target=stock&q=' + encodeURIComponent('한미'),
    { muteHttpExceptions: true }).getResponseCode();
  Logger.log('권한 승인 완료 (응답 ' + code + '). 이제 앱에서 종목 검색·시세를 쓸 수 있습니다.');
}
