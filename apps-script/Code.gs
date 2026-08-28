/**
 * IDS Lab 근태 관리 백엔드 (Google Apps Script)
 *
 * 랩장(관리자)만 기록을 수정할 수 있고, 나머지는 조회만 가능합니다.
 * 관리자 확인은 반드시 이 서버 쪽에서 이뤄집니다. 페이지의 자바스크립트는
 * 누구나 볼 수 있으므로, 열쇠는 코드가 아니라 '설정' 시트에만 둡니다.
 *
 * 처음 한 번만: 편집기에서 setupSheets() 실행 → 배포 > 새 배포 > 웹 앱
 * (실행: 나 / 액세스: 모든 사용자) → 발급된 /exec 주소를 attendance.html 에 입력.
 */

var SHEET_CONFIG  = '설정';
var SHEET_MEMBERS = '구성원';
var SHEET_LOG     = '근태기록';

/** 기록하는 상태 — '정상 출근'은 기록하지 않습니다(예외만 기록). */
var STATUSES = ['지각', '재택', '휴가', '결근'];

var DEFAULTS = {
  '벌금시작일': '2026-09-14',
  '기본벌금': 10000,
  '인상액': 5000,
  '회당상한': 30000,
  '재택_월한도': 3,
  '재택_주한도': 1,
  '관리자키': ''
};

/* ------------------------------------------------------------------ */
/* 유틸                                                                */
/* ------------------------------------------------------------------ */

/**
 * 시간대를 서울로 고정한다.
 * Session.getScriptTimeZone() 은 '스프레드시트'가 아니라 'Apps Script 프로젝트'의
 * 시간대를 돌려주는데, 이 값이 미국으로 잡혀 있으면 한국 오전에 날짜가 하루 밀린다.
 * 프로젝트 설정에 의존하지 않도록 여기서 못박는다.
 */
var TIMEZONE = 'Asia/Seoul';
function tz() { return TIMEZONE; }
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function fmtDate(d) { return Utilities.formatDate(d, tz(), 'yyyy-MM-dd'); }
function fmtStamp(d) { return Utilities.formatDate(d, tz(), 'yyyy-MM-dd HH:mm'); }
function monthKey(dateStr) { return String(dateStr).slice(0, 7); }

function isDateStr(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s)); }

/**
 * 셀 값을 사람이 읽을 문자열로 바꾼다.
 * 시트에서 '시각만' 담긴 칸은 1899-12-30 기준의 Date 로 넘어오는데,
 * 그대로 String() 하면 'Sat Dec 30 1899 …' 같은 값이 화면에 튀어나온다.
 */
function cellText(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return (v.getFullYear() < 1901)
      ? Utilities.formatDate(v, tz(), 'HH:mm')      // 시각만 담긴 칸
      : fmtDate(v);
  }
  return String(v).trim();
}

/** 'yyyy-MM-dd' 가 속한 주의 월요일 */
function weekKey(dateStr) {
  var p = String(dateStr).split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  var dow = d.getDay();
  d.setDate(d.getDate() - ((dow === 0) ? 6 : dow - 1));
  return Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* 시트 접근                                                            */
/* ------------------------------------------------------------------ */

function getConfig() {
  var sh = ss().getSheetByName(SHEET_CONFIG);
  var cfg = {};
  for (var k in DEFAULTS) cfg[k] = DEFAULTS[k];
  if (sh && sh.getLastRow() > 1) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      var key = String(rows[i][0]).trim();
      if (!key) continue;
      var val = rows[i][1];
      if (val instanceof Date) val = fmtDate(val);
      cfg[key] = val;
    }
  }
  ['기본벌금', '인상액', '회당상한', '재택_월한도', '재택_주한도'].forEach(function (k) {
    cfg[k] = Number(cfg[k]);
  });
  cfg['관리자키'] = String(cfg['관리자키']).trim();
  return cfg;
}

function getMembers() {
  var sh = ss().getSheetByName(SHEET_MEMBERS);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var name = String(rows[i][0]).trim();
    if (!name) continue;
    var active = String(rows[i][1]).trim().toUpperCase();
    if (active === 'N' || active === 'FALSE') continue;
    out.push({ name: name, note: String(rows[i][2] || '').trim() });
  }
  return out;
}

/** 근태기록: 날짜 | 이름 | 상태 | 비고 | 수정시각 */
function getRecords() {
  var sh = ss().getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var date = rows[i][0];
    if (!date) continue;
    date = (date instanceof Date) ? fmtDate(date) : String(date).trim();
    var name = cellText(rows[i][1]);
    var status = cellText(rows[i][2]);
    if (!name || !status) continue;
    if (STATUSES.indexOf(status) < 0) continue;   // '출근' 등 옛 구조의 값은 무시
    out.push({
      date: date, name: name, status: status,
      note: cellText(rows[i][3]),
      updated: cellText(rows[i][4])
    });
  }
  return out;
}

/** 기록 전체를 시트에 다시 씀 (날짜·이름 순 정렬) */
function writeRecords(records) {
  var sh = ss().getSheetByName(SHEET_LOG);
  records.sort(function (a, b) {
    return a.date < b.date ? -1 : a.date > b.date ? 1 : a.name.localeCompare(b.name, 'ko');
  });
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 5).clearContent();
  if (!records.length) return;
  var values = records.map(function (r) {
    return [r.date, r.name, r.status, r.note, r.updated];
  });
  sh.getRange(2, 1, values.length, 5).setValues(values);
}

/* ------------------------------------------------------------------ */
/* 벌금 계산 — 회차 누진, 월 합산                                       */
/* ------------------------------------------------------------------ */

/**
 * n번째 지각 = min(기본벌금 + 인상액 * (n-1), 회당상한)
 * 월 납부액 = 그 달 지각들의 합
 */
function calcFine(lateDates, cfg) {
  var dates = lateDates.slice().sort();
  var each = [], total = 0;
  for (var i = 0; i < dates.length; i++) {
    var amount = Math.min(cfg['기본벌금'] + cfg['인상액'] * i, cfg['회당상한']);
    total += amount;
    each.push({ nth: i + 1, date: dates[i], amount: amount });
  }
  return { count: dates.length, total: total, each: each };
}

/** 벌금 부과 대상인 지각인지 */
function isFineable(rec, cfg) {
  if (rec.status !== '지각') return false;
  if (rec.note.indexOf('면제') >= 0) return false;
  if (rec.date < String(cfg['벌금시작일'])) return false;
  return true;
}

/** 월 단위 전체 요약 */
function buildSummary(month, cfg, members, records) {
  var byName = {};
  members.forEach(function (m) {
    byName[m.name] = { name: m.name, note: m.note, lateDates: [], late: 0, remote: 0, leave: 0, absent: 0 };
  });
  records.forEach(function (r) {
    if (monthKey(r.date) !== month) return;
    if (!byName[r.name]) {
      byName[r.name] = { name: r.name, note: '(퇴소)', lateDates: [], late: 0, remote: 0, leave: 0, absent: 0 };
    }
    var b = byName[r.name];
    if (r.status === '지각') { b.late++; if (isFineable(r, cfg)) b.lateDates.push(r.date); }
    else if (r.status === '재택') b.remote++;
    else if (r.status === '휴가') b.leave++;
    else if (r.status === '결근') b.absent++;
  });

  var out = [];
  Object.keys(byName).forEach(function (name) {
    var b = byName[name];
    var f = calcFine(b.lateDates, cfg);
    out.push({
      name: name, note: b.note,
      late: b.late, remote: b.remote, leave: b.leave, absent: b.absent,
      fine: f.total,
      nextFine: Math.min(cfg['기본벌금'] + cfg['인상액'] * f.count, cfg['회당상한'])
    });
  });
  out.sort(function (a, b) { return b.fine - a.fine || a.name.localeCompare(b.name, 'ko'); });
  return out;
}

/* ------------------------------------------------------------------ */
/* 관리자 확인                                                          */
/* ------------------------------------------------------------------ */

function isAdmin(key, cfg) {
  var real = String(cfg['관리자키'] || '').trim();
  if (!real) return false;                       // 열쇠 미설정이면 아무도 관리자가 아님
  return String(key || '').trim() === real;
}

/* ------------------------------------------------------------------ */
/* API                                                                 */
/* ------------------------------------------------------------------ */

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    return json(bootstrap(params.month, params.key));
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = {};
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
    if (body.action === 'save') return json(saveDay(body.key, body.date, body.entries));
    return json({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function bootstrap(month, key) {
  var cfg = getConfig();
  var now = new Date();
  var today = fmtDate(now);
  month = isDateStr(month + '-01') ? month : monthKey(today);

  var members = getMembers();
  var records = getRecords();
  var admin = isAdmin(key, cfg);

  var monthRecords = records.filter(function (r) { return monthKey(r.date) === month; })
    .map(function (r) { return { date: r.date, name: r.name, status: r.status, note: r.note }; });

  return {
    ok: true,
    admin: admin,
    today: today,
    month: month,
    statuses: STATUSES,
    config: {
      fineStartDate: String(cfg['벌금시작일']),
      base: cfg['기본벌금'], step: cfg['인상액'], cap: cfg['회당상한'],
      remoteMonthly: cfg['재택_월한도'], remoteWeekly: cfg['재택_주한도']
    },
    members: members,
    records: monthRecords,
    todayRecords: records.filter(function (r) { return r.date === today; })
      .map(function (r) { return { name: r.name, status: r.status, note: r.note }; }),
    summary: buildSummary(month, cfg, members, records)
  };
}

/**
 * 하루치 예외를 통째로 저장한다.
 * entries: [{ name, status, note }] — status 가 빈 값이면 그 사람의 그날 기록을 지운다.
 * 넘기지 않은 사람의 기존 기록은 건드리지 않는다.
 */
function saveDay(key, date, entries) {
  var cfg = getConfig();
  if (!isAdmin(key, cfg)) {
    return { ok: false, error: '수정 권한이 없습니다. 관리자 링크로 접속했는지 확인해 주세요.' };
  }
  if (!isDateStr(date)) return { ok: false, error: '날짜 형식이 올바르지 않습니다.' };
  if (!entries || !entries.length) return { ok: false, error: '저장할 내용이 없습니다.' };

  var members = getMembers();
  var valid = {};
  members.forEach(function (m) { valid[m.name] = true; });

  var touched = {}, keep = [];
  var stamp = fmtStamp(new Date());
  var warnings = [];

  entries.forEach(function (en) {
    var name = String(en.name || '').trim();
    if (!name || !valid[name]) { warnings.push(name + ' — 구성원 목록에 없어 건너뜀'); return; }
    touched[name] = true;
    var status = String(en.status || '').trim();
    if (!status) return;                                  // 지움
    if (STATUSES.indexOf(status) < 0) { warnings.push(name + ' — 알 수 없는 상태 ' + status); return; }
    keep.push({ date: date, name: name, status: status, note: String(en.note || '').trim(), updated: stamp });
  });

  var records = getRecords().filter(function (r) {
    return !(r.date === date && touched[r.name]);          // 이번에 지정한 사람의 그날 기록만 교체
  });
  writeRecords(records.concat(keep));

  var out = bootstrap(monthKey(date), key);
  out.saved = keep.length;
  out.cleared = Object.keys(touched).length - keep.length;
  if (warnings.length) out.warnings = warnings;
  return out;
}

/* ------------------------------------------------------------------ */
/* 옛 시트 구조 이전 (한 번만 실행)                                      */
/* ------------------------------------------------------------------ */

/**
 * 예전 '출근 체크' 방식으로 만들어진 근태기록 시트를 새 구조로 옮긴다.
 *   옛: 날짜 | 이름 | 구분 | 체크시각 | 지각 | 비고
 *   새: 날짜 | 이름 | 상태 | 비고     | 수정시각
 *
 * 정상 출근(지각 아님)은 기록하지 않는 방식이므로 그런 행은 사라진다.
 * 지각 · 재택 · 휴가 · 결근만 남는다.
 */
function migrateOldSheet() {
  var sh = ss().getSheetByName(SHEET_LOG);
  if (!sh) { alertSafe('근태기록 시트를 찾을 수 없습니다.'); return; }

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  var header = lastRow ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(cellText) : [];

  if (header.indexOf('구분') < 0 && header.indexOf('체크시각') < 0) {
    alertSafe('이미 새 구조입니다. 옮길 것이 없습니다.');
    return;
  }

  var moved = [], dropped = 0;
  var stamp = fmtStamp(new Date());
  if (lastRow > 1) {
    var rows = sh.getRange(2, 1, lastRow - 1, Math.max(lastCol, 6)).getValues();
    for (var i = 0; i < rows.length; i++) {
      var date = rows[i][0];
      date = (date instanceof Date) ? fmtDate(date) : cellText(date);
      var name = cellText(rows[i][1]);
      if (!isDateStr(date) || !name) continue;

      var kind = cellText(rows[i][2]);             // 구분: 출근 / 재택 / 휴가
      var late = cellText(rows[i][4]).toUpperCase() === 'Y';
      var note = cellText(rows[i][5]);

      var status = '';
      if (kind === '출근') { if (late) status = '지각'; else { dropped++; continue; } }
      else if (STATUSES.indexOf(kind) >= 0) status = kind;
      else { dropped++; continue; }

      moved.push({ date: date, name: name, status: status, note: note, updated: stamp });
    }
  }

  // 시트를 새 머리글로 갈아엎고 옮긴 행을 쓴다
  sh.clear();
  sh.getRange(1, 1, 1, 5)
    .setValues([['날짜', '이름', '상태', '비고', '수정시각']]).setFontWeight('bold');
  sh.getRange('A:A').setNumberFormat('@');
  sh.getRange('D:E').setNumberFormat('@');
  sh.setFrozenRows(1);
  sh.setColumnWidth(4, 200);
  sh.setColumnWidth(5, 140);
  writeRecords(moved);

  alertSafe('이전 완료\n\n' +
    '옮긴 기록: ' + moved.length + '건 (지각 · 재택 · 휴가 · 결근)\n' +
    '정리한 행: ' + dropped + '건 (정상 출근 등, 새 방식에서는 기록하지 않습니다)\n\n' +
    '이제 배포 > 배포 관리 > 편집 > 버전: 새 버전 > 배포 를 해주세요.');
}

function alertSafe(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/* ------------------------------------------------------------------ */
/* 최초 1회 실행                                                        */
/* ------------------------------------------------------------------ */

function setupSheets() {
  var book = ss();
  book.setSpreadsheetTimeZone('Asia/Seoul');

  var cfg = book.getSheetByName(SHEET_CONFIG) || book.insertSheet(SHEET_CONFIG);
  if (cfg.getLastRow() === 0) {
    cfg.getRange(1, 1, 1, 2).setValues([['항목', '값']]).setFontWeight('bold');
    var rows = [];
    for (var k in DEFAULTS) {
      rows.push([k, k === '관리자키' ? Utilities.getUuid().replace(/-/g, '').slice(0, 12) : DEFAULTS[k]]);
    }
    cfg.getRange(2, 1, rows.length, 2).setValues(rows);
    cfg.getRange(2, 2).setNumberFormat('@');     // 벌금시작일을 문자열로
    cfg.setColumnWidth(1, 140);
    cfg.setColumnWidth(2, 220);
  }

  var mem = book.getSheetByName(SHEET_MEMBERS) || book.insertSheet(SHEET_MEMBERS);
  if (mem.getLastRow() === 0) {
    mem.getRange(1, 1, 1, 3).setValues([['이름', '활성(Y/N)', '비고']]).setFontWeight('bold');
    mem.getRange(2, 1, 1, 3).setValues([['김종엽', 'Y', '랩장']]);
    mem.setColumnWidth(1, 110);
    mem.setFrozenRows(1);
  }

  var log = book.getSheetByName(SHEET_LOG) || book.insertSheet(SHEET_LOG);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, 5)
      .setValues([['날짜', '이름', '상태', '비고', '수정시각']]).setFontWeight('bold');
    log.getRange('A:A').setNumberFormat('@');
    log.setFrozenRows(1);
    log.setColumnWidth(4, 200);
    log.setColumnWidth(5, 140);
  }

  var sheet1 = book.getSheetByName('시트1') || book.getSheetByName('Sheet1');
  if (sheet1 && book.getSheets().length > 3) book.deleteSheet(sheet1);

  var key = getConfig()['관리자키'];
  alertSafe(
    '시트 준비 완료\n\n' +
    '1) 구성원 시트에 랩원 이름을 채워 주세요.\n' +
    '2) 관리자 열쇠: ' + key + '\n' +
    '   관리자 링크 → attendance.html?admin=' + key + '\n\n' +
    '이 열쇠는 랩원에게 공유하지 마세요. 설정 시트에서 언제든 바꿀 수 있습니다.');
}
