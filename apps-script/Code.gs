/**
 * IDS Lab 근태 관리 백엔드 (Google Apps Script)
 *
 * 처음 한 번만: 스크립트 편집기에서 setupSheets() 를 실행한 뒤,
 * 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자) 으로 배포하세요.
 * 발급된 /exec URL 을 attendance.html 의 API_URL 에 넣으면 됩니다.
 */

var SHEET_CONFIG  = '설정';
var SHEET_MEMBERS = '구성원';
var SHEET_LOG     = '근태기록';

var DEFAULTS = {
  '코어타임시작': '10:00',
  '벌금시작일': '2026-09-14',
  '기본벌금': 10000,
  '인상액': 5000,
  '회당상한': 30000,
  '재택_월한도': 3,
  '재택_주한도': 1,
  '관리자키': 'idslab'
};

/* ------------------------------------------------------------------ */
/* 유틸                                                                */
/* ------------------------------------------------------------------ */

function tz() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }
function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function fmtDate(d) { return Utilities.formatDate(d, tz(), 'yyyy-MM-dd'); }
function fmtTime(d) { return Utilities.formatDate(d, tz(), 'HH:mm'); }

/** 'HH:mm' -> 분 단위 정수 */
function toMinutes(hhmm) {
  var p = String(hhmm).split(':');
  return parseInt(p[0], 10) * 60 + parseInt(p[1], 10);
}

/** 'yyyy-MM-dd' 의 월 키 'yyyy-MM' */
function monthKey(dateStr) { return String(dateStr).slice(0, 7); }

/** 'yyyy-MM-dd' 가 속한 주의 월요일 (주 = 월~일) */
function weekKey(dateStr) {
  var p = String(dateStr).split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  var dow = d.getDay();                 // 0=일 … 6=토
  var back = (dow === 0) ? 6 : dow - 1; // 월요일까지 며칠 전인지
  d.setDate(d.getDate() - back);
  return Utilities.formatDate(d, tz(), 'yyyy-MM-dd');
}

/** 주말 여부 */
function isWeekend(dateStr) {
  var p = String(dateStr).split('-');
  var dow = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getDay();
  return dow === 0 || dow === 6;
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
      if (val instanceof Date) {
        val = (key === '코어타임시작') ? fmtTime(val) : fmtDate(val);
      }
      cfg[key] = val;
    }
  }
  ['기본벌금', '인상액', '회당상한', '재택_월한도', '재택_주한도'].forEach(function (k) {
    cfg[k] = Number(cfg[k]);
  });
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

/** 근태기록 전체를 객체 배열로 */
function getRecords() {
  var sh = ss().getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var date = rows[i][0];
    if (!date) continue;
    date = (date instanceof Date) ? fmtDate(date) : String(date).trim();
    var time = rows[i][3];
    time = (time instanceof Date) ? fmtTime(time) : String(time || '').trim();
    out.push({
      row: i + 2,
      date: date,
      name: String(rows[i][1]).trim(),
      mode: String(rows[i][2]).trim(),
      time: time,
      late: String(rows[i][4]).trim() === 'Y',
      note: String(rows[i][5] || '').trim()
    });
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* 벌금 계산 — 회차 누진, 월 합산                                       */
/* ------------------------------------------------------------------ */

/**
 * 한 사람의 한 달치 지각 기록으로 벌금을 계산한다.
 * n번째 지각 = min(기본벌금 + 인상액 * (n-1), 회당상한)
 * 월 납부액 = 그 달 지각들의 합
 *
 * @param {Array} lateDates 벌금 대상 지각 날짜 배열 (같은 달, 정렬 여부 무관)
 * @param {Object} cfg 설정
 * @return {{count:number, total:number, each:Array}}
 */
function calcFine(lateDates, cfg) {
  var dates = lateDates.slice().sort();
  var each = [];
  var total = 0;
  for (var i = 0; i < dates.length; i++) {
    var amount = Math.min(cfg['기본벌금'] + cfg['인상액'] * i, cfg['회당상한']);
    total += amount;
    each.push({ nth: i + 1, date: dates[i], amount: amount });
  }
  return { count: dates.length, total: total, each: each };
}

/** 벌금 부과 대상인 지각인지 */
function isFineable(rec, cfg) {
  if (!rec.late) return false;
  if (rec.mode !== '출근') return false;
  if (rec.note.indexOf('면제') >= 0) return false;
  if (rec.date < String(cfg['벌금시작일'])) return false;
  return true;
}

/** 월 단위 전체 요약 */
function buildSummary(month, cfg, members, records) {
  var byName = {};
  members.forEach(function (m) {
    byName[m.name] = { name: m.name, note: m.note, lateDates: [], remote: 0, days: 0 };
  });
  records.forEach(function (r) {
    if (monthKey(r.date) !== month) return;
    if (!byName[r.name]) byName[r.name] = { name: r.name, note: '', lateDates: [], remote: 0, days: 0 };
    var b = byName[r.name];
    b.days++;
    if (r.mode === '재택') b.remote++;
    if (isFineable(r, cfg)) b.lateDates.push(r.date);
  });

  var out = [];
  Object.keys(byName).forEach(function (name) {
    var b = byName[name];
    var f = calcFine(b.lateDates, cfg);
    out.push({
      name: name, note: b.note, days: b.days, remote: b.remote,
      lateCount: f.count, fine: f.total,
      nextFine: Math.min(cfg['기본벌금'] + cfg['인상액'] * f.count, cfg['회당상한'])
    });
  });
  out.sort(function (a, b) { return b.fine - a.fine || a.name.localeCompare(b.name, 'ko'); });
  return out;
}

/* ------------------------------------------------------------------ */
/* API                                                                 */
/* ------------------------------------------------------------------ */

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var action = params.action || 'bootstrap';
    if (action === 'bootstrap') return json(bootstrap(params.month));
    return json({ ok: false, error: '알 수 없는 요청입니다: ' + action });
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
    if (body.action === 'checkin') return json(checkin(body.name, body.mode));
    return json({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function bootstrap(month) {
  var cfg = getConfig();
  var now = new Date();
  var today = fmtDate(now);
  month = month || monthKey(today);
  var members = getMembers();
  var records = getRecords();

  var todays = records.filter(function (r) { return r.date === today; })
    .map(function (r) { return { name: r.name, mode: r.mode, time: r.time, late: r.late }; });

  return {
    ok: true,
    today: today,
    now: fmtTime(now),
    month: month,
    config: {
      coreStart: String(cfg['코어타임시작']),
      fineStartDate: String(cfg['벌금시작일']),
      base: cfg['기본벌금'], step: cfg['인상액'], cap: cfg['회당상한'],
      remoteMonthly: cfg['재택_월한도'], remoteWeekly: cfg['재택_주한도']
    },
    members: members.map(function (m) { return m.name; }),
    todayRecords: todays,
    summary: buildSummary(month, cfg, members, records)
  };
}

function checkin(name, mode) {
  name = String(name || '').trim();
  mode = String(mode || '출근').trim();
  if (!name) return { ok: false, error: '이름을 선택해 주세요.' };
  if (['출근', '재택', '휴가'].indexOf(mode) < 0) return { ok: false, error: '알 수 없는 구분입니다.' };

  var cfg = getConfig();
  var members = getMembers();
  if (!members.some(function (m) { return m.name === name; })) {
    return { ok: false, error: name + ' 님은 구성원 목록에 없습니다. 랩장에게 문의해 주세요.' };
  }

  var now = new Date();
  var today = fmtDate(now);
  var nowTime = fmtTime(now);
  var records = getRecords();

  var dup = records.filter(function (r) { return r.date === today && r.name === name; })[0];
  if (dup) {
    return { ok: false, error: '오늘은 이미 ' + dup.mode + '으로 기록되어 있습니다 (' + dup.time + ').' };
  }

  // 재택 한도 검사
  if (mode === '재택') {
    var mine = records.filter(function (r) { return r.name === name && r.mode === '재택'; });
    var m = mine.filter(function (r) { return monthKey(r.date) === monthKey(today); }).length;
    if (m >= cfg['재택_월한도']) {
      return { ok: false, error: '이번 달 재택근무를 이미 ' + m + '회 사용하셨습니다 (월 ' + cfg['재택_월한도'] + '회).' };
    }
    var w = mine.filter(function (r) { return weekKey(r.date) === weekKey(today); }).length;
    if (w >= cfg['재택_주한도']) {
      return { ok: false, error: '이번 주 재택근무를 이미 사용하셨습니다 (주 ' + cfg['재택_주한도'] + '회).' };
    }
  }

  // 지각 판정 — 출근이면서 평일, 코어타임 시작 이후
  var late = false;
  if (mode === '출근' && !isWeekend(today) && toMinutes(nowTime) > toMinutes(cfg['코어타임시작'])) {
    late = true;
  }

  ss().getSheetByName(SHEET_LOG)
    .appendRow([today, name, mode, nowTime, late ? 'Y' : '', '']);

  var fresh = getRecords();
  var summary = buildSummary(monthKey(today), cfg, members, fresh);
  var mineSummary = summary.filter(function (s) { return s.name === name; })[0] || null;

  var message;
  if (mode === '재택') message = '재택근무로 기록했습니다. 단톡방에도 잊지 말고 알려 주세요.';
  else if (mode === '휴가') message = '휴가로 기록했습니다.';
  else if (late) {
    if (today >= String(cfg['벌금시작일']) && mineSummary) {
      var nth = mineSummary.lateCount;
      var thisFine = Math.min(cfg['기본벌금'] + cfg['인상액'] * (nth - 1), cfg['회당상한']);
      message = '지각으로 기록했습니다. 이번 달 ' + nth + '회차 · '
        + thisFine.toLocaleString('ko-KR') + '원 · 누적 '
        + mineSummary.fine.toLocaleString('ko-KR') + '원';
    } else {
      message = '지각으로 기록했습니다. 벌금은 ' + cfg['벌금시작일'] + '부터 적용됩니다.';
    }
  } else message = '출근 완료. 좋은 하루 되세요.';

  return {
    ok: true, name: name, mode: mode, time: nowTime, late: late,
    message: message, mine: mineSummary, summary: summary
  };
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
    for (var k in DEFAULTS) rows.push([k, DEFAULTS[k]]);
    cfg.getRange(2, 1, rows.length, 2).setValues(rows);
    cfg.getRange('B2:B3').setNumberFormat('@');   // 시각·날짜를 문자열로
    cfg.setColumnWidth(1, 140);
  }

  var mem = book.getSheetByName(SHEET_MEMBERS) || book.insertSheet(SHEET_MEMBERS);
  if (mem.getLastRow() === 0) {
    mem.getRange(1, 1, 1, 3).setValues([['이름', '활성(Y/N)', '비고']]).setFontWeight('bold');
    mem.getRange(2, 1, 1, 3).setValues([['김종엽', 'Y', '랩장']]);
    mem.setColumnWidth(1, 110);
  }

  var log = book.getSheetByName(SHEET_LOG) || book.insertSheet(SHEET_LOG);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, 6)
      .setValues([['날짜', '이름', '구분', '체크시각', '지각', '비고']]).setFontWeight('bold');
    log.getRange('A:A').setNumberFormat('@');
    log.getRange('D:D').setNumberFormat('@');
    log.setFrozenRows(1);
  }

  var sheet1 = book.getSheetByName('시트1') || book.getSheetByName('Sheet1');
  if (sheet1 && book.getSheets().length > 3) book.deleteSheet(sheet1);

  try { SpreadsheetApp.getUi().alert('시트 준비 완료 — 구성원 시트에 랩원 이름을 채워 주세요.'); } catch (ignore) {}
}
