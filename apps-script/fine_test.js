/**
 * 벌금 계산 · 권한 판정 로직 테스트
 *   node apps-script/fine_test.js
 */
const fs = require('fs');
const vm = require('vm');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8');

// Apps Script 전역 최소 스텁
const sandbox = {
  Session: { getScriptTimeZone: () => 'Asia/Seoul' },
  SpreadsheetApp: {}, ContentService: {}, LockService: {},
  Utilities: {
    getUuid: () => 'test-uuid',
    formatDate: (d, tzs, fmt) => {
      const p = n => String(n).padStart(2, '0');
      if (fmt === 'yyyy-MM-dd') return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
      if (fmt === 'yyyy-MM-dd HH:mm') return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
      throw new Error('fmt? ' + fmt);
    }
  },
  console
};
vm.createContext(sandbox);
vm.runInContext(src, sandbox);

const cfg = {
  '벌금시작일': '2026-09-14',
  '기본벌금': 10000, '인상액': 5000, '회당상한': 30000,
  '재택_월한도': 3, '재택_주한도': 1, '관리자키': 'abc123def456'
};

let pass = 0, fail = 0;
function eq(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}  →  ${JSON.stringify(got)}${ok ? '' : '   (기대: ' + JSON.stringify(want) + ')'}`);
}
const day = n => `2026-09-${String(n).padStart(2, '0')}`;

console.log('— 지각 n회 시 월 합산 —');
[[1, 10000], [2, 25000], [3, 45000], [4, 70000], [5, 100000], [6, 130000], [7, 160000]]
  .forEach(([n, total]) => eq(`${n}회 지각`,
    sandbox.calcFine(Array.from({ length: n }, (_, i) => day(14 + i)), cfg).total, total));

console.log('\n— 회당 금액 상한 · 입력 순서 —');
eq('회차별 금액',
  sandbox.calcFine(Array.from({ length: 7 }, (_, i) => day(14 + i)), cfg).each.map(e => e.amount),
  [10000, 15000, 20000, 25000, 30000, 30000, 30000]);
eq('역순 입력도 동일', sandbox.calcFine([day(30), day(14), day(21)], cfg).total, 45000);

console.log('\n— 벌금 대상 판정 —');
const rec = o => Object.assign({ date: day(20), name: '김종엽', status: '지각', note: '' }, o);
eq('시행 이후 지각', sandbox.isFineable(rec({}), cfg), true);
eq('시행 첫날(9/14)', sandbox.isFineable(rec({ date: day(14) }), cfg), true);
eq('시행 전 지각(9/1)', sandbox.isFineable(rec({ date: day(1) }), cfg), false);
eq('8월 지각', sandbox.isFineable(rec({ date: '2026-08-31' }), cfg), false);
eq('재택은 제외', sandbox.isFineable(rec({ status: '재택' }), cfg), false);
eq('휴가는 제외', sandbox.isFineable(rec({ status: '휴가' }), cfg), false);
eq('결근은 벌금 없음', sandbox.isFineable(rec({ status: '결근' }), cfg), false);
eq('비고 면제', sandbox.isFineable(rec({ note: '수업으로 면제' }), cfg), false);

console.log('\n— 월 요약 —');
const members = [{ name: '김종엽', note: '랩장' }, { name: '홍길동', note: '' }];
const records = [
  rec({ date: day(2) }),                          // 시행 전 → 벌금 0, 지각 횟수엔 포함
  rec({ date: day(15) }),                         // 1회차
  rec({ date: day(16) }),                         // 2회차
  rec({ date: day(17), status: '재택' }),
  rec({ date: day(18), status: '휴가' }),
  rec({ date: day(21), status: '결근' }),
  rec({ date: '2026-10-05' }),                    // 다음 달
  rec({ name: '홍길동', date: day(15), status: '재택' })
];
const sep = sandbox.buildSummary('2026-09', cfg, members, records);
const me = sep.find(s => s.name === '김종엽');
eq('9월 지각 횟수(기록 기준)', me.late, 3);
eq('9월 누적 벌금(시행 후 2건)', me.fine, 25000);
eq('9월 재택', me.remote, 1);
eq('9월 휴가', me.leave, 1);
eq('9월 결근', me.absent, 1);
eq('다음 지각 시 금액', me.nextFine, 20000);
eq('10월은 초기화', (() => { const o = sandbox.buildSummary('2026-10', cfg, members, records).find(s => s.name === '김종엽'); return [o.late, o.fine]; })(), [1, 10000]);
eq('지각 없는 사람', sep.find(s => s.name === '홍길동').fine, 0);
eq('기록 없는 달', sandbox.buildSummary('2026-11', cfg, members, records).every(s => s.fine === 0 && s.late === 0), true);

console.log('\n— 관리자 판정 —');
eq('올바른 열쇠', sandbox.isAdmin('abc123def456', cfg), true);
eq('앞뒤 공백 허용', sandbox.isAdmin('  abc123def456 ', cfg), true);
eq('틀린 열쇠', sandbox.isAdmin('wrong', cfg), false);
eq('빈 열쇠', sandbox.isAdmin('', cfg), false);
eq('열쇠 누락', sandbox.isAdmin(undefined, cfg), false);
eq('설정에 열쇠가 없으면 전원 거부', sandbox.isAdmin('', Object.assign({}, cfg, { '관리자키': '' })), false);
eq('설정 미설정 + 아무 값', sandbox.isAdmin('anything', Object.assign({}, cfg, { '관리자키': '' })), false);

console.log('\n— 날짜 유틸 —');
eq('주 시작(수)', sandbox.weekKey(day(16)), day(14));
eq('주 시작(일)', sandbox.weekKey(day(20)), day(14));
eq('주 시작(월)', sandbox.weekKey(day(21)), day(21));
eq('날짜 형식 검사', [sandbox.isDateStr('2026-09-14'), sandbox.isDateStr('2026-9-14'), sandbox.isDateStr('')], [true, false, false]);

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
