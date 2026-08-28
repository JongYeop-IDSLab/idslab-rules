const fs = require('fs');
const vm = require('vm');
const src = fs.readFileSync(process.env.HOME + '/mnt/repos/idslab-rules/apps-script/Code.gs', 'utf8');

// Apps Script 전역 최소 스텁
const sandbox = {
  Session: { getScriptTimeZone: () => 'Asia/Seoul' },
  SpreadsheetApp: {},
  ContentService: {},
  LockService: {},
  Utilities: {
    formatDate: (d, tzs, fmt) => {
      const p = n => String(n).padStart(2, '0');
      if (fmt === 'yyyy-MM-dd') return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
      if (fmt === 'HH:mm') return `${p(d.getHours())}:${p(d.getMinutes())}`;
      throw new Error('fmt? ' + fmt);
    }
  },
  console
};
vm.createContext(sandbox);
vm.runInContext(src, sandbox);

const cfg = {
  '코어타임시작': '10:00', '벌금시작일': '2026-09-14',
  '기본벌금': 10000, '인상액': 5000, '회당상한': 30000,
  '재택_월한도': 3, '재택_주한도': 1
};

let pass = 0, fail = 0;
function eq(label, got, want) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  ok ? pass++ : fail++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${label}  →  ${JSON.stringify(got)}${ok ? '' : '   (기대: ' + JSON.stringify(want) + ')'}`);
}

// 1) 회차별 누진 + 월 합산
const expected = [
  [1, 10000], [2, 25000], [3, 45000], [4, 70000], [5, 100000], [6, 130000], [7, 160000]
];
console.log('— 지각 n회 시 월 합산 —');
for (const [n, total] of expected) {
  const dates = Array.from({ length: n }, (_, i) => `2026-09-${String(14 + i).padStart(2, '0')}`);
  eq(`${n}회 지각`, sandbox.calcFine(dates, cfg).total, total);
}

// 2) 회당 상한 3만원
console.log('\n— 회당 금액 상한 —');
const each = sandbox.calcFine(
  Array.from({ length: 7 }, (_, i) => `2026-09-${String(14 + i).padStart(2, '0')}`), cfg).each;
eq('회차별 금액', each.map(e => e.amount), [10000, 15000, 20000, 25000, 30000, 30000, 30000]);

// 3) 정렬 여부와 무관하게 같은 결과
console.log('\n— 입력 순서 무관 —');
eq('역순 입력', sandbox.calcFine(['2026-09-30', '2026-09-14', '2026-09-21'], cfg).total, 45000);

// 4) 벌금 대상 판정
console.log('\n— 벌금 대상 판정 —');
const rec = (o) => Object.assign({ date: '2026-09-20', name: '김종엽', mode: '출근', time: '10:12', late: true, note: '' }, o);
eq('9/14 이후 지각', sandbox.isFineable(rec({}), cfg), true);
eq('9/14 이전 지각(9/1)', sandbox.isFineable(rec({ date: '2026-09-01' }), cfg), false);
eq('시행 첫날(9/14)', sandbox.isFineable(rec({ date: '2026-09-14' }), cfg), true);
eq('8월 지각', sandbox.isFineable(rec({ date: '2026-08-31' }), cfg), false);
eq('재택은 제외', sandbox.isFineable(rec({ mode: '재택' }), cfg), false);
eq('휴가는 제외', sandbox.isFineable(rec({ mode: '휴가' }), cfg), false);
eq('비고 면제', sandbox.isFineable(rec({ note: '수업으로 면제' }), cfg), false);
eq('정시 출근', sandbox.isFineable(rec({ late: false }), cfg), false);

// 5) 월 요약 — 9월 이전 지각은 집계 제외, 월 경계 분리
console.log('\n— 월 요약 —');
const members = [{ name: '김종엽', note: '랩장' }, { name: '홍길동', note: '' }];
const records = [
  rec({ date: '2026-09-02' }),                        // 시행 전 지각 → 벌금 0
  rec({ date: '2026-09-15' }),                        // 1회차
  rec({ date: '2026-09-16' }),                        // 2회차
  rec({ date: '2026-09-17', mode: '재택', late: false }),
  rec({ date: '2026-10-05' }),                        // 다음 달
  rec({ name: '홍길동', date: '2026-09-15', late: false })
];
const sep = sandbox.buildSummary('2026-09', cfg, members, records);
const me = sep.find(s => s.name === '김종엽');
eq('9월 지각 횟수', me.lateCount, 2);
eq('9월 누적 벌금', me.fine, 25000);
eq('9월 재택 사용', me.remote, 1);
eq('9월 기록일 수', me.days, 4);
eq('다음 지각 시 금액', me.nextFine, 20000);
const oct = sandbox.buildSummary('2026-10', cfg, members, records).find(s => s.name === '김종엽');
eq('10월은 초기화', [oct.lateCount, oct.fine], [1, 10000]);
eq('지각 없는 사람', sep.find(s => s.name === '홍길동').fine, 0);

// 6) 주 경계 (월~일)
console.log('\n— 주 경계 (월요일 시작) —');
eq('2026-09-16(수) 주', sandbox.weekKey('2026-09-16'), '2026-09-14');
eq('2026-09-20(일) 주', sandbox.weekKey('2026-09-20'), '2026-09-14');
eq('2026-09-21(월) 주', sandbox.weekKey('2026-09-21'), '2026-09-21');
eq('토요일 판정', sandbox.isWeekend('2026-09-19'), true);
eq('평일 판정', sandbox.isWeekend('2026-09-18'), false);

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
