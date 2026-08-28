const fs=require('fs'), vm=require('vm'), path=require('path');
const src=fs.readFileSync(process.env.HOME+'/mnt/repos/idslab-rules/apps-script/Code.gs','utf8');

// 옛 구조 시트를 흉내낸 가짜 시트
const OLD = [
  ['날짜','이름','구분','체크시각','지각','비고'],
  ['2026-08-28','김종엽','출근', new Date(1899,11,30,16,11), '',  ''],
  ['2026-08-28','이상후','휴가', new Date(1899,11,30,16,25), '',  ''],
  ['2026-08-28','박진형','재택', new Date(1899,11,30,16,30), '',  ''],
  ['2026-08-28','공재영','휴가', new Date(1899,11,30,16,30), '',  ''],
  ['2026-08-28','김범진','출근', new Date(1899,11,30,16,31), 'Y', '수업'],
  ['2026-08-28','박원종','재택', new Date(1899,11,30,16,33), '',  ''],
];
let sheet = OLD.map(r=>r.slice());
const alerts=[];
function range(r,c,nr,nc){
  return {
    getValues(){ const o=[]; for(let i=0;i<nr;i++){ const row=[]; for(let j=0;j<nc;j++) row.push((sheet[r-1+i]||[])[c-1+j] ?? ''); o.push(row);} return o; },
    setValues(v){ v.forEach((row,i)=>{ sheet[r-1+i]=sheet[r-1+i]||[]; row.forEach((val,j)=>sheet[r-1+i][c-1+j]=val); }); return range(r,c,nr,nc); },
    setFontWeight(){return this}, setNumberFormat(){return this}, clearContent(){
      for(let i=0;i<nr;i++) for(let j=0;j<nc;j++){ if(sheet[r-1+i]) sheet[r-1+i][c-1+j]=''; } return this; },
  };
}
const LOG = {
  getLastRow:()=>sheet.length, getLastColumn:()=>6,
  getRange:(a,b,c,d)=> typeof a==='string' ? {setNumberFormat(){return this}} : range(a,b,c,d),
  clear(){ sheet=[]; return this; },
  setFrozenRows(){return this}, setColumnWidth(){return this},
};
const sandbox={
  Session:{getScriptTimeZone:()=>'Asia/Seoul'},
  SpreadsheetApp:{ getActiveSpreadsheet:()=>({ getSheetByName:n=> n==='근태기록'?LOG:null }),
                   getUi:()=>({alert:m=>alerts.push(m)}) },
  ContentService:{}, LockService:{}, Logger:{log:m=>alerts.push(m)},
  Utilities:{ getUuid:()=>'u', formatDate:(d,t,f)=>{const p=n=>String(n).padStart(2,'0');
    if(f==='yyyy-MM-dd')return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`;
    if(f==='HH:mm')return `${p(d.getHours())}:${p(d.getMinutes())}`;
    if(f==='yyyy-MM-dd HH:mm')return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
    throw new Error(f);}},
  console,
  // vm 안팎의 Date 가 다른 realm 이면 instanceof 가 실패한다 (테스트 한정 문제)
  Date };
vm.createContext(sandbox); vm.runInContext(src,sandbox);

let pass=0,fail=0;
const eq=(l,g,w)=>{const ok=JSON.stringify(g)===JSON.stringify(w);ok?pass++:fail++;
  console.log(`${ok?'PASS':'FAIL'}  ${l}  →  ${JSON.stringify(g)}${ok?'':'  (기대 '+JSON.stringify(w)+')'}`);};

console.log('— 옛 구조를 그냥 읽었을 때 (마이그레이션 전) —');
const before = sandbox.getRecords();
eq('출근 행은 무시됨', before.every(r=>r.status!=='출근'), true);
eq('1899 날짜가 비고로 새지 않음', before.every(r=>!/1899/.test(r.note)), true);
eq('재택·휴가만 남음', before.map(r=>r.status).sort(), ['재택','재택','휴가','휴가']);

console.log('\n— 마이그레이션 —');
sandbox.migrateOldSheet();
eq('머리글 교체', sheet[0], ['날짜','이름','상태','비고','수정시각']);
const after = sandbox.getRecords();
eq('옮긴 건수', after.length, 5);
eq('지각으로 변환(김범진)', after.filter(r=>r.name==='김범진').map(r=>[r.status,r.note]), [['지각','수업']]);
eq('정상 출근은 제거(김종엽)', after.filter(r=>r.name==='김종엽').length, 0);
eq('상태 분포', after.map(r=>r.status).sort(), ['재택','재택','지각','휴가','휴가']);
eq('비고에 1899 없음', after.every(r=>!/1899/.test(r.note)), true);
eq('안내문에 건수 표기', /옮긴 기록: 5건/.test(alerts.join('\n')), true);

console.log('\n— 다시 실행해도 안전한가 —');
alerts.length=0;
sandbox.migrateOldSheet();
eq('두 번째 실행은 건너뜀', /이미 새 구조/.test(alerts.join('\n')), true);
eq('데이터 보존', sandbox.getRecords().length, 5);

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail?1:0);
