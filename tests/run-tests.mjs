import * as XLSX from 'xlsx';
import fs from 'node:fs';
import path from 'node:path';
import { parseSheet, unmergeCells, readGrid, toAOA, pickColumns } from '../lib/parse.js';

XLSX.set_fs(fs);
const DIR = path.join(import.meta.dirname, 'fixtures');

let pass = 0, fail = 0;
function check(label, ok, got, want) {
  if (ok) { pass++; console.log(`   OK   ${label}`); }
  else { fail++; console.log(`   FAIL ${label}\n        기대: ${JSON.stringify(want)}\n        실제: ${JSON.stringify(got)}`); }
}
const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);

function load(name) {
  const wb = XLSX.readFile(path.join(DIR, name + '.xlsx'), { cellDates: true });
  return wb.Sheets[wb.SheetNames[0]];
}

function report(name, res, { showRows = 99 } = {}) {
  console.log(`\n${'═'.repeat(78)}\n【${name}】`);
  console.log(' 열 이름 :', res.columns.map(c => `${c.letter}:${c.name}`).join(' | '));
  console.log(` 헤더행  : 엑셀 ${JSON.stringify(res.stats.headerExcelRows)}행` +
              `   병합 ${res.stats.mergeCount}곳 → 채운 칸 ${res.stats.filledCellCount}개`);
  console.log(` 데이터  : ${res.stats.rowCount}행 x ${res.stats.colCount}열, 결측 ${res.stats.missingTotal}개`,
              JSON.stringify(res.stats.missingByColumn));
  res.rows.slice(0, showRows).forEach((r, i) =>
    console.log(`   [엑셀${String(res.rowMeta[i].excelRow).padStart(3)}행]`, JSON.stringify(r)));
  if (res.rows.length > showRows) console.log(`   ... (총 ${res.rows.length}행)`);
  for (const n of res.notices) console.log(`   · (${n.level}/${n.type}) ${n.message}`);
  for (const e of res.excludedRows) console.log(`   × 제외 엑셀${e.excelRow}행 [${e.reason}]`,
    JSON.stringify(e.values.filter(v => v !== null)));
  console.log(' 헤더 후보(점수순):');
  res.headerInfo.candidates.slice(0, 3).forEach(c =>
    console.log(`   ${c.score.toString().padStart(6)}  행${JSON.stringify(c.rowIdxs.map(i => i + 1))}` +
                `  → ${c.preview.join(' | ')}`));
}

/* ═══ A ═══ */
{
  const res = parseSheet(load('A_제목병합'));
  report('A) 1행 제목 가로병합 + 2행 헤더', res);
  check('A 열이름', eq(res.columns.map(c => c.name), ['지역', '인구', '면적', '세대수']),
        res.columns.map(c => c.name), ['지역', '인구', '면적', '세대수']);
  check('A 데이터 3행', res.rows.length === 3, res.rows.length, 3);
  check('A 첫행 값', res.rows[0]['지역'] === '서울특별시' && res.rows[0]['인구'] === 9411000,
        res.rows[0], '서울특별시/9411000');
  check('A 제목행은 데이터에서 제외', res.excludedRows.length === 1 && res.excludedRows[0].excelRow === 1,
        res.excludedRows, '엑셀1행 제외');
}

/* ═══ B ═══ */
{
  const res = parseSheet(load('B_2단헤더'));
  report('B) 헤더 2행 (지역 / 인구=남·여 / 면적)', res);
  const want = ['지역', '인구 남', '인구 여', '면적'];
  check('B 열이름', eq(res.columns.map(c => c.name), want), res.columns.map(c => c.name), want);
  check('B 데이터 3행', res.rows.length === 3, res.rows.length, 3);
  check('B 값 매핑', res.rows[0]['인구 남'] === 4600000 && res.rows[0]['인구 여'] === 4811000,
        res.rows[0], '4600000/4811000');
}

/* ═══ C ═══ */
{
  const res = parseSheet(load('C_데이터세로병합'));
  report('C) 데이터 영역 세로병합', res);
  check('C 세로병합 3행 모두 서울특별시',
        res.rows.slice(0, 3).every(r => r['시도'] === '서울특별시'),
        res.rows.map(r => r['시도']), ['서울특별시', '서울특별시', '서울특별시', '부산광역시', '부산광역시']);
  check('C 부산 2행', res.rows.slice(3).every(r => r['시도'] === '부산광역시'),
        res.rows.map(r => r['시도']), '부산광역시x2');
  check('C 데이터 5행', res.rows.length === 5, res.rows.length, 5);
  check('C 구 값 유지', eq(res.rows.map(r => r['구']), ['종로구', '중구', '용산구', '중구', '서구']),
        res.rows.map(r => r['구']), ['종로구', '중구', '용산구', '중구', '서구']);
}

/* ═══ D ═══ */
{
  const res = parseSheet(load('D_최악'));
  report('D) 최악(제목+2단헤더+세로병합+빈행+주석+결측)', res);
  const want = ['시도', '구군', '인구 남', '인구 여', '면적'];
  check('D 열이름', eq(res.columns.map(c => c.name), want), res.columns.map(c => c.name), want);
  check('D 데이터 5행', res.rows.length === 5, res.rows.length, 5);
  check('D 시도 채움', eq(res.rows.map(r => r['시도']),
        ['서울특별시', '서울특별시', '서울특별시', '부산광역시', '부산광역시']),
        res.rows.map(r => r['시도']), '서울x3 부산x2');
  check('D 결측 유지(3개)', res.stats.missingTotal === 3, res.stats.missingByColumn, 3);
  check('D 주석 2줄 제외', res.excludedRows.filter(e => e.reason.includes('설명')).length === 2,
        res.excludedRows, 2);
  check('D 원본 행번호 대조', eq(res.rowMeta.map(m => m.excelRow), [4, 5, 6, 8, 9]),
        res.rowMeta.map(m => m.excelRow), [4, 5, 6, 8, 9]);
}

/* ═══ E ═══ */
{
  const res = parseSheet(load('E_평범'));
  report('E) 병합 없는 평범한 파일', res);
  check('E 열이름', eq(res.columns.map(c => c.name), ['연도', '학생수', '학급수', '학교수']),
        res.columns.map(c => c.name), ['연도', '학생수', '학급수', '학교수']);
  check('E 데이터 4행', res.rows.length === 4, res.rows.length, 4);
  check('E 결측 0', res.stats.missingTotal === 0, res.stats.missingTotal, 0);
  check('E 제외행 없음', res.excludedRows.length === 0, res.excludedRows, []);
}

/* ═══ F ═══ */
{
  const res = parseSheet(load('F_3단헤더'));
  report('F) 헤더 3행', res);
  const want = ['지역', '학교 초등학교 학생수', '학교 초등학교 학급수', '학교 중학교 학생수', '학교 중학교 학급수'];
  check('F 열이름', eq(res.columns.map(c => c.name), want), res.columns.map(c => c.name), want);
  check('F 데이터 2행', res.rows.length === 2, res.rows.length, 2);
}

/* ═══ G ═══ */
{
  const res = parseSheet(load('G_빈병합'));
  report('G) 병합인데 값이 빈 칸', res);
  check('G 죽지 않음', res.columns.length === 3, res.columns.length, 3);
  check('G 세로병합 채움', res.rows.every(r => r[res.columns[0].name] === '서울'),
        res.rows, '서울 x2');
}

/* ═══ H ═══ */
{
  const res = parseSheet(load('H_숫자헤더'));
  report('H) 열 이름이 숫자', res);
  check('H 열이름', eq(res.columns.map(c => c.name), ['지역', '2020', '2021', '2022']),
        res.columns.map(c => c.name), ['지역', '2020', '2021', '2022']);
  check('H 데이터 2행', res.rows.length === 2, res.rows.length, 2);
}

/* ═══ I ═══ */
{
  const res = parseSheet(load('I_중복이름'));
  report('I) 중복 열 이름', res);
  check('I 유일화', eq(res.columns.map(c => c.name), ['지역', '인구', '인구(2)', '면적']),
        res.columns.map(c => c.name), ['지역', '인구', '인구(2)', '면적']);
  check('I 값 안 겹침', res.rows[0]['인구'] === 9411000 && res.rows[0]['인구(2)'] === 9400000,
        res.rows[0], '9411000/9400000');
}

/* ═══ J ═══ */
{
  const ws = load('J_이상한병합');
  // 저장할 수 없는 "완전히 망가진 병합"을 메모리에서 끼워 넣는다
  ws['!merges'] = (ws['!merges'] || []).concat([null, { s: null, e: null },
    { s: { r: 'x' }, e: {} }, { s: { r: -5, c: -2 }, e: { r: 1, c: 1 } }, 'abc', 42]);
  let res;
  try { res = parseSheet(ws); } catch (e) { res = null; console.log('   예외:', e.message); }
  if (res) report('J) 이상한 병합(겹침·역방향·범위밖·망가짐)', res);
  check('J 예외 없이 처리', !!res && res.rows.length === 2, res && res.rows.length, 2);
  check('J 헤더가 망가지지 않음', !!res && eq(res.columns.map(c => c.name), ['지역', '인구']),
        res && res.columns.map(c => c.name), ['지역', '인구']);
  check('J 겹치는 병합: 먼저 나온 것이 이김', !!res && res.rows.every(r => r['지역'] === '서울'),
        res && res.rows.map(r => r['지역']), ['서울', '서울']);
  check('J 엉터리 큰 병합으로 표가 커지지 않음', !!res && res.stats.colCount === 2,
        res && res.stats.colCount, 2);
}

/* ═══ K ═══ */
{
  const res = parseSheet(load('K_전부글자'));
  report('K) 데이터가 전부 글자', res);
  check('K 헤더 1행만', eq(res.stats.headerExcelRows, [1]), res.stats.headerExcelRows, [1]);
  check('K 데이터 3행', res.rows.length === 3, res.rows.length, 3);
}

/* ═══ L ═══ */
{
  const ws = load('L_대용량');
  const t0 = Date.now();
  const res = parseSheet(ws);
  const ms = Date.now() - t0;
  report('L) 1000행 + 세로병합 200개 (성능)', res, { showRows: 3 });
  console.log(`   ⏱ parseSheet ${ms}ms`);
  check('L 데이터 1000행', res.rows.length === 1000, res.rows.length, 1000);
  check('L 마지막 그룹 채움', res.rows[999]['시도'] === '시도199', res.rows[999], '시도199');
  check('L 200ms 이내', ms < 200, ms + 'ms', '<200ms');
}

/* ═══ M ═══ */
{
  const res = parseSheet(load('M_빈열'));
  report('M) 중간에 빈 열', res);
  check('M 빈 열 제거', eq(res.columns.map(c => c.name), ['지역', '인구']),
        res.columns.map(c => c.name), ['지역', '인구']);
  check('M 열 문자 보존(A,C)', eq(res.columns.map(c => c.letter), ['A', 'C']),
        res.columns.map(c => c.letter), ['A', 'C']);
}

/* ═══ 극단 케이스 ═══ */
{
  console.log(`\n${'═'.repeat(78)}\n【극단 케이스】`);
  const cases = [
    ['빈 워크시트', {}],
    ['null', null],
    ['!ref 만 있음', { '!ref': 'A1:C3' }],
    ['!ref 없음', { A1: { t: 's', v: '지역' }, A2: { t: 's', v: '서울' } }],
    ['!ref 가 틀림(작게)', { '!ref': 'A1:A1', A1: { t: 's', v: '지역' }, B1: { t: 's', v: '인구' },
                            A2: { t: 's', v: '서울' }, B2: { t: 'n', v: 5 } }],
    ['헤더만 있고 데이터 0행', { '!ref': 'A1:B1', A1: { t: 's', v: '지역' }, B1: { t: 's', v: '인구' } }],
    ['오류셀', { '!ref': 'A1:B2', A1: { t: 's', v: '지역' }, B1: { t: 's', v: '인구' },
                A2: { t: 's', v: '서울' }, B2: { t: 'e', v: 0x07, w: '#DIV/0!' } }],
  ];
  for (const [label, ws] of cases) {
    try {
      const r = parseSheet(ws);
      console.log(`   OK   ${label} → ${r.stats.rowCount}행 ${r.stats.colCount}열 [${r.columns.map(c => c.name).join(',')}]`);
      pass++;
    } catch (e) { console.log(`   FAIL ${label} → 예외 ${e.message}`); fail++; }
  }
  // 사용자가 헤더를 직접 고르는 경우
  const res = parseSheet(load('D_최악'), { headerExcelRows: [2] });
  console.log('   사용자가 엑셀 2행만 헤더로 고름 →', res.columns.map(c => c.name).join(' | '),
              `/ ${res.rows.length}행`);
  check('사용자 지정 헤더 반영', res.columns.map(c => c.name).join()==='시도,구군,인구,인구(2),면적',
        res.columns.map(c => c.name), ['시도','구군','인구','인구(2)','면적']);
  const res2 = parseSheet(load('D_최악'), { keepNotes: true });
  check('keepNotes:true 면 주석도 데이터', res2.rows.length === 7, res2.rows.length, 7);
  // 원본 불변 확인
  const ws2 = load('C_데이터세로병합');
  const before = JSON.stringify(ws2);
  parseSheet(ws2);
  check('원본 워크시트 불변', JSON.stringify(ws2) === before, 'changed', 'unchanged');

  // 숫자 열 이름일 때 Object.keys 순서가 뒤바뀌는 문제 → toAOA 는 안전한가
  const h = parseSheet(load('H_숫자헤더'));
  console.log('   Object.keys(row) 순서 :', Object.keys(h.rows[0]).join(','), ' ← 뒤바뀜(쓰면 안 됨)');
  console.log('   toAOA 순서            :', toAOA(h)[0].join(','));
  check('toAOA 는 열 순서 보존', eq(toAOA(h)[0], ['지역', '2020', '2021', '2022']),
        toAOA(h)[0], ['지역', '2020', '2021', '2022']);
  check('toAOA 데이터 행', eq(toAOA(h)[1], ['서울', 970, 960, 950]), toAOA(h)[1], ['서울', 970, 960, 950]);

  // 속성 고르기
  const d = parseSheet(load('D_최악'));
  const p = pickColumns(d, ['시도', '면적']);
  check('pickColumns', eq(p.columns.map(c => c.name), ['시도', '면적']) && p.rows.length === 5,
        p.columns.map(c => c.name), ['시도', '면적']);

  // 열 이름이 __proto__ 인 최악의 경우
  const ws3 = XLSX.utils.aoa_to_sheet([['__proto__', '값'], ['가', 1], ['나', 2]]);
  const r3 = parseSheet(ws3);
  check('__proto__ 열 이름도 값이 살아 있음',
        r3.rows.length === 2 && r3.rows[0][r3.columns[0].name] === '가',
        r3.columns.map(c => c.name) + ' / ' + JSON.stringify(r3.rows), '값 보존');

  // 왕복 검증: parseSheet → toAOA → 엑셀 저장 → 다시 parseSheet
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(toAOA(d)), '정제결과');
  const tmp = path.join(DIR, '_왕복.xlsx');
  XLSX.writeFile(wb, tmp);
  const wbBack = XLSX.readFile(tmp, { cellDates: true });
  const back = parseSheet(wbBack.Sheets[wbBack.SheetNames[0]]);
  check('왕복 후 열 이름 동일', eq(back.columns.map(c => c.name), d.columns.map(c => c.name)),
        back.columns.map(c => c.name), d.columns.map(c => c.name));
  check('왕복 후 행 수 동일', back.rows.length === d.rows.length, back.rows.length, d.rows.length);
  console.log('   왕복 결과 첫 행:', JSON.stringify(back.rows[0]));
}

/* ═══ 새 규칙의 오작동(진짜 데이터를 지우지 않는지) 확인 ═══ */
{
  console.log(`\n${'═'.repeat(78)}\n【새 규칙 오작동 확인】`);
  const S = (aoa, merges) => {
    const w = XLSX.utils.aoa_to_sheet(aoa);
    if (merges) w['!merges'] = merges;
    return w;
  };

  // "출처불명", "비고사항" 처럼 주석 낱말로 시작하는 진짜 데이터는 지우면 안 된다
  const a = parseSheet(S([['지역', '상태'], ['출처불명', '확인중'], ['비고사항', '없음'],
                          ['참고자료실', '운영']]));
  check('주석 낱말로 시작하는 데이터는 남는다', a.rows.length === 3, a.rows.length, 3);

  // 값이 1칸뿐인 진짜 데이터 행도 남아야 한다
  const b = parseSheet(S([['지역', '인구'], ['서울', null], ['부산', null]]));
  check('값이 적은 데이터 행도 남는다', b.rows.length === 2, b.rows.length, 2);

  // 긴 표 중간에 열 이름 줄이 다시 나오는 경우(인쇄용 파일에서 흔함)
  const aoa = [['시도', '구군', '인구']];
  for (let i = 0; i < 30; i++) {
    if (i === 10 || i === 20) aoa.push(['시도', '구군', '인구']);   // 반복된 헤더
    aoa.push(['서울', '구' + i, 1000 + i]);
  }
  const c = parseSheet(S(aoa));
  check('반복된 헤더 줄 2개만 제외', c.rows.length === 30 &&
        c.excludedRows.filter(e => e.reason.includes('다시 나온')).length === 2,
        c.rows.length + '행/' + c.excludedRows.length + '제외', '30행/2제외');

  // 주석이 두 번째 열에 있는 경우
  const d = parseSheet(S([['지역', '인구'], ['서울', 100], [null, '※ 자료: 통계청']]));
  check('둘째 칸의 주석도 감지', d.rows.length === 1, d.rows.length, 1);

  // 데이터가 0행이 되어도 죽지 않는가
  const e = parseSheet(S([['지역', '인구'], ['※ 자료: 통계청', null]]));
  check('주석만 남아도 안전', e.rows.length === 0 && e.columns.length === 2,
        e.rows.length + '/' + e.columns.length, '0/2');

  // 날짜 서식 숫자(cellDates 없이 읽은 경우)
  const ws = { '!ref': 'A1:B2', A1: { t: 's', v: '지역' }, B1: { t: 's', v: '기준일' },
               A2: { t: 's', v: '서울' }, B2: { t: 'n', v: 45322, z: 'yyyy-mm-dd' } };
  const f = parseSheet(ws);
  check('날짜 숫자를 날짜로 바꿈', f.rows[0]['기준일'] === '2024-01-31', f.rows[0]['기준일'], '2024-01-31');
}

console.log(`\n${'═'.repeat(78)}\n결과: 통과 ${pass} / 실패 ${fail}`);
process.exit(fail ? 1 : 0);
