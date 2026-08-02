// 스스로 반례를 더 만들어 깨뜨려 보는 시험
import * as XLSX from 'xlsx';
import fs from 'node:fs';
import { parseSheet, unmergeCells, trimEmpty, normName, buildColumnNames } from '../lib/parse.js';
XLSX.set_fs(fs);

const M = (r1, c1, r2, c2) => ({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });
function sheet(aoa, merges) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  if (merges && merges.length) ws['!merges'] = merges;
  return ws;
}
function show(label, ws, opts) {
  let r;
  try { r = parseSheet(ws, opts); }
  catch (e) { console.log(`\n▶ ${label}\n   ✗ 예외: ${e.message}`); return null; }
  console.log(`\n▶ ${label}`);
  console.log(`   열: ${r.columns.map(c => c.letter + ':' + c.name).join(' | ')}`);
  console.log(`   헤더: 엑셀${JSON.stringify(r.stats.headerExcelRows)}행  데이터 ${r.stats.rowCount}행  결측 ${r.stats.missingTotal}`);
  r.rows.slice(0, 6).forEach((x, i) => console.log(`     [${r.rowMeta[i].excelRow}] ${JSON.stringify(x)}`));
  r.excludedRows.forEach(e => console.log(`     × ${e.excelRow}행 [${e.reason}] ${JSON.stringify(e.values.filter(v => v !== null))}`));
  r.notices.filter(n => n.level !== 'info').forEach(n => console.log(`     ! ${n.type}: ${n.message}`));
  return r;
}

/* 1) 헤더 이름에 줄바꿈·전각공백·단위 괄호 */
show('1. 헤더에 줄바꿈/전각공백', sheet([
  ['지  역', '인구\n(명)', '면적\r\n(㎢)', '　세대수　'],
  ['서울', 9411000, 605.2, 4400000],
]));

/* 2) 한 시트에 표가 두 개 (중간에 빈 행 + 두 번째 제목) */
show('2. 한 시트에 표 두 개 ← 못 푸는 경우', sheet([
  ['1. 인구', null],
  ['지역', '인구'],
  ['서울', 9411000],
  [null, null],
  ['2. 면적', null],
  ['지역', '면적'],
  ['서울', 605.2],
], [M(0, 0, 0, 1), M(4, 0, 4, 1)]));

/* 3) 헤더 첫 칸이 비어 있음 (엑셀에서 아주 흔함) */
show('3. 헤더 첫 칸이 빈 칸', sheet([
  [null, '2021', '2022'],
  ['서울', 970, 960],
  ['부산', 340, 335],
]));

/* 4) 숫자가 문자로 저장됨 ("1,234") + 날짜 */
show('4. 문자로 저장된 숫자와 날짜', sheet([
  ['지역', '인구', '기준일'],
  ['서울', '9,411,000', new Date(2024, 0, 31)],
  ['부산', '3,350,000', new Date(2024, 0, 31)],
]));

/* 5) 제목·설명이 4줄이나 있고 헤더는 5행 */
show('5. 헤더가 5행에 있음', sheet([
  ['2024년 통계자료', null, null],
  [null, null, null],
  ['(단위: 명)', null, null],
  ['작성: 기획예산과', null, null],
  ['지역', '인구', '면적'],
  ['서울', 9411000, 605.2],
  ['부산', 3350000, 770.1],
], [M(0, 0, 0, 2)]));

/* 6) 단위 안내가 헤더 바로 아래(표 중간)에 끼어 있음 */
show('6. 단위 안내가 헤더 바로 아래', sheet([
  ['지역', '인구', '면적'],
  ['※ 단위: 명, ㎢', null, null],
  ['서울', 9411000, 605.2],
  ['부산', 3350000, 770.1],
]));

/* 7) 병합이 헤더와 데이터를 함께 덮음(망가진 파일) */
show('7. 병합이 헤더+데이터를 함께 덮음', sheet([
  ['지역', '인구'],
  [null, 9411000],
  [null, 3350000],
], [M(0, 0, 2, 0)]));

/* 8) 아주 넓은 시트(30열) → 열 문자 AA, AB 확인 */
{
  const head = [], row = [];
  for (let i = 0; i < 30; i++) { head.push('항목' + i); row.push(i); }
  const r = show('8. 30열(열 문자 AA·AD)', sheet([head, row]));
  console.log('   마지막 3열 문자:', r.columns.slice(-3).map(c => c.letter).join(','));
}

/* 9) 2단 헤더인데 하위 이름이 상위와 같음 */
show('9. 상위=하위 이름', sheet([
  ['지역', '인구', null],
  ['지역', '인구', '인구'],
  ['서울', 9411000, 9400000],
], [M(0, 1, 0, 2)]));

/* 10) 세로 병합이 데이터 맨 끝까지 이어짐 */
show('10. 세로 병합이 표 끝까지', sheet([
  ['시도', '구'],
  ['서울', '종로구'],
  [null, '중구'],
  [null, '용산구'],
], [M(1, 0, 3, 0)]));

/* 11) 결측이 아주 많은 행(전부 빈 칸은 아님) */
show('11. 거의 빈 행', sheet([
  ['지역', '인구', '면적', '세대수'],
  ['서울', null, null, null],
  ['부산', 3350000, 770.1, 1500000],
]));

/* 12) 헤더 아래 데이터가 1행뿐 */
show('12. 데이터 1행', sheet([
  ['2024년 자료', null],
  ['지역', '인구'],
  ['서울', 9411000],
], [M(0, 0, 0, 1)]));

/* ── 순수 함수 단위 시험 ─────────────────────────────────── */
console.log('\n▶ 순수 함수 단위 시험');
// 들쭉날쭉한 배열을 unmergeCells 에 직접 넣기
const rag = unmergeCells([[1], [1, 2, 3], []], { merges: [M(0, 0, 2, 2)] });
console.log('   들쭉날쭉 배열 →', JSON.stringify(rag.grid));
// 완전히 빈 표
console.log('   전부 빈 표 trimEmpty →', JSON.stringify(trimEmpty([[null, null], [null, null]]).rows));
console.log('   trimEmpty([]) →', JSON.stringify(trimEmpty([]).rows));
// 전각공백/줄바꿈 정리
console.log('   normName("지  역\\n(명)　") →', JSON.stringify(normName('지  역\n(명)　')));
console.log('   normName(0) →', JSON.stringify(normName(0)), ' normName(false) →', JSON.stringify(normName(false)));
// 헤더 행 번호가 이상할 때
console.log('   buildColumnNames(rows, [99,-1]) →',
  JSON.stringify(buildColumnNames([['가', '나']], [99, -1]).columns.map(c => c.name)));
// 병합 정보가 배열이 아닐 때
console.log('   !merges 가 문자열 →',
  JSON.stringify(parseSheet(Object.assign(sheet([['가'], [1]]), { '!merges': 'nope' })).columns.map(c => c.name)));

/* 13) 성능: 20000행 x 12열 + 4000 병합 */
{
  const aoa = [['시도', '구군', ...Array.from({ length: 10 }, (_, i) => '항목' + i)]];
  const merges = [];
  for (let g = 0; g < 4000; g++) {
    const start = aoa.length;
    for (let k = 0; k < 5; k++) aoa.push([k === 0 ? '시도' + g : null, '구' + g + '_' + k,
      ...Array.from({ length: 10 }, (_, i) => g * 10 + i)]);
    merges.push(M(start, 0, aoa.length - 1, 0));
  }
  const ws = sheet(aoa, merges);
  const t0 = Date.now();
  const r = parseSheet(ws);
  console.log(`\n▶ 13. 성능 20000행 x 12열, 병합 4000곳 → ${r.stats.rowCount}행 ${r.stats.colCount}열, ${Date.now() - t0}ms`);
  console.log('   마지막 행 시도 =', r.rows[r.rows.length - 1]['시도']);
}
