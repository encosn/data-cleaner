/**
 * KOSIS 통계표(반복 생략 모양) 처리 시험
 *
 * 실제 파일: 성씨_본관별_인구_시군구.xlsx (KOSIS DT_1IN15SB)
 *   「성씨, 본관별」 열은 같은 값이 이어질 때 아래 칸을 비워 두었다.
 *   병합 정보가 아예 없어서 병합 해제로는 살릴 수 없고, 반복 생략으로 알아봐야 한다.
 *
 * 파일이 없으면 같은 모양을 직접 만들어 시험한다 (그래서 언제든 돌릴 수 있다).
 */
import * as XLSX from 'xlsx';
import fs from 'node:fs';
import path from 'node:path';
import { parseSheet } from '../lib/parse.js';
import {
  detectRepeatedOmission, fillDownColumns, detectTotalRows,
  analyzeAll, applyRowFilter, valueCounts, filterableColumns,
  recommend, isRecommendedColumn, ACTIONS,
} from '../lib/clean.js';

let pass = 0, fail = 0;
const check = (name, cond, got, want) => {
  if (cond) { pass++; console.log('   OK  ', name); }
  else { fail++; console.log('   FAIL', name, '\n        기대:', want, '\n        실제:', got); }
};

/** 실제 파일이 있으면 쓰고, 없으면 같은 모양을 만든다 */
function loadSheet() {
  const real = path.join(import.meta.dirname, '..', '..', '성씨_본관별_인구_시군구.xlsx');
  try {
    // 브라우저와 같은 방식으로 읽는다: 바이트를 직접 넘긴다.
    // (XLSX.readFile 은 한글 경로에서 실패할 수 있고, 브라우저에는 그 함수가 아예 없다)
    const buf = fs.readFileSync(real);
    const wb = XLSX.read(buf, { type: 'buffer', cellDates: true, dateNF: 'yyyy-mm-dd' });
    return { ws: wb.Sheets[wb.SheetNames[0]], real: true };
  } catch {
    /* 파일이 없거나 열려 있어 못 읽으면 아래 대체본으로 시험한다 */
  }
  // 대체본: 성씨 4개 × (전국 + 시도 3개) = 16행, 아래 칸 생략
  const aoa = [['성씨, 본관별', '행정구역별(시군구)', '2015']];
  const regions = ['전국', '서울특별시', '부산광역시', '대구광역시'];
  [['계', 49705663], ['김(金)', 10689959], ['이(李)', 7306828], ['박(朴)', 4192074]]
    .forEach(([name, total]) => {
      regions.forEach((rg, i) => {
        aoa.push([i === 0 ? name : null, rg, i === 0 ? total : Math.round(total / (i + 4))]);
      });
    });
  return { ws: XLSX.utils.aoa_to_sheet(aoa), real: false };
}

const { ws, real } = loadSheet();
console.log(`${'═'.repeat(78)}\n【KOSIS 반복 생략 처리】 자료: ${real ? '실제 파일' : '대체본(직접 만든 표)'}\n`);

const parsed = parseSheet(ws);
console.log('   열 이름:', parsed.columns.map(c => `${c.letter}:${c.name}`).join(' | '));
console.log('   데이터 행:', parsed.rows.length);
check('열 이름 3개', parsed.columns.length === 3, parsed.columns.length, 3);
check('머리글은 1행', String(parsed.stats.headerExcelRows) === '1', parsed.stats.headerExcelRows, '[1]');
check('병합 정보 없음(이 파일의 특징)', parsed.stats.mergeCount === 0, parsed.stats.mergeCount, 0);

const nameCol = parsed.columns[0].name;
const blank = v => v === null || v === undefined || String(v).trim() === '';
const blanksBefore = parsed.rows.filter(r => blank(r[nameCol])).length;
console.log(`\n   「${nameCol}」 빈 칸: ${blanksBefore} / ${parsed.rows.length}`
  + ` = ${(blanksBefore / parsed.rows.length * 100).toFixed(1)}%`);

// ── 반복 생략 감지
const om = detectRepeatedOmission(parsed.rows, parsed.columns);
console.log('\n   감지 결과:');
om.forEach(o => console.log(`     · ${o.name}: 빈칸 ${o.blanks}, 묶음 ${o.groups}, 평균 ${o.avgRun}칸,`
  + ` 다른열채움 ${(o.otherFilledRate * 100).toFixed(0)}%, 확실=${o.confident}`));
check('성씨 열을 반복 생략으로 찾았다', om.some(o => o.name === nameCol && o.confident),
      JSON.stringify(om.map(o => o.name + ':' + o.confident)), nameCol + ':true');
check('행정구역 열은 반복 생략이 아니다 (빈 칸 없음)',
      !om.some(o => o.name === parsed.columns[1].name), om.map(o => o.name).join(','), '없음');

// ── 채우기
const filled = fillDownColumns(parsed.rows, [nameCol]);
const blanksAfter = filled.filter(r => blank(r[nameCol])).length;
check('채운 뒤 성씨 열에 빈 칸 0개', blanksAfter === 0, blanksAfter, 0);
check('행 수는 그대로', filled.length === parsed.rows.length, filled.length, parsed.rows.length);
check('원본은 바뀌지 않았다',
      parsed.rows.filter(r => blank(r[nameCol])).length === blanksBefore,
      parsed.rows.filter(r => blank(r[nameCol])).length, blanksBefore);
console.log('\n   채운 뒤 앞 5행:');
filled.slice(0, 5).forEach((r, i) => console.log('    ', i, JSON.stringify(r)));

// ── 채우기가 통계 판정을 어떻게 바꾸는가 (이게 이 기능의 존재 이유다)
const statsBefore = analyzeAll(parsed.rows, parsed.columns, new Set());
const statsAfter = analyzeAll(filled, parsed.columns, new Set());
const sb = statsBefore[nameCol], sa = statsAfter[nameCol];
console.log(`\n   채우기 전: 결측률 ${(sb.missingRate * 100).toFixed(1)}%,`
  + ` 추천=${recommend(sb).action}, 쓸만한열=${isRecommendedColumn(sb)}`);
console.log(`   채우기 후: 결측률 ${(sa.missingRate * 100).toFixed(1)}%,`
  + ` 추천=${recommend(sa).action}, 쓸만한열=${isRecommendedColumn(sa)}`);
check('채우기 전에는 이 열을 쓰지 말라고 권한다 (문제 상황)',
      isRecommendedColumn(sb) === false, isRecommendedColumn(sb), false);
check('채운 뒤에는 쓸만한 열로 바뀐다 (해결)',
      isRecommendedColumn(sa) === true, isRecommendedColumn(sa), true);
check('채운 뒤 성씨 값 종류가 2개 이상', sa.distinct >= 2, sa.distinct, '>=2');

// ── 합계 줄 감지 (「계」와 「전국」)
const totals = detectTotalRows(filled, parsed.columns);
const labelSet = new Set(totals.labels);
console.log(`\n   합계로 본 줄: ${totals.idxs.length}개, 라벨 종류: ${[...labelSet].slice(0, 5).join(', ')}`);
check('「계」 또는 「전국」 줄을 찾았다', totals.idxs.length > 0, totals.idxs.length, '>0');

// ── 행 고르기
const fcols = filterableColumns(filled, parsed.columns, statsAfter);
console.log(`\n   행 고르기에 쓸 수 있는 열: ${fcols.map(c => c.name).join(', ') || '(없음)'}`);
check('행정구역 열로 행을 고를 수 있다', fcols.some(c => c.name === parsed.columns[1].name),
      fcols.map(c => c.name).join(','), parsed.columns[1].name);

const regionCol = parsed.columns[1].name;
const vc = valueCounts(filled, regionCol);
console.log(`   ${regionCol} 값 ${vc.length}개, 앞 3개:`, vc.slice(0, 3).map(v => `${v.value}(${v.n}행)`).join(' '));
const seoul = vc.find(v => String(v.value).includes('서울'));
if (seoul) {
  const f = applyRowFilter(filled, { [regionCol]: new Set([seoul.value]) });
  check('서울만 골라내면 서울 행만 남는다',
        f.rows.length === seoul.n && f.rows.every(r => r[regionCol] === seoul.value),
        f.rows.length, seoul.n);
  check('골라낸 행에도 성씨 값이 살아 있다',
        f.rows.every(r => !blank(r[nameCol])), 'blank 있음', '전부 채워짐');
  console.log('   서울 앞 3행:', f.rows.slice(0, 3).map(r => JSON.stringify(r)).join(' '));
}

// 빈 Set = 전부 빼기
const none = applyRowFilter(filled, { [regionCol]: new Set() });
check('빈 Set 은 「전부 빼기」 (0행)', none.rows.length === 0, none.rows.length, 0);
// 키 없음 = 걸러내지 않음
const all = applyRowFilter(filled, {});
check('키가 없으면 걸러내지 않는다', all.rows.length === filled.length, all.rows.length, filled.length);

// ── 반복 생략으로 잘못 볼 위험이 없는지 (진짜 결측치는 채우면 안 된다)
const wsMiss = XLSX.utils.aoa_to_sheet([
  ['지역', '인구', '비고'],
  ['서울', 100, '많음'],
  ['부산', null, null],      // 진짜 결측 (다른 열도 비었다)
  ['대구', 300, null],
]);
const pm = parseSheet(wsMiss);
const om2 = detectRepeatedOmission(pm.rows, pm.columns);
console.log('\n   진짜 결측치가 있는 표:', om2.length ? om2.map(o => o.name + ':' + o.confident).join(', ') : '반복 생략 없음');
check('진짜 결측치를 반복 생략으로 오인하지 않는다',
      !om2.some(o => o.confident), JSON.stringify(om2.map(o => o.name + ':' + o.confident)), '확실한 것 없음');

console.log(`\n${'═'.repeat(78)}\n결과: 통과 ${pass} / 실패 ${fail}`);
process.exit(fail ? 1 : 0);
