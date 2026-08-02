/**
 * main.js — 데이터 정제 실습실 (화면과 조작)
 * ─────────────────────────────────────────────────────────────────────────────
 * 계산은 lib/ 폴더의 세 파일이 맡고, 이 파일은 "화면을 그리고 클릭을 받는 일"만 한다.
 *   lib/parse.js  … 엑셀 → 열 이름 기준의 표 (병합 셀 풀기, 머리글 찾기)
 *   lib/clean.js  … 빈 칸 찾기, 열 성격 살피기, 「원본 + 설정 = 결과」 계산
 *   lib/io.js     … 파일 읽기(한글 깨짐 방지), 엑셀·CSV 저장
 *
 * ★ 가장 중요한 설계 ★
 *   원본(S.parsed)은 절대 바꾸지 않는다. 학생의 선택은 S.plan / S.picked 에만 쌓고,
 *   화면을 그릴 때마다 applyPlan() 으로 다시 계산한다.
 *   → 무엇을 눌러도 되돌릴 수 있고, "누른 순서 때문에 결과가 달라지는" 버그가 없다.
 * ─────────────────────────────────────────────────────────────────────────────
 */

import * as XLSX from 'xlsx';
import { parseSheet } from './lib/parse.js';
import {
  MISSING_TOKEN_GROUPS, isMissing, countTokens, analyzeAll, ACTIONS,
  optionsFor, recommend, warnFor, applyPlan, isRecommendedColumn, coachColumn,
  detectTotalRows, detectRepeatedOmission, fillDownColumns,
  filterableColumns, valueCounts, applyRowFilter,
  fmtNum, pct, cellText, toNumber, roundLike,
} from './lib/clean.js';
import { readWorkbook, saveXlsx, saveCsv, MAX_BYTES, badFileNameChars } from './lib/io.js';

/* ═══════════════════════════════════════════════════════════════════════════
   상태
   ═══════════════════════════════════════════════════════════════════════════ */

const S = {
  step: 1,
  fileName: null,
  fileSize: 0,
  rawFile: null,          // 인코딩을 바꿔 다시 읽을 때 쓴다
  workbook: null,
  sheetNames: [],
  sheetName: null,
  encoding: null,
  forcedEncoding: null,   // 자동 판정이 틀렸을 때 학생이 직접 고른 글자 방식
  ioNotices: [],
  parsed: null,           // ← 원본. 절대 바꾸지 않는다
  headerExcelRows: null,  // 학생이 머리글 줄을 직접 고른 경우
  tokens: new Set(),      // 빈 칸으로 볼 표기 (기본은 아무것도 안 켬)
  tokenCounts: {},
  stats: {},
  plan: {},               // { 열이름: {action, value} }
  picked: [],             // 3단계에서 고른 열 이름 (순서 그대로)
  renames: {},
  goal: '',
  previewMode: 'after',
  previewRows: 20,
  result: null,           // applyPlan 결과 — 고른 열 기준 (4단계 저장용)
  resultAll: null,        // applyPlan 결과 — 모든 열 기준 (3단계 요약용)
  filledRows: null,       // 반복 생략만 되살린 자료 (행 걸러내기 전)
  statsUnfiltered: {},    // 걸러내기 전 통계 — 행 고르기 목록을 만들 때 쓴다
  quiz: null,
  quizPicked: {},
  quizDone: {},
  warnedOnce: {},         // 같은 경고를 반복해서 띄우지 않기 위해
  undoStack: [],          // 되돌리기용 이전 설정 스냅샷
  redoStack: [],
  totalRows: { idxs: [], labels: [] },  // 합계·전체로 보이는 행
  excludeTotals: false,   // ★ 기본은 끔 — 합계 줄을 남기면 평균이 어떻게 달라지는지 보는 것이 학습 목표
  omission: [],           // 반복 생략으로 보이는 열 목록
  fillDown: new Set(),    // 「위 값으로 채우기」를 적용할 열 이름
  rowFilter: {},          // { 열이름: Set<string> } — 그 열에서 남길 값들
  base: null,             // 반복생략 채우기 + 행 걸러내기를 마친 작업용 자료
};

/* ═══════════════════════════════════════════════════════════════════════════
   되돌리기
   ───────────────────────────────────────────────────────────────────────────
   원본을 절대 바꾸지 않고 "설정"만 쌓아 두는 구조라서, 되돌리기는
   설정을 예전 것으로 갈아 끼우는 것만으로 끝난다. 표를 되돌릴 필요가 없다.

   되돌리기는 기능이 아니라 교육 장치다.
   되돌릴 수 있으니 학생이 겁내지 않고 여러 방법을 눌러 보고,
   숫자가 어떻게 바뀌는지 관찰한다 — 이 앱의 학습 효과는 거기서 나온다.
   ═══════════════════════════════════════════════════════════════════════════ */

/** 지금 설정을 사진처럼 찍어 둔다 */
function snapshot() {
  const rf = {};
  for (const k of Object.keys(S.rowFilter)) rf[k] = [...S.rowFilter[k]];
  return {
    plan: JSON.parse(JSON.stringify(S.plan)),
    picked: S.picked.slice(),
    renames: JSON.parse(JSON.stringify(S.renames)),
    tokens: [...S.tokens],
    excludeTotals: S.excludeTotals,
    fillDown: [...S.fillDown],
    rowFilter: rf,
    label: '',
  };
}

/** 무언가를 바꾸기 **직전에** 부른다. label 은 토스트에 보여 줄 이름. */
function pushUndo(label) {
  const snap = snapshot();
  snap.label = label;
  S.undoStack.push(snap);
  if (S.undoStack.length > 40) S.undoStack.shift();   // 너무 많이 쌓이지 않게
  S.redoStack.length = 0;                            // 새 조작을 하면 '다시 하기'는 사라진다
}

function restore(snap) {
  S.plan = JSON.parse(JSON.stringify(snap.plan));
  S.picked = snap.picked.slice();
  S.renames = JSON.parse(JSON.stringify(snap.renames));
  S.tokens = new Set(snap.tokens);
  S.excludeTotals = !!snap.excludeTotals;
  S.fillDown = new Set(snap.fillDown || []);
  S.rowFilter = {};
  for (const k of Object.keys(snap.rowFilter || {})) S.rowFilter[k] = new Set(snap.rowFilter[k]);
  recompute();
}

function doUndo() {
  if (!S.undoStack.length) { toast('되돌릴 것이 없어요.'); return; }
  const cur = snapshot();
  const snap = S.undoStack.pop();
  cur.label = snap.label;
  S.redoStack.push(cur);
  restore(snap);
  render();
  toast(`되돌렸어요${snap.label ? ` · 「${snap.label}」 취소` : ''}`, '다시 하기', doRedo);
}

function doRedo() {
  if (!S.redoStack.length) { toast('다시 할 것이 없어요.'); return; }
  const cur = snapshot();
  const snap = S.redoStack.pop();
  S.undoStack.push(cur);
  restore(snap);
  render();
  toast('다시 했어요.');
}

const $ = (id) => document.getElementById(id);

/** 파일에서 온 글자를 화면에 넣을 때는 반드시 이걸 통과시킨다 (안전) */
function esc(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/* ═══════════════════════════════════════════════════════════════════════════
   작은 화면 도구들
   ═══════════════════════════════════════════════════════════════════════════ */

function noticeHtml(level, text, strong) {
  const cls = { info: 'notice-info', ok: 'notice-ok', warn: 'notice-warn', error: 'notice-err' }[level] || 'notice-info';
  const icon = { info: 'ℹ', ok: '✔', warn: '⚠', error: '✖' }[level] || 'ℹ';
  return `<div class="notice ${cls}"><span class="notice-icon" aria-hidden="true">${icon}</span>`
       + `<div>${strong ? `<strong>${esc(strong)}</strong>` : ''}${esc(text)}</div></div>`;
}

function toast(text, actionLabel, actionFn) {
  const area = $('toastArea');
  const el = document.createElement('div');
  el.className = 'toast';
  const span = document.createElement('span');
  span.style.flex = '1';
  span.textContent = text;
  el.appendChild(span);
  if (actionLabel) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'btn';
    b.textContent = actionLabel;
    b.onclick = () => { el.remove(); actionFn && actionFn(); };
    el.appendChild(b);
  }
  area.appendChild(el);
  setTimeout(() => el.remove(), 6000);
}

/** 확인 창. 학생이 「확인」을 누르면 true */
function ask(title, bodyHtml, okLabel = '확인', cancelLabel = '그만두기') {
  return new Promise((resolve) => {
    const dlg = $('askDialog');
    $('askTitle').textContent = title;
    $('askBody').innerHTML = bodyHtml;
    $('askOk').textContent = okLabel;
    $('askCancel').textContent = cancelLabel;
    const done = (v) => {
      $('askOk').onclick = null;
      $('askCancel').onclick = null;
      dlg.close();
      resolve(v);
    };
    $('askOk').onclick = () => done(true);
    $('askCancel').onclick = () => done(false);
    dlg.showModal();
  });
}

/* ═══════════════════════════════════════════════════════════════════════════
   미리보기 표 (2·3·4단계가 함께 쓴다)
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * @param {HTMLElement} box
 * @param {object} p
 *   columns   : [{name, letter}]
 *   rows      : 보여 줄 행 (객체 배열)
 *   totalRows : 전체 행 수 (안내 문구용)
 *   filledMask: 행마다 { 열이름: true } — 우리가 채운 칸
 *   dropMask  : 행마다 true — 지워질 행
 *   rowLabels : 행마다 표시할 번호 (원본 엑셀 행 번호)
 *   stats     : 열 소제목에 쓸 통계
 */
function renderTable(box, p) {
  const limit = S.previewRows;
  const show = p.rows.slice(0, limit);

  const tools = document.createElement('div');
  tools.className = 'table-tools';
  const info = document.createElement('span');
  info.textContent = p.columns.length === 0
    ? '보여 줄 열이 없어요.'
    : `모두 ${fmtNum(p.totalRows)}행 × ${p.columns.length}열 · 아래는 앞에서 ${Math.min(limit, p.rows.length)}행만 보여 줍니다.`;
  tools.appendChild(info);
  for (const n of [20, 50, 100]) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'btn btn-sm';
    b.textContent = `${n}행 보기`;
    if (n === limit) b.disabled = true;
    b.onclick = () => { S.previewRows = n; render(); };
    tools.appendChild(b);
  }
  box.innerHTML = '';
  box.appendChild(tools);

  if (!p.columns.length) return;

  const wrapper = document.createElement('div');
  wrapper.className = 'table-box';
  const table = document.createElement('table');
  table.className = 'grid';

  // ── 머리글
  const thead = document.createElement('thead');
  const tr = document.createElement('tr');
  const th0 = document.createElement('th');
  th0.className = 'rownum';
  th0.textContent = '엑셀행';
  tr.appendChild(th0);
  for (const c of p.columns) {
    const th = document.createElement('th');
    th.textContent = c.name;
    const st = p.stats && p.stats[c.name];
    if (st) {
      const sub = document.createElement('span');
      sub.className = 'th-sub';
      const kind = { number: '숫자', text: '글자', date: '날짜', empty: '빈열' }[st.type] || '글자';
      sub.textContent = `${c.letter}열 · ${kind} · 빈 칸 ${pct(st.missingRate)}`;
      th.appendChild(sub);
    }
    tr.appendChild(th);
  }
  thead.appendChild(tr);
  table.appendChild(thead);

  // ── 본문
  const tbody = document.createElement('tbody');
  show.forEach((row, i) => {
    const rtr = document.createElement('tr');
    if (p.dropMask && p.dropMask[i]) rtr.className = 'row-drop';
    const rn = document.createElement('td');
    rn.className = 'rownum';
    rn.textContent = p.rowLabels ? p.rowLabels[i] : (i + 1);
    if (p.dropMask && p.dropMask[i]) rn.title = '이 행은 지워집니다';
    rtr.appendChild(rn);

    for (const c of p.columns) {
      const td = document.createElement('td');
      const v = row[c.name];
      const st = p.stats && p.stats[c.name];
      if (isMissing(v, S.tokens)) {
        td.className = 'cell-missing';
        td.textContent = '빈 칸';
        td.title = '값이 없는 칸(결측치)입니다';
      } else {
        td.className = (st && st.type === 'number') ? 'cell-num' : 'cell-text';
        td.textContent = cellText(v, st);
        if (p.filledMask && p.filledMask[i] && p.filledMask[i][c.name]) {
          td.classList.add('cell-filled');
          const tag = document.createElement('span');
          tag.className = 'tag-filled';
          tag.textContent = '채움';
          td.appendChild(tag);
          td.title = '원래는 빈 칸이었고, 우리가 값을 채웠어요';
        }
      }
      rtr.appendChild(td);
    }
    tbody.appendChild(rtr);
  });
  table.appendChild(tbody);
  wrapper.appendChild(table);
  box.appendChild(wrapper);
}

/* ═══════════════════════════════════════════════════════════════════════════
   1단계 · 파일 읽기
   ═══════════════════════════════════════════════════════════════════════════ */

async function openFile(file) {
  $('fileError').innerHTML = '';
  const name = file.name || '파일';
  const ext = name.slice(name.lastIndexOf('.')).toLowerCase();

  if (!['.xlsx', '.xls', '.csv', '.txt'].includes(ext)) {
    $('fileError').innerHTML = noticeHtml('error',
      `이 파일은 열 수 없어요. .xlsx, .xls, .csv 파일을 골라 주세요. (지금 고른 파일: ${ext || '알 수 없음'})`);
    return;
  }
  if (file.size > MAX_BYTES) {
    $('fileError').innerHTML = noticeHtml('error',
      `파일이 너무 커요(${(file.size / 1048576).toFixed(1)}MB). 20MB보다 작은 파일을 써 주세요. `
      + '엑셀에서 필요한 부분만 남겨 다시 저장하면 됩니다.');
    return;
  }
  if (S.parsed) {
    const ok = await ask('새 파일을 열까요?',
      '<p>이미 열어 둔 자료가 있어요. 새 파일을 열면 지금까지 한 정리가 사라집니다.</p>',
      '새 파일 열기');
    if (!ok) return;
  }

  try {
    const res = await readWorkbook(file, S.forcedEncoding ? { encoding: S.forcedEncoding } : {});
    S.rawFile = file;
    S.fileName = name;
    S.fileSize = file.size;
    S.workbook = res.workbook;
    S.sheetNames = res.workbook.SheetNames || [];
    S.encoding = res.encoding;
    S.ioNotices = res.notices || [];
    if (!S.sheetNames.length) {
      $('fileError').innerHTML = noticeHtml('error', '파일에 자료가 없어요. 다른 파일을 열어 보세요.');
      return;
    }
    S.sheetName = S.sheetNames[0];
    S.headerExcelRows = null;
    reparse();
  } catch (e) {
    $('fileError').innerHTML = noticeHtml('error',
      '파일을 읽는 중에 문제가 생겼어요. 파일이 엑셀에서 열려 있으면 닫고 다시 시도해 보세요. '
      + `(${e && e.message ? e.message : '알 수 없는 오류'})`);
  }
}

/** 시트나 머리글 줄이 바뀌면 여기서 다시 읽는다. 학생의 선택(plan)은 초기화한다. */
function reparse() {
  const ws = S.workbook.Sheets[S.sheetName];
  const opts = {};
  if (S.headerExcelRows && S.headerExcelRows.length) opts.headerExcelRows = S.headerExcelRows;
  S.parsed = parseSheet(ws, opts);

  // 열이 바뀌었으니 설정을 처음부터 다시 만든다
  S.plan = {};
  S.renames = {};
  S.quiz = null;
  S.quizPicked = {};
  S.quizDone = {};
  S.warnedOnce = {};
  // 열 구성이 달라지면 예전 설정으로 되돌릴 수가 없다 (없는 열을 가리키게 된다)
  S.undoStack = [];
  S.redoStack = [];
  S.excludeTotals = false;      // 기본은 남겨 둔다 (학생이 평균 변화를 관찰하게)
  S.rowFilter = {};

  // ★ 반복 생략 찾기 — 결측치를 세기 전에 해야 한다
  S.omission = detectRepeatedOmission(S.parsed.rows, S.parsed.columns);
  // 확실한 것만 자동으로 켠다. 안 켜면 그 열이 "94% 비어 있는 열" 로 판정되어
  // 앱이 자료를 버리라고 권하게 된다 — 그게 훨씬 큰 사고다.
  // 대신 무엇을 했는지 화면에 크게 알리고 한 번 눌러 되돌릴 수 있게 한다.
  S.fillDown = new Set(S.omission.filter(o => o.confident).map(o => o.name));

  S.tokenCounts = countTokens(fillDownColumns(S.parsed.rows, [...S.fillDown]), S.parsed.columns);
  recompute();
  // 2단계 기본 선택 = 추천 열
  S.picked = S.parsed.columns.filter(c => isRecommendedColumn(S.stats[c.name])).map(c => c.name);
  if (!S.picked.length) S.picked = S.parsed.columns.map(c => c.name);
  applyRecommendedPlan();
  render();
}

/**
 * 열 통계와 결과를 다시 계산한다. 화면을 그리기 전에 항상 호출한다.
 *
 * 결과를 두 가지로 만드는 이유:
 *   2단계에서는 아직 "쓸 열 고르기"를 하지 않았다. 그런데 3단계 선택을 2단계 통계에 섞으면
 *   학생이 아무것도 안 했는데 "열이 1개 줄었어요" 같은 문구가 떠서 혼란스럽다.
 *   → S.resultAll : 모든 열 기준  (2단계 요약 카드·미리보기용)
 *     S.result    : 고른 열 기준  (3·4단계용)
 */
/**
 * 열 통계와 결과를 다시 계산한다. 화면을 그리기 전에 항상 호출한다.
 *
 * 계산 순서 (이 순서가 중요하다)
 *   ① 반복 생략 되살리기 — 비어 보이지만 위 값과 같은 칸을 먼저 채운다.
 *      ★ 결측치를 세기 **전에** 해야 한다. 안 그러면 KOSIS 파일의 성씨 열이
 *        "94% 비어 있는 쓸모없는 열" 로 판정되어 자료를 통째로 잃는다.
 *   ② 행 걸러내기 (2단계에서 고른 값만)
 *   ③ 합계 줄 빼기 (학생이 켰을 때)
 *   ④ 열 통계 — 위 세 가지를 마친 자료 기준. 그래서 화면의 숫자가 "실제로 저장될 자료" 를 말한다
 *   ⑤ 빈 칸 처리 적용
 *
 * 결과를 두 가지로 만든다:
 *   S.resultAll : 모든 열 기준 (3단계 요약 카드·미리보기용 — 열 선택과 섞이지 않게)
 *   S.result    : 고른 열 기준 (4단계 저장용)
 */
function recompute() {
  if (!S.parsed) { S.stats = {}; S.result = null; S.resultAll = null; S.base = null; return; }

  // ① 반복 생략 되살리기
  const filled = fillDownColumns(S.parsed.rows, [...S.fillDown]);
  S.filledRows = filled;

  // 행 고르기 목록에 쓸 통계는 **걸러내기 전** 자료로 낸다.
  // ★ 왜: 걸러낸 결과로 목록을 만들면, 행을 0개로 걸러낸 순간 고를 값 목록이 사라져
  //   되돌릴 수가 없다. 고를 수 있는 값은 내가 무엇을 골랐는지와 무관해야 한다.
  S.statsUnfiltered = analyzeAll(filled, S.parsed.columns, S.tokens);

  // ② 행 걸러내기 — 원본 행 번호(엑셀 몇 행이었나)를 함께 들고 간다
  const rf = applyRowFilter(filled, S.rowFilter);
  const baseRows = [];
  const baseMeta = [];
  const baseSrcIdx = [];      // S.parsed.rows 의 몇 번째였나
  rf.keep.forEach((k, i) => {
    if (!k) return;
    baseRows.push(filled[i]);
    baseMeta.push(S.parsed.rowMeta[i]);
    baseSrcIdx.push(i);
  });

  // ③ 합계 줄은 걸러낸 뒤의 자료에서 다시 찾는다 (번호가 맞아야 한다)
  S.totalRows = detectTotalRows(baseRows, S.parsed.columns);
  const skip = S.excludeTotals ? new Set(S.totalRows.idxs) : null;

  S.base = { rows: baseRows, columns: S.parsed.columns, rowMeta: baseMeta, srcIdx: baseSrcIdx };

  // ④ 통계 — 합계 줄을 뺀 상태 기준 (그래서 합계 줄 토글이 평균을 눈에 보이게 바꾼다)
  const statRows = skip ? baseRows.filter((_, i) => !skip.has(i)) : baseRows;
  S.stats = analyzeAll(statRows, S.parsed.columns, S.tokens);

  // ⑤ 빈 칸 처리
  const allNames = S.parsed.columns.map(c => c.name);
  S.resultAll = applyPlan(S.base, S.plan, S.stats, S.tokens, allNames, skip);
  S.result = applyPlan(S.base, S.plan, S.stats, S.tokens, S.picked.length ? S.picked : allNames, skip);
}

/** 추천 처리 방법을 한 번에 적용한다 */
function applyRecommendedPlan() {
  for (const c of S.parsed.columns) {
    const st = S.stats[c.name];
    const rec = recommend(st);
    const opts = optionsFor(st);
    // 추천한 방법이 이 열의 선택지에 없으면 「그대로 두기」로 둔다
    const ok = opts.some(o => o.action === rec.action);
    S.plan[c.name] = { action: ok ? rec.action : ACTIONS.KEEP, value: undefined };
  }
  recompute();
}

/* ═══════════════════════════════════════════════════════════════════════════
   연습용 예시 파일 — 병합 셀·빈 칸·주석 행이 모두 들어 있다
   ═══════════════════════════════════════════════════════════════════════════ */

function makeSampleWorkbook() {
  const aoa = [
    ['2024년 지역별 공공도서관 현황', '', '', '', '', ''],
    ['시도', '도서관명', '자료수', '', '좌석수', '홈페이지'],
    ['', '', '도서', '전자자료', '', ''],
    ['서울특별시', '가온도서관', 52300, 4100, 320, 'https://example.kr/1'],
    ['', '나래도서관', 11900, 900, null, ''],
    ['', '다솜도서관', 33500, 2100, 210, ''],
    ['', '라온도서관', null, 1500, 150, 'https://example.kr/4'],
    ['부산광역시', '마루도서관', 28800, 1800, 190, ''],
    ['', '바다도서관', 15200, '-', 120, ''],
    ['', '새롬도서관', 41000, 3200, null, ''],
    ['대구광역시', '아름도서관', 22400, 1200, 160, ''],
    ['', '자람도서관', 9800, 'N/A', 90, ''],
    ['', '차오름도서관', null, 800, 110, ''],
    ['인천광역시', '한빛도서관', 31200, 2400, 200, 'https://example.kr/11'],
    ['', '해솔도서관', 18700, 1100, null, ''],
    ['', '', '', '', '', ''],
    ['※ 자료: 예시로 만든 연습용 자료입니다', '', '', '', '', ''],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },   // 제목 가로 병합
    { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } },   // 시도 (세로)
    { s: { r: 1, c: 1 }, e: { r: 2, c: 1 } },   // 도서관명 (세로)
    { s: { r: 1, c: 2 }, e: { r: 1, c: 3 } },   // 자료수 (가로) → 아래에 도서/전자자료
    { s: { r: 1, c: 4 }, e: { r: 2, c: 4 } },   // 좌석수 (세로)
    { s: { r: 1, c: 5 }, e: { r: 2, c: 5 } },   // 홈페이지 (세로)
    { s: { r: 3, c: 0 }, e: { r: 6, c: 0 } },   // 서울특별시 4행 세로 병합
    { s: { r: 7, c: 0 }, e: { r: 9, c: 0 } },   // 부산광역시 3행
    { s: { r: 10, c: 0 }, e: { r: 12, c: 0 } }, // 대구광역시 3행
    { s: { r: 13, c: 0 }, e: { r: 14, c: 0 } }, // 인천광역시 2행
  ];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '2024년 현황');
  return wb;
}

function loadSample() {
  S.rawFile = null;
  S.fileName = '연습용_공공도서관현황.xlsx';
  S.fileSize = 0;
  S.workbook = makeSampleWorkbook();
  S.sheetNames = S.workbook.SheetNames;
  S.sheetName = S.sheetNames[0];
  S.encoding = null;
  S.ioNotices = [{ level: 'info', type: '예시',
    message: '연습용 예시 파일을 열었습니다. 제목 줄, 두 줄 머리글, 병합된 셀, 빈 칸, 아래쪽 설명 줄이 모두 들어 있어요.' }];
  S.headerExcelRows = null;
  S.forcedEncoding = null;
  reparse();
}

/* ═══════════════════════════════════════════════════════════════════════════
   화면 그리기 — 1단계
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 「비어 보이지만 빈 칸이 아닌 칸」 카드.
 *
 * KOSIS 같은 통계표는 같은 값이 반복되면 아래 칸을 비워 둔다.
 * 그걸 빈 칸으로 세면 그 열이 "거의 다 비어 있는 쓸모없는 열" 이 되어 자료를 통째로 잃는다.
 * 그래서 확실한 경우는 미리 채워 두고, **무엇을 했는지 크게 알리고 한 번에 되돌릴 수 있게** 한다.
 */
function renderOmissionCard() {
  const card = $('cardOmission');
  if (!S.omission.length) { card.hidden = true; return; }
  card.hidden = false;

  const box = $('omissionBox');
  box.innerHTML = '';
  for (const o of S.omission) {
    const on = S.fillDown.has(o.name);
    const wrap = document.createElement('div');
    wrap.style.marginBottom = '14px';
    wrap.innerHTML = noticeHtml(on ? 'ok' : 'warn',
      o.reason
      + (on
        ? ` 그래서 빈 칸 ${fmtNum(o.blanks)}개를 바로 위 값으로 채워 두었습니다.`
        : ` 지금은 채우지 않았습니다. 이대로 두면 이 열의 빈 칸 ${fmtNum(o.blanks)}개가 결측치로 세어집니다.`),
      `「${o.name}」 열 (원본 ${o.letter}열) — 빈 칸 ${fmtNum(o.blanks)}개 중 ${fmtNum(o.groups)}개 묶음`);

    const row = document.createElement('div');
    row.className = 'btn-row';
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'btn btn-sm' + (on ? '' : ' btn-primary');
    b.textContent = on ? '↺ 채우지 않기 (빈 칸으로 두기)' : '위 값으로 채우기';
    b.onclick = () => {
      pushUndo(`${o.name}: ${on ? '채우기 끄기' : '위 값으로 채우기'}`);
      if (on) S.fillDown.delete(o.name); else S.fillDown.add(o.name);
      S.tokenCounts = countTokens(fillDownColumns(S.parsed.rows, [...S.fillDown]), S.parsed.columns);
      recompute();
      render();
      toast(on ? `「${o.name}」 을 빈 칸으로 두었어요.` : `「${o.name}」 을 위 값으로 채웠어요.`,
            '되돌리기', doUndo);
    };
    row.appendChild(b);
    wrap.appendChild(row);
    box.appendChild(wrap);
  }

  const help = document.createElement('p');
  help.className = 'muted';
  help.innerHTML = '통계표는 같은 값이 이어질 때 아래 칸을 비워 두는 일이 많아요. '
    + '<b>그 칸은 빈 칸이 아니라 「위와 같다」는 뜻</b>입니다. '
    + '빈 칸으로 세면 그 열이 거의 다 비어 있는 것처럼 보여서 자료를 잃게 됩니다. '
    + '미리보기 표에서 채워진 값을 확인해 보세요.';
  box.appendChild(help);
}

function renderStep1() {
  const loaded = !!S.parsed;
  $('dropWrap').hidden = loaded;
  $('cardLoaded').hidden = !loaded;
  $('cardColNames').hidden = !loaded;
  $('cardPreview1').hidden = !loaded;
  if (!loaded) { $('cardOmission').hidden = true; return; }
  renderOmissionCard();

  const st = S.parsed.stats;
  $('loadedTitle').textContent =
    `✔ ${S.fileName} 을(를) 읽었습니다. 모두 ${fmtNum(st.rowCount)}행 × ${st.colCount}열이에요.`;

  // ── 시트 고르기
  const sp = $('sheetPick');
  if (S.sheetNames.length > 1) {
    sp.innerHTML = `<h3>시트(표) 고르기</h3>`
      + `<p class="muted">시트가 ${S.sheetNames.length}개 있어요. 쓸 시트를 고르세요.</p>`;
    const row = document.createElement('div');
    row.className = 'btn-row';
    S.sheetNames.forEach((n) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'btn btn-sm' + (n === S.sheetName ? ' btn-primary' : '');
      b.textContent = n;
      b.onclick = () => { S.sheetName = n; S.headerExcelRows = null; reparse(); };
      row.appendChild(b);
    });
    sp.appendChild(row);
  } else {
    sp.innerHTML = `<p class="muted">시트가 하나뿐이라 바로 사용합니다. (시트 이름: ${esc(S.sheetName)})</p>`;
  }

  // ── 글자 방식 바꾸기 (한글이 깨져 보일 때)
  const ep = $('encodingPick');
  if (S.encoding) {
    const cur = S.forcedEncoding || S.encoding;
    ep.innerHTML = `<p class="muted">글자 방식: <b>${cur === 'euc-kr' ? 'CP949(엑셀 한글)' : 'UTF-8'}</b>`
      + ` — 한글이 깨져 보이면 아래에서 바꿔 보세요.</p>`;
    const row = document.createElement('div');
    row.className = 'btn-row';
    [['자동', null], ['UTF-8', 'utf-8'], ['CP949(엑셀 한글)', 'euc-kr']].forEach(([label, val]) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'btn btn-sm' + ((S.forcedEncoding || null) === val ? ' btn-primary' : '');
      b.textContent = label;
      b.disabled = !S.rawFile;
      b.onclick = async () => { S.forcedEncoding = val; await openFile(S.rawFile); };
      row.appendChild(b);
    });
    ep.appendChild(row);
  } else {
    ep.innerHTML = '';
  }

  // ── 머리글 줄 고르기
  const hp = $('headerPick');
  const usedRows = st.headerExcelRows || [];
  hp.innerHTML = `<p class="lead">이 파일은 <b>엑셀 ${usedRows.join(', ')}행</b>이 열 이름으로 보입니다. 맞나요?</p>`
    + `<p class="muted">틀렸으면 아래에서 다른 줄을 골라 주세요. 고른 줄 위쪽은 자료에서 빠집니다.</p>`;
  const cands = (S.parsed.headerInfo.candidates || []).slice(0, 5);
  const crow = document.createElement('div');
  crow.className = 'btn-row';
  cands.forEach((cd) => {
    // 후보의 행 번호를 엑셀 행 번호로 바꿔서 보여 준다
    const b = document.createElement('button');
    b.type = 'button';
    const same = JSON.stringify(cd.preview) === JSON.stringify(S.parsed.columns.map(c => c.name));
    b.className = 'btn btn-sm' + (same ? ' btn-primary' : '');
    b.textContent = `${cd.preview.slice(0, 4).join(' / ')}${cd.preview.length > 4 ? ' …' : ''}`;
    b.title = cd.preview.join(' | ');
    b.onclick = () => {
      // 후보의 rowIdxs 는 "빈 행을 지운 뒤" 기준이므로 그대로 넘긴다
      S.headerExcelRows = null;
      S.parsed = parseSheet(S.workbook.Sheets[S.sheetName], { headerRowIdxs: cd.rowIdxs });
      S.plan = {}; S.renames = {}; S.quiz = null; S.quizPicked = {}; S.quizDone = {};
      S.undoStack = []; S.redoStack = [];
      S.excludeTotals = false; S.rowFilter = {};
      S.omission = detectRepeatedOmission(S.parsed.rows, S.parsed.columns);
      S.fillDown = new Set(S.omission.filter(o => o.confident).map(o => o.name));
      S.tokenCounts = countTokens(fillDownColumns(S.parsed.rows, [...S.fillDown]), S.parsed.columns);
      recompute();
      S.picked = S.parsed.columns.filter(c => isRecommendedColumn(S.stats[c.name])).map(c => c.name);
      if (!S.picked.length) S.picked = S.parsed.columns.map(c => c.name);
      applyRecommendedPlan();
      render();
    };
    crow.appendChild(b);
  });
  hp.appendChild(crow);

  // ── 앱이 자동으로 고친 것들 알림
  const nb = $('parseNotices');
  nb.innerHTML = '';
  const all = [...S.ioNotices, ...S.parsed.notices];
  const seen = new Set();
  for (const n of all) {
    const key = n.type + '|' + n.message;
    if (seen.has(key)) continue;
    seen.add(key);
    nb.innerHTML += noticeHtml(n.level === 'error' ? 'error' : n.level, n.message);
  }
  if (st.mergeCount > 0) {
    nb.innerHTML += noticeHtml('ok',
      `이제부터는 칸 위치가 아니라 아래의 열 이름으로 자료를 다룹니다. `
      + `병합돼 있던 칸도 열 이름만으로 값을 읽을 수 있어요.`,
      `병합된 셀 ${st.mergeCount}곳을 풀어서 빈 칸 ${st.filledCellCount}개에 값을 채웠어요.`);
  }
  nb.innerHTML += noticeHtml('info',
    '이 앱은 자료를 저장하지 않습니다. 새로 고치면 파일을 다시 열어야 해요.');

  // ── 만들어진 열 이름 목록
  const cl = $('colNameList');
  cl.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'pick-grid';
  S.parsed.columns.forEach((c) => {
    const stt = S.stats[c.name];
    const box = document.createElement('div');
    box.className = 'pick';
    const kind = { number: '숫자', text: '글자', date: '날짜', empty: '빈열' }[stt.type] || '글자';
    const kindCls = { number: 'badge-num', text: 'badge-text', date: 'badge-date', empty: 'badge-warn' }[stt.type] || 'badge-text';
    box.innerHTML =
      `<div class="pick-top"><span class="pick-name">${esc(c.name)}</span>`
      + `<span class="badge ${kindCls}">${kind}</span></div>`
      + `<div class="pick-stat">원본 ${esc(c.letter)}열`
      + (c.parts && c.parts.length > 1 ? ` · 「${c.parts.map(esc).join('」 + 「')}」 을 합친 이름` : '')
      + (c.renamedFrom ? ` · 이름이 겹쳐서 「${esc(c.renamedFrom)}」 에서 바꿈` : '')
      + (c.autoNamed ? ' · 이름이 비어 있어 임시로 붙인 이름' : '')
      + `</div>`
      + `<div class="pick-sample">예: ${stt.samples.map(v => esc(cellText(v, stt))).join(' / ') || '(값 없음)'}</div>`;
    grid.appendChild(box);
  });
  cl.appendChild(grid);

  // ── 미리보기: 반복 생략을 채운 상태로 보여 준다 (행 걸러내기는 2단계이므로 여기선 전부)
  const shown = fillDownColumns(S.parsed.rows, [...S.fillDown]);
  const fdSet = new Set(S.fillDown);
  const blankNow = (v) => v === null || v === undefined || String(v).trim() === '';
  renderTable($('preview1'), {
    columns: S.parsed.columns,
    rows: shown,
    totalRows: shown.length,
    rowLabels: S.parsed.rowMeta.map(m => m.excelRow),
    // 우리가 채워 넣은 칸에 "채움" 꼬리표를 붙여 학생이 눈으로 구분하게 한다
    filledMask: fdSet.size
      ? S.parsed.rows.map((orig) => {
          const m = {};
          for (const n of fdSet) if (blankNow(orig[n])) m[n] = true;
          return m;
        })
      : null,
    stats: S.stats,
  });
}

/* ═══════════════════════════════════════════════════════════════════════════
   화면 그리기 — 2단계
   ═══════════════════════════════════════════════════════════════════════════ */

function renderStats(box) {
  // 3단계 요약은 "고른 행 + 모든 열" 기준으로 센다 (열 선택과 섞이지 않게)
  const p = S.base, r = S.resultAll;
  // 표기 칩을 켜면 빈 칸이 늘어난다 → 처리 전 결측도 다시 센다
  let missNow0 = 0;
  for (const row of p.rows) for (const c of p.columns) if (isMissing(row[c.name], S.tokens)) missNow0++;
  const cells0 = p.rows.length * p.columns.length;

  const cards = [
    { label: '행(자료)', value: r.rows.length, from: p.rows.length },
    { label: '열(속성)', value: r.columns.length, from: p.columns.length },
    { label: '빈 칸', value: r.missingLeft, from: missNow0 },
    { label: '빈 칸 비율', value: null,
      text: pct(r.rows.length * r.columns.length ? r.missingLeft / (r.rows.length * r.columns.length) : 0),
      fromText: pct(cells0 ? missNow0 / cells0 : 0) },
  ];
  box.innerHTML = cards.map((c) => {
    const now = c.text !== undefined ? c.text : fmtNum(c.value);
    const before = c.fromText !== undefined ? c.fromText : fmtNum(c.from);
    let delta = '';
    if (c.value !== null && c.from !== undefined && c.value !== c.from) {
      const diff = c.from - c.value;
      const risky = c.label === '행(자료)' && c.from > 0 && diff / c.from > 0.3;
      delta = `<div class="stat-delta ${risky ? 'risk' : 'good'}">`
            + `▼ ${fmtNum(diff)}개 ${c.label === '빈 칸' ? '정리했어요' : '줄었어요'}</div>`;
    } else if (before !== now) {
      delta = `<div class="stat-delta good">처음 ${before}</div>`;
    } else {
      delta = `<div class="stat-delta">처음과 같아요 (${before})</div>`;
    }
    return `<div class="stat"><div class="stat-label">${c.label}</div>`
         + `<div class="stat-value">${now}</div>${delta}</div>`;
  }).join('');
}

function renderTokenChips() {
  const box = $('tokenChips');
  box.innerHTML = '';
  // 항상 켜져 있는 "정말 빈 칸" 은 안내만 한다
  const fixed = document.createElement('span');
  fixed.className = 'chip';
  fixed.setAttribute('aria-pressed', 'true');
  fixed.style.cursor = 'default';
  fixed.textContent = '빈 칸 (언제나)';
  box.appendChild(fixed);

  for (const g of MISSING_TOKEN_GROUPS) {
    const n = S.tokenCounts[g.id] || 0;
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'chip' + (g.risky ? ' chip-risky' : '');
    b.setAttribute('aria-pressed', S.tokens.has(g.id) ? 'true' : 'false');
    b.innerHTML = `${esc(g.label)} <span class="chip-count">(${fmtNum(n)}개)</span>`;
    b.title = g.help;
    if (n === 0) b.style.opacity = '.5';
    b.onclick = async () => {
      if (!S.tokens.has(g.id) && g.risky) {
        const ok = await ask('정말 켤까요?',
          `<p>숫자 0도 빈 칸으로 처리하면 「좌석수 0석」 같은 <b>진짜 값이 사라질 수 있어요.</b></p>`
          + `<p class="muted">이 자료에는 0이 ${fmtNum(n)}개 있습니다.</p>`, '그래도 켜기');
        if (!ok) return;
      }
      pushUndo(`${g.label} 을 빈 칸으로 ${S.tokens.has(g.id) ? '보지 않기' : '보기'}`);
      if (S.tokens.has(g.id)) S.tokens.delete(g.id); else S.tokens.add(g.id);
      // 표기가 바뀌면 추천도 달라질 수 있으므로 통계만 다시 내고 설정은 유지한다
      recompute();
      render();
      toast(`「${g.label}」 을 ${S.tokens.has(g.id) ? '빈 칸으로 봅니다' : '값으로 봅니다'}.`);
    };
    box.appendChild(b);
  }

  // 켜지지 않았는데 한 번이라도 나온 표기는 물어본다.
  // ★ 자동으로 바꾸지 않는다 — 되돌릴 수 없는 오해를 막고, 판단 자체를 학습 내용으로 남긴다.
  const hints = MISSING_TOKEN_GROUPS
    .filter(g => !g.risky && !S.tokens.has(g.id) && (S.tokenCounts[g.id] || 0) >= 1);
  $('tokenHelp').innerHTML = hints.length
    ? noticeHtml('warn',
        hints.map(g => `「${g.label}」 ${fmtNum(S.tokenCounts[g.id])}번`).join(', ')
        + ' — 공공데이터에서는 조사하지 못한 칸을 이렇게 적는 경우가 많아요. '
        + '원본 자료의 설명을 확인하고, 빈 칸으로 볼지 위에서 직접 정해 주세요.',
        '빈 칸을 뜻하는 표시가 있을 수 있어요.')
    : '';
}

function renderColList() {
  const box = $('colList');
  box.innerHTML = '';
  // 2단계에서 고른 열만 보여 준다. 쓰지 않을 열의 빈 칸을 고민하게 만들 이유가 없다.
  const shown = S.picked.length
    ? S.picked.map(n => S.parsed.columns.find(c => c.name === n)).filter(Boolean)
    : S.parsed.columns;
  if (!shown.length) {
    box.innerHTML = '<p class="muted">고른 열이 없어요. 2단계에서 쓸 열을 먼저 고르세요.</p>';
    return;
  }
  for (const c of shown) {
    const st = S.stats[c.name];
    const rec = recommend(st);
    const choice = S.plan[c.name] || { action: ACTIONS.KEEP };
    const opts = optionsFor(st);
    const isRec = choice.action === rec.action;

    const row = document.createElement('div');
    row.className = 'col-row' + (st.missing === 0 ? ' clean' : (st.missingRate >= 0.5 ? ' risky' : ''));

    // ① 열 이름
    const c1 = document.createElement('div');
    const kind = { number: '숫자', text: '글자', date: '날짜', empty: '빈열' }[st.type] || '글자';
    c1.innerHTML = `<div class="col-name">${esc(c.name)}</div>`
      + `<div class="col-letter">원본 ${esc(c.letter)}열 · ${kind}`
      + (st.isId ? ' · 이름표 열' : '') + (st.isPrivacy ? ' · 개인정보 주의' : '') + '</div>';
    row.appendChild(c1);

    // ② 종류 배지
    const c2 = document.createElement('div');
    const badge = rec.badge
      ? `<span class="badge ${st.missing === 0 ? 'badge-ok' : 'badge-warn'}">${esc(rec.badge)}</span>`
      : (isRec ? '<span class="badge badge-ok">추천</span>' : '<span class="badge badge-num">내가 고름</span>');
    c2.innerHTML = badge;
    row.appendChild(c2);

    // ③ 빈 칸 개수와 막대
    const c3 = document.createElement('div');
    c3.innerHTML = `<div class="miss-count">${fmtNum(st.missing)}개</div>`
      + `<div class="muted">${pct(st.missingRate)}</div>`
      + `<div class="miss-bar"><i class="${st.missingRate >= 0.5 ? 'high' : ''}" `
      + `style="width:${Math.max(2, Math.round(st.missingRate * 100))}%"></i></div>`;
    row.appendChild(c3);

    // ④ 처리 방법 고르기
    const c4 = document.createElement('div');
    if (st.missing === 0) {
      c4.innerHTML = '<div class="muted">손댈 것이 없어요.</div>';
    } else {
      const sel = document.createElement('select');
      sel.className = 'fill-select';
      sel.setAttribute('aria-label', `${c.name} 열의 빈 칸 처리 방법`);
      opts.forEach((o) => {
        const op = document.createElement('option');
        op.value = o.action;
        op.textContent = o.label + (o.action === rec.action ? '  ← 추천' : '');
        if (o.action === choice.action) op.selected = true;
        sel.appendChild(op);
      });
      c4.appendChild(sel);

      // 직접 입력 칸
      const inp = document.createElement('input');
      inp.className = 'fill-input';
      inp.type = 'text';
      inp.placeholder = st.type === 'number' ? '넣을 숫자를 적으세요' : '넣을 값을 적으세요';
      inp.value = choice.value === undefined ? '' : choice.value;
      inp.hidden = choice.action !== ACTIONS.CUSTOM;
      inp.style.marginTop = '8px';
      inp.setAttribute('aria-label', `${c.name} 열에 채울 값`);
      c4.appendChild(inp);

      sel.onchange = () => {
        pushUndo(`${c.name}: ${sel.options[sel.selectedIndex].textContent.replace(/\s*←.*$/, '').trim()}`);
        S.plan[c.name] = { action: sel.value, value: sel.value === ACTIONS.CUSTOM ? inp.value : undefined };
        checkAndRender(c.name);
      };
      inp.oninput = () => {
        S.plan[c.name] = { action: ACTIONS.CUSTOM, value: inp.value };
        recompute();
        renderStats($('stats3'));
        renderPreview3();
        renderNoteFor(c.name, note);
      };

      // 추천 이유 + 경고
      const note = document.createElement('div');
      c4.appendChild(note);
      renderNoteFor(c.name, note);
    }
    row.appendChild(c4);
    box.appendChild(row);
  }
}

/** 열 하나에 대한 안내·경고 문구를 다시 그린다 */
function renderNoteFor(name, note) {
  const st = S.stats[name];
  const rec = recommend(st);
  const choice = S.plan[name] || { action: ACTIONS.KEEP };
  let html = `<div class="fill-note">💡 ${esc(rec.why)}</div>`;

  if (choice.action === ACTIONS.CUSTOM) {
    if (choice.value === undefined || choice.value === '') {
      html += `<div class="fill-note warn">채울 값을 적어 주세요. 적기 전에는 빈 칸이 그대로 남습니다.</div>`;
    } else if (st.type === 'number' && toNumber(choice.value) === null) {
      html += `<div class="fill-note warn">이 열은 숫자 열이에요. 「${esc(choice.value)}」 은 숫자가 아닙니다. `
            + `숫자로 적거나 다른 방법을 골라 보세요.</div>`;
    }
  }
  if ([ACTIONS.MEAN, ACTIONS.MEDIAN].includes(choice.action) && st.decimals === 0 && st.mean !== null) {
    const raw = choice.action === ACTIONS.MEAN ? st.mean : st.median;
    if (Math.abs(raw - Math.round(raw)) > 1e-9) {
      html += `<div class="fill-note">원본이 정수뿐이라 ${fmtNum(roundLike(raw, 1))} → `
            + `${fmtNum(roundLike(raw, 0))} 으로 반올림해서 넣습니다.</div>`;
    }
  }
  for (const w of warnFor(st, choice)) {
    html += `<div class="fill-note warn">⚠ ${esc(w.text)}</div>`;
  }
  note.innerHTML = html;
}

/** 처리 방법을 바꿨을 때: 위험한 결과면 한 번 확인하고, 화면을 다시 그린다 */
async function checkAndRender(changedName) {
  recompute();
  render();

  // 행이 너무 많이 사라지면 알려 준다 (막지는 않는다)
  const before = S.parsed.rows.length;
  const after = S.resultAll.rows.length;
  if (after === 0 && before > 0 && !S.warnedOnce.zeroRows) {
    S.warnedOnce.zeroRows = true;
    await ask('남는 자료가 없어요',
      `<p>지금 설정으로는 표가 <b>완전히 비어 버립니다.</b></p>`
      + `<p class="muted">「행 지우기」를 고른 열이 많으면 지워지는 행이 계속 합쳐집니다. `
      + `한두 열에만 쓰고 나머지는 다른 방법을 골라 보세요.</p>`, '알겠어요', '닫기');
  } else if (before > 0 && (before - after) / before > 0.3 && !S.warnedOnce.manyRows) {
    S.warnedOnce.manyRows = true;
    toast(`자료가 ${fmtNum(before)}행 → ${fmtNum(after)}행으로 줄었어요. 결과를 믿기 어려울 수 있어요.`);
  }
}

function renderPreview3() {
  // 3단계 미리보기는 "고른 행" 안에서 모든 열을 보여 준다 (열은 2단계에서 이미 골랐지만,
  // 빈 칸 처리를 정할 때는 빠진 열도 눈으로 확인할 수 있어야 판단이 된다)
  const b = S.base, r = S.resultAll;
  if (S.previewMode === 'before') {
    // 정리 전: 처리 전 자료를 보여 주고, 지워질 행에 표시를 한다
    const keep = new Set(r.keptRowIndex);
    renderTable($('preview3'), {
      columns: b.columns,
      rows: b.rows,
      totalRows: b.rows.length,
      rowLabels: b.rowMeta.map(m => m.excelRow),
      dropMask: b.rows.map((_, i) => !keep.has(i)),
      stats: S.stats,
    });
  } else {
    renderTable($('preview3'), {
      columns: b.columns,
      rows: r.rows,
      totalRows: r.rows.length,
      rowLabels: r.keptRowIndex.map(i => b.rowMeta[i].excelRow),
      filledMask: r.filledMask,
      stats: S.stats,
    });
  }
}

/**
 * 합계·소계 행 카드.
 * 여기가 수업의 하이라이트다 — 합계 줄을 넣었을 때와 뺐을 때 평균이 달라지는 것을
 * 학생이 직접 눌러 보게 한다. 그래서 기본값은 "남겨 둠"이고, 앱이 몰래 지우지 않는다.
 */
function renderTotalsCard() {
  const card = $('cardTotals');
  const n = S.totalRows.idxs.length;
  card.hidden = n === 0;
  if (!n) return;

  // 합계 줄을 넣었을 때와 뺐을 때 어떤 숫자 열의 평균이 얼마나 달라지는지 한 줄로 보여 준다
  const numCol = S.parsed.columns.find(c => S.stats[c.name] && S.stats[c.name].type === 'number');
  let compare = '';
  if (numCol) {
    const name = numCol.name;
    const meanOf = (rows) => {
      const ns = rows.map(r => toNumber(r[name])).filter(v => v !== null);
      return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : null;
    };
    const skip = new Set(S.totalRows.idxs);
    const withTotal = meanOf(S.parsed.rows);
    const without = meanOf(S.parsed.rows.filter((_, i) => !skip.has(i)));
    if (withTotal !== null && without !== null && Math.abs(withTotal - without) > 1e-9) {
      compare = `<p><b>「${esc(name)}」 열의 평균</b> — `
        + `합계 줄을 넣으면 <b>${fmtNum(roundLike(withTotal, 1))}</b>, `
        + `빼면 <b>${fmtNum(roundLike(without, 1))}</b> 입니다. `
        + `합계 줄 하나가 평균을 이만큼 바꿉니다.</p>`;
    }
  }

  // 라벨이 많을 수 있으니(예: 「전국」이 155번) 묶어서 개수로 보여 준다
  const byLabel = new Map();
  S.totalRows.labels.forEach((l) => byLabel.set(l, (byLabel.get(l) || 0) + 1));
  const rowsList = [...byLabel.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .map(([l, k]) => `${esc(l)}${k > 1 ? ` ${fmtNum(k)}줄` : ''}`)
    .join(', ') + (byLabel.size > 8 ? ` 등 ${byLabel.size}가지` : '');

  card.querySelector('#totalsBox').innerHTML =
    noticeHtml('warn',
      `찾은 줄: ${rowsList}. 합계 줄은 아래 자료들을 이미 다 더한 값이에요. `
      + '그대로 두면 같은 자료를 두 번 세는 셈이 되어 평균과 그래프가 틀어집니다.',
      `합계로 보이는 줄 ${n}개를 찾았어요.`)
    + compare
    + '<div class="btn-row" style="margin-top:12px">'
    + `<button type="button" class="btn btn-sm ${S.excludeTotals ? 'btn-primary' : ''}" id="btnToggleTotals">`
    + (S.excludeTotals ? '✔ 합계 줄을 뺐어요 (다시 넣기)' : '합계 줄 빼기') + '</button>'
    + '</div>'
    + '<p class="muted" style="margin-top:10px">'
    + '앱이 알아서 지우지 않습니다. 넣었을 때와 뺐을 때 위의 숫자가 어떻게 달라지는지 직접 눌러 보세요.</p>';

  $('btnToggleTotals').onclick = () => {
    pushUndo(S.excludeTotals ? '합계 줄 다시 넣기' : '합계 줄 빼기');
    S.excludeTotals = !S.excludeTotals;
    recompute();
    render();
    toast(S.excludeTotals ? '합계 줄을 뺐어요. 평균이 어떻게 바뀌었는지 보세요.' : '합계 줄을 다시 넣었어요.',
          '되돌리기', doUndo);
  };
}

/* ═══════════════════════════════════════════════════════════════════════════
   화면 그리기 — 3단계 (빈 칸 정리)
   ═══════════════════════════════════════════════════════════════════════════ */

function renderStep3() {
  renderStats($('stats3'));
  renderTokenChips();
  renderTotalsCard();
  renderColList();
  renderPreview3();
  // 고른 자료에 빈 칸이 아예 없으면 그렇다고 알려 준다 (학생이 뭘 해야 할지 헤매지 않게)
  const noMissing = S.parsed.columns.every(c => S.stats[c.name].missing === 0);
  $('cardNoMissing').hidden = !noMissing;
}

/* ═══════════════════════════════════════════════════════════════════════════
   화면 그리기 — 2단계 (행·열 고르기)
   ═══════════════════════════════════════════════════════════════════════════ */

function renderStep2() {
  const cols = S.parsed.columns;
  $('pickCount').textContent = `고른 열: ${S.picked.length} / ${cols.length}개`;
  $('pickCount').style.color = S.picked.length === 0 ? 'var(--c-danger)' : '';

  // ── 속성 카드
  const grid = $('pickGrid');
  grid.innerHTML = '';
  for (const c of cols) {
    const st = S.stats[c.name];
    const on = S.picked.includes(c.name);
    const coach = coachColumn(st);
    const kind = { number: '숫자', text: '글자', date: '날짜', empty: '빈열' }[st.type] || '글자';
    const kindCls = { number: 'badge-num', text: 'badge-text', date: 'badge-date', empty: 'badge-warn' }[st.type] || 'badge-text';

    const label = document.createElement('label');
    label.className = 'pick' + (on ? ' on' : '');
    label.innerHTML =
      `<div class="pick-top">`
      + `<input type="checkbox" ${on ? 'checked' : ''} aria-label="${esc(c.name)} 속성 사용">`
      + `<span class="pick-name">${esc(c.name)}</span>`
      + `<span class="badge ${kindCls}">${kind}</span></div>`
      + `<div class="pick-stat">빈 칸 ${pct(st.missingRate)} · 서로 다른 값 ${fmtNum(st.distinct)}개 · 원본 ${esc(c.letter)}열</div>`
      + `<div class="pick-sample">예: ${st.samples.map(v => esc(cellText(v, st))).join(' / ') || '(값 없음)'}</div>`
      + `<div class="pick-coach ${coach.kind}">${coach.kind === 'ok' ? '✔ ' : (coach.kind === 'info' ? 'ⓘ ' : '⚠ ')}${esc(coach.text)}</div>`;
    label.querySelector('input').onchange = (e) => {
      pushUndo(`${c.name} ${e.target.checked ? '고르기' : '빼기'}`);
      if (e.target.checked) { if (!S.picked.includes(c.name)) S.picked.push(c.name); }
      else S.picked = S.picked.filter(n => n !== c.name);
      recompute();
      render();
    };
    grid.appendChild(label);
  }

  // ── 코칭
  const numPicked = S.picked.filter(n => S.stats[n].type === 'number').length;
  let coachHtml = '';
  if (S.picked.length === 0) {
    coachHtml = noticeHtml('warn', '쓰고 싶은 열의 네모(☐)를 눌러 보세요.', '아직 고른 속성이 없어요.');
  } else if (S.picked.length === cols.length && cols.length > 3) {
    coachHtml = noticeHtml('info',
      `${cols.length}개를 모두 골랐어요. 정말 다 필요한가요? 알아보려는 것과 관계없는 열은 `
      + '빼는 편이 표를 읽기 쉽게 만듭니다.');
  } else if (S.picked.length === 1) {
    coachHtml = noticeHtml('info',
      '속성이 1개면 「무엇과 무엇을 비교」하기 어려워요. 2개 이상 골라 보면 어떨까요?');
  }
  if (numPicked === 0 && S.picked.length > 0) {
    coachHtml += noticeHtml('info',
      '숫자 열을 하나도 고르지 않았어요. 숫자 열이 있으면 평균·합계·그래프를 만들 수 있습니다.');
  }
  const risky = S.picked.filter(n => S.stats[n].missingRate >= 0.5);
  if (risky.length) {
    coachHtml += noticeHtml('warn',
      `빈 칸이 많은 열(${risky.map(esc).join(', ')})을 골랐어요. 결과를 발표할 때 `
      + '「이 열은 빈 칸이 많았다」는 점을 꼭 함께 말해 주세요.');
  }
  $('pickCoach').innerHTML = coachHtml;

  // ── 골라 둔 순서
  const ol = $('orderList');
  ol.innerHTML = '';
  S.picked.forEach((name, i) => {
    const item = document.createElement('div');
    item.className = 'order-item';
    item.innerHTML = `<span class="order-no">${i + 1}</span>`
      + `<span class="order-name">${esc(name)}</span>`;
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.className = 'text-input';
    inp.placeholder = '새 이름 (비워 두면 그대로)';
    inp.value = S.renames[name] || '';
    inp.setAttribute('aria-label', `${name} 의 새 이름`);
    inp.oninput = () => {
      const v = inp.value.trim();
      if (v) S.renames[name] = v; else delete S.renames[name];
    };
    item.appendChild(inp);
    const up = document.createElement('button');
    up.type = 'button'; up.className = 'icon-btn'; up.textContent = '▲';
    up.title = '위로'; up.disabled = i === 0;
    up.onclick = () => {
      pushUndo(`${name} 순서 올리기`);
      const a = S.picked; [a[i - 1], a[i]] = [a[i], a[i - 1]]; recompute(); render();
    };
    const dn = document.createElement('button');
    dn.type = 'button'; dn.className = 'icon-btn'; dn.textContent = '▼';
    dn.title = '아래로'; dn.disabled = i === S.picked.length - 1;
    dn.onclick = () => {
      pushUndo(`${name} 순서 내리기`);
      const a = S.picked; [a[i + 1], a[i]] = [a[i], a[i + 1]]; recompute(); render();
    };
    item.appendChild(up); item.appendChild(dn);
    ol.appendChild(item);
  });
  if (!S.picked.length) ol.innerHTML = '<p class="muted">고른 열이 없어요.</p>';

  // ── 행 고르기
  renderRowFilter();

  // ── 미리보기 (고른 행·열 기준)
  renderTable($('preview2'), {
    columns: S.result.columns,
    rows: S.result.rows,
    totalRows: S.result.rows.length,
    rowLabels: S.result.keptRowIndex.map(i => S.base.rowMeta[i].excelRow),
    filledMask: S.result.filledMask,
    stats: S.stats,
  });
}

/**
 * 행 고르기 — 값으로 걸러내기.
 * 예) 성씨 155개 중 김·이·박만, 지역 18개 중 서울·부산만.
 * 값이 많은 열은 찾기 칸으로 걸러 보게 한다 (155개를 눈으로 훑을 수는 없다).
 */
function renderRowFilter() {
  const box = $('rowFilterBox');
  // 걸러내기 전 자료로 목록을 만든다 (위 recompute 의 설명 참고)
  const filled = S.filledRows;
  const cands = filterableColumns(filled, S.parsed.columns, S.statsUnfiltered);

  // 현재 상태 문구
  const total = S.parsed.rows.length;
  const now = S.base ? S.base.rows.length : total;
  const active = Object.keys(S.rowFilter).filter(n => S.rowFilter[n] instanceof Set);
  $('rowFilterStatus').innerHTML = active.length
    ? noticeHtml(now === 0 ? 'error' : 'ok',
        now === 0
          ? '지금 조건에 맞는 행이 하나도 없어요. 값을 다시 켜 주세요.'
          : `${active.map(n => `「${n}」`).join(', ')} 조건으로 걸러서 `
            + `전체 ${fmtNum(total)}행 중 ${fmtNum(now)}행이 남았습니다.`,
        now === 0 ? '남는 행이 없어요' : `행 ${fmtNum(now)}개를 골랐어요`)
    : `<p class="muted">지금은 걸러내지 않아서 <b>${fmtNum(total)}행 전부</b>가 들어갑니다.</p>`;

  box.innerHTML = '';
  if (!cands.length) {
    box.innerHTML = '<p class="filter-empty">값으로 골라낼 수 있는 열이 없어요. '
      + '(값 종류가 2개 미만이거나 너무 많은 열은 목록으로 고를 수 없습니다)</p>';
    return;
  }

  for (const c of cands) {
    const counts = valueCounts(filled, c.name);
    const sel = S.rowFilter[c.name];
    const isActive = sel instanceof Set;
    const chosen = isActive ? counts.filter(v => sel.has(v.value)).length : counts.length;

    const det = document.createElement('details');
    det.className = 'filter-col' + (isActive ? ' active' : '');
    if (isActive) det.open = true;

    const sum = document.createElement('summary');
    sum.innerHTML = `${esc(c.name)} `
      + `<span class="badge ${isActive ? 'badge-warn' : 'badge-num'}">`
      + `${fmtNum(chosen)} / ${fmtNum(counts.length)}개 값</span>`
      + (isActive ? '' : ' <span class="muted" style="font-size:var(--fs-200)">전부 넣는 중</span>');
    det.appendChild(sum);

    const body = document.createElement('div');
    body.className = 'filter-body';

    const tools = document.createElement('div');
    tools.className = 'filter-tools';
    const search = document.createElement('input');
    search.type = 'search';
    search.placeholder = `값 찾기 (${fmtNum(counts.length)}개 중에서)`;
    search.setAttribute('aria-label', `${c.name} 값 찾기`);
    tools.appendChild(search);
    const bAll = document.createElement('button');
    bAll.type = 'button'; bAll.className = 'btn btn-sm'; bAll.textContent = '전부 넣기';
    const bNone = document.createElement('button');
    bNone.type = 'button'; bNone.className = 'btn btn-sm'; bNone.textContent = '전부 빼기';
    tools.appendChild(bAll); tools.appendChild(bNone);
    body.appendChild(tools);

    const list = document.createElement('div');
    list.className = 'filter-values';
    body.appendChild(list);

    // 검색어에 맞는 값만 그린다. 값이 많을 수 있어 한 번에 300개까지만.
    const paint = () => {
      const q = search.value.trim().toLowerCase();
      const cur = S.rowFilter[c.name];
      const on = (v) => !(cur instanceof Set) || cur.has(v);
      const shown = (q ? counts.filter(v => String(v.value).toLowerCase().includes(q)) : counts).slice(0, 300);
      list.innerHTML = '';
      if (!shown.length) {
        list.innerHTML = '<p class="filter-empty">찾는 값이 없어요.</p>';
        return;
      }
      for (const v of shown) {
        const lab = document.createElement('label');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = on(v.value);
        cb.onchange = () => {
          pushUndo(`${c.name} 행 고르기`);
          // 처음 끄는 순간, "전부 켜진 상태" 를 실제 집합으로 만들어 둔다
          let set = S.rowFilter[c.name];
          set = (set instanceof Set) ? new Set(set) : new Set(counts.map(x => x.value));
          if (cb.checked) set.add(v.value); else set.delete(v.value);
          // 전부 켜졌으면 걸러내지 않는 상태로 되돌린다 (문구가 깔끔해진다)
          if (set.size === counts.length) delete S.rowFilter[c.name];
          else S.rowFilter[c.name] = set;
          recompute();
          render();
        };
        lab.appendChild(cb);
        const nm = document.createElement('span');
        nm.className = 'fv-name';
        nm.textContent = v.value;
        nm.title = v.value;
        lab.appendChild(nm);
        const cn = document.createElement('span');
        cn.className = 'fv-count';
        cn.textContent = fmtNum(v.n) + '행';
        lab.appendChild(cn);
        list.appendChild(lab);
      }
      if (!q && counts.length > 300) {
        const more = document.createElement('p');
        more.className = 'filter-empty';
        more.textContent = `값이 ${fmtNum(counts.length)}개예요. 앞 300개만 보여 줍니다 — 위의 찾기 칸을 쓰세요.`;
        list.appendChild(more);
      }
    };
    search.oninput = paint;
    bAll.onclick = () => {
      pushUndo(`${c.name} 전부 넣기`);
      delete S.rowFilter[c.name];
      recompute(); render();
    };
    bNone.onclick = () => {
      pushUndo(`${c.name} 전부 빼기`);
      // 빈 집합 = "이 열로 걸러내는데 남길 값이 없다" → 남는 행 0개.
      // (applyRowFilter 는 열 이름이 있으면 걸러낸다. 빈 Set 이 곧 전부 빼기다)
      S.rowFilter[c.name] = new Set();
      recompute(); render();
    };
    paint();

    det.appendChild(body);
    box.appendChild(det);
  }
}

/* ═══════════════════════════════════════════════════════════════════════════
   화면 그리기 — 4단계
   ═══════════════════════════════════════════════════════════════════════════ */

/** 저장할 2차원 배열. 열 이름은 학생이 바꾼 이름을 쓴다. */
function buildAOA() {
  const cols = S.result.columns;
  const head = cols.map(c => S.renames[c.name] || c.name);
  const body = S.result.rows.map(r => cols.map(c => {
    const v = r[c.name];
    return v === null || v === undefined ? '' : v;
  }));
  return [head, ...body];
}

/** 「내가 한 일」 글 목록 */
function buildReport() {
  const lines = [];
  const p = S.parsed;
  lines.push(`「${S.fileName}」 의 「${S.sheetName}」 시트를 열었습니다.`);
  const hr = p.stats.headerExcelRows || [];
  let l2 = `엑셀 ${hr.join(', ')}행을 열 이름으로 정했습니다.`;
  if (p.stats.mergeCount > 0) {
    l2 += ` 병합된 셀 ${p.stats.mergeCount}곳을 풀어서 빈 칸 ${p.stats.filledCellCount}개에 값을 채웠습니다.`;
  }
  lines.push(l2);
  // 반복 생략 되살리기 — 무엇을 왜 채웠는지 반드시 남긴다 (자료를 고친 기록이므로)
  for (const o of S.omission) {
    if (S.fillDown.has(o.name)) {
      lines.push(`「${o.name}」 열은 같은 값이 이어질 때 아래 칸이 생략된 표여서, `
        + `빈 칸 ${fmtNum(o.blanks)}개를 바로 위 값으로 채웠습니다. (원래 값을 되살린 것입니다)`);
    } else {
      lines.push(`「${o.name}」 열은 생략된 칸으로 보였지만 채우지 않고 빈 칸으로 두었습니다.`);
    }
  }
  if (S.tokens.size) {
    const names = MISSING_TOKEN_GROUPS.filter(g => S.tokens.has(g.id)).map(g => g.label);
    lines.push(`「${names.join('」, 「')}」 를 빈 칸으로 보기로 했습니다.`);
  }
  if (S.totalRows.idxs.length) {
    lines.push(S.excludeTotals
      ? `합계로 보이는 줄 ${S.totalRows.idxs.length}개(${S.totalRows.labels.join(', ')})를 뺐습니다.`
      : `합계로 보이는 줄 ${S.totalRows.idxs.length}개(${S.totalRows.labels.join(', ')})는 그대로 두었습니다.`);
  }
  // 열마다 한 일
  for (const c of p.columns) {
    const st = S.stats[c.name];
    const ch = S.plan[c.name];
    if (!ch || st.missing === 0) continue;
    if (ch.action === ACTIONS.KEEP) continue;
    if (ch.action === ACTIONS.DROP) {
      lines.push(`「${c.name}」 의 빈 칸이 있는 행을 지우기로 했습니다. (빈 칸 ${st.missing}개)`);
    } else if (ch.action === ACTIONS.MEAN) {
      lines.push(`「${c.name}」 의 빈 칸 ${st.missing}개를 평균값 ${fmtNum(roundLike(st.mean, st.decimals))} 으로 채웠습니다.`);
    } else if (ch.action === ACTIONS.MEDIAN) {
      lines.push(`「${c.name}」 의 빈 칸 ${st.missing}개를 가운데 값 ${fmtNum(roundLike(st.median, st.decimals))} 으로 채웠습니다.`);
    } else if (ch.action === ACTIONS.TEXT || ch.action === ACTIONS.CUSTOM) {
      const opt = optionsFor(st).find(o => o.action === ch.action);
      const v = ch.value !== undefined && ch.value !== '' ? ch.value : (opt && opt.value);
      if (v !== undefined && v !== '') {
        lines.push(`「${c.name}」 의 빈 칸 ${st.missing}개를 「${v}」 로 채웠습니다.`);
      }
    }
  }
  if (S.result.droppedRows > 0) {
    lines.push(`빈 칸이 남은 행 ${fmtNum(S.result.droppedRows)}개를 지웠습니다.`);
  }
  const out = p.columns.filter(c => !S.picked.includes(c.name)).map(c => c.name);
  if (out.length) lines.push(`쓰지 않을 열 ${out.length}개를 뺐습니다: ${out.join(', ')}`);
  lines.push(`쓸 열 ${S.picked.length}개를 골랐습니다: ${S.picked.map(n => S.renames[n] || n).join(', ')}`);
  // 행 걸러내기
  for (const name of Object.keys(S.rowFilter)) {
    const set = S.rowFilter[name];
    if (!(set instanceof Set)) continue;
    const vals = [...set];
    lines.push(`「${name}」 이(가) ${vals.length ? vals.slice(0, 12).join(', ')
      + (vals.length > 12 ? ` 등 ${vals.length}개` : '') : '(없음)'} 인 행만 골랐습니다.`);
  }
  if (S.base && S.base.rows.length !== S.parsed.rows.length) {
    lines.push(`그 결과 전체 ${fmtNum(S.parsed.rows.length)}행 중 ${fmtNum(S.base.rows.length)}행이 남았습니다.`);
  }
  if (S.goal) lines.push(`알아보고 싶은 것: ${S.goal}`);
  if (S.result.missingLeft > 0) {
    lines.push(`남은 빈 칸 ${fmtNum(S.result.missingLeft)}개는 그대로 두었습니다.`);
  }
  return lines;
}

function renderStep4() {
  const p = S.parsed, r = S.result;
  // 「정리 전」 기준은 "반복 생략을 되살린 원본" 이다.
  // 반복 생략 되살리기는 학생의 정리 작업이 아니라 파일을 제대로 읽은 것이므로,
  // 그걸 "빈 칸 2,614개를 없앴다" 로 세면 한 일을 부풀리게 된다.
  const before = fillDownColumns(p.rows, [...S.fillDown]);
  let miss0 = 0;
  for (const row of before) for (const c of p.columns) if (isMissing(row[c.name], S.tokens)) miss0++;
  const cells0 = before.length * p.columns.length;
  const cells1 = r.rows.length * r.columns.length;

  const rows = [
    ['행(자료)', before.length, r.rows.length],
    ['열(속성)', p.columns.length, r.columns.length],
    ['빈 칸', miss0, r.missingLeft],
  ];
  // 행이 줄어든 이유를 원인별로 나눠 적는다.
  // ★ 그냥 "2,615개 지움" 이라고 쓰면 결측치 때문에 잃은 것처럼 읽힌다.
  //   골라내서 뺀 것과 빈 칸 때문에 지운 것은 뜻이 완전히 다르므로 구별해서 보여 준다.
  const byFilter = before.length - S.base.rows.length;
  const byTotals = r.excludedRows || 0;
  const byMissing = r.droppedRows || 0;
  const rowReason = () => {
    const parts = [];
    if (byFilter) parts.push(`골라내서 뺌 ${fmtNum(byFilter)}`);
    if (byTotals) parts.push(`합계 줄 뺌 ${fmtNum(byTotals)}`);
    if (byMissing) parts.push(`빈 칸 때문에 지움 ${fmtNum(byMissing)}`);
    return parts.length ? ` (${parts.join(' · ')})` : '';
  };

  let html = '<table class="compare"><thead><tr><th>항목</th><th>정리 전</th><th>정리 후</th><th>변화</th></tr></thead><tbody>';
  for (const [label, a, b] of rows) {
    const diff = a - b;
    // "위험한 변화" 는 **빈 칸 때문에 자료를 잃은 경우**만이다.
    // 내가 일부러 골라낸 것은 위험이 아니라 의도한 일이다.
    const risky = label === '행(자료)' && a > 0 && byMissing / a > 0.3;
    let chg;
    if (diff === 0) chg = '그대로';
    else if (label === '행(자료)') chg = `▼ ${fmtNum(diff)}개${rowReason()}`;
    else if (label === '열(속성)') chg = `▼ ${fmtNum(diff)}개 뺌`;
    else chg = `▼ ${fmtNum(diff)}개 없앰`;
    html += `<tr><th>${label}</th><td class="num">${fmtNum(a)}</td><td class="num">${fmtNum(b)}</td>`
          + `<td class="change ${diff === 0 ? '' : (risky ? 'risk' : 'good')}">${esc(chg)}</td></tr>`;
  }
  const r0 = cells0 ? miss0 / cells0 : 0, r1 = cells1 ? r.missingLeft / cells1 : 0;
  html += `<tr><th>빈 칸 비율</th><td class="num">${pct(r0)}</td><td class="num">${pct(r1)}</td>`
        + `<td class="change ${r0 === r1 ? '' : 'good'}">${r0 === r1 ? '그대로' : '▼ ' + Math.round((r0 - r1) * 100) + '%p'}</td></tr>`;
  html += '</tbody></table>';
  $('compareTable').innerHTML = html;

  // 내가 한 일
  const rep = buildReport();
  $('reportList').innerHTML = rep.map(l => `<li>${esc(l)}</li>`).join('');

  // 저장 파일 이름 기본값
  if (!$('saveName').value) {
    const base = (S.fileName || '자료').replace(/\.(xlsx|xls|csv|txt)$/i, '');
    $('saveName').value = base + '_정리';
  }

  renderQuiz();
}

/* ═══════════════════════════════════════════════════════════════════════════
   확인 문제 — 학생이 방금 다룬 자기 자료에서 문제를 만든다
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 보기 순서를 섞는다.
 * ★ 왜 필요한가: 섞지 않으면 정답이 늘 첫 번째 보기가 되어 학생이 바로 눈치챈다.
 *   문제는 한 번 만들어 두고 계속 쓰므로(S.quiz), 여기서 한 번만 섞으면 순서가 흔들리지 않는다.
 */
function shuffled(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function buildQuiz() {
  const p = S.parsed;
  const byMiss = p.columns
    .map(c => ({ name: c.name, m: S.stats[c.name].missing }))
    .sort((a, b) => b.m - a.m);
  const top = byMiss[0];

  // 1번: 빈 칸이 가장 많았던 열.
  //   빈 칸이 아예 없는 자료에서는 물어볼 것이 없으므로 이 문항을 빼고, 열 개수가 부족해도 뺀다.
  //   1등이 여러 열과 같은 개수면 정답이 하나가 아니게 되므로 그때도 뺀다.
  const tiedTop = byMiss.filter(x => x.m === top.m).length;
  const q1 = (top.m > 0 && byMiss.length >= 3 && tiedTop === 1) ? {
    q: '정리하기 전, 빈 칸이 가장 많았던 열은 무엇인가요?',
    options: shuffled([top.name, ...byMiss.slice(1, 4).map(x => x.name)]),
    answer: top.name,
    hint: '3단계 표의 「빈 칸」 칸을 다시 보세요. 막대가 가장 긴 열이에요.',
    why: `「${top.name}」 열에 빈 칸이 ${fmtNum(top.m)}개 있었어요.`,
  } : null;

  const q2 = {
    q: '빈 칸이 있는 행을 지우면 자료(행) 수는 어떻게 되나요?',
    options: shuffled(['늘어난다', '줄어든다', '그대로다']),
    answer: '줄어든다',
    hint: '행을 지웠으니 자료 개수가 어떻게 될지 생각해 보세요.',
    why: '행을 빼면 자료 수가 줄어듭니다. 그래서 너무 많이 지우면 결과를 믿기 어려워져요.',
  };
  const q3 = {
    q: '값이 모두 같은 열을 빼는 이유는 무엇인가요?',
    options: shuffled([
      '비교할 수 있는 차이가 없어서',
      '파일 크기가 커져서',
      '빈 칸이 많아서',
      '엑셀에서 열 수 없어서',
    ]),
    answer: '비교할 수 있는 차이가 없어서',
    hint: '모두 「서울」이라면 지역별로 비교가 될까요?',
    why: '모든 값이 같으면 무엇과 무엇을 비교할 수가 없습니다.',
  };
  const q4 = {
    q: '빈 칸을 그 열의 평균값으로 채우면 어떤 일이 생기나요?',
    options: shuffled([
      '자료 수는 그대로지만, 채운 칸이 모두 똑같은 값이 된다',
      '자료 수가 줄어든다',
      '원래 값들이 모두 바뀐다',
      '아무 것도 달라지지 않는다',
    ]),
    answer: '자료 수는 그대로지만, 채운 칸이 모두 똑같은 값이 된다',
    hint: '평균은 하나의 숫자예요. 빈 칸 여러 개에 그 하나를 넣으면 어떻게 될까요?',
    why: '행을 지우지 않으니 자료 수는 유지되지만, 채운 칸이 전부 같은 값이 되어 자료가 실제보다 고르게 보입니다.',
  };
  // 반복 생략을 채운 자료라면 그 개념을 묻는다 — 이 자료에서 가장 중요한 학습 지점이다
  const om = S.omission.filter(o => S.fillDown.has(o.name))[0];
  const q5 = om ? {
    q: `「${om.name}」 열은 아래 칸이 비어 있었는데 위 값으로 채웠어요. 왜 그렇게 했을까요?`,
    options: shuffled([
      '같은 값이 이어질 때 아래 칸을 생략해서 적은 표라서',
      '값을 조사하지 못해서 비어 있었기 때문에',
      '빈 칸을 없애면 파일이 작아지기 때문에',
      '평균을 구하기 쉽게 만들려고',
    ]),
    answer: '같은 값이 이어질 때 아래 칸을 생략해서 적은 표라서',
    hint: '원본 표에서 그 칸들이 정말 「값이 없는」 칸이었는지 생각해 보세요.',
    why: '통계표는 같은 값이 반복될 때 아래 칸을 비워 둡니다. '
       + '그건 「값이 없다」가 아니라 「위와 같다」는 뜻이라서, 빈 칸으로 세면 자료를 잃게 됩니다.',
  } : null;

  return [q1, q5, q2, q3, q4].filter(q => q && q.options.length >= 2);
}

function renderQuiz() {
  if (!S.quiz) S.quiz = buildQuiz();
  const box = $('quizList');
  box.innerHTML = '';
  let score = 0;
  S.quiz.forEach((q, qi) => {
    if (S.quizDone[qi] === true) score++;
    const div = document.createElement('div');
    div.className = 'quiz';
    const h = document.createElement('div');
    h.className = 'quiz-q';
    h.textContent = `${qi + 1}) ${q.q}`;
    div.appendChild(h);

    const opts = document.createElement('div');
    opts.className = 'quiz-opts';
    q.options.forEach((o, oi) => {
      const lab = document.createElement('label');
      lab.className = 'quiz-opt' + (S.quizPicked[qi] === o ? ' picked' : '');
      const rd = document.createElement('input');
      rd.type = 'radio';
      rd.name = `quiz${qi}`;
      rd.checked = S.quizPicked[qi] === o;
      rd.onchange = () => { S.quizPicked[qi] = o; renderQuiz(); };
      lab.appendChild(rd);
      const sp = document.createElement('span');
      sp.textContent = o;
      lab.appendChild(sp);
      opts.appendChild(lab);
    });
    div.appendChild(opts);

    const row = document.createElement('div');
    row.className = 'btn-row';
    const bChk = document.createElement('button');
    bChk.type = 'button'; bChk.className = 'btn btn-sm'; bChk.textContent = '정답 확인';
    const bHint = document.createElement('button');
    bHint.type = 'button'; bHint.className = 'btn btn-sm'; bHint.textContent = '힌트';
    row.appendChild(bChk); row.appendChild(bHint);
    div.appendChild(row);

    const fb = document.createElement('div');
    fb.className = 'quiz-feedback';
    if (S.quizDone[qi] === true) { fb.classList.add('ok'); fb.textContent = `정답입니다! ${q.why}`; }
    else if (S.quizDone[qi] === false) { fb.classList.add('no'); fb.textContent = '아쉬워요. 다시 한 번 생각해 볼까요? 「힌트」를 눌러도 됩니다.'; }
    div.appendChild(fb);

    bChk.onclick = () => {
      if (S.quizPicked[qi] === undefined) {
        fb.className = 'quiz-feedback hint';
        fb.textContent = '답을 고른 뒤 확인해 주세요.';
        return;
      }
      S.quizDone[qi] = S.quizPicked[qi] === q.answer;
      renderQuiz();
    };
    bHint.onclick = () => { fb.className = 'quiz-feedback hint'; fb.textContent = `힌트: ${q.hint}`; };

    box.appendChild(div);
  });
  $('quizScore').textContent = `맞힌 개수: ${score} / ${S.quiz.length}`;
  if (score === S.quiz.length && S.quiz.length > 0) {
    $('quizScore').textContent += ' — 모두 맞혔어요! 데이터 정제의 이유를 잘 이해했습니다.';
  }
}

/* ═══════════════════════════════════════════════════════════════════════════
   단계 이동
   ═══════════════════════════════════════════════════════════════════════════ */

/** 그 단계로 갈 수 있는지, 못 가면 이유는 무엇인지 */
function stepGate(n) {
  if (n === 1) return { ok: true };
  if (!S.parsed || !S.parsed.columns.length) {
    return { ok: false, why: '먼저 1단계에서 파일을 여세요.' };
  }
  if (n === 2) return { ok: true };
  // 3·4단계는 쓸 열이 1개 이상 있어야 의미가 있다 (열을 안 골랐으면 정리할 것도 저장할 것도 없다)
  if (S.picked.length === 0) {
    return { ok: false, why: '2단계에서 쓸 열을 1개 이상 골라야 해요.' };
  }
  if (S.base && S.base.rows.length === 0) {
    return { ok: false, why: '2단계에서 고른 조건에 맞는 행이 없어요. 값을 다시 켜 주세요.' };
  }
  return { ok: true };
}

async function go(n) {
  const gate = stepGate(n);
  if (!gate.ok) {
    $('lockNotice').innerHTML = noticeHtml('warn', gate.why, '잠겨 있어요.');
    const chip = document.querySelector(`.step[data-step="${n}"]`);
    if (chip) { chip.classList.add('shake'); setTimeout(() => chip.classList.remove('shake'), 320); }
    return;
  }
  // 3단계를 빈 칸이 남은 채로 떠날 때 한 번만 확인
  if (S.step === 3 && n === 4 && S.resultAll && S.resultAll.missingLeft > 0 && !S.warnedOnce.leaveStep3) {
    S.warnedOnce.leaveStep3 = true;
    const ok = await ask('빈 칸이 남아 있어요',
      `<p>빈 칸이 아직 <b>${fmtNum(S.resultAll.missingLeft)}개</b> 남아 있어요.</p>`
      + `<p class="muted">그대로 가도 괜찮지만, 남은 빈 칸은 저장한 파일에도 그대로 들어갑니다. `
      + `「그대로 두기」도 하나의 선택이에요 — 왜 그대로 두는 것이 괜찮은지 생각해 보세요.</p>`,
      '그대로 갈게요', '더 정리할게요');
    if (!ok) return;
  }
  $('lockNotice').innerHTML = '';
  S.step = n;
  render();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function renderNav() {
  const n = S.step;
  const labels = ['', '행·열 고르기', '빈 칸 정리', '엑셀로 저장'];
  const prev = $('btnPrev'), next = $('btnNext');

  prev.disabled = n === 1;
  prev.textContent = n === 1 ? '◀ 이전' : `◀ 이전: ${['', '파일 열기', '행·열 고르기', '빈 칸 정리'][n - 1]}`;

  if (n === 4) {
    next.hidden = true;
  } else {
    next.hidden = false;
    next.textContent = `다음: ${labels[n]} ▶`;
    const gate = stepGate(n + 1);
    next.disabled = !gate.ok;
  }

  // 가운데 상태 문구
  let status = `${n} / 4 단계 · `;
  if (!S.parsed) status += '아직 파일을 열지 않았어요';
  else if (n === 1) status += `${S.fileName} · ${fmtNum(S.parsed.rows.length)}행 × ${S.parsed.columns.length}열`;
  else if (n === 2) status += `열 ${S.picked.length}/${S.parsed.columns.length}개 · `
    + `행 ${fmtNum(S.base.rows.length)}/${fmtNum(S.parsed.rows.length)}개를 골랐어요`;
  else if (n === 3) status += S.resultAll.missingLeft === 0
    ? '빈 칸을 모두 정리했어요' : `빈 칸 ${fmtNum(S.resultAll.missingLeft)}개 남음`;
  else status += `저장할 표: ${fmtNum(S.result.rows.length)}행 × ${S.result.columns.length}열`;
  const gate = stepGate(n + 1);
  $('navStatus').innerHTML = esc(status)
    + (!gate.ok && n < 4 ? `<span class="nav-reason">${esc(gate.why)}</span>` : '');

  // 진행 표시 칩
  document.querySelectorAll('.step').forEach((chip) => {
    const k = Number(chip.dataset.step);
    chip.className = 'step';
    const st = chip.querySelector('.step-state');
    chip.removeAttribute('aria-current');
    chip.removeAttribute('aria-disabled');
    if (k === n) {
      chip.classList.add('current'); st.textContent = '지금 여기';
      chip.setAttribute('aria-current', 'step');
    } else if (!stepGate(k).ok) {
      chip.classList.add('locked'); st.textContent = '잠겨 있어요';
      chip.setAttribute('aria-disabled', 'true');
    } else if (k < n) {
      chip.classList.add('done'); st.textContent = '완료';
    } else {
      chip.classList.add('ready'); st.textContent = '이제 할 수 있어요';
    }
  });
  $('progressTrack').querySelectorAll('span').forEach((s, i) => {
    s.className = i < n ? 'filled' : '';
  });
}

/* ═══════════════════════════════════════════════════════════════════════════
   전체 다시 그리기
   ═══════════════════════════════════════════════════════════════════════════ */

function render() {
  for (let i = 1; i <= 4; i++) $('panel' + i).hidden = (i !== S.step);
  // 되돌릴 것이 없으면 버튼을 회색으로 (숨기지 않는다 — 사라지면 학생이 당황한다)
  document.querySelectorAll('.js-undo').forEach(b => {
    b.disabled = S.undoStack.length === 0;
    b.textContent = S.undoStack.length ? `↺ 되돌리기 (${S.undoStack.length})` : '↺ 되돌리기';
  });
  document.querySelectorAll('.js-redo').forEach(b => { b.disabled = S.redoStack.length === 0; });
  if (S.step === 1) renderStep1();
  if (S.parsed) {
    if (S.step === 2) renderStep2();
    if (S.step === 3) renderStep3();
    if (S.step === 4) renderStep4();
  }
  renderNav();
}

/* ═══════════════════════════════════════════════════════════════════════════
   클릭 연결
   ═══════════════════════════════════════════════════════════════════════════ */

function wire() {
  // ── 1단계 파일 열기
  $('btnPickFile').onclick = () => $('fileInput').click();
  $('fileInput').onchange = (e) => {
    const f = e.target.files && e.target.files[0];
    if (f) { S.forcedEncoding = null; openFile(f); }
    e.target.value = '';   // 같은 파일을 다시 골라도 동작하게
  };
  $('btnSample').onclick = loadSample;
  $('btnOtherFile').onclick = () => { $('fileInput').click(); };

  const dz = $('dropzone');
  ['dragenter', 'dragover'].forEach(ev => dz.addEventListener(ev, (e) => {
    e.preventDefault(); dz.classList.add('over');
    dz.querySelector('.dropzone-main').textContent = '놓으면 파일을 읽습니다';
  }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, (e) => {
    e.preventDefault(); dz.classList.remove('over');
    dz.querySelector('.dropzone-main').textContent = '여기로 파일을 끌어다 놓으세요';
  }));
  dz.addEventListener('drop', (e) => {
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) { S.forcedEncoding = null; openFile(f); }
  });

  // ── 2단계
  $('btnRecommendAll').onclick = () => {
    pushUndo('추천대로 한 번에 맞추기');
    applyRecommendedPlan();
    render();
    toast('추천대로 맞췄어요. 마음에 안 들면 열마다 바꿀 수 있어요.', '되돌리기', doUndo);
  };
  $('btnResetPlan').onclick = () => {
    pushUndo('모두 그대로 두기로 바꾸기');
    for (const c of S.parsed.columns) S.plan[c.name] = { action: ACTIONS.KEEP, value: undefined };
    recompute(); render();
    toast('모두 「빈 칸 그대로 두기」로 되돌렸어요.', '되돌리기', doUndo);
  };
  // 되돌리기 / 다시 하기 (2·3단계 도구줄에 같은 버튼이 하나씩 있다)
  document.querySelectorAll('.js-undo').forEach(b => { b.onclick = doUndo; });
  document.querySelectorAll('.js-redo').forEach(b => { b.onclick = doRedo; });
  document.querySelectorAll('input[name="prevMode"]').forEach((r) => {
    r.onchange = () => { S.previewMode = r.value; renderPreview3(); };
  });
  // 행 고르기 초기화
  $('btnRowFilterReset').onclick = () => {
    if (!Object.keys(S.rowFilter).length) { toast('이미 모든 행이 들어가 있어요.'); return; }
    pushUndo('모든 행 넣기');
    S.rowFilter = {};
    recompute(); render();
    toast('모든 행을 다시 넣었어요.', '되돌리기', doUndo);
  };

  // ── 3단계
  $('btnPickAll').onclick = () => {
    pushUndo('전체 고르기');
    S.picked = S.parsed.columns.map(c => c.name); recompute(); render();
    toast('모든 속성을 골랐어요.', '되돌리기', doUndo);
  };
  $('btnPickNone').onclick = async () => {
    if (!S.picked.length) return;
    const ok = await ask('전체 지우기',
      `<p>골라 둔 ${S.picked.length}개를 모두 지울까요?</p>`
      + '<p class="muted">되돌리기로 다시 살릴 수 있어요.</p>', '지우기');
    if (!ok) return;
    pushUndo('전체 지우기');
    S.picked = []; recompute(); render();
    toast('선택을 모두 지웠어요.', '되살리기', doUndo);
  };
  $('btnPickRecommend').onclick = () => {
    pushUndo('추천만 고르기');
    S.picked = S.parsed.columns.filter(c => isRecommendedColumn(S.stats[c.name])).map(c => c.name);
    recompute(); render();
    toast(`추천 속성 ${S.picked.length}개를 골랐어요.`, '되돌리기', doUndo);
  };
  $('goalInput').oninput = (e) => { S.goal = e.target.value.trim(); };

  // ── 4단계 저장
  const validateName = () => {
    const name = $('saveName').value.trim();
    if (!name) { $('saveError').innerHTML = noticeHtml('error', '파일 이름을 적어 주세요.'); return null; }
    const bad = badFileNameChars(name);
    if (bad.length) {
      $('saveError').innerHTML = noticeHtml('error',
        `파일 이름에 ${bad.join(' ')} 는 쓸 수 없어요. 다른 이름을 적어 주세요.`);
      return null;
    }
    $('saveError').innerHTML = '';
    return name;
  };
  $('btnSaveXlsx').onclick = () => {
    const name = validateName();
    if (!name) return;
    try {
      saveXlsx(buildAOA(), {
        fileName: name,
        sheetName: $('saveSheet').value.trim() || '정리한자료',
        reportLines: $('chkReport').checked ? buildReport() : null,
      });
      $('saveResult').innerHTML = noticeHtml('ok',
        `「${name}.xlsx」 를 저장했어요. 내려받기 폴더를 확인해 보세요. `
        + '이제 엑셀에서 열어 그래프를 만들어 보세요.');
    } catch (e) {
      $('saveResult').innerHTML = noticeHtml('error',
        '파일을 만들지 못했어요. 자료가 너무 크면 3단계에서 열을 더 줄여 보세요. '
        + `(${e && e.message ? e.message : ''})`);
    }
  };
  $('btnSaveCsv').onclick = () => {
    const name = validateName();
    if (!name) return;
    try {
      saveCsv(buildAOA(), { fileName: name });
      $('saveResult').innerHTML = noticeHtml('ok',
        `「${name}.csv」 를 저장했어요. 한글이 깨지지 않게 UTF-8 표시를 붙였습니다.`);
    } catch (e) {
      $('saveResult').innerHTML = noticeHtml('error', '파일을 만들지 못했어요.');
    }
  };
  $('btnCopyReport').onclick = async () => {
    const text = buildReport().map((l, i) => `${i + 1}. ${l}`).join('\n');
    try {
      await navigator.clipboard.writeText(text);
      toast('복사했어요. 보고서나 발표 자료에 붙여 넣으세요.');
    } catch {
      toast('복사하지 못했어요. 아래 목록을 직접 골라서 복사해 주세요.');
    }
  };
  $('btnRestart').onclick = async () => {
    const ok = await ask('처음부터 다시 하기',
      '<p>지금까지 한 정리가 모두 사라집니다. 정말 처음부터 다시 할까요?</p>', '다시 하기');
    if (!ok) return;
    location.reload();
  };

  // ── 도움말
  $('btnHelp').onclick = () => ask('도움말',
    '<ol class="report">'
    + '<li>1단계에서 공공데이터포털에서 받은 엑셀·CSV 파일을 엽니다. 열 이름이 맞는지 확인하세요.</li>'
    + '<li>2단계에서 열마다 빈 칸을 어떻게 할지 고릅니다. 위의 숫자가 어떻게 바뀌는지 보세요.</li>'
    + '<li>3단계에서 쓸 열만 고르고, 4단계에서 엑셀 파일로 저장합니다.</li>'
    + '</ol>'
    + '<p class="muted">이 앱은 파일을 인터넷으로 보내지 않고 브라우저 안에서만 처리합니다. '
    + '자료를 저장하지도 않으므로 새로 고치면 처음부터 다시 해야 합니다. '
    + '개인정보가 든 파일은 열지 않는 것이 좋아요.</p>', '닫기', '');

  // ── 단계 이동
  $('btnPrev').onclick = () => go(Math.max(1, S.step - 1));
  $('btnNext').onclick = () => go(Math.min(4, S.step + 1));
  document.querySelectorAll('.step').forEach((chip) => {
    chip.onclick = () => go(Number(chip.dataset.step));
  });
  // 키보드 단축키
  document.addEventListener('keydown', (e) => {
    // 글자를 적는 중에는 단축키가 끼어들지 않게 한다
    const t = e.target;
    const typing = t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT');
    if (e.altKey) {
      if (e.key === 'ArrowRight') { e.preventDefault(); go(Math.min(4, S.step + 1)); }
      if (e.key === 'ArrowLeft') { e.preventDefault(); go(Math.max(1, S.step - 1)); }
      return;
    }
    if ((e.ctrlKey || e.metaKey) && !typing && S.parsed) {
      const k = e.key.toLowerCase();
      if (k === 'z' && !e.shiftKey) { e.preventDefault(); doUndo(); }
      if (k === 'y' || (k === 'z' && e.shiftKey)) { e.preventDefault(); doRedo(); }
    }
  });

  // 확인 창을 ESC 로 닫을 때도 약속(Promise)이 끝나게
  $('askDialog').addEventListener('cancel', () => { $('askCancel').click(); });
}

wire();
render();
