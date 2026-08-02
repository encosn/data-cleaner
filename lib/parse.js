/**
 * xlsx-parse.js
 * ─────────────────────────────────────────────────────────────────────────────
 * 엑셀 시트(SheetJS worksheet 객체) → "열 이름을 기준으로 다룰 수 있는 표"로 바꿔 주는
 * 순수 함수 모음입니다.
 *
 * 왜 필요한가?
 *   공공데이터포털에서 받은 엑셀은 보기 좋게 만들려고 셀을 많이 "병합"해 둡니다.
 *   병합된 셀은 실제로는 왼쪽 위 한 칸에만 값이 있고 나머지는 빈 칸입니다.
 *   그래서 그냥 읽으면 "서울특별시"가 첫 줄에만 있고 나머지 줄은 비어 보입니다.
 *   이 파일의 함수들은 그런 엑셀을 컴퓨터가 다루기 쉬운 "네모난 표"로 펴 줍니다.
 *
 * 규칙
 *   - DOM/브라우저 전용 API를 쓰지 않습니다 → 브라우저와 Node 양쪽에서 동작합니다.
 *   - SheetJS 를 import 하지 않습니다 → worksheet 를 "그냥 객체"로 읽습니다.
 *     (필요한 셀 주소 계산은 아래에서 직접 합니다. 의존성 0)
 *   - 모든 함수는 입력을 바꾸지 않고(불변) 새 값을 돌려줍니다.
 * ─────────────────────────────────────────────────────────────────────────────
 */

/* ═══════════════════════════════════════════════════════════════════════════
   0. 작은 도구들
   ═══════════════════════════════════════════════════════════════════════════ */

/** 0 → "A", 25 → "Z", 26 → "AA" (엑셀 열 문자) */
export function colToLetter(idx) {
  let n = Number(idx);
  if (!Number.isFinite(n) || n < 0) return '';
  let s = '';
  // 26진법과 비슷하지만 "0이 없는" 진법이라 매번 1을 빼 준다.
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

/**
 * "A1" → {r:0, c:0} 를 정규식 없이 빠르게 계산한다.
 * 셀이 20만 개인 파일에서는 정규식 20만 번이 눈에 보일 만큼 느려진다(측정: 420ms → 90ms).
 * 대문자+숫자 형태만 처리하고, 그 밖의 모양은 null 을 돌려주어 decodeAddr 가 맡는다.
 */
function fastAddr(k) {
  const n = k.length;
  let i = 0, c = 0;
  while (i < n) {
    const ch = k.charCodeAt(i);
    if (ch >= 65 && ch <= 90) { c = c * 26 + (ch - 64); i++; } else break;
  }
  if (i === 0 || i > 3 || i === n) return null;
  let r = 0;
  while (i < n) {
    const ch = k.charCodeAt(i);
    if (ch >= 48 && ch <= 57) { r = r * 10 + (ch - 48); i++; } else return null;
  }
  return r === 0 ? null : { r: r - 1, c: c - 1 };
}

/** "A1" → {r:0, c:0}.  이상한 주소면 null */
export function decodeAddr(addr) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(addr).trim());
  if (!m) return null;
  let c = 0;
  const L = m[1].toUpperCase();
  for (let i = 0; i < L.length; i++) c = c * 26 + (L.charCodeAt(i) - 64);
  const r = parseInt(m[2], 10);
  if (!(r >= 1) || !(c >= 1)) return null;
  return { r: r - 1, c: c - 1 };
}

/** {r:0,c:0} → "A1" */
export function encodeAddr(r, c) {
  return colToLetter(c) + (r + 1);
}

/** 값이 "비었다"고 볼 수 있는가? (null, undefined, 공백만 있는 문자열) */
export function isBlank(v) {
  if (v === null || v === undefined) return true;
  if (typeof v === 'string') return v.trim() === '';
  if (typeof v === 'number') return Number.isNaN(v);
  return false;
}

/**
 * 헤더 이름 정리용: 줄바꿈/탭/여러 칸 공백을 한 칸으로 줄이고 앞뒤를 자른다.
 * \u00A0(줄바꿈 없는 공백)·\u3000(전각 공백)은 눈에는 공백처럼 보이지만 보통 공백과 다른 글자다.
 * 웹에서 복사해 붙인 자료에 자주 섞여 들어오는데, 안 없애면 "서울 시" 와 "서울 시" 가
 * 서로 다른 이름이 되어 열 이름 대조가 실패한다.
 */
export function normName(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return Number.isNaN(v) ? '' : String(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v).replace(/[\s\u00A0\u3000]+/g, ' ').trim();
}

/** "1,234", "12.5%", " 30 " 처럼 사람이 보기 좋게 쓴 숫자도 숫자로 본다. */
export function looksNumeric(v) {
  if (typeof v === 'number') return !Number.isNaN(v);
  if (typeof v !== 'string') return false;
  const s = v.replace(/[,\s\u00A0\u3000₩원$%]/g, '');
  if (s === '' || s === '-') return false;
  return !Number.isNaN(Number(s));
}

/** "2024-01-02", "2024.1.2", "2024/1/2" 같은 날짜 모양인가? */
export function looksDate(v) {
  if (v instanceof Date) return true;
  if (typeof v !== 'string') return false;
  return /^\d{4}[-./]\s?\d{1,2}(\s?[-./]\s?\d{1,2})?\.?$/.test(v.trim());
}

/** 데이터(값)처럼 보이는가 = 숫자나 날짜.  헤더는 보통 글자다. */
function bodyLike(v) {
  return looksNumeric(v) || looksDate(v);
}


/* ═══════════════════════════════════════════════════════════════════════════
   1. 워크시트에서 셀 값만 뽑아 2차원 배열(grid)로 만들기
   ═══════════════════════════════════════════════════════════════════════════ */

/** 엑셀의 날짜 숫자(1900년 1월 1일 = 1)를 "2024-01-31" 모양으로 바꾼다. */
function serialToDateText(serial) {
  // 엑셀은 1900년을 윤년으로 착각하는 유명한 버그가 있어 기준을 1899-12-30 으로 잡는다.
  const ms = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000;
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return null;
  const p = n => String(n).padStart(2, '0');
  return `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())}`;
}

/**
 * 셀 객체 하나에서 "우리가 쓸 값" 하나를 뽑는다.
 *  - 숫자는 숫자로 (나중에 평균 같은 계산을 할 수 있게)
 *  - 날짜는 화면에 보이던 글자로 (엑셀은 날짜를 숫자로 저장한다)
 *  - #N/A 같은 오류 셀은 빈 값으로
 * ※ 파일을 읽을 때 XLSX.read(..., { cellDates: true }) 를 주는 것을 권한다.
 *   그러면 날짜가 t:'d' 로 와서 가장 정확하다. 안 줬을 때를 대비한 예비 처리도 아래에 있다.
 */
export function cellValue(cell) {
  if (cell === null || cell === undefined) return null;
  if (typeof cell !== 'object') return cell;          // 혹시 값만 들어 있으면 그대로
  if (cell.t === 'e') return null;                    // 엑셀 오류값(#DIV/0! 등)
  if (cell.t === 'd' || cell.v instanceof Date) {     // 날짜(cellDates: true 로 읽은 경우)
    if (typeof cell.w === 'string' && cell.w.trim() !== '') return cell.w.trim();
    if (cell.v instanceof Date) return cell.v.toISOString().slice(0, 10);
  }
  // 날짜인데 숫자로 들어온 경우: 서식(z)에 y/m/d 가 있으면 날짜다
  if (cell.t === 'n' && typeof cell.v === 'number' && typeof cell.z === 'string'
      && /[ymd]/.test(cell.z) && !/[eE]\+/.test(cell.z)) {
    if (typeof cell.w === 'string' && cell.w.trim() !== '') return cell.w.trim();
    const txt = serialToDateText(cell.v);
    if (txt) return txt;
  }
  let v = cell.v;
  if (v === null || v === undefined) {
    // 값은 없고 서식만 있는 셀. 표시 문자열이라도 있으면 그걸 쓴다.
    v = (typeof cell.w === 'string' && cell.w.trim() !== '') ? cell.w : null;
  }
  if (typeof v === 'string') {
    const t = v.trim();
    return t === '' ? null : t;
  }
  if (typeof v === 'number' && Number.isNaN(v)) return null;
  return v === undefined ? null : v;
}

/**
 * 병합 정보를 안전한 모양으로 정리한다.
 *  - 항목이 null이거나 s/e가 없으면 버린다
 *  - 좌표가 숫자가 아니거나 음수면 버린다 (해석할 방법이 없다.
 *    0으로 당겨 붙이면 엉뚱한 칸을 덮어써서 헤더가 망가진다 — 실제로 겪은 버그)
 *  - s와 e가 뒤집혀 있으면 min/max로 바로잡는다
 *  - 1칸짜리 "병합"은 병합이 아니므로 버린다
 */
export function normalizeMerges(merges) {
  const out = [];
  if (!Array.isArray(merges)) return out;
  for (const m of merges) {
    if (!m || typeof m !== 'object') continue;
    const s = m.s, e = m.e;
    if (!s || typeof s !== 'object' || !e || typeof e !== 'object') continue;
    const sr = Number(s.r), sc = Number(s.c), er = Number(e.r), ec = Number(e.c);
    if (![sr, sc, er, ec].every(Number.isFinite)) continue;
    if (sr < 0 || sc < 0 || er < 0 || ec < 0) continue;      // 음수 좌표는 버린다
    const r0 = Math.min(sr, er), r1 = Math.max(sr, er);
    const c0 = Math.min(sc, ec), c1 = Math.max(sc, ec);
    if (r0 === r1 && c0 === c1) continue;
    out.push({ s: { r: r0, c: c0 }, e: { r: r1, c: c1 } });
  }
  // 위→아래, 왼→오른쪽 순서로 정렬해 두면 병합이 겹칠 때도 결과가 항상 같아진다.
  out.sort((a, b) => (a.s.r - b.s.r) || (a.s.c - b.s.c));
  return out;
}

/**
 * 워크시트를 2차원 배열로 읽는다.
 * 범위는 (1) !ref (2) 실제로 존재하는 셀 주소 (3) 병합 범위 를 모두 합쳐서 정한다.
 * → !ref 가 틀렸거나 병합이 !ref 밖으로 삐져나가도 죽지 않는다.
 */
export function readGrid(ws) {
  const notices = [];
  if (!ws || typeof ws !== 'object') {
    return { grid: [], nRows: 0, nCols: 0, merges: [], notices: [
      { level: 'error', type: '빈시트', message: '시트를 읽을 수 없습니다.' }
    ] };
  }

  let maxR = -1, maxC = -1;
  // (1) !ref
  const ref = typeof ws['!ref'] === 'string' ? ws['!ref'] : '';
  if (ref) {
    const parts = ref.split(':');
    const a = decodeAddr(parts[0]);
    const b = decodeAddr(parts[parts.length - 1]);
    if (a && b) { maxR = Math.max(maxR, a.r, b.r); maxC = Math.max(maxC, a.c, b.c); }
  }
  // (2) 실제로 존재하는 셀 주소들 (r, c 를 따로 담아 두면 두 번째 훑을 때 다시 계산하지 않는다)
  const allKeys = Object.keys(ws);
  const keys = [];
  const rr = new Int32Array(allKeys.length);
  const cc = new Int32Array(allKeys.length);
  for (let i = 0; i < allKeys.length; i++) {
    const k = allKeys[i];
    if (k.charCodeAt(0) === 33 /* ! */) continue;
    const p = fastAddr(k) || decodeAddr(k);
    if (!p) continue;
    const j = keys.length;
    keys.push(k); rr[j] = p.r; cc[j] = p.c;
    if (p.r > maxR) maxR = p.r;
    if (p.c > maxC) maxC = p.c;
  }
  // (3) 병합 범위 — 표 크기를 넘어가는 병합은 표 안으로 잘라 넣는다.
  //     표 밖으로 나간 부분에는 원래 셀이 없으므로 값을 채울 필요가 없다.
  //     (표를 억지로 늘리면 999999행 같은 엉터리 병합 하나에 메모리가 터진다)
  const rawMerges = normalizeMerges(ws['!merges']);
  const merges = [];
  let clamped = 0, droppedMerges = 0;
  for (const m of rawMerges) {
    if (m.s.r > maxR || m.s.c > maxC) { droppedMerges++; continue; }   // 시작이 표 밖
    const er = Math.min(m.e.r, maxR), ec = Math.min(m.e.c, maxC);
    if (er !== m.e.r || ec !== m.e.c) clamped++;
    if (er === m.s.r && ec === m.s.c) { droppedMerges++; continue; }   // 자르니 1칸
    merges.push({ s: { r: m.s.r, c: m.s.c }, e: { r: er, c: ec } });
  }
  const brokenMerges = (Array.isArray(ws['!merges']) ? ws['!merges'].length : 0) - rawMerges.length;
  if (clamped || droppedMerges || brokenMerges) {
    notices.push({ level: 'warn', type: '이상한병합',
      message: `표 범위를 벗어나거나 망가진 병합 정보를 정리했습니다` +
        ` (잘라냄 ${clamped}건, 무시 ${droppedMerges + brokenMerges}건).` });
  }

  const nRows = maxR + 1, nCols = maxC + 1;
  if (nRows <= 0 || nCols <= 0) {
    return { grid: [], nRows: 0, nCols: 0, merges,
      notices: [{ level: 'error', type: '빈시트', message: '시트가 비어 있습니다.' }] };
  }

  // 네모난 배열을 먼저 null 로 채우고, 있는 셀만 값을 넣는다.
  const grid = new Array(nRows);
  for (let r = 0; r < nRows; r++) grid[r] = new Array(nCols).fill(null);
  for (let j = 0; j < keys.length; j++) {
    const r = rr[j], c = cc[j];
    if (r < nRows && c < nCols) grid[r][c] = cellValue(ws[keys[j]]);
  }
  return { grid, nRows, nCols, merges, notices };
}


/* ═══════════════════════════════════════════════════════════════════════════
   2. 병합 풀기 (이 앱의 핵심)
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * unmergeCells — 병합된 칸을 "값 복사"로 펴 준다.
 *
 * 핵심 아이디어
 *   병합 범위의 왼쪽 위(anchor) 칸에만 값이 있으므로,
 *   그 값을 범위 안의 모든 칸에 복사해 넣는다.
 *   그러면 세로로 병합된 "서울특별시"가 3줄 모두에 들어가고,
 *   가로로 병합된 "인구"가 두 열 모두에 들어간다.
 *
 * 왜 원본을 바꾸지 않고 새 배열을 돌려주는가?
 *   ① 학생이 "헤더 행을 다시 고르기" 같은 조작을 하면 처음부터 다시 계산해야 한다.
 *      원본을 망가뜨렸으면 되돌릴 수 없다.
 *   ② 워크시트 객체는 SheetJS 가 내부적으로도 쓰므로 건드리면 저장할 때 문제가 생길 수 있다.
 *   ③ 순수 함수라서 테스트하기 쉽다.
 *
 * 겹치는 병합(정상 엑셀에는 없지만 손으로 만든 파일엔 있을 수 있음)은
 *   "먼저 나온 병합이 이긴다"로 정해서 결과가 항상 같게 만든다.
 *
 * @param {object|Array} ws  워크시트 객체 또는 이미 만든 grid(2차원 배열)
 * @param {object} [opts]    {merges} — grid 를 직접 넘길 때 병합 정보를 함께 준다
 * @returns {{grid:Array, merges:Array, filledCount:number, notices:Array}}
 */
export function unmergeCells(ws, opts = {}) {
  let grid, merges, notices = [];
  if (Array.isArray(ws)) {
    grid = ws;
    merges = normalizeMerges(opts.merges);
  } else {
    const read = readGrid(ws);
    grid = read.grid;
    merges = read.merges;
    notices = read.notices.slice();
  }
  const nRows = grid.length;
  const nCols = nRows ? Math.max(...grid.map(r => (r ? r.length : 0))) : 0;

  // 새 배열(깊은 복사 1단계). 원본은 절대 건드리지 않는다.
  const out = new Array(nRows);
  for (let r = 0; r < nRows; r++) {
    const src = grid[r] || [];
    const row = new Array(nCols);
    for (let c = 0; c < nCols; c++) row[c] = c < src.length ? src[c] : null;
    out[r] = row;
  }

  const claimed = new Set();   // 이미 어떤 병합이 차지한 칸 (겹침 처리용)
  let filledCount = 0;
  let emptyAnchor = 0;

  for (const m of merges) {
    const r0 = m.s.r, c0 = m.s.c;
    const r1 = Math.min(m.e.r, nRows - 1), c1 = Math.min(m.e.c, nCols - 1);
    if (r0 >= nRows || c0 >= nCols || r1 < r0 || c1 < c0) continue;   // 범위 밖 → 무시

    // 값은 "원본"에서 읽는다. (다른 병합이 덮어쓴 값을 다시 퍼뜨리지 않도록)
    const anchor = (grid[r0] && grid[r0][c0] !== undefined) ? grid[r0][c0] : null;
    if (isBlank(anchor)) emptyAnchor++;

    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const key = r * 100000 + c;
        if (claimed.has(key)) continue;      // 먼저 나온 병합이 이긴다
        claimed.add(key);
        if (r === r0 && c === c0) continue;  // 왼쪽 위 칸은 이미 값이 있다
        out[r][c] = anchor;
        filledCount++;
      }
    }
  }

  if (merges.length > 0) {
    notices.push({ level: 'info', type: '병합해제',
      message: `병합된 칸 ${merges.length}곳을 풀어서 빈 칸 ${filledCount}개에 값을 채웠습니다.` });
  }
  if (emptyAnchor > 0) {
    notices.push({ level: 'warn', type: '빈병합',
      message: `병합돼 있지만 값이 비어 있는 칸이 ${emptyAnchor}곳 있습니다.` });
  }
  // 병합 정보는 "이미 풀었다"는 뜻으로 빈 배열을 함께 돌려준다.
  return { grid: out, merges, mergesAfter: [], filledCount, notices };
}


/* ═══════════════════════════════════════════════════════════════════════════
   3. 완전히 빈 행 / 빈 열 정리
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * trimEmpty — 값이 하나도 없는 행과 열을 없앤다.
 *
 * 판단
 *   · "완전히" 빈 행/열만 지운다. 일부만 빈 칸은 결측치이므로 절대 지우지 않는다
 *     (결측치를 찾아 주는 것이 이 앱의 목적이다).
 *   · 표 중간의 빈 행도 지운다. 남겨 두면 "모든 값이 빈 행"이 생겨 통계가 망가진다.
 *     대신 앞/중간/뒤 중 어디였는지 알려 주어 학생이 확인할 수 있게 한다.
 *     (중간 빈 행은 "표가 두 개"라는 신호일 수도 있어서 경고를 띄운다)
 *   · 지운 뒤에도 원래 위치를 알 수 있게 keptRowIdxs / keptColIdxs 를 함께 준다.
 */
export function trimEmpty(rows) {
  const src = Array.isArray(rows) ? rows : [];
  const nRows = src.length;
  const nCols = nRows ? Math.max(0, ...src.map(r => (r ? r.length : 0))) : 0;

  const rowEmpty = new Array(nRows).fill(true);
  const colEmpty = new Array(nCols).fill(true);
  for (let r = 0; r < nRows; r++) {
    const row = src[r] || [];
    for (let c = 0; c < nCols; c++) {
      if (!isBlank(row[c])) { rowEmpty[r] = false; colEmpty[c] = false; }
    }
  }

  const keptRowIdxs = [], keptColIdxs = [];
  for (let r = 0; r < nRows; r++) if (!rowEmpty[r]) keptRowIdxs.push(r);
  for (let c = 0; c < nCols; c++) if (!colEmpty[c]) keptColIdxs.push(c);

  const first = keptRowIdxs.length ? keptRowIdxs[0] : nRows;
  const last = keptRowIdxs.length ? keptRowIdxs[keptRowIdxs.length - 1] : -1;
  const emptyRows = { leading: [], middle: [], trailing: [] };
  for (let r = 0; r < nRows; r++) {
    if (!rowEmpty[r]) continue;
    if (r < first) emptyRows.leading.push(r);
    else if (r > last) emptyRows.trailing.push(r);
    else emptyRows.middle.push(r);
  }
  const emptyColIdxs = [];
  for (let c = 0; c < nCols; c++) if (colEmpty[c]) emptyColIdxs.push(c);

  const out = keptRowIdxs.map(r => keptColIdxs.map(c => {
    const row = src[r] || [];
    return row[c] === undefined ? null : row[c];
  }));

  return { rows: out, keptRowIdxs, keptColIdxs, emptyRows, emptyColIdxs };
}


/* ═══════════════════════════════════════════════════════════════════════════
   4. 맨 아래 주석 행 찾기
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 표 아래 주석 줄을 알아보는 무늬.
 *
 * ⚠ 여기서 크게 조심할 것이 있다.
 *   처음에는 "출처" "비고" "참고" "단위" 로 시작하면 주석으로 봤는데,
 *   그러면 「출처불명」 「비고사항」 「참고자료실」 같은 **진짜 데이터가 조용히 사라진다.**
 *   (지역명·부서명·기관명에 이런 낱말이 들어가는 일은 흔하다!)
 *   → 그래서 낱말로 시작하는 경우는 반드시 뒤에 **쌍점(:)** 이 있어야 주석으로 본다.
 *     ※ ＊ * 【 注 처럼 사람이 주석 표시로만 쓰는 기호는 쌍점 없이도 주석으로 본다.
 */
const NOTE_HEAD = new RegExp(
  '^(' +
  '[※＊*【注]' +                                  // 주석 표시 기호
  '|주\\s*\\d*\\s*[):.]' +                        // 주) 주1) 주.
  '|주석' +
  '|\\(\\s*단위' +                                // (단위: 명)
  '|(자료출처|자료|출처|참고|비고|단위|비주)\\s*[:：]' +  // 반드시 쌍점이 있어야 한다
  ')');

/**
 * detectTrailingNotes — 표 맨 아래에 붙는 설명 줄을 찾는다.
 *   예) "※ 자료: 통계청", "* 단위: 명", "출처: 행정안전부"
 *
 * 판단
 *   · 조건: 표의 아래쪽에 있고, 값이 들어 있는 칸이 1~2개뿐이며(대개 첫 칸),
 *     그 값이 ※ · * · "자료:" · "출처" 같은 말로 시작하는 글자.
 *   · 자동으로 지우지 않고 "찾아서 알려 준 뒤 기본값으로 제외"한다.
 *     - 표에 남겨 두면 그 줄만 값이 텅 비어서 결측치 통계와 그래프가 다 틀어진다.
 *     - 그렇다고 몰래 지우면 학생이 자료를 잃어버린 줄 모른다.
 *     → 그래서 notices 로 보여 주고, UI에서 "되살리기"를 누를 수 있게 한다
 *       (parseSheet 의 keepNotes 옵션).
 */
export function detectTrailingNotes(rows) {
  const src = Array.isArray(rows) ? rows : [];
  const nCols = src.length ? Math.max(0, ...src.map(r => (r ? r.length : 0))) : 0;
  const noteIdxs = [];
  const details = [];

  for (let r = src.length - 1; r >= 0; r--) {
    const row = src[r] || [];
    const filled = [];
    for (let c = 0; c < nCols; c++) if (!isBlank(row[c])) filled.push(c);

    if (filled.length === 0) continue;                 // 빈 행은 건너뛰고 더 위를 본다
    const onlyLeft = filled.length <= 2 && filled[0] <= 1;
    const text = normName(row[filled[0]]);
    if (onlyLeft && NOTE_HEAD.test(text)) {
      noteIdxs.push(r);
      details.push({ rowIdx: r, text });
      continue;                                        // 주석이 여러 줄일 수 있다
    }
    break;                                             // 주석이 아닌 실제 데이터 → 멈춤
  }
  noteIdxs.reverse(); details.reverse();
  return { noteIdxs, details };
}


/* ═══════════════════════════════════════════════════════════════════════════
   5. 헤더(열 이름) 행 자동 추정 — "추천"만 한다
   ═══════════════════════════════════════════════════════════════════════════ */

/** 한 행의 기본 통계 */
function rowStats(row, nCols) {
  let nonEmpty = 0, strings = 0, numeric = 0;
  const seen = new Set();
  for (let c = 0; c < nCols; c++) {
    const v = row ? row[c] : null;
    if (isBlank(v)) continue;
    nonEmpty++;
    seen.add(normName(v));
    if (bodyLike(v)) numeric++; else strings++;
  }
  return { nonEmpty, strings, numeric, distinct: seen.size,
           fill: nCols ? nonEmpty / nCols : 0 };
}

/**
 * 제목 행인가?
 *  ① 같은 값이 여러 칸에 걸쳐 있다 → 가로로 넓게 병합된 제목
 *  ② 값이 1~2칸뿐인데, 아래쪽 행들은 훨씬 많이 채워져 있다 → 제목이나 안내문
 *     ("(단위: 명)", "작성: OO과" 처럼 헤더 위에 여러 줄 있는 경우가 흔하다)
 */
function titleLike(rows, r, nCols) {
  const row = rows[r];
  const st = rowStats(row, nCols);
  if (st.nonEmpty === 0) return 0;
  if (st.nonEmpty >= 2 && st.distinct === 1) return 1;
  if (st.nonEmpty <= 2) {
    let maxBelow = 0;
    for (let k = r + 1; k < Math.min(rows.length, r + 6); k++) {
      maxBelow = Math.max(maxBelow, rowStats(rows[k], nCols).nonEmpty);
    }
    if (maxBelow >= st.nonEmpty + 2) return 0.8;
    if (st.nonEmpty === 1 && nCols >= 3 && st.fill <= 0.34) return 0.7;
  }
  return 0;
}

/** 칸의 종류: 0=빈칸, 1=숫자·날짜, 2=글자 */
function cellKind(v) { return isBlank(v) ? 0 : (bodyLike(v) ? 1 : 2); }

/**
 * 근거 ⑤: "이 행이 그냥 또 하나의 데이터 행처럼 보이는가?"
 *   데이터 행들은 서로 모양이 닮았다(예: 글자·숫자·숫자).
 *   후보 행의 모양이 아래 데이터들의 대표 모양과 똑같으면 그건 헤더가 아니라 데이터다.
 *   열 이름이 숫자(2020, 2021…)여서 다른 근거가 안 통할 때 이 근거가 도움이 된다.
 */
function bodyPatternMatch(row, body, nCols) {
  if (!body.length) return 0;
  let same = 0, counted = 0;
  for (let c = 0; c < nCols; c++) {
    const cnt = [0, 0, 0];
    for (const b of body) cnt[cellKind(b ? b[c] : null)]++;
    let modal = 0;
    if (cnt[1] >= cnt[2] && cnt[1] >= cnt[0]) modal = 1;
    else if (cnt[2] >= cnt[0]) modal = 2;
    counted++;
    if (cellKind(row ? row[c] : null) === modal) same++;
  }
  return counted ? same / counted : 0;
}

/** 주석 행인가? */
function noteLike(row, nCols) {
  const st = rowStats(row, nCols);
  if (st.nonEmpty === 0 || st.nonEmpty > 2) return 0;
  for (let c = 0; c < nCols; c++) {
    if (!isBlank(row[c])) return (c <= 1 && NOTE_HEAD.test(normName(row[c]))) ? 1 : 0;
  }
  return 0;
}

/**
 * 근거 ①: "상위 이름 + 하위 이름" 2단 헤더의 흔적
 *   가로 병합을 풀면 상위 행에 같은 값이 옆으로 반복된다(예: 인구 | 인구).
 *   그 아래에 서로 다른 값(남 | 여)이 있으면 2단 헤더가 거의 확실하다.
 *   단, 반복이 표 전체 너비를 덮으면 그건 헤더가 아니라 "제목"이다.
 */
function groupLabelEvidence(rows, r, nCols) {
  const row = rows[r] || [], next = rows[r + 1] || [];
  let i = 0;
  while (i < nCols) {
    const v = normName(row[i]);
    if (v === '') { i++; continue; }
    let j = i;
    while (j + 1 < nCols && normName(row[j + 1]) === v) j++;
    const runLen = j - i + 1;
    if (runLen >= 2 && runLen < nCols) {
      const below = new Set();
      for (let k = i; k <= j; k++) { const b = normName(next[k]); if (b !== '') below.add(b); }
      if (below.size >= 2) return 1;
    }
    i = j + 1;
  }
  return 0;
}

/** 근거 ②: 위 행이 빈 칸인데 아래 행에 값이 있다(병합을 안 쓴 2단 헤더) */
function complementEvidence(rows, r, nCols) {
  const row = rows[r] || [], next = rows[r + 1] || [];
  let holes = 0, fills = 0;
  for (let c = 0; c < nCols; c++) {
    if (isBlank(row[c])) { holes++; if (!isBlank(next[c])) fills++; }
  }
  return holes > 0 && fills >= 1 && fills === holes ? 1 : 0;
}

/**
 * 근거 ③(가장 강력): 헤더에 있는 "세로 병합"의 높이
 *   공공데이터 파일은 지역·시도처럼 한 칸짜리 헤더를 헤더 높이만큼 세로로 병합한다.
 *   그 병합의 높이가 곧 헤더 행 수다.  예) 지역 A2:A4 → 헤더는 3행
 */
function vMergeHeight(merges, startRow) {
  let h = 1;
  for (const m of merges || []) {
    if (m.s.r === startRow && m.e.r > m.s.r) h = Math.max(h, m.e.r - m.s.r + 1);
  }
  return h;
}

/**
 * detectHeaderRows — 헤더가 몇 번째 행(들)인지 점수로 추정한다.
 *
 * ★ 이 함수는 "추천"만 한다.  자동 추정은 틀릴 수 있으므로
 *   화면에서 학생/교사가 헤더 행을 직접 고를 수 있어야 한다(parseSheet 의 headerRowIdxs).
 *
 * @param {Array<Array>} rows  병합을 이미 푼 2차원 배열
 * @param {object} [opts] {merges, maxScanRows, maxHeaderRows, sep}
 * @returns {{best:number[], candidates:Array}}  candidates 는 점수 높은 순
 */
export function detectHeaderRows(rows, opts = {}) {
  const src = Array.isArray(rows) ? rows : [];
  const nRows = src.length;
  const nCols = nRows ? Math.max(0, ...src.map(r => (r ? r.length : 0))) : 0;
  if (nRows === 0 || nCols === 0) return { best: [], candidates: [] };

  const merges = opts.merges || [];
  const maxScan = Math.min(nRows, opts.maxScanRows || 12);
  const maxHeaderRows = opts.maxHeaderRows || 4;
  const sep = opts.sep || ' ';
  const noteSet = new Set((opts.noteIdxs || []));

  // 근거 ④: "제목도 주석도 빈 행도 아닌 첫 번째 행"은 헤더일 가능성이 매우 높다.
  //   (엑셀에서 표를 만들면 거의 항상 그렇다. 열 이름이 숫자여서 다른 근거가
  //    모두 무력해지는 경우 — 예: 지역|2020|2021 — 에 이 근거가 결정적이다)
  let firstEligible = -1;
  for (let r = 0; r < nRows; r++) {
    if (noteSet.has(r)) continue;
    if (rowStats(src[r], nCols).nonEmpty === 0) continue;
    if (titleLike(src, r, nCols) > 0) continue;
    if (noteLike(src[r], nCols) > 0) continue;
    firstEligible = r; break;
  }

  const cand = [];

  for (let start = 0; start < maxScan; start++) {
    if (noteSet.has(start)) continue;
    if (rowStats(src[start], nCols).nonEmpty === 0) continue;   // 빈 행은 헤더가 아니다

    // 시험해 볼 헤더 높이들: 1,2,3 + 세로 병합이 알려 주는 높이
    const vh = vMergeHeight(merges, start);
    const lens = new Set([1, 2, 3, vh]);

    for (const len of lens) {
      if (len < 1 || len > maxHeaderRows) continue;
      if (start + len > nRows - 1) continue;         // 데이터가 한 줄은 남아야 한다
      const idxs = [];
      for (let k = 0; k < len; k++) idxs.push(start + k);

      // ── 이 블록으로 열 이름을 만들어 본다
      const built = buildColumnNames(src, idxs, { sep, quiet: true });
      const names = built.columns.map(c => c.rawName);      // 유일화 전 이름
      const namedCols = names.filter(n => n !== '').length;
      const headerFill = nCols ? namedCols / nCols : 0;
      const distinct = new Set(names.filter(n => n !== '')).size;
      const nameUnique = namedCols ? distinct / namedCols : 0;
      const dupFrac = namedCols ? (namedCols - distinct) / namedCols : 0;

      // ── 헤더 블록 안이 "글자"로만 되어 있는가 (데이터 행을 잡아먹지 않았는가)
      let hCells = 0, hStr = 0, dataLikeRows = 0;
      for (const r of idxs) {
        const st = rowStats(src[r], nCols);
        hCells += st.nonEmpty; hStr += st.strings;
        if (st.numeric >= 1 && st.fill >= 0.5) dataLikeRows++;
      }
      const allString = hCells ? hStr / hCells : 0;
      const dataInBlock = len ? dataLikeRows / len : 0;

      // ── 본문 표본: 헤더 아래에서 빈 행·주석 행을 뺀 최대 20행
      const body = [];
      for (let r = start + len; r < nRows && body.length < 20; r++) {
        if (noteSet.has(r)) continue;
        if (rowStats(src[r], nCols).nonEmpty === 0) continue;
        body.push(src[r]);
      }
      let bodyFillSum = 0;
      for (const b of body) bodyFillSum += rowStats(b, nCols).fill;
      const bodyFill = body.length ? bodyFillSum / body.length : 0;

      // ── 타입 대비: 헤더는 글자인데 아래는 숫자/날짜 → 헤더라는 강한 증거
      let eligible = 0, contrast = 0;
      for (let c = 0; c < nCols; c++) {
        const nm = names[c];
        if (nm === '' || bodyLike(nm)) continue;        // 헤더가 숫자면 비교 불가
        let cnt = 0, num = 0;
        for (const b of body) { const v = b ? b[c] : null; if (isBlank(v)) continue; cnt++; if (bodyLike(v)) num++; }
        if (cnt === 0) continue;
        eligible++;
        if (num / cnt >= 0.6) contrast++;
      }
      const typeContrast = eligible ? contrast / eligible : 0;

      const tl = titleLike(src, start, nCols);
      const nl = noteLike(src[start], nCols);
      const ge = len > 1 ? Math.max(groupLabelEvidence(src, start, nCols),
                                    complementEvidence(src, start, nCols)) : 0;
      const vhMatch = (len === vh && vh > 1) ? 1 : (len === 1 && vh === 1 ? 1 : 0);
      const isFirst = start === firstEligible ? 1 : 0;
      const bpm = bodyPatternMatch(src[start], body, nCols);

      const reasons = {
        headerFill, allString, nameUnique, dupFrac, bodyFill, typeContrast,
        titleLike: tl, noteLike: nl, groupEvidence: ge, vMergeMatch: vhMatch,
        firstRealRow: isFirst, bodyPatternMatch: bpm, dataInBlock,
        bodySample: body.length, headerHeightFromMerge: vh,
      };

      const score =
          30 * headerFill              // 그 행에 이름이 채워진 비율
        + 20 * allString               // 헤더는 보통 글자다
        + 20 * nameUnique              // 열 이름은 서로 달라야 한다
        +  8 * bodyFill                // 아래쪽에 값이 잘 채워져 있는가
        + 22 * typeContrast            // 헤더는 글자, 아래는 숫자 → 강한 근거
        + 30 * vhMatch                 // 세로 병합 높이와 일치 → 가장 강한 근거
        + 18 * isFirst                 // 제목·주석이 아닌 첫 행 (엑셀에서 거의 항상 헤더)
        + (len > 1 ? 22 * ge : 0)      // 2단 헤더의 흔적
        - 20 * bpm                     // 아래 데이터들과 모양이 똑같다 = 데이터 행
        - 25 * dupFrac                 // 이름이 겹친다 = 헤더를 덜 읽었다
        - 45 * tl                      // 제목 행
        - 60 * nl                      // 주석 행
        - 35 * dataInBlock             // 데이터 행을 헤더로 삼았다
        - 12 * (len - 1)               // 헤더를 여러 행으로 보려면 근거가 있어야 한다
        -  3 * start;                  // 같은 조건이면 위쪽 행을 헤더로 본다

      cand.push({ rowIdxs: idxs, score: Math.round(score * 10) / 10, reasons,
                  preview: built.columns.map(c => c.name) });
    }
  }

  cand.sort((a, b) => b.score - a.score || a.rowIdxs.length - b.rowIdxs.length
                       || a.rowIdxs[0] - b.rowIdxs[0]);
  return { best: cand.length ? cand[0].rowIdxs : [], candidates: cand.slice(0, 8) };
}


/* ═══════════════════════════════════════════════════════════════════════════
   6. 열 이름 만들기
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * buildColumnNames — 헤더 행(들)을 합쳐 "열 이름"을 만든다.
 *
 * 규칙(정해 둔 것)
 *   1) 헤더가 1행이면 그 값을 그대로 쓴다.
 *   2) 2행 이상이면 위 → 아래 순서로 이어 붙인다.  구분자는 공백 한 칸.
 *      예) 인구 + 남  →  "인구 남"
 *   3) 위와 아래가 같은 말이면 한 번만 쓴다.  (세로 병합을 풀면 지역/지역 이 되기 때문)
 *      예) 지역 + 지역 →  "지역"
 *   4) 빈 값은 건너뛴다.
 *   5) 줄바꿈·탭·이중 공백은 공백 한 칸으로 정리하고 앞뒤 공백을 자른다.
 *   6) 이름이 끝까지 비면  "이름없음(E열)"  처럼 엑셀 열 문자를 넣어 만든다.
 *   7) 이름이 겹치면  "인구", "인구(2)", "인구(3)" 으로 유일하게 만든다.
 *      → 이후 모든 처리가 "열 이름"을 열쇠로 쓰므로 이름은 반드시 유일해야 한다.
 *   8) 엑셀 열 문자(letter)와 원래 열 번호(srcIndex)를 함께 돌려준다.
 *      → 학생이 원본 엑셀과 눈으로 대조할 수 있다.
 *
 * @param {Array<Array>} rows        병합을 푼 2차원 배열
 * @param {number[]} headerRowIdxs   헤더로 쓸 행 번호들 (rows 기준, 0부터)
 * @param {object} [opts] {sep, colLetters, quiet}
 *        colLetters: rows 의 열이 원본 엑셀의 어느 열인지 (빈 열을 지운 뒤라면 필요)
 */
export function buildColumnNames(rows, headerRowIdxs, opts = {}) {
  const src = Array.isArray(rows) ? rows : [];
  const nCols = src.length ? Math.max(0, ...src.map(r => (r ? r.length : 0))) : 0;
  const sep = opts.sep === undefined ? ' ' : opts.sep;
  const idxs = (Array.isArray(headerRowIdxs) ? headerRowIdxs : [])
                 .filter(i => Number.isInteger(i) && i >= 0 && i < src.length);
  const letters = opts.colLetters || null;
  const notices = [];

  const columns = [];
  const used = new Map();      // 이름 → 몇 번 나왔는지
  let emptyNameCount = 0;

  for (let c = 0; c < nCols; c++) {
    const letter = letters ? (letters[c] || colToLetter(c)) : colToLetter(c);

    // (2)(4)(5) 위에서 아래로 모으며 빈 값은 건너뛴다
    const parts = [];
    for (const r of idxs) {
      const t = normName(src[r] ? src[r][c] : null);
      if (t === '') continue;
      // (3) 바로 위 조각과 같으면 중복이므로 건너뛴다
      if (parts.length && parts[parts.length - 1] === t) continue;
      parts.push(t);
    }
    // 떨어져 있는 같은 말도 한 번만 (예: 인구 / 남 / 인구 같은 이상한 경우)
    const uniqParts = parts.filter((p, i) => parts.indexOf(p) === i);
    let rawName = uniqParts.join(sep).replace(/\s+/g, ' ').trim();

    // (6) 이름이 비었을 때
    let name = rawName;
    let autoNamed = false;
    if (name === '') {
      name = `이름없음(${letter}열)`;
      autoNamed = true;
      emptyNameCount++;
    }
    // 자바스크립트에서 "__proto__" 는 객체의 특별한 열쇠라서 값을 넣어도 사라진다.
    // 열 이름이 우연히 그 이름이면 살짝 바꿔 준다.
    if (name === '__proto__') name = '__proto__(열)';

    // (7) 같은 이름 유일화
    let finalName = name;
    if (used.has(finalName)) {
      let k = used.get(finalName) + 1;
      while (used.has(`${name}(${k})`)) k++;
      used.set(name, k);
      finalName = `${name}(${k})`;
    }
    used.set(finalName, used.get(finalName) || 1);

    columns.push({
      name: finalName,
      rawName,                       // 유일화·자동이름 적용 전 (점수 계산용)
      letter,
      srcIndex: c,
      parts: uniqParts,
      autoNamed,
      renamedFrom: finalName !== name ? name : null,
    });
  }

  if (!opts.quiet) {
    const dup = columns.filter(c => c.renamedFrom);
    if (dup.length) {
      notices.push({ level: 'warn', type: '이름중복',
        message: `열 이름이 겹쳐서 이름을 바꿨습니다: ` +
          dup.map(d => `${d.letter}열 "${d.renamedFrom}" → "${d.name}"`).join(', ') });
    }
    if (emptyNameCount) {
      notices.push({ level: 'warn', type: '이름없음',
        message: `이름이 비어 있는 열이 ${emptyNameCount}개 있어 임시 이름을 붙였습니다.` });
    }
  }
  return { columns, notices };
}


/* ═══════════════════════════════════════════════════════════════════════════
   7. 전체 조합 — parseSheet
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * parseSheet — 워크시트 하나를 "열 이름 기준의 표"로 바꾼다.
 *
 * 처리 순서
 *   ① 셀 값을 2차원 배열로 읽는다        readGrid
 *   ② 병합을 푼다(값 복사)               unmergeCells   ← 사용자 요구사항의 핵심
 *   ③ 완전히 빈 행/열을 없앤다           trimEmpty
 *   ④ 맨 아래 주석 행을 찾는다           detectTrailingNotes
 *   ⑤ 헤더 행을 추정한다(추천)           detectHeaderRows
 *   ⑥ 열 이름을 만든다                   buildColumnNames
 *   ⑦ 데이터 행을 "열 이름: 값" 객체로 만든다
 *
 * @param {object} ws  SheetJS worksheet
 * @param {object} [opts]
 *   headerRowIdxs : 사용자가 직접 고른 헤더 행 (trim 후 배열 기준). 주면 자동 추정을 덮어쓴다.
 *   headerExcelRows : 사용자가 엑셀 행 번호(1부터)로 고를 때. 위보다 쓰기 쉽다.
 *   keepNotes     : true 면 주석 행도 데이터로 남긴다 (기본 false)
 *   sep           : 여러 행 헤더를 이을 구분자 (기본 ' ')
 * @returns {{columns, rows, rowMeta, notices, headerInfo, excludedRows, stats}}
 */
export function parseSheet(ws, opts = {}) {
  const notices = [];
  const read = readGrid(ws);
  notices.push(...read.notices);
  if (!read.grid.length) {
    return { columns: [], rows: [], rowMeta: [], notices, headerInfo: { best: [], candidates: [] },
             excludedRows: [], stats: { rowCount: 0, colCount: 0 } };
  }

  // ② 병합 풀기
  const un = unmergeCells(read.grid, { merges: read.merges });
  notices.push(...un.notices.filter(n => n.type !== '빈시트'));
  const filled = un.grid;

  // ③ 빈 행/열 정리 (원래 위치를 기억해 둔다)
  const tr = trimEmpty(filled);
  const grid = tr.rows;
  const excelRowOf = i => (tr.keptRowIdxs[i] ?? i) + 1;                 // 엑셀 행 번호(1부터)
  const colLetters = tr.keptColIdxs.map(c => colToLetter(c));
  if (tr.emptyRows.middle.length) {
    notices.push({ level: 'warn', type: '중간빈행',
      message: `표 중간에 완전히 빈 행이 ${tr.emptyRows.middle.length}개 있어 지웠습니다` +
        ` (엑셀 ${tr.emptyRows.middle.map(r => r + 1).join(', ')}행).` +
        ` 한 시트에 표가 두 개 들어 있지 않은지 확인해 주세요.` });
  }
  if (tr.emptyColIdxs.length) {
    notices.push({ level: 'info', type: '빈열',
      message: `값이 하나도 없는 열 ${tr.emptyColIdxs.length}개를 지웠습니다` +
        ` (${tr.emptyColIdxs.map(c => colToLetter(c) + '열').join(', ')}).` });
  }
  if (!grid.length) {
    return { columns: [], rows: [], rowMeta: [], notices,
             headerInfo: { best: [], candidates: [] }, excludedRows: [],
             stats: { rowCount: 0, colCount: 0 } };
  }

  // 병합 정보도 "빈 행/열을 지운 뒤"의 좌표로 옮긴다.
  // (헤더 추정에서 세로 병합의 높이를 근거로 쓰기 때문에 꼭 필요하다)
  // 원본 번호 → 정리 후 위치를 미리 표로 만들어 둔다. 지워진 행이면 그 앞의 위치.
  const rowPos = new Map(tr.keptRowIdxs.map((v, i) => [v, i]));
  const colPos = new Map(tr.keptColIdxs.map((v, i) => [v, i]));
  const posLE = (kept, n) => {           // n칸짜리 조회표를 한 번만 만든다
    const arr = new Array(n).fill(-1);
    let p = -1, k = 0;
    for (let i = 0; i < n; i++) { if (kept[k] === i) { p = k; k++; } arr[i] = p; }
    return arr;
  };
  const nRowsF = filled.length, nColsF = filled.length ? filled[0].length : 0;
  const rowLE = posLE(tr.keptRowIdxs, nRowsF);
  const colLE = posLE(tr.keptColIdxs, nColsF);
  const mergesT = [];
  for (const m of read.merges) {
    const sr = rowPos.get(m.s.r), sc = colPos.get(m.s.c);
    if (sr === undefined || sc === undefined) continue;   // 시작 행/열이 지워졌다
    const er = rowLE[Math.min(m.e.r, nRowsF - 1)];
    const ec = colLE[Math.min(m.e.c, nColsF - 1)];
    mergesT.push({ s: { r: sr, c: sc }, e: { r: Math.max(sr, er), c: Math.max(sc, ec) } });
  }

  // ④ 주석 행
  const notes = detectTrailingNotes(grid);
  const keepNotes = !!opts.keepNotes;
  if (notes.noteIdxs.length) {
    notices.push({ level: keepNotes ? 'info' : 'warn', type: '주석행',
      message: `표 아래에 설명으로 보이는 줄 ${notes.noteIdxs.length}개를 찾았습니다` +
        (keepNotes ? ' (데이터로 남겼습니다).' : ' (데이터에서 빼 두었습니다).'),
      detail: notes.details.map(d => ({ excelRow: excelRowOf(d.rowIdx), text: d.text })) });
  }

  // ⑤ 헤더 추정 (또는 사용자가 고른 값)
  const headerInfo = detectHeaderRows(grid, {
    merges: mergesT, sep: opts.sep, noteIdxs: notes.noteIdxs,
  });
  let headerRowIdxs = headerInfo.best;
  if (Array.isArray(opts.headerExcelRows) && opts.headerExcelRows.length) {
    headerRowIdxs = opts.headerExcelRows
      .map(n => tr.keptRowIdxs.indexOf(n - 1))
      .filter(i => i >= 0);
  } else if (Array.isArray(opts.headerRowIdxs) && opts.headerRowIdxs.length) {
    headerRowIdxs = opts.headerRowIdxs.filter(i => Number.isInteger(i) && i >= 0 && i < grid.length);
  }
  if (!headerRowIdxs.length) {
    headerRowIdxs = [0];
    notices.push({ level: 'warn', type: '헤더추정실패',
      message: '헤더 행을 찾지 못해 첫 줄을 열 이름으로 썼습니다. 직접 골라 주세요.' });
  }
  headerRowIdxs = [...headerRowIdxs].sort((a, b) => a - b);

  // 제목 행이 헤더 위에 있으면 알려 준다 (학생에게 "왜 1행을 안 썼는지" 설명)
  const firstHeader = headerRowIdxs[0];
  const nColsT = grid[0].length;
  const skippedTitles = [];
  for (let r = 0; r < firstHeader; r++) if (titleLike(grid, r, nColsT) > 0) skippedTitles.push(r);
  if (skippedTitles.length) {
    notices.push({ level: 'info', type: '제목행',
      message: `엑셀 ${skippedTitles.map(r => excelRowOf(r)).join(', ')}행은 표 제목으로 보여 열 이름에서 뺐습니다.` });
  }

  // ⑥ 열 이름
  const built = buildColumnNames(grid, headerRowIdxs, { sep: opts.sep, colLetters });
  notices.push(...built.notices);
  const columns = built.columns.map(c => ({
    name: c.name, letter: c.letter, srcIndex: tr.keptColIdxs[c.srcIndex],
    parts: c.parts, autoNamed: c.autoNamed, renamedFrom: c.renamedFrom,
  }));

  // 병합이 헤더 행과 데이터 행을 함께 덮고 있으면 알려 준다 (잘못 만든 파일의 신호)
  const lastHeaderRow = headerRowIdxs[headerRowIdxs.length - 1];
  if (mergesT.some(m => m.s.r <= lastHeaderRow && m.e.r > lastHeaderRow)) {
    notices.push({ level: 'warn', type: '헤더데이터병합',
      message: '열 이름 행과 데이터 행이 하나로 병합된 곳이 있습니다.' +
               ' 열 이름이 데이터 칸에도 복사되니 결과를 꼭 확인해 주세요.' });
  }

  // ⑦ 데이터 행 → 열 이름을 열쇠로 하는 객체
  const noteSet = new Set(keepNotes ? [] : notes.noteIdxs);
  const headerSet = new Set(headerRowIdxs);
  const lastHeader = lastHeaderRow;
  const rows = [], rowMeta = [], excludedRows = [];
  let inlineNotes = 0;
  let repeatedHeaders = 0;

  // 한 줄의 값들을 하나의 글자열(서명)로 만들 때 쓰는 구분자.
  // 빈 문자열로 이으면 ["가","나"] 와 ["가나"] 가 같은 서명이 되어 엉뚱한 줄을 지운다.
  // 그래서 자료에 절대 나오지 않는 제어문자(\u0001)를 구분자로 쓴다.
  const SIG_SEP = '\u0001';

  // 표가 길면 인쇄할 때 페이지마다 열 이름 줄을 다시 넣는 일이 흔하다.
  // 그 줄을 데이터로 두면 "서울"이 하나 더 세어지고 평균이 틀어지므로 빼야 한다.
  // 판단은 아주 엄격하게: 헤더 행과 값이 "완전히 똑같은" 줄만 뺀다.
  const headerSigs = new Set(
    headerRowIdxs.map(r => (grid[r] || []).map(v => normName(v)).join(SIG_SEP))
  );

  for (let r = 0; r < grid.length; r++) {
    if (headerSet.has(r)) continue;
    if (r < lastHeader) {                        // 헤더보다 위 = 제목/설명
      excludedRows.push({ excelRow: excelRowOf(r), reason: '헤더보다 위에 있는 줄(표 제목 등)',
                          values: grid[r].slice() });
      continue;
    }
    if (noteSet.has(r)) {
      excludedRows.push({ excelRow: excelRowOf(r), reason: '표 아래 설명 줄',
                          values: grid[r].slice() });
      continue;
    }
    // 표 중간에 끼어 있는 설명 줄("※ 단위: 명" 등)도 데이터가 아니다.
    // 조건이 까다로워서(값이 1~2칸 + ※·자료·출처·단위 로 시작) 진짜 데이터를 잡아먹을 위험은 낮다.
    if (!keepNotes && noteLike(grid[r], nColsT) > 0) {
      excludedRows.push({ excelRow: excelRowOf(r), reason: '표 중간에 끼어 있는 설명 줄',
                          values: grid[r].slice() });
      inlineNotes++;
      continue;
    }
    // 열 이름 줄이 표 중간에 다시 나온 경우 (인쇄용 파일에서 흔하다)
    if (headerSigs.has((grid[r] || []).map(v => normName(v)).join(SIG_SEP))) {
      excludedRows.push({ excelRow: excelRowOf(r), reason: '열 이름 줄이 다시 나온 줄',
                          values: grid[r].slice() });
      repeatedHeaders++;
      continue;
    }
    const obj = {};
    let missing = 0;
    for (let c = 0; c < columns.length; c++) {
      const v = grid[r][c];
      obj[columns[c].name] = isBlank(v) ? null : v;
      if (isBlank(v)) missing++;
    }
    rows.push(obj);
    rowMeta.push({ excelRow: excelRowOf(r), missingCount: missing });
  }

  if (inlineNotes > 0) {
    notices.push({ level: 'warn', type: '중간설명행',
      message: `표 중간에 설명으로 보이는 줄 ${inlineNotes}개를 찾아 데이터에서 빼 두었습니다.` });
  }
  if (repeatedHeaders > 0) {
    notices.push({ level: 'warn', type: '반복헤더행',
      message: `표 중간에 열 이름 줄이 ${repeatedHeaders}번 더 나와서 데이터에서 빼 두었습니다.` +
               ' 인쇄용으로 만든 파일에서 흔한 일이에요. 그대로 두면 자료가 실제보다 많아집니다.' });
  }

  // 결측치 요약 (2단계 화면에서 바로 쓸 수 있게)
  const missingByColumn = {};
  for (const col of columns) {
    let n = 0;
    for (const row of rows) if (isBlank(row[col.name])) n++;
    missingByColumn[col.name] = n;
  }

  return {
    columns, rows, rowMeta, notices, headerInfo, excludedRows,
    stats: {
      rowCount: rows.length,
      colCount: columns.length,
      headerExcelRows: headerRowIdxs.map(excelRowOf),
      mergeCount: read.merges.length,
      filledCellCount: un.filledCount,
      missingByColumn,
      missingTotal: Object.values(missingByColumn).reduce((a, b) => a + b, 0),
    },
  };
}


/* ═══════════════════════════════════════════════════════════════════════════
   8. 다음 단계(속성 고르기 → 엑셀 저장)에서 쓰는 도구
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * ⚠️ 꼭 기억할 것
 *   rows 의 각 줄은 "열 이름"이 열쇠인 객체다. 그런데 자바스크립트는
 *   "2020" 처럼 숫자로 보이는 열쇠를 객체 앞쪽으로 몰아 넣는다.
 *   그래서 Object.keys(row) 로 열 순서를 얻으면 순서가 뒤바뀐다!
 *   → 열 순서가 필요할 때는 항상 columns 배열을 쓴다. 아래 함수들이 그렇게 한다.
 */

/** 고른 열만 남긴 새 결과를 만든다 (3단계: 활용할 속성 선택) */
export function pickColumns(parsed, names) {
  const want = new Set(names);
  const columns = parsed.columns.filter(c => want.has(c.name));
  const rows = parsed.rows.map(r => {
    const o = {};
    for (const c of columns) o[c.name] = r[c.name] ?? null;
    return o;
  });
  return { ...parsed, columns, rows };
}

/** 결과를 2차원 배열로 바꾼다 (4단계: 엑셀 저장 — XLSX.utils.aoa_to_sheet 에 그대로 넣는다) */
export function toAOA(parsed, opts = {}) {
  const head = parsed.columns.map(c => c.name);
  const body = parsed.rows.map(r => parsed.columns.map(c => {
    const v = r[c.name];
    return v === null || v === undefined ? (opts.blank ?? '') : v;
  }));
  return [head, ...body];
}

export default {
  colToLetter, decodeAddr, encodeAddr, isBlank, normName, looksNumeric, looksDate,
  cellValue, normalizeMerges, readGrid, unmergeCells, trimEmpty,
  detectTrailingNotes, detectHeaderRows, buildColumnNames, parseSheet,
  pickColumns, toAOA,
};
