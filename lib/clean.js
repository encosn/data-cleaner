/**
 * clean.js
 * ─────────────────────────────────────────────────────────────────────────────
 * "빈 칸(결측치)을 찾고 정리하는" 계산을 모아 둔 파일입니다.
 * 화면을 그리는 코드는 여기 없습니다 (main.js 가 담당). 그래서 시험하기 쉽습니다.
 *
 * ★ 이 파일에서 가장 중요한 약속 두 가지 ★
 *
 *  1) 원본은 절대 바꾸지 않는다.
 *     학생이 고른 처리 방법은 "설정(plan)"으로만 저장해 두고,
 *     화면에 보여 줄 때마다 「원본 + 설정」을 다시 계산한다.
 *     → 되돌리기가 공짜로 되고, "순서 때문에 결과가 달라지는" 버그가 아예 생기지 않는다.
 *
 *  2) 평균·중앙값 같은 대표값은 항상 "원본" 기준으로 구한다.
 *     다른 열을 정리한 결과나 지워진 행에 영향을 받지 않는다.
 *
 * 처리 순서도 고정해 둔다:  ① 값 채우기(모든 열)  →  ② 행 지우기(모든 열의 합집합)
 * 채우기를 먼저 해야 "채웠으니 이 행은 안 지워도 된다"가 자연스럽게 성립한다.
 * ─────────────────────────────────────────────────────────────────────────────
 */

/* ═══════════════════════════════════════════════════════════════════════════
   1. "빈 칸"을 알아보기
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 공공데이터에서 빈 칸 대신 자주 쓰이는 말들.
 * 뜻이 조금씩 다르므로 묶어서 정리해 두고, 화면에서 학생이 켜고 끌 수 있게 한다.
 * ★ 기본은 모두 꺼짐(=빈 칸으로 보지 않음)이다. 몰래 바꾸면 학생이 자료를 잃은 줄 모른다.
 */
export const MISSING_TOKEN_GROUPS = [
  {
    id: 'dash',
    label: '- (붙임표)',
    help: '통계표에서 「해당하는 수치가 아예 없음」을 뜻하는 관례입니다. 숫자 0과는 다릅니다.',
    test: (s) => /^[-‐-―−－ー]+$/.test(s),
  },
  {
    id: 'dot',
    label: '. 또는 …',
    help: '「조사는 했지만 알 수 없음(미상)」을 뜻하는 관례입니다.',
    test: (s) => /^[.·]{1,3}$/.test(s) || s === '…',
  },
  {
    id: 'na',
    label: 'N/A · NULL',
    help: '컴퓨터가 자동으로 넣은 「값 없음」 표시입니다.',
    test: (s) => ['n/a', 'na', 'n.a.', 'null', 'nil', 'none', 'nan', '?', '??'].includes(s),
  },
  {
    id: 'excelerr',
    label: '#N/A · #VALUE!',
    help: '엑셀 수식이 깨진 채로 저장된 값입니다. 계산에 쓸 수 없습니다.',
    test: (s) => /^#(n\/a|value!|name\?|div\/0!|ref!|null!|num!)$/.test(s),
  },
  {
    id: 'none_kr',
    label: '없음 · 해당없음',
    help: '「해당없음」은 애초에 조사 대상이 아니었다는 뜻이고, 「실적없음」은 0에 가깝습니다. 뜻이 다르니 원본 설명을 확인하세요.',
    test: (s) => ['없음', '해당없음', '해당사항없음', '비해당', '대상없음', '실적없음',
                  '자료없음', '데이터없음', '정보없음', '내역없음'].includes(s),
  },
  {
    id: 'unknown_kr',
    label: '미상 · 알수없음',
    help: '값은 있는데 우리가 모르는 경우입니다.',
    test: (s) => ['미상', '불명', '알수없음', '확인불가', '확인안됨', '미확인', '조사중', '파악중'].includes(s),
  },
  {
    id: 'secret_kr',
    label: '비공개 · 미공개',
    help: '값이 있는데 일부러 숨긴 것입니다. 평균 같은 값으로 채우면 사실을 크게 왜곡합니다.',
    test: (s) => ['비공개', '미공개', '공개불가', '공개제한', '미제공', '제공안함', '비밀', '대외비'].includes(s),
  },
  {
    id: 'blank_kr',
    label: '미기재 · 무응답',
    help: '적어야 하는데 비워 둔 경우입니다.',
    test: (s) => ['미기재', '미입력', '미작성', '기재없음', '공란', '무응답', '응답거부', '결측', '누락'].includes(s),
  },
  {
    id: 'star',
    label: '* 또는 x',
    help: '통계표에서 「비밀보호를 위해 공개하지 않음」을 뜻하기도 합니다.',
    test: (s) => /^[*x×※#]{1,2}$/.test(s),
  },
  {
    id: 'zero',
    label: '숫자 0',
    risky: true,
    help: '주의! 0을 빈 칸으로 보면 「좌석수 0석」처럼 진짜 0인 값까지 사라집니다. 정말 필요할 때만 켜세요.',
    test: (s) => /^0+(\.0+)?$/.test(s),
  },
];

/** 눈에 안 보이는 공백까지 없애고, 전각 문자를 반각으로 통일해서 비교하기 쉽게 만든다. */
function canon(v) {
  return String(v)
    .normalize('NFKC')
    .replace(/[\u00A0\u3000\u200B\uFEFF\u202F\u2009]/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * 이 값이 빈 칸인가?
 * @param {*} v
 * @param {Set<string>} [activeTokens]  학생이 "빈 칸으로 보겠다"고 켠 표기 묶음 id 들
 */
export function isMissing(v, activeTokens) {
  if (v === null || v === undefined) return true;
  if (typeof v === 'number') return Number.isNaN(v);
  if (v instanceof Date) return Number.isNaN(v.getTime());
  const s = canon(v);
  if (s === '') return true;
  if (!activeTokens || activeTokens.size === 0) return false;
  const squeezed = s.replace(/\s+/g, '');
  for (const g of MISSING_TOKEN_GROUPS) {
    if (!activeTokens.has(g.id)) continue;
    if (g.test(squeezed)) return true;
  }
  return false;
}

/** 켜지지 않은 표기라도 몇 번 나오는지 세어 두면 "이거 빈 칸 아닐까요?" 하고 물어볼 수 있다. */
export function countTokens(rows, columns) {
  const counts = {};
  for (const g of MISSING_TOKEN_GROUPS) counts[g.id] = 0;
  for (const row of rows) {
    for (const col of columns) {
      const v = row[col.name];
      if (v === null || v === undefined) continue;
      const s = canon(v).replace(/\s+/g, '');
      if (s === '') continue;
      for (const g of MISSING_TOKEN_GROUPS) if (g.test(s)) counts[g.id]++;
    }
  }
  return counts;
}


/* ═══════════════════════════════════════════════════════════════════════════
   2. 숫자로 읽어 보기
   ═══════════════════════════════════════════════════════════════════════════ */

/** 값 뒤에 붙는 단위 글자들. "1,234명" 의 "명" 같은 것. */
const UNIT_TAIL = /(명|개소|개|건|대|호|세대|가구|인|원|천원|백만원|억원|조원|%|퍼센트|㎡|㎢|㎞|km|㎏|kg|톤|ha|℃|년|월|일|시간|분|초|회|점|권|부|실|동|층|권역)$/;

/**
 * 사람이 보기 좋게 쓴 숫자를 컴퓨터가 계산할 수 있는 숫자로 바꾼다.
 *   "1,234"   → 1234      (천 단위 콤마 제거)
 *   "1,234명" → 1234      (단위 글자 제거)
 *   "(1,234)" → -1234     (회계에서 괄호는 마이너스라는 약속)
 *   "△12"    → -12       (통계표에서 △는 줄었다는 뜻)
 *   "１２３"   → 123       (전각 숫자를 반각으로)
 * 숫자로 볼 수 없으면 null 을 돌려준다.
 */
export function toNumber(v) {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return Number.isNaN(v) ? null : v;
  if (v instanceof Date) return null;
  let s = String(v)
    .normalize('NFKC')
    .replace(/[\s\u00A0\u202F\u200B]/g, '')
    .replace(/^'/, '');                       // 엑셀이 "글자로 취급하라"고 붙인 따옴표
  if (s === '') return null;
  let neg = false;
  const paren = s.match(/^\(([\d,.]+)\)$/);
  if (paren) { s = paren[1]; neg = true; }
  if (/^[△▲]/.test(s)) { s = s.slice(1); neg = true; }
  if (/^[▽▼]/.test(s)) { s = s.slice(1); }
  s = s.replace(UNIT_TAIL, '').replace(/,/g, '');
  if (!/^[+-]?\d+(\.\d+)?$/.test(s)) return null;
  const n = parseFloat(s);
  if (Number.isNaN(n)) return null;
  return neg ? -Math.abs(n) : n;
}

/** 소수점 아래 자릿수 (표시용 반올림 자리를 정할 때 쓴다) */
function decimalsOf(n) {
  if (!Number.isFinite(n)) return 0;
  const s = String(n);
  const i = s.indexOf('.');
  return i < 0 ? 0 : Math.min(s.length - i - 1, 6);
}


/* ═══════════════════════════════════════════════════════════════════════════
   3. 열마다 성격을 살펴보기
   ═══════════════════════════════════════════════════════════════════════════ */

/** 열 이름이 "이름표"처럼 보이는가? (번호·코드는 채워 넣으면 없는 대상을 만드는 셈이다) */
const ID_NAME = /(번호|코드|id|일련|고유|우편|사업자|법정동|행정동|학교명|기관명|성명)/i;
/** 개인정보가 들어 있을 수 있는 열 이름 */
const PRIVACY_NAME = /(이름|성명|전화|휴대|연락처|이메일|주민|생년|주소|학번|담당자)/i;

/**
 * 열 하나의 통계를 낸다. 항상 "원본 rows" 를 받는다.
 * @returns {{
 *   name, letter, type, total, missing, missingRate,
 *   numericRate, distinct, distinctRate, isId, isPrivacy,
 *   mean, median, min, max, decimals, mode, modeTies, samples
 * }}
 */
export function analyzeColumn(rows, col, activeTokens) {
  const total = rows.length;
  const live = [];            // 빈 칸이 아닌 값들
  let missing = 0;
  for (const row of rows) {
    const v = row[col.name];
    if (isMissing(v, activeTokens)) missing++;
    else live.push(v);
  }

  // 숫자로 바꿀 수 있는 값이 얼마나 되나?
  const nums = [];
  for (const v of live) { const n = toNumber(v); if (n !== null) nums.push(n); }
  const numericRate = live.length ? nums.length / live.length : 0;

  // 날짜처럼 보이는 값이 대부분인가?
  let dateCount = 0;
  for (const v of live) {
    if (v instanceof Date) { dateCount++; continue; }
    if (typeof v === 'string' && /^\d{4}[-./]\s?\d{1,2}([-./]\s?\d{1,2})?$/.test(v.trim())) dateCount++;
  }
  const dateRate = live.length ? dateCount / live.length : 0;

  let type = 'text';
  if (live.length === 0) type = 'empty';
  else if (dateRate >= 0.8) type = 'date';
  else if (numericRate >= 0.9) type = 'number';

  // 이름표(번호·코드·이름) 열은 채워 넣으면 없는 대상을 만드는 셈이라 채우기를 막는다.
  //
  // ⚠ 여기서 조심할 것: "값이 모두 다르면 이름표"라는 규칙을 숫자 열에 그대로 쓰면
  //   인구·좌석수·자료수처럼 값이 겹치지 않는 평범한 측정값 열까지 이름표로 잘못 보게 된다.
  //   (교실에서 쓰는 자료는 20~30행이라 숫자가 겹칠 일이 거의 없다!)
  //   그러면 평균·중앙값 채우기가 아예 나오지 않아 수업의 핵심 활동을 못 하게 된다.
  //   → 숫자 열은 "열 이름"으로만 판단하고(번호·코드·ID 등), 값의 다양함은 보지 않는다.
  const distinctSet = new Set(live.map(v => canon(v)));
  const distinct = distinctSet.size;
  const distinctRate = live.length ? distinct / live.length : 0;
  const nameLooksId = ID_NAME.test(col.name);
  const isId = type === 'number'
    ? nameLooksId
    : (nameLooksId || (live.length > 3 && distinctRate >= 0.95));
  const isPrivacy = PRIVACY_NAME.test(col.name);

  // 숫자 열의 대표값들 — 원본 기준으로 딱 한 번만 계산한다
  let mean = null, median = null, min = null, max = null, decimals = 0;
  if (type === 'number' && nums.length) {
    const sorted = [...nums].sort((a, b) => a - b);
    min = sorted[0];
    max = sorted[sorted.length - 1];
    mean = nums.reduce((a, b) => a + b, 0) / nums.length;
    const mid = Math.floor(sorted.length / 2);
    median = sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
    // 채울 값의 소수 자릿수는 원본을 따른다. 최대 2자리.
    decimals = Math.min(2, Math.max(...nums.map(decimalsOf)));
  }

  // 가장 많이 나온 값 (문자 열에서 참고용으로 보여 준다)
  let mode = null, modeTies = 0;
  if (live.length) {
    const cnt = new Map();
    for (const v of live) {
      const k = canon(v);
      const e = cnt.get(k);
      if (e) e.n++; else cnt.set(k, { n: 1, raw: v });
    }
    let best = 0;
    for (const e of cnt.values()) if (e.n > best) best = e.n;
    if (best > 1) {
      const tied = [...cnt.values()].filter(e => e.n === best);
      mode = tied[0].raw;
      modeTies = tied.length;
    }
  }

  // 연도 열은 천 단위 쉼표를 붙이면 안 된다 (1998 이 "1,998" 로 보이면 학생이 헷갈린다).
  // 열 이름이 연도처럼 보이고, 값이 모두 네 자리 정수일 때만 그렇게 본다.
  const isYear = type === 'number'
    && /(연도|년도|년$|연$|year)/i.test(col.name)
    && nums.length > 0
    && nums.every(n => Number.isInteger(n) && n >= 1000 && n <= 2999);

  return {
    name: col.name, letter: col.letter, type,
    total, missing, missingRate: total ? missing / total : 0,
    numericRate, distinct, distinctRate, isId, isPrivacy, isYear,
    mean, median, min, max, decimals, mode, modeTies,
    samples: live.slice(0, 3),
  };
}

/** 모든 열의 통계 */
export function analyzeAll(rows, columns, activeTokens) {
  const out = {};
  for (const col of columns) out[col.name] = analyzeColumn(rows, col, activeTokens);
  return out;
}


/* ═══════════════════════════════════════════════════════════════════════════
   3-1. 「반복 생략」 찾기 — 빈 칸처럼 보이지만 결측치가 아닌 것
   ───────────────────────────────────────────────────────────────────────────
   ★ 이 앱에서 가장 중요한 기능 중 하나다. 왜 필요한지 꼭 읽어라.

   공공데이터(특히 KOSIS 국가통계포털)에서 받은 엑셀은 이렇게 생겼다.

        성씨, 본관별 | 행정구역별 | 2015
        ------------ | ---------- | ----------
        김(金)        | 전국        | 10,689,959
                     | 서울특별시   |  2,000,000
                     | 부산광역시   |    700,000
        이(李)        | 전국        |  7,306,828
                     | 서울특별시   |  1,400,000

   엑셀에서 보면 「김(金)」이 여러 줄에 걸쳐 병합된 것처럼 보인다.
   그런데 파일 안에는 **병합 정보가 아예 없고**, 첫 줄에만 값이 있고 나머지는 빈 칸이다.
   (같은 값을 반복해서 쓰지 않고 생략한 것 — KOSIS 파일이 거의 다 이렇다. 실제 파일로 확인했다)

   이 빈 칸을 결측치로 보면 어떻게 되는가?
     · 「성씨」 열이 "94% 비어 있는 쓸모없는 열"로 판정된다
     · 앱이 그 열을 빼라고 권한다 → **성씨 정보를 통째로 잃는다**
     · 행을 지우면 전체 자료의 94%가 사라진다
   즉 자료를 완전히 망친다. 그래서 결측치로 세기 **전에** 알아보고 위 값으로 채워야 한다.

   ★ 다만 아무 때나 채우면 안 된다. 진짜 결측치일 수도 있기 때문이다.
     그래서 아래 조건을 모두 만족할 때만 "반복 생략"으로 보고, 그때도 화면에 알리고
     학생이 되돌릴 수 있게 한다.
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 반복 생략으로 보이는 열을 찾는다.
 *
 * 판단 조건 (모두 만족해야 한다)
 *   ① 글자 열이다 (숫자 열의 빈 칸은 반복 생략일 가능성이 낮다)
 *   ② 첫 행에는 값이 있다 (물려받을 값이 없으면 생략이 아니다)
 *   ③ 빈 칸이 「값 → 빈 칸 여러 개 → 값」 묶음 모양으로 2번 이상 반복된다
 *   ④ 빈 칸인 행에서도 **다른 열에는 값이 있다** (그 행이 진짜 자료라는 뜻)
 *      → 이 조건이 핵심이다. 행 전체가 비었으면 그냥 빈 행이지 반복 생략이 아니다
 *   ⑤ 묶음의 크기가 어느 정도 고르다 (통계표는 규칙적으로 반복된다)
 *
 * @returns {Array<{name, letter, blanks, groups, avgRun, otherFilledRate, confident, reason}>}
 */
export function detectRepeatedOmission(rows, columns) {
  const out = [];
  if (!rows.length || !columns.length) return out;

  const blank = (v) => v === null || v === undefined || String(v).trim() === '';

  for (const col of columns) {
    const vals = rows.map(r => r[col.name]);

    // ② 첫 행이 비어 있으면 물려받을 값이 없다
    if (blank(vals[0])) continue;

    const blanks = vals.filter(blank).length;
    if (blanks === 0) continue;

    // ① 숫자 열이면 넘어간다 (빈 칸이 아닌 값들로 판단)
    const live = vals.filter(v => !blank(v));
    if (live.length && live.filter(v => toNumber(v) !== null).length / live.length >= 0.9) continue;

    // ③ 「값 → 빈 칸 묶음」 세기
    const runs = [];
    let cur = 0;
    for (const v of vals) {
      if (blank(v)) cur++;
      else { if (cur) runs.push(cur); cur = 0; }
    }
    if (cur) runs.push(cur);
    if (runs.length < 2) continue;

    // ④ 빈 칸인 행에서 다른 열이 채워져 있는가
    const others = columns.filter(c => c.name !== col.name);
    if (!others.length) continue;
    let blankRows = 0, blankRowsWithData = 0;
    rows.forEach((r) => {
      if (!blank(r[col.name])) return;
      blankRows++;
      if (others.some(c => !blank(r[c.name]))) blankRowsWithData++;
    });
    const otherFilledRate = blankRows ? blankRowsWithData / blankRows : 0;

    // ⑤ 묶음 크기가 고른가 (가장 흔한 크기가 전체의 절반 이상)
    const cnt = new Map();
    for (const n of runs) cnt.set(n, (cnt.get(n) || 0) + 1);
    let modeRun = 0, modeN = 0;
    for (const [n, k] of cnt) if (k > modeN) { modeN = k; modeRun = n; }
    const regular = modeN / runs.length >= 0.5;

    const avgRun = runs.reduce((a, b) => a + b, 0) / runs.length;
    const confident = otherFilledRate >= 0.9 && runs.length >= 2 && regular;

    out.push({
      name: col.name, letter: col.letter,
      blanks, groups: runs.length, avgRun: Math.round(avgRun * 10) / 10,
      otherFilledRate, confident,
      reason: confident
        ? `값 ${runs.length}개가 각각 아래로 평균 ${Math.round(avgRun)}칸씩 이어집니다.`
          + ` 빈 칸인 줄에도 다른 열에는 값이 있어서(${Math.round(otherFilledRate * 100)}%),`
          + ` 같은 값을 반복해서 쓰지 않고 생략한 것으로 보입니다.`
        : `묶음 모양이 규칙적이지 않아 확실하지 않습니다. 원본을 보고 직접 정해 주세요.`,
    });
  }
  return out;
}

/**
 * 지정한 열들의 빈 칸을 「바로 위 값」으로 채운다 (반복 생략 되살리기).
 * 원본을 바꾸지 않고 새 배열을 돌려준다.
 */
export function fillDownColumns(rows, names) {
  if (!names || !names.length) return rows;
  const blank = (v) => v === null || v === undefined || String(v).trim() === '';
  const last = {};
  return rows.map((r) => {
    const o = { ...r };
    for (const n of names) {
      if (blank(o[n])) { if (last[n] !== undefined) o[n] = last[n]; }
      else last[n] = o[n];
    }
    return o;
  });
}


/* ═══════════════════════════════════════════════════════════════════════════
   3-2. 합계·소계 행 찾기
   ───────────────────────────────────────────────────────────────────────────
   공공데이터 파일에는 보고서 표를 그대로 옮긴 것이 많아 「합계」 줄이 섞여 있다.
   그 줄을 그냥 두면 서울 인구를 두 번 세는 셈이 되어 평균이 크게 틀어진다.

   ★ 그런데 자동으로 지우지 않는다.
     「기획총괄계」 같은 부서명이나 지명에도 '계' 가 들어갈 수 있어서 잘못 지울 위험이 있고,
     무엇보다 "합계 줄을 남기면 평균이 어떻게 달라지는가" 를 눈으로 확인하는 것이
     이 수업에서 가장 중요한 장면이다. 그래서 찾아서 알려 주고 결정은 학생이 한다.
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * 합계·전체를 뜻하는 낱말.
 * '전국' 도 넣는다 — 시도별 자료에서 「전국」 줄은 아래 시도들을 이미 다 더한 값이라,
 * 그대로 두고 평균을 내면 같은 사람을 두 번 세게 된다. (KOSIS 통계표에 거의 항상 있다)
 */
const TOTAL_WORD = /^(합계|총계|소계|누계|계|전체|전국|평균|총합|합|total|sum|subtotal|average)$/i;

/**
 * 합계로 보이는 행을 찾는다.
 * @returns {{idxs:number[], labels:string[]}}  idxs 는 rows 배열의 몇 번째인지
 */
export function detectTotalRows(rows, columns) {
  // 글자 열 중 앞의 두 개만 본다 (합계 표시는 보통 가장 왼쪽 칸에 있다)
  const textCols = columns.filter(c => {
    const vs = rows.map(r => r[c.name]).filter(v => !isMissing(v, null));
    if (!vs.length) return false;
    return vs.filter(v => toNumber(v) !== null).length / vs.length < 0.5;
  }).slice(0, 2);
  if (!textCols.length) return { idxs: [], labels: [] };

  const idxs = [], labels = [];
  rows.forEach((row, i) => {
    for (const c of textCols) {
      const v = row[c.name];
      if (isMissing(v, null)) continue;
      const s = canon(v).replace(/\s+/g, '');
      if (TOTAL_WORD.test(s)) { idxs.push(i); labels.push(String(v).trim()); return; }
    }
  });
  return { idxs, labels };
}


/* ═══════════════════════════════════════════════════════════════════════════
   3-3. 행 고르기 — 값으로 걸러내기
   ───────────────────────────────────────────────────────────────────────────
   "새 엑셀에 넣을 행을 고르고 싶다" 는 요구를 위한 기능이다.
   예) 성씨 155개 중 김·이·박만, 지역 18개 중 서울·부산만 골라서 새 파일 만들기.
   ═══════════════════════════════════════════════════════════════════════════ */

/** 행 고르기에 쓸 수 있는 열인가? (값 종류가 너무 많으면 체크박스로 고를 수 없다) */
export function filterableColumns(rows, columns, stats, maxDistinct = 300) {
  return columns.filter(c => {
    const st = stats[c.name];
    if (!st || st.type === 'empty') return false;
    if (st.distinct < 2) return false;               // 값이 한 가지면 걸러낼 것이 없다
    if (st.distinct > maxDistinct) return false;     // 너무 많으면 목록이 쓸모없다
    // 값이 거의 모두 다른 열(이름표)도 걸러내기에 부적합
    if (st.distinctRate > 0.9 && st.distinct > 50) return false;
    return true;
  });
}

/** 그 열에 어떤 값들이 몇 번 나오는지 (많이 나온 순) */
export function valueCounts(rows, name) {
  const m = new Map();
  for (const r of rows) {
    const key = filterKey(r[name]);
    const e = m.get(key);
    if (e) e.n++; else m.set(key, { value: key, n: 1 });
  }
  return [...m.values()].sort((a, b) => b.n - a.n || String(a.value).localeCompare(b.value, 'ko'));
}

/** 값 하나를 걸러내기 열쇠로 바꾼다 (빈 칸도 하나의 값으로 고를 수 있게) */
export function filterKey(v) {
  return v === null || v === undefined || String(v).trim() === '' ? '(빈 칸)' : String(v);
}

/**
 * 행 걸러내기를 적용한다.
 *
 * ★ 약속: **열 이름이 filters 에 있으면** 그 열로 걸러낸다.
 *   · 열 이름이 없으면        → 그 열로는 걸러내지 않는다 (전부 넣기)
 *   · 열 이름이 있고 Set 이 비었으면 → 남는 행이 0개 (전부 빼기)
 *   빈 Set 을 "걸러내지 않음" 으로 보면 「전부 빼기」를 표현할 수 없어서 이렇게 정했다.
 *
 * @param {Array} rows
 * @param {object} filters  { 열이름: Set<string> } — 그 열에서 남길 값들
 * @returns {{rows:Array, keep:boolean[]}}
 */
export function applyRowFilter(rows, filters) {
  const names = Object.keys(filters || {}).filter(n => filters[n] instanceof Set);
  if (!names.length) return { rows, keep: rows.map(() => true) };
  const keep = rows.map((r) => names.every((n) => filters[n].has(filterKey(r[n]))));
  return { rows: rows.filter((_, i) => keep[i]), keep };
}


/* ═══════════════════════════════════════════════════════════════════════════
   4. 처리 방법 (열마다 하나씩 고른다)
   ═══════════════════════════════════════════════════════════════════════════ */

export const ACTIONS = {
  KEEP:   'keep',      // 빈 칸 그대로 두기
  DROP:   'dropRow',   // 빈 칸이 있는 행 지우기
  MEAN:   'mean',      // 평균값으로 채우기
  MEDIAN: 'median',    // 가운데 값(중앙값)으로 채우기
  TEXT:   'noneText',  // '없음' 같은 글자로 채우기
  CUSTOM: 'custom',    // 내가 정한 값으로 채우기
};

/** 채울 값의 소수점을 원본에 맞춰 반올림한다. 화면에 보인 값과 저장되는 값을 똑같이 만든다. */
export function roundLike(value, decimals) {
  const p = Math.pow(10, decimals || 0);
  return Math.round(value * p) / p;
}

/**
 * 이 열에서 고를 수 있는 방법 목록. 열의 성격에 따라 다르게 준다.
 * 선택지가 너무 많으면 학생이 헤맨다 → 열마다 최대 5개.
 */
export function optionsFor(stat) {
  const opts = [{ action: ACTIONS.KEEP, label: '빈 칸 그대로 두기' }];
  if (stat.missing === 0) return opts;

  opts.push({ action: ACTIONS.DROP, label: `빈 칸 있는 행 지우기 (${stat.missing}행)` });

  if (stat.type === 'number' && !stat.isId && stat.mean !== null) {
    const m = roundLike(stat.mean, stat.decimals);
    const md = roundLike(stat.median, stat.decimals);
    opts.push({ action: ACTIONS.MEAN, label: `평균값으로 채우기 (${fmtNum(m)})`, value: m });
    opts.push({ action: ACTIONS.MEDIAN, label: `가운데 값으로 채우기 (${fmtNum(md)})`, value: md });
  } else if (stat.type !== 'number' && !stat.isId) {
    // 이미 '없음'이 값으로 있으면 원본과 구별할 수 없으니 다른 말을 쓴다
    const hasNone = stat.samples.some(v => canon(v) === '없음');
    const word = hasNone ? '조사안됨' : '없음';
    opts.push({ action: ACTIONS.TEXT, label: `「${word}」으로 채우기`, value: word });
  }

  opts.push({ action: ACTIONS.CUSTOM, label: '내가 정한 값으로 채우기' });
  return opts;
}

/** 이 열에는 어떤 방법을 권할까? 위에서부터 먼저 걸리는 규칙 하나만 쓴다. */
export function recommend(stat) {
  if (stat.missing === 0) {
    return { action: ACTIONS.KEEP, badge: '깨끗함', why: '빈 칸이 없어요. 그대로 쓸 수 있어요.' };
  }
  if (stat.missingRate >= 0.9) {
    return { action: ACTIONS.KEEP, badge: '거의 비어 있음', excludeHint: true,
      why: '이 열은 90%가 넘게 비어 있어요. 채워도 대부분 우리가 만든 값이 됩니다. 3단계에서 빼는 걸 권해요.' };
  }
  if (stat.isId) {
    return {
      action: stat.missingRate < 0.05 ? ACTIONS.DROP : ACTIONS.KEEP,
      badge: '이름표 열',
      why: '이 열은 이름표(번호·코드·이름)라서 채워 넣으면 세상에 없는 대상을 만드는 셈이 돼요.',
    };
  }
  if (stat.type === 'number') {
    if (stat.missingRate < 0.03) {
      return { action: ACTIONS.DROP,
        why: '빈 칸이 3%도 안 돼요. 그 행만 빼면 값을 꾸며내지 않고 정리할 수 있어요.' };
    }
    if (stat.missingRate < 0.30) {
      // 아주 크거나 작은 값이 섞이면 평균이 흔들린다 → 그럴 때는 가운데 값
      const skewed = Math.abs(stat.mean - stat.median) > 0.1 * Math.abs(stat.mean || 1);
      return skewed
        ? { action: ACTIONS.MEDIAN, why: '아주 크거나 작은 값이 섞여서 평균이 흔들려요. 가운데 값이 더 안전해요.' }
        : { action: ACTIONS.MEAN, why: '빈 칸이 많지 않고 값이 고르게 퍼져 있어서 평균으로 채워도 크게 왜곡되지 않아요.' };
    }
    return { action: ACTIONS.KEEP, badge: '채우기 주의',
      why: '빈 칸이 절반 가까이 돼요. 평균으로 채우면 이 열의 절반이 똑같은 값이 됩니다. 이 열을 꼭 써야 하는지 먼저 생각해 봐요.' };
  }
  // 글자 열도 빈 칸이 절반을 넘으면 채우기를 권하지 않는다.
  // (앱이 추천한 방법에 앱이 곧바로 경고를 띄우면 학생이 무엇을 믿어야 할지 몰라 혼란스럽다)
  if (stat.missingRate >= 0.5) {
    return { action: ACTIONS.KEEP, badge: '채우기 주의', excludeHint: true,
      why: `빈 칸이 ${pct(stat.missingRate)}예요. 채우면 이 열의 절반 이상이 똑같은 값이 됩니다. `
         + '3단계에서 이 열을 빼는 것도 좋은 선택이에요.' };
  }
  return { action: ACTIONS.TEXT,
    why: '빈 칸이었다는 사실을 그대로 남기면서 표를 깔끔하게 만들 수 있어요.' };
}

/**
 * 학생의 선택이 데이터를 크게 왜곡할 수 있으면 경고를 만든다. 막지는 않는다.
 * (막으면 학생이 실험을 못 한다. 알려 주고 결정은 학생이 한다.)
 */
export function warnFor(stat, choice) {
  const w = [];
  if (!choice) return w;
  const filling = [ACTIONS.MEAN, ACTIONS.MEDIAN, ACTIONS.TEXT, ACTIONS.CUSTOM].includes(choice.action);

  // 0으로 채우기 — 가장 흔한 오개념
  if (choice.action === ACTIONS.CUSTOM && toNumber(choice.value) === 0
      && stat.type === 'number' && stat.min !== null && stat.min > 0) {
    const after = (stat.mean * (stat.total - stat.missing)) / stat.total;
    w.push({
      level: 'warn',
      text: `0은 「없다」가 아니라 「숫자 0」이에요. 이 열의 가장 작은 값은 ${fmtNum(stat.min)}입니다. `
          + `여기에 0을 ${stat.missing}개 넣으면 평균이 ${fmtNum(roundLike(stat.mean, 1))} → `
          + `${fmtNum(roundLike(after, 1))} 으로 내려갑니다. `
          + `조사를 못 한 칸이라면 「그대로 두기」나 「행 지우기」가 더 정직해요.`,
    });
  }
  // 결측이 많은데 대표값으로 채우기
  if (filling && stat.missingRate >= 0.30) {
    w.push({
      level: 'warn',
      text: `이 열의 ${pct(stat.missingRate)}가 똑같은 값이 됩니다. 그래프를 그리면 한곳에 뭉쳐서 `
          + `실제보다 자료가 고른 것처럼 보여요. 이건 우리가 만든 모양입니다.`,
    });
  }
  // 이름표 열을 채우기
  if (filling && stat.isId) {
    w.push({
      level: 'warn',
      text: `없는 사실을 만들게 돼요. 이름·번호·코드 열의 빈 칸을 채우면 세상에 없는 기록이 표에 생깁니다. `
          + `이런 열은 빈 칸으로 두거나 그 행을 빼는 것이 맞아요.`,
    });
  }
  return w;
}


/* ═══════════════════════════════════════════════════════════════════════════
   5. 「원본 + 설정 = 결과」 다시 계산하기
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * applyPlan — 학생이 고른 설정대로 정리한 표를 만든다.
 *
 * 순서 (고정)
 *   ① 모든 열의 빈 칸을 채운다
 *   ② 「행 지우기」를 고른 열들에서 아직 빈 칸인 행을 모아 한꺼번에 지운다
 *   ③ 고른 열만 남긴다 (3단계 선택)
 *
 * @param {object} parsed   parse.js 의 parseSheet 결과 (원본, 절대 바뀌지 않는다)
 * @param {object} plan     { [열이름]: {action, value} }
 * @param {object} stats    analyzeAll 결과 (원본 기준 대표값)
 * @param {Set} activeTokens
 * @param {string[]} [pickedNames]  3단계에서 고른 열 이름 (없으면 전부)
 * @param {Set<number>} [excludeIdx] 아예 뺄 행 번호 (합계 행 등)
 * @returns {{ rows, columns, filledCells, droppedRows, keptRowIndex, log }}
 */
export function applyPlan(parsed, plan, stats, activeTokens, pickedNames, excludeIdx) {
  const columns = pickedNames
    ? parsed.columns.filter(c => pickedNames.includes(c.name))
    : parsed.columns.slice();

  const fillValueOf = (name) => {
    const ch = plan[name];
    if (!ch) return undefined;
    const st = stats[name];
    switch (ch.action) {
      case ACTIONS.MEAN:   return st && st.mean !== null ? roundLike(st.mean, st.decimals) : undefined;
      case ACTIONS.MEDIAN: return st && st.median !== null ? roundLike(st.median, st.decimals) : undefined;
      case ACTIONS.TEXT:   return ch.value === undefined || ch.value === '' ? '없음' : ch.value;
      case ACTIONS.CUSTOM: {
        if (ch.value === undefined || ch.value === '') return undefined;
        // 숫자 열이면 숫자로 넣는다 (글자로 넣으면 엑셀에서 계산이 안 된다)
        if (st && st.type === 'number') {
          const n = toNumber(ch.value);
          return n === null ? ch.value : n;
        }
        return ch.value;
      }
      default: return undefined;   // keep / dropRow 는 채우지 않는다
    }
  };

  // 채울 값을 미리 한 번만 구해 둔다 (행마다 다시 계산하면 느리다)
  const fills = new Map();
  for (const c of parsed.columns) {
    const v = fillValueOf(c.name);
    if (v !== undefined) fills.set(c.name, v);
  }
  // ★ 「행 지우기」는 **고른 열만** 보고 판단한다.
  //   쓰지 않을 열의 빈 칸 때문에 행이 사라지면 학생이 이유를 알 수 없다.
  const dropCols = columns
    .filter(c => plan[c.name] && plan[c.name].action === ACTIONS.DROP)
    .map(c => c.name);

  const rows = [];
  const keptRowIndex = [];      // 원본 rows 의 몇 번째였나 (미리보기에서 대조용)
  const filledMask = [];        // 행마다 "이 열은 우리가 채운 값" 표시
  let filledCells = 0;
  let droppedRows = 0;

  let excludedRows = 0;
  for (let i = 0; i < parsed.rows.length; i++) {
    // 학생이 "합계 줄 빼기" 를 켰으면 그 줄은 아예 없는 것으로 본다
    if (excludeIdx && excludeIdx.has(i)) { excludedRows++; continue; }
    const src = parsed.rows[i];
    const out = {};
    const mask = {};
    // ① 채우기
    for (const c of parsed.columns) {
      const v = src[c.name];
      if (isMissing(v, activeTokens)) {
        const f = fills.get(c.name);
        if (f !== undefined) { out[c.name] = f; mask[c.name] = true; filledCells++; }
        else out[c.name] = null;
      } else {
        out[c.name] = v;
      }
    }
    // ② 행 지우기 — 채운 뒤에도 빈 칸인 열이 하나라도 있으면 지운다
    let drop = false;
    for (const name of dropCols) {
      if (isMissing(out[name], activeTokens)) { drop = true; break; }
    }
    if (drop) { droppedRows++; continue; }

    // ③ 고른 열만 남기기
    const picked = {};
    for (const c of columns) picked[c.name] = out[c.name] === undefined ? null : out[c.name];
    rows.push(picked);
    filledMask.push(mask);
    keptRowIndex.push(i);
  }

  // 남은 빈 칸 세기
  let missingLeft = 0;
  for (const row of rows) {
    for (const c of columns) if (isMissing(row[c.name], activeTokens)) missingLeft++;
  }

  return { rows, columns, filledCells, droppedRows, excludedRows, keptRowIndex, filledMask, missingLeft };
}


/* ═══════════════════════════════════════════════════════════════════════════
   6. 3단계 · 속성 고르기를 돕는 판정
   ═══════════════════════════════════════════════════════════════════════════ */

/** 이 열을 쓰라고 권할 만한가? (기본 체크 상태를 정하는 데 쓴다) */
export function isRecommendedColumn(stat) {
  if (stat.type === 'empty') return false;
  if (stat.missingRate > 0.30) return false;
  if (stat.distinct <= 1) return false;
  if (stat.isPrivacy) return false;
  return true;
}

/** 속성 카드에 붙일 한 줄 코칭 문구 */
export function coachColumn(stat) {
  if (stat.isPrivacy) {
    return { kind: 'danger', text: '개인정보가 들어 있을 수 있는 열이에요. 수업 자료로는 빼는 것이 좋습니다.' };
  }
  if (stat.type === 'empty') {
    return { kind: 'warn', text: '값이 하나도 없는 열이에요.' };
  }
  if (stat.distinct <= 1) {
    return { kind: 'warn', text: '값이 모두 같아요. 비교에 도움이 되지 않습니다.' };
  }
  if (stat.missingRate >= 0.8) {
    return { kind: 'warn', text: `빈 칸이 ${pct(stat.missingRate)}나 됩니다. 빼는 쪽을 권합니다.` };
  }
  if (stat.missingRate >= 0.5) {
    return { kind: 'warn', text: `빈 칸이 ${pct(stat.missingRate)}예요. 이 열로 계산하면 결과를 믿기 어렵습니다.` };
  }
  if (stat.isId && stat.type !== 'number') {
    return { kind: 'info', text: '값이 거의 모두 달라요. 이름표 역할이라 비교에는 잘 쓰지 않아요.' };
  }
  return { kind: 'ok', text: '묶어서 비교하기 좋은 열이에요.' };
}


/* ═══════════════════════════════════════════════════════════════════════════
   7. 화면에 숫자를 예쁘게 보여 주기
   ═══════════════════════════════════════════════════════════════════════════ */

/** 1234567 → "1,234,567" */
export function fmtNum(n) {
  if (n === null || n === undefined || n === '') return '';
  if (typeof n !== 'number') {
    const p = toNumber(n);
    if (p === null) return String(n);
    n = p;
  }
  if (!Number.isFinite(n)) return String(n);
  return n.toLocaleString('ko-KR', { maximumFractionDigits: 6 });
}

/** 0.034 → "3%" (0.5% 미만은 "1% 미만") */
export function pct(rate) {
  if (!Number.isFinite(rate)) return '0%';
  if (rate > 0 && rate < 0.005) return '1% 미만';
  return Math.round(rate * 100) + '%';
}

/**
 * 셀 값을 화면에 보여 줄 글자로.
 * @param {*} v
 * @param {object} [stat]  그 열의 통계. 연도 열이면 쉼표를 붙이지 않는다.
 */
export function cellText(v, stat) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    // toISOString 은 세계시로 바꿔 버려서 하루가 밀릴 수 있다 → 그 지역 날짜로 직접 만든다
    const p = n => String(n).padStart(2, '0');
    return `${v.getFullYear()}-${p(v.getMonth() + 1)}-${p(v.getDate())}`;
  }
  if (typeof v === 'number') {
    if (stat && stat.isYear) return String(v);   // 1998 → "1998" (쉼표 없이)
    return fmtNum(v);
  }
  return String(v);
}
