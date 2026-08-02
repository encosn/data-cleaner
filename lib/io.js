/**
 * io.js
 * ─────────────────────────────────────────────────────────────────────────────
 * 파일을 "읽어 오는 일"과 "내려받게 하는 일"을 담당합니다.
 *
 * 여기서 가장 중요한 문제: 한글 깨짐.
 *   공공데이터포털의 CSV 파일은 대부분 CP949(EUC-KR) 로 저장되어 있습니다.
 *   윈도우 엑셀에서 "CSV(쉼표로 분리)"로 저장하면 기본이 그 방식이기 때문입니다.
 *   그런데 웹은 UTF-8 이 기본이라, 그냥 읽으면 "지역" 이 "������" 처럼 깨집니다.
 *   → 그래서 파일의 바이트를 먼저 살펴보고 어떤 방식인지 스스로 알아냅니다.
 * ─────────────────────────────────────────────────────────────────────────────
 */

import * as XLSX from 'xlsx';

/** 이 앱이 열 수 있는 파일 종류 */
export const ACCEPT = '.xlsx,.xls,.csv,.txt';
/** 교실 노트북에서 버틸 수 있는 크기 한계 */
export const MAX_BYTES = 20 * 1024 * 1024;

/**
 * 파일이 어떤 글자 방식으로 저장되었는지 알아낸다.
 *
 *  ① 맨 앞 3바이트가 EF BB BF 이면 "UTF-8 도장(BOM)"이 찍힌 UTF-8 이다.
 *  ② UTF-8 규칙에 맞는지 엄격하게(fatal) 검사해 본다. 통과하면 UTF-8.
 *  ③ 검사에서 걸리면 CP949(EUC-KR) 로 본다.
 *     브라우저의 'euc-kr' 디코더가 CP949·windows-949 를 모두 처리해 준다.
 */
export function detectEncoding(buf) {
  const u8 = new Uint8Array(buf);
  if (u8[0] === 0xEF && u8[1] === 0xBB && u8[2] === 0xBF) return 'utf-8-bom';
  // 파일이 크면 앞부분(1MB)만 검사해도 충분하다. 단 여러 바이트 글자가 잘리면
  // 잘못 판정할 수 있으므로, 잘린 끝부분은 조금 잘라내고 본다.
  const probe = u8.length > 1_000_000 ? u8.subarray(0, 1_000_000 - 4) : u8;
  try {
    new TextDecoder('utf-8', { fatal: true }).decode(probe);
    return 'utf-8';
  } catch {
    return 'euc-kr';
  }
}

/**
 * 바이트를 글자로 바꾼다.
 * @param {ArrayBuffer} buf
 * @param {string} [force]  'utf-8' | 'euc-kr' 로 강제하고 싶을 때 (자동 판정이 틀렸을 때 대비)
 */
export function decodeText(buf, force) {
  const enc = force || detectEncoding(buf);
  const label = enc === 'euc-kr' ? 'euc-kr' : 'utf-8';
  const text = new TextDecoder(label).decode(buf);
  // BOM(U+FEFF) 을 떼지 않으면 첫 열 이름이 "(보이지 않는 글자)시도" 가 되어 이름 대조가 전부 실패한다.
  // ※ 눈에 보이지 않는 글자라서 소스에 직접 쓰지 않고 \uFEFF 로 적는다 (편집기가 지워 버릴 수 있다).
  return { text: text.replace(/^\uFEFF/, ''), encoding: enc };
}

/** 글자가 깨졌는지 대략 확인한다 (자동 판정이 틀렸을 때 학생에게 알려 주기 위해) */
export function looksGarbled(text) {
  const sample = text.slice(0, 4000);
  if (sample.length === 0) return false;
  // U+FFFD(대체 문자)가 많거나, 완성형 한글이 아닌 한글 낱자가 많으면 깨진 것이다
  const bad = (sample.match(/[�]/g) || []).length
            + (sample.match(/[ᄀ-ᇿ㄰-㆏]/g) || []).length;
  return bad / sample.length > 0.02;
}

/** CSV 의 칸 구분 기호를 알아낸다. 공공데이터에는 세미콜론·탭 파일도 있다. */
export function detectDelimiter(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim() !== '').slice(0, 8);
  if (!lines.length) return ',';
  const cands = [',', ';', '\t', '|'];
  let best = ',', bestScore = -1;
  for (const d of cands) {
    // 줄마다 칸 개수를 세서, 개수가 일정하고 많을수록 좋은 구분자다
    const counts = lines.map(l => l.split(d).length);
    const avg = counts.reduce((a, b) => a + b, 0) / counts.length;
    if (avg < 2) continue;
    const variance = counts.reduce((a, b) => a + (b - avg) ** 2, 0) / counts.length;
    const score = avg - variance * 3;
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return best;
}

/**
 * 파일 하나를 읽어 워크북으로 만든다.
 * @param {File} file
 * @param {object} [opts] {encoding, delimiter}
 * @returns {Promise<{workbook, kind, encoding, delimiter, garbled, notices}>}
 */
export async function readWorkbook(file, opts = {}) {
  const notices = [];
  const name = (file.name || '').toLowerCase();
  const ext = name.slice(name.lastIndexOf('.'));
  const buf = await file.arrayBuffer();

  // 확장자만 믿지 않는다. 실제 바이트를 보고 판단한다.
  //  PK      = xlsx (zip 파일)
  //  D0CF11E0= 옛 xls
  //  '<'     = 이름만 .xls 인 HTML 표 (공공기관 웹시스템이 이렇게 내려주는 일이 있다)
  const u8 = new Uint8Array(buf);
  const isZip  = u8[0] === 0x50 && u8[1] === 0x4B;
  const isOldXls = u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0;
  const isHtml = u8[0] === 0x3C || (u8[0] === 0xEF && u8[3] === 0x3C);

  let workbook, kind, encoding = null, delimiter = null, garbled = false;

  if (isZip || isOldXls) {
    // 진짜 엑셀 파일 — 인코딩 걱정이 없다 (파일 안에 UTF-8로 들어 있다)
    kind = isZip ? 'xlsx' : 'xls';
    // dateNF 를 주지 않으면 날짜가 "7/29/26" 같은 미국식 글자로 들어와
    // 화면과 저장 파일에 그대로 나간다. 한국 수업 자료이므로 yyyy-mm-dd 로 고정한다.
    workbook = XLSX.read(buf, {
      type: 'array', cellDates: true, cellText: true, dateNF: 'yyyy-mm-dd', dense: false,
    });
    if (ext === '.csv') {
      notices.push({ level: 'info', type: '형식다름',
        message: '이름은 .csv 인데 실제로는 엑셀 파일이었어요. 엑셀 파일로 읽었습니다.' });
    }
  } else {
    // 글자 파일(CSV/TSV) 또는 HTML 표 → 인코딩을 알아내서 글자로 바꾼 뒤 읽는다
    const dec = decodeText(buf, opts.encoding);
    encoding = dec.encoding;
    garbled = looksGarbled(dec.text);
    if (isHtml) {
      kind = 'html';
      workbook = XLSX.read(dec.text, { type: 'string', cellDates: true, dateNF: 'yyyy-mm-dd' });
      notices.push({ level: 'warn', type: '형식다름',
        message: '이름은 엑셀인데 속은 웹페이지(HTML 표)였어요. 표만 뽑아서 읽었습니다.' });
    } else {
      kind = 'csv';
      delimiter = opts.delimiter || detectDelimiter(dec.text);
      workbook = XLSX.read(dec.text, {
        type: 'string', raw: false, cellDates: true, dateNF: 'yyyy-mm-dd', FS: delimiter,
      });
      if (delimiter !== ',') {
        notices.push({ level: 'info', type: '구분기호',
          message: `칸을 나누는 기호가 쉼표가 아니라 「${delimiter === '\t' ? '탭' : delimiter}」 이었어요. 그에 맞춰 읽었습니다.` });
      }
    }
    if (encoding === 'euc-kr') {
      notices.push({ level: 'info', type: '인코딩',
        message: '이 파일은 CP949(엑셀 한글) 방식으로 저장되어 있었어요. 한글이 깨지지 않게 그 방식으로 읽었습니다.' });
    }
    if (garbled) {
      notices.push({ level: 'warn', type: '글자깨짐',
        message: '글자가 깨져 보이는 곳이 있어요. 아래에서 글자 방식을 직접 바꿔 보세요.' });
    }
  }

  return { workbook, kind, encoding, delimiter, garbled, notices };
}


/* ═══════════════════════════════════════════════════════════════════════════
   내려받기
   ═══════════════════════════════════════════════════════════════════════════ */

/** 엑셀 시트 이름에 쓸 수 없는 글자를 걸러내고 31자로 줄인다 (엑셀 규칙) */
export function safeSheetName(name) {
  const s = String(name || '시트1').replace(/[\\/?*[\]:]/g, ' ').trim();
  return (s === '' ? '시트1' : s).slice(0, 31);
}

/** 파일 이름에 쓸 수 없는 글자가 있는지 */
export function badFileNameChars(name) {
  const m = String(name || '').match(/[\\/:*?"<>|]/g);
  return m ? [...new Set(m)] : [];
}

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  // 바로 지우면 브라우저가 저장을 못 끝내는 경우가 있어 조금 기다린다
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

/**
 * 2차원 배열을 엑셀 파일로 저장한다.
 * @param {Array<Array>} aoa      첫 줄이 열 이름
 * @param {object} opts { fileName, sheetName, reportLines }
 */
export function saveXlsx(aoa, opts = {}) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // 열 너비를 내용에 맞춰 잡아 준다. 안 하면 열이 좁아서 ###### 로 보인다.
  const nCols = aoa.length ? Math.max(...aoa.map(r => r.length)) : 0;
  const cols = [];
  for (let c = 0; c < nCols; c++) {
    let w = 8;
    for (let r = 0; r < Math.min(aoa.length, 200); r++) {
      const v = aoa[r][c];
      if (v === null || v === undefined) continue;
      // 한글은 폭을 2배로 계산해야 화면에서 맞는다
      const s = String(v);
      const len = s.length + (s.match(/[가-힣ㄱ-ㅎㅏ-ㅣ]/g) || []).length;
      if (len > w) w = len;
    }
    cols.push({ wch: Math.min(w + 2, 40) });
  }
  ws['!cols'] = cols;
  XLSX.utils.book_append_sheet(wb, ws, safeSheetName(opts.sheetName || '정리한자료'));

  // "내가 한 일" 을 두 번째 시트로 함께 넣는다.
  // → 채워 넣은 값과 원래 값을 구분할 수 있게 기록을 남기는 것이 정직한 데이터 처리다.
  if (Array.isArray(opts.reportLines) && opts.reportLines.length) {
    const rep = [['내가 한 일'], ...opts.reportLines.map(l => [l])];
    const ws2 = XLSX.utils.aoa_to_sheet(rep);
    ws2['!cols'] = [{ wch: 90 }];
    XLSX.utils.book_append_sheet(wb, ws2, '정제기록');
  }

  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
           (opts.fileName || '정리한자료') + '.xlsx');
}

/**
 * CSV 로 저장한다.
 * ★ 맨 앞에 BOM(\uFEFF, UTF-8 도장)을 반드시 붙인다.
 *   안 붙이면 학생이 집에서 엑셀로 열었을 때 한글이 다시 깨진다.
 */
export function saveCsv(aoa, opts = {}) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const csv = XLSX.utils.sheet_to_csv(ws);
  download(new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8' }),
           (opts.fileName || '정리한자료') + '.csv');
}
