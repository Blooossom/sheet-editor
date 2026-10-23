import type { Cell, WorkbookModel } from '../core/types';

export interface CsvDialect {
  delimiter: ',' | '\t' | ';' | '|';
  newline: '\n' | '\r\n';
  quoteAll: boolean; // 비어 있지 않은 필드를 모두 따옴표로 감쌌는가
  quoteEmpty: boolean; // 빈 필드를 "" 로 썼는가
  trailingNewline: boolean;
}

export const DEFAULT_DIALECT: CsvDialect = {
  delimiter: ',',
  newline: '\n',
  quoteAll: false,
  quoteEmpty: false,
  trailingNewline: true,
};

const CANDIDATES: CsvDialect['delimiter'][] = [',', '\t', ';', '|'];
const SNIFF_LINES = 20;

interface Field {
  value: string;
  quoted: boolean;
}

// 따옴표 안의 구분자·줄바꿈은 데이터로 취급한다. 따옴표는 필드 맨 앞에서만 여는 따옴표로 본다.
function scan(text: string, delimiter: string): Field[][] {
  const rows: Field[][] = [];
  let row: Field[] = [];
  let value = '';
  let quoted = false;
  let started = false; // 현재 필드가 시작됐는가 (빈 따옴표 필드 포함)
  let inQuotes = false;

  const endField = () => {
    row.push({ value, quoted });
    value = '';
    quoted = false;
    started = false;
  };
  const endRow = () => {
    endField();
    rows.push(row);
    row = [];
  };

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          value += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        value += ch;
      }
    } else if (ch === '"' && !started) {
      inQuotes = true;
      quoted = true;
      started = true;
    } else if (ch === delimiter) {
      endField();
    } else if (ch === '\n') {
      endRow();
    } else if (ch === '\r' && text[i + 1] === '\n') {
      endRow();
      i++;
    } else {
      value += ch;
      started = true;
    }
  }
  if (started || value !== '' || row.length > 0) endRow();
  return rows;
}

export function detectDialect(text: string): CsvDialect {
  if (text === '') return { ...DEFAULT_DIALECT, trailingNewline: false };

  // 따옴표 밖 줄바꿈 기준으로 앞의 20줄에서 후보 구분자 개수를 센다. 빈 줄은 건너뛴다.
  let newline: CsvDialect['newline'] | null = null;
  const lines: Record<string, number>[] = [];
  let cur: Record<string, number> = {};
  let curLen = 0;
  let inQuotes = false;
  const endLine = () => {
    if (curLen > 0) lines.push(cur);
    cur = {};
    curLen = 0;
  };
  for (let i = 0; i < text.length && lines.length < SNIFF_LINES; i++) {
    const ch = text[i];
    if (ch === '"') inQuotes = !inQuotes;
    if (inQuotes) {
      curLen++;
    } else if (ch === '\n') {
      if (newline === null) newline = text[i - 1] === '\r' ? '\r\n' : '\n';
      endLine();
    } else if (ch !== '\r') {
      curLen++;
      if (CANDIDATES.includes(ch as CsvDialect['delimiter'])) cur[ch] = (cur[ch] ?? 0) + 1;
    }
  }
  if (lines.length < SNIFF_LINES) endLine();

  let delimiter: CsvDialect['delimiter'] = ',';
  let best = 0;
  for (const cand of CANDIDATES) {
    const n = lines[0]?.[cand] ?? 0;
    if (n > best && lines.every((c) => (c[cand] ?? 0) === n)) {
      best = n;
      delimiter = cand;
    }
  }

  const rows = scan(text, delimiter);
  let nonEmpty = 0;
  let nonEmptyQuoted = 0;
  let empty = 0;
  let emptyQuoted = 0;
  for (const r of rows) {
    if (r.length === 1 && r[0].value === '' && !r[0].quoted) continue; // 빈 줄
    for (const f of r) {
      if (f.value === '') {
        empty++;
        if (f.quoted) emptyQuoted++;
      } else {
        nonEmpty++;
        if (f.quoted) nonEmptyQuoted++;
      }
    }
  }

  return {
    delimiter,
    newline: newline ?? '\n',
    quoteAll: nonEmpty > 0 && nonEmptyQuoted === nonEmpty,
    quoteEmpty: empty > 0 && emptyQuoted === empty,
    trailingNewline: text.endsWith('\n'),
  };
}

export function parseCsv(text: string, d: CsvDialect): string[][] {
  return scan(text, d.delimiter).map((r) => r.map((f) => f.value));
}

export function serializeCsv(rows: string[][], d: CsvDialect): string {
  const needsQuote = (v: string) =>
    v.includes(d.delimiter) || v.includes('"') || v.includes('\r') || v.includes('\n');
  const field = (v: string) => {
    if (v === '') return d.quoteEmpty ? '""' : '';
    if (d.quoteAll || needsQuote(v)) return `"${v.replace(/"/g, '""')}"`;
    return v;
  };
  if (rows.length === 0) return '';
  const body = rows.map((r) => r.map(field).join(d.delimiter)).join(d.newline);
  return d.trailingNewline ? body + d.newline : body;
}

export function csvToWorkbook(rows: string[][], name: string): WorkbookModel {
  const cells = new Map<number, Map<number, Cell>>();
  let colCount = 0;
  rows.forEach((r, ri) => {
    colCount = Math.max(colCount, r.length);
    r.forEach((v, ci) => {
      if (v === '') return;
      let rowMap = cells.get(ri);
      if (!rowMap) cells.set(ri, (rowMap = new Map()));
      rowMap.set(ci, { value: v, text: true });
    });
  });
  return {
    kind: 'csv',
    sheets: [{ name, cells, rowCount: rows.length, colCount, blockers: [], tableRanges: [] }],
    dateStyleIds: new Set(),
    date1904: false,
  };
}

export function workbookToCsvRows(wb: WorkbookModel): string[][] {
  const sheet = wb.sheets[0];
  const rows: string[][] = [];
  for (let r = 0; r < sheet.rowCount; r++) {
    const rowMap = sheet.cells.get(r);
    const row: string[] = [];
    for (let c = 0; c < sheet.colCount; c++) {
      const v = rowMap?.get(c)?.value ?? null;
      row.push(v === null ? '' : String(v));
    }
    rows.push(row);
  }
  return rows;
}
