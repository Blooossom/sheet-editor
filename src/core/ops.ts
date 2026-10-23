import { parseA1, shiftFormula, shiftRange } from './refs';
import type { Axis } from './refs';
import type { Cell, Op, SheetModel, WorkbookModel } from './types';

interface ParseCtx {
  kind: 'xlsx' | 'csv';
  prev?: Cell;
  isDate: boolean;
  date1904: boolean;
}

const NUMBER_RE = /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/;
const DATE_RE = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2})(?::(\d{2}))?)?$/;
const DAY_MS = 86400000;

/** 날짜 문자열 → 일련번호. 형식이 아니거나 존재하지 않는 날짜면 null. UTC로 계산해 DST 영향이 없다. */
function parseDate(input: string, date1904: boolean): number | null {
  const m = DATE_RE.exec(input);
  if (!m) return null;
  const [y, mo, d, h = 0, mi = 0, s = 0] = m.slice(1).map((x) => (x === undefined ? 0 : Number(x)));
  const ms = Date.UTC(y, mo - 1, d);
  const check = new Date(ms);
  if (check.getUTCFullYear() !== y || check.getUTCMonth() !== mo - 1 || check.getUTCDate() !== d) return null;
  if (h > 23 || mi > 59 || s > 59) return null;
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (ms - epoch) / DAY_MS + (h * 3600 + mi * 60 + s) / 86400;
}

export function parseInput(input: string, ctx: ParseCtx): Cell | null {
  const styleId = ctx.prev?.styleId;
  const withStyle = (cell: Cell): Cell => {
    if (styleId !== undefined) cell.styleId = styleId;
    cell.edited = true;
    return cell;
  };
  if (input === '') return styleId !== undefined ? { value: null, styleId, edited: true } : null;
  if (ctx.kind === 'csv') return { value: input, text: true, edited: true };
  if (input.startsWith('=') && input.length > 1) return withStyle({ formula: input.slice(1), value: null });
  // 수식 셀의 문자열 결과(text)는 상속하지 않는다
  const prevIsText = ctx.prev?.text === true && ctx.prev.formula === undefined;
  if (prevIsText || /^-?0\d/.test(input) || /^-?\d{16,}/.test(input)) {
    return withStyle({ value: input, text: true });
  }
  if (ctx.isDate) {
    const serial = parseDate(input, ctx.date1904);
    if (serial !== null) return withStyle({ value: serial });
  }
  if (NUMBER_RE.test(input)) {
    const n = Number(input);
    if (Number.isFinite(n)) return withStyle({ value: n === 0 ? 0 : n }); // -0 → 0
  }
  return withStyle({ value: input, text: true });
}

export function getCell(sheet: SheetModel, row: number, col: number): Cell | undefined {
  return sheet.cells.get(row)?.get(col);
}

export function makeSetOp(wb: WorkbookModel, sheet: number, row: number, col: number, input: string): Op {
  const prev = getCell(wb.sheets[sheet], row, col);
  const isDate = prev?.styleId !== undefined && wb.dateStyleIds.has(prev.styleId);
  const cell = parseInput(input, { kind: wb.kind, prev, isDate, date1904: wb.date1904 });
  return { kind: 'set', sheet, row, col, cell };
}

function setCell(sheet: SheetModel, row: number, col: number, cell: Cell | null): void {
  if (cell === null) {
    const r = sheet.cells.get(row);
    if (r) {
      r.delete(col);
      if (r.size === 0) sheet.cells.delete(row);
    }
    return;
  }
  let r = sheet.cells.get(row);
  if (!r) sheet.cells.set(row, (r = new Map()));
  r.set(col, { ...cell }); // op 객체와 모델이 섞이지 않게 복사
  sheet.rowCount = Math.max(sheet.rowCount, row + 1);
  sheet.colCount = Math.max(sheet.colCount, col + 1);
}

/** 이동 후 인덱스. 삭제된 구간이면 null. */
function mapIndex(i: number, at: number, count: number, insert: boolean): number | null {
  if (i < at) return i;
  if (insert) return i + count;
  if (i < at + count) return null;
  return i - count;
}

function applyStructural(wb: WorkbookModel, op: Extract<Op, { at: number }>): void {
  const target = wb.sheets[op.sheet];
  const axis: Axis = op.kind === 'insertRows' || op.kind === 'deleteRows' ? 'row' : 'col';
  const insert = op.kind === 'insertRows' || op.kind === 'insertCols';
  const { at, count } = op;
  if (!(count >= 1)) throw new Error(`count는 1 이상이어야 합니다: ${count}`);

  const moved = new Map<number, Map<number, Cell>>();
  for (const [r, cols] of target.cells) {
    for (const [c, cell] of cols) {
      const nr = axis === 'row' ? mapIndex(r, at, count, insert) : r;
      const nc = axis === 'col' ? mapIndex(c, at, count, insert) : c;
      if (nr === null || nc === null) continue;
      let row = moved.get(nr);
      if (!row) moved.set(nr, (row = new Map()));
      row.set(nc, cell);
    }
  }
  target.cells = moved;

  const key = axis === 'row' ? 'rowCount' : 'colCount';
  const n = target[key];
  if (insert) {
    if (at < n) target[key] = n + count;
  } else {
    target[key] = n - Math.max(0, Math.min(n, at + count) - at);
  }

  const signed = insert ? count : -count; // refs.ts: 삭제는 음수
  target.tableRanges = target.tableRanges
    .map((range) => shiftRange(range, axis, at, signed))
    .filter((range): range is string => range !== null);

  const shift = { sheet: target.name, axis, at, count: signed };
  for (const sheet of wb.sheets) {
    for (const cols of sheet.cells.values()) {
      for (const cell of cols.values()) {
        if (cell.formula !== undefined) cell.formula = shiftFormula(cell.formula, sheet.name, shift);
      }
    }
  }
}

export function applyOp(wb: WorkbookModel, op: Op): void {
  if (op.kind === 'set') setCell(wb.sheets[op.sheet], op.row, op.col, op.cell);
  else applyStructural(wb, op);
}

export function structuralBlockReason(wb: WorkbookModel, op: Op): string | null {
  if (op.kind === 'set') return null;
  const sheet = wb.sheets[op.sheet];
  if (sheet.blockers.length > 0) {
    return `이 시트에는 ${sheet.blockers.join(', ')}이(가) 있어 행·열을 추가/삭제할 수 없습니다.`;
  }
  if (op.kind === 'insertCols' || op.kind === 'deleteCols') {
    for (const range of sheet.tableRanges) {
      const [a, b] = range.split(':');
      const start = parseA1(a).col;
      const end = parseA1(b ?? a).col;
      const overlaps =
        op.kind === 'insertCols'
          ? op.at > start && op.at <= end
          : op.at <= end && op.at + op.count - 1 >= start;
      if (overlaps) return '표(table) 범위 안에서는 열을 추가/삭제할 수 없습니다.';
    }
  }
  if (op.kind === 'deleteRows') {
    const lo = op.at;
    const hi = op.at + op.count - 1;
    for (const range of sheet.tableRanges) {
      const [a, b] = range.split(':');
      const start = parseA1(a).row; // 머리글 행 1개로 가정
      const end = parseA1(b ?? a).row;
      const header = lo <= start && start <= hi;
      const allData = end > start && lo <= start + 1 && hi >= end;
      if (header || allData) return '표(table)의 머리글 행이나 모든 데이터 행은 삭제할 수 없습니다.';
    }
  }
  return null;
}

export function cloneWorkbook(wb: WorkbookModel): WorkbookModel {
  return {
    kind: wb.kind,
    date1904: wb.date1904,
    dateStyleIds: new Set(wb.dateStyleIds),
    sheets: wb.sheets.map((s) => ({
      name: s.name,
      rowCount: s.rowCount,
      colCount: s.colCount,
      blockers: [...s.blockers],
      tableRanges: [...s.tableRanges],
      cells: new Map([...s.cells].map(([r, cols]) => [r, new Map([...cols].map(([c, cell]) => [c, { ...cell }]))])),
    })),
  };
}

export function replay(base: WorkbookModel, ops: Op[]): WorkbookModel {
  const wb = cloneWorkbook(base);
  for (const op of ops) applyOp(wb, op);
  return wb;
}
