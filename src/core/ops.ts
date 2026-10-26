import { isDateXf } from '../xlsx/styles';
import { addMerge, removeMerges, setColWidth, setRowHeight, shiftLayout } from './layoutOps';
import { parseA1, shiftFormula, shiftRange } from './refs';
import { applyStylePatch, styleCellsOf } from './styleOps';
import { effectiveXf } from '../grid/resolve';
import type { Axis } from './refs';
import type { Cell, Op, SheetModel, StyleTable, WorkbookModel } from './types';

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
  let prev = getCell(wb.sheets[sheet], row, col);
  // 칸이 없으면 행·열 서식을 물려받는다(R1, Excel과 같음). 빈 입력은 칸을 만들지 않는다.
  if (!prev && input !== '' && wb.kind === 'xlsx') {
    const xf = effectiveXf(wb.sheets[sheet], row, col);
    if (xf !== 0) prev = { value: null, styleId: xf };
  }
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
  target.tables = target.tables.flatMap((t) => {
    const ref = shiftRange(t.ref, axis, at, signed);
    return ref === null ? [] : [{ ...t, ref }];
  });

  shiftLayout(target.layout, axis, at, signed);

  const shift = { sheet: target.name, axis, at, count: signed };
  for (const sheet of wb.sheets) {
    for (const cols of sheet.cells.values()) {
      for (const cell of cols.values()) {
        if (cell.formula !== undefined) cell.formula = shiftFormula(cell.formula, sheet.name, shift);
      }
    }
  }
}

// R3: 서식이 그대로인 칸도 대상이면 고친 칸으로 표시한다
function applyStyle(wb: WorkbookModel, op: Extract<Op, { kind: 'style' }>): void {
  const sheet = wb.sheets[op.sheet];
  const memo = new Map<string, number>(); // 같은 (base, 가장자리)는 같은 결과
  for (const { row, col, base, pos } of styleCellsOf(sheet, op)) {
    const key = `${base}:${+pos.top}${+pos.bottom}${+pos.left}${+pos.right}`;
    let id = memo.get(key);
    if (id === undefined) {
      id = applyStylePatch(wb.styles, base, op.patch, pos);
      memo.set(key, id);
      if (isDateXf(wb.styles, wb.styles.xfs[id])) wb.dateStyleIds.add(id);
    }
    const prev = getCell(sheet, row, col);
    setCell(sheet, row, col, prev ? { ...prev, styleId: id, edited: true } : { value: null, styleId: id, edited: true });
  }
}

export function applyOp(wb: WorkbookModel, op: Op): void {
  switch (op.kind) {
    case 'set':
      setCell(wb.sheets[op.sheet], op.row, op.col, op.cell);
      return;
    case 'style':
      applyStyle(wb, op);
      return;
    case 'colWidth':
      setColWidth(wb.sheets[op.sheet].layout, op.c1, op.c2, op.width);
      return;
    case 'rowHeight':
      setRowHeight(wb.sheets[op.sheet].layout, op.rows, op.height);
      return;
    case 'merge':
      addMerge(wb.sheets[op.sheet].layout, op.range);
      return;
    case 'unmerge':
      removeMerges(wb.sheets[op.sheet].layout, op.range);
      return;
    default:
      applyStructural(wb, op);
  }
}

export function structuralBlockReason(wb: WorkbookModel, op: Op): string | null {
  if (op.kind !== 'insertRows' && op.kind !== 'deleteRows' && op.kind !== 'insertCols' && op.kind !== 'deleteCols') return null;
  const sheet = wb.sheets[op.sheet];
  if (sheet.blockers.length > 0) {
    return `이 시트에는 ${sheet.blockers.join(', ')}이(가) 있어 행·열을 추가/삭제할 수 없습니다.`;
  }
  if (op.kind === 'insertCols' || op.kind === 'deleteCols') {
    for (const { ref } of sheet.tables) {
      const [a, b] = ref.split(':');
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
    // 저장 쪽 checkTableRowDelete와 같은 규칙이어야 한다 (모델은 허용했는데 저장이 거부하면 사용자가 막힌다)
    const lo = op.at;
    const hi = op.at + op.count - 1;
    const overlaps = (x: number, y: number) => x <= y && lo <= y && hi >= x;
    for (const { ref, headerRows, totalsRows } of sheet.tables) {
      const [a, b] = ref.split(':');
      const start = parseA1(a).row;
      const end = parseA1(b ?? a).row;
      const dataLo = start + headerRows;
      const dataHi = end - totalsRows;
      const header = headerRows > 0 && overlaps(start, start + headerRows - 1);
      const totals = totalsRows > 0 && overlaps(end - totalsRows + 1, end);
      const allData = dataLo <= dataHi && lo <= dataLo && hi >= dataHi;
      if (header || totals || allData) return '표(table)의 머리글·요약 행이나 모든 데이터 행은 삭제할 수 없습니다.';
    }
  }
  return null;
}

// 배열·Map은 복사하고 안의 객체는 얕게만 복사한다(서식 항목은 읽은 뒤 고정 취급)
function cloneStyles(t: StyleTable): StyleTable {
  return {
    numFmts: new Map(t.numFmts),
    fonts: t.fonts.map((f) => ({ ...f })),
    fills: t.fills.map((f) => ({ ...f })),
    borders: t.borders.map((b) => ({ ...b })),
    xfs: t.xfs.map((x) => ({ ...x })),
    themeColors: [...t.themeColors],
    ...(t.indexedColors ? { indexedColors: [...t.indexedColors] } : {}),
    baseCounts: { ...t.baseCounts },
    ...(t.lossy
      ? { lossy: { fonts: new Set(t.lossy.fonts), fills: new Set(t.lossy.fills), borders: new Set(t.lossy.borders), xfs: new Set(t.lossy.xfs) } }
      : {}),
  };
}

export function cloneWorkbook(wb: WorkbookModel): WorkbookModel {
  return {
    kind: wb.kind,
    date1904: wb.date1904,
    dateStyleIds: new Set(wb.dateStyleIds),
    styles: cloneStyles(wb.styles),
    sheets: wb.sheets.map((s) => ({
      name: s.name,
      rowCount: s.rowCount,
      colCount: s.colCount,
      blockers: [...s.blockers],
      tables: s.tables.map((t) => ({ ...t })),
      layout: {
        ...s.layout,
        colRuns: s.layout.colRuns.map((r) => ({ ...r, info: { ...r.info } })),
        rows: new Map([...s.layout.rows].map(([k, v]) => [k, { ...v }])),
        merges: s.layout.merges.map((m) => ({ ...m })),
        freeze: { ...s.layout.freeze },
      },
      cells: new Map([...s.cells].map(([r, cols]) => [r, new Map([...cols].map(([c, cell]) => [c, { ...cell }]))])),
    })),
  };
}

export function replay(base: WorkbookModel, ops: Op[]): WorkbookModel {
  const wb = cloneWorkbook(base);
  for (const op of ops) applyOp(wb, op);
  return wb;
}
