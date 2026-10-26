import { emptyLayout } from '../xlsx/layout';
import { defaultStyleTable } from '../xlsx/styles';
import { applyOp, cloneWorkbook, getCell, makeSetOp, replay, structuralBlockReason } from '../core/ops';
import { displayText, editText } from '../core/display';
import type { Cell, CellRange, Op, StructuralKind, StylePatch, WorkbookModel } from '../core/types';
import { parseA1 } from '../core/refs';
import { maxDigitWidth, pxToColWidth, pxToPt } from '../grid/metrics';
import { DEFAULT_DIALECT, csvToWorkbook, detectDialect, parseCsv, serializeCsv, workbookToCsvRows } from '../csv/csv';
import type { CsvDialect } from '../csv/csv';
import { recalcWorkbook } from '../formula/recalc';
import { readXlsx } from '../xlsx/read';
import { saveXlsx } from '../xlsx/save';
import { createXlsx } from '../xlsx/create';
import type { DocumentIo } from './io';

export type Source =
  | { format: 'xlsx'; bytes: Uint8Array }
  | { format: 'csv'; dialect: CsvDialect; encoding: 'utf-8' | 'euc-kr'; bom: boolean }
  | { format: 'new' };

const CSV_BLOCK = '이 형식(csv)은 서식을 저장할 수 없습니다.';
const MAX_FORMAT_CELLS = 200000;
const MAX_COL_CHARS = 255; // Excel 한계
const MAX_ROW_PT = 409.5;
const FORMAT_CAP_MESSAGE = '한 번에 서식을 바꿀 수 있는 칸은 200,000개까지입니다.';

function formatOf(path: string): 'xlsx' | 'csv' {
  const m = /\.([^./\\]+)$/.exec(path);
  const ext = m ? m[1].toLowerCase() : '';
  if (ext === 'xlsx' || ext === 'csv') return ext;
  throw new Error('지원하지 않는 파일 형식입니다');
}

function baseName(path: string): string {
  const file = path.split(/[/\\]/).pop() ?? path;
  return file.replace(/\.[^.]*$/, '');
}

// tauri invoke는 Error가 아닌 문자열로 reject 한다
function asError(e: unknown): Error {
  return e instanceof Error ? e : new Error(String(e));
}

function clearEdited(wb: WorkbookModel): void {
  for (const sheet of wb.sheets) {
    for (const cols of sheet.cells.values()) for (const cell of cols.values()) delete cell.edited;
  }
}

function emptyWorkbook(): WorkbookModel {
  return {
    kind: 'xlsx',
    sheets: [{ name: 'Sheet1', cells: new Map(), rowCount: 0, colCount: 0, blockers: [], tables: [], layout: emptyLayout() }],
    dateStyleIds: new Set(),
    date1904: false,
    styles: defaultStyleTable(),
  };
}

/** 연속 구간으로 묶는다. 오름차순. */
function runs(indices: number[]): { start: number; count: number }[] {
  const sorted = [...new Set(indices)].sort((a, b) => a - b);
  const out: { start: number; count: number }[] = [];
  for (const i of sorted) {
    const last = out[out.length - 1];
    if (last && last.start + last.count === i) last.count++;
    else out.push({ start: i, count: 1 });
  }
  return out;
}

// 새 문서를 csv로 저장한 뒤의 모델: 모든 셀이 표시 문자열
function toCsvModel(wb: WorkbookModel): WorkbookModel {
  const out = cloneWorkbook(wb);
  out.kind = 'csv';
  out.dateStyleIds = new Set();
  out.styles = defaultStyleTable();
  for (const [si, sheet] of out.sheets.entries()) {
    sheet.layout = emptyLayout(); // csv는 열 너비·병합 등을 저장하지 않는다
    for (const [r, cols] of [...sheet.cells]) {
      for (const c of [...cols.keys()]) {
        const text = displayText(wb.sheets[si].cells.get(r)?.get(c), wb);
        if (text === '') cols.delete(c);
        else cols.set(c, { value: text, text: true });
      }
      if (cols.size === 0) sheet.cells.delete(r);
    }
  }
  return out;
}

export class DocumentStore {
  path: string | null = null;
  source: Source = { format: 'new' };
  wb: WorkbookModel = emptyWorkbook();
  version = 0;

  private base: WorkbookModel = emptyWorkbook();
  private steps: Op[][] = [];
  private redoStack: Op[][] = [];
  private listeners = new Set<() => void>();
  private _saving = false;

  constructor(private io: DocumentIo) {}

  get dirty(): boolean {
    return this.steps.length > 0;
  }
  get saving(): boolean {
    return this._saving;
  }
  get canUndo(): boolean {
    return this.steps.length > 0;
  }
  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  subscribe(fn: () => void): () => void {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }

  private notify(): void {
    for (const fn of [...this.listeners]) {
      try {
        fn();
      } catch (e) {
        console.error(e);
      }
    }
  }

  private changed(): void {
    this.version++;
    this.notify();
  }

  private assertIdle(): void {
    if (this._saving) throw new Error('저장 중에는 편집할 수 없습니다.');
  }

  private install(path: string | null, source: Source, wb: WorkbookModel): void {
    recalcWorkbook(wb); // 던지면 기존 상태를 건드리지 않는다
    this.path = path;
    this.source = source;
    this.wb = wb;
    this.base = cloneWorkbook(wb);
    this.steps = [];
    this.redoStack = [];
    this.changed();
  }

  async load(path: string): Promise<void> {
    this.assertIdle();
    const format = formatOf(path);
    try {
      if (format === 'xlsx') {
        const bytes = await this.io.readFile(path);
        this.install(path, { format: 'xlsx', bytes }, readXlsx(bytes));
      } else {
        const file = await this.io.readText(path);
        const dialect = detectDialect(file.text);
        const wb = csvToWorkbook(parseCsv(file.text, dialect), baseName(path));
        this.install(path, { format: 'csv', dialect, encoding: file.encoding, bom: file.bom }, wb);
      }
    } catch (e) {
      throw asError(e);
    }
  }

  newDocument(): void {
    this.assertIdle();
    this.install(null, { format: 'new' }, emptyWorkbook());
  }

  private rebuild(): void {
    this.wb = replay(this.base, this.steps.flat());
    recalcWorkbook(this.wb);
  }

  setCells(edits: { sheet: number; row: number; col: number; input: string }[]): void {
    this.assertIdle();
    const ops: Op[] = [];
    try {
      for (const e of edits) {
        const current = getCell(this.wb.sheets[e.sheet], e.row, e.col);
        // editText가 왕복하지 않는 값(소수 초 날짜, 15자리 초과 숫자, '=' 시작 문자열)도 입력이 같으면 건너뛴다
        if (e.input === editText(current, this.wb)) continue;
        const op = makeSetOp(this.wb, e.sheet, e.row, e.col, e.input);
        applyOp(this.wb, op);
        ops.push(op);
      }
    } catch (err) {
      this.rebuild();
      this.changed();
      throw err;
    }
    if (ops.length === 0) return;
    this.steps.push(ops);
    this.redoStack = [];
    recalcWorkbook(this.wb);
    this.changed();
  }

  structural(sheet: number, kind: StructuralKind, rows: number[]): string | null {
    this.assertIdle();
    const groups = runs(rows);
    if (groups.length === 0) return null;
    const ops: Op[] =
      kind === 'insertRows' || kind === 'insertCols'
        ? [{ kind, sheet, at: groups[0].start, count: new Set(rows).size }]
        : groups.reverse().map((g) => ({ kind, sheet, at: g.start, count: g.count }));
    // 모든 연산을 복사본에 차례로 적용하며 차단 여부를 먼저 확인한다
    const trial = cloneWorkbook(this.wb);
    for (const op of ops) {
      const reason = structuralBlockReason(trial, op);
      if (reason !== null) return reason;
      applyOp(trial, op);
    }
    this.wb = trial;
    this.steps.push(ops);
    this.redoStack = [];
    recalcWorkbook(this.wb);
    this.changed();
    return null;
  }

  /** ops를 한 단계로 적용한다. 도중에 던지면 단계를 만들지 않고 이전 상태로 되돌린다. */
  private commit(ops: Op[], recalc = false): void {
    try {
      for (const op of ops) applyOp(this.wb, op);
    } catch (err) {
      this.rebuild();
      this.changed();
      throw err;
    }
    this.steps.push(ops);
    this.redoStack = [];
    if (recalc) recalcWorkbook(this.wb);
    this.changed();
  }

  /**
   * 서식 적용. 차단 이유 또는 null.
   * - rows는 원래 행 번호. 필터면 보이는 행만 넘어오므로 그 행에만 적용한다.
   * - rows·c2가 'all'이면 사용 범위(rowCount·colCount)까지만. 빈 결과면 단계를 만들지 않는다.
   * - 병합 일부를 덮으면 열은 c1·c2를 넓힌다. 행은 rows가 연속일 때(필터 없음)만 병합 행까지 더한다. 필터 중에는 안 보이는 행을 건드리지 않는다.
   */
  formatCells(sheet: number, sel: { rows: number[] | 'all'; c1: number; c2: number | 'all' }, patch: StylePatch): string | null {
    this.assertIdle();
    if (this.wb.kind === 'csv') return CSV_BLOCK;
    if (Object.keys(patch).length === 0) return null;
    const sh = this.wb.sheets[sheet];
    // 'all'은 사용 범위(rowCount·colCount)로 줄인다. 필터 중에는 UI가 보이는 원래 행 번호 배열을 넘긴다.
    let rows = sel.rows === 'all' ? Array.from({ length: sh.rowCount }, (_, i) => i) : [...new Set(sel.rows)].sort((a, b) => a - b);
    let c1 = sel.c1;
    let c2 = sel.c2 === 'all' ? sh.colCount - 1 : sel.c2;
    if (rows.length === 0 || c2 < c1) return null;

    const contiguous = rows[rows.length - 1] - rows[0] + 1 === rows.length;
    for (let changed = true; changed; ) {
      changed = false;
      for (const m of sh.layout.merges) {
        if (m.c2 < c1 || m.c1 > c2 || !rows.some((r) => r >= m.r1 && r <= m.r2)) continue;
        if (m.c1 < c1 || m.c2 > c2) {
          c1 = Math.min(c1, m.c1);
          c2 = Math.max(c2, m.c2);
          changed = true;
        }
        if (contiguous && (m.r1 < rows[0] || m.r2 > rows[rows.length - 1])) {
          const lo = Math.min(rows[0], m.r1);
          const hi = Math.max(rows[rows.length - 1], m.r2);
          rows = Array.from({ length: hi - lo + 1 }, (_, i) => lo + i);
          changed = true;
        }
      }
    }
    if (rows.length * (c2 - c1 + 1) > MAX_FORMAT_CELLS) return FORMAT_CAP_MESSAGE;
    this.commit([{ kind: 'style', sheet, rows, c1, c2, patch }]);
    return null;
  }

  /** 열 너비(px). null = 기본 너비. 연속 구간마다 Op 하나, 전체가 한 단계. */
  setColWidth(sheet: number, cols: number[], width: number | null): string | null {
    this.assertIdle();
    if (this.wb.kind === 'csv') return CSV_BLOCK;
    if (width !== null && !(Number.isFinite(width) && width > 0)) return '올바른 크기가 아닙니다.';
    const groups = runs(cols);
    if (groups.length === 0) return null;
    let chars: number | null = null;
    if (width !== null) {
      const mdw = maxDigitWidth(this.wb.styles.fonts[0]);
      chars = Math.min(pxToColWidth(width, mdw), MAX_COL_CHARS);
    }
    this.commit(groups.map((g) => ({ kind: 'colWidth', sheet, c1: g.start, c2: g.start + g.count - 1, width: chars })));
    return null;
  }

  /** 행 높이(px). null = 기본 높이. */
  setRowHeight(sheet: number, rows: number[], height: number | null): string | null {
    this.assertIdle();
    if (this.wb.kind === 'csv') return CSV_BLOCK;
    if (height !== null && !(Number.isFinite(height) && height > 0)) return '올바른 크기가 아닙니다.';
    const sorted = [...new Set(rows)].sort((a, b) => a - b);
    if (sorted.length === 0) return null;
    this.commit([{ kind: 'rowHeight', sheet, rows: sorted, height: height === null ? null : Math.min(pxToPt(height), MAX_ROW_PT) }]);
    return null;
  }

  /** 범위와 일부만 겹치는 병합은 통째로 포함하도록 넓힌다(addMerge가 겹친 병합을 통째로 지우므로). */
  private mergeTarget(sheet: number, range: CellRange): CellRange {
    const merges = this.wb.sheets[sheet].layout.merges;
    const t = { ...range };
    for (let changed = true; changed; ) {
      changed = false;
      for (const m of merges) {
        if (m.r2 < t.r1 || m.r1 > t.r2 || m.c2 < t.c1 || m.c1 > t.c2) continue;
        const next = { r1: Math.min(t.r1, m.r1), c1: Math.min(t.c1, m.c1), r2: Math.max(t.r2, m.r2), c2: Math.max(t.c2, m.c2) };
        if (next.r1 !== t.r1 || next.c1 !== t.c1 || next.r2 !== t.r2 || next.c2 !== t.c2) {
          Object.assign(t, next);
          changed = true;
        }
      }
    }
    return t;
  }

  /** rows: 화면에 보이는 원래 행 번호. 'all' = 필터 없음. 넓힌 대상의 행을 빠짐없이 덮지 않으면 필터 중으로 본다. */
  mergeCheck(sheet: number, range: CellRange, rows: number[] | 'all'): { ok: true; lossy: boolean } | { ok: false; reason: string } {
    if (this.wb.kind === 'csv') return { ok: false, reason: CSV_BLOCK };
    const sh = this.wb.sheets[sheet];
    const t = this.mergeTarget(sheet, range);
    if (rows !== 'all') {
      const seen = new Set<number>();
      for (const r of rows) if (r >= t.r1 && r <= t.r2) seen.add(r);
      if (seen.size !== t.r2 - t.r1 + 1) return { ok: false, reason: '필터를 해제한 뒤 병합하세요.' };
    }
    for (const table of sh.tables) {
      const [a, b] = table.ref.split(':');
      const s = parseA1(a);
      const e = parseA1(b ?? a);
      if (!(e.row < t.r1 || s.row > t.r2 || e.col < t.c1 || s.col > t.c2)) {
        return { ok: false, reason: '표(table) 안에서는 병합할 수 없습니다.' };
      }
    }
    return { ok: true, lossy: this.nonAnchorCells(sheet, t).length > 0 };
  }

  /** 앵커(왼쪽 위)를 뺀 값·수식이 있는 칸 */
  private nonAnchorCells(sheet: number, t: CellRange): { row: number; col: number; cell: Cell }[] {
    const out: { row: number; col: number; cell: Cell }[] = [];
    for (const [row, cols] of this.wb.sheets[sheet].cells) {
      if (row < t.r1 || row > t.r2) continue;
      for (const [col, cell] of cols) {
        if (col < t.c1 || col > t.c2 || (row === t.r1 && col === t.c1)) continue;
        if (cell.value !== null || cell.formula !== undefined) out.push({ row, col, cell });
      }
    }
    return out.sort((a, b) => a.row - b.row || a.col - b.col);
  }

  merge(sheet: number, range: CellRange, rows: number[] | 'all'): string | null {
    this.assertIdle();
    const check = this.mergeCheck(sheet, range, rows);
    if (!check.ok) return check.reason;
    const t = this.mergeTarget(sheet, range);
    if (t.r1 === t.r2 && t.c1 === t.c2) return null;
    if (this.wb.sheets[sheet].layout.merges.some((m) => m.r1 === t.r1 && m.c1 === t.c1 && m.r2 === t.r2 && m.c2 === t.c2)) return null;
    // 값은 지우되 서식은 남긴다. 칸 항목을 없애면 행·열 서식으로 바뀌어 보이므로(R1) styleId 0도 명시한다.
    const ops: Op[] = this.nonAnchorCells(sheet, t).map(({ row, col, cell }) => ({
      kind: 'set',
      sheet,
      row,
      col,
      cell: { value: null, styleId: cell.styleId ?? 0, edited: true },
    }));
    ops.push({ kind: 'merge', sheet, range: t });
    this.commit(ops, true);
    return null;
  }

  unmerge(sheet: number, range: CellRange): string | null {
    this.assertIdle();
    if (this.wb.kind === 'csv') return CSV_BLOCK;
    const hit = this.wb.sheets[sheet].layout.merges.some((m) => !(m.r2 < range.r1 || m.r1 > range.r2 || m.c2 < range.c1 || m.c1 > range.c2));
    if (!hit) return null;
    this.commit([{ kind: 'unmerge', sheet, range }]);
    return null;
  }

  /** 현재 상태가 기본 서식 외 서식(칸 서식·열 너비·행 높이·병합)을 쓰는지. csv는 항상 false. */
  get hasFormatting(): boolean {
    if (this.wb.kind === 'csv') return false;
    for (const sh of this.wb.sheets) {
      const l = sh.layout;
      if (l.merges.length > 0 || l.rows.size > 0 || l.colRuns.length > 0) return true;
      for (const cols of sh.cells.values()) for (const cell of cols.values()) if (cell.styleId) return true;
    }
    return false;
  }

  /** 되돌린 단계의 연산들(원래 적용된 방향). 되돌릴 것이 없으면 null. */
  undo(): Op[] | null {
    this.assertIdle();
    const step = this.steps.pop();
    if (!step) return null;
    this.redoStack.push(step);
    this.rebuild();
    this.changed();
    return step;
  }

  /** 다시 적용한 단계의 연산들. 다시 할 것이 없으면 null. */
  redo(): Op[] | null {
    this.assertIdle();
    const step = this.redoStack.pop();
    if (!step) return null;
    this.steps.push(step);
    this.rebuild();
    this.changed();
    return step;
  }

  async save(newPath?: string): Promise<void> {
    this.assertIdle();
    if (!this.dirty && this.source.format !== 'new') return; // 고친 것이 없으면 파일을 다시 쓰지 않는다
    this._saving = true;
    this.notify();
    try {
      await this.doSave(newPath);
    } catch (e) {
      throw asError(e);
    } finally {
      this._saving = false;
      this.notify();
    }
  }

  private async doSave(newPath?: string): Promise<void> {
    const src = this.source;
    let nextPath = this.path;
    let nextSource: Source = src;
    let nextWb = this.wb;
    if (src.format === 'new') {
      if (!newPath) throw new Error('저장할 경로가 필요합니다');
      if (formatOf(newPath) === 'xlsx') {
        const bytes = createXlsx(this.wb);
        await this.io.writeFile(newPath, bytes);
        nextSource = { format: 'xlsx', bytes };
      } else {
        nextWb = toCsvModel(this.wb); // 파일과 메모리가 같은 표시 문자열을 갖도록 먼저 변환한다
        const text = serializeCsv(workbookToCsvRows(nextWb), DEFAULT_DIALECT);
        await this.io.writeText(newPath, { text, encoding: 'utf-8', bom: false });
        nextSource = { format: 'csv', dialect: DEFAULT_DIALECT, encoding: 'utf-8', bom: false };
      }
      nextPath = newPath;
    } else if (src.format === 'xlsx') {
      if (!this.path) throw new Error('저장할 경로가 없습니다');
      const bytes = saveXlsx(src.bytes, this.steps.flat());
      await this.io.writeFile(this.path, bytes);
      nextSource = { format: 'xlsx', bytes };
    } else {
      if (!this.path) throw new Error('저장할 경로가 없습니다');
      const text = serializeCsv(workbookToCsvRows(this.wb), src.dialect);
      await this.io.writeText(this.path, { text, encoding: src.encoding, bom: src.bom });
    }
    this.path = nextPath;
    this.source = nextSource;
    this.wb = nextWb;
    clearEdited(this.wb);
    this.base = cloneWorkbook(this.wb);
    this.steps = [];
    this.redoStack = [];
    this.changed();
  }
}
