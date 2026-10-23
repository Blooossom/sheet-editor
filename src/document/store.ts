import { applyOp, cloneWorkbook, getCell, makeSetOp, replay, structuralBlockReason } from '../core/ops';
import { displayText, editText } from '../core/display';
import type { Op, StructuralKind, WorkbookModel } from '../core/types';
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
    sheets: [{ name: 'Sheet1', cells: new Map(), rowCount: 0, colCount: 0, blockers: [], tables: [] }],
    dateStyleIds: new Set(),
    date1904: false,
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
  for (const [si, sheet] of out.sheets.entries()) {
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
