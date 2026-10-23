import FormulaParser from 'fast-formula-parser';
import type { Cell, CellValue, WorkbookModel } from '../core/types';

type Pos = { sheet: string; row: number; col: number };

// 파서는 범위 셀당 약 8µs가 들어서, 큰 범위의 단순 집계는 모델에서 직접 계산한다.
const FAST_RE = /^(SUM|COUNT|COUNTA|AVERAGE|MIN|MAX)\(\s*(?:(?:'((?:[^']|'')+)'|([\w.]+))!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)\s*\)$/i;
const ERRORS = new Set(['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!']);
const MAX_DEPTH = 50; // 파서 중첩 호출이 스택을 넘지 않도록 하는 한계
const MAX_PASSES = 4;

interface Fast { fn: string; sheet: string | undefined; r1: number; r2: number; c1: number; c2: number }

function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}

function parseFast(formula: string): Fast | null {
  const m = FAST_RE.exec(formula.trim());
  if (!m) return null;
  const ref = m[4].replace(/\$/g, '');
  const sheet = m[2] !== undefined ? m[2].replace(/''/g, "'") : m[3];
  const [a, b = a] = ref.split(':');
  const cell = /^([A-Za-z]+)(\d+)$/;
  const ca = cell.exec(a);
  const cb = cell.exec(b);
  let r1: number, r2: number, c1: number, c2: number;
  if (ca && cb) {
    r1 = Number(ca[2]) - 1; r2 = Number(cb[2]) - 1; c1 = colIndex(ca[1]); c2 = colIndex(cb[1]);
  } else if (/^\d+$/.test(a)) {
    r1 = Number(a) - 1; r2 = Number(b) - 1; c1 = 0; c2 = Infinity; // 행 전체
  } else {
    r1 = 0; r2 = Infinity; c1 = colIndex(a); c2 = colIndex(b); // 열 전체
  }
  return { fn: m[1].toUpperCase(), sheet, r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
}

// 수식 셀 값을 모델에서 다시 계산한다. xlsx만 대상이며, 계산할 수 없는 셀은 기존 값을 유지한다.
export function recalcWorkbook(wb: WorkbookModel): void {
  if (wb.kind !== 'xlsx') return;
  if (!wb.sheets.some((sh) => [...sh.cells.values()].some((row) => [...row.values()].some((c) => c.formula !== undefined)))) return;

  const sheetIndex = new Map<string, number>();
  wb.sheets.forEach((s, i) => sheetIndex.set(s.name.toLowerCase(), i));

  let done = new Set<string>();
  let cyclic = new Set<string>();
  const stack: string[] = [];
  let truncated = false;
  // 파서는 내부 상태(position)를 가지므로, 중첩 계산(참조한 수식 셀을 먼저 계산)마다 다른 인스턴스를 쓴다
  const parsers: FormulaParser[] = [];

  const keyOf = (s: number, r: number, c: number) => `${s}:${r}:${c}`;

  const valueAt = (s: number, r: number, c: number): CellValue => {
    const cell = wb.sheets[s].cells.get(r)?.get(c);
    if (!cell) return null;
    if (cell.formula !== undefined) evaluate(s, r, c, cell);
    return cell.value;
  };

  const resolveSheet = (name: string | undefined): number => (name === undefined ? -1 : sheetIndex.get(name.toLowerCase()) ?? -1);

  const makeParser = (): FormulaParser =>
    new FormulaParser({
      onCell: ({ sheet, row, col }) => {
        const s = resolveSheet(sheet);
        return s < 0 ? null : valueAt(s, row - 1, col - 1);
      },
      onRange: ({ sheet, from, to }) => {
        const s = resolveSheet(sheet);
        if (s < 0) return [[null]];
        const sh = wb.sheets[s];
        // 열 전체 참조(A:A)는 to.row가 매우 커서 사용 범위로 자른다
        const lastRow = Math.min(to.row, Math.max(sh.rowCount, from.row));
        const lastCol = Math.min(to.col, Math.max(sh.colCount, from.col));
        const out: CellValue[][] = [];
        for (let r = from.row; r <= lastRow; r++) {
          const line: CellValue[] = [];
          for (let c = from.col; c <= lastCol; c++) line.push(valueAt(s, r - 1, c - 1));
          out.push(line);
        }
        return out;
      },
    });

  // 처리하지 못하면 undefined(→ 파서로), 아니면 계산 결과 (null이면 기존 값 유지는 아님: 값 자체)
  function fastEval(f: Fast, ownSheet: number): CellValue | undefined {
    const s = f.sheet === undefined ? ownSheet : resolveSheet(f.sheet);
    if (s < 0) return undefined;
    let sum = 0, count = 0, counta = 0, min = Infinity, max = -Infinity;
    for (const [r, row] of wb.sheets[s].cells) {
      if (r < f.r1 || r > f.r2) continue;
      for (const [c, cell] of row) {
        if (c < f.c1 || c > f.c2) continue;
        const v = valueAt(s, r, c);
        if (v === null) continue;
        counta++;
        if (typeof v === 'string' && cell.formula !== undefined && ERRORS.has(v)) {
          if (f.fn === 'COUNTA') continue;
          if (f.fn !== 'COUNT') return v;
          continue;
        }
        if (typeof v !== 'number') continue;
        count++; sum += v;
        if (v < min) min = v;
        if (v > max) max = v;
      }
    }
    switch (f.fn) {
      case 'SUM': return sum;
      case 'COUNT': return count;
      case 'COUNTA': return counta;
      case 'AVERAGE': return count === 0 ? '#DIV/0!' : sum / count;
      case 'MIN': return count === 0 ? 0 : min;
      default: return count === 0 ? 0 : max;
    }
  }

  function evaluate(s: number, r: number, c: number, cell: Cell): void {
    const key = keyOf(s, r, c);
    if (done.has(key)) return;
    const at = stack.indexOf(key);
    if (at >= 0) {
      for (let i = at; i < stack.length; i++) cyclic.add(stack[i]); // 순환에 속한 모든 셀
      return;
    }
    if (stack.length >= MAX_DEPTH) {
      truncated = true; // 이번 패스에서는 기존 값으로 두고 다음 패스에서 다시 계산한다
      return;
    }
    stack.push(key);
    let next: CellValue | undefined;
    try {
      const fast = parseFast(cell.formula!);
      const fastResult = fast ? fastEval(fast, s) : undefined;
      if (fastResult !== undefined) {
        next = fastResult;
      } else {
        const pos: Pos = { sheet: wb.sheets[s].name, row: r + 1, col: c + 1 };
        const parser = (parsers[stack.length - 1] ??= makeParser());
        next = toValue(parser.parse(cell.formula!, pos));
      }
    } catch {
      next = undefined; // 계산 못 한 수식(스택 초과 포함)은 파일에 있던 값을 그대로 둔다
    } finally {
      stack.pop();
    }
    if (next !== undefined && !cyclic.has(key)) cell.value = next;
    done.add(key);
  }

  for (let pass = 0; pass < MAX_PASSES; pass++) {
    done = new Set();
    cyclic = new Set();
    truncated = false;
    // 깊은 수식 사슬은 방향에 따라 앞/뒤 순서로 번갈아 계산해야 한 번에 풀린다
    const order: [number, number, number, Cell][] = [];
    wb.sheets.forEach((sh, s) => {
      for (const [r, row] of sh.cells) for (const [c, cell] of row) if (cell.formula !== undefined) order.push([s, r, c, cell]);
    });
    if (pass % 2 === 1) order.reverse();
    for (const [s, r, c, cell] of order) evaluate(s, r, c, cell);
    if (!truncated) break;
  }
}

// undefined = 기존 값 유지
function toValue(result: unknown): CellValue | undefined {
  if (Array.isArray(result)) return toValue(result[0]?.[0]);
  if (result instanceof FormulaParser.FormulaError) {
    return result.error === '#NAME?' ? undefined : result.error;
  }
  if (result === null || result === undefined) return 0;
  if (typeof result === 'number') return Number.isFinite(result) ? result : undefined;
  if (typeof result === 'string' || typeof result === 'boolean') return result;
  return undefined;
}
