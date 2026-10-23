import FormulaParser from 'fast-formula-parser';
import type { Cell, CellValue, WorkbookModel } from '../core/types';

type Pos = { sheet: string; row: number; col: number };

// 수식 셀 값을 모델에서 다시 계산한다. xlsx만 대상이며, 계산할 수 없는 셀은 기존 값을 유지한다.
export function recalcWorkbook(wb: WorkbookModel): void {
  if (wb.kind !== 'xlsx') return;
  const sheetIndex = new Map<string, number>();
  wb.sheets.forEach((s, i) => sheetIndex.set(s.name.toLowerCase(), i));

  const done = new Set<string>();
  const inProgress = new Set<string>();
  const cyclic = new Set<string>();
  // 같은 범위를 여러 수식이 참조하므로 한 번 만든 배열을 재사용한다 (계산 중 값은 바뀌지 않는다)
  const rangeCache = new Map<string, CellValue[][]>();
  // 파서는 내부 상태(position)를 가지므로, 중첩 계산(참조한 수식 셀을 먼저 계산)마다 다른 인스턴스를 쓴다
  const parsers: FormulaParser[] = [];
  let depth = 0;

  const keyOf = (s: number, r: number, c: number) => `${s}:${r}:${c}`;

  const valueAt = (s: number, r: number, c: number): CellValue => {
    const cell = wb.sheets[s].cells.get(r)?.get(c);
    if (!cell) return null;
    if (cell.formula !== undefined) evaluate(s, r, c, cell);
    return cell.value;
  };

  const resolveSheet = (name: string | undefined, fallback: string): number => sheetIndex.get((name ?? fallback).toLowerCase()) ?? -1;

  const makeParser = (): FormulaParser =>
    new FormulaParser({
      onCell: ({ sheet, row, col }) => {
        const s = resolveSheet(sheet, '');
        return s < 0 ? null : valueAt(s, row - 1, col - 1);
      },
      onRange: ({ sheet, from, to }) => {
        const s = resolveSheet(sheet, '');
        if (s < 0) return [[null]];
        const cacheKey = `${s}:${from.row}:${from.col}:${to.row}:${to.col}`;
        const hit = rangeCache.get(cacheKey);
        if (hit) return hit;
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
        rangeCache.set(cacheKey, out);
        return out;
      },
    });

  function evaluate(s: number, r: number, c: number, cell: Cell): void {
    const key = keyOf(s, r, c);
    if (done.has(key)) return;
    if (inProgress.has(key)) {
      cyclic.add(key);
      return;
    }
    inProgress.add(key);
    // 첫 번째 시트 이름은 파서가 위치의 시트로 보충하므로 이름 그대로 넘긴다
    const pos: Pos = { sheet: wb.sheets[s].name, row: r + 1, col: c + 1 };
    const parser = (parsers[depth] ??= makeParser());
    depth++;
    try {
      const result = parser.parse(cell.formula!, pos);
      if (!cyclic.has(key)) {
        const next = toValue(result);
        if (next !== undefined) cell.value = next;
      }
    } catch {
      // 계산 못 한 수식은 파일에 있던 값을 그대로 둔다
    } finally {
      depth--;
      inProgress.delete(key);
      done.add(key);
    }
  }

  wb.sheets.forEach((sh, s) => {
    for (const [r, row] of sh.cells) {
      for (const [c, cell] of row) if (cell.formula !== undefined) evaluate(s, r, c, cell);
    }
  });
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
