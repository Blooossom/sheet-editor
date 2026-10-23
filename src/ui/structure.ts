import type { Op, StructuralKind } from '../core/types';
import type { Selection } from './Grid';

export type StructuralAction =
  | 'insertAbove'
  | 'insertBelow'
  | 'deleteRows'
  | 'insertLeft'
  | 'insertRight'
  | 'deleteCols'
  | 'insertShortcut'
  | 'deleteShortcut';

const range = (at: number, count: number) => Array.from({ length: count }, (_, i) => at + i);

/** 메뉴 동작을 store.structural 인자로 바꾼다. 번호는 모두 원래 행 번호·시트 열 번호다. */
export function structuralArgs(
  action: StructuralAction,
  sel: Selection,
): { kind: StructuralKind; rows: number[] } | null {
  if (sel.stale) return null; // 선택이 보이는 행 밖으로 밀려났다. 엉뚱한 행을 건드리지 않는다.
  if (action === 'insertShortcut') action = sel.cols.length > 0 ? 'insertLeft' : 'insertAbove';
  if (action === 'deleteShortcut') action = sel.cols.length > 0 ? 'deleteCols' : 'deleteRows';
  const rows = sel.rows.length > 0 ? sel.rows : [sel.row];
  const cols = sel.cols.length > 0 ? sel.cols : [sel.col];
  switch (action) {
    case 'insertAbove':
      return { kind: 'insertRows', rows };
    case 'insertBelow':
      return { kind: 'insertRows', rows: range(Math.max(...rows) + 1, rows.length) };
    case 'deleteRows':
      return { kind: 'deleteRows', rows };
    case 'insertLeft':
      return { kind: 'insertCols', rows: cols };
    case 'insertRight':
      return { kind: 'insertCols', rows: range(Math.max(...cols) + 1, cols.length) };
    case 'deleteCols':
      return { kind: 'deleteCols', rows: cols };
  }
}

/**
 * 구조 연산 뒤 필터 열 번호. 필터 열이 삭제되면 null.
 * undo = true면 연산을 거꾸로(삽입↔삭제) 적용한 것으로 본다.
 */
export function shiftFilterCol(col: number, sheet: number, ops: readonly Op[], undo = false): number | null {
  let c: number | null = col;
  for (const op of undo ? [...ops].reverse() : ops) {
    if (c === null) break;
    if (op.kind !== 'insertCols' && op.kind !== 'deleteCols') continue;
    if (op.sheet !== sheet) continue;
    const insert = (op.kind === 'insertCols') !== undo;
    if (insert) {
      if (op.at <= c) c += op.count;
    } else if (c >= op.at && c < op.at + op.count) {
      c = null;
    } else if (op.at < c) {
      c -= op.count;
    }
  }
  return c;
}

/** structuralArgs 결과를 shiftFilterCol이 쓰는 연산 목록으로 바꾼다. */
export function argsToOps(sheet: number, kind: StructuralKind, rows: number[]): Op[] {
  const sorted = [...new Set(rows)].sort((a, b) => a - b);
  if (kind === 'insertRows' || kind === 'insertCols') return [{ kind, sheet, at: sorted[0], count: sorted.length }];
  return sorted.reverse().map((at) => ({ kind, sheet, at, count: 1 }));
}
