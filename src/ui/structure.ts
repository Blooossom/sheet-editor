import type { StructuralKind } from '../core/types';
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
export function structuralArgs(action: StructuralAction, sel: Selection): { kind: StructuralKind; rows: number[] } {
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
