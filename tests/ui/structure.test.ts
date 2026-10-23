import { describe, expect, it } from 'vitest';
import { structuralArgs } from '../../src/ui/structure';
import type { Selection } from '../../src/ui/Grid';

const sel = (p: Partial<Selection>): Selection => ({ sheet: 0, row: 4, col: 2, rows: [], cols: [], ...p });

describe('structuralArgs', () => {
  it('셀 선택은 그 셀의 행·열을 대상으로 한다', () => {
    expect(structuralArgs('insertAbove', sel({}))).toEqual({ kind: 'insertRows', rows: [4] });
    expect(structuralArgs('insertBelow', sel({}))).toEqual({ kind: 'insertRows', rows: [5] });
    expect(structuralArgs('deleteRows', sel({}))).toEqual({ kind: 'deleteRows', rows: [4] });
    expect(structuralArgs('insertLeft', sel({}))).toEqual({ kind: 'insertCols', rows: [2] });
    expect(structuralArgs('insertRight', sel({}))).toEqual({ kind: 'insertCols', rows: [3] });
    expect(structuralArgs('deleteCols', sel({}))).toEqual({ kind: 'deleteCols', rows: [2] });
  });

  it('여러 행을 선택하면 아래 삽입은 마지막 행 다음에 선택한 개수만큼 넣는다', () => {
    const s = sel({ rows: [10, 3, 7] });
    expect(structuralArgs('insertBelow', s)).toEqual({ kind: 'insertRows', rows: [11, 12, 13] });
    expect(structuralArgs('insertAbove', s)).toEqual({ kind: 'insertRows', rows: [10, 3, 7] });
    expect(structuralArgs('deleteRows', s)).toEqual({ kind: 'deleteRows', rows: [10, 3, 7] });
  });

  it('여러 열을 선택하면 오른쪽 삽입은 마지막 열 다음에 개수만큼 넣는다', () => {
    const s = sel({ cols: [1, 4] });
    expect(structuralArgs('insertRight', s)).toEqual({ kind: 'insertCols', rows: [5, 6] });
    expect(structuralArgs('deleteCols', s)).toEqual({ kind: 'deleteCols', rows: [1, 4] });
  });

  it('단축키는 열이 선택돼 있으면 열을, 아니면 행을 대상으로 한다', () => {
    expect(structuralArgs('insertShortcut', sel({ cols: [2, 3] }))).toEqual({ kind: 'insertCols', rows: [2, 3] });
    expect(structuralArgs('deleteShortcut', sel({ cols: [2] }))).toEqual({ kind: 'deleteCols', rows: [2] });
    expect(structuralArgs('insertShortcut', sel({ rows: [8, 9] }))).toEqual({ kind: 'insertRows', rows: [8, 9] });
    expect(structuralArgs('deleteShortcut', sel({}))).toEqual({ kind: 'deleteRows', rows: [4] });
  });
});
