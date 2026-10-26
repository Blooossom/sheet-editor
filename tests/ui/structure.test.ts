import { describe, expect, it } from 'vitest';
import { argsToOps, shiftFilterCol, structuralArgs } from '../../src/ui/structure';
import type { Selection } from '../../src/ui/selectionTypes';

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

describe('stale 선택', () => {
  it('보이는 행 밖으로 밀려난 선택은 구조 연산 대상이 없다', () => {
    expect(structuralArgs('deleteShortcut', sel({ stale: true }))).toBeNull();
    expect(structuralArgs('insertBelow', sel({ stale: true }))).toBeNull();
  });
});

describe('shiftFilterCol', () => {
  const ins = (at: number, count: number) => ({ kind: 'insertCols' as const, sheet: 0, at, count });
  const del = (at: number, count: number) => ({ kind: 'deleteCols' as const, sheet: 0, at, count });
  it('삽입', () => {
    expect(shiftFilterCol(2, 0, [ins(2, 1)])).toBe(3);
    expect(shiftFilterCol(2, 0, [ins(3, 1)])).toBe(2);
    expect(shiftFilterCol(2, 0, [ins(0, 2)])).toBe(4);
  });
  it('삭제', () => {
    expect(shiftFilterCol(2, 0, [del(0, 1)])).toBe(1);
    expect(shiftFilterCol(2, 0, [del(2, 1)])).toBeNull();
    expect(shiftFilterCol(2, 0, [del(1, 2)])).toBeNull();
    expect(shiftFilterCol(2, 0, [del(3, 1)])).toBe(2);
  });
  it('undo는 반대로 적용하고 다른 시트·행 연산은 무시한다', () => {
    expect(shiftFilterCol(3, 0, [ins(2, 1)], true)).toBe(2);
    expect(shiftFilterCol(2, 0, [del(0, 1)], true)).toBe(3);
    expect(shiftFilterCol(2, 1, [ins(0, 1)])).toBe(2);
    expect(shiftFilterCol(2, 0, [{ kind: 'insertRows', sheet: 0, at: 0, count: 1 }])).toBe(2);
  });
  it('여러 열 삭제 인자를 연산으로 바꿔 적용한다', () => {
    expect(shiftFilterCol(5, 0, argsToOps(0, 'deleteCols', [1, 3]))).toBe(3);
    expect(shiftFilterCol(3, 0, argsToOps(0, 'deleteCols', [1, 3]))).toBeNull();
    expect(shiftFilterCol(1, 0, argsToOps(0, 'insertCols', [1, 2]))).toBe(3);
  });
});
