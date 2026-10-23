import { describe, it, expect } from 'vitest';
import { parseInput, makeSetOp, applyOp, structuralBlockReason, replay, getCell, cloneWorkbook } from '../../src/core/ops';
import { makeWorkbook } from '../helpers/model';
const X = { kind: 'xlsx' as const, isDate: false, date1904: false };

describe('parseInput', () => {
  it('숫자처럼 보이는 문자열을 보존한다', () => {
    expect(parseInput('00123', X)).toMatchObject({ value: '00123', text: true });
    expect(parseInput('1234567890123456789', X)).toMatchObject({ value: '1234567890123456789', text: true });
    expect(parseInput('42', { ...X, prev: { value: 'x', text: true } })).toMatchObject({ value: '42', text: true });
    expect(parseInput('3.5', X)).toMatchObject({ value: 3.5 });
  });
  it('csv는 항상 문자열', () => {
    expect(parseInput('=A1', { ...X, kind: 'csv' })).toEqual({ value: '=A1', text: true, edited: true });
    expect(parseInput('007', { ...X, kind: 'csv' })).toMatchObject({ value: '007' });
  });
  it('수식, 날짜, 스타일 유지, 지우기', () => {
    expect(parseInput('=B3*C3', X)).toMatchObject({ formula: 'B3*C3', value: null });
    expect(parseInput('2026-10-23', { ...X, isDate: true })).toMatchObject({ value: 46318 });
    expect(parseInput('hi', { ...X, prev: { value: 1, styleId: 4 } })).toMatchObject({ value: 'hi', styleId: 4 });
    expect(parseInput('', { ...X, prev: { value: 1, styleId: 4 } })).toEqual({ value: null, styleId: 4, edited: true });
    expect(parseInput('', X)).toBeNull();
  });
});

describe('applyOp', () => {
  it('행 삽입: 셀과 모든 시트의 수식을 옮긴다', () => {
    const wb = makeWorkbook('xlsx', [
      { name: 'Sheet1', grid: [[1], [2], [{ f: 'A1+A2', v: 3 }]] },
      { name: 'Sum', grid: [[{ f: 'Sheet1!A3*2', v: 6 }]] },
    ]);
    applyOp(wb, { kind: 'insertRows', sheet: 0, at: 1, count: 1 });
    expect(getCell(wb.sheets[0], 1, 0)).toBeUndefined();
    expect(getCell(wb.sheets[0], 3, 0)?.formula).toBe('A1+A3');
    expect(getCell(wb.sheets[1], 0, 0)?.formula).toBe('Sheet1!A4*2');
    expect(wb.sheets[0].rowCount).toBe(4);
  });
  it('열 삭제: 셀을 당기고 지워진 참조는 #REF!', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'Sheet1', grid: [[1, 2, { f: 'A1+B1', v: 3 }]] }]);
    applyOp(wb, { kind: 'deleteCols', sheet: 0, at: 1, count: 1 });
    expect(getCell(wb.sheets[0], 0, 1)?.formula).toBe('A1+#REF!');
    expect(wb.sheets[0].colCount).toBe(2);
  });
  it('set은 edited 표시를 남기고, replay는 순차 적용과 같다', () => {
    const base = makeWorkbook('csv', [{ name: 'data', grid: [['a', 'b'], ['1', '2']] }]);
    const ops = [makeSetOp(base, 0, 1, 1, '9'), { kind: 'deleteRows' as const, sheet: 0, at: 0, count: 1 }];
    const out = replay(base, ops);
    expect(getCell(out.sheets[0], 0, 1)).toMatchObject({ value: '9', edited: true });
    expect(getCell(base.sheets[0], 1, 1)?.value).toBe('2'); // base는 그대로
  });
});

describe('structuralBlockReason', () => {
  it('차단 요소와 표 열 범위', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1]] }]);
    wb.sheets[0].blockers = ['그림·차트'];
    expect(structuralBlockReason(wb, { kind: 'insertRows', sheet: 0, at: 0, count: 1 })).toContain('그림·차트');
    wb.sheets[0].blockers = [];
    wb.sheets[0].tableRanges = ['B1:D10'];
    expect(structuralBlockReason(wb, { kind: 'deleteCols', sheet: 0, at: 2, count: 1 })).toContain('표');
    expect(structuralBlockReason(wb, { kind: 'insertCols', sheet: 0, at: 1, count: 1 })).toBeNull();
    expect(structuralBlockReason(wb, { kind: 'insertRows', sheet: 0, at: 3, count: 1 })).toBeNull();
  });
  it('표 경계', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1]] }]);
    wb.sheets[0].tableRanges = ['B1:D10'];
    expect(structuralBlockReason(wb, { kind: 'insertCols', sheet: 0, at: 2, count: 1 })).toContain('표');
    expect(structuralBlockReason(wb, { kind: 'insertCols', sheet: 0, at: 3, count: 1 })).toContain('표');
    expect(structuralBlockReason(wb, { kind: 'insertCols', sheet: 0, at: 4, count: 1 })).toBeNull();
    expect(structuralBlockReason(wb, { kind: 'deleteCols', sheet: 0, at: 0, count: 2 })).toContain('표');
    expect(structuralBlockReason(wb, { kind: 'deleteCols', sheet: 0, at: 3, count: 3 })).toContain('표');
    expect(structuralBlockReason(wb, { kind: 'deleteCols', sheet: 0, at: 4, count: 1 })).toBeNull();
  });
});

describe('parseInput 추가 규칙', () => {
  it('부호 있는 앞자리 0 문자열 보존', () => {
    expect(parseInput('-00123', X)).toMatchObject({ value: '-00123', text: true });
    expect(parseInput('-0123', X)).toMatchObject({ value: '-0123', text: true });
    expect(parseInput('-0.5', X)).toMatchObject({ value: -0.5 });
  });
  it('유한하지 않은 수는 문자열, -0은 0', () => {
    expect(parseInput('1e400', X)).toMatchObject({ value: '1e400', text: true });
    expect(Object.is(parseInput('-0', X)?.value, 0)).toBe(true);
  });
  it('텍스트 상속은 수식 없는 이전 셀만', () => {
    const prev = { formula: 'A1', value: 'x', text: true };
    expect(parseInput('42', { ...X, prev })).toMatchObject({ value: 42 });
    expect(parseInput('42', { ...X, prev })?.text).toBeUndefined();
  });
  it('빈 수식 "="는 문자열', () => {
    expect(parseInput('=', X)).toMatchObject({ value: '=', text: true });
    expect(parseInput('=', X)?.formula).toBeUndefined();
  });
  it('날짜: 1904, 시간, 잘못된 날짜', () => {
    const D = { ...X, isDate: true };
    expect(parseInput('2026-10-23', { ...D, date1904: true })).toMatchObject({ value: 46318 - 1462 });
    expect(parseInput('2026-10-23 12:00', D)).toMatchObject({ value: 46318.5 });
    expect(parseInput('2026-13-40', D)).toMatchObject({ value: '2026-13-40', text: true });
  });
});

describe('applyOp 추가', () => {
  it('set은 op.cell과 모델을 별칭 없이 복사한다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1]] }]);
    const op = { kind: 'set' as const, sheet: 0, row: 0, col: 0, cell: { value: 5 } };
    applyOp(wb, op);
    op.cell.value = 99;
    expect(getCell(wb.sheets[0], 0, 0)?.value).toBe(5);
    getCell(wb.sheets[0], 0, 0)!.value = 7;
    expect(op.cell.value).toBe(99);
  });
  it('count < 1 구조 연산은 예외', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1]] }]);
    expect(() => applyOp(wb, { kind: 'insertRows', sheet: 0, at: 0, count: 0 })).toThrow();
  });
  it('행 삭제: 다른 시트 수식과 rowCount', () => {
    const wb = makeWorkbook('xlsx', [
      { name: 'S1', grid: [[1], [2], [3], [{ f: 'A3', v: 3 }]] },
      { name: 'S2', grid: [[{ f: 'S1!A3+S1!A2', v: 5 }]] },
    ]);
    applyOp(wb, { kind: 'deleteRows', sheet: 0, at: 1, count: 1 });
    expect(getCell(wb.sheets[1], 0, 0)?.formula).toBe('S1!A2+#REF!');
    expect(getCell(wb.sheets[0], 2, 0)?.formula).toBe('A2');
    expect(wb.sheets[0].rowCount).toBe(3);
  });
  it('열 삽입', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1, 2, { f: 'A1+B1', v: 3 }]] }]);
    applyOp(wb, { kind: 'insertCols', sheet: 0, at: 1, count: 2 });
    expect(getCell(wb.sheets[0], 0, 4)?.formula).toBe('A1+D1');
    expect(getCell(wb.sheets[0], 0, 1)).toBeUndefined();
    expect(wb.sheets[0].colCount).toBe(5);
  });
  it('rowCount 이상 위치 삽입은 rowCount 유지', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1], [2]] }]);
    applyOp(wb, { kind: 'insertRows', sheet: 0, at: 5, count: 3 });
    expect(wb.sheets[0].rowCount).toBe(2);
  });
  it('tableRanges 이동과 삭제 시 제거', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1]] }]);
    wb.sheets[0].tableRanges = ['A2:C5', 'A2'];
    applyOp(wb, { kind: 'insertRows', sheet: 0, at: 0, count: 1 });
    expect(wb.sheets[0].tableRanges).toEqual(['A3:C6', 'A3']);
    applyOp(wb, { kind: 'deleteRows', sheet: 0, at: 2, count: 1 });
    expect(wb.sheets[0].tableRanges).toEqual(['A3:C5']);
  });
});

describe('replay 격리', () => {
  it('base 불변, 순차 적용과 동일, 두 번 해도 동일', () => {
    const base = makeWorkbook('xlsx', [
      { name: 'S1', grid: [[1], [2], [{ f: 'A1+A2', v: 3 }]] },
      { name: 'S2', grid: [[{ f: 'S1!A3', v: 3 }]] },
    ]);
    const ops = [
      { kind: 'insertRows' as const, sheet: 0, at: 0, count: 1 },
      makeSetOp(base, 0, 0, 0, '=A2'),
      { kind: 'deleteRows' as const, sheet: 0, at: 1, count: 1 },
    ];
    const out = replay(base, ops);
    expect(getCell(base.sheets[0], 2, 0)?.formula).toBe('A1+A2');
    expect(getCell(base.sheets[1], 0, 0)?.formula).toBe('S1!A3');
    const seq = cloneWorkbook(base);
    ops.forEach((op) => applyOp(seq, op));
    expect(out).toEqual(seq);
    expect(replay(base, ops)).toEqual(out);
  });
});
