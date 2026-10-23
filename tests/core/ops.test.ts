import { describe, it, expect } from 'vitest';
import { parseInput, makeSetOp, applyOp, structuralBlockReason, replay, getCell } from '../../src/core/ops';
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
});
