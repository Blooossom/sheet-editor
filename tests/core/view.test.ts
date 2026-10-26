import { describe, it, expect } from 'vitest';
import { buildRowView, distinctValues, keepRows, viewNeedsUpdate } from '../../src/core/view';
import { makeWorkbook } from '../helpers/model';
const wb = makeWorkbook('csv', [{ name: 'd', grid: [['id', 'name', 'status'], ['1', '김민지', 'active'], ['2', '이준호', 'inactive'], ['3', '김서연', 'active']] }]);
const s = wb.sheets[0];
it('검색·필터는 원래 행 번호를 돌려준다', () => {
  expect(buildRowView(s, wb, '', null)).toBeNull();
  expect(buildRowView(s, wb, '김', null)).toEqual([1, 3]);
  expect(buildRowView(s, wb, 'ACTIVE', null)).toEqual([1, 2, 3]); // 대소문자 무시, 'inactive'도 포함
  expect(buildRowView(s, wb, '', { col: 2, values: new Set(['active']) })).toEqual([1, 3]);
  expect(buildRowView(s, wb, '서연', { col: 2, values: new Set(['active']) })).toEqual([3]);
  expect(distinctValues(s, wb, 2)).toEqual([{ value: 'active', count: 2 }, { value: 'inactive', count: 1 }, { value: 'status', count: 1 }]);
});

it('distinctValues는 시작 행부터 센다', () => {
  
  expect(distinctValues(s, wb, 2, 500, 1)).toEqual([{ value: 'active', count: 2 }, { value: 'inactive', count: 1 }]);
});

describe('되돌리기 뒤 보이는 행 다시 계산', () => {
  it('구조 연산(행·열 넣기·지우기)이 있을 때만 다시 계산한다', () => {
    expect(viewNeedsUpdate([{ kind: 'set', sheet: 0, row: 0, col: 0, cell: null }])).toBe(false);
    expect(viewNeedsUpdate([{ kind: 'style', sheet: 0, rows: [1], c1: 0, c2: 0, patch: { font: { bold: true } } }])).toBe(false);
    expect(viewNeedsUpdate([{ kind: 'colWidth', sheet: 0, c1: 0, c2: 0, width: 9 }, { kind: 'merge', sheet: 0, range: { r1: 0, c1: 0, r2: 1, c2: 1 } }])).toBe(false);
    expect(viewNeedsUpdate([{ kind: 'set', sheet: 0, row: 0, col: 0, cell: null }, { kind: 'deleteRows', sheet: 0, at: 1, count: 1 }])).toBe(true);
    expect(viewNeedsUpdate([{ kind: 'insertCols', sheet: 0, at: 0, count: 1 }])).toBe(true);
    expect(viewNeedsUpdate(null)).toBe(false);
  });
  it('내용이 같으면 이전 배열을 그대로 쓴다(선택이 처음 칸으로 돌아가지 않게)', () => {
    const prev = [0, 2, 3];
    expect(keepRows(prev, [0, 2, 3])).toBe(prev);
    expect(keepRows(prev, [0, 2])).toEqual([0, 2]);
    expect(keepRows(null, null)).toBeNull();
    expect(keepRows(prev, null)).toBeNull();
  });
});
