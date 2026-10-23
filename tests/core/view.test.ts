import { it, expect } from 'vitest';
import { buildRowView, distinctValues } from '../../src/core/view';
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
