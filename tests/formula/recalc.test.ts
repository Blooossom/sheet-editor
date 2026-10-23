import { it, expect } from 'vitest';
import { recalcWorkbook } from '../../src/formula/recalc';
import { applyOp, makeSetOp, getCell } from '../../src/core/ops';
import { makeWorkbook } from '../helpers/model';
it('값을 바꾸면 참조하는 수식이 다시 계산되고, 모르는 함수는 기존 값 유지', () => {
  const wb = makeWorkbook('xlsx', [
    { name: 'S', grid: [[2, 3, { f: 'A1*B1', v: 6 }, { f: 'FOOBAR(A1)', v: 7 }, { f: 'C1+Data!A1', v: 16 }]] },
    { name: 'Data', grid: [[10]] },
  ]);
  applyOp(wb, makeSetOp(wb, 0, 0, 0, '4'));
  recalcWorkbook(wb);
  expect(getCell(wb.sheets[0], 0, 2)!.value).toBe(12);
  expect(getCell(wb.sheets[0], 0, 3)!.value).toBe(7);
  expect(getCell(wb.sheets[0], 0, 4)!.value).toBe(22);
});
it('오류·순환·범위·csv', () => {
  const wb = makeWorkbook('xlsx', [
    { name: 'S', grid: [
      [1, 2, { f: 'A1/0', v: null }, { f: 'SUM(A:A)', v: null }, { f: 'E1+1', v: 9 }],
      [3, null, { f: 'SUM(A1:B2)', v: null }],
    ] },
  ]);
  recalcWorkbook(wb);
  const s = wb.sheets[0];
  expect(getCell(s, 0, 2)!.value).toBe('#DIV/0!');
  expect(getCell(s, 0, 3)!.value).toBe(4);
  expect(getCell(s, 0, 4)!.value).toBe(9); // 순환: 기존 값 유지
  expect(getCell(s, 1, 2)!.value).toBe(6);
  const csv = makeWorkbook('csv', [{ name: 'S', grid: [[1, { f: 'A1+1', v: 5 }]] }]);
  recalcWorkbook(csv);
  expect(getCell(csv.sheets[0], 0, 1)!.value).toBe(5);
});
it('성능: 2만 행 값 + 수식 50개(열 전체 SUM 1개 포함)가 500ms 안에 계산된다', () => {
  const grid: (number | { f: string; v: null })[][] = [];
  for (let i = 0; i < 20000; i++) grid.push([i, i * 2]);
  grid[0].push({ f: 'SUM(A:A)', v: null });
  for (let i = 1; i < 50; i++) grid[i].push({ f: `SUM(A${i * 100}:A${i * 100 + 199})+B${i + 1}`, v: null });
  const wb = makeWorkbook('xlsx', [{ name: 'S', grid }]);
  const t = performance.now();
  recalcWorkbook(wb);
  expect(performance.now() - t).toBeLessThan(500);
  expect(getCell(wb.sheets[0], 0, 2)!.value).toBe(199990000);
  expect(getCell(wb.sheets[0], 1, 2)!.value).toBe(39700 + 2); // SUM(A100:A299)=99..298 합 + B2
});
