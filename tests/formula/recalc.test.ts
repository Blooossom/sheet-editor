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
it('성능: SUM(A:A) 3개(5만 행) + 작은 수식 47개가 500ms 안에 계산된다', () => {
  const grid: (number | { f: string; v: null })[][] = [];
  for (let i = 0; i < 50000; i++) grid.push([i, i * 2]);
  for (let i = 0; i < 3; i++) grid[i].push({ f: 'SUM(A:A)', v: null });
  for (let i = 3; i < 50; i++) grid[i].push({ f: `A${i + 1}*2+B${i + 1}`, v: null });
  const wb = makeWorkbook('xlsx', [{ name: 'S', grid }]);
  const t = performance.now();
  recalcWorkbook(wb);
  expect(performance.now() - t).toBeLessThan(500);
  expect(getCell(wb.sheets[0], 2, 2)!.value).toBe(1249975000);
  expect(getCell(wb.sheets[0], 3, 2)!.value).toBe(3 * 2 + 6);
});
it('빠른 경로가 파서와 같은 결과를 낸다 (텍스트·불리언·빈 칸 섞임)', () => {
  const wb = makeWorkbook('xlsx', [
    { name: 'S', grid: [[1], ['x'], [null], [4], [2.5]] },
    { name: 'My Sheet', grid: [[1], ['x'], [null], [4], [2.5]] },
    { name: 'T', grid: [] },
  ]);
  for (const sh of wb.sheets.slice(0, 2)) {
    sh.cells.set(5, new Map([[0, { value: true }]]));
    sh.rowCount = 6;
  }
  const fns = ['SUM', 'COUNT', 'COUNTA', 'AVERAGE', 'MIN', 'MAX'];
  const refs = ['S!A1:A6', 'S!A:A', "'My Sheet'!A1:A6", "'My Sheet'!A:A", 'S!A7:A9', 'S!Z1:Z3', 'S!A1', 'S!1:1'];
  const t = wb.sheets[2];
  const setF = (r: number, c: number, f: string) => {
    if (!t.cells.has(r)) t.cells.set(r, new Map());
    t.cells.get(r)!.set(c, { formula: f, value: null });
  };
  fns.forEach((fn, i) => refs.forEach((ref, j) => {
    setF(i, j * 2, `${fn}(${ref})`); // 빠른 경로
    setF(i, j * 2 + 1, `${fn}(${ref})+0`); // 파서
  }));
  t.rowCount = fns.length; t.colCount = refs.length * 2;
  recalcWorkbook(wb);
  fns.forEach((fn, i) => refs.forEach((ref, j) => {
    if (getCell(t, i, j * 2 + 1)!.value === null) return; // 라이브러리가 못 계산한 경우(COUNTA 등)는 비교 대상이 없다
    if (getCell(t, i, j * 2)!.value === '#DIV/0!') return; // 빈 범위 AVERAGE: 라이브러리는 #VALUE!, Excel은 #DIV/0!
    expect(getCell(t, i, j * 2)!.value, `${fn}(${ref})`).toEqual(getCell(t, i, j * 2 + 1)!.value);
  }));
  expect(getCell(t, 0, 0)!.value).toBe(7.5); // 텍스트·불리언 무시
  expect([1, 3, 4, 5].map((i) => getCell(t, i, 0)!.value)).toEqual([3, 2.5, 1, 4]); // COUNT, AVERAGE, MIN, MAX
  expect(getCell(t, 2, 0)!.value).toBe(5); // COUNTA: 1,'x',4,2.5,true
  expect(getCell(t, 3, 8)!.value).toBe('#DIV/0!'); // 빈 범위 AVERAGE
});
it('서로 참조하는 두 셀은 모두 기존 값을 유지한다', () => {
  const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[{ f: 'B1+1', v: 7 }, { f: 'A1+1', v: 8 }, { f: 'A1*2', v: 0 }]] }]);
  recalcWorkbook(wb);
  expect(getCell(wb.sheets[0], 0, 0)!.value).toBe(7);
  expect(getCell(wb.sheets[0], 0, 1)!.value).toBe(8);
});
it('5000개 수식 사슬은 앞·뒤 방향 모두 던지지 않고 계산된다', () => {
  const n = 5000;
  const fwd = Array.from({ length: n }, (_, i) => [i === 0 ? 1 : { f: `A${i}+1`, v: null }]);
  const rev = Array.from({ length: n }, (_, i) => [i === n - 1 ? 1 : { f: `A${i + 2}+1`, v: null }]);
  const wb = makeWorkbook('xlsx', [{ name: 'F', grid: fwd }, { name: 'R', grid: rev }]);
  expect(() => recalcWorkbook(wb)).not.toThrow();
  expect(getCell(wb.sheets[0], n - 1, 0)!.value).toBe(n);
  expect(getCell(wb.sheets[1], 0, 0)!.value).toBe(n);
});
