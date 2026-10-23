import { it, expect } from 'vitest';
import { unzipSync, strFromU8 } from 'fflate';
import { createXlsx } from '../../src/xlsx/create';
import { readXlsx } from '../../src/xlsx/read';
import { makeWorkbook } from '../helpers/model';
it('만든 파일을 다시 읽으면 값이 같다', () => {
  const wb = makeWorkbook('xlsx', [{ name: 'Sheet1', grid: [['이름', 3], [null, { f: 'B1*2', v: null }]] }]);
  const back = readXlsx(createXlsx(wb)).sheets[0];
  expect(back.name).toBe('Sheet1');
  expect(back.cells.get(0)!.get(0)!.value).toBe('이름');
  expect(back.cells.get(0)!.get(1)!.value).toBe(3);
  expect(back.cells.get(1)!.get(1)!.formula).toBe('B1*2');
});
it('패키지 구조: Content_Types 첫 항목, 채움 2개, dimension', () => {
  const wb = makeWorkbook('xlsx', [{ name: 'Sheet1', grid: [['a', 1], [null, null, 5]] }]);
  const files = unzipSync(createXlsx(wb));
  expect(Object.keys(files)[0]).toBe('[Content_Types].xml');
  expect(strFromU8(files['xl/styles.xml']).match(/<fill>/g)).toHaveLength(2);
  expect(strFromU8(files['xl/worksheets/sheet1.xml'])).toContain('<dimension ref="A1:C2"/>');
  expect(strFromU8(files['xl/workbook.xml'])).toContain('fullCalcOnLoad="1"');
});
it('빈 통합 문서도 시트 하나로 읽힌다', () => {
  const wb = makeWorkbook('xlsx', [{ name: 'Sheet1', grid: [] }]);
  const back = readXlsx(createXlsx(wb));
  expect(back.sheets).toHaveLength(1);
  expect(back.sheets[0].name).toBe('Sheet1');
});
