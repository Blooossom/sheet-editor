import { it, expect } from 'vitest';
import { displayText, editText } from '../../src/core/display';
import { makeWorkbook } from '../helpers/model';
it('표시 형식', () => {
  const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1]] }]); wb.dateStyleIds.add(1);
  expect(displayText({ value: 46318, styleId: 1 }, wb)).toBe('2026-10-23');
  expect(displayText({ value: 46318.5, styleId: 1 }, wb)).toBe('2026-10-23 12:00:00');
  expect(editText({ value: 46318, styleId: 1 }, wb)).toBe('2026-10-23');
  expect(displayText({ value: 0.1 + 0.2 }, wb)).toBe('0.3');
  expect(displayText({ value: true }, wb)).toBe('TRUE');
  expect(editText({ value: 6, formula: 'A1*B1' }, wb)).toBe('=A1*B1');
  expect(displayText(undefined, wb)).toBe('');
});
it('1904 기준·초 반올림·날짜 아닌 값', () => {
  const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1]] }]); wb.dateStyleIds.add(1);
  wb.date1904 = true;
  expect(displayText({ value: 0, styleId: 1 }, wb)).toBe('1904-01-01');
  wb.date1904 = false;
  expect(displayText({ value: 46318.49999999, styleId: 1 }, wb)).toBe('2026-10-23 12:00:00');
  expect(displayText({ value: 'abc', styleId: 1 }, wb)).toBe('abc');
  expect(displayText({ value: 5, styleId: 2 }, wb)).toBe('5');
  expect(editText({ value: 'x', text: true }, wb)).toBe('x');
  expect(editText({ value: false }, wb)).toBe('FALSE');
  expect(editText(undefined, wb)).toBe('');
});
